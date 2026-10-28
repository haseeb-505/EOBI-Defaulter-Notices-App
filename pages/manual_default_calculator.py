
import os
import re
import tempfile
import zipfile
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import Workbook

from generator import (
    calculate_default_contribution,
    create_assessment_sheet,
    fill_notice,
)


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Manual Default Calculation",
    page_icon="🧮",
    layout="wide",
)


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>
        .main-title {
            font-size: 32px;
            font-weight: 700;
            margin-bottom: 5px;
        }

        .sub-title {
            font-size: 17px;
            color: #666;
            margin-bottom: 20px;
        }

        div.stButton > button,
        div.stFormSubmitButton > button {
            border-radius: 6px;
            font-weight: 600;
        }

        div[data-testid="stMetric"] {
            border: 1px solid #ddd;
            border-radius: 8px;
            padding: 12px;
        }

        div[data-testid="stHorizontalBlock"]
        > div:nth-child(2) div[data-testid="stMetric"] {
            border: 1px solid #2e7d32;
            background-color: #2e7d32;
        }

        div[data-testid="stHorizontalBlock"]
        > div:nth-child(2) div[data-testid="stMetricLabel"],
        div[data-testid="stHorizontalBlock"]
        > div:nth-child(2) div[data-testid="stMetricValue"] {
            color: white !important;
        }

        div[data-testid="stHorizontalBlock"]
        > div:nth-child(5) div[data-testid="stMetric"] {
            border: 1px solid #d32f2f;
            background-color: #d32f2f;
        }

        div[data-testid="stHorizontalBlock"]
        > div:nth-child(5) div[data-testid="stMetricLabel"],
        div[data-testid="stHorizontalBlock"]
        > div:nth-child(5) div[data-testid="stMetricValue"] {
            color: white !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SESSION STATE
# ============================================================

DEFAULT_STATE = {
    "manual_default_result": None,
    "manual_default_periods": None,
    "manual_default_ips": None,
    "manual_notice_zip": None,
    "manual_notice_zip_name": None,
}

for key, default_value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = default_value


# ============================================================
# HELPER FUNCTIONS
# ============================================================

MONTH_NAMES = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]


def clean_text(value):
    """Convert an input value to clean text."""

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    if isinstance(value, float) and value.is_integer():
        value = int(value)

    return str(value).strip()


def safe_filename(value):
    """Remove characters that are unsuitable for filenames."""

    value = clean_text(value)

    value = re.sub(r'[\\/:*?"<>|]+', "_", value)
    value = re.sub(r"\s+", "_", value)
    value = value.strip("._ ")

    return value or "Payment_Notice"


def get_inclusive_end_date(exclusive_end_date):
    """
    The calculator uses an exclusive end date.
    Convert it to the last month included in the period.
    """

    if exclusive_end_date.month == 1:
        return datetime(
            exclusive_end_date.year - 1,
            12,
            1,
        )

    return datetime(
        exclusive_end_date.year,
        exclusive_end_date.month - 1,
        1,
    )


def format_notice_periods(periods):
    """Format the selected periods for the Urdu notice."""

    formatted_periods = []

    for start_date, exclusive_end_date in periods:
        inclusive_end_date = get_inclusive_end_date(
            exclusive_end_date
        )

        start_text = start_date.strftime("%Y-%m")
        end_text = inclusive_end_date.strftime("%Y-%m")

        formatted_periods.append(
            f"{start_text} تا {end_text}"
        )

    return "، ".join(formatted_periods)


def build_assessment_row(
    establishment_name,
    registration_code,
    sub_code,
    ips,
    periods,
):
    """
    Build the single-establishment data row expected by
    create_assessment_sheet() in generator.py.

    The existing assessment-sheet function uses From1/To1
    through From4/To4, so this page supports up to four
    separate default periods.
    """

    row_data = {
        "Name": "",
        "Name of the Establishment": establishment_name,
        "Code": registration_code,
        "Sub_Code": sub_code or "0",
        "IPS": int(ips),
    }

    for index in range(1, 5):
        row_data[f"From{index}"] = None
        row_data[f"To{index}"] = None

    for index, (start_date, exclusive_end_date) in enumerate(
        periods,
        start=1,
    ):
        row_data[f"From{index}"] = start_date
        row_data[f"To{index}"] = exclusive_end_date

    return pd.Series(row_data)


def create_assessment_workbook(row_data, sheet_name):
    """
    Create an assessment workbook using the existing
    create_assessment_sheet() function.
    """

    workbook = Workbook()

    # Remove the default blank worksheet.
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    create_assessment_sheet(
        workbook,
        row_data,
        sheet_name,
    )

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return output.getvalue()


def create_notice_document(template_bytes, notice_data):
    """Fill the uploaded Word template and return its DOCX bytes."""

    temporary_path = None

    try:
        with tempfile.NamedTemporaryFile(
            suffix=".docx",
            delete=False,
        ) as temporary_file:
            temporary_file.write(template_bytes)
            temporary_path = temporary_file.name

        document = fill_notice(
            temporary_path,
            notice_data,
        )

        output = BytesIO()
        document.save(output)
        output.seek(0)

        return output.getvalue()

    finally:
        if temporary_path and os.path.exists(temporary_path):
            try:
                os.remove(temporary_path)
            except OSError:
                pass


def create_output_zip(
    notice_bytes,
    assessment_bytes,
    notice_filename,
    assessment_filename,
):
    """Package the individual notice and assessment sheet together."""

    output = BytesIO()

    with zipfile.ZipFile(
        output,
        mode="w",
        compression=zipfile.ZIP_DEFLATED,
    ) as zip_file:

        zip_file.writestr(
            notice_filename,
            notice_bytes,
        )

        zip_file.writestr(
            assessment_filename,
            assessment_bytes,
        )

    output.seek(0)
    return output.getvalue()


# ============================================================
# PAGE HEADER
# ============================================================

st.markdown(
    '<div class="main-title">'
    '🧮 Manual Default Calculation'
    '</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="sub-title">'
    'Calculate default contributions and automatically generate '
    'an individual Urdu payment notice together with its '
    'Defaulters Assessment Sheet.'
    '</div>',
    unsafe_allow_html=True,
)


# ============================================================
# NAVIGATION
# ============================================================

if st.button("⬅️ Back to Notice Generator"):
    st.switch_page("app.py")

st.divider()


# ============================================================
# SECTION 1: DEFAULT CALCULATION INPUTS
# ============================================================

st.header("1. Calculate Default Contribution")

col1, col2 = st.columns(2)

with col1:
    number_of_periods = st.number_input(
        "Number of Default Periods",
        min_value=1,
        max_value=4,
        value=1,
        step=1,
        key="manual_number_of_periods",
        help=(
            "Up to four separate periods are supported by the "
            "existing assessment-sheet generator."
        ),
    )

with col2:
    number_of_ips = st.number_input(
        "Number of Insured Persons (IPs)",
        min_value=1,
        value=1,
        step=1,
        key="manual_number_of_ips",
    )


# ============================================================
# SECTION 2: SELECT PERIODS
# ============================================================

st.subheader("Default Periods")

st.info(
    "Select the starting and ending months. Both months are "
    "included in the calculation. For example, July 2023 to "
    "June 2024 represents 12 months."
)

manual_periods = []

for i in range(int(number_of_periods)):

    st.markdown(f"**Period {i + 1}**")

    from_col, to_col = st.columns(2)

    with from_col:
        st.markdown("From")

        month_col, year_col = st.columns(2)

        with month_col:
            from_month = st.selectbox(
                "Starting Month",
                options=range(1, 13),
                format_func=lambda value: MONTH_NAMES[value - 1],
                index=6,
                key=f"from_month_{i}",
            )

        with year_col:
            from_year = st.number_input(
                "Starting Year",
                min_value=1900,
                max_value=2100,
                value=2023,
                step=1,
                key=f"from_year_{i}",
            )

    with to_col:
        st.markdown("To")

        month_col, year_col = st.columns(2)

        with month_col:
            to_month = st.selectbox(
                "Ending Month",
                options=range(1, 13),
                format_func=lambda value: MONTH_NAMES[value - 1],
                index=5,
                key=f"to_month_{i}",
            )

        with year_col:
            to_year = st.number_input(
                "Ending Year",
                min_value=1900,
                max_value=2100,
                value=2024,
                step=1,
                key=f"to_year_{i}",
            )

    start_date = datetime(
        int(from_year),
        int(from_month),
        1,
    )

    # Calculate an exclusive end date for the contribution function.
    if int(to_month) == 12:
        exclusive_end_date = datetime(
            int(to_year) + 1,
            1,
            1,
        )
    else:
        exclusive_end_date = datetime(
            int(to_year),
            int(to_month) + 1,
            1,
        )

    manual_periods.append(
        (start_date, exclusive_end_date)
    )

    if exclusive_end_date <= start_date:
        st.warning(
            f"Period {i + 1}: the ending month must not be "
            "earlier than the starting month."
        )

    st.divider()


# ============================================================
# SECTION 3: CALCULATE CONTRIBUTIONS
# ============================================================

if st.button(
    "🧮 Calculate Default Contribution",
    type="primary",
    width="stretch",
    key="calculate_manual_default",
):

    try:
        for start_date, exclusive_end_date in manual_periods:
            if exclusive_end_date <= start_date:
                raise ValueError(
                    "The ending month cannot be earlier than "
                    "the starting month."
                )

        result = calculate_default_contribution(
            periods=manual_periods,
            ips=int(number_of_ips),
        )

        st.session_state.manual_default_result = result
        st.session_state.manual_default_periods = list(
            manual_periods
        )
        st.session_state.manual_default_ips = int(number_of_ips)

        # Remove previously generated files when recalculating.
        st.session_state.manual_notice_zip = None
        st.session_state.manual_notice_zip_name = None

        st.success(
            "Default contribution calculated successfully."
        )

    except ValueError as error:
        st.error(str(error))

    except Exception as error:
        st.error(
            "An error occurred while calculating the contribution."
        )
        st.exception(error)


# ============================================================
# SECTION 4: DISPLAY RESULTS
# ============================================================

result = st.session_state.manual_default_result

if result is None:
    st.info(
        "Calculate the default contribution above to continue "
        "to the notice and assessment-sheet generator."
    )
    st.stop()


st.divider()
st.header("Calculation Results")

st.subheader("Detailed Calculation")

st.dataframe(
    result["rows"],
    width="stretch",
    hide_index=True,
)


# ============================================================
# CALCULATION METRICS
# ============================================================

st.subheader("Calculation Summary")

col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    st.metric(
        "Total Months",
        f"{result['total_months']:,}",
    )

with col2:
    st.metric(
        "Principal Contribution",
        f"{result['total_assessed']:,.0f}",
    )

with col3:
    st.metric(
        "Employer Contribution (5/6)",
        f"{(result['total_assessed'] / 6) * 5:,.0f}",
    )

with col4:
    st.metric(
        "Employee Contribution (1/6)",
        f"{result['total_assessed'] / 6:,.0f}",
    )

with col5:
    st.metric(
        "50% Statutory Increase",
        f"{result['statutory_amount']:,.0f}",
    )

with col6:
    st.metric(
        "Total Payable",
        f"{result['final_total']:,.0f}",
    )

st.subheader("Calculation Method")

st.markdown(
    """
    - **Principal Contribution:** Calculated using the applicable
      financial-year contribution rates and number of insured persons.
    - **Statutory Increase:** 50% of the principal contribution.
    - **Total Payable:** Principal contribution plus statutory increase.
    """
)


# ============================================================
# SECTION 5: NOTICE AND ESTABLISHMENT DETAILS
# ============================================================

st.divider()
st.header("2. Generate Notice and Assessment Sheet")

st.write(
    "Enter the establishment details and upload your Urdu "
    "payment notice template. The app will generate both the "
    "individual notice and the matching assessment sheet. "
    "No Excel process sheet is required."
)

template_file = st.file_uploader(
    "Upload Urdu Payment Notice Template (.docx)",
    type=["docx"],
    accept_multiple_files=False,
    key="manual_individual_notice_template",
)

if template_file is not None:
    st.success(
        f"Word template uploaded: {template_file.name}"
    )


saved_periods = st.session_state.manual_default_periods
saved_ips = st.session_state.manual_default_ips

notice_period_text = format_notice_periods(saved_periods)

st.text_input(
    "Calculated Default Period(s)",
    value=notice_period_text,
    disabled=True,
    key="manual_notice_period_display",
)

st.number_input(
    "Calculated Number of Insured Persons",
    min_value=1,
    value=int(saved_ips),
    disabled=True,
    key="manual_notice_ips_display",
)


# ============================================================
# ESTABLISHMENT DETAILS FORM
# ============================================================

with st.form("manual_notice_form"):

    st.subheader("Establishment Information")

    col1, col2 = st.columns(2)

    with col1:
        registration_code = st.text_input(
            "Registration Code *",
            key="notice_registration_code",
        )

        sub_code = st.text_input(
            "Sub-Code",
            value="0",
            key="notice_sub_code",
        )

        establishment_name = st.text_input(
            "Name of the Establishment *",
            key="notice_establishment_name",
        )

        city = st.text_input(
            "City *",
            key="notice_city",
        )

    with col2:
        address = st.text_area(
            "Address *",
            height=120,
            key="notice_address",
        )

        contact_person = st.text_input(
            "Name of Employer / Contact Person",
            key="notice_contact_person",
        )

        contact_number = st.text_input(
            "Contact Number",
            key="notice_contact_number",
        )

    generate_button = st.form_submit_button(
        "📄 Generate Payment Notice + Assessment Sheet",
        type="primary",
        width="stretch",
    )


# ============================================================
# GENERATE BOTH OUTPUTS
# ============================================================

if generate_button:

    if template_file is None:
        st.error(
            "Please upload the Urdu Payment Notice Word template."
        )

    elif not registration_code.strip():
        st.error("Please enter the Registration Code.")

    elif not establishment_name.strip():
        st.error("Please enter the Name of the Establishment.")

    elif not address.strip():
        st.error("Please enter the Address.")

    elif not city.strip():
        st.error("Please enter the City.")

    else:

        try:
            with st.spinner(
                "Generating the Urdu notice and assessment sheet..."
            ):

                # ------------------------------------------------
                # 1. PREPARE NOTICE PLACEHOLDER MAPPING
                # ------------------------------------------------

                notice_data = {
                    "Code": registration_code.strip(),
                    "Sub_Code": sub_code.strip() or "0",
                    "Name": contact_person.strip(),
                    "Name_of_the_Establishment": (
                        establishment_name.strip()
                    ),
                    "Address": address.strip(),
                    "City": city.strip(),
                    "contact": contact_number.strip(),
                    "Period": notice_period_text,
                    "Ref. Letter No.": "",
                }

                # ------------------------------------------------
                # 2. GENERATE INDIVIDUAL URDU NOTICE
                # ------------------------------------------------

                notice_bytes = create_notice_document(
                    template_file.getvalue(),
                    notice_data,
                )

                # ------------------------------------------------
                # 3. BUILD ASSESSMENT SHEET DATA
                # ------------------------------------------------

                assessment_row = build_assessment_row(
                    establishment_name=(
                        establishment_name.strip()
                    ),
                    registration_code=registration_code.strip(),
                    sub_code=sub_code.strip() or "0",
                    ips=int(saved_ips),
                    periods=saved_periods,
                )

                sheet_name = safe_filename(
                    establishment_name.strip()
                )[:31]

                # ------------------------------------------------
                # 4. GENERATE ASSESSMENT WORKBOOK
                # ------------------------------------------------

                assessment_bytes = create_assessment_workbook(
                    assessment_row,
                    sheet_name,
                )

                # ------------------------------------------------
                # 5. PACKAGE BOTH FILES IN ONE ZIP
                # ------------------------------------------------

                code_for_filename = safe_filename(
                    registration_code.strip()
                )

                notice_filename = (
                    f"Payment_Notice_{code_for_filename}.docx"
                )

                assessment_filename = (
                    f"Defaulters_Assessment_{code_for_filename}.xlsx"
                )

                zip_bytes = create_output_zip(
                    notice_bytes=notice_bytes,
                    assessment_bytes=assessment_bytes,
                    notice_filename=notice_filename,
                    assessment_filename=assessment_filename,
                )

                st.session_state.manual_notice_zip = zip_bytes
                st.session_state.manual_notice_zip_name = (
                    f"EOBI_Notice_and_Assessment_"
                    f"{code_for_filename}.zip"
                )

            st.success(
                "Both the individual payment notice and the "
                "Defaulters Assessment Sheet were generated successfully."
            )

        except Exception as error:
            st.error(
                "The notice or assessment sheet could not be generated. "
                "Please check the template and the existing functions "
                "in generator.py."
            )
            st.exception(error)


# ============================================================
# DOWNLOAD BOTH FILES
# ============================================================

if st.session_state.manual_notice_zip is not None:

    st.divider()
    st.header("3. Download Generated Files")

    st.write(
        "The ZIP file contains the individual Urdu payment notice "
        "and its corresponding Defaulters Assessment Sheet."
    )

    st.download_button(
        label="⬇️ Download Notice + Assessment Sheet (.zip)",
        data=st.session_state.manual_notice_zip,
        file_name=st.session_state.manual_notice_zip_name,
        mime="application/zip",
        type="primary",
        width="stretch",
        key="download_notice_assessment_zip",
    )

    st.markdown(
        """
        **Files included:**
        - Individual Urdu Payment Notice (`.docx`)
        - Defaulters Assessment Sheet (`.xlsx`)
        """
    )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "EOBI Manual Default Calculation, Payment Notice "
    "& Defaulters Assessment Sheet Generator"
)
