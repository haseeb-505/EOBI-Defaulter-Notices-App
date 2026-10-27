
import os
import io
import html
import shutil
import tempfile
import zipfile

import pandas as pd
import streamlit as st

from generator import generate_all


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="EOBI Defaulter Notice Generator",
    page_icon="📄",
    layout="wide",
)


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>
    div.stButton > button[kind="primary"] {
        background-color: #1976D2;
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        font-size: 17px;
        padding: 0.65rem 1rem;
    }

    div.stButton > button[kind="primary"]:hover {
        background-color: #1565C0;
        color: white;
    }

    div.stDownloadButton > button {
        background-color: #2E7D32;
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        font-size: 17px;
        padding: 0.65rem 1rem;
    }

    div.stDownloadButton > button:hover {
        background-color: #1B5E20;
        color: white;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# ADDRESS LABEL GENERATOR
# ============================================================

def create_address_labels_html(df):
    """
    Create printable A4 address labels from the uploaded
    establishment Excel data.

    Layout:
    - A4 portrait
    - 2 columns
    - 11 rows per page
    - 22 labels per page
    """

    def normalize_column(column):
        return (
            str(column)
            .strip()
            .lower()
            .replace("_", " ")
            .replace("-", " ")
            .replace(".", "")
            .replace(":", "")
        )

    # Map different possible Excel column headings to
    # the required address-label fields.

    column_aliases = {
        "establishment": [
            "name of the establishment",
            "name of establishment",
            "establishment name",
            "establishment",
            "establishment name / title",
        ],
        "address": [
            "address",
            "postal address",
            "establishment address",
        ],
        "city": [
            "city",
            "town",
        ],
        "name": [
            "name",
            "contact name",
            "contact person",
            "contact person name",
        ],
        "contact": [
            "contact",
            "contact number",
            "contact no",
            "phone",
            "telephone",
            "mobile",
        ],
    }

    actual_columns = {
        normalize_column(column): column
        for column in df.columns
    }

    selected_columns = {}

    for field, aliases in column_aliases.items():
        for alias in aliases:
            normalized_alias = normalize_column(alias)

            if normalized_alias in actual_columns:
                selected_columns[field] = actual_columns[
                    normalized_alias
                ]
                break

    # Establishment name, address, and city are essential.
    required_fields = ["establishment", "address", "city"]

    missing_fields = [
        field
        for field in required_fields
        if field not in selected_columns
    ]

    if missing_fields:
        raise ValueError(
            "Cannot create address labels because these "
            "Excel columns could not be identified: "
            + ", ".join(missing_fields)
            + ". Please check the column headings."
        )

    def clean_value(value):
        if pd.isna(value):
            return ""

        # Avoid displaying numbers such as 12345.0.
        if isinstance(value, float) and value.is_integer():
            value = int(value)

        return str(value).strip()

    def get_value(row, field):
        column = selected_columns.get(field)

        if column is None:
            return ""

        return clean_value(row[column])

    def safe_text(value):
        return html.escape(value, quote=True)

    labels = []

    for _, row in df.iterrows():

        establishment = get_value(row, "establishment")
        address = get_value(row, "address")
        city = get_value(row, "city")
        contact_name = get_value(row, "name")
        contact_number = get_value(row, "contact")

        # Combine address and city without unnecessary commas.
        address_parts = [
            part for part in [address, city] if part
        ]
        full_address = ", ".join(address_parts)

        # Display contact name followed by telephone number.
        contact_parts = []

        if contact_name:
            contact_parts.append(safe_text(contact_name))

        if contact_number:
            contact_parts.append(
                f"[{safe_text(contact_number)}]"
            )

        contact_line = " ".join(contact_parts)

        labels.append(
            f"""
            <div class="label">
                <div class="est">
                    {safe_text(establishment)}
                </div>
                <div>{safe_text(full_address)}</div>
                <div>{contact_line}</div>
            </div>
            """
        )

    # Build the HTML document.
    # Blank cells are added to complete the final page.

    pages = []

    for start in range(0, len(labels), 22):

        page_labels = labels[start:start + 22]

        while len(page_labels) < 22:
            page_labels.append(
                '<div class="label"></div>'
            )

        pages.append(
            f"""
            <section class="page">
                <div class="grid">
                    {''.join(page_labels)}
                </div>
            </section>
            """
        )

    # Also handle an empty Excel worksheet.
    if not pages:
        pages.append(
            """
            <section class="page">
                <div class="grid">
                </div>
            </section>
            """
        )

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>EOBI Address Labels</title>

<style>
    @page {{
        size: A4 portrait;
        margin: 8mm;
    }}

    * {{
        box-sizing: border-box;
    }}

    body {{
        margin: 0;
        font-family: Arial, sans-serif;
    }}

    .page {{
        width: 194mm;
        height: 281mm;
        page-break-after: always;
        break-after: page;
        overflow: hidden;
    }}

    .page:last-child {{
        page-break-after: auto;
        break-after: auto;
    }}

    .grid {{
        display: grid;
        grid-template-columns: 1fr 1fr;
        grid-template-rows: repeat(11, 25.545mm);
    }}

    .label {{
        border: 0.35pt solid #999;
        padding: 2mm 2.5mm;
        display: flex;
        flex-direction: column;
        justify-content: center;
        gap: 1.5mm;
        overflow: hidden;
        overflow-wrap: anywhere;
        font-size: 9pt;
        line-height: 1.25;
    }}

    .est {{
        font-weight: bold;
    }}

    @media screen {{
        body {{
            background: #ddd;
        }}

        .page {{
            background: white;
            margin: 10mm auto;
            box-shadow: 0 1px 5px #aaa;
        }}
    }}

    @media print {{
        .page {{
            margin: 0;
            box-shadow: none;
        }}
    }}
</style>
</head>

<body>
{''.join(pages)}
</body>
</html>
"""



# ============================================================
# HEADER AND NAVIGATION BUTTONS
# ============================================================

col_left, col_manual, col_labels = st.columns([4, 1.4, 1.4])

with col_left:
    st.markdown(
        """
        <div style="font-size: 27px; font-weight: 700;">
            📄 EOBI Defaulter Notice & Assessment Generator
        </div>
        """,
        unsafe_allow_html=True,
    )

with col_manual:
    if st.button(
        "🧮 Manual Default Calculation",
        type="secondary",
        width="stretch",
    ):
        st.switch_page("pages/manual_default_calculator.py")

with col_labels:
    if st.button(
        "🏷️ Address Labels",
        type="secondary",
        width="stretch",
    ):
        st.switch_page("pages/address_label_generator.py")

st.write(
    "Upload your establishment Excel file and Urdu payment "
    "notice template to generate individual notices, a combined "
    "notice document, Defaulters Assessment Sheets, and "
    "printable address labels."
)

st.divider()


# ============================================================
# FILE UPLOADERS
# ============================================================

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Upload Defaulter Data")

    data_file = st.file_uploader(
        "Upload ONE Excel file only",
        type=["xlsx", "xls"],
        accept_multiple_files=False,
        help=(
            "Upload one Excel file containing the establishment "
            "names, addresses, cities, contact names, and numbers."
        ),
    )

    if data_file is not None:
        st.success(f"✅ Excel file uploaded: {data_file.name}")

with col2:
    st.subheader("2. Upload Urdu Payment Notice")

    urdu_notice_file = st.file_uploader(
        "Upload ONE Urdu Payment Notice",
        type=["docx"],
        accept_multiple_files=False,
        help="Upload the payment notice template in .docx format.",
    )

    if urdu_notice_file is not None:
        st.success(
            f"✅ Urdu Payment Notice uploaded: "
            f"{urdu_notice_file.name}"
        )


# ============================================================
# DATA PREVIEW
# ============================================================

if data_file is not None:

    st.subheader("Data Preview")

    try:
        preview_df = pd.read_excel(
            io.BytesIO(data_file.getvalue())
        )

        st.write(f"**Records found:** {len(preview_df)}")

        st.dataframe(
            preview_df.head(10),
            width="stretch",
        )

    except Exception as e:
        st.error(f"Could not read the Excel file: {e}")


# ============================================================
# GENERATE NOTICES, ASSESSMENT SHEETS AND LABELS
# ============================================================

st.divider()

generate_button = st.button(
    "Generate Notices, Assessment Sheets & Address Labels",
    type="primary",
    width="stretch",
)

if generate_button:

    if data_file is None:
        st.error("Please upload the Excel data file first.")
        st.stop()

    if urdu_notice_file is None:
        st.error("Please upload the Urdu Word notice template first.")
        st.stop()

    working_dir = tempfile.mkdtemp(prefix="eobi_generator_")

    try:
        # ----------------------------------------------------
        # SAVE INPUT FILES
        # ----------------------------------------------------

        data_path = os.path.join(
            working_dir,
            os.path.basename(data_file.name),
        )

        with open(data_path, "wb") as file:
            file.write(data_file.getvalue())

        template_path = os.path.join(
            working_dir,
            os.path.basename(urdu_notice_file.name),
        )

        with open(template_path, "wb") as file:
            file.write(urdu_notice_file.getvalue())

        output_dir = os.path.join(working_dir, "output")
        os.makedirs(output_dir, exist_ok=True)

        # ----------------------------------------------------
        # GENERATE ADDRESS LABELS FROM THE SAME EXCEL FILE
        # ----------------------------------------------------

        label_df = pd.read_excel(data_path)

        labels_html = create_address_labels_html(label_df)

        labels_path = os.path.join(
            output_dir,
            "Address_Labels.html",
        )

        with open(
            labels_path,
            "w",
            encoding="utf-8",
        ) as file:
            file.write(labels_html)

        # ----------------------------------------------------
        # GENERATE EXISTING DOCUMENTS
        # ----------------------------------------------------

        with st.spinner(
            "Generating notices, assessment sheets, "
            "and address labels. Please wait..."
        ):

            result = generate_all(
                data_file=data_path,
                notice_template=template_path,
                output_dir=output_dir,
            )

            # ------------------------------------------------
            # CREATE ZIP PACKAGE
            # ------------------------------------------------

            zip_path = os.path.join(
                working_dir,
                "EOBI_Defaulter_Output.zip",
            )

            with zipfile.ZipFile(
                zip_path,
                "w",
                zipfile.ZIP_DEFLATED,
            ) as zip_file:

                # Individual notices
                for file_path in result["individual_files"]:
                    zip_file.write(
                        file_path,
                        arcname=os.path.join(
                            "Individual_Notices",
                            os.path.basename(file_path),
                        ),
                    )

                # Combined notice
                if os.path.exists(result["combined_notice"]):
                    zip_file.write(
                        result["combined_notice"],
                        arcname="All_Payment_Notices.docx",
                    )

                # Assessment workbook
                zip_file.write(
                    result["assessment"],
                    arcname="Defaulters_Assessment_Sheets.xlsx",
                )

                # Address labels
                zip_file.write(
                    labels_path,
                    arcname="Address_Labels.html",
                )

        # ----------------------------------------------------
        # READ ZIP INTO MEMORY BEFORE CLEANUP
        # ----------------------------------------------------

        with open(zip_path, "rb") as file:
            zip_data = file.read()

        # ----------------------------------------------------
        # RESULTS
        # ----------------------------------------------------

        st.success("Generation completed successfully!")

        st.subheader("Generation Results")

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric(
                "Establishments",
                result["records"],
            )

        with col2:
            st.metric(
                "Individual Notices",
                len(result["individual_files"]),
            )

        with col3:
            st.metric(
                "Assessment Sheets",
                result["records"],
            )

        with col4:
            st.metric(
                "Address Labels",
                len(label_df),
            )

        st.divider()

        # ----------------------------------------------------
        # DOWNLOAD
        # ----------------------------------------------------

        st.subheader("Download")

        st.write(
            "The ZIP package contains all generated documents "
            "and the printable HTML address labels."
        )

        st.download_button(
            label="📦 Download All Generated Files (ZIP)",
            data=zip_data,
            file_name="EOBI_Defaulter_Output.zip",
            mime="application/zip",
            width="stretch",
        )

        st.info(
            "The ZIP includes individual Word notices, "
            "All_Payment_Notices.docx, "
            "Defaulters_Assessment_Sheets.xlsx, "
            "and Address_Labels.html."
        )

    except Exception as e:
        st.error("Generation failed.")
        st.exception(e)

    finally:
        shutil.rmtree(
            working_dir,
            ignore_errors=True,
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption("EOBI Defaulter Notice & Assessment Generator")
