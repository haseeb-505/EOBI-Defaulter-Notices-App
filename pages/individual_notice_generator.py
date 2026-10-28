import os
import re
import tempfile
from datetime import timedelta
from io import BytesIO

import streamlit as st
from generator import fill_notice


st.set_page_config(
    page_title="Individual Defaulter Notice",
    page_icon="📄",
    layout="wide",
)


def build_notice_period_string(periods):
    """Format saved calculator periods; their end dates are exclusive."""
    formatted_periods = []

    for start_date, exclusive_end_date in periods:
        if start_date is None or exclusive_end_date is None:
            continue

        end_display = (
            (exclusive_end_date.replace(day=1) - timedelta(days=1)).replace(day=1)
            if exclusive_end_date.day == 1
            else exclusive_end_date
        )
        formatted_periods.append(
            f"{start_date.strftime('%Y-%m')} تا {end_display.strftime('%Y-%m')}"
        )

    return "، ".join(formatted_periods) if formatted_periods else "________"


def clean_text(value):
    return str(value).strip() if value is not None else ""


def safe_filename(value):
    value = re.sub(r"[^\w\s-]", "", clean_text(value), flags=re.UNICODE)
    value = re.sub(r"\s+", "_", value).strip("_")
    return value[:60] or "Individual_Defaulter_Notice"


st.title("📄 Individual Defaulter Notice Generator")
st.write(
    "Upload the Urdu Payment Notice template, enter one employer's "
    "information, and generate one individual notice."
)

if st.button("⬅️ Back to Notice Generator"):
    st.switch_page("app.py")

st.divider()

saved_periods = st.session_state.get("manual_default_periods")
saved_ips = st.session_state.get("manual_default_ips")
saved_result = st.session_state.get("manual_default_result")

if not saved_periods or not saved_ips:
    st.warning(
        "First open Manual Default Calculation, enter the default periods "
        "and number of IPs, and click Calculate Default Contribution. "
        "Then return here. The saved periods and IP count will be reused."
    )
    if st.button("🧮 Open Manual Default Calculation", type="primary"):
        st.switch_page("pages/manual_default_calculator.py")
    st.stop()

st.subheader("Default Information Reused from the Calculator")
period_text = build_notice_period_string(saved_periods)
st.write(f"**Default period(s):** {period_text}")
st.write(f"**Number of insured persons (IPs):** {saved_ips}")

if saved_result:
    col1, col2, col3 = st.columns(3)
    col1.metric("Principal Contribution", f"{saved_result['total_assessed']:,}")
    col2.metric("50% Statutory Increase", f"{saved_result['statutory_amount']:,}")
    col3.metric("Total Payable", f"{saved_result['final_total']:,}")

st.divider()
st.subheader("1. Upload Urdu Payment Notice")

template_file = st.file_uploader(
    "Upload ONE Urdu Payment Notice template (.docx)",
    type=["docx"],
    accept_multiple_files=False,
    key="individual_notice_template",
    help="Use the same Urdu Payment Notice template used by the main generator.",
)

st.subheader("2. Employer Information")

with st.form("individual_employer_form"):
    col1, col2 = st.columns(2)

    with col1:
        code = st.text_input("Code *", key="individual_code")
        sub_code = st.text_input(
            "Sub_Code",
            value="0",
            key="individual_sub_code",
            help="Enter 0 if the establishment has no sub-code.",
        )
        establishment_name = st.text_input(
            "Name of the Establishment *",
            key="individual_establishment_name",
        )
        address = st.text_area("Address *", key="individual_address")

    with col2:
        city = st.text_input("City *", key="individual_city")
        contact_name = st.text_input("Name", key="individual_contact_name")
        contact = st.text_input(
            "Contact",
            key="individual_contact",
            help="Enter the employer/contact number as it should appear on the notice.",
        )

    submitted = st.form_submit_button(
        "📄 Generate Individual Notice",
        type="primary",
        use_container_width=True,
    )

if submitted:
    missing = []
    if not template_file:
        missing.append("Urdu Payment Notice template")
    if not clean_text(code):
        missing.append("Code")
    if not clean_text(establishment_name):
        missing.append("Name of the Establishment")
    if not clean_text(address):
        missing.append("Address")
    if not clean_text(city):
        missing.append("City")

    if missing:
        st.error("Please complete the following required fields: " + ", ".join(missing))
        st.stop()

    clean_sub_code = clean_text(sub_code) or "0"
    notice_data = {
        "Code": clean_text(code),
        "Sub_Code": clean_sub_code,
        "Name_of_the_Establishment": clean_text(establishment_name),
        "Address": clean_text(address),
        "City": clean_text(city),
        "Name": clean_text(contact_name),
        "contact": clean_text(contact),
        "Period": period_text,
        "Ref. Letter No.": "",
    }

    try:
        with st.spinner("Generating the individual Urdu Payment Notice..."):
            template_file.seek(0)
            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as temp_template:
                temp_template.write(template_file.getvalue())
                temp_template_path = temp_template.name

            try:
                document = fill_notice(temp_template_path, notice_data)
            finally:
                try:
                    os.remove(temp_template_path)
                except OSError:
                    pass

            output = BytesIO()
            document.save(output)
            output_bytes = output.getvalue()

        filename = safe_filename(establishment_name) + "_Defaulter_Notice.docx"
        st.success("Individual notice generated successfully.")
        st.download_button(
            label="⬇️ Download Individual Notice (.docx)",
            data=output_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )

        with st.expander("Review the information used"):
            st.write(f"**Code:** {clean_text(code)}")
            st.write(f"**Sub_Code:** {clean_sub_code}")
            st.write(f"**Establishment:** {clean_text(establishment_name)}")
            st.write(f"**Address:** {clean_text(address)}")
            st.write(f"**City:** {clean_text(city)}")
            st.write(f"**Name:** {clean_text(contact_name) or '—'}")
            st.write(f"**Contact:** {clean_text(contact) or '—'}")
            st.write(f"**Default period(s):** {period_text}")
            st.write(f"**IPs:** {saved_ips}")

    except Exception as exc:
        st.error("The notice could not be generated.")
        st.exception(exc)
