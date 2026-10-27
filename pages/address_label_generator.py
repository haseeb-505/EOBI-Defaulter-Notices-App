
import html
import pandas as pd
import streamlit as st


# ---------------------------------------------------------
# PAGE CONFIGURATION
# ---------------------------------------------------------
st.set_page_config(
    page_title="EOBI Address Label Generator",
    page_icon="🏷️",
    layout="wide",
)

st.title("🏷️ EOBI Address Label Generator")
st.write(
    "Upload an Excel file to generate printable establishment address labels. "
    "Owner/Contact and Phone are optional."
)


# ---------------------------------------------------------
# HELPER FUNCTIONS
# ---------------------------------------------------------
def normalize_column(column):
    """Normalize Excel column headings for flexible matching."""
    text = str(column).strip().lower()
    for character in ("_", "-", ".", ":", "#", "/", "\\", "(", ")"):
        text = text.replace(character, " ")
    return " ".join(text.split())


def is_missing(value):
    """Return True for blank, NaN, or None values."""
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return not str(value).strip()


def clean_value(value):
    """Convert a cell value to printable text."""
    if is_missing(value):
        return ""
    return str(value).strip()


def escape_html(value):
    """Escape text before inserting it into HTML."""
    return html.escape(clean_value(value))


def find_column(df, aliases):
    """Find a column using its possible heading names."""
    normalized_columns = {
        normalize_column(column): column for column in df.columns
    }
    for alias in aliases:
        normalized_alias = normalize_column(alias)
        if normalized_alias in normalized_columns:
            return normalized_columns[normalized_alias]
    return None


# ---------------------------------------------------------
# EXCEL COLUMN MAPPING
# ---------------------------------------------------------
COLUMN_ALIASES = {
    "establishment": [
        "Name_of_Establishment",
        "Name of Establishment",
        "Establishment Name",
        "Name",
    ],
    "address": ["Address", "Establishment Address"],
    "city": ["City", "District"],
    "owner": [
        "Owner/Contact",
        "Owner Contact",
        "Owner",
        "Contact Person",
        "Name of Owner",
    ],
    "phone": [
        "Phone",
        "Phone Number",
        "Contact",
        "Contact Number",
        "Telephone",
        "Mobile",
    ],
    "courier": ["Courier #", "Courier Number", "Courier No", "Courier"],
}

# Reference priority: the first non-empty reference wins.
REFERENCE_PAIRS = [
    ("SC Ref2", "SC date2"),
    ("SC Ref", "SC date"),
    ("Ref3", "Date3"),
    ("Ref2", "Date2"),
    ("REF1", "Date1"),
]


def map_columns(df):
    """Map uploaded Excel headings to expected fields."""
    mapped = {
        field: find_column(df, aliases)
        for field, aliases in COLUMN_ALIASES.items()
    }
    for ref_name, date_name in REFERENCE_PAIRS:
        mapped[ref_name] = find_column(df, [ref_name])
        mapped[date_name] = find_column(df, [date_name])
    return mapped


# ---------------------------------------------------------
# REFERENCE SELECTION
# ---------------------------------------------------------
def get_reference_details(row, mapped):
    """Return the highest-priority available reference and its date."""
    for ref_name, date_name in REFERENCE_PAIRS:
        ref_column = mapped.get(ref_name)
        if ref_column is None:
            continue

        reference = clean_value(row.get(ref_column))
        if not reference:
            continue

        date_column = mapped.get(date_name)
        reference_date = (
            clean_value(row.get(date_column))
            if date_column is not None
            else ""
        )

        if date_column is not None:
            raw_date = row.get(date_column)
            if not is_missing(raw_date):
                try:
                    parsed_date = pd.to_datetime(raw_date, errors="coerce")
                    if not pd.isna(parsed_date):
                        reference_date = parsed_date.strftime("%d-%m-%Y")
                except (TypeError, ValueError):
                    pass

        return reference, reference_date

    return "", ""


# ---------------------------------------------------------
# ADDRESS LABEL HTML GENERATOR
# ---------------------------------------------------------
def create_address_labels_html(df, mapped):
    """Create printable HTML with 16 labels per A4 page."""
    labels = []

    for _, row in df.iterrows():
        establishment = clean_value(row.get(mapped["establishment"]))
        address = clean_value(row.get(mapped["address"]))
        city = clean_value(row.get(mapped["city"]))

        # Owner/contact and phone are optional.
        owner_column = mapped.get("owner")
        phone_column = mapped.get("phone")

        owner = (
            clean_value(row.get(owner_column))
            if owner_column is not None
            else ""
        )
        phone = (
            clean_value(row.get(phone_column))
            if phone_column is not None
            else ""
        )

        # Courier number is optional.
        courier_column = mapped.get("courier")
        courier = (
            clean_value(row.get(courier_column))
            if courier_column is not None
            else ""
        )

        reference, reference_date = get_reference_details(row, mapped)

        reference_line = ""
        if reference:
            reference_text = f"Ref: {escape_html(reference)}"
            if reference_date:
                reference_text += (
                    f" &nbsp; | &nbsp; Date: {escape_html(reference_date)}"
                )
            reference_line = (
                f'<div class="reference">{reference_text}</div>'
            )

        owner_line = (
            f'<div class="owner">{escape_html(owner)}</div>'
            if owner
            else ""
        )

        # Preserve phone content, including multiple numbers separated
        # by "|" or "/".
        phone_line = (
            f'<div class="phone">[{escape_html(phone)}]</div>'
            if phone
            else ""
        )

        courier_line = (
            f'<div class="courier">Courier #: {escape_html(courier)}</div>'
            if courier
            else ""
        )

        labels.append(
            f"""
            <div class="label">
                <div class="label-content">
                    {reference_line}
                    <div class="establishment">
                        {escape_html(establishment)}
                    </div>
                    <div class="address">{escape_html(address)}</div>
                    <div class="city">{escape_html(city)}</div>
                    {owner_line}
                    {phone_line}
                    {courier_line}
                </div>
            </div>
            """
        )

    # A4 portrait: 2 columns × 8 rows = 16 labels per page.
    labels_per_page = 16
    pages = []

    for start in range(0, len(labels), labels_per_page):
        page_labels = labels[start:start + labels_per_page]
        pages.append(
            f'<div class="page">{"".join(page_labels)}</div>'
        )

    document = f"""<!DOCTYPE html>
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
        padding: 0;
        font-family: Arial, Helvetica, sans-serif;
        color: #000;
    }}

    .page {{
        width: 194mm;
        height: 281mm;
        margin: 0 auto;
        display: grid;
        grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
        grid-template-rows: repeat(8, 34.8mm);
        column-gap: 4mm;
        row-gap: 0.5mm;
        align-content: start;
        page-break-after: always;
        break-after: page;
    }}

    .page:last-child {{
        page-break-after: auto;
        break-after: auto;
    }}

    .label {{
        height: 34.8mm;
        min-width: 0;
        min-height: 0;
        overflow: hidden;
        border: 1px dashed #999;
        padding: 2.5mm 3mm;
        display: flex;
        align-items: center;
        justify-content: center;
        overflow-wrap: anywhere;
        word-break: normal;
        line-height: 1.18;
    }}

    .label-content {{
        width: 100%;
        min-width: 0;
        max-height: 100%;
        font-size: 10pt;
        line-height: 1.18;
        overflow-wrap: anywhere;
        overflow: hidden;
    }}

    .label-content > div {{
        max-width: 100%;
        overflow-wrap: anywhere;
        word-wrap: break-word;
    }}

    .reference {{
        font-size: 0.82em;
        margin-bottom: 0.7mm;
        line-height: 1.1;
    }}

    .establishment {{
        font-size: 1.12em;
        font-weight: bold;
        margin-bottom: 0.7mm;
        line-height: 1.12;
    }}

    .address,
    .city,
    .owner,
    .phone {{
        font-size: 0.96em;
        line-height: 1.15;
    }}

    .city {{
        font-weight: bold;
        margin-bottom: 0.5mm;
    }}

    .courier {{
        font-size: 0.82em;
        margin-top: 0.5mm;
        line-height: 1.1;
    }}

    @media screen {{
        body {{
            background: #eee;
            padding: 10px;
        }}

        .page {{
            background: #fff;
            margin: 0 auto 15px auto;
            box-shadow: 0 0 4px #aaa;
        }}
    }}

    @media print {{
        .label {{
            border: 1px dashed #bbb;
        }}
    }}
</style>
</head>

<body>
    {''.join(pages)}

<script>
    // Reduce the font size on labels whose content does not fit.
    // Runs on screen and again immediately before printing.
    function fitAddressLabels() {{
        document.querySelectorAll('.label').forEach(function(label) {{
            const content = label.querySelector('.label-content');
            if (!content) return;

            let fontSize = 10;
            content.style.fontSize = fontSize + 'pt';
            content.style.lineHeight = '1.18';

            // Reduce font size gradually until the content fits.
            while (
                (content.scrollHeight > content.clientHeight + 1 ||
                 content.scrollWidth > content.clientWidth + 1) &&
                fontSize > 5
            ) {{
                fontSize -= 0.25;
                content.style.fontSize = fontSize + 'pt';
            }}

            // Tighten line spacing if the content is still too tall.
            if (content.scrollHeight > content.clientHeight + 1) {{
                content.style.lineHeight = '1.0';

                while (
                    content.scrollHeight > content.clientHeight + 1 &&
                    fontSize > 4
                ) {{
                    fontSize -= 0.25;
                    content.style.fontSize = fontSize + 'pt';
                }}
            }}
        }});
    }}

    window.addEventListener('load', fitAddressLabels);
    window.addEventListener('beforeprint', fitAddressLabels);
</script>
</body>
</html>
"""
    return document


# ---------------------------------------------------------
# FILE UPLOAD
# ---------------------------------------------------------
uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx", "xls"],
    help="Upload an Excel workbook containing establishment addresses.",
)

if uploaded_file is not None:
    try:
        if uploaded_file.name.lower().endswith(".xlsx"):
            df = pd.read_excel(
                uploaded_file,
                dtype=object,
                engine="openpyxl",
            )
        else:
            df = pd.read_excel(uploaded_file, dtype=object)

        # Remove completely empty rows.
        df = df.dropna(how="all").reset_index(drop=True)

        if df.empty:
            st.error("The uploaded Excel file contains no records.")
            st.stop()

        mapped = map_columns(df)

        # Only establishment name, address and city are mandatory.
        mandatory_fields = {
            "establishment": "Name of Establishment",
            "address": "Address",
            "city": "City",
        }

        missing_columns = [
            display_name
            for field, display_name in mandatory_fields.items()
            if mapped.get(field) is None
        ]

        if missing_columns:
            st.error(
                "The following mandatory columns were not found: "
                + ", ".join(missing_columns)
            )
            st.write("Columns found in your Excel file:")
            st.write([str(column) for column in df.columns])
            st.stop()

        # Validate only the three mandatory fields.
        errors = []

        for index, row in df.iterrows():
            missing_values = []

            for field, display_name in mandatory_fields.items():
                if is_missing(row.get(mapped[field])):
                    missing_values.append(display_name)

            if missing_values:
                excel_row = index + 2
                errors.append(
                    f"Excel row {excel_row}: missing "
                    + ", ".join(missing_values)
                )

        if errors:
            st.error(
                "Some records have missing mandatory values. "
                "Please correct these rows in Excel and upload again:"
            )
            for error in errors:
                st.write(f"- {error}")
            st.stop()

        st.success(
            f"Excel file loaded successfully: {len(df)} records."
        )
        st.caption(
            "Layout: 2 columns × 8 rows (16 labels per A4 page). "
            "Owner/Contact, Phone, Courier #, references and dates are optional."
        )

        with st.expander("Preview uploaded records"):
            st.dataframe(df.head(10), use_container_width=True)

        if st.button(
            "🏷️ Generate Address Labels",
            type="primary",
            use_container_width=True,
        ):
            html_document = create_address_labels_html(df, mapped)

            st.download_button(
                label="⬇️ Download Address Labels (HTML)",
                data=html_document.encode("utf-8"),
                file_name="EOBI_Address_Labels.html",
                mime="text/html",
                use_container_width=True,
            )

            st.info(
                "Download the HTML file and open it in your browser. "
                "Press Ctrl+P to print. Select A4 paper, Portrait "
                "orientation, and 100% scale. Disable browser headers "
                "and footers in print settings if needed."
            )

    except Exception as exc:
        st.error(f"Unable to process the Excel file: {exc}")
