import os
import shutil
import tempfile
import zipfile

import streamlit as st
import pandas as pd

from generator import generate_all


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="EOBI Defaulter Notice Generator",
    page_icon="📄",
    layout="wide"
)


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>

    /* ======================================================
       GENERATE BUTTON - BLUE
       ====================================================== */

    div.stButton > button[kind="primary"] {
        background-color: #1976D2;
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        font-size: 17px;
        padding: 0.65rem 1rem;
        transition: all 0.2s ease;
    }

    div.stButton > button[kind="primary"]:hover {
        background-color: #1565C0;
        color: white;
        border: none;
    }

    div.stButton > button[kind="primary"]:active {
        background-color: #0D47A1;
        color: white;
    }


    /* ======================================================
       DOWNLOAD BUTTON - GREEN
       ====================================================== */

    div.stDownloadButton > button {
        background-color: #2E7D32;
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 600;
        font-size: 17px;
        padding: 0.65rem 1rem;
        transition: all 0.2s ease;
    }

    div.stDownloadButton > button:hover {
        background-color: #1B5E20;
        color: white;
        border: none;
    }

    div.stDownloadButton > button:active {
        background-color: #145A18;
        color: white;
    }


    /* ======================================================
       INFO / SUCCESS BOXES
       ====================================================== */

    .generation-title {
        font-size: 18px;
        font-weight: 600;
        margin-bottom: 8px;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# HEADER
# ============================================================

st.title(
    "EOBI Defaulter Notice & Assessment Generator"
)

st.write(
    "Upload your defaulter data and Urdu payment notice "
    "template to generate individual notices, a combined "
    "notice document, and Defaulters Assessment Sheets."
)

st.divider()


# ============================================================
# FILE UPLOADERS
# ============================================================

col1, col2 = st.columns(2)


with col1:

    st.subheader(
        "1. Upload Defaulter Data"
    )

    data_file = st.file_uploader(
        "Excel or CSV file",
        type=[
            "xlsx",
            "xls",
            "csv"
        ],
        help=(
            "Upload the Excel/CSV file containing "
            "the establishment records."
        )
    )


with col2:

    st.subheader(
        "2. Upload Notice Template"
    )

    template_file = st.file_uploader(
        "Urdu Payment Notice Word Template",
        type=["docx"],
        help=(
            "Upload the Word notice template containing "
            "placeholders such as «Name», «Period», etc."
        )
    )


st.divider()


# ============================================================
# DATA PREVIEW
# ============================================================

if data_file is not None:

    st.subheader(
        "Data Preview"
    )

    try:

        # Read the uploaded file
        # without changing the uploaded object.

        if data_file.name.lower().endswith(".csv"):

            preview_df = pd.read_csv(
                data_file
            )

        else:

            preview_df = pd.read_excel(
                data_file
            )

        st.write(
            f"**Records found:** {len(preview_df)}"
        )

        st.dataframe(
            preview_df.head(10),
            width="stretch"
        )

    except Exception as e:

        st.error(
            f"Could not read the uploaded data file:\n\n{e}"
        )


# ============================================================
# TEMPLATE INFORMATION
# ============================================================

if template_file is not None:

    st.success(
        f"Notice template selected: "
        f"**{template_file.name}**"
    )


# ============================================================
# GENERATE BUTTON
# ============================================================

st.divider()

generate_button = st.button(
    "🚀 Generate Notices & Assessment Sheets",
    type="primary",
    width="stretch"
)


# ============================================================
# GENERATION
# ============================================================

if generate_button:

    # --------------------------------------------------------
    # Validate uploads
    # --------------------------------------------------------

    if data_file is None:

        st.error(
            "Please upload the Excel/CSV data file first."
        )

        st.stop()


    if template_file is None:

        st.error(
            "Please upload the Word notice template first."
        )

        st.stop()


    # --------------------------------------------------------
    # Temporary working directory
    # --------------------------------------------------------

    working_dir = tempfile.mkdtemp(
        prefix="eobi_generator_"
    )


    try:

        # ====================================================
        # SAVE UPLOADED DATA
        # ====================================================

        data_path = os.path.join(
            working_dir,
            data_file.name
        )

        with open(
            data_path,
            "wb"
        ) as f:

            f.write(
                data_file.getbuffer()
            )


        # ====================================================
        # SAVE WORD TEMPLATE
        # ====================================================

        template_path = os.path.join(
            working_dir,
            template_file.name
        )

        with open(
            template_path,
            "wb"
        ) as f:

            f.write(
                template_file.getbuffer()
            )


        # ====================================================
        # OUTPUT DIRECTORY
        # ====================================================

        output_dir = os.path.join(
            working_dir,
            "output"
        )

        os.makedirs(
            output_dir,
            exist_ok=True
        )


        # ====================================================
        # GENERATION SPINNER
        # ====================================================

        with st.spinner(
            "Generating notices and assessment sheets... "
            "Please wait."
        ):

            # -----------------------------------------------
            # Generate all documents
            # -----------------------------------------------

            result = generate_all(
                data_file=data_path,
                notice_template=template_path,
                output_dir=output_dir
            )


            # -----------------------------------------------
            # Create ZIP
            # -----------------------------------------------

            zip_path = os.path.join(
                working_dir,
                "EOBI_Defaulter_Output.zip"
            )


            with zipfile.ZipFile(
                zip_path,
                "w",
                zipfile.ZIP_DEFLATED
            ) as zip_file:

                # -------------------------------------------
                # Individual notices
                # -------------------------------------------

                for file_path in result[
                    "individual_files"
                ]:

                    zip_file.write(
                        file_path,
                        arcname=os.path.join(
                            "Individual_Notices",
                            os.path.basename(
                                file_path
                            )
                        )
                    )


                # -------------------------------------------
                # Combined notice
                # -------------------------------------------

                if os.path.exists(
                    result["combined_notice"]
                ):

                    zip_file.write(
                        result["combined_notice"],
                        arcname=(
                            "All_Payment_Notices.docx"
                        )
                    )


                # -------------------------------------------
                # Assessment workbook
                # -------------------------------------------

                zip_file.write(
                    result["assessment"],
                    arcname=(
                        "Defaulters_Assessment_Sheets.xlsx"
                    )
                )


        # ====================================================
        # READ ZIP INTO MEMORY
        # ====================================================

        with open(
            zip_path,
            "rb"
        ) as f:

            zip_data = f.read()


        # ====================================================
        # GENERATION COMPLETE
        # ====================================================

        st.success(
            "Generation completed successfully!"
        )


        # ====================================================
        # RESULTS
        # ====================================================

        st.subheader(
            "Generation Results"
        )

        col1, col2, col3 = st.columns(3)


        with col1:

            st.metric(
                "Establishments",
                result["records"]
            )


        with col2:

            st.metric(
                "Individual Notices",
                len(
                    result["individual_files"]
                )
            )


        with col3:

            st.metric(
                "Assessment Sheets",
                result["records"]
            )


        st.divider()


        # ====================================================
        # DOWNLOAD AREA
        # ====================================================

        st.subheader(
            "Download"
        )

        st.write(
            "Your files are ready. Click the green button "
            "below to download the complete ZIP package."
        )


        # ====================================================
        # DOWNLOAD BUTTON - GREEN
        # ====================================================

        st.download_button(
            label=(
                "📦 Download All Generated Files (ZIP)"
            ),
            data=zip_data,
            file_name=(
                "EOBI_Defaulter_Output.zip"
            ),
            mime="application/zip",
            width="stretch"
        )


        st.info(
            "The ZIP contains individual Word notices, "
            "the combined All_Payment_Notices.docx, "
            "and the Defaulters Assessment Sheets.xlsx file."
        )


    except Exception as e:

        st.error(
            "Generation failed."
        )

        st.exception(
            e
        )


    finally:

        # ====================================================
        # CLEANUP
        # ====================================================

        shutil.rmtree(
            working_dir,
            ignore_errors=True
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "EOBI Defaulter Notice & Assessment Generator"
)