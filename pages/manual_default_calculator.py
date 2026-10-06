import streamlit as st
from datetime import datetime

from generator import calculate_default_contribution


# ============================================================
# PAGE CONFIG
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

        div.stButton > button {
            background-color: #1976D2;
            color: white;
            border: none;
            border-radius: 6px;
            font-weight: 600;
        }

        div.stButton > button:hover {
            background-color: #125EA8;
            color: white;
        }

        div[data-testid="stMetric"] {
            border: 1px solid #ddd;
            border-radius: 8px;
            padding: 12px;
        }
        
        /* Red styling for the 5th metric */
        div[data-testid="stHorizontalBlock"] > div:nth-child(5) div[data-testid="stMetric"] {
            border: 1px solid #ff4b4b;
            background-color: #d32f2f;;
        }

        div[data-testid="stHorizontalBlock"] > div:nth-child(5) div[data-testid="stMetricLabel"] {
            color: #d32f2f !important;
        }

        div[data-testid="stHorizontalBlock"] > div:nth-child(5) div[data-testid="stMetricValue"] {
            color: #ffffff !important;
        }
        
         /* Green styling for the 2nd metric */
        div[data-testid="stHorizontalBlock"] > div:nth-child(2) div[data-testid="stMetric"] {
            border: 1px solid #2e7d32;
            background-color: #2e7d32;
        }

        div[data-testid="stHorizontalBlock"] > div:nth-child(2) div[data-testid="stMetricLabel"] {
            color: #ffffff !important;
        }

        div[data-testid="stHorizontalBlock"] > div:nth-child(2) div[data-testid="stMetricValue"] {
            color: #ffffff !important;
        }
        
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# HEADER
# ============================================================

st.markdown(
    '<div class="main-title">🧮 Manual Default Calculation</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="sub-title">'
    "Calculate EOBI contributions manually for one or multiple "
    "default periods."
    "</div>",
    unsafe_allow_html=True,
)


# ============================================================
# BACK BUTTON
# ============================================================

if st.button("⬅️ Back to Notice Generator"):
    st.switch_page("app.py")


st.divider()


# ============================================================
# BASIC INPUTS
# ============================================================

st.subheader("Calculation Details")

col1, col2 = st.columns(2)

with col1:
    number_of_periods = st.number_input(
        "Number of Default Periods",
        min_value=1,
        max_value=20,
        value=1,
        step=1,
    )

with col2:
    number_of_ips = st.number_input(
        "Number of Insured Persons (IPs)",
        min_value=1,
        value=1,
        step=1,
    )


# ============================================================
# MONTH / YEAR OPTIONS
# ============================================================

month_names = [
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


# ============================================================
# DEFAULT PERIODS
# ============================================================

st.subheader("Default Periods")

st.info(
    "Enter only the month and year. Exact dates are not required. "
    "For example, July 2023 to June 2024 represents 12 months."
)

manual_periods = []


for i in range(int(number_of_periods)):

    st.markdown(f"### Period {i + 1}")

    col1, col2 = st.columns(2)

    # --------------------------------------------------------
    # FROM
    # --------------------------------------------------------

    with col1:

        st.markdown("**From**")

        from_col1, from_col2 = st.columns(2)

        with from_col1:
            from_month = st.selectbox(
                "Month",
                options=range(1, 13),
                format_func=lambda x: month_names[x - 1],
                index=6,  # July
                key=f"from_month_{i}",
            )

        with from_col2:
            from_year = st.number_input(
                "Year",
                min_value=1900,
                max_value=2100,
                value=2023,
                step=1,
                key=f"from_year_{i}",
            )

    # --------------------------------------------------------
    # TO
    # --------------------------------------------------------

    with col2:

        st.markdown("**To**")

        to_col1, to_col2 = st.columns(2)

        with to_col1:
            to_month = st.selectbox(
                "Month",
                options=range(1, 13),
                format_func=lambda x: month_names[x - 1],
                index=5,  # June
                key=f"to_month_{i}",
            )

        with to_col2:
            to_year = st.number_input(
                "Year",
                min_value=1900,
                max_value=2100,
                value=2024,
                step=1,
                key=f"to_year_{i}",
            )

    # --------------------------------------------------------
    # INTERNAL DATE CONVERSION
    #
    # The calculator does not require the user to enter days.
    #
    # Example:
    #
    # July 2023 -> June 2024
    #
    # is internally converted to:
    #
    # 01-07-2023 -> 01-07-2024
    #
    # because split_into_fy() treats the end date as exclusive
    # when it falls on the first day of a month.
    #
    # IMPORTANT:
    # Use datetime.datetime here, not datetime.date,
    # because generator.py expects datetime objects.
    # --------------------------------------------------------

    from_date = datetime(
        int(from_year),
        int(from_month),
        1,
    )

    if int(to_month) == 12:
        calculation_to_date = datetime(
            int(to_year) + 1,
            1,
            1,
        )
    else:
        calculation_to_date = datetime(
            int(to_year),
            int(to_month) + 1,
            1,
        )

    manual_periods.append(
        (
            from_date,
            calculation_to_date,
        )
    )

    st.divider()


# ============================================================
# CALCULATE BUTTON
# ============================================================

calculate_button = st.button(
    "🧮 Calculate Default Contribution",
    type="primary",
    width="stretch",
)


# ============================================================
# CALCULATION
# ============================================================

if calculate_button:

    try:

        with st.spinner("Calculating default contribution..."):

            result = calculate_default_contribution(
                periods=manual_periods,
                ips=int(number_of_ips),
            )

        st.success("Calculation completed successfully.")


        # ====================================================
        # DETAILED CALCULATION TABLE
        # ====================================================

        st.subheader("Calculation Details")

        st.dataframe(
            result["rows"],
            width="stretch",
            hide_index=True,
            column_config={
                "From": st.column_config.TextColumn(
                    "From",
                ),
                "To": st.column_config.TextColumn(
                    "To",
                ),
                "No. of Months": st.column_config.NumberColumn(
                    "No. of Months",
                    format="%d",
                ),
                "No. IPs": st.column_config.NumberColumn(
                    "No. IPs",
                    format="%d",
                ),
                "Minimum Wages": st.column_config.NumberColumn(
                    "Minimum Wages",
                    format="%,d",
                ),
                "Contribution Rate": st.column_config.NumberColumn(
                    "Contribution Rate",
                    format="%,d",
                ),
                "Assessed Amount": st.column_config.NumberColumn(
                    "Assessed Amount",
                    format="%,d",
                ),
                "Paid Amount": st.column_config.TextColumn(
                    "Paid Amount",
                ),
                "Principal Payable": st.column_config.NumberColumn(
                    "Principal Payable",
                    format="%,d",
                ),
            },
        )


        # ====================================================
        # SUMMARY
        # ====================================================

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
                f"{result['total_assessed']:,}",
            ) 
           
        with col3:
            st.metric(
                "Employer Contribution",
               f"{(result['total_assessed'] / 6) * 5:,.0f}"
            ) 

        with col4:
            st.metric(
                "Employees Contribution",
                f"{(result['total_assessed'] / 6):,.0f}",
            )

        with col5:
            st.metric(
                "50% Statutory Increase",
                f"{result['statutory_amount']:,}",
            )

        with col6:
            st.metric(
                "Total Payable",
                f"{result['final_total']:,}",
            )


        # ====================================================
        # CALCULATION METHOD
        # ====================================================

        st.subheader("Calculation Method")

        st.markdown(
            """
            **Assessed Contribution**

            `Contribution Rate × Number of IPs × Number of Months`

            **Statutory Increase**

            `50% × Principal Contribution`

            **Total Payable**

            `Principal Contribution + Statutory Increase`
            """
        )

        st.info(
            "The applicable minimum wage and contribution rate are "
            "selected according to the Financial Year rates defined "
            "in generator.py."
        )


    except ValueError as e:

        st.error(str(e))


    except Exception as e:

        st.error("An unexpected error occurred during calculation.")
        st.exception(e)