import streamlit as st
import pandas as pd
import sqlite3
from io import BytesIO

st.title("Transcript Tracker Backend Generator")

# Upload only 1 file
uploaded_file = st.file_uploader("Upload Final Output file (CSV or Excel)", type=["csv", "xlsx"])

if uploaded_file:
    # Load file
    if uploaded_file.name.endswith("xlsx"):
        df = pd.read_excel(uploaded_file, sheet_name="Calendar-Transcript dates")
    else:
        df = pd.read_csv(uploaded_file)

    # Create SQLite in-memory DB
    conn = sqlite3.connect(":memory:")
    df.to_sql("sharepoint_table", conn, index=False, if_exists="replace")

    # SQL Query
    query = """
WITH unpivoted AS (
    SELECT 
        "ADEK Applicant ID",
        "Student Name",
        "College as per Master",
        "Academic Status as per Master",
        Term,
        Country,
        "Current Mentor",
        "Team Lead",
        "ADEK Advisor",
        "Transcript Dates - July Tier",
        "Transcript Dates - Aug Tier",
        "Transcript Dates - Sep Tier",
        'Jul-25' AS month_col,
        "Jul-25" AS status
    FROM sharepoint_table

    UNION ALL

    SELECT 
        "ADEK Applicant ID",
        "Student Name",
        "College as per Master",
        "Academic Status as per Master",
        Term,
        Country,
        "Current Mentor",
        "Team Lead",
        "ADEK Advisor",
        "Transcript Dates - July Tier",
        "Transcript Dates - Aug Tier",
        "Transcript Dates - Sep Tier",
        'Aug-25' AS month_col,
        "Aug-25" AS status
    FROM sharepoint_table

    UNION ALL

    SELECT 
        "ADEK Applicant ID",
        "Student Name",
        "College as per Master",
        "Academic Status as per Master",
        Term,
        Country,
        "Current Mentor",
        "Team Lead",
        "ADEK Advisor",
        "Transcript Dates - July Tier",
        "Transcript Dates - Aug Tier",
        "Transcript Dates - Sep Tier",
        'Sep-25' AS month_col,
        "Sep-25" AS status
    FROM sharepoint_table

    UNION ALL

    SELECT 
        "ADEK Applicant ID",
        "Student Name",
        "College as per Master",
        "Academic Status as per Master",
        Term,
        Country,
        "Current Mentor",
        "Team Lead",
        "ADEK Advisor",
        "Transcript Dates - July Tier",
        "Transcript Dates - Aug Tier",
        "Transcript Dates - Sep Tier",
        'Oct-25' AS month_col,
        "Oct-25" AS status
    FROM sharepoint_table

    UNION ALL

    SELECT 
        "ADEK Applicant ID",
        "Student Name",
        "College as per Master",
        "Academic Status as per Master",
        Term,
        Country,
        "Current Mentor",
        "Team Lead",
        "ADEK Advisor",
        "Transcript Dates - July Tier",
        "Transcript Dates - Aug Tier",
        "Transcript Dates - Sep Tier",
        'Nov-25' AS month_col,
        "Nov-25" AS status
    FROM sharepoint_table

    UNION ALL

    SELECT 
        "ADEK Applicant ID",
        "Student Name",
        "College as per Master",
        "Academic Status as per Master",
        Term,
        Country,
        "Current Mentor",
        "Team Lead",
        "ADEK Advisor",
        "Transcript Dates - July Tier",
        "Transcript Dates - Aug Tier",
        "Transcript Dates - Sep Tier",
        'Dec-25' AS month_col,
        "Dec-25" AS status
    FROM sharepoint_table

    UNION ALL

    SELECT 
        "ADEK Applicant ID",
        "Student Name",
        "College as per Master",
        "Academic Status as per Master",
        Term,
        Country,
        "Current Mentor",
        "Team Lead",
        "ADEK Advisor",
        "Transcript Dates - July Tier",
        "Transcript Dates - Aug Tier",
        "Transcript Dates - Sep Tier",
        'Jan-26' AS month_col,
        "Jan-26" AS status
    FROM sharepoint_table

    UNION ALL

    SELECT 
        "ADEK Applicant ID",
        "Student Name",
        "College as per Master",
        "Academic Status as per Master",
        Term,
        Country,
        "Current Mentor",
        "Team Lead",
        "ADEK Advisor",
        "Transcript Dates - July Tier",
        "Transcript Dates - Aug Tier",
        "Transcript Dates - Sep Tier",
        'Feb-26' AS month_col,
        "Feb-26" AS status
    FROM sharepoint_table

    UNION ALL

    SELECT 
        "ADEK Applicant ID",
        "Student Name",
        "College as per Master",
        "Academic Status as per Master",
        Term,
        Country,
        "Current Mentor",
        "Team Lead",
        "ADEK Advisor",
        "Transcript Dates - July Tier",
        "Transcript Dates - Aug Tier",
        "Transcript Dates - Sep Tier",
        'Mar-26' AS month_col,
        "Mar-26" AS status
    FROM sharepoint_table

    UNION ALL

    SELECT 
        "ADEK Applicant ID",
        "Student Name",
        "College as per Master",
        "Academic Status as per Master",
        Term,
        Country,
        "Current Mentor",
        "Team Lead",
        "ADEK Advisor",
        "Transcript Dates - July Tier",
        "Transcript Dates - Aug Tier",
        "Transcript Dates - Sep Tier",
        'Apr-26' AS month_col,
        "Apr-26" AS status
    FROM sharepoint_table

    UNION ALL

    SELECT 
        "ADEK Applicant ID",
        "Student Name",
        "College as per Master",
        "Academic Status as per Master",
        Term,
        Country,
        "Current Mentor",
        "Team Lead",
        "ADEK Advisor",
        "Transcript Dates - July Tier",
        "Transcript Dates - Aug Tier",
        "Transcript Dates - Sep Tier",
        'May-26' AS month_col,
        "May-26" AS status
    FROM sharepoint_table

    UNION ALL

    SELECT 
        "ADEK Applicant ID",
        "Student Name",
        "College as per Master",
        "Academic Status as per Master",
        Term,
        Country,
        "Current Mentor",
        "Team Lead",
        "ADEK Advisor",
        "Transcript Dates - July Tier",
        "Transcript Dates - Aug Tier",
        "Transcript Dates - Sep Tier",
        'Jun-26' AS month_col,
        "Jun-26" AS status
    FROM sharepoint_table

    UNION ALL

    SELECT 
        "ADEK Applicant ID",
        "Student Name",
        "College as per Master",
        "Academic Status as per Master",
        Term,
        Country,
        "Current Mentor",
        "Team Lead",
        "ADEK Advisor",
        "Transcript Dates - July Tier",
        "Transcript Dates - Aug Tier",
        "Transcript Dates - Sep Tier",
        'Jul-26' AS month_col,
        "Jul-26" AS status
    FROM sharepoint_table

    UNION ALL

    SELECT 
        "ADEK Applicant ID",
        "Student Name",
        "College as per Master",
        "Academic Status as per Master",
        Term,
        Country,
        "Current Mentor",
        "Team Lead",
        "ADEK Advisor",
        "Transcript Dates - July Tier",
        "Transcript Dates - Aug Tier",
        "Transcript Dates - Sep Tier",
        'Aug-26' AS month_col,
        "Aug-26" AS status
    FROM sharepoint_table

    UNION ALL

    SELECT 
        "ADEK Applicant ID",
        "Student Name",
        "College as per Master",
        "Academic Status as per Master",
        Term,
        Country,
        "Current Mentor",
        "Team Lead",
        "ADEK Advisor",
        "Transcript Dates - July Tier",
        "Transcript Dates - Aug Tier",
        "Transcript Dates - Sep Tier",
        'Sep-26' AS month_col,
        "Sep-26" AS status
    FROM sharepoint_table
)
SELECT
    "ADEK Applicant ID",
    "Student Name",
    "College as per Master",
    "Academic Status as per Master",
    Term,
    Country,
    "Current Mentor",
    "Team Lead",
    "ADEK Advisor",
    "Transcript Dates - July Tier",
    "Transcript Dates - Aug Tier",
    "Transcript Dates - Sep Tier",
    TRIM(GROUP_CONCAT(CASE WHEN status = 'TEC' THEN month_col END, ', ')) AS "Transcript Expected by College",
    TRIM(GROUP_CONCAT(CASE WHEN status = 'TEM' THEN month_col END, ', ')) AS "Transcript Expected by Mentor",
    TRIM(GROUP_CONCAT(CASE WHEN status = 'TE' THEN month_col END, ', ')) AS "Transcript Expected by College & Mentor",
    TRIM(GROUP_CONCAT(CASE WHEN status = 'TR' THEN month_col END, ', ')) AS "Transcript Received"
FROM unpivoted
GROUP BY
    "ADEK Applicant ID",
    "Student Name",
    "College as per Master",
    "Academic Status as per Master",
    Term,
    Country,
    "Current Mentor",
    "Team Lead",
    "ADEK Advisor",
    "Transcript Dates - July Tier",
    "Transcript Dates - Aug Tier",
    "Transcript Dates - Sep Tier";

    """

    # Run query
    result_df = pd.read_sql_query(query, conn)

    st.subheader("Filtered Result")
    st.dataframe(result_df)

    # Create Excel file in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        result_df.to_excel(writer, index=False, sheet_name="Data")

    excel_data = output.getvalue()

    # Download button for Excel
    st.download_button(
        label="📥 Download Result as Excel",
        data=excel_data,
        file_name="Transcript Tracker Dashboard Backend 1.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
