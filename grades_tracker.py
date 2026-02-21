import streamlit as st
import pandas as pd
import sqlite3
from io import BytesIO

st.title("Grades Tracker Backend Generator")

# Upload only 1 file
uploaded_file = st.file_uploader("Upload Final Output file (CSV or Excel)", type=["csv", "xlsx"])

if uploaded_file:
    # Load file
    if uploaded_file.name.endswith("xlsx"):
        df = pd.read_excel(uploaded_file, sheet_name="Transcript Trajectory Tracker")
    else:
        df = pd.read_csv(uploaded_file)

    # Instead of SQL UNPIVOT, use pandas.melt
    value_vars_grades = [
        "Grade / ESL Level (Feb-25)",
        "Grade / ESL Level (May-25)",
        "Grade / ESL Level (Jun-25)",
        "Grade / ESL Level (Jul-25)",
        "Grade / ESL Level (Feb-26)"
    ]

    value_vars_cgpa = [
        "CGPA Hours / Satisfactory & Unsatisfactory (Feb-25)",
        "CGPA Hours / Satisfactory & Unsatisfactory (May-25)",
        "CGPA Hours / Satisfactory & Unsatisfactory (Jun-25)",
        "CGPA Hours / Satisfactory & Unsatisfactory (Jul-25)",
        "CGPA Hours / Satisfactory & Unsatisfactory (Feb-26)"
    ]

    id_vars = [
        "ADEK Applicant ID",
        "Student Name",
        "Khotwa Program Status",
        "Academic Pathway  - Tier Report Latest",
        "Country",
        "College",
        "Major Standardized",
        "Academic Calendar System",
        "Current Mentor",
        "Team Lead"
    ]

    # Melt grade level
    df_grades = df.melt(
        id_vars=id_vars,
        value_vars=value_vars_grades,
        var_name="Grade / ESL Level Month",
        value_name="Grade / ESL Level"
    )

    # Melt CGPA
    df_cgpa = df.melt(
        id_vars=id_vars,
        value_vars=value_vars_cgpa,
        var_name="CGPA Hours / Satisfactory & Unsatisfactory Month",
        value_name="CGPA Hours / Satisfactory & Unsatisfactory"
    )

    # Align months (assuming order is same)
    result_df = pd.concat([df_grades, df_cgpa.drop(columns=id_vars)], axis=1)
    

    # Load into SQLite
    conn = sqlite3.connect(":memory:")
    result_df.to_sql("reshaped_table", conn, index=False, if_exists="replace")

    # Example SQL query
    query = """

    WITH base1 AS (SELECT "ADEK Applicant ID",
        "Student Name",
        "Khotwa Program Status",
        "Academic Pathway  - Tier Report Latest",
        Country,
        College,
        "Major Standardized",
        "Academic Calendar System",	
        "Current Mentor",
        "Team Lead",
        CASE WHEN "Grade / ESL Level Month" LIKE '%Feb-25%' THEN 'Feb-25'
          WHEN "Grade / ESL Level Month" LIKE '%May-25%' THEN 'May-25'
          WHEN "Grade / ESL Level Month" LIKE '%Jun-25%' THEN 'Jun-25'
          WHEN "Grade / ESL Level Month" LIKE '%Jul-25%' THEN 'Jul-25'
          WHEN "Grade / ESL Level Month" LIKE '%Feb-26%' THEN 'Feb-26'
          END AS Month,
        TRIM("Grade / ESL Level") AS "Grade / ESL Level",
        TRIM("CGPA Hours / Satisfactory & Unsatisfactory") AS "CGPA Hours / Satisfactory & Unsatisfactory"
        FROM reshaped_table)

    SELECT "ADEK Applicant ID",
        "Student Name",
        "Khotwa Program Status",
        "Academic Pathway  - Tier Report Latest",
        Country,
        College,
        "Major Standardized",
        "Academic Calendar System",	
        "Current Mentor",
        "Team Lead",
        Month,
        "Grade / ESL Level",
        CASE WHEN LOWER("Grade / ESL Level") LIKE '%level%' OR LOWER("Grade / ESL Level") LIKE '%fail%' OR LOWER("Grade / ESL Level") LIKE '%pass%' THEN "N/A"
            WHEN LENGTH("Grade / ESL Level") = 0 THEN "N/A"
            ELSE "Grade / ESL Level" END AS "Academic Grades",
        CASE WHEN LOWER("Grade / ESL Level") = 'level 0' THEN 0
            WHEN LOWER("Grade / ESL Level") = 'level 1' THEN 1
            WHEN LOWER("Grade / ESL Level") = 'level 2' THEN 2
            WHEN LOWER("Grade / ESL Level") = 'level 3' THEN 3
            WHEN LOWER("Grade / ESL Level") = 'level 4' THEN 4
            WHEN LOWER("Grade / ESL Level") = 'level 5' THEN 5
            WHEN LOWER("Grade / ESL Level") = 'level 6' THEN 6
            WHEN LOWER("Grade / ESL Level") = 'level 7' THEN 7
            WHEN LOWER("Grade / ESL Level") = 'level 8' THEN 8
            WHEN LOWER("Grade / ESL Level") = 'level 9' THEN 9
            ELSE "N/A" END AS 'ESL Level',
        "CGPA Hours / Satisfactory & Unsatisfactory"
    FROM base1;

    """

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
        file_name="Grades Tracker Dashboard Backend.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
