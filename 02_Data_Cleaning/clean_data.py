import pandas as pd
from pathlib import Path

# =========================================================
# FILE LOCATIONS
# =========================================================

input_file = Path(
    r"C:\Users\mrfqu\Desktop\University_Academic_Analytics\01_Dataset\Academic_Data_Model.xlsx"
)

output_file = Path(
    r"C:\Users\mrfqu\Desktop\University_Academic_Analytics\02_Data_Cleaning\Cleaned_Academic_Data.xlsx"
)

if not input_file.exists():
    raise FileNotFoundError(
        f"Input Excel file not found: {input_file}"
    )

print()
print("=" * 60)
print("UNIVERSITY ACADEMIC ANALYTICS - DATA CLEANING")
print("=" * 60)

# =========================================================
# 1. LOAD DATA
# =========================================================

students = pd.read_excel(
    input_file,
    sheet_name="Students"
)

subjects = pd.read_excel(
    input_file,
    sheet_name="Subjects"
)

academic = pd.read_excel(
    input_file,
    sheet_name="Academic_Records"
)

attendance = pd.read_excel(
    input_file,
    sheet_name="Attendance"
)

reexam = pd.read_excel(
    input_file,
    sheet_name="Reexam_Records"
)

semester = pd.read_excel(
    input_file,
    sheet_name="Semester_Performance"
)

print("\nData loaded successfully.")

# =========================================================
# 2. REMOVE EXTRA SPACES FROM COLUMN NAMES
# =========================================================

datasets = {
    "Students": students,
    "Subjects": subjects,
    "Academic_Records": academic,
    "Attendance": attendance,
    "Reexam_Records": reexam,
    "Semester_Performance": semester
}

for name, df in datasets.items():
    df.columns = df.columns.str.strip()

print("Column names standardized.")

# =========================================================
# 3. CLEAN TEXT COLUMNS
# =========================================================

text_columns = {

    "Students": [
        "Student_ID",
        "Department"
    ],

    "Subjects": [
        "Subject_ID",
        "Subject_Name",
        "Department",
        "Course_Type"
    ],

    "Academic_Records": [
        "Record_ID",
        "Student_ID",
        "Subject_ID",
        "Grade",
        "Result"
    ],

    "Attendance": [
        "Student_ID",
        "Subject_ID"
    ],

    "Reexam_Records": [
        "Reexam_ID",
        "Student_ID",
        "Subject_ID",
        "Grade",
        "Result"
    ],

    "Semester_Performance": [
        "Student_ID"
    ]
}

for table_name, columns in text_columns.items():

    df = datasets[table_name]

    for column in columns:

        if column in df.columns:

            df[column] = (
                df[column]
                .astype(str)
                .str.strip()
            )

print("Text values standardized.")

# =========================================================
# 4. STANDARDIZE CATEGORY VALUES
# =========================================================

students["Department"] = (
    students["Department"]
    .str.upper()
)

subjects["Department"] = (
    subjects["Department"]
    .str.upper()
)

academic["Result"] = (
    academic["Result"]
    .str.title()
)

reexam["Result"] = (
    reexam["Result"]
    .str.title()
)

academic["Grade"] = (
    academic["Grade"]
    .str.upper()
)

reexam["Grade"] = (
    reexam["Grade"]
    .str.upper()
)

print("Categorical values standardized.")

# =========================================================
# 5. CONVERT NUMERIC COLUMNS
# =========================================================

students["Admission_Year"] = pd.to_numeric(
    students["Admission_Year"],
    errors="coerce"
)

numeric_subject_columns = [
    "Semester",
    "Credits"
]

for column in numeric_subject_columns:

    subjects[column] = pd.to_numeric(
        subjects[column],
        errors="coerce"
    )

numeric_academic_columns = [
    "Internal_Marks",
    "External_Marks",
    "Assignment_Score",
    "Total_Marks"
]

for column in numeric_academic_columns:

    academic[column] = pd.to_numeric(
        academic[column],
        errors="coerce"
    )

numeric_attendance_columns = [
    "Classes_Held",
    "Classes_Attended",
    "Attendance_Percentage"
]

for column in numeric_attendance_columns:

    attendance[column] = pd.to_numeric(
        attendance[column],
        errors="coerce"
    )

numeric_reexam_columns = [
    "Original_Semester",
    "Attempt_Semester",
    "Attempt_Number",
    "Original_Marks",
    "Reexam_Marks"
]

for column in numeric_reexam_columns:

    reexam[column] = pd.to_numeric(
        reexam[column],
        errors="coerce"
    )

numeric_semester_columns = [
    "Semester",
    "SGPA",
    "CGPA",
    "New_Backlogs",
    "Cleared_Backlogs",
    "Active_Backlogs",
    "Semester_Credits_Cleared",
    "Cumulative_Credits_Cleared"
]

for column in numeric_semester_columns:

    semester[column] = pd.to_numeric(
        semester[column],
        errors="coerce"
    )

print("Numeric data types standardized.")

# =========================================================
# 6. ROUND DECIMAL VALUES
# =========================================================

attendance["Attendance_Percentage"] = (
    attendance[
        "Attendance_Percentage"
    ].round(2)
)

semester["SGPA"] = (
    semester["SGPA"].round(2)
)

semester["CGPA"] = (
    semester["CGPA"].round(2)
)

print("Decimal values standardized.")

# =========================================================
# 7. REMOVE EXACT DUPLICATE ROWS
# =========================================================

before_counts = {}

after_counts = {}

for table_name, df in datasets.items():

    before_counts[table_name] = len(df)

    df.drop_duplicates(
        inplace=True
    )

    after_counts[table_name] = len(df)

print("Duplicate row check completed.")

# =========================================================
# 8. MISSING VALUE CHECK
# =========================================================

print()
print("=" * 60)
print("MISSING VALUE CHECK")
print("=" * 60)

total_missing = 0

for table_name, df in datasets.items():

    missing = int(
        df.isnull().sum().sum()
    )

    total_missing += missing

    print(
        f"{table_name}: {missing}"
    )

# =========================================================
# 9. CREATE ANALYTICS MASTER TABLE
# =========================================================
#
# One row represents one student's performance
# in one subject.
#
# This table will be very useful for Python EDA.
# =========================================================

analytics_master = (
    academic
    .merge(
        students,
        on="Student_ID",
        how="left"
    )
    .merge(
        subjects,
        on="Subject_ID",
        how="left",
        suffixes=(
            "_Student",
            "_Subject"
        )
    )
    .merge(
        attendance,
        on=[
            "Student_ID",
            "Subject_ID"
        ],
        how="left"
    )
)

# Remove duplicated department information
# after confirming both match.

if (
    "Department_Student"
    in analytics_master.columns
    and
    "Department_Subject"
    in analytics_master.columns
):

    department_mismatch = (
        analytics_master[
            "Department_Student"
        ]
        !=
        analytics_master[
            "Department_Subject"
        ]
    ).sum()

    if department_mismatch != 0:

        raise ValueError(
            "Department mismatch detected while creating Analytics Master."
        )

    analytics_master.rename(
        columns={
            "Department_Student":
            "Department"
        },
        inplace=True
    )

    analytics_master.drop(
        columns=[
            "Department_Subject"
        ],
        inplace=True
    )

# =========================================================
# 10. REORDER MASTER TABLE COLUMNS
# =========================================================

master_columns = [
    "Record_ID",
    "Student_ID",
    "Department",
    "Admission_Year",
    "Subject_ID",
    "Subject_Name",
    "Semester",
    "Course_Type",
    "Credits",
    "Internal_Marks",
    "External_Marks",
    "Assignment_Score",
    "Total_Marks",
    "Grade",
    "Result",
    "Classes_Held",
    "Classes_Attended",
    "Attendance_Percentage"
]

analytics_master = analytics_master[
    master_columns
]

# =========================================================
# 11. CREATE STUDENT SUMMARY TABLE
# =========================================================

student_summary = (
    analytics_master
    .groupby(
        [
            "Student_ID",
            "Department"
        ],
        as_index=False
    )
    .agg(
        Average_Score=(
            "Total_Marks",
            "mean"
        ),
        Average_Attendance=(
            "Attendance_Percentage",
            "mean"
        ),
        Passed_Subjects=(
            "Result",
            lambda x:
            (x == "Pass").sum()
        ),
        Failed_Subjects=(
            "Result",
            lambda x:
            (x == "Fail").sum()
        )
    )
)

student_summary[
    "Average_Score"
] = student_summary[
    "Average_Score"
].round(2)

student_summary[
    "Average_Attendance"
] = student_summary[
    "Average_Attendance"
].round(2)

student_summary[
    "Pass_Percentage"
] = (
    student_summary[
        "Passed_Subjects"
    ]
    /
    (
        student_summary[
            "Passed_Subjects"
        ]
        +
        student_summary[
            "Failed_Subjects"
        ]
    )
    *
    100
).round(2)

# Add final CGPA and final active backlogs

final_semester = semester[
    semester["Semester"] == 8
][
    [
        "Student_ID",
        "CGPA",
        "Active_Backlogs",
        "Cumulative_Credits_Cleared"
    ]
]

student_summary = (
    student_summary
    .merge(
        final_semester,
        on="Student_ID",
        how="left"
    )
)

student_summary.rename(
    columns={
        "CGPA":
        "Final_CGPA",

        "Active_Backlogs":
        "Final_Active_Backlogs",

        "Cumulative_Credits_Cleared":
        "Final_Credits_Cleared"
    },
    inplace=True
)

# =========================================================
# 12. SAVE CLEANED DATA
# =========================================================

with pd.ExcelWriter(
    output_file,
    engine="openpyxl"
) as writer:

    students.to_excel(
        writer,
        sheet_name="Students",
        index=False
    )

    subjects.to_excel(
        writer,
        sheet_name="Subjects",
        index=False
    )

    academic.to_excel(
        writer,
        sheet_name="Academic_Records",
        index=False
    )

    attendance.to_excel(
        writer,
        sheet_name="Attendance",
        index=False
    )

    reexam.to_excel(
        writer,
        sheet_name="Reexam_Records",
        index=False
    )

    semester.to_excel(
        writer,
        sheet_name="Semester_Performance",
        index=False
    )

    analytics_master.to_excel(
        writer,
        sheet_name="Analytics_Master",
        index=False
    )

    student_summary.to_excel(
        writer,
        sheet_name="Student_Summary",
        index=False
    )

# =========================================================
# 13. FINAL CLEANING REPORT
# =========================================================

print()
print("=" * 60)
print("DATA CLEANING SUMMARY")
print("=" * 60)

for table_name in datasets:

    removed = (
        before_counts[table_name]
        -
        after_counts[table_name]
    )

    print(
        f"{table_name}: "
        f"{before_counts[table_name]} rows -> "
        f"{after_counts[table_name]} rows "
        f"(Duplicates removed: {removed})"
    )

print()
print(
    "Total Missing Values:",
    total_missing
)

print()
print(
    "Analytics Master Rows:",
    len(analytics_master)
)

print(
    "Student Summary Rows:",
    len(student_summary)
)

print()
print(
    "Cleaned dataset saved to:"
)

print(
    output_file
)

print()
print("=" * 60)

if (
    total_missing == 0
    and
    len(analytics_master) == 16000
    and
    len(student_summary) == 500
):

    print(
        "DATA CLEANING RESULT: SUCCESS"
    )

else:

    print(
        "DATA CLEANING RESULT: CHECK REQUIRED"
    )

print("=" * 60)