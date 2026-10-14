import pandas as pd
from pathlib import Path

# =========================================================
# FILE LOCATIONS
# =========================================================

input_file = Path(
    r"C:\Users\mrfqu\Desktop\University_Academic_Analytics\02_Data_Cleaning\Cleaned_Academic_Data.xlsx"
)

output_file = Path(
    r"C:\Users\mrfqu\Desktop\University_Academic_Analytics\03_Python_EDA\Engineered_Academic_Data.xlsx"
)

if not input_file.exists():

    raise FileNotFoundError(
        f"Cleaned dataset not found: {input_file}"
    )

print()
print("=" * 65)
print("UNIVERSITY ACADEMIC ANALYTICS - FEATURE ENGINEERING")
print("=" * 65)

# =========================================================
# 1. LOAD DATA
# =========================================================

master = pd.read_excel(
    input_file,
    sheet_name="Analytics_Master"
)

student_summary = pd.read_excel(
    input_file,
    sheet_name="Student_Summary"
)

semester = pd.read_excel(
    input_file,
    sheet_name="Semester_Performance"
)

print()
print("Datasets loaded successfully.")

print(
    "Analytics Master:",
    len(master)
)

print(
    "Student Summary:",
    len(student_summary)
)

print(
    "Semester Performance:",
    len(semester)
)

# =========================================================
# 2. SCORE CATEGORY FUNCTION
# =========================================================

def score_category(score):

    if score >= 90:
        return "Excellent"

    elif score >= 80:
        return "Very Good"

    elif score >= 70:
        return "Good"

    elif score >= 60:
        return "Average"

    elif score >= 50:
        return "Needs Improvement"

    else:
        return "Fail"


# =========================================================
# 3. ATTENDANCE CATEGORY FUNCTION
# =========================================================

def attendance_category(attendance):

    if attendance < 75:
        return "Low Attendance"

    elif attendance < 85:
        return "Satisfactory Attendance"

    else:
        return "High Attendance"


# =========================================================
# 4. SUBJECT-LEVEL FEATURES
# =========================================================

master[
    "Score_Category"
] = master[
    "Total_Marks"
].apply(
    score_category
)

master[
    "Attendance_Category"
] = master[
    "Attendance_Percentage"
].apply(
    attendance_category
)

master[
    "Pass_Flag"
] = (
    master["Result"] == "Pass"
).astype(int)

master[
    "Fail_Flag"
] = (
    master["Result"] == "Fail"
).astype(int)

master[
    "Attendance_Shortage_Flag"
] = (
    master[
        "Attendance_Percentage"
    ] < 75
).astype(int)

print()
print("Subject-level analytical features created.")

# =========================================================
# 5. STUDENT PERFORMANCE CATEGORY
# =========================================================

student_summary[
    "Performance_Category"
] = student_summary[
    "Average_Score"
].apply(
    score_category
)

student_summary[
    "Attendance_Category"
] = student_summary[
    "Average_Attendance"
].apply(
    attendance_category
)

# =========================================================
# 6. ACADEMIC STATUS
# =========================================================

student_summary[
    "Academic_Status"
] = student_summary[
    "Final_Active_Backlogs"
].apply(
    lambda x:
    "All Clear"
    if x == 0
    else "Active Backlog"
)

# =========================================================
# 7. OVERALL STUDENT RANK
# =========================================================

student_summary = (
    student_summary
    .sort_values(
        by=[
            "Average_Score",
            "Final_CGPA",
            "Average_Attendance"
        ],
        ascending=[
            False,
            False,
            False
        ]
    )
    .reset_index(
        drop=True
    )
)

student_summary[
    "Overall_Rank"
] = range(
    1,
    len(student_summary) + 1
)

# =========================================================
# 8. DEPARTMENT RANK
# =========================================================

student_summary[
    "Department_Rank"
] = (
    student_summary
    .groupby(
        "Department"
    )
    .cumcount()
    + 1
)

# =========================================================
# 9. TOP PERFORMER FLAGS
# =========================================================

student_summary[
    "Top_10_Performer"
] = student_summary[
    "Overall_Rank"
].apply(
    lambda x:
    "Yes"
    if x <= 10
    else "No"
)

student_summary[
    "Department_Topper"
] = student_summary[
    "Department_Rank"
].apply(
    lambda x:
    "Yes"
    if x == 1
    else "No"
)

print(
    "Student ranking features created."
)

# =========================================================
# 10. PERFORMANCE TREND
# =========================================================
#
# Compare Semester 1 SGPA with Semester 8 SGPA.
#
# More than +0.50 = Improving
# Between -0.50 and +0.50 = Stable
# Less than -0.50 = Declining
#
# =========================================================

semester_1 = (
    semester[
        semester["Semester"] == 1
    ][
        [
            "Student_ID",
            "SGPA"
        ]
    ]
    .rename(
        columns={
            "SGPA":
            "Semester_1_SGPA"
        }
    )
)

semester_8 = (
    semester[
        semester["Semester"] == 8
    ][
        [
            "Student_ID",
            "SGPA"
        ]
    ]
    .rename(
        columns={
            "SGPA":
            "Semester_8_SGPA"
        }
    )
)

student_summary = (
    student_summary
    .merge(
        semester_1,
        on="Student_ID",
        how="left"
    )
    .merge(
        semester_8,
        on="Student_ID",
        how="left"
    )
)

student_summary[
    "SGPA_Change"
] = (
    student_summary[
        "Semester_8_SGPA"
    ]
    -
    student_summary[
        "Semester_1_SGPA"
    ]
).round(2)


def performance_trend(change):

    if change > 0.50:
        return "Improving"

    elif change < -0.50:
        return "Declining"

    else:
        return "Stable"


student_summary[
    "Performance_Trend"
] = student_summary[
    "SGPA_Change"
].apply(
    performance_trend
)

print(
    "Semester performance trend created."
)

# =========================================================
# 11. SUBJECT PERFORMANCE SUMMARY
# =========================================================

subject_features = (
    master
    .groupby(
        [
            "Subject_ID",
            "Subject_Name",
            "Department",
            "Semester"
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

        Total_Students=(
            "Student_ID",
            "count"
        ),

        Pass_Count=(
            "Pass_Flag",
            "sum"
        ),

        Fail_Count=(
            "Fail_Flag",
            "sum"
        )
    )
)

subject_features[
    "Average_Score"
] = subject_features[
    "Average_Score"
].round(2)

subject_features[
    "Average_Attendance"
] = subject_features[
    "Average_Attendance"
].round(2)

subject_features[
    "Pass_Percentage"
] = (
    subject_features[
        "Pass_Count"
    ]
    /
    subject_features[
        "Total_Students"
    ]
    *
    100
).round(2)

subject_features[
    "Failure_Percentage"
] = (
    100
    -
    subject_features[
        "Pass_Percentage"
    ]
).round(2)

# =========================================================
# 12. SUBJECT DIFFICULTY CATEGORY
# =========================================================

def difficulty_category(
    pass_percentage
):

    if pass_percentage < 80:
        return "High Difficulty"

    elif pass_percentage < 90:
        return "Moderate Difficulty"

    else:
        return "Lower Difficulty"


subject_features[
    "Difficulty_Category"
] = subject_features[
    "Pass_Percentage"
].apply(
    difficulty_category
)

# =========================================================
# 13. SUBJECT RANK WITHIN DEPARTMENT
# =========================================================

subject_features = (
    subject_features
    .sort_values(
        by=[
            "Department",
            "Average_Score"
        ],
        ascending=[
            True,
            False
        ]
    )
)

subject_features[
    "Department_Subject_Rank"
] = (
    subject_features
    .groupby(
        "Department"
    )
    .cumcount()
    + 1
)

# =========================================================
# 14. DEPARTMENT SUMMARY
# =========================================================

department_features = (
    master
    .groupby(
        "Department",
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

        Total_Records=(
            "Record_ID",
            "count"
        ),

        Pass_Count=(
            "Pass_Flag",
            "sum"
        )
    )
)

department_features[
    "Average_Score"
] = department_features[
    "Average_Score"
].round(2)

department_features[
    "Average_Attendance"
] = department_features[
    "Average_Attendance"
].round(2)

department_features[
    "Pass_Percentage"
] = (
    department_features[
        "Pass_Count"
    ]
    /
    department_features[
        "Total_Records"
    ]
    *
    100
).round(2)

# =========================================================
# 15. STUDENTS WITH ACTIVE BACKLOGS BY DEPARTMENT
# =========================================================

backlog_summary = (
    student_summary
    .groupby(
        "Department",
        as_index=False
    )
    .agg(

        Total_Students=(
            "Student_ID",
            "count"
        ),

        Students_With_Backlogs=(
            "Final_Active_Backlogs",
            lambda x:
            (x > 0).sum()
        )
    )
)

department_features = (
    department_features
    .merge(
        backlog_summary,
        on="Department",
        how="left"
    )
)

department_features[
    "Students_Without_Backlogs"
] = (
    department_features[
        "Total_Students"
    ]
    -
    department_features[
        "Students_With_Backlogs"
    ]
)

# =========================================================
# 16. DEPARTMENT PERFORMANCE RANK
# =========================================================

department_features = (
    department_features
    .sort_values(
        by=[
            "Average_Score",
            "Pass_Percentage"
        ],
        ascending=[
            False,
            False
        ]
    )
    .reset_index(
        drop=True
    )
)

department_features[
    "Department_Rank"
] = range(
    1,
    len(department_features) + 1
)

# =========================================================
# 17. FEATURE VALIDATION
# =========================================================

print()
print("=" * 65)
print("FEATURE VALIDATION")
print("=" * 65)

print(
    "Master Feature Rows:",
    len(master)
)

print(
    "Student Feature Rows:",
    len(student_summary)
)

print(
    "Subject Feature Rows:",
    len(subject_features)
)

print(
    "Department Feature Rows:",
    len(department_features)
)

print()

print(
    "Missing values in Master Features:",
    master.isnull().sum().sum()
)

print(
    "Missing values in Student Features:",
    student_summary.isnull().sum().sum()
)

print(
    "Missing values in Subject Features:",
    subject_features.isnull().sum().sum()
)

print(
    "Missing values in Department Features:",
    department_features.isnull().sum().sum()
)

# =========================================================
# 18. FEATURE SUMMARY
# =========================================================

print()
print("=" * 65)
print("STUDENT PERFORMANCE CATEGORIES")
print("=" * 65)

print(
    student_summary[
        "Performance_Category"
    ].value_counts()
)

print()
print("=" * 65)
print("STUDENT ATTENDANCE CATEGORIES")
print("=" * 65)

print(
    student_summary[
        "Attendance_Category"
    ].value_counts()
)

print()
print("=" * 65)
print("PERFORMANCE TREND")
print("=" * 65)

print(
    student_summary[
        "Performance_Trend"
    ].value_counts()
)

print()
print("=" * 65)
print("ACADEMIC STATUS")
print("=" * 65)

print(
    student_summary[
        "Academic_Status"
    ].value_counts()
)

print()
print("=" * 65)
print("SUBJECT DIFFICULTY")
print("=" * 65)

print(
    subject_features[
        "Difficulty_Category"
    ].value_counts()
)

# =========================================================
# 19. TOP 10 STUDENTS AFTER RANKING
# =========================================================

print()
print("=" * 65)
print("TOP 10 STUDENTS")
print("=" * 65)

print(
    student_summary[
        [
            "Overall_Rank",
            "Student_ID",
            "Department",
            "Average_Score",
            "Final_CGPA",
            "Average_Attendance",
            "Department_Rank"
        ]
    ]
    .head(10)
    .to_string(
        index=False
    )
)

# =========================================================
# 20. DEPARTMENT RANKING
# =========================================================

print()
print("=" * 65)
print("DEPARTMENT RANKING")
print("=" * 65)

print(
    department_features[
        [
            "Department_Rank",
            "Department",
            "Average_Score",
            "Average_Attendance",
            "Pass_Percentage",
            "Students_With_Backlogs"
        ]
    ].to_string(
        index=False
    )
)

# =========================================================
# 21. SAVE ENGINEERED DATA
# =========================================================

with pd.ExcelWriter(
    output_file,
    engine="openpyxl"
) as writer:

    master.to_excel(
        writer,
        sheet_name="Master_Features",
        index=False
    )

    student_summary.to_excel(
        writer,
        sheet_name="Student_Features",
        index=False
    )

    subject_features.to_excel(
        writer,
        sheet_name="Subject_Features",
        index=False
    )

    department_features.to_excel(
        writer,
        sheet_name="Department_Features",
        index=False
    )

print()
print("=" * 65)
print("FEATURE ENGINEERING COMPLETED SUCCESSFULLY")
print("=" * 65)

print()
print(
    "Engineered dataset saved to:"
)

print(
    output_file
)

print()
print(
    "Step 8 - Feature Engineering is complete."
)

print("=" * 65)