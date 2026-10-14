from pathlib import Path
import pandas as pd
import numpy as np


# ============================================================
# UNIVERSITY ACADEMIC PERFORMANCE ANALYTICS
# STEP 16 - EXPLAINABLE ACADEMIC RISK INDEX
# ============================================================


# ------------------------------------------------------------
# 1. FILE PATHS
# ------------------------------------------------------------

PROJECT_ROOT = Path(__file__).resolve().parents[1]

INPUT_FILE = (
    PROJECT_ROOT
    / "02_Data_Cleaning"
    / "Cleaned_Academic_Data.xlsx"
)

OUTPUT_FILE = (
    PROJECT_ROOT
    / "03_Python_EDA"
    / "Academic_Risk_Index.xlsx"
)


# ------------------------------------------------------------
# 2. CHECK INPUT FILE
# ------------------------------------------------------------

if not INPUT_FILE.exists():
    raise FileNotFoundError(
        f"\nInput file not found:\n{INPUT_FILE}\n"
        "\nMake sure Cleaned_Academic_Data.xlsx is inside "
        "02_Data_Cleaning."
    )


print("=" * 65)
print("STEP 16 - ACADEMIC RISK INDEX")
print("=" * 65)

print("\nLoading cleaned academic dataset...")


# ------------------------------------------------------------
# 3. LOAD REQUIRED SHEETS
# ------------------------------------------------------------

students = pd.read_excel(
    INPUT_FILE,
    sheet_name="Students"
)

student_summary = pd.read_excel(
    INPUT_FILE,
    sheet_name="Student_Summary"
)

semester_performance = pd.read_excel(
    INPUT_FILE,
    sheet_name="Semester_Performance"
)

reexam_records = pd.read_excel(
    INPUT_FILE,
    sheet_name="Reexam_Records"
)


# ------------------------------------------------------------
# 4. VALIDATE REQUIRED COLUMNS
# ------------------------------------------------------------

required_summary_columns = [
    "Student_ID",
    "Average_Score",
    "Average_Attendance",
    "Final_CGPA",
    "Final_Active_Backlogs"
]

required_semester_columns = [
    "Student_ID",
    "Semester",
    "SGPA"
]

required_reexam_columns = [
    "Student_ID",
    "Subject_ID"
]


def check_columns(df, required_columns, sheet_name):
    missing = [
        col for col in required_columns
        if col not in df.columns
    ]

    if missing:
        raise ValueError(
            f"\nMissing columns in {sheet_name}: {missing}\n"
            f"Available columns: {list(df.columns)}"
        )


check_columns(
    student_summary,
    required_summary_columns,
    "Student_Summary"
)

check_columns(
    semester_performance,
    required_semester_columns,
    "Semester_Performance"
)

check_columns(
    reexam_records,
    required_reexam_columns,
    "Reexam_Records"
)


# ------------------------------------------------------------
# 5. ADD DEPARTMENT IF REQUIRED
# ------------------------------------------------------------

risk_profile = student_summary.copy()

if "Department" not in risk_profile.columns:

    risk_profile = risk_profile.merge(
        students[
            [
                "Student_ID",
                "Department"
            ]
        ],
        on="Student_ID",
        how="left"
    )


# ------------------------------------------------------------
# 6. PERFORMANCE TRAJECTORY
# ------------------------------------------------------------
# Calculates semester-wise SGPA slope.
#
# Positive slope  -> improving
# Near zero       -> stable
# Negative slope  -> declining


trajectory_records = []


for student_id, group in semester_performance.groupby(
    "Student_ID"
):

    temp = group[
        [
            "Semester",
            "SGPA"
        ]
    ].dropna()

    temp = temp.sort_values(
        "Semester"
    )

    if len(temp) >= 2:

        x = temp["Semester"].astype(float).to_numpy()
        y = temp["SGPA"].astype(float).to_numpy()

        slope = np.polyfit(
            x,
            y,
            1
        )[0]

    else:
        slope = 0.0

    trajectory_records.append(
        {
            "Student_ID": student_id,
            "SGPA_Trend_Slope": slope
        }
    )


trajectory_df = pd.DataFrame(
    trajectory_records
)


# ------------------------------------------------------------
# 7. TRAJECTORY CLASSIFICATION
# ------------------------------------------------------------

def classify_trajectory(slope):

    if slope >= 0.10:
        return "Improving"

    elif slope >= -0.05:
        return "Stable"

    elif slope >= -0.20:
        return "Declining"

    else:
        return "Strongly Declining"


trajectory_df[
    "Performance_Trajectory"
] = trajectory_df[
    "SGPA_Trend_Slope"
].apply(
    classify_trajectory
)


# ------------------------------------------------------------
# 8. RE-EXAMINATION BURDEN
# ------------------------------------------------------------
# We calculate:
#
# Total Re-exam Attempts
# Number of subjects requiring re-exam
# Number of repeated attempts


subject_attempts = (
    reexam_records
    .groupby(
        [
            "Student_ID",
            "Subject_ID"
        ]
    )
    .size()
    .reset_index(
        name="Attempts_For_Subject"
    )
)


reexam_summary_records = []


for student_id, group in subject_attempts.groupby(
    "Student_ID"
):

    unique_subjects = len(group)

    total_attempts = int(
        group["Attempts_For_Subject"].sum()
    )

    repeated_attempts = int(
        (
            group["Attempts_For_Subject"] - 1
        )
        .clip(lower=0)
        .sum()
    )

    reexam_summary_records.append(
        {
            "Student_ID": student_id,
            "Reexam_Subjects": unique_subjects,
            "Total_Reexam_Attempts": total_attempts,
            "Repeated_Reexam_Attempts": repeated_attempts
        }
    )


reexam_summary = pd.DataFrame(
    reexam_summary_records
)


# ------------------------------------------------------------
# 9. MERGE INTELLIGENCE FEATURES
# ------------------------------------------------------------

risk_profile = risk_profile.merge(
    trajectory_df,
    on="Student_ID",
    how="left"
)


risk_profile = risk_profile.merge(
    reexam_summary,
    on="Student_ID",
    how="left"
)


reexam_columns = [
    "Reexam_Subjects",
    "Total_Reexam_Attempts",
    "Repeated_Reexam_Attempts"
]


risk_profile[
    reexam_columns
] = risk_profile[
    reexam_columns
].fillna(0)


risk_profile[
    "SGPA_Trend_Slope"
] = risk_profile[
    "SGPA_Trend_Slope"
].fillna(0)


risk_profile[
    "Performance_Trajectory"
] = risk_profile[
    "Performance_Trajectory"
].fillna(
    "Stable"
)


# ------------------------------------------------------------
# 10. RISK COMPONENT FUNCTIONS
# ------------------------------------------------------------


# ------------------------------------------------------------
# PERFORMANCE RISK
# Maximum = 25
#
# Score >= 75 -> almost no risk
# Score <= 40 -> maximum risk
# ------------------------------------------------------------

def calculate_performance_risk(score):

    risk = (
        (75 - score)
        / (75 - 40)
        * 25
    )

    return float(
        np.clip(
            risk,
            0,
            25
        )
    )


# ------------------------------------------------------------
# ATTENDANCE RISK
# Maximum = 20
#
# Attendance >= 80 -> no major risk
# Attendance <= 50 -> maximum risk
# ------------------------------------------------------------

def calculate_attendance_risk(attendance):

    risk = (
        (80 - attendance)
        / (80 - 50)
        * 20
    )

    return float(
        np.clip(
            risk,
            0,
            20
        )
    )


# ------------------------------------------------------------
# BACKLOG RISK
# Maximum = 20
#
# 0 backlog  -> 0
# 1 backlog  -> ~6.67
# 2 backlog  -> ~13.33
# 3+ backlog -> 20
# ------------------------------------------------------------

def calculate_backlog_risk(backlogs):

    risk = (
        backlogs
        / 3
        * 20
    )

    return float(
        np.clip(
            risk,
            0,
            20
        )
    )


# ------------------------------------------------------------
# TREND RISK
# Maximum = 15
# ------------------------------------------------------------

def calculate_trend_risk(slope):

    if slope >= -0.05:
        return 0.0

    elif slope >= -0.20:
        return 9.0

    else:
        return 15.0


# ------------------------------------------------------------
# RE-EXAMINATION RISK
# Maximum = 10
#
# A single re-exam gives low historical risk.
# Multiple subjects and repeated attempts increase risk.
# ------------------------------------------------------------

def calculate_reexam_risk(
    reexam_subjects,
    repeated_attempts
):

    if reexam_subjects == 0:
        return 0.0

    risk = 2.5

    # Additional re-exam subjects
    risk += max(
        reexam_subjects - 1,
        0
    ) * 1.5

    # Repeated attempts have stronger impact
    risk += repeated_attempts * 2.0

    return float(
        min(
            risk,
            10
        )
    )


# ------------------------------------------------------------
# CGPA RISK
# Maximum = 10
#
# CGPA >= 8 -> low/no CGPA risk
# CGPA <= 5 -> maximum CGPA risk
# ------------------------------------------------------------

def calculate_cgpa_risk(cgpa):

    risk = (
        (8 - cgpa)
        / (8 - 5)
        * 10
    )

    return float(
        np.clip(
            risk,
            0,
            10
        )
    )


# ------------------------------------------------------------
# 11. CALCULATE RISK COMPONENTS
# ------------------------------------------------------------

risk_profile[
    "Performance_Risk"
] = risk_profile[
    "Average_Score"
].apply(
    calculate_performance_risk
)


risk_profile[
    "Attendance_Risk"
] = risk_profile[
    "Average_Attendance"
].apply(
    calculate_attendance_risk
)


risk_profile[
    "Backlog_Risk"
] = risk_profile[
    "Final_Active_Backlogs"
].apply(
    calculate_backlog_risk
)


risk_profile[
    "Trend_Risk"
] = risk_profile[
    "SGPA_Trend_Slope"
].apply(
    calculate_trend_risk
)


risk_profile[
    "Reexam_Risk"
] = risk_profile.apply(
    lambda row: calculate_reexam_risk(
        row["Reexam_Subjects"],
        row["Repeated_Reexam_Attempts"]
    ),
    axis=1
)


risk_profile[
    "CGPA_Risk"
] = risk_profile[
    "Final_CGPA"
].apply(
    calculate_cgpa_risk
)


# ------------------------------------------------------------
# 12. TOTAL ACADEMIC RISK INDEX
# ------------------------------------------------------------

risk_components = [
    "Performance_Risk",
    "Attendance_Risk",
    "Backlog_Risk",
    "Trend_Risk",
    "Reexam_Risk",
    "CGPA_Risk"
]


risk_profile[
    "Academic_Risk_Index"
] = risk_profile[
    risk_components
].sum(
    axis=1
)


risk_profile[
    "Academic_Risk_Index"
] = risk_profile[
    "Academic_Risk_Index"
].clip(
    0,
    100
).round(
    2
)


# Round component values

for column in risk_components:

    risk_profile[
        column
    ] = risk_profile[
        column
    ].round(
        2
    )


risk_profile[
    "SGPA_Trend_Slope"
] = risk_profile[
    "SGPA_Trend_Slope"
].round(
    4
)


# ------------------------------------------------------------
# 13. RISK CATEGORY
# ------------------------------------------------------------

def classify_risk(score):

    if score < 30:
        return "Low Risk"

    elif score < 60:
        return "Moderate Risk"

    elif score < 80:
        return "High Risk"

    else:
        return "Critical Risk"


risk_profile[
    "Risk_Category"
] = risk_profile[
    "Academic_Risk_Index"
].apply(
    classify_risk
)


# ------------------------------------------------------------
# 14. RISK PRIORITY NUMBER
# ------------------------------------------------------------
# Makes later sorting easier in Power BI.

priority_map = {
    "Low Risk": 1,
    "Moderate Risk": 2,
    "High Risk": 3,
    "Critical Risk": 4
}


risk_profile[
    "Risk_Priority"
] = risk_profile[
    "Risk_Category"
].map(
    priority_map
)


# ------------------------------------------------------------
# 15. SORT OUTPUT
# ------------------------------------------------------------

risk_profile = risk_profile.sort_values(
    by=[
        "Risk_Priority",
        "Academic_Risk_Index"
    ],
    ascending=[
        False,
        False
    ]
).reset_index(
    drop=True
)


# ------------------------------------------------------------
# 16. RISK DISTRIBUTION
# ------------------------------------------------------------

risk_order = [
    "Low Risk",
    "Moderate Risk",
    "High Risk",
    "Critical Risk"
]


risk_distribution = (
    risk_profile[
        "Risk_Category"
    ]
    .value_counts()
    .reindex(
        risk_order,
        fill_value=0
    )
    .reset_index()
)


risk_distribution.columns = [
    "Risk_Category",
    "Student_Count"
]


risk_distribution[
    "Percentage"
] = (
    risk_distribution[
        "Student_Count"
    ]
    / len(
        risk_profile
    )
    * 100
).round(
    2
)


# ------------------------------------------------------------
# 17. DEPARTMENT RISK SUMMARY
# ------------------------------------------------------------

department_risk = (
    risk_profile
    .groupby(
        "Department",
        as_index=False
    )
    .agg(
        Students=(
            "Student_ID",
            "nunique"
        ),
        Average_Risk_Index=(
            "Academic_Risk_Index",
            "mean"
        ),
        Average_Score=(
            "Average_Score",
            "mean"
        ),
        Average_Attendance=(
            "Average_Attendance",
            "mean"
        ),
        Average_Final_CGPA=(
            "Final_CGPA",
            "mean"
        ),
        Total_Active_Backlogs=(
            "Final_Active_Backlogs",
            "sum"
        )
    )
)


department_risk[
    "High_or_Critical_Students"
] = department_risk[
    "Department"
].map(
    risk_profile[
        risk_profile[
            "Risk_Category"
        ].isin(
            [
                "High Risk",
                "Critical Risk"
            ]
        )
    ]
    .groupby(
        "Department"
    )[
        "Student_ID"
    ]
    .nunique()
).fillna(
    0
).astype(
    int
)


department_risk[
    "High_or_Critical_Percentage"
] = (
    department_risk[
        "High_or_Critical_Students"
    ]
    / department_risk[
        "Students"
    ]
    * 100
).round(
    2
)


department_risk[
    "Average_Risk_Index"
] = department_risk[
    "Average_Risk_Index"
].round(
    2
)


department_risk[
    "Average_Score"
] = department_risk[
    "Average_Score"
].round(
    2
)


department_risk[
    "Average_Attendance"
] = department_risk[
    "Average_Attendance"
].round(
    2
)


department_risk[
    "Average_Final_CGPA"
] = department_risk[
    "Average_Final_CGPA"
].round(
    2
)


# ------------------------------------------------------------
# 18. METHODOLOGY TABLE
# ------------------------------------------------------------

methodology = pd.DataFrame(
    [
        {
            "Risk_Component": "Performance",
            "Maximum_Weight": 25,
            "Description":
            "Risk increases as average score falls below 75; "
            "maximum risk at 40 or below."
        },
        {
            "Risk_Component": "Attendance",
            "Maximum_Weight": 20,
            "Description":
            "Risk increases as average attendance falls below 80%; "
            "maximum risk at 50% or below."
        },
        {
            "Risk_Component": "Active Backlogs",
            "Maximum_Weight": 20,
            "Description":
            "Risk increases with unresolved active backlogs; "
            "three or more reaches maximum contribution."
        },
        {
            "Risk_Component": "Performance Trend",
            "Maximum_Weight": 15,
            "Description":
            "Semester-wise SGPA slope identifies improving, stable, "
            "declining and strongly declining performance."
        },
        {
            "Risk_Component": "Re-examination History",
            "Maximum_Weight": 10,
            "Description":
            "Multiple re-exam subjects and repeated attempts increase "
            "historical academic-risk contribution."
        },
        {
            "Risk_Component": "CGPA",
            "Maximum_Weight": 10,
            "Description":
            "Risk increases as final CGPA falls below 8.0; "
            "maximum contribution at 5.0 or below."
        }
    ]
)


# ------------------------------------------------------------
# 19. VALIDATION
# ------------------------------------------------------------

print("\nValidating Academic Risk Index...")


if len(risk_profile) != 500:
    print(
        f"WARNING: Expected 500 students, "
        f"but found {len(risk_profile)}"
    )


if risk_profile[
    "Academic_Risk_Index"
].isna().any():

    raise ValueError(
        "Validation failed: "
        "Academic_Risk_Index contains missing values."
    )


if (
    risk_profile[
        "Academic_Risk_Index"
    ] < 0
).any():

    raise ValueError(
        "Validation failed: "
        "Risk score below 0 detected."
    )


if (
    risk_profile[
        "Academic_Risk_Index"
    ] > 100
).any():

    raise ValueError(
        "Validation failed: "
        "Risk score above 100 detected."
    )


print("Validation passed.")


# ------------------------------------------------------------
# 20. SAVE RESULTS
# ------------------------------------------------------------

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    risk_profile.to_excel(
        writer,
        sheet_name="Student_Risk_Profile",
        index=False
    )

    risk_distribution.to_excel(
        writer,
        sheet_name="Risk_Distribution",
        index=False
    )

    department_risk.to_excel(
        writer,
        sheet_name="Department_Risk",
        index=False
    )

    methodology.to_excel(
        writer,
        sheet_name="Risk_Methodology",
        index=False
    )


# ------------------------------------------------------------
# 21. TERMINAL SUMMARY
# ------------------------------------------------------------

print("\n" + "=" * 65)
print("ACADEMIC RISK INDEX COMPLETED")
print("=" * 65)

print(
    f"\nStudents analysed: "
    f"{len(risk_profile)}"
)

print(
    f"Average Risk Index: "
    f"{risk_profile['Academic_Risk_Index'].mean():.2f}"
)

print(
    f"Lowest Risk Index: "
    f"{risk_profile['Academic_Risk_Index'].min():.2f}"
)

print(
    f"Highest Risk Index: "
    f"{risk_profile['Academic_Risk_Index'].max():.2f}"
)


print("\nRisk Distribution:")

for _, row in risk_distribution.iterrows():

    print(
        f"{row['Risk_Category']}: "
        f"{int(row['Student_Count'])} students "
        f"({row['Percentage']:.2f}%)"
    )


print("\nTop 10 Highest-Risk Students:")

print(
    risk_profile[
        [
            "Student_ID",
            "Department",
            "Academic_Risk_Index",
            "Risk_Category",
            "Performance_Trajectory",
            "Final_Active_Backlogs"
        ]
    ]
    .head(10)
    .to_string(
        index=False
    )
)


print(
    f"\nOutput saved successfully to:\n"
    f"{OUTPUT_FILE}"
)

print("\nSTEP 16 COMPLETED SUCCESSFULLY")
print("=" * 65)