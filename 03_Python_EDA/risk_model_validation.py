from pathlib import Path
import pandas as pd
import numpy as np


# ============================================================
# UNIVERSITY ACADEMIC PERFORMANCE ANALYTICS
# STEP 16B - ACADEMIC RISK MODEL VALIDATION
# ============================================================


PROJECT_ROOT = Path(__file__).resolve().parents[1]

INPUT_FILE = (
    PROJECT_ROOT
    / "03_Python_EDA"
    / "Academic_Risk_Index.xlsx"
)

OUTPUT_FILE = (
    PROJECT_ROOT
    / "03_Python_EDA"
    / "Risk_Model_Validation.xlsx"
)


# ------------------------------------------------------------
# 1. CHECK INPUT FILE
# ------------------------------------------------------------

if not INPUT_FILE.exists():
    raise FileNotFoundError(
        f"\nFile not found:\n{INPUT_FILE}\n"
        "Run academic_risk_index.py first."
    )


print("=" * 70)
print("STEP 16B - ACADEMIC RISK MODEL VALIDATION")
print("=" * 70)


# ------------------------------------------------------------
# 2. LOAD RISK PROFILE
# ------------------------------------------------------------

risk = pd.read_excel(
    INPUT_FILE,
    sheet_name="Student_Risk_Profile"
)

print(f"\nStudents loaded: {len(risk)}")


# ------------------------------------------------------------
# 3. CREATE VALIDATION GROUPS
# ------------------------------------------------------------

# Backlog groups
risk["Backlog_Group"] = pd.cut(
    risk["Final_Active_Backlogs"],
    bins=[-1, 0, 1, 2, np.inf],
    labels=[
        "0 Backlogs",
        "1 Backlog",
        "2 Backlogs",
        "3+ Backlogs"
    ]
)


# Score groups
risk["Score_Group"] = pd.cut(
    risk["Average_Score"],
    bins=[-np.inf, 49.99, 59.99, 74.99, np.inf],
    labels=[
        "Below 50",
        "50-59",
        "60-74",
        "75 and Above"
    ]
)


# Attendance groups
risk["Attendance_Group"] = pd.cut(
    risk["Average_Attendance"],
    bins=[-np.inf, 64.99, 74.99, 79.99, np.inf],
    labels=[
        "Below 65",
        "65-74",
        "75-79",
        "80 and Above"
    ]
)


# CGPA groups
risk["CGPA_Group"] = pd.cut(
    risk["Final_CGPA"],
    bins=[-np.inf, 5.99, 6.99, 7.99, np.inf],
    labels=[
        "Below 6",
        "6.00-6.99",
        "7.00-7.99",
        "8.00 and Above"
    ]
)


# ------------------------------------------------------------
# 4. FUNCTION FOR GROUP SUMMARIES
# ------------------------------------------------------------

def group_summary(df, group_column):

    result = (
        df.groupby(
            group_column,
            observed=False
        )
        .agg(
            Students=("Student_ID", "nunique"),
            Average_Risk=("Academic_Risk_Index", "mean"),
            Minimum_Risk=("Academic_Risk_Index", "min"),
            Maximum_Risk=("Academic_Risk_Index", "max")
        )
        .reset_index()
    )

    result["Average_Risk"] = (
        result["Average_Risk"]
        .round(2)
    )

    result["Minimum_Risk"] = (
        result["Minimum_Risk"]
        .round(2)
    )

    result["Maximum_Risk"] = (
        result["Maximum_Risk"]
        .round(2)
    )

    return result


backlog_validation = group_summary(
    risk,
    "Backlog_Group"
)

score_validation = group_summary(
    risk,
    "Score_Group"
)

attendance_validation = group_summary(
    risk,
    "Attendance_Group"
)

cgpa_validation = group_summary(
    risk,
    "CGPA_Group"
)

trajectory_validation = group_summary(
    risk,
    "Performance_Trajectory"
)


# ------------------------------------------------------------
# 5. IDENTIFY SEVERE ACADEMIC CONDITIONS
# ------------------------------------------------------------

risk["Severe_Backlog_Flag"] = (
    risk["Final_Active_Backlogs"] >= 5
).astype(int)


risk["Very_Low_Score_Flag"] = (
    risk["Average_Score"] < 50
).astype(int)


risk["Very_Low_Attendance_Flag"] = (
    risk["Average_Attendance"] < 65
).astype(int)


risk["Low_CGPA_Flag"] = (
    risk["Final_CGPA"] < 6
).astype(int)


risk["Repeated_Reexam_Flag"] = (
    risk["Repeated_Reexam_Attempts"] >= 2
).astype(int)


risk["Strong_Decline_Flag"] = (
    risk["Performance_Trajectory"]
    == "Strongly Declining"
).astype(int)


severe_columns = [
    "Severe_Backlog_Flag",
    "Very_Low_Score_Flag",
    "Very_Low_Attendance_Flag",
    "Low_CGPA_Flag",
    "Repeated_Reexam_Flag",
    "Strong_Decline_Flag"
]


risk["Severe_Condition_Count"] = (
    risk[severe_columns]
    .sum(axis=1)
)


# A student with two or more serious warning conditions
# is considered a severe-case candidate for later early-warning logic.

risk["Severe_Case_Candidate"] = np.where(
    risk["Severe_Condition_Count"] >= 2,
    "Yes",
    "No"
)


# ------------------------------------------------------------
# 6. SEVERE CASE TABLE
# ------------------------------------------------------------

severe_cases = (
    risk[
        risk["Severe_Case_Candidate"] == "Yes"
    ]
    .sort_values(
        [
            "Severe_Condition_Count",
            "Academic_Risk_Index"
        ],
        ascending=[
            False,
            False
        ]
    )
)


severe_case_columns = [
    "Student_ID",
    "Department",
    "Average_Score",
    "Average_Attendance",
    "Final_CGPA",
    "Final_Active_Backlogs",
    "Performance_Trajectory",
    "Repeated_Reexam_Attempts",
    "Academic_Risk_Index",
    "Risk_Category",
    "Severe_Condition_Count",
    "Severe_Backlog_Flag",
    "Very_Low_Score_Flag",
    "Very_Low_Attendance_Flag",
    "Low_CGPA_Flag",
    "Repeated_Reexam_Flag",
    "Strong_Decline_Flag"
]

severe_cases = severe_cases[
    severe_case_columns
]


# ------------------------------------------------------------
# 7. COMPONENT CORRELATION WITH TOTAL RISK
# ------------------------------------------------------------

risk_components = [
    "Performance_Risk",
    "Attendance_Risk",
    "Backlog_Risk",
    "Trend_Risk",
    "Reexam_Risk",
    "CGPA_Risk"
]


correlation_records = []


for component in risk_components:

    correlation = risk[
        component
    ].corr(
        risk["Academic_Risk_Index"]
    )

    correlation_records.append(
        {
            "Risk_Component": component,
            "Correlation_With_Total_Risk":
                round(correlation, 4)
        }
    )


component_correlations = pd.DataFrame(
    correlation_records
)


# ------------------------------------------------------------
# 8. OVERALL VALIDATION SUMMARY
# ------------------------------------------------------------

summary = pd.DataFrame(
    [
        {
            "Metric": "Students Analysed",
            "Value": len(risk)
        },
        {
            "Metric": "Average Risk Index",
            "Value": round(
                risk["Academic_Risk_Index"].mean(),
                2
            )
        },
        {
            "Metric": "Median Risk Index",
            "Value": round(
                risk["Academic_Risk_Index"].median(),
                2
            )
        },
        {
            "Metric": "Maximum Risk Index",
            "Value": round(
                risk["Academic_Risk_Index"].max(),
                2
            )
        },
        {
            "Metric": "Low Risk Students",
            "Value": int(
                (
                    risk["Risk_Category"]
                    == "Low Risk"
                ).sum()
            )
        },
        {
            "Metric": "Moderate Risk Students",
            "Value": int(
                (
                    risk["Risk_Category"]
                    == "Moderate Risk"
                ).sum()
            )
        },
        {
            "Metric": "High Risk Students",
            "Value": int(
                (
                    risk["Risk_Category"]
                    == "High Risk"
                ).sum()
            )
        },
        {
            "Metric": "Critical Risk Students",
            "Value": int(
                (
                    risk["Risk_Category"]
                    == "Critical Risk"
                ).sum()
            )
        },
        {
            "Metric": "Severe Case Candidates",
            "Value": int(
                (
                    risk["Severe_Case_Candidate"]
                    == "Yes"
                ).sum()
            )
        }
    ]
)


# ------------------------------------------------------------
# 9. SAVE VALIDATION WORKBOOK
# ------------------------------------------------------------

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    summary.to_excel(
        writer,
        sheet_name="Validation_Summary",
        index=False
    )

    backlog_validation.to_excel(
        writer,
        sheet_name="Risk_By_Backlogs",
        index=False
    )

    score_validation.to_excel(
        writer,
        sheet_name="Risk_By_Score",
        index=False
    )

    attendance_validation.to_excel(
        writer,
        sheet_name="Risk_By_Attendance",
        index=False
    )

    cgpa_validation.to_excel(
        writer,
        sheet_name="Risk_By_CGPA",
        index=False
    )

    trajectory_validation.to_excel(
        writer,
        sheet_name="Risk_By_Trajectory",
        index=False
    )

    severe_cases.to_excel(
        writer,
        sheet_name="Severe_Case_Candidates",
        index=False
    )

    component_correlations.to_excel(
        writer,
        sheet_name="Component_Correlations",
        index=False
    )


# ------------------------------------------------------------
# 10. TERMINAL OUTPUT
# ------------------------------------------------------------

print("\n" + "=" * 70)
print("RISK BY ACTIVE BACKLOGS")
print("=" * 70)

print(
    backlog_validation.to_string(
        index=False
    )
)


print("\n" + "=" * 70)
print("RISK BY AVERAGE SCORE")
print("=" * 70)

print(
    score_validation.to_string(
        index=False
    )
)


print("\n" + "=" * 70)
print("RISK BY ATTENDANCE")
print("=" * 70)

print(
    attendance_validation.to_string(
        index=False
    )
)


print("\n" + "=" * 70)
print("RISK BY FINAL CGPA")
print("=" * 70)

print(
    cgpa_validation.to_string(
        index=False
    )
)


print("\n" + "=" * 70)
print("RISK BY PERFORMANCE TRAJECTORY")
print("=" * 70)

print(
    trajectory_validation.to_string(
        index=False
    )
)


print("\n" + "=" * 70)
print("SEVERE CASE CHECK")
print("=" * 70)

print(
    f"Students with 2 or more severe conditions: "
    f"{len(severe_cases)}"
)


if len(severe_cases) > 0:

    print(
        "\nTop severe-condition candidates:"
    )

    print(
        severe_cases[
            [
                "Student_ID",
                "Department",
                "Academic_Risk_Index",
                "Risk_Category",
                "Final_Active_Backlogs",
                "Severe_Condition_Count"
            ]
        ]
        .head(10)
        .to_string(
            index=False
        )
    )


print("\n" + "=" * 70)
print("VALIDATION COMPLETE")
print("=" * 70)

print(
    f"\nValidation workbook saved to:\n"
    f"{OUTPUT_FILE}"
)

print(
    "\nDo not change the risk formula yet. "
    "Review these results first."
)