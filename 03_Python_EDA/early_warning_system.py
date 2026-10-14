from pathlib import Path
import pandas as pd
import numpy as np


# ============================================================
# UNIVERSITY ACADEMIC PERFORMANCE ANALYTICS
# STEP 17 - EARLY WARNING SYSTEM
# + EXPLAINABLE RISK DRIVERS
# CORRECTED DRIVER RANKING VERSION
# ============================================================


# ------------------------------------------------------------
# 1. FILE PATHS
# ------------------------------------------------------------

PROJECT_ROOT = Path(__file__).resolve().parents[1]

INPUT_FILE = (
    PROJECT_ROOT
    / "03_Python_EDA"
    / "Academic_Risk_Index.xlsx"
)

OUTPUT_FILE = (
    PROJECT_ROOT
    / "03_Python_EDA"
    / "Early_Warning_System.xlsx"
)


# ------------------------------------------------------------
# 2. CHECK INPUT FILE
# ------------------------------------------------------------

if not INPUT_FILE.exists():
    raise FileNotFoundError(
        f"\nInput file not found:\n{INPUT_FILE}\n"
        "Run academic_risk_index.py first."
    )


print("=" * 72)
print("STEP 17 - EARLY WARNING SYSTEM + EXPLAINABLE RISK DRIVERS")
print("=" * 72)

print("\nLoading Academic Risk Index...")


# ------------------------------------------------------------
# 3. LOAD STUDENT RISK PROFILE
# ------------------------------------------------------------

risk = pd.read_excel(
    INPUT_FILE,
    sheet_name="Student_Risk_Profile"
)


required_columns = [
    "Student_ID",
    "Department",
    "Average_Score",
    "Average_Attendance",
    "Final_CGPA",
    "Final_Active_Backlogs",
    "Performance_Trajectory",
    "Reexam_Subjects",
    "Repeated_Reexam_Attempts",
    "Performance_Risk",
    "Attendance_Risk",
    "Backlog_Risk",
    "Trend_Risk",
    "Reexam_Risk",
    "CGPA_Risk",
    "Academic_Risk_Index",
    "Risk_Category"
]


missing_columns = [
    col for col in required_columns
    if col not in risk.columns
]


if missing_columns:
    raise ValueError(
        f"\nMissing required columns:\n{missing_columns}"
    )


print(f"Students loaded: {len(risk)}")


# ------------------------------------------------------------
# 4. INDIVIDUAL WARNING FLAGS
# ------------------------------------------------------------

def performance_warning(score):

    if score < 50:
        return "Critical"

    elif score < 60:
        return "Major"

    elif score < 65:
        return "Minor"

    else:
        return "None"


risk["Performance_Warning"] = (
    risk["Average_Score"]
    .apply(performance_warning)
)


def attendance_warning(attendance):

    if attendance < 65:
        return "Critical"

    elif attendance < 75:
        return "Major"

    elif attendance < 80:
        return "Minor"

    else:
        return "None"


risk["Attendance_Warning"] = (
    risk["Average_Attendance"]
    .apply(attendance_warning)
)


def backlog_warning(backlogs):

    if backlogs >= 5:
        return "Critical"

    elif backlogs >= 3:
        return "Major"

    elif backlogs >= 1:
        return "Minor"

    else:
        return "None"


risk["Backlog_Warning"] = (
    risk["Final_Active_Backlogs"]
    .apply(backlog_warning)
)


def trajectory_warning(trajectory):

    if trajectory == "Strongly Declining":
        return "Critical"

    elif trajectory == "Declining":
        return "Major"

    else:
        return "None"


risk["Trajectory_Warning"] = (
    risk["Performance_Trajectory"]
    .apply(trajectory_warning)
)


def reexam_warning(row):

    subjects = row["Reexam_Subjects"]
    repeated = row["Repeated_Reexam_Attempts"]

    if repeated >= 3:
        return "Critical"

    elif repeated >= 2:
        return "Major"

    elif repeated >= 1 or subjects >= 2:
        return "Minor"

    else:
        return "None"


risk["Reexam_Warning"] = (
    risk.apply(
        reexam_warning,
        axis=1
    )
)


def cgpa_warning(cgpa):

    if cgpa < 6:
        return "Critical"

    elif cgpa < 7:
        return "Major"

    elif cgpa < 7.5:
        return "Minor"

    else:
        return "None"


risk["CGPA_Warning"] = (
    risk["Final_CGPA"]
    .apply(cgpa_warning)
)


# ------------------------------------------------------------
# 5. SEVERITY SCORES
# ------------------------------------------------------------

severity_score = {
    "None": 0,
    "Minor": 1,
    "Major": 2,
    "Critical": 3
}


warning_columns = [
    "Performance_Warning",
    "Attendance_Warning",
    "Backlog_Warning",
    "Trajectory_Warning",
    "Reexam_Warning",
    "CGPA_Warning"
]


for col in warning_columns:

    score_col = col.replace(
        "_Warning",
        "_Severity"
    )

    risk[score_col] = (
        risk[col]
        .map(severity_score)
    )


severity_columns = [
    "Performance_Severity",
    "Attendance_Severity",
    "Backlog_Severity",
    "Trajectory_Severity",
    "Reexam_Severity",
    "CGPA_Severity"
]


risk["Warning_Severity_Total"] = (
    risk[severity_columns]
    .sum(axis=1)
)


risk["Active_Warning_Count"] = (
    risk[severity_columns]
    .gt(0)
    .sum(axis=1)
)


risk["Major_or_Critical_Count"] = (
    risk[severity_columns]
    .ge(2)
    .sum(axis=1)
)


risk["Critical_Warning_Count"] = (
    risk[severity_columns]
    .eq(3)
    .sum(axis=1)
)


# ------------------------------------------------------------
# 6. SEVERE CONDITION COUNT
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


# ------------------------------------------------------------
# 7. EARLY WARNING STATUS
# ------------------------------------------------------------

def early_warning_status(row):

    if (
        row["Risk_Category"] == "Critical Risk"
        or row["Severe_Condition_Count"] >= 2
        or row["Final_Active_Backlogs"] >= 5
        or row["Critical_Warning_Count"] >= 2
    ):
        return "Immediate Intervention"

    elif (
        row["Risk_Category"] == "High Risk"
        or row["Severe_Condition_Count"] == 1
        or row["Major_or_Critical_Count"] >= 2
        or row["Academic_Risk_Index"] >= 50
    ):
        return "Attention Required"

    elif (
        row["Risk_Category"] == "Moderate Risk"
        or row["Active_Warning_Count"] >= 1
    ):
        return "Watch"

    else:
        return "Normal"


risk["Early_Warning_Status"] = (
    risk.apply(
        early_warning_status,
        axis=1
    )
)


warning_priority_map = {
    "Normal": 1,
    "Watch": 2,
    "Attention Required": 3,
    "Immediate Intervention": 4
}


risk["Warning_Priority"] = (
    risk["Early_Warning_Status"]
    .map(warning_priority_map)
)


# ------------------------------------------------------------
# 8. HUMAN-READABLE RISK DRIVERS
# ------------------------------------------------------------

def build_risk_drivers(row):

    drivers = []

    if row["Performance_Warning"] == "Critical":
        drivers.append(
            "Very low average academic score"
        )

    elif row["Performance_Warning"] == "Major":
        drivers.append(
            "Low average academic score"
        )

    elif row["Performance_Warning"] == "Minor":
        drivers.append(
            "Average score requires monitoring"
        )


    if row["Attendance_Warning"] == "Critical":
        drivers.append(
            "Very low attendance"
        )

    elif row["Attendance_Warning"] == "Major":
        drivers.append(
            "Low attendance"
        )

    elif row["Attendance_Warning"] == "Minor":
        drivers.append(
            "Attendance below preferred level"
        )


    if row["Backlog_Warning"] == "Critical":
        drivers.append(
            f"Severe unresolved backlog burden "
            f"({int(row['Final_Active_Backlogs'])} active backlogs)"
        )

    elif row["Backlog_Warning"] == "Major":
        drivers.append(
            f"Multiple active backlogs "
            f"({int(row['Final_Active_Backlogs'])})"
        )

    elif row["Backlog_Warning"] == "Minor":
        drivers.append(
            f"Active backlog present "
            f"({int(row['Final_Active_Backlogs'])})"
        )


    if row["Trajectory_Warning"] == "Critical":
        drivers.append(
            "Strongly declining semester performance"
        )

    elif row["Trajectory_Warning"] == "Major":
        drivers.append(
            "Declining semester performance"
        )


    if row["Reexam_Warning"] == "Critical":
        drivers.append(
            "Repeated re-examination difficulty"
        )

    elif row["Reexam_Warning"] == "Major":
        drivers.append(
            "Multiple repeated re-examination attempts"
        )

    elif row["Reexam_Warning"] == "Minor":
        drivers.append(
            "Re-examination history requires monitoring"
        )


    if row["CGPA_Warning"] == "Critical":
        drivers.append(
            "Very low cumulative academic performance"
        )

    elif row["CGPA_Warning"] == "Major":
        drivers.append(
            "Low cumulative GPA"
        )

    elif row["CGPA_Warning"] == "Minor":
        drivers.append(
            "CGPA requires monitoring"
        )


    if len(drivers) == 0:
        return "No significant academic warning detected"

    return " | ".join(drivers)


risk["Explainable_Risk_Drivers"] = (
    risk.apply(
        build_risk_drivers,
        axis=1
    )
)


# ------------------------------------------------------------
# 9. CORRECTED PRIMARY AND SECONDARY RISK DRIVERS
# ------------------------------------------------------------
#
# IMPORTANT CORRECTION:
#
# Drivers are ranked using their ACTUAL WEIGHTED CONTRIBUTION
# to the Academic Risk Index.
#
# Example:
#
# Performance Risk = 12
# Trend Risk       = 3
#
# Performance must be the stronger driver because it adds
# more actual points to the total Academic Risk Index.
# ------------------------------------------------------------

component_mapping = {
    "Performance": "Performance_Risk",
    "Attendance": "Attendance_Risk",
    "Active Backlogs": "Backlog_Risk",
    "Performance Trend": "Trend_Risk",
    "Re-examination History": "Reexam_Risk",
    "CGPA": "CGPA_Risk"
}


def identify_top_drivers(row):

    component_values = {}

    for name, column in component_mapping.items():

        component_values[name] = float(
            row[column]
        )


    sorted_components = sorted(
        component_values.items(),
        key=lambda item: item[1],
        reverse=True
    )


    primary_name = sorted_components[0][0]
    primary_value = sorted_components[0][1]

    secondary_name = sorted_components[1][0]
    secondary_value = sorted_components[1][1]


    # If every risk contribution is zero,
    # there is no meaningful primary driver.

    if primary_value == 0:

        primary_name = "No Significant Driver"
        secondary_name = "No Significant Driver"


    elif secondary_value == 0:

        secondary_name = "No Significant Driver"


    return pd.Series(
        [
            primary_name,
            round(primary_value, 2),
            secondary_name,
            round(secondary_value, 2)
        ]
    )


risk[
    [
        "Primary_Risk_Driver",
        "Primary_Driver_Contribution",
        "Secondary_Risk_Driver",
        "Secondary_Driver_Contribution"
    ]
] = risk.apply(
    identify_top_drivers,
    axis=1
)


# ------------------------------------------------------------
# 10. PRIMARY DRIVER SHARE OF TOTAL RISK
# ------------------------------------------------------------

risk[
    "Primary_Driver_Risk_Share_Percent"
] = np.where(
    risk["Academic_Risk_Index"] > 0,

    (
        risk["Primary_Driver_Contribution"]
        / risk["Academic_Risk_Index"]
        * 100
    ),

    0
)


risk[
    "Primary_Driver_Risk_Share_Percent"
] = (
    risk[
        "Primary_Driver_Risk_Share_Percent"
    ]
    .round(2)
)


# ------------------------------------------------------------
# 11. PRIMARY DRIVER SEVERITY
# ------------------------------------------------------------

def primary_driver_severity(row):

    driver = row[
        "Primary_Risk_Driver"
    ]

    driver_warning_map = {
        "Performance":
            row["Performance_Warning"],

        "Attendance":
            row["Attendance_Warning"],

        "Active Backlogs":
            row["Backlog_Warning"],

        "Performance Trend":
            row["Trajectory_Warning"],

        "Re-examination History":
            row["Reexam_Warning"],

        "CGPA":
            row["CGPA_Warning"],

        "No Significant Driver":
            "None"
    }


    result = driver_warning_map.get(
        driver,
        "None"
    )


    if result == "None":
        return "Low"

    return result


risk["Primary_Driver_Severity"] = (
    risk.apply(
        primary_driver_severity,
        axis=1
    )
)


# ------------------------------------------------------------
# 12. WARNING EXPLANATION
# ------------------------------------------------------------

def warning_explanation(row):

    status = row[
        "Early_Warning_Status"
    ]

    primary = row[
        "Primary_Risk_Driver"
    ]

    secondary = row[
        "Secondary_Risk_Driver"
    ]


    if status == "Normal":

        return (
            "No significant early academic warning "
            "is currently detected."
        )


    elif status == "Watch":

        if secondary == "No Significant Driver":

            return (
                f"Early monitoring is recommended. "
                f"The strongest current risk factor is "
                f"{primary}."
            )

        return (
            f"Early monitoring is recommended. "
            f"The strongest current risk factors are "
            f"{primary} and {secondary}."
        )


    elif status == "Attention Required":

        if secondary == "No Significant Driver":

            return (
                f"Academic attention is required. "
                f"The major contributing factor is "
                f"{primary}."
            )

        return (
            f"Academic attention is required. "
            f"The major contributing factors are "
            f"{primary} and {secondary}."
        )


    else:

        if secondary == "No Significant Driver":

            return (
                f"Immediate academic intervention is recommended. "
                f"The strongest contributing factor is "
                f"{primary}, with "
                f"{int(row['Severe_Condition_Count'])} "
                f"severe condition(s) detected."
            )

        return (
            f"Immediate academic intervention is recommended. "
            f"The strongest contributing factors are "
            f"{primary} and {secondary}, with "
            f"{int(row['Severe_Condition_Count'])} "
            f"severe condition(s) detected."
        )


risk["Warning_Explanation"] = (
    risk.apply(
        warning_explanation,
        axis=1
    )
)


# ------------------------------------------------------------
# 13. SORT STUDENTS BY WARNING PRIORITY
# ------------------------------------------------------------

risk = risk.sort_values(
    by=[
        "Warning_Priority",
        "Severe_Condition_Count",
        "Academic_Risk_Index"
    ],
    ascending=[
        False,
        False,
        False
    ]
).reset_index(
    drop=True
)


# ------------------------------------------------------------
# 14. WARNING DISTRIBUTION
# ------------------------------------------------------------

warning_order = [
    "Normal",
    "Watch",
    "Attention Required",
    "Immediate Intervention"
]


warning_distribution = (
    risk[
        "Early_Warning_Status"
    ]
    .value_counts()
    .reindex(
        warning_order,
        fill_value=0
    )
    .reset_index()
)


warning_distribution.columns = [
    "Early_Warning_Status",
    "Student_Count"
]


warning_distribution[
    "Percentage"
] = (
    warning_distribution[
        "Student_Count"
    ]
    / len(risk)
    * 100
).round(2)


# ------------------------------------------------------------
# 15. PRIMARY DRIVER SUMMARY
# ------------------------------------------------------------

driver_summary = (
    risk[
        "Primary_Risk_Driver"
    ]
    .value_counts()
    .reset_index()
)


driver_summary.columns = [
    "Primary_Risk_Driver",
    "Student_Count"
]


driver_summary[
    "Percentage"
] = (
    driver_summary[
        "Student_Count"
    ]
    / len(risk)
    * 100
).round(2)


# ------------------------------------------------------------
# 16. DEPARTMENT WARNING SUMMARY
# ------------------------------------------------------------

department_summary = (
    risk
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
        Average_Warning_Severity=(
            "Warning_Severity_Total",
            "mean"
        )
    )
)


urgent_counts = (
    risk[
        risk[
            "Early_Warning_Status"
        ].isin(
            [
                "Attention Required",
                "Immediate Intervention"
            ]
        )
    ]
    .groupby(
        "Department"
    )[
        "Student_ID"
    ]
    .nunique()
)


immediate_counts = (
    risk[
        risk[
            "Early_Warning_Status"
        ]
        == "Immediate Intervention"
    ]
    .groupby(
        "Department"
    )[
        "Student_ID"
    ]
    .nunique()
)


department_summary[
    "Attention_or_Immediate"
] = (
    department_summary[
        "Department"
    ]
    .map(urgent_counts)
    .fillna(0)
    .astype(int)
)


department_summary[
    "Immediate_Intervention"
] = (
    department_summary[
        "Department"
    ]
    .map(immediate_counts)
    .fillna(0)
    .astype(int)
)


department_summary[
    "Urgent_Percentage"
] = (
    department_summary[
        "Attention_or_Immediate"
    ]
    / department_summary[
        "Students"
    ]
    * 100
).round(2)


department_summary[
    "Average_Risk_Index"
] = (
    department_summary[
        "Average_Risk_Index"
    ]
    .round(2)
)


department_summary[
    "Average_Warning_Severity"
] = (
    department_summary[
        "Average_Warning_Severity"
    ]
    .round(2)
)


# ------------------------------------------------------------
# 17. TOP PRIORITY STUDENTS
# ------------------------------------------------------------

priority_students = risk[
    risk[
        "Early_Warning_Status"
    ].isin(
        [
            "Attention Required",
            "Immediate Intervention"
        ]
    )
].copy()


priority_columns = [
    "Student_ID",
    "Department",
    "Academic_Risk_Index",
    "Risk_Category",
    "Early_Warning_Status",
    "Average_Score",
    "Average_Attendance",
    "Final_CGPA",
    "Final_Active_Backlogs",
    "Performance_Trajectory",
    "Repeated_Reexam_Attempts",
    "Primary_Risk_Driver",
    "Primary_Driver_Contribution",
    "Primary_Driver_Risk_Share_Percent",
    "Primary_Driver_Severity",
    "Secondary_Risk_Driver",
    "Secondary_Driver_Contribution",
    "Severe_Condition_Count",
    "Explainable_Risk_Drivers",
    "Warning_Explanation"
]


priority_students = priority_students[
    priority_columns
]


# ------------------------------------------------------------
# 18. METHODOLOGY
# ------------------------------------------------------------

methodology = pd.DataFrame(
    [
        {
            "Rule": "Normal",
            "Definition":
            "Low-risk student with no significant warning flags."
        },
        {
            "Rule": "Watch",
            "Definition":
            "Moderate risk or at least one early warning indicator."
        },
        {
            "Rule": "Attention Required",
            "Definition":
            "High risk, one severe condition, multiple major warnings, "
            "or Academic Risk Index of at least 50."
        },
        {
            "Rule": "Immediate Intervention",
            "Definition":
            "Critical risk, at least two severe conditions, five or more "
            "active backlogs, or at least two critical warning indicators."
        },
        {
            "Rule": "Explainable Risk Drivers",
            "Definition":
            "Human-readable reasons generated from performance, attendance, "
            "backlogs, trajectory, re-examination history and CGPA."
        },
        {
            "Rule": "Primary Risk Driver",
            "Definition":
            "The component contributing the largest actual weighted number "
            "of points to the Academic Risk Index."
        },
        {
            "Rule": "Secondary Risk Driver",
            "Definition":
            "The component contributing the second-largest actual weighted "
            "number of points to the Academic Risk Index."
        },
        {
            "Rule": "Primary Driver Risk Share",
            "Definition":
            "Percentage of the student's total Academic Risk Index "
            "contributed by the primary risk driver."
        }
    ]
)


# ------------------------------------------------------------
# 19. VALIDATION
# ------------------------------------------------------------

print("\nValidating Early Warning System...")


if len(risk) != 500:

    print(
        f"WARNING: Expected 500 students, "
        f"found {len(risk)}."
    )


if risk[
    "Early_Warning_Status"
].isna().any():

    raise ValueError(
        "Validation failed: missing warning status."
    )


if risk[
    "Primary_Risk_Driver"
].isna().any():

    raise ValueError(
        "Validation failed: missing primary risk driver."
    )


if (
    warning_distribution[
        "Student_Count"
    ].sum()
    != len(risk)
):

    raise ValueError(
        "Validation failed: warning distribution "
        "does not equal total students."
    )


# Check that primary contribution is never lower
# than secondary contribution.

invalid_driver_order = (
    risk[
        "Primary_Driver_Contribution"
    ]
    <
    risk[
        "Secondary_Driver_Contribution"
    ]
).sum()


if invalid_driver_order > 0:

    raise ValueError(
        "Validation failed: secondary driver exceeds "
        "primary driver contribution."
    )


print("Validation passed.")


# ------------------------------------------------------------
# 20. SAVE OUTPUT
# ------------------------------------------------------------

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    risk.to_excel(
        writer,
        sheet_name="Student_Early_Warning",
        index=False
    )

    warning_distribution.to_excel(
        writer,
        sheet_name="Warning_Distribution",
        index=False
    )

    driver_summary.to_excel(
        writer,
        sheet_name="Primary_Driver_Summary",
        index=False
    )

    department_summary.to_excel(
        writer,
        sheet_name="Department_Warnings",
        index=False
    )

    priority_students.to_excel(
        writer,
        sheet_name="Priority_Students",
        index=False
    )

    methodology.to_excel(
        writer,
        sheet_name="Warning_Methodology",
        index=False
    )


# ------------------------------------------------------------
# 21. TERMINAL SUMMARY
# ------------------------------------------------------------

print("\n" + "=" * 72)
print("EARLY WARNING SYSTEM COMPLETED")
print("=" * 72)

print(
    f"\nStudents analysed: "
    f"{len(risk)}"
)


print("\nEarly Warning Distribution:")

for _, row in warning_distribution.iterrows():

    print(
        f"{row['Early_Warning_Status']}: "
        f"{int(row['Student_Count'])} students "
        f"({row['Percentage']:.2f}%)"
    )


print("\nCorrected Primary Risk Driver Distribution:")

print(
    driver_summary.to_string(
        index=False
    )
)


print(
    "\nStudents requiring Attention "
    "or Immediate Intervention:"
)

print(
    len(priority_students)
)


immediate_count = int(
    (
        risk[
            "Early_Warning_Status"
        ]
        == "Immediate Intervention"
    ).sum()
)


print(
    f"\nImmediate Intervention students: "
    f"{immediate_count}"
)


print("\nTop 10 Highest-Priority Students:")

print(
    risk[
        [
            "Student_ID",
            "Department",
            "Academic_Risk_Index",
            "Risk_Category",
            "Early_Warning_Status",
            "Primary_Risk_Driver",
            "Primary_Driver_Contribution",
            "Secondary_Risk_Driver",
            "Secondary_Driver_Contribution",
            "Severe_Condition_Count"
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

print("\nSTEP 17 COMPLETED SUCCESSFULLY")
print("=" * 72)