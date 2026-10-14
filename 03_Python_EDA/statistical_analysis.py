import pandas as pd
import numpy as np
from scipy import stats
from pathlib import Path

# =========================================================
# FILE LOCATIONS
# =========================================================

cleaned_file = Path(
    r"C:\Users\mrfqu\Desktop\University_Academic_Analytics\02_Data_Cleaning\Cleaned_Academic_Data.xlsx"
)

engineered_file = Path(
    r"C:\Users\mrfqu\Desktop\University_Academic_Analytics\03_Python_EDA\Engineered_Academic_Data.xlsx"
)

output_file = Path(
    r"C:\Users\mrfqu\Desktop\University_Academic_Analytics\03_Python_EDA\Statistical_Analysis.xlsx"
)

if not cleaned_file.exists():
    raise FileNotFoundError(
        f"Cleaned dataset not found: {cleaned_file}"
    )

if not engineered_file.exists():
    raise FileNotFoundError(
        f"Engineered dataset not found: {engineered_file}"
    )

print()
print("=" * 70)
print("UNIVERSITY ACADEMIC PERFORMANCE ANALYTICS")
print("STATISTICAL ANALYSIS")
print("=" * 70)

# =========================================================
# 1. LOAD DATA
# =========================================================

master = pd.read_excel(
    cleaned_file,
    sheet_name="Analytics_Master"
)

student_features = pd.read_excel(
    engineered_file,
    sheet_name="Student_Features"
)

master_features = pd.read_excel(
    engineered_file,
    sheet_name="Master_Features"
)

print()
print("Datasets loaded successfully.")
print("Academic Records:", len(master))
print("Students:", len(student_features))

# =========================================================
# 2. DESCRIPTIVE STATISTICS
# =========================================================

descriptive_statistics = pd.DataFrame({
    "Variable": [
        "Total Marks",
        "Attendance Percentage",
        "Internal Marks",
        "External Marks",
        "Assignment Score"
    ],

    "Mean": [
        master["Total_Marks"].mean(),
        master["Attendance_Percentage"].mean(),
        master["Internal_Marks"].mean(),
        master["External_Marks"].mean(),
        master["Assignment_Score"].mean()
    ],

    "Median": [
        master["Total_Marks"].median(),
        master["Attendance_Percentage"].median(),
        master["Internal_Marks"].median(),
        master["External_Marks"].median(),
        master["Assignment_Score"].median()
    ],

    "Standard_Deviation": [
        master["Total_Marks"].std(),
        master["Attendance_Percentage"].std(),
        master["Internal_Marks"].std(),
        master["External_Marks"].std(),
        master["Assignment_Score"].std()
    ],

    "Minimum": [
        master["Total_Marks"].min(),
        master["Attendance_Percentage"].min(),
        master["Internal_Marks"].min(),
        master["External_Marks"].min(),
        master["Assignment_Score"].min()
    ],

    "Maximum": [
        master["Total_Marks"].max(),
        master["Attendance_Percentage"].max(),
        master["Internal_Marks"].max(),
        master["External_Marks"].max(),
        master["Assignment_Score"].max()
    ]
})

numeric_columns = [
    "Mean",
    "Median",
    "Standard_Deviation",
    "Minimum",
    "Maximum"
]

descriptive_statistics[numeric_columns] = (
    descriptive_statistics[numeric_columns]
    .round(2)
)

print()
print("=" * 70)
print("1. DESCRIPTIVE STATISTICS")
print("=" * 70)

print(
    descriptive_statistics.to_string(
        index=False
    )
)

# =========================================================
# 3. CORRELATION MATRIX
# =========================================================

correlation_columns = [
    "Total_Marks",
    "Attendance_Percentage",
    "Internal_Marks",
    "External_Marks",
    "Assignment_Score"
]

correlation_matrix = (
    master[
        correlation_columns
    ]
    .corr()
    .round(3)
)

print()
print("=" * 70)
print("2. CORRELATION MATRIX")
print("=" * 70)

print(correlation_matrix)

# =========================================================
# 4. PEARSON CORRELATION
#    ATTENDANCE VS TOTAL MARKS
# =========================================================

pearson_r, pearson_p = stats.pearsonr(
    master["Attendance_Percentage"],
    master["Total_Marks"]
)

pearson_r = float(pearson_r)
pearson_p = float(pearson_p)

r_squared = pearson_r ** 2

if abs(pearson_r) >= 0.70:
    correlation_strength = "Strong"

elif abs(pearson_r) >= 0.40:
    correlation_strength = "Moderate"

elif abs(pearson_r) >= 0.20:
    correlation_strength = "Weak"

else:
    correlation_strength = "Very Weak"

if pearson_r > 0:
    correlation_direction = "Positive"
elif pearson_r < 0:
    correlation_direction = "Negative"
else:
    correlation_direction = "No"

print()
print("=" * 70)
print("3. ATTENDANCE VS MARKS - PEARSON CORRELATION")
print("=" * 70)

print(
    "Pearson Correlation (r):",
    round(pearson_r, 3)
)

print(
    "R-squared:",
    round(r_squared, 3)
)

print(
    "P-value:",
    f"{pearson_p:.6g}"
)

print(
    "Relationship:",
    correlation_strength,
    correlation_direction
)

# =========================================================
# 5. DEPARTMENT DESCRIPTIVE STATISTICS
# =========================================================

department_statistics = (
    master
    .groupby(
        "Department",
        as_index=False
    )
    .agg(
        Student_Records=(
            "Record_ID",
            "count"
        ),

        Mean_Score=(
            "Total_Marks",
            "mean"
        ),

        Median_Score=(
            "Total_Marks",
            "median"
        ),

        Score_Std_Dev=(
            "Total_Marks",
            "std"
        ),

        Mean_Attendance=(
            "Attendance_Percentage",
            "mean"
        )
    )
)

department_statistics[
    [
        "Mean_Score",
        "Median_Score",
        "Score_Std_Dev",
        "Mean_Attendance"
    ]
] = (
    department_statistics[
        [
            "Mean_Score",
            "Median_Score",
            "Score_Std_Dev",
            "Mean_Attendance"
        ]
    ]
    .round(2)
)

# =========================================================
# 6. ONE-WAY ANOVA - DEPARTMENT SCORES
# =========================================================

department_groups = [
    group["Total_Marks"].values
    for _, group
    in master.groupby("Department")
]

department_f, department_p = stats.f_oneway(
    *department_groups
)

department_f = float(department_f)
department_p = float(department_p)

overall_mean = (
    master["Total_Marks"].mean()
)

ss_between_department = sum(
    len(group) *
    (
        np.mean(group)
        -
        overall_mean
    ) ** 2
    for group
    in department_groups
)

ss_total = sum(
    (
        master["Total_Marks"]
        -
        overall_mean
    ) ** 2
)

department_eta_squared = (
    ss_between_department
    /
    ss_total
)

print()
print("=" * 70)
print("4. DEPARTMENT PERFORMANCE - ANOVA")
print("=" * 70)

print(
    department_statistics.to_string(
        index=False
    )
)

print()
print(
    "ANOVA F-statistic:",
    round(
        department_f,
        3
    )
)

print(
    "ANOVA P-value:",
    f"{department_p:.6g}"
)

print(
    "Eta Squared:",
    round(
        department_eta_squared,
        4
    )
)

if department_p < 0.05:

    department_test_result = (
        "Statistically significant difference"
    )

else:

    department_test_result = (
        "No statistically significant difference"
    )

print(
    "Result:",
    department_test_result
)

# =========================================================
# 7. SEMESTER DESCRIPTIVE STATISTICS
# =========================================================

semester_statistics = (
    master
    .groupby(
        "Semester",
        as_index=False
    )
    .agg(
        Records=(
            "Record_ID",
            "count"
        ),

        Mean_Score=(
            "Total_Marks",
            "mean"
        ),

        Median_Score=(
            "Total_Marks",
            "median"
        ),

        Score_Std_Dev=(
            "Total_Marks",
            "std"
        ),

        Mean_Attendance=(
            "Attendance_Percentage",
            "mean"
        )
    )
)

semester_statistics[
    [
        "Mean_Score",
        "Median_Score",
        "Score_Std_Dev",
        "Mean_Attendance"
    ]
] = (
    semester_statistics[
        [
            "Mean_Score",
            "Median_Score",
            "Score_Std_Dev",
            "Mean_Attendance"
        ]
    ]
    .round(2)
)

# =========================================================
# 8. ONE-WAY ANOVA - SEMESTER SCORES
# =========================================================

semester_groups = [
    group["Total_Marks"].values
    for _, group
    in master.groupby("Semester")
]

semester_f, semester_p = stats.f_oneway(
    *semester_groups
)

semester_f = float(semester_f)
semester_p = float(semester_p)

ss_between_semester = sum(
    len(group) *
    (
        np.mean(group)
        -
        overall_mean
    ) ** 2
    for group
    in semester_groups
)

semester_eta_squared = (
    ss_between_semester
    /
    ss_total
)

print()
print("=" * 70)
print("5. SEMESTER PERFORMANCE - ANOVA")
print("=" * 70)

print(
    semester_statistics.to_string(
        index=False
    )
)

print()
print(
    "ANOVA F-statistic:",
    round(
        semester_f,
        3
    )
)

print(
    "ANOVA P-value:",
    f"{semester_p:.6g}"
)

print(
    "Eta Squared:",
    round(
        semester_eta_squared,
        4
    )
)

if semester_p < 0.05:

    semester_test_result = (
        "Statistically significant difference"
    )

else:

    semester_test_result = (
        "No statistically significant difference"
    )

print(
    "Result:",
    semester_test_result
)

# =========================================================
# 9. ATTENDANCE CATEGORY VS RESULT
#    CHI-SQUARE TEST
# =========================================================

attendance_result_table = pd.crosstab(
    master_features[
        "Attendance_Category"
    ],
    master_features[
        "Result"
    ]
)

chi2, chi_p, chi_dof, expected = (
    stats.chi2_contingency(
        attendance_result_table
    )
)

chi2 = float(chi2)
chi_p = float(chi_p)

total_observations = (
    attendance_result_table
    .to_numpy()
    .sum()
)

min_dimension = min(
    attendance_result_table.shape[0] - 1,
    attendance_result_table.shape[1] - 1
)

cramers_v = np.sqrt(
    chi2
    /
    (
        total_observations
        *
        min_dimension
    )
)

print()
print("=" * 70)
print("6. ATTENDANCE CATEGORY VS RESULT - CHI-SQUARE")
print("=" * 70)

print(attendance_result_table)

print()
print(
    "Chi-square statistic:",
    round(
        chi2,
        3
    )
)

print(
    "Degrees of Freedom:",
    chi_dof
)

print(
    "P-value:",
    f"{chi_p:.6g}"
)

print(
    "Cramer's V:",
    round(
        cramers_v,
        3
    )
)

if chi_p < 0.05:

    chi_result = (
        "Statistically significant association"
    )

else:

    chi_result = (
        "No statistically significant association"
    )

print(
    "Result:",
    chi_result
)

# =========================================================
# 10. PERFORMANCE BY ATTENDANCE CATEGORY
# =========================================================

attendance_statistics = (
    master_features
    .groupby(
        "Attendance_Category",
        observed=False,
        as_index=False
    )
    .agg(
        Records=(
            "Record_ID",
            "count"
        ),

        Average_Score=(
            "Total_Marks",
            "mean"
        ),

        Pass_Count=(
            "Pass_Flag",
            "sum"
        )
    )
)

attendance_statistics[
    "Pass_Percentage"
] = (
    attendance_statistics[
        "Pass_Count"
    ]
    /
    attendance_statistics[
        "Records"
    ]
    *
    100
)

attendance_statistics[
    "Average_Score"
] = (
    attendance_statistics[
        "Average_Score"
    ].round(2)
)

attendance_statistics[
    "Pass_Percentage"
] = (
    attendance_statistics[
        "Pass_Percentage"
    ].round(2)
)

print()
print("=" * 70)
print("7. PERFORMANCE BY ATTENDANCE CATEGORY")
print("=" * 70)

print(
    attendance_statistics.to_string(
        index=False
    )
)

# =========================================================
# 11. BACKLOG VS STUDENT PERFORMANCE
# =========================================================

backlog_statistics = (
    student_features
    .groupby(
        "Academic_Status",
        as_index=False
    )
    .agg(
        Students=(
            "Student_ID",
            "count"
        ),

        Average_Score=(
            "Average_Score",
            "mean"
        ),

        Average_CGPA=(
            "Final_CGPA",
            "mean"
        ),

        Average_Attendance=(
            "Average_Attendance",
            "mean"
        )
    )
)

backlog_statistics[
    [
        "Average_Score",
        "Average_CGPA",
        "Average_Attendance"
    ]
] = (
    backlog_statistics[
        [
            "Average_Score",
            "Average_CGPA",
            "Average_Attendance"
        ]
    ]
    .round(2)
)

print()
print("=" * 70)
print("8. ACADEMIC STATUS COMPARISON")
print("=" * 70)

print(
    backlog_statistics.to_string(
        index=False
    )
)

# =========================================================
# 12. CREATE STATISTICAL TEST SUMMARY
# =========================================================

statistical_tests = pd.DataFrame({
    "Test": [
        "Pearson Correlation - Attendance vs Marks",
        "One-Way ANOVA - Department vs Marks",
        "One-Way ANOVA - Semester vs Marks",
        "Chi-Square - Attendance Category vs Result"
    ],

    "Statistic": [
        pearson_r,
        department_f,
        semester_f,
        chi2
    ],

    "P_Value": [
        pearson_p,
        department_p,
        semester_p,
        chi_p
    ],

    "Effect_Size": [
        r_squared,
        department_eta_squared,
        semester_eta_squared,
        cramers_v
    ],

    "Effect_Size_Type": [
        "R-Squared",
        "Eta Squared",
        "Eta Squared",
        "Cramer's V"
    ],

    "Conclusion": [
        (
            correlation_strength
            + " "
            + correlation_direction
            + " relationship"
        ),

        department_test_result,

        semester_test_result,

        chi_result
    ]
})

statistical_tests[
    [
        "Statistic",
        "P_Value",
        "Effect_Size"
    ]
] = (
    statistical_tests[
        [
            "Statistic",
            "P_Value",
            "Effect_Size"
        ]
    ]
    .round(6)
)

# =========================================================
# 13. SAVE STATISTICAL RESULTS
# =========================================================

with pd.ExcelWriter(
    output_file,
    engine="openpyxl"
) as writer:

    descriptive_statistics.to_excel(
        writer,
        sheet_name="Descriptive_Statistics",
        index=False
    )

    correlation_matrix.to_excel(
        writer,
        sheet_name="Correlation_Matrix"
    )

    department_statistics.to_excel(
        writer,
        sheet_name="Department_Statistics",
        index=False
    )

    semester_statistics.to_excel(
        writer,
        sheet_name="Semester_Statistics",
        index=False
    )

    attendance_result_table.to_excel(
        writer,
        sheet_name="Attendance_Result_Table"
    )

    attendance_statistics.to_excel(
        writer,
        sheet_name="Attendance_Statistics",
        index=False
    )

    backlog_statistics.to_excel(
        writer,
        sheet_name="Backlog_Statistics",
        index=False
    )

    statistical_tests.to_excel(
        writer,
        sheet_name="Statistical_Tests",
        index=False
    )

# =========================================================
# 14. FINAL VALIDATION
# =========================================================

print()
print("=" * 70)
print("STATISTICAL ANALYSIS SUMMARY")
print("=" * 70)

print()
print(
    "Attendance-Marks Correlation:",
    round(
        pearson_r,
        3
    )
)

print(
    "Department ANOVA P-value:",
    f"{department_p:.6g}"
)

print(
    "Semester ANOVA P-value:",
    f"{semester_p:.6g}"
)

print(
    "Attendance-Result Chi-square P-value:",
    f"{chi_p:.6g}"
)

print()

print(
    "Statistical analysis saved to:"
)

print(
    output_file
)

print()
print("=" * 70)
print("STATISTICAL ANALYSIS COMPLETED SUCCESSFULLY")
print("=" * 70)

print()
print(
    "Step 9 - Statistical Analysis is complete."
)

print("=" * 70)