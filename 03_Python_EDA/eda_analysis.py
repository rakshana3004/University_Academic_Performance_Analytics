import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path

# =========================================================
# FILE LOCATIONS
# =========================================================

input_file = Path(
    r"C:\Users\mrfqu\Desktop\University_Academic_Analytics\02_Data_Cleaning\Cleaned_Academic_Data.xlsx"
)

output_folder = Path(
    r"C:\Users\mrfqu\Desktop\University_Academic_Analytics\03_Python_EDA"
)

output_folder.mkdir(
    parents=True,
    exist_ok=True
)

if not input_file.exists():
    raise FileNotFoundError(
        f"Cleaned dataset not found: {input_file}"
    )

print()
print("=" * 65)
print("UNIVERSITY ACADEMIC PERFORMANCE ANALYTICS - EDA")
print("=" * 65)

# =========================================================
# 1. LOAD CLEANED DATA
# =========================================================

master = pd.read_excel(
    input_file,
    sheet_name="Analytics_Master"
)

student_summary = pd.read_excel(
    input_file,
    sheet_name="Student_Summary"
)

semester_performance = pd.read_excel(
    input_file,
    sheet_name="Semester_Performance"
)

print()
print("Cleaned data loaded successfully.")

print(
    "Analytics Master Records:",
    len(master)
)

print(
    "Student Summary Records:",
    len(student_summary)
)

# =========================================================
# 2. OVERALL PROJECT METRICS
# =========================================================

total_students = (
    master["Student_ID"]
    .nunique()
)

total_subjects = (
    master["Subject_ID"]
    .nunique()
)

average_score = (
    master[
        "Total_Marks"
    ].mean()
)

average_attendance = (
    master[
        "Attendance_Percentage"
    ].mean()
)

pass_count = (
    master[
        "Result"
    ] == "Pass"
).sum()

fail_count = (
    master[
        "Result"
    ] == "Fail"
).sum()

pass_percentage = (
    pass_count
    /
    len(master)
    *
    100
)

fail_percentage = (
    fail_count
    /
    len(master)
    *
    100
)

print()
print("=" * 65)
print("1. OVERALL ACADEMIC PERFORMANCE")
print("=" * 65)

print(
    "Total Students:",
    total_students
)

print(
    "Total Subjects:",
    total_subjects
)

print(
    "Average Score:",
    round(
        average_score,
        2
    )
)

print(
    "Average Attendance:",
    round(
        average_attendance,
        2
    ),
    "%"
)

print(
    "Pass Percentage:",
    round(
        pass_percentage,
        2
    ),
    "%"
)

print(
    "Fail Percentage:",
    round(
        fail_percentage,
        2
    ),
    "%"
)

# =========================================================
# 3. DEPARTMENT PERFORMANCE
# =========================================================

department_performance = (
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
            "Result",
            lambda x:
            (x == "Pass").sum()
        )
    )
)

department_performance[
    "Pass_Percentage"
] = (
    department_performance[
        "Pass_Count"
    ]
    /
    department_performance[
        "Total_Records"
    ]
    *
    100
)

department_performance[
    "Average_Score"
] = (
    department_performance[
        "Average_Score"
    ].round(2)
)

department_performance[
    "Average_Attendance"
] = (
    department_performance[
        "Average_Attendance"
    ].round(2)
)

department_performance[
    "Pass_Percentage"
] = (
    department_performance[
        "Pass_Percentage"
    ].round(2)
)

department_performance = (
    department_performance
    .sort_values(
        "Average_Score",
        ascending=False
    )
)

print()
print("=" * 65)
print("2. DEPARTMENT PERFORMANCE")
print("=" * 65)

print(
    department_performance[
        [
            "Department",
            "Average_Score",
            "Average_Attendance",
            "Pass_Percentage"
        ]
    ].to_string(
        index=False
    )
)

# =========================================================
# 4. SEMESTER-WISE PERFORMANCE
# =========================================================

semester_analysis = (
    master
    .groupby(
        "Semester",
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
            "Result",
            lambda x:
            (x == "Pass").sum()
        )
    )
)

semester_analysis[
    "Pass_Percentage"
] = (
    semester_analysis[
        "Pass_Count"
    ]
    /
    semester_analysis[
        "Total_Records"
    ]
    *
    100
)

semester_analysis[
    "Average_Score"
] = (
    semester_analysis[
        "Average_Score"
    ].round(2)
)

semester_analysis[
    "Average_Attendance"
] = (
    semester_analysis[
        "Average_Attendance"
    ].round(2)
)

semester_analysis[
    "Pass_Percentage"
] = (
    semester_analysis[
        "Pass_Percentage"
    ].round(2)
)

print()
print("=" * 65)
print("3. SEMESTER-WISE PERFORMANCE")
print("=" * 65)

print(
    semester_analysis[
        [
            "Semester",
            "Average_Score",
            "Average_Attendance",
            "Pass_Percentage"
        ]
    ].to_string(
        index=False
    )
)

# =========================================================
# 5. SUBJECT PERFORMANCE
# =========================================================

subject_performance = (
    master
    .groupby(
        [
            "Subject_ID",
            "Subject_Name",
            "Department"
        ],
        as_index=False
    )
    .agg(
        Average_Score=(
            "Total_Marks",
            "mean"
        ),

        Total_Students=(
            "Student_ID",
            "count"
        ),

        Passed_Students=(
            "Result",
            lambda x:
            (x == "Pass").sum()
        )
    )
)

subject_performance[
    "Pass_Percentage"
] = (
    subject_performance[
        "Passed_Students"
    ]
    /
    subject_performance[
        "Total_Students"
    ]
    *
    100
)

subject_performance[
    "Average_Score"
] = (
    subject_performance[
        "Average_Score"
    ].round(2)
)

subject_performance[
    "Pass_Percentage"
] = (
    subject_performance[
        "Pass_Percentage"
    ].round(2)
)

best_subjects = (
    subject_performance
    .sort_values(
        "Average_Score",
        ascending=False
    )
    .head(10)
)

difficult_subjects = (
    subject_performance
    .sort_values(
        "Average_Score",
        ascending=True
    )
    .head(10)
)

print()
print("=" * 65)
print("4. TOP 10 SUBJECTS BY AVERAGE SCORE")
print("=" * 65)

print(
    best_subjects[
        [
            "Subject_ID",
            "Subject_Name",
            "Department",
            "Average_Score",
            "Pass_Percentage"
        ]
    ].to_string(
        index=False
    )
)

print()
print("=" * 65)
print("5. 10 LOWEST-PERFORMING SUBJECTS")
print("=" * 65)

print(
    difficult_subjects[
        [
            "Subject_ID",
            "Subject_Name",
            "Department",
            "Average_Score",
            "Pass_Percentage"
        ]
    ].to_string(
        index=False
    )
)

# =========================================================
# 6. TOP PERFORMERS
# =========================================================

top_performers = (
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
    .head(10)
)

top_performers = (
    top_performers
    .reset_index(
        drop=True
    )
)

top_performers.insert(
    0,
    "Rank",
    range(
        1,
        len(top_performers) + 1
    )
)

print()
print("=" * 65)
print("6. TOP 10 STUDENT PERFORMERS")
print("=" * 65)

print(
    top_performers[
        [
            "Rank",
            "Student_ID",
            "Department",
            "Average_Score",
            "Average_Attendance",
            "Pass_Percentage",
            "Final_CGPA"
        ]
    ].to_string(
        index=False
    )
)

# =========================================================
# 7. DEPARTMENT TOPPERS
# =========================================================

department_toppers = (
    student_summary
    .sort_values(
        [
            "Department",
            "Average_Score",
            "Final_CGPA"
        ],
        ascending=[
            True,
            False,
            False
        ]
    )
    .groupby(
        "Department",
        as_index=False
    )
    .first()
)

print()
print("=" * 65)
print("7. DEPARTMENT TOPPERS")
print("=" * 65)

print(
    department_toppers[
        [
            "Department",
            "Student_ID",
            "Average_Score",
            "Average_Attendance",
            "Final_CGPA"
        ]
    ].to_string(
        index=False
    )
)

# =========================================================
# 8. GRADE DISTRIBUTION
# =========================================================

grade_distribution = (
    master[
        "Grade"
    ]
    .value_counts()
    .reindex(
        [
            "A++",
            "A+",
            "B++",
            "B+",
            "C",
            "RA"
        ],
        fill_value=0
    )
    .reset_index()
)

grade_distribution.columns = [
    "Grade",
    "Count"
]

print()
print("=" * 65)
print("8. GRADE DISTRIBUTION")
print("=" * 65)

print(
    grade_distribution.to_string(
        index=False
    )
)

# =========================================================
# 9. ATTENDANCE VS MARKS CORRELATION
# =========================================================

attendance_correlation = (
    master[
        [
            "Attendance_Percentage",
            "Total_Marks"
        ]
    ]
    .corr()
    .iloc[
        0,
        1
    ]
)

print()
print("=" * 65)
print("9. ATTENDANCE VS MARKS")
print("=" * 65)

print(
    "Correlation between Attendance and Marks:",
    round(
        attendance_correlation,
        3
    )
)

if attendance_correlation >= 0.7:

    correlation_interpretation = (
        "Strong positive relationship"
    )

elif attendance_correlation >= 0.4:

    correlation_interpretation = (
        "Moderate positive relationship"
    )

elif attendance_correlation >= 0.2:

    correlation_interpretation = (
        "Weak positive relationship"
    )

else:

    correlation_interpretation = (
        "Very weak relationship"
    )

print(
    "Interpretation:",
    correlation_interpretation
)

# =========================================================
# 10. ATTENDANCE GROUP ANALYSIS
# =========================================================

master[
    "Attendance_Group"
] = pd.cut(
    master[
        "Attendance_Percentage"
    ],
    bins=[
        0,
        60,
        75,
        85,
        100
    ],
    labels=[
        "Below 60%",
        "60-75%",
        "75-85%",
        "Above 85%"
    ],
    include_lowest=True
)

attendance_group_analysis = (
    master
    .groupby(
        "Attendance_Group",
        observed=False,
        as_index=False
    )
    .agg(
        Average_Score=(
            "Total_Marks",
            "mean"
        ),

        Number_of_Records=(
            "Record_ID",
            "count"
        )
    )
)

attendance_group_analysis[
    "Average_Score"
] = (
    attendance_group_analysis[
        "Average_Score"
    ].round(2)
)

print()
print("=" * 65)
print("10. PERFORMANCE BY ATTENDANCE GROUP")
print("=" * 65)

print(
    attendance_group_analysis.to_string(
        index=False
    )
)

# =========================================================
# 11. BACKLOG ANALYSIS
# =========================================================

final_semester = (
    semester_performance[
        semester_performance[
            "Semester"
        ] == 8
    ]
)

students_no_backlogs = (
    final_semester[
        "Active_Backlogs"
    ] == 0
).sum()

students_with_backlogs = (
    final_semester[
        "Active_Backlogs"
    ] > 0
).sum()

print()
print("=" * 65)
print("11. FINAL BACKLOG STATUS")
print("=" * 65)

print(
    "Students with no active backlogs:",
    int(
        students_no_backlogs
    )
)

print(
    "Students with active backlogs:",
    int(
        students_with_backlogs
    )
)

# =========================================================
# 12. CREATE OUTPUT EXCEL REPORT
# =========================================================

eda_report_file = (
    output_folder
    /
    "EDA_Results.xlsx"
)

overall_metrics = pd.DataFrame(
    {
        "Metric": [
            "Total Students",
            "Total Subjects",
            "Average Score",
            "Average Attendance",
            "Pass Percentage",
            "Fail Percentage",
            "Attendance-Marks Correlation",
            "Students With Active Backlogs"
        ],

        "Value": [
            total_students,
            total_subjects,
            round(
                average_score,
                2
            ),
            round(
                average_attendance,
                2
            ),
            round(
                pass_percentage,
                2
            ),
            round(
                fail_percentage,
                2
            ),
            round(
                attendance_correlation,
                3
            ),
            int(
                students_with_backlogs
            )
        ]
    }
)

with pd.ExcelWriter(
    eda_report_file,
    engine="openpyxl"
) as writer:

    overall_metrics.to_excel(
        writer,
        sheet_name="Overall_Metrics",
        index=False
    )

    department_performance.to_excel(
        writer,
        sheet_name="Department_Performance",
        index=False
    )

    semester_analysis.to_excel(
        writer,
        sheet_name="Semester_Performance",
        index=False
    )

    subject_performance.to_excel(
        writer,
        sheet_name="Subject_Performance",
        index=False
    )

    top_performers.to_excel(
        writer,
        sheet_name="Top_Performers",
        index=False
    )

    department_toppers.to_excel(
        writer,
        sheet_name="Department_Toppers",
        index=False
    )

    grade_distribution.to_excel(
        writer,
        sheet_name="Grade_Distribution",
        index=False
    )

    attendance_group_analysis.to_excel(
        writer,
        sheet_name="Attendance_Analysis",
        index=False
    )

print()
print(
    "EDA results saved to:",
    eda_report_file
)

# =========================================================
# 13. CHART 1 - DEPARTMENT AVERAGE SCORE
# =========================================================

plt.figure(
    figsize=(9, 5)
)

plt.bar(
    department_performance[
        "Department"
    ],
    department_performance[
        "Average_Score"
    ]
)

plt.title(
    "Department-wise Average Academic Score"
)

plt.xlabel(
    "Department"
)

plt.ylabel(
    "Average Score"
)

plt.ylim(
    0,
    100
)

plt.tight_layout()

plt.savefig(
    output_folder
    /
    "01_Department_Average_Score.png",
    dpi=300
)

plt.close()

# =========================================================
# 14. CHART 2 - DEPARTMENT PASS PERCENTAGE
# =========================================================

plt.figure(
    figsize=(9, 5)
)

plt.bar(
    department_performance[
        "Department"
    ],
    department_performance[
        "Pass_Percentage"
    ]
)

plt.title(
    "Department-wise Pass Percentage"
)

plt.xlabel(
    "Department"
)

plt.ylabel(
    "Pass Percentage"
)

plt.ylim(
    0,
    100
)

plt.tight_layout()

plt.savefig(
    output_folder
    /
    "02_Department_Pass_Percentage.png",
    dpi=300
)

plt.close()

# =========================================================
# 15. CHART 3 - SEMESTER PERFORMANCE TREND
# =========================================================

plt.figure(
    figsize=(9, 5)
)

plt.plot(
    semester_analysis[
        "Semester"
    ],
    semester_analysis[
        "Average_Score"
    ],
    marker="o"
)

plt.title(
    "Semester-wise Average Score Trend"
)

plt.xlabel(
    "Semester"
)

plt.ylabel(
    "Average Score"
)

plt.xticks(
    range(
        1,
        9
    )
)

plt.ylim(
    0,
    100
)

plt.tight_layout()

plt.savefig(
    output_folder
    /
    "03_Semester_Performance_Trend.png",
    dpi=300
)

plt.close()

# =========================================================
# 16. CHART 4 - PASS / FAIL DISTRIBUTION
# =========================================================

plt.figure(
    figsize=(7, 6)
)

plt.pie(
    [
        pass_count,
        fail_count
    ],
    labels=[
        "Pass",
        "Fail"
    ],
    autopct="%1.1f%%",
    startangle=90
)

plt.title(
    "Overall Pass and Fail Distribution"
)

plt.tight_layout()

plt.savefig(
    output_folder
    /
    "04_Pass_Fail_Distribution.png",
    dpi=300
)

plt.close()

# =========================================================
# 17. CHART 5 - GRADE DISTRIBUTION
# =========================================================

plt.figure(
    figsize=(9, 5)
)

plt.bar(
    grade_distribution[
        "Grade"
    ],
    grade_distribution[
        "Count"
    ]
)

plt.title(
    "Academic Grade Distribution"
)

plt.xlabel(
    "Grade"
)

plt.ylabel(
    "Number of Records"
)

plt.tight_layout()

plt.savefig(
    output_folder
    /
    "05_Grade_Distribution.png",
    dpi=300
)

plt.close()

# =========================================================
# 18. CHART 6 - ATTENDANCE VS MARKS
# =========================================================

plt.figure(
    figsize=(9, 6)
)

plt.scatter(
    master[
        "Attendance_Percentage"
    ],
    master[
        "Total_Marks"
    ],
    alpha=0.25
)

plt.title(
    "Relationship Between Attendance and Academic Score"
)

plt.xlabel(
    "Attendance Percentage"
)

plt.ylabel(
    "Total Marks"
)

plt.xlim(
    0,
    100
)

plt.ylim(
    0,
    100
)

plt.tight_layout()

plt.savefig(
    output_folder
    /
    "06_Attendance_vs_Marks.png",
    dpi=300
)

plt.close()

# =========================================================
# 19. CHART 7 - TOP 10 PERFORMERS
# =========================================================

top_chart = (
    top_performers
    .sort_values(
        "Average_Score",
        ascending=True
    )
)

plt.figure(
    figsize=(10, 6)
)

plt.barh(
    top_chart[
        "Student_ID"
    ],
    top_chart[
        "Average_Score"
    ]
)

plt.title(
    "Top 10 Student Performers"
)

plt.xlabel(
    "Average Score"
)

plt.ylabel(
    "Student ID"
)

plt.xlim(
    0,
    100
)

plt.tight_layout()

plt.savefig(
    output_folder
    /
    "07_Top_10_Performers.png",
    dpi=300
)

plt.close()

# =========================================================
# 20. CHART 8 - ATTENDANCE GROUP PERFORMANCE
# =========================================================

plt.figure(
    figsize=(9, 5)
)

plt.bar(
    attendance_group_analysis[
        "Attendance_Group"
    ].astype(str),
    attendance_group_analysis[
        "Average_Score"
    ]
)

plt.title(
    "Academic Performance by Attendance Group"
)

plt.xlabel(
    "Attendance Group"
)

plt.ylabel(
    "Average Score"
)

plt.ylim(
    0,
    100
)

plt.tight_layout()

plt.savefig(
    output_folder
    /
    "08_Attendance_Group_Performance.png",
    dpi=300
)

plt.close()

# =========================================================
# FINAL MESSAGE
# =========================================================

print()
print("=" * 65)
print("EDA COMPLETED SUCCESSFULLY")
print("=" * 65)

print()
print(
    "8 analytical charts have been created."
)

print(
    "EDA_Results.xlsx has been created."
)

print()
print(
    "Files are saved inside:"
)

print(
    output_folder
)

print()
print(
    "Step 7 - Exploratory Data Analysis is complete."
)

print("=" * 65)