import sqlite3
import pandas as pd
from pathlib import Path

# =========================================================
# FILE LOCATIONS
# =========================================================

database_file = Path(
    r"C:\Users\mrfqu\Desktop\University_Academic_Analytics\04_SQL_Analysis\University_Academic_Analytics.db"
)

output_excel = Path(
    r"C:\Users\mrfqu\Desktop\University_Academic_Analytics\04_SQL_Analysis\SQL_Analysis_Results.xlsx"
)

output_sql = Path(
    r"C:\Users\mrfqu\Desktop\University_Academic_Analytics\04_SQL_Analysis\Academic_Analytics_Queries.sql"
)

if not database_file.exists():
    raise FileNotFoundError(
        f"SQLite database not found: {database_file}"
    )

print()
print("=" * 75)
print("UNIVERSITY ACADEMIC PERFORMANCE ANALYTICS")
print("SQL ANALYTICAL QUERY EXECUTION")
print("=" * 75)

# =========================================================
# DATABASE CONNECTION
# =========================================================

connection = sqlite3.connect(
    database_file
)

print()
print("Connected to SQLite database successfully.")

# =========================================================
# SQL QUERY 1 - OVERALL KPIs
# =========================================================

query_overall_kpis = """
SELECT

    COUNT(DISTINCT student_id)
        AS total_students,

    COUNT(DISTINCT subject_id)
        AS total_subjects,

    COUNT(*)
        AS total_academic_records,

    ROUND(
        AVG(total_marks),
        2
    )
        AS average_score,

    ROUND(
        AVG(attendance_percentage),
        2
    )
        AS average_attendance,

    SUM(
        CASE
            WHEN result = 'Pass'
            THEN 1
            ELSE 0
        END
    )
        AS pass_records,

    SUM(
        CASE
            WHEN result = 'Fail'
            THEN 1
            ELSE 0
        END
    )
        AS fail_records,

    ROUND(
        100.0
        *
        SUM(
            CASE
                WHEN result = 'Pass'
                THEN 1
                ELSE 0
            END
        )
        /
        COUNT(*),
        2
    )
        AS pass_percentage

FROM analytics_master;
"""

overall_kpis = pd.read_sql_query(
    query_overall_kpis,
    connection
)

# =========================================================
# SQL QUERY 2 - DEPARTMENT PERFORMANCE
# =========================================================

query_department_performance = """
SELECT

    department,

    COUNT(DISTINCT student_id)
        AS total_students,

    COUNT(*)
        AS academic_records,

    ROUND(
        AVG(total_marks),
        2
    )
        AS average_score,

    ROUND(
        AVG(attendance_percentage),
        2
    )
        AS average_attendance,

    SUM(
        CASE
            WHEN result = 'Pass'
            THEN 1
            ELSE 0
        END
    )
        AS pass_count,

    SUM(
        CASE
            WHEN result = 'Fail'
            THEN 1
            ELSE 0
        END
    )
        AS fail_count,

    ROUND(
        100.0
        *
        SUM(
            CASE
                WHEN result = 'Pass'
                THEN 1
                ELSE 0
            END
        )
        /
        COUNT(*),
        2
    )
        AS pass_percentage

FROM analytics_master

GROUP BY department

ORDER BY
    average_score DESC;
"""

department_performance = pd.read_sql_query(
    query_department_performance,
    connection
)

# =========================================================
# SQL QUERY 3 - DEPARTMENT RANKING
# =========================================================

query_department_ranking = """
WITH department_summary AS
(
    SELECT

        department,

        ROUND(
            AVG(total_marks),
            2
        )
            AS average_score,

        ROUND(
            100.0
            *
            SUM(
                CASE
                    WHEN result = 'Pass'
                    THEN 1
                    ELSE 0
                END
            )
            /
            COUNT(*),
            2
        )
            AS pass_percentage,

        ROUND(
            AVG(attendance_percentage),
            2
        )
            AS average_attendance

    FROM analytics_master

    GROUP BY department
)

SELECT

    ROW_NUMBER() OVER
    (
        ORDER BY
            average_score DESC,
            pass_percentage DESC
    )
        AS department_rank,

    department,

    average_score,

    pass_percentage,

    average_attendance

FROM department_summary

ORDER BY department_rank;
"""

department_ranking = pd.read_sql_query(
    query_department_ranking,
    connection
)

# =========================================================
# SQL QUERY 4 - SEMESTER PERFORMANCE
# =========================================================

query_semester_performance = """
SELECT

    semester,

    COUNT(*)
        AS total_records,

    ROUND(
        AVG(total_marks),
        2
    )
        AS average_score,

    ROUND(
        AVG(attendance_percentage),
        2
    )
        AS average_attendance,

    SUM(
        CASE
            WHEN result = 'Pass'
            THEN 1
            ELSE 0
        END
    )
        AS pass_count,

    ROUND(
        100.0
        *
        SUM(
            CASE
                WHEN result = 'Pass'
                THEN 1
                ELSE 0
            END
        )
        /
        COUNT(*),
        2
    )
        AS pass_percentage

FROM analytics_master

GROUP BY semester

ORDER BY semester;
"""

semester_performance = pd.read_sql_query(
    query_semester_performance,
    connection
)

# =========================================================
# SQL QUERY 5 - SUBJECT PERFORMANCE
# =========================================================

query_subject_performance = """
SELECT

    subject_id,

    subject_name,

    department,

    semester,

    COUNT(DISTINCT student_id)
        AS total_students,

    ROUND(
        AVG(total_marks),
        2
    )
        AS average_score,

    ROUND(
        AVG(attendance_percentage),
        2
    )
        AS average_attendance,

    SUM(
        CASE
            WHEN result = 'Pass'
            THEN 1
            ELSE 0
        END
    )
        AS pass_count,

    SUM(
        CASE
            WHEN result = 'Fail'
            THEN 1
            ELSE 0
        END
    )
        AS fail_count,

    ROUND(
        100.0
        *
        SUM(
            CASE
                WHEN result = 'Pass'
                THEN 1
                ELSE 0
            END
        )
        /
        COUNT(*),
        2
    )
        AS pass_percentage

FROM analytics_master

GROUP BY
    subject_id,
    subject_name,
    department,
    semester

ORDER BY
    average_score DESC;
"""

subject_performance = pd.read_sql_query(
    query_subject_performance,
    connection
)

# =========================================================
# SQL QUERY 6 - TOP 10 SUBJECTS
# =========================================================

query_top_subjects = """
SELECT

    subject_id,

    subject_name,

    department,

    semester,

    ROUND(
        AVG(total_marks),
        2
    )
        AS average_score,

    ROUND(
        100.0
        *
        SUM(
            CASE
                WHEN result = 'Pass'
                THEN 1
                ELSE 0
            END
        )
        /
        COUNT(*),
        2
    )
        AS pass_percentage

FROM analytics_master

GROUP BY
    subject_id,
    subject_name,
    department,
    semester

ORDER BY
    average_score DESC

LIMIT 10;
"""

top_subjects = pd.read_sql_query(
    query_top_subjects,
    connection
)

# =========================================================
# SQL QUERY 7 - LOWEST 10 SUBJECTS
# =========================================================

query_lowest_subjects = """
SELECT

    subject_id,

    subject_name,

    department,

    semester,

    ROUND(
        AVG(total_marks),
        2
    )
        AS average_score,

    ROUND(
        100.0
        *
        SUM(
            CASE
                WHEN result = 'Pass'
                THEN 1
                ELSE 0
            END
        )
        /
        COUNT(*),
        2
    )
        AS pass_percentage

FROM analytics_master

GROUP BY
    subject_id,
    subject_name,
    department,
    semester

ORDER BY
    average_score ASC

LIMIT 10;
"""

lowest_subjects = pd.read_sql_query(
    query_lowest_subjects,
    connection
)

# =========================================================
# SQL QUERY 8 - TOP 10 STUDENTS
# =========================================================

query_top_students = """
SELECT

    ROW_NUMBER() OVER
    (
        ORDER BY
            average_score DESC,
            final_cgpa DESC,
            average_attendance DESC
    )
        AS overall_rank,

    student_id,

    department,

    average_score,

    average_attendance,

    pass_percentage,

    final_cgpa,

    final_active_backlogs

FROM student_summary

ORDER BY
    average_score DESC,
    final_cgpa DESC,
    average_attendance DESC

LIMIT 10;
"""

top_students = pd.read_sql_query(
    query_top_students,
    connection
)

# =========================================================
# SQL QUERY 9 - DEPARTMENT TOPPERS
# =========================================================

query_department_toppers = """
WITH ranked_students AS
(
    SELECT

        student_id,

        department,

        average_score,

        average_attendance,

        pass_percentage,

        final_cgpa,

        ROW_NUMBER() OVER
        (
            PARTITION BY department

            ORDER BY
                average_score DESC,
                final_cgpa DESC,
                average_attendance DESC
        )
            AS department_rank

    FROM student_summary
)

SELECT

    department,

    student_id,

    average_score,

    average_attendance,

    pass_percentage,

    final_cgpa

FROM ranked_students

WHERE department_rank = 1

ORDER BY department;
"""

department_toppers = pd.read_sql_query(
    query_department_toppers,
    connection
)

# =========================================================
# SQL QUERY 10 - GRADE DISTRIBUTION
# =========================================================

query_grade_distribution = """
SELECT

    grade,

    COUNT(*)
        AS grade_count,

    ROUND(
        100.0
        *
        COUNT(*)
        /
        (
            SELECT COUNT(*)
            FROM academic_records
        ),
        2
    )
        AS percentage

FROM academic_records

GROUP BY grade

ORDER BY

    CASE grade

        WHEN 'A++'
            THEN 1

        WHEN 'A+'
            THEN 2

        WHEN 'B++'
            THEN 3

        WHEN 'B+'
            THEN 4

        WHEN 'C'
            THEN 5

        WHEN 'RA'
            THEN 6

    END;
"""

grade_distribution = pd.read_sql_query(
    query_grade_distribution,
    connection
)

# =========================================================
# SQL QUERY 11 - ATTENDANCE GROUP ANALYSIS
# =========================================================

query_attendance_groups = """
WITH attendance_groups AS
(
    SELECT

        record_id,

        total_marks,

        result,

        CASE

            WHEN attendance_percentage < 75
                THEN 'Low Attendance'

            WHEN attendance_percentage < 85
                THEN 'Satisfactory Attendance'

            ELSE
                'High Attendance'

        END
            AS attendance_category

    FROM analytics_master
)

SELECT

    attendance_category,

    COUNT(*)
        AS total_records,

    ROUND(
        AVG(total_marks),
        2
    )
        AS average_score,

    SUM(
        CASE
            WHEN result = 'Pass'
            THEN 1
            ELSE 0
        END
    )
        AS pass_count,

    SUM(
        CASE
            WHEN result = 'Fail'
            THEN 1
            ELSE 0
        END
    )
        AS fail_count,

    ROUND(
        100.0
        *
        SUM(
            CASE
                WHEN result = 'Pass'
                THEN 1
                ELSE 0
            END
        )
        /
        COUNT(*),
        2
    )
        AS pass_percentage

FROM attendance_groups

GROUP BY attendance_category

ORDER BY

    CASE attendance_category

        WHEN 'Low Attendance'
            THEN 1

        WHEN 'Satisfactory Attendance'
            THEN 2

        WHEN 'High Attendance'
            THEN 3

    END;
"""

attendance_groups = pd.read_sql_query(
    query_attendance_groups,
    connection
)

# =========================================================
# SQL QUERY 12 - FINAL BACKLOG STATUS
# =========================================================

query_backlog_status = """
SELECT

    CASE

        WHEN active_backlogs = 0
            THEN 'All Clear'

        ELSE
            'Active Backlog'

    END
        AS academic_status,

    COUNT(*)
        AS total_students,

    ROUND(
        AVG(cgpa),
        2
    )
        AS average_final_cgpa,

    ROUND(
        AVG(active_backlogs),
        2
    )
        AS average_active_backlogs,

    ROUND(
        AVG(cumulative_credits_cleared),
        2
    )
        AS average_credits_cleared

FROM semester_performance

WHERE semester = 8

GROUP BY academic_status

ORDER BY academic_status;
"""

backlog_status = pd.read_sql_query(
    query_backlog_status,
    connection
)

# =========================================================
# SQL QUERY 13 - BACKLOGS BY DEPARTMENT
# =========================================================

query_department_backlogs = """
SELECT

    s.department,

    COUNT(*)
        AS total_students,

    SUM(
        CASE
            WHEN sp.active_backlogs > 0
            THEN 1
            ELSE 0
        END
    )
        AS students_with_backlogs,

    SUM(
        CASE
            WHEN sp.active_backlogs = 0
            THEN 1
            ELSE 0
        END
    )
        AS all_clear_students,

    SUM(
        sp.active_backlogs
    )
        AS total_active_backlogs,

    ROUND(
        AVG(
            sp.active_backlogs
        ),
        2
    )
        AS average_active_backlogs

FROM semester_performance AS sp

INNER JOIN students AS s

    ON sp.student_id
    =
    s.student_id

WHERE sp.semester = 8

GROUP BY s.department

ORDER BY
    students_with_backlogs DESC;
"""

department_backlogs = pd.read_sql_query(
    query_department_backlogs,
    connection
)

# =========================================================
# SQL QUERY 14 - RE-EXAMINATION PERFORMANCE
# =========================================================

query_reexam_performance = """
SELECT

    attempt_number,

    COUNT(*)
        AS total_attempts,

    SUM(
        CASE
            WHEN result = 'Pass'
            THEN 1
            ELSE 0
        END
    )
        AS passed_attempts,

    SUM(
        CASE
            WHEN result = 'Fail'
            THEN 1
            ELSE 0
        END
    )
        AS failed_attempts,

    ROUND(
        100.0
        *
        SUM(
            CASE
                WHEN result = 'Pass'
                THEN 1
                ELSE 0
            END
        )
        /
        COUNT(*),
        2
    )
        AS pass_percentage,

    ROUND(
        AVG(original_marks),
        2
    )
        AS average_original_marks,

    ROUND(
        AVG(reexam_marks),
        2
    )
        AS average_reexam_marks

FROM reexam_records

GROUP BY attempt_number

ORDER BY attempt_number;
"""

reexam_performance = pd.read_sql_query(
    query_reexam_performance,
    connection
)

# =========================================================
# SQL QUERY 15 - RE-EXAMS BY SEMESTER
# =========================================================

query_reexam_semester = """
SELECT

    attempt_semester,

    COUNT(*)
        AS total_attempts,

    SUM(
        CASE
            WHEN result = 'Pass'
            THEN 1
            ELSE 0
        END
    )
        AS passed_attempts,

    SUM(
        CASE
            WHEN result = 'Fail'
            THEN 1
            ELSE 0
        END
    )
        AS failed_attempts,

    ROUND(
        AVG(reexam_marks),
        2
    )
        AS average_reexam_marks

FROM reexam_records

GROUP BY attempt_semester

ORDER BY attempt_semester;
"""

reexam_semester = pd.read_sql_query(
    query_reexam_semester,
    connection
)

# =========================================================
# SQL QUERY 16 - SEMESTER CGPA / BACKLOG PROGRESSION
# =========================================================

query_semester_progression = """
SELECT

    semester,

    ROUND(
        AVG(sgpa),
        2
    )
        AS average_sgpa,

    ROUND(
        AVG(cgpa),
        2
    )
        AS average_cgpa,

    SUM(new_backlogs)
        AS new_backlogs,

    SUM(cleared_backlogs)
        AS cleared_backlogs,

    SUM(active_backlogs)
        AS active_backlogs

FROM semester_performance

GROUP BY semester

ORDER BY semester;
"""

semester_progression = pd.read_sql_query(
    query_semester_progression,
    connection
)

# =========================================================
# PRINT RESULTS
# =========================================================

print()
print("=" * 75)
print("1. OVERALL ACADEMIC KPIs")
print("=" * 75)

print(
    overall_kpis.to_string(
        index=False
    )
)

print()
print("=" * 75)
print("2. DEPARTMENT PERFORMANCE")
print("=" * 75)

print(
    department_performance.to_string(
        index=False
    )
)

print()
print("=" * 75)
print("3. DEPARTMENT RANKING")
print("=" * 75)

print(
    department_ranking.to_string(
        index=False
    )
)

print()
print("=" * 75)
print("4. SEMESTER PERFORMANCE")
print("=" * 75)

print(
    semester_performance.to_string(
        index=False
    )
)

print()
print("=" * 75)
print("5. TOP 10 SUBJECTS")
print("=" * 75)

print(
    top_subjects.to_string(
        index=False
    )
)

print()
print("=" * 75)
print("6. LOWEST 10 SUBJECTS")
print("=" * 75)

print(
    lowest_subjects.to_string(
        index=False
    )
)

print()
print("=" * 75)
print("7. TOP 10 STUDENTS")
print("=" * 75)

print(
    top_students.to_string(
        index=False
    )
)

print()
print("=" * 75)
print("8. DEPARTMENT TOPPERS")
print("=" * 75)

print(
    department_toppers.to_string(
        index=False
    )
)

print()
print("=" * 75)
print("9. ATTENDANCE GROUP ANALYSIS")
print("=" * 75)

print(
    attendance_groups.to_string(
        index=False
    )
)

print()
print("=" * 75)
print("10. FINAL BACKLOG STATUS")
print("=" * 75)

print(
    backlog_status.to_string(
        index=False
    )
)

print()
print("=" * 75)
print("11. DEPARTMENT BACKLOG ANALYSIS")
print("=" * 75)

print(
    department_backlogs.to_string(
        index=False
    )
)

print()
print("=" * 75)
print("12. RE-EXAMINATION PERFORMANCE")
print("=" * 75)

print(
    reexam_performance.to_string(
        index=False
    )
)

# =========================================================
# SAVE ALL SQL RESULTS INTO EXCEL
# =========================================================

with pd.ExcelWriter(
    output_excel,
    engine="openpyxl"
) as writer:

    overall_kpis.to_excel(
        writer,
        sheet_name="Overall_KPIs",
        index=False
    )

    department_performance.to_excel(
        writer,
        sheet_name="Department_Performance",
        index=False
    )

    department_ranking.to_excel(
        writer,
        sheet_name="Department_Ranking",
        index=False
    )

    semester_performance.to_excel(
        writer,
        sheet_name="Semester_Performance",
        index=False
    )

    subject_performance.to_excel(
        writer,
        sheet_name="Subject_Performance",
        index=False
    )

    top_subjects.to_excel(
        writer,
        sheet_name="Top_Subjects",
        index=False
    )

    lowest_subjects.to_excel(
        writer,
        sheet_name="Lowest_Subjects",
        index=False
    )

    top_students.to_excel(
        writer,
        sheet_name="Top_Students",
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

    attendance_groups.to_excel(
        writer,
        sheet_name="Attendance_Analysis",
        index=False
    )

    backlog_status.to_excel(
        writer,
        sheet_name="Backlog_Status",
        index=False
    )

    department_backlogs.to_excel(
        writer,
        sheet_name="Department_Backlogs",
        index=False
    )

    reexam_performance.to_excel(
        writer,
        sheet_name="Reexam_Performance",
        index=False
    )

    reexam_semester.to_excel(
        writer,
        sheet_name="Reexam_By_Semester",
        index=False
    )

    semester_progression.to_excel(
        writer,
        sheet_name="Semester_Progression",
        index=False
    )

# =========================================================
# SAVE PURE SQL QUERIES INTO .SQL FILE
# =========================================================

queries = {
    "01 - Overall Academic KPIs":
        query_overall_kpis,

    "02 - Department Performance":
        query_department_performance,

    "03 - Department Ranking":
        query_department_ranking,

    "04 - Semester Performance":
        query_semester_performance,

    "05 - Subject Performance":
        query_subject_performance,

    "06 - Top 10 Subjects":
        query_top_subjects,

    "07 - Lowest 10 Subjects":
        query_lowest_subjects,

    "08 - Top 10 Students":
        query_top_students,

    "09 - Department Toppers":
        query_department_toppers,

    "10 - Grade Distribution":
        query_grade_distribution,

    "11 - Attendance Group Analysis":
        query_attendance_groups,

    "12 - Final Backlog Status":
        query_backlog_status,

    "13 - Department Backlogs":
        query_department_backlogs,

    "14 - Re-examination Performance":
        query_reexam_performance,

    "15 - Re-examinations by Semester":
        query_reexam_semester,

    "16 - Semester Progression":
        query_semester_progression
}

with open(
    output_sql,
    "w",
    encoding="utf-8"
) as sql_file:

    sql_file.write(
        "-- =====================================================\n"
    )

    sql_file.write(
        "-- UNIVERSITY ACADEMIC PERFORMANCE ANALYTICS\n"
    )

    sql_file.write(
        "-- SQL ANALYTICAL QUERIES\n"
    )

    sql_file.write(
        "-- =====================================================\n\n"
    )

    for title, query in queries.items():

        sql_file.write(
            f"-- {title}\n"
        )

        sql_file.write(
            "-- -----------------------------------------------------\n"
        )

        sql_file.write(
            query.strip()
        )

        sql_file.write(
            "\n\n"
        )

# =========================================================
# CROSS-VALIDATION
# =========================================================

print()
print("=" * 75)
print("SQL RESULT VALIDATION")
print("=" * 75)

validation_errors = []

# Overall known values
overall_row = overall_kpis.iloc[0]

if int(
    overall_row["total_students"]
) != 500:

    validation_errors.append(
        "Total student count incorrect."
    )

if int(
    overall_row["total_subjects"]
) != 160:

    validation_errors.append(
        "Total subject count incorrect."
    )

if int(
    overall_row["total_academic_records"]
) != 16000:

    validation_errors.append(
        "Academic record count incorrect."
    )

if int(
    overall_row["pass_records"]
) != 14130:

    validation_errors.append(
        "Pass record count incorrect."
    )

if int(
    overall_row["fail_records"]
) != 1870:

    validation_errors.append(
        "Fail record count incorrect."
    )

# Top student should match previous Python analysis
if (
    top_students.iloc[0][
        "student_id"
    ]
    !=
    "ST210"
):

    validation_errors.append(
        "Top student does not match Python analysis."
    )

# Final backlog total check
final_backlogs = pd.read_sql_query(
    """
    SELECT

        SUM(active_backlogs)
            AS final_active_backlogs

    FROM semester_performance

    WHERE semester = 8;
    """,
    connection
)

if int(
    final_backlogs.iloc[0][
        "final_active_backlogs"
    ]
) != 342:

    validation_errors.append(
        "Final active backlog count incorrect."
    )

print(
    "Total Students:",
    int(
        overall_row[
            "total_students"
        ]
    )
)

print(
    "Total Subjects:",
    int(
        overall_row[
            "total_subjects"
        ]
    )
)

print(
    "Academic Records:",
    int(
        overall_row[
            "total_academic_records"
        ]
    )
)

print(
    "Pass Percentage:",
    overall_row[
        "pass_percentage"
    ],
    "%"
)

print(
    "Top Student:",
    top_students.iloc[0][
        "student_id"
    ]
)

print(
    "Final Active Backlogs:",
    int(
        final_backlogs.iloc[0][
            "final_active_backlogs"
        ]
    )
)

# =========================================================
# FINAL RESULT
# =========================================================

print()
print("=" * 75)

if len(validation_errors) == 0:

    print(
        "SQL ANALYSIS RESULT: SUCCESS"
    )

    print()
    print(
        "All SQL queries executed and"
    )

    print(
        "cross-validation checks passed."
    )

else:

    print(
        "SQL ANALYSIS RESULT: CHECK REQUIRED"
    )

    print()

    for error in validation_errors:

        print(
            "-",
            error
        )

print()
print(
    "SQL results saved to:"
)

print(
    output_excel
)

print()
print(
    "SQL query file saved to:"
)

print(
    output_sql
)

print("=" * 75)

# =========================================================
# CLOSE DATABASE CONNECTION
# =========================================================

connection.close()