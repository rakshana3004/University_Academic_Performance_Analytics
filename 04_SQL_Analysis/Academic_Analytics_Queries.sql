-- =====================================================
-- UNIVERSITY ACADEMIC PERFORMANCE ANALYTICS
-- SQL ANALYTICAL QUERIES
-- =====================================================

-- 01 - Overall Academic KPIs
-- -----------------------------------------------------
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

-- 02 - Department Performance
-- -----------------------------------------------------
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

-- 03 - Department Ranking
-- -----------------------------------------------------
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

-- 04 - Semester Performance
-- -----------------------------------------------------
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

-- 05 - Subject Performance
-- -----------------------------------------------------
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

-- 06 - Top 10 Subjects
-- -----------------------------------------------------
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

-- 07 - Lowest 10 Subjects
-- -----------------------------------------------------
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

-- 08 - Top 10 Students
-- -----------------------------------------------------
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

-- 09 - Department Toppers
-- -----------------------------------------------------
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

-- 10 - Grade Distribution
-- -----------------------------------------------------
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

-- 11 - Attendance Group Analysis
-- -----------------------------------------------------
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

-- 12 - Final Backlog Status
-- -----------------------------------------------------
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

-- 13 - Department Backlogs
-- -----------------------------------------------------
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

-- 14 - Re-examination Performance
-- -----------------------------------------------------
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

-- 15 - Re-examinations by Semester
-- -----------------------------------------------------
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

-- 16 - Semester Progression
-- -----------------------------------------------------
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

