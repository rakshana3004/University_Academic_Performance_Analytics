import pandas as pd
import sqlite3
from pathlib import Path

# =========================================================
# FILE LOCATIONS
# =========================================================

excel_file = Path(
    r"C:\Users\mrfqu\Desktop\University_Academic_Analytics\02_Data_Cleaning\Cleaned_Academic_Data.xlsx"
)

database_file = Path(
    r"C:\Users\mrfqu\Desktop\University_Academic_Analytics\04_SQL_Analysis\University_Academic_Analytics.db"
)

if not excel_file.exists():
    raise FileNotFoundError(
        f"Cleaned Excel dataset not found: {excel_file}"
    )

# Remove old database if one already exists
# so that this script always creates a fresh database.

if database_file.exists():
    database_file.unlink()

print()
print("=" * 70)
print("UNIVERSITY ACADEMIC ANALYTICS - SQL DATABASE CREATION")
print("=" * 70)

# =========================================================
# 1. LOAD CLEANED EXCEL DATA
# =========================================================

students = pd.read_excel(
    excel_file,
    sheet_name="Students"
)

subjects = pd.read_excel(
    excel_file,
    sheet_name="Subjects"
)

academic = pd.read_excel(
    excel_file,
    sheet_name="Academic_Records"
)

attendance = pd.read_excel(
    excel_file,
    sheet_name="Attendance"
)

reexam = pd.read_excel(
    excel_file,
    sheet_name="Reexam_Records"
)

semester = pd.read_excel(
    excel_file,
    sheet_name="Semester_Performance"
)

print()
print("Excel data loaded successfully.")

print(
    "Students:",
    len(students)
)

print(
    "Subjects:",
    len(subjects)
)

print(
    "Academic Records:",
    len(academic)
)

print(
    "Attendance Records:",
    len(attendance)
)

print(
    "Re-exam Records:",
    len(reexam)
)

print(
    "Semester Performance:",
    len(semester)
)

# =========================================================
# 2. STANDARDIZE COLUMN NAMES FOR SQL
# =========================================================

def prepare_for_sql(df):

    result = df.copy()

    result.columns = [
        column.strip().lower()
        for column in result.columns
    ]

    return result


students = prepare_for_sql(
    students
)

subjects = prepare_for_sql(
    subjects
)

academic = prepare_for_sql(
    academic
)

attendance = prepare_for_sql(
    attendance
)

reexam = prepare_for_sql(
    reexam
)

semester = prepare_for_sql(
    semester
)

# =========================================================
# 3. CONNECT TO SQLITE DATABASE
# =========================================================

connection = sqlite3.connect(
    database_file
)

connection.execute(
    "PRAGMA foreign_keys = ON;"
)

cursor = connection.cursor()

print()
print("SQLite database connection created.")

# =========================================================
# 4. CREATE STUDENTS TABLE
# =========================================================

cursor.execute(
    """
    CREATE TABLE students
    (
        student_id TEXT PRIMARY KEY,

        department TEXT NOT NULL,

        admission_year INTEGER NOT NULL
    );
    """
)

# =========================================================
# 5. CREATE SUBJECTS TABLE
# =========================================================

cursor.execute(
    """
    CREATE TABLE subjects
    (
        subject_id TEXT PRIMARY KEY,

        subject_name TEXT NOT NULL,

        department TEXT NOT NULL,

        semester INTEGER NOT NULL
            CHECK (
                semester BETWEEN 1 AND 8
            ),

        course_type TEXT NOT NULL,

        credits INTEGER NOT NULL
            CHECK (
                credits IN (3, 4)
            )
    );
    """
)

# =========================================================
# 6. CREATE ACADEMIC RECORDS TABLE
# =========================================================

cursor.execute(
    """
    CREATE TABLE academic_records
    (
        record_id TEXT PRIMARY KEY,

        student_id TEXT NOT NULL,

        subject_id TEXT NOT NULL,

        internal_marks INTEGER NOT NULL
            CHECK (
                internal_marks BETWEEN 0 AND 30
            ),

        external_marks INTEGER NOT NULL
            CHECK (
                external_marks BETWEEN 0 AND 60
            ),

        assignment_score INTEGER NOT NULL
            CHECK (
                assignment_score BETWEEN 0 AND 10
            ),

        total_marks INTEGER NOT NULL
            CHECK (
                total_marks BETWEEN 0 AND 100
            ),

        grade TEXT NOT NULL
            CHECK (
                grade IN (
                    'A++',
                    'A+',
                    'B++',
                    'B+',
                    'C',
                    'RA'
                )
            ),

        result TEXT NOT NULL
            CHECK (
                result IN (
                    'Pass',
                    'Fail'
                )
            ),

        FOREIGN KEY (
            student_id
        )
        REFERENCES students (
            student_id
        ),

        FOREIGN KEY (
            subject_id
        )
        REFERENCES subjects (
            subject_id
        ),

        UNIQUE (
            student_id,
            subject_id
        )
    );
    """
)

# =========================================================
# 7. CREATE ATTENDANCE TABLE
# =========================================================

cursor.execute(
    """
    CREATE TABLE attendance
    (
        student_id TEXT NOT NULL,

        subject_id TEXT NOT NULL,

        classes_held INTEGER NOT NULL
            CHECK (
                classes_held > 0
            ),

        classes_attended INTEGER NOT NULL
            CHECK (
                classes_attended >= 0
            ),

        attendance_percentage REAL NOT NULL
            CHECK (
                attendance_percentage
                BETWEEN 0 AND 100
            ),

        PRIMARY KEY (
            student_id,
            subject_id
        ),

        FOREIGN KEY (
            student_id
        )
        REFERENCES students (
            student_id
        ),

        FOREIGN KEY (
            subject_id
        )
        REFERENCES subjects (
            subject_id
        )
    );
    """
)

# =========================================================
# 8. CREATE RE-EXAMINATION TABLE
# =========================================================

cursor.execute(
    """
    CREATE TABLE reexam_records
    (
        reexam_id TEXT PRIMARY KEY,

        student_id TEXT NOT NULL,

        subject_id TEXT NOT NULL,

        original_semester INTEGER NOT NULL
            CHECK (
                original_semester
                BETWEEN 1 AND 8
            ),

        attempt_semester INTEGER NOT NULL
            CHECK (
                attempt_semester
                BETWEEN 1 AND 8
            ),

        attempt_number INTEGER NOT NULL
            CHECK (
                attempt_number
                BETWEEN 1 AND 3
            ),

        original_marks INTEGER NOT NULL
            CHECK (
                original_marks
                BETWEEN 0 AND 100
            ),

        reexam_marks INTEGER NOT NULL
            CHECK (
                reexam_marks
                BETWEEN 0 AND 100
            ),

        grade TEXT NOT NULL
            CHECK (
                grade IN (
                    'A++',
                    'A+',
                    'B++',
                    'B+',
                    'C',
                    'RA'
                )
            ),

        result TEXT NOT NULL
            CHECK (
                result IN (
                    'Pass',
                    'Fail'
                )
            ),

        CHECK (
            attempt_semester
            >
            original_semester
        ),

        FOREIGN KEY (
            student_id
        )
        REFERENCES students (
            student_id
        ),

        FOREIGN KEY (
            subject_id
        )
        REFERENCES subjects (
            subject_id
        )
    );
    """
)

# =========================================================
# 9. CREATE SEMESTER PERFORMANCE TABLE
# =========================================================

cursor.execute(
    """
    CREATE TABLE semester_performance
    (
        student_id TEXT NOT NULL,

        semester INTEGER NOT NULL
            CHECK (
                semester BETWEEN 1 AND 8
            ),

        sgpa REAL NOT NULL
            CHECK (
                sgpa BETWEEN 0 AND 10
            ),

        cgpa REAL NOT NULL
            CHECK (
                cgpa BETWEEN 0 AND 10
            ),

        new_backlogs INTEGER NOT NULL
            CHECK (
                new_backlogs >= 0
            ),

        cleared_backlogs INTEGER NOT NULL
            CHECK (
                cleared_backlogs >= 0
            ),

        active_backlogs INTEGER NOT NULL
            CHECK (
                active_backlogs >= 0
            ),

        semester_credits_cleared INTEGER
            NOT NULL
            CHECK (
                semester_credits_cleared >= 0
            ),

        cumulative_credits_cleared INTEGER
            NOT NULL
            CHECK (
                cumulative_credits_cleared
                BETWEEN 0 AND 112
            ),

        PRIMARY KEY (
            student_id,
            semester
        ),

        FOREIGN KEY (
            student_id
        )
        REFERENCES students (
            student_id
        )
    );
    """
)

print()
print("SQL tables created successfully.")

# =========================================================
# 10. INSERT DATA INTO SQL TABLES
# =========================================================

students.to_sql(
    "students",
    connection,
    if_exists="append",
    index=False
)

subjects.to_sql(
    "subjects",
    connection,
    if_exists="append",
    index=False
)

academic.to_sql(
    "academic_records",
    connection,
    if_exists="append",
    index=False
)

attendance.to_sql(
    "attendance",
    connection,
    if_exists="append",
    index=False
)

reexam.to_sql(
    "reexam_records",
    connection,
    if_exists="append",
    index=False
)

semester.to_sql(
    "semester_performance",
    connection,
    if_exists="append",
    index=False
)

print(
    "All cleaned data inserted into SQL successfully."
)

# =========================================================
# 11. CREATE DATABASE INDEXES
# =========================================================

cursor.execute(
    """
    CREATE INDEX idx_academic_student
    ON academic_records (
        student_id
    );
    """
)

cursor.execute(
    """
    CREATE INDEX idx_academic_subject
    ON academic_records (
        subject_id
    );
    """
)

cursor.execute(
    """
    CREATE INDEX idx_attendance_student
    ON attendance (
        student_id
    );
    """
)

cursor.execute(
    """
    CREATE INDEX idx_attendance_subject
    ON attendance (
        subject_id
    );
    """
)

cursor.execute(
    """
    CREATE INDEX idx_reexam_student
    ON reexam_records (
        student_id
    );
    """
)

cursor.execute(
    """
    CREATE INDEX idx_reexam_subject
    ON reexam_records (
        subject_id
    );
    """
)

cursor.execute(
    """
    CREATE INDEX idx_semester_student
    ON semester_performance (
        student_id
    );
    """
)

print(
    "Database indexes created successfully."
)

# =========================================================
# 12. CREATE ANALYTICS MASTER SQL VIEW
# =========================================================

cursor.execute(
    """
    CREATE VIEW analytics_master AS

    SELECT

        ar.record_id,

        ar.student_id,

        st.department,

        st.admission_year,

        ar.subject_id,

        su.subject_name,

        su.semester,

        su.course_type,

        su.credits,

        ar.internal_marks,

        ar.external_marks,

        ar.assignment_score,

        ar.total_marks,

        ar.grade,

        ar.result,

        at.classes_held,

        at.classes_attended,

        at.attendance_percentage

    FROM academic_records AS ar

    INNER JOIN students AS st
        ON ar.student_id
        =
        st.student_id

    INNER JOIN subjects AS su
        ON ar.subject_id
        =
        su.subject_id

    INNER JOIN attendance AS at
        ON ar.student_id
        =
        at.student_id

        AND

        ar.subject_id
        =
        at.subject_id;
    """
)

# =========================================================
# 13. CREATE STUDENT SUMMARY SQL VIEW
# =========================================================

cursor.execute(
    """
    CREATE VIEW student_summary AS

    WITH academic_summary AS
    (
        SELECT

            student_id,

            department,

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
            AS passed_subjects,

            SUM(
                CASE
                    WHEN result = 'Fail'
                    THEN 1
                    ELSE 0
                END
            )
            AS failed_subjects,

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
            student_id,
            department
    )

    SELECT

        a.student_id,

        a.department,

        a.average_score,

        a.average_attendance,

        a.passed_subjects,

        a.failed_subjects,

        a.pass_percentage,

        sp.cgpa
        AS final_cgpa,

        sp.active_backlogs
        AS final_active_backlogs,

        sp.cumulative_credits_cleared
        AS final_credits_cleared

    FROM academic_summary AS a

    INNER JOIN semester_performance AS sp

        ON a.student_id
        =
        sp.student_id

    WHERE sp.semester = 8;
    """
)

print(
    "SQL analytical views created successfully."
)

# =========================================================
# 14. COMMIT DATABASE
# =========================================================

connection.commit()

# =========================================================
# 15. TABLE COUNT VERIFICATION
# =========================================================

print()
print("=" * 70)
print("SQL TABLE VERIFICATION")
print("=" * 70)

expected_counts = {
    "students": 500,
    "subjects": 160,
    "academic_records": 16000,
    "attendance": 16000,
    "reexam_records": 2037,
    "semester_performance": 4000
}

validation_errors = []

for table_name, expected in expected_counts.items():

    cursor.execute(
        f"""
        SELECT COUNT(*)
        FROM {table_name};
        """
    )

    actual = cursor.fetchone()[0]

    print(
        f"{table_name}: "
        f"{actual} rows"
    )

    if actual != expected:

        validation_errors.append(
            f"{table_name}: "
            f"expected {expected}, "
            f"found {actual}"
        )

# =========================================================
# 16. VIEW VERIFICATION
# =========================================================

cursor.execute(
    """
    SELECT COUNT(*)
    FROM analytics_master;
    """
)

master_count = cursor.fetchone()[0]

cursor.execute(
    """
    SELECT COUNT(*)
    FROM student_summary;
    """
)

student_summary_count = (
    cursor.fetchone()[0]
)

print()
print(
    "analytics_master view:",
    master_count,
    "rows"
)

print(
    "student_summary view:",
    student_summary_count,
    "rows"
)

if master_count != 16000:

    validation_errors.append(
        "analytics_master view count is incorrect"
    )

if student_summary_count != 500:

    validation_errors.append(
        "student_summary view count is incorrect"
    )

# =========================================================
# 17. FOREIGN KEY VALIDATION
# =========================================================

foreign_key_errors = cursor.execute(
    """
    PRAGMA foreign_key_check;
    """
).fetchall()

print()
print(
    "Foreign key violations:",
    len(foreign_key_errors)
)

if len(foreign_key_errors) > 0:

    validation_errors.append(
        "Foreign key violations detected"
    )

# =========================================================
# 18. SAMPLE SQL ANALYTICS QUERY
# =========================================================

print()
print("=" * 70)
print("SAMPLE SQL QUERY - TOP 5 STUDENTS")
print("=" * 70)

top_students = pd.read_sql_query(
    """
    SELECT

        student_id,

        department,

        average_score,

        average_attendance,

        pass_percentage,

        final_cgpa

    FROM student_summary

    ORDER BY

        average_score DESC,

        final_cgpa DESC

    LIMIT 5;
    """,
    connection
)

print(
    top_students.to_string(
        index=False
    )
)

# =========================================================
# 19. DATABASE OBJECTS
# =========================================================

objects = pd.read_sql_query(
    """
    SELECT

        type,

        name

    FROM sqlite_master

    WHERE type IN (
        'table',
        'view',
        'index'
    )

    AND name NOT LIKE
        'sqlite_%'

    ORDER BY
        type,
        name;
    """,
    connection
)

print()
print("=" * 70)
print("DATABASE OBJECTS")
print("=" * 70)

print(
    objects.to_string(
        index=False
    )
)

# =========================================================
# 20. FINAL RESULT
# =========================================================

print()
print("=" * 70)

if len(validation_errors) == 0:

    print(
        "SQL DATABASE CREATION RESULT: SUCCESS"
    )

    print()
    print(
        "All tables, relationships, indexes and views"
    )

    print(
        "were created successfully."
    )

else:

    print(
        "SQL DATABASE CREATION RESULT: CHECK REQUIRED"
    )

    print()

    for error in validation_errors:

        print(
            "-",
            error
        )

print()
print(
    "Database saved to:"
)

print(
    database_file
)

print("=" * 70)

# =========================================================
# CLOSE CONNECTION
# =========================================================

connection.close()