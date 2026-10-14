import pandas as pd
from pathlib import Path

# =========================================================
# FILE LOCATION
# =========================================================

file_path = Path(
    r"C:\Users\mrfqu\Desktop\University_Academic_Analytics\01_Dataset\Academic_Data_Model.xlsx"
)

if not file_path.exists():
    raise FileNotFoundError(
        f"Excel file not found: {file_path}"
    )

print()
print("=" * 60)
print("UNIVERSITY ACADEMIC ANALYTICS - FINAL DATA VALIDATION")
print("=" * 60)

# =========================================================
# LOAD ALL TABLES
# =========================================================

students = pd.read_excel(
    file_path,
    sheet_name="Students"
)

subjects = pd.read_excel(
    file_path,
    sheet_name="Subjects"
)

academic = pd.read_excel(
    file_path,
    sheet_name="Academic_Records"
)

attendance = pd.read_excel(
    file_path,
    sheet_name="Attendance"
)

reexam = pd.read_excel(
    file_path,
    sheet_name="Reexam_Records"
)

semester = pd.read_excel(
    file_path,
    sheet_name="Semester_Performance"
)

# =========================================================
# ERROR TRACKER
# =========================================================

errors = []


def check(condition, message):

    if condition:
        print(f"PASS : {message}")

    else:
        print(f"FAIL : {message}")
        errors.append(message)


# =========================================================
# 1. ROW COUNTS
# =========================================================

print()
print("=" * 60)
print("1. TABLE SIZE VALIDATION")
print("=" * 60)

print("Students:", len(students))
print("Subjects:", len(subjects))
print("Academic Records:", len(academic))
print("Attendance Records:", len(attendance))
print("Re-exam Records:", len(reexam))
print("Semester Performance Records:", len(semester))

print()

check(
    len(students) == 500,
    "Students table contains exactly 500 students"
)

check(
    len(subjects) == 160,
    "Subjects table contains exactly 160 subjects"
)

check(
    len(academic) == 16000,
    "Academic Records contains exactly 16,000 records"
)

check(
    len(attendance) == 16000,
    "Attendance contains exactly 16,000 records"
)

check(
    len(semester) == 4000,
    "Semester Performance contains exactly 4,000 records"
)

check(
    len(reexam) > 0,
    "Re-examination table contains records"
)

# =========================================================
# 2. MISSING VALUE CHECK
# =========================================================

print()
print("=" * 60)
print("2. MISSING VALUE VALIDATION")
print("=" * 60)

tables = {
    "Students": students,
    "Subjects": subjects,
    "Academic Records": academic,
    "Attendance": attendance,
    "Reexam Records": reexam,
    "Semester Performance": semester
}

for table_name, dataframe in tables.items():

    missing = int(
        dataframe.isnull().sum().sum()
    )

    print(
        f"{table_name}: {missing} missing values"
    )

    check(
        missing == 0,
        f"{table_name} has no missing values"
    )

# =========================================================
# 3. PRIMARY KEY / DUPLICATE CHECKS
# =========================================================

print()
print("=" * 60)
print("3. DUPLICATE VALIDATION")
print("=" * 60)

check(
    students["Student_ID"].duplicated().sum() == 0,
    "Student IDs are unique"
)

check(
    subjects["Subject_ID"].duplicated().sum() == 0,
    "Subject IDs are unique"
)

check(
    academic["Record_ID"].duplicated().sum() == 0,
    "Academic Record IDs are unique"
)

check(
    reexam["Reexam_ID"].duplicated().sum() == 0,
    "Re-examination IDs are unique"
)

academic_pair_duplicates = (
    academic[
        [
            "Student_ID",
            "Subject_ID"
        ]
    ]
    .duplicated()
    .sum()
)

check(
    academic_pair_duplicates == 0,
    "No duplicate Student-Subject pairs in Academic Records"
)

attendance_pair_duplicates = (
    attendance[
        [
            "Student_ID",
            "Subject_ID"
        ]
    ]
    .duplicated()
    .sum()
)

check(
    attendance_pair_duplicates == 0,
    "No duplicate Student-Subject pairs in Attendance"
)

semester_pair_duplicates = (
    semester[
        [
            "Student_ID",
            "Semester"
        ]
    ]
    .duplicated()
    .sum()
)

check(
    semester_pair_duplicates == 0,
    "No duplicate Student-Semester records"
)

# =========================================================
# 4. FOREIGN KEY VALIDATION
# =========================================================

print()
print("=" * 60)
print("4. RELATIONSHIP / FOREIGN KEY VALIDATION")
print("=" * 60)

student_ids = set(
    students["Student_ID"]
)

subject_ids = set(
    subjects["Subject_ID"]
)

check(
    set(academic["Student_ID"]).issubset(student_ids),
    "All Academic Record Student IDs exist in Students"
)

check(
    set(attendance["Student_ID"]).issubset(student_ids),
    "All Attendance Student IDs exist in Students"
)

check(
    set(reexam["Student_ID"]).issubset(student_ids),
    "All Re-exam Student IDs exist in Students"
)

check(
    set(semester["Student_ID"]).issubset(student_ids),
    "All Semester Performance Student IDs exist in Students"
)

check(
    set(academic["Subject_ID"]).issubset(subject_ids),
    "All Academic Subject IDs exist in Subjects"
)

check(
    set(attendance["Subject_ID"]).issubset(subject_ids),
    "All Attendance Subject IDs exist in Subjects"
)

check(
    set(reexam["Subject_ID"]).issubset(subject_ids),
    "All Re-exam Subject IDs exist in Subjects"
)

# =========================================================
# 5. DEPARTMENT-SUBJECT CONSISTENCY
# =========================================================

print()
print("=" * 60)
print("5. DEPARTMENT CONSISTENCY")
print("=" * 60)

department_check = (
    academic
    .merge(
        students[
            [
                "Student_ID",
                "Department"
            ]
        ],
        on="Student_ID",
        how="left"
    )
    .merge(
        subjects[
            [
                "Subject_ID",
                "Department"
            ]
        ],
        on="Subject_ID",
        how="left",
        suffixes=(
            "_Student",
            "_Subject"
        )
    )
)

wrong_department_records = (
    department_check[
        "Department_Student"
    ]
    !=
    department_check[
        "Department_Subject"
    ]
).sum()

check(
    wrong_department_records == 0,
    "Every student is linked only to subjects from their department"
)

# =========================================================
# 6. RECORD COUNT PER STUDENT
# =========================================================

print()
print("=" * 60)
print("6. STUDENT RECORD COVERAGE")
print("=" * 60)

academic_per_student = (
    academic
    .groupby("Student_ID")
    .size()
)

attendance_per_student = (
    attendance
    .groupby("Student_ID")
    .size()
)

semester_per_student = (
    semester
    .groupby("Student_ID")
    .size()
)

check(
    (academic_per_student == 32).all(),
    "Every student has exactly 32 Academic Records"
)

check(
    (attendance_per_student == 32).all(),
    "Every student has exactly 32 Attendance Records"
)

check(
    (semester_per_student == 8).all(),
    "Every student has exactly 8 Semester Performance records"
)

# =========================================================
# 7. SUBJECT RECORD COVERAGE
# =========================================================

print()
print("=" * 60)
print("7. SUBJECT RECORD COVERAGE")
print("=" * 60)

academic_per_subject = (
    academic
    .groupby("Subject_ID")
    .size()
)

attendance_per_subject = (
    attendance
    .groupby("Subject_ID")
    .size()
)

check(
    (academic_per_subject == 100).all(),
    "Every subject has exactly 100 Academic Records"
)

check(
    (attendance_per_subject == 100).all(),
    "Every subject has exactly 100 Attendance Records"
)

# =========================================================
# 8. MARKS VALIDATION
# =========================================================

print()
print("=" * 60)
print("8. MARKS VALIDATION")
print("=" * 60)

incorrect_totals = (
    (
        academic["Internal_Marks"]
        +
        academic["External_Marks"]
        +
        academic["Assignment_Score"]
    )
    !=
    academic["Total_Marks"]
).sum()

check(
    incorrect_totals == 0,
    "Total Marks equals Internal + External + Assignment"
)

check(
    academic["Internal_Marks"].between(
        0,
        30
    ).all(),
    "All Internal Marks are between 0 and 30"
)

check(
    academic["External_Marks"].between(
        0,
        60
    ).all(),
    "All External Marks are between 0 and 60"
)

check(
    academic["Assignment_Score"].between(
        0,
        10
    ).all(),
    "All Assignment Scores are between 0 and 10"
)

check(
    academic["Total_Marks"].between(
        0,
        100
    ).all(),
    "All Total Marks are between 0 and 100"
)

# =========================================================
# 9. GRADE AND RESULT VALIDATION
# =========================================================

print()
print("=" * 60)
print("9. GRADE / RESULT VALIDATION")
print("=" * 60)


def expected_grade(mark):

    if mark >= 90:
        return "A++"

    elif mark >= 80:
        return "A+"

    elif mark >= 70:
        return "B++"

    elif mark >= 60:
        return "B+"

    elif mark >= 50:
        return "C"

    else:
        return "RA"


calculated_grades = (
    academic["Total_Marks"]
    .apply(expected_grade)
)

incorrect_grades = (
    calculated_grades
    != academic["Grade"]
).sum()

check(
    incorrect_grades == 0,
    "Academic grades match Total Marks"
)

expected_results = (
    academic["Total_Marks"]
    .apply(
        lambda mark:
        "Pass"
        if mark >= 50
        else "Fail"
    )
)

incorrect_results = (
    expected_results
    != academic["Result"]
).sum()

check(
    incorrect_results == 0,
    "Pass/Fail results match Total Marks"
)

# =========================================================
# 10. ATTENDANCE VALIDATION
# =========================================================

print()
print("=" * 60)
print("10. ATTENDANCE VALIDATION")
print("=" * 60)

check(
    (
        attendance["Classes_Attended"]
        <=
        attendance["Classes_Held"]
    ).all(),
    "Classes Attended never exceeds Classes Held"
)

check(
    (
        attendance["Classes_Held"] > 0
    ).all(),
    "All subjects have positive Classes Held"
)

check(
    attendance[
        "Attendance_Percentage"
    ].between(
        0,
        100
    ).all(),
    "All Attendance Percentages are between 0 and 100"
)

calculated_attendance = (
    attendance[
        "Classes_Attended"
    ]
    /
    attendance[
        "Classes_Held"
    ]
    * 100
).round(2)

attendance_difference = (
    calculated_attendance
    -
    attendance[
        "Attendance_Percentage"
    ]
).abs()

check(
    (
        attendance_difference <= 0.01
    ).all(),
    "Attendance Percentages are calculated correctly"
)

# =========================================================
# 11. ACADEMIC / ATTENDANCE PAIR MATCHING
# =========================================================

print()
print("=" * 60)
print("11. MARKS-ATTENDANCE MATCHING")
print("=" * 60)

academic_pairs = set(
    zip(
        academic["Student_ID"],
        academic["Subject_ID"]
    )
)

attendance_pairs = set(
    zip(
        attendance["Student_ID"],
        attendance["Subject_ID"]
    )
)

check(
    academic_pairs == attendance_pairs,
    "Academic Records and Attendance contain identical Student-Subject pairs"
)

# =========================================================
# 12. SUBJECT CREDIT VALIDATION
# =========================================================

print()
print("=" * 60)
print("12. SUBJECT CREDIT VALIDATION")
print("=" * 60)

check(
    subjects["Credits"].isin(
        [3, 4]
    ).all(),
    "Every subject has either 3 or 4 credits"
)

credit_distribution = (
    subjects[
        "Credits"
    ]
    .value_counts()
)

print()
print("Credit Distribution:")
print(credit_distribution)

check(
    int(
        credit_distribution.get(
            3,
            0
        )
    ) == 80,
    "Exactly 80 subjects have 3 credits"
)

check(
    int(
        credit_distribution.get(
            4,
            0
        )
    ) == 80,
    "Exactly 80 subjects have 4 credits"
)

# =========================================================
# 13. RE-EXAMINATION VALIDATION
# =========================================================

print()
print("=" * 60)
print("13. RE-EXAMINATION VALIDATION")
print("=" * 60)

check(
    (
        reexam[
            "Attempt_Semester"
        ]
        >
        reexam[
            "Original_Semester"
        ]
    ).all(),
    "Every re-exam occurs after the original semester"
)

check(
    reexam[
        "Attempt_Number"
    ].between(
        1,
        3
    ).all(),
    "All Re-exam Attempt Numbers are between 1 and 3"
)

check(
    reexam[
        "Reexam_Marks"
    ].between(
        0,
        100
    ).all(),
    "All Re-exam Marks are between 0 and 100"
)

expected_reexam_results = (
    reexam[
        "Reexam_Marks"
    ]
    .apply(
        lambda mark:
        "Pass"
        if mark >= 50
        else "Fail"
    )
)

check(
    (
        expected_reexam_results
        ==
        reexam["Result"]
    ).all(),
    "All Re-exam Results match Re-exam Marks"
)

expected_reexam_grades = (
    reexam[
        "Reexam_Marks"
    ]
    .apply(expected_grade)
)

check(
    (
        expected_reexam_grades
        ==
        reexam["Grade"]
    ).all(),
    "All Re-exam Grades match Re-exam Marks"
)

# ---------------------------------------------------------
# Every re-exam must belong to an originally failed course
# ---------------------------------------------------------

failed_pairs = set(
    zip(
        academic[
            academic["Result"] == "Fail"
        ]["Student_ID"],
        academic[
            academic["Result"] == "Fail"
        ]["Subject_ID"]
    )
)

reexam_pairs = set(
    zip(
        reexam["Student_ID"],
        reexam["Subject_ID"]
    )
)

check(
    reexam_pairs.issubset(
        failed_pairs
    ),
    "Every Re-exam belongs to an originally failed course"
)

# =========================================================
# 14. SEMESTER PERFORMANCE VALIDATION
# =========================================================

print()
print("=" * 60)
print("14. SEMESTER PERFORMANCE VALIDATION")
print("=" * 60)

check(
    semester["Semester"].between(
        1,
        8
    ).all(),
    "All Semester values are between 1 and 8"
)

check(
    semester["SGPA"].between(
        0,
        10
    ).all(),
    "All SGPA values are between 0 and 10"
)

check(
    semester["CGPA"].between(
        0,
        10
    ).all(),
    "All CGPA values are between 0 and 10"
)

check(
    (
        semester[
            "New_Backlogs"
        ] >= 0
    ).all(),
    "New Backlogs never becomes negative"
)

check(
    (
        semester[
            "Cleared_Backlogs"
        ] >= 0
    ).all(),
    "Cleared Backlogs never becomes negative"
)

check(
    (
        semester[
            "Active_Backlogs"
        ] >= 0
    ).all(),
    "Active Backlogs never becomes negative"
)

check(
    (
        semester[
            "Cumulative_Credits_Cleared"
        ]
        <= 112
    ).all(),
    "Cumulative Credits Cleared never exceeds 112"
)

# =========================================================
# 15. BACKLOG CROSS-CHECK
# =========================================================

print()
print("=" * 60)
print("15. BACKLOG CROSS-CHECK")
print("=" * 60)

original_failures = int(
    (
        academic["Result"]
        ==
        "Fail"
    ).sum()
)

cleared_courses = (
    reexam[
        reexam["Result"] == "Pass"
    ][
        [
            "Student_ID",
            "Subject_ID"
        ]
    ]
    .drop_duplicates()
)

courses_cleared = len(
    cleared_courses
)

expected_active = (
    original_failures
    -
    courses_cleared
)

semester_8 = semester[
    semester["Semester"] == 8
]

calculated_active = int(
    semester_8[
        "Active_Backlogs"
    ].sum()
)

print(
    "Original Failed Courses:",
    original_failures
)

print(
    "Courses Cleared Through Re-exam:",
    courses_cleared
)

print(
    "Expected Active Backlogs:",
    expected_active
)

print(
    "Calculated Active Backlogs:",
    calculated_active
)

check(
    expected_active
    ==
    calculated_active,
    "Final Active Backlogs match Re-examination records"
)

check(
    int(
        semester[
            "New_Backlogs"
        ].sum()
    )
    ==
    original_failures,
    "Total New Backlogs equals original failed courses"
)

check(
    int(
        semester[
            "Cleared_Backlogs"
        ].sum()
    )
    ==
    courses_cleared,
    "Total Cleared Backlogs equals cleared Re-exam courses"
)

# =========================================================
# 16. FINAL CREDIT CROSS-CHECK
# =========================================================

print()
print("=" * 60)
print("16. FINAL CREDIT CROSS-CHECK")
print("=" * 60)

subject_credit_map = dict(
    zip(
        subjects["Subject_ID"],
        subjects["Credits"]
    )
)

expected_final_credits = {}

for student_id in students[
    "Student_ID"
]:

    student_failed = set(
        academic[
            (academic["Student_ID"] == student_id)
            &
            (academic["Result"] == "Fail")
        ]["Subject_ID"]
    )

    student_cleared = set(
        reexam[
            (reexam["Student_ID"] == student_id)
            &
            (reexam["Result"] == "Pass")
        ]["Subject_ID"]
    )

    active_failed = (
        student_failed
        -
        student_cleared
    )

    active_failed_credits = sum(
        subject_credit_map[
            subject
        ]
        for subject in active_failed
    )

    expected_final_credits[
        student_id
    ] = (
        112
        -
        active_failed_credits
    )

final_credit_errors = 0

for _, row in semester_8.iterrows():

    student_id = row[
        "Student_ID"
    ]

    actual_credits = int(
        row[
            "Cumulative_Credits_Cleared"
        ]
    )

    expected_credits = int(
        expected_final_credits[
            student_id
        ]
    )

    if actual_credits != expected_credits:
        final_credit_errors += 1

check(
    final_credit_errors == 0,
    "Final Credits Cleared correctly reflect remaining active backlogs"
)

# =========================================================
# 17. PROJECT METRICS
# =========================================================

print()
print("=" * 60)
print("17. FINAL PROJECT METRICS")
print("=" * 60)

pass_count = int(
    (
        academic["Result"] == "Pass"
    ).sum()
)

fail_count = int(
    (
        academic["Result"] == "Fail"
    ).sum()
)

pass_percentage = (
    pass_count
    /
    len(academic)
    *
    100
)

average_score = (
    academic[
        "Total_Marks"
    ].mean()
)

average_attendance = (
    attendance[
        "Attendance_Percentage"
    ].mean()
)

average_final_cgpa = (
    semester_8[
        "CGPA"
    ].mean()
)

students_with_active_backlogs = (
    semester_8[
        "Active_Backlogs"
    ] > 0
).sum()

print(
    "Total Students:",
    len(students)
)

print(
    "Total Subjects:",
    len(subjects)
)

print(
    "Total Academic Records:",
    len(academic)
)

print(
    "Pass Records:",
    pass_count
)

print(
    "Fail Records:",
    fail_count
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
    "Average Final CGPA:",
    round(
        average_final_cgpa,
        2
    )
)

print(
    "Students with Active Backlogs:",
    int(
        students_with_active_backlogs
    )
)

# =========================================================
# FINAL RESULT
# =========================================================

print()
print("=" * 60)

if len(errors) == 0:

    print(
        "FINAL VALIDATION RESULT: PASSED"
    )

    print()
    print(
        "The complete academic dataset is internally consistent"
    )

    print(
        "and is ready for the Data Analytics stage."
    )

else:

    print(
        "FINAL VALIDATION RESULT: FAILED"
    )

    print()
    print(
        "Problems found:",
        len(errors)
    )

    for number, error in enumerate(
        errors,
        start=1
    ):

        print(
            f"{number}. {error}"
        )

print("=" * 60)