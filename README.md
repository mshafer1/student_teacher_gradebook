# Student-Teacher-Gradebook

This is a utility script (written in Python) that allows a teacher to track all student assignments in one workbook while cloning student specific information to a workbook for each student (to allow for sharing with the specific student).

# Requirements:

* Windows
  
  Currently, only Microsoft Excel spreadsheets are supported, and they are controlled using the pywin32 libary. This combination necessitates running on Windows.

* Python 3.8 or higher

# Installation

Option 1:
<details open>
<summary>With <a href="">pipx</a></summary>

`pipx install git+https://github.com/mshafer1/student_teacher_gradebook.git`

**NOTE**: Using this method will require that Git setting for long-paths is enable<sup>[1]</sup>
</details>

Option 2:
<details>
<summary>From source (with <a href="https://python-poetry.org/">poetry</a>)</summary>

```
git clone https://github.com/mshafer1/student_teacher_gradebook.git

cd student_teacher_gradebook

poetry install
```

**NOTE**: Using this method will require that Git setting for long-paths is enable<sup>[1]</sup>
</details>

Option 3:
<details>
<summary>From source (with self managed venv)</summary>

```
git clone https://github.com/mshafer1/student_teacher_gradebook.git

cd student_teacher_gradebook

poetry install
```
**NOTE**: Using this method will require that Git setting for long-paths is enable<sup>[1]</sup>
</details>

# Source files

Example `TeacherBook.xlsx` and example `studentTemplate.xlsx` are available in the `source` directory.

When executed, `student-teacher-gradebook` loads configuration from the `Config` sheet and uses the `Roster` sheet for tracking sutdent sheets.

To setup:
* Fill out the `Roster` sheet with names of students (leave the workbook section alone).
* Edit/Adjust the variables in `Config` sheet as needed
* Run `student-teacher-gradebook populate-student-sheets` to generate student sheets and have them added to `Roster`


# Usage

When running a script, it will look in the current working directory for TeacherBook.xlsx.

This module provides the following scripts:

## `student-teacher-gradebook populate-student-sheets`

Reads the `Roster` sheet and copies the student template for each student listed.

The name of the student sheet is determined by the value of the `Student Filename Format String` field in the `Config` sheet.

Place holders are allowed. Allowed values are currently:
* `id_` A script generated unique ID for each student (based on the student name and time of assignment)
* `name` The name of the student as found in the B column of the `Roster` sheet.

Files are copied from the `studentTemplate.xlsx` file.

Output is stored in the current working directory (unless a path in `Student Filename Format String` overrides this)


## `update-student-sheets`

Reads the `Roster` sheet and copies data from the remaining sheets to each student's workbook.

The location of the student workbook is determined by:
  * relative to teacher sheet
  * UNLESS absolute path is stored in `Roster`

For each sheet (other than `Config` and `Roster`), for which a student has their name occur in the first column AND at least one non-empty cell in their row, a corresponding row is created in the student sheet with the name of the teacher's sheet replacing their name.

Student sheets are generated fresh each run (removing all other sheets in the student workbook).



# Known issues

* Excel formatting values as dates

  Due to how Excel accepts values in, values like `8/10` copied from the teacher sheet to the student sheet may render as "10-Oct"

[1]: https://stackoverflow.com/a/59052951/8100990 "Git for Windows enable long paths within Git."