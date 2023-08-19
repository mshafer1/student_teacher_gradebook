# Student-Teacher-Gradebook

This is a utility script (written in Python) that allows a teacher to track all student assignments in one workbook while cloning student specific information to a workbook for each student (to allow for sharing with the specific student).

# Requirements:

* Windows
  
  Currently, only Microsoft Excel spreadsheets are supported, and they are controlled using the pywin32 libary. This combination necessitates running on Windows.

* Python 3.8 or higher


# Usage

## `populate-student-sheets`

Reads the `Roster` sheet and copies the student template for each student listed.

The name of the student sheet is determined by the value of the `Student Filename Format String` field in the `Config` sheet.

Place holders are allowed. Allowed values are currently:
* `id_` A script generated unique ID for each student (based on the student name and time of assignment)
* `name` The name of the student as found in the B column of the `Roster` sheet.

Files are copied from the `studentTemplate.xlsx` file.

Output is stored in the current working directory (unless a path in `Student Filename Format String` overrides this)



# Known issues

* Excel formatting values as dates

  Due to how Excel accepts values in, values like `8/10` copied from the teacher sheet to the student sheet may render as "10-Oct"
