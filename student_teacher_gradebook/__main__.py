"""A utility script for a teacher to track all students' assignments in one main spreadsheet and copy out to individual spreadsheets for students."""  # noqa: W505 - docs
import logging
import pathlib
import shutil

import click

import student_teacher_gradebook
import student_teacher_gradebook._config

_MODULE_LOGGER = logging.getLogger(__name__)


@click.group()
def _cli():
    ...


@_cli.command()
def update_student_sheets():
    """Update students' sheets."""
    _MODULE_LOGGER.info("Loading main workbook...")
    with student_teacher_gradebook.MainWorkbook(
        student_teacher_gradebook._config.TEACHER_BOOK
    ) as main_workbook:
        # for each other worksheet besides 'Roster' and 'Config'
        print("Loading data...")
        worksheets_to_process = [
            ws
            for ws in main_workbook.worksheet_names()
            if ws
            not in {
                student_teacher_gradebook._config.CONFIG_SHEET_NAME,
                student_teacher_gradebook._config.ROSTER_SHEET_NAME,
            }
        ]
        student_data_mapping = main_workbook.get_student_values_from_sheets(worksheets_to_process)

        for student_name, data in student_data_mapping.items():
            student = main_workbook.roster_as_mapping[student_name]
            with main_workbook.open_student_workbook(student) as student_book:
                temp_new_sheet_name = "Progress_new"
                sheet_name = "Progress"
                student_book.add_sheet(temp_new_sheet_name)
                for sheet in student_book.worksheet_names():
                    if sheet == temp_new_sheet_name:
                        continue
                    student_book.remove_sheet(sheet)
                student_book.rename_sheet(temp_new_sheet_name, sheet_name)
                for i, value in enumerate(data):
                    student_book.set_row_range(
                        sheet_name=sheet_name, start_column_index="A", row_index=1 + i, values=value
                    )
                student_book.save()


@_cli.command()
def populate_student_sheets():
    """Generate students' sheets from template."""
    _MODULE_LOGGER.info("Loading main workbook...")
    with student_teacher_gradebook.MainWorkbook(
        student_teacher_gradebook._config.TEACHER_BOOK
    ) as main_workbook:
        for i, student in enumerate(main_workbook.roster):
            print("Evaluating:", student)
            if student.student_file is None:
                _MODULE_LOGGER.info("Student %s does not have a student sheet yet.", student.name)
                try:
                    student_file = main_workbook.config.student_filename_format_string.format(
                        **student._asdict()
                    )
                except KeyError:
                    _MODULE_LOGGER.warning(
                        "Error, invalid student template name in %s sheet. Valid place holder values are: %s",
                        student_teacher_gradebook._config.CONFIG_SHEET_NAME,
                        student._fields,
                    )
                    raise

            if not pathlib.Path(student_file).is_file():
                if student.student_file is not None:
                    _MODULE_LOGGER.warning(
                        "File for student %s is missing, recreating.", student.name
                    )
                _MODULE_LOGGER.info(
                    "Creating student sheet for %s from %s",
                    student.name,
                    main_workbook.config.student_template_filename,
                )
                shutil.copy2(
                    student_teacher_gradebook._config.SOURCE_DIR
                    / main_workbook.config.student_template_filename,
                    student_file,
                )
            if student.student_file is None:
                new_student_data = student_teacher_gradebook.StudentData(
                    **{**student._asdict(), "student_file": student_file}
                )
                main_workbook.update_student_values(i, new_student_data)


if __name__ == "__main__":
    _cli()
