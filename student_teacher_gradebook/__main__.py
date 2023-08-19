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
    ...


@_cli.command()
def populate_student_sheets():
    """Generate students' sheets from template."""
    _MODULE_LOGGER.info("Loading main workbook...")
    with student_teacher_gradebook._MainWorkbook(
        student_teacher_gradebook._config.TEACHER_BOOK
    ) as main_workbook:
        for i, student in enumerate(main_workbook.roster):
            print(student)
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
                new_student_data = student_teacher_gradebook.StudentData(**{**student._asdict(), "student_file": student_file})
                main_workbook.update_student_values(i, new_student_data)


if __name__ == "__main__":
    _cli()
