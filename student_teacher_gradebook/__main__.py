import logging
import pathlib
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
    with student_teacher_gradebook._MainWorkbook(student_teacher_gradebook._config.TEACHER_BOOK) as main_workbook:
        ...
        # title_stem = "MUSC184_F21 Progress tracking sheet for "
        # template_filename = main_workbook.progress_sheet[
        #     grades_sheet_manager._consts.TEMPLATE_SHEET_LOCATION
        # ].value
        # template_file = pathlib.Path(template_filename)
        # if not template_file.is_file():  # if workbook doesn't exist, try relative to main workbook
        #     template_file = master_sheet.parent / template_filename

        # _MODULE_LOGGER.info("Copying %d workbooks...", len(main_workbook.roster))
        # for i, student_name, student_workbook in tqdm(
        #     zip(
        #         range(len(main_workbook.roster)),
        #         main_workbook.roster,
        #         main_workbook.student_workbooks,
        #     )
        # ):
        #     target_filename = template_file.resolve().parent / (title_stem + student_name + ".xlsx")
        #     if student_workbook is None:
        #         student_workbook = target_filename
        #         _MODULE_LOGGER.debug("Copying file to: %s", student_workbook)
        #     elif pathlib.Path(student_workbook).is_file():
        #         _MODULE_LOGGER.warning("Workbook (%s) already exists, skipping", student_workbook)
        #         continue
        #     else:
        #         _MODULE_LOGGER.debug("Workbook (%s) does not exist, recreating", student_workbook)

        #     shutil.copy(
        #         str(template_file),
        #         str(target_filename),
        #     )

        #     main_workbook.set_workbook(i, student_workbook)



if __name__ == "__main__":
    _cli()