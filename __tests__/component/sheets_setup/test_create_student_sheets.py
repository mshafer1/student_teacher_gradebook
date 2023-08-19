import pathlib
import typing

import pytest
import pytest_snapshot.plugin

import __tests__._utils
import student_teacher_gradebook

MODULE_DIR = pathlib.Path(__file__).parent


@pytest.mark.parametrize(
    ["roster"],
    [
        pytest.param([]),
        pytest.param(["John Doe", "Molly Doe", "Stephen Jane"], id="roster1"),
    ],
)
def test____teacher_book_with_roster___populate_student_sheets___creates_expected_workbooks(
    roster: typing.Iterable[str],
    temp_teacher_workbook: pathlib.Path,
    console_runner: __tests__._utils.RUNNER_TYPE,
    snapshot: pytest_snapshot.plugin.Snapshot,
    request,
):
    with student_teacher_gradebook._MainWorkbook(temp_teacher_workbook) as teacher_book:
        teacher_book.set_column_range("Roster", 2, "B", roster)

    result = console_runner("populate-student-sheets")

    assert not result.exception
    output_dir = temp_teacher_workbook.parent

    snapshot.snapshot_dir = (
        MODULE_DIR / "snapshots/populate_student_sheets" / request.node.callspec.id
    )
    __tests__._utils._assert_excel_data_in_dir(output_dir, snapshot)
