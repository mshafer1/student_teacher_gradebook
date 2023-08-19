import typing
import pytest

import __tests__._utils


@pytest.mark.parametrize(
    ["roster"],
    [
        pytest.param([]),
        pytest.param(["John Doe", "Molly Doe", "Stephen Jane"]),
    ],
)
def test____teacher_book_with_roster___populate_student_sheets___creates_expected_workbooks(
    roster: typing.Iterable[str],
    temp_teacher_workbook,
    console_runner: __tests__._utils.RUNNER_TYPE,
):
    

    result = console_runner("populate-student-sheets")

    assert not result.exception
