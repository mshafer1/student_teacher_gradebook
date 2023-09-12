import pathlib
import shutil
import typing

import pytest
import pytest_snapshot.plugin

import student_teacher_gradebook
import tests._utils

MODULE_DIR = pathlib.Path(__file__).parent
TEST_CASE_DIR = MODULE_DIR / "test_cases"

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
    console_runner: tests._utils.RUNNER_TYPE,
    snapshot: pytest_snapshot.plugin.Snapshot,
    request,
):
    with student_teacher_gradebook.MainWorkbook(temp_teacher_workbook) as teacher_book:
        teacher_book.set_column_range("Roster", 2, "B", roster)

    result = console_runner("populate-student-sheets")

    assert not result.exception
    output_dir = temp_teacher_workbook.parent

    snapshot.snapshot_dir = (
        MODULE_DIR / "snapshots/populate_student_sheets" / request.node.callspec.id
    )
    tests._utils.assert_excel_data_in_dir(output_dir, snapshot)


@pytest.fixture(
    params=[item for item in TEST_CASE_DIR.iterdir() if item.is_dir()],
    ids=lambda o: o.name,
)
def test_case(request):
    return request.param


@pytest.fixture()
def test_case_with_temp_cwd(test_case: pathlib.Path, temp_cwd: pathlib.Path):
    for source in test_case.glob("*.xlsx"):
        shutil.copy2(source, temp_cwd)

    teacher_book = temp_cwd / source.name

    old_value = student_teacher_gradebook._config.TEACHER_BOOK
    student_teacher_gradebook._config.TEACHER_BOOK = teacher_book
    yield test_case
    student_teacher_gradebook._config.TEACHER_BOOK = old_value


@pytest.fixture()
def test_id_as_file(temp_cwd: pathlib.Path, request):
    (temp_cwd / request.node.callspec.id).touch()

def test____populate_student_sheets___expected_sheets(
    test_case_with_temp_cwd: pathlib.Path,
    temp_cwd: pathlib.Path,
    console_runner: tests._utils.RUNNER_TYPE,
    snapshot: pytest_snapshot.plugin.Snapshot,
    test_id_as_file,
    fixed_datetime,
):
    result = console_runner("populate-student-sheets")

    assert not result.exception
    output_dir = temp_cwd

    snapshot.snapshot_dir = test_case_with_temp_cwd / "output" / "update_student_sheets"
    tests._utils.assert_excel_data_in_dir(output_dir, snapshot)
