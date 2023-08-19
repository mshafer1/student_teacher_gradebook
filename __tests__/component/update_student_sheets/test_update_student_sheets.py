import pathlib
import shutil

import pytest
import pytest_snapshot.plugin

import __tests__._utils
import student_teacher_gradebook

MODULE_DIR = pathlib.Path(__file__).parent
TEST_CASE_DIR = MODULE_DIR / "test_cases"

@pytest.fixture
def test_case():
    for item in TEST_CASE_DIR.iterdir():
        if item.is_dir():
            yield item

@pytest.fixture()
def test_case_with_temp_cwd(test_case: pathlib.Path, temp_cwd: pathlib.Path):
    for source in test_case.glob("*.xlsx"):
        shutil.copy2(source, temp_cwd)

    teacher_book = temp_cwd / source.name

    old_value = student_teacher_gradebook._config.TEACHER_BOOK
    student_teacher_gradebook._config.TEACHER_BOOK = teacher_book
    yield test_case
    student_teacher_gradebook._config.TEACHER_BOOK = old_value


def test____teacher_book_with_roster___update_student_sheets___expected_spreadsheets_out(
    test_case_with_temp_cwd: pathlib.Path,
    temp_cwd: pathlib.Path,
    console_runner: __tests__._utils.RUNNER_TYPE,
    snapshot: pytest_snapshot.plugin.Snapshot,
    request,
):
    result = console_runner("update-student-sheets")

    assert not result.exception
    output_dir = temp_cwd

    snapshot.snapshot_dir = (
        test_case_with_temp_cwd / "output" / "update_student_sheets"
    )
    __tests__._utils.assert_excel_data_in_dir(output_dir, snapshot)