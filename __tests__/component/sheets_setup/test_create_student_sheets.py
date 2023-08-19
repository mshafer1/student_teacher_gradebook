import fnmatch
import pathlib
import re
import typing
import xml.dom.minidom
import zipfile

import pytest
import pytest_snapshot.plugin

import __tests__._utils
import student_teacher_gradebook

MODULE_DIR = pathlib.Path(__file__).parent


def _pretty_xml(file: pathlib.Path):
    with file.open("r", encoding="UTF-8") as fin:
        data = xml.dom.minidom.parse(fin)
    # file.write_text(xmltodict.unparse(data, short_empty_elements=True, indent=" "*4, pretty=True), encoding="UTF-8")
    result = data.toprettyxml()
    for remove, insert in [
        (re.escape(str(pathlib.Path(".").resolve())), "{cwd}"),
        (r"\<xr:revisionPtr.+?/\>", "<revision value=redacted/>"),
    ]:
        result = re.sub(remove, insert, result)
    return result


_IGNORE_FILE_PATTERNS = ("docProps/core.xml",)


def _filter_ignore_files(files: typing.Iterable[pathlib.Path], root: pathlib.Path):
    # return filter(lambda o: not any([fnmatch.fnmatch(data_file, ignore_pattern) for ignore_pattern in _IGNORE_FILE_PATTERNS]), files)
    result = []
    for file in files:
        if not any(
            [
                fnmatch.fnmatch(file.relative_to(root).as_posix(), ignore_pattern)
                for ignore_pattern in _IGNORE_FILE_PATTERNS
            ]
        ):
            result.append(file)
    return result


def _assert_excel_data_in_dir(dir: pathlib.Path, snapshot: pytest_snapshot.plugin.Snapshot):
    excel_files = dir.rglob("*.xlsx")
    data = {}
    for file in excel_files:
        unzipped_dir = file.parent / f"{file.name}._unzipped"
        with zipfile.ZipFile(file) as fin:
            fin.extractall(unzipped_dir)
        data.update(
            **{
                data_file.relative_to(unzipped_dir.parent)
                .as_posix()
                .replace("/", "___"): _pretty_xml(data_file)
                for data_file in _filter_ignore_files(
                    (file.parent / f"{file.name}._unzipped").rglob("*"), root=unzipped_dir
                )
                if data_file.is_file()
            }
        )
    snapshot.assert_match_dir(data, "expected_files")


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
    _assert_excel_data_in_dir(output_dir, snapshot)
