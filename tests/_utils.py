import csv
import fnmatch
import io
import pathlib
import re
import typing
import xml.dom.minidom
import zipfile

import click.testing
import pytest_snapshot.plugin

import student_teacher_gradebook
from tests import conftest

RUNNER_TYPE = typing.Callable[[typing.List[str]], click.testing.Result]
""" Type hint that specifies the result from `console_runner`, see link below. """
_ = conftest.console_runner


def _pretty_xml(file: pathlib.Path):
    with file.open("r", encoding="UTF-8", errors="ignore") as fin:
        try:
            data = xml.dom.minidom.parse(fin)
        except Exception as e:
            return None

    result = data.toprettyxml()
    for remove, insert in [
        (re.escape(str(pathlib.Path(".").resolve())), "{cwd}"),
        (re.escape(str(student_teacher_gradebook._config.MODULE_DIR)), "{source_dir}"),
        (r"\<xr:revisionPtr.+?/\>", '<revision value="redacted"/>'),
        (r'xr:uid=".+?"', 'xr:uid="do_not_care"'),
    ]:
        result = re.sub(remove, insert, result, flags=re.IGNORECASE)
    return result


_IGNORE_FILE_PATTERNS = ("docProps/core.xml",)


def _filter_ignore_files(files: typing.Iterable[pathlib.Path], root: pathlib.Path):
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


def assert_excel_data_in_dir(dir: pathlib.Path, snapshot: pytest_snapshot.plugin.Snapshot):
    """Assert that each *.xlsx file under `dir` matches snapshot.

    This is accomplished by unzipping the files and doing a little sanitization
    (e.g., replacing the CWD).
    """
    excel_files = dir.rglob("*.xlsx")
    data = {}
    for file in excel_files:
        unzipped_dir = file.parent / f"{file.name}._unzipped"
        with zipfile.ZipFile(file) as fin:
            fin.extractall(unzipped_dir)
        for data_file in _filter_ignore_files(
                    (file.parent / f"{file.name}._unzipped").rglob("*"), root=unzipped_dir
                ):
            if not data_file.is_file():
                continue
            value = _pretty_xml(data_file)
            if value is None:
                continue
            data[data_file.relative_to(unzipped_dir.parent)
                .as_posix()
                .replace("/", "___")] = value
        wb = student_teacher_gradebook._BaseWorkBook(file)
        try:
            wb.open()
            for sheet in wb.worksheet_names():
                stream = io.StringIO()
                writer = csv.writer(stream)
                for row in wb.get_cells_value_range(sheet, 1, "A"):
                    writer.writerow(row)
                data[file.stem + f"_as_csv__{sheet}.csv"] = stream.getvalue().encode("UTF-8")
        finally:
            wb.close()
    snapshot.assert_match_dir(data, "expected_files")
