import fnmatch
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
    with file.open("r", encoding="UTF-8") as fin:
        data = xml.dom.minidom.parse(fin)

    result = data.toprettyxml()
    for remove, insert in [
        (re.escape(str(pathlib.Path(".").resolve())), "{cwd}"),
        (re.escape(str(student_teacher_gradebook._config._MODULE_DIR)), "{source_dir}"),
        (r"\<xr:revisionPtr.+?/\>", '<revision value="redacted"/>'),
    ]:
        result = re.sub(remove, insert, result, flags=re.IGNORECASE)
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


def assert_excel_data_in_dir(dir: pathlib.Path, snapshot: pytest_snapshot.plugin.Snapshot):
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
