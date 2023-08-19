import functools
import pathlib
import shutil
import unittest.mock

import click.testing
import pytest
from pytest_mock.plugin import MockerFixture

import student_teacher_gradebook
import student_teacher_gradebook.__main__


@pytest.fixture()
def temp_cwd(monkeypatch: pytest.MonkeyPatch, tmp_path: pathlib.Path):
    monkeypatch.chdir(str(tmp_path))
    yield tmp_path


@pytest.fixture()
def temp_teacher_workbook(temp_cwd, tmp_path: pathlib.Path, mocker: MockerFixture):
    source = student_teacher_gradebook._config.TEACHER_BOOK
    shutil.copy2(source, tmp_path)

    teacher_book = tmp_path / source.name

    old_value = student_teacher_gradebook._config.TEACHER_BOOK
    student_teacher_gradebook._config.TEACHER_BOOK = teacher_book
    yield teacher_book
    student_teacher_gradebook._config.TEACHER_BOOK = old_value


@pytest.fixture()
def console_runner():
    runner = click.testing.CliRunner(mix_stderr=False)
    main = functools.partial(
        runner.invoke, student_teacher_gradebook.__main__._cli, standalone_mode=False, catch_exceptions=False
    )
    return main
