"""pytest setup config."""
import datetime
import functools
import importlib
import pathlib
import shutil

import click.testing
import freezegun.api
import pytest
from pytest_mock.plugin import MockerFixture

import student_teacher_gradebook
import student_teacher_gradebook.__main__


@pytest.fixture()
def temp_cwd(monkeypatch: pytest.MonkeyPatch, tmp_path: pathlib.Path):
    """Set the cwd to tmp_path for the test."""
    monkeypatch.chdir(str(tmp_path))
    yield tmp_path

@pytest.fixture(autouse=True)
def use_source_dir(monkeypatch: pytest.MonkeyPatch):
    # source_dir = student_teacher_gradebook._config.MODULE_DIR / "source"
    # old_value = student_teacher_gradebook._config.SOURCE_DIR
    # student_teacher_gradebook._config.SOURCE_DIR = source_dir.resolve()
    # yield source_dir
    # student_teacher_gradebook._config.SOURCE_DIR = old_value
    source_dir = student_teacher_gradebook._config.MODULE_DIR / "../source"
    monkeypatch.setenv("STUDENT_TEACHER_GRADEBOOK__SOURCE_DIR", str(source_dir))
    importlib.reload(student_teacher_gradebook._config)
    yield source_dir

    

@pytest.fixture()
def temp_teacher_workbook(
    temp_cwd: pathlib.Path, tmp_path: pathlib.Path, fixed_datetime, use_source_dir
):
    """Copy the source workbook to tmp_path, set cwd, and patch the config to use this."""
    source = student_teacher_gradebook._config.TEACHER_BOOK
    shutil.copy2(source, tmp_path)

    teacher_book = tmp_path / source.name

    old_value = student_teacher_gradebook._config.TEACHER_BOOK
    student_teacher_gradebook._config.TEACHER_BOOK = teacher_book
    yield teacher_book
    student_teacher_gradebook._config.TEACHER_BOOK = old_value


@pytest.fixture()
def console_runner():
    """Fixture that provides a convenience wrapper around click.testing.CliRunner.invoke.

    The wrapper sets the following.
    mix_stderr=False / allowing for stdout and stderr to be checked seperately
    standalone_mode=False / don't sys.exit on an issue
    catch_exceptions=False / let Pytest handle errors.

    The final return uses functools.partial to call 'student_teacher_gradebook.__main__._cli'
    This allows for testing the cli by calling this fixture with intended args.
    """
    runner = click.testing.CliRunner(mix_stderr=False)
    main = functools.partial(
        runner.invoke,
        student_teacher_gradebook.__main__._cli,
        standalone_mode=False,
        catch_exceptions=False,
    )
    return main


@pytest.fixture
def fixed_datetime(freezer: freezegun.api.FrozenDateTimeFactory):
    """Wrap the freezer fixture with a default time (01-01-2000 00:00)."""
    freezer.move_to(datetime.datetime(2000, 1, 1))
    yield freezer
