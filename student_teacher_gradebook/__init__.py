import logging
import pathlib
import typing

import win32com.client as win32

from student_teacher_gradebook import _config

_MODULE_LOGGER = logging.getLogger(__name__)
_MODULE_LOGGER.addHandler(logging.NullHandler())
_StrOrPath = typing.Union[str, pathlib.Path]


class _VBA_Consts:
    xlUp = -4162


class _BaseWorkBook:
    def __init__(self, path: _StrOrPath) -> None:
        self._path = pathlib.Path(path).resolve()


def _openWorkbook(xlapp, xlfile):
    """from https://stackoverflow.com/a/39880844/8100990"""
    xlwb = xlapp.Workbooks.Open(xlfile)
    return xlwb


def _to_snake_case(value: str) -> str:
    """Split value on words and make snake case

    >>> _to_snake_case("Student Template Filename")
    'student_template_filename'
    """
    return "_".join(x.lower() for x in value.split())


class _Config(typing.NamedTuple):
    student_template_filename: str
    student_filename_format_string: str


class _MainWorkbook(_BaseWorkBook):
    def __init__(self, path: pathlib.Path) -> None:
        super().__init__(path)
        self._app = win32.gencache.EnsureDispatch("Excel.Application")
        self._app.Visible = True
        self._workbook = None
        self._config = None
        self.student_workbooks: typing.Tuple[str, ...] = ()
        self.roster: typing.Tuple[str, ...] = ()

    def __del__(self):
        self._app.Quit()

    @property
    def progress_sheet(self):
        ...

    def set_workbook(self, roster_index, value):
        ...

    def _load_config(self):
        student_name_to_workbook_mapping = {}
        config_sheet = self._workbook.Worksheets(_config.CONFIG_SHEET_NAME)
        max_rows = config_sheet.UsedRange.Rows.Count

        _config_keys = _Config._fields

        config_sheet_data = {}
        for i in range(1, max_rows + 1):
            key = (config_sheet.Cells(i, 1).Value or "").rstrip(":")
            value = config_sheet.Cells(i, 2).Value
            if not key:
                continue
            key = _to_snake_case(key)
            if key in _config_keys:
                config_sheet_data[key] = value
        try:
            self._config = _Config(**config_sheet_data)
        except TypeError:
            _MODULE_LOGGER.warning(
                "Error, all required keys must be provided on the %s sheet.",
                _config.CONFIG_SHEET_NAME,
            )
            raise

    def __enter__(self):
        self._workbook = _openWorkbook(self._app, str(self._path))
        self._load_config()
        return self

    def __exit__(self, *exc):
        if self._workbook is not None:
            _MODULE_LOGGER.info("Saving changes to main workbook")
            self._workbook.Close()
        self._workbook = None
