

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


class _MainWorkbook(_BaseWorkBook):
    def __init__(self, path: pathlib.Path) -> None:
        super().__init__(path)
        self._app = win32.gencache.EnsureDispatch('Excel.Application')
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
        
    def __enter__(self):
        self._workbook = _openWorkbook(self._app, str(self._path))
        self._load_config()
        return self

    def __exit__(self, *exc):
        if self._workbook is not None:
            _MODULE_LOGGER.info("Saving changes to main workbook")
            self._workbook.Close()
        self._workbook = None
