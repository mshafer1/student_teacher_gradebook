"""Utility for cloning student grades from a single teacher's workbook to one workbook per student."""  # noqa: W505
import contextlib
import datetime
import hashlib
import logging
import pathlib
import typing

import win32com.client as win32

from student_teacher_gradebook import _config

_MODULE_LOGGER = logging.getLogger(__name__)
_MODULE_LOGGER.addHandler(logging.NullHandler())
_StrOrPath = typing.Union[str, pathlib.Path]

_EXCEL_FIRST_ROW_OF_DATA = 1
_TABLE_OFFSET = 1


class _VBA_Consts:  # noqa: N801
    xlUp = -4162  # noqa: N815 - matching Excel convention


def _excel_column_number_to_name(column_number):
    """Type convert from integer to base 26 (A=1).

    >>> _excel_column_number_to_name(1)
    'A'

    >>> _excel_column_number_to_name(2)
    'B'

    >>> _excel_column_number_to_name(26)
    'Z'

    >>> _excel_column_number_to_name(27)
    'AA'
    """
    output = ""
    index = column_number - 1
    while index >= 0:
        character = chr((index % 26) + ord("A"))
        output = output + character
        index = index // 26 - 1

    return output[::-1]


def _excel_column_name_to_number(column_name: str):
    """Type convert from base 26 (A=1) to decimal.

    >>> _excdel_column_name_to_number('A')
    1
    >>> _excdel_column_name_to_number('B')
    2
    >>> _excdel_column_name_to_number('Z')
    26
    >>> _excdel_column_name_to_number('AA')
    27
    """
    result = 0
    for i, char in enumerate(column_name[::-1]):
        result += 26**i * (ord(char) - ord("A") + 1)
    return result


class _BaseWorkBook:
    def __init__(self, path: _StrOrPath) -> None:
        self._path = pathlib.Path(path).resolve()
        self._workbook = None
        self._app = win32.gencache.EnsureDispatch("Excel.Application")
        self._app.Visible = True

    def _workbook_must_be_opened(inner):  # noqa: N805 - class-level decorator.
        def wrapper(self, *args, **kwargs):
            if self._workbook is None:
                raise Exception()
            return inner(self, *args, **kwargs)

        return wrapper

    @staticmethod
    def _open_workbook(xlapp, xlfile):
        xlwb = xlapp.Workbooks.Open(xlfile)
        return xlwb

    def open(self):
        self._workbook = _BaseWorkBook._open_workbook(self._app, str(self._path))

    def save(self):
        self._workbook.Save()

    def close(self):
        self._workbook.Close()

    @_workbook_must_be_opened
    def set_column_range(
        self,
        sheet_name: str,
        start_row_index: int,
        column_index: typing.Union[int, str],
        values: typing.Iterable[typing.Any],
    ):
        sheet = self._workbook.Worksheets(sheet_name)
        if isinstance(column_index, str):
            column_index = _excel_column_name_to_number(column_index)
        for i, value in enumerate(values):
            sheet.Cells(start_row_index + i, column_index).Value = value

    @_workbook_must_be_opened
    def set_row_range(
        self,
        sheet_name: str,
        start_column_index: typing.Union[int, str],
        row_index: int,
        values: typing.Iterable[typing.Any],
    ):
        sheet = self._workbook.Worksheets(sheet_name)
        if isinstance(start_column_index, str):
            start_column_index = _excel_column_name_to_number(start_column_index)
        for i, value in enumerate(values):
            sheet.Cells(row_index, start_column_index + i).Value = str(value)

    @_workbook_must_be_opened
    def get_cells_value_range(
        self,
        sheet_name: str,
        start_row_index: int,
        column_index: typing.Union[int, str],
        end_row_index: typing.Optional[int] = None,
    ):
        sheet = self._workbook.Worksheets(sheet_name)
        if isinstance(column_index, str):
            column_index = _excel_column_name_to_number(column_index)
        if end_row_index is None:
            end_row_index = sheet.UsedRange.Rows.Count

        end_column_index = sheet.UsedRange.Columns.Count

        for i in range(start_row_index, end_row_index + 1):
            values = [sheet.Cells(i, j).Value for j in range(column_index, end_column_index + 1)]
            yield values

    @_workbook_must_be_opened
    def worksheet_names(self):
        for sheet in self._workbook.Worksheets:
            yield sheet.Name

    @_workbook_must_be_opened
    def remove_sheet(self, worksheet_name: str):
        self._workbook.Sheets(worksheet_name).Delete()

    @_workbook_must_be_opened
    def copy_sheet_from(self, source_workbook: pathlib.Path, sheet_index=1, *_, new_name="progress_new"):
        _temp_workbook = self._open_workbook(self._app, str(source_workbook))
        try:
            _temp_workbook.Sheets(sheet_index).Name = new_name
            _temp_workbook.Sheets(sheet_index).Copy(Before=self._workbook.Sheets(1))
        finally:
            _temp_workbook.Close(SaveChanges=False)

    @_workbook_must_be_opened
    def add_sheet(self, worksheet_name: str):
        self._workbook.Sheets.Add().Name = worksheet_name

    @_workbook_must_be_opened
    def rename_sheet(self, old_name: str, new_name: str):
        self._workbook.Sheets(old_name).Name = new_name


def _to_snake_case(value: str) -> str:
    """Split value on words and make snake case.

    >>> _to_snake_case("Student Template Filename")
    'student_template_filename'
    """
    return "_".join(x.lower() for x in value.split())


class Config(typing.NamedTuple):
    """Data model for app settings."""

    student_template_filename: str
    student_filename_format_string: str


class StudentData(typing.NamedTuple):
    """Data model for student info."""

    id_: str
    name: str
    student_file: typing.Optional[pathlib.Path]


class StudentWorkbook(_BaseWorkBook):
    """Class to load and work on a student's workbook."""

    def __init__(self, path: pathlib.Path) -> None:
        """Initialize student workbook."""
        super().__init__(path)


class MainWorkbook(_BaseWorkBook):
    """Model for interacting with the main workbook."""

    def __init__(self, path: pathlib.Path) -> None:
        """Initialize Excel app and prep to load config."""
        super().__init__(path)

        self._config = None
        self.student_workbooks: typing.Tuple[str, ...] = ()
        self._roster: typing.Tuple[StudentData, ...] = ()

    def __del__(self):
        """Make sure we close out if we haven't already."""
        self._app.Quit()

    def _load_config(self):
        config_sheet = self._workbook.Worksheets(_config.CONFIG_SHEET_NAME)
        max_rows = config_sheet.UsedRange.Rows.Count

        _config_keys = Config._fields

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
            self._config = Config(**config_sheet_data)
        except TypeError:
            _MODULE_LOGGER.warning(
                "Error, all required keys must be provided on the %s sheet.",
                _config.CONFIG_SHEET_NAME,
            )
            raise

        self._load_roster()

    def _load_roster(self):
        roster = []
        for row in self.get_cells_value_range(
            _config.ROSTER_SHEET_NAME, start_row_index=2, column_index="A"
        ):
            data = row[:3]
            if data[0] is None:
                data[0] = hashlib.md5(
                    (datetime.datetime.now().isoformat() + "-" + data[1]).encode("UTF-8")
                ).hexdigest()[:8]
            roster.append(StudentData(*data))

        self._roster = tuple(roster)

    @contextlib.contextmanager
    def open_student_workbook(self, student: StudentData):
        """Open workbook for student."""
        target_file = student.student_file
        if not pathlib.Path(target_file).is_absolute():
            target_file = _config.TEACHER_BOOK.parent / target_file
        try:
            student_workbook = StudentWorkbook(student.student_file)
            student_workbook.open()
            yield student_workbook
        except Exception as e:
            print(e)
            raise
        finally:
            student_workbook._workbook.Close()

    def get_student_values_from_sheets(self, sheet_names: typing.Iterable[str]):
        """Load student data for all sheets in sheet_names."""
        student_names = set([student.name for student in self.roster])
        student_data_mapping: typing.Dict[str, list] = {}
        for sheet_name in sheet_names:
            for row in self.get_cells_value_range(
                sheet_name=sheet_name, start_row_index=1, column_index="A"
            ):
                if row[0] in student_names and any(val is not None for val in row[1:]):
                    name = row[0]
                    values = [val or "" for val in row[1:]]
                    if name not in student_data_mapping:
                        student_data_mapping[name] = []
                    student_data_mapping[name].append([sheet_name] + values)
        return student_data_mapping

    def update_student_values(self, student_index, student: StudentData):
        """Update student in roster sheet."""
        self.set_row_range(
            _config.ROSTER_SHEET_NAME,
            "A",
            student_index + _EXCEL_FIRST_ROW_OF_DATA + _TABLE_OFFSET,
            student,
        )
        self._load_roster()

    @property
    def config(self):
        """App configuration."""
        return self._config

    @property
    def roster(self):
        """Object representing the current roster."""
        return tuple(self._roster)  # makes a copy on property read for iterating.

    @property
    def roster_as_mapping(self):
        """Student name to student in roster."""
        return {s.name: s for s in self.roster}

    def __enter__(self):
        """Open workbook and load config."""
        self.open()
        self._load_config()
        return self

    def __exit__(self, *exc):
        """After context, save and close workbook."""
        if self._workbook is not None:
            _MODULE_LOGGER.info("Saving changes to main workbook")
            self.save()
            self.close()
        self._workbook = None
