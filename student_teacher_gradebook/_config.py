import pathlib
import typing
from decouple import config

_MODULE_DIR = pathlib.Path(__file__).parent
SOURCE_DIR: pathlib.Path = config(
    "STUDENT_TEACHER_GRADEBOOK__SOURCE_DIR",
    default=_MODULE_DIR / "../source",
    cast=pathlib.Path,
)

TEACHER_BOOK: pathlib.Path = SOURCE_DIR / config(
    "STUDENT_TEACHER_GRADEBOOK__TEACHER_BOOK_FILENAME", default="TeacherBook.xlsx"
)

STUDENT_TEMPLATE: pathlib.Path = SOURCE_DIR / config(
    "STUDENT_TEACHER_GRADEBOOK__STUDENT_TEMPLATE_FILENAME",
    default="studentTemplate.xlsx",
)

CONFIG_SHEET_NAME: typing.Optional[str] = config(
    "STUDENT_TEACHER_GRADEBOOK__CONFIG_SHEET_NAME", default="Config"
)
