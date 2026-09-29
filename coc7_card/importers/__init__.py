"""Read investigator documents without executing embedded content."""

from .excel import ExcelImportError, import_investigator
from .attributes import import_attributes

__all__ = ["ExcelImportError", "import_investigator", "import_attributes"]
