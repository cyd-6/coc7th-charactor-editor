from __future__ import annotations

import sys

from ..catalog import TemplateCatalog
from .excel import ExcelExporter


def excel_exporter(catalog: TemplateCatalog) -> ExcelExporter:
    if sys.platform == "win32":
        return ExcelExporter(catalog)
    from .excel_linux import LinuxExcelExporter

    return LinuxExcelExporter(catalog)
