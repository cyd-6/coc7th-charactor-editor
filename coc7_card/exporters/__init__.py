"""Export backends for the investigator card builder.

Both backends are imported lazily so that a missing optional dependency in one
exporter never prevents the whole application from starting.
"""

__all__ = ["ExcelExporter", "PdfExporter"]


def __getattr__(name: str):
    if name == "ExcelExporter":
        from .excel import ExcelExporter

        return ExcelExporter
    if name == "PdfExporter":
        from .pdf import PdfExporter

        return PdfExporter
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
