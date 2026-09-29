"""Linux XLSX export: preserve the template, calculate a disposable copy."""

from __future__ import annotations

import copy
import hashlib
import os
import shutil
import signal
import subprocess
import tempfile
import threading
import zipfile
from datetime import datetime
from pathlib import Path

from lxml import etree as ET

from ..models import CharacterDraft, ExcelExportResult
from ..rules import RuleEngine
from .excel import ExcelExporter, ExcelExportError, ExcelUnavailableError, _safe_filename
from .xlsx_template import TemplateWorkbook, tag, worksheet_paths


EXCEL_SLOT = threading.BoundedSemaphore(1)
FORMULA_ERRORS = {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!"}


class LinuxExcelExporter(ExcelExporter):
    def export(
        self,
        draft: CharacterDraft,
        portrait_bytes: bytes | None = None,
        *,
        generated_at: datetime | None = None,
    ) -> ExcelExportResult:
        report = RuleEngine.validate(draft)
        if not report.can_export:
            raise ExcelExportError("调查员数据未通过导出检查：" + "；".join(issue.message for issue in report.errors))
        executable = shutil.which("libreoffice") or shutil.which("soffice")
        if not executable:
            raise ExcelUnavailableError("Excel 导出需要 LibreOffice Calc。请使用项目提供的 Docker 镜像，或安装 libreoffice-calc。")
        try:
            timeout = int(os.environ.get("COC7_EXCEL_TIMEOUT", "120"))
            if timeout <= 0:
                raise ValueError
        except ValueError as exc:
            raise ExcelExportError("COC7_EXCEL_TIMEOUT 必须为正整数秒数。") from exc
        if not EXCEL_SLOT.acquire(timeout=timeout):
            raise ExcelExportError("Excel 导出繁忙，请稍后重试。")
        try:
            return self._export(draft, portrait_bytes, generated_at, executable, timeout, report)
        finally:
            EXCEL_SLOT.release()

    def _export(self, draft, portrait_bytes, generated_at, executable, timeout, report) -> ExcelExportResult:
        source_hash = hashlib.sha256(self.catalog.template_path.read_bytes()).hexdigest()
        if source_hash != self.catalog.source_sha256:
            raise ExcelExportError("模板文件在程序运行期间发生变化，已停止导出。")
        timestamp = generated_at or datetime.now()
        filename = f"COC7_{_safe_filename(draft.identity.name)}_{timestamp:%Y%m%d-%H%M%S}.xlsx"
        try:
            with tempfile.TemporaryDirectory(prefix="coc7-excel-") as directory:
                temp = Path(directory)
                incoming = temp / "input.xlsx"
                output_directory = temp / "calculated"
                output_directory.mkdir()
                workbook = TemplateWorkbook(self.catalog.template_path)
                try:
                    self._patch_template(workbook)
                    self._write_character(workbook, draft, portrait_bytes, temp)
                    # LibreOffice may trust Excel's existing caches on import,
                    # even with fullCalcOnLoad. Remove them to force calculation
                    # from the character values just written to the template.
                    for sheet in workbook.sheets.values():
                        for cell in sheet.cells.values():
                            if cell.find(tag("f")) is not None:
                                cached = cell.find(tag("v"))
                                if cached is not None:
                                    cell.remove(cached)
                                cell.attrib.pop("t", None)
                    workbook.save(incoming)
                    self._recalculate(executable, incoming, output_directory, temp / "profile", timeout)
                    self._merge_calculated_values(workbook, output_directory / incoming.name)
                    output = temp / filename
                    workbook.save(output)
                    sheet_count = self._verify_export(output)
                    data = output.read_bytes()
                finally:
                    workbook.close()
        except (ExcelExportError, ExcelUnavailableError):
            raise
        except Exception as exc:
            raise ExcelExportError(f"Excel 导出失败：{exc}") from exc
        if hashlib.sha256(self.catalog.template_path.read_bytes()).hexdigest() != source_hash:
            raise ExcelExportError("导出后检测到原模板发生变化，结果已被拒绝。")
        return ExcelExportResult(
            filename=filename,
            data=data,
            source_sha256=source_hash,
            sheet_count=sheet_count,
            warnings=tuple(issue.message for issue in report.warnings),
        )

    @staticmethod
    def _write_portrait(sheet, portrait_bytes: bytes, temp_path: Path) -> None:
        sheet.Parent.add_portrait(sheet, portrait_bytes)

    @staticmethod
    def _recalculate(executable: str, incoming: Path, output: Path, profile: Path, timeout: int) -> None:
        # Each request gets a separate profile. Never attach to a user's open
        # LibreOffice instance, and terminate the whole process group on timeout.
        command = [
            executable, f"-env:UserInstallation={profile.as_uri()}",
            "--headless", "--nologo", "--nodefault", "--nofirststartwizard",
            "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(output), str(incoming),
        ]
        process = subprocess.Popen(
            command, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace", start_new_session=True,
        )
        try:
            logs, _ = process.communicate(timeout=timeout)
        except subprocess.TimeoutExpired as exc:
            try:
                os.killpg(process.pid, signal.SIGKILL)
            except ProcessLookupError:
                pass
            process.communicate()
            raise ExcelExportError(f"Excel 公式重算超过 {timeout} 秒，请稍后重试。") from exc
        result = output / incoming.name
        if process.returncode or not result.is_file() or not zipfile.is_zipfile(result):
            raise ExcelExportError(f"LibreOffice 未能完成公式重算：{logs.strip()[-1200:] or '没有生成工作簿'}")

    @staticmethod
    def _merge_calculated_values(workbook: TemplateWorkbook, calculated_path: Path) -> None:
        # Import only cell caches; do not deliver LibreOffice's rewritten file.
        # This preserves original Excel formulas, styles, validations and charts.
        with zipfile.ZipFile(calculated_path) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        paths = worksheet_paths(parts)
        if tuple(paths) != tuple(workbook.paths):
            raise ExcelExportError("重算后的工作表数量或顺序与原模板不一致。")
        strings = []
        if "xl/sharedStrings.xml" in parts:
            strings = ["".join(node.itertext()) for node in ET.fromstring(parts["xl/sharedStrings.xml"])]
        for name, sheet in workbook.sheets.items():
            calculated = ET.fromstring(parts[paths[name]])
            cells = {cell.get("r"): cell for cell in calculated.iter(tag("c"))}
            for address, cell in sheet.cells.items():
                if cell.find(tag("f")) is None:
                    continue
                cached = cells.get(address)
                if cached is None:
                    raise ExcelExportError(f"公式缺少重算结果：{name}!{address}")
                value = cached.find(tag("v"))
                kind = cached.get("t")
                if kind == "e" or value is not None and value.text in FORMULA_ERRORS:
                    raise ExcelExportError(f"导出结果含公式错误 {value.text if value is not None else ''}：{name}!{address}")
                for child in list(cell):
                    if child.tag in {tag("v"), tag("is")}:
                        cell.remove(child)
                cell.attrib.pop("t", None)
                if kind in {"s", "inlineStr"}:
                    text = strings[int(value.text)] if kind == "s" else "".join(cached.find(tag("is")).itertext())
                    cell.set("t", "str")
                    ET.SubElement(cell, tag("v")).text = text
                elif value is not None:
                    if kind:
                        cell.set("t", kind)
                    cell.append(copy.deepcopy(value))
                elif kind == "str":
                    cell.set("t", "str")
                    ET.SubElement(cell, tag("v"))
                else:
                    raise ExcelExportError(f"公式缺少缓存值：{name}!{address}")
