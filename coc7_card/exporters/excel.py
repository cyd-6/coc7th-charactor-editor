from __future__ import annotations

import hashlib
import re
import shutil
import tempfile
import threading
from datetime import datetime
from pathlib import Path

from ..branch_workbook import write_branch_section
from ..catalog import TemplateCatalog, split_occupation_skill_token
from ..models import CharacterDraft, ExcelExportResult
from ..portraits import prepare_portrait
from ..rules import RuleEngine
from ..skill_specializations import is_branch_slot


FORMULA_ERRORS = {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!"}
WINDOWS_EXCEL_SLOT = threading.BoundedSemaphore(1)


class ExcelUnavailableError(RuntimeError):
    pass


class ExcelExportError(RuntimeError):
    pass


def _safe_filename(value: str) -> str:
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value.strip())
    return text.strip(" .") or "未命名调查员"


def _set_cell_value(cell, value) -> None:
    # COM interprets strings (e.g. =1+1 or 00123) while OOXML already stores
    # them as literal text. Excel's quote prefix is not part of the saved value.
    if isinstance(value, str) and value and hasattr(cell, '_oleobj_'):
        value = "'" + value
    cell.Value2 = value


class ExcelExporter:
    def __init__(self, catalog: TemplateCatalog):
        self.catalog = catalog

    def export(
        self,
        draft: CharacterDraft,
        portrait_bytes: bytes | None = None,
        *,
        generated_at: datetime | None = None,
    ) -> ExcelExportResult:
        report = RuleEngine.validate(draft)
        if not report.can_export:
            messages = "；".join(issue.message for issue in report.errors)
            raise ExcelExportError(f"调查员数据未通过导出检查：{messages}")

        if not WINDOWS_EXCEL_SLOT.acquire(timeout=120):
            raise ExcelExportError("Excel 导出繁忙，请稍后重试。")
        try:
            return self._export_windows(draft, portrait_bytes, generated_at, report)
        finally:
            WINDOWS_EXCEL_SLOT.release()

    def _export_windows(self, draft, portrait_bytes, generated_at, report):
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise ExcelUnavailableError(
                "完整 Excel 导出需要 Windows 桌面版 Microsoft Excel 和 pywin32。"
                "请使用项目虚拟环境执行 python -m pip install -r requirements.txt。"
            ) from exc

        source_hash = hashlib.sha256(self.catalog.template_path.read_bytes()).hexdigest()
        if source_hash != self.catalog.source_sha256:
            raise ExcelExportError("模板文件在程序运行期间发生变化，已停止导出以保护原文件。")

        timestamp = generated_at or datetime.now()
        filename = f"COC7_{_safe_filename(draft.identity.name)}_{timestamp:%Y%m%d-%H%M%S}.xlsx"
        excel = None
        workbook = None
        temp_manager = tempfile.TemporaryDirectory(prefix="coc7-card-", ignore_cleanup_errors=True)
        temp_path = Path(temp_manager.name)
        output_path = temp_path / filename
        pythoncom.CoInitialize()
        try:
            shutil.copy2(self.catalog.template_path, output_path)

            try:
                excel = win32com.client.DispatchEx("Excel.Application")
            except Exception as exc:
                raise ExcelUnavailableError(
                    "无法启动 Windows 桌面版 Microsoft Excel。请确认已安装并完成首次启动，"
                    "能够手动打开 XLSX；PDF 导出不需要 Excel。"
                ) from exc
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False
            excel.EnableEvents = False
            excel.ScreenUpdating = False
            excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
            workbook = excel.Workbooks.Open(
                str(output_path),
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
                CorruptLoad=0,
            )
            excel.Calculation = -4135  # xlCalculationManual, only this isolated instance.
            self._write_character(workbook, draft, portrait_bytes, temp_path)
            excel.CalculateFullRebuild()
            # Deliver an automatically recalculating workbook to the user.
            excel.Calculation = -4105  # xlCalculationAutomatic
            workbook.Save()
            workbook.Close(SaveChanges=False)
            workbook = None
            excel.Quit()
            excel = None

            data = output_path.read_bytes()
            sheet_count = self._verify_export(output_path)
        except (ExcelUnavailableError, ExcelExportError):
            raise
        except Exception as exc:  # COM reports many environment-specific exception types.
            raise ExcelExportError(f"Excel 导出失败：{exc}") from exc
        finally:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            workbook = None
            excel = None
            pythoncom.CoUninitialize()
            temp_manager.cleanup()

        if hashlib.sha256(self.catalog.template_path.read_bytes()).hexdigest() != source_hash:
            raise ExcelExportError("导出后检测到原模板发生变化，结果已被拒绝。")
        return ExcelExportResult(
            filename=filename,
            data=data,
            source_sha256=source_hash,
            sheet_count=sheet_count,
            warnings=tuple(issue.message for issue in report.warnings),
        )

    def _write_character(
        self,
        workbook,
        draft: CharacterDraft,
        portrait_bytes: bytes | None,
        temp_path: Path,
    ) -> None:
        sheet = workbook.Worksheets("人物卡")
        identity = draft.identity
        occupation = draft.occupation
        formula = occupation.point_formula
        derived = RuleEngine.calculate(draft.attributes, identity.age, formula, draft.experience.san_loss)

        values = {
            "E3": identity.name,
            "E4": identity.player,
            "E5": occupation.name,
            "M4": identity.era,
            "E6": int(identity.age),
            "M6": identity.gender,
            "E7": identity.residence,
            "M7": identity.birthplace,
            "U3": draft.attributes.STR,
            "U5": draft.attributes.CON,
            "U7": draft.attributes.SIZ,
            "AA3": draft.attributes.DEX,
            "AA5": draft.attributes.APP,
            "AA7": draft.attributes.INT,
            "AG3": draft.attributes.POW,
            "AG5": draft.attributes.EDU,
            "AG7": draft.attributes.Luck,
            "E10": derived.hp,
            "N10": derived.san,
            "W10": derived.mp,
        }
        for cell, value in values.items():
            _set_cell_value(sheet.Range(cell), value)

        self._write_date(sheet, identity.current_date)
        if occupation.is_custom:
            sheet.Range("M5").Value2 = 1
            self._write_custom_occupation(workbook, draft)
        else:
            sheet.Range("M5").Value2 = occupation.occupation_id

        definitions = {definition.template_slot: definition for definition in self.catalog.skills}
        for skill in draft.skills:
            cell_map = self.catalog.skill_cell(skill.template_slot)
            if cell_map is None:
                continue
            _set_cell_value(sheet.Range(cell_map.name_cell), skill.name)
            if cell_map.specialization_cell:
                _set_cell_value(sheet.Range(cell_map.specialization_cell), skill.specialization)
            definition = definitions.get(skill.template_slot)
            expected_dynamic_base = None
            if definition and definition.base_formula:
                expected_dynamic_base = RuleEngine.skill_base_value(
                    definition.base_formula,
                    draft.attributes,
                    definition.base_value,
                )
            if expected_dynamic_base is None or int(skill.base_value) != expected_dynamic_base:
                sheet.Range(cell_map.base_cell).Value2 = int(skill.base_value)
            sheet.Range(cell_map.extra_cell).Value2 = int(skill.extra_final)
            sheet.Range(cell_map.experience_cell).Value2 = int(skill.experience_points)
            if "克苏鲁神话" not in skill.name:
                sheet.Range(cell_map.occupation_cell).Value2 = int(skill.occupation_points)
                sheet.Range(cell_map.interest_cell).Value2 = int(skill.interest_points)

        write_branch_section(workbook, draft, _set_cell_value)
        self._write_background(sheet, draft)
        experience = draft.experience
        _set_cell_value(sheet.Range("F113"), experience.name or "无")
        if experience.name:
            _set_cell_value(sheet.Range("BC26"), experience.name)
            sheet.Range("BC28").Value2 = experience.san_loss
            sheet.Range("BG28").Value2 = experience.skill_points
            _set_cell_value(sheet.Range("BC30"), experience.notes)
        self._write_assets(sheet, draft)
        self._write_weapons(sheet, draft)
        self._write_inventory(sheet, draft)
        if portrait_bytes:
            self._write_portrait(sheet, portrait_bytes, temp_path)

    @staticmethod
    def _write_date(sheet, value: str) -> None:
        if not value:
            return
        try:
            parsed = datetime.fromisoformat(value)
        except ValueError:
            return
        sheet.Range("G8").Value2 = parsed.year
        _set_cell_value(sheet.Range("J8"), f"{parsed.month}月")
        _set_cell_value(sheet.Range("L8"), f"{parsed.day}日")
        _set_cell_value(sheet.Range("N8"), parsed.strftime("%H：%M"))

    @staticmethod
    def _write_custom_occupation(workbook, draft: CharacterDraft) -> None:
        occupation = draft.occupation
        sheet = workbook.Worksheets("职业列表")
        _set_cell_value(sheet.Range("C3"), occupation.name)
        _set_cell_value(sheet.Range("D3"), occupation.credit_range_text)
        sheet.Range("F3").Formula = f"={occupation.point_formula}"
        _set_cell_value(sheet.Range("G3"), occupation.summary)
        _set_cell_value(sheet.Range("K3"), occupation.contacts)
        _set_cell_value(sheet.Range("M3"), occupation.description)
        # Credit is always selected separately in the Web UI. It must not take
        # one of the eight editable occupation slots needed for XLSX round trips.
        selected = [skill for skill in draft.skills if skill.selected_occupation and skill.template_slot != "F26"]
        for offset in range(8):
            cell = f"I{3 + offset}"
            if offset < len(selected):
                token_name, _ = split_occupation_skill_token(selected[offset].name)
                if is_branch_slot(selected[offset].template_slot):
                    token_name = selected[offset].display_name
                _set_cell_value(sheet.Range(cell), token_name)
            else:
                sheet.Range(cell).MergeArea.ClearContents()
        sheet.Range("J11").Value2 = max(0, 8 - min(8, len(selected)))

    @staticmethod
    def _write_background(sheet, draft: CharacterDraft) -> None:
        background = draft.background
        mapping = {
            "AA61": background.appearance,
            "AA63": background.beliefs,
            "AA65": background.significant_people,
            "AA67": background.meaningful_places,
            "AA69": background.treasured_possessions,
            "AA71": background.traits,
            "AA73": background.scars,
            "AA75": background.phobias_manias,
            "W77": background.personal_story,
        }
        for cell, value in mapping.items():
            _set_cell_value(sheet.Range(cell), value)

        key_rows = {
            "形象描述": 61,
            "思想与信念": 63,
            "重要之人": 65,
            "意义非凡之地": 67,
            "宝贵之物": 69,
            "特质": 71,
            "伤口和疤痕": 73,
            "恐惧症和躁狂症": 75,
        }
        for row in key_rows.values():
            sheet.Range(f"AR{row}").Value2 = "☐"
        for connection in background.key_connection.split("、"):
            selected_row = key_rows.get(connection.strip())
            if selected_row:
                sheet.Range(f"AR{selected_row}").Value2 = "☒"

    @staticmethod
    def _write_assets(sheet, draft: CharacterDraft) -> None:
        assets = draft.assets
        # Separate USD inputs and exchange assumptions from the existing card
        # layout. No row/column insertion or merging is performed.
        fx = sheet.Parent.Worksheets("货币汇率")
        if not fx.PageSetup.PrintArea:
            fx.PageSetup.PrintArea = "$A$1:$H$31"
        # The template owns the annual lookup, source notes and validations.
        # Only input cells are changed here so exported workbooks stay editable.
        fx.Range("K1").Value2 = assets.exchange_year
        if str(fx.Range("K2").Formula).replace("'", "").replace("$", "") == "=人物卡!S62":
            _set_cell_value(sheet.Range("S62"), assets.currency)
        else:
            _set_cell_value(fx.Range("K2"), assets.currency)
        credit = "'人物卡'!$R$26"
        base_formulas = [
            f'LOOKUP({credit},{{0,1,10,50,90,99}},{{0.5,2,10,50,250,5000}})',
            f'IF({credit}=0,0.5,IF({credit}=99,50000,{credit}*LOOKUP({credit},{{1,10,50,90}},{{1,2,5,20}})))',
            f'IF({credit}=0,0,IF({credit}=99,5000000,{credit}*LOOKUP({credit},{{1,10,50,90}},{{10,50,500,2000}})))',
        ]
        for row, value, auto, formula in zip((9, 10, 11),
                (assets.usd_spending, assets.usd_cash, assets.usd_assets),
                (assets.auto_spending, assets.auto_cash, assets.auto_assets), base_formulas):
            if auto:
                fx.Range(f"K{row}").Formula = f'={formula}*IF($K$1=2026,20,1)'
            else:
                fx.Range(f"K{row}").Value2 = float(value)
        # Fixed conversion formulas and labels belong to the prepared template.
        # Keep the template's General rate format. Some localized Excel COM
        # installations reject the English "General" format string.
        for cell in ("I62", "O62", "L62", "B75", "F75", "J75", "N75", "R75"):
            sheet.Range(cell).MergeArea.NumberFormat = "#,##0.00"
            sheet.Range(cell).MergeArea.ShrinkToFit = True
        sheet.Range("S62").MergeArea.ShrinkToFit = True
        sheet.Range("O61").Value2 = "当前现金"
        mapping = {
            "F62": assets.living_standard,
            "L63": assets.asset_description,
            "B70": assets.vehicles,
            "F70": assets.residence,
            "J70": assets.luxuries,
            "N70": assets.securities,
            "R70": assets.other,
        }
        for cell, value in mapping.items():
            if value != "":
                _set_cell_value(sheet.Range(cell), value)

    @staticmethod
    def _write_weapons(sheet, draft: CharacterDraft) -> None:
        skill_values = {skill.display_name: skill.final_value for skill in draft.skills}
        skill_values.update({skill.name.rstrip("：:"): skill.final_value for skill in draft.skills})
        if draft.weapons:
            for row in range(53, 59):
                for cell in ("B", "G", "M", "Q", "W", "AA", "AC", "AE", "AG", "AJ"):
                    sheet.Range(f"{cell}{row}").MergeArea.ClearContents()
        for row, weapon in zip(range(53, 59), draft.weapons, strict=False):
            values = {
                f"B{row}": weapon.name,
                f"G{row}": weapon.category,
                f"M{row}": weapon.skill,
                f"Q{row}": skill_values.get(weapon.skill, ""),
                f"W{row}": weapon.damage,
                f"AA{row}": weapon.range,
                f"AE{row}": weapon.attacks,
                f"AG{row}": weapon.ammo,
                f"AJ{row}": weapon.malfunction,
            }
            for cell, value in values.items():
                _set_cell_value(sheet.Range(cell), value)

    @staticmethod
    def _write_inventory(sheet, draft: CharacterDraft) -> None:
        for row, item in zip(range(79, 94), draft.inventory, strict=False):
            _set_cell_value(sheet.Range(f"B{row}"), item.status)
            _set_cell_value(sheet.Range(f"D{row}"), item.location)
            _set_cell_value(sheet.Range(f"F{row}"), item.name)
            _set_cell_value(sheet.Range(f"N{row}"), item.backpack_slot)

    @staticmethod
    def _write_portrait(sheet, portrait_bytes: bytes, temp_path: Path) -> None:
        portrait_path = temp_path / "portrait.png"
        with prepare_portrait(portrait_bytes) as converted:
            converted.save(portrait_path, format="PNG")
            image_width, image_height = converted.size

        target = sheet.Range("AL3:AS9")
        target_width, target_height = float(target.Width), float(target.Height)
        source_ratio = image_width / max(image_height, 1)
        target_ratio = target_width / max(target_height, 1)
        if source_ratio > target_ratio:
            width = target_width
            height = width / source_ratio
        else:
            height = target_height
            width = height * source_ratio
        left = float(target.Left) + (target_width - width) / 2
        top = float(target.Top) + (target_height - height) / 2
        sheet.Shapes.AddPicture(str(portrait_path), False, True, left, top, width, height)

    def _verify_export(self, output_path: Path) -> int:
        from .xlsx_verify import verify_workbook

        return verify_workbook(output_path, self.catalog.sheet_names,
                               getattr(self.catalog, 'template_revision_note', None))
