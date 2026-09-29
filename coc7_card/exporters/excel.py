from __future__ import annotations

import hashlib
import io
import re
import shutil
import tempfile
import warnings
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from PIL import Image

from ..catalog import TemplateCatalog, split_occupation_skill_token
from ..models import CharacterDraft, ExcelExportResult
from ..rules import RuleEngine


FORMULA_ERRORS = {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!"}


class ExcelUnavailableError(RuntimeError):
    pass


class ExcelExportError(RuntimeError):
    pass


def _safe_filename(value: str) -> str:
    text = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value.strip())
    return text.strip(" .") or "未命名调查员"


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

        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise ExcelUnavailableError(
                "完整 Excel 导出需要 Windows 桌面版 Microsoft Excel 和 pywin32。请运行 run.bat 安装依赖。"
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

            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False
            excel.EnableEvents = False
            workbook = excel.Workbooks.Open(
                str(output_path),
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
            )
            self._write_character(workbook, draft, portrait_bytes, temp_path)
            excel.CalculateFullRebuild()
            workbook.Save()
            workbook.Close(SaveChanges=True)
            workbook = None
            excel.Quit()
            excel = None

            data = output_path.read_bytes()
            sheet_count = self._verify_export(output_path)
        except ExcelUnavailableError:
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
            sheet.Range(cell).Value2 = value

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
            sheet.Range(cell_map.name_cell).Value2 = skill.name
            if cell_map.specialization_cell:
                sheet.Range(cell_map.specialization_cell).Value2 = skill.specialization
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

        self._write_background(sheet, draft)
        experience = draft.experience
        sheet.Range("F113").Value2 = experience.name or "无"
        if experience.name:
            sheet.Range("BC26").Value2 = experience.name
            sheet.Range("BC28").Value2 = experience.san_loss
            sheet.Range("BG28").Value2 = experience.skill_points
            sheet.Range("BC30").Value2 = experience.notes
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
        sheet.Range("J8").Value2 = f"{parsed.month}月"
        sheet.Range("L8").Value2 = f"{parsed.day}日"
        sheet.Range("N8").Value2 = parsed.strftime("%H：%M")

    @staticmethod
    def _write_custom_occupation(workbook, draft: CharacterDraft) -> None:
        occupation = draft.occupation
        sheet = workbook.Worksheets("职业列表")
        sheet.Range("C3").Value2 = occupation.name
        sheet.Range("D3").Value2 = occupation.credit_range_text
        sheet.Range("F3").Formula = f"={occupation.point_formula}"
        sheet.Range("G3").Value2 = occupation.summary
        sheet.Range("K3").Value2 = occupation.contacts
        sheet.Range("M3").Value2 = occupation.description
        # Credit is always selected separately in the Web UI. It must not take
        # one of the eight editable occupation slots needed for XLSX round trips.
        selected = [skill for skill in draft.skills if skill.selected_occupation and skill.template_slot != "F26"]
        for offset in range(8):
            cell = f"I{3 + offset}"
            if offset < len(selected):
                token_name, _ = split_occupation_skill_token(selected[offset].name)
                sheet.Range(cell).Value2 = token_name
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
            sheet.Range(cell).Value2 = value

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
        fx.Range("K2").Value2 = assets.currency
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
        sheet.Range("BO37").NumberFormat = "General"
        for cell in ("I62", "O62", "L62", "B75", "F75", "J75", "N75", "R75"):
            sheet.Range(cell).NumberFormat = "#,##0.00"
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
                sheet.Range(cell).Value2 = value

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
                sheet.Range(cell).Value2 = value

    @staticmethod
    def _write_inventory(sheet, draft: CharacterDraft) -> None:
        for row, item in zip(range(79, 94), draft.inventory, strict=False):
            sheet.Range(f"B{row}").Value2 = item.status
            sheet.Range(f"D{row}").Value2 = item.location
            sheet.Range(f"F{row}").Value2 = item.name
            sheet.Range(f"N{row}").Value2 = item.backpack_slot

    @staticmethod
    def _write_portrait(sheet, portrait_bytes: bytes, temp_path: Path) -> None:
        portrait_path = temp_path / "portrait.png"
        with Image.open(io.BytesIO(portrait_bytes)) as image:
            converted = image.convert("RGBA")
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
        with warnings.catch_warnings():
            warnings.filterwarnings(
                "ignore",
                message="Data Validation extension is not supported and will be removed",
            )
            # Normal mode materializes every sheet while the warning filter is
            # active. The books remain verification-only and are never saved.
            formula_book = load_workbook(output_path, read_only=False, data_only=False)
            value_book = load_workbook(output_path, read_only=False, data_only=True)
        try:
            if tuple(formula_book.sheetnames) != self.catalog.sheet_names:
                raise ExcelExportError("导出工作簿的工作表数量或顺序与原模板不一致。")
            if formula_book["更新说明"]["P3"].value not in (None, ""):
                raise ExcelExportError("导出工作簿仍含模板版本说明：更新说明!P3")
            expected = {
                ("简化卡 骰娘导入", "T29"): '=IF(W29=0,"",人物卡!AB42)',
                ("简化卡 骰娘导入", "T34"): '=IF(W34=0,"",人物卡!AB47)',
            }
            for (sheet_name, cell), formula in expected.items():
                actual = str(formula_book[sheet_name][cell].value or "").replace("'", "")
                if actual != formula:
                    raise ExcelExportError(f"模板兼容性修复未写入：{sheet_name}!{cell}")
            for cell in ("AT15", "AT16", "AT17"):
                actual = str(formula_book["附表"][cell].value or "").replace("'", "")
                if "附表!$V$226" not in actual or "#REF!" in actual:
                    raise ExcelExportError(f"模板兼容性修复未写入：附表!{cell}")
            for sheet in formula_book.worksheets:
                for row in sheet.iter_rows():
                    for cell in row:
                        value = cell.value
                        if not isinstance(value, str) or not value.startswith("="):
                            continue
                        if "#REF!" in value:
                            raise ExcelExportError(f"导出结果仍含断裂引用：{sheet.title}!{cell.coordinate}")
                        calculated = value_book[sheet.title][cell.coordinate].value
                        if calculated in FORMULA_ERRORS:
                            raise ExcelExportError(
                                f"导出结果含公式错误 {calculated}：{sheet.title}!{cell.coordinate}"
                            )
            return len(formula_book.sheetnames)
        finally:
            formula_book.close()
            value_book.close()
