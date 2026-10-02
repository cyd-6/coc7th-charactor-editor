"""Fill the user's two-page 1920s sheet, retaining its original artwork."""
from __future__ import annotations

import io
import re
import threading
from datetime import datetime
from pathlib import Path

from PIL import Image, ImageOps
from pypdf import PdfReader, PdfWriter
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

from ..models import CharacterDraft, PdfExportResult
from ..portraits import prepare_portrait
from ..rules import RuleEngine

FONT_NAME = "NotoSansSC-Card"
FONT_LOCK = threading.Lock()
PREVIEW_LOCK = threading.Lock()
TEMPLATE_PATH = Path(__file__).resolve().parents[2] / "assets/templates/1920sCha.pdf"


class PdfExportError(RuntimeError):
    pass


class PdfLayoutOverflowError(PdfExportError):
    pass


def _safe_filename(value):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value.strip()).strip(" .") or "未命名调查员"


def _register_font(path):
    with FONT_LOCK:
        if FONT_NAME not in pdfmetrics.getRegisteredFontNames():
            if not path.is_file():
                raise PdfExportError(f"找不到中文字体：{path}")
            pdfmetrics.registerFont(TTFont(FONT_NAME, str(path)))


def _string(value):
    return "" if value is None else str(value)


class ReferencePainter:
    """All input coordinates are points measured from the page's top left."""
    def __init__(self, pdf, height):
        self.pdf, self.height = pdf, height

    def text(self, value, x, baseline, width, *, size=8.5, min_size=7, align="left"):
        text = _string(value).replace("\n", " / ")
        if not text:
            return
        while size > min_size and pdfmetrics.stringWidth(text, FONT_NAME, size) > width:
            size = max(min_size, size - .25)
        if pdfmetrics.stringWidth(text, FONT_NAME, size) > width:
            raise PdfLayoutOverflowError(f"内容超出原版 PDF 栏位，请缩短后重试：{text[:70]}")
        self.pdf.setFillColor(colors.HexColor("#111111"))
        self.pdf.setFont(FONT_NAME, size)
        if align == "center":
            self.pdf.drawCentredString(x + width / 2, self.height - baseline, text)
        elif align == "right":
            self.pdf.drawRightString(x + width, self.height - baseline, text)
        else:
            self.pdf.drawString(x, self.height - baseline, text)

    def white(self, rect):
        x, y, right, bottom = rect
        self.pdf.setFillColor(colors.white)
        self.pdf.rect(x, self.height - bottom, right - x, bottom - y, fill=1, stroke=0)

    def box(self, value, rect, size=10, min_size=6.5):
        x, y, right, bottom = rect
        self.text(value, x + 1, (y + bottom) / 2 + size * .34, right - x - 2,
                  size=size, min_size=min_size, align="center")

    @staticmethod
    def _layout(value, slots, size):
        remaining, lines = value, []
        for x, baseline, width in slots:
            if not remaining:
                break
            count, measured = 0, 0.0
            while count < len(remaining) and remaining[count] != "\n":
                advance = pdfmetrics.stringWidth(remaining[count], FONT_NAME, size)
                if measured + advance > width:
                    break
                measured += advance
                count += 1
            # Keep closing punctuation with at least one preceding character.
            if count > 1 and count < len(remaining) and remaining[count] in "，。；：！？、）】》/／,.;:!?)]":
                count -= 1
            # Apply this after punctuation, so moving a semicolon does not
            # split e.g. '40；' into a trailing '4' and a leading '0；'.
            if 0 < count < len(remaining) and re.match(r"[A-Za-z0-9./_+:-]", remaining[count]):
                word = re.search(r"[A-Za-z0-9][A-Za-z0-9./_+:-]*$", remaining[:count])
                if word and word.start() > 0:
                    count = word.start()
            line, remaining = remaining[:count], remaining[count:]
            if remaining.startswith("\n"):
                remaining = remaining[1:]
            lines.append((x, baseline, width, line))
        return lines, remaining

    def flow(self, value, slots, *, size=8.5, min_size=7):
        text = _string(value).replace("\r\n", "\n").replace("\r", "\n")
        if not text:
            return ""
        while True:
            lines, remaining = self._layout(text, slots, size)
            if not remaining or size <= min_size:
                break
            size = max(min_size, size - .25)
        for x, baseline, width, line in lines:
            self.text(line, x, baseline, width, size=size, min_size=size)
        return remaining

    def circle(self, x, y, radius=5.7):
        self.pdf.setStrokeColor(colors.black)
        self.pdf.setLineWidth(.75)
        self.pdf.circle(x, self.height - y, radius, fill=0, stroke=1)


SKILL_COLUMNS = (
    ("F16", "F17", "F18", "F19", "F20", "F21", "F22", "F23", "F24", "F26", "F27", "F28", "F29", "F30", "F31"),
    ("F33", "F34", "F35", "F36", "F38", "F39", "F40", "F42", "F43", "F44", "F45", "F46", "F47", "F48", "F49"),
    tuple(f"AB{i}" for i in range(16, 31)),
    ("AB31", "AB32", "AB33", "AB34", "AB35", "AB36", "AB37", "AB38", "AB39", "AB40", None, None, None, None, None),
)
OPTIONAL_BASES = {**{f"F{i}": 5 for i in (20, 21, 22)},
                  **{f"F{i}": 1 for i in (32, 35, 36, 37, 39, 40, 41, 46, 47, 48)}, "F25": 5,
                  **{f"AB{i}": 1 for i in (27, 31, 32, 33, 42, 43, 44, 45, 46, 47, 48)}, "AB37": 10, "AB41": 5}


def _used_skill(skill):
    return (skill.template_slot not in OPTIONAL_BASES or skill.specialization.strip()
            or skill.selected_occupation or any((skill.occupation_points, skill.interest_points,
                                                skill.extra_final, skill.experience_points))
            or skill.base_value != OPTIONAL_BASES[skill.template_slot])


class PdfExporter:
    def __init__(self, font_path: str | Path, template_path: str | Path | None = None):
        self.font_path = Path(font_path).resolve()
        self.template_path = Path(template_path or TEMPLATE_PATH).resolve()
        _register_font(self.font_path)

    def export(self, draft: CharacterDraft, portrait_bytes: bytes | None = None, *,
               generated_at: datetime | None = None, include_preview: bool = True) -> PdfExportResult:
        report = RuleEngine.validate(draft)
        if not report.can_export:
            raise PdfExportError("调查员数据未通过导出检查：" + "；".join(issue.message for issue in report.errors))
        if not self.template_path.is_file():
            raise PdfExportError("找不到 1920s PDF 底版，请检查部署资源。")
        try:
            source = PdfReader(io.BytesIO(self.template_path.read_bytes()))
            if len(source.pages) != 2 or source.get_fields():
                raise PdfExportError("1920s PDF 底版应为两页静态卡片。")
        except PdfExportError:
            raise
        except Exception as exc:
            raise PdfExportError("无法读取 1920s PDF 底版。") from exc
        formula = draft.occupation.point_formula
        derived = RuleEngine.calculate(draft.attributes, draft.identity.age, formula, draft.experience.san_loss)
        timestamp = generated_at or datetime.now()
        overlay = io.BytesIO()
        pdf = canvas.Canvas(overlay, pagesize=(float(source.pages[0].mediabox.width), float(source.pages[0].mediabox.height)), pageCompression=1, invariant=1)
        notes = []
        self._draw_front(ReferencePainter(pdf, float(source.pages[0].mediabox.height)), draft, derived, portrait_bytes, notes)
        pdf.showPage()
        second = source.pages[1]
        pdf.setPageSize((float(second.mediabox.width), float(second.mediabox.height)))
        from .pdf_reference_back import draw_back
        draw_back(ReferencePainter(pdf, float(second.mediabox.height)), draft, notes)
        pdf.showPage()
        pdf.save()
        additions, writer = PdfReader(io.BytesIO(overlay.getvalue())), PdfWriter()
        for original, values in zip(source.pages, additions.pages, strict=True):
            page = writer.add_page(original)
            page.merge_page(values)
        writer.add_metadata({"/Title": f"COC7 调查员表 - {draft.identity.name}", "/Author": draft.identity.name, "/Creator": "coc7车卡器"})
        output = io.BytesIO()
        writer.write(output)
        data = output.getvalue()
        page_count = self._verify_static_pdf(data)
        previews = self.render_preview_images(data) if include_preview else ()
        return PdfExportResult(filename=f"COC7_{_safe_filename(draft.identity.name)}_1920s调查员表_{timestamp:%Y%m%d-%H%M%S}.pdf",
                               data=data, page_count=page_count, preview_images=previews)

    def _draw_front(self, p, draft, derived, portrait, notes):
        identity = draft.identity
        for value, x, baseline, width in (
            (identity.name, 114, 103, 93), (identity.player, 114, 118.7, 93),
            (draft.occupation.name, 114, 134.3, 93),
            (identity.age, 114, 150, 38), (identity.gender, 176, 150, 31),
            (identity.residence, 114, 165.6, 93), (identity.birthplace, 114, 181.3, 93),
        ):
            p.text(value, x, baseline, width, size=9, min_size=7)
        context = " · ".join(str(v) for v in (identity.era, identity.story_year, identity.current_date) if v)
        if context:
            notes.append(f"时代／日期：{context}")
        positions = {"STR": (255.5, 99.5), "CON": (255.5, 131), "SIZ": (255.5, 161),
                     "DEX": (340, 99.5), "APP": (340, 131), "EDU": (340, 161),
                     "INT": (429.5, 99.5), "POW": (429.5, 131)}
        for key, (x, y) in positions.items():
            self._triplet(p, getattr(draft.attributes, key), (x, y, x + 46, y + 26), split=25, size=13)
        p.box(derived.mov, (429.5, 161, 457.5, 187), size=13)
        if portrait:
            self._draw_portrait(p, portrait)
        by_slot = {skill.template_slot: skill for skill in draft.skills}
        mythos = by_slot.get("F27")
        for label, value, rect in (
            ("最大", derived.hp, (136, 194.5, 179.5, 210.5)),
            ("初始", derived.san, (323.5, 194.5, 370.5, 210.5)),
            ("上限", max(0, 99 - (mythos.final_value if mythos else 0)), (370.5, 194.5, 417.5, 210.5)),
            ("最大", derived.mp, (491.5, 252.5, 535.5, 268.5)),
        ):
            x, y, right, bottom = rect
            p.white((x + 5, y + 3, right - 5, bottom - 3))
            p.box(f"{label} {value}", rect, size=8.2)
        p.text(f"幸运 {draft.attributes.Luck}", 112, 284, 78, size=8, align="center")
        self._status_marks(p, derived, draft.attributes.Luck)
        self._draw_skills(p, draft, notes)
        self._draw_weapons(p, draft, notes)
        p.box(derived.damage_bonus, (528.5, 678.5, 575, 705.5), size=11)
        p.box(derived.build, (528.5, 707.5, 575, 734), size=12)
        dodge = by_slot.get("F29")
        self._triplet(p, dodge.final_value if dodge else derived.dodge, (528.5, 737, 575, 763), split=25, size=12)

    @staticmethod
    def _triplet(p, value, rect, *, split=17, size=9):
        x, y, right, bottom = rect
        half = (y + bottom) / 2
        p.box(value, (x, y, x + split, bottom), size=size)
        p.box(value // 2, (x + split, y, right, half), size=6.3, min_size=5.6)
        p.box(value // 5, (x + split, half, right, bottom), size=6.3, min_size=5.6)

    @staticmethod
    def _status_marks(p, derived, luck):
        if 0 <= derived.hp <= 20:
            if derived.hp < 6:
                x, y = 154.4 + (derived.hp % 3) * 20.8, 216.3 + (derived.hp // 3) * 13.1
            else:
                x, y = 112 + ((derived.hp - 6) % 5) * 20.8, 242.6 + ((derived.hp - 6) // 5) * 13.1
            p.circle(x, y)
        if 0 <= derived.mp <= 24:
            p.circle(472.5 + (derived.mp % 5) * 20.7, 275 + (derived.mp // 5) * 13.1)
        for value, is_luck in ((derived.san, False), (luck, True)):
            if not 0 <= value <= 99:
                continue
            if value < 8:
                if value == 0 and not is_luck:
                    continue
                x = (321 if is_luck else 472.5) + (value if is_luck else value - 1) * 14.35
                y = 286.5 if is_luck else 204.5
            else:
                x = (111.3 if is_luck else 243.5) + ((value - 8) % 23) * 14.35
                y = (297 if is_luck else 214.5) + ((value - 8) // 23) * 10.4
            p.circle(x, y, 5.5)

    def _draw_skills(self, p, draft, notes):
        skills = {skill.template_slot: skill for skill in draft.skills}
        placed, available = set(), []
        for column, slots in enumerate(SKILL_COLUMNS):
            for row, slot in enumerate(slots):
                skill = skills.get(slot)
                if skill is not None and _used_skill(skill):
                    self._skill(p, skill, column, row, replace_label=slot in OPTIONAL_BASES or slot in {"F34", "F38", "F49"})
                    placed.add(slot)
                elif slot is None or slot in OPTIONAL_BASES:
                    available.append((column, row))
        available.sort(key=lambda position: SKILL_COLUMNS[position[0]][position[1]] is not None)
        for skill in draft.skills:
            if skill.template_slot in placed or not _used_skill(skill):
                continue
            if available:
                column, row = available.pop(0)
                self._skill(p, skill, column, row, replace_label=True)
            else:
                notes.append(f"技能：{skill.display_name} {skill.final_value}/{skill.hard_value}/{skill.extreme_value}")

    def _skill(self, p, skill, column, row, *, replace_label):
        x = (175.5, 298.5, 421.5, 545)[column]
        y = 353 + row * 19.55
        self._triplet(p, skill.final_value, (x, y, x + 31.5, y + 18), size=9)
        if replace_label:
            left = x - 71
            p.white((left - .5, y + 1, x - 3, y + 17.5))
            remaining = p.flow(skill.display_name, [(left, y + 8, 67), (left, y + 16, 67)], size=7.1, min_size=6.5)
            if remaining:
                raise PdfLayoutOverflowError(f"技能名称超过原版栏位：{skill.display_name}")

    @staticmethod
    def _draw_weapons(p, draft, notes):
        lookup = {}
        for skill in draft.skills:
            for name in (skill.display_name, skill.name.rstrip("：:"), skill.specialization):
                if name:
                    lookup.setdefault(name, skill.final_value)
        widths = ((96.5, 73), (176.5, 25.5), (209, 23.5), (239, 26), (272, 45),
                  (324, 34), (364.5, 29.5), (400.5, 37.5), (445, 29))
        if not draft.weapons:
            fighting = next((s.final_value for s in draft.skills if s.template_slot == "F34"), 25)
            for value, (x, width) in zip((fighting, fighting // 2, fighting // 5), widths[1:4], strict=True):
                p.text(value, x, 685.5, width, align="center", size=8)
            return
        for index, weapon in enumerate(draft.weapons):
            baseline = 685.5 + 14.5 * index
            score = lookup.get(weapon.skill)
            values = (weapon.name, score, score // 2 if score is not None else None,
                      score // 5 if score is not None else None, weapon.damage, weapon.range,
                      weapon.attacks, weapon.ammo, weapon.malfunction)
            for value, (x, width) in zip(values, widths, strict=True):
                if index == 0:
                    p.white((x, baseline - 9, x + width, baseline + 3.3))
                p.text(value, x, baseline, width, size=8, min_size=6.5, align="left" if x == 96.5 else "center")
            extra = "；".join(v for v in (weapon.category, f"技能：{weapon.skill}" if weapon.skill else "", weapon.notes) if v)
            if extra:
                notes.append(f"{weapon.name}：{extra}")

    @staticmethod
    def _draw_portrait(p, data):
        try:
            with prepare_portrait(data) as image, image.convert("RGB") as rgb:
                with ImageOps.fit(rgb, (376, 460), method=Image.Resampling.LANCZOS) as fitted:
                    p.pdf.drawImage(ImageReader(fitted), 486.5, p.height - 190.5, width=93, height=115, mask="auto")
        except Exception as exc:
            raise PdfExportError(f"头像无法读取：{exc}") from exc

    @staticmethod
    def render_preview_images(pdf_bytes: bytes, scale: float = 1.35) -> tuple[bytes, ...]:
        try:
            import fitz
        except ImportError:
            return ()
        with PREVIEW_LOCK:
            try:
                with fitz.open(stream=pdf_bytes, filetype="pdf") as document:
                    return tuple(page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False).tobytes("png") for page in document)
            finally:
                fitz.TOOLS.store_shrink(100)

    @staticmethod
    def _verify_static_pdf(data: bytes) -> int:
        reader = PdfReader(io.BytesIO(data))
        if len(reader.pages) != 2:
            raise PdfExportError("纸质调查员表必须恰好两页。")
        if reader.get_fields() or "/AcroForm" in reader.trailer["/Root"]:
            raise PdfExportError("静态 PDF 中意外出现了表单字段。")
        for page in reader.pages:
            if any(item.get_object().get("/Subtype") == "/Widget" for item in page.get("/Annots", [])):
                raise PdfExportError("静态 PDF 中意外出现了可填写控件。")
        return 2
