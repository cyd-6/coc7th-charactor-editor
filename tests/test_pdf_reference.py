from __future__ import annotations

import hashlib
import io
from datetime import datetime
from pathlib import Path

import fitz
import pytest
from PIL import Image, ImageChops
from pypdf import PdfReader

from app import FONT_PATH, get_catalog
from coc7_card.exporters.pdf import PdfExporter, PdfLayoutOverflowError
from coc7_card.models import InventoryItem, Weapon
from coc7_card.web import build_draft


REFERENCE = Path(__file__).resolve().parents[1] / "assets/templates/1920sCha.pdf"


@pytest.fixture
def draft():
    catalog = get_catalog()
    occupation = next(item for item in catalog.occupations if item.credit_min == 0)
    return build_draft({
        "identity": {
            "name": "林知秋", "player": "顾言", "age": 32, "gender": "女",
            "era": "1920s", "story_year": 1920, "current_date": "1920-04-12",
            "birthplace": "上海", "residence": "阿卡姆",
        },
        "attributes": {"STR": 60, "CON": 65, "SIZ": 55, "DEX": 70, "APP": 50,
                       "INT": 75, "POW": 60, "EDU": 70, "Luck": 55},
        "occupation_id": occupation.occupation_id,
        "skills": [
            {"template_slot": "F20", "specialization": "摄影", "interest_points": 20,
             "extra_final": 7, "experience_points": 8},
            {"template_slot": "AB48", "specialization": "异闻辨析", "interest_points": 13},
        ],
        "experience": {"selection": "custom", "name": "山村调查", "skill_points": 10,
                       "san_loss": 2, "notes": "曾见红月"},
        "background": {
            "appearance": "短发灰衣", "beliefs": "寻找事实", "significant_people": "老师周远",
            "meaningful_places": "旧城码头", "treasured_possessions": "祖父怀表", "traits": "谨慎坚毅",
            "scars": "左腕旧疤", "phobias_manias": "害怕深海", "personal_story": "寻找故友留下的信",
            "key_connection": "失踪的妹妹", "contacts_notes": "联络记者沈言",
        },
        "assets": {
            "currency_code": "GBP", "usd_spending": "0", "usd_cash": "0", "usd_assets": "0",
            "living_standard": "节俭", "asset_description": "继承祖产", "vehicles": "旧自行车",
            "residence": "沿河小屋", "luxuries": "银质袖扣", "securities": "铁路债券", "other": "古董书架",
        },
        "weapons": [{"name": "小刀", "category": "冷兵器", "skill": "斗殴", "damage": "1D4+DB",
                     "range": "接触", "attacks": "1", "ammo": "0", "malfunction": "0", "notes": "刻有家徽"}],
        "inventory": [{"name": "调查笔记", "status": "携带", "location": "外套口袋", "backpack_slot": "甲一"}],
    }, catalog)


def export_pdf(draft, portrait=None):
    return PdfExporter(FONT_PATH).export(
        draft, portrait, generated_at=datetime(2026, 9, 29, 12, 0), include_preview=False,
    )


def compact_text(data):
    reader = PdfReader(io.BytesIO(data))
    return "".join("".join(page.extract_text() for page in reader.pages).split())


def test_wrapped_punctuation_keeps_multi_digit_values_together():
    from reportlab.pdfbase import pdfmetrics
    from coc7_card.exporters.pdf import FONT_NAME, ReferencePainter
    PdfExporter(FONT_PATH)
    width = pdfmetrics.stringWidth("甲乙40", FONT_NAME, 8) + .01
    lines, remaining = ReferencePainter._layout("甲乙40；丙", [(0, 10, width), (0, 20, 100)], 8)
    assert [line[3] for line in lines] == ["甲乙", "40；丙"]
    assert not remaining


def test_reference_pages_preserve_size_static_state_and_original_source(draft):
    source_hash = hashlib.sha256(REFERENCE.read_bytes()).hexdigest()
    result = export_pdf(draft)
    original, output = PdfReader(REFERENCE), PdfReader(io.BytesIO(result.data))
    assert len(output.pages) == result.page_count == len(original.pages) == 2
    assert not output.get_fields()
    assert "/AcroForm" not in output.trailer["/Root"]
    for page, reference in zip(output.pages, original.pages, strict=True):
        assert tuple(float(value) for value in page.mediabox) == pytest.approx(tuple(float(value) for value in reference.mediabox), abs=0.01)
        assert tuple(float(value) for value in page.cropbox) == pytest.approx(tuple(float(value) for value in reference.cropbox), abs=0.01)
        assert page.get("/Rotate", 0) == reference.get("/Rotate", 0)
        assert all(annotation.get_object().get("/Subtype") != "/Widget" for annotation in page.get("/Annots", []))
    assert hashlib.sha256(REFERENCE.read_bytes()).hexdigest() == source_hash


def test_reference_artwork_retains_original_resolution_and_appearance(draft):
    result = export_pdf(draft)
    with fitz.open(REFERENCE) as source, fitz.open(stream=result.data, filetype="pdf") as output:
        for original, actual in zip(source, output, strict=True):
            # The source artwork is a 2550 x 3300 raster image; preserve its
            # native pixels, along with the PDF's remaining vector marks.
            original_images = original.get_images(full=True)
            actual_images = actual.get_images(full=True)
            assert any(image[2:4] == (2550, 3300) for image in original_images)
            actual_native = [image for image in actual_images if image[2:4] == (2550, 3300)]
            assert actual_native
            expected_pixels = {hashlib.sha256(fitz.Pixmap(source, image[0]).samples).hexdigest()
                               for image in original_images if image[2:4] == (2550, 3300)}
            actual_pixels = {hashlib.sha256(fitz.Pixmap(output, image[0]).samples).hexdigest()
                             for image in actual_native}
            assert expected_pixels <= actual_pixels
            assert len(actual.get_drawings()) >= len(original.get_drawings())
            for region in [(40, 40, 85, 185), (48, 739, 121, 796)]:
                clip = fitz.Rect(region)
                expected_pix = original.get_pixmap(matrix=fitz.Matrix(2, 2), clip=clip, alpha=False)
                actual_pix = actual.get_pixmap(matrix=fitz.Matrix(2, 2), clip=clip, alpha=False)
                expected = Image.frombytes("RGB", (expected_pix.width, expected_pix.height), expected_pix.samples)
                actual_image = Image.frombytes("RGB", (actual_pix.width, actual_pix.height), actual_pix.samples)
                assert expected.getextrema()[0][0] < 240  # The clip contains real ornamentation.
                difference = ImageChops.difference(expected, actual_image)
                pixels = difference.tobytes()
                changed = sum(max(pixels[index:index + 3]) > 8 for index in range(0, len(pixels), 3))
                assert changed / (difference.width * difference.height) < 0.001


def test_reference_export_retains_investigator_and_supplemental_details(draft):
    text = compact_text(export_pdf(draft).data)
    expected = [
        "林知秋", "顾言", "上海", "阿卡姆", "1920-04-12", "摄影", "异闻辨析",
        "短发灰衣", "寻找事实", "老师周远", "旧城码头", "祖父怀表", "谨慎坚毅",
        "左腕旧疤", "害怕深海", "寻找故友留下的信", "失踪的妹妹", "联络记者沈言",
        "山村调查", "曾见红月", "继承祖产", "旧自行车", "沿河小屋", "银质袖扣", "铁路债券", "古董书架",
        "小刀", "冷兵器", "斗殴", "1D4+DB", "接触", "刻有家徽",
        "调查笔记", "携带", "外套口袋", "甲一",
        draft.assets.currency, draft.assets.exchange_source, draft.assets.exchange_basis, draft.assets.exchange_date,
    ]
    for value in expected:
        assert "".join(value.split()) in text, f"PDF lost a supplied value: {value}"
    for value in draft.attributes.as_mapping().values():
        assert str(value) in text


def test_reference_zero_money_damage_bonus_and_build_are_printed(draft):
    result = export_pdf(draft)
    assert compact_text(result.data).count("0.00") >= 3
    with fitz.open(stream=result.data, filetype="pdf") as output:
        # STR + SIZ = 115 gives exactly zero for damage bonus and build.
        # Inspect the two individual output boxes, rather than accepting an
        # unrelated "0" elsewhere on the printed score tracks.
        assert output[0].get_textbox(fitz.Rect(527, 677, 575, 705)).strip() == "0"
        assert output[0].get_textbox(fitz.Rect(527, 705, 575, 734)).strip() == "0"


def test_reference_portrait_occupies_original_portrait_frame(draft):
    portrait = io.BytesIO()
    Image.new("RGB", (80, 120), (76, 112, 132)).save(portrait, format="PNG")
    result = export_pdf(draft, portrait.getvalue())
    with fitz.open(REFERENCE) as original, fitz.open(stream=result.data, filetype="pdf") as output:
        assert len(output[0].get_images(full=True)) > len(original[0].get_images(full=True))
        pix = output[0].get_pixmap(clip=fitz.Rect(532, 127, 536, 131), alpha=False)
        assert pix.pixel(pix.width // 2, pix.height // 2) == (76, 112, 132)


def test_reference_legal_weapon_and_inventory_capacity_has_no_dropped_rows(draft):
    draft.weapons = [Weapon(name=f"样本兵器{i}", skill="斗殴", damage="1D4", range="接触", attacks="1") for i in range(1, 7)]
    draft.inventory = [InventoryItem(name=f"随身物件{i}", status="携带", location="背包", backpack_slot=str(i)) for i in range(1, 16)]
    text = compact_text(export_pdf(draft).data)
    for weapon in draft.weapons:
        assert weapon.name in text
    for item in draft.inventory:
        assert item.name in text


@pytest.mark.parametrize("slot,field,value", [
    ("F25", "interest_points", 13),       # Modern skill omitted from the printed 1920s labels.
    ("F32", "base_value", 30),           # Imported edited base, with no points allocated.
    ("AB41", "selected_occupation", True),  # Selected occupation skill at its default base.
    ("F37", "extra_final", 2),           # Fourth melee specialty beyond the printed slots.
    ("F41", "experience_points", 2),     # Fourth ranged specialty funded by an experience package.
])
def test_reference_preserves_used_optional_skills(draft, slot, field, value):
    skill = next(skill for skill in draft.skills if skill.template_slot == slot)
    setattr(skill, field, value)
    text = compact_text(export_pdf(draft).data)
    assert "".join(skill.display_name.split()) in text


def test_reference_crowded_optional_skills_are_retained_or_reported_as_overflow(draft):
    # Activate every catalog row: the original has 60 skill boxes, while this
    # catalog has 67 rows. A second-page continuation is allowed; silent loss
    # of extra modern/specialty skills is not.
    for skill in draft.skills:
        skill.extra_final += 1
    extras = [skill for skill in draft.skills if skill.template_slot in {
        "F25", "F32", "F37", "F41", "AB41", "AB42", "AB43", "AB44", "AB45", "AB46", "AB47", "AB48",
    }]
    try:
        result = export_pdf(draft)
    except PdfLayoutOverflowError as error:
        assert str(error).strip()
    else:
        text = compact_text(result.data)
        for skill in extras:
            assert "".join(skill.display_name.split()) in text


@pytest.mark.parametrize("field", ["name", "personal_story", "asset_description"])
def test_reference_rejects_content_that_cannot_fit_instead_of_truncating(draft, field):
    owner = draft.identity if field == "name" else draft.background if field == "personal_story" else draft.assets
    setattr(owner, field, "不能静默丢失的完整内容" * 1000)
    with pytest.raises(PdfLayoutOverflowError):
        export_pdf(draft)
