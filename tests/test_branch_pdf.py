"""Dynamic skill branches retain every value within the two-page PDF contract."""

import hashlib
import io

from PIL import Image
from pypdf import PdfReader
import pytest

from app import FONT_PATH, get_catalog
from coc7_card.exporters.pdf import PdfExporter, PdfLayoutOverflowError, TEMPLATE_PATH
from coc7_card.models import ExperiencePackage, SkillAllocation
from coc7_card.rules import RuleEngine
from coc7_card.web import build_draft, catalog_payload


def _base_draft():
    catalog = get_catalog()
    occupation = next(item for item in catalog.occupations if item.credit_min == 0)
    definitions = catalog_payload(catalog)["skills"]
    return build_draft({
        "identity": {"name": "分支排版样本", "age": 30},
        "attributes": {key: 50 for key in ("STR", "CON", "SIZ", "DEX", "APP", "INT", "POW", "EDU", "Luck")},
        "occupation_id": occupation.occupation_id,
        "skills": [{
            "template_slot": item["template_slot"],
            "specialization": item["default_specialization"] or item["preset_specialization"],
        } for item in definitions],
    }, catalog)


def _branch(index, group="科学", specialization=None, base=1, **points):
    slot = f"branch-{index:032x}"
    return SkillAllocation(
        key=slot,
        template_slot=slot,
        name=group,
        specialization=specialization or f"分支{index}",
        base_value=base,
        **points,
    )


def _many_branches_draft():
    draft = _base_draft()
    definitions = [
        ("射击", "喷射器", 10), ("射击", "机枪", 10), ("射击", "重武器", 10),
        ("科学", "数学", 10), ("科学", "天文学", 1), ("科学", "地质学", 1),
        ("科学", "密码学", 1), ("科学", "药学", 1),
        ("技艺", "表演", 5), ("技艺", "伪造", 5), ("技艺", "书法", 5), ("技艺", "厨艺", 5),
    ]
    for index, (group, specialization, base) in enumerate(definitions, 1):
        draft.skills.append(_branch(index, group, specialization, base,
                                    interest_points=2, extra_final=index))
    draft.skills[-12].occupation_points = 3
    draft.skills[-12].selected_occupation = True
    draft.skills[-12].experience_points = 1
    draft.experience = ExperiencePackage(name="合成经历", skill_points=1)
    return draft


@pytest.fixture
def draft():
    return _base_draft()


def _compact_text(page):
    return "".join(page.extract_text().split())


def test_dynamic_branches_reach_both_original_pdf_pages_and_previews(tmp_path):
    draft = _many_branches_draft()
    branches = [skill for skill in draft.skills if skill.template_slot.startswith("branch-")]
    source_hash = hashlib.sha256(TEMPLATE_PATH.read_bytes()).hexdigest()
    assert len(draft.skills) == 67 + 12
    result = PdfExporter(FONT_PATH).export(draft, include_preview=True)
    (tmp_path / "many-branches.pdf").write_bytes(result.data)
    original = PdfReader(TEMPLATE_PATH)
    reader = PdfReader(io.BytesIO(result.data))
    assert len(reader.pages) == result.page_count == len(result.preview_images) == 2
    assert not reader.get_fields()
    assert "/AcroForm" not in reader.trailer["/Root"]
    for actual, source, preview in zip(reader.pages, original.pages, result.preview_images, strict=True):
        assert tuple(actual.mediabox) == tuple(source.mediabox)
        assert tuple(actual.cropbox) == tuple(source.cropbox)
        with Image.open(io.BytesIO(preview)) as image:
            assert image.format == "PNG"
            assert image.width > 800 and image.height > 1000
    text = "".join(_compact_text(page) for page in reader.pages)
    for skill in branches:
        label = skill.display_name
        # Front boxes print success values before the label; the reverse-side
        # continuation writes the label followed by all three success values.
        front = f"{skill.final_value}{skill.hard_value}{skill.extreme_value}{label}"
        back = f"技能：{label}{skill.final_value}/{skill.hard_value}/{skill.extreme_value}"
        assert front in text or back in text, label
    assert "射击（喷射器）" in _compact_text(reader.pages[0])
    assert "技能：技艺（厨艺）" in _compact_text(reader.pages[1])
    assert hashlib.sha256(TEMPLATE_PATH.read_bytes()).hexdigest() == source_hash


def test_too_many_dynamic_branches_raise_clear_two_page_overflow(draft):
    draft.skills.extend(_branch(index) for index in range(1, 101))
    assert RuleEngine.validate(draft).can_export
    with pytest.raises(PdfLayoutOverflowError, match="两页空白栏已满") as error:
        PdfExporter(FONT_PATH).export(draft, include_preview=False)
    assert "未能完整排入的内容" in str(error.value)
    assert "分支" in str(error.value)


def test_dynamic_branch_name_too_long_for_front_box_raises_overflow(draft):
    draft.skills.append(_branch(1, "科学", "完整技能名称" * 10))
    with pytest.raises(PdfLayoutOverflowError, match="技能名称超过原版栏位"):
        PdfExporter(FONT_PATH).export(draft, include_preview=False)


@pytest.mark.parametrize("field,allocations,code", [
    ("occupation_points", (70, 70, 61), "OCCUPATION_POINTS_TOTAL"),
    ("interest_points", (34, 34, 33), "INTEREST_POINTS_TOTAL"),
    ("experience_points", (4, 4, 3), "EXPERIENCE_BUDGET"),
])
def test_dynamic_branch_allocations_count_toward_shared_budgets(draft, field, allocations, code):
    draft.experience = ExperiencePackage(name="合成经历", skill_points=10)
    draft.skills.extend(_branch(index, **{field: points}, selected_occupation=True)
                        for index, points in enumerate(allocations, 1))
    report = RuleEngine.validate(draft)
    assert code in {issue.code for issue in report.errors}
    # Each branch remains below 99; the error is the combined point budget.
    assert "SKILL_MAXIMUM" not in {issue.code for issue in report.errors}


def test_dynamic_branches_keep_growth_separate_and_occupation_eligibility_independent(draft):
    draft.experience = ExperiencePackage(name="合成经历", skill_points=3)
    chosen = _branch(1, "射击", "喷射器", 10, occupation_points=7,
                     interest_points=2, extra_final=11, experience_points=3,
                     selected_occupation=True)
    other = _branch(2, "射击", "机枪", 10, occupation_points=1)
    draft.skills.extend([chosen, other])
    report = RuleEngine.validate(draft)
    assert chosen.final_value == 33
    assert "EXPERIENCE_BUDGET" not in {issue.code for issue in report.errors}
    assert [(issue.code, issue.field) for issue in report.errors] == [
        ("SKILL_NON_OCCUPATION", other.template_slot),
    ]
    other.occupation_points = 0
    assert RuleEngine.validate(draft).can_export


def test_duplicate_dynamic_branch_names_are_still_reported(draft):
    draft.skills.extend([_branch(1, "科学", "分子研究"), _branch(2, "科学", "分子研究")])
    duplicate = [issue for issue in RuleEngine.validate(draft).warnings if issue.code == "SKILL_DUPLICATE"]
    assert len(duplicate) == 1
    assert duplicate[0].field == draft.skills[-1].template_slot
    assert "科学（分子研究）" in duplicate[0].message
