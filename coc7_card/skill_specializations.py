"""Named skill specializations shared by browser catalogs and draft conversion."""

from __future__ import annotations

import re
import unicodedata
from collections.abc import Iterable

from .models import SkillDefinition


# Values follow the template's 附表!M219:V246 and combat lookup formulas.
# 数学 is 10 in 人物卡!AF31:AF33 / 附表!C187; the auxiliary list's 1
# conflicts with those actual character-sheet formulas.
SPECIALIZATIONS: dict[str, tuple[tuple[str, int], ...]] = {
    "技艺": tuple((name, 5) for name in (
        "表演", "美术", "摄影", "伪造", "写作", "书法", "乐理", "厨艺", "裁缝", "理发",
        "建筑", "舞蹈", "酿酒", "捕鱼", "歌唱", "制陶", "雕塑", "杂技", "风水", "技术制图",
        "耕作", "打字", "速记", "木匠", "莫里斯舞蹈", "歌剧歌唱", "粉刷匠与油漆工", "吹真空管",
    )),
    "格斗": (("斗殴", 25), ("鞭子", 5), ("电锯", 10), ("链枷", 10), ("绞具", 15),
           ("斧", 15), ("剑", 20), ("矛", 20)),
    "射击": (("手枪", 20), ("步枪/霰弹枪", 25), ("冲锋枪", 15), ("弓术", 15),
           ("喷射器", 10), ("机枪", 10), ("重武器", 10)),
    "科学": tuple((name, 10 if name == "数学" else 1) for name in (
        "地质学", "化学", "生物学", "数学", "天文学", "物理学", "药学", "植物学", "动物学",
        "密码学", "工程学", "气象学", "司法科学",
    )),
    "外语": (),
    "驾驶": (("飞行器", 1), ("船", 1)),
    "生存": (),
    "学识": (),
}

PRESET_SPECIALIZATIONS = {
    "F20": "摄影", "F21": "美术", "F22": "写作",
    "F34": "斗殴", "F35": "斧", "F36": "剑", "F37": "矛",
    "F38": "手枪", "F39": "步枪/霰弹枪", "F40": "冲锋枪", "F41": "弓术",
    "AB31": "物理学", "AB32": "化学", "AB33": "生物学",
}

# Short labels already used in the template's occupation choices. Keep their
# original text in the draft while applying the corresponding skill base.
SPECIALIZATION_ALIASES = {
    "射击": {"步枪/霰弹枪": ("步/霰", "来复", "霰弹")},
    "格斗": {"电锯": ("链锯",), "鞭子": ("鞭",)},
}

BRANCH_SLOT_PATTERN = re.compile(r"branch-[0-9a-f]{32}")
MAX_BRANCH_SPECIALIZATION_LENGTH = 80
MAX_SKILL_BRANCHES = 200
GROUP_BASE_VALUES = {group: {"技艺": 5, "生存": 10}.get(group, 1) for group in SPECIALIZATIONS}


def is_branch_slot(value: object) -> bool:
    return isinstance(value, str) and BRANCH_SLOT_PATTERN.fullmatch(value) is not None


def validate_branch_identity(slot: str, name: str, specialization: str) -> tuple[str, str, str]:
    """Validate shared branch identifiers without changing their stable identity."""
    if not is_branch_slot(slot):
        raise ValueError("分支技能标识无效，请重新添加该分支技能。")
    if not isinstance(name, str) or name.strip() not in SPECIALIZATIONS:
        raise ValueError("分支技能的所属大项无效，请选择已有技能大项。")
    if not isinstance(specialization, str) or not specialization.strip():
        raise ValueError("请填写分支技能名称。")
    if len(specialization.strip()) > MAX_BRANCH_SPECIALIZATION_LENGTH:
        raise ValueError(f"分支技能名称不能超过 {MAX_BRANCH_SPECIALIZATION_LENGTH} 个字符。")
    return slot, name.strip(), specialization.strip()


def _normalized_specialization(value: str) -> str:
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", value)).strip().casefold()


def canonical_specialization(group: str, specialization: str) -> str:
    """Normalize aliases for duplicate detection, keeping displayed names intact."""
    target = _normalized_specialization(specialization)
    for name, _base in SPECIALIZATIONS.get(group, ()):
        names = (name, *SPECIALIZATION_ALIASES.get(group, {}).get(name, ()))
        if any(target == _normalized_specialization(candidate) for candidate in names):
            return _normalized_specialization(name)
    return target


def skill_group(name: str) -> str:
    """Recognize template category names without regrouping renamed skills."""
    group = re.sub(r"[：:①②③]+$", "", name.strip()).strip()
    return group if group in SPECIALIZATIONS else ""


def matching_branch_slot(name: str, specialization: str, rows: Iterable[dict]) -> str | None:
    """Reuse one existing branch for an occupation's explicit specialization."""
    group = skill_group(name)
    if not group or not specialization.strip() or specialization.strip() in {"任一", "任意", "任选", "自选"}:
        return None
    target = canonical_specialization(group, specialization)
    for row in rows:
        slot = row.get("template_slot") or row.get("key")
        branch_name = row.get("specialization")
        if (is_branch_slot(slot) and row.get("name") == group and isinstance(branch_name, str)
                and canonical_specialization(group, branch_name) == target):
            return slot
    return None


def _specialization_options(group: str) -> list[dict]:
    return [
        {"name": name, "base_value": base,
         "aliases": list(SPECIALIZATION_ALIASES.get(group, {}).get(name, ()))}
        for name, base in SPECIALIZATIONS.get(group, ())
    ]


def specialization_metadata(definition: SkillDefinition) -> dict:
    group = skill_group(definition.name)
    return {
        "group": group,
        "specialization_options": _specialization_options(group),
        "preset_specialization": PRESET_SPECIALIZATIONS.get(definition.template_slot, ""),
    }


def skill_groups_metadata() -> list[dict]:
    return [
        {"name": group, "group": group, "base_value": GROUP_BASE_VALUES[group],
         "specialization_options": _specialization_options(group)}
        for group in SPECIALIZATIONS
    ]


def specialization_base_value(name: str, specialization: str, fallback: int) -> int:
    """Known named subskills use their own base; custom/empty names keep it."""
    group = skill_group(name)
    specialization = canonical_specialization(group, specialization)
    for name, base in SPECIALIZATIONS.get(group, ()):
        if specialization == _normalized_specialization(name):
            return base
    return fallback
