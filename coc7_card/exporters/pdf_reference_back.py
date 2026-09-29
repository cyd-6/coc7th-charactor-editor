"""Fill the reverse side of the supplied 1920s investigator sheet.

Coordinates are measured in points from the reference page's top-left corner.
The original PDF supplies all borders, rules, headings and decorative artwork.
Only the two tiny labels inside a used companion box need to be cleared when
placing the application's free-form contact notes there.
"""

from __future__ import annotations

from typing import Protocol

from ..models import CharacterDraft


Slot = tuple[float, float, float]


class _Painter(Protocol):
    def text(self, value: object, x: float, baseline: float, width: float,
             size: float = 8.5, min_size: float = 7, align: str = "left") -> None: ...

    def flow(self, value: str, slots: list[Slot], size: float = 8.5,
             min_size: float = 7) -> str: ...

    def white(self, rect: tuple[float, float, float, float]) -> None: ...


# field, printed heading, first-line x, full-line x, right edge, first baseline,
# number of ruled lines. The traits field stops before the original silhouette.
_BACKGROUND = (
    ("appearance", "形象描述", 143, 95, 328, 114.5, 4),
    ("beliefs", "思想与信念", 154, 95, 328, 172.0, 4),
    ("significant_people", "重要之人", 144, 95, 328, 229.5, 4),
    ("meaningful_places", "意义非凡之地", 164, 95, 328, 287.0, 4),
    ("treasured_possessions", "宝贵之物", 144, 95, 328, 344.5, 5),
    ("traits", "特质", 363, 340, 494, 114.5, 4),
    ("scars", "创伤和疤痕", 407, 340, 573, 172.0, 4),
    ("phobias_manias", "恐惧症和躁狂症", 428, 340, 573, 229.5, 4),
)

_COMPANION_BOXES = (
    (302, 645, 391, 667),
    (395, 637.7, 485, 659.7),
    (488, 645, 577, 667),
    (302, 691, 391, 713),
    (488, 691, 577, 713),
    (302, 736.5, 391, 758.5),
    (395, 745, 485, 767),
    (488, 736.5, 577, 758.5),
)


def _slots(first_x: float, full_x: float, right: float, baseline: float,
           count: int) -> list[Slot]:
    return [
        (first_x if index == 0 else full_x, baseline + index * 14.5,
         right - (first_x if index == 0 else full_x))
        for index in range(count)
    ]


def _string(value: object) -> str:
    return "" if value is None else str(value).strip()


def _flow_with_remainder(painter: _Painter, text: str, slots: list[Slot],
                         label: str, extra: list[str], *, size: float = 8.5) -> None:
    if not text:
        return
    remaining = painter.flow(text, slots, size=size, min_size=7)
    if remaining:
        extra.append(f"{label}续：{remaining}")


def _draw_contacts(painter: _Painter, value: str, extra: list[str]) -> None:
    paragraphs = [line.strip() for line in value.splitlines() if line.strip()]
    remaining = paragraphs.pop(0) if paragraphs else ""
    for x, y, right, bottom in _COMPANION_BOXES:
        if not remaining:
            if not paragraphs:
                break
            remaining = paragraphs.pop(0)
        # The free-form model has no distinct role/player fields. Remove only
        # those printed words, leaving the box edges and connecting artwork.
        painter.white((x + 1.2, y + 1.2, x + 19.5, bottom - 1.2))
        remaining = painter.flow(
            remaining,
            [(x + 2, y + 8.6, right - x - 4),
             (x + 2, y + 18.1, right - x - 4)],
            size=7.5, min_size=7,
        )
    if remaining or paragraphs:
        extra.append("联系人续：" + "\n".join(
            part for part in [remaining, *paragraphs] if part
        ))


def _inventory_text(draft: CharacterDraft) -> str:
    lines: list[str] = []
    for item in draft.inventory:
        if not any(_string(value) for value in
                   (item.name, item.status, item.location, item.backpack_slot)):
            continue
        details: list[str] = []
        if _string(item.status):
            details.append(_string(item.status))
        if _string(item.location):
            details.append(_string(item.location))
        if _string(item.backpack_slot):
            details.append(f"格{item.backpack_slot}")
        text = _string(item.name)
        if details:
            text += "（" + "，".join(details) + "）"
        lines.append(text)
    return "\n".join(lines)


def _asset_notes(draft: CharacterDraft) -> str:
    asset = draft.assets
    currency = asset.currency
    if asset.currency_code and asset.currency_code not in currency:
        currency += f"（{asset.currency_code}）"
    parts = [
        f"{asset.exchange_year}年 {currency}；"
        f"1 USD={asset.exchange_rate} {asset.currency}",
    ]
    if asset.living_standard:
        parts.append(f"生活水平：{asset.living_standard}")
    if asset.exchange_basis:
        parts.append(f"口径：{asset.exchange_basis}")
    if asset.exchange_date:
        parts.append(f"期间：{asset.exchange_date}")
    if asset.exchange_source:
        parts.append(f"来源：{asset.exchange_source}")
    if asset.exchange_warning:
        parts.append(f"提示：{asset.exchange_warning}")
    for label, value in (
        ("明细", asset.asset_description),
        ("交通工具", asset.vehicles),
        ("住所", asset.residence),
        ("奢侈品", asset.luxuries),
        ("股票/证券", asset.securities),
        ("其他", asset.other),
    ):
        if _string(value):
            parts.append(f"{label}：{value}")
    return "；".join(parts)


def draw_back(painter: _Painter, draft: CharacterDraft, notes: list[str]) -> None:
    """Fill the original back page, or explicitly reject unprintable overflow."""
    from .pdf import PdfLayoutOverflowError

    extra: list[str] = []
    for field, label, first_x, full_x, right, baseline, count in _BACKGROUND:
        _flow_with_remainder(
            painter, _string(getattr(draft.background, field)),
            _slots(first_x, full_x, right, baseline, count), label, extra,
        )

    # The sheet offers twenty ruled equipment lines; use all of them in reading
    # order and retain every submitted item rather than slicing the input list.
    equipment_slots = [
        (x, 464 + index * 14.5, 130)
        for x in (95, 231.5) for index in range(10)
    ]
    _flow_with_remainder(painter, _inventory_text(draft), equipment_slots,
                         "装备和物品", extra, size=8)

    for label, value, x, baseline in (
        ("消费水平", draft.assets.spending_level, 431.5, 464),
        ("现金", draft.assets.cash, 410, 478.5),
        ("资产", draft.assets.other_assets, 410, 493),
    ):
        _flow_with_remainder(painter, _string(value), [(x, baseline, 573 - x)],
                             label, extra)
    asset_slots = [(384.5, 507.5 + index * 14.5, 188.5) for index in range(7)]
    _flow_with_remainder(painter, _asset_notes(draft), asset_slots,
                         "资产说明", extra, size=8)

    _draw_contacts(painter, draft.background.contacts_notes, extra)

    supplement: list[str] = []
    if draft.background.personal_story:
        supplement.append(f"个人故事：{draft.background.personal_story}")
    if draft.background.key_connection:
        supplement.append(f"关键连接：{draft.background.key_connection}")
    experience = draft.experience
    if (experience.name or experience.skill_points or experience.san_loss
            or experience.minimum_age or experience.notes):
        detail = (f"经历：{experience.name}；经历包点{experience.skill_points}；"
                  f"SAN减{experience.san_loss}")
        if experience.minimum_age:
            detail += f"；最低年龄{experience.minimum_age}"
        if experience.notes:
            detail += f"；{experience.notes}"
        supplement.append(detail)
    supplement.extend(_string(note) for note in notes if _string(note))
    supplement.extend(extra)

    # These two original sections have no standalone web fields. Their ruled
    # blanks carry explicitly prefixed supplemental data; all printed headings
    # and the illustration remain intact.
    supplement_slots = (
        _slots(459, 340, 573, 287, 4)
        + _slots(407, 340, 573, 344.5, 5)
    )
    remaining = painter.flow("；".join(supplement), supplement_slots, size=8.3, min_size=7)
    if remaining:
        raise PdfLayoutOverflowError(
            "参考调查员卡的两页空白栏已满，请缩短背景、资产明细或补充备注后再导出。"
            f"未能完整排入的内容：{remaining[:120]}"
        )
