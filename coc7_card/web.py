from __future__ import annotations

from dataclasses import asdict
from typing import Any

from .catalog import TemplateCatalog, normalize_skill_name, split_occupation_skill_token
from .models import (
    Assets,
    Attributes,
    Background,
    CharacterDraft,
    ExperiencePackage,
    InventoryItem,
    InvestigatorIdentity,
    OccupationDefinition,
    SkillAllocation,
    Weapon,
)
from .rules import RuleEngine
from .currency import currency_options, legacy_currency_code, parse_usd, format_amount, usd_reference, YEAR_START, YEAR_END


class DraftPayloadError(ValueError):
    """Raised when a browser payload cannot be converted into a character draft."""


def catalog_payload(catalog: TemplateCatalog) -> dict[str, Any]:
    return {
        "meta": {
            "source_sha256": catalog.source_sha256,
            "sheet_count": len(catalog.sheet_names),
            "occupation_count": len(catalog.occupations),
            "skill_count": len(catalog.skills),
        },
        "occupations": [asdict(item) for item in catalog.occupations],
        "skills": [asdict(item) for item in catalog.skills],
        "weapons": [asdict(item) for item in catalog.weapons],
        "inventory": [asdict(item) for item in catalog.inventory],
        "experience_packages": list(catalog.experience_packages),
        "currencies": {str(year): currency_options(catalog.currency_quotes, year)
                       for year in range(YEAR_START, YEAR_END + 1)},
        "currency_missing": {str(year): currency_options(catalog.currency_quotes, year, available=False)
                             for year in range(YEAR_START, YEAR_END + 1)},
        "currency_metadata": catalog.currency_metadata,
    }


def report_payload(report) -> dict[str, Any]:  # noqa: ANN001
    return {
        "can_export": report.can_export,
        "counts": {
            "error": len(report.errors),
            "warning": len(report.warnings),
            "info": len(report.infos),
        },
        "issues": [
            {
                "severity": issue.severity.value,
                "code": issue.code,
                "message": issue.message,
                "field": issue.field,
            }
            for issue in report.issues
        ],
    }


def derived_payload(derived) -> dict[str, Any]:  # noqa: ANN001
    return asdict(derived)


def build_draft(payload: dict[str, Any], catalog: TemplateCatalog) -> CharacterDraft:
    if not isinstance(payload, dict):
        raise DraftPayloadError("提交的数据不是有效的调查员草稿。")

    identity_raw = _mapping(payload.get("identity"))
    attributes = Attributes.from_mapping(_mapping(payload.get("attributes")))
    occupation = _parse_occupation(payload, catalog)
    group_choices = {
        str(marker): [_text(item) for item in values if _text(item)]
        for marker, values in _mapping(payload.get("group_choices")).items()
        if isinstance(values, list)
    }
    free_skill_choices = [
        _text(item) for item in _list(payload.get("free_skill_choices")) if _text(item)
    ]

    skill_rows = {
        _text(item.get("template_slot") or item.get("key")): item
        for item in _list(payload.get("skills"))
        if isinstance(item, dict) and _text(item.get("template_slot") or item.get("key"))
    }
    selected_slots = _selected_occupation_slots(
        occupation,
        catalog,
        group_choices,
        free_skill_choices,
        [_text(item) for item in _list(payload.get("custom_skill_slots"))],
        skill_rows,
    )
    skills: list[SkillAllocation] = []
    for definition in catalog.skills:
        row = skill_rows.get(definition.template_slot, {})
        base_default = RuleEngine.skill_base_value(
            definition.base_formula,
            attributes,
            definition.base_value,
        )
        # A spreadsheet may carry an explicitly edited starting skill value.
        # Keep that input separate from growth/experience; ordinary drafts still
        # derive starting values from the catalog and current attributes.
        base_override = row.get("base_override")
        if base_override is not None:
            if isinstance(base_override, bool) or not isinstance(base_override, (int, float)):
                raise DraftPayloadError(f"{definition.name}的导入基础值必须为 0—999 的整数。")
            if not 0 <= base_override <= 999 or int(base_override) != base_override:
                raise DraftPayloadError(f"{definition.name}的导入基础值必须为 0—999 的整数。")
            base_default = int(base_override)
        specialization = _text(row.get("specialization")) or definition.default_specialization
        if definition.template_slot in selected_slots and selected_slots[definition.template_slot]:
            specialization = specialization or selected_slots[definition.template_slot]
        skills.append(
            SkillAllocation(
                key=definition.template_slot,
                template_slot=definition.template_slot,
                name=_text(row.get("name")) or definition.name,
                specialization=specialization,
                base_value=base_default,
                occupation_points=_integer(row.get("occupation_points"), 0),
                interest_points=_integer(row.get("interest_points"), 0),
                extra_final=_integer(row.get("extra_final"), 0),
                experience_points=_integer(row.get("experience_points"), 0),
                selected_occupation=definition.template_slot in selected_slots,
            )
        )

    credit_rating = next(
        (skill.final_value for skill in skills if skill.name.rstrip("：:").strip() == "信用评级"),
        0,
    )
    background_raw = _mapping(payload.get("background"))
    assets_raw = _mapping(payload.get("assets"))
    story_year = identity_raw.get("story_year")
    has_story_year = story_year is not None and bool(str(story_year).strip())
    try:
        exchange_year = int(str(story_year if has_story_year else assets_raw.get("exchange_year", 1920)))
    except ValueError as exc:
        raise DraftPayloadError("故事年份／汇率年份请输入整数年份。") from exc
    story_year = exchange_year if has_story_year else None
    code = _text(assets_raw.get("currency_code", "USD"))
    if payload.get("version") == 2:
        code = legacy_currency_code(code, exchange_year)
    quote = next((item for item in currency_options(catalog.currency_quotes, exchange_year) if item["code"] == code), None)
    if quote is None:
        raise DraftPayloadError("该币种在所选年份没有可用汇率，请重新选择币种。美元输入已保留。")
    defaults = usd_reference(credit_rating, exchange_year)
    raw_amounts = [assets_raw.get(key, assets_raw.get(legacy, "")) for key, legacy in
                   [("usd_spending", "spending_level"), ("usd_cash", "cash"), ("usd_assets", "other_assets")]]
    usd_amounts = [parse_usd(value, default) for value, default in zip(raw_amounts, defaults)]

    return CharacterDraft(
        identity=InvestigatorIdentity(
            name=_text(identity_raw.get("name")),
            player=_text(identity_raw.get("player")),
            occupation_name=occupation.name if occupation else _text(identity_raw.get("occupation_name")),
            age=_integer(identity_raw.get("age"), 30),
            gender=_text(identity_raw.get("gender")),
            residence=_text(identity_raw.get("residence")),
            birthplace=_text(identity_raw.get("birthplace")),
            era=_text(identity_raw.get("era")) or "1920s",
            current_date=_text(identity_raw.get("current_date")),
            story_year=story_year,
        ),
        attributes=attributes,
        occupation=occupation,
        skills=skills,
        background=Background(
            appearance=_text(background_raw.get("appearance")),
            beliefs=_text(background_raw.get("beliefs")),
            significant_people=_text(background_raw.get("significant_people")),
            meaningful_places=_text(background_raw.get("meaningful_places")),
            treasured_possessions=_text(background_raw.get("treasured_possessions")),
            traits=_text(background_raw.get("traits")),
            scars=_text(background_raw.get("scars")),
            phobias_manias=_text(background_raw.get("phobias_manias")),
            personal_story=_text(background_raw.get("personal_story")),
            key_connection=_text(background_raw.get("key_connection")),
            contacts_notes=_text(background_raw.get("contacts_notes")),
        ),
        assets=Assets(
            credit_rating=credit_rating,
            living_standard=_text(assets_raw.get("living_standard")),
            spending_level=format_amount(usd_amounts[0], quote["rate"]),
            cash=format_amount(usd_amounts[1], quote["rate"]),
            other_assets=format_amount(usd_amounts[2], quote["rate"]),
            asset_description=_text(assets_raw.get("asset_description")),
            vehicles=_text(assets_raw.get("vehicles")),
            residence=_text(assets_raw.get("residence")),
            luxuries=_text(assets_raw.get("luxuries")),
            securities=_text(assets_raw.get("securities")),
            other=_text(assets_raw.get("other")),
            currency=quote["name"], currency_code=code, exchange_year=exchange_year,
            exchange_rate=quote["rate"], exchange_date=quote["date"], exchange_source=quote["source"],
            exchange_basis=quote["basis"], exchange_warning=quote["warning"],
            exchange_source_url=quote["source_url"],
            usd_spending="" if raw_amounts[0] is None or not str(raw_amounts[0]).strip() else str(usd_amounts[0]),
            usd_cash="" if raw_amounts[1] is None or not str(raw_amounts[1]).strip() else str(usd_amounts[1]),
            usd_assets="" if raw_amounts[2] is None or not str(raw_amounts[2]).strip() else str(usd_amounts[2]),
            auto_spending=raw_amounts[0] is None or not str(raw_amounts[0]).strip(),
            auto_cash=raw_amounts[1] is None or not str(raw_amounts[1]).strip(),
            auto_assets=raw_amounts[2] is None or not str(raw_amounts[2]).strip(),
        ),
        weapons=[_parse_weapon(item) for item in _list(payload.get("weapons")) if _named_row(item)],
        inventory=[
            _parse_inventory_item(item)
            for item in _list(payload.get("inventory"))
            if _named_row(item)
        ],
        group_choices=group_choices,
        experience=_parse_experience(payload.get("experience"), catalog),
        free_skill_choices=free_skill_choices if occupation and not occupation.is_custom else [],
        nonstandard_override=bool(payload.get("nonstandard_override", False)),
    )


def _parse_occupation(payload: dict[str, Any], catalog: TemplateCatalog) -> OccupationDefinition | None:
    raw = payload.get("occupation")
    if raw in (None, "", 0):
        occupation_id = _integer(payload.get("occupation_id"), 0)
        return catalog.occupation_by_id(occupation_id) if occupation_id else None
    if isinstance(raw, (int, float, str)):
        occupation = catalog.occupation_by_id(_integer(raw, 0))
        if occupation is None:
            raise DraftPayloadError("所选职业不存在或已失效。")
        return occupation
    raw = _mapping(raw)
    is_custom = bool(raw.get("is_custom")) or _integer(raw.get("occupation_id"), 0) == 1
    if not is_custom:
        occupation = catalog.occupation_by_id(_integer(raw.get("occupation_id"), 0))
        if occupation is None:
            raise DraftPayloadError("所选职业不存在或已失效。")
        return occupation

    low = max(0, min(99, _integer(raw.get("credit_min"), 0)))
    high = max(0, min(99, _integer(raw.get("credit_max"), 99)))
    if low > high:
        low, high = high, low
    return OccupationDefinition(
        occupation_id=1,
        name=_text(raw.get("name")) or "自定义职业",
        credit_min=low,
        credit_max=high,
        point_formula=_text(raw.get("point_formula")) or "EDU*4",
        summary=_text(raw.get("summary")) or "由玩家与 KP 共同定义的职业。",
        contacts=_text(raw.get("contacts")),
        description=_text(raw.get("description")),
        fixed_skills=tuple(_text(item) for item in _list(raw.get("fixed_skills")) if _text(item)),
        is_custom=True,
    )


def _selected_occupation_slots(
    occupation: OccupationDefinition | None,
    catalog: TemplateCatalog,
    group_choices: dict[str, list[str]],
    free_skill_choices: list[str],
    custom_skill_slots: list[str],
    raw_skill_rows: dict[str, dict[str, Any]],
) -> dict[str, str]:
    selected: dict[str, str] = {"F26": ""}
    if occupation is None:
        return selected
    if occupation.is_custom:
        slots = [slot for slot in custom_skill_slots if catalog.skill_cell(slot)]
        if not slots:
            slots = [
                slot
                for slot, row in raw_skill_rows.items()
                if bool(row.get("selected_occupation")) and catalog.skill_cell(slot)
            ]
        selected.update({slot: "" for slot in slots[:8]})
        return selected

    tokens = list(occupation.fixed_skills)
    for group in occupation.choice_groups:
        tokens.extend(group_choices.get(group.marker, []))
    definitions_by_name: dict[str, list[Any]] = {}
    for definition in catalog.skills:
        definitions_by_name.setdefault(normalize_skill_name(definition.name), []).append(definition)
    for token in tokens:
        base_name, specialization = split_occupation_skill_token(token)
        matches = definitions_by_name.get(base_name, [])
        if matches:
            selected[matches[0].template_slot] = specialization
    for slot in free_skill_choices:
        if catalog.skill_cell(slot):
            selected[slot] = ""
    return selected


def _parse_weapon(value: Any) -> Weapon:
    raw = _mapping(value)
    return Weapon(
        name=_text(raw.get("name")),
        category=_text(raw.get("category")),
        skill=_text(raw.get("skill")),
        damage=_text(raw.get("damage")),
        range=_text(raw.get("range")),
        attacks=_text(raw.get("attacks")),
        ammo=_text(raw.get("ammo")),
        malfunction=_text(raw.get("malfunction")),
        notes=_text(raw.get("notes")),
    )


def _parse_experience(value: Any, catalog: TemplateCatalog) -> ExperiencePackage:
    raw = _mapping(value)
    selected = _text(raw.get("selection"))
    # Also accept the canonical model representation used in export tests.
    if "selection" not in raw:
        selected = _text(raw.get("name"))
    if not selected or selected == "无":
        return ExperiencePackage()
    if selected == "custom" or ("selection" not in raw and raw.get("is_custom") is True):
        name = _text(raw.get("name"))
        if not name:
            raise DraftPayloadError("请填写自定义经历包名称。")
        return ExperiencePackage(name=name, skill_points=_integer(raw.get("skill_points"), 0),
                                 san_loss=_integer(raw.get("san_loss"), 0), notes=_text(raw.get("notes")), is_custom=True)
    definition = next((item for item in catalog.experience_packages if item["name"] == selected), None)
    if definition is None:
        raise DraftPayloadError("请选择有效的经历包，或使用自定义经历包。")
    return ExperiencePackage(name=selected, skill_points=definition["skill_points"],
                             san_loss=_integer(raw.get("san_loss"), 0),
                             notes=definition["notes"], minimum_age=definition["minimum_age"])


def _parse_inventory_item(value: Any) -> InventoryItem:
    raw = _mapping(value)
    return InventoryItem(
        name=_text(raw.get("name")),
        status=_text(raw.get("status")),
        location=_text(raw.get("location")),
        backpack_slot=_text(raw.get("backpack_slot")),
    )


def _named_row(value: Any) -> bool:
    return isinstance(value, dict) and bool(_text(value.get("name")))


def _mapping(value: Any) -> dict[str, Any]:
    return value if isinstance(value, dict) else {}


def _list(value: Any) -> list[Any]:
    return value if isinstance(value, list) else []


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _integer(value: Any, default: int = 0) -> int:
    if value in (None, ""):
        return int(default)
    try:
        return int(float(value))
    except (TypeError, ValueError, OverflowError):
        return int(default)
