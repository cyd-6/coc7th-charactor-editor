from __future__ import annotations

from dataclasses import dataclass, field
from enum import StrEnum
from typing import Any


ATTRIBUTE_KEYS = ("STR", "CON", "SIZ", "DEX", "APP", "INT", "POW", "EDU", "Luck")


@dataclass(slots=True)
class Attributes:
    STR: int = 50
    CON: int = 50
    SIZ: int = 50
    DEX: int = 50
    APP: int = 50
    INT: int = 50
    POW: int = 50
    EDU: int = 50
    Luck: int = 50

    def as_mapping(self) -> dict[str, int]:
        return {key: int(getattr(self, key)) for key in ATTRIBUTE_KEYS}

    @classmethod
    def from_mapping(cls, values: dict[str, Any]) -> "Attributes":
        return cls(**{key: int(values.get(key, 50)) for key in ATTRIBUTE_KEYS})


@dataclass(slots=True)
class InvestigatorIdentity:
    name: str = ""
    player: str = ""
    occupation_name: str = ""
    age: int = 30
    gender: str = ""
    residence: str = ""
    birthplace: str = ""
    era: str = "1920s"
    current_date: str = ""
    story_year: int | None = None


@dataclass(frozen=True, slots=True)
class SkillDefinition:
    key: str
    name: str
    base_value: int
    template_slot: str
    base_formula: str | None = None
    specializable: bool = False
    default_specialization: str = ""


@dataclass(slots=True)
class SkillAllocation:
    key: str
    name: str
    base_value: int
    template_slot: str
    specialization: str = ""
    occupation_points: int = 0
    interest_points: int = 0
    extra_final: int = 0
    experience_points: int = 0
    selected_occupation: bool = False

    @property
    def display_name(self) -> str:
        clean = self.name.rstrip("：:").strip()
        return f"{clean}（{self.specialization.strip()}）" if self.specialization.strip() else clean

    @property
    def final_value(self) -> int:
        return int(self.base_value + self.occupation_points + self.interest_points + self.extra_final + self.experience_points)

    @property
    def hard_value(self) -> int:
        return self.final_value // 2

    @property
    def extreme_value(self) -> int:
        return self.final_value // 5


@dataclass(frozen=True, slots=True)
class SkillChoiceGroup:
    marker: str
    label: str
    required_count: int
    candidates: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class OccupationDefinition:
    occupation_id: int
    name: str
    credit_min: int
    credit_max: int
    point_formula: str
    summary: str = ""
    contacts: str = ""
    description: str = ""
    fixed_skills: tuple[str, ...] = ()
    choice_groups: tuple[SkillChoiceGroup, ...] = ()
    free_choices: int = 0
    is_custom: bool = False

    @property
    def credit_range_text(self) -> str:
        return f"{self.credit_min}-{self.credit_max}"


@dataclass(slots=True)
class Background:
    appearance: str = ""
    beliefs: str = ""
    significant_people: str = ""
    meaningful_places: str = ""
    treasured_possessions: str = ""
    traits: str = ""
    scars: str = ""
    phobias_manias: str = ""
    personal_story: str = ""
    key_connection: str = ""
    contacts_notes: str = ""


@dataclass(slots=True)
class Assets:
    credit_rating: int = 0
    living_standard: str = ""
    spending_level: str = ""
    cash: str = ""
    other_assets: str = ""
    asset_description: str = ""
    vehicles: str = ""
    residence: str = ""
    luxuries: str = ""
    securities: str = ""
    other: str = ""
    currency: str = "美元"
    currency_code: str = "USD"
    exchange_year: int = 1920
    exchange_rate: str = "1"
    exchange_date: str = ""
    exchange_source: str = ""
    exchange_basis: str = ""
    exchange_warning: str = ""
    exchange_source_url: str = ""
    usd_spending: str = ""
    usd_cash: str = ""
    usd_assets: str = ""
    auto_spending: bool = True
    auto_cash: bool = True
    auto_assets: bool = True


@dataclass(slots=True)
class Weapon:
    name: str = ""
    category: str = ""
    skill: str = ""
    damage: str = ""
    range: str = ""
    attacks: str = ""
    ammo: str = ""
    malfunction: str = ""
    notes: str = ""


@dataclass(slots=True)
class InventoryItem:
    name: str = ""
    status: str = ""
    location: str = ""
    backpack_slot: str = ""


@dataclass(frozen=True, slots=True)
class DerivedStats:
    hp: int
    mp: int
    san: int
    mov: int
    dodge: int
    damage_bonus: str
    build: int
    occupation_points: int
    interest_points: int


@dataclass(slots=True)
class ExperiencePackage:
    name: str = ""
    skill_points: int = 0
    san_loss: int = 0
    notes: str = ""
    minimum_age: int = 0
    is_custom: bool = False

    @property
    def summary(self) -> str:
        if not self.name:
            return ""
        return f"经历包：{self.name}；技能增长 {self.skill_points}；SAN 减少 {self.san_loss}。{self.notes}"


@dataclass(slots=True)
class CharacterDraft:
    identity: InvestigatorIdentity = field(default_factory=InvestigatorIdentity)
    attributes: Attributes = field(default_factory=Attributes)
    occupation: OccupationDefinition | None = None
    skills: list[SkillAllocation] = field(default_factory=list)
    background: Background = field(default_factory=Background)
    assets: Assets = field(default_factory=Assets)
    weapons: list[Weapon] = field(default_factory=list)
    inventory: list[InventoryItem] = field(default_factory=list)
    group_choices: dict[str, list[str]] = field(default_factory=dict)
    free_skill_choices: list[str] = field(default_factory=list)
    nonstandard_override: bool = False
    experience: ExperiencePackage = field(default_factory=ExperiencePackage)


class IssueSeverity(StrEnum):
    ERROR = "error"
    WARNING = "warning"
    INFO = "info"


@dataclass(frozen=True, slots=True)
class ValidationIssue:
    severity: IssueSeverity
    code: str
    message: str
    field: str = ""


@dataclass(slots=True)
class ValidationReport:
    issues: list[ValidationIssue] = field(default_factory=list)

    @property
    def errors(self) -> list[ValidationIssue]:
        return [issue for issue in self.issues if issue.severity == IssueSeverity.ERROR]

    @property
    def warnings(self) -> list[ValidationIssue]:
        return [issue for issue in self.issues if issue.severity == IssueSeverity.WARNING]

    @property
    def infos(self) -> list[ValidationIssue]:
        return [issue for issue in self.issues if issue.severity == IssueSeverity.INFO]

    @property
    def can_export(self) -> bool:
        return not self.errors

    def add(
        self,
        severity: IssueSeverity,
        code: str,
        message: str,
        field: str = "",
    ) -> None:
        self.issues.append(ValidationIssue(severity, code, message, field))


@dataclass(frozen=True, slots=True)
class ExcelExportResult:
    filename: str
    data: bytes
    source_sha256: str
    sheet_count: int
    warnings: tuple[str, ...] = ()


@dataclass(frozen=True, slots=True)
class PdfExportResult:
    filename: str
    data: bytes
    page_count: int
    preview_images: tuple[bytes, ...] = ()
