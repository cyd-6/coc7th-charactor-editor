from __future__ import annotations

import ast
from collections.abc import Mapping

from .models import (
    Attributes,
    CharacterDraft,
    DerivedStats,
    IssueSeverity,
    ValidationReport,
)
from .skill_specializations import MAX_SKILL_BRANCHES, is_branch_slot


class FormulaError(ValueError):
    """Raised when an occupation point formula contains unsupported syntax."""


class SafeFormulaEvaluator(ast.NodeVisitor):
    """Evaluate the small arithmetic language used by the workbook occupations."""

    def __init__(self, variables: Mapping[str, int]):
        self.variables = {str(key).upper(): int(value) for key, value in variables.items()}

    def evaluate(self, expression: str) -> int:
        text = expression.strip().lstrip("=").replace("×", "*").replace("，", ",")
        if not text:
            raise FormulaError("职业点公式为空")
        try:
            tree = ast.parse(text, mode="eval")
        except SyntaxError as exc:
            raise FormulaError(f"无法解析职业点公式：{expression}") from exc
        value = self.visit(tree.body)
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise FormulaError("职业点公式没有得到数值")
        return int(value)

    def visit_Constant(self, node: ast.Constant) -> int | float:
        if isinstance(node.value, bool) or not isinstance(node.value, (int, float)):
            raise FormulaError("公式只允许整数常量")
        return node.value

    def visit_Name(self, node: ast.Name) -> int:
        key = node.id.upper()
        if key not in self.variables:
            raise FormulaError(f"公式引用了不允许的属性：{node.id}")
        return self.variables[key]

    def visit_BinOp(self, node: ast.BinOp) -> int | float:
        left = self.visit(node.left)
        right = self.visit(node.right)
        if isinstance(node.op, ast.Add):
            return left + right
        if isinstance(node.op, ast.Sub):
            return left - right
        if isinstance(node.op, ast.Mult):
            return left * right
        if isinstance(node.op, ast.Div):
            if right == 0:
                raise FormulaError("职业点公式不能除以零")
            return left / right
        raise FormulaError(f"公式不允许运算符：{type(node.op).__name__}")

    def visit_UnaryOp(self, node: ast.UnaryOp) -> int | float:
        operand = self.visit(node.operand)
        if isinstance(node.op, ast.UAdd):
            return operand
        if isinstance(node.op, ast.USub):
            return -operand
        raise FormulaError("公式只允许正负号")

    def visit_Call(self, node: ast.Call) -> int | float:
        if not isinstance(node.func, ast.Name):
            raise FormulaError("公式只允许调用 SUM 或 MAX")
        function_name = node.func.id.upper()
        if function_name not in {"SUM", "MAX"} or node.keywords:
            raise FormulaError("公式只允许调用 SUM 或 MAX")
        values = [self.visit(argument) for argument in node.args]
        if not values:
            raise FormulaError(f"{function_name} 至少需要一个参数")
        return sum(values) if function_name == "SUM" else max(values)

    def generic_visit(self, node):  # noqa: ANN001, ANN201
        raise FormulaError(f"公式包含不允许的语法：{type(node).__name__}")


class RuleEngine:
    """Deterministic COC7 calculations and validation for final entered values."""

    @staticmethod
    def skill_base_value(base_formula: str | None, attributes: Attributes, fallback: int) -> int:
        if not base_formula:
            return int(fallback)
        evaluator = SafeFormulaEvaluator(attributes.as_mapping())
        return evaluator.evaluate(base_formula)

    @staticmethod
    def occupation_points(point_formula: str, attributes: Attributes) -> int:
        return SafeFormulaEvaluator(attributes.as_mapping()).evaluate(point_formula)

    @staticmethod
    def calculate(
        final_attributes: Attributes,
        age: int,
        occupation_formula: str = "EDU*4",
        san_loss: int = 0,
    ) -> DerivedStats:
        values = final_attributes.as_mapping()
        hp = (values["CON"] + values["SIZ"]) // 10
        mp = values["POW"] // 5
        san = max(0, values["POW"] - int(san_loss))
        dodge = values["DEX"] // 2
        mov = RuleEngine._movement(values["STR"], values["DEX"], values["SIZ"], age)
        damage_bonus, build = RuleEngine._damage_bonus_and_build(values["STR"] + values["SIZ"])
        occupation_points = RuleEngine.occupation_points(occupation_formula, final_attributes)
        return DerivedStats(
            hp=hp,
            mp=mp,
            san=san,
            mov=mov,
            dodge=dodge,
            damage_bonus=damage_bonus,
            build=build,
            occupation_points=occupation_points,
            interest_points=values["INT"] * 2,
        )

    @staticmethod
    def _movement(strength: int, dexterity: int, size: int, age: int) -> int:
        if strength < size and dexterity < size:
            movement = 7
        elif strength > size and dexterity > size:
            movement = 9
        else:
            movement = 8
        if 15 <= age <= 19:
            movement += 1
        elif age >= 40:
            movement -= min(5, ((age - 40) // 10) + 1)
        return max(1, movement)

    @staticmethod
    def _damage_bonus_and_build(total: int) -> tuple[str, int]:
        if total <= 64:
            return "-2", -2
        if total <= 84:
            return "-1", -1
        if total <= 124:
            return "0", 0
        if total <= 164:
            return "+1D4", 1
        if total <= 204:
            return "+1D6", 2
        extra_steps = ((total - 205) // 80) + 1
        dice = 1 + extra_steps
        return f"+{dice}D6", 2 + extra_steps

    @staticmethod
    def asset_reference(credit_rating: int) -> dict[str, str]:
        """Return the standard 1920s dollar reference; every field remains editable."""
        credit = max(0, min(99, int(credit_rating)))
        if credit == 0:
            return {"living": "身无分文", "spending": "$0.50", "cash": "$0.50", "assets": "$0"}
        if credit <= 9:
            return {"living": "贫穷", "spending": "$2", "cash": f"${credit}", "assets": f"${credit * 10}"}
        if credit <= 49:
            return {"living": "标准", "spending": "$10", "cash": f"${credit * 2}", "assets": f"${credit * 50}"}
        if credit <= 89:
            return {"living": "小康", "spending": "$50", "cash": f"${credit * 5}", "assets": f"${credit * 500}"}
        if credit <= 98:
            return {"living": "富裕", "spending": "$250", "cash": f"${credit * 20}", "assets": f"${credit * 2_000}"}
        return {"living": "豪富", "spending": "$5,000", "cash": "$50,000", "assets": "$5,000,000+"}

    @staticmethod
    def validate(draft: CharacterDraft) -> ValidationReport:
        report = ValidationReport()
        if sum(is_branch_slot(skill.template_slot) for skill in draft.skills) > MAX_SKILL_BRANCHES:
            report.add(IssueSeverity.ERROR, "SKILL_BRANCH_LIMIT", f"新增技能分支不能超过 {MAX_SKILL_BRANCHES} 个。", "skills")
        if not draft.identity.name.strip():
            report.add(IssueSeverity.ERROR, "IDENTITY_NAME_REQUIRED", "调查员姓名不能为空。", "identity.name")
        if not 1 <= int(draft.identity.age) <= 120:
            report.add(IssueSeverity.ERROR, "AGE_RANGE", "年龄必须在 1 到 120 之间。", "identity.age")
        elif draft.identity.age <= 14 or draft.identity.age >= 90:
            report.add(IssueSeverity.WARNING, "AGE_UNUSUAL", "该年龄超出常规人类调查员范围，请与 KP 确认。", "identity.age")

        for key, value in draft.attributes.as_mapping().items():
            low = 9 if key == "SIZ" else 1
            if not low <= value <= 99:
                severity = IssueSeverity.WARNING if draft.nonstandard_override else IssueSeverity.ERROR
                report.add(
                    severity,
                    "ATTRIBUTE_RANGE",
                    f"{key}={value} 超出常规范围 {low}-99。" + ("已按 KP 特批继续。" if draft.nonstandard_override else "请修正或启用 KP 特批。"),
                    f"attributes.{key}",
                )

        if draft.occupation is None:
            report.add(IssueSeverity.ERROR, "OCCUPATION_REQUIRED", "必须选择或创建一个职业。", "occupation")
            expected_occupation_points = 0
            expected_interest_points = draft.attributes.INT * 2
        else:
            try:
                derived = RuleEngine.calculate(draft.attributes, draft.identity.age, draft.occupation.point_formula)
                expected_occupation_points = derived.occupation_points
                expected_interest_points = derived.interest_points
            except FormulaError as exc:
                report.add(IssueSeverity.ERROR, "OCCUPATION_FORMULA", str(exc), "occupation.point_formula")
                expected_occupation_points = 0
                expected_interest_points = draft.attributes.INT * 2

            for group in draft.occupation.choice_groups:
                selected = list(dict.fromkeys(draft.group_choices.get(group.marker, [])))
                if len(selected) != group.required_count:
                    report.add(
                        IssueSeverity.ERROR if len(selected) > group.required_count else IssueSeverity.INFO,
                        "OCCUPATION_GROUP_COUNT",
                        f"{group.label}最多选择 {group.required_count} 项，当前选择 {len(selected)} 项。",
                        f"group_choices.{group.marker}",
                    )
                invalid = [item for item in selected if item not in group.candidates]
                if invalid:
                    report.add(
                        IssueSeverity.ERROR,
                        "OCCUPATION_GROUP_INVALID",
                        f"{group.label}包含不属于该组的技能：{'、'.join(invalid)}。",
                        f"group_choices.{group.marker}",
                    )
            if len(set(draft.free_skill_choices)) != draft.occupation.free_choices:
                report.add(
                    IssueSeverity.ERROR if len(set(draft.free_skill_choices)) > draft.occupation.free_choices else IssueSeverity.INFO,
                    "OCCUPATION_FREE_COUNT",
                    f"任意特长最多选择 {draft.occupation.free_choices} 项，当前选择 {len(set(draft.free_skill_choices))} 项。",
                    "free_skill_choices",
                )

        used_occupation = sum(max(0, int(skill.occupation_points)) for skill in draft.skills)
        used_interest = sum(max(0, int(skill.interest_points)) for skill in draft.skills)
        if used_occupation != expected_occupation_points:
            report.add(
                IssueSeverity.ERROR if used_occupation > expected_occupation_points else IssueSeverity.INFO,
                "OCCUPATION_POINTS_TOTAL",
                f"职业点预算 {expected_occupation_points} 点，当前使用 {used_occupation} 点；未用完不影响导出。",
                "skills",
            )
        if used_interest != expected_interest_points:
            report.add(
                IssueSeverity.ERROR if used_interest > expected_interest_points else IssueSeverity.INFO,
                "INTEREST_POINTS_TOTAL",
                f"兴趣点预算 {expected_interest_points} 点，当前使用 {used_interest} 点；未用完不影响导出。",
                "skills",
            )

        credit_skill = None
        experience = draft.experience
        experience_used = sum(skill.experience_points for skill in draft.skills)
        if experience.name:
            if experience.san_loss < 0 or experience.san_loss > draft.attributes.POW or experience.skill_points < 0:
                report.add(IssueSeverity.ERROR, "EXPERIENCE_VALUES", "经历包技能点不能为负，SAN 减少值应在 0 与 POW 之间。", "experience")
            if draft.identity.age < experience.minimum_age:
                report.add(IssueSeverity.WARNING, "EXPERIENCE_AGE", f"{experience.name}建议年龄至少为 {experience.minimum_age} 岁，请与 KP 确认。", "experience")
            report.add(IssueSeverity.INFO, "EXPERIENCE_BACKGROUND", experience.notes or "请补充经历包相关背景。", "experience")
        if experience_used > (experience.skill_points if experience.name else 0):
            report.add(IssueSeverity.ERROR, "EXPERIENCE_BUDGET", "经历点超出所选经历包预算，请调整分配或选择经历包。", "experience")
        seen_names: set[str] = set()
        for skill in draft.skills:
            if min(skill.base_value, skill.occupation_points, skill.interest_points, skill.extra_final, skill.experience_points) < 0:
                report.add(IssueSeverity.ERROR, "SKILL_NEGATIVE", f"{skill.display_name} 的点数不能为负数。", skill.template_slot)
            if skill.final_value > 99:
                severity = IssueSeverity.WARNING if draft.nonstandard_override else IssueSeverity.ERROR
                report.add(
                    severity,
                    "SKILL_MAXIMUM",
                    f"{skill.display_name} 的最终值为 {skill.final_value}，高于常规上限 99。",
                    skill.template_slot,
                )
            if skill.occupation_points and skill.interest_points:
                report.add(IssueSeverity.WARNING, "SKILL_MIXED_POINTS", f"{skill.display_name} 同时使用了职业点和兴趣点（混点）。", skill.template_slot)
            if skill.occupation_points and not skill.selected_occupation:
                report.add(IssueSeverity.ERROR, "SKILL_NON_OCCUPATION", f"{skill.display_name} 不是已选择的职业技能，不能投入职业点。请退回这些点数，或将该技能选为职业技能。", skill.template_slot)
            if "克苏鲁神话" in skill.name:
                if skill.experience_points:
                    report.add(IssueSeverity.ERROR, "MYTHOS_EXPERIENCE", "克苏鲁神话请使用 KP 确定的额外增量，不使用普通经历点。", skill.template_slot)
                if skill.occupation_points or skill.interest_points:
                    report.add(IssueSeverity.ERROR, "MYTHOS_POINT_ALLOCATION", "初始克苏鲁神话不能投入职业点或兴趣点。", skill.template_slot)
                if skill.extra_final:
                    report.add(IssueSeverity.WARNING, "MYTHOS_OVERRIDE", "克苏鲁神话使用了 KP 特批的直接最终增量。", skill.template_slot)
            if skill.name.rstrip("：:") == "信用评级":
                credit_skill = skill
            folded = skill.display_name.casefold()
            if folded in seen_names and skill.final_value:
                report.add(IssueSeverity.WARNING, "SKILL_DUPLICATE", f"技能名称重复：{skill.display_name}。", skill.template_slot)
            seen_names.add(folded)

        if draft.occupation and credit_skill:
            credit_value = credit_skill.final_value
            if not draft.occupation.credit_min <= credit_value <= draft.occupation.credit_max:
                report.add(
                    IssueSeverity.ERROR,
                    "CREDIT_RANGE",
                    f"信用评级 {credit_value} 不在职业要求的 {draft.occupation.credit_range_text} 范围内。",
                    credit_skill.template_slot,
                )
            if draft.assets.credit_rating != credit_value:
                report.add(
                    IssueSeverity.WARNING,
                    "CREDIT_ASSET_MISMATCH",
                    f"资产页信用评级 {draft.assets.credit_rating} 与技能表信用评级 {credit_value} 不一致；导出以技能表为准。",
                    "assets.credit_rating",
                )
        elif draft.occupation:
            report.add(IssueSeverity.ERROR, "CREDIT_SKILL_MISSING", "技能表中缺少信用评级。", "skills")

        if len(draft.weapons) > 6:
            report.add(IssueSeverity.ERROR, "WEAPON_LIMIT", "纸质卡和模板最多导出六项武器。", "weapons")
        if len(draft.inventory) > 15:
            report.add(IssueSeverity.ERROR, "INVENTORY_LIMIT", "原模板最多导出十五项随身物品。", "inventory")

        optional_fields = {
            "形象描述": draft.background.appearance,
            "思想与信念": draft.background.beliefs,
            "重要之人": draft.background.significant_people,
            "背景故事": draft.background.personal_story,
            "资产说明": draft.assets.asset_description,
        }
        missing = [label for label, value in optional_fields.items() if not value.strip()]
        if missing:
            report.add(
                IssueSeverity.INFO,
                "PROFILE_INCOMPLETE",
                f"以下资料尚未填写，但不阻止导出：{'、'.join(missing)}。",
                "background",
            )
        return report
