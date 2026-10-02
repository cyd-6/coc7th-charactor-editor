"""Exercise the browser's skill state and grouped rendering without a browser."""

import json
import os
from pathlib import Path
import shutil
import subprocess

import pytest

from app import get_catalog
from coc7_card.web import catalog_payload


ROOT = Path(__file__).resolve().parents[1]
NODE_HARNESS = r"""
const fs = require("node:fs");
const vm = require("node:vm");
const assert = require("node:assert/strict");
const {catalog, scenario} = JSON.parse(fs.readFileSync(0, "utf8"));

// Only DOM storage is simulated. State transitions, HTML generation, budgets,
// skill totals, equipment rendering, and draft serialization use app.js itself.
function element(tagName = "div") {
  const selected = new Map();
  return {
    tagName, value: "", textContent: "", innerHTML: "", hidden: false,
    style: {}, dataset: {}, children: [],
    classList: {toggle() {}},
    addEventListener() {},
    focus() {}, scrollIntoView() {},
    showModal() { this.open = true; },
    close() { this.open = false; },
    setCustomValidity(message) { this.validationMessage = message; },
    reportValidity() { return !this.validationMessage; },
    append(...children) { this.children.push(...children); },
    appendChild(child) { this.children.push(child); },
    querySelector(selector) {
      const child = this.children.find((item) => item.tagName === selector);
      if (child) return child;
      if (!selected.has(selector)) selected.set(selector, element());
      return selected.get(selector);
    },
    querySelectorAll() { return []; },
  };
}
const elements = new Map();
const document = {
  addEventListener() {},
  getElementById(id) {
    if (!elements.has(id)) elements.set(id, element());
    return elements.get(id);
  },
  createElement: element,
  querySelector() { return null; },
  querySelectorAll() { return []; },
};
document.getElementById("skill-filter").value = "all";
const context = vm.createContext({
  assert, catalog, document, element, structuredClone, crypto: require("node:crypto").webcrypto,
  // Do not start the application or persist a synthetic draft asynchronously.
  setTimeout() { return 1; }, clearTimeout() {},
});
vm.runInContext(fs.readFileSync("static/app.js", "utf8"), context, {filename: "static/app.js"});
vm.runInContext("bootstrapData = JSON.parse(JSON.stringify(catalog)); state = createDefaultState();", context);
vm.runInContext(scenario, context, {filename: "skill-group-regression.js"});
"""


@pytest.fixture(scope="module")
def run_browser_scenario():
    configured_node = os.environ.get("COC7_NODE")
    if configured_node is not None:
        node = shutil.which(configured_node)
        if not node:
            pytest.fail("COC7_NODE must point to an executable Node.js program")
    else:
        node = shutil.which("node")
        if not node:
            pytest.skip("Node.js is required to verify browser skill groups")
    catalog = catalog_payload(get_catalog())

    def run(scenario):
        result = subprocess.run(
            [node, "-e", NODE_HARNESS],
            input=json.dumps({"catalog": catalog, "scenario": scenario}),
            text=True,
            capture_output=True,
            cwd=ROOT,
            timeout=15,
        )
        assert result.returncode == 0, result.stdout + result.stderr

    return run


def test_new_draft_has_independent_common_shooting_specializations(run_browser_scenario):
    run_browser_scenario(r"""
const shooting = state.skills.filter((skill) => skillGroupName(skill) === "射击");
assert.deepEqual(shooting.map((skill) => [skill.template_slot, skill.specialization, skill.base_value]), [
  ["F38", "手枪", 20], ["F39", "步枪/霰弹枪", 25],
  ["F40", "冲锋枪", 15], ["F41", "弓术", 15],
]);
shooting.forEach((skill, index) => { skill.interest_points = index + 1; });
assert.deepEqual(shooting.map(skillFinal), [21, 27, 18, 19]);
assert.deepEqual(shooting.map(skillLabel), [
  "射击（手枪）", "射击（步枪/霰弹枪）", "射击（冲锋枪）", "射击（弓术）",
]);
assert.equal(state.skills.length, 67);
assert.equal(new Set(state.skills.map((skill) => skill.template_slot)).size, 67);

// Real occupation tokens retain their short names while using the named base.
for (const [occupationName, slot, specialization, base] of [
  ["绅士、淑女", "F39", "步/霰", 25],
  ["工人-伐木工", "F35", "链锯", 10],
]) {
  state = createDefaultState();
  state.occupation_id = bootstrapData.occupations.find((item) => item.name === occupationName).occupation_id;
  syncOccupationSkills();
  assert.equal(skillBySlot(slot).specialization, specialization);
  assert.equal(skillBySlot(slot).base_value, base);
}

// A custom skill occupying a former shooting slot is no longer a shooting subskill.
Object.assign(skillBySlot("F40"), {name: "档案研究", specialization: "冲锋枪"});
syncDynamicSkillBases();
assert.equal(skillGroupName(skillBySlot("F40")), "");
assert.equal(skillBySlot("F40").base_value, 1);
assert.equal(skillBySlot("F35").base_value, 10);
""")


def test_normalizing_legacy_draft_preserves_empty_names_bases_and_allocations(run_browser_scenario):
    run_browser_scenario(r"""
state.occupation_mode = "custom";
state.skills = [
  {template_slot: "F39", specialization: "", base_override: 0,
   occupation_points: 11, interest_points: 12, extra_final: 13, experience_points: 14},
  {template_slot: "F40", name: "自制投射器", specialization: "自订型号", base_override: 37,
   occupation_points: 2, interest_points: 3, extra_final: 4, experience_points: 5},
];
normalizeState();
syncOccupationSkills();
const blank = skillBySlot("F39");
assert.equal(blank.specialization, "");
assert.equal(blank.auto_specialization, false);
assert.equal(blank.base_override, 0);
assert.equal(blank.base_value, 0);
assert.deepEqual([blank.occupation_points, blank.interest_points, blank.extra_final, blank.experience_points], [11, 12, 13, 14]);
assert.equal(skillFinal(blank), 50);
const custom = skillBySlot("F40");
assert.equal(custom.name, "自制投射器");
assert.equal(custom.specialization, "自订型号");
assert.equal(custom.base_value, 37);
assert.equal(skillGroupName(custom), "");
assert.deepEqual([custom.occupation_points, custom.interest_points, custom.extra_final, custom.experience_points], [2, 3, 4, 5]);
assert.equal(skillBySlot("F41").specialization, "弓术");
assert.equal(state.skills.length, 67);
""")


def test_occupation_specialization_replaces_only_untouched_presets(run_browser_scenario):
    run_browser_scenario(r"""
state.occupation_id = 172; // 女学生 requires 射击①（弓术） in the actual catalog.
syncOccupationSkills();
assert.equal(skillBySlot("F39").specialization, "弓术");
assert.equal(skillBySlot("F39").base_value, 15);
assert.equal(skillBySlot("F39").selected_occupation, true);
assert.equal(skillBySlot("F41").specialization, "步枪/霰弹枪");
assert.equal(skillBySlot("F41").base_value, 25);
assert.equal(skillBySlot("F35").specialization, "矛");
assert.equal(skillBySlot("F35").base_value, 20);
assert.equal(skillBySlot("F37").specialization, "斧");
assert.equal(skillBySlot("F37").base_value, 15);
for (const group of ["射击", "格斗"]) {
  const skills = state.skills.filter((skill) => skillGroupName(skill) === group);
  assert.equal(new Set(skills.map((skill) => skill.specialization)).size, 4);
}
state.occupation_mode = "custom";
syncOccupationSkills();
assert.equal(skillBySlot("F39").specialization, "步枪/霰弹枪");
assert.equal(skillBySlot("F39").base_value, 25);
assert.equal(skillBySlot("F41").specialization, "弓术");
assert.equal(skillBySlot("F35").specialization, "斧");
assert.equal(skillBySlot("F37").specialization, "矛");

for (const edited of [false, true]) {
  state = createDefaultState();
  state.occupation_id = 172;
  syncOccupationSkills();
  const bow = skillBySlot("F39");
  if (edited) bow.auto_specialization = false;
  else bow.interest_points = 5;
  state.occupation_mode = "custom";
  syncOccupationSkills();
  assert.equal(bow.specialization, "弓术");
  assert.equal(bow.base_value, 15);
  assert.equal(bow.interest_points, edited ? 0 : 5);
  assert.equal(skillBySlot("F41").specialization, "步枪/霰弹枪");
  assert.equal(skillBySlot("F41").base_value, 25);
  assert.equal(new Set(state.skills.filter((skill) => skillGroupName(skill) === "射击")
    .map((skill) => skill.specialization)).size, 4);
}

for (const field of ["occupation_points", "interest_points", "extra_final", "experience_points"]) {
  state = createDefaultState();
  const skill = skillBySlot("F39");
  skill[field] = 7;
  state.occupation_id = 172;
  syncOccupationSkills();
  assert.equal(skill.specialization, "步枪/霰弹枪", field);
  assert.equal(skill[field], 7);
  assert.equal(skill.base_value, 25);
}
state = createDefaultState();
Object.assign(skillBySlot("F39"), {specialization: "手工改名", auto_specialization: false});
state.occupation_id = 172;
syncOccupationSkills();
assert.equal(skillBySlot("F39").specialization, "手工改名");
state = createDefaultState();
skillBySlot("F39").base_override = 0;
state.occupation_id = 172;
syncOccupationSkills();
assert.equal(skillBySlot("F39").specialization, "步枪/霰弹枪");
assert.equal(skillBySlot("F39").base_value, 0);
""")


def test_search_expands_matching_subskill_without_losing_folded_rows(run_browser_scenario):
    run_browser_scenario(r"""
const body = document.getElementById("skill-table-body");
const search = document.getElementById("skill-search");
const rows = () => [...body.innerHTML.matchAll(/<tr\b[^>]*data-skill-row="([^"]+)"[^>]*>/g)];
renderSkills();
assert.equal(rows().length, 67);
assert.equal(new Set(rows().map((match) => match[1])).size, 67);
for (const group of ["射击", "格斗", "科学", "技艺", "外语"]) {
  assert.ok(body.innerHTML.includes(`data-skill-group-toggle="${group}" aria-expanded="false"`), group);
}
assert.ok(rows().filter((match) => match[0].includes('data-skill-group="射击"')).every((match) => /\bhidden\b/.test(match[0])));
search.value = "冲锋枪";
renderSkills();
assert.deepEqual(rows().map((match) => match[1]), ["F40"]);
assert.ok(!/\bhidden\b/.test(rows()[0][0]));
assert.ok(body.innerHTML.includes('data-skill-group-toggle="射击" aria-expanded="true"'));
search.value = "";
renderSkills();
assert.equal(rows().length, 67);
assert.ok(body.innerHTML.includes('data-skill-group-toggle="射击" aria-expanded="false"'));
assert.equal(serializeDraft().skills.length, 67);
search.value = "不存在的技能分项";
renderSkills();
assert.equal(rows().length, 0);
assert.equal(document.getElementById("skill-empty").hidden, false);
assert.equal(serializeDraft().skills.length, 67);
""")


def test_rename_keeps_all_point_kinds_and_updates_only_matching_weapon(run_browser_scenario):
    run_browser_scenario(r"""
const skill = skillBySlot("F40");
Object.assign(skill, {occupation_points: 11, interest_points: 12, extra_final: 13, experience_points: 14});
const oldName = skillDisplayName(skill);
state.weapons = [
  {name: "随改名更新的武器", skill: oldName},
  {name: "独立武器", skill: "射击（其他自订分项）"},
];
const siblingBefore = JSON.stringify(skillBySlot("F39"));
const row = element("tr");
const input = element("input");
input.dataset.specializationSlot = "F40";
input.value = " 机枪 ";
input.closest = (selector) => selector === "tr" ? row : input;
handleSpecializationChange({target: input});
assert.equal(skill.specialization, "机枪");
assert.equal(input.value, "机枪");
assert.equal(skill.template_slot, "F40");
assert.equal(skill.name, "射击②");
assert.equal(skill.auto_specialization, false);
assert.equal(skill.base_value, 10);
assert.deepEqual([skill.occupation_points, skill.interest_points, skill.extra_final, skill.experience_points], [11, 12, 13, 14]);
assert.equal(JSON.stringify(skillBySlot("F39")), siblingBefore);
assert.equal(state.weapons[0].skill, "射击②（机枪）");
assert.equal(state.weapons[1].skill, "射击（其他自订分项）");
assert.equal(row.querySelector('[data-result="base"]').textContent, 10);
assert.equal(row.querySelector('[data-result="final"]').textContent, 60);
assert.equal(row.querySelector('[data-result="hard"]').textContent, 30);
assert.equal(row.querySelector('[data-result="extreme"]').textContent, 12);
assert.ok(document.getElementById("weapon-table-body").innerHTML.includes('value="射击②（机枪）" selected>射击（机枪） · 60'));
const serialized = serializeDraft();
assert.equal(serialized.skills.find((item) => item.template_slot === "F40").experience_points, 14);
assert.equal(serialized.weapons[0].skill, "射击②（机枪）");
""")


def test_additional_branches_survive_normalization_with_independent_points(run_browser_scenario):
    run_browser_scenario(r"""
const originalShooting = JSON.stringify(state.skills.filter((skill) => skillGroupName(skill) === "射击"));
document.getElementById("skill-branch-group").value = "射击";
for (const name of ["机枪", "重武器", "自定义枪械"]) {
  document.getElementById("skill-branch-name").value = name;
  addSkillBranch({preventDefault() {}});
}
const branches = state.skills.filter((skill) => isBranchSlot(skill.template_slot));
assert.equal(branches.length, 3);
assert.equal(new Set(branches.map((skill) => skill.template_slot)).size, 3);
assert.equal(state.skills.filter((skill) => skillGroupName(skill) === "射击").length, 7);
assert.equal(JSON.stringify(state.skills.filter((skill) => skillGroupName(skill) === "射击" && !isBranchSlot(skill.template_slot))), originalShooting);
assert.deepEqual(branches.map((skill) => skill.base_value), [10, 10, 1]);
branches.forEach((skill, index) => { skill.interest_points = 10 + index; skill.experience_points = 2; skill.extra_final = 3; });
branches[1].base_override = 0;
state.occupation_mode = "custom";
state.custom_skill_slots = [branches[0].template_slot];
syncOccupationSkills();
branches[0].occupation_points = 5;
const before = JSON.parse(JSON.stringify(state.skills));
normalizeState();
syncOccupationSkills();
assert.deepEqual(state.skills, before);
assert.equal(skillBySlot(branches[1].template_slot).base_value, 0);
assert.equal(skillBySlot(branches[0].template_slot).selected_occupation, true);
assert.equal(skillBySlot(branches[1].template_slot).selected_occupation, false);
renderSkills();
assert.equal(serializeDraft().skills.length, 70);
document.getElementById("skill-search").value = "重武器";
renderSkills();
assert.ok(document.getElementById("skill-table-body").innerHTML.includes(`data-skill-row="${branches[1].template_slot}"`));
assert.equal(serializeDraft().skills.length, 70);
""")


def test_duplicate_branch_rejected_and_remove_refunds_only_target(run_browser_scenario):
    run_browser_scenario(r"""
document.getElementById("skill-branch-group").value = "射击";
const nameInput = document.getElementById("skill-branch-name");
for (const name of ["步/霰", "手枪"]) {
  nameInput.value = name;
  addSkillBranch({preventDefault() {}});
  assert.ok(nameInput.validationMessage.includes("已存在"));
  assert.equal(state.skills.length, 67);
}
for (const name of ["机枪", "重武器"]) {
  nameInput.value = name;
  addSkillBranch({preventDefault() {}});
}
const [target, sibling] = state.skills.filter((skill) => isBranchSlot(skill.template_slot));
Object.assign(target, {interest_points: 15, extra_final: 4, experience_points: 3});
sibling.interest_points = 7;
state.occupation_mode = "custom";
state.custom_skill_slots = [target.template_slot, sibling.template_slot];
syncOccupationSkills();
target.occupation_points = 10;
state.weapons = [{name: "测试武器", skill: skillDisplayName(target), damage: "1D6"}];
confirmRemoveSkillBranch(target.template_slot);
assert.equal(state.skills.length, 69); // Opening confirmation leaves the draft intact.
assert.equal(skillBySlot(target.template_slot).interest_points, 15);
document.getElementById("skill-branch-remove-dialog").close();
assert.equal(state.skills.length, 69);
confirmRemoveSkillBranch(target.template_slot);
removeSkillBranch();
assert.equal(state.skills.length, 68);
assert.equal(skillBySlot(target.template_slot), undefined);
assert.equal(skillBySlot(sibling.template_slot).interest_points, 7);
assert.deepEqual(state.custom_skill_slots, [sibling.template_slot]);
assert.equal(document.getElementById("int-used").textContent, "7");
assert.equal(document.getElementById("occ-used").textContent, "0");
assert.equal(state.weapons[0].name, "测试武器");
assert.equal(state.weapons[0].damage, "1D6");
assert.equal(state.weapons[0].skill, "");
""")


def test_specific_occupation_reuses_existing_branch_without_duplicate_presets(run_browser_scenario):
    run_browser_scenario(r"""
document.getElementById("skill-branch-group").value = "科学";
document.getElementById("skill-branch-name").value = "数学";
addSkillBranch({preventDefault() {}});
const branch = state.skills.find((skill) => isBranchSlot(skill.template_slot));
branch.interest_points = 3;
const architect = bootstrapData.occupations.find((occupation) => occupation.name === "建筑师");
assert.ok(architect);
state.occupation_id = architect.occupation_id;
syncOccupationSkills();
assert.equal(branch.selected_occupation, true);
assert.equal(branch.base_value, 10);
assert.equal(branch.interest_points, 3);
assert.notEqual(skillBySlot("AB31").specialization, "数学");
assert.equal(skillBySlot("AB31").selected_occupation, false);
assert.equal(state.skills.filter((skill) => skillGroupName(skill) === "科学" && skill.specialization === "数学").length, 1);
confirmRemoveSkillBranch(branch.template_slot);
removeSkillBranch();
assert.equal(skillBySlot("AB31").specialization, "数学");
assert.equal(skillBySlot("AB31").selected_occupation, true);
assert.equal(skillBySlot("AB31").base_value, 10);
assert.equal(state.skills.filter((skill) => isBranchSlot(skill.template_slot)).length, 0);
""")


def test_branch_limit_allows_two_hundred_then_preserves_existing_points(run_browser_scenario):
    run_browser_scenario(r"""
assert.equal(bootstrapData.meta.max_skill_branches, 200);
state.occupation_mode = "custom";
for (let index = 1; index <= 199; index++) {
  const slot = `branch-${index.toString(16).padStart(32, "0")}`;
  state.skills.push({key: slot, template_slot: slot, name: "科学", specialization: `测试领域${index}`,
    auto_specialization: false, base_value: 1, occupation_points: 0, interest_points: index % 2,
    extra_final: index % 3, experience_points: 0, selected_occupation: false});
}
skillBySlot("F16").interest_points = 9;
const oldPoints = () => state.skills.slice(0, 67 + 199).map((skill) => [skill.template_slot,
  skill.occupation_points, skill.interest_points, skill.extra_final, skill.experience_points]);
const before = JSON.stringify(oldPoints());
document.getElementById("skill-branch-group").value = "科学";
document.getElementById("skill-branch-name").value = "第200个分支";
addSkillBranch({preventDefault() {}});
assert.equal(state.skills.length, 67 + 200);
assert.equal(JSON.stringify(oldPoints()), before);
const accepted = JSON.stringify(state.skills);
document.getElementById("skill-branch-name").value = "第201个分支";
addSkillBranch({preventDefault() {}});
assert.equal(JSON.stringify(state.skills), accepted);
assert.ok(document.getElementById("toast-region").children.at(-1).textContent.includes("不能超过 200 个"));
assert.equal(serializeDraft().skills.length, 67 + 200);
""")


def test_legacy_oversized_draft_is_not_truncated_and_can_remove_one_branch(run_browser_scenario):
    run_browser_scenario(r"""
state.occupation_mode = "custom";
for (let index = 1; index <= 201; index++) {
  const slot = `branch-${index.toString(16).padStart(32, "0")}`;
  state.skills.push({key: slot, template_slot: slot, name: "科学", specialization: `旧领域${index}`,
    auto_specialization: false, base_value: 1, occupation_points: 0, interest_points: index % 2,
    extra_final: index % 3, experience_points: 0, selected_occupation: false});
}
const branches = () => state.skills.filter((skill) => isBranchSlot(skill.template_slot));
const before = JSON.parse(JSON.stringify(branches()));
normalizeState();
assert.deepEqual(branches(), before);
assert.equal(serializeDraft().skills.length, 67 + 201);
const removed = branches().at(-1).template_slot;
const survivors = JSON.parse(JSON.stringify(branches().slice(0, -1)));
confirmRemoveSkillBranch(removed);
assert.equal(branches().length, 201);
removeSkillBranch();
assert.equal(branches().length, 200);
assert.deepEqual(branches(), survivors);
assert.equal(skillBySlot(removed), undefined);
assert.equal(serializeDraft().skills.length, 67 + 200);
""")
