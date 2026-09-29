"use strict";

const STORAGE_KEY = "coc7-investigator-draft-v2";
const LOCAL_SERVICE_MESSAGE = "无法连接车卡服务，请确认服务已启动、访问地址正确，然后重试。";
const SERVICE_TIMEOUT_MESSAGE = "车卡服务响应时间过长，请稍后重试。";
const ATTRIBUTE_KEYS = ["STR", "CON", "SIZ", "DEX", "APP", "INT", "POW", "EDU", "Luck"];
const ATTRIBUTE_LABELS = { STR: "力量", CON: "体质", SIZ: "体型", DEX: "敏捷", APP: "外貌", INT: "智力", POW: "意志", EDU: "教育", Luck: "幸运" };
const STEP_LABELS = ["调查员信息", "最终属性", "职业配置", "技能分配", "背景资产", "武器物品", "检查导出"];
const BACKGROUND_FIELDS = {
  "bg-appearance": "appearance",
  "bg-beliefs": "beliefs",
  "bg-people": "significant_people",
  "bg-places": "meaningful_places",
  "bg-possessions": "treasured_possessions",
  "bg-traits": "traits",
  "bg-scars": "scars",
  "bg-phobias": "phobias_manias",
  "bg-story": "personal_story",
  "bg-contacts": "contacts_notes",
};
const ASSET_FIELDS = {
  "asset-living": "living_standard",
  "asset-spending": "usd_spending",
  "asset-cash": "usd_cash",
  "asset-other-assets": "usd_assets",
  "asset-description": "asset_description",
  "asset-vehicles": "vehicles",
  "asset-residence": "residence",
  "asset-luxuries": "luxuries",
  "asset-securities": "securities",
  "asset-other": "other",
};

let bootstrapData = null;
let state = null;
let currentStep = 0;
let portraitFile = null;
let portraitUrl = null;
let pdfObjectUrl = null;
let validationData = null;
let saveTimer = null;
let derivedTimer = null;
let responsiveFrame = null;
let responsiveObserver = null;
let densityMediaQuery = null;
let densityMediaHandler = null;
let lastResponsiveSignature = "";
let lastToastMessage = "";
let lastToastAt = 0;
let draftGeneration = 0;
let importBusy = false;
let pendingImport = null;
let lastImportSnapshot = null;
let pendingAttributeImport = null;
let lastAttributeImportSnapshot = null;

document.addEventListener("DOMContentLoaded", initialize);

async function initialize() {
  try {
    installResponsiveController();
    const response = await apiFetch("/api/bootstrap", { cache: "no-store" }, { timeoutMs: 12000, retries: 1 });
    if (!response.ok) throw new Error("无法读取职业与技能目录。");
    bootstrapData = await readJsonResponse(response, "无法读取职业与技能目录。");
    state = loadState() || createDefaultState();
    normalizeState();
    bindStaticControls();
    hydrateControls();
    renderCatalogMeta();
    renderOccupation();
    syncDynamicSkillBases();
    syncOccupationSkills();
    renderSkills();
    renderEquipment();
    renderAssetReference();
    renderLiveSummary();
    goToStep(Number(state.current_step || 0), false);
    await refreshDerived(true);
    document.getElementById("boot-screen").hidden = true;
    document.getElementById("app-frame").hidden = false;
  } catch (error) {
    renderBootFailure(error);
  }
}

function renderBootFailure(error) {
  const message = friendlyErrorMessage(error, "车卡器初始化失败，请重新连接车卡服务。");
  const bootScreen = document.getElementById("boot-screen");
  bootScreen.innerHTML = `<div class="boot-seal">!</div><p>${escapeHtml(message)}</p><button class="boot-retry" type="button">重新连接</button>`;
  bootScreen.querySelector(".boot-retry").addEventListener("click", () => window.location.reload());
}

async function apiFetch(resource, options = {}, policy = {}) {
  const timeoutMs = Number(policy.timeoutMs ?? 15000);
  const retries = Math.max(0, Number(policy.retries ?? 0));
  const retryDelayMs = Math.max(0, Number(policy.retryDelayMs ?? 450));

  for (let attempt = 0; attempt <= retries; attempt += 1) {
    const controller = new AbortController();
    let timedOut = false;
    const timeout = timeoutMs > 0 ? window.setTimeout(() => {
      timedOut = true;
      controller.abort();
    }, timeoutMs) : null;

    try {
      return await fetch(resource, { ...options, signal: controller.signal });
    } catch (error) {
      const mayRetry = attempt < retries && (timedOut || isNetworkFailure(error));
      if (mayRetry) {
        await delay(retryDelayMs * (attempt + 1));
        continue;
      }
      if (timedOut) throw new Error(SERVICE_TIMEOUT_MESSAGE);
      throw new Error(friendlyErrorMessage(error));
    } finally {
      if (timeout !== null) window.clearTimeout(timeout);
    }
  }

  throw new Error(LOCAL_SERVICE_MESSAGE);
}

async function readJsonResponse(response, fallbackMessage) {
  try {
    return await response.json();
  } catch (_) {
    if (!response.ok) throw new Error(fallbackMessage);
    throw new Error("车卡服务返回的数据无法识别，请重新操作一次。");
  }
}

function isNetworkFailure(error) {
  if (error?.name === "AbortError") return true;
  const message = String(error?.message || error || "").toLowerCase();
  return /failed\s+to\s+fetch|network\s*error|network request failed|load failed|err_connection|connection refused/.test(message);
}

function friendlyErrorMessage(error, fallbackMessage = "操作未完成，请稍后重试。") {
  if (typeof error === "string" && error.trim()) {
    if (isNetworkFailure(error)) return LOCAL_SERVICE_MESSAGE;
    return error.trim();
  }
  if (error?.name === "AbortError") return SERVICE_TIMEOUT_MESSAGE;
  const message = String(error?.message || error || "").trim();
  if (!message) return fallbackMessage;
  if (isNetworkFailure(error)) return LOCAL_SERVICE_MESSAGE;
  if (/unexpected (token|end)|json/i.test(message)) return "车卡服务返回的数据无法识别，请重新操作一次。";
  return message;
}

function delay(milliseconds) {
  return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
}

function installResponsiveController() {
  const schedule = () => {
    if (responsiveFrame !== null) cancelAnimationFrame(responsiveFrame);
    responsiveFrame = requestAnimationFrame(() => {
      responsiveFrame = null;
      applyResponsiveLayout();
    });
  };
  window.addEventListener("resize", schedule, { passive: true });
  window.addEventListener("orientationchange", schedule, { passive: true });
  if (window.visualViewport) {
    window.visualViewport.addEventListener("resize", schedule, { passive: true });
    window.visualViewport.addEventListener("scroll", schedule, { passive: true });
  }
  if (window.ResizeObserver) {
    responsiveObserver = new ResizeObserver(schedule);
    responsiveObserver.observe(document.documentElement);
  }
  watchDevicePixelRatio(schedule);
  applyResponsiveLayout();
}

function watchDevicePixelRatio(schedule) {
  if (!window.matchMedia) return;
  if (densityMediaQuery && densityMediaHandler) {
    if (densityMediaQuery.removeEventListener) densityMediaQuery.removeEventListener("change", densityMediaHandler);
    else if (densityMediaQuery.removeListener) densityMediaQuery.removeListener(densityMediaHandler);
  }
  densityMediaQuery = window.matchMedia(`(resolution: ${window.devicePixelRatio || 1}dppx)`);
  densityMediaHandler = () => {
    watchDevicePixelRatio(schedule);
    schedule();
  };
  if (densityMediaQuery.addEventListener) densityMediaQuery.addEventListener("change", densityMediaHandler, { once: true });
  else if (densityMediaQuery.addListener) densityMediaQuery.addListener(densityMediaHandler);
}

function applyResponsiveLayout() {
  const visualWidth = window.visualViewport?.width || document.documentElement.clientWidth || window.innerWidth;
  const visualHeight = window.visualViewport?.height || document.documentElement.clientHeight || window.innerHeight;
  const width = Math.max(320, Math.round(Math.min(window.innerWidth || visualWidth, visualWidth)));
  const height = Math.max(320, Math.round(visualHeight));
  let layout = "desktop";
  if (width <= 650) layout = "compact";
  else if (width <= 900) layout = "tablet";
  else if (width <= 1280) layout = "laptop";
  else if (width >= 2560) layout = "ultra";
  else if (width >= 1800) layout = "wide";
  const orientation = width >= height ? "landscape" : "portrait";
  const density = Math.max(0.9, Math.min(1.18, width / 1600));
  const workspaceWidth = calculateWorkspaceWidth(width);
  const signature = `${layout}:${orientation}:${width}:${height}:${window.devicePixelRatio || 1}`;
  if (signature === lastResponsiveSignature) return;
  lastResponsiveSignature = signature;
  const root = document.documentElement;
  root.dataset.layout = layout;
  root.dataset.orientation = orientation;
  root.style.setProperty("--app-height", `${height}px`);
  root.style.setProperty("--visual-width", `${width}px`);
  root.style.setProperty("--adaptive-density", density.toFixed(3));
  root.style.setProperty("--workspace-max", `${workspaceWidth}px`);
}

function calculateWorkspaceWidth(viewportWidth) {
  const width = Math.max(320, Number(viewportWidth) || 320);
  if (width <= 1800) return Math.round(width);
  return Math.round(Math.min(2160, width * 0.45 + 990));
}

function createDefaultState() {
  const today = new Date();
  const localDate = new Date(today.getTime() - today.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
  const attributes = Object.fromEntries(ATTRIBUTE_KEYS.map((key) => [key, 50]));
  return {
    version: 3,
    current_step: 0,
    identity: { name: "", player: "", age: 30, gender: "", residence: "", birthplace: "", era: "1920s", current_date: localDate },
    attributes,
    nonstandard_override: false,
    occupation_mode: "catalog",
    occupation_id: bootstrapData.occupations[0]?.occupation_id || 2,
    custom_occupation: { name: "自定义职业", credit_min: 10, credit_max: 50, point_formula_kind: "EDU*4", secondary: "DEX" },
    custom_skill_slots: [],
    group_choices: {},
    free_skill_choices: [],
    experience: { selection: "", name: "", san_loss: 0, skill_points: 0, notes: "" },
    skills: bootstrapData.skills.map((definition) => makeSkillState(definition, attributes)),
    background: { appearance: "", beliefs: "", significant_people: "", meaningful_places: "", treasured_possessions: "", traits: "", scars: "", phobias_manias: "", personal_story: "", key_connection: "", contacts_notes: "" },
    assets: { living_standard: "", spending_level: "", cash: "", other_assets: "", asset_description: "", vehicles: "", residence: "", luxuries: "", securities: "", other: "", currency: "美元" },
    weapons: [],
    inventory: [],
    derived: { hp: 10, mp: 10, san: 50, mov: 8, dodge: 25, damage_bonus: "0", build: 0, occupation_points: 200, interest_points: 100 },
  };
}

function makeSkillState(definition, attributes) {
  return {
    key: definition.template_slot,
    template_slot: definition.template_slot,
    name: definition.name,
    specialization: definition.default_specialization || "",
    base_value: dynamicBase(definition, attributes),
    occupation_points: 0,
    interest_points: 0,
    extra_final: 0,
    experience_points: 0,
    selected_occupation: false,
  };
}

function loadState() {
  try {
    const raw = sessionStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed && [2, 3].includes(parsed.version) ? parsed : null;
  } catch (_) {
    return null;
  }
}

function normalizeState() {
  const defaults = createDefaultState();
  state.identity = { ...defaults.identity, ...(state.identity || {}) };
  state.attributes = { ...defaults.attributes, ...(state.attributes || {}) };
  state.custom_occupation = { ...defaults.custom_occupation, ...(state.custom_occupation || {}) };
  state.background = { ...defaults.background, ...(state.background || {}) };
  state.experience = { ...defaults.experience, ...(state.experience || {}) };
  state.assets = { ...defaults.assets, ...(state.assets || {}) };
  state.assets.exchange_year = numeric(state.assets.exchange_year, state.identity.era === "现代" ? 2026 : 1920);
  state.identity.story_year = state.identity.story_year ?? state.assets.exchange_year;
  state.assets.exchange_year = Number(state.identity.story_year);
  state.assets.currency_code = state.assets.currency_code ?? "USD";
  if (state.version === 2) state.assets.currency_code = Coc7Currency.migrateCode(state.assets.currency_code, state.assets.exchange_year);
  state.version = 3;
  for (const [key, old] of [["usd_spending", "spending_level"], ["usd_cash", "cash"], ["usd_assets", "other_assets"]]) {
    if (!(key in state.assets)) state.assets[key] = state.assets[old] || "";
  }
  state.group_choices = state.group_choices && typeof state.group_choices === "object" ? state.group_choices : {};
  state.free_skill_choices = Array.isArray(state.free_skill_choices) ? state.free_skill_choices : [];
  state.custom_skill_slots = Array.isArray(state.custom_skill_slots) ? state.custom_skill_slots : [];
  state.weapons = Array.isArray(state.weapons) ? state.weapons.slice(0, 6) : [];
  state.inventory = Array.isArray(state.inventory) ? state.inventory.slice(0, 15) : [];
  state.occupation_mode = state.occupation_mode === "custom" ? "custom" : "catalog";
  if (!bootstrapData.occupations.some((item) => item.occupation_id === Number(state.occupation_id))) {
    state.occupation_id = bootstrapData.occupations[0]?.occupation_id || 2;
  }
  const oldBySlot = new Map((Array.isArray(state.skills) ? state.skills : []).map((item) => [item.template_slot, item]));
  state.skills = bootstrapData.skills.map((definition) => ({
    ...makeSkillState(definition, state.attributes),
    ...(oldBySlot.get(definition.template_slot) || {}),
    key: definition.template_slot,
    template_slot: definition.template_slot,
    name: oldBySlot.get(definition.template_slot)?.name || definition.name,
  }));
  state.current_step = clampNumber(state.current_step, 0, 6, 0);
}

function bindStaticControls() {
  document.querySelectorAll("[data-step-target]").forEach((button) => button.addEventListener("click", () => goToStep(Number(button.dataset.stepTarget))));
  document.getElementById("brand-home").addEventListener("click", (event) => { event.preventDefault(); goToStep(0); });
  document.getElementById("previous-step").addEventListener("click", () => goToStep(currentStep - 1));
  document.getElementById("next-step").addEventListener("click", () => goToStep(currentStep + 1));

  const identityBindings = {
    "identity-name": "name", "identity-player": "player", "identity-age": "age", "identity-gender": "gender",
    "identity-residence": "residence", "identity-birthplace": "birthplace", "identity-era": "era", "identity-date": "current_date",
  };
  Object.entries(identityBindings).forEach(([id, key]) => {
    document.getElementById(id).addEventListener("input", (event) => {
      state.identity[key] = key === "age" ? numeric(event.target.value, 30) : event.target.value;
      markChanged();
      renderLiveSummary();
      if (key === "age") refreshDerived();
    });
  });

  document.querySelectorAll("[data-attribute]").forEach((input) => input.addEventListener("input", (event) => {
    state.attributes[event.target.dataset.attribute] = numeric(event.target.value, 0);
    syncDynamicSkillBases();
    markChanged();
    renderSkills();
    refreshDerived();
  }));
  document.getElementById("nonstandard-override").addEventListener("change", (event) => { state.nonstandard_override = event.target.checked; markChanged(); });

  document.querySelectorAll("[data-occupation-mode]").forEach((button) => button.addEventListener("click", () => {
    state.occupation_mode = button.dataset.occupationMode;
    state.group_choices = {};
    state.free_skill_choices = [];
    renderOccupation();
    syncOccupationSkills();
    renderSkills();
    markChanged();
    refreshDerived();
  }));
  document.getElementById("occupation-search").addEventListener("input", renderOccupationOptions);
  document.getElementById("occupation-select").addEventListener("change", (event) => {
    state.occupation_id = Number(event.target.value);
    state.group_choices = {};
    state.free_skill_choices = [];
    renderOccupation();
    syncOccupationSkills();
    renderSkills();
    markChanged();
    refreshDerived();
  });
  bindCustomOccupationControls();
  document.getElementById("free-choice-search").addEventListener("input", renderFreeChoices);
  document.getElementById("custom-skill-search").addEventListener("input", renderCustomSkillChoices);

  document.getElementById("skill-search").addEventListener("input", renderSkills);
  document.getElementById("skill-filter").addEventListener("change", renderSkills);
  document.getElementById("skill-table-body").addEventListener("input", handleSkillInput);
  document.getElementById("skill-table-body").addEventListener("click", (event) => {
    const button = event.target.closest("[data-clear-occupation]");
    if (!button) return;
    const skill = skillBySlot(button.dataset.clearOccupation);
    if (!skill || skill.selected_occupation) return;
    skill.occupation_points = 0;
    renderSkills(); renderAssetReference(); markChanged();
  });

  Object.entries(BACKGROUND_FIELDS).forEach(([id, key]) => document.getElementById(id).addEventListener("input", (event) => { state.background[key] = event.target.value; markChanged(); }));
  Object.entries(ASSET_FIELDS).forEach(([id, key]) => document.getElementById(id).addEventListener("input", (event) => { state.assets[key] = event.target.value; renderAssetReference(); markChanged(); }));
  document.getElementById("identity-story-year").addEventListener("change", (event) => {
    if (!event.target.value || !Number.isInteger(Number(event.target.value))) {
      event.target.value = state.identity.story_year;
      toast("故事年份请输入完整的整数年份。", true); return;
    }
    state.identity.story_year = Number(event.target.value);
    state.assets.exchange_year = state.identity.story_year;
    renderCurrencyControls(); renderAssetReference(); renderLiveSummary(); markChanged();
  });
  document.getElementById("asset-currency").addEventListener("change", (event) => {
    state.assets.currency_code = event.target.value;
    renderAssetReference(); markChanged();
  });
  document.getElementById("apply-asset-reference").addEventListener("click", applyAssetReference);

  document.getElementById("add-weapon").addEventListener("click", addWeapon);
  document.getElementById("add-item").addEventListener("click", addInventoryItem);
  document.getElementById("weapon-table-body").addEventListener("input", handleWeaponInput);
  document.getElementById("weapon-table-body").addEventListener("change", handleEquipmentSelection);
  document.getElementById("weapon-table-body").addEventListener("click", handleEquipmentRemove);
  document.getElementById("inventory-table-body").addEventListener("input", handleInventoryInput);
  document.getElementById("experience-selection").addEventListener("change", (event) => {
    state.experience.selection = event.target.value;
    state.experience.san_loss = 0;
    renderExperience(); renderSkills(); refreshDerived(); markChanged();
  });
  for (const [id, field] of [["experience-san", "san_loss"], ["experience-budget", "skill_points"], ["experience-name", "name"], ["experience-notes", "notes"]]) {
    document.getElementById(id).addEventListener("input", (event) => {
      state.experience[field] = ["san_loss", "skill_points"].includes(field) ? numeric(event.target.value) : event.target.value;
      renderBudgets(); if (field === "san_loss") refreshDerived(); markChanged();
    });
  }
  document.getElementById("inventory-table-body").addEventListener("change", handleInventoryInput);
  document.getElementById("inventory-table-body").addEventListener("click", handleEquipmentRemove);

  document.getElementById("validate-now").addEventListener("click", () => validateDraft(true));
  document.getElementById("export-pdf").addEventListener("click", exportPdf);
  document.getElementById("export-excel").addEventListener("click", exportExcel);
  document.getElementById("issue-list").addEventListener("click", handleIssueNavigation);

  bindPortraitControls();
  bindImportControls();
  bindAttributeImportControls();
  const dialog = document.getElementById("reset-dialog");
  document.getElementById("reset-open").addEventListener("click", () => dialog.showModal());
  document.getElementById("reset-confirm").addEventListener("click", () => { clearAttributeImport(); sessionStorage.removeItem(STORAGE_KEY); location.reload(); });
}

function bindCustomOccupationControls() {
  const fields = {
    "custom-name": "name", "custom-credit-min": "credit_min", "custom-credit-max": "credit_max",
    "custom-formula-kind": "point_formula_kind", "custom-secondary": "secondary",
  };
  Object.entries(fields).forEach(([id, key]) => document.getElementById(id).addEventListener("input", (event) => {
    state.custom_occupation[key] = ["credit_min", "credit_max"].includes(key) ? numeric(event.target.value, 0) : event.target.value;
    document.getElementById("custom-secondary-wrap").hidden = state.custom_occupation.point_formula_kind !== "secondary";
    renderLiveSummary();
    markChanged();
    refreshDerived();
  }));
}

function hydrateControls() {
  document.getElementById("identity-story-year").value = state.identity.story_year;
  const identityBindings = {
    "identity-name": "name", "identity-player": "player", "identity-age": "age", "identity-gender": "gender",
    "identity-residence": "residence", "identity-birthplace": "birthplace", "identity-era": "era", "identity-date": "current_date",
  };
  Object.entries(identityBindings).forEach(([id, key]) => { document.getElementById(id).value = state.identity[key] ?? ""; });
  ATTRIBUTE_KEYS.forEach((key) => { document.getElementById(`attr-${key}`).value = state.attributes[key]; });
  document.getElementById("nonstandard-override").checked = Boolean(state.nonstandard_override);
  document.getElementById("custom-name").value = state.custom_occupation.name;
  document.getElementById("custom-credit-min").value = state.custom_occupation.credit_min;
  document.getElementById("custom-credit-max").value = state.custom_occupation.credit_max;
  document.getElementById("custom-formula-kind").value = state.custom_occupation.point_formula_kind;
  document.getElementById("custom-secondary").value = state.custom_occupation.secondary;
  document.getElementById("custom-secondary-wrap").hidden = state.custom_occupation.point_formula_kind !== "secondary";
  Object.entries(BACKGROUND_FIELDS).forEach(([id, key]) => { document.getElementById(id).value = state.background[key] || ""; });
  const connections = String(state.background.key_connection || "").split("、");
  const connectionGrid = document.getElementById("bg-key");
  connectionGrid.innerHTML = "";
  for (const name of ["形象描述", "思想与信念", "重要之人", "意义非凡之地", "宝贵之物", "特质", "伤口和疤痕", "恐惧症和躁狂症"]) {
    const chip = createCheckChip(name, connections.includes(name));
    chip.querySelector("input").value = name;
    chip.querySelector("input").addEventListener("change", () => {
      state.background.key_connection = [...connectionGrid.querySelectorAll("input:checked")].map((input) => input.value).join("、");
      markChanged();
    });
    connectionGrid.appendChild(chip);
  }
  Object.entries(ASSET_FIELDS).forEach(([id, key]) => { document.getElementById(id).value = state.assets[key] || ""; });
  renderCurrencyControls();
}

function renderCatalogMeta() {
  const meta = bootstrapData.meta;
  document.getElementById("catalog-meta").textContent = `${meta.occupation_count} 个职业 · ${meta.skill_count} 个技能`;
}

function goToStep(step, persist = true) {
  currentStep = clampNumber(step, 0, 6, 0);
  state.current_step = currentStep;
  document.querySelectorAll(".wizard-step").forEach((section) => section.classList.toggle("active", Number(section.dataset.step) === currentStep));
  document.querySelectorAll(".step-link").forEach((button) => button.classList.toggle("active", Number(button.dataset.stepTarget) === currentStep));
  const fraction = ((currentStep + 1) / 7) * 100;
  const degrees = fraction * 3.6;
  document.getElementById("progress-emblem").style.setProperty("--progress", `${degrees}deg`);
  document.getElementById("progress-percent").textContent = `${Math.round(fraction)}%`;
  document.getElementById("step-caption").textContent = `步骤 ${currentStep + 1} / 7 · ${STEP_LABELS[currentStep]}`;
  document.getElementById("footer-progress").style.width = `${fraction}%`;
  document.getElementById("previous-step").disabled = currentStep === 0;
  document.getElementById("next-step").hidden = currentStep === 6;
  if (persist) scheduleSave();
  if (currentStep === 4) renderAssetReference();
  if (currentStep === 5) renderEquipment();
  if (currentStep === 6) validateDraft(false);
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function currentOccupation() {
  if (state.occupation_mode === "custom") {
    const low = Math.min(numeric(state.custom_occupation.credit_min), numeric(state.custom_occupation.credit_max));
    const high = Math.max(numeric(state.custom_occupation.credit_min), numeric(state.custom_occupation.credit_max));
    const formula = state.custom_occupation.point_formula_kind === "secondary"
      ? `EDU*2+${state.custom_occupation.secondary}*2`
      : "EDU*4";
    return {
      occupation_id: 1,
      name: state.custom_occupation.name.trim() || "自定义职业",
      credit_min: low,
      credit_max: high,
      point_formula: formula,
      summary: "由玩家与 KP 共同定义的职业。",
      contacts: "",
      description: "",
      fixed_skills: state.custom_skill_slots.map((slot) => skillBySlot(slot)?.name).filter(Boolean),
      choice_groups: [],
      free_choices: 0,
      is_custom: true,
    };
  }
  return bootstrapData.occupations.find((item) => item.occupation_id === Number(state.occupation_id)) || null;
}

function renderOccupationOptions() {
  const select = document.getElementById("occupation-select");
  const query = normalizeText(document.getElementById("occupation-search").value);
  const matches = bootstrapData.occupations.filter((item) => !query || normalizeText(`${item.name} ${item.summary} ${item.description}`).includes(query));
  select.innerHTML = matches.map((item) => `<option value="${item.occupation_id}">${escapeHtml(item.name)} · 信用 ${item.credit_min}-${item.credit_max}</option>`).join("");
  if (matches.some((item) => item.occupation_id === Number(state.occupation_id))) select.value = String(state.occupation_id);
  else if (matches.length) {
    state.occupation_id = matches[0].occupation_id;
    select.value = String(state.occupation_id);
  }
}

function renderOccupation() {
  renderExperience();
  document.querySelectorAll("[data-occupation-mode]").forEach((button) => button.classList.toggle("active", button.dataset.occupationMode === state.occupation_mode));
  document.getElementById("catalog-occupation-panel").hidden = state.occupation_mode !== "catalog";
  document.getElementById("custom-occupation-panel").hidden = state.occupation_mode !== "custom";
  if (state.occupation_mode === "custom") {
    renderCustomSkillChoices();
    return;
  }
  renderOccupationOptions();
  const occupation = currentOccupation();
  if (!occupation) return;
  document.getElementById("occupation-select").value = String(occupation.occupation_id);
  document.getElementById("occupation-title").textContent = occupation.name;
  document.getElementById("occupation-summary").textContent = occupation.summary || "暂无资料";
  document.getElementById("occupation-id").textContent = `#${occupation.occupation_id}`;
  document.getElementById("occupation-credit").textContent = `${occupation.credit_min}–${occupation.credit_max}`;
  document.getElementById("occupation-formula").textContent = formulaLabel(occupation.point_formula);
  document.getElementById("occupation-points").textContent = state.derived?.occupation_points ?? "—";
  document.getElementById("occupation-free-count").textContent = occupation.free_choices;
  document.getElementById("fixed-skill-chips").innerHTML = occupation.fixed_skills.length
    ? occupation.fixed_skills.map((skill) => `<span>${escapeHtml(skill)}</span>`).join("")
    : "<span>无固定技能</span>";
  document.getElementById("occupation-contacts").textContent = occupation.contacts || "暂无资料";
  document.getElementById("occupation-description").textContent = occupation.description || "暂无资料";
  renderOccupationChoiceGroups(occupation);
  renderFreeChoices();
}

function renderOccupationChoiceGroups(occupation) {
  const host = document.getElementById("occupation-choice-groups");
  host.innerHTML = "";
  occupation.choice_groups.forEach((group) => {
    const current = (state.group_choices[group.marker] || []).filter((item) => group.candidates.includes(item));
    state.group_choices[group.marker] = current.slice(0, group.required_count);
    const card = document.createElement("div");
    card.className = "paper-card choice-card";
    card.innerHTML = `<div class="choice-heading"><div><span class="choice-marker">${escapeHtml(group.marker)}</span><h3>${escapeHtml(group.label)}</h3></div><span class="choice-counter">${state.group_choices[group.marker].length} / ${group.required_count}</span></div><div class="checkbox-chip-grid"></div>`;
    const grid = card.querySelector(".checkbox-chip-grid");
    group.candidates.forEach((candidate) => {
      const label = createCheckChip(candidate, state.group_choices[group.marker].includes(candidate));
      const input = label.querySelector("input");
      input.addEventListener("change", () => {
        const values = state.group_choices[group.marker] || [];
        if (input.checked && !values.includes(candidate)) {
          if (values.length >= group.required_count) { input.checked = false; toast(`“${group.label}”最多选择 ${group.required_count} 项。`, true); return; }
          values.push(candidate);
        } else if (!input.checked) state.group_choices[group.marker] = values.filter((item) => item !== candidate);
        card.querySelector(".choice-counter").textContent = `${state.group_choices[group.marker].length} / ${group.required_count}`;
        syncOccupationSkills();
        renderSkills();
        markChanged();
      });
      grid.appendChild(label);
    });
    host.appendChild(card);
  });
}

function renderFreeChoices() {
  const occupation = currentOccupation();
  const card = document.getElementById("free-choice-card");
  if (!occupation || state.occupation_mode !== "catalog" || occupation.free_choices === 0) { card.hidden = true; state.free_skill_choices = []; return; }
  card.hidden = false;
  state.free_skill_choices = state.free_skill_choices.filter((slot) => bootstrapData.skills.some((item) => item.template_slot === slot)).slice(0, occupation.free_choices);
  document.getElementById("free-choice-caption").textContent = "";
  document.getElementById("free-choice-counter").textContent = `${state.free_skill_choices.length} / ${occupation.free_choices}`;
  const query = normalizeText(document.getElementById("free-choice-search").value);
  const grid = document.getElementById("free-choice-grid");
  grid.innerHTML = "";
  bootstrapData.skills.filter((skill) => !query || normalizeText(`${skill.name} ${skill.default_specialization}`).includes(query)).forEach((skill) => {
    const labelText = skill.default_specialization ? `${skill.name}（${skill.default_specialization}）` : skill.name;
    const label = createCheckChip(labelText, state.free_skill_choices.includes(skill.template_slot));
    const input = label.querySelector("input");
    input.addEventListener("change", () => {
      if (input.checked && !state.free_skill_choices.includes(skill.template_slot)) {
        if (state.free_skill_choices.length >= occupation.free_choices) { input.checked = false; toast(`任意特长最多选择 ${occupation.free_choices} 项。`, true); return; }
        state.free_skill_choices.push(skill.template_slot);
      } else if (!input.checked) state.free_skill_choices = state.free_skill_choices.filter((slot) => slot !== skill.template_slot);
      document.getElementById("free-choice-counter").textContent = `${state.free_skill_choices.length} / ${occupation.free_choices}`;
      syncOccupationSkills(); renderSkills(); markChanged();
    });
    grid.appendChild(label);
  });
}

function renderCustomSkillChoices() {
  state.custom_skill_slots = state.custom_skill_slots.filter((slot) => bootstrapData.skills.some((item) => item.template_slot === slot)).slice(0, 8);
  document.getElementById("custom-skill-counter").textContent = `${state.custom_skill_slots.length} / 8`;
  const query = normalizeText(document.getElementById("custom-skill-search").value);
  const grid = document.getElementById("custom-skill-grid");
  grid.innerHTML = "";
  bootstrapData.skills.filter((skill) => !query || normalizeText(`${skill.name} ${skill.default_specialization}`).includes(query)).forEach((skill) => {
    const labelText = skill.default_specialization ? `${skill.name}（${skill.default_specialization}）` : skill.name;
    const label = createCheckChip(labelText, state.custom_skill_slots.includes(skill.template_slot));
    const input = label.querySelector("input");
    input.addEventListener("change", () => {
      if (input.checked && !state.custom_skill_slots.includes(skill.template_slot)) {
        if (state.custom_skill_slots.length >= 8) { input.checked = false; toast("自定义职业最多选择八个职业技能。", true); return; }
        state.custom_skill_slots.push(skill.template_slot);
      } else if (!input.checked) state.custom_skill_slots = state.custom_skill_slots.filter((slot) => slot !== skill.template_slot);
      document.getElementById("custom-skill-counter").textContent = `${state.custom_skill_slots.length} / 8`;
      syncOccupationSkills(); renderSkills(); renderLiveSummary(); markChanged();
    });
    grid.appendChild(label);
  });
}

function createCheckChip(text, checked) {
  const label = document.createElement("label");
  label.className = "check-chip";
  const input = document.createElement("input");
  input.type = "checkbox";
  input.checked = checked;
  const span = document.createElement("span");
  span.textContent = text;
  label.append(input, span);
  return label;
}

function selectedOccupationSlots() {
  const selected = new Map([["F26", ""]]);
  const occupation = currentOccupation();
  if (!occupation) return selected;
  if (occupation.is_custom) {
    state.custom_skill_slots.forEach((slot) => selected.set(slot, ""));
    return selected;
  }
  const tokens = [...occupation.fixed_skills];
  occupation.choice_groups.forEach((group) => tokens.push(...(state.group_choices[group.marker] || [])));
  tokens.forEach((token) => {
    const [name, specialization] = splitSkillToken(token);
    const definition = bootstrapData.skills.find((item) => normalizeSkillName(item.name) === name);
    if (definition) selected.set(definition.template_slot, specialization);
  });
  state.free_skill_choices.forEach((slot) => selected.set(slot, ""));
  return selected;
}

function syncOccupationSkills() {
  const selected = selectedOccupationSlots();
  state.skills.forEach((skill) => {
    skill.selected_occupation = selected.has(skill.template_slot);
    const specialization = selected.get(skill.template_slot);
    if (skill.selected_occupation && specialization && !String(skill.specialization || "").trim()) skill.specialization = specialization;
  });
  renderBudgets();
  renderLiveSummary();
}

function syncDynamicSkillBases() {
  const definitions = new Map(bootstrapData.skills.map((item) => [item.template_slot, item]));
  state.skills.forEach((skill) => {
    const definition = definitions.get(skill.template_slot);
    if (definition) skill.base_value = Number.isInteger(skill.base_override) && skill.base_override >= 0
      ? skill.base_override
      : dynamicBase(definition, state.attributes);
  });
}

function dynamicBase(definition, attributes) {
  if (!definition.base_formula) return numeric(definition.base_value, 0);
  if (definition.base_formula.toUpperCase() === "DEX/2") return Math.floor(numeric(attributes.DEX) / 2);
  if (definition.base_formula.toUpperCase() === "EDU") return numeric(attributes.EDU);
  return numeric(definition.base_value, 0);
}

function renderSkills() {
  const body = document.getElementById("skill-table-body");
  const query = normalizeText(document.getElementById("skill-search").value);
  const filter = document.getElementById("skill-filter").value;
  const visible = state.skills.filter((skill) => {
    const matchesQuery = !query || normalizeText(`${skill.name} ${skill.specialization}`).includes(query);
    const allocated = numeric(skill.occupation_points) + numeric(skill.interest_points) + numeric(skill.extra_final) + numeric(skill.experience_points) > 0;
    const matchesFilter = filter === "all" || (filter === "occupation" && skill.selected_occupation) || (filter === "allocated" && allocated) || (filter === "specialized" && (skill.specialization || skill.name.endsWith("：") || skill.name.endsWith(":")));
    return matchesQuery && matchesFilter;
  });
  body.innerHTML = visible.map((skill) => skillRowHtml(skill)).join("");
  document.getElementById("skill-empty").hidden = visible.length > 0;
  renderBudgets();
}

function renderExperience() {
  const value = state.experience;
  const definitions = bootstrapData.experience_packages || [];
  const definition = definitions.find((item) => item.name === value.selection);
  document.getElementById("experience-selection").innerHTML = '<option value="">无</option>' + definitions.map((item) => `<option value="${escapeAttr(item.name)}">${escapeHtml(item.name)}</option>`).join("") + '<option value="custom">自定义经历包</option>';
  document.getElementById("experience-selection").value = value.selection;
  const custom = value.selection === "custom";
  document.getElementById("experience-name-wrap").hidden = !custom;
  document.getElementById("experience-notes-wrap").hidden = !custom;
  document.getElementById("experience-name").value = value.name;
  document.getElementById("experience-notes").value = value.notes;
  document.getElementById("experience-san").value = value.san_loss;
  document.getElementById("experience-san").disabled = !value.selection;
  document.getElementById("experience-budget").readOnly = !custom;
  document.getElementById("experience-budget").value = experienceBudget();
  document.getElementById("experience-details").hidden = !definition;
  setText("experience-description", definition ? `SAN 减少：${definition.san_description}。${definition.notes}` : "");
}

function experienceBudget() {
  return state.experience.selection === "custom" ? numeric(state.experience.skill_points) : (bootstrapData.experience_packages || []).find((item) => item.name === state.experience.selection)?.skill_points || 0;
}

function skillRowHtml(skill) {
  const finalValue = skillFinal(skill);
  const isSpecialized = Boolean(String(skill.specialization || "").trim()) || /[：:]$/.test(skill.name);
  const warning = finalValue > 99 || (skill.name.includes("克苏鲁神话") && (numeric(skill.occupation_points) || numeric(skill.interest_points)));
  const badges = `${skill.selected_occupation ? '<span class="skill-badge occupation">★</span>' : '<span class="skill-badge">—</span>'}${isSpecialized ? '<span class="skill-badge specialized">专</span>' : ""}`;
  return `<tr data-skill-row="${escapeAttr(skill.template_slot)}" class="${warning ? "warning" : ""}">
    <td>${badges}</td>
    <td class="skill-name">${escapeHtml(skill.name.replace(/[：:]$/, ""))}${skill.specialization ? `（${escapeHtml(skill.specialization)}）` : ""}</td>
    <td class="readonly-value" data-result="base">${numeric(skill.base_value)}</td>
    <td><input type="number" min="0" max="500" data-skill-slot="${escapeAttr(skill.template_slot)}" data-skill-field="occupation_points" value="${numeric(skill.occupation_points)}" aria-label="${escapeAttr(skill.name)}职业点" ${skill.selected_occupation ? "" : 'disabled title="职业点仅可用于职业技能"'}>${!skill.selected_occupation && numeric(skill.occupation_points) ? `<button type="button" data-clear-occupation="${escapeAttr(skill.template_slot)}">退回职业点</button>` : ""}</td>
    <td><input type="number" min="0" max="500" data-skill-slot="${escapeAttr(skill.template_slot)}" data-skill-field="interest_points" value="${numeric(skill.interest_points)}" aria-label="${escapeAttr(skill.name)}兴趣点"></td>
    <td><input type="number" min="0" max="500" data-skill-slot="${escapeAttr(skill.template_slot)}" data-skill-field="experience_points" value="${numeric(skill.experience_points)}" aria-label="${escapeAttr(skill.name)}经历包点"></td>
    <td><input type="number" min="0" max="500" data-skill-slot="${escapeAttr(skill.template_slot)}" data-skill-field="extra_final" value="${numeric(skill.extra_final)}" aria-label="${escapeAttr(skill.name)}成长点数"></td>
    <td class="readonly-value" data-result="final">${finalValue}</td><td class="readonly-value" data-result="hard">${Math.floor(finalValue / 2)}</td><td class="readonly-value" data-result="extreme">${Math.floor(finalValue / 5)}</td>
  </tr>`;
}

function handleSkillInput(event) {
  const input = event.target.closest("[data-skill-slot]");
  if (!input) return;
  const skill = skillBySlot(input.dataset.skillSlot);
  if (!skill) return;
  const field = input.dataset.skillField;
  if (!["occupation_points", "interest_points", "extra_final", "experience_points"].includes(field)) return;
  if (field === "occupation_points" && !skill.selected_occupation) return;
  skill[field] = Math.max(0, numeric(input.value, 0));
  const row = input.closest("tr");
  const finalValue = skillFinal(skill);
  row.querySelector('[data-result="final"]').textContent = finalValue;
  row.querySelector('[data-result="hard"]').textContent = Math.floor(finalValue / 2);
  row.querySelector('[data-result="extreme"]').textContent = Math.floor(finalValue / 5);
  row.classList.toggle("warning", finalValue > 99 || (skill.name.includes("克苏鲁神话") && (numeric(skill.occupation_points) || numeric(skill.interest_points))));
  renderBudgets();
  renderAssetReference();
  markChanged();
}

function skillBySlot(slot) { return state.skills.find((item) => item.template_slot === slot); }
function skillFinal(skill) { return numeric(skill.base_value) + numeric(skill.occupation_points) + numeric(skill.interest_points) + numeric(skill.extra_final) + numeric(skill.experience_points); }

function renderBudgets() {
  const experienceUsed = state.skills.reduce((total, skill) => total + numeric(skill.experience_points), 0);
  document.getElementById("experience-remaining").hidden = !state.experience.selection;
  setText("experience-remaining", `经历包点 ${experienceUsed} / ${experienceBudget()} · ${budgetMessage(experienceUsed, experienceBudget())}`);
  const occUsed = state.skills.reduce((total, skill) => total + numeric(skill.occupation_points), 0);
  const intUsed = state.skills.reduce((total, skill) => total + numeric(skill.interest_points), 0);
  const occTotal = numeric(state.derived?.occupation_points, 0);
  const intTotal = numeric(state.derived?.interest_points, numeric(state.attributes.INT) * 2);
  setText("occ-used", occUsed); setText("occ-total", occTotal); setText("int-used", intUsed); setText("int-total", intTotal);
  setText("occ-remaining", budgetMessage(occUsed, occTotal)); setText("int-remaining", budgetMessage(intUsed, intTotal));
  setMeter("occ-meter", occUsed, occTotal); setMeter("int-meter", intUsed, intTotal);
  document.querySelector(".occupation-budget")?.classList.toggle("over", occUsed > occTotal);
  document.querySelector(".interest-budget")?.classList.toggle("over", intUsed > intTotal);
  setText("mini-occ-budget", `${occUsed} / ${occTotal}`); setText("mini-int-budget", `${intUsed} / ${intTotal}`);
  setMeter("mini-occ-meter", occUsed, occTotal); setMeter("mini-int-meter", intUsed, intTotal);
  setText("mini-budget-state", occUsed > occTotal || intUsed > intTotal ? "超出预算" : "预算内");
}

function budgetMessage(used, total) {
  const difference = total - used;
  if (difference === 0) return "预算已恰好用完";
  return difference > 0 ? `剩余 ${difference} 点` : `超出 ${Math.abs(difference)} 点`;
}

function setMeter(id, used, total) {
  const percent = total > 0 ? Math.min(100, (used / total) * 100) : 0;
  document.getElementById(id).style.width = `${percent}%`;
}

async function refreshDerived(immediate = false) {
  clearTimeout(derivedTimer);
  const generation = draftGeneration;
  const run = async () => {
    try {
      const response = await apiFetch("/api/calculate", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ attributes: state.attributes, age: state.identity.age, occupation_formula: currentOccupation()?.point_formula || "EDU*4", san_loss: state.experience.selection ? numeric(state.experience.san_loss) : 0 }),
      }, { timeoutMs: 15000 });
      const data = await readJsonResponse(response, "派生值计算失败。");
      if (generation !== draftGeneration) return;
      if (!response.ok) throw new Error(data.detail || "派生值计算失败。");
      state.derived = data.derived;
      renderDerived(); renderBudgets(); setText("occupation-points", state.derived.occupation_points); renderLiveSummary(); scheduleSave();
    } catch (error) { if (generation === draftGeneration) toast(friendlyErrorMessage(error, "派生值计算失败，请稍后重试。"), true); }
  };
  if (immediate) await run(); else derivedTimer = setTimeout(run, 220);
}

function renderDerived() {
  if (!state.derived) return;
  const mapping = { hp: "derived-hp", mp: "derived-mp", san: "derived-san", mov: "derived-mov", dodge: "derived-dodge", damage_bonus: "derived-damage", build: "derived-build", interest_points: "derived-interest" };
  Object.entries(mapping).forEach(([key, id]) => setText(id, state.derived[key]));
}

function renderLiveSummary() {
  const occupation = currentOccupation();
  setText("mini-name", state.identity.name.trim() || "未命名调查员");
  setText("mini-occupation", occupation?.name || "尚未选择职业");
  setText("mini-era", `${state.identity.era || "时代未定"} · ${numeric(state.identity.age, 30)} 岁`);
  setText("mini-hp", state.derived?.hp ?? "—"); setText("mini-mp", state.derived?.mp ?? "—"); setText("mini-san", state.derived?.san ?? "—"); setText("mini-mov", state.derived?.mov ?? "—");
  renderBudgets();
}

function renderAssetReference() {
  const creditSkill = state.skills.find((skill) => skill.name.replace(/[：:]$/, "") === "信用评级");
  const credit = creditSkill ? skillFinal(creditSkill) : 0;
  const reference = assetReference(credit);
  const quote = currentCurrency();
  setText("asset-credit", credit); setText("asset-ref-living", reference.living);
  renderCurrencySource(quote);
  if (!quote) {
    for (const id of ["asset-ref-spending", "asset-ref-cash", "asset-ref-assets", "asset-converted-spending", "asset-converted-cash", "asset-converted-assets"]) setText(id, "暂无可用汇率");
    const validYear = Boolean(bootstrapData.currencies[String(state.assets.exchange_year)]);
    setText("asset-fx-note", validYear ? "原币种在该年份没有可用报价，请重新选择币种。美元输入已保留。" : "汇率年份支持 1920—2026 年，请修改故事年份后导出。");
    document.getElementById("export-pdf").disabled = true;
    document.getElementById("export-excel").disabled = true;
    return;
  }
  const factor = state.assets.exchange_year === 2026 ? 20 : 1;
  const converted = (value) => {
    try { return `${Coc7Currency.formatAmount(value, quote.rate)} ${quote.name}`; }
    catch (error) { return error.message; }
  };
  for (const [key, ref, id] of [["usd_spending", "spending", "spending"], ["usd_cash", "cash", "cash"], ["usd_assets", "assets", "assets"]]) {
    const baseline = Number(String(reference[ref]).replace(/[$,+]/g, "")) * factor;
    setText(`asset-ref-${id}`, converted(baseline));
    const input = state.assets[key];
    setText(`asset-converted-${id}`, converted(input != null && String(input).trim() ? input : baseline));
  }
  state.assets.currency = quote.name;
  setText("asset-fx-note", `1 USD = ${quote.rate} ${quote.name}（${state.assets.exchange_year} 年）`);
}

function currentCurrency() {
  return (bootstrapData.currencies[String(state.assets.exchange_year)] || []).find((item) => item.code === state.assets.currency_code);
}

function renderCurrencySource(quote) {
  const missing = bootstrapData.currency_missing[String(state.assets.exchange_year)] || [];
  const selected = quote || missing.find((item) => item.code === state.assets.currency_code);
  setText("asset-source-summary", selected ? `数据来源：${selected.source} · ${selected.basis} · ${selected.date}` : "选择币种后查看来源");
  const warning = document.getElementById("asset-fx-warning");
  warning.hidden = !selected?.warning && !selected?.missing_reason;
  warning.textContent = selected?.warning || selected?.missing_reason || "";
  document.getElementById("asset-source-details").hidden = !selected;
  const content = document.getElementById("asset-source-content");
  content.replaceChildren();
  if (selected) {
    const add = (label, value) => {
      const dt = document.createElement("dt"), dd = document.createElement("dd");
      dt.textContent = label; dd.textContent = value || "来源表未提供";
      content.append(dt, dd); return dd;
    };
    add("机构或文献", selected.source);
    add("统计口径与期间", `${selected.basis}；${selected.date}`);
    const linkCell = add("原始资料链接", selected.source_url ? "" : (selected.code === "USD" ? "美元基准恒等，无外部报价。" : "来源表未提供原始链接。"));
    if (/^https?:\/\//i.test(selected.source_url)) {
      const link = document.createElement("a");
      link.href = selected.source_url; link.target = "_blank"; link.rel = "noopener noreferrer";
      link.textContent = selected.source; linkCell.replaceChildren(link);
    }
    add("引用数据文件", bootstrapData.currency_metadata.source_file);
    add("表内位置", selected.source_cells.join("、") + (selected.raw_cell ? `；${selected.raw_cell}` : ""));
    if (selected.raw_quote) add("原始报价", `${selected.raw_quote} ${selected.raw_direction === "USD/LCU" ? "美元 / 每单位本币" : "本币 / 每美元"}`);
    if (selected.conversion) add("换算说明", selected.conversion);
    if (selected.source_grade) add("原表来源评级", selected.source_grade);
    add("说明与限制", selected.missing_reason || selected.notes || "来源表未列出额外限制。");
  }
  document.getElementById("asset-missing-details").hidden = !missing.length;
  setText("asset-missing-summary", `缺失报价（${missing.length} 种）`);
  const list = document.getElementById("asset-missing-list");
  list.replaceChildren();
  for (const item of missing) {
    const li = document.createElement("li");
    li.textContent = `${item.name}：${item.missing_reason}`; list.append(li);
  }
}

function renderCurrencyControls() {
  const options = bootstrapData.currencies[String(state.assets.exchange_year)] || [];
  const missing = bootstrapData.currency_missing[String(state.assets.exchange_year)] || [];
  const valid = options.some((item) => item.code === state.assets.currency_code);
  const select = document.getElementById("asset-currency");
  select.disabled = options.length === 0;
  document.getElementById("asset-year").value = state.assets.exchange_year;
  select.innerHTML = (valid ? "" : '<option value="" selected>请选择该年份的币种</option>')
    + options.map((item) => `<option value="${escapeAttr(item.code)}">${escapeHtml(item.name)}</option>`).join("")
    + (missing.length ? '<optgroup label="缺少报价，不可选">' + missing.map((item) => `<option disabled value="${escapeAttr(item.code)}">${escapeHtml(item.name)}（缺失）</option>`).join("") + '</optgroup>' : "");
  select.value = valid ? state.assets.currency_code : "";
}

function assetReference(value) {
  const credit = Math.max(0, Math.min(99, numeric(value)));
  if (credit === 0) return { living: "身无分文", spending: "$0.50", cash: "$0.50", assets: "$0" };
  if (credit <= 9) return { living: "贫穷", spending: "$2", cash: `$${credit}`, assets: `$${credit * 10}` };
  if (credit <= 49) return { living: "标准", spending: "$10", cash: `$${credit * 2}`, assets: `$${credit * 50}` };
  if (credit <= 89) return { living: "小康", spending: "$50", cash: `$${credit * 5}`, assets: `$${credit * 500}` };
  if (credit <= 98) return { living: "富裕", spending: "$250", cash: `$${credit * 20}`, assets: `$${credit * 2000}` };
  return { living: "豪富", spending: "$5,000", cash: "$50,000", assets: "$5,000,000+" };
}

function applyAssetReference() {
  const creditSkill = state.skills.find((skill) => skill.name.replace(/[：:]$/, "") === "信用评级");
  const reference = assetReference(creditSkill ? skillFinal(creditSkill) : 0);
  state.assets.living_standard = reference.living;
  state.assets.usd_spending = ""; state.assets.usd_cash = ""; state.assets.usd_assets = "";
  hydrateAssetControls(); renderAssetReference(); markChanged(); toast("已恢复自动计算资产。", false);
}

function hydrateAssetControls() { Object.entries(ASSET_FIELDS).forEach(([id, key]) => { document.getElementById(id).value = state.assets[key] || ""; }); }

function addWeapon() {
  if (state.weapons.length >= 6) { toast("最多只能添加六项常用武器。", true); return; }
  state.weapons.push({ name: "", category: "", skill: "", damage: "", range: "", attacks: "", ammo: "", malfunction: "", notes: "" });
  renderEquipment(); markChanged();
}

function addInventoryItem() {
  if (state.inventory.length >= 15) { toast("最多只能添加十五项随身物品。", true); return; }
  state.inventory.push({ name: "", status: "", location: "", backpack_slot: "" });
  renderEquipment(); markChanged();
}

function renderEquipment() {
  document.getElementById("weapon-table-body").innerHTML = state.weapons.map((weapon, index) => `<tr data-equipment-kind="weapon" data-equipment-index="${index}"><td>${index + 1}</td><td><input data-field="name" value="${escapeAttr(weapon.name)}" placeholder="武器名称" aria-label="第 ${index + 1} 项武器名称"></td><td>${equipmentChoiceHtml("weapon", weapon, index)}</td><td><select data-field="skill" aria-label="第 ${index + 1} 项武器技能">${weaponSkillOptions(weapon.skill)}</select></td><td><input data-field="damage" value="${escapeAttr(weapon.damage)}" placeholder="1D8" aria-label="第 ${index + 1} 项武器伤害"></td><td><input data-field="range" value="${escapeAttr(weapon.range)}"></td><td><input data-field="attacks" value="${escapeAttr(weapon.attacks)}"></td><td><input data-field="ammo" value="${escapeAttr(weapon.ammo)}"></td><td><input data-field="malfunction" value="${escapeAttr(weapon.malfunction)}"></td><td><button type="button" class="remove-row" data-remove-row="weapon" aria-label="移除武器">×</button></td></tr>`).join("");
  document.getElementById("inventory-table-body").innerHTML = state.inventory.map((item, index) => `<tr data-equipment-kind="inventory" data-equipment-index="${index}"><td>${index + 1}</td><td><input data-field="name" value="${escapeAttr(item.name)}" placeholder="物品名称" aria-label="第 ${index + 1} 项物品名称"></td><td><input data-field="status" value="${escapeAttr(item.status)}" aria-label="第 ${index + 1} 项物品状态"></td><td><input data-field="location" list="inventory-locations" value="${escapeAttr(item.location)}" placeholder="选择或填写部位" aria-label="第 ${index + 1} 项物品携带部位"></td><td><input data-field="backpack_slot" value="${escapeAttr(item.backpack_slot)}"></td><td><button type="button" class="remove-row" data-remove-row="inventory" aria-label="移除物品">×</button></td></tr>`).join("");
  document.getElementById("weapon-empty").hidden = state.weapons.length > 0;
  document.getElementById("inventory-empty").hidden = state.inventory.length > 0;
  document.getElementById("add-weapon").disabled = state.weapons.length >= 6;
  document.getElementById("add-item").disabled = state.inventory.length >= 15;
}

function equipmentCatalog(kind) {
  return (kind === "weapon" ? bootstrapData.weapons : bootstrapData.inventory) || [];
}

function equipmentChoiceHtml(kind, item, index) {
  const catalog = equipmentCatalog(kind);
  const isWeapon = kind === "weapon";
  const field = isWeapon ? "category" : "name";
  const selectedId = item.catalog_id || catalog.find((entry) => entry.name === item[field])?.catalog_id || (item[field] ? "custom" : "");
  const selected = catalog.find((entry) => entry.catalog_id === selectedId);
  const custom = selectedId === "custom" || Boolean(selectedId && !selected);
  const groups = new Map();
  const era = state.identity.era;
  const ordered = [...catalog].sort((a, b) => Number(b.era.includes(era)) - Number(a.era.includes(era)));
  ordered.forEach((entry) => {
    const group = isWeapon ? entry.group : `${entry.era} · ${entry.group}`;
    if (!groups.has(group)) groups.set(group, []);
    const suffix = entry.pack ? ` · ${entry.pack} 发/盒` : "";
    const eraHint = isWeapon && era !== "其他" && !entry.era.includes(era) ? ` · ${entry.era}` : "";
    groups.get(group).push(`<option value="${escapeAttr(entry.catalog_id)}" ${entry.catalog_id === selectedId ? "selected" : ""}>${escapeHtml(entry.name + suffix + eraHint)}</option>`);
  });
  const options = [...groups].map(([group, entries]) => `<optgroup label="${escapeAttr(group)}">${entries.join("")}</optgroup>`).join("");
  const label = isWeapon ? "武器类型" : "物品名称";
  const customInput = custom ? `<input data-field="${field}" value="${escapeAttr(item[field] || "")}" placeholder="自定义${label}" aria-label="第 ${index + 1} 项自定义${label}">` : "";
  const reference = !isWeapon && selected ? `<small class="equipment-choice-note">${escapeHtml(selected.era)} · 参考 $${escapeHtml(selected.price)}${selected.pack ? ` / ${escapeHtml(selected.pack)} 发` : ""}</small>` : "";
  return `<div class="equipment-choice"><select data-catalog-kind="${kind}" aria-label="第 ${index + 1} 项${label}"><option value="" ${!selectedId ? "selected" : ""} disabled>选择${label}</option><option value="custom" ${custom ? "selected" : ""}>自定义填写</option>${options}</select>${customInput}${reference}</div>`;
}

function weaponSkillOptions(current) {
  const options = state.skills.map((skill) => ({ name: skillDisplayName(skill), value: skillFinal(skill) }));
  if (current && !options.some((skill) => skill.name === current)) options.push({ name: current, value: null });
  return '<option value="">选择技能</option>' + options.map((skill) => `<option value="${escapeAttr(skill.name)}" ${skill.name === current ? "selected" : ""}>${escapeHtml(skill.name)}${skill.value === null ? "" : ` · ${skill.value}`}</option>`).join("");
}

function resolveWeaponSkill(sourceSkill) {
  const specialization = sourceSkill === "弓" ? "弓术" : sourceSkill;
  const existing = state.skills.find((skill) => String(skill.specialization || "").trim() === specialization || skillDisplayName(skill) === specialization);
  if (existing) return skillDisplayName(existing);
  if (["斗殴", "鞭子", "电锯", "链枷", "绞具", "斧", "剑", "矛"].includes(specialization)) return `格斗（${specialization}）`;
  if (["手枪", "步枪/霰弹枪", "冲锋枪", "弓术", "喷射器", "机枪", "重武器"].includes(specialization)) return `射击（${specialization}）`;
  return sourceSkill;
}

function applyEquipmentDefinition(kind, item, definition) {
  if (kind === "weapon") {
    if (!item.name || item.name === item.category) item.name = definition.name;
    item.category = definition.name;
    item.skill = resolveWeaponSkill(definition.skill);
    for (const field of ["damage", "range", "attacks", "ammo", "malfunction"]) item[field] = definition[field];
  } else {
    item.name = definition.name;
  }
  item.catalog_id = definition.catalog_id;
}

function handleEquipmentSelection(event) {
  const kind = event.target.dataset.catalogKind;
  if (!kind) return;
  const row = event.target.closest("[data-equipment-index]");
  if (!row) return;
  const item = (kind === "weapon" ? state.weapons : state.inventory)[Number(row.dataset.equipmentIndex)];
  if (!item) return;
  if (event.target.value === "custom") item.catalog_id = "custom";
  else {
    const definition = equipmentCatalog(kind).find((entry) => entry.catalog_id === event.target.value);
    if (!definition) return;
    applyEquipmentDefinition(kind, item, definition);
  }
  renderEquipment(); markChanged();
}

function handleWeaponInput(event) {
  const row = event.target.closest("[data-equipment-index]"); if (!row || !event.target.dataset.field) return;
  state.weapons[Number(row.dataset.equipmentIndex)][event.target.dataset.field] = event.target.value; markChanged();
}
function handleInventoryInput(event) {
  const row = event.target.closest("[data-equipment-index]"); if (!row || !event.target.dataset.field) return;
  state.inventory[Number(row.dataset.equipmentIndex)][event.target.dataset.field] = event.target.value; markChanged();
}
function handleEquipmentRemove(event) {
  const button = event.target.closest("[data-remove-row]"); if (!button) return;
  const row = button.closest("[data-equipment-index]"); const index = Number(row.dataset.equipmentIndex);
  if (button.dataset.removeRow === "weapon") state.weapons.splice(index, 1); else state.inventory.splice(index, 1);
  renderEquipment(); markChanged();
}

function serializeDraft() {
  return {
    version: state.version,
    identity: { ...state.identity, occupation_name: currentOccupation()?.name || "" },
    attributes: { ...state.attributes },
    occupation: currentOccupation(),
    custom_skill_slots: [...state.custom_skill_slots],
    skills: state.skills.map((skill) => ({ ...skill })),
    background: { ...state.background }, assets: { ...state.assets },
    experience: { ...state.experience },
    weapons: state.weapons.filter((item) => item.name.trim()), inventory: state.inventory.filter((item) => item.name.trim()),
    group_choices: structuredClone(state.group_choices), free_skill_choices: [...state.free_skill_choices],
    nonstandard_override: Boolean(state.nonstandard_override),
  };
}

async function validateDraft(showBusy) {
  const generation = draftGeneration;
  if (showBusy) setBusy(true, "正在检查", "核对点数与导出条件……");
  try {
    const response = await apiFetch(
      "/api/validate",
      { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(serializeDraft()) },
      { timeoutMs: 30000 },
    );
    const data = await readJsonResponse(response, "规则检查失败。");
    if (generation !== draftGeneration) return null;
    if (!response.ok) throw new Error(data.detail || "规则检查失败。");
    validationData = data.validation;
    state.derived = data.derived;
    renderDerived(); renderBudgets(); renderReview(); renderLiveSummary();
    return validationData;
  } catch (error) {
    if (generation !== draftGeneration) return null;
    const message = friendlyErrorMessage(error, "规则检查失败，请稍后重试。");
    validationData = null; renderReviewError(message); toast(message, true); return null;
  } finally { if (showBusy) setBusy(false); }
}

function renderReview() {
  if (!validationData) return;
  const canExport = validationData.can_export;
  const ring = document.getElementById("review-ring");
  ring.style.setProperty("--review-color", canExport ? "var(--teal)" : "var(--danger)");
  setText("review-score", canExport ? "通过" : "阻断");
  setText("review-title", canExport ? "可以导出" : "请修正错误");
  setText("review-summary", canExport ? "" : "点击问题旁的按钮前往修正。" );
  setText("error-count", validationData.counts.error); setText("warning-count", validationData.counts.warning); setText("info-count", validationData.counts.info);
  const host = document.getElementById("issue-list");
  const visibleIssues = validationData.issues.filter((issue) => issue.severity !== "info");
  const infoIssues = validationData.issues.filter((issue) => issue.severity === "info");
  const infoDetails = infoIssues.length
    ? `<details class="compact-details review-info-details"><summary>完整性提醒（${infoIssues.length}）</summary>${infoIssues.map(reviewIssueHtml).join("")}</details>`
    : "";
  host.innerHTML = validationData.issues.length
    ? visibleIssues.map(reviewIssueHtml).join("") + infoDetails
    : '<div class="empty-state">检查通过</div>';
  document.getElementById("export-pdf").disabled = !canExport;
  document.getElementById("export-excel").disabled = !canExport;
  setText("mini-issue-count", `${validationData.issues.length} 项`);
  document.getElementById("mini-notice").innerHTML = canExport ? "<span>✓</span><p>可以导出</p>" : `<span>!</span><p>${validationData.counts.error} 个错误待修正</p>`;
}

function reviewIssueHtml(issue) {
  const diagnostic = `${issue.code}${issue.field ? ` · ${issue.field}` : ""}`;
  return `<div class="issue-item ${escapeAttr(issue.severity)}"><span>${issue.severity === "error" ? "!" : issue.severity === "warning" ? "△" : "i"}</span><div><strong title="${escapeAttr(diagnostic)}">${escapeHtml(issue.message)}</strong></div><button type="button" data-issue-field="${escapeAttr(issue.field)}">前往 →</button></div>`;
}

function renderReviewError(message) {
  setText("review-score", "异常"); setText("review-title", "检查未完成"); setText("review-summary", message);
  document.getElementById("issue-list").innerHTML = `<div class="issue-item error"><span>!</span><div><strong>${escapeHtml(message)}</strong></div></div>`;
  document.getElementById("export-pdf").disabled = true; document.getElementById("export-excel").disabled = true;
}

function handleIssueNavigation(event) {
  const button = event.target.closest("[data-issue-field]"); if (!button) return;
  const field = button.dataset.issueField;
  let step = 6;
  if (field.startsWith("identity")) step = 0; else if (field.startsWith("attributes")) step = 1; else if (field.startsWith("occupation") || field.startsWith("group_choices") || field === "free_skill_choices") step = 2; else if (field === "skills" || /^F|^AB/.test(field)) step = 3; else if (field.startsWith("background") || field.startsWith("assets")) step = 4; else if (field.startsWith("weapons") || field.startsWith("inventory")) step = 5;
  goToStep(step);
}

async function exportPdf() {
  if (importBusy || hasPendingImport()) return;
  const generation = draftGeneration;
  const validation = await validateDraft(false);
  if (generation !== draftGeneration || importBusy || hasPendingImport()) return;
  if (!validation?.can_export) { toast("请先修正阻断错误，再生成 PDF。", true); return; }
  setBusy(true, "正在生成 PDF", "排版并生成预览……");
  try {
    const form = exportFormData();
    const response = await apiFetch("/api/export/pdf", { method: "POST", body: form }, { timeoutMs: 180000 });
    const data = await readJsonResponse(response, "PDF 生成失败。");
    if (generation !== draftGeneration) return;
    if (!response.ok) throw new Error(data.detail || "PDF 生成失败。");
    const blob = base64Blob(data.pdf_base64, "application/pdf");
    if (pdfObjectUrl) URL.revokeObjectURL(pdfObjectUrl);
    pdfObjectUrl = URL.createObjectURL(blob);
    const link = document.getElementById("download-pdf"); link.href = pdfObjectUrl; link.download = data.filename; link.hidden = false;
    if (data.previews?.length >= 2) {
      document.getElementById("pdf-preview-section").hidden = false;
      document.getElementById("pdf-preview-1").src = `data:image/png;base64,${data.previews[0]}`;
      document.getElementById("pdf-preview-2").src = `data:image/png;base64,${data.previews[1]}`;
    }
    toast("PDF 已生成。", false);
  } catch (error) { toast(friendlyErrorMessage(error, "PDF 生成失败，请稍后重试。"), true); }
  finally { setBusy(false); }
}

async function exportExcel() {
  if (importBusy || hasPendingImport()) return;
  const generation = draftGeneration;
  const validation = await validateDraft(false);
  if (generation !== draftGeneration || importBusy || hasPendingImport()) return;
  if (!validation?.can_export) { toast("请先修正阻断错误，再生成 Excel。", true); return; }
  setBusy(true, "正在生成 Excel", "填写资料并重算公式……");
  try {
    const response = await apiFetch("/api/export/excel", { method: "POST", body: exportFormData() }, { timeoutMs: 600000 });
    if (!response.ok) { const data = await readJsonResponse(response, "Excel 生成失败。"); throw new Error(data.detail || "Excel 生成失败。"); }
    const blob = await response.blob();
    if (generation !== draftGeneration) return;
    const filename = responseFilename(response.headers.get("Content-Disposition")) || "COC7_调查员.xlsx";
    const url = URL.createObjectURL(blob); const link = document.createElement("a"); link.href = url; link.download = filename; document.body.appendChild(link); link.click(); link.remove(); setTimeout(() => URL.revokeObjectURL(url), 1000);
    toast("Excel 已生成。", false);
  } catch (error) { toast(friendlyErrorMessage(error, "Excel 生成失败，请稍后重试。"), true); }
  finally { setBusy(false); }
}

function exportFormData() {
  const form = new FormData(); form.append("draft_json", JSON.stringify(serializeDraft())); if (portraitFile) form.append("portrait", portraitFile, portraitFile.name); return form;
}

function bindImportControls() {
  const input = document.getElementById("import-excel-input");
  const dialog = document.getElementById("import-dialog");
  document.getElementById("import-excel").addEventListener("click", () => {
    if (importBusy || hasPendingImport()) return;
    input.value = "";
    input.click();
  });
  input.addEventListener("change", () => {
    const file = input.files?.[0];
    input.value = "";
    if (file) previewExcelImport(file);
  });
  dialog.addEventListener("close", () => { pendingImport = null; });
  document.getElementById("import-confirm").addEventListener("click", confirmExcelImport);
  document.getElementById("import-undo").addEventListener("click", undoExcelImport);
}

function setImportBusy(active, title = "正在读取 Excel") {
  importBusy = active;
  for (const id of ["import-excel", "import-undo", "import-confirm", "import-attributes", "attribute-import-undo", "attribute-import-confirm"]) document.getElementById(id).disabled = active;
  setBusy(active, title, "请稍候……");
}

async function previewExcelImport(file) {
  if (importBusy || hasPendingImport()) return;
  if (!/\.xlsx$/i.test(file.name)) { toast("请选择 .xlsx 调查员表格。", true); return; }
  if (file.size > 12 * 1024 * 1024) { toast("Excel 文件不能超过 12 MB。", true); return; }
  setImportBusy(true);
  try {
    const form = new FormData();
    form.append("workbook", file, file.name);
    const response = await apiFetch("/api/import/excel", { method: "POST", body: form }, { timeoutMs: 60000 });
    const data = await readJsonResponse(response, "Excel 读取失败。");
    if (!response.ok) throw new Error(data.detail || "Excel 读取失败。");
    if (!data.draft || typeof data.draft !== "object" || Array.isArray(data.draft)) throw new Error("表格中未找到可导入的调查员资料。");
    const portrait = importedPortraitFile(data.portrait);
    const draft = structuredClone(data.draft);
    const summary = data.summary || {};
    setText("import-filename", file.name);
    const entries = [
      ["姓名", summary.name || draft.identity?.name || "未命名"],
      ["职业", summary.occupation || "未填写"],
      ["时代", summary.era || draft.identity?.era || "未填写"],
      ["资料", `${numeric(summary.skill_count)} 项技能 · ${numeric(summary.weapon_count)} 件武器 · ${numeric(summary.inventory_count)} 件物品`],
    ];
    const summaryList = document.getElementById("import-summary");
    summaryList.replaceChildren();
    for (const [label, value] of entries) {
      const term = document.createElement("dt"); term.textContent = label;
      const description = document.createElement("dd"); description.textContent = String(value);
      summaryList.append(term, description);
    }
    const warnings = Array.isArray(data.warnings) ? data.warnings : [];
    document.getElementById("import-warnings").hidden = warnings.length === 0;
    const warningList = document.getElementById("import-warning-list");
    warningList.replaceChildren();
    for (const warning of warnings) {
      const item = document.createElement("li"); item.textContent = String(warning); warningList.append(item);
    }
    pendingImport = { draft, portrait };
    const dialog = document.getElementById("import-dialog");
    dialog.returnValue = "";
    dialog.showModal();
  } catch (error) {
    pendingImport = null;
    toast(friendlyErrorMessage(error, "Excel 读取失败，请检查文件后重试。"), true);
  } finally { setImportBusy(false); }
}

function importedPortraitFile(portrait) {
  if (!portrait) return null;
  if (!["image/png", "image/jpeg", "image/webp"].includes(portrait.mime_type)) throw new Error("表格中的头像格式无法读取。");
  const blob = base64Blob(portrait.data_base64, portrait.mime_type);
  if (blob.size > 8 * 1024 * 1024) throw new Error("表格中的头像超过 8 MB。");
  return new File([blob], String(portrait.filename || "investigator-portrait"), { type: portrait.mime_type });
}

function hasPendingImport() { return Boolean(pendingImport || pendingAttributeImport); }

function bindAttributeImportControls() {
  const input = document.getElementById("import-attributes-input");
  document.getElementById("import-attributes").addEventListener("click", () => {
    if (importBusy || hasPendingImport()) return;
    input.value = "";
    input.click();
  });
  input.addEventListener("change", () => {
    const file = input.files?.[0];
    input.value = "";
    if (file) previewAttributeImport(file);
  });
  document.getElementById("attribute-import-dialog").addEventListener("close", () => { pendingAttributeImport = null; });
  document.getElementById("attribute-import-confirm").addEventListener("click", confirmAttributeImport);
  document.getElementById("attribute-import-undo").addEventListener("click", undoAttributeImport);
}

async function previewAttributeImport(file) {
  if (importBusy || hasPendingImport()) return;
  if (!/\.(xlsx|csv|tsv)$/i.test(file.name)) { toast("请选择 XLSX、CSV 或 TSV 属性简化表。", true); return; }
  if (file.size > 12 * 1024 * 1024) { toast("属性表不能超过 12 MB。", true); return; }
  const generation = draftGeneration;
  setImportBusy(true, "正在读取属性表");
  try {
    const form = new FormData();
    form.append("workbook", file, file.name);
    const response = await apiFetch("/api/import/attributes", { method: "POST", body: form }, { timeoutMs: 60000 });
    const data = await readJsonResponse(response, "属性表读取失败。");
    if (generation !== draftGeneration) return;
    if (!response.ok) throw new Error(data.detail || "属性表读取失败。");
    if (!data.attributes || typeof data.attributes !== "object" || Array.isArray(data.attributes)) throw new Error("表格中未找到可导入的属性。");
    const attributes = {};
    for (const key of ATTRIBUTE_KEYS) {
      if (!Object.prototype.hasOwnProperty.call(data.attributes, key)) continue;
      const value = data.attributes[key];
      if (!Number.isInteger(value) || value < 0 || value > 300) throw new Error(`${ATTRIBUTE_LABELS[key]}必须为 0—300 的整数。`);
      attributes[key] = value;
    }
    if (!Object.keys(attributes).length) throw new Error("表格中未找到可导入的属性。");
    setText("attribute-import-filename", file.name);
    const summaryList = document.getElementById("attribute-import-summary");
    summaryList.replaceChildren();
    for (const [key, value] of Object.entries(attributes)) {
      const term = document.createElement("dt"); term.textContent = `${ATTRIBUTE_LABELS[key]} ${key}`;
      const description = document.createElement("dd");
      description.textContent = `${state.attributes[key]} → ${value}`;
      if (state.attributes[key] === value) description.className = "unchanged";
      summaryList.append(term, description);
    }
    const warnings = Array.isArray(data.warnings) ? data.warnings : [];
    document.getElementById("attribute-import-warnings").hidden = warnings.length === 0;
    const warningList = document.getElementById("attribute-import-warning-list");
    warningList.replaceChildren();
    for (const warning of warnings) {
      const item = document.createElement("li"); item.textContent = String(warning); warningList.append(item);
    }
    pendingAttributeImport = attributes;
    const dialog = document.getElementById("attribute-import-dialog");
    dialog.returnValue = "";
    dialog.showModal();
  } catch (error) {
    pendingAttributeImport = null;
    if (generation === draftGeneration) toast(friendlyErrorMessage(error, "属性表读取失败，请检查文件后重试。"), true);
  } finally { setImportBusy(false); }
}

async function confirmAttributeImport() {
  if (importBusy || !pendingAttributeImport) return;
  const attributes = pendingAttributeImport;
  const previous = {};
  const changed = {};
  for (const [key, value] of Object.entries(attributes)) {
    if (state.attributes[key] === value) continue;
    previous[key] = state.attributes[key];
    changed[key] = value;
  }
  document.getElementById("attribute-import-dialog").close("confirm");
  pendingAttributeImport = null;
  if (!Object.keys(changed).length) { toast("属性与当前值相同。", false); return; }
  setImportBusy(true, "正在导入属性");
  try {
    applyAttributeValues(changed);
    lastAttributeImportSnapshot = previous;
    document.getElementById("attribute-import-undo").hidden = false;
    await refreshDerived(true);
    toast(`已导入 ${Object.keys(changed).length} 项属性。`, false);
  } finally { setImportBusy(false); }
}

async function undoAttributeImport() {
  if (importBusy || hasPendingImport() || !lastAttributeImportSnapshot) return;
  const previous = lastAttributeImportSnapshot;
  setImportBusy(true, "正在撤销属性导入");
  try {
    applyAttributeValues(previous);
    lastAttributeImportSnapshot = null;
    document.getElementById("attribute-import-undo").hidden = true;
    await refreshDerived(true);
    toast("已恢复本次导入前的属性。", false);
  } finally { setImportBusy(false); }
}

function applyAttributeValues(attributes) {
  draftGeneration += 1;
  clearTimeout(saveTimer);
  clearTimeout(derivedTimer);
  Object.assign(state.attributes, attributes);
  for (const key of Object.keys(attributes)) document.getElementById(`attr-${key}`).value = state.attributes[key];
  state.derived = null;
  for (const id of ["derived-hp", "derived-mp", "derived-san", "derived-mov", "derived-dodge", "derived-damage", "derived-build", "derived-interest"]) setText(id, "—");
  syncDynamicSkillBases();
  clearInvestigatorOutputs();
  renderSkills();
  setText("occupation-points", "—");
  renderLiveSummary();
  markChanged();
}

function clearAttributeImport() {
  pendingAttributeImport = null;
  lastAttributeImportSnapshot = null;
  document.getElementById("attribute-import-undo").hidden = true;
  const dialog = document.getElementById("attribute-import-dialog");
  if (dialog.open) dialog.close();
}

async function confirmExcelImport() {
  if (importBusy || !pendingImport) return;
  const imported = pendingImport;
  const previous = { draft: structuredClone(state), portrait: portraitFile };
  document.getElementById("import-dialog").close("confirm");
  pendingImport = null;
  setImportBusy(true, "正在导入调查员");
  try {
    replaceInvestigator(imported.draft, imported.portrait, true);
    lastImportSnapshot = previous;
    document.getElementById("import-undo").hidden = false;
    await refreshDerived(true);
    toast("调查员已导入，可继续编辑。", false);
  } catch (error) {
    replaceInvestigator(previous.draft, previous.portrait);
    toast(friendlyErrorMessage(error, "导入未完成，已保留原调查员。"), true);
  } finally { setImportBusy(false); }
}

async function undoExcelImport() {
  if (importBusy || hasPendingImport() || !lastImportSnapshot) return;
  const previous = lastImportSnapshot;
  setImportBusy(true, "正在撤销导入");
  try {
    replaceInvestigator(previous.draft, previous.portrait);
    lastImportSnapshot = null;
    document.getElementById("import-undo").hidden = true;
    await refreshDerived(true);
    toast("已恢复导入前的调查员。", false);
  } finally { setImportBusy(false); }
}

function replaceInvestigator(draft, portrait, startAtIdentity = false) {
  clearAttributeImport();
  draftGeneration += 1;
  clearTimeout(saveTimer);
  clearTimeout(derivedTimer);
  state = { ...createDefaultState(), ...structuredClone(draft) };
  normalizeState();
  if (startAtIdentity) state.current_step = 0;
  for (const id of ["occupation-search", "free-choice-search", "custom-skill-search", "skill-search"]) document.getElementById(id).value = "";
  document.getElementById("skill-filter").value = "all";
  for (const id of ["occupation-details", "experience-details", "asset-source-details"]) document.getElementById(id).open = false;
  clearInvestigatorOutputs();
  hydrateControls();
  renderOccupation();
  syncOccupationSkills();
  renderSkills();
  renderEquipment();
  renderAssetReference();
  renderDerived();
  renderLiveSummary();
  setPortrait(portrait);
  document.getElementById("portrait-input").value = "";
  goToStep(state.current_step, false);
  markChanged();
}

function clearInvestigatorOutputs() {
  validationData = null;
  if (pdfObjectUrl) URL.revokeObjectURL(pdfObjectUrl);
  pdfObjectUrl = null;
  const link = document.getElementById("download-pdf");
  link.hidden = true; link.removeAttribute("href"); link.removeAttribute("download");
  document.getElementById("pdf-preview-section").hidden = true;
  for (const id of ["pdf-preview-1", "pdf-preview-2"]) document.getElementById(id).removeAttribute("src");
  document.getElementById("review-ring").style.removeProperty("--review-color");
  setText("review-score", "—"); setText("review-title", "等待检查"); setText("review-summary", "");
  for (const id of ["error-count", "warning-count", "info-count"]) setText(id, 0);
  document.getElementById("issue-list").innerHTML = '<div class="empty-state">尚未执行检查。</div>';
  document.getElementById("export-pdf").disabled = true;
  document.getElementById("export-excel").disabled = true;
  setText("mini-issue-count", "0 项");
  document.getElementById("mini-notice").innerHTML = "<span>i</span><p>等待检查</p>";
}

function bindPortraitControls() {
  const input = document.getElementById("portrait-input"); const drop = document.getElementById("portrait-drop");
  input.addEventListener("change", () => setPortrait(input.files?.[0] || null));
  ["dragenter", "dragover"].forEach((name) => drop.addEventListener(name, (event) => { event.preventDefault(); drop.classList.add("dragover"); }));
  ["dragleave", "drop"].forEach((name) => drop.addEventListener(name, (event) => { event.preventDefault(); drop.classList.remove("dragover"); }));
  drop.addEventListener("drop", (event) => setPortrait(event.dataTransfer.files?.[0] || null));
  document.getElementById("portrait-remove").addEventListener("click", () => setPortrait(null));
}

function setPortrait(file) {
  if (file && !["image/png", "image/jpeg", "image/webp"].includes(file.type)) { toast("头像只支持 PNG、JPEG 或 WebP。", true); return; }
  if (file && file.size > 8 * 1024 * 1024) { toast("头像文件不能超过 8 MB。", true); return; }
  if (portraitUrl) URL.revokeObjectURL(portraitUrl);
  portraitFile = file; portraitUrl = file ? URL.createObjectURL(file) : null;
  const preview = document.getElementById("portrait-preview"); const placeholder = document.getElementById("portrait-placeholder"); const mini = document.getElementById("mini-portrait-image"); const miniPlaceholder = document.getElementById("mini-portrait-placeholder");
  preview.hidden = !file; placeholder.hidden = Boolean(file); mini.hidden = !file; miniPlaceholder.hidden = Boolean(file); document.getElementById("portrait-remove").hidden = !file;
  if (file) { preview.src = portraitUrl; mini.src = portraitUrl; } else { preview.removeAttribute("src"); mini.removeAttribute("src"); document.getElementById("portrait-input").value = ""; }
}

function markChanged() { validationData = null; document.getElementById("save-state").textContent = "正在保存……"; scheduleSave(); }
function scheduleSave() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    try { sessionStorage.setItem(STORAGE_KEY, JSON.stringify(state)); document.getElementById("save-state").textContent = "会话已保存"; }
    catch (_) { document.getElementById("save-state").textContent = "会话存储空间不足"; }
  }, 180);
}

function setBusy(active, title = "正在处理……", detail = "请保持此页面打开。") {
  document.getElementById("busy-overlay").hidden = !active; setText("busy-title", title); setText("busy-detail", detail);
}

function toast(message, isError) {
  const displayMessage = isError ? friendlyErrorMessage(message) : String(message || "");
  const now = Date.now();
  if (displayMessage === lastToastMessage && now - lastToastAt < 1800) return;
  lastToastMessage = displayMessage;
  lastToastAt = now;
  const node = document.createElement("div"); node.className = `toast${isError ? " error" : ""}`; node.textContent = displayMessage; document.getElementById("toast-region").appendChild(node); setTimeout(() => node.remove(), 4200);
}

function base64Blob(value, type) {
  const binary = atob(value); const bytes = new Uint8Array(binary.length); for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index); return new Blob([bytes], { type });
}
function responseFilename(header) { const match = String(header || "").match(/filename\*=UTF-8''([^;]+)/i); return match ? decodeURIComponent(match[1]) : ""; }
function skillDisplayName(skill) { const clean = skill.name.replace(/[：:]$/, "").trim(); return String(skill.specialization || "").trim() ? `${clean}（${skill.specialization.trim()}）` : clean; }
function formulaLabel(value) { return String(value || "").replaceAll("*", "×"); }
function splitSkillToken(token) { const match = String(token || "").trim().match(/^(.+?)（(.+)）$/); return match ? [normalizeSkillName(match[1]), match[2].trim()] : [normalizeSkillName(token), ""]; }
function normalizeSkillName(value) { return String(value || "").replaceAll(" Ω", "").replaceAll("Ω", "").trim().replace(/\s+/g, " "); }
function normalizeText(value) { return String(value || "").trim().toLocaleLowerCase("zh-CN").replace(/\s+/g, " "); }
function numeric(value, fallback = 0) { const parsed = Number(value); return Number.isFinite(parsed) ? Math.trunc(parsed) : fallback; }
function clampNumber(value, min, max, fallback) { return Math.max(min, Math.min(max, numeric(value, fallback))); }
function setText(id, value) { const element = document.getElementById(id); if (element) element.textContent = String(value ?? ""); }
function escapeHtml(value) { return String(value ?? "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char])); }
function escapeAttr(value) { return escapeHtml(value).replace(/`/g, "&#96;"); }
