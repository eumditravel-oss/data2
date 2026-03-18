"use strict";

const $ = (id) => document.getElementById(id);

const fileInputs = { current: $("file-current"), A: $("file-a"), B: $("file-b"), C: $("file-c") };
const fileListEls = { current: $("file-current-list"), A: $("file-a-list"), B: $("file-b-list"), C: $("file-c-list") };

const compareBody = $("compare-body");
const statusBox = $("status-box");
const logBox = $("log-box");
const tabContainer = $("tab-container");

const btnExtract = $("btn-extract");
const btnOpenMapping = $("btn-open-mapping");
const btnCalc = $("btn-calc");
const btnExportCsv = $("btn-export-csv");
const btnReset = $("btn-reset");

const mappingModal = $("mapping-modal");
const mappingBackdrop = $("mapping-backdrop");
const btnCloseMapping = $("btn-close-mapping");
const btnApplySuggestions = $("btn-apply-suggestions");
const btnSaveMapping = $("btn-save-mapping");
const mappingBody = $("mapping-body");
const typeaheadRoot = $("typeahead-root");

const PROJECT_KEYS = ["current", "A", "B", "C"];
const PROJECT_LABELS = { current: "현재 프로젝트", A: "A 프로젝트", B: "B 프로젝트", C: "C 프로젝트" };

const CATEGORY_OPTIONS = [
  { value: "", label: "선택" }, { value: "레미콘", label: "레미콘" },
  { value: "거푸집", label: "거푸집" }, { value: "철근", label: "철근" }, { value: "제외", label: "제외" }
];
const INCLUDE_OPTIONS = [ { value: "Y", label: "반영" }, { value: "N", label: "제외" } ];
const DEFAULT_ITEM_CODE_OPTIONS = ["240", "270", "300", "180", "3회", "4회", "유로", "알폼", "갱폼", "합벽", "보밑면", "데크", "방수턱", "H10", "H13", "H16", "H19", "H22", "H25", "H29"];

// 기존의 "섹션(동)" 개념이 사라졌으므로, 뼈대가 되는 베이스 레이아웃 1세트만 정의
const BASE_LAYOUT = [
  { itemCode: "240", item: "레미콘", spec: "25-24-15", category: "레미콘" },
  { itemCode: "270", item: "레미콘", spec: "25-27-15", category: "레미콘" },
  { itemCode: "300", item: "레미콘", spec: "25-30-15", category: "레미콘" },
  { itemCode: "180", item: "레미콘", spec: "25-18-08", category: "레미콘" },
  { itemCode: "3회", item: "거푸집", spec: "3회", category: "거푸집" },
  { itemCode: "4회", item: "거푸집", spec: "4회", category: "거푸집" },
  { itemCode: "유로", item: "거푸집", spec: "유로", category: "거푸집" },
  { itemCode: "알폼", item: "거푸집", spec: "알폼", category: "거푸집" },
  { itemCode: "갱폼", item: "거푸집", spec: "갱폼", category: "거푸집" },
  { itemCode: "합벽", item: "거푸집", spec: "합벽", category: "거푸집" },
  { itemCode: "보밑면", item: "거푸집", spec: "보밑면", category: "거푸집" },
  { itemCode: "데크", item: "거푸집", spec: "데크플레이트", category: "거푸집" },
  { itemCode: "방수턱", item: "거푸집", spec: "방수턱", category: "거푸집" },
  { itemCode: "H10", item: "철근", spec: "H10", category: "철근" },
  { itemCode: "H13", item: "철근", spec: "H13", category: "철근" },
  { itemCode: "H16", item: "철근", spec: "H16", category: "철근" },
  { itemCode: "H19", item: "철근", spec: "H19", category: "철근" },
  { itemCode: "H22", item: "철근", spec: "H22", category: "철근" },
  { section: "APT", itemCode: "H25", item: "철근", spec: "H25", category: "철근" },
  { section: "APT", itemCode: "H29", item: "철근", spec: "H29", category: "철근" },
];

const state = {
  rawEntriesByProject: { current: [], A: [], B: [], C: [] },
  fileSummaryByProject: { current: [], A: [], B: [], C: [] },
  uniqueNames: [],
  mappingConfig: {},
  lastCompareRows: [],
  itemCodeOptions: [...DEFAULT_ITEM_CODE_OPTIONS],
  typeahead: { targetInput: null, targetKey: "", items: [], activeIndex: -1 },
  sections: [], // 엑셀에서 추출한 층 목록을 동적으로 저장 (지상1층, 지상2층 등)
  activeTab: "", 
};

function setStatus(text) { statusBox.textContent = text; }
function setLog(text) { logBox.textContent = text; }
function escapeHtml(value) { return String(value ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&#039;"); }
function normalizeText(value) { return String(value ?? "").replace(/\s+/g, "").trim().toUpperCase(); }
function normalizeDisplayText(value) { return String(value ?? "").replace(/\s+/g, " ").trim(); }
function toNumber(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const cleaned = String(value ?? "").replace(/,/g, "").trim();
  if (!cleaned) return 0;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}
function fmtNumber(value) { return Number(value || 0).toLocaleString("ko-KR", { maximumFractionDigits: 0 }); }
function fmtRatio(value) {
  if (!value && value !== 0) return "0%";
  return (Number(value) * 100).toFixed(0) + "%";
}
function ratioClass(value) {
  if (!value) return "";
  if (value < 0.9) return "bad";
  if (value <= 1.1) return "good";
  return "warn";
}
function updateFileListText() {
  for (const key of PROJECT_KEYS) {
    const files = Array.from(fileInputs[key].files || []);
    if (!files.length) { fileListEls[key].textContent = "선택된 파일 없음"; continue; }
    fileListEls[key].textContent = files.map((f) => f.name).join("\n");
  }
}
function openModal() { mappingModal.classList.add("is-open"); mappingModal.setAttribute("aria-hidden", "false"); }
function closeModal() { mappingModal.classList.remove("is-open"); mappingModal.setAttribute("aria-hidden", "true"); hideTypeahead(); }
function sortUniqueStrings(list) { return [...new Set(list.filter(Boolean))].sort((a, b) => a.localeCompare(b, "ko")); }
function ensureItemCodeOption(value) {
  const text = normalizeDisplayText(value);
  if (!text) return;
  if (!state.itemCodeOptions.includes(text)) {
    state.itemCodeOptions.push(text);
    state.itemCodeOptions = sortUniqueStrings(state.itemCodeOptions);
  }
}

/** -----------------------------
 * 탭 동적 렌더링 및 UI 업데이트
 * ----------------------------- */
function renderTabs() {
  if (state.sections.length === 0) {
    tabContainer.innerHTML = "";
    return;
  }
  
  tabContainer.innerHTML = state.sections.map(sec => {
    const isActive = sec === state.activeTab;
    const opacity = isActive ? '1' : '0.4';
    const border = isActive ? '2px' : '1px';
    const bgColor = isActive ? '#eef2f7' : '#ffffff';
    return `<span class="legend__chip legend__chip--dynamic" data-tab="${escapeHtml(sec)}" 
                  style="cursor: pointer; opacity: ${opacity}; border-width: ${border}; background-color: ${bgColor}; border-style: solid; border-color: #d8deea; padding: 6px 14px; border-radius: 20px; font-weight: bold;">
              ${escapeHtml(sec)}
            </span>`;
  }).join("");

  const chips = tabContainer.querySelectorAll(".legend__chip--dynamic");
  chips.forEach(chip => {
    chip.addEventListener("click", (e) => {
      state.activeTab = e.currentTarget.getAttribute("data-tab");
      renderTabs(); // 탭 색상 새로고침
      if (state.lastCompareRows.length > 0) renderCompareTable(state.lastCompareRows);
    });
  });
}

/** -----------------------------
 * 엑셀 파싱 (완전히 새로 설계된 '층별(2행 묶음)' 추출 로직)
 * ----------------------------- */
function parseSheetEntries(rows, projectKey, fileName) {
  const headerRow3 = rows[2] || []; // 노란색 타이틀
  const headerRow4 = rows[3] || []; // 주황색 타이틀

  const colStart = 1; 
  const colEnd = Math.max(headerRow3.length, headerRow4.length) - 1;
  const entries = [];
  const skipHeaders = ["합계", "비고", "누계", "소계", ""];

  // 데이터는 6행(인덱스 5)부터 시작하며 2줄씩 짝지어서 읽음
  for (let r = 5; r < rows.length; r += 2) {
    const row1 = rows[r] || [];     // 층의 첫 번째 행 (노란색 수량 매칭)
    const row2 = rows[r + 1] || []; // 층의 두 번째 행 (주황색 면적 매칭)
    
    // A열, B열, C열 중 텍스트가 있는 것을 '층 이름'으로 인식
    let floorName = normalizeDisplayText(row1[0]) || normalizeDisplayText(row1[1]) || normalizeDisplayText(row1[2]);
    
    // 합계 텍스트가 들어간 줄이거나 층 이름이 없으면 건너뜀
    if (!floorName || floorName.includes("합계") || floorName.includes("소계") || floorName === "계") {
      continue; 
    }

    for (let c = colStart; c <= colEnd; c += 1) {
      const name3 = normalizeDisplayText(headerRow3[c]);
      const name4 = normalizeDisplayText(headerRow4[c]);

      // 3행(노란색) 타이틀이 있으면 첫 번째 줄(row1) 데이터와 연결
      if (name3 && !skipHeaders.includes(name3)) {
        const val = toNumber(row1[c]);
        if (val !== 0) {
          entries.push({
            projectKey, projectLabel: PROJECT_LABELS[projectKey], section: floorName, sourceFile: fileName, rowIndex: r + 1, rowLabel: "1행", rawName: name3, normalizedName: normalizeText(name3), qty: val
          });
        }
      }

      // 4행(주황색) 타이틀이 있으면 두 번째 줄(row2) 데이터와 연결
      if (name4 && !skipHeaders.includes(name4)) {
        const val = toNumber(row2[c]);
        if (val !== 0) {
          entries.push({
            projectKey, projectLabel: PROJECT_LABELS[projectKey], section: floorName, sourceFile: fileName, rowIndex: r + 2, rowLabel: "2행", rawName: name4, normalizedName: normalizeText(name4), qty: val
          });
        }
      }
    }
  }

  return entries;
}

async function parseWorkbookFile(file, projectKey) {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const allEntries = []; let totalEntryCount = 0; 

  // 파일 내 모든 시트를 돌면서 데이터 추출
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
    const entries = parseSheetEntries(rows, projectKey, file.name);
    allEntries.push(...entries);
    totalEntryCount += entries.length;
  }
  return { projectKey, fileName: file.name, entryCount: totalEntryCount, entries: allEntries };
}

/** -----------------------------
 * 명칭 매핑 및 제안
 * ----------------------------- */
function suggestMappingByName(rawName) {
  const t = normalizeText(rawName);
  const result = { include: "Y", category: "", itemCode: "", note: "" };

  if (t.includes("24MPA")) { result.category = "레미콘"; result.itemCode = "240"; result.note = "24MPA → 240"; return result; }
  if (t.includes("27MPA")) { result.category = "레미콘"; result.itemCode = "270"; result.note = "27MPA → 270"; return result; }
  if (t.includes("30MPA")) { result.category = "레미콘"; result.itemCode = "300"; result.note = "30MPA → 300"; return result; }
  if (t.includes("현치/무근") || t.includes("버림")) { result.category = "레미콘"; result.itemCode = "180"; result.note = "버림/현치무근 → 180"; return result; }
  if (t.includes("현치/구체") || t === "합벽채움" || t.includes("ACT-INSIDE")) { result.category = "레미콘"; result.itemCode = "240"; result.note = "구체/합벽/ACT-INSIDE 계열 → 240"; return result; }

  if (t === "3회") { result.category = "거푸집"; result.itemCode = "3회"; result.note = "3회 → 3회"; return result; }
  if (t === "4회") { result.category = "거푸집"; result.itemCode = "4회"; result.note = "4회 → 4회"; return result; }
  if (t === "유로") { result.category = "거푸집"; result.itemCode = "유로"; result.note = "유로 → 유로"; return result; }
  if (t === "원형" || t === "CURVED") { result.category = "거푸집"; result.itemCode = "유로"; result.note = "원형/CURVED → 유로"; return result; }
  if (t === "알폼-H" || t === "알폼-V" || t === "알폼") { result.category = "거푸집"; result.itemCode = "알폼"; result.note = "알폼-H/V → 알폼"; return result; }
  if (t === "갱폼" || t === "갱폼-2") { result.category = "거푸집"; result.itemCode = "갱폼"; result.note = "갱폼/갱폼-2 → 갱폼"; return result; }
  if (t === "합벽") { result.category = "거푸집"; result.itemCode = "합벽"; result.note = "합벽 → 합벽"; return result; }
  if (t === "보밑면") { result.category = "거푸집"; result.itemCode = "보밑면"; result.note = "보밑면 → 보밑면"; return result; }
  if (t === "DECK") { result.category = "거푸집"; result.itemCode = "데크"; result.note = "DECK → 데크"; return result; }
  if (t === "방수턱") { result.category = "거푸집"; result.itemCode = "방수턱"; result.note = "방수턱 → 방수턱"; return result; }

  if (t === "H10") { result.category = "철근"; result.itemCode = "H10"; result.note = "H10 → H10"; return result; }
  if (t === "H13") { result.category = "철근"; result.itemCode = "H13"; result.note = "H13 → H13"; return result; }
  if (t === "H16") { result.category = "철근"; result.itemCode = "H16"; result.note = "H16 → H16"; return result; }
  if (t === "H19") { result.category = "철근"; result.itemCode = "H19"; result.note = "H19 → H19"; return result; }
  if (t === "D22" || t === "H22") { result.category = "철근"; result.itemCode = "H22"; result.note = "D22/H22 → H22"; return result; }
  if (t === "H25") { result.category = "철근"; result.itemCode = "H25"; result.note = "H25 → H25"; return result; }
  if (t === "H29") { result.category = "철근"; result.itemCode = "H29"; result.note = "H29 → H29"; return result; }

  result.include = "N"; result.category = "제외"; result.itemCode = ""; result.note = "자동 제안 없음";
  return result;
}

function buildUniqueNamesFromEntries() {
  const nameMap = new Map();
  for (const projectKey of PROJECT_KEYS) {
    for (const entry of state.rawEntriesByProject[projectKey]) {
      if (!nameMap.has(entry.normalizedName)) {
        nameMap.set(entry.normalizedName, { rawName: entry.rawName, normalizedName: entry.normalizedName });
      }
    }
  }
  state.uniqueNames = Array.from(nameMap.values()).sort((a, b) => a.rawName.localeCompare(b.rawName, "ko"));
}

function ensureMappingConfig() {
  for (const item of state.uniqueNames) {
    if (!state.mappingConfig[item.normalizedName]) {
      const suggested = suggestMappingByName(item.rawName);
      state.mappingConfig[item.normalizedName] = suggested;
      if (suggested.itemCode) ensureItemCodeOption(suggested.itemCode);
    }
  }
}

function optionHtml(list, selectedValue) {
  return list.map((opt) => {
    const selected = String(opt.value) === String(selectedValue) ? "selected" : "";
    return `<option value="${escapeHtml(opt.value)}" ${selected}>${escapeHtml(opt.label)}</option>`;
  }).join("");
}

function getCategoryClass(category) {
  if (category === "레미콘") return "is-concrete";
  if (category === "거푸집") return "is-form";
  if (category === "철근") return "is-rebar";
  return "is-exclude";
}
function getRowIncludeClass(include) { return include === "Y" ? "is-included" : "is-excluded"; }

function renderMappingTable() {
  if (!state.uniqueNames.length) {
    mappingBody.innerHTML = `<tr><td colspan="5">추출된 명칭이 없습니다.</td></tr>`;
    return;
  }
  const html = state.uniqueNames.map((item) => {
    const config = state.mappingConfig[item.normalizedName] || suggestMappingByName(item.rawName);
    const suggestion = suggestMappingByName(item.rawName);
    return `
      <tr class="map-row ${getRowIncludeClass(config.include)}" data-name-key="${escapeHtml(item.normalizedName)}">
        <td>${escapeHtml(item.rawName)}</td>
        <td><select class="map-include">${optionHtml(INCLUDE_OPTIONS, config.include)}</select></td>
        <td><select class="map-category ${getCategoryClass(config.category)}">${optionHtml(CATEGORY_OPTIONS, config.category)}</select></td>
        <td class="itemcode-cell">
          <div class="itemcode-editor">
            <input type="text" class="itemcode-input" value="${escapeHtml(config.itemCode || "")}" autocomplete="off" placeholder="직접 입력 또는 Alt+↓" />
            <button type="button" class="itemcode-add-btn">추가</button>
          </div>
        </td>
        <td class="mapping-suggest">${escapeHtml(suggestion.note || "-")}</td>
      </tr>`;
  }).join("");
  mappingBody.innerHTML = html;
  bindMappingRowBehaviors();
}

function applyVisualStateToRow(row) {
  const includeSelect = row.querySelector(".map-include");
  const categorySelect = row.querySelector(".map-category");
  const include = includeSelect?.value || "N";
  const category = categorySelect?.value || "";
  row.classList.remove("is-included", "is-excluded");
  row.classList.add(include === "Y" ? "is-included" : "is-excluded");
  categorySelect.classList.remove("is-concrete", "is-form", "is-rebar", "is-exclude");
  categorySelect.classList.add(getCategoryClass(category));
}

function bindMappingRowBehaviors() {
  const rows = Array.from(mappingBody.querySelectorAll("tr[data-name-key]"));
  rows.forEach((row) => {
    const includeSelect = row.querySelector(".map-include");
    const categorySelect = row.querySelector(".map-category");
    const itemInput = row.querySelector(".itemcode-input");
    const addBtn = row.querySelector(".itemcode-add-btn");

    applyVisualStateToRow(row);
    includeSelect.addEventListener("change", () => applyVisualStateToRow(row));
    categorySelect.addEventListener("change", () => applyVisualStateToRow(row));
    itemInput.addEventListener("focus", () => showTypeaheadForInput(itemInput, row.dataset.nameKey));
    itemInput.addEventListener("input", () => showTypeaheadForInput(itemInput, row.dataset.nameKey));
    itemInput.addEventListener("keydown", (e) => {
      if (e.altKey && e.key === "ArrowDown") { e.preventDefault(); showTypeaheadForInput(itemInput, row.dataset.nameKey, true); return; }
      if (typeaheadRoot.hidden || state.typeahead.targetInput !== itemInput) return;
      if (e.key === "ArrowDown") { e.preventDefault(); moveTypeahead(1); return; }
      if (e.key === "ArrowUp") { e.preventDefault(); moveTypeahead(-1); return; }
      if (e.key === "Enter") {
        if (state.typeahead.activeIndex >= 0) { e.preventDefault(); commitTypeaheadSelection(state.typeahead.activeIndex); }
        else { const text = normalizeDisplayText(itemInput.value); if (text) { e.preventDefault(); commitManualItemCode(row.dataset.nameKey, itemInput, text); } }
        return;
      }
      if (e.key === "Escape") hideTypeahead();
    });
    itemInput.addEventListener("blur", () => { setTimeout(() => { if (document.activeElement && typeaheadRoot.contains(document.activeElement)) return; hideTypeahead(); }, 120); });
    addBtn.addEventListener("click", () => { const text = normalizeDisplayText(itemInput.value); if (!text) return; commitManualItemCode(row.dataset.nameKey, itemInput, text); });
  });
}

function saveMappingFromUI() {
  const rows = Array.from(mappingBody.querySelectorAll("tr[data-name-key]"));
  rows.forEach((row) => {
    const key = row.dataset.nameKey;
    const include = row.querySelector(".map-include")?.value || "Y";
    const category = row.querySelector(".map-category")?.value || "";
    const itemCode = normalizeDisplayText(row.querySelector(".itemcode-input")?.value || "");
    if (itemCode) ensureItemCodeOption(itemCode);
    state.mappingConfig[key] = { include, category, itemCode, note: suggestMappingByName(key).note || "" };
  });
}

function applyAutoSuggestionsToCurrentMapping() {
  for (const item of state.uniqueNames) {
    const suggested = suggestMappingByName(item.rawName);
    state.mappingConfig[item.normalizedName] = suggested;
    if (suggested.itemCode) ensureItemCodeOption(suggested.itemCode);
  }
  renderMappingTable();
}

/** -----------------------------
 * 타입어헤드 (자동완성)
 * ----------------------------- */
function getFilteredItemCodeOptions(keyword, forceAll = false) {
  const q = normalizeText(keyword);
  const list = sortUniqueStrings(state.itemCodeOptions);
  if (forceAll || !q) return list.slice(0, 200);
  const starts = []; const includes = [];
  for (const item of list) { const n = normalizeText(item); if (n.startsWith(q)) { starts.push(item); } else if (n.includes(q)) { includes.push(item); } }
  return [...starts, ...includes].slice(0, 200);
}

function positionTypeahead(input) {
  const rect = input.getBoundingClientRect();
  const width = Math.max(rect.width, 240);
  typeaheadRoot.style.left = `${rect.left + window.scrollX}px`;
  typeaheadRoot.style.top = `${rect.bottom + window.scrollY + 4}px`;
  typeaheadRoot.style.width = `${width}px`;
}

function renderTypeahead(items) {
  state.typeahead.items = items; state.typeahead.activeIndex = items.length ? 0 : -1;
  if (!items.length) { typeaheadRoot.innerHTML = `<div class="typeahead-empty">일치하는 항목이 없습니다. 직접 입력 후 추가할 수 있습니다.</div>`; return; }
  typeaheadRoot.innerHTML = `<div class="typeahead-list">${items.map((item, idx) => `<button type="button" class="typeahead-item ${idx === 0 ? "is-active" : ""}" data-index="${idx}">${escapeHtml(item)}</button>`).join("")}</div>`;
  Array.from(typeaheadRoot.querySelectorAll(".typeahead-item")).forEach((btn) => {
    btn.addEventListener("mousedown", (e) => e.preventDefault());
    btn.addEventListener("click", () => commitTypeaheadSelection(Number(btn.dataset.index)));
  });
}

function showTypeaheadForInput(input, nameKey, forceAll = false) {
  state.typeahead.targetInput = input; state.typeahead.targetKey = nameKey;
  const items = getFilteredItemCodeOptions(input.value, forceAll);
  positionTypeahead(input); renderTypeahead(items); typeaheadRoot.hidden = false;
}

function hideTypeahead() {
  typeaheadRoot.hidden = true; typeaheadRoot.innerHTML = "";
  state.typeahead.targetInput = null; state.typeahead.targetKey = ""; state.typeahead.items = []; state.typeahead.activeIndex = -1;
}

function refreshTypeaheadActive() {
  Array.from(typeaheadRoot.querySelectorAll(".typeahead-item")).forEach((el, idx) => { el.classList.toggle("is-active", idx === state.typeahead.activeIndex); });
}

function moveTypeahead(delta) {
  const len = state.typeahead.items.length; if (!len) return;
  const next = state.typeahead.activeIndex + delta;
  if (next < 0) { state.typeahead.activeIndex = len - 1; } else if (next >= len) { state.typeahead.activeIndex = 0; } else { state.typeahead.activeIndex = next; }
  refreshTypeaheadActive();
}

function commitTypeaheadSelection(index) {
  const value = state.typeahead.items[index]; if (!value || !state.typeahead.targetInput) return;
  state.typeahead.targetInput.value = value; ensureItemCodeOption(value); hideTypeahead(); state.typeahead.targetInput?.focus();
}

function commitManualItemCode(nameKey, input, text) {
  ensureItemCodeOption(text); input.value = text;
  const row = input.closest("tr[data-name-key]");
  const category = row.querySelector(".map-category")?.value || "";
  const include = row.querySelector(".map-include")?.value || "Y";
  state.mappingConfig[nameKey] = { include, category, itemCode: text, note: suggestMappingByName(nameKey).note || "" };
  hideTypeahead(); setStatus(`아이템구분 추가: ${text}`);
}

/** -----------------------------
 * 데이터 취합 및 계산
 * ----------------------------- */
function clearDataState() {
  state.rawEntriesByProject = { current: [], A: [], B: [], C: [] };
  state.fileSummaryByProject = { current: [], A: [], B: [], C: [] };
  state.uniqueNames = []; state.mappingConfig = {}; state.lastCompareRows = []; state.itemCodeOptions = [...DEFAULT_ITEM_CODE_OPTIONS];
  state.sections = []; state.activeTab = "";
}

async function extractNamesFromFiles() {
  clearDataState();
  const logs = []; let totalFileCount = 0;
  for (const projectKey of PROJECT_KEYS) {
    const files = Array.from(fileInputs[projectKey].files || []);
    totalFileCount += files.length;
    for (const file of files) {
      const parsed = await parseWorkbookFile(file, projectKey);
      state.rawEntriesByProject[projectKey].push(...parsed.entries);
      state.fileSummaryByProject[projectKey].push(parsed);
      logs.push(`[${PROJECT_LABELS[projectKey]}] 파일: ${parsed.fileName}`, `- 총 추출건수: ${parsed.entryCount}`, "");
    }
  }
  if (!totalFileCount) throw new Error("업로드된 파일이 없습니다.");
  
  // 파일 전체에서 유일한 층(Floor) 이름들을 긁어모아 탭 목록 생성
  const allSections = new Set();
  PROJECT_KEYS.forEach(pk => {
    state.rawEntriesByProject[pk].forEach(e => allSections.add(e.section));
  });
  state.sections = Array.from(allSections);
  if (state.sections.length > 0) state.activeTab = state.sections[0];
  
  buildUniqueNamesFromEntries(); ensureMappingConfig();
  renderTabs(); // 탭 그리기
  
  btnOpenMapping.disabled = state.uniqueNames.length === 0; btnCalc.disabled = state.uniqueNames.length === 0;
  setLog(logs.join("\n") || "로그가 없습니다."); setStatus(`명칭 추출 완료: ${state.uniqueNames.length}개`);
}

function getMappedEntriesByProject() {
  const result = { current: [], A: [], B: [], C: [] };
  for (const projectKey of PROJECT_KEYS) {
    result[projectKey] = state.rawEntriesByProject[projectKey].map((entry) => {
      const config = state.mappingConfig[entry.normalizedName];
      if (!config || config.include !== "Y" || !config.category || config.category === "제외" || !config.itemCode) return null;
      return { ...entry, mappedCategory: config.category, mappedItemCode: config.itemCode };
    }).filter(Boolean);
  }
  return result;
}

function buildProjectAggregate(mappedEntries) {
  const aggregate = { current: {}, A: {}, B: {}, C: {} };
  for (const projectKey of PROJECT_KEYS) {
    const bucket = {};
    for (const entry of mappedEntries[projectKey]) {
      const key = `${entry.section}__${entry.mappedCategory}__${entry.mappedItemCode}`;
      bucket[key] = (bucket[key] || 0) + entry.qty;
    }
    aggregate[projectKey] = bucket;
  }
  return aggregate;
}

function calcCompareRows() {
  const mappedEntries = getMappedEntriesByProject();
  const aggregate = buildProjectAggregate(mappedEntries);
  
  const allKeys = new Set();
  for (const projectKey of PROJECT_KEYS) { for (const k of Object.keys(aggregate[projectKey])) { allKeys.add(k); } }
  
  const dynamicItems = {}; const hardcodedKeys = new Set();
  for (const row of BASE_LAYOUT) { hardcodedKeys.add(`__${row.category}__${row.itemCode}`); }
  
  for (const k of allKeys) {
    const parts = k.split("__"); const sec = parts[0]; const cat = parts[1]; const ic = parts[2];
    if (!hardcodedKeys.has(`__${cat}__${ic}`)) {
      if (!dynamicItems[sec]) dynamicItems[sec] = {}; if (!dynamicItems[sec][cat]) dynamicItems[sec][cat] = [];
      dynamicItems[sec][cat].push(ic);
    }
  }

  const rows = []; 
  const categoryOrder = ["레미콘", "거푸집", "철근"]; 

  // 추출된 모든 층에 대하여 고정된 뼈대 레이아웃(레미콘, 거푸집, 철근)을 똑같이 복제하여 계산
  for (const sec of state.sections) {
    rows.push({ type: "section", section: sec });

    for (const cat of categoryOrder) {
      const hardcodedForCat = BASE_LAYOUT.filter(r => r.category === cat);
      for (const row of hardcodedForCat) {
        const key = `${sec}__${cat}__${row.itemCode}`;
        const cur = aggregate.current[key] || 0; const A = aggregate.A[key] || 0; const B = aggregate.B[key] || 0; const C = aggregate.C[key] || 0;
        const avg = (A + B + C) / 3; const ratio = cur === 0 ? 0 : avg / cur;
        rows.push({ ...row, section: sec, current: cur, A, B, C, avg, ratio, note: row.note || "" });
      }
      if (dynamicItems[sec] && dynamicItems[sec][cat]) {
        for (const ic of dynamicItems[sec][cat]) {
          const key = `${sec}__${cat}__${ic}`;
          const cur = aggregate.current[key] || 0; const A = aggregate.A[key] || 0; const B = aggregate.B[key] || 0; const C = aggregate.C[key] || 0;
          const avg = (A + B + C) / 3; const ratio = cur === 0 ? 0 : avg / cur;
          rows.push({ section: sec, itemCode: ic, item: cat, spec: ic, category: cat, current: cur, A, B, C, avg, ratio, note: "사용자 추가 항목" });
        }
      }
    }
    
    if (dynamicItems[sec]) {
      for (const [dynCat, dynIcs] of Object.entries(dynamicItems[sec])) {
        if (!categoryOrder.includes(dynCat)) {
          for (const ic of dynIcs) {
            const key = `${sec}__${dynCat}__${ic}`;
            const cur = aggregate.current[key] || 0; const A = aggregate.A[key] || 0; const B = aggregate.B[key] || 0; const C = aggregate.C[key] || 0;
            const avg = (A + B + C) / 3; const ratio = cur === 0 ? 0 : avg / cur;
            rows.push({ section: sec, itemCode: ic, item: dynCat, spec: ic, category: dynCat, current: cur, A, B, C, avg, ratio, note: "사용자 추가 항목" });
          }
        }
      }
    }
  }
  state.lastCompareRows = rows; return rows;
}

function renderCompareTable(rows) {
  if (!rows.length) {
    compareBody.innerHTML = `<tr><td colspan="11" class="empty-row">비교표가 아직 생성되지 않았습니다.</td></tr>`;
    return;
  }
  
  // 화면에는 현재 클릭된 탭의 내용만 보여줌
  const filteredRows = rows.filter(r => r.section === state.activeTab && r.type !== "section");
  if (!filteredRows.length) {
    compareBody.innerHTML = `<tr><td colspan="11" class="empty-row">[${state.activeTab}] 층에 해당하는 데이터가 없습니다.</td></tr>`;
    return;
  }
  
  const html = filteredRows.map((row) => {
    return `
      <tr>
        <td>${escapeHtml(row.section)}</td>
        <td>${escapeHtml(row.itemCode)}</td>
        <td>${escapeHtml(row.item)}</td>
        <td>${escapeHtml(row.spec)}</td>
        <td class="num">${fmtNumber(row.current)}</td>
        <td class="num">${fmtNumber(row.A)}</td>
        <td class="num">${fmtNumber(row.B)}</td>
        <td class="num">${fmtNumber(row.C)}</td>
        <td class="num">${fmtNumber(row.avg)}</td>
        <td class="num ratio ${ratioClass(row.ratio)}">${fmtRatio(row.ratio)}</td>
        <td>${escapeHtml(row.note || "")}</td>
      </tr>`;
  }).join("");
  compareBody.innerHTML = html;
}

function validateBeforeCalc() {
  if (!state.uniqueNames.length) throw new Error("먼저 명칭을 추출해야 합니다.");
  const missing = [];
  for (const item of state.uniqueNames) {
    const config = state.mappingConfig[item.normalizedName];
    if (!config) { missing.push(`${item.rawName} : 설정 없음`); continue; }
    if (config.include === "Y") {
      if (!config.category || config.category === "제외") missing.push(`${item.rawName} : 분류 미설정`);
      if (!config.itemCode) missing.push(`${item.rawName} : 아이템구분 미설정`);
    }
  }
  if (missing.length) throw new Error(`명칭 설정이 완료되지 않았습니다.\n\n${missing.slice(0, 30).join("\n")}${missing.length > 30 ? "\n..." : ""}`);
}


/** -----------------------------
 * 엑셀 다운로드 (비교표_갑지 양식 100% 동일 복제, 색상, 테두리, 행 높이 지원)
 * ----------------------------- */
function exportCompareExcel() {
  if (!state.lastCompareRows.length) {
    setStatus("내보낼 비교표가 없습니다.");
    return;
  }

  const ws = {};
  const merges = [];
  const rowsFormat = []; 
  
  const fontName = "맑은 고딕";
  const borderThin = { style: "thin", color: { rgb: "000000" } };
  const borderAll = { top: borderThin, bottom: borderThin, left: borderThin, right: borderThin };

  const titleStyle = { font: { name: fontName, bold: true, sz: 16, color: { rgb: "000000" } }, alignment: { horizontal: "center", vertical: "center" } };
  const headerStyle = { fill: { fgColor: { rgb: "F2F2F2" } }, font: { name: fontName, bold: true, sz: 10, color: { rgb: "000000" } }, alignment: { horizontal: "center", vertical: "center" }, border: borderAll };
  const sectionStyle = { fill: { fgColor: { rgb: "E2EFDA" } }, font: { name: fontName, bold: true, sz: 10, color: { rgb: "000000" } }, alignment: { horizontal: "center", vertical: "center" }, border: borderAll };
  const centerStyle = { font: { name: fontName, sz: 10, color: { rgb: "000000" } }, alignment: { horizontal: "center", vertical: "center" }, border: borderAll };
  const numberStyle = { font: { name: fontName, sz: 10, color: { rgb: "000000" } }, alignment: { horizontal: "right", vertical: "center" }, border: borderAll, numFmt: "#,##0" };
  const ratioStyle = { font: { name: fontName, sz: 10, color: { rgb: "000000" } }, alignment: { horizontal: "center", vertical: "center" }, border: borderAll, numFmt: "0%" };

  ws[XLSX.utils.encode_cell({ c: 0, r: 0 })] = { v: "ㅇㅇ 프로젝트 비교분석자료", t: "s", s: titleStyle };
  merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 9 } }); 
  rowsFormat[0] = { hpt: 40 }; 

  const headers = ["코드", "품명", "규격", "현재 프로젝트", "'A' 프로젝트", "'B' 프로젝트", "'C' 프로젝트", "평균치(A~C프로젝트)", "비율", "비고"];
  for (let c = 0; c < headers.length; c++) {
    ws[XLSX.utils.encode_cell({ c: c, r: 1 })] = { v: headers[c], t: "s", s: headerStyle };
  }
  rowsFormat[1] = { hpt: 25 };

  let r = 2; 
  state.lastCompareRows.forEach((row) => {
    if (row.type === "section") {
      // 층 간 띄어쓰기 공백행 (테두리 없음)
      for (let c = 0; c < 10; c++) {
        ws[XLSX.utils.encode_cell({ c: c, r: r })] = { v: "", t: "s" };
      }
      rowsFormat[r] = { hpt: 12 };
      r++;

      // 층 타이틀행 (B열에 이름, 연두색 배경)
      for (let c = 0; c < 10; c++) {
        let val = c === 1 ? row.section : "";
        ws[XLSX.utils.encode_cell({ c: c, r: r })] = { v: val, t: "s", s: sectionStyle };
      }
      rowsFormat[r] = { hpt: 22 };
      r++;
    } else {
      const ratioVal = Number.isFinite(row.ratio) ? Number(row.ratio) : 0;
      const rowData = [
        { v: row.itemCode, t: "s", s: centerStyle },
        { v: row.item, t: "s", s: centerStyle },
        { v: row.spec, t: "s", s: centerStyle },
        { v: Math.round(row.current), t: "n", s: numberStyle },
        { v: Math.round(row.A), t: "n", s: numberStyle },
        { v: Math.round(row.B), t: "n", s: numberStyle },
        { v: Math.round(row.C), t: "n", s: numberStyle },
        { v: Math.round(row.avg), t: "n", s: numberStyle },
        { v: ratioVal, t: "n", s: ratioStyle }, 
        { v: row.note || "", t: "s", s: centerStyle },
      ];
      for (let c = 0; c < rowData.length; c++) {
        ws[XLSX.utils.encode_cell({ c: c, r: r })] = rowData[c];
      }
      rowsFormat[r] = { hpt: 20 };
      r++;
    }
  });

  ws['!ref'] = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: 9, r: r - 1 } });
  ws['!merges'] = merges;
  ws['!rows'] = rowsFormat; 
  ws['!cols'] = [ 
    { wch: 10 }, { wch: 12 }, { wch: 16 }, { wch: 15 }, 
    { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 20 }, 
    { wch: 10 }, { wch: 15 }
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "비교표");
  XLSX.writeFile(wb, "비교표_결과.xlsx");
}

/** -----------------------------
 * 기타 이벤트 제어
 * ----------------------------- */
function resetAll() {
  for (const key of PROJECT_KEYS) fileInputs[key].value = "";
  updateFileListText(); clearDataState(); renderTabs();
  btnOpenMapping.disabled = true; btnCalc.disabled = true; btnExportCsv.disabled = true;
  compareBody.innerHTML = `<tr><td colspan="11" class="empty-row">비교표가 아직 생성되지 않았습니다.</td></tr>`;
  mappingBody.innerHTML = ""; setStatus("대기 중"); setLog("로그가 여기에 표시됩니다."); closeModal();
}

for (const key of PROJECT_KEYS) fileInputs[key].addEventListener("change", updateFileListText);

btnExportCsv.textContent = "엑셀 내보내기";
btnExtract.addEventListener("click", async () => { try { setStatus("명칭 추출 중..."); setLog("파일을 분석하고 있습니다..."); await extractNamesFromFiles(); renderMappingTable(); openModal(); } catch (error) { console.error(error); setStatus("오류 발생"); setLog(error?.message || String(error)); } });
btnOpenMapping.addEventListener("click", () => { renderMappingTable(); openModal(); });
btnCloseMapping.addEventListener("click", closeModal); mappingBackdrop.addEventListener("click", closeModal);
btnApplySuggestions.addEventListener("click", () => { applyAutoSuggestionsToCurrentMapping(); });
btnSaveMapping.addEventListener("click", () => { saveMappingFromUI(); closeModal(); setStatus("명칭 설정 저장 완료"); });
btnCalc.addEventListener("click", () => { try { if (mappingModal.classList.contains("is-open")) saveMappingFromUI(); validateBeforeCalc(); const rows = calcCompareRows(); renderCompareTable(rows); btnExportCsv.disabled = rows.length === 0; setStatus("비교표 생성 완료"); } catch (error) { console.error(error); setStatus("오류 발생"); setLog(error?.message || String(error)); } });
btnExportCsv.addEventListener("click", exportCompareExcel);
btnReset.addEventListener("click", resetAll);

document.addEventListener("click", (e) => { if (!typeaheadRoot.hidden) { const clickedInsideTypeahead = typeaheadRoot.contains(e.target); const clickedInput = e.target.closest(".itemcode-input"); if (!clickedInsideTypeahead && !clickedInput) hideTypeahead(); } });
window.addEventListener("resize", () => { if (!typeaheadRoot.hidden && state.typeahead.targetInput) positionTypeahead(state.typeahead.targetInput); });
window.addEventListener("scroll", () => { if (!typeaheadRoot.hidden && state.typeahead.targetInput) positionTypeahead(state.typeahead.targetInput); }, true);

updateFileListText(); renderTabs(); setStatus("대기 중"); setLog("로그가 여기에 표시됩니다.");
