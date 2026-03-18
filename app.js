"use strict";

const $ = (id) => document.getElementById(id);

const fileInputs = { current: $("file-current"), A: $("file-a"), B: $("file-b"), C: $("file-c") };
const fileListEls = { current: $("file-current-list"), A: $("file-a-list"), B: $("file-b-list"), C: $("file-c-list") };

const theadEl = $("compare-thead");
const tbodyEl = $("compare-body");
const statusBox = $("status-box");
const logBox = $("log-box");

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

// 💡 1. 카테고리(중분류) 옵션 업데이트: 콘크리트, 거푸집, 철근, 잡
const CATEGORY_OPTIONS = [
  { value: "", label: "선택" }, 
  { value: "콘크리트", label: "콘크리트" },
  { value: "거푸집", label: "거푸집" }, 
  { value: "철근", label: "철근" }, 
  { value: "잡", label: "잡" }, 
  { value: "제외", label: "제외" }
];
const INCLUDE_OPTIONS = [ { value: "Y", label: "반영" }, { value: "N", label: "제외" } ];
const DEFAULT_ITEM_CODE_OPTIONS = ["240", "270", "300", "180", "3회", "4회", "유로", "알폼", "갱폼", "합벽", "보밑면", "데크", "방수턱", "H10", "H13", "H16", "H19", "H22", "H25", "H29"];

const state = {
  rawEntriesByProject: { current: [], A: [], B: [], C: [] },
  uniqueNames: [], mappingConfig: {}, itemCodeOptions: [...DEFAULT_ITEM_CODE_OPTIONS],
  typeahead: { targetInput: null, targetKey: "", items: [], activeIndex: -1 },
  lastData: null
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
  return Number.isFinite(Number(cleaned)) ? Number(cleaned) : 0;
}
function fmtNumber(value) { return Number(value || 0).toLocaleString("ko-KR", { maximumFractionDigits: 0 }); }
function fmtIndex(value) { return (!value || !isFinite(value)) ? "-" : Number(value).toFixed(2); }
function indexClass(value) {
  if (!value) return "";
  if (value > 1.1) return "bad"; 
  if (value < 0.9) return "good"; 
  return "warn"; 
}

function updateFileListText() {
  for (const key of PROJECT_KEYS) {
    const files = Array.from(fileInputs[key].files || []);
    fileListEls[key].textContent = files.length ? files.map(f => f.name).join("\n") : "선택된 파일 없음";
  }
}
function openModal() { mappingModal.classList.add("is-open"); mappingModal.setAttribute("aria-hidden", "false"); }
function closeModal() { mappingModal.classList.remove("is-open"); mappingModal.setAttribute("aria-hidden", "true"); hideTypeahead(); }
function sortUniqueStrings(list) { return [...new Set(list.filter(Boolean))].sort((a, b) => a.localeCompare(b, "ko")); }
function ensureItemCodeOption(value) {
  const text = normalizeDisplayText(value);
  if (text && !state.itemCodeOptions.includes(text)) {
    state.itemCodeOptions.push(text); state.itemCodeOptions = sortUniqueStrings(state.itemCodeOptions);
  }
}

function parseFloor(f) {
  if (!f) return { type: 'unknown', num: 0, raw: f, label: f };
  const m = f.match(/\d+/);
  const num = m ? parseInt(m[0], 10) : 0;
  if (f.includes("지하") || f.toUpperCase().includes("B")) return { type: 'B', num: num, raw: f, label: `B${num}F` };
  if (f.includes("지상") || f.includes("층") || f.toUpperCase().includes("F")) return { type: 'F', num: num, raw: f, label: `${num}F` };
  return { type: 'unknown', num: num, raw: f, label: f };
}
function sortFloors(floors) {
  return floors.map(parseFloor).sort((a, b) => {
    if (a.type === 'B' && b.type === 'B') return b.num - a.num; 
    if (a.type === 'B' && b.type !== 'B') return -1;
    if (a.type !== 'B' && b.type === 'B') return 1;
    if (a.type === 'F' && b.type === 'F') return a.num - b.num; 
    return 0;
  });
}

function parseSheetEntries(rows, projectKey, fileName, sheetName) {
  let buildingName = normalizeDisplayText(sheetName);
  for(let r=0; r<4; r++) {
    if(!rows[r]) continue;
    for(let c=0; c<10; c++) {
      let val = normalizeDisplayText(rows[r][c]);
      if(val && val.endsWith("동")) { buildingName = val; break; }
    }
  }

  const headerRow3 = rows[2] || []; 
  const headerRow4 = rows[3] || []; 
  const colStart = 2; 
  const colEnd = Math.max(headerRow3.length, headerRow4.length) - 1;
  const entries = [];
  const skipHeaders = ["합계", "비고", "누계", "소계", ""];

  for(let r=4; r<rows.length; r++) {
    let row1 = rows[r] || [];
    let col0 = normalizeDisplayText(row1[0]);
    let col1 = normalizeDisplayText(row1[1]);
    let possibleFloor = col0 || col1;

    if(possibleFloor) {
      let currentFloor = possibleFloor;
      if(currentFloor === "계" || currentFloor.includes("합계")) currentFloor = "합계";

      for(let c=colStart; c<=colEnd; c++) {
        let name3 = normalizeDisplayText(headerRow3[c]);
        if(name3 && !skipHeaders.includes(name3)) {
          let val = toNumber(row1[c]);
          if(val !== 0) entries.push({ projectKey, building: buildingName, floor: currentFloor, rawName: name3, normalizedName: normalizeText(name3), qty: val });
        }
      }

      if(r+1 < rows.length) {
        let row2 = rows[r+1] || [];
        let nextCol0 = normalizeDisplayText(row2[0]);
        let nextCol1 = normalizeDisplayText(row2[1]);
        if(!nextCol0 && !nextCol1) {
          for(let c=colStart; c<=colEnd; c++) {
            let name4 = normalizeDisplayText(headerRow4[c]);
            if(name4 && !skipHeaders.includes(name4)) {
              let val = toNumber(row2[c]);
              if(val !== 0) entries.push({ projectKey, building: buildingName, floor: currentFloor, rawName: name4, normalizedName: normalizeText(name4), qty: val });
            }
          }
          r++; 
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
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
    const entries = parseSheetEntries(rows, projectKey, file.name, sheetName);
    allEntries.push(...entries);
    totalEntryCount += entries.length;
  }
  return { projectKey, fileName: file.name, entryCount: totalEntryCount, entries: allEntries };
}

// 💡 2. 매핑 자동제안 시 "레미콘"을 "콘크리트"로 제안하도록 변경
function suggestMappingByName(rawName) {
  const t = normalizeText(rawName);
  const result = { include: "Y", category: "", itemCode: "", note: "" };

  if (t.includes("24MPA")) { result.category = "콘크리트"; result.itemCode = "240"; result.note = "24MPA → 240"; return result; }
  if (t.includes("27MPA")) { result.category = "콘크리트"; result.itemCode = "270"; result.note = "27MPA → 270"; return result; }
  if (t.includes("30MPA")) { result.category = "콘크리트"; result.itemCode = "300"; result.note = "30MPA → 300"; return result; }
  if (t.includes("현치/무근") || t.includes("버림")) { result.category = "콘크리트"; result.itemCode = "180"; result.note = "버림/현치무근 → 180"; return result; }
  if (t.includes("현치/구체") || t === "합벽채움" || t.includes("ACT-INSIDE")) { result.category = "콘크리트"; result.itemCode = "240"; result.note = "구체/합벽/ACT-INSIDE 계열 → 240"; return result; }
  
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

  // 자동 제안에 매칭되지 않으면 수동으로 지정할 수 있도록 비워둠
  result.include = "N"; result.category = "제외"; result.itemCode = ""; result.note = "자동 제안 없음";
  return result;
}

function buildUniqueNamesFromEntries() {
  const nameMap = new Map();
  for (const pk of PROJECT_KEYS) {
    for (const entry of state.rawEntriesByProject[pk]) {
      if (!nameMap.has(entry.normalizedName)) nameMap.set(entry.normalizedName, { rawName: entry.rawName, normalizedName: entry.normalizedName });
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
  return list.map((opt) => `<option value="${escapeHtml(opt.value)}" ${String(opt.value) === String(selectedValue) ? "selected" : ""}>${escapeHtml(opt.label)}</option>`).join("");
}

// 💡 3. 테마 컬러 매핑 (콘크리트, 거푸집, 철근에 맞춰 색상 지정)
function getCategoryClass(category) {
  if (category === "콘크리트") return "is-concrete"; // 기존 레미콘 스타일(붉은색 톤) 그대로 사용
  if (category === "거푸집") return "is-form";
  if (category === "철근") return "is-rebar";
  if (category === "잡") return "is-exclude"; // 잡 항목은 회색 톤으로 적용
  return "is-exclude";
}

function getRowIncludeClass(include) { return include === "Y" ? "is-included" : "is-excluded"; }

function renderMappingTable() {
  if (!state.uniqueNames.length) { mappingBody.innerHTML = `<tr><td colspan="5">추출된 명칭이 없습니다.</td></tr>`; return; }
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
  const includeSelect = row.querySelector(".map-include"); const categorySelect = row.querySelector(".map-category");
  const include = includeSelect?.value || "N"; const category = categorySelect?.value || "";
  row.classList.remove("is-included", "is-excluded"); row.classList.add(include === "Y" ? "is-included" : "is-excluded");
  categorySelect.classList.remove("is-concrete", "is-form", "is-rebar", "is-exclude"); categorySelect.classList.add(getCategoryClass(category));
}

function bindMappingRowBehaviors() {
  const rows = Array.from(mappingBody.querySelectorAll("tr[data-name-key]"));
  rows.forEach((row) => {
    const includeSelect = row.querySelector(".map-include"); const categorySelect = row.querySelector(".map-category");
    const itemInput = row.querySelector(".itemcode-input"); const addBtn = row.querySelector(".itemcode-add-btn");
    applyVisualStateToRow(row);
    includeSelect.addEventListener("change", () => applyVisualStateToRow(row)); categorySelect.addEventListener("change", () => applyVisualStateToRow(row));
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
    const key = row.dataset.nameKey; const include = row.querySelector(".map-include")?.value || "Y";
    const category = row.querySelector(".map-category")?.value || ""; const itemCode = normalizeDisplayText(row.querySelector(".itemcode-input")?.value || "");
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

function getFilteredItemCodeOptions(keyword, forceAll = false) {
  const q = normalizeText(keyword); const list = sortUniqueStrings(state.itemCodeOptions);
  if (forceAll || !q) return list.slice(0, 200);
  const starts = []; const includes = [];
  for (const item of list) { const n = normalizeText(item); if (n.startsWith(q)) starts.push(item); else if (n.includes(q)) includes.push(item); }
  return [...starts, ...includes].slice(0, 200);
}
function positionTypeahead(input) { const rect = input.getBoundingClientRect(); const width = Math.max(rect.width, 240); typeaheadRoot.style.left = `${rect.left + window.scrollX}px`; typeaheadRoot.style.top = `${rect.bottom + window.scrollY + 4}px`; typeaheadRoot.style.width = `${width}px`; }
function renderTypeahead(items) {
  state.typeahead.items = items; state.typeahead.activeIndex = items.length ? 0 : -1;
  if (!items.length) { typeaheadRoot.innerHTML = `<div class="typeahead-empty">일치하는 항목이 없습니다.</div>`; return; }
  typeaheadRoot.innerHTML = `<div class="typeahead-list">${items.map((item, idx) => `<button type="button" class="typeahead-item ${idx === 0 ? "is-active" : ""}" data-index="${idx}">${escapeHtml(item)}</button>`).join("")}</div>`;
  Array.from(typeaheadRoot.querySelectorAll(".typeahead-item")).forEach((btn) => { btn.addEventListener("mousedown", (e) => e.preventDefault()); btn.addEventListener("click", () => commitTypeaheadSelection(Number(btn.dataset.index))); });
}
function showTypeaheadForInput(input, nameKey, forceAll = false) { state.typeahead.targetInput = input; state.typeahead.targetKey = nameKey; const items = getFilteredItemCodeOptions(input.value, forceAll); positionTypeahead(input); renderTypeahead(items); typeaheadRoot.hidden = false; }
function hideTypeahead() { typeaheadRoot.hidden = true; typeaheadRoot.innerHTML = ""; state.typeahead.targetInput = null; state.typeahead.targetKey = ""; state.typeahead.items = []; state.typeahead.activeIndex = -1; }
function refreshTypeaheadActive() { Array.from(typeaheadRoot.querySelectorAll(".typeahead-item")).forEach((el, idx) => { el.classList.toggle("is-active", idx === state.typeahead.activeIndex); }); }
function moveTypeahead(delta) {
  const len = state.typeahead.items.length; if (!len) return; const next = state.typeahead.activeIndex + delta;
  if (next < 0) state.typeahead.activeIndex = len - 1; else if (next >= len) state.typeahead.activeIndex = 0; else state.typeahead.activeIndex = next;
  refreshTypeaheadActive();
}
function commitTypeaheadSelection(index) {
  const value = state.typeahead.items[index]; if (!value || !state.typeahead.targetInput) return;
  state.typeahead.targetInput.value = value; ensureItemCodeOption(value); hideTypeahead(); state.typeahead.targetInput?.focus();
}
function commitManualItemCode(nameKey, input, text) {
  ensureItemCodeOption(text); input.value = text;
  const row = input.closest("tr[data-name-key]"); const category = row.querySelector(".map-category")?.value || ""; const include = row.querySelector(".map-include")?.value || "Y";
  state.mappingConfig[nameKey] = { include, category, itemCode: text, note: suggestMappingByName(nameKey).note || "" };
  hideTypeahead(); setStatus(`아이템구분 추가: ${text}`);
}

function clearDataState() {
  state.rawEntriesByProject = { current: [], A: [], B: [], C: [] };
  state.uniqueNames = []; state.mappingConfig = {}; state.itemCodeOptions = [...DEFAULT_ITEM_CODE_OPTIONS];
  state.lastData = null;
}

async function extractNamesFromFiles() {
  clearDataState();
  const logs = []; let totalFileCount = 0;
  for (const pk of PROJECT_KEYS) {
    const files = Array.from(fileInputs[pk].files || []);
    totalFileCount += files.length;
    for (const file of files) {
      const parsed = await parseWorkbookFile(file, pk);
      state.rawEntriesByProject[pk].push(...parsed.entries);
      logs.push(`[${PROJECT_LABELS[pk]}] 파일: ${file.name}`, `- 총 추출건수: ${parsed.entryCount}`, "");
    }
  }
  if (!totalFileCount) throw new Error("업로드된 파일이 없습니다.");
  
  buildUniqueNamesFromEntries(); ensureMappingConfig();
  
  btnOpenMapping.disabled = state.uniqueNames.length === 0; btnCalc.disabled = state.uniqueNames.length === 0;
  setLog(logs.join("\n") || "로그가 없습니다."); setStatus(`명칭 추출 완료: ${state.uniqueNames.length}개`);
}

function getMappedEntriesByProject() {
  const result = { current: [], A: [], B: [], C: [] };
  for (const pk of PROJECT_KEYS) {
    result[pk] = state.rawEntriesByProject[pk].map((entry) => {
      const config = state.mappingConfig[entry.normalizedName];
      if (!config || config.include !== "Y" || !config.category || config.category === "제외" || !config.itemCode) return null;
      return { ...entry, mappedCategory: config.category }; 
    }).filter(Boolean);
  }
  return result;
}

// 💡 4. 4가지 중분류 기준(콘크리트, 거푸집, 철근, 잡)으로 QS 지표 연산
function calcCompareRows() {
  const mappedEntries = getMappedEntriesByProject();
  const currentEntries = mappedEntries.current.filter(e => e.floor && e.floor !== "합계");
  const rawFloors = new Set(currentEntries.map(e => e.floor));
  const parsedFloors = sortFloors(Array.from(rawFloors)); 
  
  const abcAvgs = { "콘크리트": 0, "거푸집": 0, "철근": 0, "잡": 0 };
  for (const cat of Object.keys(abcAvgs)) {
    let sumAvg = 0; let projCount = 0;
    for (const pk of ["A", "B", "C"]) {
      const entries = mappedEntries[pk].filter(e => e.mappedCategory === cat && e.floor !== "합계");
      if (entries.length === 0) continue;
      
      const uniqueFloors = new Set();
      let totalQty = 0;
      for (const e of entries) {
         uniqueFloors.add(`${e.building}__${e.floor}`); 
         totalQty += e.qty;
      }
      if (uniqueFloors.size > 0) {
         sumAvg += (totalQty / uniqueFloors.size);
         projCount++;
      }
    }
    abcAvgs[cat] = projCount > 0 ? sumAvg / projCount : 0;
  }

  const buildings = Array.from(new Set(currentEntries.map(e => e.building))).sort();
  const categories = ["콘크리트", "거푸집", "철근", "잡"]; // 정렬 순서 강제 고정
  const rows = [];

  for (const bldg of buildings) {
    for (const cat of categories) {
      const row = {
        building: bldg,
        category: cat,
        floors: {},
        currentAvg: 0,
        abcAvg: abcAvgs[cat],
        index: 0
      };

      let bldgTotalQty = 0;

      for (const pf of parsedFloors) {
        const qty = currentEntries
          .filter(e => e.building === bldg && e.floor === pf.raw && e.mappedCategory === cat)
          .reduce((sum, e) => sum + e.qty, 0);
        
        row.floors[pf.label] = qty;
        
        const floorExistsInBldg = currentEntries.some(e => e.building === bldg && e.floor === pf.raw);
        if (floorExistsInBldg) {
          bldgTotalQty += qty;
        }
      }
      
      const uniqueFloorsForBldg = new Set(currentEntries.filter(e => e.building === bldg).map(e => e.floor));
      row.currentAvg = uniqueFloorsForBldg.size > 0 ? (bldgTotalQty / uniqueFloorsForBldg.size) : 0;
      row.index = row.abcAvg > 0 ? row.currentAvg / row.abcAvg : 0;
      
      // 해당 동/카테고리에 수량이 1이라도 존재할 때만 표에 표시 (빈 줄 방지용, 원치 않으면 삭제 가능)
      if(bldgTotalQty > 0) rows.push(row);
    }
  }

  const data = { rows, parsedFloors };
  state.lastData = data;
  return data;
}

function renderCompareTable(data) {
  const { rows, parsedFloors } = data;
  if (!rows.length) {
    tbodyEl.innerHTML = `<tr><td colspan="${4 + parsedFloors.length}" class="empty-row">비교표가 아직 생성되지 않았습니다.</td></tr>`;
    return;
  }

  let thHtml = `<tr><th>동</th><th>중분류(아이템)</th>`;
  for (const pf of parsedFloors) thHtml += `<th>${escapeHtml(pf.label)}</th>`;
  thHtml += `<th>층평균</th><th>3개프로젝트평균</th><th>Index(현재/평균)</th></tr>`;
  theadEl.innerHTML = thHtml;

  const html = rows.map(row => {
    let tdHtml = `<tr>
      <td style="font-weight:bold;">${escapeHtml(row.building)}</td>
      <td style="font-weight:bold; color:#1d4ed8;">${escapeHtml(row.category)}</td>`;
    
    for (const pf of parsedFloors) {
      const val = row.floors[pf.label];
      tdHtml += `<td class="num">${val ? fmtNumber(val) : "-"}</td>`;
    }

    tdHtml += `
      <td class="num" style="font-weight:bold;">${fmtNumber(row.currentAvg)}</td>
      <td class="num" style="font-weight:bold;">${fmtNumber(row.abcAvg)}</td>
      <td class="num ratio ${indexClass(row.index)}">${fmtIndex(row.index)}</td>
    </tr>`;
    return tdHtml;
  }).join("");

  tbodyEl.innerHTML = html;
}

function validateBeforeCalc() {
  if (!state.uniqueNames.length) throw new Error("먼저 명칭을 추출해야 합니다.");
  const missing = [];
  for (const item of state.uniqueNames) {
    const config = state.mappingConfig[item.normalizedName];
    if (!config) { missing.push(`${item.rawName} : 설정 없음`); continue; }
    if (config.include === "Y") {
      if (!config.category || config.category === "제외") missing.push(`${item.rawName} : 중분류 미설정`);
    }
  }
  if (missing.length) throw new Error(`명칭 설정이 완료되지 않았습니다.\n\n${missing.slice(0, 30).join("\n")}${missing.length > 30 ? "\n..." : ""}`);
}

/** -----------------------------
 * 엑셀 다운로드 (QS 벤치마크 매트릭스 양식)
 * ----------------------------- */
function exportCompareExcel() {
  if (!state.lastData || !state.lastData.rows.length) { setStatus("내보낼 비교표가 없습니다."); return; }

  const { rows, parsedFloors } = state.lastData;
  const ws = {}; const merges = []; const rowsFormat = []; 
  
  const fontName = "맑은 고딕";
  const borderThin = { style: "thin", color: { rgb: "000000" } };
  const borderAll = { top: borderThin, bottom: borderThin, left: borderThin, right: borderThin };

  const headerStyle = { fill: { fgColor: { rgb: "D9D9D9" } }, font: { name: fontName, bold: true, sz: 10, color: { rgb: "000000" } }, alignment: { horizontal: "center", vertical: "center" }, border: borderAll };
  const centerStyle = { font: { name: fontName, sz: 10, color: { rgb: "000000" } }, alignment: { horizontal: "center", vertical: "center" }, border: borderAll };
  const numberStyle = { font: { name: fontName, sz: 10, color: { rgb: "000000" } }, alignment: { horizontal: "right", vertical: "center" }, border: borderAll, numFmt: "#,##0" };
  const ratioStyle  = { font: { name: fontName, sz: 10, color: { rgb: "000000" } }, alignment: { horizontal: "center", vertical: "center" }, border: borderAll, numFmt: "0.00" };

  const headers = ["동", "아이템", ...parsedFloors.map(pf => pf.label), "층평균", "3개프로젝트평균", "Index(현재/평균)"];
  for (let c = 0; c < headers.length; c++) ws[XLSX.utils.encode_cell({ c: c, r: 0 })] = { v: headers[c], t: "s", s: headerStyle };
  rowsFormat[0] = { hpt: 25 };

  let r = 1; 
  for (const row of rows) {
    let c = 0;
    ws[XLSX.utils.encode_cell({ c: c++, r: r })] = { v: row.building, t: "s", s: centerStyle };
    ws[XLSX.utils.encode_cell({ c: c++, r: r })] = { v: row.category, t: "s", s: centerStyle };
    
    for (const pf of parsedFloors) {
      const val = row.floors[pf.label] || 0;
      ws[XLSX.utils.encode_cell({ c: c++, r: r })] = { v: val === 0 ? "" : Math.round(val), t: val === 0 ? "s" : "n", s: numberStyle };
    }

    ws[XLSX.utils.encode_cell({ c: c++, r: r })] = { v: Math.round(row.currentAvg), t: "n", s: numberStyle };
    ws[XLSX.utils.encode_cell({ c: c++, r: r })] = { v: Math.round(row.abcAvg), t: "n", s: numberStyle };
    ws[XLSX.utils.encode_cell({ c: c++, r: r })] = { v: row.index, t: "n", s: ratioStyle };
    
    rowsFormat[r] = { hpt: 20 }; r++;
  }

  ws['!ref'] = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: headers.length - 1, r: r - 1 } });
  ws['!rows'] = rowsFormat; 
  
  const cols = [{ wch: 12 }, { wch: 10 }];
  for(let i=0; i<parsedFloors.length; i++) cols.push({ wch: 10 }); 
  cols.push({ wch: 14 }, { wch: 18 }, { wch: 16 }); 
  ws['!cols'] = cols;

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "QS_Comparison");
  XLSX.writeFile(wb, "QS_Benchmark_결과.xlsx");
}

function resetAll() {
  for (const pk of PROJECT_KEYS) fileInputs[pk].value = "";
  updateFileListText(); clearDataState();
  btnOpenMapping.disabled = true; btnCalc.disabled = true; btnExportCsv.disabled = true;
  tbodyEl.innerHTML = `<tr><td colspan="5" class="empty-row">비교표가 아직 생성되지 않았습니다.</td></tr>`;
  theadEl.innerHTML = `<tr><th>동</th><th>아이템</th><th>층평균</th><th>3개프로젝트평균</th><th>Index(현재/평균)</th></tr>`;
  mappingBody.innerHTML = ""; setStatus("대기 중"); setLog("로그가 여기에 표시됩니다."); closeModal();
}

for (const pk of PROJECT_KEYS) fileInputs[pk].addEventListener("change", updateFileListText);

btnExportCsv.textContent = "엑셀 내보내기";
btnExtract.addEventListener("click", async () => { try { setStatus("명칭 추출 중..."); setLog("파일을 분석하고 있습니다..."); await extractNamesFromFiles(); renderMappingTable(); openModal(); } catch (error) { console.error(error); setStatus("오류 발생"); setLog(error?.message || String(error)); } });
btnOpenMapping.addEventListener("click", () => { renderMappingTable(); openModal(); });
btnCloseMapping.addEventListener("click", closeModal); mappingBackdrop.addEventListener("click", closeModal);
btnApplySuggestions.addEventListener("click", () => { applyAutoSuggestionsToCurrentMapping(); });
btnSaveMapping.addEventListener("click", () => { saveMappingFromUI(); closeModal(); setStatus("명칭 설정 저장 완료"); });
btnCalc.addEventListener("click", () => { try { if (mappingModal.classList.contains("is-open")) saveMappingFromUI(); validateBeforeCalc(); const data = calcCompareRows(); renderCompareTable(data); btnExportCsv.disabled = data.rows.length === 0; setStatus("비교표 생성 완료"); } catch (error) { console.error(error); setStatus("오류 발생"); setLog(error?.message || String(error)); } });
btnExportCsv.addEventListener("click", exportCompareExcel);
btnReset.addEventListener("click", resetAll);

document.addEventListener("click", (e) => { if (!typeaheadRoot.hidden) { const clickedInsideTypeahead = typeaheadRoot.contains(e.target); const clickedInput = e.target.closest(".itemcode-input"); if (!clickedInsideTypeahead && !clickedInput) hideTypeahead(); } });
window.addEventListener("resize", () => { if (!typeaheadRoot.hidden && state.typeahead.targetInput) positionTypeahead(state.typeahead.targetInput); });
window.addEventListener("scroll", () => { if (!typeaheadRoot.hidden && state.typeahead.targetInput) positionTypeahead(state.typeahead.targetInput); }, true);

updateFileListText(); setStatus("대기 중"); setLog("로그가 여기에 표시됩니다.");
