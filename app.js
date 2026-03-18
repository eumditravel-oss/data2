"use strict";

const $ = (id) => document.getElementById(id);

const fileInputs = { current: $("file-current"), A: $("file-a"), B: $("file-b"), C: $("file-c") };
const fileListEls = { current: $("file-current-list"), A: $("file-a-list"), B: $("file-b-list"), C: $("file-c-list") };

const compareBody = $("compare-body");
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

// UI Controls
const controlsBar = $("controls-bar");
const selectBuilding = $("select-building");
const selectFloor = $("select-floor");
const wrapFloor = $("wrap-floor");
const sumOptions = $("sum-options");
const sumCheckboxes = $("sum-checkboxes");

const PROJECT_KEYS = ["current", "A", "B", "C"];
const PROJECT_LABELS = { current: "현재 프로젝트", A: "A 프로젝트", B: "B 프로젝트", C: "C 프로젝트" };

const CATEGORY_OPTIONS = [
  { value: "", label: "선택" }, { value: "레미콘", label: "레미콘" },
  { value: "거푸집", label: "거푸집" }, { value: "철근", label: "철근" }, { value: "제외", label: "제외" }
];
const INCLUDE_OPTIONS = [ { value: "Y", label: "반영" }, { value: "N", label: "제외" } ];
const DEFAULT_ITEM_CODE_OPTIONS = ["240", "270", "300", "180", "3회", "4회", "유로", "알폼", "갱폼", "합벽", "보밑면", "데크", "방수턱", "H10", "H13", "H16", "H19", "H22", "H25", "H29"];

// 뼈대 레이아웃 (동/층은 동적으로 주입됨)
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
  { itemCode: "H25", item: "철근", spec: "H25", category: "철근" },
  { itemCode: "H29", item: "철근", spec: "H29", category: "철근" },
];

const state = {
  rawEntriesByProject: { current: [], A: [], B: [], C: [] },
  uniqueNames: [], mappingConfig: {}, itemCodeOptions: [...DEFAULT_ITEM_CODE_OPTIONS],
  typeahead: { targetInput: null, targetKey: "", items: [], activeIndex: -1 },
  buildings: [], floorsByBuilding: {}, 
  activeBuilding: "", activeFloor: "", summationChecked: new Set()
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
  const n = Number(cleaned); return Number.isFinite(n) ? n : 0;
}
function fmtNumber(value) { return Number(value || 0).toLocaleString("ko-KR", { maximumFractionDigits: 0 }); }
function fmtRatio(value) { return (!value && value !== 0) ? "0%" : (Number(value) * 100).toFixed(0) + "%"; }
function ratioClass(value) { if (!value) return ""; if (value < 0.9) return "bad"; if (value <= 1.1) return "good"; return "warn"; }
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
  if (!state.itemCodeOptions.includes(text)) { state.itemCodeOptions.push(text); state.itemCodeOptions = sortUniqueStrings(state.itemCodeOptions); }
}

/** -----------------------------
 * 핵심 로직: 엑셀 파싱 (2행 묶음 매칭 및 동 추출)
 * ----------------------------- */
function parseSheetEntries(rows, projectKey, fileName, sheetName) {
  // 1. 0~3행 안에서 "동"으로 끝나는 텍스트를 찾아 건물 이름으로 지정 (없으면 시트명 사용)
  let buildingName = normalizeDisplayText(sheetName);
  for(let r=0; r<4; r++) {
    if(!rows[r]) continue;
    for(let c=0; c<10; c++) {
      let val = normalizeDisplayText(rows[r][c]);
      if(val && val.endsWith("동")) { buildingName = val; break; }
    }
  }

  const headerRow3 = rows[2] || []; // 노란색 타이틀
  const headerRow4 = rows[3] || []; // 주황색 타이틀
  const colStart = 1; 
  const colEnd = Math.max(headerRow3.length, headerRow4.length) - 1;
  const entries = [];
  const skipHeaders = ["합계", "비고", "누계", "소계", ""];

  let currentFloor = "";
  
  // 2. 5행(데이터 시작)부터 아래로 읽으며 층과 2줄 묶음 처리
  for(let r=4; r<rows.length; r++) {
    let row1 = rows[r] || [];
    let col0 = normalizeDisplayText(row1[0]);
    let col1 = normalizeDisplayText(row1[1]);
    let possibleFloor = col0 || col1;

    // 첫 칸에 텍스트가 있다면 층이 시작되는 첫 번째 줄(3행 타이틀에 대응)
    if(possibleFloor) {
      // "계" -> "합계"로 변환
      if(possibleFloor === "계" || possibleFloor === "층 계" || possibleFloor.includes("합계")) {
        currentFloor = "합계";
      } else {
        currentFloor = possibleFloor;
      }

      if(!state.floorsByBuilding[buildingName]) state.floorsByBuilding[buildingName] = new Set();
      state.floorsByBuilding[buildingName].add(currentFloor);

      // (A) 첫 번째 줄: 3행 타이틀과 매칭하여 데이터 추출
      for(let c=colStart; c<=colEnd; c++) {
        let name3 = normalizeDisplayText(headerRow3[c]);
        if(name3 && !skipHeaders.includes(name3)) {
          let val = toNumber(row1[c]);
          if(val !== 0) entries.push({ projectKey, building: buildingName, floor: currentFloor, rawName: name3, normalizedName: normalizeText(name3), qty: val });
        }
      }

      // (B) 두 번째 줄(병합된 층의 아래칸): 4행 타이틀과 매칭
      if(r+1 < rows.length) {
        let row2 = rows[r+1] || [];
        let nextCol0 = normalizeDisplayText(row2[0]);
        let nextCol1 = normalizeDisplayText(row2[1]);
        // 다음 줄의 첫 칸이 비어있다면 병합된 것으로 간주하고 파싱 진행
        if(!nextCol0 && !nextCol1) {
          for(let c=colStart; c<=colEnd; c++) {
            let name4 = normalizeDisplayText(headerRow4[c]);
            if(name4 && !skipHeaders.includes(name4)) {
              let val = toNumber(row2[c]);
              if(val !== 0) entries.push({ projectKey, building: buildingName, floor: currentFloor, rawName: name4, normalizedName: normalizeText(name4), qty: val });
            }
          }
          r++; // 두 번째 줄까지 처리했으므로 루프 건너뛰기
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

/** -----------------------------
 * UI 및 컨트롤러 업데이트 로직
 * ----------------------------- */
function updateControlsUI() {
  controlsBar.style.display = "block";
  selectBuilding.innerHTML = "";
  
  // 1. 건물(동) 드롭다운 채우기
  state.buildings.forEach(b => {
    const opt = document.createElement("option");
    opt.value = b; opt.textContent = b;
    if(state.activeBuilding === b) opt.selected = true;
    selectBuilding.appendChild(opt);
  });
  
  // 전체 합산 옵션 추가
  const sumOpt = document.createElement("option");
  sumOpt.value = "SUM_MODE"; sumOpt.textContent = "➡ 선택 동 합산 (전체 합계)";
  if(state.activeBuilding === "SUM_MODE") sumOpt.selected = true;
  selectBuilding.appendChild(sumOpt);

  // 2. 모드에 따른 층 선택 UI 조작
  if(state.activeBuilding === "SUM_MODE") {
    wrapFloor.style.display = "none";
    sumOptions.style.display = "block";
    
    // 체크박스 렌더링
    sumCheckboxes.innerHTML = state.buildings.map(b => `
      <label style="display:flex; align-items:center; gap:4px; font-size:13px;">
        <input type="checkbox" value="${b}" class="sum-chk" ${state.summationChecked.has(b) ? 'checked' : ''}>
        ${b}
      </label>
    `).join("");

    document.querySelectorAll(".sum-chk").forEach(chk => {
      chk.addEventListener("change", (e) => {
        if(e.target.checked) state.summationChecked.add(e.target.value);
        else state.summationChecked.delete(e.target.value);
        triggerRecalc();
      });
    });

  } else {
    wrapFloor.style.display = "flex";
    sumOptions.style.display = "none";
    
    // 층 드롭다운 채우기 (해당 동에 있는 층들만)
    selectFloor.innerHTML = "";
    const floors = Array.from(state.floorsByBuilding[state.activeBuilding] || []);
    
    // "합계"가 있으면 맨 위로, 나머지는 1층, 2층 오름차순 정렬
    floors.sort((a,b) => {
      if(a==="합계") return -1; if(b==="합계") return 1;
      return a.localeCompare(b, "ko", {numeric:true});
    });

    if(floors.length > 0 && !floors.includes(state.activeFloor)) {
      state.activeFloor = floors[0];
    }

    floors.forEach(f => {
      const opt = document.createElement("option");
      opt.value = f; opt.textContent = f;
      if(state.activeFloor === f) opt.selected = true;
      selectFloor.appendChild(opt);
    });
  }
}

selectBuilding.addEventListener("change", (e) => {
  state.activeBuilding = e.target.value;
  updateControlsUI();
  triggerRecalc();
});

selectFloor.addEventListener("change", (e) => {
  state.activeFloor = e.target.value;
  triggerRecalc();
});

function triggerRecalc() {
  if (state.uniqueNames.length === 0) return;
  const rows = calcCompareRows();
  renderCompareTable(rows);
}

/** -----------------------------
 * 명칭 매핑 등 기본 로직 (생략없이 포함)
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
  if (!items.length) { typeaheadRoot.innerHTML = `<div class="typeahead-empty">일치하는 항목이 없습니다. 직접 입력 후 추가할 수 있습니다.</div>`; return; }
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
  state.uniqueNames = []; state.mappingConfig = {}; state.lastCompareRows = []; state.itemCodeOptions = [...DEFAULT_ITEM_CODE_OPTIONS];
  state.buildings = []; state.floorsByBuilding = {}; state.activeBuilding = ""; state.activeFloor = ""; state.summationChecked.clear();
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
  
  // 데이터에서 존재하는 고유 동(Building) 추출
  const buildingSet = new Set();
  PROJECT_KEYS.forEach(pk => state.rawEntriesByProject[pk].forEach(e => buildingSet.add(e.building)));
  state.buildings = Array.from(buildingSet).sort();
  
  if(state.buildings.length > 0) {
    state.activeBuilding = state.buildings[0];
    state.buildings.forEach(b => state.summationChecked.add(b)); // 전체 동 기본 선택
  }

  buildUniqueNamesFromEntries(); ensureMappingConfig();
  updateControlsUI(); // 동/층 선택 UI 업데이트
  
  btnOpenMapping.disabled = state.uniqueNames.length === 0; btnCalc.disabled = state.uniqueNames.length === 0;
  setLog(logs.join("\n") || "로그가 없습니다."); setStatus(`명칭 추출 완료: ${state.uniqueNames.length}개`);
}

/** -----------------------------
 * 핵심 로직: 비교표 계산 (선택된 UI 모드에 따라 필터 및 합산)
 * ----------------------------- */
function getMappedEntriesByProject() {
  const result = { current: [], A: [], B: [], C: [] };
  for (const pk of PROJECT_KEYS) {
    result[pk] = state.rawEntriesByProject[pk].map((entry) => {
      const config = state.mappingConfig[entry.normalizedName];
      if (!config || config.include !== "Y" || !config.category || config.category === "제외" || !config.itemCode) return null;
      return { ...entry, mappedCategory: config.category, mappedItemCode: config.itemCode };
    }).filter(Boolean);
  }
  return result;
}

function calcCompareRows() {
  const mappedEntries = getMappedEntriesByProject();
  
  // 데이터를 [키: 동__층__카테고리__아이템코드] 로 합산 생성
  const agg = { current: {}, A: {}, B: {}, C: {} };
  const allDynamicKeys = new Set();
  const hardcodedKeys = new Set(BASE_LAYOUT.map(r => `__${r.category}__${r.itemCode}`));

  for (const pk of PROJECT_KEYS) {
    for (const entry of mappedEntries[pk]) {
      const key = `${entry.building}__${entry.floor}__${entry.mappedCategory}__${entry.mappedItemCode}`;
      agg[pk][key] = (agg[pk][key] || 0) + entry.qty;
      
      if (!hardcodedKeys.has(`__${entry.mappedCategory}__${entry.mappedItemCode}`)) {
        allDynamicKeys.add(`${entry.mappedCategory}__${entry.mappedItemCode}`);
      }
    }
  }

  const rows = [];
  const b = state.activeBuilding;
  const f = state.activeFloor;
  const isSumMode = (b === "SUM_MODE");

  const secTitle = isSumMode ? "선택 동 통합 합계" : `${b} - ${f}`;
  rows.push({ type: "section", section: secTitle });

  const categoryOrder = ["레미콘", "거푸집", "철근"];

  // 1. 하드코딩된 레이아웃 순회
  for (const cat of categoryOrder) {
    const hardcodedForCat = BASE_LAYOUT.filter(r => r.category === cat);
    for (const row of hardcodedForCat) {
      let cur = 0, A = 0, B = 0, C = 0;
      
      if (isSumMode) {
        // 합산 모드: 선택된 모든 동의 "합계" 층만 싹 긁어서 더함
        for(const chk of state.summationChecked) {
          const key = `${chk}__합계__${cat}__${row.itemCode}`;
          cur += agg.current[key] || 0; A += agg.A[key] || 0; B += agg.B[key] || 0; C += agg.C[key] || 0;
        }
      } else {
        // 개별 출력 모드: 현재 선택된 동과 층의 데이터만 가져옴
        const key = `${b}__${f}__${cat}__${row.itemCode}`;
        cur = agg.current[key] || 0; A = agg.A[key] || 0; B = agg.B[key] || 0; C = agg.C[key] || 0;
      }

      const avg = (A + B + C) / 3; const ratio = cur === 0 ? 0 : avg / cur;
      rows.push({ ...row, section: secTitle, current: cur, A, B, C, avg, ratio, note: row.note || "" });
    }

    // 2. 사용자가 추가한 다이나믹 항목 이어서 순회
    for (const dynKey of allDynamicKeys) {
      const parts = dynKey.split("__");
      if (parts[0] === cat) {
        const ic = parts[1];
        let cur = 0, A = 0, B = 0, C = 0;
        
        if (isSumMode) {
          for(const chk of state.summationChecked) {
            const key = `${chk}__합계__${cat}__${ic}`;
            cur += agg.current[key] || 0; A += agg.A[key] || 0; B += agg.B[key] || 0; C += agg.C[key] || 0;
          }
        } else {
          const key = `${b}__${f}__${cat}__${ic}`;
          cur = agg.current[key] || 0; A = agg.A[key] || 0; B = agg.B[key] || 0; C = agg.C[key] || 0;
        }

        const avg = (A + B + C) / 3; const ratio = cur === 0 ? 0 : avg / cur;
        rows.push({ section: secTitle, itemCode: ic, item: cat, spec: ic, category: cat, current: cur, A, B, C, avg, ratio, note: "사용자 추가" });
      }
    }
  }
  
  state.lastCompareRows = rows;
  return rows;
}

function renderCompareTable(rows) {
  if (!rows.length) {
    compareBody.innerHTML = `<tr><td colspan="11" class="empty-row">비교표가 아직 생성되지 않았습니다.</td></tr>`;
    return;
  }
  // 섹션 라인 제거하고 순수 데이터만 출력 (표 타이틀은 고정)
  const filteredRows = rows.filter(r => r.type !== "section");
  const html = filteredRows.map((row) => {
    return `
      <tr>
        <td style="font-weight:bold; color:#1d4ed8;">${escapeHtml(row.section)}</td>
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
 * 엑셀 다운로드 (갑지 양식 & 현재 화면 동기화)
 * ----------------------------- */
function exportCompareExcel() {
  if (!state.lastCompareRows.length) { setStatus("내보낼 비교표가 없습니다."); return; }

  const ws = {}; const merges = []; const rowsFormat = []; 
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
  merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 9 } }); rowsFormat[0] = { hpt: 40 }; 

  const headers = ["코드", "품명", "규격", "현재 프로젝트", "'A' 프로젝트", "'B' 프로젝트", "'C' 프로젝트", "평균치(A~C프로젝트)", "비율", "비고"];
  for (let c = 0; c < headers.length; c++) ws[XLSX.utils.encode_cell({ c: c, r: 1 })] = { v: headers[c], t: "s", s: headerStyle };
  rowsFormat[1] = { hpt: 25 };

  let r = 2; 
  state.lastCompareRows.forEach((row) => {
    if (row.type === "section") {
      for (let c = 0; c < 10; c++) ws[XLSX.utils.encode_cell({ c: c, r: r })] = { v: "", t: "s" };
      rowsFormat[r] = { hpt: 12 }; r++;

      for (let c = 0; c < 10; c++) {
        let val = c === 1 ? row.section : "";
        ws[XLSX.utils.encode_cell({ c: c, r: r })] = { v: val, t: "s", s: sectionStyle };
      }
      rowsFormat[r] = { hpt: 22 }; r++;
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
      for (let c = 0; c < rowData.length; c++) ws[XLSX.utils.encode_cell({ c: c, r: r })] = rowData[c];
      rowsFormat[r] = { hpt: 20 }; r++;
    }
  });

  ws['!ref'] = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: 9, r: r - 1 } });
  ws['!merges'] = merges; ws['!rows'] = rowsFormat; 
  ws['!cols'] = [ { wch: 10 }, { wch: 12 }, { wch: 16 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 15 }, { wch: 20 }, { wch: 10 }, { wch: 15 } ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "비교표");
  XLSX.writeFile(wb, "비교표_결과.xlsx");
}

function resetAll() {
  for (const pk of PROJECT_KEYS) fileInputs[pk].value = "";
  updateFileListText(); clearDataState(); controlsBar.style.display = "none";
  btnOpenMapping.disabled = true; btnCalc.disabled = true; btnExportCsv.disabled = true;
  compareBody.innerHTML = `<tr><td colspan="11" class="empty-row">비교표가 아직 생성되지 않았습니다.</td></tr>`;
  mappingBody.innerHTML = ""; setStatus("대기 중"); setLog("로그가 여기에 표시됩니다."); closeModal();
}

for (const pk of PROJECT_KEYS) fileInputs[pk].addEventListener("change", updateFileListText);

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

updateFileListText(); setStatus("대기 중"); setLog("로그가 여기에 표시됩니다.");
