"use strict";

/** -----------------------------
 * DOM
 * ----------------------------- */
const $ = (id) => document.getElementById(id);

const fileInputs = {
  current: $("file-current"),
  A: $("file-a"),
  B: $("file-b"),
  C: $("file-c"),
};

const fileListEls = {
  current: $("file-current-list"),
  A: $("file-a-list"),
  B: $("file-b-list"),
  C: $("file-c-list"),
};

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

// 탭 버튼 요소 가져오기
const legendChips = document.querySelectorAll(".legend__chip");

/** -----------------------------
 * 상수
 * ----------------------------- */
const PROJECT_KEYS = ["current", "A", "B", "C"];

const PROJECT_LABELS = {
  current: "현재 프로젝트",
  A: "A 프로젝트",
  B: "B 프로젝트",
  C: "C 프로젝트",
};

const SECTION_DISPLAY_CLASS = {
  APT: "section-row--apt",
  PIT: "section-row--pit",
  주차장: "section-row--parking",
  부대동: "section-row--ancillary",
};

const CATEGORY_OPTIONS = [
  { value: "", label: "선택" },
  { value: "레미콘", label: "레미콘" },
  { value: "거푸집", label: "거푸집" },
  { value: "철근", label: "철근" },
  { value: "제외", label: "제외" },
];

const INCLUDE_OPTIONS = [
  { value: "Y", label: "반영" },
  { value: "N", label: "제외" },
];

const DEFAULT_ITEM_CODE_OPTIONS = [
  "240", "270", "300", "180",
  "3회", "4회", "유로", "알폼", "갱폼", "합벽", "보밑면", "데크", "방수턱",
  "H10", "H13", "H16", "H19", "H22", "H25", "H29"
];

const COMPARE_LAYOUT = [
  { type: "section", section: "APT" },
  { section: "APT", itemCode: "240", item: "레미콘", spec: "25-24-15", category: "레미콘" },
  { section: "APT", itemCode: "270", item: "레미콘", spec: "25-27-15", category: "레미콘" },
  { section: "APT", itemCode: "300", item: "레미콘", spec: "25-30-15", category: "레미콘" },
  { section: "APT", itemCode: "180", item: "레미콘", spec: "25-18-08", category: "레미콘" },
  { section: "APT", itemCode: "3회", item: "거푸집", spec: "3회", category: "거푸집" },
  { section: "APT", itemCode: "4회", item: "거푸집", spec: "4회", category: "거푸집" },
  { section: "APT", itemCode: "유로", item: "거푸집", spec: "유로", category: "거푸집" },
  { section: "APT", itemCode: "알폼", item: "거푸집", spec: "알폼", category: "거푸집" },
  { section: "APT", itemCode: "갱폼", item: "거푸집", spec: "갱폼", category: "거푸집" },
  { section: "APT", itemCode: "합벽", item: "거푸집", spec: "합벽", category: "거푸집" },
  { section: "APT", itemCode: "보밑면", item: "거푸집", spec: "보밑면", category: "거푸집" },
  { section: "APT", itemCode: "데크", item: "거푸집", spec: "데크플레이트", category: "거푸집" },
  { section: "APT", itemCode: "방수턱", item: "거푸집", spec: "방수턱", category: "거푸집" },
  { section: "APT", itemCode: "H10", item: "철근", spec: "H10", category: "철근" },
  { section: "APT", itemCode: "H13", item: "철근", spec: "H13", category: "철근" },
  { section: "APT", itemCode: "H16", item: "철근", spec: "H16", category: "철근" },
  { section: "APT", itemCode: "H19", item: "철근", spec: "H19", category: "철근" },
  { section: "APT", itemCode: "H22", item: "철근", spec: "H22", category: "철근" },
  { section: "APT", itemCode: "H25", item: "철근", spec: "H25", category: "철근" },
  { section: "APT", itemCode: "H29", item: "철근", spec: "H29", category: "철근" },

  { type: "section", section: "PIT" },
  { section: "PIT", itemCode: "240", item: "레미콘", spec: "25-24-15", category: "레미콘" },
  { section: "PIT", itemCode: "270", item: "레미콘", spec: "25-27-15", category: "레미콘" },
  { section: "PIT", itemCode: "300", item: "레미콘", spec: "25-30-15", category: "레미콘" },
  { section: "PIT", itemCode: "180", item: "레미콘", spec: "25-18-08", category: "레미콘" },
  { section: "PIT", itemCode: "3회", item: "거푸집", spec: "3회", category: "거푸집" },
  { section: "PIT", itemCode: "4회", item: "거푸집", spec: "4회", category: "거푸집" },
  { section: "PIT", itemCode: "유로", item: "거푸집", spec: "유로", category: "거푸집" },
  { section: "PIT", itemCode: "알폼", item: "거푸집", spec: "알폼", category: "거푸집" },
  { section: "PIT", itemCode: "갱폼", item: "거푸집", spec: "갱폼", category: "거푸집" },
  { section: "PIT", itemCode: "합벽", item: "거푸집", spec: "합벽", category: "거푸집" },
  { section: "PIT", itemCode: "보밑면", item: "거푸집", spec: "보밑면", category: "거푸집" },
  { section: "PIT", itemCode: "데크", item: "거푸집", spec: "데크플레이트", category: "거푸집" },
  { section: "PIT", itemCode: "방수턱", item: "거푸집", spec: "방수턱", category: "거푸집" },
  { section: "PIT", itemCode: "H10", item: "철근", spec: "H10", category: "철근" },
  { section: "PIT", itemCode: "H13", item: "철근", spec: "H13", category: "철근" },
  { section: "PIT", itemCode: "H16", item: "철근", spec: "H16", category: "철근" },
  { section: "PIT", itemCode: "H19", item: "철근", spec: "H19", category: "철근" },
  { section: "PIT", itemCode: "H22", item: "철근", spec: "H22", category: "철근" },
  { section: "PIT", itemCode: "H25", item: "철근", spec: "H25", category: "철근" },
  { section: "PIT", itemCode: "H29", item: "철근", spec: "H29", category: "철근" },

  { type: "section", section: "주차장" },
  { section: "주차장", itemCode: "240", item: "레미콘", spec: "25-24-15", category: "레미콘" },
  { section: "주차장", itemCode: "270", item: "레미콘", spec: "25-27-15", category: "레미콘" },
  { section: "주차장", itemCode: "300", item: "레미콘", spec: "25-30-15", category: "레미콘" },
  { section: "주차장", itemCode: "180", item: "레미콘", spec: "25-18-08", category: "레미콘" },
  { section: "주차장", itemCode: "3회", item: "거푸집", spec: "3회", category: "거푸집" },
  { section: "주차장", itemCode: "4회", item: "거푸집", spec: "4회", category: "거푸집" },
  { section: "주차장", itemCode: "유로", item: "거푸집", spec: "유로", category: "거푸집" },
  { section: "주차장", itemCode: "알폼", item: "거푸집", spec: "알폼", category: "거푸집" },
  { section: "주차장", itemCode: "갱폼", item: "거푸집", spec: "갱폼", category: "거푸집" },
  { section: "주차장", itemCode: "합벽", item: "거푸집", spec: "합벽", category: "거푸집" },
  { section: "주차장", itemCode: "보밑면", item: "거푸집", spec: "보밑면", category: "거푸집" },
  { section: "주차장", itemCode: "데크", item: "거푸집", spec: "데크플레이트", category: "거푸집" },
  { section: "주차장", itemCode: "방수턱", item: "거푸집", spec: "방수턱", category: "거푸집" },
  { section: "주차장", itemCode: "H10", item: "철근", spec: "H10", category: "철근" },
  { section: "주차장", itemCode: "H13", item: "철근", spec: "H13", category: "철근" },
  { section: "주차장", itemCode: "H16", item: "철근", spec: "H16", category: "철근" },
  { section: "주차장", itemCode: "H19", item: "철근", spec: "H19", category: "철근" },
  { section: "주차장", itemCode: "H22", item: "철근", spec: "H22", category: "철근" },
  { section: "주차장", itemCode: "H25", item: "철근", spec: "H25", category: "철근" },
  { section: "주차장", itemCode: "H29", item: "철근", spec: "H29", category: "철근" },

  { type: "section", section: "부대동" },
  { section: "부대동", itemCode: "240", item: "레미콘", spec: "25-24-15", category: "레미콘" },
  { section: "부대동", itemCode: "270", item: "레미콘", spec: "25-27-15", category: "레미콘" },
  { section: "부대동", itemCode: "300", item: "레미콘", spec: "25-30-15", category: "레미콘" },
  { section: "부대동", itemCode: "180", item: "레미콘", spec: "25-18-08", category: "레미콘" },
  { section: "부대동", itemCode: "3회", item: "거푸집", spec: "3회", category: "거푸집" },
  { section: "부대동", itemCode: "4회", item: "거푸집", spec: "4회", category: "거푸집" },
  { section: "부대동", itemCode: "유로", item: "거푸집", spec: "유로", category: "거푸집" },
  { section: "부대동", itemCode: "알폼", item: "거푸집", spec: "알폼", category: "거푸집" },
  { section: "부대동", itemCode: "갱폼", item: "거푸집", spec: "갱폼", category: "거푸집" },
  { section: "부대동", itemCode: "합벽", item: "거푸집", spec: "합벽", category: "거푸집" },
  { section: "부대동", itemCode: "보밑면", item: "거푸집", spec: "보밑면", category: "거푸집" },
  { section: "부대동", itemCode: "데크", item: "거푸집", spec: "데크플레이트", category: "거푸집" },
  { section: "부대동", itemCode: "방수턱", item: "거푸집", spec: "방수턱", category: "거푸집" },
  { section: "부대동", itemCode: "H10", item: "철근", spec: "H10", category: "철근" },
  { section: "부대동", itemCode: "H13", item: "철근", spec: "H13", category: "철근" },
  { section: "부대동", itemCode: "H16", item: "철근", spec: "H16", category: "철근" },
  { section: "부대동", itemCode: "H19", item: "철근", spec: "H19", category: "철근" },
  { section: "부대동", itemCode: "H22", item: "철근", spec: "H22", category: "철근" },
  { section: "부대동", itemCode: "H25", item: "철근", spec: "H25", category: "철근" },
  { section: "부대동", itemCode: "H29", item: "철근", spec: "H29", category: "철근" },
];

/** -----------------------------
 * 상태
 * ----------------------------- */
const state = {
  rawEntriesByProject: {
    current: [],
    A: [],
    B: [],
    C: [],
  },
  fileSummaryByProject: {
    current: [],
    A: [],
    B: [],
    C: [],
  },
  uniqueNames: [],
  mappingConfig: {},
  lastCompareRows: [],
  itemCodeOptions: [...DEFAULT_ITEM_CODE_OPTIONS],
  typeahead: {
    targetInput: null,
    targetKey: "",
    items: [],
    activeIndex: -1,
  },
  activeTab: "APT", 
};

/** -----------------------------
 * 공통 유틸
 * ----------------------------- */
function setStatus(text) {
  statusBox.textContent = text;
}

function setLog(text) {
  logBox.textContent = text;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function normalizeText(value) {
  return String(value ?? "").replace(/\s+/g, "").trim().toUpperCase();
}

function normalizeDisplayText(value) {
  return String(value ?? "").replace(/\s+/g, " ").trim();
}

function toNumber(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }
  const cleaned = String(value ?? "").replace(/,/g, "").trim();
  if (!cleaned) return 0;
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function fmtNumber(value) {
  return Number(value || 0).toLocaleString("ko-KR", {
    maximumFractionDigits: 0,
  });
}

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
    if (!files.length) {
      fileListEls[key].textContent = "선택된 파일 없음";
      continue;
    }
    fileListEls[key].textContent = files.map((f) => f.name).join("\n");
  }
}

function openModal() {
  mappingModal.classList.add("is-open");
  mappingModal.setAttribute("aria-hidden", "false");
}

function closeModal() {
  mappingModal.classList.remove("is-open");
  mappingModal.setAttribute("aria-hidden", "true");
  hideTypeahead();
}

function sortUniqueStrings(list) {
  return [...new Set(list.filter(Boolean))].sort((a, b) => a.localeCompare(b, "ko"));
}

function ensureItemCodeOption(value) {
  const text = normalizeDisplayText(value);
  if (!text) return;
  if (!state.itemCodeOptions.includes(text)) {
    state.itemCodeOptions.push(text);
    state.itemCodeOptions = sortUniqueStrings(state.itemCodeOptions);
  }
}

/** -----------------------------
 * 탭 관련 UI 업데이트
 * ----------------------------- */
function updateTabUI() {
  legendChips.forEach(chip => {
    const chipText = chip.textContent.trim();
    if (chipText === state.activeTab) {
      chip.classList.add("is-active");
      chip.style.opacity = "1";
      chip.style.borderWidth = "2px";
    } else {
      chip.classList.remove("is-active");
      chip.style.opacity = "0.4";
      chip.style.borderWidth = "1px";
    }
  });
}

/** -----------------------------
 * 엑셀 파싱
 * ----------------------------- */
function getSectionNameFromSheetName(sheetName) {
  const name = normalizeText(sheetName);
  if (name.includes("APT") || name.includes("아파트")) return "APT";
  if (name.includes("PIT") || name.includes("피트")) return "PIT";
  if (name.includes("주차장")) return "주차장";
  if (name.includes("부대")) return "부대동";
  return ""; 
}

function parseSheetEntries(rows, sectionName, projectKey, fileName) {
  const headerRow3 = rows[2] || []; 
  const headerRow4 = rows[3] || []; 

  let rowIdxSum3 = -1; 
  let rowIdxSum4 = -1; 

  for (let r = 5; r < rows.length; r++) {
    const text = (rows[r] || []).slice(0, 10).map(x => normalizeText(String(x))).join("");
    if (text.includes("수량합계") && rowIdxSum3 === -1) {
      rowIdxSum3 = r;
    }
    if (text.includes("면적합계") && rowIdxSum4 === -1) {
      rowIdxSum4 = r;
    }
  }

  if (rowIdxSum3 === -1 && rows.length > 37) rowIdxSum3 = 37;
  if (rowIdxSum4 === -1 && rows.length > 38) rowIdxSum4 = 38;

  const colStart = 1;
  const colEnd = Math.max(headerRow3.length, headerRow4.length) - 1;
  const entries = [];
  
  const skipHeaders = ["합계", "비고", "누계", "소계"];

  for (let c = colStart; c <= colEnd; c += 1) {
    const name3 = normalizeDisplayText(headerRow3[c]);
    const name4 = normalizeDisplayText(headerRow4[c]);

    if (name3 && !skipHeaders.includes(name3)) {
      const val = rowIdxSum3 !== -1 ? toNumber(rows[rowIdxSum3][c]) : 0;
      if (val !== 0) {
        entries.push({
          projectKey,
          projectLabel: PROJECT_LABELS[projectKey],
          section: sectionName || "미확인",
          sourceFile: fileName,
          rowIndex: rowIdxSum3 + 1,
          rowLabel: "수량합계",
          rawName: name3,
          normalizedName: normalizeText(name3),
          qty: val,
        });
      }
    }

    if (name4 && !skipHeaders.includes(name4)) {
      const val = rowIdxSum4 !== -1 ? toNumber(rows[rowIdxSum4][c]) : 0;
      if (val !== 0) {
        entries.push({
          projectKey,
          projectLabel: PROJECT_LABELS[projectKey],
          section: sectionName || "미확인",
          sourceFile: fileName,
          rowIndex: rowIdxSum4 + 1,
          rowLabel: "면적합계",
          rawName: name4,
          normalizedName: normalizeText(name4),
          qty: val,
        });
      }
    }
  }

  return entries;
}

async function parseWorkbookFile(file, projectKey) {
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  
  const allEntries = [];
  let totalEntryCount = 0;
  const parsedSheets = [];

  for (const sheetName of workbook.SheetNames) {
    const sectionName = getSectionNameFromSheetName(sheetName);
    
    if (!sectionName) continue;

    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: true,
      defval: "",
    });

    const entries = parseSheetEntries(rows, sectionName, projectKey, file.name);
    allEntries.push(...entries);
    totalEntryCount += entries.length;
    parsedSheets.push(`${sheetName}(${sectionName})`);
  }

  return {
    projectKey,
    fileName: file.name,
    sheetNames: parsedSheets.join(", "),
    entryCount: totalEntryCount,
    entries: allEntries,
  };
}

/** -----------------------------
 * 명칭 처리 제안
 * ----------------------------- */
function suggestMappingByName(rawName) {
  const t = normalizeText(rawName);

  const result = {
    include: "Y",
    category: "",
    itemCode: "",
    note: "",
  };

  if (t.includes("24MPA")) {
    result.category = "레미콘";
    result.itemCode = "240";
    result.note = "24MPA → 240";
    return result;
  }
  if (t.includes("27MPA")) {
    result.category = "레미콘";
    result.itemCode = "270";
    result.note = "27MPA → 270";
    return result;
  }
  if (t.includes("30MPA")) {
    result.category = "레미콘";
    result.itemCode = "300";
    result.note = "30MPA → 300";
    return result;
  }
  if (t.includes("현치/무근") || t.includes("버림")) {
    result.category = "레미콘";
    result.itemCode = "180";
    result.note = "버림/현치무근 → 180";
    return result;
  }
  if (t.includes("현치/구체") || t === "합벽채움" || t.includes("ACT-INSIDE")) {
    result.category = "레미콘";
    result.itemCode = "240";
    result.note = "구체/합벽/ACT-INSIDE 계열 → 240";
    return result;
  }

  if (t === "3회") {
    result.category = "거푸집";
    result.itemCode = "3회";
    result.note = "3회 → 3회";
    return result;
  }
  if (t === "4회") {
    result.category = "거푸집";
    result.itemCode = "4회";
    result.note = "4회 → 4회";
    return result;
  }
  if (t === "유로") {
    result.category = "거푸집";
    result.itemCode = "유로";
    result.note = "유로 → 유로";
    return result;
  }
  if (t === "원형" || t === "CURVED") {
    result.category = "거푸집";
    result.itemCode = "유로";
    result.note = "원형/CURVED → 유로";
    return result;
  }
  if (t === "알폼-H" || t === "알폼-V" || t === "알폼") {
    result.category = "거푸집";
    result.itemCode = "알폼";
    result.note = "알폼-H/V → 알폼";
    return result;
  }
  if (t === "갱폼" || t === "갱폼-2") {
    result.category = "거푸집";
    result.itemCode = "갱폼";
    result.note = "갱폼/갱폼-2 → 갱폼";
    return result;
  }
  if (t === "합벽") {
    result.category = "거푸집";
    result.itemCode = "합벽";
    result.note = "합벽 → 합벽";
    return result;
  }
  if (t === "보밑면") {
    result.category = "거푸집";
    result.itemCode = "보밑면";
    result.note = "보밑면 → 보밑면";
    return result;
  }
  if (t === "DECK") {
    result.category = "거푸집";
    result.itemCode = "데크";
    result.note = "DECK → 데크";
    return result;
  }
  if (t === "방수턱") {
    result.category = "거푸집";
    result.itemCode = "방수턱";
    result.note = "방수턱 → 방수턱";
    return result;
  }

  if (t === "H10") {
    result.category = "철근";
    result.itemCode = "H10";
    result.note = "H10 → H10";
    return result;
  }
  if (t === "H13") {
    result.category = "철근";
    result.itemCode = "H13";
    result.note = "H13 → H13";
    return result;
  }
  if (t === "H16") {
    result.category = "철근";
    result.itemCode = "H16";
    result.note = "H16 → H16";
    return result;
  }
  if (t === "H19") {
    result.category = "철근";
    result.itemCode = "H19";
    result.note = "H19 → H19";
    return result;
  }
  if (t === "D22" || t === "H22") {
    result.category = "철근";
    result.itemCode = "H22";
    result.note = "D22/H22 → H22";
    return result;
  }
  if (t === "H25") {
    result.category = "철근";
    result.itemCode = "H25";
    result.note = "H25 → H25";
    return result;
  }
  if (t === "H29") {
    result.category = "철근";
    result.itemCode = "H29";
    result.note = "H29 → H29";
    return result;
  }

  result.include = "N";
  result.category = "제외";
  result.itemCode = "";
  result.note = "자동 제안 없음";
  return result;
}

/** -----------------------------
 * 명칭 추출 / 매핑 렌더링
 * ----------------------------- */
function buildUniqueNamesFromEntries() {
  const nameMap = new Map();

  for (const projectKey of PROJECT_KEYS) {
    for (const entry of state.rawEntriesByProject[projectKey]) {
      if (!nameMap.has(entry.normalizedName)) {
        nameMap.set(entry.normalizedName, {
          rawName: entry.rawName,
          normalizedName: entry.normalizedName,
        });
      }
    }
  }

  state.uniqueNames = Array.from(nameMap.values()).sort((a, b) => {
    return a.rawName.localeCompare(b.rawName, "ko");
  });
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
  return list
    .map((opt) => {
      const selected = String(opt.value) === String(selectedValue) ? "selected" : "";
      return `<option value="${escapeHtml(opt.value)}" ${selected}>${escapeHtml(opt.label)}</option>`;
    })
    .join("");
}

function getCategoryClass(category) {
  if (category === "레미콘") return "is-concrete";
  if (category === "거푸집") return "is-form";
  if (category === "철근") return "is-rebar";
  return "is-exclude";
}

function getRowIncludeClass(include) {
  return include === "Y" ? "is-included" : "is-excluded";
}

function renderMappingTable() {
  if (!state.uniqueNames.length) {
    mappingBody.innerHTML = `
      <tr>
        <td colspan="5">추출된 명칭이 없습니다.</td>
      </tr>
    `;
    return;
  }

  const html = state.uniqueNames
    .map((item) => {
      const config = state.mappingConfig[item.normalizedName] || suggestMappingByName(item.rawName);
      const suggestion = suggestMappingByName(item.rawName);

      return `
        <tr class="map-row ${getRowIncludeClass(config.include)}" data-name-key="${escapeHtml(item.normalizedName)}">
          <td>${escapeHtml(item.rawName)}</td>
          <td>
            <select class="map-include">
              ${optionHtml(INCLUDE_OPTIONS, config.include)}
            </select>
          </td>
          <td>
            <select class="map-category ${getCategoryClass(config.category)}">
              ${optionHtml(CATEGORY_OPTIONS, config.category)}
            </select>
          </td>
          <td class="itemcode-cell">
            <div class="itemcode-editor">
              <input
                type="text"
                class="itemcode-input"
                value="${escapeHtml(config.itemCode || "")}"
                autocomplete="off"
                placeholder="직접 입력 또는 Alt+↓"
              />
              <button type="button" class="itemcode-add-btn">추가</button>
            </div>
            <div class="itemcode-help">Alt+↓ : 목록 열기 / 직접 입력 가능 / 신규 코드 추가 가능</div>
          </td>
          <td class="mapping-suggest">${escapeHtml(suggestion.note || "-")}</td>
        </tr>
      `;
    })
    .join("");

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

    includeSelect.addEventListener("change", () => {
      applyVisualStateToRow(row);
    });

    categorySelect.addEventListener("change", () => {
      applyVisualStateToRow(row);
    });

    itemInput.addEventListener("focus", () => {
      showTypeaheadForInput(itemInput, row.dataset.nameKey);
    });

    itemInput.addEventListener("input", () => {
      showTypeaheadForInput(itemInput, row.dataset.nameKey);
    });

    itemInput.addEventListener("keydown", (e) => {
      if (e.altKey && e.key === "ArrowDown") {
        e.preventDefault();
        showTypeaheadForInput(itemInput, row.dataset.nameKey, true);
        return;
      }

      if (typeaheadRoot.hidden) return;

      if (state.typeahead.targetInput !== itemInput) return;

      if (e.key === "ArrowDown") {
        e.preventDefault();
        moveTypeahead(1);
        return;
      }
      if (e.key === "ArrowUp") {
        e.preventDefault();
        moveTypeahead(-1);
        return;
      }
      if (e.key === "Enter") {
        if (state.typeahead.activeIndex >= 0) {
          e.preventDefault();
          commitTypeaheadSelection(state.typeahead.activeIndex);
        } else {
          const text = normalizeDisplayText(itemInput.value);
          if (text) {
            e.preventDefault();
            commitManualItemCode(row.dataset.nameKey, itemInput, text);
          }
        }
        return;
      }
      if (e.key === "Escape") {
        hideTypeahead();
      }
    });

    itemInput.addEventListener("blur", () => {
      setTimeout(() => {
        if (document.activeElement && typeaheadRoot.contains(document.activeElement)) return;
        hideTypeahead();
      }, 120);
    });

    addBtn.addEventListener("click", () => {
      const text = normalizeDisplayText(itemInput.value);
      if (!text) return;
      commitManualItemCode(row.dataset.nameKey, itemInput, text);
    });
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

    state.mappingConfig[key] = {
      include,
      category,
      itemCode,
      note: suggestMappingByName(key).note || "",
    };
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
 * 타입어헤드
 * ----------------------------- */
function getFilteredItemCodeOptions(keyword, forceAll = false) {
  const q = normalizeText(keyword);
  const list = sortUniqueStrings(state.itemCodeOptions);

  if (forceAll || !q) return list.slice(0, 200);

  const starts = [];
  const includes = [];

  for (const item of list) {
    const n = normalizeText(item);
    if (n.startsWith(q)) {
      starts.push(item);
    } else if (n.includes(q)) {
      includes.push(item);
    }
  }

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
  state.typeahead.items = items;
  state.typeahead.activeIndex = items.length ? 0 : -1;

  if (!items.length) {
    typeaheadRoot.innerHTML = `<div class="typeahead-empty">일치하는 항목이 없습니다. 직접 입력 후 추가할 수 있습니다.</div>`;
    return;
  }

  typeaheadRoot.innerHTML = `
    <div class="typeahead-list">
      ${items.map((item, idx) => `
        <button type="button" class="typeahead-item ${idx === 0 ? "is-active" : ""}" data-index="${idx}">
          ${escapeHtml(item)}
        </button>
      `).join("")}
    </div>
  `;

  Array.from(typeaheadRoot.querySelectorAll(".typeahead-item")).forEach((btn) => {
    btn.addEventListener("mousedown", (e) => {
      e.preventDefault();
    });
    btn.addEventListener("click", () => {
      const idx = Number(btn.dataset.index);
      commitTypeaheadSelection(idx);
    });
  });
}

function showTypeaheadForInput(input, nameKey, forceAll = false) {
  state.typeahead.targetInput = input;
  state.typeahead.targetKey = nameKey;

  const items = getFilteredItemCodeOptions(input.value, forceAll);
  positionTypeahead(input);
  renderTypeahead(items);
  typeaheadRoot.hidden = false;
}

function hideTypeahead() {
  typeaheadRoot.hidden = true;
  typeaheadRoot.innerHTML = "";
  state.typeahead.targetInput = null;
  state.typeahead.targetKey = "";
  state.typeahead.items = [];
  state.typeahead.activeIndex = -1;
}

function refreshTypeaheadActive() {
  Array.from(typeaheadRoot.querySelectorAll(".typeahead-item")).forEach((el, idx) => {
    el.classList.toggle("is-active", idx === state.typeahead.activeIndex);
  });
}

function moveTypeahead(delta) {
  const len = state.typeahead.items.length;
  if (!len) return;
  const next = state.typeahead.activeIndex + delta;
  if (next < 0) {
    state.typeahead.activeIndex = len - 1;
  } else if (next >= len) {
    state.typeahead.activeIndex = 0;
  } else {
    state.typeahead.activeIndex = next;
  }
  refreshTypeaheadActive();
}

function commitTypeaheadSelection(index) {
  const value = state.typeahead.items[index];
  if (!value || !state.typeahead.targetInput) return;

  state.typeahead.targetInput.value = value;
  ensureItemCodeOption(value);
  hideTypeahead();
  state.typeahead.targetInput?.focus();
}

function commitManualItemCode(nameKey, input, text) {
  ensureItemCodeOption(text);
  input.value = text;

  const row = input.closest("tr[data-name-key]");
  const category = row.querySelector(".map-category")?.value || "";
  const include = row.querySelector(".map-include")?.value || "Y";

  state.mappingConfig[nameKey] = {
    include,
    category,
    itemCode: text,
    note: suggestMappingByName(nameKey).note || "",
  };

  hideTypeahead();
  setStatus(`아이템구분 추가: ${text}`);
}

/** -----------------------------
 * 데이터 취합
 * ----------------------------- */
function clearDataState() {
  state.rawEntriesByProject = {
    current: [],
    A: [],
    B: [],
    C: [],
  };
  state.fileSummaryByProject = {
    current: [],
    A: [],
    B: [],
    C: [],
  };
  state.uniqueNames = [];
  state.mappingConfig = {};
  state.lastCompareRows = [];
  state.itemCodeOptions = [...DEFAULT_ITEM_CODE_OPTIONS];
}

async function extractNamesFromFiles() {
  clearDataState();

  const logs = [];
  let totalFileCount = 0;

  for (const projectKey of PROJECT_KEYS) {
    const files = Array.from(fileInputs[projectKey].files || []);
    totalFileCount += files.length;

    for (const file of files) {
      const parsed = await parseWorkbookFile(file, projectKey);
      state.rawEntriesByProject[projectKey].push(...parsed.entries);
      state.fileSummaryByProject[projectKey].push(parsed);

      logs.push(
        `[${PROJECT_LABELS[projectKey]}] 파일: ${parsed.fileName}`,
        `- 인식된 시트: ${parsed.sheetNames || "없음"}`,
        `- 총 추출건수: ${parsed.entryCount}`,
        ""
      );
    }
  }

  if (!totalFileCount) {
    throw new Error("업로드된 파일이 없습니다.");
  }

  buildUniqueNamesFromEntries();
  ensureMappingConfig();

  btnOpenMapping.disabled = state.uniqueNames.length === 0;
  btnCalc.disabled = state.uniqueNames.length === 0;

  setLog(logs.join("\n") || "로그가 없습니다.");
  setStatus(`명칭 추출 완료: ${state.uniqueNames.length}개`);
}

/** -----------------------------
 * 비교표 계산
 * ----------------------------- */
function getMappedEntriesByProject() {
  const result = {
    current: [],
    A: [],
    B: [],
    C: [],
  };

  for (const projectKey of PROJECT_KEYS) {
    result[projectKey] = state.rawEntriesByProject[projectKey]
      .map((entry) => {
        const config = state.mappingConfig[entry.normalizedName];
        if (!config) return null;
        if (config.include !== "Y") return null;
        if (!config.category || config.category === "제외") return null;
        if (!config.itemCode) return null;

        return {
          ...entry,
          mappedCategory: config.category,
          mappedItemCode: config.itemCode,
        };
      })
      .filter(Boolean);
  }

  return result;
}

function buildProjectAggregate(mappedEntries) {
  const aggregate = {
    current: {},
    A: {},
    B: {},
    C: {},
  };

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
  for (const projectKey of PROJECT_KEYS) {
    for (const k of Object.keys(aggregate[projectKey])) {
      allKeys.add(k);
    }
  }

  const dynamicItems = {};
  const hardcodedKeys = new Set();

  for (const row of COMPARE_LAYOUT) {
    if (row.type !== "section") {
      hardcodedKeys.add(`${row.section}__${row.category}__${row.itemCode}`);
    }
  }

  for (const k of allKeys) {
    if (!hardcodedKeys.has(k)) {
      const parts = k.split("__");
      const sec = parts[0];
      const cat = parts[1];
      const ic = parts[2];

      if (!dynamicItems[sec]) dynamicItems[sec] = {};
      if (!dynamicItems[sec][cat]) dynamicItems[sec][cat] = [];
      
      dynamicItems[sec][cat].push(ic);
    }
  }

  const rows = [];
  const sections = ["APT", "PIT", "주차장", "부대동"];
  const categoryOrder = ["레미콘", "거푸집", "철근"]; 

  for (const sec of sections) {
    rows.push({ type: "section", section: sec });

    for (const cat of categoryOrder) {
      const hardcodedForCat = COMPARE_LAYOUT.filter(
        r => r.type !== "section" && r.section === sec && r.category === cat
      );
      
      for (const row of hardcodedForCat) {
        const key = `${sec}__${cat}__${row.itemCode}`;
        const cur = aggregate.current[key] || 0;
        const A = aggregate.A[key] || 0;
        const B = aggregate.B[key] || 0;
        const C = aggregate.C[key] || 0;
        const avg = (A + B + C) / 3;
        const ratio = cur === 0 ? 0 : avg / cur;

        rows.push({
          ...row,
          current: cur, A, B, C, avg, ratio, note: row.note || ""
        });
      }

      if (dynamicItems[sec] && dynamicItems[sec][cat]) {
        for (const ic of dynamicItems[sec][cat]) {
          const key = `${sec}__${cat}__${ic}`;
          const cur = aggregate.current[key] || 0;
          const A = aggregate.A[key] || 0;
          const B = aggregate.B[key] || 0;
          const C = aggregate.C[key] || 0;
          const avg = (A + B + C) / 3;
          const ratio = cur === 0 ? 0 : avg / cur;

          rows.push({
            section: sec,
            itemCode: ic,
            item: cat,       
            spec: ic,        
            category: cat,
            current: cur, A, B, C, avg, ratio, note: "사용자 추가 항목"
          });
        }
      }
    }
    
    if (dynamicItems[sec]) {
      for (const [dynCat, dynIcs] of Object.entries(dynamicItems[sec])) {
        if (!categoryOrder.includes(dynCat)) {
          for (const ic of dynIcs) {
            const key = `${sec}__${dynCat}__${ic}`;
            const cur = aggregate.current[key] || 0;
            const A = aggregate.A[key] || 0;
            const B = aggregate.B[key] || 0;
            const C = aggregate.C[key] || 0;
            const avg = (A + B + C) / 3;
            const ratio = cur === 0 ? 0 : avg / cur;

            rows.push({
              section: sec,
              itemCode: ic,
              item: dynCat,
              spec: ic,
              category: dynCat,
              current: cur, A, B, C, avg, ratio, note: "사용자 추가 항목"
            });
          }
        }
      }
    }
  }

  state.lastCompareRows = rows;
  return rows;
}

// UI 표 렌더링 (화면에는 탭으로 보여줌)
function renderCompareTable(rows) {
  if (!rows.length) {
    compareBody.innerHTML = `
      <tr>
        <td colspan="11" class="empty-row">비교표가 아직 생성되지 않았습니다.</td>
      </tr>
    `;
    return;
  }

  const filteredRows = rows.filter(r => r.section === state.activeTab && r.type !== "section");

  if (!filteredRows.length) {
    compareBody.innerHTML = `
      <tr>
        <td colspan="11" class="empty-row">[${state.activeTab}] 영역에 해당하는 데이터가 없습니다.</td>
      </tr>
    `;
    return;
  }

  const html = filteredRows
    .map((row) => {
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
        </tr>
      `;
    })
    .join("");

  compareBody.innerHTML = html;
}

function validateBeforeCalc() {
  if (!state.uniqueNames.length) {
    throw new Error("먼저 명칭을 추출해야 합니다.");
  }

  const missing = [];

  for (const item of state.uniqueNames) {
    const config = state.mappingConfig[item.normalizedName];
    if (!config) {
      missing.push(`${item.rawName} : 설정 없음`);
      continue;
    }
    if (config.include === "Y") {
      if (!config.category || config.category === "제외") {
        missing.push(`${item.rawName} : 분류 미설정`);
      }
      if (!config.itemCode) {
        missing.push(`${item.rawName} : 아이템구분 미설정`);
      }
    }
  }

  if (missing.length) {
    throw new Error(
      `명칭 설정이 완료되지 않았습니다.\n\n${missing.slice(0, 30).join("\n")}${missing.length > 30 ? "\n..." : ""}`
    );
  }
}

/** -----------------------------
 * 엑셀 다운로드 (원하시는 "비교표_갑지" 양식과 100% 동일하게 구현)
 * ----------------------------- */
function exportCompareExcel() {
  if (!state.lastCompareRows.length) {
    setStatus("내보낼 비교표가 없습니다.");
    return;
  }

  const ws = {};
  const merges = [];
  const rowsFormat = []; // 횡방향 행간 두께(높이) 조절용 배열
  
  // 1. 공통 폰트 및 테두리 두께 지정
  const fontName = "맑은 고딕";
  const borderAll = {
    top: { style: "thin", color: { rgb: "000000" } },
    bottom: { style: "thin", color: { rgb: "000000" } },
    left: { style: "thin", color: { rgb: "000000" } },
    right: { style: "thin", color: { rgb: "000000" } }
  };

  // 2. 각 영역별 스타일 디테일 지정 (컬러, 폰트 크기, 정렬 등)
  const titleStyle = {
    font: { name: fontName, bold: true, sz: 16, color: { rgb: "000000" } },
    alignment: { horizontal: "center", vertical: "center" }
  };

  const headerStyle = {
    fill: { fgColor: { rgb: "F2F2F2" } }, // 회색 배경
    font: { name: fontName, bold: true, sz: 10, color: { rgb: "000000" } },
    alignment: { horizontal: "center", vertical: "center" },
    border: borderAll
  };

  const sectionStyle = {
    fill: { fgColor: { rgb: "E2EFDA" } }, // 연초록색 배경
    font: { name: fontName, bold: true, sz: 10, color: { rgb: "000000" } },
    alignment: { horizontal: "center", vertical: "center" },
    border: borderAll
  };

  const centerStyle = {
    font: { name: fontName, sz: 10, color: { rgb: "000000" } },
    alignment: { horizontal: "center", vertical: "center" },
    border: borderAll
  };

  const numberStyle = {
    font: { name: fontName, sz: 10, color: { rgb: "000000" } },
    alignment: { horizontal: "right", vertical: "center" },
    border: borderAll,
    numFmt: "#,##0" // 소수점 버리고 천단위 콤마
  };

  const ratioStyle = {
    font: { name: fontName, sz: 10, color: { rgb: "000000" } },
    alignment: { horizontal: "center", vertical: "center" },
    border: borderAll,
    numFmt: "0%" // 비율을 완벽한 100% 엑셀 서식으로 지정
  };

  // 3. 최상단 타이틀 입력 및 행 높이 설정
  ws[XLSX.utils.encode_cell({ c: 0, r: 0 })] = { v: "ㅇㅇ 프로젝트 비교분석자료", t: "s", s: titleStyle };
  merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 9 } }); // A1 ~ J1 병합
  rowsFormat[0] = { hpt: 40 }; // 1행 높이 크게

  // 4. 헤더 설정 (10열 양식)
  const headers = ["코드", "품명", "규격", "현재 프로젝트", "'A' 프로젝트", "'B' 프로젝트", "'C' 프로젝트", "평균치(A~C프로젝트)", "비율", "비고"];
  for (let c = 0; c < headers.length; c++) {
    ws[XLSX.utils.encode_cell({ c: c, r: 1 })] = { v: headers[c], t: "s", s: headerStyle };
  }
  rowsFormat[1] = { hpt: 25 }; // 2행 높이

  // 5. 실제 데이터 입력
  let r = 2; 
  state.lastCompareRows.forEach((row) => {
    if (row.type === "section") {
      // 5-1. 동(APT, PIT 등) 영역 전환 시 한 줄 빈 행 추가
      for (let c = 0; c < 10; c++) {
        ws[XLSX.utils.encode_cell({ c: c, r: r })] = { v: "", t: "s" };
      }
      rowsFormat[r] = { hpt: 12 }; // 빈 줄은 높이를 얇게
      r++;

      // 5-2. 섹션 타이틀 행 추가 (B열에 타이틀 배치, A~J 배경 칠하기)
      for (let c = 0; c < 10; c++) {
        let val = c === 1 ? row.section : "";
        ws[XLSX.utils.encode_cell({ c: c, r: r })] = { v: val, t: "s", s: sectionStyle };
      }
      rowsFormat[r] = { hpt: 22 }; // 섹션 행 높이
      r++;
    } else {
      // 5-3. 데이터 렌더링
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
        { v: ratioVal, t: "n", s: ratioStyle }, // 100% 변환
        { v: row.note || "", t: "s", s: centerStyle },
      ];

      for (let c = 0; c < rowData.length; c++) {
        ws[XLSX.utils.encode_cell({ c: c, r: r })] = rowData[c];
      }
      rowsFormat[r] = { hpt: 20 }; // 일반 데이터 행 높이
      r++;
    }
  });

  // 6. 셀 크기, 병합, 행 높이 등 최종 옵션 주입
  ws['!ref'] = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: 9, r: r - 1 } });
  ws['!merges'] = merges;
  ws['!rows'] = rowsFormat; // 행 높이(횡방향 두께) 적용
  ws['!cols'] = [ // 열 너비(종방향 두께) 적용
    { wch: 10 }, // A: 코드
    { wch: 12 }, // B: 품명
    { wch: 16 }, // C: 규격
    { wch: 15 }, // D: 현재 프로젝트
    { wch: 15 }, // E: A 프로젝트
    { wch: 15 }, // F: B 프로젝트
    { wch: 15 }, // G: C 프로젝트
    { wch: 20 }, // H: 평균치
    { wch: 10 }, // I: 비율
    { wch: 15 }, // J: 비고
  ];

  // 7. 엑셀 파일 생성
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "비교표");
  XLSX.writeFile(wb, "비교표_결과.xlsx");
}

/** -----------------------------
 * 초기화
 * ----------------------------- */
function resetAll() {
  for (const key of PROJECT_KEYS) {
    fileInputs[key].value = "";
  }
  updateFileListText();
  clearDataState();

  btnOpenMapping.disabled = true;
  btnCalc.disabled = true;
  btnExportCsv.disabled = true;

  compareBody.innerHTML = `
    <tr>
      <td colspan="11" class="empty-row">비교표가 아직 생성되지 않았습니다.</td>
    </tr>
  `;

  mappingBody.innerHTML = "";
  setStatus("대기 중");
  setLog("로그가 여기에 표시됩니다.");
  closeModal();
}

/** -----------------------------
 * 이벤트 연결
 * ----------------------------- */
for (const key of PROJECT_KEYS) {
  fileInputs[key].addEventListener("change", updateFileListText);
}

// 탭 버튼 클릭 이벤트 등록
legendChips.forEach(chip => {
  chip.addEventListener("click", (e) => {
    state.activeTab = e.target.textContent.trim();
    updateTabUI();
    if (state.lastCompareRows.length > 0) {
      renderCompareTable(state.lastCompareRows);
    }
  });
});

btnExportCsv.textContent = "엑셀 내보내기";

btnExtract.addEventListener("click", async () => {
  try {
    setStatus("명칭 추출 중...");
    setLog("파일을 분석하고 있습니다...");
    await extractNamesFromFiles();
    renderMappingTable();
    openModal();
  } catch (error) {
    console.error(error);
    setStatus("오류 발생");
    setLog(error?.message || String(error));
  }
});

btnOpenMapping.addEventListener("click", () => {
  renderMappingTable();
  openModal();
});

btnCloseMapping.addEventListener("click", closeModal);
mappingBackdrop.addEventListener("click", closeModal);

btnApplySuggestions.addEventListener("click", () => {
  applyAutoSuggestionsToCurrentMapping();
});

btnSaveMapping.addEventListener("click", () => {
  saveMappingFromUI();
  closeModal();
  setStatus("명칭 설정 저장 완료");
});

btnCalc.addEventListener("click", () => {
  try {
    if (mappingModal.classList.contains("is-open")) {
      saveMappingFromUI();
    }

    validateBeforeCalc();
    const rows = calcCompareRows();
    renderCompareTable(rows);

    btnExportCsv.disabled = rows.length === 0;
    setStatus("비교표 생성 완료");
  } catch (error) {
    console.error(error);
    setStatus("오류 발생");
    setLog(error?.message || String(error));
  }
});

btnExportCsv.addEventListener("click", exportCompareExcel);
btnReset.addEventListener("click", resetAll);

document.addEventListener("click", (e) => {
  if (!typeaheadRoot.hidden) {
    const clickedInsideTypeahead = typeaheadRoot.contains(e.target);
    const clickedInput = e.target.closest(".itemcode-input");
    if (!clickedInsideTypeahead && !clickedInput) {
      hideTypeahead();
    }
  }
});

window.addEventListener("resize", () => {
  if (!typeaheadRoot.hidden && state.typeahead.targetInput) {
    positionTypeahead(state.typeahead.targetInput);
  }
});

window.addEventListener("scroll", () => {
  if (!typeaheadRoot.hidden && state.typeahead.targetInput) {
    positionTypeahead(state.typeahead.targetInput);
  }
}, true);

/** -----------------------------
 * 최초 상태 초기화
 * ----------------------------- */
updateFileListText();
updateTabUI();
setStatus("대기 중");
setLog("로그가 여기에 표시됩니다.");
