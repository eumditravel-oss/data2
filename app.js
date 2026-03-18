"use strict";

/* =========================
   기본 상태
========================= */
const PROJECTS = [
  { key: "current", name: "현재 프로젝트" },
  { key: "a", name: "A 프로젝트" },
  { key: "b", name: "B 프로젝트" },
  { key: "c", name: "C 프로젝트" }
];

const state = {
  projects: {
    current: createEmptyProjectState(),
    a: createEmptyProjectState(),
    b: createEmptyProjectState(),
    c: createEmptyProjectState()
  },
  mappings: {},           // rawKey => canonicalName
  parsedReady: false,
  mappedReady: false
};

function createEmptyProjectState() {
  return {
    files: [],
    rawItems: [],         // 원본 아이템명 목록
    dongs: [],            // 동 목록
    floors: [],           // 층 목록
    data: {}              // data[dong][rawItem][floor] = value
  };
}

/* =========================
   DOM
========================= */
const $ = (id) => document.getElementById(id);

const dom = {
  tabs: [...document.querySelectorAll(".tab")],
  tabPanels: {
    upload: $("tab-upload"),
    mapping: $("tab-mapping"),
    compare: $("tab-compare")
  },

  fileInputs: {
    current: $("file-current"),
    a: $("file-a"),
    b: $("file-b"),
    c: $("file-c")
  },

  fileNames: {
    current: $("name-current"),
    a: $("name-a"),
    b: $("name-b"),
    c: $("name-c")
  },

  fileLists: {
    current: $("list-current"),
    a: $("list-a"),
    b: $("list-b"),
    c: $("list-c")
  },

  btnParse: $("btn-parse"),
  btnReset: $("btn-reset"),

  uploadStatus: $("upload-status"),

  mappingProjects: $("mapping-projects"),
  mappingTbody: $("mapping-tbody"),
  mappingStatus: $("mapping-status"),
  mappingSearch: $("mapping-search"),
  btnAutofill: $("btn-autofill"),
  btnApplyMapping: $("btn-apply-mapping"),
  canonicalOptions: $("canonical-options"),

  filterDong: $("filter-dong"),
  filterItem: $("filter-item"),
  filterMode: $("filter-mode"),
  btnRenderCompare: $("btn-render-compare"),
  compareThead: $("compare-thead"),
  compareTbody: $("compare-tbody"),
  compareStatus: $("compare-status"),
  summaryCards: $("summary-cards")
};

/* =========================
   탭 전환
========================= */
dom.tabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    const target = tab.dataset.tab;
    setActiveTab(target);
  });
});

function setActiveTab(tabKey) {
  dom.tabs.forEach((btn) => btn.classList.toggle("is-active", btn.dataset.tab === tabKey));
  Object.entries(dom.tabPanels).forEach(([k, panel]) => {
    panel.classList.toggle("is-active", k === tabKey);
  });
}

/* =========================
   파일 선택 UI
========================= */
PROJECTS.forEach(({ key }) => {
  dom.fileInputs[key].addEventListener("change", (e) => {
    const files = [...(e.target.files || [])];
    state.projects[key].files = files;
    renderFileUI(key);
  });
});

function renderFileUI(projectKey) {
  const files = state.projects[projectKey].files;
  dom.fileNames[projectKey].textContent = files.length
    ? `${files.length}개 파일 선택됨`
    : "선택된 파일 없음";

  dom.fileLists[projectKey].innerHTML = files.length
    ? files.map(f => `<span class="file-chip">${escapeHtml(f.name)}</span>`).join("")
    : "";
}

/* =========================
   초기화
========================= */
dom.btnReset.addEventListener("click", () => {
  PROJECTS.forEach(({ key }) => {
    state.projects[key] = createEmptyProjectState();
    dom.fileInputs[key].value = "";
    renderFileUI(key);
  });

  state.mappings = {};
  state.parsedReady = false;
  state.mappedReady = false;

  dom.uploadStatus.textContent = "";
  dom.mappingProjects.innerHTML = "";
  dom.mappingTbody.innerHTML = `<tr><td colspan="4" class="empty">업로드 후 아이템을 추출하면 여기에 표시됩니다.</td></tr>`;
  dom.mappingStatus.textContent = "";
  dom.filterDong.innerHTML = "";
  dom.compareThead.innerHTML = "";
  dom.compareTbody.innerHTML = `<tr><td class="empty">비교할 데이터를 먼저 준비해 주세요.</td></tr>`;
  dom.compareStatus.textContent = "";
  dom.summaryCards.innerHTML = "";

  setActiveTab("upload");
});

/* =========================
   EXCEL 읽기 / 파싱
========================= */
dom.btnParse.addEventListener("click", async () => {
  const hasAny = PROJECTS.some(({ key }) => state.projects[key].files.length > 0);
  if (!hasAny) {
    dom.uploadStatus.textContent = "업로드된 파일이 없습니다.";
    return;
  }

  dom.uploadStatus.textContent = "EXCEL 파일을 읽는 중입니다...";

  try {
    for (const { key } of PROJECTS) {
      const projectState = createEmptyProjectState();
      projectState.files = state.projects[key].files.slice();

      for (const file of state.projects[key].files) {
        const parsed = await parseExcelFile(file);
        mergeParsedProject(projectState, parsed);
      }

      projectState.rawItems = uniqueSort(projectState.rawItems);
      projectState.dongs = uniqueSort(projectState.dongs, dongSorter);
      projectState.floors = uniqueSort(projectState.floors, floorSorter);

      state.projects[key] = projectState;
    }

    state.parsedReady = true;
    state.mappedReady = false;

    renderMappingProjectLists();
    renderMappingTable();
    populateCanonicalDatalist();

    const statusLines = PROJECTS.map(({ key, name }) => {
      const p = state.projects[key];
      return `${name} : 파일 ${p.files.length}개 / 동 ${p.dongs.length}개 / 아이템 ${p.rawItems.length}개`;
    });

    dom.uploadStatus.textContent = "업로드 및 아이템 추출이 완료되었습니다.\n" + statusLines.join("\n");

    setActiveTab("mapping");
  } catch (err) {
    console.error(err);
    dom.uploadStatus.textContent = `오류가 발생했습니다.\n${err.message || err}`;
  }
});

async function parseExcelFile(file) {
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: "array" });

  const aggregated = {
    rawItems: [],
    dongs: [],
    floors: [],
    data: {}
  };

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
    const mergedRows = applyMerges(rows, ws["!merges"] || []);

    const parsedSheet = parseWorksheetRows(mergedRows);
    mergeParsedProject(aggregated, parsedSheet);
  }

  return aggregated;
}

function applyMerges(rows, merges) {
  const grid = rows.map(r => r.slice());

  for (const merge of merges) {
    const startRow = merge.s.r;
    const endRow = merge.e.r;
    const startCol = merge.s.c;
    const endCol = merge.e.c;

    const topLeftValue = (((grid[startRow] || [])[startCol]) ?? "");

    for (let r = startRow; r <= endRow; r++) {
      if (!grid[r]) grid[r] = [];
      for (let c = startCol; c <= endCol; c++) {
        if (grid[r][c] === undefined || grid[r][c] === "") {
          grid[r][c] = topLeftValue;
        }
      }
    }
  }

  return grid;
}

function parseWorksheetRows(rows) {
  const result = {
    rawItems: [],
    dongs: [],
    floors: [],
    data: {}
  };

  if (!rows || rows.length < 5) return result;

  // 엑셀 기준 3행, 4행
  const row3 = rows[2] || [];
  const row4 = rows[3] || [];

  let currentDong = "";
  let currentFloor = "";
  let sameFloorCount = 0;
  let previousFloor = null;

  for (let r = 4; r < rows.length; r++) {
    const row = rows[r] || [];
    const rowText = row.map(v => String(v ?? "").trim()).join(" | ");

    // 동 인식: [101동] 형태
    const dongMatch = rowText.match(/\[([^\]]+)\]/);
    if (dongMatch) {
      currentDong = dongMatch[1].trim();
      if (currentDong) {
        pushUnique(result.dongs, currentDong);
        if (!result.data[currentDong]) result.data[currentDong] = {};
      }
      currentFloor = "";
      sameFloorCount = 0;
      previousFloor = null;
      continue;
    }

    if (!currentDong) continue;

    const floorCandidate = normalizeFloorCell(row[0]);
    if (floorCandidate !== "") {
      currentFloor = floorCandidate;
      pushUnique(result.floors, currentFloor);

      if (previousFloor === currentFloor) {
        sameFloorCount += 1;
      } else {
        sameFloorCount = 1;
        previousFloor = currentFloor;
      }
    } else if (!currentFloor) {
      continue;
    } else {
      // A열이 비었는데 병합 해제 상황일 수 있어 이전 층 유지
      if (previousFloor === currentFloor) {
        sameFloorCount += 1;
      } else {
        sameFloorCount = 1;
        previousFloor = currentFloor;
      }
    }

    // floor 당 2행 구조
    // 1행차 = 3행 아이템, 2행차 = 4행 아이템
    const itemRowSource = sameFloorCount % 2 === 1 ? row3 : row4;

    for (let c = 1; c < row.length; c++) {
      const rawLabel = normalizeItemName(itemRowSource[c]);
      if (!rawLabel) continue;

      const val = toNumber(row[c]);
      if (val === null) continue;

      pushUnique(result.rawItems, rawLabel);

      if (!result.data[currentDong][rawLabel]) result.data[currentDong][rawLabel] = {};
      const oldVal = Number(result.data[currentDong][rawLabel][currentFloor] || 0);
      result.data[currentDong][rawLabel][currentFloor] = oldVal + val;
    }
  }

  return result;
}

function mergeParsedProject(target, parsed) {
  if (!parsed) return;

  parsed.rawItems?.forEach(item => pushUnique(target.rawItems, item));
  parsed.dongs?.forEach(d => pushUnique(target.dongs, d));
  parsed.floors?.forEach(f => pushUnique(target.floors, f));

  target.data = target.data || {};

  for (const dong of Object.keys(parsed.data || {})) {
    if (!target.data[dong]) target.data[dong] = {};

    for (const rawItem of Object.keys(parsed.data[dong] || {})) {
      if (!target.data[dong][rawItem]) target.data[dong][rawItem] = {};

      for (const floor of Object.keys(parsed.data[dong][rawItem] || {})) {
        const oldVal = Number(target.data[dong][rawItem][floor] || 0);
        const addVal = Number(parsed.data[dong][rawItem][floor] || 0);
        target.data[dong][rawItem][floor] = oldVal + addVal;
      }
    }
  }
}

/* =========================
   아이템 통일 설정
========================= */
function renderMappingProjectLists() {
  dom.mappingProjects.innerHTML = PROJECTS.map(({ key, name }) => {
    const items = state.projects[key].rawItems || [];
    const body = items.length
      ? items.map(item => `<div class="item-badge">${escapeHtml(item)}</div>`).join("")
      : `<div class="item-badge">아이템 없음</div>`;

    return `
      <div class="project-item-card">
        <div class="project-item-card__head">${escapeHtml(name)}</div>
        <div class="project-item-card__body">${body}</div>
      </div>
    `;
  }).join("");
}

function renderMappingTable() {
  const rows = [];

  PROJECTS.forEach(({ key, name }) => {
    const rawItems = state.projects[key].rawItems || [];
    rawItems.forEach(rawItem => {
      const rawKey = makeRawKey(key, rawItem);
      const suggested = suggestCanonicalName(rawItem);
      const mapped = state.mappings[rawKey] ?? suggested;

      rows.push({
        projectKey: key,
        projectName: name,
        rawItem,
        rawKey,
        suggested,
        mapped
      });
    });
  });

  if (!rows.length) {
    dom.mappingTbody.innerHTML = `<tr><td colspan="4" class="empty">업로드 후 아이템을 추출하면 여기에 표시됩니다.</td></tr>`;
    return;
  }

  const q = dom.mappingSearch.value.trim().toLowerCase();

  const filtered = rows.filter(row =>
    !q ||
    row.projectName.toLowerCase().includes(q) ||
    row.rawItem.toLowerCase().includes(q) ||
    row.suggested.toLowerCase().includes(q) ||
    String(row.mapped).toLowerCase().includes(q)
  );

  dom.mappingTbody.innerHTML = filtered.length
    ? filtered.map(row => `
      <tr>
        <td>${escapeHtml(row.projectName)}</td>
        <td>${escapeHtml(row.rawItem)}</td>
        <td>${escapeHtml(row.suggested)}</td>
        <td>
          <input
            class="mapping-input"
            data-rawkey="${escapeHtml(row.rawKey)}"
            list="canonical-options"
            value="${escapeHtmlAttr(row.mapped)}"
            placeholder="통일할 아이템명 입력"
          />
        </td>
      </tr>
    `).join("")
    : `<tr><td colspan="4" class="empty">검색 결과가 없습니다.</td></tr>`;

  bindMappingInputs();
}

function bindMappingInputs() {
  dom.mappingTbody.querySelectorAll(".mapping-input").forEach(input => {
    input.addEventListener("input", (e) => {
      const rawKey = e.target.dataset.rawkey;
      state.mappings[rawKey] = e.target.value.trim();
    });
  });
}

dom.mappingSearch.addEventListener("input", renderMappingTable);

dom.btnAutofill.addEventListener("click", () => {
  PROJECTS.forEach(({ key }) => {
    (state.projects[key].rawItems || []).forEach(rawItem => {
      const rawKey = makeRawKey(key, rawItem);
      if (!state.mappings[rawKey]) {
        state.mappings[rawKey] = suggestCanonicalName(rawItem);
      }
    });
  });

  renderMappingTable();
  dom.mappingStatus.textContent = "자동 제안명으로 통일명을 채웠습니다. 필요하면 수정 후 적용해 주세요.";
});

dom.btnApplyMapping.addEventListener("click", () => {
  if (!state.parsedReady) {
    dom.mappingStatus.textContent = "먼저 EXCEL 파일을 읽어 주세요.";
    return;
  }

  // 빈값 보정
  PROJECTS.forEach(({ key }) => {
    (state.projects[key].rawItems || []).forEach(rawItem => {
      const rawKey = makeRawKey(key, rawItem);
      if (!state.mappings[rawKey] || !state.mappings[rawKey].trim()) {
        state.mappings[rawKey] = suggestCanonicalName(rawItem);
      }
    });
  });

  state.mappedReady = true;
  populateDongFilter();
  renderSummaryCards();
  renderCompareTable();

  dom.mappingStatus.textContent = "아이템 통일명이 적용되었습니다.";
  setActiveTab("compare");
});

function populateCanonicalDatalist() {
  const allNames = [];

  PROJECTS.forEach(({ key }) => {
    (state.projects[key].rawItems || []).forEach(rawItem => {
      pushUnique(allNames, suggestCanonicalName(rawItem));
    });
  });

  dom.canonicalOptions.innerHTML = uniqueSort(allNames).map(name =>
    `<option value="${escapeHtmlAttr(name)}"></option>`
  ).join("");
}

function makeRawKey(projectKey, rawItem) {
  return `${projectKey}::${rawItem}`;
}

function suggestCanonicalName(rawItem) {
  return String(rawItem || "")
    .replace(/\s+/g, " ")
    .replace(/[_\-]+/g, "-")
    .replace(/\s*\/\s*/g, "/")
    .trim();
}

/* =========================
   비교표
========================= */
function populateDongFilter() {
  const allDongs = uniqueSort(
    PROJECTS.flatMap(({ key }) => state.projects[key].dongs || []),
    dongSorter
  );

  dom.filterDong.innerHTML = allDongs.length
    ? allDongs.map(d => `<option value="${escapeHtmlAttr(d)}">${escapeHtml(d)}</option>`).join("")
    : `<option value="">동 없음</option>`;
}

dom.btnRenderCompare.addEventListener("click", renderCompareTable);
dom.filterDong.addEventListener("change", () => {
  renderSummaryCards();
  renderCompareTable();
});
dom.filterItem.addEventListener("input", renderCompareTable);
dom.filterMode.addEventListener("change", renderCompareTable);

function renderSummaryCards() {
  if (!state.mappedReady) {
    dom.summaryCards.innerHTML = "";
    return;
  }

  const dong = dom.filterDong.value;
  if (!dong) {
    dom.summaryCards.innerHTML = "";
    return;
  }

  const viewData = buildUnifiedDataByDong(dong);
  const unifiedItems = Object.keys(viewData.items);
  const floors = viewData.floors;

  let ratioCount = 0;
  let ratioHigh = 0;
  let ratioLow = 0;

  unifiedItems.forEach(item => {
    floors.forEach(floor => {
      const row = viewData.items[item][floor];
      if (row && row.ratio !== null) {
        ratioCount++;
        if (row.ratio >= 110) ratioHigh++;
        if (row.ratio <= 90) ratioLow++;
      }
    });
  });

  dom.summaryCards.innerHTML = `
    <div class="summary-card">
      <div class="summary-card__label">선택 동</div>
      <div class="summary-card__value">${escapeHtml(dong)}</div>
    </div>
    <div class="summary-card">
      <div class="summary-card__label">비교 아이템 수</div>
      <div class="summary-card__value">${unifiedItems.length}</div>
    </div>
    <div class="summary-card">
      <div class="summary-card__label">비율 계산 셀 수</div>
      <div class="summary-card__value">${ratioCount}</div>
    </div>
    <div class="summary-card">
      <div class="summary-card__label">과다 / 과소 셀 수</div>
      <div class="summary-card__value">${ratioHigh} / ${ratioLow}</div>
    </div>
  `;
}

function renderCompareTable() {
  if (!state.mappedReady) {
    dom.compareThead.innerHTML = "";
    dom.compareTbody.innerHTML = `<tr><td class="empty">아이템 통일명을 먼저 적용해 주세요.</td></tr>`;
    dom.compareStatus.textContent = "";
    return;
  }

  const dong = dom.filterDong.value;
  if (!dong) {
    dom.compareThead.innerHTML = "";
    dom.compareTbody.innerHTML = `<tr><td class="empty">동을 선택해 주세요.</td></tr>`;
    dom.compareStatus.textContent = "";
    return;
  }

  const keyword = dom.filterItem.value.trim().toLowerCase();
  const ratioMode = dom.filterMode.value;

  const viewData = buildUnifiedDataByDong(dong);
  let items = Object.keys(viewData.items);
  const floors = viewData.floors;

  if (keyword) {
    items = items.filter(item => item.toLowerCase().includes(keyword));
  }

  items = uniqueSort(items);

  if (!items.length) {
    dom.compareThead.innerHTML = "";
    dom.compareTbody.innerHTML = `<tr><td class="empty">조건에 맞는 아이템이 없습니다.</td></tr>`;
    dom.compareStatus.textContent = "";
    return;
  }

  renderCompareHeader(floors);
  renderCompareBody(items, floors, viewData, ratioMode);

  dom.compareStatus.textContent =
    `동 ${dong} 기준 / 아이템 ${items.length}개 / 층 ${floors.length}개`;
}

function buildUnifiedDataByDong(dong) {
  const unified = {
    floors: [],
    items: {}
  };

  const allFloors = uniqueSort(
    PROJECTS.flatMap(({ key }) => {
      const rawDongData = state.projects[key].data[dong] || {};
      return Object.values(rawDongData).flatMap(itemObj => Object.keys(itemObj || {}));
    }),
    floorSorter
  );

  unified.floors = allFloors;

  PROJECTS.forEach(({ key }) => {
    const dongData = state.projects[key].data[dong] || {};

    Object.keys(dongData).forEach(rawItem => {
      const rawKey = makeRawKey(key, rawItem);
      const canonical = (state.mappings[rawKey] || suggestCanonicalName(rawItem)).trim();
      if (!canonical) return;

      if (!unified.items[canonical]) unified.items[canonical] = {};
      allFloors.forEach(floor => {
        if (!unified.items[canonical][floor]) {
          unified.items[canonical][floor] = {
            current: null,
            a: null,
            b: null,
            c: null,
            avg: null,
            ratio: null
          };
        }

        const val = dongData[rawItem]?.[floor];
        if (typeof val === "number" && !Number.isNaN(val)) {
          unified.items[canonical][floor][key] =
            Number(unified.items[canonical][floor][key] || 0) + val;
        }
      });
    });
  });

  Object.keys(unified.items).forEach(item => {
    allFloors.forEach(floor => {
      const cell = unified.items[item][floor];
      const simVals = [cell.a, cell.b, cell.c].filter(v => typeof v === "number");
      cell.avg = simVals.length ? average(simVals) : null;
      cell.ratio = (typeof cell.current === "number" && typeof cell.avg === "number" && cell.avg !== 0)
        ? (cell.current / cell.avg) * 100
        : null;
    });
  });

  return unified;
}

function renderCompareHeader(floors) {
  const groups = [
    { key: "current", label: "현재 프로젝트", cls: "compare-group-current" },
    { key: "a", label: "유사 A", cls: "compare-group-a" },
    { key: "b", label: "유사 B", cls: "compare-group-b" },
    { key: "c", label: "유사 C", cls: "compare-group-c" },
    { key: "avg", label: "유사 3개 평균", cls: "compare-group-avg" },
    { key: "ratio", label: "현재 / 평균 비율(%)", cls: "compare-group-ratio" }
  ];

  const top = [];
  top.push(`<th class="sticky-col" rowspan="2" style="min-width:110px">아이템</th>`);
  top.push(`<th class="sticky-col-2" rowspan="2" style="min-width:140px">구분</th>`);

  groups.forEach(g => {
    top.push(`<th class="${g.cls}" colspan="${floors.length}">${g.label}</th>`);
  });

  const bottom = [];
  groups.forEach(g => {
    floors.forEach(floor => {
      bottom.push(`<th class="compare-floor-head ${g.cls}">${escapeHtml(floor)}</th>`);
    });
  });

  dom.compareThead.innerHTML = `
    <tr>${top.join("")}</tr>
    <tr>${bottom.join("")}</tr>
  `;
}

function renderCompareBody(items, floors, viewData, ratioMode) {
  const rows = [];

  items.forEach(item => {
    const tr = [];
    tr.push(`<td class="sticky-col row-item">${escapeHtml(item)}</td>`);
    tr.push(`<td class="sticky-col-2 row-item-sub">수량 / 평균 / 비율</td>`);

    // current
    floors.forEach(floor => {
      tr.push(`<td>${formatValue(viewData.items[item][floor].current)}</td>`);
    });

    // a
    floors.forEach(floor => {
      tr.push(`<td>${formatValue(viewData.items[item][floor].a)}</td>`);
    });

    // b
    floors.forEach(floor => {
      tr.push(`<td>${formatValue(viewData.items[item][floor].b)}</td>`);
    });

    // c
    floors.forEach(floor => {
      tr.push(`<td>${formatValue(viewData.items[item][floor].c)}</td>`);
    });

    // avg
    floors.forEach(floor => {
      tr.push(`<td>${formatValue(viewData.items[item][floor].avg)}</td>`);
    });

    // ratio
    floors.forEach(floor => {
      const ratio = viewData.items[item][floor].ratio;
      const cls = ratioClass(ratio, ratioMode);
      tr.push(`<td class="${cls}">${formatRatio(ratio)}</td>`);
    });

    rows.push(`<tr>${tr.join("")}</tr>`);
  });

  dom.compareTbody.innerHTML = rows.join("");
}

function ratioClass(ratio, mode) {
  if (ratio === null) return "value-empty";
  if (mode === "ratio") {
    if (ratio >= 110) return "ratio-high";
    if (ratio <= 90) return "ratio-low";
    return "ratio-mid";
  }
  if (ratio >= 110) return "ratio-high";
  if (ratio <= 90) return "ratio-low";
  return "";
}

/* =========================
   유틸
========================= */
function normalizeItemName(v) {
  return String(v ?? "").replace(/\s+/g, " ").trim();
}

function normalizeFloorCell(v) {
  const s = String(v ?? "").trim();
  if (!s) return "";
  // 1, 2, 3 ... 형태를 1F, 2F로 표준화
  if (/^\d+$/.test(s)) return `${s}F`;
  return s;
}

function toNumber(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return Number.isFinite(v) ? v : null;

  const s = String(v).replace(/,/g, "").trim();
  if (!s) return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function average(arr) {
  if (!arr.length) return null;
  return arr.reduce((a, b) => a + b, 0) / arr.length;
}

function pushUnique(arr, value) {
  if (!arr.includes(value)) arr.push(value);
}

function uniqueSort(arr, sorter) {
  const unique = [...new Set(arr)];
  return sorter ? unique.sort(sorter) : unique.sort();
}

function floorSorter(a, b) {
  const na = parseFloorNumber(a);
  const nb = parseFloorNumber(b);
  if (na !== null && nb !== null) return na - nb;
  return String(a).localeCompare(String(b), "ko");
}

function dongSorter(a, b) {
  const na = parseLeadingNumber(a);
  const nb = parseLeadingNumber(b);
  if (na !== null && nb !== null) return na - nb;
  return String(a).localeCompare(String(b), "ko");
}

function parseFloorNumber(v) {
  const m = String(v).match(/-?\d+/);
  return m ? Number(m[0]) : null;
}

function parseLeadingNumber(v) {
  const m = String(v).match(/\d+/);
  return m ? Number(m[0]) : null;
}

function formatValue(v) {
  if (v === null || v === undefined || Number.isNaN(v)) {
    return `<span class="value-empty">-</span>`;
  }
  return Number(v).toLocaleString("ko-KR", { maximumFractionDigits: 3 });
}

function formatRatio(v) {
  if (v === null || v === undefined || Number.isNaN(v)) {
    return `<span class="value-empty">-</span>`;
  }
  return `${v.toFixed(1)}%`;
}

function escapeHtml(str) {
  return String(str ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function escapeHtmlAttr(str) {
  return escapeHtml(str).replaceAll("`", "&#96;");
}
