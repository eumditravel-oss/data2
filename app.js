"use strict";

/* =========================
   상태
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
  mappingGroups: [],       // [{groupId, signature, suggested, canonical, itemsByProject}]
  mappings: {},            // rawKey => canonicalName
  parsedReady: false,
  mappedReady: false
};

function createEmptyProjectState() {
  return {
    files: [],
    rawItems: [],
    dongs: [],
    floors: [],
    data: {}
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

  mappingSearch: $("mapping-search"),
  btnAutofill: $("btn-autofill"),
  btnApplyMapping: $("btn-apply-mapping"),
  mappingGroupList: $("mapping-group-list"),
  mappingStatus: $("mapping-status"),
  canonicalOptions: $("canonical-options"),

  filterDong: $("filter-dong"),
  filterItem: $("filter-item"),
  filterMode: $("filter-mode"),
  btnRenderCompare: $("btn-render-compare"),
  summaryCards: $("summary-cards"),
  compareCardList: $("compare-card-list"),
  compareStatus: $("compare-status")
};

/* =========================
   탭
========================= */
dom.tabs.forEach((tab) => {
  tab.addEventListener("click", () => setActiveTab(tab.dataset.tab));
});

function setActiveTab(tabKey) {
  dom.tabs.forEach((btn) => btn.classList.toggle("is-active", btn.dataset.tab === tabKey));
  Object.entries(dom.tabPanels).forEach(([k, panel]) => {
    panel.classList.toggle("is-active", k === tabKey);
  });
}

/* =========================
   파일 UI
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
  dom.fileNames[projectKey].textContent = files.length ? `${files.length}개 파일 선택됨` : "선택된 파일 없음";
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

  state.mappingGroups = [];
  state.mappings = {};
  state.parsedReady = false;
  state.mappedReady = false;

  dom.uploadStatus.textContent = "";
  dom.mappingGroupList.innerHTML = "";
  dom.mappingStatus.textContent = "";
  dom.filterDong.innerHTML = "";
  dom.summaryCards.innerHTML = "";
  dom.compareCardList.innerHTML = `<div class="empty-box">비교할 데이터를 먼저 준비해 주세요.</div>`;
  dom.compareStatus.textContent = "";

  setActiveTab("upload");
});

/* =========================
   업로드 및 파싱
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

    buildMappingGroups();
    renderMappingGroupList();
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

  const row3 = rows[2] || [];
  const row4 = rows[3] || [];

  let currentDong = "";
  let currentFloor = "";
  let sameFloorCount = 0;
  let previousFloor = null;

  for (let r = 4; r < rows.length; r++) {
    const row = rows[r] || [];
    const rowText = row.map(v => String(v ?? "").trim()).join(" | ");

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

      if (previousFloor === currentFloor) sameFloorCount += 1;
      else {
        sameFloorCount = 1;
        previousFloor = currentFloor;
      }
    } else if (!currentFloor) {
      continue;
    } else {
      if (previousFloor === currentFloor) sameFloorCount += 1;
      else {
        sameFloorCount = 1;
        previousFloor = currentFloor;
      }
    }

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
   아이템 그룹화
========================= */
function buildMappingGroups() {
  const grouped = new Map();

  PROJECTS.forEach(({ key }) => {
    const items = state.projects[key].rawItems || [];
    items.forEach(rawItem => {
      const signature = makeItemSignature(rawItem);
      if (!grouped.has(signature)) {
        grouped.set(signature, {
          groupId: signature,
          signature,
          suggested: suggestCanonicalName(rawItem),
          canonical: suggestCanonicalName(rawItem),
          itemsByProject: {
            current: [],
            a: [],
            b: [],
            c: []
          }
        });
      }
      grouped.get(signature).itemsByProject[key].push(rawItem);
    });
  });

  state.mappingGroups = [...grouped.values()]
    .map(group => {
      group.canonical = chooseBestCanonical(group);
      return group;
    })
    .sort((a, b) => a.canonical.localeCompare(b.canonical, "ko"));

  state.mappings = {};
  state.mappingGroups.forEach(group => {
    PROJECTS.forEach(({ key }) => {
      group.itemsByProject[key].forEach(rawItem => {
        state.mappings[makeRawKey(key, rawItem)] = group.canonical;
      });
    });
  });
}

function makeItemSignature(rawItem) {
  const s = String(rawItem || "").toUpperCase().trim();
  return s
    .replace(/\s+/g, "")
    .replace(/[_\-]/g, "")
    .replace(/[(){}\[\]]/g, "")
    .replace(/\/+/g, "/")
    .replace(/MPA/g, "MPA")
    .replace(/WH빔/g, "WHB")
    .replace(/빔/g, "B")
    .replace(/합벽/g, "합벽")
    .replace(/CURVED/g, "CURVED");
}

function chooseBestCanonical(group) {
  const all = [
    ...group.itemsByProject.current,
    ...group.itemsByProject.a,
    ...group.itemsByProject.b,
    ...group.itemsByProject.c
  ];
  if (!all.length) return group.suggested;
  return uniqueSort(all, (x, y) => String(x).length - String(y).length || String(x).localeCompare(String(y), "ko"))[0];
}

function renderMappingGroupList() {
  const q = dom.mappingSearch.value.trim().toLowerCase();

  const groups = state.mappingGroups.filter(group => {
    const texts = [
      group.canonical,
      ...group.itemsByProject.current,
      ...group.itemsByProject.a,
      ...group.itemsByProject.b,
      ...group.itemsByProject.c
    ].join(" ").toLowerCase();

    return !q || texts.includes(q);
  });

  if (!groups.length) {
    dom.mappingGroupList.innerHTML = `<div class="empty-box">업로드 후 아이템 그룹이 여기에 표시됩니다.</div>`;
    return;
  }

  dom.mappingGroupList.innerHTML = groups.map(group => {
    return `
      <div class="mapping-group-card">
        <div class="mapping-group-card__top">
          <div class="mapping-group-card__left">
            <div class="mapping-group-card__title">${escapeHtml(group.canonical)}</div>
            <div class="mapping-group-card__meta">
              유사 아이템을 한 그룹으로 묶었습니다. 이 그룹의 모든 아이템은 아래 통일명으로 적용됩니다.
            </div>
          </div>
          <div class="mapping-group-card__right">
            <div class="mapping-canonical">
              <label>통일 아이템명</label>
              <input
                class="group-canonical-input"
                data-groupid="${escapeHtmlAttr(group.groupId)}"
                list="canonical-options"
                value="${escapeHtmlAttr(group.canonical)}"
                placeholder="통일 아이템명 입력"
              />
            </div>
          </div>
        </div>

        <div class="mapping-project-grid">
          ${renderMappingProjectColumn("current", "현재 프로젝트", group.itemsByProject.current)}
          ${renderMappingProjectColumn("a", "A 프로젝트", group.itemsByProject.a)}
          ${renderMappingProjectColumn("b", "B 프로젝트", group.itemsByProject.b)}
          ${renderMappingProjectColumn("c", "C 프로젝트", group.itemsByProject.c)}
        </div>
      </div>
    `;
  }).join("");

  bindGroupCanonicalInputs();
}

function renderMappingProjectColumn(projectKey, title, items) {
  const body = items.length
    ? items.map(item => `<span class="mapping-item-chip">${escapeHtml(item)}</span>`).join("")
    : `<span class="mapping-item-chip is-empty">해당 없음</span>`;

  return `
    <div class="mapping-project-col ${projectKey}">
      <div class="mapping-project-col__head">${escapeHtml(title)}</div>
      <div class="mapping-project-col__body">${body}</div>
    </div>
  `;
}

function bindGroupCanonicalInputs() {
  document.querySelectorAll(".group-canonical-input").forEach(input => {
    input.addEventListener("input", (e) => {
      const groupId = e.target.dataset.groupid;
      const group = state.mappingGroups.find(g => g.groupId === groupId);
      if (!group) return;
      group.canonical = e.target.value.trim();

      PROJECTS.forEach(({ key }) => {
        group.itemsByProject[key].forEach(rawItem => {
          state.mappings[makeRawKey(key, rawItem)] = group.canonical;
        });
      });
    });
  });
}

dom.mappingSearch.addEventListener("input", renderMappingGroupList);

dom.btnAutofill.addEventListener("click", () => {
  state.mappingGroups.forEach(group => {
    group.canonical = chooseBestCanonical(group);
    PROJECTS.forEach(({ key }) => {
      group.itemsByProject[key].forEach(rawItem => {
        state.mappings[makeRawKey(key, rawItem)] = group.canonical;
      });
    });
  });

  renderMappingGroupList();
  dom.mappingStatus.textContent = "유사 아이템 그룹 기준으로 통일명을 자동 채웠습니다.";
});

dom.btnApplyMapping.addEventListener("click", () => {
  if (!state.parsedReady) {
    dom.mappingStatus.textContent = "먼저 EXCEL 파일을 읽어 주세요.";
    return;
  }

  state.mappingGroups.forEach(group => {
    if (!group.canonical || !group.canonical.trim()) {
      group.canonical = chooseBestCanonical(group);
    }
    PROJECTS.forEach(({ key }) => {
      group.itemsByProject[key].forEach(rawItem => {
        state.mappings[makeRawKey(key, rawItem)] = group.canonical;
      });
    });
  });

  state.mappedReady = true;
  populateCanonicalDatalist();
  populateDongFilter();
  renderSummaryCards();
  renderCompareCards();

  dom.mappingStatus.textContent = "아이템 통일명이 적용되었습니다.";
  setActiveTab("compare");
});

function populateCanonicalDatalist() {
  const allNames = uniqueSort(state.mappingGroups.map(g => g.canonical).filter(Boolean));
  dom.canonicalOptions.innerHTML = allNames.map(name =>
    `<option value="${escapeHtmlAttr(name)}"></option>`
  ).join("");
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

dom.btnRenderCompare.addEventListener("click", () => {
  renderSummaryCards();
  renderCompareCards();
});
dom.filterDong.addEventListener("change", () => {
  renderSummaryCards();
  renderCompareCards();
});
dom.filterItem.addEventListener("input", renderCompareCards);
dom.filterMode.addEventListener("change", renderCompareCards);

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

function renderCompareCards() {
  if (!state.mappedReady) {
    dom.compareCardList.innerHTML = `<div class="empty-box">아이템 통일명을 먼저 적용해 주세요.</div>`;
    dom.compareStatus.textContent = "";
    return;
  }

  const dong = dom.filterDong.value;
  if (!dong) {
    dom.compareCardList.innerHTML = `<div class="empty-box">동을 선택해 주세요.</div>`;
    dom.compareStatus.textContent = "";
    return;
  }

  const keyword = dom.filterItem.value.trim().toLowerCase();
  const ratioMode = dom.filterMode.value;
  const viewData = buildUnifiedDataByDong(dong);
  const floors = viewData.floors;

  let items = Object.keys(viewData.items);
  if (keyword) items = items.filter(item => item.toLowerCase().includes(keyword));
  items = uniqueSort(items);

  if (!items.length) {
    dom.compareCardList.innerHTML = `<div class="empty-box">조건에 맞는 아이템이 없습니다.</div>`;
    dom.compareStatus.textContent = "";
    return;
  }

  dom.compareCardList.innerHTML = items.map(item => renderCompareCard(item, floors, viewData, ratioMode)).join("");
  dom.compareStatus.textContent = `동 ${dong} 기준 / 아이템 ${items.length}개 / 층 ${floors.length}개`;
}

function renderCompareCard(item, floors, viewData, ratioMode) {
  const values = viewData.items[item];

  let ratioValues = floors.map(f => values[f]?.ratio).filter(v => typeof v === "number");
  const ratioAvg = ratioValues.length ? average(ratioValues) : null;

  return `
    <div class="compare-card">
      <div class="compare-card__head">
        <div>
          <div class="compare-card__title">${escapeHtml(item)}</div>
          <div class="compare-card__meta">층별 수량 / 평균 / 비율 비교</div>
        </div>
        <div class="compare-card__meta">평균 비율 ${ratioAvg === null ? "-" : `${ratioAvg.toFixed(1)}%`}</div>
      </div>

      <div class="compare-card__body">
        <table class="compare-matrix">
          <thead>
            <tr>
              <th class="label-col">구분</th>
              ${floors.map(f => `<th>${escapeHtml(f)}</th>`).join("")}
            </tr>
          </thead>
          <tbody>
            <tr class="row-current">
              <td class="label-col">현재 프로젝트</td>
              ${floors.map(f => `<td>${formatValue(values[f]?.current)}</td>`).join("")}
            </tr>
            <tr class="row-a">
              <td class="label-col">유사 A</td>
              ${floors.map(f => `<td>${formatValue(values[f]?.a)}</td>`).join("")}
            </tr>
            <tr class="row-b">
              <td class="label-col">유사 B</td>
              ${floors.map(f => `<td>${formatValue(values[f]?.b)}</td>`).join("")}
            </tr>
            <tr class="row-c">
              <td class="label-col">유사 C</td>
              ${floors.map(f => `<td>${formatValue(values[f]?.c)}</td>`).join("")}
            </tr>
            <tr class="row-avg">
              <td class="label-col">유사 3개 평균</td>
              ${floors.map(f => `<td>${formatValue(values[f]?.avg)}</td>`).join("")}
            </tr>
            <tr class="row-ratio">
              <td class="label-col">현재 / 평균 비율(%)</td>
              ${floors.map(f => {
                const ratio = values[f]?.ratio ?? null;
                const cls = ratioClass(ratio, ratioMode);
                return `<td class="${cls}">${formatRatio(ratio)}</td>`;
              }).join("")}
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  `;
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

function normalizeItemName(v) {
  return String(v ?? "").replace(/\s+/g, " ").trim();
}

function normalizeFloorCell(v) {
  const s = String(v ?? "").trim();
  if (!s) return "";
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
