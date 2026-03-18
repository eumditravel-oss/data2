"use strict";

/* =========================
   상태 및 상수
========================= */
const PROJECTS = [
  { key: "current", name: "현재 프로젝트" },
  { key: "a", name: "A 프로젝트" },
  { key: "b", name: "B 프로젝트" },
  { key: "c", name: "C 프로젝트" }
];

const CATEGORIES = ["콘크리트", "거푸집", "철근", "잡/기타"];

const state = {
  projects: {
    current: createEmptyProjectState(),
    a: createEmptyProjectState(),
    b: createEmptyProjectState(),
    c: createEmptyProjectState()
  },
  mappingGroups: [],
  mappings: {}, // { "projectKey::rawItem": { canonical, category } }
  selectedGroupIds: new Set(),
  selectedSplitItems: new Set(),
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
    current: $("file-current"), a: $("file-a"), b: $("file-b"), c: $("file-c")
  },
  fileNames: {
    current: $("name-current"), a: $("name-a"), b: $("name-b"), c: $("name-c")
  },
  fileLists: {
    current: $("list-current"), a: $("list-a"), b: $("list-b"), c: $("list-c")
  },
  btnParse: $("btn-parse"),
  btnReset: $("btn-reset"),
  uploadStatus: $("upload-status"),
  mappingSearch: $("mapping-search"),
  btnAutofill: $("btn-autofill"),
  btnMergeGroups: $("btn-merge-groups"),
  btnSplitSelected: $("btn-split-selected"),
  btnApplyMapping: $("btn-apply-mapping"),
  mappingGroupList: $("mapping-group-list"),
  mappingStatus: $("mapping-status"),
  canonicalOptions: $("canonical-options"),
  filterDong: $("filter-dong"),
  filterCategory: $("filter-category"), // 중분류 필터 추가
  filterItem: $("filter-item"),
  filterMode: $("filter-mode"),
  btnRenderCompare: $("btn-render-compare"),
  summaryCards: $("summary-cards"),
  compareCardList: $("compare-card-list"),
  compareStatus: $("compare-status")
};

/* =========================
   이벤트 바인딩 및 탭
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

PROJECTS.forEach(({ key }) => {
  dom.fileInputs[key].addEventListener("change", (e) => {
    const files = [...(e.target.files || [])];
    state.projects[key].files = files;
    renderFileUI(key);
  });
});

dom.btnReset.addEventListener("click", () => {
  location.reload(); // 단순화를 위한 리로드
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

    dom.uploadStatus.textContent = "추출 완료. [2. 아이템 통일 설정] 탭으로 이동합니다.";
    setActiveTab("mapping");
  } catch (err) {
    console.error(err);
    dom.uploadStatus.textContent = `오류 발생: ${err.message}`;
  }
});

async function parseExcelFile(file) {
  const buffer = await file.arrayBuffer();
  const wb = XLSX.read(buffer, { type: "array" });
  const aggregated = createEmptyProjectState();

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
    const startRow = merge.s.r, endRow = merge.e.r;
    const startCol = merge.s.c, endCol = merge.e.c;
    const val = (grid[startRow] && grid[startRow][startCol]) ?? "";
    for (let r = startRow; r <= endRow; r++) {
      if (!grid[r]) grid[r] = [];
      for (let c = startCol; c <= endCol; c++) {
        if (grid[r][c] === "" || grid[r][c] === undefined) grid[r][c] = val;
      }
    }
  }
  return grid;
}

function parseWorksheetRows(rows) {
  const result = createEmptyProjectState();
  if (!rows || rows.length < 5) return result;

  const row3 = rows[2] || [];
  const row4 = rows[3] || [];
  let currentDong = "", currentFloor = "", sameFloorCount = 0, previousFloor = null;

  for (let r = 4; r < rows.length; r++) {
    const row = rows[r] || [];
    const rowText = row.map(v => String(v ?? "").trim()).join(" | ");
    const dongMatch = rowText.match(/\[([^\]]+)\]/);

    if (dongMatch) {
      currentDong = dongMatch[1].trim();
      pushUnique(result.dongs, currentDong);
      if (!result.data[currentDong]) result.data[currentDong] = {};
      currentFloor = ""; sameFloorCount = 0; previousFloor = null;
      continue;
    }

    if (!currentDong) continue;

    const floorCandidate = normalizeFloorCell(row[0]);
    if (floorCandidate !== "") {
      currentFloor = floorCandidate;
      pushUnique(result.floors, currentFloor);
      if (previousFloor === currentFloor) sameFloorCount++;
      else { sameFloorCount = 1; previousFloor = currentFloor; }
    } else if (currentFloor) {
      sameFloorCount++;
    } else continue;

    const itemRowSource = sameFloorCount % 2 === 1 ? row3 : row4;

    for (let c = 1; c < row.length; c++) {
      const rawLabel = normalizeItemName(itemRowSource[c]);
      if (!rawLabel) continue;
      const val = toNumber(row[c]);
      if (val === null) continue;

      pushUnique(result.rawItems, rawLabel);
      if (!result.data[currentDong][rawLabel]) result.data[currentDong][rawLabel] = {};
      result.data[currentDong][rawLabel][currentFloor] = (result.data[currentDong][rawLabel][currentFloor] || 0) + val;
    }
  }
  return result;
}

function mergeParsedProject(target, parsed) {
  parsed.rawItems?.forEach(item => pushUnique(target.rawItems, item));
  parsed.dongs?.forEach(d => pushUnique(target.dongs, d));
  parsed.floors?.forEach(f => pushUnique(target.floors, f));
  for (const dong in parsed.data) {
    if (!target.data[dong]) target.data[dong] = {};
    for (const item in parsed.data[dong]) {
      if (!target.data[dong][item]) target.data[dong][item] = {};
      for (const floor in parsed.data[dong][item]) {
        target.data[dong][item][floor] = (target.data[dong][item][floor] || 0) + parsed.data[dong][item][floor];
      }
    }
  }
}

/* =========================
   자동 분류 및 그룹화 로직
========================= */
function determineCategory(name) {
  const s = String(name).toUpperCase().trim();
  // 1. 철근: H 또는 D 포함
  if (s.includes("H") || s.includes("D")) return "철근";
  // 2. 콘크리트: MPA 포함, 숫자만 있거나 강도 패턴(25-240-15)인 경우
  if (s.includes("MPA") || /^\d+$/.test(s) || /\d+-\d+-\d+/.test(s)) return "콘크리트";
  // 3. 거푸집/기타: 한글이 포함되어 있거나 그 외
  if (/[가-힣]/.test(s)) return "거푸집";
  return "잡/기타";
}

function buildMappingGroups() {
  const grouped = new Map();

  PROJECTS.forEach(({ key }) => {
    state.projects[key].rawItems.forEach(rawItem => {
      const signature = makeItemSignature(rawItem);
      if (!grouped.has(signature)) {
        const initialCat = determineCategory(rawItem);
        grouped.set(signature, {
          groupId: `grp_${makeUid()}`,
          signature,
          canonical: suggestCanonicalName(rawItem),
          category: initialCat,
          itemsByProject: { current: [], a: [], b: [], c: [] }
        });
      }
      grouped.get(signature).itemsByProject[key].push(rawItem);
    });
  });

  state.mappingGroups = [...grouped.values()].sort((a, b) => a.canonical.localeCompare(b.canonical, "ko"));
  rebuildMappingsFromGroups();
}

function rebuildMappingsFromGroups() {
  state.mappings = {};
  state.mappingGroups.forEach(group => {
    PROJECTS.forEach(({ key }) => {
      group.itemsByProject[key].forEach(rawItem => {
        state.mappings[makeRawKey(key, rawItem)] = {
          canonical: group.canonical,
          category: group.category
        };
      });
    });
  });
}

/* =========================
   매핑 UI 렌더링
========================= */
function renderMappingGroupList() {
  const q = dom.mappingSearch.value.trim().toLowerCase();
  const groups = state.mappingGroups.filter(g => {
    const text = (g.canonical + Object.values(g.itemsByProject).flat().join("")).toLowerCase();
    return !q || text.includes(q);
  });

  dom.mappingGroupList.innerHTML = groups.map(group => `
    <div class="mapping-group-card ${state.selectedGroupIds.has(group.groupId) ? "is-selected" : ""}">
      <div class="mapping-group-card__top">
        <div class="mapping-group-card__pick">
          <input type="checkbox" class="group-select-checkbox" data-groupid="${group.groupId}" ${state.selectedGroupIds.has(group.groupId) ? "checked" : ""} />
        </div>
        <div class="mapping-group-card__left">
          <div class="mapping-group-card__title">${escapeHtml(group.canonical)}</div>
          <div class="mapping-group-card__meta">분류: <strong>${group.category}</strong></div>
        </div>
        <div class="mapping-group-card__right">
          <div class="mapping-canonical">
            <div style="display:flex; gap:10px;">
              <div style="flex:1">
                <label>통일 아이템명</label>
                <input class="group-canonical-input" data-groupid="${group.groupId}" value="${escapeHtml(group.canonical)}" list="canonical-options" />
              </div>
              <div style="width:120px">
                <label>중분류 수정</label>
                <select class="group-category-select" data-groupid="${group.groupId}">
                  ${CATEGORIES.map(c => `<option value="${c}" ${group.category === c ? "selected" : ""}>${c}</option>`).join("")}
                </select>
              </div>
            </div>
          </div>
        </div>
      </div>
      <div class="mapping-project-grid">
        ${PROJECTS.map(p => renderMappingProjectColumn(group.groupId, p.key, p.name, group.itemsByProject[p.key])).join("")}
      </div>
    </div>
  `).join("");
  bindGroupControls();
}

function renderMappingProjectColumn(groupId, projectKey, title, items) {
  const body = items.length ? items.map(item => {
    const splitKey = `${groupId}::${projectKey}::${item}`;
    return `
      <label class="mapping-item-chip">
        <input type="checkbox" class="split-item-checkbox" data-splitkey="${splitKey}" ${state.selectedSplitItems.has(splitKey) ? "checked" : ""} />
        <span>${escapeHtml(item)}</span>
      </label>
    `;
  }).join("") : `<span class="mapping-item-chip is-empty">해당 없음</span>`;

  return `<div class="mapping-project-col ${projectKey}"><div class="mapping-project-col__head">${title}</div><div class="mapping-project-col__body">${body}</div></div>`;
}

function bindGroupControls() {
  document.querySelectorAll(".group-select-checkbox").forEach(el => {
    el.onclick = (e) => {
      const id = e.target.dataset.groupid;
      e.target.checked ? state.selectedGroupIds.add(id) : state.selectedGroupIds.delete(id);
      renderMappingGroupList();
    };
  });

  document.querySelectorAll(".group-canonical-input").forEach(el => {
    el.oninput = (e) => {
      const g = state.mappingGroups.find(x => x.groupId === e.target.dataset.groupid);
      if (g) g.canonical = e.target.value.trim();
    };
  });

  document.querySelectorAll(".group-category-select").forEach(el => {
    el.onchange = (e) => {
      const g = state.mappingGroups.find(x => x.groupId === e.target.dataset.groupid);
      if (g) {
        g.category = e.target.value;
        dom.mappingStatus.textContent = `[${g.canonical}]의 분류가 ${g.category}(으)로 변경되었습니다.`;
      }
    };
  });

  document.querySelectorAll(".split-item-checkbox").forEach(el => {
    el.onclick = (e) => {
      const key = e.target.dataset.splitkey;
      e.target.checked ? state.selectedSplitItems.add(key) : state.selectedSplitItems.delete(key);
    };
  });
}

/* =========================
   비교표 렌더링 및 필터링
========================= */
dom.btnApplyMapping.onclick = () => {
  rebuildMappingsFromGroups();
  state.mappedReady = true;
  populateDongFilter();
  renderCompareCards();
  setActiveTab("compare");
};

// 필터 변경 시 자동 갱신
[dom.filterDong, dom.filterCategory, dom.filterMode].forEach(el => {
  el.onchange = renderCompareCards;
});
dom.filterItem.oninput = renderCompareCards;

function renderCompareCards() {
  if (!state.mappedReady) return;
  const dong = dom.filterDong.value;
  const categoryFilter = dom.filterCategory.value;
  const keyword = dom.filterItem.value.toLowerCase();
  
  if (!dong) return;

  const viewData = buildUnifiedDataByDong(dong);
  let items = Object.keys(viewData.items);

  // 필터 1: 중분류
  if (categoryFilter !== "all") {
    items = items.filter(itemName => {
      const group = state.mappingGroups.find(g => g.canonical === itemName);
      return group && group.category === categoryFilter;
    });
  }

  // 필터 2: 검색어
  if (keyword) items = items.filter(item => item.toLowerCase().includes(keyword));

  dom.compareCardList.innerHTML = items.length ? items.map(item => renderCompareCard(item, viewData.floors, viewData)).join("") 
    : `<div class="empty-box">선택한 조건(분류: ${categoryFilter})에 맞는 데이터가 없습니다.</div>`;
}

function buildUnifiedDataByDong(dong) {
  const unified = { floors: [], items: {} };
  const allFloors = uniqueSort(PROJECTS.flatMap(({key}) => {
    const dData = state.projects[key].data[dong] || {};
    return Object.values(dData).flatMap(obj => Object.keys(obj));
  }), floorSorter);
  unified.floors = allFloors;

  PROJECTS.forEach(({key}) => {
    const dongData = state.projects[key].data[dong] || {};
    for (const rawItem in dongData) {
      const mapInfo = state.mappings[makeRawKey(key, rawItem)];
      if (!mapInfo) continue;
      const canonical = mapInfo.canonical;
      if (!unified.items[canonical]) unified.items[canonical] = {};
      
      allFloors.forEach(f => {
        if (!unified.items[canonical][f]) unified.items[canonical][f] = { current:null, a:null, b:null, c:null, avg:null, ratio:null };
        const val = dongData[rawItem][f];
        if (typeof val === "number") unified.items[canonical][f][key] = (unified.items[canonical][f][key] || 0) + val;
      });
    }
  });

  // 평균 및 비율 계산
  for (const item in unified.items) {
    allFloors.forEach(f => {
      const cell = unified.items[item][f];
      const sims = [cell.a, cell.b, cell.c].filter(v => v !== null);
      if (sims.length) cell.avg = sims.reduce((a,b)=>a+b,0) / sims.length;
      if (cell.current !== null && cell.avg) cell.ratio = (cell.current / cell.avg) * 100;
    });
  }
  return unified;
}

/* =========================
   나머지 유틸리티 함수 (기존과 동일)
========================= */
function makeRawKey(pKey, item) { return `${pKey}::${item}`; }
function makeItemSignature(name) { return String(name).replace(/\s+/g, "").toUpperCase(); }
function suggestCanonicalName(name) { return String(name).trim(); }
function normalizeItemName(v) { return String(v ?? "").trim(); }
function normalizeFloorCell(v) { 
  const s = String(v ?? "").trim();
  return /^\d+$/.test(s) ? s + "F" : s;
}
function toNumber(v) {
  const n = parseFloat(String(v).replace(/,/g, ""));
  return isNaN(n) ? null : n;
}
function pushUnique(arr, v) { if (v && !arr.includes(v)) arr.push(v); }
function uniqueSort(arr, sorter) { 
  const u = [...new Set(arr)]; 
  return sorter ? u.sort(sorter) : u.sort(); 
}
function dongSorter(a,b) { return (parseInt(a) || 0) - (parseInt(b) || 0); }
function floorSorter(a,b) { return (parseInt(a) || 0) - (parseInt(b) || 0); }
function escapeHtml(s) { return String(s).replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":"&#39;"}[m])); }
function makeUid() { return Math.random().toString(36).substr(2, 9); }
function renderFileUI(k) { dom.fileNames[k].textContent = `${state.projects[k].files.length}개 파일`; }
function populateDongFilter() {
  const dongs = uniqueSort(PROJECTS.flatMap(p => state.projects[p.key].dongs), dongSorter);
  dom.filterDong.innerHTML = dongs.map(d => `<option value="${d}">${d}</option>`).join("");
}

function renderCompareCard(item, floors, viewData) {
  const values = viewData.items[item];
  return `
    <div class="compare-card">
      <div class="compare-card__head"><strong>${escapeHtml(item)}</strong></div>
      <div class="compare-card__body">
        <table class="compare-matrix">
          <thead><tr><th>구분</th>${floors.map(f=>`<th>${f}</th>`).join("")}</tr></thead>
          <tbody>
            <tr class="row-current"><td>현재</td>${floors.map(f=>`<td>${formatVal(values[f].current)}</td>`).join("")}</tr>
            <tr class="row-avg"><td>유사평균</td>${floors.map(f=>`<td>${formatVal(values[f].avg)}</td>`).join("")}</tr>
            <tr class="row-ratio"><td>비율(%)</td>${floors.map(f=>`<td class="${getRatioClass(values[f].ratio)}">${values[f].ratio ? values[f].ratio.toFixed(1)+'%' : '-'}</td>`).join("")}</tr>
          </tbody>
        </table>
      </div>
    </div>`;
}
function formatVal(v) { return v === null ? "-" : v.toLocaleString(); }
function getRatioClass(r) { 
  if (!r) return "";
  if (r >= 110) return "ratio-high";
  if (r <= 90) return "ratio-low";
  return "";
}
