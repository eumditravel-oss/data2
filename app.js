"use strict";

const PROJECTS = [
  { key: "current", name: "현재 프로젝트" },
  { key: "a", name: "A 프로젝트" },
  { key: "b", name: "B 프로젝트" },
  { key: "c", name: "C 프로젝트" }
];
const CATEGORIES = ["콘크리트", "거푸집", "철근", "잡/기타"];

const state = {
  projects: { current: createEmpty(), a: createEmpty(), b: createEmpty(), c: createEmpty() },
  mappingGroups: [],
  mappings: {}, // { key: { canonical, category } }
  parsedReady: false,
  mappedReady: false
};

function createEmpty() { return { files: [], rawItems: [], dongs: [], floors: [], data: {} }; }

// 자동 분류 함수 (요청하신 기준)
function determineCategory(name) {
  const s = String(name).toUpperCase();
  if (s.includes("H") || s.includes("D")) return "철근";
  if (s.includes("MPA") || /^\d+$/.test(s) || /\d+-\d+-\d+/.test(s)) return "콘크리트";
  if (/[가-힣]/.test(s)) return "거푸집";
  return "잡/기타";
}

const $ = (id) => document.getElementById(id);
const dom = {
  tabs: document.querySelectorAll(".tab"),
  tabPanels: { upload: $("tab-upload"), mapping: $("tab-mapping"), compare: $("tab-compare") },
  fileInputs: { current: $("file-current"), a: $("file-a"), b: $("file-b"), c: $("file-c") },
  fileNames: { current: $("name-current"), a: $("name-a"), b: $("name-b"), c: $("name-c") },
  btnParse: $("btn-parse"),
  uploadStatus: $("upload-status"),
  mappingSearch: $("mapping-search"),
  btnApplyMapping: $("btn-apply-mapping"),
  mappingGroupList: $("mapping-group-list"),
  filterDong: $("filter-dong"),
  filterCategory: $("filter-category"),
  filterItem: $("filter-item"),
  filterMode: $("filter-mode"),
  summaryCards: $("summary-cards"),
  compareCardList: $("compare-card-list"),
  compareStatus: $("compare-status")
};

// --- 기존 탭 및 파일 UI 로직 유지 ---
dom.tabs.forEach(tab => {
  tab.onclick = () => {
    dom.tabs.forEach(b => b.classList.remove("is-active"));
    tab.classList.add("is-active");
    Object.entries(dom.tabPanels).forEach(([k, p]) => p.classList.toggle("is-active", k === tab.dataset.tab));
  };
});

PROJECTS.forEach(({key}) => {
  dom.fileInputs[key].onchange = (e) => {
    state.projects[key].files = [...e.target.files];
    dom.fileNames[key].textContent = `${e.target.files.length}개 파일 선택됨`;
  };
});

// --- 기존 파싱 로직 유지 ---
dom.btnParse.onclick = async () => {
  dom.uploadStatus.textContent = "엑셀 분석 중...";
  for (const {key} of PROJECTS) {
    const pState = createEmpty();
    for (const file of state.projects[key].files) {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array" });
      wb.SheetNames.forEach(sn => {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "" });
        const parsed = parseSheet(rows);
        mergeData(pState, parsed);
      });
    }
    state.projects[key] = pState;
  }
  buildMappingGroups();
  renderMappingGroups();
  dom.uploadStatus.textContent = "분석 완료!";
  dom.tabs[1].click();
};

function parseSheet(rows) {
  const res = createEmpty();
  let dong = "", floor = "", prevF = null, sameCnt = 0;
  const r3 = rows[2] || [], r4 = rows[3] || [];
  for (let r = 4; r < rows.length; r++) {
    const txt = rows[r].join("|");
    const m = txt.match(/\[([^\]]+)\]/);
    if (m) {
      dong = m[1].trim();
      if (!res.dongs.includes(dong)) res.dongs.push(dong);
      res.data[dong] = {}; floor = ""; sameCnt = 0; continue;
    }
    if (!dong) continue;
    const fRaw = String(rows[r][0]).trim();
    if (fRaw !== "") {
      floor = /^\d+$/.test(fRaw) ? fRaw + "F" : fRaw;
      if (!res.floors.includes(floor)) res.floors.push(floor);
      sameCnt = (prevF === floor) ? sameCnt + 1 : 1;
      prevF = floor;
    } else if (floor) sameCnt++;
    else continue;
    const head = (sameCnt % 2 === 1) ? r3 : r4;
    for (let c = 1; c < rows[r].length; c++) {
      const item = String(head[c] || "").trim();
      const val = parseFloat(String(rows[r][c]).replace(/,/g, ""));
      if (!item || isNaN(val)) continue;
      if (!res.rawItems.includes(item)) res.rawItems.push(item);
      if (!res.data[dong][item]) res.data[dong][item] = {};
      res.data[dong][item][floor] = (res.data[dong][item][floor] || 0) + val;
    }
  }
  return res;
}

function mergeData(target, source) {
  source.rawItems.forEach(i => !target.rawItems.includes(i) && target.rawItems.push(i));
  source.dongs.forEach(d => !target.dongs.includes(d) && target.dongs.push(d));
  source.floors.forEach(f => !target.floors.includes(f) && target.floors.push(f));
  for (const d in source.data) {
    if (!target.data[d]) target.data[d] = {};
    for (const i in source.data[d]) {
      if (!target.data[d][i]) target.data[d][i] = {};
      for (const f in source.data[d][i]) {
        target.data[d][i][f] = (target.data[d][i][f] || 0) + source.data[d][i][f];
      }
    }
  }
}

// --- 그룹화 및 렌더링 (중분류 select 추가) ---
function buildMappingGroups() {
  const grouped = new Map();
  PROJECTS.forEach(p => {
    state.projects[p.key].rawItems.forEach(raw => {
      const sig = raw.replace(/\s+/g, "").toUpperCase();
      if (!grouped.has(sig)) {
        grouped.set(sig, {
          groupId: Math.random().toString(36).substr(2, 9),
          canonical: raw,
          category: determineCategory(raw), // 자동 분류
          itemsByProject: { current: [], a: [], b: [], c: [] }
        });
      }
      grouped.get(sig).itemsByProject[p.key].push(raw);
    });
  });
  state.mappingGroups = [...grouped.values()];
}

function renderMappingGroups() {
  const q = dom.mappingSearch.value.toLowerCase();
  dom.mappingGroupList.innerHTML = state.mappingGroups
    .filter(g => g.canonical.toLowerCase().includes(q))
    .map(group => `
      <div class="mapping-group-card">
        <div class="mapping-group-card__top">
          <div class="mapping-group-card__left">
            <div class="mapping-group-card__title">${group.canonical}</div>
          </div>
          <div class="mapping-group-card__right">
            <div class="mapping-canonical">
              <label>중분류 / 통일명</label>
              <div style="display:flex; gap:5px;">
                <select class="group-category-select" onchange="updateCat('${group.groupId}', this.value)">
                  ${CATEGORIES.map(c => `<option value="${c}" ${group.category === c ? 'selected' : ''}>${c}</option>`).join("")}
                </select>
                <input class="group-canonical-input" value="${group.canonical}" oninput="updateName('${group.groupId}', this.value)" style="flex:1" />
              </div>
            </div>
          </div>
        </div>
        <div class="mapping-project-grid">
          ${PROJECTS.map(p => `
            <div class="mapping-project-col ${p.key}">
              <div class="mapping-project-col__head">${p.name}</div>
              <div class="mapping-project-col__body">
                ${group.itemsByProject[p.key].map(i => `<span class="mapping-item-chip">${i}</span>`).join("") || '-'}
              </div>
            </div>
          `).join("")}
        </div>
      </div>
    `).join("");
}

window.updateCat = (id, val) => { state.mappingGroups.find(g => g.groupId === id).category = val; };
window.updateName = (id, val) => { state.mappingGroups.find(g => g.groupId === id).canonical = val; };
dom.mappingSearch.oninput = renderMappingGroups;

// --- 비교표 적용 및 렌더링 (필터 로직 추가) ---
dom.btnApplyMapping.onclick = () => {
  state.mappings = {};
  state.mappingGroups.forEach(g => {
    PROJECTS.forEach(p => {
      g.itemsByProject[p.key].forEach(raw => {
        state.mappings[`${p.key}::${raw}`] = { canonical: g.canonical, category: g.category };
      });
    });
  });
  state.mappedReady = true;
  const dongs = [...new Set(PROJECTS.flatMap(p => state.projects[p.key].dongs))].sort();
  dom.filterDong.innerHTML = dongs.map(d => `<option value="${d}">${d}</option>`).join("");
  renderCompare();
  dom.tabs[2].click();
};

[dom.filterDong, dom.filterCategory, dom.filterMode].forEach(el => el.onchange = renderCompare);
dom.filterItem.oninput = renderCompare;

function renderCompare() {
  if (!state.mappedReady) return;
  const dong = dom.filterDong.value;
  const cat = dom.filterCategory.value;
  const keyw = dom.filterItem.value.toLowerCase();
  
  const unified = { floors: [], items: {} };
  PROJECTS.forEach(p => {
    const dData = state.projects[p.key].data[dong] || {};
    for (const raw in dData) {
      const m = state.mappings[`${p.key}::${raw}`];
      if (!m || (cat !== 'all' && m.category !== cat) || (keyw && !m.canonical.toLowerCase().includes(keyw))) continue;
      
      if (!unified.items[m.canonical]) unified.items[m.canonical] = {};
      for (const f in dData[raw]) {
        if (!unified.floors.includes(f)) unified.floors.push(f);
        if (!unified.items[m.canonical][f]) unified.items[m.canonical][f] = { current:0, a:0, b:0, c:0 };
        unified.items[m.canonical][f][p.key] += dData[raw][f];
      }
    }
  });
  unified.floors.sort();

  dom.compareCardList.innerHTML = Object.keys(unified.items).map(name => {
    const vals = unified.items[name];
    return `
      <div class="compare-card">
        <div class="compare-card__head"><strong>${name}</strong></div>
        <div class="compare-card__body">
          <table class="compare-matrix">
            <thead><tr><th>구분</th>${unified.floors.map(f=>`<th>${f}</th>`).join("")}</tr></thead>
            <tbody>
              <tr><td>현재</td>${unified.floors.map(f=>`<td>${(vals[f]?.current||0).toLocaleString()}</td>`).join("")}</tr>
              <tr style="background:#f9f9f9;"><td>유사평균</td>${unified.floors.map(f=>{
                const avg = ((vals[f]?.a||0) + (vals[f]?.b||0) + (vals[f]?.c||0)) / 3;
                return `<td>${avg.toLocaleString(undefined, {maximumFractionDigits:1})}</td>`;
              }).join("")}</tr>
            </tbody>
          </table>
        </div>
      </div>`;
  }).join("") || "검색 결과 없음";
}
