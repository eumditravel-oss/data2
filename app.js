"use strict";

const PROJECTS = [
  { key: "current", name: "현재 프로젝트" },
  { key: "a", name: "A 프로젝트" },
  { key: "b", name: "B 프로젝트" },
  { key: "c", name: "C 프로젝트" }
];
const CATEGORIES = ["콘크리트", "거푸집", "철근", "잡/기타"];

const state = {
  projects: { current: emptyState(), a: emptyState(), b: emptyState(), c: emptyState() },
  mappingGroups: [],
  mappings: {},
  mappedReady: false
};

function emptyState() { return { files: [], rawItems: [], dongs: [], floors: [], data: {} }; }

const $ = (id) => document.getElementById(id);
const dom = {
  tabs: document.querySelectorAll(".tab"),
  tabPanels: { upload: $("tab-upload"), mapping: $("tab-mapping"), compare: $("tab-compare") },
  fileInputs: { current: $("file-current"), a: $("file-a"), b: $("file-b"), c: $("file-c") },
  fileNames: { current: $("name-current"), a: $("name-a"), b: $("name-b"), c: $("name-c") },
  fileLists: { current: $("list-current"), a: $("list-a"), b: $("list-b"), c: $("list-c") },
  btnParse: $("btn-parse"),
  uploadStatus: $("upload-status"),
  mappingSearch: $("mapping-search"),
  btnApplyMapping: $("btn-apply-mapping"),
  mappingGroupList: $("mapping-group-list"),
  filterDong: $("filter-dong"),
  filterCategory: $("filter-category"),
  filterItem: $("filter-item"),
  compareCardList: $("compare-card-list")
};

/* 파일 업로드 시 칩 표시 */
function renderFileUI(key) {
  const files = state.projects[key].files;
  dom.fileNames[key].textContent = files.length ? `${files.length}개 파일` : "선택된 파일 없음";
  dom.fileLists[key].innerHTML = files.map(f => `<span class="file-chip">${f.name}</span>`).join("");
}

PROJECTS.forEach(({key}) => {
  dom.fileInputs[key].onchange = (e) => {
    state.projects[key].files = [...e.target.files];
    renderFileUI(key);
  };
});

/* 자동 분류 */
function determineCategory(name) {
  const s = String(name).toUpperCase();
  if (s.includes("H") || s.includes("D")) return "철근";
  if (s.includes("MPA") || /^\d+$/.test(s) || /\d+-\d+-\d+/.test(s)) return "콘크리트";
  if (/[가-힣]/.test(s)) return "거푸집";
  return "잡/기타";
}

/* 탭 제어 */
dom.tabs.forEach(tab => {
  tab.onclick = () => {
    dom.tabs.forEach(b => b.classList.remove("is-active"));
    tab.classList.add("is-active");
    Object.entries(dom.tabPanels).forEach(([k, p]) => p.classList.toggle("is-active", k === tab.dataset.tab));
  };
});

/* 파싱 */
dom.btnParse.onclick = async () => {
  dom.uploadStatus.textContent = "분석 중...";
  for (const {key} of PROJECTS) {
    const pState = emptyState();
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
  buildGroups();
  renderGroups();
  dom.uploadStatus.textContent = "분석 완료!";
  dom.tabs[1].click();
};

function parseSheet(rows) {
  const res = emptyState();
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

function buildGroups() {
  const grouped = new Map();
  PROJECTS.forEach(p => {
    state.projects[p.key].rawItems.forEach(raw => {
      const sig = raw.replace(/\s+/g, "").toUpperCase();
      if (!grouped.has(sig)) {
        grouped.set(sig, {
          groupId: Math.random().toString(36).substr(2, 9),
          canonical: raw,
          category: determineCategory(raw),
          items: { current: [], a: [], b: [], c: [] }
        });
      }
      const g = grouped.get(sig);
      if (!g.items[p.key].includes(raw)) g.items[p.key].push(raw);
    });
  });
  state.mappingGroups = [...grouped.values()];
}

function renderGroups() {
  const q = dom.mappingSearch.value.toLowerCase();
  dom.mappingGroupList.innerHTML = state.mappingGroups
    .filter(g => g.canonical.toLowerCase().includes(q))
    .map(g => `
      <div class="mapping-group-card">
        <div class="mapping-group-card__top">
          <div class="mapping-group-card__left"><strong>${g.canonical}</strong></div>
          <div class="mapping-group-card__right">
            <select onchange="updateCat('${g.groupId}', this.value)">
              ${CATEGORIES.map(c => `<option value="${c}" ${g.category === c ? 'selected' : ''}>${c}</option>`).join("")}
            </select>
          </div>
        </div>
        <div class="mapping-project-grid">
          ${PROJECTS.map(p => `
            <div class="mapping-project-col ${p.key}">
              <div class="mapping-project-col__head">${p.name}</div>
              <div class="mapping-project-col__body">${g.items[p.key].join(", ") || '-'}</div>
            </div>`).join("")}
        </div>
      </div>`).join("");
}

window.updateCat = (id, val) => { state.mappingGroups.find(g => g.groupId === id).category = val; };
dom.mappingSearch.oninput = renderGroups;

dom.btnApplyMapping.onclick = () => {
  state.mappings = {};
  state.mappingGroups.forEach(g => {
    PROJECTS.forEach(p => { g.items[p.key].forEach(raw => {
      state.mappings[`${p.key}::${raw}`] = { canonical: g.canonical, category: g.category };
    }); });
  });
  state.mappedReady = true;
  const dongs = [...new Set(PROJECTS.flatMap(p => state.projects[p.key].dongs))].sort();
  dom.filterDong.innerHTML = dongs.map(d => `<option value="${d}">${d}</option>`).join("");
  renderCompare();
  dom.tabs[2].click();
};

[dom.filterDong, dom.filterCategory].forEach(el => el.onchange = renderCompare);
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
              <tr class="row-current"><td>현재 프로젝트</td>${unified.floors.map(f=>`<td>${(vals[f]?.current||0).toLocaleString()}</td>`).join("")}</tr>
              <tr><td>유사 A</td>${unified.floors.map(f=>`<td>${(vals[f]?.a||0).toLocaleString()}</td>`).join("")}</tr>
              <tr><td>유사 B</td>${unified.floors.map(f=>`<td>${(vals[f]?.b||0).toLocaleString()}</td>`).join("")}</tr>
              <tr><td>유사 C</td>${unified.floors.map(f=>`<td>${(vals[f]?.c||0).toLocaleString()}</td>`).join("")}</tr>
              <tr style="background:#f4f7fd; font-weight:bold;"><td>유사 평균</td>${unified.floors.map(f=>{
                const avg = ((vals[f]?.a||0) + (vals[f]?.b||0) + (vals[f]?.c||0)) / 3;
                return `<td>${avg.toLocaleString(undefined, {maximumFractionDigits:1})}</td>`;
              }).join("")}</tr>
              <tr style="font-weight:bold;"><td>비율(%)</td>${unified.floors.map(f=>{
                const avg = ((vals[f]?.a||0) + (vals[f]?.b||0) + (vals[f]?.c||0)) / 3;
                const r = avg > 0 ? (vals[f].current / avg * 100) : 0;
                let cls = (r >= 110) ? "ratio-high" : (r > 0 && r <= 90) ? "ratio-low" : "";
                return `<td class="${cls}">${r ? r.toFixed(1)+'%' : '-'}</td>`;
              }).join("")}</tr>
            </tbody>
          </table>
        </div>
      </div>`;
  }).join("") || "검색 결과 없음";
}

$ ("btn-reset").onclick = () => location.reload();
