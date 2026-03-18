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

/* 자동 분류 엔진  */
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

/* 파일 선택 */
PROJECTS.forEach(({key}) => {
  dom.fileInputs[key].onchange = (e) => {
    state.projects[key].files = [...e.target.files];
    dom.fileNames[key].textContent = `${e.target.files.length}개 파일 선택됨`;
  };
});

/* 파싱 로직  */
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
  dom.uploadStatus.textContent = "완료! 2번 탭으로 가세요.";
};

function parseSheet(rows) {
  const res = emptyState();
  let dong = "", floor = "", prevF = null, sameCnt = 0;
  const r3 = rows[2] || [], r4 = rows[3] || [];

  for (let r = 4; r < rows.length; r++) {
    const row = rows[r];
    const txt = row.join("|");
    const m = txt.match(/\[([^\]]+)\]/);
    if (m) {
      dong = m[1].trim();
      if (!res.dongs.includes(dong)) res.dongs.push(dong);
      res.data[dong] = {}; floor = ""; sameCnt = 0;
      continue;
    }
    if (!dong) continue;

    const fRaw = String(row[0]).trim();
    if (fRaw !== "") {
      floor = /^\d+$/.test(fRaw) ? fRaw + "F" : fRaw;
      if (!res.floors.includes(floor)) res.floors.push(floor);
      sameCnt = (prevF === floor) ? sameCnt + 1 : 1;
      prevF = floor;
    } else if (floor) sameCnt++;
    else continue;

    const head = (sameCnt % 2 === 1) ? r3 : r4;
    for (let c = 1; c < row.length; c++) {
      const item = String(head[c] || "").trim();
      const val = parseFloat(String(row[c]).replace(/,/g, ""));
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
      <div class="mapping-group-card" style="border:1px solid #ddd; margin-bottom:10px; padding:10px; border-radius:8px;">
        <div style="display:flex; justify-content:space-between; align-items:center;">
          <strong>${g.canonical}</strong>
          <select onchange="updateGroupCat('${g.groupId}', this.value)">
            ${CATEGORIES.map(c => `<option value="${c}" ${g.category === c ? "selected" : ""}>${c}</option>`).join("")}
          </select>
        </div>
        <div style="font-size:12px; color:#888; margin-top:5px;">
          추출된 원본: ${Object.values(g.items).flat().filter(Boolean).join(", ")}
        </div>
      </div>
    `).join("");
}

window.updateGroupCat = (gid, val) => {
  const g = state.mappingGroups.find(x => x.groupId === gid);
  if (g) g.category = val;
};

dom.mappingSearch.oninput = renderGroups;

/* 비교표 렌더링  */
dom.btnApplyMapping.onclick = () => {
  state.mappings = {};
  state.mappingGroups.forEach(g => {
    PROJECTS.forEach(p => {
      g.items[p.key].forEach(raw => {
        state.mappings[`${p.key}::${raw}`] = { canonical: g.canonical, category: g.category };
      });
    });
  });
  state.mappedReady = true;
  const allDongs = [...new Set(PROJECTS.flatMap(p => state.projects[p.key].dongs))].sort();
  dom.filterDong.innerHTML = allDongs.map(d => `<option value="${d}">${d}</option>`).join("");
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
      if (!m || (cat !== "all" && m.category !== cat) || (keyw && !m.canonical.toLowerCase().includes(keyw))) continue;
      
      if (!unified.items[m.canonical]) unified.items[m.canonical] = {};
      for (const f in dData[raw]) {
        if (!unified.floors.includes(f)) unified.floors.push(f);
        if (!unified.items[m.canonical][f]) unified.items[m.canonical][f] = { current:0, a:0, b:0, c:0 };
        unified.items[m.canonical][f][p.key] += dData[raw][f];
      }
    }
  });
  unified.floors.sort();

  dom.compareCardList.innerHTML = Object.keys(unified.items).map(name => `
    <div class="compare-card" style="border:1px solid #eee; margin-bottom:15px; padding:10px;">
      <h4>${name}</h4>
      <table style="width:100%; border-collapse:collapse; font-size:12px;">
        <thead><tr><th>구분</th>${unified.floors.map(f=>`<th>${f}</th>`).join("")}</tr></thead>
        <tbody>
          <tr><td>현재</td>${unified.floors.map(f=>`<td>${(unified.items[name][f]?.current || 0).toLocaleString()}</td>`).join("")}</tr>
          <tr style="background:#f9f9f9;"><td>유사평균</td>${unified.floors.map(f=>{
            const v = unified.items[name][f];
            const avg = ((v?.a||0) + (v?.b||0) + (v?.c||0)) / 3;
            return `<td>${avg.toLocaleString(undefined, {maximumFractionDigits:1})}</td>`;
          }).join("")}</tr>
        </tbody>
      </table>
    </div>
  `).join("") || "데이터 없음";
}
