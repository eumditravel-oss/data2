"use strict";

const PROJECTS = [
  { key: "current", name: "현재" }, { key: "a", name: "유사A" }, 
  { key: "b", name: "유사B" }, { key: "c", name: "유사C" }
];
const CATEGORIES = ["콘크리트", "거푸집", "철근", "잡/기타"];

const state = {
  projects: { current: emptyState(), a: emptyState(), b: emptyState(), c: emptyState() },
  mappingGroups: [],
  selectedGroupIds: new Set(),
  selectedSplitKeys: new Set(),
  mappings: {},
  mappedReady: false
};

function emptyState() { return { files: [], rawItems: [], dongs: [], floors: [], data: {}, dongTypes: {} }; }

const $ = (id) => document.getElementById(id);
const dom = {
  tabs: document.querySelectorAll(".tab"),
  tabPanels: { upload: $("tab-upload"), mapping: $("tab-mapping"), compare: $("tab-compare") },
  fileInputs: { current: $("file-current"), a: $("file-a"), b: $("file-b"), c: $("file-c") },
  fileLists: { current: $("list-current"), a: $("list-a"), b: $("list-b"), c: $("list-c") },
  btnParse: $("btn-parse"),
  mappingGroupList: $("mapping-group-list"),
  filterDong: $("filter-dong"),
  filterCategory: $("filter-category"),
  filterType: $("filter-type"),
  compareCardList: $("compare-card-list")
};

/* 대분류 자동 판단 */
function determineDongType(dongName) {
  const n = dongName.toUpperCase();
  if (n.includes("PIT")) return "PIT";
  if (n.includes("주차장")) return "주차장";
  if (n.includes("경비") || n.includes("사무소") || n.includes("상가") || n.includes("동")) {
     if (/\d{3,4}동/.test(n)) return "APT";
     return "부속동";
  }
  return "APT";
}

/* 중분류 자동 판단 */
function determineCategory(name) {
  const s = String(name).toUpperCase();
  if (s.includes("H") || s.includes("D")) return "철근";
  if (s.includes("MPA") || /^\d+$/.test(s) || /\d+-\d+-\d+/.test(s)) return "콘크리트";
  if (/[가-힣]/.test(s)) return "거푸집";
  return "잡/기타";
}

/* 파일 업로드 시각화 */
PROJECTS.forEach(({key}) => {
  dom.fileInputs[key].onchange = (e) => {
    state.projects[key].files = [...e.target.files];
    dom.fileLists[key].innerHTML = state.projects[key].files.map(f => `<span class="file-chip">${f.name}</span>`).join("");
  };
});

/* 엑셀 파싱 및 데이터 수집 */
dom.btnParse.onclick = async () => {
  for (const {key} of PROJECTS) {
    const pState = emptyState();
    for (const file of state.projects[key].files) {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array" });
      wb.SheetNames.forEach(sn => {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "" });
        parseSheet(rows, pState);
      });
    }
    state.projects[key] = pState;
  }
  autoGroup();
  renderMapping();
  dom.tabs[1].click();
};

function parseSheet(rows, pState) {
  let dong = "", floor = "", prevF = null, sameCnt = 0;
  const r3 = rows[2] || [], r4 = rows[3] || [];
  for (let r = 4; r < rows.length; r++) {
    const txt = rows[r].join("|");
    const m = txt.match(/\[([^\]]+)\]/);
    if (m) {
      dong = m[1].trim();
      if (!pState.dongs.includes(dong)) {
        pState.dongs.push(dong);
        pState.dongTypes[dong] = determineDongType(dong); // 대분류 저장
      }
      pState.data[dong] = pState.data[dong] || {};
      floor = ""; sameCnt = 0; continue;
    }
    if (!dong) continue;
    const fRaw = String(rows[r][0]).trim();
    if (fRaw !== "") {
      floor = /^\d+$/.test(fRaw) ? fRaw + "F" : fRaw;
      if (!pState.floors.includes(floor)) pState.floors.push(floor);
      sameCnt = (prevF === floor) ? sameCnt + 1 : 1;
      prevF = floor;
    } else if (floor) sameCnt++;
    const head = (sameCnt % 2 === 1) ? r3 : r4;
    for (let c = 1; c < rows[r].length; c++) {
      const item = String(head[c] || "").trim();
      const val = parseFloat(String(rows[r][c]).replace(/,/g, ""));
      if (!item || isNaN(val)) continue;
      if (!pState.rawItems.includes(item)) pState.rawItems.push(item);
      pState.data[dong][item] = pState.data[dong][item] || {};
      pState.data[dong][item][floor] = (pState.data[dong][item][floor] || 0) + val;
    }
  }
}

/* 그룹화 및 매핑 UI (편집 도구 포함) */
function autoGroup() {
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

function renderMapping() {
  const q = $("mapping-search").value.toLowerCase();
  dom.mappingGroupList.innerHTML = state.mappingGroups
    .filter(g => g.canonical.toLowerCase().includes(q))
    .map(g => `
      <div class="mapping-group-card ${state.selectedGroupIds.has(g.groupId) ? 'is-selected' : ''}">
        <div class="mapping-group-card__top">
          <div class="mapping-group-card__left">
            <input type="checkbox" onchange="toggleGroupSelection('${g.groupId}', this.checked)" ${state.selectedGroupIds.has(g.groupId) ? 'checked' : ''} />
            <input class="group-canonical-input" value="${g.canonical}" oninput="updateGroupName('${g.groupId}', this.value)" />
          </div>
          <div class="mapping-group-card__right">
            <select onchange="updateGroupCat('${g.groupId}', this.value)">
              ${CATEGORIES.map(c => `<option value="${c}" ${g.category === c ? 'selected' : ''}>${c}</option>`).join("")}
            </select>
          </div>
        </div>
        <div class="mapping-project-grid">
          ${PROJECTS.map(p => `
            <div class="mapping-project-col">
              <div class="mapping-project-col__head">${p.name}</div>
              <div class="mapping-project-col__body">
                ${g.items[p.key].map(i => `
                  <label class="mapping-item-chip">
                    <input type="checkbox" onchange="toggleItemSelection('${g.groupId}','${p.key}','${i}', this.checked)" />
                    <span>${i}</span>
                  </label>
                `).join("")}
              </div>
            </div>
          `).join("")}
        </div>
      </div>
    `).join("");
}

/* 편집 핸들러 */
window.toggleGroupSelection = (id, checked) => { checked ? state.selectedGroupIds.add(id) : state.selectedGroupIds.delete(id); renderMapping(); };
window.updateGroupName = (id, val) => { state.mappingGroups.find(x => x.groupId === id).canonical = val; };
window.updateGroupCat = (id, val) => { state.mappingGroups.find(x => x.groupId === id).category = val; };
window.toggleItemSelection = (gid, pkey, item, checked) => {
  const key = `${gid}|${pkey}|${item}`;
  checked ? state.selectedSplitKeys.add(key) : state.selectedSplitKeys.delete(key);
};

/* 병합/분리 기능 */
$("btn-merge-groups").onclick = () => {
  if (state.selectedGroupIds.size < 2) return alert("병합할 그룹을 2개 이상 선택하세요.");
  const selected = state.mappingGroups.filter(g => state.selectedGroupIds.has(g.groupId));
  const newGroup = {
    groupId: Math.random().toString(36).substr(2, 9),
    canonical: selected[0].canonical,
    category: selected[0].category,
    items: { current: [], a: [], b: [], c: [] }
  };
  selected.forEach(g => PROJECTS.forEach(p => newGroup.items[p.key].push(...g.items[p.key])));
  state.mappingGroups = state.mappingGroups.filter(g => !state.selectedGroupIds.has(g.groupId));
  state.mappingGroups.push(newGroup);
  state.selectedGroupIds.clear();
  renderMapping();
};

$("btn-split-items").onclick = () => {
  if (state.selectedSplitKeys.size === 0) return alert("분리할 아이템을 체크하세요.");
  state.selectedSplitKeys.forEach(key => {
    const [gid, pkey, item] = key.split("|");
    const group = state.mappingGroups.find(g => g.groupId === gid);
    group.items[pkey] = group.items[pkey].filter(i => i !== item);
    state.mappingGroups.push({
      groupId: Math.random().toString(36).substr(2, 9),
      canonical: item,
      category: group.category,
      items: { current: [], a: [], b: [], c: [] }
    });
    state.mappingGroups[state.mappingGroups.length-1].items[pkey].push(item);
  });
  state.selectedSplitKeys.clear();
  renderMapping();
};

/* 최종 적용 및 비교표 */
$("btn-apply-mapping").onclick = () => {
  state.mappings = {};
  state.mappingGroups.forEach(g => {
    PROJECTS.forEach(p => g.items[p.key].forEach(raw => {
      state.mappings[`${p.key}::${raw}`] = { canonical: g.canonical, category: g.category };
    }));
  });
  state.mappedReady = true;
  updateDongFilters();
  dom.tabs[2].click();
};

function updateDongFilters() {
  const type = dom.filterType.value;
  const dongs = [];
  PROJECTS.forEach(p => {
    state.projects[p.key].dongs.forEach(d => {
      if ((type === "all" || state.projects[p.key].dongTypes[d] === type) && !dongs.includes(d)) dongs.push(d);
    });
  });
  dom.filterDong.innerHTML = dongs.sort().map(d => `<option value="${d}">${d}</option>`).join("");
  renderCompare();
}

dom.filterType.onchange = updateDongFilters;
dom.filterDong.onchange = renderCompare;
dom.filterCategory.onchange = renderCompare;

function renderCompare() {
  if (!state.mappedReady) return;
  const dong = dom.filterDong.value;
  const catFilter = dom.filterCategory.value;
  const unified = { floors: [], items: {} };

  PROJECTS.forEach(p => {
    const dData = state.projects[p.key].data[dong] || {};
    for (const raw in dData) {
      const m = state.mappings[`${p.key}::${raw}`];
      if (!m || (catFilter !== "all" && m.category !== catFilter)) continue;
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
        <table class="compare-matrix">
          <thead><tr><th>구분</th>${unified.floors.map(f=>`<th>${f}</th>`).join("")}</tr></thead>
          <tbody>
            <tr class="row-current"><td>현재</td>${unified.floors.map(f=>`<td>${(vals[f]?.current||0).toLocaleString()}</td>`).join("")}</tr>
            <tr><td>유사A</td>${unified.floors.map(f=>`<td>${(vals[f]?.a||0).toLocaleString()}</td>`).join("")}</tr>
            <tr><td>유사B</td>${unified.floors.map(f=>`<td>${(vals[f]?.b||0).toLocaleString()}</td>`).join("")}</tr>
            <tr><td>유사C</td>${unified.floors.map(f=>`<td>${(vals[f]?.c||0).toLocaleString()}</td>`).join("")}</tr>
            <tr style="background:#f4f7fd; font-weight:bold;"><td>유사평균</td>${unified.floors.map(f=>{
              const avg = ((vals[f]?.a||0) + (vals[f]?.b||0) + (vals[f]?.c||0)) / 3;
              return `<td>${avg.toLocaleString(undefined, {maximumFractionDigits:1})}</td>`;
            }).join("")}</tr>
          </tbody>
        </table>
      </div>`;
  }).join("") || "데이터 없음";
}

dom.tabs.forEach(tab => {
  tab.onclick = () => {
    dom.tabs.forEach(b => b.classList.remove("is-active"));
    tab.classList.add("is-active");
    Object.entries(dom.tabPanels).forEach(([k, p]) => p.classList.toggle("is-active", k === tab.dataset.tab));
  };
});
