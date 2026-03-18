"use strict";

const PROJECTS = [
  { key: "current", name: "현재" }, { key: "a", name: "A" }, 
  { key: "b", name: "B" }, { key: "c", name: "C" }
];
const CATEGORIES = ["콘크리트", "거푸집", "철근", "잡/기타"];

const state = {
  projects: { current: emptyState(), a: emptyState(), b: emptyState(), c: emptyState() },
  dongMap: {}, // { "current::101동": "101동" }
  mappingGroups: [],
  selectedGroupIds: new Set()
};

function emptyState() { return { rawItems: [], dongs: [], data: {} }; }

const $ = (id) => document.getElementById(id);
const dom = {
  tabs: document.querySelectorAll(".tab"),
  subTabs: document.querySelectorAll(".sub-tab"),
  fileInputs: { current: $("file-current"), a: $("file-a"), b: $("file-b"), c: $("file-c") },
  fileLists: { current: $("list-current"), a: $("list-a"), b: $("list-b"), c: $("list-c") },
  btnParse: $("btn-parse"),
  dongList: $("dong-mapping-list"),
  itemList: $("mapping-group-list"),
  filterDong: $("filter-dong"),
  compareList: $("compare-card-list")
};

/* 1. 데이터 분석 및 동 추출 */
dom.btnParse.onclick = async () => {
  for (const p of PROJECTS) {
    const pState = state.projects[p.key];
    const files = Array.from($(`file-${p.key}`).files);
    for (const file of files) {
      const rows = XLSX.utils.sheet_to_json(XLSX.read(await file.arrayBuffer(), {type:'array'}).Sheets[XLSX.read(await file.arrayBuffer(), {type:'array'}).SheetNames[0]], {header:1, defval:""});
      parseData(rows, pState, p.key);
    }
  }
  initDongMapping();
  buildItemGroups();
  renderDongMapping();
  renderItemMapping();
  $("tab-mapping").click(); 
  dom.subTabs[0].click();
};

function parseData(rows, pState, pKey) {
  let dong = "";
  const r3 = rows[2] || [], r4 = rows[3] || [];
  for (let r = 4; r < rows.length; r++) {
    const m = rows[r].join("|").match(/\[([^\]]+)\]/);
    if (m) {
      dong = m[1].trim();
      if (!pState.dongs.includes(dong)) pState.dongs.push(dong);
      pState.data[dong] = pState.data[dong] || {};
      continue;
    }
    if (!dong) continue;
    const itemRow = (r % 2 === 1) ? r3 : r4; // 간략화된 2행 1세트 로직
    for (let c = 1; c < rows[r].length; c++) {
      const item = String(itemRow[c] || "").trim();
      const val = parseFloat(String(rows[r][c]).replace(/,/g,""));
      if (!item || isNaN(val)) continue;
      if (!pState.rawItems.includes(item)) pState.rawItems.push(item);
      pState.data[dong][item] = pState.data[dong][item] || {};
      const f = rows[r][0] || "1F";
      pState.data[dong][item][f] = (pState.data[dong][item][f] || 0) + val;
    }
  }
}

/* 2. 동 명칭 통일 관리 */
function initDongMapping() {
  PROJECTS.forEach(p => {
    state.projects[p.key].dongs.forEach(d => {
      const stdName = d.replace(/1BL_|2BL_|_PIT|동/g, ""); // 자동 정제 예시
      state.dongMap[`${p.key}::${d}`] = stdName;
    });
  });
}

function renderDongMapping() {
  const allRawDongs = Object.keys(state.dongMap).sort();
  dom.dongList.innerHTML = allRawDongs.map(key => {
    const [pKey, dName] = key.split("::");
    return `
      <div class="dong-row">
        <div class="col-p">[${pKey.toUpperCase()}] ${dName}</div>
        <div class="col-arrow">→</div>
        <div class="col-std"><input class="dong-std-input" data-key="${key}" value="${state.dongMap[key]}" /></div>
      </div>
    `;
  }).join("");

  document.querySelectorAll(".dong-std-input").forEach(el => {
    el.oninput = (e) => state.dongMap[e.target.dataset.key] = e.target.value.trim();
    applyArrowNav(el);
  });
}

/* 3. 아이템 통일 관리 (리스트형) */
function buildItemGroups() {
  const grouped = new Map();
  PROJECTS.forEach(p => {
    state.projects[p.key].rawItems.forEach(raw => {
      const sig = raw.replace(/\s+/g,"").toUpperCase();
      if (!grouped.has(sig)) {
        grouped.set(sig, { id: Math.random().toString(36).substr(2,9), canonical: raw, category: "잡/기타", items: { current:[], a:[], b:[], c:[] } });
      }
      const g = grouped.get(sig);
      if (!g.items[p.key].includes(raw)) g.items[p.key].push(raw);
    });
  });
  state.mappingGroups = [...grouped.values()];
}

function renderItemMapping() {
  dom.itemList.innerHTML = state.mappingGroups.map((g, idx) => `
    <div class="item-row">
      <div class="col-check"><input type="checkbox" onchange="toggleSel('${g.id}', this.checked)"></div>
      <div class="col-orig">
        ${PROJECTS.map(p => `<span class="p-chip ${p.key}">${g.items[p.key][0] || '-'}</span>`).join("")}
      </div>
      <div class="col-edit"><input class="item-std-input" data-id="${g.id}" value="${g.canonical}" /></div>
      <div class="col-cat">
        <select class="item-cat-select" data-id="${g.id}">
          ${CATEGORIES.map(c => `<option value="${c}" ${g.category === c ? 'selected' : ''}>${c}</option>`).join("")}
        </select>
      </div>
    </div>
  `).join("");

  document.querySelectorAll(".item-std-input").forEach(el => {
    el.oninput = (e) => state.mappingGroups.find(x => x.id === e.target.dataset.id).canonical = e.target.value;
    applyArrowNav(el);
  });
  document.querySelectorAll(".item-cat-select").forEach(el => {
    el.onchange = (e) => state.mappingGroups.find(x => x.id === e.target.dataset.id).category = e.target.value;
    applyArrowNav(el);
  });
}

/* 방향키 네비게이션 지원 */
function applyArrowNav(el) {
  el.onkeydown = (e) => {
    const inputs = Array.from(document.querySelectorAll('.dong-std-input, .item-std-input, .item-cat-select'));
    const idx = inputs.indexOf(el);
    if (e.key === "ArrowDown" && idx < inputs.length - 1) { e.preventDefault(); inputs[idx + 1].focus(); }
    if (e.key === "ArrowUp" && idx > 0) { e.preventDefault(); inputs[idx - 1].focus(); }
  };
}

/* 4. 비교표 적용 */
$("btn-apply-all").onclick = () => {
  const stdDongs = [...new Set(Object.values(state.dongMap))].sort();
  dom.filterDong.innerHTML = stdDongs.map(d => `<option value="${d}">${d}</option>`).join("");
  renderCompare();
  $("tab-compare").click();
};

function renderCompare() {
  const targetDong = dom.filterDong.value;
  const unified = { items: {} };
  
  PROJECTS.forEach(p => {
    state.projects[p.key].dongs.forEach(d => {
      if (state.dongMap[`${p.key}::${d}`] !== targetDong) return;
      const dData = state.projects[p.key].data[d];
      for (const raw in dData) {
        const group = state.mappingGroups.find(g => g.items[p.key].includes(raw));
        if (!group) continue;
        const name = group.canonical;
        unified.items[name] = unified.items[name] || { current:0, a:0, b:0, c:0 };
        for (const f in dData[raw]) unified.items[name][p.key] += dData[raw][f];
      }
    });
  });

  dom.compareList.innerHTML = Object.keys(unified.items).map(name => `
    <div class="compare-card">
      <strong>${name}</strong>
      <table>
        <tr><td>현재</td><td>${unified.items[name].current.toLocaleString()}</td></tr>
        <tr><td>평균(ABC)</td><td>${((unified.items[name].a + unified.items[name].b + unified.items[name].c)/3).toFixed(1)}</td></tr>
      </table>
    </div>
  `).join("");
}

/* 공통 UI 제어 */
dom.tabs.forEach(t => t.onclick = () => {
  dom.tabs.forEach(x => x.classList.remove('is-active')); t.classList.add('is-active');
  document.querySelectorAll('.tab-panel').forEach(x => x.classList.remove('is-active')); $(`tab-${t.dataset.tab}`).classList.add('is-active');
});
dom.subTabs.forEach(t => t.onclick = () => {
  dom.subTabs.forEach(x => x.classList.remove('is-active')); t.classList.add('is-active');
  document.querySelectorAll('.sub-panel').forEach(x => x.classList.remove('is-active')); $(t.dataset.sub).classList.add('is-active');
});
