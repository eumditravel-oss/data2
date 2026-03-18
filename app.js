"use strict";

const PROJECTS = [
  { key: "current", name: "현재" }, { key: "a", name: "유사A" }, 
  { key: "b", name: "유사B" }, { key: "c", name: "유사C" }
];
const CATEGORIES = ["콘크리트", "거푸집", "철근", "잡/기타"];

const state = {
  projects: { current: emptyState(), a: emptyState(), b: emptyState(), c: emptyState() },
  dongMap: {}, 
  mappingGroups: [],
  selectedGroupIds: new Set()
};

function emptyState() { return { rawItems: [], dongs: [], data: {} }; }

const $ = (id) => document.getElementById(id);
const dom = {
  tabs: document.querySelectorAll(".tab"),
  subTabs: document.querySelectorAll(".sub-tab"),
  btnParse: $("btn-parse"),
  dongList: $("dong-mapping-list"),
  itemList: $("mapping-group-list"),
  filterDong: $("filter-dong"),
  compareList: $("compare-card-list")
};

/* 데이터 분석 */
dom.btnParse.onclick = async () => {
  for (const p of PROJECTS) {
    const files = Array.from($(`file-${p.key}`).files);
    for (const file of files) {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, {type:'array'});
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1, defval:""});
      parseData(rows, state.projects[p.key]);
    }
  }
  initDongMap();
  buildItemGroups();
  renderDongUI();
  renderItemUI();
  dom.tabs[1].click();
};

function parseData(rows, pState) {
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
    const itemRow = (r % 2 === 1) ? r3 : r4;
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

/* 동 명칭 관리 */
function initDongMap() {
  PROJECTS.forEach(p => {
    state.projects[p.key].dongs.forEach(d => {
      const key = `${p.key}::${d}`;
      const numOnly = d.match(/\d+/);
      state.dongMap[key] = numOnly ? numOnly[0] : d;
    });
  });
}

function renderDongUI() {
  dom.dongList.innerHTML = Object.keys(state.dongMap).sort().map(key => {
    const [pk, dn] = key.split("::");
    return `
      <div class="dong-row">
        <div class="col-p-name"><strong>[${pk.toUpperCase()}]</strong> ${dn}</div>
        <div class="col-arrow">→</div>
        <div class="col-std"><input class="dong-std-input" data-key="${key}" value="${state.dongMap[key]}" /></div>
      </div>
    `;
  }).join("");
  document.querySelectorAll(".dong-std-input").forEach(el => {
    el.oninput = (e) => state.dongMap[e.target.dataset.key] = e.target.value;
    applyNav(el);
  });
}

/* 아이템 관리 */
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
  state.mappingGroups = [...grouped.values()].sort((a,b)=>a.canonical.localeCompare(b.canonical));
}

function renderItemUI() {
  dom.itemList.innerHTML = state.mappingGroups.map(g => `
    <div class="item-row">
      <div class="col-check"><input type="checkbox" onchange="toggleSel('${g.id}', this.checked)"></div>
      <div class="col-orig">
        ${PROJECTS.map(p => `<span class="p-chip ${p.key}" title="${g.items[p.key][0]||''}">${g.items[p.key][0]||'-'}</span>`).join("")}
      </div>
      <div class="col-edit"><input class="item-std-input" data-id="${g.id}" value="${g.canonical}" /></div>
      <div class="col-cat">
        <select class="item-cat-select" data-id="${g.id}">
          ${CATEGORIES.map(c => `<option value="${c}" ${g.category === c ? 'selected' : ''}>${c}</option>`).join("")}
        </select>
      </div>
    </div>
  `).join("");
  document.querySelectorAll(".item-std-input, .item-cat-select").forEach(el => {
    el.onchange = el.oninput = (e) => {
      const g = state.mappingGroups.find(x => x.id === e.target.dataset.id);
      if (e.target.tagName === 'INPUT') g.canonical = e.target.value;
      else g.category = e.target.value;
    };
    applyNav(el);
  });
}

function applyNav(el) {
  el.onkeydown = (e) => {
    if (e.key === "ArrowDown" || e.key === "ArrowUp") {
      const all = Array.from(document.querySelectorAll('.dong-std-input, .item-std-input, .item-cat-select'));
      const idx = all.indexOf(el) + (e.key === "ArrowDown" ? 1 : -1);
      if (all[idx]) { e.preventDefault(); all[idx].focus(); }
    }
  };
}

/* 탭 전환 */
dom.tabs.forEach(t => t.onclick = () => {
  dom.tabs.forEach(x => x.classList.remove('is-active')); t.classList.add('is-active');
  document.querySelectorAll('.tab-panel').forEach(x => x.classList.remove('is-active'));
  $(`tab-${t.dataset.tab}`).classList.add('is-active');
});
dom.subTabs.forEach(t => t.onclick = () => {
  dom.subTabs.forEach(x => x.classList.remove('is-active')); t.classList.add('is-active');
  document.querySelectorAll('.sub-panel').forEach(x => x.classList.remove('is-active'));
  $(t.dataset.sub).classList.add('is-active');
});
