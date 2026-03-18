"use strict";

const PROJECTS = [{key:"current", name:"현재"}, {key:"a", name:"유사A"}, {key:"b", name:"유사B"}, {key:"c", name:"유사C"}];
const CATEGORIES = ["콘크리트", "거푸집", "철근", "잡/기타"];

const state = {
  projects: { current: emptyState(), a: emptyState(), b: emptyState(), c: emptyState() },
  dongMap: {}, 
  mappingGroups: [],
  mappedReady: false
};

function emptyState() { return { rawItems: [], dongs: [], floors: [], data: {} }; }

const $ = (id) => document.getElementById(id);
const dom = {
  tabs: document.querySelectorAll(".tab"),
  subTabs: document.querySelectorAll(".sub-tab"),
  btnParse: $("btn-parse"),
  dongList: $("dong-mapping-list"),
  itemList: $("mapping-group-list"),
  filterDong: $("filter-dong"),
  compareList: $("compare-card-list"),
  uploadStatus: $("upload-status")
};

/* 중분류 예측 알고리즘 */
function predictCategory(name) {
  const s = String(name).toUpperCase().replace(/\s+/g, "");
  if (/(H|D|HD|SD)\d+/.test(s) || s.includes("철근")) return "철근";
  if (s.includes("MPA") || /\d+-\d+-\d+/.test(s) || (/^\d+$/.test(s) && parseInt(s) >= 150)) return "콘크리트";
  const formKeywords = ["폼", "FORM", "유로", "알폼", "갱폼", "합벽", "회"];
  if (formKeywords.some(k => s.includes(k)) || /[가-힣]/.test(s)) return "거푸집";
  return "잡/기타";
}

/* 데이터 파싱 */
dom.btnParse.onclick = async () => {
  dom.uploadStatus.textContent = "데이터 분석 중...";
  try {
    for (const p of PROJECTS) {
      const files = Array.from($(`file-${p.key}`).files);
      const pState = state.projects[p.key];
      for (const file of files) {
        const rows = XLSX.utils.sheet_to_json(XLSX.read(await file.arrayBuffer(),{type:'array'}).Sheets[XLSX.read(await file.arrayBuffer(),{type:'array'}).SheetNames[0]], {header:1, defval:""});
        parseSheet(rows, pState);
      }
    }
    initDongMapping();
    buildItemGroups();
    renderDongUI();
    renderItemUI();
    dom.tabs[1].click();
  } catch(e) { dom.uploadStatus.textContent = "오류: " + e.message; }
};

function parseSheet(rows, pState) {
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
    const fRaw = String(rows[r][0]).trim();
    let floor = fRaw !== "" ? (/^\d+$/.test(fRaw) ? fRaw + "F" : fRaw) : "1F";
    if (!pState.floors.includes(floor)) pState.floors.push(floor);
    const head = (r % 2 === 1) ? r3 : r4;
    for (let c = 1; c < rows[r].length; c++) {
      const item = String(head[c] || "").trim();
      const val = parseFloat(String(rows[r][c]).replace(/,/g,""));
      if (!item || isNaN(val)) continue;
      if (!pState.rawItems.includes(item)) pState.rawItems.push(item);
      pState.data[dong][item] = pState.data[dong][item] || {};
      pState.data[dong][item][floor] = (pState.data[dong][item][floor] || 0) + val;
    }
  }
}

/* 동 명칭 관리 */
function initDongMapping() {
  PROJECTS.forEach(p => state.projects[p.key].dongs.forEach(d => {
    const key = `${p.key}::${d}`;
    const num = d.match(/\d+/);
    state.dongMap[key] = num ? num[0] : d;
  }));
}

function renderDongUI() {
  const q = $("dong-search").value.toLowerCase();
  dom.dongList.innerHTML = Object.keys(state.dongMap)
    .filter(key => key.toLowerCase().includes(q))
    .sort()
    .map(key => `
      <div class="dong-row">
        <div class="col-p-name"><strong>[${key.split("::")[0].toUpperCase()}]</strong> ${key.split("::")[1]}</div>
        <div style="width:40px; text-align:center; color:#ccc;">→</div>
        <div style="flex:1"><input class="dong-std-input" data-key="${key}" value="${state.dongMap[key]}" /></div>
      </div>`).join("");
  
  document.querySelectorAll(".dong-std-input").forEach(el => {
    el.oninput = (e) => state.dongMap[e.target.dataset.key] = e.target.value.trim();
    applyNav(el);
  });
}

/* 아이템 관리 */
function buildItemGroups() {
  const grouped = new Map();
  PROJECTS.forEach(p => state.projects[p.key].rawItems.forEach(raw => {
    const sig = raw.replace(/\s+/g,"").toUpperCase();
    if (!grouped.has(sig)) grouped.set(sig, { id:Math.random().toString(36).substr(2,9), canonical:raw, category:predictCategory(raw), items:{current:[],a:[],b:[],c:[]} });
    if (!grouped.get(sig).items[p.key].includes(raw)) grouped.get(sig).items[p.key].push(raw);
  }));
  state.mappingGroups = [...grouped.values()].sort((a,b)=>a.canonical.localeCompare(b.canonical));
}

function renderItemUI() {
  const q = $("mapping-search").value.toLowerCase();
  dom.itemList.innerHTML = state.mappingGroups
    .filter(g => g.canonical.toLowerCase().includes(q))
    .map(g => `
    <div class="item-row">
      <div class="col-check"><input type="checkbox"></div>
      <div class="col-orig">${PROJECTS.map(p=>`<span class="p-chip ${p.key}" title="${g.items[p.key][0]||''}">${g.items[p.key][0]||'-'}</span>`).join("")}</div>
      <div class="col-edit"><input class="item-std-input" data-id="${g.id}" value="${g.canonical}" /></div>
      <div class="col-cat"><select class="item-cat-select" data-id="${g.id}">${CATEGORIES.map(c=>`<option value="${c}" ${g.category===c?'selected':''}>${c}</option>`).join("")}</select></div>
    </div>`).join("");
  
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

/* 비교표 생성 */
$("btn-apply-all").onclick = () => {
  state.mappedReady = true;
  const stdDongs = [...new Set(Object.values(state.dongMap))].sort();
  dom.filterDong.innerHTML = stdDongs.map(d => `<option value="${d}">${d}</option>`).join("");
  renderCompare();
  dom.tabs[2].click();
};

$("dong-search").oninput = renderDongUI;
$("mapping-search").oninput = renderItemUI;
[$("filter-dong"), $("filter-category")].forEach(el => el.onchange = renderCompare);
$("filter-item").oninput = renderCompare;

function renderCompare() {
  if (!state.mappedReady) return;
  const targetStd = dom.filterDong.value;
  const cat = $("filter-category").value;
  const keyw = $("filter-item").value.toLowerCase();
  const unified = { floors: [], items: {} };

  PROJECTS.forEach(p => state.projects[p.key].dongs.forEach(orig => {
    if (state.dongMap[`${p.key}::${orig}`] !== targetStd) return;
    const dData = state.projects[p.key].data[orig];
    for (const raw in dData) {
      const g = state.mappingGroups.find(x => x.items[p.key].includes(raw));
      if (!g || (cat !== 'all' && g.category !== cat) || (keyw && !g.canonical.toLowerCase().includes(keyw))) continue;
      unified.items[g.canonical] = unified.items[g.canonical] || {};
      for (const f in dData[raw]) {
        if (!unified.floors.includes(f)) unified.floors.push(f);
        if (!unified.items[g.canonical][f]) unified.items[g.canonical][f] = {current:0,a:0,b:0,c:0};
        unified.items[g.canonical][f][p.key] += dData[raw][f];
      }
    }
  }));

  unified.floors.sort();
  dom.compareList.innerHTML = Object.keys(unified.items).map(name => {
    const v = unified.items[name];
    return `
      <div class="compare-card">
        <div class="compare-card__head"><strong>${name}</strong></div>
        <div class="compare-card__body">
          <table class="compare-matrix">
            <thead><tr><th>구분</th>${unified.floors.map(f=>`<th>${f}</th>`).join("")}</tr></thead>
            <tbody>
              <tr class="row-current"><td>현재</td>${unified.floors.map(f=>`<td>${(v[f]?.current||0).toLocaleString()}</td>`).join("")}</tr>
              <tr><td>유사A</td>${unified.floors.map(f=>`<td>${(v[f]?.a||0).toLocaleString()}</td>`).join("")}</tr>
              <tr><td>유사B</td>${unified.floors.map(f=>`<td>${(v[f]?.b||0).toLocaleString()}</td>`).join("")}</tr>
              <tr><td>유사C</td>${unified.floors.map(f=>`<td>${(v[f]?.c||0).toLocaleString()}</td>`).join("")}</tr>
              <tr style="background:#f4f7fd; font-weight:bold;"><td>평균(ABC)</td>${unified.floors.map(f=>{
                const avg = ((v[f]?.a||0) + (v[f]?.b||0) + (v[f]?.c||0)) / 3;
                return `<td>${avg.toLocaleString(undefined,{maximumFractionDigits:1})}</td>`;
              }).join("")}</tr>
            </tbody>
          </table>
        </div>
      </div>`;
  }).join("") || "<div class='panel' style='text-align:center; color:#999;'>조건에 맞는 데이터가 없습니다.</div>";
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
