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
  filterCategory: $("filter-category"),
  compareList: $("compare-card-list"),
  ratioBoard: $("floor-ratio-board"),
  uploadStatus: $("upload-status")
};

/* 중분류 예측 (사용자 요구사항 반영) */
function predictCategory(name) {
  const s = String(name).toUpperCase().replace(/\s+/g, "");
  if (/(H|D|HD|SD|D)\d+/.test(s) || s.includes("철근") || s.startsWith("H1") || s.startsWith("D1")) return "철근";
  if (s.includes("MPA") || /\d+-\d+-\d+/.test(s) || (/^\d+$/.test(s) && parseInt(s) >= 150)) return "콘크리트";
  if (["폼","FORM","회","회","알폼","갱폼","합벽"].some(k => s.includes(k)) || /[가-힣]/.test(s)) return "거푸집";
  return "잡/기타";
}

/* 데이터 파싱 시작 */
dom.btnParse.onclick = async () => {
  dom.uploadStatus.textContent = "데이터를 정밀 분석 중입니다. 잠시만 기다려주세요...";
  try {
    for (const p of PROJECTS) {
      const files = Array.from($(`file-${p.key}`).files);
      const pState = state.projects[p.key];
      for (const file of files) {
        const rows = XLSX.utils.sheet_to_json(XLSX.read(await file.arrayBuffer(),{type:'array'}).Sheets[XLSX.read(await file.arrayBuffer(),{type:'array'}).SheetNames[0]], {header:1, defval:""});
        parseSheetData(rows, pState);
      }
    }
    initDongMapping(); buildItemGroups(); renderDongUI(); renderItemUI();
    dom.uploadStatus.textContent = "분석 완료! '설정' 탭에서 데이터를 검토하세요.";
    dom.tabs[1].click();
  } catch(e) { dom.uploadStatus.textContent = "분석 중 오류: " + e.message; }
};

/** 엑셀 데이터 추출 핵심 로직 **/
function parseSheetData(rows, pState) {
  let currentDong = "";
  // 엑셀 3행(index 2)과 4행(index 3)에서 자재명/규격 결합 추출
  const headerRow1 = rows[2] || [];
  const headerRow2 = rows[3] || [];
  const itemNames = headerRow1.map((val, idx) => {
    const combined = (String(val) + String(headerRow2[idx])).trim();
    return combined;
  });

  for (let r = 4; r < rows.length; r++) {
    const row = rows[r];
    const rowText = row.join("|");
    
    // 1. [동 명칭] 추출 정규식
    const dongMatch = rowText.match(/\[([^\]]+)\]/);
    if (dongMatch) {
      currentDong = dongMatch[1].trim();
      if (!pState.dongs.includes(currentDong)) pState.dongs.push(currentDong);
      pState.data[currentDong] = pState.data[currentDong] || {};
      continue;
    }
    if (!currentDong) continue;

    // 2. 층 이름 추출 (1열) - 비어있으면 데이터 행이 아니거나 이전 행의 연속임
    const fRaw = String(row[0]).trim();
    if (fRaw === "" || fRaw.includes("층") || fRaw.includes("계")) continue;
    
    let floor = /^\d+$/.test(fRaw) ? fRaw + "F" : fRaw;
    if (!pState.floors.includes(floor)) pState.floors.push(floor);

    // 3. 수량 데이터 수집
    for (let c = 1; c < row.length; c++) {
      const name = itemNames[c];
      const val = parseFloat(String(row[c]).replace(/,/g, ""));
      
      if (!name || isNaN(val) || val === 0) continue;
      
      if (!pState.rawItems.includes(name)) pState.rawItems.push(name);
      pState.data[currentDong][name] = pState.data[currentDong][name] || {};
      pState.data[currentDong][name][floor] = (pState.data[currentDong][name][floor] || 0) + val;
    }
  }
}

/* 동 명칭 관리 */
function initDongMapping() {
  PROJECTS.forEach(p => state.projects[p.key].dongs.forEach(d => {
    const key = `${p.key}::${d}`;
    const clean = d.replace(/1BL_|2BL_|_PIT|동/g, "").trim();
    state.dongMap[key] = clean;
  }));
}

function renderDongUI() {
  const q = $("dong-search").value.toLowerCase();
  dom.dongList.innerHTML = Object.keys(state.dongMap).filter(k=>k.toLowerCase().includes(q)).sort().map(key => `
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

/* 아이템 그룹화 */
function buildItemGroups() {
  const grouped = new Map();
  PROJECTS.forEach(p => state.projects[p.key].rawItems.forEach(raw => {
    const sig = raw.replace(/\s+/g,"").toUpperCase();
    if (!grouped.has(sig)) {
      grouped.set(sig, { id:Math.random().toString(36).substr(2,9), canonical:raw, category:predictCategory(raw), items:{current:[],a:[],b:[],c:[]} });
    }
    if (!grouped.get(sig).items[p.key].includes(raw)) grouped.get(sig).items[p.key].push(raw);
  }));
  state.mappingGroups = [...grouped.values()].sort((a,b)=>a.canonical.localeCompare(b.canonical));
}

function renderItemUI() {
  const q = $("mapping-search").value.toLowerCase();
  dom.itemList.innerHTML = state.mappingGroups.filter(g=>g.canonical.toLowerCase().includes(q)).map(g => `
    <div class="item-row">
      <div class="col-check"><input type="checkbox"></div>
      <div class="col-orig">${PROJECTS.map(p=>`<span class="p-chip ${p.key}">${g.items[p.key][0]||'-'}</span>`).join("")}</div>
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

/* 3. 비교표 렌더링 (A, B, C 상세 데이터 포함) */
$("btn-apply-all").onclick = () => {
  state.mappedReady = true;
  const stdDongs = [...new Set(Object.values(state.dongMap))].sort();
  dom.filterDong.innerHTML = stdDongs.map(d => `<option value="${d}">${d}</option>`).join("");
  renderCompare();
  dom.tabs[2].click();
};

[$("filter-dong"), $("filter-category")].forEach(el => el.onchange = renderCompare);
$("dong-search").oninput = renderDongUI;
$("mapping-search").oninput = renderItemUI;

function renderCompare() {
  if (!state.mappedReady) return;
  const targetStd = dom.filterDong.value;
  const catFilter = $("filter-category").value;
  const unified = { floors: [], items: {} };

  PROJECTS.forEach(p => state.projects[p.key].dongs.forEach(orig => {
    if (state.dongMap[`${p.key}::${orig}`] !== targetStd) return;
    const dData = state.projects[p.key].data[orig];
    for (const raw in dData) {
      const g = state.mappingGroups.find(x => x.items[p.key].includes(raw));
      if (!g || (catFilter !== 'all' && g.category !== catFilter)) continue;
      
      unified.items[g.canonical] = unified.items[g.canonical] || { cat: g.category };
      for (const f in dData[raw]) {
        if (!unified.floors.includes(f)) unified.floors.push(f);
        if (!unified.items[g.canonical][f]) unified.items[g.canonical][f] = {current:0,a:0,b:0,c:0};
        unified.items[g.canonical][f][p.key] += dData[raw][f];
      }
    }
  }));

  unified.floors.sort((x,y) => (parseInt(x)||0) - (parseInt(y)||0));

  // 분석 보드 및 카드 렌더링
  renderRatioBoard(unified);
  dom.compareList.innerHTML = Object.keys(unified.items).map(name => {
    const v = unified.items[name];
    return `
      <div class="compare-card">
        <div class="compare-card__head"><strong>${name}</strong> <small>(${unified.items[name].cat})</small></div>
        <div class="compare-card__body">
          <table class="compare-matrix">
            <thead><tr><th>구분</th>${unified.floors.map(f=>`<th>${f}</th>`).join("")}</tr></thead>
            <tbody>
              <tr class="row-current"><td>현재 수량</td>${unified.floors.map(f=>`<td>${(v[f]?.current||0).toLocaleString()}</td>`).join("")}</tr>
              <tr><td>유사 A</td>${unified.floors.map(f=>`<td>${(v[f]?.a||0).toLocaleString()}</td>`).join("")}</tr>
              <tr><td>유사 B</td>${unified.floors.map(f=>`<td>${(v[f]?.b||0).toLocaleString()}</td>`).join("")}</tr>
              <tr><td>유사 C</td>${unified.floors.map(f=>`<td>${(v[f]?.c||0).toLocaleString()}</td>`).join("")}</tr>
              <tr style="background:#f4f7fd; font-weight:bold;"><td>평균(ABC)</td>${unified.floors.map(f=>{
                const avg = ((v[f]?.a||0) + (v[f]?.b||0) + (v[f]?.c||0)) / 3;
                return `<td>${avg.toLocaleString(undefined,{maximumFractionDigits:1})}</td>`;
              }).join("")}</tr>
            </tbody>
          </table>
        </div>
      </div>`;
  }).join("");
}

function renderRatioBoard(unified) {
  const floors = unified.floors;
  const pKeys = ['current', 'a', 'b', 'c'];
  const data = pKeys.reduce((acc, k) => { acc[k] = {}; return acc; }, {});

  floors.forEach(f => {
    const getSum = (pKey, cat) => Object.keys(unified.items)
      .filter(name => unified.items[name].cat === cat)
      .reduce((sum, name) => sum + (unified.items[name][f]?.[pKey] || 0), 0);

    pKeys.forEach(p => {
      const conc = getSum(p, '콘크리트');
      const rebar = getSum(p, '철근');
      data[p][f] = conc > 0 ? (rebar / conc).toFixed(4) : "0.0000";
    });
  });

  dom.ratioBoard.innerHTML = `
    <table class="compare-matrix">
      <thead><tr><th>구분 (Ton/m³)</th>${floors.map(f=>`<th>${f}</th>`).join("")}</tr></thead>
      <tbody>
        <tr class="row-current"><td>현재 프로젝트</td>${floors.map(f=>`<td>${data.current[f]}</td>`).join("")}</tr>
        <tr><td>유사 A</td>${floors.map(f=>`<td>${data.a[f]}</td>`).join("")}</tr>
        <tr><td>유사 B</td>${floors.map(f=>`<td>${data.b[f]}</td>`).join("")}</tr>
        <tr><td>유사 C</td>${floors.map(f=>`<td>${data.c[f]}</td>`).join("")}</tr>
        <tr style="background:#f4f7fd; font-weight:bold;"><td>평균(ABC)</td>${floors.map(f=>{
          const avg = (parseFloat(data.a[f]) + parseFloat(data.b[f]) + parseFloat(data.c[f])) / 3;
          return `<td>${avg.toFixed(4)}</td>`;
        }).join("")}</tr>
      </tbody>
    </table>`;
}

/* 메뉴 제어 */
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
