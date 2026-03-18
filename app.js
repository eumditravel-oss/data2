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
  btnExcel: $("btn-excel"),
  dongList: $("dong-mapping-list"),
  itemList: $("mapping-group-list"),
  filterDong: $("filter-dong"),
  filterCategory: $("filter-category"),
  compareList: $("compare-card-list"),
  ratioBoard: $("floor-ratio-board"),
  uploadStatus: $("upload-status")
};

/* 층 정렬 로직 */
function floorSorter(a, b) {
  const getRank = (name) => {
    const s = String(name).toUpperCase().trim();
    if (s.startsWith('B')) return 1000 - (parseInt(s.replace('B', '')) || 0);
    if (s === 'FT') return 2000;
    if (s.endsWith('F') || /^\d+$/.test(s)) return 3000 + (parseInt(s.replace('F', '')) || 0);
    if (s.startsWith('PH')) return 4000 + (parseInt(s.replace('PH', '')) || 0);
    return 5000;
  };
  return getRank(a) - getRank(b);
}

/* 중분류 예측 */
function predictCategory(name) {
  const s = String(name).toUpperCase().replace(/\s+/g, "");
  if (/(H|D|HD|SD|D)\d+/.test(s) || s.includes("철근") || s.startsWith("H1") || s.startsWith("D1")) return "철근";
  if (s.includes("MPA") || /\d+-\d+-\d+/.test(s) || (/^\d+$/.test(s) && parseInt(s) >= 150)) return "콘크리트";
  if (["폼","FORM","회","알폼","갱폼","합벽"].some(k => s.includes(k)) || /[가-힣]/.test(s)) return "거푸집";
  return "잡/기타";
}

/* 파싱 및 분석 */
dom.btnParse.onclick = async () => {
  dom.uploadStatus.textContent = "엑셀 데이터를 정밀 분석 중입니다...";
  try {
    for (const p of PROJECTS) {
      const files = Array.from($(`file-${p.key}`).files);
      if(files.length === 0) continue;
      const pState = state.projects[p.key];
      for (const file of files) {
        const rows = XLSX.utils.sheet_to_json(XLSX.read(await file.arrayBuffer(),{type:'array'}).Sheets[XLSX.read(await file.arrayBuffer(),{type:'array'}).SheetNames[0]], {header:1, defval:""});
        parseSheetData(rows, pState);
      }
    }
    initDongMapping(); buildItemGroups(); renderDongUI(); renderItemUI();
    dom.uploadStatus.textContent = "분석 완료!";
    dom.tabs[1].click();
  } catch(e) { dom.uploadStatus.textContent = "오류: " + e.message; }
};

function parseSheetData(rows, pState) {
  let currentDong = "";
  const r3 = rows[2] || [], r4 = rows[3] || [];
  for (let r = 4; r < rows.length; r++) {
    const m = rows[r].join("|").match(/동\s*명\s*:\s*\[([^\]]+)\]/);
    if (m) {
      const raw = m[1].trim();
      if (raw && !PROJECTS.some(p => raw.toLowerCase().includes(p.key))) {
        currentDong = raw;
        if (!pState.dongs.includes(currentDong)) pState.dongs.push(currentDong);
        pState.data[currentDong] = {};
      }
      continue;
    }
    if (!currentDong) continue;
    const fRaw = String(rows[r][0]).trim();
    if (fRaw === "" || fRaw.includes("계") || fRaw.includes("공사명")) continue;
    let floor = /^\d+$/.test(fRaw) ? fRaw + "F" : fRaw;
    if (!pState.floors.includes(floor)) pState.floors.push(floor);
    for (let c = 1; c < rows[r].length; c++) {
      const item = String(r3[c] || r4[c] || "").trim();
      const val = parseFloat(String(rows[r][c]).replace(/,/g,""));
      if (!item || isNaN(val) || val === 0) continue;
      if (!pState.rawItems.includes(item)) pState.rawItems.push(item);
      pState.data[currentDong][item] = pState.data[currentDong][item] || {};
      pState.data[currentDong][item][floor] = (pState.data[currentDong][item][floor] || 0) + val;
    }
  }
}

function initDongMapping() {
  PROJECTS.forEach(p => state.projects[p.key].dongs.forEach(d => {
    state.dongMap[`${p.key}::${d}`] = d.replace(/1BL_|2BL_|_PIT|동/g, "").trim();
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

function buildItemGroups() {
  const grouped = new Map();
  PROJECTS.forEach(p => state.projects[p.key].rawItems.forEach(raw => {
    const sig = raw.replace(/\s+/g, "").toUpperCase();
    if (!grouped.has(sig)) grouped.set(sig, { id:Math.random().toString(36).substr(2,9), canonical:raw, category:predictCategory(raw), items:{current:[],a:[],b:[],c:[]} });
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
      <div class="col-cat">
        <select class="item-cat-select" data-id="${g.id}">${CATEGORIES.map(c => `<option value="${c}" ${g.category === c ? 'selected' : ''}>${c}</option>`).join("")}</select>
      </div>
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

/* 데이터 가공 로직 (공통 사용) */
function getUnifiedData() {
  const targetStd = dom.filterDong.value;
  const catFilter = dom.filterCategory.value;
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
        if (!unified.items[g.canonical][f]) unified.items[g.canonical][f] = {current:0, a:0, b:0, c:0};
        unified.items[g.canonical][f][p.key] += dData[raw][f];
      }
    }
  }));
  unified.floors.sort(floorSorter);
  return unified;
}

/* 화면 렌더링 */
function renderCompare() {
  if (!state.mappedReady) return;
  const unified = getUnifiedData();
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
              <tr class="row-current"><td>현재</td>${unified.floors.map(f=>`<td>${(v[f]?.current||0).toLocaleString()}</td>`).join("")}</tr>
              <tr><td>유사 A</td>${unified.floors.map(f=>`<td>${(v[f]?.a||0).toLocaleString()}</td>`).join("")}</tr>
              <tr><td>유사 B</td>${unified.floors.map(f=>`<td>${(v[f]?.b||0).toLocaleString()}</td>`).join("")}</tr>
              <tr><td>유사 C</td>${unified.floors.map(f=>`<td>${(v[f]?.c||0).toLocaleString()}</td>`).join("")}</tr>
              <tr style="background:#f4f7fd; font-weight:bold;"><td>평균</td>${unified.floors.map(f=>{
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
  const data = { current:{}, a:{}, b:{}, c:{} };
  floors.forEach(f => {
    ['current','a','b','c'].forEach(p => {
      const conc = Object.keys(unified.items).filter(n=>unified.items[n].cat==='콘크리트').reduce((s,n)=>s+(unified.items[n][f]?.[p]||0),0);
      const rebar = Object.keys(unified.items).filter(n=>unified.items[n].cat==='철근').reduce((s,n)=>s+(unified.items[n][f]?.[p]||0),0);
      data[p][f] = conc > 0 ? (rebar / conc).toFixed(4) : "0.0000";
    });
  });
  dom.ratioBoard.innerHTML = `
    <table class="compare-matrix">
      <thead><tr><th>구분 (Ton/m³)</th>${floors.map(f=>`<th>${f}</th>`).join("")}</tr></thead>
      <tbody>
        <tr class="row-current"><td>현재</td>${floors.map(f=>`<td>${data.current[f]}</td>`).join("")}</tr>
        <tr><td>유사 A</td>${floors.map(f=>`<td>${data.a[f]}</td>`).join("")}</tr>
        <tr><td>유사 B</td>${floors.map(f=>`<td>${data.b[f]}</td>`).join("")}</tr>
        <tr><td>유사 C</td>${floors.map(f=>`<td>${data.c[f]}</td>`).join("")}</tr>
      </tbody>
    </table>`;
}

/* Excel 다운로드 핵심 로직 */
dom.btnExcel.onclick = () => {
  if (!state.mappedReady) return alert("먼저 분석을 완료해주세요.");
  const dong = dom.filterDong.value;
  const unified = getUnifiedData();
  const floors = unified.floors;
  const aoa = [];

  // 1. 비율 분석표 추가
  aoa.push([`동: ${dong} - 층별 철근 지표 분석 (Ton / m³)`]);
  aoa.push(["구분", ...floors]);
  ['current', 'a', 'b', 'c'].forEach(p => {
    const row = [p === 'current' ? '현재 프로젝트' : `유사 ${p.toUpperCase()}`];
    floors.forEach(f => {
      const conc = Object.keys(unified.items).filter(n=>unified.items[n].cat==='콘크리트').reduce((s,n)=>s+(unified.items[n][f]?.[p]||0),0);
      const rebar = Object.keys(unified.items).filter(n=>unified.items[n].cat==='철근').reduce((s,n)=>s+(unified.items[n][f]?.[p]||0),0);
      row.push(conc > 0 ? parseFloat((rebar / conc).toFixed(4)) : 0);
    });
    aoa.push(row);
  });
  aoa.push([]); // 빈 줄

  // 2. 상세 아이템 비교표 추가
  aoa.push(["아이템별 상세 비교표"]);
  Object.keys(unified.items).forEach(name => {
    const v = unified.items[name];
    aoa.push([`항목: ${name}`, `분류: ${unified.items[name].cat}`]);
    aoa.push(["구분", ...floors]);
    
    PROJECTS.forEach(p => {
      const row = [p.name];
      floors.forEach(f => row.push(v[f]?.[p.key] || 0));
      aoa.push(row);
    });
    
    // 평균 행 추가
    const avgRow = ["유사 평균"];
    floors.forEach(f => avgRow.push(((v[f]?.a||0) + (v[f]?.b||0) + (v[f]?.c||0)) / 3));
    aoa.push(avgRow);
    aoa.push([]); // 아이템 간 간격
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  XLSX.utils.book_append_sheet(wb, ws, `${dong}_비교분석`);
  XLSX.writeFile(wb, `QS_비교표_${dong}.xlsx`);
};

/* 기타 버튼 이벤트 */
$("btn-apply-all").onclick = () => {
  state.mappedReady = true;
  const stdDongs = [...new Set(Object.values(state.dongMap))].filter(Boolean).sort();
  dom.filterDong.innerHTML = stdDongs.map(d => `<option value="${d}">${d}</option>`).join("");
  renderCompare();
  dom.tabs[2].click();
};

[$("filter-dong"), $("filter-category")].forEach(el => el.onchange = renderCompare);
$("dong-search").oninput = renderDongUI;
$("mapping-search").oninput = renderItemUI;

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
