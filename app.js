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

/* 중분류 예측 알고리즘 */
function predictCategory(name) {
  const s = String(name).toUpperCase().replace(/\s+/g, "");
  if (/(H|D|HD|SD)\d+/.test(s) || s.includes("철근") || s.startsWith("H") || s.startsWith("D")) return "철근";
  if (s.includes("MPA") || /\d+-\d+-\d+/.test(s) || (/^\d+$/.test(s) && parseInt(s) >= 150)) return "콘크리트";
  if (["폼","FORM","회","유로","알폼","갱폼","합벽"].some(k => s.includes(k)) || /[가-힣]/.test(s)) return "거푸집";
  return "잡/기타";
}

/* 데이터 분석 및 파싱 */
dom.btnParse.onclick = async () => {
  dom.uploadStatus.textContent = "데이터를 분석 중입니다. 유효하지 않은 명칭을 필터링합니다...";
  try {
    for (const p of PROJECTS) {
      const input = $(`file-${p.key}`);
      const files = Array.from(input.files || []);
      if(files.length === 0) continue;

      const pState = state.projects[p.key];
      for (const file of files) {
        const buffer = await file.arrayBuffer();
        const wb = XLSX.read(buffer, { type: 'array' });
        wb.SheetNames.forEach(sn => {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "" });
          parseSheetData(rows, pState);
        });
      }
    }
    initDongMapping();
    buildItemGroups();
    renderDongUI();
    renderItemUI();
    dom.uploadStatus.textContent = "분석 완료!";
    dom.tabs[1].click();
  } catch(e) { 
    dom.uploadStatus.textContent = "오류 발생: " + e.message; 
    console.error(e);
  }
};

/* 엑셀 구조 정밀 파싱 */
function parseSheetData(rows, pState) {
  let currentDong = "";
  let lastFloor = "";
  const row3 = rows[2] || [];
  const row4 = rows[3] || [];

  for (let r = 4; r < rows.length; r++) {
    const row = rows[r];
    if (!row || row.length === 0) continue;

    const rowText = row.join("|");
    
    // [수정] "동 명 : [101동]" 형태만 추출하고 시스템 변수([CURRENT] 등)는 무시
    const dongMatch = rowText.match(/동\s*명\s*:\s*\[([^\]]+)\]/);
    if (dongMatch) {
      const extractedName = dongMatch[1].trim();
      // 대문자만 있거나 비어있는 값은 제외 (예: CURRENT, TEMP 등 필터링)
      if (extractedName && !/^[A-Z_]+$/.test(extractedName)) {
        currentDong = extractedName;
        if (!pState.dongs.includes(currentDong)) pState.dongs.push(currentDong);
        pState.data[currentDong] = pState.data[currentDong] || {};
        lastFloor = "";
      }
      continue;
    }
    if (!currentDong) continue;

    const fRaw = String(row[0]).trim();
    if (fRaw.includes("계") || fRaw.includes("공사명") || fRaw === "층" || fRaw === "") {
        // 층 이름이 없어도 이전 층의 연속 행(row4 매칭행)으로 판단
    } else {
        lastFloor = /^\d+$/.test(fRaw) ? fRaw + "F" : fRaw;
        if (!pState.floors.includes(lastFloor)) pState.floors.push(lastFloor);
    }
    
    if (!lastFloor) continue;

    for (let c = 1; c < row.length; c++) {
      const val = parseFloat(String(row[c]).replace(/,/g, ""));
      if (isNaN(val) || val === 0) continue;

      let itemName = (fRaw !== "" && fRaw !== "층") ? String(row3[c] || "").trim() : String(row4[c] || "").trim();
      if (!itemName) itemName = String(row3[c] || row4[c] || "").trim();
      if (!itemName) continue;

      if (!pState.rawItems.includes(itemName)) pState.rawItems.push(itemName);
      if (!pState.data[currentDong][itemName]) pState.data[currentDong][itemName] = {};
      
      pState.data[currentDong][itemName][lastFloor] = (pState.data[currentDong][itemName][lastFloor] || 0) + val;
    }
  }
}

/* 동 명칭 관리 (필터링 강화) */
function initDongMapping() {
  state.dongMap = {}; // 초기화
  PROJECTS.forEach(p => {
    state.projects[p.key].dongs.forEach(d => {
      // 프로젝트 내부 키값이나 비정상적인 동 이름은 맵에 넣지 않음
      if (d === "current" || d === "a" || d === "b" || d === "c") return;
      
      const key = `${p.key}::${d}`;
      state.dongMap[key] = d.replace(/1BL_|2BL_|_PIT|동/g, "").trim();
    });
  });
}

function renderDongUI() {
  const q = $("dong-search").value.toLowerCase();
  const filteredKeys = Object.keys(state.dongMap).filter(k => k.toLowerCase().includes(q));
  
  if (filteredKeys.length === 0) {
    dom.dongList.innerHTML = "<div style='padding:20px; color:#999; text-align:center;'>추출된 동 정보가 없습니다.</div>";
    return;
  }

  dom.dongList.innerHTML = filteredKeys.sort().map(key => `
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
    if (!grouped.has(sig)) {
      grouped.set(sig, { id: Math.random().toString(36).substr(2, 9), canonical: raw, category: predictCategory(raw), items: { current: [], a: [], b: [], c: [] } });
    }
    if (!grouped.get(sig).items[p.key].includes(raw)) grouped.get(sig).items[p.key].push(raw);
  }));
  state.mappingGroups = [...grouped.values()].sort((a, b) => a.canonical.localeCompare(b.canonical));
}

function renderItemUI() {
  const q = $("mapping-search").value.toLowerCase();
  dom.itemList.innerHTML = state.mappingGroups.filter(g=>g.canonical.toLowerCase().includes(q)).map(g => `
    <div class="item-row">
      <div class="col-check"><input type="checkbox"></div>
      <div class="col-orig">${PROJECTS.map(p=>`<span class="p-chip ${p.key}">${g.items[p.key][0]||'-'}</span>`).join("")}</div>
      <div class="col-edit"><input class="item-std-input" data-id="${g.id}" value="${g.canonical}" /></div>
      <div class="col-cat">
        <select class="item-cat-select" data-id="${g.id}">
          ${CATEGORIES.map(c => `<option value="${c}" ${g.category === c ? 'selected' : ''}>${c}</option>`).join("")}
        </select>
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

$("btn-apply-all").onclick = () => {
  if (Object.keys(state.dongMap).length === 0) {
    alert("분석된 동 정보가 없습니다. 엑셀 파일을 다시 확인해주세요.");
    return;
  }
  state.mappedReady = true;
  const stdDongs = [...new Set(Object.values(state.dongMap))].filter(Boolean).sort();
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
  if (!targetStd) return;

  const unified = { floors: [], items: {} };

  PROJECTS.forEach(p => state.projects[p.key].dongs.forEach(orig => {
    if (state.dongMap[`${p.key}::${orig}`] !== targetStd) return;
    const dData = state.projects[p.key].data[orig];
    for (const raw in dData) {
      const group = state.mappingGroups.find(x => x.items[p.key].includes(raw));
      if (!group || (catFilter !== 'all' && group.category !== catFilter)) continue;
      
      unified.items[group.canonical] = unified.items[group.canonical] || { cat: group.category };
      for (const f in dData[raw]) {
        if (!unified.floors.includes(f)) unified.floors.push(f);
        if (!unified.items[group.canonical][f]) unified.items[group.canonical][f] = {current:0, a:0, b:0, c:0};
        unified.items[group.canonical][f][p.key] += dData[raw][f];
      }
    }
  }));

  unified.floors.sort((x,y) => (parseInt(x)||0) - (parseInt(y)||0));
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
        <tr class="row-current"><td>현재</td>${floors.map(f=>`<td>${data.current[f]}</td>`).join("")}</tr>
        <tr><td>유사 A</td>${floors.map(f=>`<td>${data.a[f]}</td>`).join("")}</tr>
        <tr><td>유사 B</td>${floors.map(f=>`<td>${data.b[f]}</td>`).join("")}</tr>
        <tr><td>유사 C</td>${floors.map(f=>`<td>${data.c[f]}</td>`).join("")}</tr>
      </tbody>
    </table>`;
}

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
