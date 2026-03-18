"use strict";

const CATEGORIES = ["콘크리트", "거푸집", "철근", "잡/기타"];

const state = {
  rawItems: [],
  dongs: [],
  floors: [],
  data: {}, // { dong: { item: { floor: val } } }
  mappings: [], // { id, original, canonical, category }
  ready: false
};

const $ = (id) => document.getElementById(id);

/* 지능형 층 정렬 (요청사항 반영) */
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

/* 카테고리 예측 */
function predictCategory(name) {
  const s = String(name).toUpperCase().replace(/\s+/g, "");
  if (/(H|D|HD|SD)\d+/.test(s) || s.includes("철근")) return "철근";
  if (s.includes("MPA") || /\d+-\d+-\d+/.test(s) || (/^\d+$/.test(s) && parseInt(s) >= 150)) return "콘크리트";
  if (["폼","FORM","회","알폼","갱폼","합벽"].some(k => s.includes(k)) || /[가-힣]/.test(s)) return "거푸집";
  return "잡/기타";
}

/* 파싱 로직 */
$("btn-parse").onclick = async () => {
  const files = Array.from($('file-main').files);
  if (files.length === 0) return alert("파일을 선택해주세요.");

  state.rawItems = []; state.dongs = []; state.floors = []; state.data = {};

  for (const file of files) {
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array' });
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: "" });
    parseRows(rows);
  }

  buildMapping();
  renderMapping();
  $('tab-mapping').click();
};

function parseRows(rows) {
  let currentDong = "";
  let lastFloor = "";
  const row3 = rows[2] || []; 
  const row4 = rows[3] || [];

  for (let r = 4; r < rows.length; r++) {
    const row = rows[r];
    const rowText = row.join("|");
    const dongMatch = rowText.match(/동\s*명\s*:\s*\[([^\]]+)\]/);
    
    if (dongMatch) {
      currentDong = dongMatch[1].trim();
      if (!state.dongs.includes(currentDong)) state.dongs.push(currentDong);
      state.data[currentDong] = state.data[currentDong] || {};
      lastFloor = ""; continue;
    }
    if (!currentDong) continue;

    const fRaw = String(row[0]).trim();
    if (fRaw !== "" && !fRaw.includes("계") && !fRaw.includes("층")) {
      lastFloor = /^\d+$/.test(fRaw) ? fRaw + "F" : fRaw;
      if (!state.floors.includes(lastFloor)) state.floors.push(lastFloor);
    }
    if (!lastFloor) continue;

    for (let c = 1; c < row.length; c++) {
      const val = parseFloat(String(row[c]).replace(/,/g, ""));
      if (isNaN(val) || val === 0) continue;

      let itemName = (fRaw !== "") ? String(row3[c] || "").trim() : String(row4[c] || "").trim();
      if (!itemName) itemName = String(row3[c] || row4[c] || "").trim();
      if (!itemName) continue;

      if (!state.rawItems.includes(itemName)) state.rawItems.push(itemName);
      state.data[currentDong][itemName] = state.data[currentDong][itemName] || {};
      state.data[currentDong][itemName][lastFloor] = (state.data[currentDong][itemName][lastFloor] || 0) + val;
    }
  }
}

function buildMapping() {
  state.mappings = state.rawItems.map((item, idx) => ({
    id: idx,
    original: item,
    canonical: item,
    category: predictCategory(item)
  }));
}

function renderMapping() {
  const list = $("mapping-list");
  list.innerHTML = state.mappings.map(m => `
    <div class="item-row">
      <div style="width:60px; text-align:center; color:#999;">${m.id + 1}</div>
      <div style="flex:1; font-weight:bold;">${m.original}</div>
      <div style="width:200px;"><input class="input" value="${m.canonical}" oninput="updateMapping(${m.id}, 'canonical', this.value)" /></div>
      <div style="width:150px;">
        <select class="input" onchange="updateMapping(${m.id}, 'category', this.value)">
          ${CATEGORIES.map(c => `<option value="${c}" ${m.category === c ? 'selected' : ''}>${c}</option>`).join("")}
        </select>
      </div>
    </div>
  `).join("");
}

window.updateMapping = (id, field, val) => { state.mappings[id][field] = val; };

/* 결과 생성 및 렌더링 */
$("btn-apply").onclick = () => {
  state.ready = true;
  const filterDong = $("filter-dong");
  filterDong.innerHTML = state.dongs.sort().map(d => `<option value="${d}">${d}</option>`).join("");
  renderView();
  $('tab-view').click();
};

$("filter-dong").onchange = renderView;

function renderView() {
  if (!state.ready) return;
  const dong = $("filter-dong").value;
  const floors = state.floors.sort(floorSorter);
  const dongData = state.data[dong] || {};

  // 1. 헤더 생성
  const headRow = $("table-header-row");
  headRow.innerHTML = `<th>아이템 명칭</th><th>분류</th><th>단위</th>` + floors.map(f => `<th>${f}</th>`).join("") + `<th>합계</th>`;

  // 2. 데이터 그룹화
  const grouped = {};
  state.mappings.forEach(m => {
    const qtyByFloor = dongData[m.original] || {};
    if (Object.keys(qtyByFloor).length === 0) return;
    
    if (!grouped[m.canonical]) grouped[m.canonical] = { category: m.category, floors: {} };
    floors.forEach(f => {
      grouped[m.canonical].floors[f] = (grouped[m.canonical].floors[f] || 0) + (qtyByFloor[f] || 0);
    });
  });

  // 3. 지표 계산 (철근 톤당 루베)
  const stats = { conc: {}, rebar: {}, ratio: {} };
  floors.forEach(f => {
    stats.conc[f] = Object.keys(grouped).filter(n => grouped[n].category === '콘크리트').reduce((s, n) => s + grouped[n].floors[f], 0);
    stats.rebar[f] = Object.keys(grouped).filter(n => grouped[n].category === '철근').reduce((s, n) => s + grouped[n].floors[f], 0);
    stats.ratio[f] = stats.conc[f] > 0 ? (stats.rebar[f] / stats.conc[f]).toFixed(4) : "0.0000";
  });

  // 4. 테이블 바디 생성
  const body = $("table-body");
  let html = "";

  // 아이템 수량행
  Object.keys(grouped).sort().forEach(name => {
    const item = grouped[name];
    const unit = item.category === '철근' ? 'TON' : (item.category === '콘크리트' ? 'M3' : 'M2');
    const total = floors.reduce((s, f) => s + item.floors[f], 0);
    html += `<tr>
      <td class="text-left">${name}</td><td>${item.category}</td><td>${unit}</td>
      ${floors.map(f => `<td>${item.floors[f].toLocaleString(undefined, {maximumFractionDigits:2})}</td>`).join("")}
      <td class="col-total">${total.toLocaleString(undefined, {maximumFractionDigits:2})}</td>
    </tr>`;
  });

  // 비율(지표) 행 추가
  html += `<tr class="row-ratio">
    <td colspan="3" style="text-align:right; font-weight:800; background:#fff4e6;">층별 철근 비율 (Ton / m³) </td>
    ${floors.map(f => `<td style="background:#fff4e6; font-weight:800; color:#e67e22;">${stats.ratio[f]}</td>`).join("")}
    <td style="background:#ffd8a8;">-</td>
  </tr>`;

  body.innerHTML = html;
  
  // 대시보드 요약
  const totalConc = floors.reduce((s, f) => s + stats.conc[f], 0);
  const totalRebar = floors.reduce((s, f) => s + stats.rebar[f], 0);
  const totalRatio = totalConc > 0 ? (totalRebar / totalConc).toFixed(4) : 0;
  
  $("dong-summary").innerHTML = `
    <div class="stat-item"><span>총 콘크리트</span><strong>${totalConc.toLocaleString()} m³</strong></div>
    <div class="stat-item"><span>총 철근</span><strong>${totalRebar.toLocaleString()} Ton</strong></div>
    <div class="stat-item highlight"><span>평균 톤당 루베</span><strong>${totalRatio} Ton/m³</strong></div>
  `;
}

/* 탭 전환 로직 */
document.querySelectorAll(".tab").forEach(tab => {
  tab.onclick = () => {
    document.querySelectorAll(".tab, .tab-panel").forEach(el => el.classList.remove("is-active"));
    tab.classList.add("is-active");
    $( "tab-" + tab.dataset.tab).classList.add("is-active");
  };
});
