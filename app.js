"use strict";

const CATEGORIES = ["콘크리트", "거푸집", "철근", "잡/기타"];

const state = {
  rawItems: [], dongs: [], floors: [], data: {}, mappings: [], ready: false
};

const $ = (id) => document.getElementById(id);

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

function predictCategory(name) {
  const s = String(name).toUpperCase().replace(/\s+/g, "");
  if (/(H|D|HD|SD)\d+/.test(s) || s.includes("철근")) return "철근";
  if (s.includes("MPA") || /\d+-\d+-\d+/.test(s) || (/^\d+$/.test(s) && parseInt(s) >= 150)) return "콘크리트";
  if (["폼","FORM","회","알폼","갱폼","합벽"].some(k => s.includes(k)) || /[가-힣]/.test(s)) return "거푸집";
  return "잡/기타";
}

$("btn-parse").onclick = async () => {
  const files = Array.from($('file-main').files);
  if (files.length === 0) return alert("파일을 선택해주세요.");
  state.rawItems = []; state.dongs = []; state.floors = []; state.data = {};

  for (const file of files) {
    const rows = XLSX.utils.sheet_to_json(XLSX.read(await file.arrayBuffer(), {type:'array'}).Sheets[XLSX.read(await file.arrayBuffer(), {type:'array'}).SheetNames[0]], {header:1, defval:""});
    parseRows(rows);
  }
  buildMapping(); renderMapping(); switchTab('mapping');
};

function parseRows(rows) {
  let curDong = "", lastF = "";
  const r3 = rows[2] || [], r4 = rows[3] || [];
  for (let r = 4; r < rows.length; r++) {
    const row = rows[r];
    const m = row.join("|").match(/동\s*명\s*:\s*\[([^\]]+)\]/);
    if (m) { curDong = m[1].trim(); if(!state.dongs.includes(curDong)) state.dongs.push(curDong); state.data[curDong] = {}; lastF = ""; continue; }
    if (!curDong) continue;
    const fRaw = String(row[0]).trim();
    if (fRaw !== "" && !fRaw.includes("계") && !fRaw.includes("층")) {
      lastF = /^\d+$/.test(fRaw) ? fRaw + "F" : fRaw;
      if (!state.floors.includes(lastF)) state.floors.push(lastF);
    }
    if (!lastF) continue;
    for (let c = 1; c < row.length; c++) {
      const val = parseFloat(String(row[c]).replace(/,/g, ""));
      if (isNaN(val) || val === 0) continue;
      let name = (fRaw !== "") ? String(r3[c] || "").trim() : String(r4[c] || "").trim();
      if (!name) name = String(r3[c] || r4[c] || "").trim();
      if (!name) continue;
      if (!state.rawItems.includes(name)) state.rawItems.push(name);
      state.data[curDong][name] = state.data[curDong][name] || {};
      state.data[curDong][name][lastF] = (state.data[curDong][name][lastF] || 0) + val;
    }
  }
}

function buildMapping() {
  state.mappings = state.rawItems.map((item, idx) => ({ id: idx, original: item, canonical: item, category: predictCategory(item) }));
}

function renderMapping() {
  $("mapping-list").innerHTML = state.mappings.map(m => `
    <div class="item-row">
      <div class="col-num">${m.id + 1}</div>
      <div class="col-orig">${m.original}</div>
      <div class="col-edit"><input class="input" value="${m.canonical}" oninput="updateMapping(${m.id},'canonical',this.value)"/></div>
      <div class="col-cat">
        <select class="input" onchange="updateMapping(${m.id},'category',this.value)">
          ${CATEGORIES.map(c=>`<option value="${c}" ${m.category===c?'selected':''}>${c}</option>`).join("")}
        </select>
      </div>
    </div>`).join("");
}
window.updateMapping = (id, f, v) => state.mappings[id][f] = v;

$("btn-apply").onclick = () => {
  state.ready = true;
  $("filter-dong").innerHTML = state.dongs.sort().map(d => `<option value="${d}">${d}</option>`).join("");
  renderView(); switchTab('view');
};

$("filter-dong").onchange = renderView;

function renderView() {
  if (!state.ready) return;
  const dong = $("filter-dong").value;
  const floors = state.floors.sort(floorSorter);
  const dongData = state.data[dong] || {};
  const grouped = {};

  state.mappings.forEach(m => {
    const qByF = dongData[m.original] || {}; if (Object.keys(qByF).length === 0) return;
    if (!grouped[m.canonical]) grouped[m.canonical] = { category: m.category, floors: {} };
    floors.forEach(f => grouped[m.canonical].floors[f] = (grouped[m.canonical].floors[f] || 0) + (qByF[f] || 0));
  });

  let headHtml = `<tr><th rowspan="2">동</th><th rowspan="2">아이템</th><th rowspan="2">구분</th><th rowspan="2">단위</th><th colspan="${floors.length}">현재 프로젝트 수량</th><th rowspan="2">합계</th></tr><tr>`;
  floors.forEach(f => headHtml += `<th>${f}</th>`); headHtml += "</tr>";
  $("table-head").innerHTML = headHtml;

  let bodyHtml = "";
  ["콘크리트", "철근", "거푸집"].forEach(cat => {
    Object.keys(grouped).filter(n => grouped[n].category === cat).sort().forEach(name => {
      const item = grouped[name];
      const total = floors.reduce((s,f)=>s+item.floors[f],0);
      bodyHtml += `<tr><td>${dong}</td><td>${cat}</td><td>${name}</td><td>${cat==='철근'?'TON':(cat==='콘크리트'?'M3':'M2')}</td>${floors.map(f=>`<td>${item.floors[f].toLocaleString(undefined,{maximumFractionDigits:2})}</td>`).join("")}<td class="col-total">${total.toLocaleString()}</td></tr>`;
    });
    if(cat === '철근') {
      bodyHtml += `<tr class="row-ratio"><td colspan="4" style="text-align:right">톤당 루베 지표 (Ton/m³)</td>${floors.map(f => {
        const cS = Object.keys(grouped).filter(n=>grouped[n].category==='콘크리트').reduce((s,n)=>s+grouped[n].floors[f],0);
        const rS = Object.keys(grouped).filter(n=>grouped[n].category==='철근').reduce((s,n)=>s+grouped[n].floors[f],0);
        return `<td>${cS>0?(rS/cS).toFixed(4):'-'}</td>`;
      }).join("")}<td>-</td></tr>`;
    }
  });
  $("table-body").innerHTML = bodyHtml;
}

/* 엑셀 내보내기 (템플릿 양식 100% 동일화) */
$("btn-excel").onclick = () => {
  const floors = state.floors.sort(floorSorter);
  const aoa = [
    ["QS 분석용 프로젝트 통합 리포트"],
    ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "할증 후 수량 기준"],
    ["동", "아이템", "구분", "단위", "현재 프로젝트 수량", ...Array(floors.length-1).fill(""), "합계"],
    ["", "", "", "", ...floors, ""]
  ];

  const merges = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 4 + floors.length } },
    { s: { r: 2, c: 4 }, e: { r: 2, c: 4 + floors.length - 1 } }
  ];

  state.dongs.sort().forEach(dong => {
    const dongData = state.data[dong] || {};
    const grouped = {};
    state.mappings.forEach(m => {
      const qByF = dongData[m.original] || {}; if (Object.keys(qByF).length === 0) return;
      if (!grouped[m.canonical]) grouped[m.canonical] = { category: m.category, floors: {} };
      floors.forEach(f => grouped[m.canonical].floors[f] = (grouped[m.canonical].floors[f] || 0) + (qByF[f] || 0));
    });

    const startRow = aoa.length;
    ["콘크리트", "철근", "거푸집"].forEach(cat => {
      Object.keys(grouped).filter(n => grouped[n].category === cat).sort().forEach(name => {
        const item = grouped[name];
        const row = [dong, cat, name, cat==='철근'?'TON':(cat==='콘크리트'?'M3':'M2')];
        floors.forEach(f => row.push(item.floors[f] || 0));
        row.push(floors.reduce((s,f)=>s+item.floors[f], 0));
        aoa.push(row);
      });
      if (cat === '철근') {
        const ratioRow = [dong, "지표", "톤당루베", "Ton/M3"];
        floors.forEach(f => {
          const cSum = Object.keys(grouped).filter(n=>grouped[n].category==='콘크리트').reduce((s,n)=>s+grouped[n].floors[f],0);
          const rSum = Object.keys(grouped).filter(n=>grouped[n].category==='철근').reduce((s,n)=>s+grouped[n].floors[f],0);
          ratioRow.push(cSum > 0 ? parseFloat((rSum / cSum).toFixed(4)) : 0);
        });
        ratioRow.push(""); aoa.push(ratioRow);
      }
    });
    merges.push({ s: { r: startRow, c: 0 }, e: { r: aoa.length - 1, c: 0 } }); // 동 이름 열 병합
  });

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges'] = merges;
  
  // 열 너비 설정 (템플릿과 동일하게)
  const colWidths = [
    { wch: 10 }, { wch: 15 }, { wch: 20 }, { wch: 8 }, 
    ...floors.map(() => ({ wch: 10 })), { wch: 12 }
  ];
  ws['!cols'] = colWidths;

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "통합수량표");
  XLSX.writeFile(wb, `QS_통합리포트_전체.xlsx`);
};

function switchTab(id) {
  document.querySelectorAll(".tab, .tab-panel").forEach(el => el.classList.remove("is-active"));
  document.querySelector(`[data-tab="${id}"]`).classList.add("is-active");
  $("tab-" + id).classList.add("is-active");
}
document.querySelectorAll(".tab").forEach(t => t.onclick = () => switchTab(t.dataset.tab));
