"use strict";

const CATEGORIES = ["콘크리트", "거푸집", "철근", "잡/기타"];

const state = {
  rawItems: [], dongs: [], floors: [], data: {}, mappings: [], ready: false
};

const $ = (id) => document.getElementById(id);

/* 층 정렬 로직 (지하역순 -> FT -> 지상순 -> 옥탑) */
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
  if (files.length === 0) return alert("파일을 먼저 선택해주세요.");
  
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
    if (!row || row.length === 0) continue;

    const txt = row.join("|");
    const m = txt.match(/동\s*명\s*:\s*\[([^\]]+)\]/);
    if (m) { 
      const raw = m[1].trim();
      if (raw) {
        curDong = raw;
        if(!state.dongs.includes(curDong)) state.dongs.push(curDong);
        state.data[curDong] = state.data[curDong] || {};
      }
      lastF = ""; continue; 
    }
    if (!curDong) continue;

    const fRaw = String(row[0]).trim();
    if (fRaw === "층" || fRaw.includes("계") || fRaw.includes("합") || fRaw.includes("공사명")) {
      lastF = ""; continue;
    }

    if (fRaw !== "") {
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
      <div class="col-cat"><select class="input" onchange="updateMapping(${m.id},'category',this.value)">${CATEGORIES.map(c=>`<option value="${c}" ${m.category===c?'selected':''}>${c}</option>`).join("")}</select></div>
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
  ["콘크리트", "철근", "거푸집", "잡/기타"].forEach(cat => {
    const items = Object.keys(grouped).filter(n => grouped[n].category === cat).sort();
    if (items.length === 0) return;
    items.forEach(name => {
      const item = grouped[name];
      const total = floors.reduce((s,f)=>s+item.floors[f],0);
      bodyHtml += `<tr><td>${dong}</td><td>${cat==='콘크리트'?'레미콘':cat}</td><td>${name}</td><td>${cat==='철근'?'TON':(cat==='콘크리트'?'M3':'M2')}</td>${floors.map(f=>`<td>${item.floors[f].toLocaleString(undefined,{maximumFractionDigits:3})}</td>`).join("")}<td class="col-total">${total.toLocaleString(undefined,{maximumFractionDigits:3})}</td></tr>`;
    });
    if(cat === '철근') {
      bodyHtml += `<tr class="row-ratio"><td colspan="4" style="text-align:right">톤당 루베 지표 (Ton/m³)</td>${floors.map(f => {
        const cSum = Object.keys(grouped).filter(n=>grouped[n].category==='콘크리트').reduce((s,n)=>s+grouped[n].floors[f],0);
        const rSum = Object.keys(grouped).filter(n=>grouped[n].category==='철근').reduce((s,n)=>s+grouped[n].floors[f],0);
        return `<td>${cSum>0?(rSum/cSum).toFixed(4):'-'}</td>`;
      }).join("")}<td>-</td></tr>`;
    }
  });
  $("table-body").innerHTML = bodyHtml;
}

/* ★ 템플릿 100% 동일 엑셀 다운로드 (ExcelJS 적용) ★ */
$("btn-excel").onclick = async () => {
  if (!state.ready) return alert("먼저 분석을 완료해주세요.");
  
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('비교양식', {
    views: [{ state: 'frozen', ySplit: 4, xSplit: 4 }] // 틀 고정 (템플릿과 동일)
  });

  const floors = state.floors.sort(floorSorter);

  // 1. 열 너비 설정
  const cols = [
    { width: 10 }, // A: 동
    { width: 14 }, // B: 아이템 (레미콘/철근)
    { width: 20 }, // C: 구분 (24MPa, H10 등)
    { width: 10 }  // D: 단위
  ];
  floors.forEach(() => cols.push({ width: 10 })); // 층별 열
  cols.push({ width: 12 }); // 합계
  cols.push({ width: 12 }); // 비고
  ws.columns = cols;

  // 2. 상단 텍스트 추가 (1~2행)
  const r1 = ws.addRow(["QS 분석용 프로젝트 비교 템플릿"]);
  r1.height = 30;
  ws.mergeCells(1, 1, 1, 4 + floors.length + 2); // 첫 행 전체 병합
  r1.getCell(1).font = { size: 14, bold: true, name: '맑은 고딕' };
  r1.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };

  const r2 = ws.addRow([]);
  r2.height = 20;
  const r2EndCol = 4 + floors.length + 1;
  r2.getCell(r2EndCol).value = "할증 후 수량 기준";
  r2.getCell(r2EndCol).font = { size: 10, name: '맑은 고딕' };
  r2.getCell(r2EndCol).alignment = { vertical: 'middle', horizontal: 'right' };
  ws.mergeCells(2, r2EndCol - 1, 2, r2EndCol); // 할증 텍스트 우측 병합

  // 3. 헤더 추가 (3~4행)
  const r3 = ws.addRow(["동", "아이템", "구분", "단위", "현재 프로젝트"]);
  const r4 = ws.addRow(["", "", "", "", ...floors, "합계", "비고"]);
  r3.height = 25; r4.height = 25;

  // 헤더 병합 처리
  ws.mergeCells(3, 1, 4, 1); // 동
  ws.mergeCells(3, 2, 4, 2); // 아이템
  ws.mergeCells(3, 3, 4, 3); // 구분
  ws.mergeCells(3, 4, 4, 4); // 단위
  ws.mergeCells(3, 5, 3, 4 + floors.length + 1); // "현재 프로젝트" 가로 병합
  ws.mergeCells(3, 4 + floors.length + 2, 4, 4 + floors.length + 2); // 비고

  // 헤더 스타일 (테두리 및 배경색)
  const headerBorder = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
  for(let r=3; r<=4; r++) {
    for(let c=1; c<=4 + floors.length + 2; c++) {
      const cell = ws.getCell(r, c);
      cell.font = { bold: true, size: 10, name: '맑은 고딕' };
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { arg: 'FFEAEAEA' } }; // 옅은 회색
      cell.border = headerBorder;
    }
  }

  // 4. 데이터 추가 및 그룹화
  const dataBorder = { top:{style:'thin', color:{argb:'FFBFBFBF'}}, left:{style:'thin', color:{argb:'FFBFBFBF'}}, bottom:{style:'thin', color:{argb:'FFBFBFBF'}}, right:{style:'thin', color:{argb:'FFBFBFBF'}} };

  state.dongs.sort().forEach(dong => {
    const dongData = state.data[dong] || {};
    const grouped = {};
    state.mappings.forEach(m => {
      const qByF = dongData[m.original] || {}; if (Object.keys(qByF).length === 0) return;
      if (!grouped[m.canonical]) grouped[m.canonical] = { category: m.category, floors: {} };
      floors.forEach(f => grouped[m.canonical].floors[f] = (grouped[m.canonical].floors[f] || 0) + (qByF[f] || 0));
    });

    let dongStartRow = ws.rowCount + 1;

    ["콘크리트", "철근", "거푸집"].forEach(cat => {
      const items = Object.keys(grouped).filter(n => grouped[n].category === cat).sort();
      if (items.length === 0) return;

      const catSum = {}; floors.forEach(f => catSum[f] = 0);
      let totalSum = 0;

      // 상세 아이템 (엑셀 그룹화 레벨 1)
      items.forEach(name => {
        const item = grouped[name];
        const unit = cat==='철근'?'TON':(cat==='콘크리트'?'M3':'M2');
        const rowData = [dong, cat==='콘크리트'?'레미콘':cat, name, unit];
        
        let rowTotal = 0;
        floors.forEach(f => {
          const val = item.floors[f] || 0;
          rowData.push(val);
          catSum[f] += val; rowTotal += val;
        });
        rowData.push(rowTotal);
        
        const r = ws.addRow(rowData);
        r.height = 18;
        r.outlineLevel = 1; // [-] 그룹 묶기 지정
        totalSum += rowTotal;

        r.eachCell((cell, colNum) => {
          cell.border = dataBorder; cell.font = { name: '맑은 고딕', size: 10 };
          if (colNum <= 4) cell.alignment = { vertical: 'middle', horizontal: 'center' };
          else { cell.alignment = { vertical: 'middle', horizontal: 'right' }; cell.numFmt = '#,##0.000'; }
        });
      });

      // 카테고리 합계 (엑셀 그룹화 레벨 0 - 펼쳤을때 보이는 기준줄)
      const sumRowData = [dong, cat==='콘크리트'?'레미콘':cat, "합계", cat==='철근'?'TON':(cat==='콘크리트'?'M3':'M2')];
      floors.forEach(f => sumRowData.push(catSum[f]));
      sumRowData.push(totalSum);
      
      const sumRow = ws.addRow(sumRowData);
      sumRow.height = 18;
      sumRow.outlineLevel = 0; 
      sumRow.eachCell((cell, colNum) => {
        cell.font = { name: '맑은 고딕', size: 10, bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { arg: 'FFF2F2F2' } }; // 짙은 회색 배경
        cell.border = dataBorder;
        if (colNum <= 4) cell.alignment = { vertical: 'middle', horizontal: 'center' };
        else { cell.alignment = { vertical: 'middle', horizontal: 'right' }; cell.numFmt = '#,##0.000'; }
      });

      // 비율 계산 (철근 톤당 루베)
      if (cat === '철근') {
        const ratioRowData = [dong, "레미콘/철근", "비율", "M3/TON"];
        floors.forEach(f => {
          const cSum = Object.keys(grouped).filter(n=>grouped[n].category==='콘크리트').reduce((s,n)=>s+grouped[n].floors[f],0);
          ratioRowData.push(cSum > 0 ? (catSum[f] / cSum) : 0);
        });
        ratioRowData.push(""); // 마지막 합계란 비움
        const ratioRow = ws.addRow(ratioRowData);
        ratioRow.height = 18;
        ratioRow.eachCell((cell, colNum) => {
          cell.font = { name: '맑은 고딕', size: 10, bold: true, color: { arg: 'FFC00000' } }; // 붉은색 글씨
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { arg: 'FFFFE699' } }; // 주황/노란색 배경
          cell.border = dataBorder;
          if (colNum <= 4) cell.alignment = { vertical: 'middle', horizontal: 'center' };
          else { cell.alignment = { vertical: 'middle', horizontal: 'right' }; cell.numFmt = '#,##0.0000'; }
        });
      }
    });

    // 동 이름 수직 병합 (A열)
    let dongEndRow = ws.rowCount;
    if (dongStartRow < dongEndRow) {
      ws.mergeCells(dongStartRow, 1, dongEndRow, 1);
      const mergedCell = ws.getCell(dongStartRow, 1);
      mergedCell.alignment = { vertical: 'middle', horizontal: 'center' };
      mergedCell.font = { name: '맑은 고딕', size: 10, bold: true };
    }
  });

  // Blob 생성 및 다운로드 (FileSaver.js 활용)
  const buffer = await wb.xlsx.writeBuffer();
  saveAs(new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), "QS_통합템플릿_리포트.xlsx");
};

function switchTab(id) {
  document.querySelectorAll(".tab, .tab-panel").forEach(el => el.classList.remove("is-active"));
  document.querySelector(`[data-tab="${id}"]`).classList.add("is-active");
  $("tab-" + id).classList.add("is-active");
}
document.querySelectorAll(".tab").forEach(t => t.onclick = () => switchTab(t.dataset.tab));
