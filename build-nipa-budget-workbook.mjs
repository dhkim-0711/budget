import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';

const root = process.cwd();
const source = JSON.parse(fs.readFileSync(path.join(root, 'nipa-ai-budget-projects.json'), 'utf8'));
const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const hqMatch = html.match(/const NIPA_HQ_PROJECTS = (\[[\s\S]*?\]);\s*const NIPA_HEADQUARTERS/);
const hqRows = hqMatch ? JSON.parse(hqMatch[1]) : [];
const normalize = (value) => String(value || '').replace(/^\(내역\)\s*/, '').trim();
const keyOf = (row) => [row.code, normalize(row.parentTitle || row.title), normalize(row.detailTitle || row.title)].join('|');
const hqMap = new Map();
for (const row of hqRows) {
  const key = keyOf(row);
  const values = hqMap.get(key) || new Set();
  if (row.nipaHeadquarters) values.add(row.nipaHeadquarters);
  hqMap.set(key, values);
}

const layers = ['AI인프라','AI반도체','AI모델','지역AX','피지컬AI','글로벌 이니셔티브','AI생태계','산업전략','기관운영비'];
const headquarters = ['정책기획단','SW융합본부','AI인프라본부','AI반도체본부','AI활용본부','지역AX본부','글로벌본부','경영기획본부'];
const accounts = ['일반회계','특별회계','기금'];
const headers = ['관리상태','사업명','세부사업명','사업코드','회계','레이어','과기정통부 소관','NIPA 본부','2026년 예산(백만원)','개요','주요내용','추진방향'];
const rows = source.rows.map((row) => ({
  관리상태: '기존',
  사업명: row.detailTitle || normalize(row.title),
  세부사업명: row.parentTitle || row.title,
  사업코드: row.code,
  회계: row.account,
  레이어: row.layer,
  '과기정통부 소관': row.mappedOrgName,
  'NIPA 본부': [...(hqMap.get(keyOf(row)) || [])].join(' · '),
  '2026년 예산(백만원)': Number(row.budget) || 0,
  개요: row.overview || row.detail?.overview || '',
  주요내용: Array.isArray(row.sourceMainItems) ? row.sourceMainItems.join('\n') : (row.mainPoints || ''),
  추진방향: row.direction || row.detail?.direction || '',
}));

const wb = new ExcelJS.Workbook();
wb.creator = 'NIPA Budget Dashboard';
wb.created = new Date();
wb.calcProperties.fullCalcOnLoad = true;
wb.views = [{x: 0, y: 0, width: 16000, height: 9000, firstSheet: 0, activeTab: 0}];

const navy = 'FF12375B', blue = 'FF245F9E', pale = 'FFEAF2F9', line = 'FFD5E0EA', white = 'FFFFFFFF', muted = 'FF60758A';
const thinBorder = {top:{style:'thin',color:{argb:line}},left:{style:'thin',color:{argb:line}},bottom:{style:'thin',color:{argb:line}},right:{style:'thin',color:{argb:line}}};
const title = (ws, text, sub) => {
  ws.mergeCells('B2:L3'); ws.getCell('B2').value=text; ws.getCell('B2').font={name:'맑은 고딕',size:22,bold:true,color:{argb:navy}}; ws.getCell('B2').alignment={vertical:'middle'};
  ws.mergeCells('B4:L4'); ws.getCell('B4').value=sub; ws.getCell('B4').font={name:'맑은 고딕',size:10,color:{argb:muted}};
};
const nav = (ws) => {
  const links=[['B6','대시보드',"#'대시보드'!A1"],['D6','사업 DB',"#'사업 DB'!A1"],['F6','신규입력 안내',"#'신규입력 안내'!A1"],['H6','레이어 현황',"#'레이어 현황'!A1"],['J6','본부별 현황',"#'본부별 현황'!A1"]];
  for(const [cell,label,target] of links){ws.mergeCells(`${cell}:${String.fromCharCode(cell.charCodeAt(0)+1)}6`);const c=ws.getCell(cell);c.value={text:label,hyperlink:target};c.font={name:'맑은 고딕',bold:true,color:{argb:white}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:blue}};c.alignment={horizontal:'center',vertical:'middle'};c.border=thinBorder;}
  ws.getRow(6).height=25;
};

const dashboard = wb.addWorksheet('대시보드',{views:[{showGridLines:false}]});
const db = wb.addWorksheet('사업 DB',{views:[{state:'frozen',ySplit:8,xSplit:4}]});
title(db,'NIPA 2026 예산 사업 DB','기존 행은 직접 수정하고, 표 하단의 빈 행에는 신규 사업을 입력하세요. 필터와 정렬은 표 머리글에서 사용할 수 있습니다.'); nav(db);
db.addTable({name:'사업DB',ref:'A8',headerRow:true,totalsRow:false,style:{theme:'TableStyleMedium2',showRowStripes:true},columns:headers.map(name=>({name})),rows:[...rows.map(row=>headers.map(h=>row[h])),...Array.from({length:50},()=>Array(headers.length).fill(''))]});
db.columns=[7,30,35,13,12,20,48,22,20,45,55,55].map(width=>({width}));
db.getRow(8).height=32; db.getRow(8).font={bold:true,color:{argb:white}}; db.getRow(8).alignment={horizontal:'center',vertical:'middle',wrapText:true};
for(let r=9;r<=8+rows.length+50;r++){
  db.getRow(r).height=42;
  for(let c=1;c<=headers.length;c++) db.getCell(r,c).alignment={vertical:'top',wrapText:c>=7};
  db.getCell(r,1).dataValidation={type:'list',allowBlank:true,formulae:['"기존,신규,검토"']};
  db.getCell(r,5).dataValidation={type:'list',allowBlank:true,formulae:[`"${accounts.join(',')}"`]};
  db.getCell(r,6).dataValidation={type:'list',allowBlank:true,formulae:[`"${layers.join(',')}"`]};
  db.getCell(r,9).numFmt='#,##0';
}
db.autoFilter={from:'A8',to:`L${8+rows.length+50}`};
db.addConditionalFormatting({ref:`A9:A${8+rows.length+50}`,rules:[{type:'containsText',operator:'containsText',text:'신규',style:{fill:{type:'pattern',pattern:'solid',bgColor:{argb:'FFDDF4E7'}}}}]});

title(dashboard,'정보통신산업진흥원 2026 예산 DB 현황판','오프라인 Excel 버전 · 사업 DB를 수정하면 요약 수식이 자동 갱신됩니다.'); nav(dashboard);
dashboard.columns=Array.from({length:12},()=>({width:13}));
const cards=[['B','D','사업 행 수',{formula:'COUNTIF(사업DB[사업코드],"?*")'}],['E','G','총예산(백만원)',{formula:'SUM(사업DB[2026년 예산(백만원)])'}],['H','J','레이어 수',{formula:'COUNTIF(\'레이어 현황\'!C9:C17,">0")'}]];
for(const [from,to,label,value] of cards){dashboard.mergeCells(`${from}8:${to}8`);dashboard.mergeCells(`${from}9:${to}10`);const heading=dashboard.getCell(`${from}8`),c=dashboard.getCell(`${from}9`);heading.value=label;heading.font={name:'맑은 고딕',size:10,bold:true,color:{argb:muted}};heading.alignment={horizontal:'center',vertical:'middle'};c.value=value;c.numFmt='#,##0';c.font={name:'맑은 고딕',size:20,bold:true,color:{argb:navy}};c.alignment={horizontal:'center',vertical:'middle'};for(let r=8;r<=10;r++)for(let col=dashboard.getColumn(from).number;col<=dashboard.getColumn(to).number;col++){dashboard.getCell(r,col).fill={type:'pattern',pattern:'solid',fgColor:{argb:pale}};dashboard.getCell(r,col).border=thinBorder;}}
dashboard.getCell('B7').value='핵심 지표';dashboard.getCell('B7').font={bold:true,size:14,color:{argb:navy}};
dashboard.getCell('B12').value='레이어별 예산';dashboard.getCell('B12').font={bold:true,size:14,color:{argb:navy}};
dashboard.getCell('B13').value='레이어';dashboard.getCell('E13').value='사업 수';dashboard.getCell('G13').value='예산(백만원)';
for(let i=0;i<layers.length;i++){const r=14+i;dashboard.mergeCells(`B${r}:D${r}`);dashboard.getCell(`B${r}`).value=layers[i];dashboard.mergeCells(`E${r}:F${r}`);dashboard.getCell(`E${r}`).value={formula:`COUNTIF(사업DB[레이어],B${r})`};dashboard.mergeCells(`G${r}:J${r}`);dashboard.getCell(`G${r}`).value={formula:`SUMIF(사업DB[레이어],B${r},사업DB[2026년 예산(백만원)])`};dashboard.getCell(`G${r}`).numFmt='#,##0';for(let c=2;c<=10;c++)dashboard.getCell(r,c).border=thinBorder;}
dashboard.addConditionalFormatting({ref:'G14:G22',rules:[{type:'colorScale',cfvo:[{type:'min'},{type:'max'}],color:[{argb:'FFFFFFFF'},{argb:'FF7CB5E8'}]}]});

const makeSummary=(name,items,field,wildcard=false)=>{const ws=wb.addWorksheet(name,{views:[{state:'frozen',ySplit:8,showGridLines:false}]});title(ws,`NIPA 2026 ${name}`,'사업 DB의 현재 값을 기준으로 자동 집계됩니다.');nav(ws);ws.columns=[4,28,16,20,55].map(width=>({width}));ws.addRow([]);ws.addRow(['','구분','사업 수','예산(백만원)','활용 안내']);ws.getRow(8).font={bold:true,color:{argb:white}};ws.getRow(8).fill={type:'pattern',pattern:'solid',fgColor:{argb:blue}};items.forEach((item,i)=>{const r=9+i;ws.getCell(r,2).value=item;const criteria=wildcard?`"*"&B${r}&"*"`:`B${r}`;ws.getCell(r,3).value={formula:`COUNTIF(사업DB[${field}],${criteria})`};ws.getCell(r,4).value={formula:`SUMIF(사업DB[${field}],${criteria},사업DB[2026년 예산(백만원)])`};ws.getCell(r,4).numFmt='#,##0';ws.getCell(r,5).value='사업 DB에서 해당 값을 필터하면 상세 목록을 확인할 수 있습니다.';for(let c=2;c<=5;c++){ws.getCell(r,c).border=thinBorder;ws.getCell(r,c).alignment={vertical:'middle',wrapText:true};}});ws.addConditionalFormatting({ref:`D9:D${8+items.length}`,rules:[{type:'colorScale',cfvo:[{type:'min'},{type:'max'}],color:[{argb:'FFFFFFFF'},{argb:'FF7CB5E8'}]}]});return ws;};
makeSummary('레이어 현황',layers,'레이어');
makeSummary('본부별 현황',headquarters,'NIPA 본부',true);

const guide=wb.addWorksheet('신규입력 안내',{views:[{showGridLines:false}]});title(guide,'신규 사업 추가·기존 사업 수정','엑셀의 직접 편집 장점을 활용하는 업무 절차입니다.');nav(guide);guide.columns=[4,28,85].map(width=>({width}));
const guides=[['신규 사업 추가','사업 DB 표의 마지막 빈 행에 값을 입력하고 관리상태를 “신규”로 선택합니다. 사업명과 세부사업명이 다르면 내역사업으로 관리할 수 있습니다.'],['기존 사업 수정','사업 DB에서 사업코드·사업명으로 행을 찾은 뒤 해당 셀을 직접 수정합니다. 변경 검토가 필요하면 관리상태를 “검토”로 바꿉니다.'],['분류 선택','회계와 레이어는 드롭다운을 사용합니다. 공동 담당 본부는 “AI반도체본부 · 글로벌본부”처럼 가운데점으로 구분해 입력할 수 있습니다.'],['백업 권장','수정 전 파일을 다른 이름으로 저장해 버전을 남기세요. HTML 대시보드로 반영하려면 수정된 사업 DB 시트를 DB 관리 탭에 업로드합니다.']];
guide.getCell('B8').value='업무';guide.getCell('C8').value='사용 방법';for(const c of ['B8','C8']){guide.getCell(c).font={bold:true,color:{argb:white}};guide.getCell(c).fill={type:'pattern',pattern:'solid',fgColor:{argb:blue}};guide.getCell(c).border=thinBorder;}
guides.forEach((item,i)=>{const r=9+i;guide.getCell(r,2).value=item[0];guide.getCell(r,3).value=item[1];guide.getRow(r).height=55;for(let c=2;c<=3;c++){guide.getCell(r,c).border=thinBorder;guide.getCell(r,c).alignment={vertical:'middle',wrapText:true};}});

wb.worksheets.forEach(ws=>{ws.properties.defaultRowHeight=20;ws.pageSetup={orientation:'landscape',fitToPage:true,fitToWidth:1,fitToHeight:0};});
const out=path.join(root,'NIPA_2026_예산_DB_관리.xlsx');
await wb.xlsx.writeFile(out);
console.log(JSON.stringify({output:out,rows:rows.length,sheets:wb.worksheets.map(s=>s.name)},null,2));
