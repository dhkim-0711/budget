import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';

const root = process.cwd();
const sourcePath = path.join(root, 'ai-budget-ecosystem-dashboard.html');
const outputPath = path.join(root, 'nipa-ai-budget-ecosystem-dashboard.html');
const indexPath = path.join(root, 'index.html');
const dataPath = path.join(root, 'nipa-ai-budget-projects.json');
const verificationPath = path.join(root, 'nipa-budget-verification.json');
const operatorName = '정보통신산업진흥원';

if (!fs.existsSync(sourcePath)) {
  throw new Error(`기준 대시보드를 찾을 수 없습니다: ${sourcePath}`);
}

let html = fs.readFileSync(sourcePath, 'utf8');

function parseEmbeddedJson(name, nextName) {
  const pattern = new RegExp(`const ${name} = ([\\s\\S]*?);\\r?\\nconst ${nextName} =`);
  const match = html.match(pattern);
  if (!match) throw new Error(`${name} 데이터를 기준 대시보드에서 찾지 못했습니다.`);
  return JSON.parse(match[1]);
}

function replaceEmbeddedJson(name, nextName, value) {
  const pattern = new RegExp(`const ${name} = [\\s\\S]*?;\\r?\\nconst ${nextName} =`);
  html = html.replace(pattern, `const ${name} = ${JSON.stringify(value)};\nconst ${nextName} =`);
}

const allProjects = parseEmbeddedJson('PROJECTS', 'LAYERS');
const originalLayers = parseEmbeddedJson('LAYERS', 'ORGS');
const originalOrgs = parseEmbeddedJson('ORGS', 'ORG_TREE');
const originalOrgTreeMatch = html.match(/const ORG_TREE = ([\s\S]*?);\r?\n/);
if (!originalOrgTreeMatch) throw new Error('ORG_TREE 데이터를 기준 대시보드에서 찾지 못했습니다.');
const originalOrgTree = JSON.parse(originalOrgTreeMatch[1]);
let aiProjects = allProjects.filter((project) => String(project.operator || '').includes(operatorName));
const knownOrgPaths = [...new Set(allProjects.map((project) => project.mappedOrgName).filter((value) => value?.includes('/')))];
const knownOrgLeaves = knownOrgPaths.map((orgPath) => ({
  orgPath,
  leaf: orgPath.split('/').at(-1).trim(),
})).filter((item) => item.leaf.length >= 3);

function resolveSourceOrganization(sourceLines, sourceLine, focusTitle = '') {
  const headerLines = sourceLines.slice(Math.max(0, sourceLine - 2), sourceLine + 180);
  const header = headerLines.join(' ');
  const focus = String(focusTitle || '').replace(/[^\p{L}\p{N}]/gu, '');
  const focusGrams = Array.from({ length: Math.max(0, focus.length - 2) }, (_, index) => focus.slice(index, index + 3));
  const sections = [];
  for (let i = 0; i < headerLines.length; i += 1) {
    if (headerLines[i].trim() !== '소관부처') continue;
    const preceding = headerLines.slice(Math.max(0, i - 5), i).join('').replace(/[^\p{L}\p{N}]/gu, '');
    const score = focusGrams.reduce((sum, gram) => sum + (preceding.includes(gram) ? 1 : 0), 0);
    const departmentLines = [];
    for (const line of headerLines.slice(i + 1, i + 8)) {
      const value = line.trim();
      if (!value || /사업시행주체|^가\.\s|예산 총괄표|지출계획 총괄표/.test(value)) break;
      departmentLines.push(value);
    }
    sections.push({ score, text: departmentLines.join(' ') });
  }
  const departmentText = sections.sort((a, b) => b.score - a.score)[0]?.text || header;
  const departmentCompact = departmentText.replace(/\s+/g, '');
  const matches = knownOrgLeaves
    .filter((item) => departmentCompact.includes(item.leaf.replace(/\s+/g, '')))
    .sort((a, b) => b.leaf.length - a.leaf.length);
  let mappedOrgName = '';
  if (/국제협력총괄담당관/.test(departmentText)) mappedOrgName = '제1차관 / 기획조정실 / 국제협력관 / 국제협력총괄담당관';
    else if (/구주아프리카협력담당관/.test(departmentText)) mappedOrgName = '제1차관 / 기획조정실 / 국제협력관 / 구주아프리카협력담당관';
    else if (/다자협력담당관/.test(departmentText)) mappedOrgName = '제1차관 / 기획조정실 / 국제협력관 / 다자협력담당관';
    else if (/디지털신산업제도과/.test(departmentText)) mappedOrgName = '제2차관 / 정보통신정책실 / 정보통신정책관 / 디지털신산업제도과';
    else if (/정보통신정책총괄과/.test(departmentText)) mappedOrgName = '제2차관 / 정보통신정책실 / 정보통신정책관 / 정보통신정책총괄과';
    else if (/디지털플랫폼팀/.test(departmentText)) mappedOrgName = '제2차관 / 정보보호네트워크정책실 / 통신정책관 / 디지털플랫폼팀';
    else if (/정보통신방송기술정책과/.test(departmentText)) mappedOrgName = '제2차관 / 정보통신정책실 / 정보통신산업정책관 / 정보통신방송기술정책과';
    else if (/정보통신산업정책과/.test(departmentText)) mappedOrgName = '제2차관 / 정보통신정책실 / 정보통신산업정책관 / 정보통신산업정책과';
    else if (/정보통신산업기반과/.test(departmentText)) mappedOrgName = '제2차관 / 정보통신정책실 / 정보통신산업정책관 / 정보통신산업기반과';
    else if (/디바이스AX혁신팀/.test(departmentText)) mappedOrgName = '제2차관 / 정보통신정책실 / 정보통신산업정책관 / 디바이스AX혁신팀';
    else if (/디지털콘텐츠과/.test(departmentText)) mappedOrgName = '제2차관 / 정보통신정책실 / 소프트웨어정책관 / 디지털콘텐츠과';
    else if (/인공지능컴퓨팅인프라팀/.test(departmentText)) mappedOrgName = '제2차관 / 인공지능정책실 / 인공지능인프라정책관 / 인공지능컴퓨팅인프라팀';
    else if (/인공지능전환지원과/.test(departmentText)) mappedOrgName = '제2차관 / 인공지능정책실 / 인공지능인프라정책관 / 인공지능전환지원과';
    else if (/인공지능융합팀/.test(departmentText)) mappedOrgName = '제2차관 / 인공지능정책실 / 인공지능정책기획관 / 인공지능융합팀';
    else if (/디지털인재양성과/.test(departmentText)) mappedOrgName = '제2차관 / 인공지능정책실 / 인공지능정책기획관 / 디지털인재양성과';
    else if (/AI데이터진흥과|인공지능데이터진흥과/.test(departmentText)) mappedOrgName = '제2차관 / 인공지능정책실 / 인공지능인프라정책관 / 인공지능데이터진흥과';
    else if (/인공지능기반정책과/.test(departmentText)) mappedOrgName = '제2차관 / 인공지능정책실 / 인공지능기반정책관 / 인공지능기반정책과';
    else if (/디지털사회기획과/.test(departmentText)) mappedOrgName = '제2차관 / 정보통신정책실 / 정보통신정책관 / 디지털사회기획과';
    else if (/정보통신정책실[\s\S]{0,80}정보통신산업정책관/.test(departmentText)) mappedOrgName = '제2차관 / 정보통신정책실 / 정보통신산업정책관';
    else mappedOrgName = matches[0]?.orgPath || '';
  if (/3D프린팅/.test(focusTitle) && /정보통신산업정책관$/.test(mappedOrgName)) {
    mappedOrgName = '제2차관 / 정보통신정책실 / 정보통신산업정책관 / 정보통신산업기반과';
  }
  if (/ICT문화융합센터/.test(focusTitle) && /소프트웨어정책관$/.test(mappedOrgName)) {
    mappedOrgName = '제2차관 / 정보통신정책실 / 소프트웨어정책관 / 소프트웨어산업과';
  }
  if (/국가\s*NPU전용\s*컴퓨팅센터\s*광주\s*설립/.test(focusTitle)) {
    mappedOrgName = '제2차관 / 정보통신정책실 / 정보통신산업정책관 / 정보통신방송기술정책과';
  }
  if (!mappedOrgName) return null;
  const orgPathKey = originalOrgs
    .map((org) => org.id)
    .filter((id) => mappedOrgName === id || mappedOrgName.startsWith(`${id} /`))
    .sort((a, b) => b.length - a.length)[0] || mappedOrgName.split(' / ').slice(0, 3).join(' / ');
  return { mappedOrgName, orgPathKey };
}

function accountFromFile(file) {
  if (file.includes('특별회계')) return '특별회계';
  if (file.includes('일반회계')) return '일반회계';
  if (file.includes('기금')) return '기금';
  return '일반회계';
}

function legacyClassifyNipaLayer(title) {
  if (/진흥원지원|전담기관.*(?:인건비|운영비)|총액대상.*인건비/i.test(title)) return { id: 'OUT', name: '분류외·기관운영비', color: '#6b7280' };
  if (/보안|정보보호|블록체인|디지털자산|공급망|안전|신뢰/i.test(title)) return { id: 'X1', name: '평가·안전·보안', color: '#bc5090' };
  if (/규제|정책|표준|거버넌스/i.test(title)) return { id: 'X2', name: '거버넌스·주권', color: '#7a7f9a' };
  if (/인재|교육|대학|진로|역량강화/i.test(title)) return { id: 'L4', name: '모델·기초연구', color: '#43aa8b' };
  if (/반도체|NPU|컴퓨팅|클라우드|데이터센터|인프라/i.test(title)) return { id: 'L2', name: '컴퓨팅·하드웨어', color: '#0081a7' };
  if (/데이터|API/i.test(title)) return { id: 'L3', name: '데이터', color: '#2a9d8f' };
  if (/콘텐츠|미디어|VR|AR|메타버스|가상융합|공간컴퓨팅|홀로그램|XaaS/i.test(title)) return { id: 'L5', name: '에이전트·시스템레이어', color: '#6d597a' };
  if (/지역|산업|기업|창업|해외|글로벌|바우처|공공SW|규제샌드박스|지원센터|클러스터|생태계|상용화|진출/i.test(title)) return { id: 'L6', name: '기업·공공 AX', color: '#f4a261' };
  if (/피지컬|자율|로봇|3D프린팅/i.test(title)) return { id: 'L8', name: '피지컬 AI', color: '#ef476f' };
  return { id: 'L7', name: '산업 버티컬', color: '#e76f51' };
}

function classifyNipaLayer(title) {
  if (/정보통신산업진흥원\s*지원|SW정책연구소\s*운영|소프트웨어정책연구소\s*운영|전담기관.*(?:인건비|운영비)|총액대상.*인건비/i.test(title)) return { id: '기관운영비', name: '기관운영비', color: '#6b7280' };
  if (/인간\s*[-–]?\s*AI\s*협업형\s*LAM\s*개발.*글로벌\s*실증/i.test(title)) return { id: '피지컬AI', name: '피지컬AI', color: '#ef476f' };
  if (/글로벌/i.test(title)) return { id: '글로벌 이니셔티브', name: '글로벌 이니셔티브', color: '#6d597a' };
  if (/피지컬\s*AI|로봇|자율주행|무인|드론|스마트팩토리|메타팩토리|디지털트윈|협업지능|3D프린팅/i.test(title)) return { id: '피지컬AI', name: '피지컬AI', color: '#ef476f' };
  if (/AI\s*반도체|인공지능\s*반도체|NPU|온디바이스|PIM|신경망처리|반도체.*실증/i.test(title)) return { id: 'AI반도체', name: 'AI반도체', color: '#0081a7' };
  if (/GPU|컴퓨팅|클라우드|데이터센터|서버|데이터|AI허브|AI-HUB|고성능컴퓨팅|초고속망|인프라/i.test(title)) return { id: 'AI인프라', name: 'AI인프라', color: '#355c7d' };
  if (/파운데이션|거대언어|초거대|생성형|AGI|AI모델|인공지능모델|에이전틱|에이전트|모델.*개발|학습.*모델/i.test(title)) return { id: 'AI모델', name: 'AI모델', color: '#43aa8b' };
  if (/지역|권역|지역SW|지역디지털|이노베이션스퀘어|지역.*AX|AX.*지역|스마트시티|지역특화/i.test(title)) return { id: '지역AX', name: '지역AX', color: '#f4a261' };
  if (/글로벌|해외|수출|국제협력|ODA|진출|K-?테크|월드IT|이니셔티브/i.test(title)) return { id: '글로벌 이니셔티브', name: '글로벌 이니셔티브', color: '#6d597a' };
  if (/창업|인재|교육|대학|생태계|오픈소스|바우처|기업지원|사업화|콘텐츠|미디어|디지털포용|안전|보안|신뢰|규제|표준/i.test(title)) return { id: 'AI생태계', name: 'AI생태계', color: '#2a9d8f' };
  return { id: '산업전략', name: '산업전략', color: '#e76f51' };
}

function nipaLayer(id) {
  const colors = {
    AI인프라: '#355c7d', AI반도체: '#0081a7', AI모델: '#43aa8b', 지역AX: '#f4a261',
    피지컬AI: '#ef476f', '글로벌 이니셔티브': '#6d597a', AI생태계: '#2a9d8f',
    산업전략: '#e76f51', 기관운영비: '#6b7280',
  };
  return { id, name: id, color: colors[id] };
}

function policyPurposeLayer(project, fallback) {
  const title = [project.title, project.detailTitle].filter(Boolean).join(' ');

  // 제도·협력·창업·기업성장·안전 기반은 개별 기술보다 생태계 조성이 목적이다.
  if (/규제(?:샌드박스|혁신)|AI\s*산[·ㆍ]?학[·ㆍ]?연\s*협력|ICT창의기업육성|AI[·ㆍ]?공간컴퓨팅\s*창업\s*생태계|AI\s*특화지구\s*조성|디지털\s*공공\s*혁신생태계|디지털인프라\s*\(SW\)\s*진단\s*및\s*개선|SW산업기술확산역량강화|ICT문화융합센터|SW고성장클럽|SW기업성장\s*촉진지원/i.test(title)) {
    return nipaLayer('AI생태계');
  }

  // 지역 정착·지역 인재·지역 기업 육성은 기술 수단과 무관하게 지역AX가 목적이다.
  if (/지역디지털인재양성|소프트웨어인재키움|체인지업그라운드\s*지역확산|지역디지털사업\s*운영|디지털혁신거점\s*조성지원|생성형\s*AI기반\s*가상융합\s*콘텐츠\s*제작센터/i.test(title)) {
    return nipaLayer('지역AX');
  }

  if (/지역특화\s*AI\s*공간컴퓨팅\s*육성/i.test(title)) return nipaLayer('피지컬AI');

  // 모델·에이전트 자체 역량 확보가 핵심인 사업만 AI모델로 분류한다.
  if (/AI\s*AGENT\s*선도\s*국가|공공AX\s*프로젝트|AI일상화\s*확산|부처협업\s*기반\s*AI확산/i.test(title)) return nipaLayer('AI모델');

  // 특정 산업의 제품·서비스 개발, 실증, 사업화는 활용 기술보다 산업 목적을 우선한다.
  if (/VR[·ㆍ]?AR콘텐츠산업육성|디지털콘텐츠기업경쟁력강화|디지털\s*안전\s*선도모델|3D프린팅산업\s*육성기반|의료AI혁신생태계조성|K-Health\s*국민의료|AI\s*기반\s*뇌발달질환|AI\s*응용제품\s*신속\s*상용화|홀로그램기술\s*사업화|AI정밀의료SW|K-CareNetwork|데이터\s*활용\s*의료|디지털헬스케어기반\s*AI|XaaS\s*선도|AI융합\s*지능형\s*농업|제조업\s*AI융합|스마트물류\s*기술|어장\s*공간정보|농식품\s*분야\s*메타버스|디지털트윈\s*융합\s*의료|3D프린팅\s*의료공동제조소/i.test(title)) {
    return nipaLayer('산업전략');
  }

  if (/AX실증밸리조성/i.test(title)) return nipaLayer('AI인프라');

  // 국제회의·개도국 협력·현지거점은 인프라라는 표현이 있어도 대외협력이 목적이다.
  if (/디지털\s*장관회의|정보통신전문가\s*초청연수|K-Lab\s*설치|현지거점\s*활용\s*인프라/i.test(title)) {
    return nipaLayer('글로벌 이니셔티브');
  }
  return fallback;
}

function moneyTokens(text) {
  return [...String(text || '').matchAll(/(?:△\s*)?\d{1,3}(?:,\d{3})+|(?:△\s*)?\d+/g)]
    .map((match) => ({ raw: match[0], value: Number(match[0].replace(/[△,\s]/g, '')) }))
    .filter((token) => Number.isFinite(token.value) && token.value > 0 && token.value < 10000000 && ![2022, 2023, 2024, 2025, 2026].includes(token.value));
}

function extractBlockBudget(lines) {
  const titleLine = lines[0] || '';
  const code = titleLine.match(/(\d{4}-\d{3})/)?.[1] || '';
  const manual = new Map([
    ['4234-325', 3000],
    ['4231-300', 2250],
    ['4201-301', /지역 주도 디지털 혁신 지원/.test(titleLine) ? 1800 : null],
    ['2033-506', 600],
    ['2232-423', 1000],
  ]);
  if (/협업지능 피지컬AI 기반 SW플랫폼/.test(titleLine)) return 76666;
  if (manual.has(code) && manual.get(code) != null) return manual.get(code);
  const functional = lines.findIndex((line) => /기능별\s*분류\s*\(합계\)/.test(line));
  if (functional >= 0) {
    const values = lines.slice(functional + 1, functional + 30)
      .map((line) => line.trim())
      .filter((line) => /^\d{1,3}(?:,\d{3})*$/.test(line))
      .map((line) => Number(line.replace(/,/g, '')));
    if (values.length >= 4) return values[values.length - 1];
  }
  const summary = lines.findIndex((line) => /2026.*(?:예산안|지출계획).*총괄표|(?:예산안|지출계획).*총괄표/.test(line));
  if (summary >= 0) {
    const section = lines.slice(summary, summary + 100).join(' ').split(/□\s*기능별|1\)\s*사업목적/)[0];
    const tokens = moneyTokens(section);
    if (tokens.length) {
      const deltaAt = tokens.findIndex((token) => token.raw.includes('△'));
      const usable = deltaAt > 0 ? tokens.slice(0, deltaAt) : tokens;
      const large = usable.filter((token) => token.raw.includes(',') || token.value >= 100);
      if (large.length) {
        const tail = large.slice(-3);
        return tail.length >= 2 ? tail[tail.length - 2].value : tail[0].value;
      }
    }
  }
  return null;
}

function sourceOverview(lines) {
  const start = lines.findIndex((line) => /1\)\s*사업목적/.test(line));
  const candidates = lines.slice(start >= 0 ? start + 1 : 0, start >= 0 ? start + 30 : 40)
    .map((line) => line.replace(/^[□○ㅇ․ㆍ·\-\s]+/, '').trim())
    .filter((line) => line.length >= 18 && !/^(사업기간|총사업비|사업규모|연도|구분|회계|소관)/.test(line));
  return candidates[0] || '정보통신산업진흥원이 사업시행주체로 명시된 2026년 예산 사업입니다.';
}

function cleanNarrativeLine(line) {
  return String(line || '').replace(/^[□○ㅇ․ㆍ·\-※o\s]+/, '').replace(/\s+/g, ' ').trim();
}

function extractNarrative(lines, focusTitle = '') {
  const usable = (line) => {
    const value = cleanNarrativeLine(line);
    return value.length >= 24 && value.length <= 400
      && !/^(연도|사업비|구분|소관부처|사업시행|지원 금액|예산액|본예산|추경|법령상|제\d+조)/.test(value)
      && !/내역사업명구분피보조|지원 금액\(2026|성과지표구분|측정산식|자료수집방법/.test(value)
      && !/^\d{1,3}(?:,\d{3})*(?:백만원|%)?$/.test(value);
  };
  const purposeStart = lines.findIndex((line) => /1\)\s*사업목적|사업목적[·ㆍ]내용/.test(line));
  const purposeEnd = purposeStart >= 0
    ? lines.slice(purposeStart + 1).findIndex((line) => /사업근거 및 추진경위|2\)\s*사업/.test(line))
    : -1;
  const purposeLines = lines.slice(purposeStart >= 0 ? purposeStart + 1 : 0,
    purposeEnd >= 0 ? purposeStart + 1 + purposeEnd : Math.min(lines.length, (purposeStart >= 0 ? purposeStart : 0) + 80));
  const mainStart = lines.findIndex((line) => /□\s*주요내용/.test(line));
  const mainLines = mainStart >= 0 ? lines.slice(mainStart + 1, Math.min(lines.length, mainStart + 100)) : [];
  const focusTokens = focusTitle.replace(/[()]/g, ' ').split(/\s+/).filter((token) => token.length >= 3);
  const compactFocus = focusTitle.replace(/[^\p{L}\p{N}]/gu, '');
  const focusSegments = [compactFocus, ...[...focusTitle.matchAll(/\(([^()]{4,})\)/g)]
    .map((match) => match[1].replace(/[^\p{L}\p{N}]/gu, ''))]
    .filter((value) => value.length >= 4);
  const focusGrams = Array.from({ length: Math.max(0, compactFocus.length - 2) }, (_, index) => compactFocus.slice(index, index + 3));
  const focusScore = (line) => {
    const compact = line.replace(/[^\p{L}\p{N}]/gu, '');
    return focusGrams.reduce((score, gram) => score + (compact.includes(gram) ? 1 : 0), 0);
  };
  const focusedCandidates = [...purposeLines, ...mainLines]
    .filter(usable)
    .map(cleanNarrativeLine)
    .filter((line) => !focusTokens.length || focusTokens.some((token) => line.includes(token)))
    .sort((a, b) => focusScore(b) - focusScore(a));
  const bestFocusScore = focusedCandidates.length ? focusScore(focusedCandidates[0]) : 0;
  const minimumStrongScore = Math.max(2, Math.ceil(focusGrams.length * 0.45));
  const directFocused = focusedCandidates.filter((line) => {
    const compact = line.replace(/[^\p{L}\p{N}]/gu, '');
    return focusSegments.some((segment) => compact.includes(segment))
      || (bestFocusScore >= minimumStrongScore && focusScore(line) >= Math.max(minimumStrongScore, bestFocusScore * 0.72));
  });
  const focused = [...directFocused];
  for (const section of [purposeLines, mainLines]) {
    for (let index = 0; index < section.length; index += 1) {
      const current = cleanNarrativeLine(section[index]);
      if (!directFocused.includes(current)) continue;
      for (let next = index + 1; next < Math.min(section.length, index + 4); next += 1) {
        const raw = section[next];
        if (/^\s*[-○ㅇ]\s*\(/.test(raw)) break;
        if (usable(raw)) focused.push(cleanNarrativeLine(raw));
      }
    }
  }
  const purpose = purposeLines.filter(usable).map(cleanNarrativeLine);
  const main = mainLines.filter(usable).map(cleanNarrativeLine);
  return {
    overview: focused[0] || purpose[0] || main[0] || '',
    purpose: Array.from(new Set(purpose)).slice(0, 4),
    focused: Array.from(new Set(focused)).slice(0, 5),
    main: Array.from(new Set([...focused, ...main])).slice(0, 7),
  };
}

function legacyAnalyzeDirection(project) {
  const subject = project.detailTitle || project.title.replace(/^\(내역\)\s*/, '');
  const strategies = {
    L2: '장비·클라우드 공급량뿐 아니라 실제 가동률, 이용 대기시간, 기업별 활용성과와 국산 컴퓨팅 생태계 연계를 함께 관리해야 합니다.',
    L3: '데이터 구축량보다 품질·표준·재사용률과 산업 현장에서 실제 서비스로 연결된 비율을 핵심 성과로 관리해야 합니다.',
    L4: '교육·인재 배출 인원에 머물지 않고 취업·프로젝트 참여·기업 수요 충족과 후속 성장경로까지 추적해야 합니다.',
    L5: '콘텐츠·서비스 제작 건수보다 상용 배포, 이용자 확보, 반복 사용과 민간 매출 전환 여부를 중심으로 성과를 판단해야 합니다.',
    L6: '지원기업 수보다 현장 도입, 비용 절감, 매출·투자·수출 및 후속 확산으로 이어지는 전환율을 중심으로 사업을 운영해야 합니다.',
    L7: '산업별 규제와 현장 데이터 조건을 반영해 실증 이후 조달·인허가·상용 운영까지 이어지는 전주기 구조를 갖춰야 합니다.',
    L8: '기술개발과 함께 테스트베드 안전성, 현장 운용성, 수요기업의 구매·도입 조건을 동시에 검증해야 합니다.',
    X1: '보안·안전 기준을 사후 점검이 아니라 설계·실증 단계부터 적용하고, 취약점 개선과 인증·확산 성과를 함께 측정해야 합니다.',
    X2: '제도 운영 건수보다 규제 해소 기간, 사업화 전환, 표준·정책의 현장 적용성과를 중심으로 효과를 검증해야 합니다.',
    OUT: '기관 운영비는 인력·경상비·고유사업비를 구분하고, 인력 투입이 주요 사업성과와 어떻게 연결되는지 설명할 필요가 있습니다.',
  };
  const evidence = project.overview || project.sourceOverview || '';
  return `${subject}은(는) ${evidence.replace(/[.。]$/, '')}을(를) 핵심 목적으로 하는 사업입니다. ${strategies[project.layer] || strategies.L7} 2026년 예산 ${Number(project.budget || 0).toLocaleString('ko-KR')}백만원은 세부 집행항목별 산출물과 수혜대상, 완료 시점을 연결해 점검하는 것이 적절합니다.`;
}

function analyzeDirection(project) {
  const subject = project.detailTitle || project.title.replace(/^\(내역\)\s*/, '');
  const text = [subject, project.overview, ...(project.sourceMainItems || [])].join(' ');
  // 부수적으로 등장하는 '글로벌', '지역'보다 사업의 확정 레이어와 명시적 목적을 우선한다.
  if (project.layer === '피지컬AI' && /피지컬\s*AI|LAM|로봇|자율주행|디지털트윈|3D프린팅|스마트물류/i.test(text)) {
    return `${subject}: 모델·시뮬레이션 개발을 실제 장비 제어와 현장 검증까지 연결하는 통합 실증 체계가 필요합니다. 성능 수치뿐 아니라 안전성, 반복 운용 안정성, 수요기업의 도입 가능성을 함께 확인해 사업화로 이어지는 기반을 마련해야 합니다.`;
  }
  if (project.layer === '지역AX') {
    const region = text.match(/제주|세종|광주|강원|대구|부산|울산|경북|경남|전북|전남|충북|충남|경기|인천/)?.[0];
    return `${subject}: ${region ? `${region}의 ` : ''}지역 주력산업과 현안에 맞는 AX 과제를 발굴하고, 지역기업·수요기관이 함께 검증하는 구조로 추진할 필요가 있습니다. 실증 이후의 매출·고용 효과와 다른 권역으로의 확산 가능성을 기준으로 성과를 관리하는 방향이 적절합니다.`;
  }
  if (/VR|AR|가상융합|메타버스|콘텐츠|OTT|미디어/i.test(subject)) {
    return `${subject}: 제작 지원에 그치지 않고 유통 채널 확보, 반복 이용, 기업 매출로 이어지는 사업모델 검증을 강화해야 합니다. 제작 인프라·인력양성·해외진출 지원을 연계해 우수 결과물이 시장에서 지속 활용되도록 추진하는 것이 중요합니다.`;
  }
  const themes = [
    [/GPU 서버확충/i, `${subject}은 도입 물량보다 실제 가동률과 이용 대기시간을 공개하고, 중소 AI기업이 안정적으로 자원을 배정받는 운영체계를 갖추는 데 초점을 둘 필요가 있습니다. 서버 확충분이 모델 개발·서비스 출시로 이어진 비율을 핵심 성과로 관리하는 방향이 적절합니다.`],
    [/GPU 통합환경|통합운영환경/i, `${subject}은 대규모 GPU를 하나의 자원 풀로 묶는 스케줄링·장애대응·사용량 정산 체계를 우선 완성해야 합니다. 클러스터 가동률과 작업 처리시간, 장애 복구시간을 기준으로 운영 품질을 검증하는 것이 중요합니다.`],
    [/GPU임차|GPU 임차/i, `${subject}은 단기 자원 제공에 머물지 않도록 이용기업의 개발 단계와 필요한 연산량을 연동해 배분해야 합니다. 지원 종료 뒤 유료 전환, 투자유치 또는 제품 출시로 이어진 사례를 중심으로 효과를 확인할 필요가 있습니다.`],
    [/AI반도체|NPU|온디바이스/i, `${subject}은 칩 공급과 실증을 분리하지 말고 실제 서비스 워크로드에서 성능·전력효율·호환성을 검증하는 구조로 운영하는 것이 좋습니다. 검증 결과가 구매계약과 반복 도입, 국산 소프트웨어 생태계 확대로 연결되는지를 성과 기준으로 삼아야 합니다.`],
    [/데이터센터/i, `${subject}은 시설 조성 자체보다 전력·냉각 효율, 장비 가동률, 입주기업 활용성과를 함께 관리해야 합니다. 지역 인프라와 국가 AI 컴퓨팅 체계가 중복되지 않도록 기능과 수요권역을 명확히 나누는 방향이 필요합니다.`],
    [/클라우드|SaaS/i, `${subject}은 이용료 지원 중심에서 벗어나 클라우드 전환 설계, 보안·인증, 운영 최적화까지 묶은 지원체계로 고도화할 필요가 있습니다. 도입기업의 업무비용 절감과 지속 이용률, 공급기업의 반복 매출을 함께 추적하는 것이 바람직합니다.`],
    [/의료|헬스|치료|돌봄|심리케어/i, `${subject}은 기술 시연보다 의료·생활 현장에서의 안전성, 유효성, 사용자 수용성을 먼저 입증해야 합니다. 임상·실증 근거와 인허가 준비도, 의료기관의 후속 도입 여부를 연결해 사업 성과를 판단하는 방향이 적절합니다.`],
    [/인재|교육|사관학교|진로|재직자/i, `${subject}은 교육 인원 확대보다 지역·산업 수요에 맞는 프로젝트 경험과 취업·창업 연결을 강화해야 합니다. 수료자의 현장 배치, 기업의 인력 수요 충족도와 교육 이후 성장경로를 추적할 필요가 있습니다.`],
    [/글로벌|해외|수출|국제협력|장관회의|초청연수/i, `${subject}은 행사·상담 건수보다 목표시장별 현지 수요처와 규제·인증 조건을 사전에 좁혀 지원하는 방식이 효과적입니다. 현지 실증 이후 계약, 수출, 공동사업으로 전환된 성과를 중심으로 후속 지원을 차등화해야 합니다.`],
    [/규제|샌드박스|신뢰성|인증|프로세스/i, `${subject}은 제도 운영 건수보다 기업이 실제로 해소한 규제와 인증 획득, 시장 진입까지 걸린 시간을 줄이는 데 집중해야 합니다. 반복되는 애로사항은 개별 특례에 머물지 않고 공통 기준과 제도 개선으로 환류하는 구조가 필요합니다.`],
    [/피지컬AI|협업지능|LAM|로봇|자율주행|디지털트윈|3D프린팅|스마트물류/i, `${subject}은 모델·시뮬레이션 개발과 실제 장비 제어, 안전 검증, 현장 운영을 하나의 실증 체계로 묶어야 합니다. 테스트베드 성능이 수요기업의 구매와 생산공정 적용으로 이어지는지를 명확히 확인하는 방향이 좋습니다.`],
    [/지역|권역|광주|강원|제주|세종|지역AX/i, `${subject}은 지역별 시설·과제를 나열하기보다 특화산업의 수요기업과 지역 디지털기업이 함께 참여하는 대표 프로젝트를 만드는 데 집중할 필요가 있습니다. 지역기업 매출·고용과 수요기관 도입, 다른 권역으로의 확산 가능성을 함께 평가해야 합니다.`],
    [/콘텐츠|가상융합|메타버스|홀로그램|OTT|미디어/i, `${subject}은 콘텐츠 제작 건수보다 실제 유통채널 확보와 반복 이용, 기업 매출로 이어지는 사업모델 검증을 강화해야 합니다. 제작 인프라·인력양성·해외진출 사업을 연계해 우수 결과물이 시장에서 계속 활용되도록 설계하는 것이 좋습니다.`],
    [/창업|벤처|기업성장|사업화|바우처/i, `${subject}은 일률적인 공모 지원보다 기업의 성장단계와 기술성숙도에 따라 실증·투자·판로 지원을 구분해야 합니다. 지원 이후 매출, 투자유치, 신규 고객과 후속 민간자금 유입을 중심으로 선별과 후속지원을 운영할 필요가 있습니다.`],
    [/공공AX|공공 혁신/i, `${subject}은 단발성 AI 구축보다 공공업무 개선 목표와 책임 있는 운영 주체를 먼저 확정해야 합니다. 업무시간 단축, 서비스 품질, 이용자 편익과 다른 기관의 재사용 가능성을 기준으로 확산 여부를 판단하는 것이 바람직합니다.`],
    [/정보통신산업진흥원지원/i, `${subject}은 인력·경상비와 고유사업비를 구분하고, 조직 역량이 주요 사업의 기획·관리 품질 향상에 어떻게 기여하는지 보여줄 필요가 있습니다. 본부별 사업성과와 운영자원을 연결한 설명 체계를 갖추는 것이 적절합니다.`],
    [/SW정책연구소/i, `${subject}은 연구 건수보다 정책 의사결정에 활용된 분석과 산업 현장의 제도 개선 성과를 중심으로 운영할 필요가 있습니다. 연구과제 선정부터 정책 반영, 후속 평가까지 이어지는 활용경로를 공개하는 방향이 좋습니다.`],
  ];
  const matched = themes.find(([pattern]) => pattern.test(text));
  if (matched) return matched[1];
  const layerDirections = {
    AI인프라: `${subject}은 구축 규모보다 이용률과 서비스 전환 성과가 드러나도록 공급·운영 지표를 함께 관리해야 합니다. 민간 자원과의 역할 분담 및 후속 이용계획도 구체화할 필요가 있습니다.`,
    AI모델: `${subject}은 모델 규모보다 분야별 정확도와 안전성, 경량화, 실제 서비스 적용성을 균형 있게 검증해야 합니다. 개발 결과가 재사용 가능한 모델·데이터 자산으로 축적되도록 관리하는 방향이 좋습니다.`,
    AI생태계: `${subject}은 개별 지원 건수보다 인재·기업·기술이 다음 성장단계로 이동한 비율을 중심으로 성과를 관리해야 합니다. 유사 프로그램을 연결해 중복을 줄이고 후속 민간투자를 유도할 필요가 있습니다.`,
    산업전략: `${subject}은 기획·실증·사업화가 단절되지 않도록 수요산업과 구매 주체를 초기 단계부터 참여시키는 것이 중요합니다. 사업별 기술성숙도와 시장성을 기준으로 계속·조정 여부를 판단해야 합니다.`,
    기관운영비: `${subject}은 운영자원의 투입 근거와 고유기능 성과를 분리해 설명하고, 주요 사업의 품질 향상에 기여한 내용을 구체적으로 제시할 필요가 있습니다.`,
  };
  return layerDirections[project.layer] || layerDirections.산업전략;
}

function summarizePoint(value) {
  const text = cleanNarrativeLine(value)
    .replace(/^\([^)]{2,80}\)\s*/, '')
    .replace(/\s*[-–]\s*\((?:확정|요구|산출)\)\s*/g, ' ')
    .trim();
  if (text.length <= 260) return text;
  const sentence = text.slice(0, 260).match(/^(.{80,240}?[.!?]|.{80,240}?다\.)/);
  return `${(sentence?.[1] || text.slice(0, 240)).trim()}…`;
}

function bulletize(value) {
  return summarizePoint(value)
    .replace(/^[①②③④⑤⑥⑦⑧⑨⑩]\s*/, '')
    .replace(/^\([^)]{2,80}\)\s*/, '')
    .replace(/\s*[:：]\s*\((?:'|’)?\d{2}[^)]*\)[\s\S]*$/, '')
    .replace(/(?:합니다|됩니다|있습니다|필요합니다|추진합니다)[.]?$/, '')
    .replace(/[.。]\s*$/, '')
    .trim();
}

function cleanBusinessText(value) {
  return cleanNarrativeLine(value)
    .replace(/^\([^)]{2,100}\)\s*/, '')
    .replace(/^사업목적\s*[:：]?\s*/i, '')
    .replace(/^사업내용\s*[:：]?\s*/i, '')
    .replace(/\s*:\s*\((?:'|’|`)?\d{2}\)[\s\S]*$/, '')
    .replace(/\s*최근\s*5년\s*간[\s\S]*$/i, '')
    .replace(/\s*총사업비\s*[:：][\s\S]*$/i, '')
    .replace(/\s+/g, ' ')
    .replace(/^[)\]·\-–—\s]+/, '')
    .replace(/[.。]\s*$/, '')
    .trim();
}

function usableBusinessText(value) {
  const text = cleanBusinessText(value);
  return text.length >= 18 && text.length <= 320
    && !/최근 5년|사업비\s*\(|예산액기준|전년대비|본예산|예산안|계획안|요구액|산출내역|해당없음|사업규모\s*연도|사업 수혜자\s*:|법적근거|제\d+조/.test(text)
    && !/^기타\s*[:：]/.test(text)
    && !/^\*+|^(?:202[0-9]){2,}|사업구조개편|내내역.*이관/.test(text)
    && !/백만원|억원|시장선호|기술발전\s*속도|변동\s*가능|조정\s*운용|구매\s*시점의?\s*단가|비용\s*반영|[×=]/.test(text)
    && !/^[\d,\s.%억원백만원→()'’`\-–—]+$/.test(text);
}

function isBareDetailLabel(value, focusTitle) {
  const compact = cleanBusinessText(value).replace(/[^\p{L}\p{N}]/gu, '');
  const focus = String(focusTitle || '').replace(/[^\p{L}\p{N}]/gu, '');
  const withoutMarker = compact.replace(/^[①②③④⑤⑥⑦⑧⑨⑩\d]+/, '');
  return compact === focus || withoutMarker === focus;
}

function detailNarrativeOverride(project) {
  const title = project.detailTitle || '';
  if (/지역 디지털 기업 경쟁력 강화/.test(title)) {
    return '지역 내 초기·강소 디지털기업의 서비스 개발과 사업화를 지원하고, 기술 고도화·투자유치·시장진출 등 기업 수요에 맞춘 성장 지원을 추진';
  }
  if (/소프트웨어 프로세스 인증 확산/.test(title)) {
    return '중소 SW기업의 역량 강화와 제도 개선, SP인증 심사업무 및 SW 품질체계 지원을 통해 인증 활성화와 기업의 품질 경쟁력 강화를 추진';
  }
  if (/AI 미디어·문화 향유 확산.*AI 법률 보조 서비스 확산.*초거대 AI 기반 학술활동 지원/.test(title)) {
    return '초거대 AI를 미디어·문화 향유, 법률업무 보조, 학술활동에 적용하는 서비스의 개발과 실증을 지원하여 전문분야의 업무 효율과 국민의 AI 활용을 확대';
  }
  if (/초거대 AI 기반 클라우드 서비스 개발 역량 지원/.test(title)) {
    return '초거대 AI 모델을 활용한 클라우드 서비스의 개발·실증과 사업화 역량을 지원하여 경쟁력 있는 AI 응용서비스의 시장 진출을 촉진';
  }
  return '';
}

function selectBusinessOverview(candidates, focusTitle = '') {
  const focusTokens = String(focusTitle).replace(/[^\p{L}\p{N}]/gu, ' ').split(/\s+/).filter((item) => item.length >= 2);
  return candidates
    .filter(usableBusinessText)
    .map((value) => cleanBusinessText(value))
    .map((value, index) => ({
      value,
      score: (focusTokens.some((token) => value.includes(token)) ? 8 : 0)
        + (/지원|구축|개발|실증|육성|조성|확산|강화|운영|개선|양성/.test(value) ? 5 : 0)
        + Math.min(value.length, 180) / 60
        - index * 0.05,
    }))
    .sort((a, b) => b.score - a.score)[0]?.value || '';
}

function conciseMainPoints(candidates) {
  const descriptive = candidates
    .filter(usableBusinessText)
    .map((item) => bulletize(cleanBusinessText(item)))
    .filter((item) => item.length >= 18 && item.length <= 260);
  const unique = [];
  for (const item of descriptive) {
    const key = item.replace(/[^\p{L}\p{N}]/gu, '');
    if (unique.some((existing) => {
      const existingKey = existing.replace(/[^\p{L}\p{N}]/gu, '');
      return existingKey.includes(key) || key.includes(existingKey);
    })) continue;
    unique.push(item);
    if (unique.length === 3) break;
  }
  if (!unique.length && candidates.length) unique.push(bulletize(candidates[0]));
  return unique;
}

function extractNipaDetailEntries(lines) {
  const tableStart = lines.findIndex((line) => line.trim() === '내역사업명');
  if (tableStart < 0) return [];
  const endOffset = lines.slice(tableStart + 1).findIndex((line) => /^3\)\s*2026|^3\)\s*당해|^3\)\s*계획/.test(line.trim()));
  const end = endOffset >= 0 ? tableStart + 1 + endOffset : Math.min(lines.length, tableStart + 180);
  const entries = [];
  const headerPattern = /^(?:내역사업명|구분|피보조|지원|보조율|\(2026|사업명)/;
  const boundaryPattern = /(?:백만원|%|법률|진흥법|기본법|특별법|시행령|설립|제\d+조)|^\d{1,3}(?:,\d{3})*$|^(?:정보통신|산업진흥원|정보통신산업진흥원|출연|보조|직접|지자체보조)$/;

  for (let i = tableStart + 1; i < end; i += 1) {
    if (!/^(?:출연|보조|지자체보조)$/.test(lines[i].trim())) continue;
    const institutionParts = [];
    let amount = null;
    for (let j = i + 1; j < Math.min(end, i + 8); j += 1) {
      const value = lines[j].trim();
      const amountMatch = value.match(/^(\d{1,3}(?:,\d{3})*|\d+)\s*백만원?$/);
      if (amountMatch) {
        amount = Number(amountMatch[1].replace(/,/g, ''));
        break;
      }
      if (/^\d{1,3}(?:,\d{3})*$/.test(value)) {
        amount = Number(value.replace(/,/g, ''));
        break;
      }
      institutionParts.push(value);
    }
    const institution = institutionParts.join('').replace(/\s+/g, '');
    if (!institution.includes(operatorName.replace(/\s+/g, '')) || !Number.isFinite(amount)) continue;

    const titleParts = [];
    for (let k = i - 1; k >= Math.max(tableStart + 1, i - 6); k -= 1) {
      const value = lines[k].trim();
      if (!value || headerPattern.test(value) || boundaryPattern.test(value)) break;
      titleParts.unshift(value);
    }
    const title = titleParts.join(' ').replace(/\s+/g, ' ').trim();
    if (!title) continue;
    entries.push({ title, budget: amount });
  }
  return Array.from(new Map(entries.map((entry) => [`${entry.title}|${entry.budget}`, entry])).values());
}

const auditPath = path.join(root, 'all-nipa-budget-block-audit.json');
if (!fs.existsSync(auditPath)) throw new Error('먼저 node audit-all-nipa-budget-blocks.mjs를 실행해 원문 전수감사 자료를 생성하세요.');
const auditedRaw = JSON.parse(fs.readFileSync(auditPath, 'utf8')).rows;
const audited = auditedRaw.filter((row) => {
  const fullPath = path.join(root, 'work', 'hwpx_text', row.file);
  const allLines = fs.readFileSync(fullPath, 'utf8').split(/\r?\n/);
  const compactHeader = allLines.slice(Math.max(0, row.line - 6), row.line - 1)
    .reverse()
    .find((line) => line.includes('사 업 명') && line.includes('□ 사업 코드 정보')) || '';
  const operatorSegment = compactHeader.split(/사업시행주체/).slice(1).join('사업시행주체');
  const directOperator = /정보통신\s*산업진흥원/.test(operatorSegment);
  return row.title.includes(operatorName) || directOperator;
});
const mappingAuditPath = path.join(root, 'nipa-execution-mapping-audit.json');
const mappingAuditRows = fs.existsSync(mappingAuditPath)
  ? JSON.parse(fs.readFileSync(mappingAuditPath, 'utf8')).audited || []
  : [];
const mappingByParentKey = new Map(mappingAuditRows.map((row) => [
  `${row.code}|${row.title.replace(/\s+/g, '')}`,
  row,
]));
const auditedParentKeys = new Set(audited.map((row) => `${row.code}|${row.title.replace(/\s+/g, '')}`));
aiProjects = aiProjects.filter((project) => auditedParentKeys.has(
  `${project.code}|${String(project.parentTitle || project.title).replace(/\s+/g, '')}`,
));
const existingParentKeys = new Set(aiProjects.map((project) => `${project.code}|${String(project.parentTitle || project.title).replace(/\s+/g, '')}`));
const nextNo = Math.max(...aiProjects.map((project) => Number(project.no) || 0), 0) + 1;
const additions = [];

for (const [index, row] of audited.entries()) {
  const key = `${row.code}|${row.title.replace(/\s+/g, '')}`;
  if (existingParentKeys.has(key)) continue;
  const fullPath = path.join(root, 'work', 'hwpx_text', row.file);
  const allLines = fs.readFileSync(fullPath, 'utf8').split(/\r?\n/);
  const block = allLines.slice(row.line - 1, row.line - 1 + row.blockLines);
  const layer = classifyNipaLayer(row.title);
  const budget = extractBlockBudget(block);
  const overview = sourceOverview(block);
  additions.push({
    no: nextNo + index,
    sourceNo: `NIPA-${String(index + 1).padStart(3, '0')}`,
    detailNo: 1,
    pdfPage: '',
    pdfListed: 'N',
    aiBudget: 'N',
    budgetClass: layer.id === '기관운영비' ? '기관운영비' : 'NIPA 전체예산',
    aiBudgetReason: 'AI 여부와 무관하게 원문 사업시행주체에 정보통신산업진흥원이 명시되어 포함',
    title: row.title,
    parentTitle: row.title,
    detailTitle: '',
    isDetail: false,
    code: row.code,
    account: accountFromFile(row.file),
    layer: layer.id,
    layerName: layer.name,
    layerColor: layer.color,
    orgName: '과학기술정보통신부',
    mappedOrgName: '과학기술정보통신부 / 원문 소관부서',
    orgUnit: '과학기술정보통신부 / 원문 소관부서',
    orgPathKey: '과학기술정보통신부 / 원문 소관부서',
    operator: operatorName,
    budget,
    budgetSource: budget == null ? '원문 금액 추가 확인 필요' : '세부사업 2026 예산 총액',
    sourceBook: row.file,
    sourceLine: row.line,
    sourceOverview: overview,
    sourcePurposeItems: [overview],
    sourceMainItems: [],
    overview,
    mainPoints: `${overview}${budget == null ? '' : ` / 2026년 예산: ${budget.toLocaleString('ko-KR')}백만원`}`,
    direction: '정보통신산업진흥원 시행사업 전체 범위에 포함된 사업으로, AI 예산 여부와 관계없이 원문 귀속 기준으로 관리합니다.',
  });
}

const preliminaryProjects = [...aiProjects, ...additions];
const refinedProjects = [...preliminaryProjects];

for (const row of audited) {
  const fullPath = path.join(root, 'work', 'hwpx_text', row.file);
  const allLines = fs.readFileSync(fullPath, 'utf8').split(/\r?\n/);
  const block = allLines.slice(row.line - 1, row.line - 1 + row.blockLines);
  const details = extractNipaDetailEntries(block);
  if (!details.length) continue;
  const parentKey = `${row.code}|${row.title.replace(/\s+/g, '')}`;
  const parentIndexes = refinedProjects
    .map((project, index) => ({ project, index }))
    .filter(({ project }) => `${project.code}|${String(project.parentTitle || project.title).replace(/\s+/g, '')}` === parentKey);
  const replaceable = parentIndexes.filter(({ project }) => !project.isDetail);
  if (!replaceable.length) continue;
  const base = replaceable[0].project;
  const mapping = mappingByParentKey.get(parentKey);
  const primaryOperatorCell = String(mapping?.operator || '').split('소관부처')[0].replace(/\s+/g, '');
  const systemStart = block.findIndex((line) => /②\s*사업추진체계/.test(line));
  const systemOperatorLine = systemStart >= 0
    ? block.slice(systemStart, systemStart + 35).find((line) => /사업시행주체\s*[:：]/.test(line)) || ''
    : '';
  const systemOperators = systemOperatorLine.split(/사업시행주체\s*[:：]\s*/).slice(1).join('').trim();
  const normalizedSystemOperators = systemOperators.replace(/\s+/g, '');
  const bodyShowsMixedOperators = normalizedSystemOperators.includes(operatorName.replace(/\s+/g, ''))
    && normalizedSystemOperators.replace(operatorName.replace(/\s+/g, ''), '').replace(/[,，·ㆍ/]/g, '').length > 0;
  const detailOnly = !row.title.includes(operatorName)
    && (bodyShowsMixedOperators || (!systemOperators && primaryOperatorCell !== operatorName.replace(/\s+/g, '')))
    && details.length > 0;
  if (detailOnly) {
    for (const { index } of parentIndexes.sort((a, b) => b.index - a.index)) refinedProjects.splice(index, 1);
    details.forEach((detail, detailIndex) => {
      const cleanDetailTitle = detail.title.replace(/^[․ㆍ·○ㅇ\-\s]+/, '').trim();
      const layer = classifyNipaLayer(`${row.title} ${cleanDetailTitle}`);
      refinedProjects.push({
        ...base,
        no: `${base.no}-${detailIndex + 1}`,
        detailNo: detailIndex + 1,
        title: `(내역) ${cleanDetailTitle}`,
        parentTitle: row.title,
        detailTitle: cleanDetailTitle,
        isDetail: true,
        displayMode: 'nipa-detail-only',
        layer: layer.id,
        layerName: layer.name,
        layerColor: layer.color,
        budget: detail.budget,
        budgetSource: '공동 시행 모사업 중 NIPA 담당 내역 지원금액',
        overview: `${row.title} 가운데 정보통신산업진흥원이 시행하는 '${cleanDetailTitle}' 내역입니다.`,
        mainPoints: `${cleanDetailTitle} 수행에 배정된 NIPA 예산 ${detail.budget.toLocaleString('ko-KR')}백만원`,
        detailBreakdown: [{ ...detail, title: cleanDetailTitle }],
      });
    });
    continue;
  }
  const detailBudget = details.reduce((sum, detail) => sum + detail.budget, 0);
  const wholeBudget = extractBlockBudget(block) || Number(base.budget) || detailBudget;
  const detailSummary = details
    .map((detail) => `${detail.title} ${detail.budget.toLocaleString('ko-KR')}백만원`)
    .join(' · ');
  const overview = base.overview || base.sourceOverview || sourceOverview(block);
  const replacement = {
    ...base,
    title: row.title,
    parentTitle: row.title,
    detailTitle: '',
    isDetail: false,
    detailNo: 1,
    detailBreakdown: details,
    parentBudget: wholeBudget,
    budget: wholeBudget,
    budgetSource: 'NIPA 단독 시행 모사업의 2026년 예산 총액',
    overview,
    mainPoints: [base.mainPoints, detailSummary].filter(Boolean).join(' / 주요 내역: '),
  };
  const keepIndex = Math.min(...parentIndexes.map(({ index }) => index));
  for (const { index } of parentIndexes.sort((a, b) => b.index - a.index)) refinedProjects.splice(index, 1);
  refinedProjects.splice(keepIndex, 0, replacement);
}

const auditedByParentKey = new Map(audited.map((row) => [
  `${row.code}|${row.title.replace(/\s+/g, '')}`,
  row,
]));
const groupedProjects = new Map();
for (const project of refinedProjects) {
  const key = project.displayMode === 'nipa-detail-only'
    ? `${project.code}|${String(project.parentTitle).replace(/\s+/g, '')}|${String(project.detailTitle).replace(/\s+/g, '')}`
    : `${project.code}|${String(project.parentTitle || project.title).replace(/\s+/g, '')}`;
  if (!groupedProjects.has(key)) groupedProjects.set(key, []);
  groupedProjects.get(key).push(project);
}
const collapsedProjects = [...groupedProjects.values()].map((rows) => {
  if (rows.length === 1) return rows[0];
  const base = rows.find((row) => !row.isDetail) || rows[0];
  const breakdown = rows.flatMap((row) => row.detailBreakdown?.length
    ? row.detailBreakdown
    : [{ title: row.detailTitle || row.title.replace(/^\(내역\)\s*/, ''), budget: Number(row.budget) || 0 }]);
  const overview = base.overview || base.sourceOverview || rows.find((row) => row.overview)?.overview || '';
  return {
    ...base,
    title: base.parentTitle || base.title,
    detailTitle: '',
    isDetail: false,
    detailNo: 1,
    detailBreakdown: breakdown,
    budget: rows.reduce((sum, row) => sum + (Number(row.budget) || 0), 0),
    overview,
    mainPoints: [overview, breakdown.map((item) => `${item.title} ${item.budget.toLocaleString('ko-KR')}백만원`).join(' · ')]
      .filter(Boolean).join(' / 주요 내역: '),
  };
});
const projects = collapsedProjects.map((project) => {
  const parentKey = `${project.code}|${String(project.parentTitle || project.title).replace(/\s+/g, '')}`;
  const auditedRow = auditedByParentKey.get(parentKey);
  const normalized = {
    ...project,
    title: project.displayMode === 'nipa-detail-only' ? project.title : (project.parentTitle || project.title),
    detailTitle: project.displayMode === 'nipa-detail-only' ? project.detailTitle : '',
    isDetail: project.displayMode === 'nipa-detail-only',
    ...(auditedRow ? {
      account: accountFromFile(auditedRow.file),
      sourceBook: auditedRow.file,
      sourceLine: auditedRow.line,
    } : {}),
  };
  let narrative = { overview: '', purpose: [], focused: [], main: [] };
  let sourceOrganization = null;
  let sourceOrganizations = [];
  let orgDetailAssignments = [];
  if (auditedRow) {
    const sourceLines = fs.readFileSync(path.join(root, 'work', 'hwpx_text', auditedRow.file), 'utf8').split(/\r?\n/);
    const sourceBlock = sourceLines.slice(auditedRow.line - 1, auditedRow.line - 1 + auditedRow.blockLines);
    narrative = extractNarrative(sourceBlock, normalized.detailTitle || normalized.title);
    sourceOrganization = resolveSourceOrganization(sourceLines, auditedRow.line, normalized.detailTitle || normalized.title);
    const organizationFocuses = normalized.displayMode === 'nipa-detail-only'
      ? [normalized.detailTitle || normalized.title]
      : (normalized.detailBreakdown?.length
        ? normalized.detailBreakdown.map((detail) => detail.title)
        : [normalized.title]);
    sourceOrganizations = Array.from(new Map(organizationFocuses
      .map((focus) => resolveSourceOrganization(sourceLines, auditedRow.line, focus))
      .filter(Boolean)
      .map((organization) => [organization.mappedOrgName, organization])).values());
    orgDetailAssignments = (normalized.detailBreakdown?.length
      ? normalized.detailBreakdown
      : [{ title: normalized.detailTitle || normalized.title, budget: Number(normalized.budget) || 0 }])
      .map((detail) => {
        const organization = resolveSourceOrganization(sourceLines, auditedRow.line, detail.title);
        return organization ? {
          title: detail.title,
          budget: Number(detail.budget) || 0,
          department: organization.mappedOrgName,
          orgPathKey: organization.orgPathKey,
        } : null;
      })
      .filter(Boolean);
  }
  if (sourceOrganization) {
    normalized.mappedOrgName = sourceOrganization.mappedOrgName;
    normalized.orgPathKey = sourceOrganization.orgPathKey;
    normalized.orgName = sourceOrganization.mappedOrgName;
    normalized.orgUnit = sourceOrganization.mappedOrgName;
    normalized.orgDivision = sourceOrganization.mappedOrgName.split(' / ')[0];
    normalized.orgDepartments = (sourceOrganizations.length ? sourceOrganizations : [sourceOrganization])
      .map((organization) => organization.mappedOrgName);
  }
  normalized.orgDetailAssignments = orgDetailAssignments;
  if (!normalized.orgPathKey && normalized.mappedOrgName) {
    normalized.orgPathKey = originalOrgs
      .map((org) => org.id)
      .filter((id) => normalized.mappedOrgName === id || normalized.mappedOrgName.startsWith(`${id} /`))
      .sort((a, b) => b.length - a.length)[0]
      || normalized.mappedOrgName.split(' / ').slice(0, 3).join(' / ');
  }
  if (!normalized.orgDepartments?.length && normalized.mappedOrgName) normalized.orgDepartments = [normalized.mappedOrgName];
  const detailFocusedCandidates = normalized.displayMode === 'nipa-detail-only'
    ? narrative.focused.filter((value) => value && !isBareDetailLabel(value, normalized.detailTitle))
    : [];
  const detailFallback = normalized.displayMode === 'nipa-detail-only'
    ? `${normalized.parentTitle || normalized.title}의 내역사업인 '${normalized.detailTitle}'을 수행`
    : '';
  const narrativeCandidates = normalized.displayMode === 'nipa-detail-only'
    ? (detailFocusedCandidates.length ? detailFocusedCandidates : [detailFallback])
    : [narrative.overview, ...narrative.purpose, ...narrative.main, normalized.overview, normalized.sourceOverview];
  normalized.overview = selectBusinessOverview(
    narrativeCandidates.filter(Boolean),
    normalized.detailTitle || normalized.title,
  ) || cleanBusinessText(normalized.overview || normalized.sourceOverview)
    || '정보통신산업진흥원이 수행하는 2026년 예산 사업';
  normalized.sourcePurposeItems = (normalized.displayMode === 'nipa-detail-only' ? detailFocusedCandidates : narrative.purpose)
    .filter(usableBusinessText)
    .map(cleanBusinessText);
  if (!normalized.sourcePurposeItems.length) normalized.sourcePurposeItems = [normalized.overview];
  const contentCandidates = normalized.displayMode === 'nipa-detail-only'
    ? [normalized.overview, ...detailFocusedCandidates].filter(Boolean)
    : [normalized.overview, ...narrativeCandidates].filter(Boolean);
  normalized.sourceMainItems = conciseMainPoints(contentCandidates);
  if (!normalized.sourceMainItems.length) normalized.sourceMainItems = [normalized.overview];
  const overriddenDetailNarrative = normalized.displayMode === 'nipa-detail-only'
    ? detailNarrativeOverride(normalized)
    : '';
  if (overriddenDetailNarrative) {
    normalized.overview = overriddenDetailNarrative;
    normalized.sourcePurposeItems = [overriddenDetailNarrative];
    normalized.sourceMainItems = [overriddenDetailNarrative];
  }
  normalized.mainPoints = normalized.sourceMainItems.join(' / ');
  const primaryLayer = classifyNipaLayer([normalized.title, normalized.detailTitle].join(' '));
  const contentLayer = classifyNipaLayer([
    normalized.title,
    normalized.detailTitle,
    normalized.overview,
    ...(normalized.sourcePurposeItems || []),
    ...(normalized.sourceMainItems || []),
  ].join(' '));
  let semanticLayer = ['기관운영비', '글로벌 이니셔티브'].includes(primaryLayer.id)
    || ['기관운영비', '글로벌 이니셔티브'].includes(contentLayer.id)
    ? primaryLayer
    : contentLayer;
  const purposeText = [normalized.title, normalized.detailTitle, normalized.overview].join(' ');
  const regionalGoal = /지역\s*(?:교육|산업|기업|인재|경제|사회|생태계|혁신|경쟁력|격차|문제|현안)|지역정주|지역주도|지역특화|균형발전|권역별|권역\s*(?:산업|혁신|경쟁력)|(?:서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|전북|전남|경북|경남|제주).{0,30}(?:AX|산업|혁신|허브|거점|생태계)/i.test(purposeText);
  const explicitPhysicalGoal = /피지컬\s*AI|협업지능|로봇|자율주행|무인|드론|스마트팩토리|메타팩토리|디지털트윈|공간컴퓨팅/i.test(purposeText);
  if (!['기관운영비', '글로벌 이니셔티브'].includes(semanticLayer.id)) {
    if (regionalGoal) semanticLayer = { id: '지역AX', name: '지역AX', color: '#f4a261' };
    else if (explicitPhysicalGoal) semanticLayer = { id: '피지컬AI', name: '피지컬AI', color: '#ef476f' };
  }
  if (/인간\s*[-–]?\s*AI\s*협업형\s*LAM\s*개발.*글로벌\s*실증/i.test(normalized.title)) {
    semanticLayer = { id: '피지컬AI', name: '피지컬AI', color: '#ef476f' };
  }
  semanticLayer = policyPurposeLayer(normalized, semanticLayer);
  normalized.layer = semanticLayer.id;
  normalized.layerName = semanticLayer.name;
  normalized.layerColor = semanticLayer.color;
  const mainItems = normalized.sourceMainItems;
  const direction = analyzeDirection(normalized);
  normalized.direction = direction;
  const detailTableRows = normalized.detailBreakdown?.length
    ? normalized.detailBreakdown
    : [{ title: normalized.detailTitle || normalized.title.replace(/^\(내역\)\s*/, ''), budget: Number(normalized.budget) || 0 }];
  normalized.detail = {
    ...(normalized.detail || {}),
    type: normalized.detail?.type || normalized.layerName,
    displayTitle: normalized.detailTitle || normalized.title.replace(/^\(내역\)\s*/, ''),
    parentTitle: normalized.parentTitle || normalized.title,
    overview: normalized.overview,
    mainPoints: mainItems,
    direction,
    details: detailTableRows,
    source: `${normalized.sourceBook} / 원문 기준선 ${normalized.sourceLine} / 사업목적·주요내용 근거: ${normalized.overview}`,
  };
  if (!/동반친출/.test(`${normalized.title} ${normalized.detailTitle}`)) return normalized;
  return {
    ...normalized,
    title: normalized.title.replace(/동반친출/g, '동반진출'),
    detailTitle: normalized.detailTitle.replace(/동반친출/g, '동반진출'),
    mainPoints: String(normalized.mainPoints || '').replace(/동반친출/g, '동반진출'),
  };
}).sort((a, b) => (Number(b.budget) || 0) - (Number(a.budget) || 0));

const displayedDetailsByParent = new Map();
for (const project of projects.filter((item) => item.displayMode === 'nipa-detail-only')) {
  const key = `${project.code}|${String(project.parentTitle || '').replace(/\s+/g, '')}`;
  if (!displayedDetailsByParent.has(key)) displayedDetailsByParent.set(key, []);
  displayedDetailsByParent.get(key).push({
    title: project.detailTitle || project.title.replace(/^\(내역\)\s*/, ''),
    budget: Number(project.budget) || 0,
  });
}
for (const project of projects.filter((item) => item.displayMode === 'nipa-detail-only')) {
  const key = `${project.code}|${String(project.parentTitle || '').replace(/\s+/g, '')}`;
  project.detail.details = (displayedDetailsByParent.get(key) || [])
    .sort((a, b) => b.budget - a.budget);
  project.orgDepartments = [project.mappedOrgName];
}

function narrativeForOrgDetail(project, detailTitle) {
  const overridden = detailNarrativeOverride({ ...project, detailTitle });
  if (overridden) return { overview: overridden, mainPoints: [overridden] };
  const parentKey = `${project.code}|${String(project.parentTitle || project.title).replace(/\s+/g, '')}`;
  const auditedRow = auditedByParentKey.get(parentKey);
  if (!auditedRow) {
    const fallback = `${project.parentTitle || project.title}의 내역사업인 '${detailTitle}'을 수행`;
    return { overview: fallback, mainPoints: [fallback] };
  }
  const sourceLines = fs.readFileSync(path.join(root, 'work', 'hwpx_text', auditedRow.file), 'utf8').split(/\r?\n/);
  const sourceBlock = sourceLines.slice(auditedRow.line - 1, auditedRow.line - 1 + auditedRow.blockLines);
  const narrative = extractNarrative(sourceBlock, detailTitle);
  const focused = narrative.focused
    .filter((value) => usableBusinessText(value) && !isBareDetailLabel(value, detailTitle));
  const fallback = `${project.parentTitle || project.title}의 내역사업인 '${detailTitle}'을 수행`;
  const overview = selectBusinessOverview(focused, detailTitle) || fallback;
  const mainPoints = conciseMainPoints([overview, ...focused]);
  return { overview, mainPoints: mainPoints.length ? mainPoints : [overview] };
}

const orgProjects = projects.flatMap((project) => {
  const assignments = project.orgDetailAssignments || [];
  const departments = [...new Set(assignments.map((item) => item.department).filter(Boolean))];
  if (project.displayMode === 'nipa-detail-only' || departments.length <= 1) return [project];
  return assignments.map((assignment, index) => {
    const focusedNarrative = narrativeForOrgDetail(project, assignment.title);
    const splitProject = {
      ...project,
      no: `${project.no}-org-${index + 1}`,
      title: `(내역) ${assignment.title}`,
      parentTitle: project.parentTitle || project.title,
      detailTitle: assignment.title,
      isDetail: true,
      budget: assignment.budget,
      overview: focusedNarrative.overview,
      sourcePurposeItems: [focusedNarrative.overview],
      sourceMainItems: focusedNarrative.mainPoints,
      mainPoints: focusedNarrative.mainPoints.join(' / '),
      mappedOrgName: assignment.department,
      orgName: assignment.department,
      orgUnit: assignment.department,
      orgDivision: assignment.department.split(' / ')[0],
      orgPathKey: assignment.orgPathKey || assignment.department,
      orgDepartments: [assignment.department],
    };
    const direction = analyzeDirection(splitProject);
    splitProject.direction = direction;
    splitProject.detail = {
      ...project.detail,
      displayTitle: assignment.title,
      parentTitle: project.parentTitle || project.title,
      overview: focusedNarrative.overview,
      mainPoints: focusedNarrative.mainPoints,
      direction,
      source: `${project.sourceBook} / 원문 기준선 ${project.sourceLine} / 사업목적·주요내용 근거: ${focusedNarrative.overview}`,
    };
    return splitProject;
  });
}).sort((a, b) => (Number(b.budget) || 0) - (Number(a.budget) || 0));

const headquartersReviewPath = path.join(root, 'NIPA_사업별_담당본부_재검증DB.xlsx');
const headquartersReview = new Map();
if (fs.existsSync(headquartersReviewPath)) {
  const reviewWorkbook = new ExcelJS.Workbook();
  await reviewWorkbook.xlsx.readFile(headquartersReviewPath);
  const reviewSheet = reviewWorkbook.getWorksheet('사업별 본부 재검증');
  if (reviewSheet) {
    for (let rowNo = 2; rowNo <= reviewSheet.rowCount; rowNo += 1) {
      const row = reviewSheet.getRow(rowNo);
      const title = row.getCell(3).text.trim();
      if (!title) continue;
      headquartersReview.set(title, {
        headquarters: row.getCell(8).text.trim(),
        status: row.getCell(9).text.trim(),
        breakdown: row.getCell(12).text.trim(),
        sourceRows: row.getCell(13).text.trim(),
      });
    }
  }
}

function normalizeNipaHeadquarters(value) {
  return String(value || '')
    .replace(/AI반도체지원본부/g, 'AI반도체본부')
    .replace(/경영지원 관련 본부 확인 필요|SW정책연구소|확인 필요/g, '경영기획본부')
    .trim();
}

const nipaHeadquartersProjects = projects.flatMap((project, projectIndex) => {
  const review = headquartersReview.get(project.title) || {};
  if (project.layer === '기관운영비') {
    return [{ ...project, no: `${project.no}-nipa-hq-1`, nipaHeadquarters: '경영기획본부', nipaMappingStatus: '사용자 지정' }];
  }
  const lines = String(review.breakdown || '').split(/\r?\n/).filter(Boolean);
  const detailAllocations = [];
  const grouped = new Map();
  for (const line of lines) {
    const match = line.match(/^(.*?)\s*→\s*(.*?)\s*\(([\d,]+)백만원\)\s*$/);
    if (!match) continue;
    const detailTitle = match[1].replace(/^\(내\)\s*/, '').trim();
    const headquarters = normalizeNipaHeadquarters(match[2]);
    const budget = Number(match[3].replace(/,/g, ''));
    detailAllocations.push({ detailTitle, headquarters, budget });
    grouped.set(headquarters, (grouped.get(headquarters) || 0) + budget);
  }
  if (!grouped.size) {
    const headquarters = normalizeNipaHeadquarters(review.headquarters || '확인 필요');
    return [{ ...project, no: `${project.no}-nipa-hq-1`, nipaHeadquarters: headquarters, nipaMappingStatus: review.status || '확인 필요' }];
  }
  const allocations = [...grouped.entries()].map(([headquarters, budget]) => ({ headquarters, budget }));
  const matchedBudget = allocations.reduce((sum, item) => sum + item.budget, 0);
  const residual = (Number(project.budget) || 0) - matchedBudget;
  if (!project.isDetail && grouped.size > 1 && detailAllocations.length) {
    if (residual !== 0) {
      const largest = [...allocations].sort((a, b) => b.budget - a.budget)[0];
      detailAllocations.push({ detailTitle: '기타·미연결 내역', headquarters: largest.headquarters, budget: residual });
    }
    return detailAllocations.map((allocation, index) => ({
      ...project,
      no: `${project.no}-nipa-hq-detail-${index + 1}`,
      title: `(내역) ${allocation.detailTitle}`,
      parentTitle: project.parentTitle || project.title,
      detailTitle: allocation.detailTitle,
      isDetail: true,
      budget: allocation.budget,
      nipaHeadquarters: allocation.headquarters,
      nipaMappingStatus: review.status || '내부자료 연계',
      nipaSourceRows: review.sourceRows || '',
      detail: {
        ...project.detail,
        displayTitle: allocation.detailTitle,
        parentTitle: project.parentTitle || project.title,
      },
    }));
  }
  if (allocations.length > 1 && residual !== 0) {
    allocations.sort((a, b) => b.budget - a.budget)[0].budget += residual;
  }
  return allocations.map((allocation, index) => ({
    ...project,
    no: `${project.no}-nipa-hq-${index + 1}`,
    budget: allocations.length === 1 ? Number(project.budget) || 0 : allocation.budget,
    nipaHeadquarters: allocation.headquarters,
    nipaMappingStatus: review.status || '내부자료 연계',
    nipaSourceRows: review.sourceRows || '',
  }));
}).sort((a, b) => (Number(b.budget) || 0) - (Number(a.budget) || 0));

const nipaHeadquartersOrder = ['정책기획단', 'SW융합본부', 'AI인프라본부', 'AI반도체본부', 'AI활용본부', '지역AX본부', '글로벌본부', '경영기획본부'];
const nipaHeadquartersColors = {
  정책기획단: '#4f7dd9',
  SW융합본부: '#3aa7a3',
  AI인프라본부: '#3487d4',
  AI반도체본부: '#5c6fdb',
  AI활용본부: '#8a63d2',
  지역AX본부: '#e28b3f',
  글로벌본부: '#d85f82',
  경영기획본부: '#6f91a8',
};
const nipaHeadquarters = [...new Set(nipaHeadquartersProjects.map((project) => project.nipaHeadquarters))]
  .map((name) => {
    const rows = nipaHeadquartersProjects.filter((project) => project.nipaHeadquarters === name);
    return {
      name,
      color: nipaHeadquartersColors[name] || '#4f8fb3',
      count: rows.length,
      budget: rows.reduce((sum, project) => sum + (Number(project.budget) || 0), 0),
    };
  })
  .sort((a, b) => {
    const left = nipaHeadquartersOrder.indexOf(a.name);
    const right = nipaHeadquartersOrder.indexOf(b.name);
    return (left < 0 ? 999 : left) - (right < 0 ? 999 : right);
  });

const nipaLayerDefinitions = [
  { id: 'AI인프라', name: 'AI인프라', color: '#355c7d', keywords: ['GPU', '클라우드', '데이터', '컴퓨팅'] },
  { id: 'AI반도체', name: 'AI반도체', color: '#0081a7', keywords: ['AI반도체', 'NPU', '온디바이스'] },
  { id: 'AI모델', name: 'AI모델', color: '#43aa8b', keywords: ['파운데이션 모델', '생성형 AI', '에이전트'] },
  { id: '지역AX', name: '지역AX', color: '#f4a261', keywords: ['지역특화', '권역 AX', '지역기업'] },
  { id: '피지컬AI', name: '피지컬AI', color: '#ef476f', keywords: ['로봇', '자율', '디지털트윈'] },
  { id: '글로벌 이니셔티브', name: '글로벌 이니셔티브', color: '#6d597a', keywords: ['해외진출', '국제협력', '수출'] },
  { id: 'AI생태계', name: 'AI생태계', color: '#2a9d8f', keywords: ['인재', '창업', '오픈소스', '보안'] },
  { id: '산업전략', name: '산업전략', color: '#e76f51', keywords: ['산업 AX', '정책', '사업화'] },
  { id: '기관운영비', name: '기관운영비', color: '#6b7280', keywords: ['정보통신산업진흥원 지원', 'SW정책연구소 운영'] },
];
const layers = nipaLayerDefinitions.map((layer) => {
  const rows = projects.filter((project) => project.layer === layer.id);
  return {
    ...layer,
    count: rows.length,
    budget: rows.reduce((sum, project) => sum + (Number(project.budget) || 0), 0),
  };
}).filter((layer) => layer.count > 0);

const orgKey = (project) => project.orgPathKey || project.mappedOrgName || project.orgName || '';
const orgs = originalOrgs.map((org) => {
  const rows = projects.filter((project) => orgKey(project) === org.id || orgKey(project) === org.name);
  return {
    ...org,
    count: rows.length,
    budget: rows.reduce((sum, project) => sum + (Number(project.budget) || 0), 0),
  };
}).filter((org) => org.count > 0);
const knownOrgIds = new Set(orgs.map((org) => org.id));
for (const orgPathKey of [...new Set(projects.map((project) => orgKey(project)).filter(Boolean))]) {
  if (knownOrgIds.has(orgPathKey)) continue;
  const rows = projects.filter((project) => orgKey(project) === orgPathKey);
  orgs.push({
    id: orgPathKey,
    name: orgPathKey,
    division: orgPathKey.split(' / ')[0],
    count: rows.length,
    budget: rows.reduce((sum, project) => sum + (Number(project.budget) || 0), 0),
  });
  knownOrgIds.add(orgPathKey);
}

function ensureOrgTreePath(tree, orgPath) {
  const parts = orgPath.split(' / ').filter(Boolean);
  let level = tree;
  for (let index = 0; index < parts.length; index += 1) {
    const name = parts[index];
    const isLeaf = index === parts.length - 1;
    let node = level.find((item) => (typeof item === 'string' ? item : item.name) === name);
    if (!node) {
      node = isLeaf ? name : { name, children: [] };
      level.push(node);
    }
    if (isLeaf) break;
    if (typeof node === 'string') {
      const replacement = { name: node, children: [] };
      level[level.indexOf(node)] = replacement;
      node = replacement;
    }
    node.children ||= [];
    level = node.children;
  }
}
const orgTree = structuredClone(originalOrgTree);
for (const department of new Set(projects.flatMap((project) => project.orgDepartments || []))) {
  ensureOrgTreePath(orgTree, department);
}

const totalBudget = projects.reduce((sum, project) => sum + (Number(project.budget) || 0), 0);
const institutionalBudget = projects
  .filter((project) => project.layer === '기관운영비')
  .reduce((sum, project) => sum + (Number(project.budget) || 0), 0);
const budgetExcludingInstitutional = totalBudget - institutionalBudget;
const parentCount = new Set(projects.map((project) => project.sourceNo)).size;
const sharedOperatorCount = projects.filter((project) => String(project.operator).includes(',')).length;

replaceEmbeddedJson('PROJECTS', 'LAYERS', projects);
replaceEmbeddedJson('LAYERS', 'ORGS', layers);
replaceEmbeddedJson('ORGS', 'ORG_TREE', orgs);
html = html.replace(/const LAYERS =/, `const ORG_PROJECTS = ${JSON.stringify(orgProjects)};\nconst LAYERS =`);
html = html.replace(/const LAYERS =/, `const NIPA_HQ_PROJECTS = ${JSON.stringify(nipaHeadquartersProjects)};\nconst NIPA_HEADQUARTERS = ${JSON.stringify(nipaHeadquarters)};\nconst LAYERS =`);
html = html.replace(/const ORG_TREE = [\s\S]*?;\r?\n/, `const ORG_TREE = ${JSON.stringify(orgTree)};\n`);
html = html
  .replace('const project = PROJECTS.find((item) => String(item.no) === String(projectNo));', 'const project = [...PROJECTS, ...ORG_PROJECTS, ...NIPA_HQ_PROJECTS].find((item) => String(item.no) === String(projectNo));')
  .replace('  PROJECTS.forEach((project) => {\n    const paths = project.orgDepartments', '  ORG_PROJECTS.forEach((project) => {\n    const paths = project.orgDepartments')
  .replace('    return PROJECTS.filter((project) => {\n      const paths = project.orgDepartments', '    return ORG_PROJECTS.filter((project) => {\n      const paths = project.orgDepartments');

html = html
  .replace(/<title>[\s\S]*?<\/title>/, '<title>정보통신산업진흥원 2026 예산 생태계 MAP 현황판</title>')
  .replace(/<div class="eyebrow">[\s\S]*?<\/div>/, '<div class="eyebrow">NIPA Total Budget x Ecosystem Map</div>')
  .replace(/<h1>[\s\S]*?<\/h1>/, '<h1>정보통신산업진흥원 2026 예산 생태계 MAP 현황판</h1>')
  .replace(/<p class="lead">[\s\S]*?<\/p>/, `<p class="lead">과학기술정보통신부 2026년 예산 및 기금운용계획 사업설명자료 전체를 기준으로 사업시행주체에 정보통신산업진흥원이 명시된 모든 사업을 선별했습니다. 예산 레이어와 소관 조직 체계, 사업·예산·원문 근거를 한 화면에서 확인할 수 있습니다.</p>`)
  .replace(/<div class="meta">[\s\S]*?<\/div>\s*<\/div>\s*<\/header>/, `<div class="meta">
      <span class="pill">NIPA 수행 사업 ${projects.length.toLocaleString('ko-KR')}건</span>
      <span class="pill">세부사업 기준 ${parentCount.toLocaleString('ko-KR')}건</span>
      <span class="pill">총 ${budgetExcludingInstitutional.toLocaleString('ko-KR')} 백만원 · 기관운영비 포함 ${totalBudget.toLocaleString('ko-KR')} 백만원</span>
      <span class="pill">생태계 레이어 ${layers.length.toLocaleString('ko-KR')}개</span>
      <span class="pill">금액 단위: 백만원 · 원문 자동 추출값</span>
    </div>
  </div>
</header>`)
  .replace(/<div class="kpis">[\s\S]*?<\/div>\s*<div class="controls">/, `<div class="kpis">
      <div class="kpi"><span>NIPA 수행 세부사업</span><b id="kpiCount"></b></div>
      <div class="kpi"><span>추출 예산 합계 (기관운영비 포함 ${totalBudget.toLocaleString('ko-KR')}백만원)</span><b id="kpiBudget"></b></div>
    </div>
    <div class="controls">`)
  .replace(/<input id="q" type="search" placeholder="[^"]*">/, '<input id="q" type="search" placeholder="사업명, 사업코드, 소관부서 검색">')
  .replace('수요조사·전문가 의견·정책/생태계 자료 종합 추론', '정책/시장 자료 기반 종합 추론')
  .replace(/const money = \(v\) =>[^;]*;/, "const money = (v) => v ? fmt.format(v) + '\\u00a0백만원' : '미추출';")
  .replace(/document\.title = [^;]*;/g, "document.title = '정보통신산업진흥원 2026 예산 생태계 MAP 현황판';");

const directionCards = [
  ['AI인프라', 'AI인프라를 공동 활용·성과 중심 기반으로 전환', 'GPU·클라우드·데이터 자원은 공급량보다 가동률, 이용 대기시간, 기업별 서비스 전환 성과를 관리하고 국산 인프라와 연계해야 합니다.'],
  ['AI반도체', 'AI반도체는 개발·실증·수요창출을 단일 트랙으로 연결', '국산 NPU와 온디바이스 AI는 실제 워크로드 성능·전력효율 검증에서 공공·민간 구매와 상용화까지 이어지는 패키지 투자가 필요합니다.'],
  ['AI모델', 'AI모델은 도메인 성능과 활용 가능한 기술자산에 집중', '파운데이션·생성형·에이전트 모델은 규모 경쟁보다 데이터 품질, 안전성, 경량화, 산업별 서비스 적용성과를 중심으로 지원해야 합니다.'],
  ['지역AX', '지역AX는 권역별 대표 프로젝트 중심으로 재편', '지역 특화산업 로드맵과 기업지원 사업을 연결하고, 지역기업 매출·고용과 수요기관 도입, 다른 지역으로의 확산을 공통 KPI로 설정해야 합니다.'],
  ['피지컬AI', '피지컬AI는 기술·테스트베드·현장구매를 패키지화', '제조·물류·모빌리티 현장의 실증환경과 안전 검증, 운영비, 실제 구매 주체를 함께 설계해 연구개발 이후 도입 단절을 줄여야 합니다.'],
  ['글로벌 이니셔티브', '글로벌 이니셔티브는 수출 전환형 지원으로 고도화', '해외행사 참가 건수보다 목표시장별 현지 실증, 인증·규제 대응, 수요처 발굴과 계약·수출 전환을 중심으로 예산을 집중해야 합니다.'],
  ['AI생태계', 'AI생태계는 인재·창업·오픈소스·신뢰 기반을 연결', '분절된 기반 사업을 기업 성장단계와 연계하고, 인재의 현장 투입, 기술 채택, 후속투자, 보안·신뢰 확보까지 지속적으로 추적해야 합니다.'],
  ['산업전략', '산업전략은 기획에서 조달·민간투자까지 포트폴리오화', '산업별 AX 수요와 기술성숙도를 기준으로 기획·실증·사업화 사업을 연결하고 유사 사업의 중복을 줄이는 구조조정이 필요합니다.'],
].map(([id, title, body]) => `<article class="direction-card"><div class="direction-tags"><span>${id}</span></div><b>${title}</b><p>${body}</p></article>`).join('');
html = html
  .replace(/<button class="tab" data-panel="ecosystem">[\s\S]*?<\/button>/, '<button class="tab" data-panel="ecosystem">레이어 매핑</button>')
  .replace(/<div class="direction-cards">[\s\S]*<\/article>\s*<\/div>\s*<\/div>(?=\s*<\/div>\s*<\/div>\s*<div class="card dashboard-list-card")/, `<div class="direction-cards">${directionCards}</div></div>`)
  .replace(/<section id="ecosystem" class="panel">\s*<div class="note">[\s\S]*?<\/div>/, '<section id="ecosystem" class="panel">')
  .replace(/생태계 MAP 매핑/g, '레이어 매핑')
  .replace(/생태계 레이어/g, '예산 레이어')
  .replace(/<th>생태계<\/th>/g, '<th>레이어</th>')
  .replace(/<span>생태계<\/span>/g, '<span>레이어</span>');
html = html
  .replace(/layer\.id \+ ' ' \+ layer\.name/g, 'layer.name')
  .replace(/layer\.id\+' '\+layer\.name/g, 'layer.name')
  .replace(/escapeHtml\(project\.layer\)\+' '\+escapeHtml\(project\.layerName\)/g, 'escapeHtml(project.layerName)')
  .replace(/money\(p\.budget\)\+'<div class="muted">'\+p\.budgetSource\+'<\/div>'/g, 'money(p.budget)')
  .replace(/money\(p\.budget\)\+'<div class="muted">'\+escapeHtml\(p\.budgetSource\)\+'<\/div>'/g, 'money(p.budget)')
  .replace("money(p.budget)+'<div class=\"muted\">'+p.budgetSource+'</div>'", 'money(p.budget)')
  .replace("money(p.budget)+'<div class=\"muted\">'+escapeHtml(p.budgetSource)+'</div>'", 'money(p.budget)');

html = html
  .replace(/<button class="tab" data-panel="org">[\s\S]*?<\/button>/, '<button class="tab" data-panel="org">과기정통부 소관 현황</button>')
  .replace(/\s*<button class="tab" data-panel="demand-analysis">[\s\S]*?<\/button>/, '')
  .replace(/(<button class="tab" data-panel="org">과기정통부 소관 현황<\/button>)/, '$1\n    <button class="tab" data-panel="nipa-headquarters">NIPA 본부별 사업현황</button>')
  .replace(/(<section id="org" class="panel">)/, `<section id="nipa-headquarters" class="panel">
    <div class="map-grid" id="nipaHeadquartersGrid"></div>
  </section>
  $1`)
  .replace(/\s*<section id="demand-analysis" class="panel">[\s\S]*?(?=\s*<\/main>)/, '');

const nipaHeadquartersRenderer = `
function renderNipaHeadquarters() {
  const grid = $('nipaHeadquartersGrid');
  if (!grid) return;
  grid.innerHTML = NIPA_HEADQUARTERS.map((headquarters) => {
    const rows = NIPA_HQ_PROJECTS.filter((project) => project.nipaHeadquarters === headquarters.name)
      .sort((a, b) => (b.budget || 0) - (a.budget || 0));
    const items = rows.slice(0, 10);
    return '<article class="layer-card" role="button" tabindex="0" data-nipa-headquarters="'+escapeAttr(headquarters.name)+'">'
      + '<div class="layer-head" style="background:'+escapeAttr(headquarters.color)+'"><span>'+escapeHtml(headquarters.name)+'</span><span>'+fmt.format(headquarters.count)+'건 · '+money(headquarters.budget)+'</span></div>'
      + '<ul>'+items.map((project) => '<li><b>'+escapeHtml(project.title)+'</b><div class="muted">'+escapeHtml(project.code || '')+' · '+money(project.budget)+'</div></li>').join('')+'</ul></article>';
  }).join('');
  grid.querySelectorAll('[data-nipa-headquarters]').forEach((card) => {
    const open = () => {
      const name = card.dataset.nipaHeadquarters;
      const rows = NIPA_HQ_PROJECTS.filter((project) => project.nipaHeadquarters === name)
        .sort((a, b) => (b.budget || 0) - (a.budget || 0));
      openProjectModal({
        title: name,
        sub: '내부 현황조사 자료를 기준으로 매핑한 NIPA 담당본부별 사업입니다.',
        rows,
        emptyText: '매핑된 사업이 없습니다.',
        countLabel: '담당 사업',
      });
    };
    card.addEventListener('click', open);
    card.addEventListener('keydown', (event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); open(); } });
  });
}
`;
html = html
  .replace('function renderOrgChart() {', `${nipaHeadquartersRenderer}\nfunction renderOrgChart() {`)
  .replace('renderOrgChart();', 'renderOrgChart();\nrenderNipaHeadquarters();');

const excelJsBundlePath = path.join(root, 'node_modules', 'exceljs', 'dist', 'exceljs.min.js');
if (!fs.existsSync(excelJsBundlePath)) throw new Error('엑셀 기능에 필요한 exceljs.min.js를 찾을 수 없습니다. npm install을 실행하세요.');
const excelJsBundle = fs.readFileSync(excelJsBundlePath, 'utf8').replace(/<\/script/gi, '<\\/script');
const excelManagerStyles = `
<style>
.excel-card h2{transform:translateX(-4px)}
.excel-layout{display:grid;grid-template-columns:1fr 1fr;gap:16px}.excel-card{padding:18px}.excel-card h2{margin:0 0 6px}.excel-card h3{margin:18px 0 8px;font-size:15px}.excel-actions{display:flex;flex-wrap:wrap;align-items:center;gap:8px;margin:12px 0}.excel-btn{border:1px solid #b8c7d9;background:#fff;color:#17324f;border-radius:8px;padding:9px 13px;font:inherit;font-weight:800;cursor:pointer}.excel-btn.primary{background:#245f9e;color:#fff;border-color:#245f9e}.excel-btn:disabled{opacity:.45;cursor:not-allowed}.excel-fields{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:0 18px;padding:8px 12px;background:#f7fafc;border:1px solid var(--line-soft);border-radius:8px}.excel-fields label,.excel-project-list label{display:grid;grid-template-columns:18px minmax(0,1fr);gap:9px;align-items:start;font-size:13px;line-height:1.45}.excel-fields label{min-height:35px;align-items:center;padding:4px 0}.excel-fields input,.excel-project-list input{width:16px;height:16px;margin:2px 0 0;accent-color:#245f9e}.excel-project-list{max-height:310px;overflow:auto;border:1px solid var(--line-soft);border-radius:8px;padding:6px 12px;background:#fff}.excel-project-list label{min-height:48px;padding:9px 0;border-bottom:1px solid var(--line-soft)}.excel-project-list label:last-child{border-bottom:0}.excel-upload{display:block;border:1px dashed #92aac3;border-radius:8px;padding:14px;background:#f7fafc}.excel-status{margin-top:10px;padding:10px 12px;border-radius:8px;background:#eef5fb;color:#254766;font-size:13px;white-space:pre-line}.excel-status.error{background:#fff0f0;color:#922}.excel-preview{max-height:330px;overflow:auto;border:1px solid var(--line-soft);border-radius:8px}.excel-preview table{width:100%;border-collapse:collapse;font-size:12px}.excel-preview th,.excel-preview td{padding:7px;border-bottom:1px solid var(--line-soft);text-align:left;vertical-align:top}.excel-preview th{position:sticky;top:0;background:#f7fafc}.excel-help{font-size:13px;color:var(--muted);line-height:1.55}.excel-badge{display:inline-block;border-radius:999px;padding:2px 7px;font-size:11px;font-weight:800;background:#e7f1fb;color:#245f9e}.excel-badge.bad{background:#ffe8e8;color:#922}.excel-filter-summary{align-self:center;color:#516b85;font-size:12px;font-weight:700}.excel-section-divider{height:1px;margin:22px 0;background:var(--line-soft)}.excel-filter-modal{position:fixed;inset:0;z-index:1200;display:grid;place-items:center;padding:20px;background:rgba(14,32,52,.48)}.excel-filter-modal[hidden]{display:none}.excel-filter-dialog{width:min(620px,100%);max-height:90vh;overflow:auto;background:#fff;border-radius:14px;box-shadow:0 24px 70px rgba(11,35,60,.28)}.excel-filter-head{display:flex;align-items:center;justify-content:space-between;padding:18px 20px;border-bottom:1px solid var(--line-soft)}.excel-filter-head h3{margin:0;font-size:19px}.excel-filter-body{display:grid;grid-template-columns:1fr 1fr;gap:14px;padding:20px}.excel-filter-body label{display:grid;gap:6px;color:#39536e;font-size:12px;font-weight:800}.excel-filter-body label:first-child{grid-column:1/-1}.excel-filter-body input,.excel-filter-body select{width:100%;box-sizing:border-box;border:1px solid #b8c7d9;border-radius:8px;padding:10px;background:#fff;color:#17324f;font:inherit}.excel-filter-foot{display:flex;justify-content:space-between;gap:8px;padding:0 20px 20px}
@media(max-width:820px){.excel-layout{grid-template-columns:1fr}.excel-fields{grid-template-columns:1fr}}
</style>`;
const excelManagerPanel = `
<section id="excel-manager" class="panel">
  <div class="excel-layout">
    <article class="card excel-card">
      <h2>엑셀 내보내기</h2>
      <p class="excel-help">엑셀 자료관리 전용 필터로 대상을 좁힌 뒤 사업과 열을 선택해 엑셀로 저장합니다.</p>
      <h3>내보낼 항목</h3>
      <div class="excel-fields" id="excelFieldList"></div>
      <div class="excel-actions"><button class="excel-btn" id="excelSelectAllFields" type="button">항목 전체 선택</button><button class="excel-btn" id="excelClearFields" type="button">항목 해제</button></div>
      <h3>사업 선택 <span class="excel-badge" id="excelProjectCount">0건</span></h3>
      <div class="excel-actions"><button class="excel-btn primary" id="excelOpenFilter" type="button">필터 적용하기</button><span class="excel-filter-summary" id="excelFilterSummary">전체 사업</span><button class="excel-btn" id="excelSelectAllProjects" type="button">사업 전체 선택</button><button class="excel-btn" id="excelClearProjects" type="button">사업 해제</button><button class="excel-btn" id="excelQuickResetFilter" type="button">필터 초기화</button></div>
      <div class="excel-project-list" id="excelProjectList"></div>
      <div class="excel-actions"><button class="excel-btn primary" id="excelExportSelected" type="button">선택 자료 엑셀 다운로드</button><button class="excel-btn" id="excelExportAll" type="button">전체 DB 엑셀 다운로드</button></div>
      <div class="excel-status" id="excelExportStatus">현재 필터를 불러온 뒤 사업과 항목을 선택하세요.</div>
    </article>
    <article class="card excel-card">
      <h2>신규 사업 추가</h2>
      <p class="excel-help">기본 양식을 내려받아 작성한 후 업로드하세요. 사업명과 세부사업명이 다르면 내역사업 카드로 추가됩니다.</p>
      <div class="excel-actions"><button class="excel-btn primary" id="excelDownloadTemplate" type="button">신규 사업 입력 양식 다운로드</button></div>
      <label class="excel-upload">작성한 엑셀 파일 선택<input id="excelImportFile" type="file" accept=".xlsx" style="display:block;margin-top:9px"></label>
      <div class="excel-status" id="excelImportStatus">아직 선택한 파일이 없습니다.</div>
      <div class="excel-preview" id="excelImportPreview" hidden></div>
      <div class="excel-actions"><button class="excel-btn primary" id="excelConfirmImport" type="button" disabled>검증된 사업 추가</button><button class="excel-btn" id="excelClearImported" type="button">브라우저 추가자료 초기화</button></div>
      <p class="excel-help">추가자료는 현재 브라우저의 로컬 저장소에 보존됩니다. 다른 PC와 공유하려면 ‘전체 DB 엑셀 다운로드’로 파일을 저장해 전달하세요.</p>
      <div class="excel-section-divider"></div>
      <h2>기존 사업 수정</h2>
      <p class="excel-help">‘전체 DB 엑셀 다운로드’ 파일의 내용을 수정해 업로드하세요. 사업코드·세부사업명·사업명이 일치하는 기존 카드의 분류, 예산, 개요, 주요내용, 추진방향을 갱신합니다.</p>
      <label class="excel-upload">수정한 DB 엑셀 파일 선택<input id="excelUpdateFile" type="file" accept=".xlsx" style="display:block;margin-top:9px"></label>
      <div class="excel-status" id="excelUpdateStatus">아직 선택한 수정 파일이 없습니다.</div>
      <div class="excel-preview" id="excelUpdatePreview" hidden></div>
      <div class="excel-actions"><button class="excel-btn primary" id="excelConfirmUpdate" type="button" disabled>검증된 내용으로 수정</button><button class="excel-btn" id="excelClearUpdates" type="button">브라우저 수정내용 초기화</button></div>
      <p class="excel-help">원본 데이터 파일은 변경하지 않으며, 확정한 수정내용은 현재 브라우저에 저장됩니다. 수정 전 전체 DB 엑셀을 백업해 두는 것을 권장합니다.</p>
    </article>
  </div>
  <div class="excel-filter-modal" id="excelFilterModal" role="dialog" aria-modal="true" aria-labelledby="excelFilterTitle" hidden>
    <div class="excel-filter-dialog">
      <div class="excel-filter-head"><h3 id="excelFilterTitle">엑셀 내보내기 필터</h3><button class="excel-btn" id="excelCloseFilter" type="button" aria-label="닫기">닫기</button></div>
      <div class="excel-filter-body">
        <label>사업명·사업코드 검색<input id="excelFilterQuery" type="search" placeholder="검색어 입력"></label>
        <label>레이어<select id="excelFilterLayer"></select></label>
        <label>회계<select id="excelFilterAccount"></select></label>
        <label>과기정통부 소관<select id="excelFilterOrg"></select></label>
        <label>NIPA 본부<select id="excelFilterHq"></select></label>
      </div>
      <div class="excel-filter-foot"><button class="excel-btn" id="excelResetFilter" type="button">조건 초기화</button><button class="excel-btn primary" id="excelApplyFilter" type="button">필터 적용</button></div>
    </div>
  </div>
</section>`;
html = html
  .replace('</head>', `${excelManagerStyles}\n</head>`)
  .replace(/(<button class="tab" data-panel="nipa-headquarters">NIPA 본부별 사업현황<\/button>)/, '$1\n    <button class="tab" data-panel="excel-manager">DB 관리</button>')
  .replace('</main>', `${excelManagerPanel}\n</main>`);

const excelManagerScript = `
<script>${excelJsBundle}</script>
<script>
(() => {
  const STORAGE_KEY = 'nipa-budget-custom-projects-v1';
  const UPDATE_KEY = 'nipa-budget-project-updates-v2';
  const FIELD_DEFS = [
    ['사업명','title'],['세부사업명','parentTitle'],['사업코드','code'],['회계','account'],
    ['레이어','layer'],['과기정통부 소관','mappedOrgName'],['NIPA 본부','nipaHeadquarters'],['2026년 예산(백만원)','budget'],
    ['개요','overview'],['주요내용','mainPoints'],['추진방향','direction']
  ];
  const REQUIRED = ['사업명','사업코드','회계','레이어','과기정통부 소관','NIPA 본부','2026년 예산(백만원)','개요','주요내용','추진방향'];
  const VALID_ACCOUNTS = ['일반회계','특별회계','기금'];
  const VALID_LAYERS = ['AI인프라','AI반도체','AI모델','지역AX','피지컬AI','글로벌 이니셔티브','AI생태계','산업전략','기관운영비'];
  let exportRows = [];
  let pendingRows = [];
  let pendingUpdates = [];
  let excelFilters = {query:'',layer:'',account:'',org:'',hq:''};
  const byId = (id) => document.getElementById(id);
  const cellText = (value) => value == null ? '' : String(value).trim();
  const listText = (value) => Array.isArray(value) ? value.join(' / ') : cellText(value);
  const setStatus = (id, message, error = false) => { const el=byId(id); el.textContent=message; el.classList.toggle('error',error); };
  const downloadWorkbook = async (workbook, name) => {
    const buffer = await workbook.xlsx.writeBuffer();
    const url = URL.createObjectURL(new Blob([buffer], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
    const anchor = document.createElement('a'); anchor.href=url; anchor.download=name; anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  };
  const styleSheet = (sheet) => {
    sheet.views = [{state:'frozen', ySplit:1}];
    sheet.autoFilter = {from:{row:1,column:1},to:{row:Math.max(1,sheet.rowCount),column:sheet.columnCount}};
    sheet.getRow(1).eachCell((cell) => {cell.font={bold:true,color:{argb:'FFFFFFFF'}};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF245F9E'}};cell.alignment={vertical:'middle'};});
    sheet.columns.forEach((column) => {column.width=Math.min(45,Math.max(14,...column.values.slice(1).map((v)=>String(v||'').length+2)));});
  };
  const exportValue = (project, key) => {
    if (key === 'title') return project.detailTitle || String(project.title||'').replace(/^\\(내역\\)\\s*/, '');
    if (key === 'parentTitle') return project.parentTitle || project.title || '';
    if (key === 'nipaHeadquarters') return project.nipaHeadquarters || (typeof projectNipaHeadquarters==='function' ? projectNipaHeadquarters(project) : '');
    if (key === 'mainPoints') return listText(project.sourceMainItems || project.detail?.mainPoints || project.mainPoints);
    return project[key] ?? project.detail?.[key] ?? '';
  };
  const exportProjects = async (rows, fields, filename) => {
    if (!rows.length || !fields.length) throw new Error('사업과 항목을 각각 하나 이상 선택하세요.');
    const workbook = new ExcelJS.Workbook(); workbook.creator='NIPA Budget Dashboard';
    const sheet = workbook.addWorksheet('사업자료');
    sheet.addRow(fields.map(([label])=>label));
    rows.forEach((project)=>sheet.addRow(fields.map(([,key])=>exportValue(project,key))));
    styleSheet(sheet); await downloadWorkbook(workbook,filename);
  };
  const renderFields = () => {byId('excelFieldList').innerHTML=FIELD_DEFS.map(([label,key])=>'<label><input type="checkbox" data-excel-field="'+key+'" checked> '+escapeHtml(label)+'</label>').join('');};
  const renderExportProjects = () => {
    const selected = new Set([...document.querySelectorAll('[data-excel-project]:checked')].map((el)=>el.value));
    byId('excelProjectList').innerHTML=exportRows.map((project,index)=>{const id=String(project.no);return '<label><input type="checkbox" data-excel-project value="'+escapeAttr(id)+'" '+(!selected.size||selected.has(id)?'checked':'')+'> <span><b>'+escapeHtml(project.title)+'</b><br><span class="muted">'+escapeHtml(project.code||'')+' · '+escapeHtml(project.layer||'')+' · '+money(project.budget)+'</span></span></label>';}).join('')||'<div class="muted" style="padding:10px">현재 조건에 해당하는 사업이 없습니다.</div>';
    updateExportCount();
  };
  const updateExportCount = () => {const count=document.querySelectorAll('[data-excel-project]:checked').length;byId('excelProjectCount').textContent=count+'건 선택';};
  const uniqueOptions = (values) => [...new Set(values.filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),'ko'));
  const excelProjectHqs = (project) => {
    const matched=(typeof NIPA_HQ_PROJECTS!=='undefined'&&typeof projectMatchKey==='function')?NIPA_HQ_PROJECTS.filter((item)=>projectMatchKey(item)===projectMatchKey(project)).map((item)=>item.nipaHeadquarters).filter(Boolean):[];
    if(matched.length)return uniqueOptions(matched);
    const fallback=project.nipaHeadquarters||(typeof projectNipaHeadquarters==='function'?projectNipaHeadquarters(project):'');
    return uniqueOptions(cellText(fallback).split('·').map((value)=>value.trim()));
  };
  const excelProjectHq = (project) => excelProjectHqs(project).join(' · ');
  const fillFilterSelect = (id,label,values) => {byId(id).innerHTML='<option value="">'+label+'</option>'+values.map((value)=>'<option value="'+escapeAttr(value)+'">'+escapeHtml(value)+'</option>').join('');};
  const initExcelFilters = () => {
    fillFilterSelect('excelFilterLayer','전체 레이어',LAYERS.map((x)=>x.id));
    fillFilterSelect('excelFilterAccount','전체 회계',uniqueOptions(PROJECTS.map((x)=>x.account)));
    fillFilterSelect('excelFilterOrg','전체 소관',uniqueOptions(PROJECTS.map((x)=>x.mappedOrgName)));
    fillFilterSelect('excelFilterHq','전체 본부',uniqueOptions(PROJECTS.flatMap(excelProjectHqs)));
    byId('excelFilterQuery').value=excelFilters.query;
    byId('excelFilterLayer').value=excelFilters.layer;byId('excelFilterAccount').value=excelFilters.account;byId('excelFilterOrg').value=excelFilters.org;byId('excelFilterHq').value=excelFilters.hq;
  };
  const applyExcelFilters = () => {
    excelFilters={query:cellText(byId('excelFilterQuery').value).toLowerCase(),layer:byId('excelFilterLayer').value,account:byId('excelFilterAccount').value,org:byId('excelFilterOrg').value,hq:byId('excelFilterHq').value};
    exportRows=PROJECTS.filter((project)=>{const hay=[project.title,project.detailTitle,project.parentTitle,project.code].join(' ').toLowerCase();return (!excelFilters.query||hay.includes(excelFilters.query))&&(!excelFilters.layer||project.layer===excelFilters.layer)&&(!excelFilters.account||project.account===excelFilters.account)&&(!excelFilters.org||project.mappedOrgName===excelFilters.org)&&(!excelFilters.hq||excelProjectHqs(project).includes(excelFilters.hq));});
    renderExportProjects();const count=Object.values(excelFilters).filter(Boolean).length;byId('excelFilterSummary').textContent=(count?count+'개 조건 · ':'전체 · ')+exportRows.length+'건';setStatus('excelExportStatus','엑셀 자료관리 필터 결과 '+exportRows.length+'건을 불러왔습니다.');byId('excelFilterModal').hidden=true;
  };
  const resetExcelFilters = () => {byId('excelFilterQuery').value='';['excelFilterLayer','excelFilterAccount','excelFilterOrg','excelFilterHq'].forEach((id)=>byId(id).value='');};
  const loadCurrentFilter = () => {initExcelFilters();applyExcelFilters();};
  const templateWorkbook = async () => {
    const wb=new ExcelJS.Workbook(); const sheet=wb.addWorksheet('신규사업입력');
    sheet.addRow(FIELD_DEFS.map(([label])=>label));
    sheet.addRow(['예시 사업','예시 세부사업','0000-000','일반회계','AI생태계','제2차관 / 인공지능정책실 / 담당관 / 담당과','AI활용본부',1000,'사업의 목적과 지원 대상을 작성','핵심 활동을 문장 또는 줄바꿈으로 작성','사업 특성에 맞는 추진방향을 작성']);
    styleSheet(sheet); sheet.getRow(2).font={italic:true,color:{argb:'FF6B7280'}};
    const guide=wb.addWorksheet('작성안내'); guide.addRows([
      ['항목','작성 기준'],['사업명','필수. 카드에 표시할 사업 또는 내역사업명'],['세부사업명','선택. 사업명과 다르면 내역사업으로 등록'],['사업코드','필수. 0000-000 형식'],['회계','일반회계, 특별회계, 기금 중 선택'],['레이어',VALID_LAYERS.join(', ')],['과기정통부 소관','조직 경로를 / 로 구분'],['NIPA 본부','담당 본부명'],['2026년 예산(백만원)','숫자만 입력'],['개요·주요내용·추진방향','각 사업에 해당하는 내용만 작성']
    ]); styleSheet(guide); await downloadWorkbook(wb,'NIPA_신규사업_입력양식.xlsx');
  };
  const worksheetRows = (sheet) => {
    const headers=sheet.getRow(1).values.slice(1).map(cellText); const rows=[];
    for(let r=2;r<=sheet.rowCount;r++){const obj={};headers.forEach((h,i)=>obj[h]=cellText(sheet.getRow(r).getCell(i+1).value?.text ?? sheet.getRow(r).getCell(i+1).value));if(Object.values(obj).some(Boolean))rows.push({...obj,__row:r});} return {headers,rows};
  };
  const normalizedName = (value) => cellText(value).replace(/^\\(내역\\)\\s*/,'');
  const projectMatchKey = (project) => [project.code,normalizedName(project.parentTitle||project.title),normalizedName(project.detailTitle||project.title)].join('|');
  const rowMatchKey = (row) => [row['사업코드'],normalizedName(row['세부사업명']||row['사업명']),normalizedName(row['사업명'])].join('|');
  const validateUpdate = ({headers,rows}) => {
    const missingHeaders=FIELD_DEFS.map(([label])=>label).filter((label)=>!headers.includes(label));
    if(missingHeaders.length)return {valid:[],invalid:[{__row:1,__errors:['양식 열 누락: '+missingHeaders.join(', ')]}]};
    const targets=new Map();PROJECTS.forEach((project)=>{const key=projectMatchKey(project);targets.set(key,[...(targets.get(key)||[]),project]);});
    const valid=[],invalid=[];
    rows.forEach((row)=>{const errors=[];REQUIRED.forEach((key)=>{if(!cellText(row[key]))errors.push(key+' 필수');});
      if(row['사업코드']&&!/^\\d{4}-\\d{3}$/.test(row['사업코드']))errors.push('사업코드 형식 오류');
      if(row['회계']&&!VALID_ACCOUNTS.includes(row['회계']))errors.push('회계 선택값 오류');
      if(row['레이어']&&!VALID_LAYERS.includes(row['레이어']))errors.push('레이어 선택값 오류');
      const budget=Number(String(row['2026년 예산(백만원)']).replace(/,/g,''));if(!Number.isFinite(budget)||budget<0)errors.push('예산은 0 이상의 숫자');
      const matches=targets.get(rowMatchKey(row))||[];if(!matches.length)errors.push('일치하는 기존 사업 없음');if(matches.length>1)errors.push('동일 식별값의 기존 사업이 여러 건');
      const target={...row,__budget:budget,__targetNo:matches[0]?.no,__errors:errors};(errors.length?invalid:valid).push(target);
    });return {valid,invalid};
  };
  const showUpdatePreview = (result) => {
    pendingUpdates=result.valid;const all=[...result.valid,...result.invalid].sort((a,b)=>a.__row-b.__row);
    byId('excelUpdatePreview').hidden=false;byId('excelUpdatePreview').innerHTML='<table><thead><tr><th>행</th><th>사업명</th><th>코드</th><th>판정</th></tr></thead><tbody>'+all.map((r)=>'<tr><td>'+r.__row+'</td><td>'+escapeHtml(r['사업명']||'')+'</td><td>'+escapeHtml(r['사업코드']||'')+'</td><td>'+(r.__errors.length?'<span class="excel-badge bad">'+escapeHtml(r.__errors.join(' · '))+'</span>':'<span class="excel-badge">수정 가능</span>')+'</td></tr>').join('')+'</tbody></table>';
    byId('excelConfirmUpdate').disabled=!result.valid.length;setStatus('excelUpdateStatus','수정 가능 '+result.valid.length+'건 · 오류 '+result.invalid.length+'건',result.invalid.length>0);
  };
  const updatePatchFromRow = (row) => {const layer=LAYERS.find((x)=>x.id===row['레이어'])||{};const main=String(row['주요내용']).split(/\\r?\\n|\\s*\\/\\s*/).map((x)=>x.trim()).filter(Boolean);return {account:row['회계'],layer:row['레이어'],layerName:layer.name||row['레이어'],layerColor:layer.color||'#607d9b',mappedOrgName:row['과기정통부 소관'],orgName:row['과기정통부 소관'],orgUnit:row['과기정통부 소관'],orgDivision:String(row['과기정통부 소관']).split(' / ')[0],orgDepartments:[row['과기정통부 소관']],orgPathKey:String(row['과기정통부 소관']).split(' / ').slice(0,3).join(' / '),nipaHeadquarters:row['NIPA 본부'],budget:row.__budget,overview:row['개요'],sourcePurposeItems:[row['개요']],sourceMainItems:main,mainPoints:main.join(' / '),direction:row['추진방향'],detail:{overview:row['개요'],mainPoints:main,direction:row['추진방향']}};};
  const applyStoredUpdate = (record) => {
    const primary=PROJECTS.find((project)=>String(project.no)===String(record.no));if(!primary)return false;const old={layer:primary.layer,budget:Number(primary.budget)||0,hq:excelProjectHq(primary)};const patch=record.patch||{};
    [PROJECTS,ORG_PROJECTS,NIPA_HQ_PROJECTS].forEach((rows)=>rows.filter((project)=>String(project.no).replace(/-(org|hq)$/,'')===String(record.no)).forEach((project)=>{const oldDetail=project.detail||{};Object.assign(project,patch);project.detail={...oldDetail,...(patch.detail||{})};}));
    const newBudget=Number(patch.budget)||0;if(old.layer!==patch.layer){const oldLayer=LAYERS.find((x)=>x.id===old.layer),newLayer=LAYERS.find((x)=>x.id===patch.layer);if(oldLayer){oldLayer.count=Math.max(0,(oldLayer.count||0)-1);oldLayer.budget=(oldLayer.budget||0)-old.budget;}if(newLayer){newLayer.count=(newLayer.count||0)+1;newLayer.budget=(newLayer.budget||0)+newBudget;}}else{const layer=LAYERS.find((x)=>x.id===old.layer);if(layer)layer.budget=(layer.budget||0)-old.budget+newBudget;}
    const newHq=patch.nipaHeadquarters||old.hq;if(old.hq!==newHq){const oldItem=NIPA_HEADQUARTERS.find((x)=>x.name===old.hq),newItem=NIPA_HEADQUARTERS.find((x)=>x.name===newHq);if(oldItem){oldItem.count=Math.max(0,(oldItem.count||0)-1);oldItem.budget=(oldItem.budget||0)-old.budget;}if(newItem){newItem.count=(newItem.count||0)+1;newItem.budget=(newItem.budget||0)+newBudget;}else NIPA_HEADQUARTERS.push({name:newHq,color:'#607d9b',count:1,budget:newBudget});}else{const item=NIPA_HEADQUARTERS.find((x)=>x.name===old.hq);if(item)item.budget=(item.budget||0)-old.budget+newBudget;}return true;
  };
  const validateImport = ({headers,rows}) => {
    const missingHeaders=FIELD_DEFS.map(([label])=>label).filter((label)=>!headers.includes(label));
    if(missingHeaders.length)return {valid:[],invalid:[{__row:1,__errors:['양식 열 누락: '+missingHeaders.join(', ')]}]};
    const existing=new Set(PROJECTS.map((p)=>[p.code,p.parentTitle||p.title,p.detailTitle||p.title].map((v)=>cellText(v).replace(/^\\(내역\\)\\s*/,'')).join('|')));
    const seen=new Set(); const valid=[],invalid=[];
    rows.forEach((row)=>{const errors=[];REQUIRED.forEach((key)=>{if(!cellText(row[key]))errors.push(key+' 필수');});
      if(row['사업코드']&&!/^\\d{4}-\\d{3}$/.test(row['사업코드']))errors.push('사업코드 형식 오류');
      if(row['회계']&&!VALID_ACCOUNTS.includes(row['회계']))errors.push('회계 선택값 오류');
      if(row['레이어']&&!VALID_LAYERS.includes(row['레이어']))errors.push('레이어 선택값 오류');
      const budget=Number(String(row['2026년 예산(백만원)']).replace(/,/g,''));if(!Number.isFinite(budget)||budget<0)errors.push('예산은 0 이상의 숫자');
      const parent=row['세부사업명']||row['사업명'];const key=[row['사업코드'],parent,row['사업명']].join('|');if(existing.has(key)||seen.has(key))errors.push('기존 또는 파일 내 중복 사업');seen.add(key);
      const target={...row,__budget:budget,__errors:errors};(errors.length?invalid:valid).push(target);
    }); return {valid,invalid};
  };
  const showPreview = (result) => {
    pendingRows=result.valid;const all=[...result.valid,...result.invalid].sort((a,b)=>a.__row-b.__row);
    byId('excelImportPreview').hidden=false;byId('excelImportPreview').innerHTML='<table><thead><tr><th>행</th><th>사업명</th><th>코드</th><th>판정</th></tr></thead><tbody>'+all.map((r)=>'<tr><td>'+r.__row+'</td><td>'+escapeHtml(r['사업명']||'')+'</td><td>'+escapeHtml(r['사업코드']||'')+'</td><td>'+(r.__errors.length?'<span class="excel-badge bad">'+escapeHtml(r.__errors.join(' · '))+'</span>':'<span class="excel-badge">추가 가능</span>')+'</td></tr>').join('')+'</tbody></table>';
    byId('excelConfirmImport').disabled=!result.valid.length;setStatus('excelImportStatus','추가 가능 '+result.valid.length+'건 · 오류 '+result.invalid.length+'건',result.invalid.length>0);
  };
  const rowToProject = (row,index) => {
    const title=row['사업명'],parent=row['세부사업명']||title,isDetail=parent!==title,layer=LAYERS.find((x)=>x.id===row['레이어'])||{id:row['레이어'],name:row['레이어'],color:'#607d9b'};
    const main=String(row['주요내용']).split(/\\r?\\n|\\s*\\/\\s*/).map((x)=>x.trim()).filter(Boolean);
    const customId='custom-'+Date.now()+'-'+index;
    return {no:customId,sourceNo:customId,detailNo:1,isDetail,parentTitle:parent,detailTitle:isDetail?title:'',title:isDetail?'(내역) '+title:title,code:row['사업코드'],account:row['회계'],layer:layer.id,layerName:layer.name||layer.id,layerColor:layer.color||'#607d9b',mappedOrgName:row['과기정통부 소관'],orgName:row['과기정통부 소관'],orgUnit:row['과기정통부 소관'],orgDivision:String(row['과기정통부 소관']).split(' / ')[0],orgDepartments:[row['과기정통부 소관']],orgPathKey:String(row['과기정통부 소관']).split(' / ').slice(0,3).join(' / '),nipaHeadquarters:row['NIPA 본부'],operator:'정보통신산업진흥원',budget:row.__budget,budgetSource:'사용자 엑셀 추가',overview:row['개요'],sourcePurposeItems:[row['개요']],sourceMainItems:main,mainPoints:main.join(' / '),direction:row['추진방향'],displayMode:isDetail?'nipa-detail-only':'custom',detailBreakdown:[{title,budget:row.__budget}],detail:{type:layer.name||layer.id,displayTitle:title,parentTitle:parent,overview:row['개요'],mainPoints:main,direction:row['추진방향'],details:[{title,budget:row.__budget}],source:'사용자 업로드 엑셀'}};
  };
  const registerProject = (project) => {
    PROJECTS.push(project); ORG_PROJECTS.push({...project,no:project.no+'-org'}); NIPA_HQ_PROJECTS.push({...project,no:project.no+'-hq'});
    let layer=LAYERS.find((x)=>x.id===project.layer);if(layer){layer.count=(layer.count||0)+1;layer.budget=(layer.budget||0)+project.budget;}
    let hq=NIPA_HEADQUARTERS.find((x)=>x.name===project.nipaHeadquarters);if(hq){hq.count=(hq.count||0)+1;hq.budget=(hq.budget||0)+project.budget;}else NIPA_HEADQUARTERS.push({name:project.nipaHeadquarters,color:'#607d9b',count:1,budget:project.budget});
  };
  const loadStored = () => {try{JSON.parse(localStorage.getItem(STORAGE_KEY)||'[]').forEach(registerProject);}catch(error){console.warn(error);}};
  const loadUpdates = () => {try{JSON.parse(localStorage.getItem(UPDATE_KEY)||'[]').forEach(applyStoredUpdate);}catch(error){console.warn(error);}};
  const commitImport = () => {const projects=pendingRows.map(rowToProject);projects.forEach(registerProject);const stored=JSON.parse(localStorage.getItem(STORAGE_KEY)||'[]');localStorage.setItem(STORAGE_KEY,JSON.stringify([...stored,...projects]));pendingRows=[];byId('excelConfirmImport').disabled=true;renderDashboard();renderEcosystem();renderOrgChart();renderNipaHeadquarters();loadCurrentFilter();setStatus('excelImportStatus',projects.length+'건을 추가하고 현재 브라우저에 저장했습니다.');};
  const commitUpdate = () => {const records=pendingUpdates.map((row)=>({no:row.__targetNo,patch:updatePatchFromRow(row)}));const stored=JSON.parse(localStorage.getItem(UPDATE_KEY)||'[]'),merged=new Map(stored.map((item)=>[String(item.no),item]));records.forEach((item)=>{merged.set(String(item.no),item);applyStoredUpdate(item);});localStorage.setItem(UPDATE_KEY,JSON.stringify([...merged.values()]));pendingUpdates=[];byId('excelConfirmUpdate').disabled=true;renderDashboard();renderEcosystem();renderOrgChart();renderNipaHeadquarters();loadCurrentFilter();setStatus('excelUpdateStatus',records.length+'건을 수정하고 현재 브라우저에 저장했습니다.');};
  renderFields(); loadStored(); loadUpdates();
  byId('excelOpenFilter').addEventListener('click',()=>{byId('excelFilterModal').hidden=false;byId('excelFilterQuery').focus();});byId('excelCloseFilter').addEventListener('click',()=>byId('excelFilterModal').hidden=true);byId('excelFilterModal').addEventListener('click',(event)=>{if(event.target===byId('excelFilterModal'))byId('excelFilterModal').hidden=true;});byId('excelApplyFilter').addEventListener('click',applyExcelFilters);byId('excelResetFilter').addEventListener('click',()=>{resetExcelFilters();applyExcelFilters();});byId('excelProjectList').addEventListener('change',updateExportCount);
  byId('excelSelectAllFields').addEventListener('click',()=>document.querySelectorAll('[data-excel-field]').forEach((x)=>x.checked=true));byId('excelClearFields').addEventListener('click',()=>document.querySelectorAll('[data-excel-field]').forEach((x)=>x.checked=false));
  byId('excelSelectAllProjects').addEventListener('click',()=>{document.querySelectorAll('[data-excel-project]').forEach((x)=>x.checked=true);updateExportCount();});byId('excelClearProjects').addEventListener('click',()=>{document.querySelectorAll('[data-excel-project]').forEach((x)=>x.checked=false);updateExportCount();});
  byId('excelQuickResetFilter').addEventListener('click',()=>{resetExcelFilters();applyExcelFilters();});
  byId('excelExportSelected').addEventListener('click',async()=>{try{const ids=new Set([...document.querySelectorAll('[data-excel-project]:checked')].map((x)=>x.value)),rows=exportRows.filter((x)=>ids.has(String(x.no))),keys=new Set([...document.querySelectorAll('[data-excel-field]:checked')].map((x)=>x.dataset.excelField)),fields=FIELD_DEFS.filter(([,key])=>keys.has(key));await exportProjects(rows,fields,'NIPA_선택사업.xlsx');setStatus('excelExportStatus',rows.length+'건을 내려받았습니다.');}catch(error){setStatus('excelExportStatus',error.message,true);}});
  byId('excelExportAll').addEventListener('click',async()=>{try{await exportProjects(PROJECTS,FIELD_DEFS,'NIPA_전체사업DB.xlsx');setStatus('excelExportStatus',PROJECTS.length+'건의 전체 DB를 내려받았습니다.');}catch(error){setStatus('excelExportStatus',error.message,true);}});
  byId('excelDownloadTemplate').addEventListener('click',templateWorkbook);
  byId('excelImportFile').addEventListener('change',async(event)=>{try{const file=event.target.files[0];if(!file)return;const wb=new ExcelJS.Workbook();await wb.xlsx.load(await file.arrayBuffer());const sheet=wb.getWorksheet('신규사업입력')||wb.worksheets[0];showPreview(validateImport(worksheetRows(sheet)));}catch(error){pendingRows=[];byId('excelConfirmImport').disabled=true;setStatus('excelImportStatus','파일을 읽지 못했습니다: '+error.message,true);}});
  byId('excelConfirmImport').addEventListener('click',commitImport);byId('excelClearImported').addEventListener('click',()=>{if(!confirm('이 브라우저에서 추가한 사업을 모두 삭제할까요?'))return;localStorage.removeItem(STORAGE_KEY);location.reload();});
  byId('excelUpdateFile').addEventListener('change',async(event)=>{try{const file=event.target.files[0];if(!file)return;const wb=new ExcelJS.Workbook();await wb.xlsx.load(await file.arrayBuffer());const sheet=wb.getWorksheet('사업자료')||wb.worksheets[0];showUpdatePreview(validateUpdate(worksheetRows(sheet)));}catch(error){pendingUpdates=[];byId('excelConfirmUpdate').disabled=true;setStatus('excelUpdateStatus','수정 파일을 읽지 못했습니다: '+error.message,true);}});
  byId('excelConfirmUpdate').addEventListener('click',commitUpdate);byId('excelClearUpdates').addEventListener('click',()=>{if(!confirm('이 브라우저에 저장된 기존사업 수정내용을 모두 초기화할까요?'))return;localStorage.removeItem(UPDATE_KEY);location.reload();});
  document.querySelector('[data-panel="excel-manager"]').addEventListener('click',loadCurrentFilter);
  renderDashboard(); renderEcosystem(); renderOrgChart(); renderNipaHeadquarters();
})();
</script>`;
html = html.replace('</body>', `${excelManagerScript}\n</body>`);

const payload = {
  generatedAt: new Date().toISOString(),
  source: '과학기술정보통신부 2026년 예산 및 기금운용계획 사업설명자료 1~8권',
  selectionRule: `AI 여부와 관계없이 사업시행주체 문자열에 '${operatorName}' 포함`,
  projectRows: projects.length,
  parentProjects: parentCount,
  sharedOperatorRows: sharedOperatorCount,
  budgetTotalMillionKrw: totalBudget,
  budgetTotalExcludingInstitutionalMillionKrw: budgetExcludingInstitutional,
  rows: projects,
};

fs.writeFileSync(outputPath, html, 'utf8');
fs.writeFileSync(indexPath, html, 'utf8');
fs.writeFileSync(dataPath, JSON.stringify(payload, null, 2), 'utf8');
fs.writeFileSync(verificationPath, JSON.stringify({
  generatedAt: payload.generatedAt,
  sourceDashboard: path.basename(sourcePath),
  sourceProjectRows: allProjects.length,
  rawNipaTextCandidateBlocks: auditedRaw.length,
  evidenceQualifiedNipaBlocks: audited.length,
  selectedProjectRows: projects.length,
  parentProjects: parentCount,
  parentCardRows: projects.filter((project) => project.displayMode !== 'nipa-detail-only').length,
  nipaDetailOnlyRows: projects.filter((project) => project.displayMode === 'nipa-detail-only').length,
  requestedControlTotalMillionKrw: 3020457,
  calculatedTotalMillionKrw: totalBudget,
  controlTotalDifferenceMillionKrw: totalBudget - 3020457,
  completeCardNarratives: projects.every((project) => project.detail?.overview
    && project.detail?.mainPoints?.length && project.detail?.direction && project.detail?.source),
  organizationMappingComplete: projects.every((project) => project.mappedOrgName
    && project.orgPathKey && project.orgDepartments?.length),
  organizationChartGroups: orgs.length,
  organizationChartCountReconciles: orgs.reduce((sum, org) => sum + org.count, 0) === projects.length,
  organizationChartBudgetReconciles: orgs.reduce((sum, org) => sum + org.budget, 0) === totalBudget,
  sharedOperatorRows: sharedOperatorCount,
  everyRowMentionsNipa: projects.every((project) => String(project.operator || '').includes(operatorName)),
  excludedKnownFalsePositives: {
    '7176-104': !projects.some((project) => project.code === '7176-104'),
    '7176-204': !projects.some((project) => project.code === '7176-204'),
    '2533-301': !projects.some((project) => project.code === '2533-301'),
    '7002-102': !projects.some((project) => project.code === '7002-102'),
  },
  inclusionEvidenceRule: '사업 머리부의 사업시행주체에 정보통신산업진흥원이 명시되거나, 사업명 자체가 정보통신산업진흥원 지원사업인 경우만 포함. 후속 내역표는 포함 판정이 아니라 NIPA 배정액 분리에만 사용',
  layerCountsReconcile: layers.reduce((sum, layer) => sum + layer.count, 0) === projects.length,
  layerBudgetReconciles: layers.reduce((sum, layer) => sum + layer.budget, 0) === totalBudget,
  accounts: [...new Set(projects.map((project) => project.account))],
  sourceBooks: [...new Set(projects.map((project) => project.sourceBook))],
}, null, 2), 'utf8');

console.log(JSON.stringify({ outputPath, indexPath, dataPath, verificationPath, projectRows: projects.length, parentCount, totalBudget, budgetExcludingInstitutional }, null, 2));
