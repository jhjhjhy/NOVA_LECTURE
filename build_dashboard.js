const xlsx = require('xlsx');
const fs = require('fs');

// ─── 강의명 정규화 (Node + 브라우저 공용 로직) ───────────────────────────
function normalizeLectureName(name) {
  if (!name) return '';
  let n = name.trim();
  // 반복 적용: 가격 패턴 제거 (예: -145만원, -200원, -145)
  n = n.replace(/\s*[-]\s*[\d,]+\s*(만원|원)/gi, '');
  n = n.replace(/\s*[-]\s*\d+\s*$/g, ''); // 후행 하이픈+숫자만 (예: 클래스-145, 강의-100)
  // 분할결제 회차 제거 (예: (1), (2), (3))
  n = n.replace(/\s*\(\d+\)\s*$/g, '');
  // 관리용 접미사 제거
  n = n.replace(/\s*[-_]\s*(재결제|재결재|추가결제|프리미엄\s*전용|프리미엄|얼리버드|테스트|복제됨|복제|전용|추가)\s*/gi, '');
  // 후행 하이픈/공백/가 제거
  n = n.replace(/\s*가\s*$/g, ''); // 후행 '가' (ERP 변형 접미사)
  n = n.replace(/[-\s]+$/, '').replace(/\s+/g, ' ').trim();
  return n;
}

// ─── 데이터 로드 ──────────────────────────────────────────────────────────
const wb = xlsx.readFile('C:/Users/ehdtl/Downloads/lecturesales_stats_20260326_041854.xlsx');
const ws = wb.Sheets['강의매출통계'];
const raw = xlsx.utils.sheet_to_json(ws, { header: 1 });
const rows = raw.slice(1).filter(r => r[0] && r[0] !== '시스템노바 테스트용' && r[3] !== 'test');

const data = rows.map(r => ({
  플랫폼: r[0] || '',
  무료강의일: r[1] || '',
  강사: r[2] || '',
  강의명: r[3] || '',
  기수: r[4] || '-',
  무료강의신청수: Number(r[5]) || 0,
  강의총매출: Number(r[6]) || 0,
  강의총매출수강생수: Number(r[7]) || 0,
  플랫폼매출: Number(r[8]) || 0,
  PG사수수료: Number(r[12]) || 0,
  PG제외매출: Number(r[13]) || 0,
  노바수수료: Number(r[14]) || 0,
  광고비: Number(r[15]) || 0,
  기타비용: Number(r[16]) || 0,
  순매출: Number(r[17]) || 0,
  플랫폼수익금: Number(r[18]) || 0,
  인플루언서RS정산금: Number(r[19]) || 0,
  강사정산금: Number(r[20]) || 0,
}));

// ─── 강의그룹명 기준 집계 (기수별 1개 row로 합산) ─────────────────────────
function aggregateByKey(dataRows) {
  const map = {};
  dataRows.forEach(r => {
    const normName = normalizeLectureName(r.강의명);
    if (!normName) return;
    const kisu = (r.기수 && r.기수.trim() !== '') ? r.기수.trim() : '-';
    if (!map[normName]) map[normName] = {
      lecture_key: normName, 강의그룹명: normName, 기수List: [],
      강사: r.강사, 플랫폼: r.플랫폼,
      강의총매출: 0, 순매출: 0, 수강생수: 0, 무료강의신청수: 0,
      PG제외매출: 0, PG사수수료: 0, 노바수수료: 0, 광고비: 0,
      기타비용: 0, 플랫폼수익금: 0, 강사정산금: 0, rows: [], _km: {}
    };
    const d = map[normName];
    if (!d.기수List.includes(kisu)) d.기수List.push(kisu);
    d.강의총매출 += r.강의총매출; d.순매출 += r.순매출; d.수강생수 += r.강의총매출수강생수;
    d.무료강의신청수 = Math.max(d.무료강의신청수, r.무료강의신청수);
    d.PG제외매출 += r.PG제외매출; d.PG사수수료 += r.PG사수수료; d.노바수수료 += r.노바수수료;
    d.광고비 += r.광고비; d.기타비용 += r.기타비용;
    d.플랫폼수익금 += r.플랫폼수익금; d.강사정산금 += r.강사정산금;
    // 기수별 row 합산 (기수당 정확히 1개)
    if (!d._km[kisu]) d._km[kisu] = {
      무료강의일: r.무료강의일, 기수: kisu, 플랫폼: r.플랫폼, 강의명: r.강의명,
      강의총매출: 0, 순매출: 0, 수강생수: 0, 무료강의신청수: 0,
      PG제외매출: 0, 플랫폼매출: 0, 광고비: 0, 노바수수료: 0,
      PG사수수료: 0, 기타비용: 0, 강사정산금: 0
    };
    const k = d._km[kisu];
    k.강의총매출 += r.강의총매출; k.순매출 += r.순매출; k.수강생수 += r.강의총매출수강생수;
    k.무료강의신청수 = Math.max(k.무료강의신청수, r.무료강의신청수);
    k.PG제외매출 += r.PG제외매출; k.플랫폼매출 += r.플랫폼매출;
    k.광고비 += r.광고비; k.노바수수료 += r.노바수수료;
    k.PG사수수료 += r.PG사수수료; k.기타비용 += r.기타비용; k.강사정산금 += r.강사정산금;
  });
  return Object.values(map).map(d => {
    d.rows = Object.values(d._km).sort((a,b) => (a.무료강의일||'').localeCompare(b.무료강의일||''));
    delete d._km;
    return d;
  }).sort((a, b) => b.강의총매출 - a.강의총매출);
}

// ─── 날짜 정규화 (Node.js용) ──────────────────────────────────────────────
function normDateStr(s) {
  if (!s || s === '-') return '';
  let r = '', src = String(s).slice(0, 10);
  for (let i = 0; i < src.length; i++) { const c = src[i]; r += (c === '.' || c === '/') ? '-' : c; }
  return r;
}
// ─── 강의명 유사도 (공백 제거 후 문자 바이그램 Jaccard, 0~1) ─────────────
function nameSimilarity(a, b) {
  a = a.toLowerCase().replace(/\s+/g,'');
  b = b.toLowerCase().replace(/\s+/g,'');
  if (a === b) return 1;
  if (a.length < 2 || b.length < 2) return a === b ? 1 : 0;
  const bg = s => { const r = new Set(); for (let i=0; i<s.length-1; i++) r.add(s.slice(i,i+2)); return r; };
  const ba = bg(a), bb = bg(b);
  let common = 0; ba.forEach(t => { if (bb.has(t)) common++; });
  const union = ba.size + bb.size - common;
  return union > 0 ? common / union : 0;
}
// ─── 두 그룹 병합 ──────────────────────────────────────────────────────────
function mergeGroupPair(target, source) {
  const km = {};
  [...target.rows, ...source.rows].forEach(r => {
    const k = normDateStr(r.무료강의일) + '::' + r.기수;
    if (!km[k]) { km[k] = Object.assign({}, r); return; }
    const t = km[k];
    t.강의총매출 += r.강의총매출||0; t.순매출 += r.순매출||0; t.수강생수 += r.수강생수||0;
    t.무료강의신청수 = Math.max(t.무료강의신청수||0, r.무료강의신청수||0);
    t.PG제외매출 += r.PG제외매출||0; t.플랫폼매출 += r.플랫폼매출||0;
    t.광고비 += r.광고비||0; t.노바수수료 += r.노바수수료||0;
    t.PG사수수료 += r.PG사수수료||0; t.기타비용 += r.기타비용||0; t.강사정산금 += r.강사정산금||0;
  });
  target.rows = Object.values(km).sort((a,b) => (a.무료강의일||'').localeCompare(b.무료강의일||''));
  target.강의총매출 += source.강의총매출; target.순매출 += source.순매출;
  target.수강생수 += source.수강생수;
  target.무료강의신청수 = Math.max(target.무료강의신청수, source.무료강의신청수);
  target.PG제외매출 += source.PG제외매출||0; target.PG사수수료 += source.PG사수수료||0;
  target.노바수수료 += source.노바수수료||0; target.광고비 += source.광고비||0;
  target.기타비용 += source.기타비용||0; target.플랫폼수익금 += source.플랫폼수익금||0;
  target.강사정산금 += source.강사정산금||0;
  source.기수List.forEach(k => { if (!target.기수List.includes(k)) target.기수List.push(k); });
  if (!target.mergedFrom) target.mergedFrom = [];
  target.mergedFrom.push({ lecture_key: source.lecture_key, 강의그룹명: source.강의그룹명 });
  if (source.mergedFrom) target.mergedFrom.push(...source.mergedFrom);
}
// ─── 자동 그룹 병합 (강의일+기수 동일, 강의명 유사도 ≥0.5) ────────────────
function autoMergeGroups(groups, excluded) {
  excluded = excluded || new Set();
  const byDateKisu = {};
  groups.forEach(g => {
    g.rows.forEach(r => {
      const dk = normDateStr(r.무료강의일) + '::' + r.기수;
      if (!byDateKisu[dk]) byDateKisu[dk] = [];
      if (!byDateKisu[dk].find(x => x.lecture_key === g.lecture_key)) byDateKisu[dk].push(g);
    });
  });
  const gMap = {}; groups.forEach(g => gMap[g.lecture_key] = g);
  const parent = {}; groups.forEach(g => parent[g.lecture_key] = g.lecture_key);
  function find(k) { return parent[k] === k ? k : (parent[k] = find(parent[k])); }
  function union(a, b) {
    const ra = find(a), rb = find(b); if (ra === rb) return;
    const pk = [ra, rb].sort().join('|||');
    if (excluded.has(pk)) return;
    const ga = gMap[ra], gb = gMap[rb];
    if (ga && gb && ga.강의총매출 >= gb.강의총매출) parent[rb] = ra; else parent[ra] = rb;
  }
  // 강의일이 없는 버킷은 자동 병합 대상에서 제외 (날짜 없으면 동일 강의 여부 불확실)
  Object.entries(byDateKisu).forEach(([dk, gs]) => {
    if (gs.length < 2) return;
    const date = dk.split('::')[0];
    if (!date) return; // 강의일 없는 항목은 자동 병합 금지
    for (let i = 0; i < gs.length; i++)
      for (let j = i+1; j < gs.length; j++)
        if (nameSimilarity(gs[i].강의그룹명, gs[j].강의그룹명) >= 0.75) union(gs[i].lecture_key, gs[j].lecture_key);
  });
  const rootMap = {}; groups.forEach(g => {
    const root = find(g.lecture_key);
    if (!rootMap[root]) rootMap[root] = { primary: gMap[root], sources: [] };
    if (g.lecture_key !== root) rootMap[root].sources.push(g);
  });
  Object.values(rootMap).forEach(({ primary, sources }) => sources.forEach(src => mergeGroupPair(primary, src)));
  return groups.filter(g => find(g.lecture_key) === g.lecture_key);
}

const _aggregated = aggregateByKey(data);
const summary = autoMergeGroups(_aggregated);
console.log('강의 수 (집계 후):', _aggregated.length, '→ 통합 후:', summary.length);
const dataJSON = JSON.stringify(summary);

// ─── HTML 생성 ─────────────────────────────────────────────────────────────
const html = `<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Nova Analytics</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"><\/script>
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"><\/script>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
:root {
  --bg: #0B0C0B; --card: #141414; --card2: #1C1C1C; --border: #282828;
  --accent: #A3C244; --accent2: #7FA033; --green: #10b981; --yellow: #f59e0b;
  --red: #EF4444; --text: #E5E7EB; --muted: #9CA3AF; --hover: #1E1E1E; --cyan: #22d3ee;
}
body { background: var(--bg); color: var(--text); font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif; min-height: 100vh; }

/* ── 업로드 섹션 ── */
.upload-section {
  background: var(--card); border-bottom: 1px solid var(--border);
  padding: 12px 24px;
}
.upload-toggle {
  display: flex; align-items: center; gap: 8px; cursor: pointer;
  font-size: 13px; color: var(--muted); user-select: none;
}
.upload-toggle:hover { color: var(--text); }
.upload-toggle .arrow { transition: transform .2s; font-size: 10px; }
.upload-toggle.open .arrow { transform: rotate(180deg); }
.upload-body { display: none; padding-top: 14px; }
.upload-body.open { display: block; }
.upload-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px; margin-bottom: 12px; }
@media(max-width:900px) { .upload-grid { grid-template-columns: 1fr 1fr; } }
@media(max-width:600px) { .upload-grid { grid-template-columns: 1fr; } }
.upload-card {
  border: 1.5px dashed var(--border); border-radius: 10px; padding: 14px 16px;
  cursor: pointer; transition: border-color .2s; position: relative;
}
.upload-card:hover { border-color: var(--accent); }
.upload-card.has-file { border-color: var(--green); border-style: solid; }
.upload-card-label { font-size: 11px; color: var(--muted); margin-bottom: 4px; }
.upload-card-name { font-size: 13px; font-weight: 600; }
.upload-card-name.empty { color: var(--muted); font-weight: 400; }
.upload-card-name.filled { color: var(--green); }
.upload-card-badge {
  display: inline-block; padding: 1px 7px; border-radius: 4px; font-size: 10px;
  background: rgba(16,185,129,0.15); color: var(--green); margin-left: 6px;
}
.upload-input { display: none; }
.upload-actions { display: flex; align-items: center; gap: 10px; }
.apply-btn {
  padding: 8px 20px; background: var(--accent); color: #0D1500; border: none;
  border-radius: 8px; font-size: 13px; font-weight: 700; cursor: pointer; transition: opacity .2s;
}
.apply-btn:hover { opacity: 0.85; }
.apply-btn:disabled { opacity: 0.4; cursor: not-allowed; }
.upload-status { font-size: 12px; color: var(--muted); }

/* ── 안내 배너 ── */
.top-banner {
  display: flex; align-items: center; flex-wrap: wrap; gap: 8px;
  padding: 11px 24px; font-size: 13px; font-weight: 500; color: #fff;
  position: relative; z-index: 210;
}
.top-banner.banner-warn {
  background: linear-gradient(90deg, #7f1d1d 0%, #991b1b 50%, #7f1d1d 100%);
  border-bottom: 1px solid rgba(239,68,68,0.5);
}
.top-banner.banner-ok {
  background: linear-gradient(90deg, #064e3b 0%, #065f46 50%, #064e3b 100%);
  border-bottom: 1px solid rgba(16,185,129,0.4);
}
.top-banner.banner-none {
  background: linear-gradient(90deg, #1c1917 0%, #292524 50%, #1c1917 100%);
  border-bottom: 1px solid rgba(120,113,108,0.4);
}
.banner-title { font-weight: 700; white-space: nowrap; font-size: 13px; letter-spacing: .01em; }
.banner-div { opacity: .45; margin: 0 4px; font-size: 11px; }
.banner-item { display: flex; align-items: center; gap: 5px; white-space: nowrap; font-size: 12px; }
.banner-val { font-weight: 700; }
.banner-status-ok   { background: rgba(16,185,129,0.25); border: 1px solid rgba(16,185,129,0.5); color: #6ee7b7; border-radius: 5px; padding: 1px 7px; font-size: 11px; font-weight: 700; }
.banner-status-warn { background: rgba(239,68,68,0.25); border: 1px solid rgba(239,68,68,0.5); color: #fca5a5; border-radius: 5px; padding: 1px 7px; font-size: 11px; font-weight: 700; }
.banner-status-none { background: rgba(120,113,108,0.25); border: 1px solid rgba(120,113,108,0.4); color: #d6d3d1; border-radius: 5px; padding: 1px 7px; font-size: 11px; font-weight: 700; }

/* ── sticky 상단 고정 영역 ── */
.sticky-top { position: sticky; top: 0; z-index: 200; }

/* ── 헤더 ── */
.header {
  background: var(--card); border-bottom: 1px solid var(--border);
  padding: 14px 24px; display: flex; align-items: center; gap: 16px;
  flex-wrap: wrap;
}
.logo { font-size: 16px; font-weight: 800; white-space: nowrap; }
.logo span { color: var(--accent); }
.nav-tabs { display: flex; gap: 4px; }
.nav-tab {
  padding: 7px 16px; border-radius: 8px; font-size: 13px; cursor: pointer;
  color: var(--muted); border: 1px solid transparent; transition: all .2s; background: transparent; white-space: nowrap;
}
.nav-tab.active { background: var(--accent); color: #0D1500; border-color: var(--accent); font-weight: 700; }
.nav-tab:hover:not(.active) { background: var(--hover); color: var(--text); }
.filter-wrap { display: flex; align-items: center; gap: 10px; margin-left: auto; }
.filter-label { font-size: 12px; color: var(--muted); white-space: nowrap; }
.filter-wrap select {
  background: var(--card2); border: 1px solid var(--border); color: var(--text);
  padding: 7px 32px 7px 12px; border-radius: 8px; font-size: 13px; cursor: pointer;
  outline: none; max-width: 260px; appearance: none;
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%2394a3b8' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E");
  background-repeat: no-repeat; background-position: right 10px center;
}
.filter-wrap select:focus { border-color: var(--accent); }

/* ── 기간 필터 바 ── */
.date-filter-bar {
  background: var(--card); border-bottom: 1px solid var(--border);
  padding: 8px 24px; display: flex; align-items: center; gap: 12px; flex-wrap: wrap;
}
.date-filter-label { font-size: 12px; color: var(--muted); white-space: nowrap; }
.date-input {
  background: var(--card2); border: 1px solid var(--border); color: var(--text);
  padding: 5px 10px; border-radius: 7px; font-size: 12px; outline: none;
  color-scheme: dark;
}
.date-input:focus { border-color: var(--accent); }
.date-sep { font-size: 12px; color: var(--muted); }
.date-reset-btn {
  padding: 4px 12px; border-radius: 6px; font-size: 12px; cursor: pointer;
  border: 1px solid var(--border); background: transparent; color: var(--muted); transition: all .2s;
}
.date-reset-btn:hover { border-color: var(--accent); color: var(--accent); }
.date-result { font-size: 12px; color: var(--muted); margin-left: 4px; }

/* ── 메인 레이아웃 ── */
.main { padding: 24px; }
.page { display: none; }
.page.active { display: block; }

/* ── 리스트 페이지 ── */
.page-top { margin-bottom: 12px; }
.page-title { font-size: 20px; font-weight: 800; }
.page-sub { font-size: 13px; color: var(--muted); margin-top: 3px; }
/* ── 통합 필터 패널 ── */
.list-filter-panel {
  display: flex; justify-content: space-between; align-items: center;
  flex-wrap: wrap; gap: 10px;
  background: var(--card);
  border: 1px solid rgba(255,255,255,0.07);
  border-radius: 14px; padding: 14px 18px; margin-bottom: 14px;
  box-shadow: 0 2px 12px rgba(0,0,0,0.28);
}
.lf-left  { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; }
.lf-right { display: flex; align-items: center; }
.lf-sep   { width: 1px; height: 20px; background: var(--border); margin: 0 8px; flex-shrink: 0; opacity: .7; }
.lf-group { display: flex; align-items: center; gap: 6px; }
.lf-label {
  font-size: 10px; color: var(--muted); white-space: nowrap;
  font-weight: 700; letter-spacing: .06em; text-transform: uppercase;
}
.lf-select {
  background: var(--card2); border: 1px solid var(--border); color: var(--text);
  padding: 6px 10px; border-radius: 8px; font-size: 12px; cursor: pointer; outline: none; max-width: 200px;
  transition: border-color .2s, background .2s, box-shadow .2s;
  color-scheme: dark;
}
.lf-select:hover { border-color: #404040; }
.lf-select:focus { border-color: var(--accent); box-shadow: 0 0 0 2px rgba(163,194,68,0.18); }
.lf-select.lf-active {
  background: rgba(163,194,68,0.12); border-color: var(--accent);
  color: var(--text); box-shadow: 0 0 0 2px rgba(163,194,68,0.15);
}
.lf-date-input {
  background: var(--card2); border: 1px solid var(--border); color: var(--text);
  padding: 6px 8px; border-radius: 8px; font-size: 12px; outline: none; color-scheme: dark;
  transition: border-color .2s, background .2s, box-shadow .2s;
}
.lf-date-input:hover { border-color: #404040; }
.lf-date-input:focus { border-color: var(--accent); box-shadow: 0 0 0 2px rgba(163,194,68,0.18); }
.lf-date-input.lf-active {
  background: rgba(163,194,68,0.1); border-color: var(--accent);
  box-shadow: 0 0 0 2px rgba(163,194,68,0.15);
}
.lf-date-sep { font-size: 12px; color: var(--muted); }
.lf-reset-btn {
  padding: 5px 10px; border-radius: 7px; font-size: 11px; cursor: pointer;
  border: 1px solid var(--border); background: var(--card2); color: var(--muted);
  transition: all .2s; white-space: nowrap;
}
.lf-reset-btn:hover { border-color: var(--accent); color: var(--accent); background: rgba(163,194,68,0.08); }
.lf-date-result { font-size: 11px; color: var(--muted); font-weight: 600; }
.search-input {
  background: var(--card2); border: 1.5px solid var(--border); color: var(--text);
  padding: 8px 14px; border-radius: 9px; font-size: 13px; outline: none; width: 230px;
  transition: border-color .2s, box-shadow .2s, background .2s;
}
.search-input:hover { border-color: #404040; }
.search-input:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(163,194,68,0.15); background: var(--card2); }
.sort-bar { display: flex; gap: 6px; margin-bottom: 12px; flex-wrap: wrap; align-items: center; }
.sort-label { font-size: 12px; color: var(--muted); }
.sort-btn { padding: 5px 12px; border-radius: 6px; font-size: 12px; cursor: pointer; color: var(--muted); border: 1px solid var(--border); background: var(--card); transition: all .2s; }
.sort-btn.active { color: var(--accent); border-color: var(--accent); }
.table-wrap { overflow-x: auto; border-radius: 12px; border: 1px solid var(--border); }
table { width: 100%; border-collapse: collapse; min-width: 680px; }
thead { background: var(--card2); }
th { padding: 11px 14px; text-align: left; font-size: 11px; color: var(--muted); font-weight: 600; white-space: nowrap; }
th.num { text-align: right; }
tbody tr { border-top: 1px solid var(--border); cursor: pointer; transition: background .15s; }
tbody tr:hover { background: var(--hover); }
td { padding: 12px 14px; font-size: 13px; }
td.num { text-align: right; font-variant-numeric: tabular-nums; }
td.rank { color: var(--muted); font-size: 12px; }
.lec-name { font-weight: 600; max-width: 300px; line-height: 1.3; }
.lec-name small { font-size: 11px; color: var(--muted); font-weight: 400; display: block; margin-top: 2px; }
.badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; background: var(--card2); color: var(--muted); border: 1px solid var(--border); }
.kisu-badge { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 11px; background: rgba(255,255,255,0.06); color: var(--accent); border: 1px solid rgba(255,255,255,0.12); }
.merge-badge { display: inline-block; padding: 2px 7px; border-radius: 4px; font-size: 10px; background: rgba(251,191,36,0.12); color: #fbbf24; border: 1px solid rgba(251,191,36,0.3); margin-left: 5px; vertical-align: middle; cursor: pointer; }
.merge-panel { background: rgba(251,191,36,0.05); border: 1px solid rgba(251,191,36,0.2); border-radius: 10px; padding: 12px 16px; margin-bottom: 14px; }
.merge-panel-title { font-size: 12px; color: #fbbf24; font-weight: 600; margin-bottom: 8px; }
.merge-item { display: flex; align-items: center; justify-content: space-between; padding: 5px 0; border-bottom: 1px solid rgba(255,255,255,0.04); font-size: 12px; color: var(--muted); }
.merge-item:last-child { border-bottom: none; }
.merge-item-name { flex: 1; }
.merge-item-tag { font-size: 10px; padding: 1px 6px; border-radius: 4px; margin-left: 6px; }
.merge-item-tag.primary { background: rgba(255,255,255,0.07); color: var(--accent); }
.merge-item-tag.merged { background: rgba(251,191,36,0.12); color: #fbbf24; }
.merge-unlink-btn { font-size: 10px; padding: 2px 8px; border-radius: 4px; background: rgba(239,68,68,0.1); color: #ef4444; border: 1px solid rgba(239,68,68,0.2); cursor: pointer; margin-left: 8px; white-space: nowrap; }
.merge-unlink-btn:hover { background: rgba(239,68,68,0.2); }
.go-btn { font-size: 11px; color: var(--accent); }
.pagination { display: flex; justify-content: center; gap: 5px; margin-top: 20px; flex-wrap: wrap; }
.pg-btn { padding: 6px 11px; border-radius: 6px; font-size: 12px; cursor: pointer; border: 1px solid var(--border); background: var(--card); color: var(--muted); transition: all .2s; }
.pg-btn.active { background: var(--accent); color: #0D1500; border-color: var(--accent); font-weight: 700; }
.pg-btn:hover:not(.active) { background: var(--hover); color: var(--text); }

/* ── 상세 페이지 ── */
.back-btn { display: inline-flex; align-items: center; gap: 6px; color: var(--muted); font-size: 13px; cursor: pointer; padding: 6px 0; margin-bottom: 16px; transition: color .2s; }
.back-btn:hover { color: var(--text); }
.detail-title { font-size: 20px; font-weight: 800; margin-bottom: 4px; line-height: 1.4; }
.detail-meta { font-size: 13px; color: var(--muted); margin-bottom: 22px; }
.detail-meta span { margin-right: 14px; }
.detail-sticky-header {
  position: sticky;
  z-index: 150;
  background: var(--bg);
  padding-bottom: 10px;
  border-bottom: 1px solid var(--border);
  margin-bottom: 14px;
}
.detail-sticky-header .back-btn { margin-bottom: 8px; }
.detail-sticky-header .detail-meta { margin-bottom: 0; }

/* ── KPI ── */
.kpi-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 12px; margin-bottom: 14px; }
.kpi-card { background: var(--card); border: 1px solid var(--border); border-radius: 16px; padding: 20px 22px; }
.kpi-label { font-size: 11px; color: var(--muted); margin-bottom: 10px; letter-spacing: .03em; }
.kpi-value { font-size: 28px; font-weight: 800; line-height: 1; letter-spacing: -.02em; }
.kpi-unit { font-size: 14px; color: var(--muted); margin-left: 2px; font-weight: 500; }
.kpi-sub { font-size: 11px; color: var(--muted); margin-top: 6px; }
.kpi-card.green .kpi-value { color: var(--green); }
.kpi-card.yellow .kpi-value { color: var(--yellow); }
.kpi-card.purple .kpi-value { color: var(--accent2); }
.kpi-card.cyan .kpi-value { color: var(--cyan); }
/* ── 핵심지표 패널 레이아웃 ── */
.kpi-panel-wrap { margin-bottom: 18px; }
.kpi-panel-title { font-size: 13px; font-weight: 700; color: var(--text); letter-spacing: .02em; margin-bottom: 10px; }
.kpi-panel { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
.kpi-group { display: flex; flex-direction: column; gap: 8px; min-width: 0; border: 1px solid rgba(163,194,68,0.28); border-radius: 14px; padding: 12px; }
.kpi-group-label { font-size: 10px; color: rgba(163,194,68,0.75); letter-spacing: .08em; text-transform: uppercase; font-weight: 600; padding-left: 2px; }
.kpi-grid22 { display: grid; grid-template-columns: repeat(2,1fr); gap: 6px; align-items: stretch; flex: 1; }
.kpi-grid32 { display: grid; grid-template-columns: repeat(3,1fr); gap: 6px; align-items: stretch; flex: 1; }
.kpi-grid12 { display: grid; grid-template-columns: 1fr; gap: 6px; flex: 1; }
.kpi-sm { background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 12px 14px; min-width: 0; box-sizing: border-box; }
.kpi-sm .ks-label { font-size: 10px; color: var(--muted); margin-bottom: 5px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.kpi-sm .ks-val { font-size: 22px; font-weight: 800; letter-spacing: -.02em; line-height: 1.1; color: var(--accent); word-break: break-all; }
.kpi-sm .ks-val-full { font-size: 16px; font-weight: 700; letter-spacing: -.01em; line-height: 1.2; color: var(--text); word-break: break-all; }
.kpi-sm .ks-sub { font-size: 10px; color: var(--muted); margin-top: 3px; }
.ks-red { color: #F87171 !important; }
.ks-accent { color: var(--accent) !important; }
@media(max-width:960px) { .kpi-panel { grid-template-columns: 1fr 1fr; } .kpi-panel > .kpi-group:last-child { grid-column: 1 / -1; } .kpi-grid12 { grid-template-columns: repeat(2,1fr); } }
.kpi-row5 { grid-template-columns: repeat(5, 1fr); }
.kpi-row6 { grid-template-columns: repeat(6, 1fr); }
.kpi-card-conv { padding: 16px 18px; }
.kpi-conv-val { font-size: 22px; color: var(--accent); }
.kpi-card.kpi-red .kpi-conv-val { color: #F87171; }
.kpi-card.kpi-green .kpi-conv-val { color: #16A34A; }
@media(max-width:900px) { .kpi-row5 { grid-template-columns: repeat(3, 1fr); } .kpi-row6 { grid-template-columns: repeat(3, 1fr); } }
@media(max-width:600px) { .kpi-row5 { grid-template-columns: repeat(2, 1fr); } .kpi-row6 { grid-template-columns: repeat(2, 1fr); } }

/* ── 영업이익 카드 ── */
.profit-card { background: var(--card); border: 1px solid var(--border); border-radius: 16px; padding: 20px 22px; }
.profit-title { font-size: 12px; color: var(--muted); margin-bottom: 12px; display: flex; align-items: center; gap: 6px; }
.profit-title::before { content:''; display:inline-block; width:8px; height:8px; border-radius:50%; background:var(--green); }
.profit-body { display: flex; justify-content: space-between; align-items: flex-end; margin-bottom: 14px; }
.profit-main-val { font-size: 22px; font-weight: 800; }
.profit-rate { font-size: 28px; font-weight: 800; color: var(--green); }
.profit-rate span { font-size: 14px; color: var(--muted); font-weight: 500; }
.profit-bar-wrap { background: rgba(255,255,255,0.06); border-radius: 99px; height: 8px; overflow: hidden; }
.profit-bar { background: linear-gradient(90deg, #10b981, #34d399); border-radius: 99px; height: 100%; transition: width .8s ease; }
.profit-bar-label { font-size: 11px; color: var(--green); margin-top: 6px; }

/* ── 광고 성과 ── */
.ad-section { background: var(--card); border: 1px solid var(--border); border-radius: 16px; padding: 20px 22px; }
.ad-title { font-size: 12px; color: var(--muted); margin-bottom: 14px; display: flex; align-items: center; gap: 6px; }
.ad-title::before { content:''; display:inline-block; width:8px; height:8px; border-radius:50%; background:var(--accent); }
.ad-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(130px, 1fr)); gap: 16px; }
.ad-item-label { font-size: 11px; color: var(--muted); margin-bottom: 6px; }
.ad-item-value { font-size: 22px; font-weight: 800; color: var(--yellow); }
.ad-item-unit { font-size: 12px; color: var(--muted); font-weight: 500; margin-left: 2px; }
.ad-item-sub { font-size: 11px; color: var(--muted); margin-top: 4px; }

/* ── 차트 공통 ── */
.chart-card { background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 18px 20px; }
.chart-title { font-size: 14px; font-weight: 700; margin-bottom: 3px; }
.chart-sub { font-size: 11px; color: var(--muted); margin-bottom: 14px; }
.chart-wrap { position: relative; height: 220px; }
.chart-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 3px; }
.interval-select {
  background: var(--card2); border: 1px solid var(--border); color: var(--text);
  padding: 4px 24px 4px 10px; border-radius: 6px; font-size: 12px; cursor: pointer; color-scheme: dark;
  outline: none; appearance: none;
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='10' viewBox='0 0 24 24' fill='none' stroke='%2394a3b8' stroke-width='2'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E");
  background-repeat: no-repeat; background-position: right 8px center;
}

/* ── 기수별 테이블 ── */
.kisu-table { width: 100%; border-collapse: collapse; font-size: 13px; }
.kisu-table th { padding: 8px 12px; text-align: left; color: var(--muted); font-size: 11px; border-bottom: 1px solid var(--border); white-space: nowrap; }
.kisu-table th.num { text-align: right; }
.kisu-table td { padding: 10px 12px; border-bottom: 1px solid var(--border); }
.kisu-table td.num { text-align: right; font-variant-numeric: tabular-nums; }
.kisu-table tr:last-child td { border-bottom: none; }

/* ── 구매 상세 토글 ── */
.toggle-btn {
  display: flex; align-items: center; gap: 6px; cursor: pointer;
  font-size: 13px; color: var(--muted); padding: 10px 0; user-select: none; transition: color .2s;
}
.toggle-btn:hover { color: var(--text); }
.toggle-btn .arr { font-size: 10px; transition: transform .2s; }
.toggle-btn.open .arr { transform: rotate(180deg); }
.order-list-wrap { display: none; margin-top: 10px; }
.order-list-wrap.open { display: block; }
.order-list { max-height: 300px; overflow-y: auto; }
.order-row { display: flex; gap: 16px; padding: 7px 12px; border-bottom: 1px solid var(--border); font-size: 13px; font-variant-numeric: tabular-nums; }
.order-row:hover { background: var(--hover); }
.order-time { color: var(--muted); min-width: 120px; }
.order-elapsed { color: var(--accent); min-width: 80px; font-size: 12px; font-variant-numeric: tabular-nums; }
.order-amount { color: var(--text); font-weight: 600; min-width: 90px; }
.order-buyer { color: var(--text); min-width: 70px; font-size: 12px; }
.order-phone { color: var(--muted); min-width: 110px; font-size: 12px; }
.order-name { color: var(--muted); flex: 1; font-size: 11px; }
.no-order { padding: 20px; text-align: center; color: var(--muted); font-size: 13px; }

/* ── 이익+광고 2열 그리드 ── */
.profit-ad-row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 14px; }
@media(max-width:768px) { .profit-ad-row { grid-template-columns: 1fr; } }

@media(max-width:600px) {
  .header { padding: 12px 16px; }
  .main { padding: 14px; }
  .filter-wrap { width: 100%; }
  .filter-wrap select { max-width: 100%; width: 100%; }
  .upload-section { padding: 12px 16px; }
}

/* ── 성과 퍼널 ── */
.funnel-row {
  display: grid;
  grid-template-columns: 110px 1fr 100px 110px 1fr;
  gap: 10px; align-items: center;
  padding: 10px 0; border-bottom: 1px solid var(--border);
}
.funnel-row:last-child { border-bottom: none; }
.funnel-stage { font-size: 12px; color: var(--muted); white-space: nowrap; }
.funnel-bar-wrap { background: rgba(255,255,255,0.05); border-radius: 4px; height: 6px; }
.funnel-bar { height: 100%; border-radius: 4px; background: linear-gradient(90deg,#A3C244,#7FA033); transition: width .5s ease; min-width: 2px; }
.funnel-val-cell { display: flex; align-items: center; gap: 4px; justify-content: flex-end; }
.funnel-inp { background: transparent; border: 1px solid transparent; border-radius: 4px; padding: 3px 6px; color: var(--text); font-size: 13px; font-weight: 600; width: 70px; text-align: right; font-variant-numeric: tabular-nums; }
.funnel-inp:focus { border-color: var(--accent); outline: none; background: var(--card2); }
.funnel-conv-wrap { display: flex; align-items: center; gap: 6px; font-size: 12px; }
.funnel-conv { font-variant-numeric: tabular-nums; white-space: nowrap; }
.funnel-conv.above { color: var(--green); }
.funnel-conv.below { color: var(--red); }
.funnel-conv.neutral { color: var(--muted); }
.funnel-tgt-inp { background: transparent; border: 1px solid transparent; border-radius: 4px; padding: 2px 4px; color: var(--muted); font-size: 11px; width: 36px; text-align: right; }
.funnel-tgt-inp:focus { border-color: var(--border); outline: none; background: var(--card2); }
.funnel-aar-inp { background: transparent; border: 1px solid transparent; border-radius: 4px; padding: 3px 6px; font-size: 11px; width: 100%; min-width: 0; }
.funnel-aar-inp:focus { border-color: var(--accent); outline: none; background: var(--card2); }
.funnel-aar-inp.warn { color: var(--yellow); }
.funnel-aar-inp.ok { color: var(--muted); }
@media(max-width:768px) {
  .funnel-row { grid-template-columns: 90px 1fr 80px 90px 1fr; }
}

/* ── 핵심 지표 영역 ── */
.km-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin-bottom: 16px; }
.km-card { background: var(--card2); border: 1px solid var(--border); border-radius: 8px; padding: 14px 16px; text-align: center; }
.km-label { font-size: 11px; color: var(--muted); margin-bottom: 6px; }
.km-value { font-size: 20px; font-weight: 800; color: var(--text); }
.km-value .km-unit { font-size: 12px; color: var(--muted); font-weight: 400; margin-left: 2px; }
.km-sub { font-size: 11px; color: var(--muted); margin-top: 4px; }
.km-achieve { color: var(--green); }
.km-achieve.under { color: var(--red); }
.km-divider { border-top: 1px solid var(--border); margin: 4px 0 14px; }
.km-conv-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; }
.km-conv-card { background: var(--card2); border: 1px solid var(--border); border-radius: 8px; padding: 12px 14px; text-align: center; }
.km-conv-label { font-size: 11px; color: var(--muted); margin-bottom: 4px; }
.km-conv-value { font-size: 18px; font-weight: 700; color: var(--accent); }
.km-conv-value .km-unit { font-size: 11px; color: var(--muted); font-weight: 400; }
.km-conv-sub { font-size: 10px; color: var(--muted); margin-top: 4px; }
@media(max-width:768px) {
  .km-grid { grid-template-columns: repeat(2, 1fr); }
  .km-conv-grid { grid-template-columns: repeat(2, 1fr); }
}

/* ── 드롭다운 옵션 다크 스타일 ── */
select { color-scheme: dark; }
select option {
  background-color: #131510;
  color: #E2EDD0;
}
select option:checked {
  background-color: #A3C244;
  color: #0D1500;
}
select option:hover {
  background-color: #1C2518;
  color: #E2EDD0;
}
select optgroup {
  background-color: #0D0F0B;
  color: #7A9070;
}

/* ── 홈 진입 화면 ── */
.home-container { max-width: 560px; margin: 36px auto 0; padding: 0 24px; }
.home-welcome { text-align: center; margin-bottom: 28px; }
.home-welcome-title { font-size: 24px; font-weight: 800; letter-spacing: -0.03em; margin-bottom: 10px; }
.home-welcome-title span { color: var(--accent); font-weight: 900; }
.home-welcome-sub { font-size: 13px; color: var(--muted); }
.home-cards { display: flex; flex-direction: column; gap: 10px; }
.home-entry-card {
  background: var(--card); border: 1px solid var(--border); border-radius: 12px;
  padding: 20px 24px; cursor: pointer; transition: all .2s;
  display: flex; flex-direction: row; align-items: center; justify-content: space-between;
  text-align: left;
}
.home-entry-card:hover {
  border-color: var(--accent); background: var(--card2);
  transform: translateY(-2px); box-shadow: 0 8px 24px rgba(163,194,68,0.12);
}
.hcard-body { display: flex; flex-direction: column; gap: 4px; }
.hcard-title { font-size: 17px; font-weight: 700; }
.hcard-desc { font-size: 12px; color: var(--muted); line-height: 1.5; }
.hcard-arrow { font-size: 18px; color: var(--muted); transition: color .2s, transform .2s; flex-shrink: 0; margin-left: 16px; }
.home-entry-card:hover .hcard-arrow { color: var(--accent); transform: translateX(3px); }
/* 플랫폼 선택 서브뷰 */
.home-main-view.hidden { display: none; }
.home-platform-view { display: none; }
.home-platform-view.active { display: block; }
.home-back-btn {
  display: inline-flex; align-items: center; gap: 6px; margin-bottom: 22px;
  font-size: 13px; color: var(--muted); cursor: pointer; background: none; border: none;
  padding: 0; transition: color .2s;
}
.home-back-btn:hover { color: var(--text); }
.home-platform-title { font-size: 17px; font-weight: 700; margin-bottom: 6px; }
.home-platform-sub { font-size: 13px; color: var(--muted); margin-bottom: 24px; }
.platform-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(150px, 1fr)); gap: 12px; }
.platform-card {
  background: var(--card); border: 1px solid var(--border); border-radius: 12px;
  padding: 22px 16px; cursor: pointer; transition: all .2s; text-align: center;
}
.platform-card:hover {
  border-color: var(--accent); background: var(--card2);
  transform: translateY(-2px); box-shadow: 0 8px 20px rgba(163,194,68,0.12);
}
.platform-card-name { font-size: 14px; font-weight: 700; margin-bottom: 5px; }
.platform-card-count { font-size: 11px; color: var(--muted); }

/* ── 로그인 화면 ── */
#login-screen {
  display: none; position: fixed; inset: 0; z-index: 9999;
  background: var(--bg); align-items: center; justify-content: center;
}
#login-screen.active { display: flex; }
.login-card {
  background: var(--card); border: 1px solid var(--border); border-radius: 16px;
  padding: 44px 40px 36px; width: 360px; max-width: 90vw;
  box-shadow: 0 24px 60px rgba(0,0,0,0.5);
}
.login-logo {
  text-align: center; margin-bottom: 28px;
  font-size: 22px; font-weight: 800; letter-spacing: -0.03em; color: var(--text);
}
.login-logo span { color: var(--accent); }
.login-subtitle {
  text-align: center; font-size: 13px; color: var(--muted); margin-bottom: 28px;
}
.login-label { font-size: 12px; color: var(--muted); margin-bottom: 7px; }
.login-input {
  width: 100%; padding: 11px 14px; background: var(--card2); border: 1px solid var(--border);
  border-radius: 8px; color: var(--text); font-size: 14px; outline: none;
  transition: border-color .2s;
}
.login-input:focus { border-color: var(--accent); }
.login-input.error { border-color: var(--red); }
.login-error { font-size: 12px; color: var(--red); margin-top: 8px; min-height: 16px; }
.login-btn {
  width: 100%; margin-top: 20px; padding: 12px; background: var(--accent);
  color: #0D1500; border: none; border-radius: 8px; font-size: 14px; font-weight: 800;
  cursor: pointer; transition: opacity .2s; letter-spacing: 0.02em;
}
.login-btn:hover { opacity: 0.88; }

/* ── 비밀번호 변경 모달 ── */
#pw-modal-overlay {
  display: none; position: fixed; inset: 0; z-index: 8888;
  background: rgba(0,0,0,0.65); align-items: center; justify-content: center;
}
#pw-modal-overlay.active { display: flex; }
.pw-modal {
  background: var(--card); border: 1px solid var(--border); border-radius: 14px;
  padding: 32px 32px 28px; width: 360px; max-width: 92vw;
  box-shadow: 0 16px 48px rgba(0,0,0,0.5);
}
.pw-modal-title { font-size: 16px; font-weight: 700; margin-bottom: 22px; }
.pw-field { margin-bottom: 14px; }
.pw-field-label { font-size: 11px; color: var(--muted); margin-bottom: 6px; }
.pw-field-input {
  width: 100%; padding: 10px 12px; background: var(--card2); border: 1px solid var(--border);
  border-radius: 7px; color: var(--text); font-size: 13px; outline: none; transition: border-color .2s;
}
.pw-field-input:focus { border-color: var(--accent); }
.pw-msg { font-size: 12px; min-height: 16px; margin-top: 8px; }
.pw-msg.err { color: var(--red); }
.pw-msg.ok  { color: var(--green); }
.pw-modal-actions { display: flex; gap: 10px; margin-top: 20px; }
.pw-save-btn {
  flex: 1; padding: 10px; background: var(--accent); color: #0D1500; border: none;
  border-radius: 8px; font-size: 13px; font-weight: 700; cursor: pointer; transition: opacity .2s;
}
.pw-save-btn:hover { opacity: 0.88; }
.pw-cancel-btn {
  flex: 1; padding: 10px; background: transparent; color: var(--muted);
  border: 1px solid var(--border); border-radius: 8px; font-size: 13px; cursor: pointer; transition: border-color .2s;
}
.pw-cancel-btn:hover { border-color: var(--accent); color: var(--text); }

/* ── 푸터 ── */
.page-footer {
  text-align: center; padding: 24px 0 16px;
  font-size: 12px; color: #9CA3AF; opacity: 0.6; letter-spacing: 0.02em;
}
/* 로그인 화면 내부 푸터: overlay 하단에 고정 */
#login-screen .page-footer {
  position: absolute; bottom: 16px; left: 0; right: 0;
  padding: 0;
}

/* ── 헤더 인증 버튼 ── */
.header-auth { display: flex; align-items: center; gap: 8px; margin-left: auto; }
.auth-btn {
  padding: 6px 14px; font-size: 12px; border-radius: 7px; cursor: pointer;
  background: transparent; color: var(--muted); border: 1px solid var(--border);
  transition: all .2s; white-space: nowrap;
}
.auth-btn:hover { color: var(--text); border-color: var(--accent); }
.auth-btn.logout { color: var(--red); border-color: transparent; }
.auth-btn.logout:hover { border-color: var(--red); background: rgba(239,68,68,0.08); }
</style>
</head>
<body>

<!-- ══ 로그인 화면 ══ -->
<div id="login-screen">
  <div class="login-card">
    <div class="login-logo"><span>Nova</span> Analytics</div>
    <div class="login-subtitle">System Nova 관리자 전용 페이지입니다.</div>
    <input class="login-input" type="password" id="login-pw-input" placeholder="비밀번호 입력"
      onkeydown="if(event.key==='Enter')doLogin()">
    <div class="login-error" id="login-error"></div>
    <button class="login-btn" onclick="doLogin()">로그인</button>
  </div>
  <div class="page-footer">© 2026 염정하</div>
</div>


<!-- ══ 앱 래퍼 (로그인 후 표시) ══ -->
<div id="app-wrapper" style="display:none;">

<!-- ══ sticky 상단 영역 ══ -->
<div class="sticky-top">

<!-- ══ 안내 배너 ══ -->
<div id="top-banner" class="top-banner banner-none">
  <span class="banner-item">현재 반영된 최신 무료강의일: <span class="banner-val" id="banner-lecture-date">-</span> <span id="banner-status" class="banner-status-none">데이터 없음</span></span>
</div>

<!-- ══ 업로드 섹션 ══ -->
<div class="upload-section">
  <div class="upload-toggle" id="upload-toggle" onclick="toggleUpload()">
    <span>📂</span>
    <span>데이터 업로드</span>
    <span class="arrow" id="upload-arrow">▼</span>
    <span id="upload-badge" style="margin-left:4px;"></span>
  </div>
  <div class="upload-body" id="upload-body">
    <div class="upload-grid">
      <div class="upload-card" id="card1" onclick="document.getElementById('file1').click()">
        <input class="upload-input" type="file" id="file1" accept=".xlsx,.xls" onchange="handleFile(1,this)">
        <div class="upload-card-label">ERP 매출관리 엑셀파일</div>
        <div class="upload-card-name empty" id="fname1">파일을 선택하거나 드래그하세요</div>
      </div>
      <div class="upload-card" id="card2" onclick="document.getElementById('file2').click()">
        <input class="upload-input" type="file" id="file2" accept=".xlsx,.xls" onchange="handleFile(2,this)">
        <div class="upload-card-label">ERP 주문결제관리 엑셀파일</div>
        <div class="upload-card-name empty" id="fname2">파일을 선택하거나 드래그하세요</div>
      </div>
      <div class="upload-card" id="card3" onclick="document.getElementById('file3').click()">
        <input class="upload-input" type="file" id="file3" accept=".xlsx,.xls,.csv" onchange="handleFile(3,this)">
        <div class="upload-card-label">노바 강의일정 시트</div>
        <div class="upload-card-name empty" id="fname3">파일을 선택하거나 드래그하세요</div>
      </div>
    </div>
    <div class="upload-actions">
      <button class="apply-btn" id="apply-btn" onclick="applyUpload()" disabled>데이터 적용</button>
      <span class="upload-status" id="upload-status"></span>
    </div>
  </div>
</div>

<!-- ══ 헤더 ══ -->
<div class="header">
  <div class="logo" onclick="resetAll()" style="cursor:pointer;"><span>Nova</span> Analytics</div>
  <div class="nav-tabs">
    <div class="nav-tab active" id="tab-home" onclick="showPage('home')">홈</div>
    <div class="nav-tab" id="tab-list" onclick="goBackToList()">강의 리스트</div>
    <div class="nav-tab" id="tab-detail" onclick="showPage('detail')">강의 상세 분석</div>
  </div>
  <div class="header-auth">
    <button class="auth-btn logout" onclick="doLogout()">로그아웃</button>
  </div>
</div>

</div><!-- /sticky-top -->

<!-- ══ 메인 ══ -->
<div class="main">

  <!-- 홈 진입 화면 -->
  <div class="page active" id="page-home">
    <div class="home-container">
      <div class="home-welcome">
        <div class="home-welcome-title"><span>Nova</span> Analytics</div>
        <div class="home-welcome-sub">분석을 시작하려면 아래 메뉴를 선택하세요.</div>
      </div>
      <div id="home-main-view" class="home-main-view">
        <div class="home-cards">
          <div class="home-entry-card" onclick="showPage('list')">
            <div class="hcard-body">
              <div class="hcard-title">전체 강의 리스트 보기</div>
              <div class="hcard-desc">매출, 수강생, 신청수 등 통합 데이터를 확인합니다.</div>
            </div>
            <div class="hcard-arrow">›</div>
          </div>
          <div class="home-entry-card" onclick="showPlatformPicker()">
            <div class="hcard-body">
              <div class="hcard-title">플랫폼 선택하기</div>
              <div class="hcard-desc">플랫폼별로 강의를 분류해 데이터를 분석합니다.</div>
            </div>
            <div class="hcard-arrow">›</div>
          </div>
        </div>
      </div>
      <div id="home-platform-view" class="home-platform-view">
        <button class="home-back-btn" onclick="hidePlatformPicker()">← 홈으로</button>
        <div class="home-platform-title">플랫폼 선택</div>
        <div class="home-platform-sub">분석할 플랫폼을 선택하세요.</div>
        <div class="platform-grid" id="platform-grid"></div>
      </div>
    </div>
  </div>

  <!-- 강의 리스트 -->
  <div class="page" id="page-list">
    <div class="page-top">
      <div class="page-title">강의 리스트</div>
      <div class="page-sub" id="list-count"></div>
    </div>
    <!-- 통합 필터 패널 -->
    <div class="list-filter-panel">
      <div class="lf-left">
        <div class="lf-group">
          <span class="lf-label">기간</span>
          <input class="lf-date-input" type="date" id="date-from" onchange="onDateFilter();updateLfActive()">
          <span class="lf-date-sep">~</span>
          <input class="lf-date-input" type="date" id="date-to" onchange="onDateFilter();updateLfActive()">
          <button class="lf-reset-btn" onclick="resetDateFilter();updateLfActive()">초기화</button>
          <span class="lf-date-result" id="date-result"></span>
        </div>
        <div class="lf-sep"></div>
        <div class="lf-group">
          <span class="lf-label">강의 선택</span>
          <select id="global-filter" class="lf-select" onchange="onFilterChange(this.value);updateLfActive()">
            <option value="">-- 전체 보기 --</option>
          </select>
        </div>
        <div class="lf-sep"></div>
        <div class="lf-group">
          <span class="lf-label">플랫폼</span>
          <select id="platform-filter" class="lf-select" onchange="onPlatformFilter(this.value);updateLfActive()"></select>
        </div>
      </div>
      <div class="lf-right">
        <input class="search-input" type="text" placeholder="강의명, 강사명 검색..." oninput="onSearch(this.value)">
      </div>
    </div>
    <div class="sort-bar">
      <span class="sort-label">정렬:</span>
      <div class="sort-btn" data-sort="강의총매출" onclick="sortBy(this)">총 매출순</div>
      <div class="sort-btn" data-sort="순매출" onclick="sortBy(this)">순매출순</div>
      <div class="sort-btn" data-sort="수강생수" onclick="sortBy(this)">수강생순</div>
      <div class="sort-btn" data-sort="무료강의신청수" onclick="sortBy(this)">신청수순</div>
      <div class="sort-btn active" data-sort="날짜순" onclick="sortBy(this)">날짜순</div>
    </div>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th>#</th><th>일자</th><th>강의명</th><th>기수</th><th>플랫폼</th>
            <th class="num">강의 총 매출</th><th class="num">순매출</th>
            <th class="num">수강생</th><th class="num">신청수</th><th></th>
          </tr>
        </thead>
        <tbody id="list-tbody"></tbody>
      </table>
    </div>
    <div class="pagination" id="pagination"></div>
  </div>

  <!-- 강의 상세 -->
  <div class="page" id="page-detail">
    <div class="detail-sticky-header" id="detail-sticky-header">
      <div class="back-btn" onclick="goBackToList()">&#8592; 강의 리스트로</div>
      <div class="detail-title" id="detail-name"></div>
      <div class="detail-meta" id="detail-meta"></div>
      <div id="detail-webinar-link" style="margin-top:6px;font-size:12px;"></div>
    </div>
    <div id="merge-panel" style="display:none;" class="merge-panel"></div>

    <!-- 기존 KPI ID 유지 (JS 로직 보존용, 숨김) -->
    <div style="display:none;">
      <span id="kpi-cvr"></span><span id="kpi-cvr-sub"></span>
      <span id="kpi-students"></span><span id="kpi-free"></span>
    </div>

    <!-- 핵심 지표 패널 -->
    <div class="kpi-panel-wrap">
      <div class="kpi-panel-title">핵심 지표</div>
      <div class="kpi-panel">

        <!-- ① 핵심 KPI 2×3 -->
        <div class="kpi-group">
          <div class="kpi-group-label">핵심 KPI</div>
          <div class="kpi-grid32">
            <div class="kpi-sm">
              <div class="ks-label">강의 총 매출</div>
              <div class="ks-val ks-red" id="kpi-total">-</div>
            </div>
            <div class="kpi-sm">
              <div class="ks-label">수강생 수</div>
              <div class="ks-val ks-red" id="kpi-top-students">-</div>
            </div>
            <div class="kpi-sm">
              <div class="ks-label">신청자 수</div>
              <div class="ks-val" id="kpi-top-free">-</div>
            </div>
            <div class="kpi-sm">
              <div class="ks-label">목표매출</div>
              <div class="ks-val" id="kpi-top-target">-</div>
            </div>
            <div class="kpi-sm">
              <div class="ks-label">목표달성률</div>
              <div class="ks-val" id="kpi-top-achieve">-</div>
              <div class="ks-sub" id="kpi-top-achieve-sub"></div>
            </div>
            <div class="kpi-sm">
              <div class="ks-label">ROAS</div>
              <div class="ks-val" id="kpi-top-roas">-</div>
            </div>
          </div>
        </div>

        <!-- ② 전환 지표 2×3 -->
        <div class="kpi-group">
          <div class="kpi-group-label">전환 지표</div>
          <div class="kpi-grid32">
            <div class="kpi-sm">
              <div class="ks-label">톡방 입장률</div>
              <div class="ks-val" id="kpi-conv-tok">-</div>
              <div class="ks-sub" id="kpi-conv-tok-sub"></div>
            </div>
            <div class="kpi-sm">
              <div class="ks-label">라이브 입장률</div>
              <div class="ks-val" id="kpi-conv-live">-</div>
              <div class="ks-sub" id="kpi-conv-live-sub"></div>
            </div>
            <div class="kpi-sm">
              <div class="ks-label">라이브 결제전환률</div>
              <div class="ks-val ks-red" id="kpi-conv-livepay">-</div>
              <div class="ks-sub" id="kpi-conv-livepay-sub"></div>
            </div>
            <div class="kpi-sm">
              <div class="ks-label">앵콜 입장률</div>
              <div class="ks-val" id="kpi-conv-encore">-</div>
              <div class="ks-sub" id="kpi-conv-encore-sub"></div>
            </div>
            <div class="kpi-sm">
              <div class="ks-label">앵콜 결제전환률</div>
              <div class="ks-val ks-red" id="kpi-conv-encorepay">-</div>
              <div class="ks-sub" id="kpi-conv-encorepay-sub"></div>
            </div>
            <div class="kpi-sm">
              <div class="ks-label">최종 결제전환률</div>
              <div class="ks-val ks-red" id="kpi-top-finalconv">-</div>
              <div class="ks-sub" id="kpi-top-finalconv-sub"></div>
            </div>
          </div>
        </div>

      </div>
    </div>

    <!-- 영업이익 + 광고 성과 -->
    <div class="profit-ad-row">
      <div class="profit-card">
        <div class="profit-title">영업이익 현황</div>
        <div class="profit-body">
          <div>
            <div style="font-size:11px;color:var(--muted);margin-bottom:4px;">최종 순매출</div>
            <div class="profit-main-val" id="kpi-net">-</div>
          </div>
          <div style="text-align:right;">
            <div style="font-size:11px;color:var(--muted);margin-bottom:4px;">영업이익률</div>
            <div class="profit-rate" id="kpi-margin">-<span>%</span></div>
          </div>
        </div>
        <div class="profit-bar-wrap"><div class="profit-bar" id="profit-bar" style="width:0%"></div></div>
        <div class="profit-bar-label" id="profit-bar-label"></div>
      </div>
      <div class="ad-section">
        <div class="ad-title">광고 성과</div>
        <div class="ad-grid">
          <div>
            <div class="ad-item-label">ROAS (광고수익률)</div>
            <div><span class="ad-item-value" id="kpi-roas">-</span><span class="ad-item-unit">배</span></div>
            <div class="ad-item-sub" id="kpi-roas-sub"></div>
          </div>
          <div>
            <div class="ad-item-label">광고비</div>
            <div id="kpi-adcost-wrap">
              <span class="ad-item-value" id="kpi-adcost" style="color:var(--text);font-size:20px;">-</span>
            </div>
            <div style="margin-top:6px;">
              <input id="adcost-input" type="text" placeholder="직접 입력 (원)" inputmode="numeric"
                style="background:var(--card2);border:1px solid var(--border);border-radius:6px;padding:5px 10px;color:var(--text);font-size:12px;width:100%;outline:none;box-sizing:border-box;"
                onfocus="this.style.borderColor='var(--accent)'" onblur="this.style.borderColor='var(--border)'"
                oninput="onAdcostInput(this.value)">
            </div>
            <div class="ad-item-sub" id="kpi-adcost-sub"></div>
          </div>
          <div>
            <div class="ad-item-label">강사 정산금</div>
            <div><span class="ad-item-value" id="kpi-teacher" style="color:var(--text);font-size:20px;">-</span></div>
          </div>
        </div>
      </div>
    </div>

    <!-- 무료강의 후 시간별 구매 추이 -->
    <div class="chart-card full" id="time-chart-card" style="margin-bottom:14px;">
      <div class="chart-header">
        <div>
          <div class="chart-title">&#9646; 무료강의 후 시간별 구매 추이</div>
          <div class="chart-sub" id="time-chart-sub">주문 결제 데이터를 업로드하면 표시됩니다</div>
        </div>
        <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
          <select class="interval-select" id="interval-select" onchange="reRenderTimeChart()">
            <option value="1">1분</option>
            <option value="5">5분</option>
            <option value="10" selected>10분</option>
            <option value="30">30분</option>
            <option value="60">1시간</option>
            <option value="1440">24시간</option>
          </select>
          <input type="checkbox" id="day-only-check" checked style="display:none;">
          <div style="display:flex;flex-direction:column;align-items:flex-end;gap:2px;">
            <div style="display:flex;align-items:center;gap:4px;font-size:12px;color:var(--muted);flex-wrap:wrap;justify-content:flex-end;">
              <input type="date" id="range-start-date" onchange="reRenderTimeChart()"
                style="background:var(--card2);border:1px solid var(--border);border-radius:4px;color:var(--text);font-size:11px;padding:2px 4px;">
              <input type="time" id="range-start-time" value="19:30" onchange="reRenderTimeChart()"
                style="background:var(--card2);border:1px solid var(--border);border-radius:4px;color:var(--text);font-size:11px;padding:2px 4px;">
              <span style="white-space:nowrap;">~</span>
              <input type="date" id="range-end-date" onchange="reRenderTimeChart()"
                style="background:var(--card2);border:1px solid var(--border);border-radius:4px;color:var(--text);font-size:11px;padding:2px 4px;">
              <input type="time" id="range-end-time" value="03:00" onchange="reRenderTimeChart()"
                style="background:var(--card2);border:1px solid var(--border);border-radius:4px;color:var(--text);font-size:11px;padding:2px 4px;">
            </div>
            <span style="font-size:10px;color:var(--muted);opacity:0.7;">선택한 날짜+시간 범위의 데이터만 표시됩니다</span>
          </div>
          <!-- 기존 시간 입력 호환성 보존 (hidden) -->
          <input type="time" id="day-start-time" value="19:30" style="display:none;">
          <input type="time" id="day-end-time" value="03:00" style="display:none;">
        </div>
      </div>
      <div class="chart-wrap" style="height:200px;"><canvas id="chart-time"></canvas></div>
      <input id="elapsed-offset" type="number" value="0" style="display:none;">
      <div style="display:flex;align-items:flex-end;justify-content:flex-end;gap:12px;margin-top:6px;padding-right:4px;flex-wrap:wrap;">
        <div style="display:flex;flex-direction:column;align-items:flex-end;gap:2px;">
          <label style="display:flex;align-items:center;gap:4px;font-size:11px;color:var(--muted);">영상 분석용 라이브 시작시간
            <input type="time" id="live-start-time" value="19:30" oninput="reRenderTimeChart()"
              style="background:var(--card2);border:1px solid var(--border);border-radius:4px;color:var(--text);font-size:11px;padding:2px 4px;">
          </label>
          <span style="font-size:10px;color:var(--muted);opacity:0.7;">영상 기준으로 경과시간을 계산하는 기준입니다</span>
        </div>
      </div>
      <!-- 구매 상세 토글 -->
      <div style="border-top:1px solid var(--border);margin-top:14px;padding-top:10px;">
        <div class="toggle-btn" id="order-toggle" onclick="toggleOrderList()">
          <span>구매 상세 보기</span>
          <span class="arr" id="order-arr">▼</span>
          <span id="order-count-badge" style="font-size:11px;color:var(--accent);margin-left:4px;"></span>
        </div>
        <div class="order-list-wrap" id="order-list-wrap">
          <div class="order-list" id="order-list"></div>
        </div>
      </div>
    </div>

    <!-- 핵심 지표 렌더링용 hidden 컨테이너 (key-metrics-content ID 보존) -->
    <div style="display:none;"><div id="key-metrics-content"></div></div>

    <!-- 성과 퍼널 -->
    <div class="chart-card" id="funnel-card" style="margin-bottom:14px;">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;">
        <div>
          <div class="chart-title">▌ 성과 퍼널</div>
          <div class="chart-sub" id="funnel-sub">결과값을 입력하면 전환율이 자동 계산됩니다</div>
        </div>
      </div>
      <div id="funnel-alert"></div>
      <!-- 헤더 -->
      <div style="display:grid;grid-template-columns:110px 1fr 100px 110px 1fr;gap:10px;padding-bottom:6px;border-bottom:1px solid var(--border);margin-bottom:2px;">
        <span style="font-size:10px;color:var(--muted);">단계</span>
        <span style="font-size:10px;color:var(--muted);"></span>
        <span style="font-size:10px;color:var(--muted);text-align:right;">결과값</span>
        <span style="font-size:10px;color:var(--muted);">전환율 / 목표</span>
        <span style="font-size:10px;color:var(--muted);">AAR 피드백</span>
      </div>
      <div id="funnel-rows"></div>
    </div>

    <!-- 기수별 상세 테이블 -->
    <div class="chart-card" style="margin-bottom:20px;">
      <div class="chart-title" style="margin-bottom:14px;">기수별 상세 데이터</div>
      <div style="overflow-x:auto;">
        <table class="kisu-table">
          <thead>
            <tr>
              <th>기수</th><th>무료강의일</th><th>플랫폼</th>
              <th class="num">강의총매출</th><th class="num">순매출</th>
              <th class="num">수강생</th><th class="num">신청수</th>
            </tr>
          </thead>
          <tbody id="kisu-tbody"></tbody>
        </table>
      </div>
    </div>
  </div>

</div><!-- /main -->

<div class="page-footer">© 2026 염정하</div>

</div><!-- /app-wrapper -->

<script>
// ════════════════════════════════════════════════════
//  인증 (로그인)
// ════════════════════════════════════════════════════
(function initAuth() {
  const AUTH_KEY = 'nova_auth';
  const FIXED_PW = 'sn147'; // 고정 비밀번호 (변경 불가)

  function isLoggedIn()  { return sessionStorage.getItem(AUTH_KEY) === '1'; }

  function showApp() {
    document.getElementById('login-screen').classList.remove('active');
    document.getElementById('app-wrapper').style.display = '';
  }
  function showLoginScreen() {
    document.getElementById('login-screen').classList.add('active');
    document.getElementById('app-wrapper').style.display = 'none';
    setTimeout(function(){ var el=document.getElementById('login-pw-input'); if(el) el.focus(); }, 50);
  }

  window.doLogin = function() {
    var pw  = document.getElementById('login-pw-input').value;
    var inp = document.getElementById('login-pw-input');
    var err = document.getElementById('login-error');
    if (pw === FIXED_PW) {
      sessionStorage.setItem(AUTH_KEY, '1');
      inp.classList.remove('error');
      err.textContent = '';
      showApp();
    } else {
      inp.classList.add('error');
      err.textContent = '비밀번호가 올바르지 않습니다.';
      inp.value = '';
      inp.focus();
    }
  };

  window.doLogout = function() {
    sessionStorage.removeItem(AUTH_KEY);
    location.reload();
  };

  // 초기 인증 상태 확인
  if (isLoggedIn()) showApp();
  else showLoginScreen();
})();

// ════════════════════════════════════════════════════
//  상수 & 상태
// ════════════════════════════════════════════════════
const DEFAULT_DATA = ${dataJSON};

let SALES_DATA = DEFAULT_DATA.slice();
let ORDER_DATA  = [];          // [{ts: Date, amount: number, lectureName: string}]
let RAW_SALES_ROWS = [];       // parseSalesRows 결과 (재집계용)
let MERGE_OVERRIDES_EXCLUDED = new Set(); // "keyA|||keyB" 형식 쌍 (병합 제외)

let pendingSales = null;       // ArrayBuffer
let pendingOrder = null;       // ArrayBuffer
let pendingSchedule = null;    // ArrayBuffer (일정 시트 엑셀)
let SCHEDULE_DATA = [];        // [{강의명, 강사, 강의일, 시간, 플랫폼, 기수, 상태, ...}]

const PAGE_SIZE = 20;
let currentPage = 1;
let currentSort = '날짜순';
let searchQuery  = '';
let selectedKey  = '';
let platformFilter = '';
let charts = {};
let currentDetailKey = '';
let _historyLock = false; // 뒤로가기/상세 전환 시 이중 pushState 방지

// ════════════════════════════════════════════════════
//  강의명 정규화 (빌드 스크립트와 동일 로직)
// ════════════════════════════════════════════════════
function normalizeLectureName(name) {
  if (!name) return '';
  let n = name.trim();
  n = n.replace(/\\s*[-]\\s*[\\d,]+\\s*(만원|원)/gi, '');
  n = n.replace(/\\s*[-]\\s*\\d+\\s*$/g, '');
  n = n.replace(/\\s*\\(\\d+\\)\\s*$/g, '');
  n = n.replace(/\\s*[-_]\\s*(재결제|재결재|추가결제|프리미엄\\s*전용|프리미엄|얼리버드|테스트|복제됨|복제|전용|추가)\\s*/gi, '');
  n = n.replace(/\\s*가\\s*$/g, '');
  n = n.replace(/[-\\s]+$/, '').replace(/\\s+/g, ' ').trim();
  return n;
}

// ════════════════════════════════════════════════════
//  Excel → JSON 파싱 (브라우저 SheetJS)
// ════════════════════════════════════════════════════
function parseExcelBuffer(buffer) {
  const wb = XLSX.read(buffer, { type: 'array', cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
}

// 컬럼 자동 매핑
const COL_ALIAS = {
  플랫폼:    ['플랫폼','platform'],
  무료강의일: ['무료강의일','강의일','날짜','date'],
  강사:      ['강사','강사명','teacher','instructor'],
  강의명:    ['강의명','상품명','강좌명','lecture_name','product_name'],
  기수:      ['기수','기수명','round'],
  무료강의신청수: ['무료강의 신청수','무료강의신청수','신청수','applicants'],
  강의총매출: ['강의 총 매출 (A)','강의총매출','총매출','total_revenue'],
  강의총매출수강생수: ['강의 총 매출 수강생 수','수강생수','수강생','students'],
  플랫폼매출: ['플랫폼 매출','플랫폼매출','platform_revenue'],
  PG사수수료: ['PG사 수수료 (B)','PG사수수료','PG수수료','pg_fee'],
  PG제외매출: ['PG 제외 매출 (C) = (A - B)','PG제외매출','pg_excluded'],
  노바수수료: ['노바 수수료 (D)','노바수수료','nova_fee'],
  광고비:    ['광고비 (E)','광고비','ad_cost'],
  기타비용:  ['기타 비용 (F)','기타비용','other_cost'],
  순매출:    ['순매출 (G) = (C - D - E - F)','순매출','net_revenue'],
  플랫폼수익금: ['플랫폼 수익금','플랫폼수익금','platform_profit'],
  인플루언서RS정산금: ['인플루언서 RS 정산금','인플루언서RS정산금','influencer_fee'],
  강사정산금: ['강사 정산금','강사정산금','teacher_fee'],
};

const ORDER_ALIAS = {
  결제일시: ['주문일시','결제일시','결제시간','구매일시','주문시간','결제날짜','payment_time','ordered_at','datetime'],
  결제금액: ['매출금액','결제금액','주문금액','금액','구매금액','amount','price','결제액'],
  회원명:   ['회원명','구매자명','구매자','이름','name','buyer','member_name'],
  전화번호: ['휴대전화번호','전화번호','휴대폰번호','연락처','phone','mobile','tel'],
  강의명:   ['주문항목명','강의명','상품명','강좌명','lecture_name','product_name','상품'],
  주문상태: ['주문상태','order_status','상태','status'],
};

function findColIdx(headers, candidates) {
  const norm = h => (h||'').toString().toLowerCase().replace(/[\s()=\-_]/g,'');
  for (const c of candidates) {
    const idx = headers.findIndex(h => norm(h) === norm(c));
    if (idx >= 0) return idx;
  }
  return -1;
}

function parseSalesRows(rows) {
  if (!rows || rows.length < 2) return [];
  const headers = rows[0];
  const get = (aliases) => findColIdx(headers, aliases);
  const idx = {};
  Object.entries(COL_ALIAS).forEach(([k, v]) => { idx[k] = get(v); });

  return rows.slice(1).filter(r => r[idx.플랫폼] && r[idx.플랫폼] !== '시스템노바 테스트용' && r[idx.강의명] !== 'test').map(r => {
    const g = (k, def) => { const i = idx[k]; return i >= 0 ? r[i] : def; };
    return {
      플랫폼: g('플랫폼',''), 무료강의일: g('무료강의일',''), 강사: g('강사',''),
      강의명: g('강의명',''), 기수: g('기수','-') || '-',
      무료강의신청수: Number(g('무료강의신청수',0))||0,
      강의총매출: Number(g('강의총매출',0))||0,
      강의총매출수강생수: Number(g('강의총매출수강생수',0))||0,
      플랫폼매출: Number(g('플랫폼매출',0))||0,
      PG사수수료: Number(g('PG사수수료',0))||0,
      PG제외매출: Number(g('PG제외매출',0))||0,
      노바수수료: Number(g('노바수수료',0))||0,
      광고비: Number(g('광고비',0))||0,
      기타비용: Number(g('기타비용',0))||0,
      순매출: Number(g('순매출',0))||0,
      플랫폼수익금: Number(g('플랫폼수익금',0))||0,
      인플루언서RS정산금: Number(g('인플루언서RS정산금',0))||0,
      강사정산금: Number(g('강사정산금',0))||0,
    };
  });
}

function parseOrderRows(rows) {
  if (!rows || rows.length < 2) return [];
  const headers = rows[0];
  const tsIdx     = findColIdx(headers, ORDER_ALIAS.결제일시);
  const amtIdx    = findColIdx(headers, ORDER_ALIAS.결제금액);
  const lecIdx    = findColIdx(headers, ORDER_ALIAS.강의명);
  const statusIdx = findColIdx(headers, ORDER_ALIAS.주문상태);
  const nameIdx   = findColIdx(headers, ORDER_ALIAS.회원명);
  const phoneIdx  = findColIdx(headers, ORDER_ALIAS.전화번호);

  // 디버그: 매핑된 컬럼 확인 (콘솔 출력)
  console.log('[주문파싱] 헤더:', headers);
  console.log('[주문파싱] 컬럼 인덱스 → 주문일시:', tsIdx, '/ 매출금액:', amtIdx, '/ 주문항목명:', lecIdx, '/ 주문상태:', statusIdx);

  const result = [];
  const refundList = [];  // 환불/취소 건 별도 저장
  rows.slice(1).forEach(r => {
    // 빈 행 스킵
    if (!r || r.every(v => v === null || v === undefined || v === '')) return;

    // 주문상태 필터 - 환불/취소 별도 저장 후 제외
    if (statusIdx >= 0) {
      const st = String(r[statusIdx] || '');
      if (/환불|취소|cancel|refund/i.test(st)) {
        const lecName = lecIdx >= 0 ? String(r[lecIdx]||'') : '';
        const rName  = nameIdx  >= 0 ? String(r[nameIdx]||'').trim()  : '';
        const rPhone = phoneIdx >= 0 ? String(r[phoneIdx]||'').replace(/[^0-9]/g,'') : '';
        if (lecName) refundList.push({ lectureName: lecName, name: rName, phone: rPhone });
        return;
      }
    }

    const rawTs = tsIdx >= 0 ? r[tsIdx] : null;
    if (!rawTs) return;

    let ts = null;
    if (rawTs instanceof Date && !isNaN(rawTs)) {
      ts = rawTs;
    } else {
      // 다양한 날짜 포맷 처리
      let s = String(rawTs).trim()
        .replace(/년\\s*/g, '-').replace(/월\\s*/g, '-').replace(/일\\s*/g, ' ')
        .replace(/시\\s*/g, ':').replace(/분\\s*/g, ':').replace(/초/g, '')
        .replace(/\\./g, '-').replace(/\\//g, '-').trim();
      // "2024-01-15 14:03:21" → ISO 변환 (일부 브라우저 호환)
      s = s.replace(/^(\\d{4}-\\d{2}-\\d{2})\\s+(\\d{2}:\\d{2}(:\\d{2})?)$/, '$1T$2');
      const d = new Date(s);
      if (!isNaN(d)) ts = d;
    }
    if (!ts) return;

    const rawAmt = amtIdx >= 0 ? r[amtIdx] : 0;
    const amt   = typeof rawAmt === 'number' ? rawAmt : Number(String(rawAmt||0).replace(/,/g,''))||0;
    const lec   = lecIdx   >= 0 ? String(r[lecIdx]||'')   : '';
    const name  = nameIdx  >= 0 ? String(r[nameIdx]||'')  : '';
    const phone = phoneIdx >= 0 ? String(r[phoneIdx]||'') : '';
    result.push({ ts, amount: amt, lectureName: lec, name, phone });
  });

  console.log('[주문파싱] 파싱 완료:', result.length, '건 (총', rows.length-1, '행 중)', '/ 환불/취소:', refundList.length, '건');
  window._REFUND_LIST = refundList;
  return result.sort((a,b) => a.ts - b.ts);
}

// ════════════════════════════════════════════════════
//  집계
// ════════════════════════════════════════════════════
function aggregateByKey(dataRows) {
  const map = {};
  dataRows.forEach(r => {
    const normName = normalizeLectureName(r.강의명);
    if (!normName) return;
    const kisu = (r.기수 && String(r.기수).trim() !== '') ? String(r.기수).trim() : '-';
    if (!map[normName]) map[normName] = {
      lecture_key: normName, 강의그룹명: normName, 기수List: [],
      강사: r.강사, 플랫폼: r.플랫폼,
      강의총매출:0, 순매출:0, 수강생수:0, 무료강의신청수:0,
      PG제외매출:0, PG사수수료:0, 노바수수료:0, 광고비:0,
      기타비용:0, 플랫폼수익금:0, 강사정산금:0, rows:[], _km:{}
    };
    const d = map[normName];
    if (!d.기수List.includes(kisu)) d.기수List.push(kisu);
    d.강의총매출+=r.강의총매출; d.순매출+=r.순매출; d.수강생수+=r.강의총매출수강생수;
    d.무료강의신청수 = Math.max(d.무료강의신청수, r.무료강의신청수);
    d.PG제외매출+=r.PG제외매출; d.PG사수수료+=r.PG사수수료; d.노바수수료+=r.노바수수료;
    d.광고비+=r.광고비; d.기타비용+=r.기타비용;
    d.플랫폼수익금+=r.플랫폼수익금; d.강사정산금+=r.강사정산금;
    // 기수별 row 합산 (기수당 정확히 1개)
    if (!d._km[kisu]) d._km[kisu] = {
      무료강의일:r.무료강의일, 기수:kisu, 플랫폼:r.플랫폼, 강의명:r.강의명,
      강의총매출:0, 순매출:0, 수강생수:0, 무료강의신청수:0,
      PG제외매출:0, 플랫폼매출:0, 광고비:0, 노바수수료:0,
      PG사수수료:0, 기타비용:0, 강사정산금:0
    };
    const k = d._km[kisu];
    k.강의총매출+=r.강의총매출; k.순매출+=r.순매출; k.수강생수+=r.강의총매출수강생수;
    k.무료강의신청수=Math.max(k.무료강의신청수, r.무료강의신청수);
    k.PG제외매출+=r.PG제외매출; k.플랫폼매출+=r.플랫폼매출;
    k.광고비+=r.광고비; k.노바수수료+=r.노바수수료;
    k.PG사수수료+=r.PG사수수료; k.기타비용+=r.기타비용; k.강사정산금+=r.강사정산금;
  });
  return Object.values(map).map(d => {
    d.rows = Object.values(d._km).sort((a,b)=>(a.무료강의일||'').localeCompare(b.무료강의일||''));
    delete d._km;
    return d;
  }).sort((a,b) => b.강의총매출 - a.강의총매출);
}

// ════════════════════════════════════════════════════
//  자동 그룹 병합 (강의일+기수 동일, 강의명 유사도≥0.5)
// ════════════════════════════════════════════════════
function nameSimilarity(a, b) {
  a = a.toLowerCase().replace(/\\s+/g,' ').trim();
  b = b.toLowerCase().replace(/\\s+/g,' ').trim();
  if (a === b) return 1;
  if (a.includes(b) || b.includes(a)) return 0.9;
  const sa = new Set(a.split(/\\s+/)), sb = new Set(b.split(/\\s+/));
  let common = 0; sa.forEach(t => { if (sb.has(t)) common++; });
  const union = sa.size + sb.size - common;
  return union > 0 ? common / union : 0;
}
function mergeGroupPair(target, source) {
  const km = {};
  [...target.rows, ...source.rows].forEach(r => {
    const k = normDateStr(r.무료강의일) + '::' + r.기수;
    if (!km[k]) { km[k] = Object.assign({}, r); return; }
    const t = km[k];
    t.강의총매출 += r.강의총매출||0; t.순매출 += r.순매출||0; t.수강생수 += r.수강생수||0;
    t.무료강의신청수 = Math.max(t.무료강의신청수||0, r.무료강의신청수||0);
    t.PG제외매출 += r.PG제외매출||0; t.플랫폼매출 += r.플랫폼매출||0;
    t.광고비 += r.광고비||0; t.노바수수료 += r.노바수수료||0;
    t.PG사수수료 += r.PG사수수료||0; t.기타비용 += r.기타비용||0; t.강사정산금 += r.강사정산금||0;
  });
  target.rows = Object.values(km).sort((a,b) => (a.무료강의일||'').localeCompare(b.무료강의일||''));
  target.강의총매출 += source.강의총매출; target.순매출 += source.순매출;
  target.수강생수 += source.수강생수;
  target.무료강의신청수 = Math.max(target.무료강의신청수, source.무료강의신청수);
  target.PG제외매출 += source.PG제외매출||0; target.PG사수수료 += source.PG사수수료||0;
  target.노바수수료 += source.노바수수료||0; target.광고비 += source.광고비||0;
  target.기타비용 += source.기타비용||0; target.플랫폼수익금 += source.플랫폼수익금||0;
  target.강사정산금 += source.강사정산금||0;
  source.기수List.forEach(k => { if (!target.기수List.includes(k)) target.기수List.push(k); });
  if (!target.mergedFrom) target.mergedFrom = [];
  target.mergedFrom.push({ lecture_key: source.lecture_key, 강의그룹명: source.강의그룹명 });
  if (source.mergedFrom) target.mergedFrom.push(...source.mergedFrom);
}
function autoMergeGroups(groups, excluded) {
  excluded = excluded || new Set();
  const byDateKisu = {};
  groups.forEach(g => {
    g.rows.forEach(r => {
      const dk = normDateStr(r.무료강의일) + '::' + r.기수;
      if (!byDateKisu[dk]) byDateKisu[dk] = [];
      if (!byDateKisu[dk].find(x => x.lecture_key === g.lecture_key)) byDateKisu[dk].push(g);
    });
  });
  const gMap = {}; groups.forEach(g => gMap[g.lecture_key] = g);
  const parent = {}; groups.forEach(g => parent[g.lecture_key] = g.lecture_key);
  function find(k) { return parent[k] === k ? k : (parent[k] = find(parent[k])); }
  function union(a, b) {
    const ra = find(a), rb = find(b); if (ra === rb) return;
    const pk = [ra, rb].sort().join('|||');
    if (excluded.has(pk)) return;
    const ga = gMap[ra], gb = gMap[rb];
    if (ga && gb && ga.강의총매출 >= gb.강의총매출) parent[rb] = ra; else parent[ra] = rb;
  }
  Object.entries(byDateKisu).forEach(([dk, gs]) => {
    if (gs.length < 2) return;
    const date = dk.split('::')[0];
    if (!date) return;
    for (let i = 0; i < gs.length; i++)
      for (let j = i+1; j < gs.length; j++)
        if (nameSimilarity(gs[i].강의그룹명, gs[j].강의그룹명) >= 0.75) union(gs[i].lecture_key, gs[j].lecture_key);
  });
  const rootMap = {}; groups.forEach(g => {
    const root = find(g.lecture_key);
    if (!rootMap[root]) rootMap[root] = { primary: gMap[root], sources: [] };
    if (g.lecture_key !== root) rootMap[root].sources.push(g);
  });
  Object.values(rootMap).forEach(({ primary, sources }) => sources.forEach(src => mergeGroupPair(primary, src)));
  return groups.filter(g => find(g.lecture_key) === g.lecture_key);
}
function rebuildSalesData() {
  if (!RAW_SALES_ROWS.length) return;
  SALES_DATA = autoMergeGroups(aggregateByKey(RAW_SALES_ROWS), MERGE_OVERRIDES_EXCLUDED);
  initFilter(); renderList();
  if (currentDetailKey) showDetail(currentDetailKey);
}
function excludeMerge(keyA, keyB) {
  const pk = [keyA, keyB].sort().join('|||');
  MERGE_OVERRIDES_EXCLUDED.add(pk);
  idbSaveMergeExclusions(MERGE_OVERRIDES_EXCLUDED);
  rebuildSalesData();
}

// ════════════════════════════════════════════════════
//  업로드 UI
// ════════════════════════════════════════════════════
function toggleUpload() {
  const body  = document.getElementById('upload-body');
  const tog   = document.getElementById('upload-toggle');
  const arrow = document.getElementById('upload-arrow');
  const open  = body.classList.toggle('open');
  tog.classList.toggle('open', open);
  arrow.textContent = open ? '▲' : '▼';
}

function handleFile(n, input) {
  const file = input.files[0]; if (!file) return;
  const nameEl = document.getElementById('fname' + n);
  const card   = document.getElementById('card' + n);
  const reader = new FileReader();
  reader.onload = e => {
    if (n === 1) pendingSales = e.target.result;
    else if (n === 2) pendingOrder = e.target.result;
    else if (n === 3) pendingSchedule = e.target.result;
    nameEl.textContent = file.name;
    nameEl.className = 'upload-card-name filled';
    card.classList.add('has-file');
    document.getElementById('apply-btn').disabled = !(pendingSales || pendingOrder || pendingSchedule);
    // IndexedDB 저장
    idbSave(n, e.target.result, file.name);
  };
  reader.readAsArrayBuffer(file);
}

// ════════════════════════════════════════════════════
//  노바 강의일정 시트 파싱
// ════════════════════════════════════════════════════
const SCHEDULE_ALIAS = {
  강의명: ['강의명','강의제목','강좌명','상품명','lecture_name'],
  강사:   ['강사','강사명','instructor','teacher'],
  강의일: ['강의일','무료강의일','일정','날짜','date'],
  시간:   ['시간','강의시간','time'],
  플랫폼: ['플랫폼','platform'],
  기수:   ['기수','기수명','round'],
  상태:   ['상태','진행상태','status'],
  톡방인원:     ['톡방인원','톡방 인원','단톡방인원','단톡방 인원','카톡방인원','톡방유입','단톡방유입'],
  라이브참여자: ['라이브참여자','라이브 참여자','라이브참여','라이브 참여','live참여자','참여자수','참여자'],
  목표매출:     ['목표매출','목표 매출','target_revenue','목표금액','매출목표'],
  ROAS:         ['ROAS','roas','로아스','광고수익률'],
  광고집행출연료: ['광고집행+출연료','광고집행출연료','광고비+출연료','광고집행','광고비출연료','광고+출연료'],
  무료웨비나링크: ['무료웨비나 링크','무료웨비나링크','웨비나링크','웨비나 링크','webinar_link','webinar link','무료강의링크','무료강의 링크'],
};

function parseScheduleRows(rows) {
  if (!rows || rows.length < 2) return [];
  const headers = rows[0];
  console.log('[일정파싱] 헤더:', headers);
  const idx = {};
  Object.entries(SCHEDULE_ALIAS).forEach(([k, v]) => { idx[k] = findColIdx(headers, v); });
  console.log('[일정파싱] 컬럼 인덱스:', JSON.stringify(idx));

  // 유사 매핑 감지
  const _fuzzyMap = {};
  const normH = h => (h||'').toString().toLowerCase().replace(/[\\s()=\\-_]/g,'');
  Object.entries(SCHEDULE_ALIAS).forEach(([k, candidates]) => {
    if (idx[k] >= 0) { _fuzzyMap[k] = { matched:true, exactMatch:true, colName:'' }; return; }
    // 정확 매칭 실패 → 유사 매핑 시도
    const keywords = candidates.map(c => normH(c));
    let bestIdx = -1, bestHeader = '';
    headers.forEach((h,hi) => {
      const nh = normH(h);
      if (!nh) return;
      for (const kw of keywords) {
        if (nh.includes(kw) || kw.includes(nh)) {
          if (bestIdx < 0) { bestIdx = hi; bestHeader = String(h); }
        }
      }
    });
    if (bestIdx >= 0) {
      idx[k] = bestIdx;
      _fuzzyMap[k] = { matched:true, exactMatch:false, colName: bestHeader };
      console.log('[일정파싱] 유사 매핑: ' + k + ' → "' + bestHeader + '" (col ' + bestIdx + ')');
    } else {
      _fuzzyMap[k] = { matched:false, exactMatch:false, colName:'' };
    }
  });
  window._scheduleFuzzyWarnings = _fuzzyMap;

  const result = rows.slice(1).filter(r => r.some && r.some(v => v)).map(r => {
    const g = (k, def) => { const i = idx[k]; return i >= 0 && r[i] != null ? String(r[i]) : def; };
    return {
      강의명: g('강의명',''), 강사: g('강사',''), 강의일: g('강의일',''),
      시간: g('시간',''), 플랫폼: g('플랫폼',''), 기수: g('기수',''), 상태: g('상태',''),
      톡방인원: g('톡방인원',''), 라이브참여자: g('라이브참여자',''),
      목표매출: g('목표매출',''), ROAS: g('ROAS',''),
      광고집행출연료: g('광고집행출연료',''),
      무료웨비나링크: g('무료웨비나링크',''),
    };
  }).filter(r => r.강의명);
  console.log('[일정파싱] 파싱 완료:', result.length, '행');
  if (result.length > 0) console.log('[일정파싱] 첫 행:', JSON.stringify(result[0]));
  return result;
}

function applyUpload() {
  const btn = document.getElementById('apply-btn');
  const status = document.getElementById('upload-status');
  btn.disabled = true;
  status.textContent = '처리 중...';
  try {
    if (pendingSales) {
      const rows = parseExcelBuffer(pendingSales);
      RAW_SALES_ROWS = parseSalesRows(rows);
      SALES_DATA = autoMergeGroups(aggregateByKey(RAW_SALES_ROWS), MERGE_OVERRIDES_EXCLUDED);
      initFilter();
    }
    if (pendingOrder) {
      // 주문 파일은 raw:true + cellDates:true → Date 객체로 직접 수신
      const wb2 = XLSX.read(pendingOrder, { type: 'array', cellDates: true });
      const ws2 = wb2.Sheets[wb2.SheetNames[0]];
      const rows2 = XLSX.utils.sheet_to_json(ws2, { header: 1, raw: true, cellDates: true });
      ORDER_DATA = parseOrderRows(rows2);
    }
    if (pendingSchedule) {
      console.log('[applyUpload] 일정 파일 처리 시작, buffer 크기:', pendingSchedule.byteLength||'N/A');
      const wb3 = XLSX.read(pendingSchedule, { type: 'array', raw: false });
      console.log('[applyUpload] 시트 이름들:', wb3.SheetNames);
      const ws3 = wb3.Sheets[wb3.SheetNames[0]];
      const rows3 = XLSX.utils.sheet_to_json(ws3, { header: 1, raw: false });
      console.log('[applyUpload] 엑셀 행 수:', rows3.length, '| 헤더:', rows3[0]);
      SCHEDULE_DATA = parseScheduleRows(rows3);
      console.log('[applyUpload] SCHEDULE_DATA:', SCHEDULE_DATA.length, '행');
    }
    const parts = [];
    if (pendingSales) parts.push(SALES_DATA.length + '개 강의');
    if (pendingOrder) parts.push(ORDER_DATA.length > 0 ? ORDER_DATA.length + '건 주문' : '⚠️ 주문 0건 (컬럼 확인 필요)');
    if (pendingSchedule) parts.push(SCHEDULE_DATA.length + '개 일정');
    status.textContent = '✓ 적용 완료 (' + parts.join(' / ') + ')';
    const _todayStr = new Date().getFullYear() + '-' + String(new Date().getMonth()+1).padStart(2,'0') + '-' + String(new Date().getDate()).padStart(2,'0');
    idbSaveLastUpload(_todayStr);
    updateBanner();
    renderList();
    if (currentDetailKey) showDetail(currentDetailKey);
    document.getElementById('upload-badge').innerHTML =
      '<span style="background:rgba(16,185,129,0.2);color:var(--green);padding:1px 8px;border-radius:4px;font-size:11px;margin-left:4px;">적용됨</span>';
    setTimeout(() => {
      const body = document.getElementById('upload-body');
      body.classList.remove('open');
      document.getElementById('upload-toggle').classList.remove('open');
      document.getElementById('upload-arrow').textContent = '▼';
    }, 1500);
  } catch(e) {
    status.textContent = '오류: ' + e.message;
    btn.disabled = false;
  }
}

// ════════════════════════════════════════════════════
//  포맷 유틸
// ════════════════════════════════════════════════════
function fmt(n) {
  n = Math.round(n);
  if (!n) return '0';
  if (Math.abs(n) >= 100000000) return (n/100000000).toFixed(1) + '억';
  if (Math.abs(n) >= 10000) return Math.round(n/10000).toLocaleString() + '만';
  return n.toLocaleString();
}
function fmtFull(n) { return Math.round(n).toLocaleString(); }
function fmtUnit(n, unit) {
  return '<span style="font-size:28px;font-weight:800;letter-spacing:-.02em;">' + fmtFull(n) + '</span>' +
         '<span class="kpi-unit">' + unit + '</span>';
}
function pad2(n) { return String(n).padStart(2,'0'); }
function fmtTime(d) { return pad2(d.getHours()) + ':' + pad2(d.getMinutes()) + ':' + pad2(d.getSeconds()); }
function fmtDateTime(d) { return (d.getMonth()+1) + '/' + pad2(d.getDate()) + ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes()) + ':' + pad2(d.getSeconds()); }
// 라이브 시작 시간 기준 경과시간 — +H:MM:SS 형식 (offsetMin: 보정값(분), 기본 0)
function fmtElapsed(ts, offsetMin) {
  const h = ts.getHours();
  // 라이브 시작 시간 입력값 읽기 (없으면 19:30 기본값)
  const _lv = (document.getElementById('live-start-time') || {}).value || '19:30';
  const [_lh, _lm] = _lv.split(':').map(Number);
  const baseH = isNaN(_lh) ? 19 : _lh;
  const baseM = isNaN(_lm) ? 30 : _lm;
  // 익일 새벽(0~3시)이면 전날 라이브 시작 기준
  const base = h < 3
    ? new Date(ts.getFullYear(), ts.getMonth(), ts.getDate() - 1, baseH, baseM, 0)
    : new Date(ts.getFullYear(), ts.getMonth(), ts.getDate(), baseH, baseM, 0);
  const diffMs = ts.getTime() - base.getTime() + (offsetMin || 0) * 60000;
  if (diffMs < 0) return '-';
  const totalSec = Math.floor(diffMs / 1000);
  const hh = Math.floor(totalSec / 3600);
  const mm = Math.floor((totalSec % 3600) / 60);
  const ss = totalSec % 60;
  return '+' + hh + ':' + pad2(mm) + ':' + pad2(ss);
}
function getElapsedOffset() {
  const el = document.getElementById('elapsed-offset');
  return el ? (parseInt(el.value) || 0) : 0;
}
function updateDetailStickyTop() {
  const stickyNav = document.querySelector('.sticky-top');
  const detailHdr = document.getElementById('detail-sticky-header');
  if (stickyNav && detailHdr) detailHdr.style.top = stickyNav.offsetHeight + 'px';
}

// ════════════════════════════════════════════════════
//  필터 & 리스트
// ════════════════════════════════════════════════════
// 무료강의일 문자열을 YYYY/MM/DD 형식으로 변환
function formatDateDisp(s) {
  if (!s || s === '-') return '';
  let r = '', src = String(s).slice(0, 10);
  for (let i = 0; i < src.length; i++) { const c = src[i]; r += (c === '.' || c === '/') ? '/' : c; }
  return r;
}

// 날짜 비교용 정규화 (YYYY-MM-DD)
function normDateStr(s) {
  if (!s || s === '-') return '';
  let r = '', src = String(s).slice(0, 10);
  for (let i = 0; i < src.length; i++) { const c = src[i]; r += (c === '.' || c === '/') ? '-' : c; }
  return r;
}

function initPlatformFilter() {
  const sel = document.getElementById('platform-filter');
  const platforms = [...new Set(SALES_DATA.map(d => d.플랫폼).filter(Boolean))].sort();
  sel.innerHTML = '<option value="">전체</option>' +
    platforms.map(p => '<option value="'+p+'"'+(p===platformFilter?' selected':'')+'>'+p+'</option>').join('');
}

function onPlatformFilter(val) {
  platformFilter = val;
  selectedKey = '';
  currentDetailKey = '';
  currentPage = 1;
  const gfEl = document.getElementById('global-filter');
  if (gfEl) gfEl.value = '';
  rebuildGlobalFilter();
  updateLfActive();
  renderList();
}

function rebuildGlobalFilter() {
  const sel = document.getElementById('global-filter');
  if (!sel) return;
  sel.innerHTML = '<option value="">-- 전체 보기 --</option>';
  const source = platformFilter ? SALES_DATA.filter(d => d.플랫폼 === platformFilter) : SALES_DATA;
  const entries = [];
  source.forEach(d => {
    d.rows.forEach(r => {
      const ds = normDateStr(r.무료강의일) || '';
      const datePrefix = ds ? formatDateDisp(ds) + ' | ' : '';
      const kisuLabel = (r.기수 && r.기수 !== '-') ? ' (' + r.기수 + ')' : '';
      entries.push({ value: d.lecture_key + '::' + r.기수, date: ds, label: datePrefix + d.강의그룹명 + kisuLabel });
    });
  });
  entries.sort((a,b) => b.date.localeCompare(a.date));
  entries.forEach(e => {
    const opt = document.createElement('option');
    opt.value = e.value;
    opt.textContent = e.label.length > 55 ? e.label.slice(0, 55) + '…' : e.label;
    sel.appendChild(opt);
  });
}

function updateLfActive() {
  const gf = document.getElementById('global-filter');
  const pf = document.getElementById('platform-filter');
  const df = document.getElementById('date-from');
  const dt = document.getElementById('date-to');
  if (gf) gf.classList.toggle('lf-active', !!gf.value);
  if (pf) pf.classList.toggle('lf-active', !!pf.value);
  if (df) df.classList.toggle('lf-active', !!df.value);
  if (dt) dt.classList.toggle('lf-active', !!dt.value);
}

function initFilter() {
  initPlatformFilter();
  const sel = document.getElementById('global-filter');
  sel.innerHTML = '<option value="">-- 전체 보기 --</option>';
  // 기수별 개별 항목 생성, 날짜 내림차순
  const entries = [];
  SALES_DATA.forEach(d => {
    d.rows.forEach(r => {
      const ds = normDateStr(r.무료강의일) || '';
      const datePrefix = ds ? formatDateDisp(ds) + ' | ' : '';
      const kisuLabel = (r.기수 && r.기수 !== '-') ? ' (' + r.기수 + ')' : '';
      const label = datePrefix + d.강의그룹명 + kisuLabel;
      entries.push({ value: d.lecture_key + '::' + r.기수, date: ds, label });
    });
  });
  entries.sort((a,b) => b.date.localeCompare(a.date));
  entries.forEach(e => {
    const opt = document.createElement('option');
    opt.value = e.value;
    opt.textContent = e.label.length > 55 ? e.label.slice(0, 55) + '…' : e.label;
    sel.appendChild(opt);
  });
}

function onFilterChange(val) {
  selectedKey = val;
  if (val) showDetail(val);
  else { showPage('list'); renderList(); }
}

function onSearch(val) { searchQuery = val; currentPage = 1; renderList(); }

function resetAll() {
  searchQuery = ''; selectedKey = ''; platformFilter = '';
  currentSort = '날짜순'; currentPage = 1; currentDetailKey = '';
  document.getElementById('date-from').value = '';
  document.getElementById('date-to').value = '';
  document.getElementById('date-result').textContent = '';
  const searchEl = document.querySelector('.search-input');
  if (searchEl) searchEl.value = '';
  const pfEl = document.getElementById('platform-filter');
  if (pfEl) pfEl.value = '';
  const gfEl = document.getElementById('global-filter');
  if (gfEl) gfEl.value = '';
  document.querySelectorAll('.sort-btn').forEach(b => b.classList.toggle('active', b.dataset.sort === '날짜순'));
  updateLfActive();
  showPage('home');
}

function showPlatformPicker() {
  document.getElementById('home-main-view').classList.add('hidden');
  document.getElementById('home-platform-view').classList.add('active');
  const grid = document.getElementById('platform-grid');
  const platforms = [...new Set(SALES_DATA.map(d => d.플랫폼).filter(Boolean))].sort();
  window._homePlatforms = platforms;
  grid.innerHTML = platforms.map(function(p, idx) {
    const count = SALES_DATA.filter(d => d.플랫폼 === p).length;
    return '<div class="platform-card" onclick="selectPlatform(window._homePlatforms[' + idx + '])">' +
      '<div class="platform-card-name">' + p + '</div>' +
      '<div class="platform-card-count">강의 ' + count + '개</div>' +
      '</div>';
  }).join('');
}

function hidePlatformPicker() {
  var mv = document.getElementById('home-main-view');
  var pv = document.getElementById('home-platform-view');
  if (mv) mv.classList.remove('hidden');
  if (pv) pv.classList.remove('active');
}

function goBackToList() {
  selectedKey = '';
  currentDetailKey = '';
  currentPage = 1;
  const gfEl = document.getElementById('global-filter');
  if (gfEl) gfEl.value = '';
  rebuildGlobalFilter();
  updateLfActive();
  showPage('list');
}

function selectPlatform(platform) {
  platformFilter = platform;
  selectedKey = '';
  currentDetailKey = '';
  currentPage = 1;
  const pfEl = document.getElementById('platform-filter');
  if (pfEl) pfEl.value = platform;
  const gfEl = document.getElementById('global-filter');
  if (gfEl) gfEl.value = '';
  rebuildGlobalFilter();
  updateLfActive();
  showPage('list');
}

function sortBy(el) {
  currentSort = el.dataset.sort; currentPage = 1;
  document.querySelectorAll('.sort-btn').forEach(b => b.classList.remove('active'));
  el.classList.add('active'); renderList();
}

function onDateFilter() { currentPage = 1; renderList(); }
function resetDateFilter() {
  document.getElementById('date-from').value = '';
  document.getElementById('date-to').value = '';
  currentPage = 1; renderList();
}

function getFiltered() {
  let list = SALES_DATA;
  if (selectedKey) {
    const lk = selectedKey.includes('::') ? selectedKey.split('::')[0] : selectedKey;
    list = list.filter(d => d.lecture_key === lk);
  }
  if (platformFilter) {
    list = list.filter(d => d.플랫폼 === platformFilter);
  }
  if (searchQuery) {
    const q = searchQuery.toLowerCase();
    list = list.filter(d => d.강의그룹명.toLowerCase().includes(q) || d.강사.toLowerCase().includes(q));
  }
  // 날짜 범위 필터
  const dateFrom = document.getElementById('date-from').value;
  const dateTo   = document.getElementById('date-to').value;
  if (dateFrom || dateTo) {
    list = list.filter(d => d.rows.some(r => {
      const ds = normDateStr(r.무료강의일);
      if (!ds) return false;
      return (!dateFrom || ds >= dateFrom) && (!dateTo || ds <= dateTo);
    }));
    const el = document.getElementById('date-result');
    if (el) el.textContent = list.length + '개 강의 검색됨';
  } else {
    const el = document.getElementById('date-result');
    if (el) el.textContent = '';
  }
  if (currentSort === '날짜순') {
    return [...list].sort((a,b) => {
      const da = a.rows.map(r=>normDateStr(r.무료강의일)).filter(Boolean).sort().pop()||'';
      const db = b.rows.map(r=>normDateStr(r.무료강의일)).filter(Boolean).sort().pop()||'';
      return db.localeCompare(da);
    });
  }
  return [...list].sort((a,b) => b[currentSort] - a[currentSort]);
}

function renderList() {
  const list = getFiltered();
  const total = list.length;
  const totalPages = Math.ceil(total / PAGE_SIZE);
  const paged = list.slice((currentPage-1)*PAGE_SIZE, currentPage*PAGE_SIZE);
  document.getElementById('list-count').textContent = '총 ' + total.toLocaleString() + '개 강의';
  const tbody = document.getElementById('list-tbody');
  tbody.innerHTML = paged.map((d,i) => {
    const rank = (currentPage-1)*PAGE_SIZE + i + 1;
    const name = d.강의그룹명.length > 32 ? d.강의그룹명.slice(0,32)+'…' : d.강의그룹명;
    const safe = d.lecture_key.replace(/\\\\/g,'\\\\\\\\').replace(/'/g,"\\\\'");
    const kisuBadges = (d.기수List||[]).filter(k=>k!=='-').map(k=>'<span class="kisu-badge" style="margin-right:2px;">'+k+'</span>').join('') || '<span style="color:var(--muted);font-size:11px;">-</span>';
    const latestDate = d.rows.map(r=>normDateStr(r.무료강의일)).filter(Boolean).sort().pop()||'';
    const dateDisp = latestDate ? latestDate.slice(0,4)+'년 '+latestDate.slice(5,7)+'월 '+latestDate.slice(8,10)+'일' : '-';
    const mergeBadge = (d.mergedFrom && d.mergedFrom.length > 0) ? '<span class="merge-badge" title="'+d.mergedFrom.map(m=>m.강의그룹명).join(', ')+'">통합 '+(d.mergedFrom.length+1)+'개</span>' : '';
    return '<tr onclick="showDetail(\\'' + safe + '\\')">' +
      '<td class="rank">' + rank + '</td>' +
      '<td style="white-space:nowrap;font-size:12px;color:var(--muted);">' + dateDisp + '</td>' +
      '<td><div class="lec-name">' + name + mergeBadge + '<small>' + d.강사 + '</small></div></td>' +
      '<td>' + kisuBadges + '</td>' +
      '<td><span class="badge">' + d.플랫폼 + '</span></td>' +
      '<td class="num" style="font-weight:600;">' + fmt(d.강의총매출) + '원</td>' +
      '<td class="num">' + fmt(d.순매출) + '원</td>' +
      '<td class="num">' + d.수강생수.toLocaleString() + '명</td>' +
      '<td class="num">' + d.무료강의신청수.toLocaleString() + '명</td>' +
      '<td><span class="go-btn">상세 →</span></td></tr>';
  }).join('');
  const pg = document.getElementById('pagination');
  pg.innerHTML = '';
  if (totalPages <= 1) return;
  for (let p=1; p<=totalPages; p++) {
    if (totalPages > 10 && p > 3 && p < totalPages-2 && Math.abs(p-currentPage) > 2) {
      if (p===4||p===totalPages-3) pg.innerHTML += '<span style="color:var(--muted);padding:6px;">...</span>';
      continue;
    }
    pg.innerHTML += '<div class="pg-btn '+(p===currentPage?'active':'')+'" onclick="goPage('+p+')">'+p+'</div>';
  }
}

function goPage(p) { currentPage=p; renderList(); window.scrollTo(0,0); }

function showPage(name) {
  if (!_historyLock) {
    const state = { page: name };
    if (name === 'list') state.selectedKey = selectedKey;
    history.pushState(state, '');
  }
  _historyLock = false;
  document.querySelectorAll('.page').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.nav-tab').forEach(el => el.classList.remove('active'));
  document.getElementById('page-'+name).classList.add('active');
  document.getElementById('tab-'+name).classList.add('active');
  if (name==='list') renderList();
  if (name==='home') hidePlatformPicker();
}

// ════════════════════════════════════════════════════
//  상세 페이지
// ════════════════════════════════════════════════════
// ── 광고비 수동 입력 (독립 입력값, 기존 계산 로직과 무관) ──
let _manualAdcost = 0;
function onAdcostInput(raw) {
  const n = Number(String(raw).replace(/[^0-9]/g, '')) || 0;
  _manualAdcost = n;
  const adcostEl = document.getElementById('kpi-adcost');
  if (adcostEl) {
    if (n > 0) {
      adcostEl.innerHTML = '<span style="font-size:20px;font-weight:800;color:var(--text);">' + fmt(n) + '</span><span style="font-size:12px;color:var(--muted);margin-left:2px;">원</span>';
    } else {
      adcostEl.innerHTML = '<span style="font-size:16px;font-weight:600;color:var(--muted);">-</span>';
    }
  }
  // ROAS 표시 업데이트 (현재 총매출 기준, 기존 kd 접근용)
  const roasEl = document.getElementById('kpi-roas');
  const roasSubEl = document.getElementById('kpi-roas-sub');
  if (roasEl && window._currentKd) {
    const totalRev = window._currentKd.강의총매출 || 0;
    const adcostForRoas = n > 0 ? n : (window._scheduleAdcost || window._currentKd.광고비 || 0);
    const roas = adcostForRoas > 0 ? (totalRev / adcostForRoas) : 0;
    roasEl.textContent = roas > 0 ? roas.toFixed(1) : '-';
    if (roasSubEl) roasSubEl.textContent = roas > 0 ? '광고비 대비 매출' : '광고비 없음';
    const _rt = document.getElementById('kpi-top-roas');
    if (_rt) _rt.innerHTML = roas > 0 ? (roas*100).toFixed(0)+'<span class="kpi-unit">%</span>' : '-';
  }
}

function showDetail(combinedKey) {
  const sepIdx = combinedKey ? combinedKey.indexOf('::') : -1;
  const lecture_key = sepIdx >= 0 ? combinedKey.slice(0, sepIdx) : combinedKey;
  const kisuFilter  = sepIdx >= 0 ? combinedKey.slice(sepIdx + 2) : null;

  const d = SALES_DATA.find(x => x.lecture_key === lecture_key);
  if (!d) return;
  // 강의 변경 시 날짜 피커 초기화 (자동 채우기가 재실행되도록)
  if (currentDetailKey !== combinedKey) {
    const _c = (id) => { const e = document.getElementById(id); if(e) e.value = ''; };
    _c('range-start-date'); _c('range-end-date');
  }
  currentDetailKey = combinedKey;
  selectedKey = combinedKey;
  document.getElementById('global-filter').value = combinedKey;

  // 표시할 rows (기수 필터 적용)
  const dispRows = kisuFilter ? d.rows.filter(r => r.기수 === kisuFilter) : d.rows;
  // KPI 소스: 기수 필터 시 해당 row 합산, 아니면 전체 d
  const kd = kisuFilter && dispRows.length > 0 ? {
    강의총매출:   dispRows.reduce((s,r) => s+r.강의총매출, 0),
    순매출:       dispRows.reduce((s,r) => s+r.순매출, 0),
    수강생수:     dispRows.reduce((s,r) => s+(r.수강생수||0), 0),
    무료강의신청수: dispRows.reduce((s,r) => Math.max(s, r.무료강의신청수||0), 0),
    PG제외매출:   dispRows.reduce((s,r) => s+r.PG제외매출, 0),
    광고비:       dispRows.reduce((s,r) => s+(r.광고비||0), 0),
    강사정산금:   dispRows.reduce((s,r) => s+(r.강사정산금||0), 0),
  } : d;
  window._currentKd = kd;  // 광고비 입력 시 ROAS 계산에 사용

  const kisuStr = kisuFilter && kisuFilter !== '-' ? kisuFilter : (d.기수List||[]).filter(k=>k!=='-').join(', ');
  document.getElementById('detail-name').textContent = d.강의그룹명 + (kisuFilter && kisuFilter !== '-' ? ' — ' + kisuFilter : '');
  document.getElementById('detail-meta').innerHTML =
    '<span>강사: '+d.강사+'</span>' +
    '<span>플랫폼: '+d.플랫폼+'</span>' +
    (kisuStr ? '<span>기수: '+kisuStr+'</span>' : '') +
    '<span>'+dispRows.length+'개 기수 데이터</span>';
  // 무료웨비나 링크 표시 (SCHEDULE_DATA에서 강의명 매칭)
  (function() {
    const wEl = document.getElementById('detail-webinar-link');
    if (!wEl) return;
    const normKey = normalizeLectureName(d.강의그룹명);
    let link = '';
    if (SCHEDULE_DATA && SCHEDULE_DATA.length > 0) {
      // 1) 정규화 후 정확 매칭
      let matched = SCHEDULE_DATA.find(s => normalizeLectureName(s.강의명) === normKey && s.무료웨비나링크);
      // 2) 부분 문자열 매칭 (긴 쪽이 짧은 쪽 포함)
      if (!matched) matched = SCHEDULE_DATA.find(s => {
        const sn = normalizeLectureName(s.강의명);
        return s.무료웨비나링크 && (sn.includes(normKey) || normKey.includes(sn));
      });
      // 3) 링크 없어도 매칭 강의 탐색 (링크 없음 표시용)
      if (!matched) matched = SCHEDULE_DATA.find(s => normalizeLectureName(s.강의명) === normKey);
      if (!matched) matched = SCHEDULE_DATA.find(s => {
        const sn = normalizeLectureName(s.강의명);
        return sn.includes(normKey) || normKey.includes(sn);
      });
      if (matched) link = matched.무료웨비나링크 || '';
    }
    if (SCHEDULE_DATA && SCHEDULE_DATA.length > 0) {
      if (link) {
        wEl.innerHTML = '<span style="color:var(--muted);margin-right:6px;">무료웨비나</span>' +
          '<a href="' + link.replace(/"/g,'&quot;') + '" target="_blank" rel="noopener noreferrer" ' +
          'style="color:var(--cyan);text-decoration:underline;word-break:break-all;">' + link + '</a>';
      } else {
        wEl.innerHTML = '<span style="color:var(--muted);margin-right:6px;">무료웨비나</span>' +
          '<span style="color:var(--muted);">데이터 없음</span>';
      }
    } else {
      wEl.innerHTML = '';
    }
  })();
  // 통합 강의 패널
  const mergePanel = document.getElementById('merge-panel');
  if (d.mergedFrom && d.mergedFrom.length > 0 && !kisuFilter) {
    const allItems = [{ lecture_key: d.lecture_key, 강의그룹명: d.강의그룹명, isPrimary: true }]
      .concat(d.mergedFrom.map(m => ({ ...m, isPrimary: false })));
    mergePanel.style.display = '';
    mergePanel.innerHTML =
      '<div class="merge-panel-title">⚡ 통합 강의 — ERP 중복 항목 ' + allItems.length + '개가 자동으로 통합되었습니다</div>' +
      allItems.map(item => {
        const safePK = d.lecture_key.replace(/'/g,"\\'");
        const safeSK = item.lecture_key.replace(/'/g,"\\'");
        const unlinkBtn = !item.isPrimary
          ? '<button class="merge-unlink-btn" onclick="event.stopPropagation();excludeMerge(\\'' + safePK + '\\',\\'' + safeSK + '\\')">통합 해제</button>'
          : '';
        return '<div class="merge-item">' +
          '<span class="merge-item-name">' + item.강의그룹명 + '</span>' +
          '<span class="merge-item-tag ' + (item.isPrimary ? 'primary' : 'merged') + '">' + (item.isPrimary ? '기준' : '통합') + '</span>' +
          unlinkBtn + '</div>';
      }).join('');
  } else {
    mergePanel.style.display = 'none';
  }

  // KPI (기존 hidden 요소 유지)
  document.getElementById('kpi-students').innerHTML = fmtUnit(kd.수강생수,'명');
  document.getElementById('kpi-free').innerHTML = fmtUnit(kd.무료강의신청수,'명');
  // 새 Row1
  document.getElementById('kpi-total').innerHTML = fmtFull(kd.강의총매출)+'<span class="kpi-unit">원</span>';
  document.getElementById('kpi-top-free').innerHTML = (kd.무료강의신청수||0).toLocaleString()+'<span class="kpi-unit">명</span>';
  document.getElementById('kpi-top-students').innerHTML = (kd.수강생수||0).toLocaleString()+'<span class="kpi-unit">명</span>';
  // 전환지표/재무 항목은 renderFunnel 후 _renderKeyMetrics에서 업데이트
  ['kpi-top-finalconv','kpi-top-target','kpi-top-achieve','kpi-top-roas',
   'kpi-conv-tok','kpi-conv-live','kpi-conv-livepay','kpi-conv-encore','kpi-conv-encorepay'].forEach(id => {
    const el = document.getElementById(id); if(el) el.textContent = '-';
  });
  const cvr = kd.무료강의신청수 > 0 ? (kd.수강생수/kd.무료강의신청수*100) : 0;
  document.getElementById('kpi-cvr').innerHTML =
    '<span style="font-size:28px;font-weight:800;letter-spacing:-.02em;">'+cvr.toFixed(2)+'</span><span class="kpi-unit">%</span>';
  document.getElementById('kpi-cvr-sub').textContent = '신청 '+kd.무료강의신청수.toLocaleString()+'명 → 결제 '+kd.수강생수+'명';
  // 영업이익
  document.getElementById('kpi-net').innerHTML = '<span style="font-size:22px;font-weight:800;">'+fmt(kd.순매출)+'</span><span style="font-size:12px;color:var(--muted);margin-left:2px;">원</span>';
  const margin = kd.강의총매출 > 0 ? (kd.순매출/kd.강의총매출*100) : 0;
  document.getElementById('kpi-margin').innerHTML = margin.toFixed(2)+'<span>%</span>';
  document.getElementById('profit-bar').style.width = Math.min(Math.max(margin,0),100)+'%';
  document.getElementById('profit-bar-label').textContent = '이익 '+margin.toFixed(2)+'%';
  // 광고
  const roas = kd.광고비 > 0 ? (kd.강의총매출/kd.광고비) : 0;
  document.getElementById('kpi-roas').textContent = roas > 0 ? roas.toFixed(1) : '-';
  document.getElementById('kpi-roas-sub').textContent = roas > 0 ? '광고비 대비 매출' : '광고비 없음';
  const _roasTopEl = document.getElementById('kpi-top-roas');
  if (_roasTopEl) _roasTopEl.innerHTML = roas > 0 ? (roas*100).toFixed(0)+'<span class="kpi-unit">%</span>' : '-';
  document.getElementById('kpi-adcost').innerHTML = kd.광고비 > 0
    ? '<span style="font-size:20px;font-weight:800;color:var(--text);">'+fmt(kd.광고비)+'</span><span style="font-size:12px;color:var(--muted);margin-left:2px;">원</span>'
    : '<span style="font-size:16px;font-weight:600;color:var(--muted);">-</span>';
  // 광고비 입력 필드 초기화 (강의 전환 시 리셋)
  _manualAdcost = 0;
  const _adInput = document.getElementById('adcost-input');
  if (_adInput) { _adInput.value = ''; }
  document.getElementById('kpi-teacher').innerHTML =
    '<span style="font-size:20px;font-weight:800;color:var(--text);">'+fmt(kd.강사정산금)+'</span><span style="font-size:12px;color:var(--muted);margin-left:2px;">원</span>';

  renderKisuTable(dispRows);
  renderFunnel(combinedKey, kd).then(() => {
    // renderFunnel 내부에서 _uniquePayerCounts가 설정됨 → KPI 수강생수/CVR 업데이트
    const upc = window._uniquePayerCounts || {};
    if (upc.total !== null && upc.total !== undefined) {
      const uniqueStudents = upc.total;
      document.getElementById('kpi-students').innerHTML = fmtUnit(uniqueStudents,'명');
      document.getElementById('kpi-top-students').innerHTML = uniqueStudents.toLocaleString()+'<span class="kpi-unit">명</span>';
      const uCvr = kd.무료강의신청수 > 0 ? (uniqueStudents/kd.무료강의신청수*100) : 0;
      document.getElementById('kpi-cvr').innerHTML =
        '<span style="font-size:28px;font-weight:800;letter-spacing:-.02em;">'+uCvr.toFixed(2)+'</span><span class="kpi-unit">%</span>';
      document.getElementById('kpi-cvr-sub').textContent = '신청 '+kd.무료강의신청수.toLocaleString()+'명 → 결제 '+uniqueStudents+'명 (중복제거)';
    }
    // 광고비 기본값: 시트의 '광고집행+출연료' → 입력란 제외하고 표시·계산에만 적용
    const schedAdcost = window._scheduleAdcost;
    if (schedAdcost && schedAdcost > 0 && _manualAdcost === 0) {
      const adcostEl = document.getElementById('kpi-adcost');
      if (adcostEl) adcostEl.innerHTML = '<span style="font-size:20px;font-weight:800;color:var(--text);">' + fmt(schedAdcost) + '</span><span style="font-size:12px;color:var(--muted);margin-left:2px;">원</span>';
      const totalRev = window._currentKd ? (window._currentKd.강의총매출 || 0) : 0;
      const roas = totalRev / schedAdcost;
      const roasEl = document.getElementById('kpi-roas');
      const roasSubEl = document.getElementById('kpi-roas-sub');
      if (roasEl) roasEl.textContent = roas.toFixed(1);
      if (roasSubEl) roasSubEl.textContent = '광고비 대비 매출';
      const _rt = document.getElementById('kpi-top-roas');
      if (_rt) _rt.innerHTML = (roas*100).toFixed(0)+'<span class="kpi-unit">%</span>';
    }
  });
  renderTimeChart(combinedKey);
  if (!_historyLock) history.pushState({ page: 'detail', key: combinedKey }, '');
  _historyLock = true; // showPage('detail')가 중복 push하지 않도록
  showPage('detail');
  updateDetailStickyTop();
}

// ════════════════════════════════════════════════════
//  시간별 구매 추이 차트
// ════════════════════════════════════════════════════
function reRenderTimeChart() { renderTimeChart(currentDetailKey); }

function renderTimeChart(combinedKey) {
  const intervalMin = parseInt(document.getElementById('interval-select').value) || 10;
  const intervalMs  = intervalMin * 60000;
  const sub = document.getElementById('time-chart-sub');

  const sepIdx = combinedKey ? combinedKey.indexOf('::') : -1;
  const key        = sepIdx >= 0 ? combinedKey.slice(0, sepIdx) : combinedKey;
  const kisuFilter = sepIdx >= 0 ? combinedKey.slice(sepIdx + 2) : null;

  // ── 강의 주문 매칭 ──
  let orders = ORDER_DATA;
  let matchMode = 'all';
  let lectureDates = []; // YYYY-MM-DD

  if (key && ORDER_DATA.length > 0) {
    const d = SALES_DATA.find(x => x.lecture_key === key);
    if (d) {
      // 기수 필터 적용된 날짜만
      const filteredRows = kisuFilter ? d.rows.filter(r => r.기수 === kisuFilter) : d.rows;
      lectureDates = filteredRows.map(r => normDateStr(r.무료강의일)).filter(Boolean);
      const normGroup = d.강의그룹명.toLowerCase().replace(/\\s+/g,' ').trim();
      const matched = ORDER_DATA.filter(o => {
        const normOrder = normalizeLectureName(o.lectureName).toLowerCase().replace(/\\s+/g,' ').trim();
        return normOrder === normGroup || normOrder.includes(normGroup) || normGroup.includes(normOrder)
          || (normGroup.length >= 8 && normOrder.startsWith(normGroup.slice(0, Math.min(normGroup.length, 12))));
      });
      if (matched.length > 0) { orders = matched; matchMode = 'matched'; }
      else {
        orders = ORDER_DATA; matchMode = 'fallback';
        console.warn('[시간차트] 강의 매칭 실패. 검색어:', normGroup, '| 샘플:', ORDER_DATA.slice(0,3).map(o=>o.lectureName));
      }
    }
  }

  if (orders.length === 0) {
    sub.textContent = '주문 결제 데이터를 업로드하면 표시됩니다';
    if (charts.time) { charts.time.destroy(); charts.time = null; }
    renderOrderList([]);
    return;
  }

  // ── 날짜+시간 범위 피커 자동 채우기 (강의 변경 시 lecture date 기본값) ──
  const _sdEl = document.getElementById('range-start-date');
  const _edEl = document.getElementById('range-end-date');
  if (_sdEl && !_sdEl.value && lectureDates.length > 0) {
    const firstD = [...lectureDates].sort()[0];
    _sdEl.value = firstD;
    const lastD = [...lectureDates].sort().pop();
    // 종료 시간이 00:00~12:00 범위면 익일로 자동 설정
    const _etEl = document.getElementById('range-end-time');
    const etVal = (_etEl && _etEl.value) || '03:00';
    const [eh] = etVal.split(':').map(Number);
    const [stEl] = ((document.getElementById('range-start-time')||{}).value||'19:30').split(':').map(Number);
    const endDateObj = new Date(lastD);
    if (eh <= stEl) endDateObj.setDate(endDateObj.getDate() + 1); // 익일
    if (_edEl) _edEl.value = endDateObj.toISOString().slice(0,10);
  }

  // ── 절대 날짜+시간 범위 파싱 ──
  const _getAbsMs = (dateId, timeId, defTime) => {
    const dv = (document.getElementById(dateId)||{}).value;
    const tv = (document.getElementById(timeId)||{}).value || defTime;
    if (!dv) return null;
    const [yr, mo, dy] = dv.split('-').map(Number);
    const [h, m] = tv.split(':').map(Number);
    return new Date(yr, mo-1, dy, isNaN(h)?0:h, isNaN(m)?0:m, 0).getTime();
  };
  const rangeStartMs = _getAbsMs('range-start-date', 'range-start-time', '19:30');
  const rangeEndMs   = _getAbsMs('range-end-date',   'range-end-time',   '03:00');

  // ── 조회 범위 필터 ──
  const hasRange = rangeStartMs !== null && rangeEndMs !== null;
  if (hasRange) {
    orders = orders.filter(o => {
      const t = o.ts.getTime();
      return t >= rangeStartMs && t < rangeEndMs;
    });
  }

  if (matchMode === 'fallback') {
    sub.textContent = '⚠️ 강의 매칭 실패 — 전체 주문 '+orders.length+'건 표시 중 (F12 콘솔에서 강의명 확인)';
    sub.style.color = 'var(--yellow)';
  } else {
    sub.textContent = (matchMode==='matched'?'매칭':'전체')+' '+orders.length+'건 / '+intervalMin+'분 단위';
    sub.style.color = '';
  }
  updateOrderCountBadge(orders.length);

  // ── epoch ms 버킷 ──
  const buckets = {};
  orders.forEach(o => {
    const bMs = Math.floor(o.ts.getTime() / intervalMs) * intervalMs;
    if (!buckets[bMs]) buckets[bMs] = { count:0, amount:0, firstTs: o.ts };
    else if (o.ts < buckets[bMs].firstTs) buckets[bMs].firstTs = o.ts;
    buckets[bMs].count++;
    buckets[bMs].amount += o.amount;
  });

  // ── X축 시작: 범위 피커 값 우선, 없으면 강의일 19:30 fallback ──
  let startMs;
  if (rangeStartMs !== null) {
    startMs = Math.floor(rangeStartMs / intervalMs) * intervalMs;
  } else if (lectureDates.length > 0) {
    const firstDate = [...lectureDates].sort()[0];
    const [yr, mo, dy] = firstDate.split('-').map(Number);
    startMs = Math.floor(new Date(yr, mo-1, dy, 19, 30, 0).getTime() / intervalMs) * intervalMs;
  } else {
    const o = orders[0];
    startMs = Math.floor(new Date(o.ts.getFullYear(), o.ts.getMonth(), o.ts.getDate(), 19, 30, 0).getTime() / intervalMs) * intervalMs;
  }
  const existingKeys = Object.keys(buckets).map(Number).sort((a,b)=>a-b);
  // ── X축 끝: 범위 피커 종료일 기준, 날짜 없으면 마지막 데이터 기준 ──
  let endMs;
  if (rangeEndMs !== null) {
    endMs = Math.max(startMs, Math.floor((rangeEndMs - 1) / intervalMs) * intervalMs);
  } else {
    endMs = existingKeys.length > 0 ? Math.max(startMs, existingKeys[existingKeys.length-1]) : startMs;
  }

  // startMs ~ endMs 전 구간 채움 (0 포함)
  const allKeys = [];
  for (let k = startMs; k <= endMs; k += intervalMs) allKeys.push(k);

  const dtLabel = ms => {
    const d = new Date(ms);
    return (d.getMonth()+1) + '/' + pad2(d.getDate()) + ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes());
  };

  const labels   = allKeys.map(k => dtLabel(k));
  const counts   = allKeys.map(k => buckets[k] ? buckets[k].count  : 0);
  const amounts  = allKeys.map(k => buckets[k] ? buckets[k].amount : 0);
  const firstTss = allKeys.map(k => buckets[k] ? buckets[k].firstTs : null);
  const total   = orders.length;

  if (charts.time) charts.time.destroy();
  charts.time = new Chart(document.getElementById('chart-time'), {
    type: 'line',
    data: { labels, datasets: [{
      label: '구매 건수', data: counts,
      borderColor: '#A3C244', backgroundColor: 'rgba(163,194,68,0.12)',
      tension: 0.4, fill: true, pointRadius: 4, pointHoverRadius: 7,
    }]},
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { labels:{ color:'#94a3b8', font:{size:11}, boxWidth:12 } },
        tooltip: { callbacks: {
          title: ctx => dtLabel(allKeys[ctx[0].dataIndex]) + ' ~ ' + dtLabel(allKeys[ctx[0].dataIndex] + intervalMs),
          label: ctx => {
            const cnt = ctx.raw, pct = total > 0 ? (cnt/total*100).toFixed(1) : '0.0';
            const ft = firstTss[ctx.dataIndex];
            const elapsedLine = ft ? '첫 결제 경과: ' + fmtElapsed(ft, getElapsedOffset()) : null;
            const lines = ['구매 건수: '+cnt+'건', '전체 대비: '+pct+'%', '구간 매출: '+fmtFull(amounts[ctx.dataIndex])+'원'];
            if (elapsedLine) lines.push(elapsedLine);
            return lines;
          }
        }}
      },
      scales: {
        x: { ticks:{ color:'#94a3b8', font:{size:10}, maxRotation:45 }, grid:{color:'rgba(255,255,255,0.04)'} },
        y: { ticks:{ color:'#94a3b8', font:{size:10} }, grid:{color:'rgba(255,255,255,0.04)'}, beginAtZero:true }
      }
    }
  });

  renderOrderList(orders);
}

// ════════════════════════════════════════════════════
//  구매 상세 리스트 토글
// ════════════════════════════════════════════════════
function updateOrderCountBadge(n) {
  document.getElementById('order-count-badge').textContent = n > 0 ? n + '건' : '';
}

function toggleOrderList() {
  const wrap = document.getElementById('order-list-wrap');
  const btn  = document.getElementById('order-toggle');
  const arr  = document.getElementById('order-arr');
  const open = wrap.classList.toggle('open');
  btn.classList.toggle('open', open);
  arr.textContent = open ? '▲' : '▼';
}

function renderOrderList(orders) {
  const el = document.getElementById('order-list');
  if (orders.length === 0) {
    el.innerHTML = '<div class="no-order">해당 강의의 주문 데이터가 없습니다</div>';
    updateOrderCountBadge(0);
    return;
  }
  updateOrderCountBadge(orders.length);
  el.innerHTML = orders.map(o =>
    '<div class="order-row">' +
    '<span class="order-time">'+fmtDateTime(o.ts)+'</span>' +
    '<span class="order-elapsed">'+fmtElapsed(o.ts, getElapsedOffset())+'</span>' +
    '<span class="order-amount">'+fmtFull(o.amount)+'원</span>' +
    '<span class="order-buyer">'+(o.name||'-')+'</span>' +
    '<span class="order-phone">'+(o.phone||'-')+'</span>' +
    '<span class="order-name">'+o.lectureName+'</span>' +
    '</div>'
  ).join('');
}

// ════════════════════════════════════════════════════
//  기수별 테이블
// ════════════════════════════════════════════════════
function renderKisuTable(rows) {
  const sorted = [...rows].sort((a,b) => a.무료강의일.localeCompare(b.무료강의일));
  document.getElementById('kisu-tbody').innerHTML = sorted.map(r =>
    '<tr>' +
    '<td><span class="kisu-badge">'+(r.기수||'-')+'</span></td>' +
    '<td>'+(r.무료강의일||'-')+'</td>' +
    '<td><span class="badge">'+r.플랫폼+'</span></td>' +
    '<td class="num">'+fmtFull(r.강의총매출)+'원</td>' +
    '<td class="num">'+fmtFull(r.순매출)+'원</td>' +
    '<td class="num">'+r.수강생수+'명</td>' +
    '<td class="num">'+r.무료강의신청수.toLocaleString()+'명</td>' +
    '</tr>'
  ).join('');
}

// ════════════════════════════════════════════════════
//  성과 퍼널
// ════════════════════════════════════════════════════
const FUNNEL_DEFAULTS = { 라이브참여율:30, 라이브전환율:10, 앵콜참여율:20, 앵콜전환율:10, 환불율:10, 최종결제율:10 };
const FUNNEL_STAGES = [
  { key:'총신청자',    label:'총 신청자',      convKey:null,           tgtKey:null,           aarMsg:'' },
  { key:'단톡방유입',  label:'단톡방 유입',    convKey:'유입률',       tgtKey:null,           aarMsg:'타겟 설정 점검' },
  { key:'라이브참여자',label:'라이브 참여자',  convKey:'참여율',       tgtKey:'라이브참여율', aarMsg:'알림/제목 점검' },
  { key:'라이브결제자',label:'라이브 결제자',  convKey:'라이브전환율', tgtKey:'라이브전환율', aarMsg:'' },
  { key:'앵콜입장자',  label:'앵콜 입장자',    convKey:'앵콜참여율',   tgtKey:'앵콜참여율',   aarMsg:'리마케팅 점검' },
  { key:'앵콜결제자',  label:'앵콜 결제자',    convKey:'앵콜전환율',   tgtKey:'앵콜전환율',   aarMsg:'혜택/Q&A 점검' },
  { key:'환불자',      label:'환불자',          convKey:'환불율',       tgtKey:'환불율',       aarMsg:'환불 사유 확인' },
  { key:'최종결제자',  label:'최종 결제자',    convKey:'최종결제율',   tgtKey:'최종결제율',   aarMsg:'매출 기여도 점검' },
];
let _fKey = '', _fSaved = {}, _fKd = null;

function idbSaveFunnel(key, data) {
  idbOpen().then(db => { const tx = db.transaction('funnel','readwrite'); tx.objectStore('funnel').put(data,key); }).catch(()=>{});
}
function idbLoadFunnel(key) {
  return idbOpen().then(db => new Promise(resolve => {
    const r = db.transaction('funnel','readonly').objectStore('funnel').get(key);
    r.onsuccess = e => resolve(e.target.result||null); r.onerror = ()=>resolve(null);
  })).catch(()=>null);
}

function _funnelErpDefaults(kd, combinedKey) {
  const sepIdx = combinedKey ? combinedKey.indexOf('::') : -1;
  const lk = sepIdx>=0 ? combinedKey.slice(0,sepIdx) : combinedKey;
  const kf = sepIdx>=0 ? combinedKey.slice(sepIdx+2) : null;
  const d  = SALES_DATA.find(x=>x.lecture_key===lk);
  const rows = d ? (kf ? d.rows.filter(r=>r.기수===kf) : d.rows) : [];

  // ── 시트(SCHEDULE_DATA)에서 톡방인원, 라이브참여자 매핑 ──
  let 톡방인원 = null, 라이브참여자 = null;
  let refundCnt = null;
  window._funnelScheduleSource = {};  // 매핑 출처 추적
  window._scheduleTargetRevenue = null;
  window._scheduleROAS = null;
  window._scheduleAdcost = null;
  // ── 강의명 유사도 비교용: 따옴표/공백/특수문자/괄호 제거 후 비교 ──
  function _normForMatch(s) {
    if (!s) return '';
    return normalizeLectureName(s).toLowerCase()
      .replace(/[\\s'"''""\`\\(\\)\\[\\]\\{\\}\\-_.,!?@#\\$%\\^&\\*:;\\/\\\\~+|<>=]/g, '');
  }
  // 바이그램 Jaccard 유사도 (0~1)
  function _bigram(s) {
    const r = new Set();
    for (let i = 0; i < s.length - 1; i++) r.add(s.slice(i, i + 2));
    return r;
  }
  function _simScore(a, b) {
    if (a === b) return 1;
    if (!a || !b) return 0;
    if (a.includes(b) || b.includes(a)) return 0.9;
    if (a.length < 2 || b.length < 2) return 0;
    const ba = _bigram(a), bb = _bigram(b);
    let common = 0;
    ba.forEach(t => { if (bb.has(t)) common++; });
    const union = ba.size + bb.size - common;
    return union > 0 ? common / union : 0;
  }

  // ── 주문 파일 기반: 유니크 결제자 수 (이름+전화번호 중복 제거) ──
  const ng = d ? _normForMatch(d.강의그룹명) : '';
  const paidSet = new Set();       // 전체 유니크 결제자
  const livePayerSet = new Set();  // 라이브(19:30~23:00) 유니크 결제자
  const encorePayerSet = new Set(); // 앵콜(23:00~03:00) 유니크 결제자
  window._uniquePayerCounts = { total: null, live: null, encore: null };

  if (ORDER_DATA.length > 0 && d) {
    const erpDateList = rows.map(r => normDateStr(r.무료강의일)).filter(Boolean);
    ORDER_DATA.forEach(o => {
      const no = _normForMatch(o.lectureName);
      if (!no || !(no === ng || no.includes(ng) || ng.includes(no) || _simScore(no, ng) >= 0.5)) return;
      const pkey = (o.name||'').trim() + '|||' + (o.phone||'').replace(/[^0-9]/g,'');
      if (pkey === '|||') return;
      paidSet.add(pkey);
      // 시간대별 분류 (라이브 19:30~23:00, 앵콜 23:00~03:00)
      if (o.ts && erpDateList.length > 0) {
        for (const ds of erpDateList) {
          const [yr, mo, dy] = ds.split('-').map(Number);
          const liveStart = new Date(yr, mo-1, dy, 19, 30, 0);
          const liveEnd   = new Date(yr, mo-1, dy, 23, 0, 0);
          const encoreEnd = new Date(yr, mo-1, dy+1, 3, 0, 0);
          if (o.ts >= liveStart && o.ts < liveEnd) { livePayerSet.add(pkey); break; }
          if (o.ts >= liveEnd && o.ts < encoreEnd) { encorePayerSet.add(pkey); break; }
        }
      }
    });
    if (paidSet.size > 0) {
      window._uniquePayerCounts = { total: paidSet.size, live: livePayerSet.size, encore: encorePayerSet.size };
      console.log('[결제자 중복제거] 유니크 결제자:', paidSet.size, '명 (라이브:', livePayerSet.size, '/ 앵콜:', encorePayerSet.size, ')');
    }
  }

  // ── 환불자: ERP 주문 파일의 환불/취소 건수 (재결제자 제외) ──
  const refList = window._REFUND_LIST || [];
  if (refList.length > 0 && d) {
    // 이 강의의 환불 건 필터
    const matchedRefunds = refList.filter(r => {
      const nr = _normForMatch(r.lectureName);
      return nr && (nr === ng || nr.includes(ng) || ng.includes(nr) || _simScore(nr, ng) >= 0.5);
    });
    // 환불자 중 동일 이름+전화번호로 재결제한 사람 제외
    const netRefunds = matchedRefunds.filter(r => {
      const key = (r.name||'').trim() + '|||' + (r.phone||'').replace(/[^0-9]/g,'');
      if (key === '|||') return true;  // 식별 불가 시 환불로 유지
      return !paidSet.has(key);
    });
    if (matchedRefunds.length > 0) {
      refundCnt = netRefunds.length;
      console.log('[퍼널 매핑] 환불:', matchedRefunds.length, '건 → 재결제:', matchedRefunds.length - netRefunds.length, '건 제외 → 순환불자:', refundCnt, '건');
    }
  }

  // ERP 날짜 목록
  const erpDates = rows.map(r => normDateStr(r.무료강의일)).filter(Boolean);
  const normLk = _normForMatch(d.강의그룹명);

  console.log('[퍼널 매핑] SCHEDULE_DATA:', SCHEDULE_DATA.length, '행 | ERP 강의:', d.강의그룹명, '→', normLk, '| 날짜:', erpDates);

  if (SCHEDULE_DATA.length > 0 && d) {
    let matchRows = [];
    let matchMethod = '';

    // 시트 날짜 정규화 함수
    function _normSchDate(s) {
      if (!s) return '';
      // 다양한 형식 지원: 26.3.25, 2026-03-25, 2026.03.25, 26/3/25, 3/25 등
      let ds = String(s).replace(/[\\(\\)가-힣]/g,'').trim();
      // YY.M.D 또는 YYYY.M.D
      let m = ds.match(/(\\d{2,4})[\\.\\/\\-](\\d{1,2})[\\.\\/\\-](\\d{1,2})/);
      if (m) {
        let yr = Number(m[1]);
        if (yr < 100) yr += 2000;
        return yr + '-' + String(m[2]).padStart(2,'0') + '-' + String(m[3]).padStart(2,'0');
      }
      return normDateStr(s);
    }

    // 1차: 날짜 동일 + 강의명 유사도 ≥ 0.5
    SCHEDULE_DATA.forEach(sr => {
      const schDate = _normSchDate(sr.강의일);
      const dateMatch = erpDates.includes(schDate);
      if (!dateMatch) return;
      const normSr = _normForMatch(sr.강의명);
      const sim = _simScore(normLk, normSr);
      if (sim >= 0.5) {
        matchRows.push({ row: sr, sim, method: '날짜+강의명(sim=' + sim.toFixed(2) + ')' });
      }
    });
    if (matchRows.length > 0) matchMethod = '날짜+강의명';
    console.log('[퍼널 매핑] 1차(날짜+강의명유사) 매칭:', matchRows.length);

    // 2차: 강의명 유사도 ≥ 0.6 (날짜 없이)
    if (matchRows.length === 0) {
      SCHEDULE_DATA.forEach(sr => {
        const normSr = _normForMatch(sr.강의명);
        const sim = _simScore(normLk, normSr);
        if (sim >= 0.6) {
          matchRows.push({ row: sr, sim, method: '강의명only(sim=' + sim.toFixed(2) + ')' });
        }
      });
      if (matchRows.length > 0) matchMethod = '강의명유사';
      console.log('[퍼널 매핑] 2차(강의명유사) 매칭:', matchRows.length);
    }

    // 3차: 날짜 동일 + 강사명 동일
    if (matchRows.length === 0) {
      const erpInst = (d.강사||'').trim();
      if (erpInst) {
        SCHEDULE_DATA.forEach(sr => {
          const schDate = _normSchDate(sr.강의일);
          const dateMatch = erpDates.includes(schDate);
          if (!dateMatch) return;
          if ((sr.강사||'').trim() === erpInst) {
            matchRows.push({ row: sr, sim: 1, method: '날짜+강사명' });
          }
        });
        if (matchRows.length > 0) matchMethod = '날짜+강사명';
        console.log('[퍼널 매핑] 3차(날짜+강사명) 매칭:', matchRows.length);
      }
    }

    // 4차: 강사명만
    if (matchRows.length === 0) {
      const erpInst = (d.강사||'').trim();
      if (erpInst) {
        SCHEDULE_DATA.forEach(sr => {
          if ((sr.강사||'').trim() === erpInst) {
            matchRows.push({ row: sr, sim: 1, method: '강사명only' });
          }
        });
        if (matchRows.length > 0) matchMethod = '강사명';
        console.log('[퍼널 매핑] 4차(강사명only) 매칭:', matchRows.length);
      }
    }

    if (matchRows.length > 0) {
      // 유사도 높은 순 정렬
      matchRows.sort((a, b) => b.sim - a.sim);
      const best = matchRows[0];
      const sr = best.row;
      console.log('[퍼널 매핑] ✅ 매칭 성공 [' + best.method + ']:', JSON.stringify(sr));
      if (sr.톡방인원 && sr.톡방인원 !== '' && !isNaN(Number(sr.톡방인원))) {
        톡방인원 = Number(sr.톡방인원);
        window._funnelScheduleSource.단톡방유입 = '톡방인원';
      }
      if (sr.라이브참여자 && sr.라이브참여자 !== '' && !isNaN(Number(sr.라이브참여자))) {
        라이브참여자 = Number(sr.라이브참여자);
        window._funnelScheduleSource.라이브참여자 = '라이브참여자';
      }
      const _parseKoreanNum = (raw) => {
        if (!raw && raw !== 0) return null;
        let s = String(raw).replace(/[\\s₩,원]/g, '');
        if (!s) return null;
        // 억 단위: "1.5억" or "15000만"
        let v = null;
        const 억m = s.match(/^([\d.]+)억([\d.]*만)?([\d.]*)?$/);
        const 만m = s.match(/^([\d.]+)만([\d.]*)?$/);
        if (억m) {
          v = parseFloat(억m[1]) * 100000000
            + (억m[2] ? parseFloat(억m[2]) * 10000 : 0)
            + (억m[3] ? parseFloat(억m[3]) : 0);
        } else if (만m) {
          v = parseFloat(만m[1]) * 10000 + (만m[2] ? parseFloat(만m[2]) : 0);
        } else {
          const n = Number(s);
          if (!isNaN(n)) v = n;
        }
        console.log('[목표매출 파싱] 원본:', raw, '→ 결과:', v);
        return v;
      };
      const parsedTarget = _parseKoreanNum(sr.목표매출);
      if (parsedTarget !== null && parsedTarget > 0) {
        window._scheduleTargetRevenue = parsedTarget;
      }
      const parsedROAS = _parseKoreanNum(sr.ROAS);
      if (parsedROAS !== null && parsedROAS > 0) {
        window._scheduleROAS = parsedROAS;
      }
      const parsedAdcost = _parseKoreanNum(sr.광고집행출연료);
      if (parsedAdcost !== null && parsedAdcost > 0) {
        window._scheduleAdcost = parsedAdcost;
      }
      window._funnelScheduleSource._method = matchMethod;
      console.log('[퍼널 매핑] 결과 → 톡방인원:', 톡방인원, '| 라이브참여자:', 라이브참여자);
    } else {
      console.log('[퍼널 매핑] ❌ 매칭 실패! ERP:', d.강의그룹명, '| 날짜:', erpDates);
      console.log('[퍼널 매핑] 시트 강의명 샘플:', SCHEDULE_DATA.slice(0,5).map(r => r.강의명 + '(' + r.강의일 + ')'));
    }
  }

  const upc = window._uniquePayerCounts || {};
  return {
    총신청자: kd.무료강의신청수||null,
    단톡방유입: 톡방인원,
    라이브참여자: 라이브참여자,
    앵콜입장자: null,   // 수동 입력값만 사용 → "-" 표시
    앵콜결제자: null,   // 수동 입력값만 사용 → "-" 표시 (자동 계산 금지)
    환불자: refundCnt,  // ERP 주문 파일 기준 (환불/취소 건수)
    최종결제자: upc.total !== null ? upc.total : (kd.수강생수||null),  // 주문 유니크 우선, 없으면 ERP
  };
}

function _funnelBuild(kd) {
  const def = _funnelErpDefaults(kd, _fKey);
  const sv = _fSaved.values||{}, st = _fSaved.targets||{}, sf = _fSaved.feedback||{};
  const v = {};
  FUNNEL_STAGES.forEach(s=>{
    v[s.key] = sv[s.key]!==undefined ? sv[s.key] : (def[s.key]!==undefined ? def[s.key] : null);
  });
  // 라이브 결제자 기본값 (수동 입력 없을 때만 자동 계산)
  // - 앵콜 결제자 수동 입력 없음: 라이브 결제자 = 최종 결제자
  // - 앵콜 결제자 수동 입력 있음: 라이브 결제자 = 최종 결제자 - 앵콜 결제자
  if (sv.라이브결제자===undefined && v.최종결제자!==null) {
    if (sv.앵콜결제자!==undefined && v.앵콜결제자!==null) {
      v.라이브결제자 = Math.max(0, Number(v.최종결제자) - Number(v.앵콜결제자));
    } else {
      v.라이브결제자 = Number(v.최종결제자);
    }
  }
  const tgt = {};
  Object.keys(FUNNEL_DEFAULTS).forEach(k=>{ tgt[k]=st[k]!==undefined?st[k]:FUNNEL_DEFAULTS[k]; });
  const pct = (a,b) => (a!==null&&b!==null&&Number(b)>0) ? (Number(a)/Number(b)*100) : null;
  const conv = {
    유입률:       pct(v.단톡방유입,   v.총신청자),
    참여율:       pct(v.라이브참여자, v.단톡방유입),
    라이브전환율: pct(v.라이브결제자, v.라이브참여자),
    앵콜참여율:   pct(v.앵콜입장자,   v.라이브참여자),
    앵콜전환율:   pct(v.앵콜결제자,   v.앵콜입장자),
    환불율:       pct(v.환불자,        v.최종결제자),
    최종결제율: (() => {
      const a = (v.라이브결제자!==null?Number(v.라이브결제자):0) + (v.앵콜결제자!==null?Number(v.앵콜결제자):0);
      const b = (v.라이브참여자!==null?Number(v.라이브참여자):0) + (v.앵콜입장자!==null?Number(v.앵콜입장자):0);
      return b > 0 ? a/b*100 : null;
    })(),
  };
  const maxV = v.총신청자||1;
  const container = document.getElementById('funnel-rows');
  const alertEl = document.getElementById('funnel-alert');

  // ── 유사 매핑 경고 + 출처 → funnel-alert 영역에 표시 ──
  let alertHtml = '';
  const fw = window._scheduleFuzzyWarnings||{};
  const src = window._funnelScheduleSource||{};
  const fuzzyItems = [];
  const funnelToSheet = { 단톡방유입:'톡방인원', 라이브참여자:'라이브참여자' };
  Object.entries(funnelToSheet).forEach(([fk, sk]) => {
    if (fw[sk] && fw[sk].matched && !fw[sk].exactMatch) {
      fuzzyItems.push('<b>' + fk + '</b> → 시트 컬럼 "<b>' + fw[sk].colName + '</b>" (유사 매핑)');
    }
  });
  if (fuzzyItems.length > 0 && SCHEDULE_DATA.length > 0) {
    alertHtml += '<div id="funnel-fuzzy-warn" style="background:rgba(239,68,68,0.12);border:1px solid rgba(239,68,68,0.4);border-radius:6px;padding:8px 12px;margin-bottom:8px;font-size:12px;color:#fca5a5;display:flex;align-items:flex-start;gap:10px;">' +
      '<div style="flex:1;">⚠️ <b>유사 매핑 주의</b> — 정확한 컬럼명이 없어 유사한 컬럼으로 자동 매핑되었습니다: ' +
      fuzzyItems.join(', ') + '</div>' +
      '<button onclick="document.getElementById(\\\'funnel-fuzzy-warn\\\').style.display=\\\'none\\\'" style="flex-shrink:0;padding:2px 10px;border-radius:4px;border:1px solid rgba(239,68,68,0.4);background:transparent;color:#fca5a5;font-size:11px;cursor:pointer;">확인</button>' +
      '</div>';
  }
  const sourceItems = [];
  if (SCHEDULE_DATA.length > 0) {
    if (src.단톡방유입) sourceItems.push('단톡방 유입 ← 시트 "' + src.단톡방유입 + '"');
    if (src.라이브참여자) sourceItems.push('라이브 참여자 ← 시트 "' + src.라이브참여자 + '"');
    if (sourceItems.length > 0) {
      alertHtml += '<div style="font-size:11px;color:var(--muted);margin-bottom:6px;">📋 시트 매핑: ' + sourceItems.join(' | ') + '</div>';
    }
  }
  if (alertEl) alertEl.innerHTML = alertHtml;

  container.innerHTML = FUNNEL_STAGES.map(s=>{
    const val    = v[s.key];
    const convV  = s.convKey ? conv[s.convKey] : null;
    const tgtV   = s.tgtKey  ? tgt[s.tgtKey]   : null;
    const barW   = val!==null ? Math.max(2, Math.min(100, Number(val)/maxV*100)) : 0;
    const below  = s.tgtKey && convV!==null && convV<tgtV;
    const above  = s.tgtKey && convV!==null && convV>=tgtV;
    const convCls = convV===null ? 'neutral' : (above?'above':(below?'below':'neutral'));
    const convTxt = convV!==null ? convV.toFixed(1)+'%' : '-';
    const aarAuto = below ? s.aarMsg : '';
    const aarTxt  = sf[s.key]!==undefined ? sf[s.key] : '';
    const aarCls  = aarTxt ? 'ok' : 'ok';
    return '<div class="funnel-row">' +
      '<div class="funnel-stage">'+s.label+'</div>' +
      '<div class="funnel-bar-wrap"><div class="funnel-bar" style="width:'+barW+'%"></div></div>' +
      '<div class="funnel-val-cell"><input class="funnel-inp" data-fv="'+s.key+'" type="text" value="'+(val!==null?val:'')+'" placeholder="-"><span style="font-size:11px;color:var(--muted);">명</span></div>' +
      '<div class="funnel-conv-wrap">' +
        '<span class="funnel-conv '+convCls+'">'+convTxt+'</span>' +
        (s.tgtKey ? '<span style="font-size:10px;color:var(--muted);">▶</span><input class="funnel-tgt-inp" data-tv="'+s.tgtKey+'" type="text" value="'+tgtV+'"><span style="font-size:10px;color:var(--muted);">%</span>' : '<span style="font-size:11px;color:var(--muted);">-</span>') +
      '</div>' +
      '<div><input class="funnel-aar-inp '+aarCls+'" data-av="'+s.key+'" type="text" value="'+aarTxt+'" placeholder="'+(aarAuto||'피드백 입력...')+'"></div>' +
    '</div>';
  }).join('');
  container.querySelectorAll('.funnel-inp').forEach(el=>el.addEventListener('change',_funnelValChange));
  container.querySelectorAll('.funnel-tgt-inp').forEach(el=>el.addEventListener('change',_funnelTgtChange));
  container.querySelectorAll('.funnel-aar-inp').forEach(el=>el.addEventListener('change',_funnelAarChange));

  // ── 핵심 지표 영역 렌더링 ──
  _renderKeyMetrics(kd, v);
}

function _renderKeyMetrics(kd, funnelVals) {
  const el = document.getElementById('key-metrics-content');
  if (!el) return;

  // 매출 데이터 (ERP 기준)
  const totalRev = kd.강의총매출 || 0;
  // 시트 데이터
  const targetRev = window._scheduleTargetRevenue;
  const sheetROAS = window._scheduleROAS;
  // ERP ROAS (fallback)
  const erpROAS = kd.광고비 > 0 ? (kd.강의총매출 / kd.광고비) : null;
  const displayROAS = sheetROAS !== null ? sheetROAS : erpROAS;
  // 목표달성율
  const achieveRate = (targetRev && targetRev > 0) ? (totalRev / targetRev * 100) : null;

  // 전환율 계산 (퍼널 값 기반)
  const safeDiv = (a, b) => (a !== null && a !== undefined && b !== null && b !== undefined && Number(b) > 0) ? (Number(a) / Number(b) * 100) : null;
  const convTok   = safeDiv(funnelVals.단톡방유입, funnelVals.총신청자);       // 톡방 입장률
  const convLive  = safeDiv(funnelVals.라이브참여자, funnelVals.단톡방유입);    // 라이브 입장전환률
  const _lp = funnelVals.라이브결제자!==null&&funnelVals.라이브결제자!==undefined ? Number(funnelVals.라이브결제자) : 0;
  const _ep = funnelVals.앵콜결제자!==null&&funnelVals.앵콜결제자!==undefined ? Number(funnelVals.앵콜결제자) : 0;
  const _la = funnelVals.라이브참여자!==null&&funnelVals.라이브참여자!==undefined ? Number(funnelVals.라이브참여자) : 0;
  const _ea = funnelVals.앵콜입장자!==null&&funnelVals.앵콜입장자!==undefined ? Number(funnelVals.앵콜입장자) : 0;
  const convTotal = (_la+_ea) > 0 ? (_lp+_ep)/(_la+_ea)*100 : null;  // 최종 결제전환률
  const convEncore = safeDiv(funnelVals.앵콜입장자, funnelVals.라이브참여자);   // 앵콜 입장률

  const fmtR = (n) => { if (n === null) return '-'; return n >= 100000000 ? (n/100000000).toFixed(1)+'억' : n >= 10000 ? (n/10000).toFixed(0)+'만' : n.toLocaleString(); };
  const fmtPct = (n) => n !== null ? n.toFixed(1) + '%' : '-';
  const roasSrc = sheetROAS !== null ? '시트' : (erpROAS !== null ? 'ERP' : '');

  let html = '<div class="km-grid">';
  // 총 매출
  html += '<div class="km-card"><div class="km-label">총 매출 (ERP)</div><div class="km-value">' + fmtR(totalRev) + '<span class="km-unit">원</span></div></div>';
  // 목표매출
  html += '<div class="km-card"><div class="km-label">목표매출 (시트)</div><div class="km-value">' + (targetRev ? fmtR(targetRev) + '<span class="km-unit">원</span>' : '<span style="color:var(--muted);">-</span>') + '</div>';
  if (!targetRev && SCHEDULE_DATA.length === 0) html += '<div class="km-sub">일정 시트를 업로드하세요</div>';
  html += '</div>';
  // 목표달성율
  const achieveCls = achieveRate !== null ? (achieveRate >= 100 ? 'km-achieve' : 'km-achieve under') : '';
  html += '<div class="km-card"><div class="km-label">목표달성율</div><div class="km-value ' + achieveCls + '">' + fmtPct(achieveRate) + '</div>';
  if (achieveRate !== null) {
    const gap = totalRev - targetRev;
    html += '<div class="km-sub">' + (gap >= 0 ? '+' : '') + fmtR(gap) + '원</div>';
  }
  html += '</div>';
  // ROAS
  html += '<div class="km-card"><div class="km-label">ROAS' + (roasSrc ? ' (' + roasSrc + ')' : '') + '</div><div class="km-value">' + (displayROAS !== null ? displayROAS.toFixed(1) + '<span class="km-unit">배</span>' : '<span style="color:var(--muted);">-</span>') + '</div></div>';
  html += '</div>';

  // 구분선
  html += '<div class="km-divider"></div>';

  // 전환율 지표
  html += '<div style="font-size:12px;color:var(--muted);margin-bottom:8px;font-weight:600;">전환율 지표</div>';
  html += '<div class="km-conv-grid">';

  const convItems = [
    { label: '톡방 입장률', value: convTok, formula: '단톡방유입 / 총신청자', a: funnelVals.단톡방유입, b: funnelVals.총신청자 },
    { label: '라이브 입장전환률', value: convLive, formula: '라이브참여자 / 단톡방유입', a: funnelVals.라이브참여자, b: funnelVals.단톡방유입 },
    { label: '최종 결제전환률', value: convTotal, formula: '(라이브결제자+앵콜결제자)/(라이브참여자+앵콜입장자)', a: (_lp+_ep)||null, b: (_la+_ea)||null },
    { label: '앵콜 입장률', value: convEncore, formula: '앵콜입장자 / 라이브참여자', a: funnelVals.앵콜입장자, b: funnelVals.라이브참여자 },
  ];
  convItems.forEach(c => {
    const numA = c.a !== null && c.a !== undefined ? Number(c.a).toLocaleString() : '-';
    const numB = c.b !== null && c.b !== undefined ? Number(c.b).toLocaleString() : '-';
    html += '<div class="km-conv-card">' +
      '<div class="km-conv-label">' + c.label + '</div>' +
      '<div class="km-conv-value">' + fmtPct(c.value) + '</div>' +
      '<div class="km-conv-sub">' + numA + ' / ' + numB + '</div>' +
    '</div>';
  });
  html += '</div>';

  el.innerHTML = html;

  // ── 상단 새 KPI Row1/Row2 업데이트 ──
  const _setKpi = (id, val) => { const e = document.getElementById(id); if(e) e.innerHTML = val; };
  const _setSub = (id, val) => { const e = document.getElementById(id); if(e) e.textContent = val; };

  // Row1: 최종 결제전환률
  _setKpi('kpi-top-finalconv', convTotal !== null
    ? convTotal.toFixed(1)+'<span class="kpi-unit">%</span>'
    : '-');
  const _lp2 = funnelVals.라이브결제자!==null&&funnelVals.라이브결제자!==undefined ? Number(funnelVals.라이브결제자) : 0;
  const _ep2 = funnelVals.앵콜결제자!==null&&funnelVals.앵콜결제자!==undefined ? Number(funnelVals.앵콜결제자) : 0;
  const _la2 = funnelVals.라이브참여자!==null&&funnelVals.라이브참여자!==undefined ? Number(funnelVals.라이브참여자) : 0;
  const _ea2 = funnelVals.앵콜입장자!==null&&funnelVals.앵콜입장자!==undefined ? Number(funnelVals.앵콜입장자) : 0;
  _setSub('kpi-top-finalconv-sub', (_lp2||_ep2) ? (_lp2+_ep2)+'명 / '+(_la2+_ea2)+'명' : '');

  // Row1: 목표매출
  _setKpi('kpi-top-target', targetRev
    ? fmtFull(targetRev)+'<span class="kpi-unit">원</span>'
    : '-');

  // Row1: 목표달성률
  if (achieveRate !== null) {
    _setKpi('kpi-top-achieve', achieveRate.toFixed(1)+'<span class="kpi-unit">%</span>');
    const gap = totalRev - targetRev;
    _setSub('kpi-top-achieve-sub', (gap >= 0 ? '+' : '') + fmtR(gap) + '원');
  } else {
    _setKpi('kpi-top-achieve', '-');
    _setSub('kpi-top-achieve-sub', '');
  }

  // Row2: 전환률 5개
  const _fmtConv = (n) => n !== null ? n.toFixed(1)+'<span class="kpi-unit">%</span>' : '-';
  const convLivePay = (funnelVals.라이브결제자!==null && funnelVals.라이브참여자!==null && Number(funnelVals.라이브참여자)>0)
    ? Number(funnelVals.라이브결제자)/Number(funnelVals.라이브참여자)*100 : null;
  const convEncorePay = (funnelVals.앵콜결제자!==null && funnelVals.앵콜입장자!==null && Number(funnelVals.앵콜입장자)>0)
    ? Number(funnelVals.앵콜결제자)/Number(funnelVals.앵콜입장자)*100 : null;

  _setKpi('kpi-conv-tok',        _fmtConv(convTok));
  _setSub('kpi-conv-tok-sub',    (funnelVals.단톡방유입!==null?funnelVals.단톡방유입+'명':'') + (funnelVals.총신청자!==null?' / '+funnelVals.총신청자+'명':''));
  _setKpi('kpi-conv-live',       _fmtConv(convLive));
  _setSub('kpi-conv-live-sub',   (funnelVals.라이브참여자!==null?funnelVals.라이브참여자+'명':'') + (funnelVals.단톡방유입!==null?' / '+funnelVals.단톡방유입+'명':''));
  _setKpi('kpi-conv-livepay',    _fmtConv(convLivePay));
  _setSub('kpi-conv-livepay-sub',(funnelVals.라이브결제자!==null?funnelVals.라이브결제자+'명':'') + (funnelVals.라이브참여자!==null?' / '+funnelVals.라이브참여자+'명':''));
  _setKpi('kpi-conv-encore',     _fmtConv(convEncore));
  _setSub('kpi-conv-encore-sub', (funnelVals.앵콜입장자!==null?funnelVals.앵콜입장자+'명':'') + (funnelVals.라이브참여자!==null?' / '+funnelVals.라이브참여자+'명':''));
  _setKpi('kpi-conv-encorepay',  _fmtConv(convEncorePay));
  _setSub('kpi-conv-encorepay-sub',(funnelVals.앵콜결제자!==null?funnelVals.앵콜결제자+'명':'') + (funnelVals.앵콜입장자!==null?' / '+funnelVals.앵콜입장자+'명':''));
}

function _funnelValChange(e) {
  const k=e.target.dataset.fv, raw=e.target.value.trim();
  if(!_fSaved.values) _fSaved.values={};
  if(raw==='') delete _fSaved.values[k]; else _fSaved.values[k]=parseFloat(raw)||0;
  idbSaveFunnel(_fKey,_fSaved); if(_fKd) _funnelBuild(_fKd);
}
function _funnelTgtChange(e) {
  const k=e.target.dataset.tv, raw=e.target.value.trim();
  if(!_fSaved.targets) _fSaved.targets={};
  if(raw==='') delete _fSaved.targets[k]; else _fSaved.targets[k]=parseFloat(raw)||0;
  idbSaveFunnel(_fKey,_fSaved); if(_fKd) _funnelBuild(_fKd);
}
function _funnelAarChange(e) {
  const k=e.target.dataset.av, raw=e.target.value;
  if(!_fSaved.feedback) _fSaved.feedback={};
  if(raw==='') delete _fSaved.feedback[k]; else _fSaved.feedback[k]=raw;
  idbSaveFunnel(_fKey,_fSaved);
}

async function renderFunnel(combinedKey, kd) {
  _fKey=combinedKey; _fKd=kd;
  const saved = await idbLoadFunnel(combinedKey);
  _fSaved = saved || { values:{}, targets:{}, feedback:{} };
  _funnelBuild(kd);
}

// ════════════════════════════════════════════════════
//  초기화
// ════════════════════════════════════════════════════
initFilter();
renderList();

// ════════════════════════════════════════════════════
//  IndexedDB 파일 저장/복원 (대용량 엑셀 지원)
// ════════════════════════════════════════════════════
function idbOpen() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open('nova_dashboard', 3);
    req.onupgradeneeded = e => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains('files'))  db.createObjectStore('files');
      if (!db.objectStoreNames.contains('funnel')) db.createObjectStore('funnel');
      if (!db.objectStoreNames.contains('merges')) db.createObjectStore('merges');
    };
    req.onsuccess = e => resolve(e.target.result);
    req.onerror = () => reject();
  });
}
function idbSaveMergeExclusions(set) {
  idbOpen().then(db => { const tx = db.transaction('merges','readwrite'); tx.objectStore('merges').put([...set], 'exclusions'); }).catch(()=>{});
}
function idbLoadMergeExclusions() {
  return idbOpen().then(db => new Promise(resolve => {
    const r = db.transaction('merges','readonly').objectStore('merges').get('exclusions');
    r.onsuccess = e => resolve(new Set(e.target.result||[])); r.onerror = () => resolve(new Set());
  })).catch(() => new Set());
}
function idbSaveLastUpload(dateStr) {
  idbOpen().then(db => { const tx = db.transaction('merges','readwrite'); tx.objectStore('merges').put(dateStr, 'lastUpload'); }).catch(()=>{});
}
function idbLoadLastUpload() {
  return idbOpen().then(db => new Promise(resolve => {
    const r = db.transaction('merges','readonly').objectStore('merges').get('lastUpload');
    r.onsuccess = e => resolve(e.target.result||null); r.onerror = () => resolve(null);
  })).catch(() => null);
}
function updateBanner() {
  const today = new Date();
  const todayStr = today.getFullYear() + '-' + String(today.getMonth()+1).padStart(2,'0') + '-' + String(today.getDate()).padStart(2,'0');
  let latestDate = '';
  if (SALES_DATA && SALES_DATA.length > 0) {
    SALES_DATA.forEach(d => d.rows.forEach(r => { const ds = normDateStr(r.무료강의일); if (ds && ds > latestDate) latestDate = ds; }));
  }
  const fmtDate = s => s ? s.slice(0,4)+'년 '+s.slice(5,7)+'월 '+s.slice(8,10)+'일' : '-';
  const banner = document.getElementById('top-banner');
  const ldEl   = document.getElementById('banner-lecture-date');
  const stEl   = document.getElementById('banner-status');
  if (!banner) return;
  if (ldEl) ldEl.textContent = fmtDate(latestDate);
  if (!latestDate) {
    banner.className = 'top-banner banner-none';
    if (stEl) { stEl.textContent = '데이터 없음'; stEl.className = 'banner-status-none'; }
  } else if (latestDate >= todayStr) {
    banner.className = 'top-banner banner-ok';
    if (stEl) { stEl.textContent = '최신 ✓'; stEl.className = 'banner-status-ok'; }
  } else {
    banner.className = 'top-banner banner-warn';
    if (stEl) { stEl.textContent = '업데이트 필요 ⚠️'; stEl.className = 'banner-status-warn'; }
  }
}
function idbSave(n, buffer, fname) {
  idbOpen().then(db => {
    const tx = db.transaction('files', 'readwrite');
    tx.objectStore('files').put({ buffer, fname }, 'file' + n);
  }).catch(() => {});
}
function idbLoad(n) {
  return idbOpen().then(db => new Promise((resolve) => {
    const req = db.transaction('files', 'readonly').objectStore('files').get('file' + n);
    req.onsuccess = e => resolve(e.target.result || null);
    req.onerror = () => resolve(null);
  })).catch(() => null);
}

// ── 강의 상세 sticky 헤더 top 자동 업데이트 (업로드 토글 등 높이 변화 대응) ──
(function() {
  const stickyNav = document.querySelector('.sticky-top');
  if (stickyNav && window.ResizeObserver) {
    new ResizeObserver(updateDetailStickyTop).observe(stickyNav);
  }
})();

// ── 브라우저 뒤로가기 처리 ──
window.addEventListener('popstate', function(e) {
  if (!e.state) return;
  _historyLock = true;
  if (e.state.page === 'detail' && e.state.key) {
    showDetail(e.state.key);
  } else {
    if (e.state.page === 'list') {
      selectedKey = e.state.selectedKey || '';
      const gfEl = document.getElementById('global-filter');
      if (gfEl) gfEl.value = selectedKey;
    }
    showPage(e.state.page);
  }
});

// 자동 복원
(async function restoreFromIDB() {
  history.replaceState({ page: 'home' }, '');
  MERGE_OVERRIDES_EXCLUDED = await idbLoadMergeExclusions();
  const _lastUpload = await idbLoadLastUpload();
  // 기본 데이터에도 병합 적용 (업로드 전 초기 상태)
  SALES_DATA = autoMergeGroups(DEFAULT_DATA.map(d => JSON.parse(JSON.stringify(d))), MERGE_OVERRIDES_EXCLUDED);
  initFilter(); renderList();
  const results = await Promise.all([idbLoad(1), idbLoad(2), idbLoad(3)]);
  results.forEach((data, i) => {
    if (!data) return;
    const n = i + 1;
    if (n === 1) pendingSales = data.buffer;
    else if (n === 2) pendingOrder = data.buffer;
    else if (n === 3) pendingSchedule = data.buffer;
    const nameEl = document.getElementById('fname' + n);
    const card   = document.getElementById('card' + n);
    if (nameEl) { nameEl.textContent = data.fname || '(복원됨)'; nameEl.className = 'upload-card-name filled'; }
    if (card) card.classList.add('has-file');
  });
  if (pendingSales || pendingOrder || pendingSchedule) {
    document.getElementById('apply-btn').disabled = false;
    applyUpload();
  } else {
    updateBanner();
  }
})();
</script>
</body>
</html>`;

fs.writeFileSync('./강의매출대시보드.html', html, 'utf-8');
fs.writeFileSync('./index.html', html, 'utf-8');
console.log('완료! 파일 크기:', Math.round(html.length/1024), 'KB');
