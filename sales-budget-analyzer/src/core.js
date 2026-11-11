/* ===== 分析コア（ブラウザ/Node共通） ===== */
const SHEET_CUR = '貼①【集計表】';
const SHEET_PREV = '貼①【集計表】 (前年)';
const DAYTYPES = ['平日', '休前日', '土曜', '日祝日'];
const WEEKDAYS = ['月', '火', '水', '木', '金', '土', '日'];
// 時間帯（オーダー時刻ベース）。0時は24として扱う
const BANDS = [
  { key: 'lunch', name: 'ランチ', label: '11–14時', hours: [9, 10, 11, 12, 13, 14] },
  { key: 'idle', name: 'アイドル', label: '15–16時', hours: [15, 16] },
  { key: 'early', name: 'ディナー前半', label: '17–18時', hours: [17, 18] },
  { key: 'peak', name: 'ディナーピーク', label: '19–20時', hours: [19, 20] },
  { key: 'late', name: 'レイト', label: '21時以降', hours: [21, 22, 23, 24, 1, 2, 3, 4, 5, 6, 7, 8] },
];
const HOUR_TO_BAND = {};
BANDS.forEach((b, i) => b.hours.forEach(h => { HOUR_TO_BAND[h] = i; }));

function findSheet(wb, name) {
  if (wb.Sheets[name]) return wb.Sheets[name];
  const norm = s => s.replace(/[\s　]/g, '');
  const hit = wb.SheetNames.find(n => norm(n) === norm(name));
  return hit ? wb.Sheets[hit] : null;
}

function toDate(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (typeof v === 'number') { // Excelシリアル
    const d = new Date(Math.round((v - 25569) * 86400000));
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }
  if (typeof v === 'string' && v) { const d = new Date(v); if (!isNaN(d)) return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
  return null;
}
const num = v => (typeof v === 'number' && isFinite(v)) ? v : (typeof v === 'string' && v.trim() !== '' && isFinite(+v) ? +v : 0);

/** 集計表シート → 日別レコード配列 */
function extractRows(XLSX, ws) {
  if (!ws) return [];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  let hr = aoa.findIndex(r => r && r.includes('店舗コード') && r.includes('営業日'));
  if (hr < 0) throw new Error('「店舗コード」「営業日」の見出し行が見つかりません');
  const H = aoa[hr];
  const col = name => H.indexOf(name);
  const need = { code: '店舗コード', name: '店舗名', date: '営業日', wd: '曜日', dt: '区分', weather: '天気', sales: '純売上高(合計)', cust: '客数(合計)', groups: '組数(合計)', kubun: '店舗区分', budget: '売上月次予算', custBudget: '客数月次予算', disc: '税抜値引(合計)', fc: 'FC区分', area: '営業部', cost: '原価', loss: '廃棄額', buy: '仕入高',
    salesFc: '売上日別予測', custFc: '客数日別予測', staffAct: '社員労働時間(分)', paAct: 'P/A労働時間(分)', staffPlan: '社員労働時間日別予測(分)', paPlan: 'P/A労働時間日別予測(分)' };
  const C = {};
  for (const [k, n] of Object.entries(need)) C[k] = col(n);
  for (const k of ['code', 'date', 'sales', 'budget']) if (C[k] < 0) throw new Error(`列「${need[k]}」が見つかりません`);
  const hS = {}, hC = {};
  for (let h = 1; h <= 24; h++) {
    const hh = h === 24 ? 0 : h;
    hS[h] = col(`時間別売上${hh}時(オーダー時)`);
    hC[h] = col(`時間別客数${hh}時(オーダー時)`);
  }
  const out = [];
  for (let i = hr + 1; i < aoa.length; i++) {
    const r = aoa[i];
    if (!r || r[C.code] == null || r[C.code] === '') continue;
    const date = toDate(r[C.date]);
    if (!date) continue;
    const hourS = new Array(25).fill(0), hourC = new Array(25).fill(0);
    for (let h = 1; h <= 24; h++) { hourS[h] = hS[h] >= 0 ? num(r[hS[h]]) : 0; hourC[h] = hC[h] >= 0 ? num(r[hC[h]]) : 0; }
    let dt = C.dt >= 0 ? String(r[C.dt] || '') : '';
    let wd = C.wd >= 0 ? String(r[C.wd] || '') : '';
    if (!wd) wd = WEEKDAYS[(date.getDay() + 6) % 7];
    if (!DAYTYPES.includes(dt)) dt = wd === '土' ? '土曜' : wd === '日' ? '日祝日' : wd === '金' ? '休前日' : '平日';
    out.push({
      code: String(r[C.code]).trim(),
      name: String(r[C.name] || r[C.code]).replace(/^焼肉きんぐ/, ''),
      date, wd, dt,
      weather: C.weather >= 0 ? String(r[C.weather] || '') : '',
      sales: num(r[C.sales]), cust: num(r[C.cust]), groups: C.groups >= 0 ? num(r[C.groups]) : 0,
      budget: num(r[C.budget]), custBudget: C.custBudget >= 0 ? num(r[C.custBudget]) : 0,
      disc: C.disc >= 0 ? num(r[C.disc]) : 0,
      fc: C.fc >= 0 ? String(r[C.fc] || '') : '', kubun: C.kubun >= 0 ? String(r[C.kubun] || '').trim() : '',
      area: C.area >= 0 ? String(r[C.area] || '').trim() : '',
      cost: C.cost >= 0 ? num(r[C.cost]) : 0,
      loss: C.loss >= 0 ? num(r[C.loss]) : 0,
      buy: C.buy >= 0 ? num(r[C.buy]) : 0,
      // 店長の売上・客数予測（日別）と、労働時間の実績・計画（分。社員＋P/A）
      salesFc: C.salesFc >= 0 ? num(r[C.salesFc]) : 0, custFc: C.custFc >= 0 ? num(r[C.custFc]) : 0,
      staffAct: C.staffAct >= 0 ? num(r[C.staffAct]) : 0, minAct: (C.staffAct >= 0 ? num(r[C.staffAct]) : 0) + (C.paAct >= 0 ? num(r[C.paAct]) : 0),
      staffPlan: C.staffPlan >= 0 ? num(r[C.staffPlan]) : 0, minPlan: (C.staffPlan >= 0 ? num(r[C.staffPlan]) : 0) + (C.paPlan >= 0 ? num(r[C.paPlan]) : 0),
      hourS, hourC,
    });
  }
  return out;
}

/* ===== 追加データ（1時間ごと詳細・値引詳細） ===== */
function headerIndex(aoa, must, lim) {
  for (let i = 0; i < Math.min(aoa.length, lim || 10); i++) { const r = aoa[i]; if (r && must.every(m => r.includes(m))) return i; }
  return -1;
}
/* ===== 業績速報（日報）：日別・着地見込・累計。店ごとと、直営既存店・直営全店・FC全店・全体の合計 ===== */
/**
 * cur・prev：売上集計表の行（当年・前年）、stores：analyze の店（std・fc つき）、day：日報の日
 * 日別の前年は同じ曜日（364日前）、累計・着地の前年は同じ日付の範囲（ほかの分析と同じ）
 */
function perfReport(cur, prev, stores, day) {
  const t = +day, y = day.getFullYear(), m = day.getMonth(), N = new Date(y, m + 1, 0).getDate();
  const dPrev = +new Date(y, m, day.getDate() - 364);
  const sameDay = (a, b) => a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  const H = v => v / 60;
  const base = s => {
    const rs = cur.filter(r => r.code === s.code && r.date.getMonth() === m), ps = prev.filter(r => r.code === s.code);
    const d = rs.find(r => +r.date === t) || null, pd = prev.find(r => r.code === s.code && +r.date === dPrev) || null;
    const cum = rs.filter(r => +r.date <= t), pcum = ps.filter(r => r.date.getDate() <= day.getDate());
    const sd = s.std && s.std.days ? s.std.days[day.getDate() - 1] : null;
    const stdDays = s.std && s.std.days ? s.std.days.filter(x => x.date.getDate() <= day.getDate() && x.done && x.stdA) : [];
    const S = (a, k) => sum(a, r => r[k] || 0);
    const r0 = { code: s.code, name: s.name, fc: /^FC/i.test(s.fcKind || ''), kubun: (rs[0] || {}).kubun || '' };
    // 日別
    // 〜Base：予測・計画・前年・標準のある店の分だけ（合計の行で、ない店が混ざって比がおかしくならないように）
    r0.day = d ? { budget: d.budget, salesFc: d.salesFc, sales: d.sales, cust: d.cust, custFc: d.custFc, groups: d.groups, disc: d.disc, pSales: pd ? pd.sales : null, pCust: pd ? pd.cust : null,
      plan: d.minPlan ? H(d.minPlan) : 0, act: d.minAct ? H(d.minAct) : 0, std: sd && sd.stdA ? sd.stdA : 0, loss: d.loss,
      salesFcBase: d.salesFc ? d.sales : 0, custFcBase: d.custFc ? d.cust : 0, actPlanBase: d.minPlan ? H(d.minAct || 0) : 0, actStdBase: sd && sd.stdA ? H(d.minAct || 0) : 0,
      salesPBase: pd ? d.sales : 0, custPBase: pd ? d.cust : 0 } : null;
    // 累計（月初〜その日）
    r0.cum = { budget: S(cum, 'budget'), salesFc: S(cum, 'salesFc'), sales: S(cum, 'sales'), cust: S(cum, 'cust'), custFc: S(cum, 'custFc'), disc: S(cum, 'disc'), loss: S(cum, 'loss'),
      pSales: pcum.length ? S(pcum, 'sales') : null, pCust: pcum.length ? S(pcum, 'cust') : null, salesPBase: pcum.length ? S(cum, 'sales') : 0, custPBase: pcum.length ? S(cum, 'cust') : 0,
      custFcBase: sum(cum.filter(r => r.custFc), r => r.cust), custFcSum: S(cum.filter(r => r.custFc), 'custFc'),
      plan: H(S(cum, 'minPlan')), act: sum(stdDays, x => x.actH || 0), std: sum(stdDays, x => x.stdA || 0), planDone: sum(stdDays, x => x.planH || 0),
      wdAct: sum(stdDays.filter(x => x.type === 1), x => x.actH || 0), wdStd: sum(stdDays.filter(x => x.type === 1), x => x.stdA || 0),
      weAct: sum(stdDays.filter(x => x.type === 2), x => x.actH || 0), weStd: sum(stdDays.filter(x => x.type === 2), x => x.stdA || 0),
      budgetStd: s.std ? s.std.budgetRate : null };
    // 着地見込（月）：着地＝その日までの実績＋残りの日の見込み（ほかの画面と同じ着地見込み）。月の予測・予算・前年は月の合計
    const fcRest = s.fc ? sum(s.fc.days.filter(x => x.date.getDate() > day.getDate()), x => x.fc) : 0;
    const custRest = s.fc ? sum(s.fc.days.filter(x => x.date.getDate() > day.getDate()), x => x.custFc || 0) : 0;
    r0.land = { budget: S(rs, 'budget'), salesFc: S(rs, 'salesFc'), custFc: S(rs, 'custFc'), landing: r0.cum.sales + fcRest, landCust: r0.cum.cust + custRest,
      pSales: ps.length ? S(ps, 'sales') : null, pCust: ps.length ? S(ps, 'cust') : null, cumBudget: r0.cum.budget, cumSales: r0.cum.sales, cumP: r0.cum.pSales, cust: r0.cum.cust,
      landPBase: ps.length ? r0.cum.sales + fcRest : 0, custPBase: ps.length ? r0.cum.cust + custRest : 0, cumPBase: pcum.length ? r0.cum.sales : 0,
      landFcBase: S(rs, 'salesFc') ? r0.cum.sales + fcRest : 0, custFcBase: S(rs, 'custFc') ? r0.cum.cust + custRest : 0 };
    return r0;
  };
  const rows = stores.map(base).filter(r => r.day || r.cum.sales);
  const add = (list, k) => { const o = {}; list.forEach(r => { const x = r[k]; if (!x) return; Object.keys(x).forEach(f => { if (typeof x[f] === 'number') o[f] = (o[f] || 0) + x[f]; else if (x[f] == null && !(f in o)) o[f] = null; }); }); return o; };
  const group = (name, list) => list.length ? { name, total: true, n: list.length, day: add(list, 'day'), cum: { ...add(list, 'cum'), budgetStd: null }, land: add(list, 'land') } : null;
  const isDir = r => !r.fc, kizon = r => /既存/.test(r.kubun);
  const groups = [group('直営既存店', rows.filter(r => isDir(r) && kizon(r))), group('直営全店', rows.filter(isDir)), group('FC全店', rows.filter(r => r.fc)), group('エリア全体', rows)].filter(Boolean);
  // 直営既存店と直営全店が同じなら1つに
  if (groups[0] && groups[1] && groups[0].n === groups[1].n && groups[0].name === '直営既存店') groups.splice(0, 1);
  return { day, days: N, rows: [...rows.filter(isDir), ...rows.filter(r => r.fc)], groups };
}

/* ===== NPS（お客様アンケート）：業態ごとのシート。見出し行に code・NPS・推奨・批判。後半は「先月差」 ===== */
// QSCs：Q平均・S平均・C平均・sp平均の4つ。[キー, 表示, 意味, 列の見出し, 内訳の列（Q_1・Q1_1、S1_1・S2_1、C_1、sp_1・s_8）]
const NPS_QSC = [['Q', 'Q', '品質', 'Q平均', /^Q\d*_\d+$/], ['S', 'S', 'サービス', 'S平均', /^S\d_\d+$/], ['C', 'C', 'クレンリネス', 'C平均', /^C_\d+$/], ['s', 's', 'スピード', 'sp平均', /^sp?_\d+$/]];
function npsHeader(aoa) {
  for (let i = 0; i < Math.min(aoa.length, 12); i++) { const r = (aoa[i] || []).map(x => String(x == null ? '' : x).trim()); if (r.includes('code') && r.includes('NPS') && r.includes('推奨') && r.includes('批判')) return i; }
  return -1;
}
function isNpsBook(XLSX, wb) { return wb.SheetNames.some(n => npsHeader(XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null, sheetRows: 12 })) >= 0); }
function extractNps(XLSX, wb) {
  const out = { period: '', stores: {}, brands: {} };
  const pctv = v => typeof v === 'number' ? (Math.abs(v) <= 1.0001 ? v * 100 : v) : null; // 0.43 と 43 のどちらの書き方でも点に
  const numv = v => typeof v === 'number' ? v : null;
  wb.SheetNames.forEach(sn => {
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: null });
    const hr = npsHeader(aoa); if (hr < 0) return;
    const H = aoa[hr].map(x => String(x == null ? '' : x).trim()), G = (aoa[hr - 1] || []).map(x => x == null ? '' : String(x).trim());
    if (!out.period) { const pr = (aoa[0] || []).find(x => typeof x === 'string' && /\d+\/\d+/.test(x)); if (pr) out.period = pr.trim(); }
    // 「先月差」より右は先月との差
    let diffFrom = H.length; for (let r = 0; r < hr; r++) (aoa[r] || []).forEach((x, c) => { if (String(x || '').trim() === '先月差' && c < diffFrom) diffFrom = c; });
    const idx = (name, nth) => { let k = 0; for (let c = 0; c < H.length; c++) if (H[c] === name && (k++ === (nth || 0))) return c; return -1; };
    const idxIn = (name, cur) => { for (let c = cur ? 0 : diffFrom; c < (cur ? diffFrom : H.length); c++) if (H[c] === name) return c; return -1; };
    // 平均の行：見出しの上で、回答数の列に数値がある行
    const nCol = idx('回答数'); let avgRow = null; for (let r = 0; r < hr; r++) if (aoa[r] && typeof aoa[r][nCol] === 'number') avgRow = aoa[r];
    // 時間帯：最初の4つ＝土日祝、次の4つ＝平日（今月）。先月差の側も同じ順
    const bandCols = cur => { const cs = []; const names = ['ランチ', 'アイドル', 'ディナー', 'ALL']; for (let c = cur ? 0 : diffFrom; c < (cur ? diffFrom : H.length); c++) if (names.includes(H[c])) cs.push(c); return cs.slice(0, 8); };
    const bc = bandCols(true), bd = bandCols(false);
    const qsc = cur => NPS_QSC.map(([k, label, mean, head, re]) => {
      const tot = idxIn(head, cur), items = [];
      for (let c = cur ? 0 : diffFrom; c < (cur ? diffFrom : H.length); c++) if (re.test(H[c]) && G[c]) items.push({ c, label: G[c] });
      return { k, label, mean, tot, items };
    });
    const qc = qsc(true), qd = qsc(false);
    const rec = (r, isAvg) => {
      const g = c => c >= 0 ? r[c] : null;
      const bands = cs => cs.length >= 8 ? { hd: cs.slice(0, 4).map(c => pctv(g(c))), wd: cs.slice(4, 8).map(c => pctv(g(c))) } : null;
      return {
        n: numv(g(nCol)), nps: pctv(g(idx('NPS'))), crit: numv(g(idx('批判'))), neu: numv(g(idx('中立'))), pro: numv(g(idx('推奨'))),
        bands: bands(bc), bandsDiff: isAvg ? null : bands(bd),
        qsc: qc.map((q, i) => ({ k: q.k, label: q.label, mean: q.mean, v: numv(g(q.tot)), d: isAvg ? null : numv(g(qd[i].tot)),
          items: q.items.map(it => { const di = qd[i].items.find(x => x.label === it.label); return { label: it.label, v: numv(g(it.c)), d: isAvg || !di ? null : numv(g(di.c)) }; }).filter(x => x.v != null) })),
        repeat: numv(g(idx('リピート率'))),
      };
    };
    const ranked = aoa.slice(hr + 1).filter(r => r && typeof r[0] === 'number' && typeof r[idx('業態ランキング')] === 'number').length;
    if (avgRow) out.brands[sn] = rec(avgRow, true);
    const inSheet = [];
    aoa.slice(hr + 1).forEach(r => {
      if (!r || r[0] == null || !/^\d+$/.test(String(r[0]).trim())) return;
      const g = c => c >= 0 ? r[c] : null;
      const x = rec(r, false);
      Object.assign(x, { code: String(r[0]).trim(), name: String(g(idx('店舗名')) || ''), brand: sn, am: g(idx('AM')) || '', grade: g(idx('総合評価')), rank: numv(g(idx('業態ランキング'))), rankN: ranked, areaRank: numv(g(idx('エリアランキング'))) });
      if (x.nps != null) { out.stores[x.code] = x; inSheet.push(x); }
    });
    // 内訳の項目の業態平均（平均の行には内訳がないので、業態の全店を回答数で重み付けして出す）
    const B = out.brands[sn];
    if (B) B.qsc.forEach((q, i) => {
      const labels = [...new Set(inSheet.flatMap(x => x.qsc[i].items.map(it => it.label)))];
      q.items = labels.map(label => { let w = 0, t = 0; inSheet.forEach(x => { const it = x.qsc[i].items.find(y => y.label === label); if (it && it.v != null && x.n) { w += x.n; t += it.v * x.n; } }); return { label, v: w ? t / w : null }; }).filter(it => it.v != null);
    });
  });
  return out;
}
/**
 * NPSのファイルの年月。期間は「10/1（木）～10/31（土）」のように年がないので、基準の月（売上集計表の最新の月：年*12+月-1）から決める。
 * 基準の月より後になる月は前の年とみなす
 */
function npsYm(period, refK) {
  const m = String(period || '').match(/(\d{1,2})\s*\/\s*\d{1,2}/); if (!m || refK == null) return null;
  const mo = +m[1]; let y = Math.floor(refK / 12);
  if (y * 12 + mo - 1 > refK) y--; // NPSが売上集計表より先の月になることはないので、それより後なら前の年
  return `${y}-${String(mo).padStart(2, '0')}`;
}
/** NPSの履歴に残す分（対象の店だけ。1店あたり数十バイト）：NPS・回答数・推奨・批判・QSC5つ、と業態平均 */
function npsSnapshot(data, codes, ym) {
  if (!data || !ym) return null;
  const recs = Object.values(data.stores).filter(x => !codes || codes.has(x.code)).map(x => ({ c: x.code, b: x.brand, nps: x.nps, n: x.n, pro: x.pro, crit: x.crit, q: Object.fromEntries(x.qsc.filter(q => q.v != null).map(q => [q.k, q.v])), v: 2 }));
  const brands = {}; recs.forEach(r => { const a = data.brands[r.b]; if (a && !brands[r.b]) brands[r.b] = { nps: a.nps, q: Object.fromEntries(a.qsc.filter(q => q.v != null).map(q => [q.k, q.v])) }; });
  return recs.length ? { ym, recs, brands, at: Date.now() } : null;
}
/**
 * NPS・QSCsのコメント：QSCsの4つを業態平均と比べ、差の大きい順に内訳の項目まで掘り下げる。先月より下がったものも
 * np：その店、avg：業態平均（内訳の項目の平均つき）
 */
function qscLines(np, avg) {
  if (!np) return [];
  const L = [], f1 = v => v.toFixed(1), sd = v => (v >= 0 ? '+' : '▲') + Math.abs(v).toFixed(1);
  const A = k => avg && avg.qsc.find(x => x.k === k);
  L.push({ t: `・NPS：${f1(np.nps)}${avg ? `（業態平均${f1(avg.nps)}・${sd(np.nps - avg.nps)}）` : ''}${np.grade ? `、評価${np.grade}` : ''}${np.rank ? `、業態内${np.rank}位／${np.rankN}店` : ''}。推奨${Math.round(np.pro * 100)}%・批判${Math.round(np.crit * 100)}%（回答${np.n}件）。`, d: false });
  const qs = np.qsc.filter(q => q.v != null).map(q => { const a = A(q.k); return { ...q, a: a && a.v != null ? a.v : null, gap: a && a.v != null ? q.v - a.v : null, ai: a ? a.items : [] }; });
  if (!qs.length) return L;
  L.push({ t: `・QSCs：${qs.map(q => `${q.label}（${q.mean}）${f1(q.v)}${q.gap != null ? `（平均${sd(q.gap)}）` : ''}`).join('、')}。`, d: false });
  // 業態平均より低い順に、内訳の項目で平均との差が大きいもの
  const low = qs.filter(q => q.gap != null && q.gap < -0.5).sort((a, b) => a.gap - b.gap);
  low.forEach((q, i) => {
    const its = q.items.map(it => { const a = q.ai.find(x => x.label === it.label); return { ...it, gap: a ? it.v - a.v : null }; }).filter(it => it.gap != null && it.gap < -0.5).sort((a, b) => a.gap - b.gap).slice(0, 3);
    if (its.length) L.push({ t: `・QSCsの内訳・${q.label} ${q.mean}：業態平均より${f1(-q.gap)}点低い。特に低い項目は ${its.map(it => `${it.label}${f1(it.v)}（平均${sd(it.gap)}${it.d != null ? `・先月${sd(it.d)}` : ''}）`).join('、')}。`, d: i >= 2 });
  });
  const hi = qs.filter(q => q.gap != null && q.gap >= 0.5).sort((a, b) => b.gap - a.gap);
  if (hi.length) L.push({ t: `・QSCsで業態平均より高い：${hi.map(q => `${q.label}（${q.mean}）${sd(q.gap)}`).join('、')}。`, d: true });
  // 先月より下がったもの（1点以上）と、その中で下がった項目
  const down = qs.filter(q => q.d != null && q.d <= -1).sort((a, b) => a.d - b.d);
  if (down.length) L.push({ t: `・QSCsの先月差：${down.map(q => { const its = q.items.filter(it => it.d != null && it.d <= -1).sort((a, b) => a.d - b.d).slice(0, 2); return `${q.label}（${q.mean}）${sd(q.d)}${its.length ? `（${its.map(it => `${it.label}${sd(it.d)}`).join('・')}）` : ''}`; }).join('、')}。`, d: false });
  return L;
}

/** 店のNPSの推移（古い順。最大 n か月）：[{ym, nps, n, q, avg, avgQ}] */
function npsSeries(H, code, n) {
  const byYm = {};
  Object.values((H || {}).nps || {}).forEach(x => { const r = x.recs.find(y => y.c === code); if (r) byYm[x.ym] = { ym: x.ym, nps: r.nps, n: r.n, q: r.q, avg: x.brands[r.b] ? x.brands[r.b].nps : null, avgQ: x.brands[r.b] ? x.brands[r.b].q : null }; });
  return Object.values(byYm).sort((a, b) => a.ym < b.ym ? -1 : 1).slice(-(n || 6));
}

/* ===== 店舗情報（経営企業名・担当AM・ブロック）。見出しが「[10488] 経営企業名（表示）」のような形でも読む ===== */
function masterHeader(aoa) {
  for (let i = 0; i < Math.min(aoa.length, 10); i++) {
    const r = (aoa[i] || []).map(x => String(x == null ? '' : x));
    if (r.some(x => x.includes('店舗コード')) && r.some(x => x.includes('経営企業名'))) return i;
  }
  return -1;
}
function isMasterSheet(XLSX, ws) { return masterHeader(XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, sheetRows: 10 })) >= 0; }
function extractMaster(XLSX, ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: null });
  const hr = masterHeader(aoa); if (hr < 0) return [];
  const H = aoa[hr].map(x => String(x == null ? '' : x));
  const col = (inc, exc) => H.findIndex(x => x.includes(inc) && !(exc && exc.test(x)));
  const C = { code: col('店舗コード'), name: col('店舗名', /ふりがな|かな/), kind: col('直営・FC'), company: col('経営企業名'), am: col('担当AM'), block: col('ブロック'), lane: col('特急'), address: col('住所'), status: col('開店状態') };
  const v = (r, k) => C[k] >= 0 && r[C[k]] != null ? String(r[C[k]]).trim() : '';
  const out = [];
  for (let i = hr + 1; i < aoa.length; i++) {
    const r = aoa[i]; if (!r) continue;
    const code = v(r, 'code'); if (!code) continue;
    const company = v(r, 'company');
    out.push({ code, name: v(r, 'name'), kind: v(r, 'kind'), company: company === '0' ? '' : company, am: v(r, 'am'), block: v(r, 'block'), lane: v(r, 'lane'), address: v(r, 'address'), status: v(r, 'status') });
  }
  return out;
}
/** シートの種類を判定：summary（売上集計表）／hourly（1時間ごと詳細）／payment（値引・受取内訳） */
function sheetKind(XLSX, ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, sheetRows: 30 });
  if (headerIndex(aoa, ['値引後売上高', '理論原価率', '原価率理論差']) >= 0) return 'inventory';
  if (headerIndex(aoa, ['従業員コード', '確定出勤', '労働時間']) >= 0) return 'timecard';
  if (headerIndex(aoa, ['従業員コード', '労働時間(W/S)']) >= 0) return 'schedule';
  if (headerIndex(aoa, ['社員番号', '時給'], 30) >= 0) return 'wage';
  if (headerIndex(aoa, ['店舗コード', '営業日', '純売上高(合計)']) >= 0) return 'summary';
  if (headerIndex(aoa, ['店舗コード', '時刻', '売上当年']) >= 0) return 'hourly';
  if (headerIndex(aoa, ['店舗コード', '受取内訳', '金額']) >= 0) return 'payment';
  if (headerIndex(aoa, ['店舗コード', 'メニュー名', '予定原価', '販売数']) >= 0) return 'menu';
  if (headerIndex(aoa, ['店舗コード', 'メニュー名', '予定原価', '営業日別販売数']) >= 0) return 'menuDaily';
  if (headerIndex(aoa, ['店舗コード', '品目名', '実使用量(金額)', '理論使用量(金額)']) >= 0) return 'usage';
  return null;
}
function extractHourly(XLSX, ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const hr = headerIndex(aoa, ['店舗コード', '時刻', '売上当年']); const H = aoa[hr];
  const c = n => H.indexOf(n);
  const C = { code: c('店舗コード'), date: c('営業指定開始日') >= 0 ? c('営業指定開始日') : c('営業日'), hour: c('時刻'), sales: c('売上当年'), disc: c('値引額'), cust: c('客数当年'), labor: c('労働時間') };
  const out = [];
  for (let i = hr + 1; i < aoa.length; i++) {
    const r = aoa[i]; if (!r || r[C.code] == null) continue;
    const date = toDate(r[C.date]); if (!date) continue;
    let h = num(r[C.hour]); if (h === 0) h = 24;
    out.push({ code: String(r[C.code]).trim(), date, hour: h, sales: num(r[C.sales]), disc: C.disc >= 0 ? num(r[C.disc]) : 0, cust: C.cust >= 0 ? num(r[C.cust]) : 0, labor: C.labor >= 0 ? num(r[C.labor]) : 0 });
  }
  return out;
}
/* ===== 原価管理表（店舗ごとのシート：棚卸・原価率理論差・13か月推移） ===== */
function isInventorySheet(XLSX, ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, sheetRows: 10 });
  return headerIndex(aoa, ['値引後売上高', '理論原価率', '原価率理論差']) >= 0;
}
function extractInventory(XLSX, ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const hr = headerIndex(aoa, ['値引後売上高', '理論原価率', '原価率理論差']); if (hr < 0) return null;
  const H = aoa[hr].map(h => String(h || '').replace(/\s/g, '')), V = aoa[hr + 1] || [];
  const g = k => { const i = H.indexOf(k); return i >= 0 && typeof V[i] === 'number' ? V[i] : null; };
  if (g('値引後売上高') == null || !g('値引後売上高')) return null; // 空のテンプレート
  const label = k => aoa.find(r => r && r.some(v => String(v || '').trim() === k));
  const storeRow = label('店舗：'), monthRow = label('年月：');
  const valAfter = (row, k) => { if (!row) return null; const i = row.findIndex(v => String(v || '').trim() === k); return row.slice(i + 1).find(v => v != null && v !== ''); };
  const storeTxt = String(valAfter(storeRow, '店舗：') || '');
  const m = storeTxt.match(/^\s*0*(\d+)\s*(.*)$/);
  const month = toDate(valAfter(monthRow, '年月：'));
  // 大分類別の仕入・原価（「食材費」ブロック）
  const cats = [];
  const fi = aoa.findIndex(r => r && r.some(v => String(v || '').trim() === '食材費'));
  if (fi >= 0) {
    const names = aoa[fi + 1] || [];
    const rowOf = k => aoa.slice(fi + 1, fi + 8).find(r => r && r.some(v => String(v || '').trim() === k));
    const buyR = rowOf('仕入金額'), costR = rowOf('原価金額');
    names.forEach((n, j) => { if (n && typeof n === 'string' && !/^(食材費|仕入金額)$/.test(n)) { const b = buyR ? num(buyR[j]) : 0, c = costR ? num(costR[j]) : 0; if (b || c) cats.push({ cat: n.trim(), buy: b, cost: c }); } });
  }
  // 13か月の推移（「原価推移」ブロック）
  const trend = [];
  const ti = aoa.findIndex(r => r && r.some(v => String(v || '').trim() === '原価推移'));
  if (ti >= 0) {
    const months = aoa[ti + 1] || [];
    const rowOf = k => aoa.slice(ti + 1, ti + 9).find(r => r && r.some(v => String(v || '').trim() === k));
    const rT = rowOf('理論原価率'), rP = rowOf('値引前原価率'), rL = rowOf('値引前ロス率'), rG = rowOf('ロス込原価率理論差'), rI = rowOf('棚卸高');
    months.forEach((mv, j) => { const d = toDate(mv); if (d && rT && typeof rT[j] === 'number') trend.push({ date: d, theory: rT[j], pre: rP ? rP[j] : null, loss: rL ? rL[j] : null, gapExLoss: rG ? rG[j] : null, inv: rI ? rI[j] : null }); });
  }
  return {
    code: m ? m[1] : storeTxt.trim(), store: m ? m[2].replace(/^焼肉きんぐ/, '').trim() : '', date: month ? new Date(month.getFullYear(), month.getMonth() + 1, 0) : null,
    sales: g('値引後売上高'), preSales: g('値引前売上高'), open: g('前月棚卸高'), buy: g('当月仕入高'), close: g('当月棚卸高'), cost: g('原価'),
    realRate: g('実質原価率'), preRate: g('値引前原価率'), loss: g('ロス金額'), lossPre: g('値引前ロス率'), theory: g('理論原価率'),
    gap: g('原価率理論差'), gapAmt: g('原価理論差'), gapExLoss: g('ロス込原価率理論差'), lend: g('貸出食材費'), borrow: g('借入食材費'), cats, trend,
  };
}

/* ===== メニュー別（原価・喫食量・皿数） ===== */
/** 中分類 → 区分：食べ放題フード／食べ放題ドリンク／単品／コース／対象外 */
function menuGroup(cat) {
  if (/^（Ｆ）|^\(F\)/.test(cat)) return /ドリンク|アルコール/.test(cat) ? 'fDrink' : 'fFood';
  if (/^（単）|^\(単\)/.test(cat)) return 'single';
  if (cat === 'コース') return 'course';
  return 'other';
}
function extractMenu(XLSX, ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const hr = headerIndex(aoa, ['店舗コード', 'メニュー名', '予定原価', '販売数']); const H = aoa[hr];
  const c = n => H.indexOf(n);
  const C = { period: c('営業日'), code: c('店舗コード'), cat: c('メニュー中分類名称'), name: c('メニュー名'), cost: c('予定原価'), qty: c('販売数'), sales: c('メニュー別売上'), g: c('かしょく重量(g)'), cust: c('総客数') };
  const out = [];
  for (let i = hr + 1; i < aoa.length; i++) {
    const r = aoa[i]; if (!r || r[C.code] == null) continue;
    // 営業日は「2026/10/01～2026/10/31」の期間表記。終了日を日付として使う
    const ds = String(r[C.period] || '').split(/[～~〜]/).map(x => toDate(x.trim())).filter(Boolean);
    const cat = String(r[C.cat] || '');
    out.push({ code: String(r[C.code]).trim(), date: ds[ds.length - 1] || null, from: ds[0] || null, cat, grp: menuGroup(cat), name: String(r[C.name] || ''),
      unitCost: num(r[C.cost]), qty: num(r[C.qty]), sales: C.sales >= 0 ? num(r[C.sales]) : 0, g: C.g >= 0 ? num(r[C.g]) : 0, cust: C.cust >= 0 ? num(r[C.cust]) : 0 });
  }
  return out;
}
/**
 * 日別販売実績表（メニューごとの日別の販売数）→ メニュー別と同じ形に月の途中まで合計する。
 * 客数・かしょく重量の列はないので、客数はあとで売上集計表から入れる（cust: 0）
 */
function extractMenuDaily(XLSX, ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const hr = headerIndex(aoa, ['店舗コード', 'メニュー名', '予定原価', '営業日別販売数']); const H = aoa[hr];
  const c = n => H.indexOf(n);
  const C = { date: c('営業日'), code: c('店舗コード'), cat: c('メニュー中分類名称'), name: c('メニュー名'), price: c('売価'), cost: c('予定原価'), qty: c('営業日別販売数') };
  const by = {};
  for (let i = hr + 1; i < aoa.length; i++) {
    const r = aoa[i]; if (!r || r[C.code] == null) continue;
    const date = toDate(r[C.date]); if (!date) continue;
    const code = String(r[C.code]).trim(), cat = String(r[C.cat] || ''), name = String(r[C.name] || ''), k = code + '\u0000' + cat + '\u0000' + name;
    const x = (by[k] = by[k] || { code, cat, grp: menuGroup(cat), name, unitCost: num(r[C.cost]), price: num(r[C.price]), qty: 0, sales: 0, g: 0, cust: 0, date, from: date, daily: true });
    const q = num(r[C.qty]); x.qty += q; x.sales += q * num(r[C.price]);
    if (date > x.date) x.date = date; if (date < x.from) x.from = date;
  }
  const rows = Object.values(by), last = rows.length ? new Date(Math.max(...rows.map(r => +r.date))) : null, first = rows.length ? new Date(Math.min(...rows.map(r => +r.from))) : null;
  rows.forEach(r => { r.date = last; r.from = first; });
  return rows;
}
const courseType = n => String(n).replace(/[\s　]*[（(]?(大人|シニア|小学生|幼児)[）)]?[\s　]*$/, '').trim();
/** 店舗（または複数店）のメニュー行 → 1人あたり指標 */
function menuAgg(rows) {
  if (!rows.length) return null;
  const custBy = {}; rows.forEach(r => { custBy[r.code] = r.cust; });
  const cust = Object.values(custBy).reduce((a, b) => a + b, 0);
  if (!cust) return null;
  const pp = f => sum(rows, f) / cust;
  const food = rows.filter(r => r.grp === 'fFood'), drink = rows.filter(r => r.grp === 'fDrink');
  const meat = food.filter(r => /精肉/.test(r.cat));
  const cats = {};
  food.forEach(r => { const k = r.cat.replace(/^（Ｆ）/, ''); (cats[k] = cats[k] || { cat: k, qty: 0, g: 0, cost: 0 }); cats[k].qty += r.qty; cats[k].g += r.qty * r.g; cats[k].cost += r.qty * r.unitCost; });
  Object.values(cats).forEach(x => { x.qtyPP = x.qty / cust; x.gPP = x.g / cust; x.costPP = x.cost / cust; });
  const items = {};
  rows.filter(r => r.unitCost > 0).forEach(r => { (items[r.name] = items[r.name] || { name: r.name, cat: r.cat, qty: 0, cost: 0 }); items[r.name].qty += r.qty; items[r.name].cost += r.qty * r.unitCost; });
  Object.values(items).forEach(x => { x.costPP = x.cost / cust; x.qtyPP = x.qty / cust; });
  // コース構成：年齢区分（大人・シニア・小学生・幼児）を外したコースの種類ごとに合算
  const courses = {};
  rows.filter(r => r.grp === 'course' && !/飲み放題|ドリンク|アルコール/.test(r.name)).forEach(r => { const k = courseType(r.name); courses[k] = (courses[k] || 0) + r.qty; });
  const courseTotal = Object.values(courses).reduce((a, b) => a + b, 0);
  // 飲み放題比率＝注文数÷すべてのコースの注文数（大人・子どもを含む人数）。3つに分ける
  //   ソフトドリンク＝通常のソフトドリンク飲み放題、ランチソフト＝Ｌソフトドリンク飲み放題（ランチ限定）、アルコール＝生ビール飲み放題など
  //   ※Ｌソフトはランチコース以外（ランチ時間のきんぐコースなど）でも出るため、ランチコースの人数では割らない
  const dr = rows.filter(r => r.grp === 'course' && /飲み放題/.test(r.name));
  const isL = r => /^[\s　]*[ＬL]/.test(r.name);
  const soft = sum(dr.filter(r => /ソフト/.test(r.name) && !isL(r)), r => r.qty), lsoft = sum(dr.filter(r => /ソフト/.test(r.name) && isL(r)), r => r.qty), alc = sum(dr.filter(r => !/ソフト/.test(r.name)), r => r.qty);
  const dk = courseTotal && dr.length ? { soft, lsoft, alc, base: courseTotal,
    softRate: soft / courseTotal, lunchRate: lsoft / courseTotal, alcRate: alc / courseTotal, rate: (soft + lsoft + alc) / courseTotal } : null;
  return {
    cust, cost: sum(rows, r => r.qty * r.unitCost), costPP: pp(r => r.qty * r.unitCost), menuSales: sum(rows, r => r.sales),
    foodQty: sum(food, r => r.qty) / cust, foodG: sum(food, r => r.qty * r.g) / cust, foodCost: sum(food, r => r.qty * r.unitCost) / cust,
    meatQty: sum(meat, r => r.qty) / cust, meatG: sum(meat, r => r.qty * r.g) / cust, meatCost: sum(meat, r => r.qty * r.unitCost) / cust,
    drinkQty: sum(drink, r => r.qty) / cust,
    cats: Object.values(cats).sort((a, b) => b.cost - a.cost), items,
    drink: dk,
    courses: Object.entries(courses).map(([name, q]) => ({ name, q, share: courseTotal ? q / courseTotal : 0 })).sort((a, b) => b.q - a.q),
  };
}

function discCategory(name) {
  if (/従食/.test(name)) return '従業員食事';
  if (/ﾁﾗｼ|チラシ/.test(name)) return 'チラシ';
  if (/SNS/i.test(name)) return 'SNS';
  if (/優待|株主|オーナー/.test(name)) return '優待券';
  if (/アプリ|警|巡査/.test(name)) return 'アプリ会員';
  if (/個店/.test(name)) return '個店販促';
  return 'その他';
}
function extractPayments(XLSX, ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const hr = headerIndex(aoa, ['店舗コード', '受取内訳', '金額']); const H = aoa[hr];
  const c = n => H.indexOf(n);
  const C = { code: c('店舗コード'), date: c('営業日'), kind: c('受取内訳'), detail: c('内訳詳細'), count: c('件数'), amount: c('金額') };
  const out = [];
  for (let i = hr + 1; i < aoa.length; i++) {
    const r = aoa[i]; if (!r || r[C.code] == null || String(r[C.kind]) !== '値引') continue;
    const full = String(r[C.detail] || '不明');
    const name = full.split('_')[0].trim();
    out.push({ code: String(r[C.code]).trim(), date: toDate(r[C.date]), name, cat: discCategory(full), count: num(r[C.count]), amount: num(r[C.amount]) });
  }
  return out;
}

/** 追加データを店舗分析に付け足す */
function discAgg(ps) {
  if (!ps.length) return null;
  const total = sum(ps, p => p.amount), count = sum(ps, p => p.count);
  const byCat = {}, byName = {};
  ps.forEach(p => {
    (byCat[p.cat] = byCat[p.cat] || { cat: p.cat, amount: 0, count: 0 }); byCat[p.cat].amount += p.amount; byCat[p.cat].count += p.count;
    (byName[p.name] = byName[p.name] || { name: p.name, cat: p.cat, amount: 0, count: 0 }); byName[p.name].amount += p.amount; byName[p.name].count += p.count;
  });
  return { total, count, cats: Object.values(byCat).sort((a, b) => b.amount - a.amount), items: Object.values(byName).sort((a, b) => b.amount - a.amount) };
}
/** 値引の前年比較：分類ごとの件数差・金額差・1人あたり値引差（税抜換算） */
function discCompare(s) {
  const c = s.discDetail, p = s.pDiscDetail;
  if (!c || !p || !s.cust || !s.pCust) return null;
  // 税込の値引額を、売上集計表の税抜値引に合わせて換算
  const kC = c.total ? s.disc / c.total : 1, kP = p.total ? s.pDisc / p.total : 1;
  const cats = [...new Set([...c.cats.map(x => x.cat), ...p.cats.map(x => x.cat)])].map(cat => {
    const a = c.cats.find(x => x.cat === cat) || { amount: 0, count: 0 }, b = p.cats.find(x => x.cat === cat) || { amount: 0, count: 0 };
    return { cat, count: a.count, pCount: b.count, dCount: a.count - b.count, amount: a.amount, pAmount: b.amount, dAmount: a.amount - b.amount,
      // 客単価への影響（値引が増えるとマイナス）
      upEff: -(a.amount * kC / s.cust - b.amount * kP / s.pCust) };
  });
  const gp = s.groups ? s.cust / s.groups : null; // 1組あたり人数
  // 従業員食事は来店につながらないため誘客の件数から除く
  cats.forEach(x => { x.lure = x.cat !== '従業員食事'; });
  const lure = cats.filter(x => x.lure);
  const count = sum(lure, x => x.count), pCount = sum(lure, x => x.pCount);
  return { cats, count, pCount, dCount: count - pCount, gp };
}

function enrichStore(s, curRows, hourly, pays, lastDate, pPays) {
  // 値引内訳（当年・前年）
  const ps = pays.filter(p => p.code === s.code && (!p.date || p.date <= lastDate));
  if (ps.length) s.discDetail = discAgg(ps);
  const pps = (pPays || []).filter(p => p.code === s.code && (!p.date || p.date.getDate() <= lastDate.getDate()));
  if (pps.length) s.pDiscDetail = discAgg(pps);
  s.discCmp = discCompare(s);
  // 時間帯別の労働時間（人時売上）
  const hs = hourly.filter(h => h.code === s.code && h.date <= lastDate);
  if (hs.length && sum(hs, h => h.labor) > 0) {
    const dtOf = {}; curRows.filter(r => r.code === s.code).forEach(r => { dtOf[+r.date] = r.dt; });
    const g = DAYTYPES.map(() => BANDS.map(() => ({ sales: 0, labor: 0, disc: 0 })));
    hs.forEach(h => { const ti = DAYTYPES.indexOf(dtOf[+h.date]); const bi = HOUR_TO_BAND[h.hour]; if (ti < 0 || bi == null) return; g[ti][bi].sales += h.sales; g[ti][bi].labor += h.labor; g[ti][bi].disc += h.disc; });
    s.laborGrid = g;
    s.laborBand = BANDS.map((b, bi) => ({ sales: sum(g, r => r[bi].sales), labor: sum(g, r => r[bi].labor) }));
  }
}
const mhOf = c => c && c.labor > 0 ? c.sales / (c.labor / 60) : null;

const sum = (a, f) => a.reduce((s, x) => s + f(x), 0);
const isRain = w => /雨|雪/.test(w || '');

function bandTotals(days, key) { // key: 'hourS' | 'hourC'
  const t = new Array(BANDS.length).fill(0);
  for (const d of days) for (let h = 1; h <= 24; h++) { const b = HOUR_TO_BAND[h]; if (b != null) t[b] += d[key][h]; }
  return t;
}

function analyzeStore(cur, prev, monthDays) {
  const act = cur.filter(d => d.sales > 0 || d.budget > 0);
  const sales = sum(act, d => d.sales), budget = sum(act, d => d.budget);
  const cust = sum(act, d => d.cust), custBudget = sum(act, d => d.custBudget);
  const pSales = sum(prev, d => d.sales), pCust = sum(prev, d => d.cust);
  const up = cust ? sales / cust : 0;
  const bUp = custBudget ? budget / custBudget : 0;
  const pUp = pCust ? pSales / pCust : 0;
  const custEff = custBudget ? (cust - custBudget) * bUp : 0;
  const upEff = custBudget ? (up - bUp) * cust : 0;

  // 前年の区分別・時間帯構成比（予算を時間帯へ配分するため）
  const pShare = {};
  const allShare = (() => { const t = bandTotals(prev.length ? prev : act, 'hourS'); const s = t.reduce((a, b) => a + b, 0); return s ? t.map(x => x / s) : BANDS.map(() => 1 / BANDS.length); })();
  for (const dt of DAYTYPES) {
    const ds = prev.filter(d => d.dt === dt);
    const t = bandTotals(ds, 'hourS'); const s = t.reduce((a, b) => a + b, 0);
    pShare[dt] = s > 0 ? t.map(x => x / s) : allShare;
  }

  // 区分別
  const byType = DAYTYPES.map(dt => {
    const ds = act.filter(d => d.dt === dt), ps = prev.filter(d => d.dt === dt);
    const s = sum(ds, d => d.sales), b = sum(ds, d => d.budget), c = sum(ds, d => d.cust), cb = sum(ds, d => d.custBudget);
    const pss = sum(ps, d => d.sales);
    return { key: dt, days: ds.length, pDays: ps.length, sales: s, budget: b, diff: s - b, rate: b ? s / b : null,
      cust: c, custBudget: cb, avg: ds.length ? s / ds.length : 0, pAvg: ps.length ? pss / ps.length : 0, pSales: pss };
  });
  // 曜日別
  const byWd = WEEKDAYS.map(w => {
    const ds = act.filter(d => d.wd === w);
    const s = sum(ds, d => d.sales), b = sum(ds, d => d.budget);
    return { key: w, days: ds.length, sales: s, budget: b, diff: s - b, rate: b ? s / b : null };
  });

  // 時間帯別：予算を前年構成比で配分 → 区分×時間帯の予算差
  const grid = DAYTYPES.map(() => BANDS.map(() => ({ sales: 0, budget: 0, cust: 0, pSales: 0 })));
  act.forEach(d => {
    const ti = DAYTYPES.indexOf(d.dt);
    for (let h = 1; h <= 24; h++) { const b = HOUR_TO_BAND[h]; grid[ti][b].sales += d.hourS[h]; grid[ti][b].cust += d.hourC[h]; }
    // 時間帯別売上の合計と純売上の差（端数）は構成比で按分しない。予算は構成比で配分
    pShare[d.dt].forEach((sh, b) => { grid[ti][b].budget += d.budget * sh; });
  });
  prev.forEach(d => { const ti = DAYTYPES.indexOf(d.dt); for (let h = 1; h <= 24; h++) grid[ti][HOUR_TO_BAND[h]].pSales += d.hourS[h]; });
  // 前年比は曜日区分ごとの日数の違いを除くため1日あたりで比べる
  grid.forEach((row, ti) => row.forEach(c => { const d = byType[ti].days, pd = byType[ti].pDays; c.d = d; c.pd = pd;
    c.yoyDay = c.pSales > 0 && d && pd ? (c.sales / d) / (c.pSales / pd) : null; c.dayDiff = d && pd ? c.sales / d - c.pSales / pd : null; }));
  const byBand = BANDS.map((bd, b) => {
    const s = sum(grid, r => r[b].sales), bu = sum(grid, r => r[b].budget), c = sum(grid, r => r[b].cust), p = sum(grid, r => r[b].pSales);
    return { key: bd.key, name: bd.name, label: bd.label, sales: s, budget: bu, diff: s - bu, rate: bu ? s / bu : null, cust: c, pSales: p, yoy: p ? s / p : null, pDiff: s - p };
  });

  // 天候
  const rain = act.filter(d => isRain(d.weather) && d.budget > 0), fine = act.filter(d => !isRain(d.weather) && d.weather && d.budget > 0);
  const rr = ds => { const b = sum(ds, d => d.budget); return b ? sum(ds, d => d.sales) / b : null; };
  const weather = { rainDays: rain.length, rainRate: rr(rain), fineDays: fine.length, fineRate: rr(fine), rainDiff: sum(rain, d => d.sales - d.budget) };

  // 日別（ワースト・ベスト）
  const daily = act.filter(d => d.budget > 0).map(d => ({ date: d.date, wd: d.wd, dt: d.dt, weather: d.weather, sales: d.sales, budget: d.budget, diff: d.sales - d.budget, rate: d.sales / d.budget, cust: d.cust, custBudget: d.custBudget, loss: d.loss }))
    .sort((a, b) => a.date - b.date);
  const worst = [...daily].sort((a, b) => a.diff - b.diff).slice(0, 3).filter(d => d.diff < 0);
  const best = [...daily].sort((a, b) => b.diff - a.diff).slice(0, 3).filter(d => d.diff > 0);
  const closed = act.filter(d => d.sales === 0 && d.budget > 0);

  const disc = sum(act, d => d.disc), pDisc = sum(prev, d => d.disc);
  const cost = sum(act, d => d.cost), pCost = sum(prev, d => d.cost);
  // ロス（廃棄額）：突出日（1日平均の3倍以上かつ1万円以上）と曜日の偏り
  const loss = sum(act, d => d.loss), pLoss = sum(prev, d => d.loss), buy = sum(act, d => d.buy), pBuy = sum(prev, d => d.buy);
  const lossDays = act.filter(d => d.sales > 0);
  const avgLoss = lossDays.length ? loss / lossDays.length : 0;
  const lossSpikes = lossDays.filter(d => d.loss >= Math.max(avgLoss * 3, 10000)).sort((a, b) => b.loss - a.loss).slice(0, 3)
    .map(d => ({ date: d.date, wd: d.wd, loss: d.loss, rate: d.loss / d.sales }));
  const lossByWd = WEEKDAYS.map(w => ({ key: w, loss: sum(act.filter(d => d.wd === w), d => d.loss), days: act.filter(d => d.wd === w).length }));
  return {
    // 原価分析は社内の原価管理表に合わせて「値引前売上高（純売上高＋値引）」を基準にする
    preSales: sales + disc, pPreSales: pSales + pDisc,
    buy, pBuy, buyRate: buy && sales ? buy / (sales + disc) : null, pBuyRate: pBuy && pSales ? pBuy / (pSales + pDisc) : null,
    loss, pLoss, lossRate: sales && act.some(d => d.loss) ? loss / (sales + disc) : null, pLossRate: pSales && prev.some(d => d.loss) ? pLoss / (pSales + pDisc) : null, lossSpikes, lossByWd, avgLoss,
    cost, pCost, costRate: cost && sales ? cost / (sales + disc) : null, pCostRate: pCost && pSales ? pCost / (pSales + pDisc) : null,
    groups: sum(act, d => d.groups), pGroups: sum(prev, d => d.groups),
    disc, pDisc, pDiscRate: (pSales + pDisc) ? pDisc / (pSales + pDisc) : null,
    code: cur[0].code, name: cur[0].name, area: cur[0].area || '', fcKind: cur[0].fc || '',
    days: act.length, monthDays, sales, budget, diff: sales - budget, rate: budget ? sales / budget : null,
    cust, custBudget, custRate: custBudget ? cust / custBudget : null, up, bUp, pUp,
    custEff, upEff, pSales, pCust, yoy: pSales ? sales / pSales : null, custYoy: pCust ? cust / pCust : null,
    discRate: (sales + disc) ? disc / (sales + disc) : 0,
    byType, byWd, byBand, grid, weather, daily, worst, best, closed,
    hasPrev: prev.length > 0,
  };
}

/* ===== 勤怠（タイムカード・勤務スケジュール・時給） ===== */
const numc = v => typeof v === 'string' ? num(v.replace(/[,\s円]/g, '')) : num(v);
/** 時刻 → 0時からの分（"11:00"・"2026/11/1 11:00"・Date・Excelの時刻値） */
function toMin(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return v.getHours() * 60 + v.getMinutes();
  if (typeof v === 'number') return Math.round((v % 1) * 1440);
  const m = String(v).match(/(\d{1,2}):(\d{2})\s*$/);
  return m ? +m[1] * 60 + +m[2] : null;
}
/** 勤務の区間（出勤〜退勤、分）。退勤が出勤以前なら翌日 */
function shiftSpan(a, b) { const s = toMin(a), e = toMin(b); if (s == null || e == null) return null; return { s, e: e <= s ? e + 1440 : e }; }
/** 深夜（22時〜翌5時）にかかる分 */
const nightOf = sp => sp ? Math.max(0, Math.min(sp.e, 29 * 60) - Math.max(sp.s, 22 * 60)) + Math.max(0, Math.min(sp.e, 5 * 60) - sp.s) : 0;
const MONTHS_EN = { Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5, Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11 };
/** タイムカード（確定労働実績）。日付は「01(日)」、年月は「月度」（Nov-26 など）から */
function extractTimecard(XLSX, ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const hr = headerIndex(aoa, ['従業員コード', '確定出勤', '労働時間']); const H = aoa[hr];
  const c = n => H.indexOf(n);
  const C = { ym: c('月度'), code: c('店舗コード'), store: c('店舗名'), kind: c('雇用形態名称'), emp: c('従業員コード'), name: c('従業員名'), day: c('日付'),
    wsIn: c('W/S出勤'), wsOut: c('W/S退勤'), type: c('勤務種別'), in: c('確定出勤'), out: c('確定退勤'), labor: c('労働時間'), brk: c('休憩時間'), ot: c('残業時間'), night: c('深夜時間') };
  const out = [];
  for (let i = hr + 1; i < aoa.length; i++) {
    const r = aoa[i]; if (!r || r[C.emp] == null || r[C.emp] === '') continue;
    let y = null, mo = null; const ymv = r[C.ym];
    if (ymv instanceof Date) { y = ymv.getFullYear(); mo = ymv.getMonth(); }
    else { const m = String(ymv || '').match(/([A-Za-z]{3})[-\s]?(\d{2,4})|(\d{4})[\/年-](\d{1,2})/); if (m) { if (m[1]) { mo = MONTHS_EN[m[1]]; y = +m[2] < 100 ? 2000 + +m[2] : +m[2]; } else { y = +m[3]; mo = +m[4] - 1; } } }
    const dv = r[C.day]; let date = null;
    if (dv instanceof Date) date = toDate(dv); else { const d = parseInt(String(dv || ''), 10); if (d && y != null) date = new Date(y, mo, d); }
    if (!date) continue;
    const done = r[C.out] != null && r[C.out] !== '' && num(r[C.labor]) > 0;
    out.push({ code: String(r[C.code]).trim(), store: String(r[C.store] || '').replace(/^焼肉きんぐ/, ''), emp: String(r[C.emp]).trim(), name: String(r[C.name] || '').trim(),
      staff: /社員/.test(String(r[C.kind] || '')), date, type: String(r[C.type] || ''), done, labor: done ? num(r[C.labor]) : 0, ot: done ? num(r[C.ot]) : 0, night: done ? num(r[C.night]) : 0,
      ws: shiftSpan(r[C.wsIn], r[C.wsOut]) });
  }
  return out;
}
/** 勤務スケジュール表（休憩を除いた予定の労働時間） */
function extractSchedule(XLSX, ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const hr = headerIndex(aoa, ['従業員コード', '労働時間(W/S)']); const H = aoa[hr];
  const c = n => H.indexOf(n);
  const C = { area: c('営業部名'), code: c('店舗コード'), store: c('店舗名'), emp: c('従業員コード'), cls: c('従業員区分'), name: c('従業員名'), date: c('営業日'), in: c('出勤時刻(W/S)'), out: c('退勤時刻(W/S)'), mins: c('労働時間(W/S)') };
  const out = [];
  for (let i = hr + 1; i < aoa.length; i++) {
    const r = aoa[i]; if (!r || r[C.emp] == null || r[C.emp] === '') continue;
    const date = toDate(typeof r[C.date] === 'string' ? r[C.date].replace(/-/g, '/') : r[C.date]); if (!date) continue;
    const mins = num(r[C.mins]); if (!mins) continue;
    out.push({ code: String(r[C.code]).trim(), store: String(r[C.store] || '').replace(/^焼肉きんぐ/, ''), area: C.area >= 0 ? String(r[C.area] || '').trim() : '', emp: String(r[C.emp]).trim(), name: String(r[C.name] || '').trim(),
      staff: String(r[C.cls]) === '0', date, mins, span: shiftSpan(r[C.in], r[C.out]) });
  }
  return out;
}
/** 時給データ（社員番号・時給・繁忙手当） */
function extractWage(XLSX, ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const hr = headerIndex(aoa, ['社員番号', '時給'], 30); const H = aoa[hr].map(h => String(h || '').trim());
  const c = n => H.indexOf(n);
  const C = { code: c('所属コード'), emp: c('社員番号'), name: c('氏名'), wage: c('時給'), busy: c('繁忙手当'), status: c('在籍区分'), hire: c('入社年月日'), leave: c('退職年月日') };
  // 何月分の資料か（上の行の「2026年10月分」など）。なければ空
  let ym = '';
  for (let i = 0; i < hr && !ym; i++) for (const v of aoa[i] || []) { const m = String(v || '').match(/(\d{4})年\s*(\d{1,2})月分/); if (m) { ym = `${m[1]}-${String(m[2]).padStart(2, '0')}`; break; } }
  const out = [];
  for (let i = hr + 1; i < aoa.length; i++) {
    const r = aoa[i]; if (!r || r[C.emp] == null || r[C.emp] === '') continue;
    const wage = numc(r[C.wage]); if (!wage) continue;
    const hire = C.hire >= 0 ? toDate(r[C.hire]) : null, leave = C.leave >= 0 ? toDate(r[C.leave]) : null;
    out.push({ emp: String(r[C.emp]).trim(), code: C.code >= 0 ? String(r[C.code] || '').trim() : '', name: String(r[C.name] || '').trim(), wage, busy: C.busy >= 0 ? numc(r[C.busy]) : 0,
      hire: hire ? +hire : null, leave: leave ? +leave : null, status: C.status >= 0 ? String(r[C.status] || '').trim() : '', ym });
  }
  return out;
}

// 残業アラートの基準（月の残業時間。利用者確認済み：45h注意・60h要対応）
// 割増：残業25%・深夜25%、残業かつ深夜は50%（足し合わせで50%になる）。繁忙手当はなし
const OT_WARN = 45 * 60, OT_LIMIT = 60 * 60;
const PREMIUM_OT = 0.25, PREMIUM_NIGHT = 0.25;
/**
 * 当月の労働時間・雑給（パート）・残業の見込み
 * 確定日まで＝タイムカードの実績、それ以降＝勤務スケジュール（休憩を除いた予定時間）。
 * スケジュールがなければタイムカードの勤務予定から「拘束8時間以上は休憩1時間」で見込む。
 * 予定の合計が、同じ曜日区分の実績平均の半分に満たない日は「予定未作成」とみなして実績平均で補う。
 */
function analyzeLabor(tc, ws, wage, opt) {
  if (!tc.length && !ws.length) return null;
  const all = tc.concat(ws);
  const anyDate = all[0].date, y = anyDate.getFullYear(), mo = anyDate.getMonth();
  const N = new Date(y, mo + 1, 0).getDate();
  const dtOf = date => (opt.dtByDate && opt.dtByDate[+date]) || dtOfDate(date);
  // 確定日：実績のある人が、その日の予定人数の半分以上いる最後の日
  const planCnt = {}, doneCnt = {};
  (ws.length ? ws : tc.filter(r => r.ws)).forEach(r => { const d = r.date.getDate(); planCnt[d] = (planCnt[d] || 0) + 1; });
  tc.filter(r => r.done).forEach(r => { const d = r.date.getDate(); doneCnt[d] = (doneCnt[d] || 0) + 1; });
  let cut = 0;
  for (let d = 1; d <= N; d++) if ((doneCnt[d] || 0) > 0 && (doneCnt[d] || 0) >= (planCnt[d] || 0) * 0.5) cut = d;
  const wageOf = {}; wage.forEach(w => { wageOf[w.emp] = w; });
  // 予定：店舗×従業員×日
  const plan = {};
  if (ws.length) ws.forEach(r => { if (r.date.getDate() > cut) (plan[r.code] = plan[r.code] || []).push({ ...r, ot: Math.max(0, r.mins - 480), night: nightOf(r.span) }); });
  else tc.filter(r => r.ws && r.date.getDate() > cut).forEach(r => { const span = r.ws.e - r.ws.s, mins = span >= 480 ? span - 60 : span; (plan[r.code] = plan[r.code] || []).push({ ...r, mins, ot: Math.max(0, mins - 480), night: nightOf(r.ws) }); });
  const codes = [...new Set(all.map(r => r.code))].filter(c => !opt.pick || opt.pick.has(c));
  const stores = codes.map(code => {
    const act = tc.filter(r => r.code === code && r.done && r.date.getDate() <= cut);
    const pl = plan[code] || [];
    const name = (act[0] || pl[0] || all.find(r => r.code === code)).store;
    // パートの時給：時給データ。なければ店舗のパート平均
    const partW = [...new Set(act.concat(pl).filter(r => !r.staff).map(r => r.emp))].map(e => wageOf[e]).filter(Boolean).map(w => w.wage);
    const avgW = partW.length ? partW.reduce((a, b) => a + b, 0) / partW.length : null;
    const noWage = new Set();
    const cost = (r, mins) => { if (r.staff) return 0; const w = wageOf[r.emp]; if (!w) noWage.add(r.emp); const h = w ? w.wage : avgW || 0; return mins / 60 * h + (r.ot / 60) * h * PREMIUM_OT + (r.night / 60) * h * PREMIUM_NIGHT; };
    // 日別
    const days = [];
    for (let d = 1; d <= N; d++) {
      const date = new Date(y, mo, d), dt = dtOf(date);
      if (d <= cut) { const rs = act.filter(r => r.date.getDate() === d); days.push({ date, dt, done: true, mins: sum(rs, r => r.labor), staffMins: sum(rs.filter(r => r.staff), r => r.labor), cost: sum(rs, r => cost(r, r.labor)), ot: sum(rs, r => r.ot), night: sum(rs, r => r.night) }); }
      else { const rs = pl.filter(r => r.date.getDate() === d); days.push({ date, dt, done: false, mins: sum(rs, r => r.mins), staffMins: sum(rs.filter(r => r.staff), r => r.mins), cost: sum(rs, r => cost(r, r.mins)), ot: sum(rs, r => r.ot), night: sum(rs, r => r.night) }); }
    }
    // 予定未作成の日を補う（同じ曜日区分の実績平均の半分未満）
    const doneDays = days.filter(d => d.done && d.mins > 0);
    const avgBy = dt => { const ds = doneDays.filter(d => d.dt === dt); const use = ds.length ? ds : doneDays; return use.length ? { mins: sum(use, d => d.mins) / use.length, staffMins: sum(use, d => d.staffMins) / use.length, cost: sum(use, d => d.cost) / use.length } : null; };
    days.filter(d => !d.done).forEach(d => { const a = avgBy(d.dt); if (a && d.mins < a.mins * 0.5) { d.est = true; d.planMins = d.mins; d.mins = a.mins; d.staffMins = a.staffMins; d.cost = a.cost; } });
    const past = days.filter(d => d.done), fut = days.filter(d => !d.done);
    // 個人別（残業・雑給）
    const people = {};
    const add = (r, mins, isPlan) => { const p = (people[r.emp] = people[r.emp] || { emp: r.emp, name: r.name, staff: r.staff, mins: 0, planMins: 0, ot: 0, planOt: 0, night: 0, cost: 0, days: 0, planDays: 0 });
      if (isPlan) { p.planMins += mins; p.planOt += r.ot; p.planDays++; } else { p.mins += mins; p.ot += r.ot; p.days++; } p.night += r.night; p.cost += cost(r, mins); };
    act.forEach(r => add(r, r.labor, false)); pl.forEach(r => add(r, r.mins, true));
    const list = Object.values(people).map(p => ({ ...p, totalOt: p.ot + p.planOt, totalMins: p.mins + p.planMins, wage: wageOf[p.emp] ? wageOf[p.emp].wage : null }));
    const s = {
      code, name, cut, N, days, people: list,
      mins: sum(past, d => d.mins), staffMins: sum(past, d => d.staffMins), cost: sum(past, d => d.cost),
      planMins: sum(fut, d => d.mins), planStaffMins: sum(fut, d => d.staffMins), planCost: sum(fut, d => d.cost), estDays: fut.filter(d => d.est).length,
      avgWage: avgW, noWage: noWage.size, wageless: !avgW,
    };
    s.landMins = s.mins + s.planMins; s.landCost = s.cost + s.planCost;
    s.alerts = list.filter(p => p.totalOt >= OT_WARN).sort((a, b) => b.totalOt - a.totalOt);
    return s;
  }).filter(s => s.landMins > 0).sort((a, b) => a.code.localeCompare(b.code));
  const T = k => sum(stores, s => s[k]);
  const wd = stores.filter(s => !s.wageless), TW = k => sum(wd, s => s[k]);
  const area = { mins: T('mins'), staffMins: T('staffMins'), planStaffMins: T('planStaffMins'), cost: TW('cost'), planMins: T('planMins'), planCost: TW('planCost'), landMins: T('landMins'), landCost: TW('landCost'), wageless: stores.length - wd.length, wagelessNames: stores.filter(s => s.wageless).map(s => s.name),
    estDays: Math.max(0, ...stores.map(s => s.estDays)), noWage: TW('noWage'), alerts: stores.flatMap(s => s.alerts.map(p => ({ ...p, store: s.name }))).sort((a, b) => b.totalOt - a.totalOt), cut, N };
  return { stores, area, cut, N, y, mo, hasSchedule: ws.length > 0, hasTimecard: tc.length > 0, hasWage: wage.length > 0 };
}
const hrs = m => (Math.round(m / 60 * 10) / 10).toLocaleString('ja-JP') + 'h';
const hrs0 = m => Math.round(m / 60).toLocaleString('ja-JP') + 'h';
/** 人のコメント（店舗・全体共通）。f：売上の見込み（雑給率に使う） */
function laborLines(s, L, f, sales, ctx) {
  const out = [], push = (t, d) => out.push({ t, d: !!d });
  const cutTxt = L.cut ? `${L.mo + 1}/${L.cut}時点` : '実績なし';
  const land = L.cut < L.N;
  if (!L.cut) push(`【労働時間】勤務予定${hrs0(s.landMins)}（${L.mo + 1}月の勤務スケジュール。うち社員${hrs0((s.staffMins || 0) + (s.planStaffMins || 0))}）。`);
  else push(`【労働時間】実績${hrs0(s.mins)}（${cutTxt}。うち社員${hrs0(s.staffMins)}）${land ? `、着地見込み${hrs0(s.landMins)}（残り${L.N - L.cut}日の${L.fromSummary ? '計画' : '勤務予定'}${hrs0(s.planMins)}）` : ''}。`);
  if (s.estDays) push(`　※残り${s.estDays}日は${L.fromSummary ? '計画の労働時間がない' : '勤務予定が未作成（予定が同じ曜日区分の実績平均の半分未満）の'}ため、実績平均で補っています。`);
  if (L.hasWage && s.wageless === true) push('・雑給：時給データにこの店舗のパートがいないため計算できません（時給データの店舗を確認してください）。');
  else if (L.hasWage) {
    const landSales = f ? f.landing : null;
    let t = !L.cut ? `・雑給（パート）：勤務予定どおりなら${man(s.landCost)}${landSales ? `。雑給率${pct(s.landCost / landSales)}（売上見込み比）` : ''}` : `・雑給（パート）：実績${man(s.cost)}${land ? `、着地見込み${man(s.landCost)}` : ''}`;
    if (sales && L.cut) t += `。雑給率${pct(s.cost / sales)}（実績）${landSales ? `／${pct(s.landCost / landSales)}（売上・雑給とも着地見込み）` : ''}`;
    push(t + `。割増は残業25%・深夜25%（重なる時間は50%）で計算${s.wageless ? `（時給データのない${s.wagelessNames.join('・')}を除く）` : ''}。`);
    if (s.noWage) push(`　時給データにない${s.noWage}名（新規入社など）は店舗のパート平均時給${s.avgWage ? yen(s.avgWage) : ''}で計算しています。`, true);
  }
  if (L.fromSummary) { push('　※労働時間は売上集計表の実績・計画（労働時間日別予測）から。タイムカード表・勤務スケジュール表を入れると、残業・雑給・個人別も出ます。', true); return out; }
  if (s.alerts && s.alerts.length) {
    const over = s.alerts.filter(p => p.totalOt >= OT_LIMIT);
    push(`・残業アラート（月の残業見込み${OT_WARN / 60}h以上）：${s.alerts.slice(0, 5).map(p => `${p.store ? p.store + '・' : ''}${p.name}${p.staff ? '（社員）' : ''}${hrs(p.totalOt)}（実績${hrs(p.ot)}＋予定${hrs(p.planOt)}）`).join('、')}${s.alerts.length > 5 ? ` ほか${s.alerts.length - 5}名` : ''}。${over.length ? `うち${over.length}名は${OT_LIMIT / 60}h以上の見込みで、勤務予定の見直しが必要。` : ''}`);
  } else push(`・残業アラート：月の残業見込みが${OT_WARN / 60}h以上の人はいません。`);
  return out;
}

/* ===== 標準比（標準比＆人件費 見込君のマスタ） ===== */
/** 標準比のブック：シート「マスタ」「予算標準比」があるもの */
function isStdBook(wb) { return !!(findSheet(wb, 'マスタ') && findSheet(wb, '予算標準比')); }
function extractStd(XLSX, wb) {
  const aoaOf = n => { const ws = findSheet(wb, n); return ws ? XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }) : []; };
  const t = v => String(v == null ? '' : v).trim();
  // 卓数ごとの切片（ランチ無し日／ランチ有り日）
  const M = aoaOf('マスタ'), tables = [];
  M.forEach((r, i) => { const j = r ? r.findIndex(v => t(v) === '焼肉卓数') : -1; if (j < 0) return;
    for (let k = i + 1; k < M.length; k++) { const x = M[k]; if (!x || typeof x[j] !== 'number') break; tables.push({ tables: x[j], noL: num(x[j + 1]), L: num(x[j + 2]) }); } });
  // 予算標準比（店舗×月）
  const B = aoaOf('予算標準比'), budget = {};
  const bh = B.findIndex(r => r && r.some(v => t(v) === '店舗コード'));
  if (bh >= 0) { const H = B[bh], ci = H.findIndex(v => t(v) === '店舗コード');
    for (let k = bh + 1; k < B.length; k++) { const r = B[k]; if (!r || r[ci] == null || r[ci] === '') continue; const m = {};
      H.forEach((h, j) => { const mo = typeof h === 'number' ? h : parseInt(t(h), 10); if (mo >= 1 && mo <= 12 && typeof r[j] === 'number') m[mo] = r[j]; }); budget[t(r[ci])] = m; } }
  // 店舗情報（卓数）
  const S = aoaOf('店舗情報・予算標準比'), info = {};
  const sh = S.findIndex(r => r && r.some(v => t(v) === '店舗コード') && r.some(v => t(v) === '卓数'));
  if (sh >= 0) { const H = S[sh].map(t), ci = H.indexOf('店舗コード'), ti = H.indexOf('卓数'), ni = H.indexOf('店舗名'), fi = H.indexOf('フラグ'), vi = H.indexOf('変数');
    for (let k = sh + 1; k < S.length; k++) { const r = S[k]; if (!r || r[ci] == null || r[ci] === '') continue;
      const flag = fi >= 0 ? t(r[fi]) : '', per = vi >= 0 && typeof r[vi] === 'number' ? r[vi] : null;
      info[t(r[ci])] = { tables: num(r[ti]), name: ni >= 0 ? t(r[ni]).replace(/^焼肉きんぐ/, '') : '', flag, perCust: per }; } }
  // 祝日（カレンダーマスタの曜日が「祝」の日。会社の休日扱いの日を含む）
  const holidays = [];
  aoaOf('カレンダーマスタ').forEach(r => { if (!r) return; const j = r.findIndex(v => t(v) === '祝'); if (j > 0) { const d = toDate(r[j - 1]); if (d) holidays.push(+d); } });
  // 係数（原本シートの値。なければ既定値）
  const O = aoaOf('原本'), near = (label, dx) => { for (const r of O) { if (!r) continue; const j = r.findIndex(v => t(v) === label); if (j >= 0 && typeof r[j + dx] === 'number') return r[j + dx]; } return null; };
  const k = { perCust: near('変数', 1) ?? 0.33, roboNoL: near('L無し', 1) ?? 1.38, roboL: near('Lあり', 1) ?? 4.14, auto: near('自動案内係数', 1) ?? -0.007 };
  // レーン店（特急レーン）用のブック：店舗情報に「変数」（客数1人あたりの時間）があり、式は 切片＋客数×変数
  const lane = Object.values(info).some(x => x.perCust != null);
  return { kind: lane ? 'lane' : 'normal', tables, budget, info, holidays, k, date: Date.now() };
}
/** 通常店とレーン店の標準比マスタを1つにまとめる（どちらか片方だけでもよい） */
function combineStd(normal, lane) {
  if (!normal && !lane) return null;
  const base = normal || lane, info = {};
  Object.entries((normal || {}).info || {}).forEach(([c, x]) => { info[c] = { ...x, lane: false }; });
  Object.entries((lane || {}).info || {}).forEach(([c, x]) => { if (x.perCust != null && /特急|レーン/.test(x.flag || '')) info[c] = { ...(info[c] || {}), ...x, lane: true }; else if (!info[c]) info[c] = { ...x, lane: false }; });
  return {
    tables: normal ? normal.tables : [], laneTables: lane ? lane.tables : [],
    budget: { ...((lane || {}).budget || {}), ...((normal || {}).budget || {}) },
    info, holidays: [...new Set([...((normal || {}).holidays || []), ...((lane || {}).holidays || [])])], k: normal ? normal.k : base.k,
    hasNormal: !!normal, hasLane: !!lane,
  };
}
// 店舗設定の初期値（利用者の指定：自動案内あり・T8は2台）。店舗ごとに保存した設定があればそちらを使う
const STD_DEFAULT = { lunchWd: false, lunchHd: true, auto: true, t8: 2, hoursWd: 6, hoursHd: 12 };
/** 1日の標準時間（見込君「原本」シートの式と同じ）。type：1＝平日、2＝土日祝 */
function stdHours(cust, type, st, tables, M, lane) {
  const rows = lane ? M.laneTables : M.tables; if (!rows || !rows.length || !tables) return null;
  const tk = Math.min(Math.max(Math.round(tables), rows[0].tables), rows[rows.length - 1].tables);
  const row = rows.find(r => r.tables === tk) || rows[0];
  const lunch = type === 1 ? st.lunchWd : st.lunchHd;
  const adj = type === 1 ? ((st.lunchWd ? st.hoursWd - 12 : st.hoursWd - 6) * 2) : (st.hoursHd - 12) * 2;
  // レーン店：切片＋客数×変数（ロボ控除・自動案内なし）
  if (lane) return (st.perCust || 0) * cust + (lunch ? row.L : row.noL) + adj;
  let h = M.k.perCust * cust + (lunch ? row.L : row.noL) + adj - (lunch ? M.k.roboL : M.k.roboNoL) * st.t8;
  // 自動案内（原本の式どおり、平日のランチ営業日は控除しない）
  if (st.auto && !(type === 1 && st.lunchWd)) h += cust * 0.9 * M.k.auto;
  return h;
}
/** 店舗の標準比：確定日までは実績（客数＝売上集計表、時間＝タイムカード）、以降は客数見込みと勤務予定 */
/**
 * 店舗の標準比（日別）
 * 客数：過ぎた日は実績、これからの日は店長の客数予測（なければ客数見込み）
 * 労働時間：過ぎた日は実績（売上集計表、なければタイムカード）、これからの日は計画（売上集計表の労働時間日別予測、なければ勤務スケジュール）
 * あわせて「計画標準比」＝計画労働時間÷店長の客数予測からの標準時間（過ぎた日も含め、計画がどうだったか）
 */
function analyzeStd(s, lab, L, std, settings, salesLast, rows) {
  if (!std) return null;
  rows = rows || [];
  const info = std.info[s.code] || {}, set = settings && settings[s.code];
  const lane = !!info.lane;
  const st = { ...STD_DEFAULT, ...(set || {}) }, tables = (set && set.tables) || info.tables;
  if (lane) st.perCust = info.perCust;
  if (!tables) return { missing: true };
  if (lane ? !(std.laneTables && std.laneTables.length) : !(std.tables && std.tables.length)) return { missing: true, need: lane ? 'lane' : 'normal' };
  const base = lab ? lab.days[0].date : rows.length ? rows[0].date : null; if (!base) return null;
  const y = base.getFullYear(), mo0 = base.getMonth(), N = new Date(y, mo0 + 1, 0).getDate();
  const hol = new Set(std.holidays);
  const typeOf = date => hol.has(+date) || date.getDay() === 0 || date.getDay() === 6 ? 2 : 1;
  const rowOf = {};
  rows.forEach(r => { const k = r.date.getDate(), o = (rowOf[k] = rowOf[k] || { cust: 0, custFc: 0, sales: 0, minAct: 0, minPlan: 0 }); o.cust += r.cust || 0; o.custFc += r.custFc || 0; o.sales += r.sales || 0; o.minAct += r.minAct || 0; o.minPlan += r.minPlan || 0; });
  const estCust = {}; if (s.fc) s.fc.days.forEach(d => { estCust[d.date.getDate()] = d.custFc; });
  const days = [];
  for (let n = 1; n <= N; n++) {
    const date = new Date(y, mo0, n), type = typeOf(date), row = rowOf[n] || {}, ld = lab ? lab.days[n - 1] : null;
    const sh = c => c != null ? stdHours(c, type, st, tables, std, lane) : null;
    const salesDone = n <= salesLast && row.sales > 0;
    const custAct = salesDone ? row.cust : null;
    const mgrFc = row.custFc > 0, custFc = mgrFc ? row.custFc : (estCust[n] ?? null);
    const actH = row.minAct > 0 ? row.minAct / 60 : (ld && ld.done ? ld.mins / 60 : null);
    const planFromSum = row.minPlan > 0, planH = planFromSum ? row.minPlan / 60 : (ld && !ld.done && ld.mins > 0 ? ld.mins / 60 : null);
    const done = salesDone && actH != null;
    const hours = done ? actH : (planH ?? 0);
    const cust = done ? custAct : custFc;
    const stdFc = sh(custFc), stdA = sh(custAct);
    days.push({ date, type, cust, custEst: !done, custFc, custAct, mgrFc, hours, planH, actH, planSrc: planFromSum ? 'plan' : planH != null ? 'sched' : null,
      done, est: !done && !planFromSum && !!(ld && ld.est), std: sh(cust), stdFc, stdA,
      rPlan: planH != null && stdFc ? planH / stdFc : null, rAct: done && stdA ? actH / stdA : null, noPlan: !done && !(hours > 0) });
  }
  // 着地：実績の日＋計画（予定）のある日。計画のない日は除く
  const ok = days.filter(d => d.std != null && (d.done || d.hours > 0));
  const past = ok.filter(d => d.done), all = ok;
  const H = ds => sum(ds, d => d.hours), SH = ds => sum(ds, d => d.std), C = ds => sum(ds, d => d.cust);
  const mo = mo0 + 1, br = (std.budget[s.code] || {})[mo] ?? null;
  const landStd = SH(all), landH = H(all);
  const r = {
    settings: st, custom: !!set, tables, lane, budgetRate: br, days,
    cutStd: SH(past), cutHours: H(past), cutCust: C(past), cutRate: SH(past) ? H(past) / SH(past) : null,
    landStd, landHours: landH, landCust: C(all), landRate: landStd ? landH / landStd : null,
    remDays: days.filter(d => !d.done && d.hours > 0).length, noPlanDays: days.filter(d => d.noPlan).length, complete: ok.length === days.length,
  };
  // 計画標準比（店長の客数予測と計画労働時間で、月を通して）
  const pd = days.filter(d => d.planSrc === 'plan' && d.stdFc);
  r.planDays = pd.length; r.planHours = sum(pd, d => d.planH); r.planStd = sum(pd, d => d.stdFc); r.planRate = r.planStd ? r.planHours / r.planStd : null;
  r.mgrDays = days.filter(d => d.mgrFc).length;
  // 客数予測の当たり具合（過ぎた日）
  const acc = days.filter(d => d.done && d.mgrFc && d.custAct != null);
  r.fcAcc = acc.length ? { n: acc.length, ratio: sum(acc, d => d.custAct) / sum(acc, d => d.custFc) } : null;
  // 予算標準比を超える計画の日（これからの日）
  if (br != null) r.overDays = days.filter(d => !d.done && d.rPlan != null && d.rPlan > br).map(d => ({ date: d.date, rPlan: d.rPlan, overH: d.planH - d.stdFc * br })).sort((a, b) => b.overH - a.overH);
  // 予算標準比との差：目標時間＝標準時間×予算標準比
  if (br != null) { r.target = landStd * br; r.over = landH - r.target; r.overCost = r.over > 0 && lab && lab.avgWage ? r.over * lab.avgWage : null; }
  // 週別（月曜始まり）
  const wk = {}; ok.forEach(d => { const w = Math.floor((d.date.getDate() - 1 + ((new Date(d.date.getFullYear(), d.date.getMonth(), 1).getDay() + 6) % 7)) / 7); (wk[w] = wk[w] || []).push(d); });
  r.weeks = Object.keys(wk).sort((a, b) => a - b).map(w => { const ds = wk[w]; return { from: ds[0].date, to: ds[ds.length - 1].date, n: ds.length, cust: C(ds), std: SH(ds), hours: H(ds), rate: SH(ds) ? H(ds) / SH(ds) : null, done: ds.every(d => d.done), mixed: ds.some(d => d.done) && ds.some(d => !d.done) }; });
  return r;
}
/** タイムカード・勤務スケジュールがないとき、売上集計表の労働時間（実績・日別予測）から労働時間の形を作る */
function laborFromSummary(rows, lastDate) {
  if (!rows.some(r => r.minAct > 0 || r.minPlan > 0)) return null;
  const y = rows[0].date.getFullYear(), mo = rows[0].date.getMonth(), N = new Date(y, mo + 1, 0).getDate();
  const cut = lastDate.getMonth() === mo && lastDate.getFullYear() === y ? Math.max(0, ...rows.filter(r => r.minAct > 0 && r.date <= lastDate).map(r => r.date.getDate())) : 0;
  const codes = [...new Set(rows.map(r => r.code))];
  const stores = codes.map(code => {
    const rs = rows.filter(r => r.code === code), name = rs[0].name;
    const days = [];
    for (let d = 1; d <= N; d++) {
      const date = new Date(y, mo, d), xs = rs.filter(r => r.date.getDate() === d), dt = xs[0] ? xs[0].dt : dtOfDate(date), done = d <= cut;
      days.push({ date, dt, done, mins: sum(xs, r => done ? r.minAct : r.minPlan), staffMins: sum(xs, r => done ? r.staffAct : r.staffPlan), cost: 0, ot: 0, night: 0 });
    }
    // 計画のない日（これから）は、同じ曜日区分の実績平均で補う
    const dd = days.filter(d => d.done && d.mins > 0);
    days.filter(d => !d.done && !d.mins).forEach(d => { const same = dd.filter(x => x.dt === d.dt), use = same.length ? same : dd; if (use.length) { d.est = true; d.mins = sum(use, x => x.mins) / use.length; d.staffMins = sum(use, x => x.staffMins) / use.length; } });
    const past = days.filter(d => d.done), fut = days.filter(d => !d.done);
    const o = { code, name, cut, N, days, people: [], mins: sum(past, d => d.mins), staffMins: sum(past, d => d.staffMins), cost: 0, planMins: sum(fut, d => d.mins), planStaffMins: sum(fut, d => d.staffMins), planCost: 0,
      estDays: fut.filter(d => d.est).length, avgWage: null, noWage: 0, wageless: false, alerts: [] };
    o.landMins = o.mins + o.planMins; o.landCost = 0;
    return o;
  }).filter(s => s.landMins > 0);
  const T = k => sum(stores, s => s[k] || 0);
  const area = { mins: T('mins'), staffMins: T('staffMins'), planStaffMins: T('planStaffMins'), cost: 0, planMins: T('planMins'), planCost: 0, landMins: T('landMins'), landCost: 0, estDays: Math.max(0, ...stores.map(s => s.estDays)), noWage: 0, wageless: 0, wagelessNames: [], alerts: [], cut, N };
  return { stores, area, cut, N, y, mo, hasSchedule: false, hasTimecard: false, hasWage: false, fromSummary: true };
}
const p0 = v => v == null ? '－' : (v * 100).toFixed(1) + '%';
function stdLines(x, ctx) {
  if (!x) return [];
  if (x.missing) return [{ t: x.need === 'lane' ? '・標準比：特急レーン店のため、レーン店用の標準比マスタを読み込むと計算します。' : x.need === 'normal' ? '・標準比：通常店用の標準比マスタを読み込むと計算します。' : '・標準比：卓数がわからないため計算できません（標準比マスタに店舗がないか、店舗設定で卓数を入れてください）。', d: false }];
  const out = [];
  let t;
  if (!x.cutHours) t = `・標準比${x.lane ? '（特急レーン店）' : ''}：勤務予定どおりなら${p0(x.landRate)}（客数見込み${Math.round(x.landCust).toLocaleString('ja-JP')}人×勤務予定${hrs0(x.landHours * 60)}、人時来客${x.landHours ? (x.landCust / x.landHours).toFixed(2) : '－'}）`;
  else { t = `・標準比${x.lane ? '（特急レーン店）' : ''}：実績${p0(x.cutRate)}（${ctx.cutTxt}、人時来客${x.cutHours ? (x.cutCust / x.cutHours).toFixed(2) : '－'}）`;
    if (x.remDays) t += `、着地見込み${p0(x.landRate)}（残り${x.remDays}日は${x.mgrDays ? '店長の客数予測×計画労働時間' : '客数見込み×勤務予定'}）`; }
  if (x.budgetRate != null) t += `。予算標準比${p0(x.budgetRate)}`;
  out.push({ t: t + '。', d: false });
  if (x.budgetRate != null && x.remDays) {
    if (x.over > 0) out.push({ t: `　予算標準比に収めるには、${x.cutHours ? '残り' : ''}${x.remDays}日で${hrs0(x.over * 60)}（1日${(x.over / x.remDays).toFixed(1)}h）の削減が必要${x.overCost ? `（雑給換算で約${man(x.overCost)}）` : ''}。`, d: false });
    else out.push({ t: `　勤務予定どおりなら予算標準比内（余裕${hrs0(-x.over * 60)}）。`, d: false });
  }
  // 店長の計画（客数予測と計画労働時間）
  if (x.planDays) {
    let t2 = `・計画標準比（店長の客数予測×計画労働時間、${x.planDays}日分）：${p0(x.planRate)}`;
    if (x.budgetRate != null && x.overDays) t2 += x.overDays.length ? `。これからの日で予算標準比${p0(x.budgetRate)}を超える計画が${x.overDays.length}日（${x.overDays.slice(0, 4).map(d => `${md(d.date)}(${WEEKDAYS[(d.date.getDay() + 6) % 7]})${p0(d.rPlan)}・+${d.overH.toFixed(1)}h`).join('、')}${x.overDays.length > 4 ? ' ほか' : ''}）` : '。これからの日はすべて予算標準比内の計画';
    out.push({ t: t2 + '。', d: false });
  }
  if (x.fcAcc) out.push({ t: `　客数予測の当たり具合：過ぎた${x.fcAcc.n}日の実際の客数は予測の${p0(x.fcAcc.ratio)}${x.fcAcc.ratio >= 1.05 ? '（予測より多い。人が足りなくなりやすい）' : x.fcAcc.ratio <= 0.95 ? '（予測より少ない。計画の労働時間が多めになりやすい）' : ''}。`, d: false });
  if (x.noPlanDays) out.push({ t: `　※計画の労働時間がない${x.noPlanDays}日は着地見込みに含めていません。`, d: true });
  const bad = x.weeks.filter(w => w.n >= 3 && w.rate != null && w.rate >= 1.05).sort((a, b) => b.rate - a.rate)[0];
  if (bad) out.push({ t: `　週別で最も高いのは${md(bad.from)}〜${md(bad.to)}の${p0(bad.rate)}${bad.done ? '' : '（見込みを含む）'}。`, d: true });
  if (!x.custom) out.push({ t: `　※店舗設定（ランチ営業・${x.lane ? '' : 'T8台数・'}営業時間${x.lane ? '' : '・自動案内'}）が初期値のままです。「人」タブの店舗詳細で確認してください。`, d: true });
  return out;
}

/* ===== 採用・定着（タイムカード・時給データを月ごとに積み上げて分析） ===== */
const ymOf = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
const ymIdx = ym => { const [y, m] = ym.split('-').map(Number); return y * 12 + m - 1; };
const ymTxt = ym => `${+ym.slice(0, 4)}年${+ym.slice(5)}月`;
/** タイムカード → 月の個人別まとめ（履歴として保存する小さな形） */
function tcSummary(tc) {
  if (!tc.length) return null;
  const ym = ymOf(tc[0].date), by = {};
  tc.filter(r => r.done).forEach(r => { const p = (by[r.emp] = by[r.emp] || { e: r.emp, n: r.name, c: r.code, s: r.staff ? 1 : 0, m: 0, d: 0, o: 0 }); p.m += r.labor; p.d++; p.o += r.ot; });
  // 確定日（その月の実績の最終日）
  const last = Math.max(0, ...tc.filter(r => r.done).map(r => r.date.getDate()));
  return { ym, last, days: new Date(tc[0].date.getFullYear(), tc[0].date.getMonth() + 1, 0).getDate(), recs: Object.values(by), at: Date.now() };
}
/** 時給データ → 月のスナップショット（対象店舗の人だけ） */
function wageSnapshot(wage, codes, fallbackYm) {
  if (!wage.length) return null;
  const ym = wage[0].ym || fallbackYm;
  const recs = wage.filter(w => !codes || codes.has(w.code)).map(w => ({ e: w.emp, c: w.code, n: w.name, h: w.hire, l: w.leave, st: w.status, w: w.wage }));
  return { ym, recs, at: Date.now() };
}
/** 履歴をまとめる（同じ月は新しいもので置き換え） */
function mergeHistory(H, add) {
  // キーは「年月」または「年月|エリア」（ウェブアプリではエリアごとに保存する）
  const out = { tc: { ...((H || {}).tc || {}) }, wage: { ...((H || {}).wage || {}) }, nps: { ...((H || {}).nps || {}) } };
  const key = x => x.ym + (x.area ? '|' + x.area : '');
  ((add || {}).tc || []).forEach(x => { if (x && x.ym) out.tc[key(x)] = x; });
  ((add || {}).wage || []).forEach(x => { if (x && x.ym) out.wage[key(x)] = x; });
  ((add || {}).nps || []).forEach(x => { if (x && x.ym) out.nps[key(x)] = x; });
  return out;
}
/** 年月ごとにまとめ直す（エリアごとの分をつなげる） */
function byMonth(obj) {
  const out = {};
  Object.values(obj || {}).forEach(x => { const o = (out[x.ym] = out[x.ym] || { ym: x.ym, recs: [], seen: new Set(), last: x.last, days: x.days }); if (x.last != null) o.last = Math.min(o.last ?? x.last, x.last); x.recs.forEach(r => { const k = r.e + '@' + r.c; if (!o.seen.has(k)) { o.seen.add(k); o.recs.push(r); } }); });
  return out;
}
const TENURE = [
  { key: 't0', label: '1か月未満', test: k => k === 0 },
  { key: 't1', label: '1か月〜半年', test: k => k >= 1 && k <= 5 },
  { key: 't2', label: '半年〜1年', test: k => k >= 6 && k <= 11 },
  { key: 't3', label: '1年以上', test: k => k >= 12 },
];
const RAMP = [[0, 0, '入社月'], [1, 1, '2か月目'], [2, 2, '3か月目'], [3, 3, '4か月目'], [4, 5, '5〜6か月目'], [6, 8, '7〜9か月目'], [9, 11, '10〜12か月目'], [12, 23, '2年目'], [24, 999, '3年目以降']];
const RISK_WARN = 0.7, RISK_HIGH = 0.5;
/**
 * 勤務期間別の労働時間・戦力化・退社リスク（パートのみ。社員は入社日がないため対象外）
 * H：履歴 {tc:{ym:{recs}}, wage:{ym:{recs}}}、cur：当月の勤怠分析（見込みに使う）
 */
function analyzeTenure(H0, cur, opt) {
  opt = opt || {};
  const H = { tc: byMonth(H0.tc), wage: byMonth(H0.wage) };
  const pick = opt.pick, months = Object.keys(H.tc || {}).sort();
  if (!months.length) return null;
  const curYm = cur ? `${cur.y}-${String(cur.mo + 1).padStart(2, '0')}` : months[months.length - 1];
  const useMonths = months.filter(m => m <= curYm);
  // 入社日・在籍：いちばん新しいスナップショットを優先
  const wMonths = Object.keys(H.wage || {}).sort();
  const info = {};
  wMonths.forEach(m => H.wage[m].recs.forEach(r => { info[r.e] = { ...(info[r.e] || {}), ...r, seen: m, first: (info[r.e] || {}).first || m }; }));
  const lastW = wMonths[wMonths.length - 1];
  // 個人ごとの月別時間（パート）
  const P = {};
  // 月の途中までしかない月（月末後に最終のタイムカードを入れていない月）は、確定日までの実績を月換算する
  const partial = {};
  useMonths.forEach(m => { const x = H.tc[m]; if (x.last && x.days && x.last < x.days) partial[m] = x.days / x.last; });
  useMonths.forEach(m => H.tc[m].recs.forEach(r => {
    if (r.s || (pick && !pick.has(r.c))) return;
    const p = (P[r.e] = P[r.e] || { emp: r.e, name: r.n, code: r.c, hrs: {}, days: {} });
    p.hrs[m] = (p.hrs[m] || 0) + r.m / 60 * (partial[m] || 1); p.days[m] = (p.days[m] || 0) + r.d * (partial[m] || 1); p.code = r.c; p.name = r.n;
  }));
  // 当月は見込み（実績の1日あたり＋予定のある日の1日あたりで月換算）
  const curRate = {};
  if (cur && curYm === months[months.length - 1]) cur.stores.forEach(st => {
    if (pick && !pick.has(st.code)) return;
    const covered = st.days.filter(d => !d.done && !d.est).length;
    st.people.filter(x => !x.staff).forEach(x => {
      const p = (P[x.emp] = P[x.emp] || { emp: x.emp, name: x.name, code: st.code, hrs: {}, days: {} });
      const span = cur.cut + covered;
      const monthEq = span ? (x.mins + x.planMins) / 60 / span * cur.N : 0;
      p.hrs[curYm] = monthEq; p.days[curYm] = span ? (x.days + x.planDays) / span * cur.N : 0; p.code = st.code;
      curRate[x.emp] = { act: cur.cut ? x.mins / 60 / cur.cut : 0, plan: covered ? x.planMins / 60 / covered : null, actDays: x.days, planDays: x.planDays, covered };
    });
  });
  // 入社日：時給資料の入社年月日。載っていない人（利用者確認：基本的に前月入社）は推定する
  //   タイムカードの履歴で前の月にいなかったことがわかれば、初めて出てきた月に入社。わからなければ「前月入社」（履歴にもっと前からいればその月）
  const hireOf = {};
  const curI = ymIdx(curYm);
  // 店舗のパートの3割以上が時給資料にいないときは、時給資料が違う（その店舗が載っていない）とみなして推定しない
  const miss = {}, tot = {};
  Object.values(P).forEach(p => { tot[p.code] = (tot[p.code] || 0) + 1; if (!(info[p.emp] && info[p.emp].h)) miss[p.code] = (miss[p.code] || 0) + 1; });
  const noWageStores = Object.keys(tot).filter(c => (miss[c] || 0) / tot[c] > 0.3);
  Object.values(P).forEach(p => {
    const i = info[p.emp];
    if (i && i.h) { hireOf[p.emp] = { h: i.h, est: false }; return; }
    if (noWageStores.includes(p.code)) return;
    const ms = Object.keys(p.hrs).filter(m => p.hrs[m] > 0).sort();
    if (!ms.length) return;
    const first = ymIdx(ms[0]), knownBefore = useMonths.some(m => ymIdx(m) === first - 1);
    const hi = knownBefore ? first : Math.min(first, curI - 1);
    hireOf[p.emp] = { h: +new Date(Math.floor(hi / 12), hi % 12, 1), est: true };
  });
  const tenureK = (emp, ym) => { const i = hireOf[emp]; if (!i) return null; const h = new Date(i.h); return ymIdx(ym) - (h.getFullYear() * 12 + h.getMonth()); };
  const band = k => k == null ? null : TENURE.findIndex(b => b.test(k));
  // 当月の勤務期間別
  const people = Object.values(P).map(p => ({ ...p, k: tenureK(p.emp, curYm), hire: hireOf[p.emp] ? hireOf[p.emp].h : null, hireEst: !!(hireOf[p.emp] && hireOf[p.emp].est), wage: info[p.emp] ? info[p.emp].w : null, cur: p.hrs[curYm] || 0 }));
  const working = people.filter(p => p.cur > 0);
  const bandStats = list => TENURE.map((b, i) => { const xs = list.filter(p => band(p.k) === i); const h = sum(xs, p => p.cur); return { key: b.key, label: b.label, n: xs.length, hours: h, avg: xs.length ? h / xs.length : null }; })
    .concat([(() => { const xs = list.filter(p => p.k == null); const h = sum(xs, p => p.cur); return { key: 'tx', label: '入社日不明', n: xs.length, hours: h, avg: xs.length ? h / xs.length : null }; })()]);
  const bands = bandStats(working);
  const total = sum(working, p => p.cur);
  bands.forEach(b => { b.share = total ? b.hours / total : null; });
  // 入社後の月数ごとの平均時間（全期間の積み上げ。当月は見込み）
  const ramp = RAMP.map(([a, b, label]) => ({ label, n: 0, h: 0 }));
  Object.values(P).forEach(p => Object.entries(p.hrs).forEach(([m, h]) => { const k = tenureK(p.emp, m); if (k == null || k < 0 || h <= 0) return; const i = RAMP.findIndex(([a, b]) => k >= a && k <= b); ramp[i].n++; ramp[i].h += h; }));
  ramp.forEach(r => { r.avg = r.n ? r.h / r.n : null; });
  // 新規採用（当月・前月入社）と退社
  const hires = people.filter(p => p.k != null && p.k <= 1).sort((a, b) => b.hire - a.hire);
  const leavers = [];
  const prevW = wMonths.filter(m => m < lastW).pop();
  if (prevW) H.wage[prevW].recs.forEach(r => { const inLast = H.wage[lastW].recs.some(x => x.e === r.e); if (!inLast && (!pick || pick.has(r.c))) leavers.push({ emp: r.e, name: r.n, code: r.c, hire: r.h, how: `時給データ（${ymTxt(lastW)}分）から外れた`, sure: true }); });
  Object.values(info).filter(i => (i.st === '退職' || i.l) && (!pick || pick.has(i.c))).forEach(i => { if (!leavers.some(x => x.emp === i.e)) leavers.push({ emp: i.e, name: i.n, code: i.c, hire: i.h, leave: i.l, how: '退職日あり', sure: true }); });
  // 勤務が2か月続けてない人（退社の可能性）
  if (useMonths.length >= 3) { const [m2, m1, m0] = useMonths.slice(-3); Object.values(P).forEach(p => { if ((p.hrs[m2] || 0) > 0 && !(p.hrs[m1] > 0) && !(p.hrs[m0] > 0) && !leavers.some(x => x.emp === p.emp)) leavers.push({ emp: p.emp, name: p.name, code: p.code, hire: info[p.emp] && info[p.emp].h, how: `${ymTxt(m1)}・${ymTxt(m0)}の勤務なし`, sure: false }); }); }
  // 当月に勤務している人は退社扱いにしない（異動・データの都合で時給データから外れた場合など）
  for (let i = leavers.length - 1; i >= 0; i--) if (P[leavers[i].emp] && (P[leavers[i].emp].hrs[curYm] || 0) > 0) leavers.splice(i, 1);
  leavers.forEach(l => { const p = P[l.emp]; l.k = l.hire ? tenureK(l.emp, curYm) : null; l.hrs = p ? p.hrs : {}; });
  // 退社した人の、辞める前の時間の変化（履歴があるときの裏付け）
  const pre = leavers.map(l => { const ms = Object.keys(l.hrs).filter(m => l.hrs[m] > 0).sort(); if (ms.length < 3) return null; const lastH = l.hrs[ms[ms.length - 1]], before = ms.slice(0, -1).map(m => l.hrs[m]); return lastH / avgOf(before); }).filter(v => v != null && isFinite(v));
  // 退社リスク：入社1年以内のパートで、当月の時間が前の月（最大3か月）の平均の70%以下／当月の予定が実績ペースの70%以下
  const risk = [];
  working.concat(people.filter(p => p.cur === 0 && Object.keys(p.hrs).length > 1)).forEach(p => {
    if (p.k == null || p.k > 11) return;
    const prevMs = useMonths.filter(m => m < curYm && (p.hrs[m] || 0) > 0 && tenureK(p.emp, m) > 0).slice(-3);
    const prevAvg = prevMs.length ? avgOf(prevMs.map(m => p.hrs[m])) : null;
    const r = curRate[p.emp];
    // ①前の月との比較 ②当月の中で、これからの予定がこれまでの実績より減っているか。低い方を使う
    const c = [];
    // 表示は月間の総労働時間で（当月見込み＝表の「見込」と同じ値）
    // 表示は月間の総労働時間と出勤日数で（今月の見込み＝表の「見込」と同じ値）
    const sg0 = (v, u) => `${v >= 0 ? '+' : '▲'}${Math.abs(Math.round(v))}${u}`;
    const curD = p.days[curYm] || 0;
    if (prevAvg) { const prevD = avgOf(prevMs.map(m => p.days[m] || 0));
      c.push({ ratio: p.cur / prevAvg, base: prevAvg, baseD: prevD, vs: `前${prevMs.length}か月平均` }); }
    if (r && r.plan != null && r.act > 0 && r.actDays >= 3 && r.covered >= 5) { const pace = r.act * cur.N, paceD = cur.cut ? r.actDays / cur.cut * cur.N : 0;
      c.push({ ratio: r.plan / r.act, base: pace, baseD: paceD, vs: 'これまでのペース' }); }
    const best = c.sort((a, b) => a.ratio - b.ratio)[0];
    const ratio = best ? best.ratio : null;
    let why = best ? `今月見込み${Math.round(p.cur)}h・${Math.round(curD)}日／${best.vs}${Math.round(best.base)}h・${Math.round(best.baseD)}日（${sg0(p.cur - best.base, 'h')}）` : '';
    // 減り方の分解：総労働時間＝出勤日数×1回あたりの時間
    //   日数の影響＝(今月の日数−比較の日数)×比較の1回あたり、1回あたりの影響＝(今月の1回あたり−比較の1回あたり)×今月の日数（合計は時間の差と一致）
    let kind = null;
    if (best && best.baseD > 0) {
      const bh = best.base / best.baseD, ch = curD > 0 ? p.cur / curD : 0;
      const effD = (curD - best.baseD) * bh, effL = (ch - bh) * curD, tot = effD + effL;
      if (tot < 0) {
        const shD = Math.min(0, effD) / Math.min(-0.01, Math.min(0, effD) + Math.min(0, effL));
        kind = curD === 0 ? 'none' : shD >= 0.7 ? 'days' : shD <= 0.3 ? 'len' : 'both';
        if (kind !== 'none') why += `。日数で${sg0(effD, 'h')}・1回あたり${bh.toFixed(1)}→${ch.toFixed(1)}hで${sg0(effL, 'h')}`;
      }
    }
    if (ratio != null && ratio <= RISK_WARN) risk.push({ ...p, ratio, why, kind, level: ratio <= RISK_HIGH ? 'high' : 'warn' });
  });
  risk.sort((a, b) => a.ratio - b.ratio);
  // 店舗別
  const byStore = {};
  working.forEach(p => { (byStore[p.code] = byStore[p.code] || []).push(p); });
  const stores = Object.entries(byStore).map(([code, xs]) => {
    const b = bandStats(xs), t = sum(xs, p => p.cur); b.forEach(x => { x.share = t ? x.hours / t : null; });
    return { code, n: xs.length, cohorts: cohortsOf(Object.values(P).filter(p => p.code === code)), estN: xs.filter(p => p.hireEst).length, bands: b, hires: hires.filter(h => h.code === code), leavers: leavers.filter(l => l.code === code), risk: risk.filter(r => r.code === code) };
  });
  const order = p => (risk.find(r => r.emp === p.emp) ? 0 : hires.some(h => h.emp === p.emp) ? 1 : 2);
  const peopleOf = code => people.filter(p => p.code === code && Object.values(p.hrs).some(h => h > 0)).sort((a, b) => order(a) - order(b) || (a.k ?? 999) - (b.k ?? 999));
  const partialMonths = Object.keys(partial).filter(m => m !== curYm);
  // 入社月ごとの推移（直近6か月の入社）：入社月を1か月目として、月ごとの平均時間と勤務している人数
  function cohortsOf(list) {
    const by = {};
    list.forEach(p => { const h = hireOf[p.emp]; if (!h) return; const d = new Date(h.h), ci = d.getFullYear() * 12 + d.getMonth(); if (ci < curI - 5 || ci > curI) return;
      (by[ci] = by[ci] || []).push(p); });
    return Object.keys(by).map(Number).sort((a, b) => b - a).map(ci => {
      const xs = by[ci], steps = [];
      for (let k = 0; ci + k <= curI; k++) {
        const ym = `${Math.floor((ci + k) / 12)}-${String((ci + k) % 12 + 1).padStart(2, '0')}`;
        if (!useMonths.includes(ym) && ym !== curYm) { steps.push({ k, ym, n: null, avg: null }); continue; }
        const ws = xs.filter(p => (p.hrs[ym] || 0) > 0);
        steps.push({ k, ym, n: ws.length, avg: ws.length ? sum(ws, p => p.hrs[ym]) / ws.length : null, proj: ym === curYm && !!Object.keys(curRate).length });
      }
      const pts = steps.filter(x => x.avg != null);
      let trend = null;
      if (pts.length >= 2) { const a = pts[pts.length - 2].avg, b = pts[pts.length - 1].avg; trend = b >= a * 1.1 ? 'up' : b <= a * 0.9 ? 'down' : 'flat'; }
      const nowN = steps.length && steps[steps.length - 1].n != null ? steps[steps.length - 1].n : null;
      return { ci, ym: `${Math.floor(ci / 12)}-${String(ci % 12 + 1).padStart(2, '0')}`, n: xs.length, est: xs.filter(p => hireOf[p.emp].est).length, steps, trend, nowN };
    });
  }
  const allPeople = Object.values(P);
  const cohorts = cohortsOf(allPeople);
  const estN = working.filter(p => p.hireEst).length;
  return { cohorts, cohortsOf, allPeople, noWageStores, estN, partialMonths, peopleOf, curYm, months: useMonths, wMonths, hasCur: !!Object.keys(curRate).length, bands, total, ramp, hires, leavers, risk, stores, preLeave: pre.length >= 3 ? { n: pre.length, ratio: avgOf(pre) } : null, n: working.length, noHire: working.filter(p => p.k == null).length };
}
function tenureLines(T, st, isArea) {
  const out = [], push = (t, d) => out.push({ t, d: !!d });
  const B = (st ? st.bands : T.bands), n = st ? st.n : T.n;
  if (!n) return [{ t: '当月に勤務したパートがいません。', d: false }];
  const b = i => B[i];
  const hh = v => v == null ? '－' : Math.round(v) + 'h';
  push(`【勤務期間別（パート${n}名、${ymTxt(T.curYm)}${T.hasCur ? '・見込みで月換算' : ''}）】${B.filter(x => x.n && x.key !== 'tx').map(x => `${x.label}${x.n}名・平均${hh(x.avg)}`).join('、')}${b(4).n ? `、入社日不明${b(4).n}名` : ''}。`);
  // 戦力化：1か月〜半年の平均が1年以上の何割か
  if (b(1).n >= 2 && b(3).n >= 2 && b(3).avg) {
    const r = b(1).avg / b(3).avg;
    push(`・戦力化：入社1か月〜半年の平均${hh(b(1).avg)}は、1年以上（${hh(b(3).avg)}）の${Math.round(r * 100)}%${r < 0.6 ? '。新しい人の時間が少なく、戦力化が進んでいない可能性' : r >= 0.9 ? '。新しい人も1年以上の人と同じくらい勤務できている' : ''}。`);
  }
  const known = b(0).hours + b(1).hours + b(2).hours + b(3).hours;
  if (known) push(`　入社1年未満の人が労働時間の${Math.round((b(0).hours + b(1).hours + b(2).hours) / known * 100)}%を担っています（入社日のわかる人のうち）。`, true);
  const estN = st ? st.estN : T.estN;
  const nw = st ? (T.noWageStores.includes(st.code) ? [st.code] : []) : T.noWageStores;
  if (nw.length) push(`　※${st ? 'この店舗' : `${nw.length}店舗`}はパートの多くが時給資料に載っていないため、入社日を推定していません（時給資料が違う可能性。入社日不明として表示）。`);
  if (estN) push(`　※時給資料に載っていない${estN}名は前月入社として推定しています（タイムカードの履歴で入社月がわかる人はその月）。次の月の時給資料を入れると正しい入社日に置き換わります。`, true);
  const hires = st ? st.hires : T.hires, leavers = st ? st.leavers : T.leavers, risk = st ? st.risk : T.risk;
  const cs = (st ? st.cohorts : T.cohorts).filter(c => c.steps.some(x => x.avg != null));
  if (cs.length) {
    const TR = { up: '順調に伸びている', flat: '横ばい', down: '減っている・注意' };
    push(`・入社月ごとの推移：${cs.slice(0, 4).map(c => `${+c.ym.slice(5)}月入社${c.n}名${c.est ? `（うち推定${c.est}）` : ''}：${c.steps.filter(x => x.avg != null).map(x => `${x.k + 1}か月目${Math.round(x.avg)}h${x.n < c.n ? `（${x.n}名）` : ''}`).join('→')}${c.trend ? `（${TR[c.trend]}）` : ''}`).join('、')}。`);
    push('　※1か月目は月の途中からの勤務を含むため少なめに出ます。当月は見込みで月換算。', true);
  }
  const nm = p => `${isArea && p.store ? p.store + '・' : ''}${p.name}`;
  if (hires.length) { const e = hires.filter(p => p.hireEst).length;
    push(`・新規採用（当月・前月入社）：${hires.length}名${e ? `（うち${e}名は時給資料に未掲載のため入社月を推定）` : ''}。${hires.slice(0, 6).map(p => `${nm(p)}${hh(p.cur)}`).join('、')}${hires.length > 6 ? ' ほか' : ''}。`); }
  if (leavers.length) push(`・退社：${leavers.filter(l => l.sure).length}名${leavers.some(l => !l.sure) ? `、退社の可能性${leavers.filter(l => !l.sure).length}名（2か月勤務なし）` : ''}（${leavers.slice(0, 6).map(l => `${nm(l)}${l.k != null ? `・入社${l.k}か月` : ''}`).join('、')}${leavers.length > 6 ? ' ほか' : ''}）。`);
  const KIND = { days: '日数が減少', len: '1回が短く', both: '日数・1回とも', none: '今月勤務なし' };
  if (risk.length) push(`・退社リスク注意（入社1年以内で、月間の労働時間が減る見込み）：${risk.slice(0, 6).map(p => `${nm(p)}（入社${p.k}か月、${Math.round(p.cur)}h・${Math.round(p.days[T.curYm] || 0)}日${p.kind ? '、' + KIND[p.kind] : ''}）`).join('、')}${risk.length > 6 ? ` ほか${risk.length - 6}名` : ''}。`);
  if (risk.length) { const kc = k => risk.filter(x => x.kind === k).length, parts = [];
    if (kc('days')) parts.push(`日数が減っている${kc('days')}名は本人の都合（掛け持ち・学業・家庭など）の可能性→面談で確認`);
    if (kc('len')) parts.push(`1回が短くなっている${kc('len')}名はシフトの組み方（短い時間の割り当て）を確認`);
    if (kc('both')) parts.push(`日数・1回とも減っている${kc('both')}名は本人の事情とシフトの両方を確認`);
    if (kc('none')) parts.push(`今月勤務のない${kc('none')}名は在籍・連絡状況を確認`);
    push(`　減り方：${parts.join('。')}。`); }
  else push('・退社リスク注意：入社1年以内で時間が大きく減っている人はいません。');
  if (T.preLeave && !st) push(`　参考：これまでに退社した${T.preLeave.n}名は、最後の月の時間がそれまでの平均の${Math.round(T.preLeave.ratio * 100)}%でした。`, true);
  if (T.partialMonths && T.partialMonths.length && !st) push(`　※${T.partialMonths.map(ymTxt).join('・')}は月の途中までのタイムカードしかないため、確定日までの実績を月換算しています（月末後に最終のタイムカードを入れると置き換わります）。`);
  if (T.months.length < 3 && !st) push(`　※履歴が${T.months.length}か月分のため、退社の判定・時間の推移は今後の積み上げで精度が上がります。`, true);
  return out;
}

/* ===== 当月の着地見込み ===== */
const dtOfDate = date => { const wd = WEEKDAYS[(date.getDay() + 6) % 7]; return wd === '土' ? '土曜' : wd === '日' ? '日祝日' : wd === '金' ? '休前日' : '平日'; };
/**
 * 残り日の予算 × ここまでの曜日区分別の予算比（売上0の日は除く）で月末の売上を見込む。
 * 曜日区分の実績が2日未満なら全体の予算比を使う。
 * 売上集計表に月末までの予算がない日は、ここまでの同じ曜日区分の平均日予算で補う（祝日は判別できない）。
 * rows：その店の当月の全行（月末までの予算行を含む）、prevRows：前年同月の全行
 */
function forecastStore(s, rows, prevRows, lastDate, monthEnd) {
  const n = monthEnd.getDate() - lastDate.getDate();
  if (n <= 0) return null;
  const done = s.daily.filter(d => d.sales > 0 && d.budget > 0);
  const rateOf = ds => { const b = sum(ds, d => d.budget); return b ? sum(ds, d => d.sales) / b : null; };
  const pace = rateOf(done);
  if (pace == null) return null;
  const typeRate = {}, avgBudget = {}, custRate = {}, avgCustB = {};
  const cDone = done.filter(d => d.custBudget > 0), cRateOf = ds => { const b = sum(ds, d => d.custBudget); return b ? sum(ds, d => d.cust) / b : null; };
  const cPace = cRateOf(cDone);
  DAYTYPES.forEach(dt => {
    const ds = done.filter(d => d.dt === dt);
    typeRate[dt] = ds.length >= 2 ? rateOf(ds) : pace;
    const cs = cDone.filter(d => d.dt === dt);
    custRate[dt] = cs.length >= 2 ? cRateOf(cs) : cPace;
    const cb = s.daily.filter(d => d.dt === dt && d.custBudget > 0);
    avgCustB[dt] = cb.length ? sum(cb, d => d.custBudget) / cb.length : 0;
    const bs = s.daily.filter(d => d.dt === dt);
    avgBudget[dt] = bs.length ? sum(bs, d => d.budget) / bs.length : (s.daily.length ? sum(s.daily, d => d.budget) / s.daily.length : 0);
  });
  const fut = {};
  rows.filter(r => r.date > lastDate).forEach(r => { const k = +r.date; (fut[k] = fut[k] || { budget: 0, custBudget: 0, salesFc: 0, custFc: 0, dt: r.dt }); fut[k].budget += r.budget; fut[k].custBudget += r.custBudget || 0; fut[k].salesFc += r.salesFc || 0; fut[k].custFc += r.custFc || 0; });
  const days = [];
  for (let d = lastDate.getDate() + 1; d <= monthEnd.getDate(); d++) {
    const date = new Date(lastDate.getFullYear(), lastDate.getMonth(), d);
    const f = fut[+date];
    const dt = f ? f.dt : dtOfDate(date);
    const budget = f ? f.budget : avgBudget[dt];
    const cb = f && f.custBudget ? f.custBudget : avgCustB[dt];
    // 客数の見込み：客数予算×ここまでの曜日区分別の客数予算比。客数予算がなければ売上見込み÷ここまでの客単価
    let custFc = cb && custRate[dt] != null ? cb * custRate[dt] : s.up ? budget * typeRate[dt] / s.up : 0;
    let fc = budget * typeRate[dt], mgr = false;
    // 店長の日別予測があればそれを使う
    if (f && f.salesFc > 0) { fc = f.salesFc; mgr = true; if (f.custFc > 0) custFc = f.custFc; else if (s.up) custFc = fc / s.up; }
    days.push({ date, dt, budget, est: !f, fc, custFc, mgr });
  }
  const remBudget = sum(days, d => d.budget), remFc = sum(days, d => d.fc);
  const monthBudget = s.budget + remBudget, landing = s.sales + remFc, need = monthBudget - s.sales;
  // 参考：前年の残り期間の実績 × ここまでの前年比
  const pRem = prevRows.filter(r => r.date.getDate() > lastDate.getDate() && r.date.getDate() <= monthEnd.getDate());
  const prevBased = s.yoy && pRem.length ? s.sales + sum(pRem, r => r.sales) * s.yoy : null;
  return {
    remDays: n, days, estDays: days.filter(d => d.est).length,
    remByType: DAYTYPES.map(dt => ({ key: dt, days: days.filter(d => d.dt === dt).length })).filter(x => x.days),
    mgrDays: days.filter(d => d.mgr).length,
    remBudget, remFc, monthBudget, landing, landDiff: landing - monthBudget, landRate: monthBudget ? landing / monthBudget : null,
    need, needPerDay: need / n, needRate: remBudget ? need / remBudget : null, pace, typeRate,
    doneSales: sum(done, d => d.sales), doneBudget: sum(done, d => d.budget), remCust: sum(days, d => d.custFc),
    prevBased, prevBasedRate: prevBased != null && monthBudget ? prevBased / monthBudget : null,
  };
}
/** 店舗の見込みを合計して全体の見込みにする */
function forecastSum(list, area) {
  const fs = list.filter(Boolean); if (!fs.length) return null;
  const f0 = fs[0];
  const remBudget = sum(fs, f => f.remBudget), remFc = sum(fs, f => f.remFc);
  const monthBudget = area.budget + remBudget, landing = area.sales + remFc, need = monthBudget - area.sales;
  const days = f0.days.map((d, i) => ({ date: d.date, dt: d.dt, est: fs.some(f => f.days[i] && f.days[i].est), mgr: fs.every(f => f.days[i] && f.days[i].mgr), budget: sum(fs, f => f.days[i] ? f.days[i].budget : 0), fc: sum(fs, f => f.days[i] ? f.days[i].fc : 0), custFc: sum(fs, f => f.days[i] ? f.days[i].custFc : 0) }));
  const pb = fs.every(f => f.prevBased != null) ? sum(fs, f => f.prevBased) : null;
  return {
    remDays: f0.remDays, days, estDays: Math.max(...fs.map(f => f.estDays)), remByType: f0.remByType, mgrDays: Math.min(...fs.map(f => f.mgrDays || 0)), mgrStores: fs.filter(f => f.mgrDays).length, nStores: fs.length,
    remBudget, remFc, monthBudget, landing, landDiff: landing - monthBudget, landRate: monthBudget ? landing / monthBudget : null,
    need, needPerDay: need / f0.remDays, needRate: remBudget ? need / remBudget : null, pace: sum(fs, f => f.doneSales) / sum(fs, f => f.doneBudget),
    prevBased: pb, prevBasedRate: pb != null && monthBudget ? pb / monthBudget : null,
  };
}

const SCOPE_NAME = { direct: '直営', fc: 'FC', all: '全店' };
const isDirect = r => /^直営/.test(r.fc || ''), isFcStore = r => /^FC/i.test(r.fc || '');
/** 店舗が対象条件に合うか（FC区分がないデータは直営/FCの区別なし） */
function matchStore(r, scope, area) {
  if (area && r.area !== area) return false;
  if (scope === 'all' || !r.fc) return true;
  return scope === 'direct' ? isDirect(r) : isFcStore(r);
}
function listAreas(rows) { return [...new Set(rows.map(r => r.area).filter(Boolean))].sort(); }
/** 営業部ごとの直営／FC／全店の店舗数 */
function scopeCounts(rows, area) {
  const cnt = sc => new Set(rows.filter(r => matchStore(r, sc, area)).map(r => r.code)).size;
  return { direct: cnt('direct'), fc: cnt('fc'), all: cnt('all'), hasFc: rows.some(r => r.fc) };
}

function analyze(curRows, prevRows, extra) {
  extra = extra || {};
  let tc = extra.tc || [], wsch = extra.ws || [], wage = extra.wage || [];
  let hourly = extra.hourly || [], pays = extra.pays || [], pPays = extra.pPays || [], menu = extra.menu || [], pMenu = extra.pMenu || [], inv = extra.inv || [];
  // 対象店舗の絞り込み：営業部、直営／FC／全店
  const scope = extra.scope || 'all', areaSel = extra.area || '';
  const hasFc = curRows.some(r => r.fc);
  const areas = listAreas(curRows);
  const pick = new Set(curRows.filter(r => matchStore(r, scope, areaSel)).map(r => r.code));
  if (!pick.size) throw new Error(`${areaSel ? areaSel + 'に' : ''}${SCOPE_NAME[scope]}の店舗がありません`);
  curRows = curRows.filter(r => pick.has(r.code));
  prevRows = prevRows.filter(r => pick.has(r.code));
  hourly = hourly.filter(r => pick.has(r.code));
  pays = pays.filter(r => pick.has(r.code));
  pPays = pPays.filter(r => pick.has(r.code));
  menu = menu.filter(r => pick.has(r.code));
  inv = inv.filter(r => pick.has(r.code));
  pMenu = pMenu.filter(r => pick.has(r.code));
  const multiArea = !areaSel && areas.length > 1;
  const areaPart = areaSel ? areaSel + '・' : multiArea ? '全営業部・' : '';
  const scopeName = hasFc ? SCOPE_NAME[scope] : '全店';
  const scopeLabel = `${areaPart}${scopeName}計`;
  const titleLabel = `${areaPart}${scopeName}`;
  if (!curRows.length) throw new Error('当月データがありません');
  // 実績のある最終日まで
  const lastDate = new Date(Math.max(...curRows.filter(r => r.sales > 0).map(r => +r.date)));
  const firstDate = new Date(Math.min(...curRows.map(r => +r.date)));
  const cur = curRows.filter(r => r.date <= lastDate);
  const monthEnd = new Date(lastDate.getFullYear(), lastDate.getMonth() + 1, 0);
  const isFullMonth = lastDate.getDate() === monthEnd.getDate();
  // 前年は同じ日付範囲（月内の日）で比較
  const prev = prevRows.filter(r => r.date.getDate() <= lastDate.getDate());
  const fullBudget = {};
  curRows.forEach(r => { fullBudget[r.code] = (fullBudget[r.code] || 0) + r.budget; });

  const codes = [...new Set(cur.map(r => r.code))];
  const stores = codes.map(c => {
    const s = analyzeStore(cur.filter(r => r.code === c), prev.filter(r => r.code === c), monthEnd.getDate());
    s.fullBudget = fullBudget[c] || 0;
    return s;
  }).filter(s => s.budget > 0 || s.sales > 0);
  // エリア合計
  const area = analyzeStore(cur.map(r => ({ ...r, code: 'AREA', name: scopeLabel })), prev, monthEnd.getDate());
  area.fullBudget = Object.values(fullBudget).reduce((a, b) => a + b, 0);
  // 当月の着地見込み（月の途中のときだけ）
  stores.forEach(s => { s.fc = isFullMonth ? null : forecastStore(s, curRows.filter(r => r.code === s.code), prevRows.filter(r => r.code === s.code), lastDate, monthEnd); });
  area.fc = isFullMonth ? null : forecastSum(stores.map(s => s.fc), area);
  // 人：労働時間・雑給・残業
  let labor = null;
  if (tc.length || wsch.length) {
    const dtByDate = {}; curRows.forEach(r => { dtByDate[+r.date] = r.dt; });
    labor = analyzeLabor(tc, wsch, wage, { pick, dtByDate });
  } else labor = laborFromSummary(curRows, lastDate); // タイムカードなどがなければ売上集計表の労働時間（実績・日別予測）
    const byDate = {};
  cur.forEach(r => { if (r.sales > 0 && r.budget > 0) (byDate[+r.date] = byDate[+r.date] || []).push(r); });
  // 共通の不振日：対象店の85%以上が予算比85%未満、かつ同じ曜日区分の月平均より15pt以上低い日
  const typeRate = {};
  DAYTYPES.forEach(dt => { const rs = cur.filter(r => r.dt === dt && r.sales > 0); const b = sum(rs, r => r.budget); typeRate[dt] = b ? sum(rs, r => r.sales) / b : 1; });
  const commonWeak = Object.entries(byDate).filter(([, rs]) => rs.length >= 3
      && rs.filter(r => r.sales / r.budget < 0.85).length / rs.length >= 0.85
      && sum(rs, r => r.sales) / sum(rs, r => r.budget) <= typeRate[rs[0].dt] - 0.15)
    .map(([k, rs]) => ({ date: new Date(+k), wd: rs[0].wd, dt: rs[0].dt, rate: sum(rs, r => r.sales) / sum(rs, r => r.budget) }));
  stores.forEach(s => enrichStore(s, cur, hourly, pays, lastDate, pPays));
  // エリア直営平均の人時売上（区分×時間帯）
  const withLabor = stores.filter(s => s.laborGrid);
  const areaLabor = withLabor.length ? DAYTYPES.map((_, ti) => BANDS.map((_, bi) => ({ sales: sum(withLabor, s => s.laborGrid[ti][bi].sales), labor: sum(withLabor, s => s.laborGrid[ti][bi].labor) }))) : null;
  if (pays.length) {
    area.code = 'AREA';
    enrichStore(area, [], [], pays.map(p => ({ ...p, code: 'AREA' })), lastDate, pPays.map(p => ({ ...p, code: 'AREA' })));
  }
  area.laborStores = withLabor.length;
  // メニュー別：店舗ごとの1人あたり指標と、対象店平均（ベンチマーク）
  // 原価管理表（正式な原価率理論差）
  stores.forEach(s => { s.inv = inv.find(r => r.code === s.code) || null; });
  // 当月の在庫の動き：月初在庫＝前月の棚卸分析表の当月棚卸高。当月の棚卸がまだない（月の途中など）ときだけ
  const prevInv = (extra.prevInv || []).filter(r => pick.has(r.code));
  const stockOf = (x, open) => {
    if (open == null || x.inv || !(x.buy || x.cost)) return null;
    const use = x.cost || 0, loss = x.loss || 0, now = open + (x.buy || 0) - use - loss;
    const out = { open, buy: x.buy || 0, use, loss, now, delta: now - open };
    if (x.fc && x.sales) { // 月末までの見込み：着地売上（値引前）×理論原価率・ロス率
      const landPre = x.fc.landing * ((x.preSales || x.sales) / x.sales), rate = x.costRate || 0, lr = x.lossRate || 0;
      out.landPre = landPre; out.monthUse = landPre * rate; out.monthLoss = landPre * lr;
      out.remainBuy = out.monthUse + out.monthLoss - out.buy; out.remDays = x.fc.remDays;
    }
    return out;
  };
  stores.forEach(s => { const p = prevInv.find(r => r.code === s.code); s.stock = p ? stockOf(s, p.close) : null; });
  { const ws = stores.filter(s => s.stock); if (ws.length) { const T = k => sum(ws, s => s.stock[k] || 0); area.stock = { open: T('open'), buy: T('buy'), use: T('use'), loss: T('loss'), now: T('now'), delta: T('delta'), n: ws.length, remainBuy: ws.every(s => s.stock.remainBuy != null) ? T('remainBuy') : null, remDays: ws[0].stock.remDays }; } }
  const withInv = stores.filter(s => s.inv);
  if (withInv.length) {
    const T = (f) => sum(withInv, s => f(s.inv));
    const pre = T(i => i.preSales), cost = T(i => i.cost), th = T(i => i.theory * i.preSales), loss = T(i => i.loss);
    area.inv = { preSales: pre, cost, preRate: cost / pre, theory: th / pre, gap: (cost - th) / pre, gapAmt: cost - th, lossPre: loss / pre, gapExLoss: (cost - th - loss) / pre, open: T(i => i.open), close: T(i => i.close), n: withInv.length };
    const catSum = {}; withInv.forEach(s => s.inv.cats.forEach(c => { catSum[c.cat] = (catSum[c.cat] || 0) + c.cost; }));
    area.inv.catRate = Object.fromEntries(Object.entries(catSum).map(([k, v]) => [k, v / pre]));
  }
  // 店舗情報（経営企業名・担当AM）を付ける
  const master = extra.master && extra.master.length ? Object.fromEntries(extra.master.map(m => [m.code, m])) : null;
  if (master) stores.forEach(s => { const m = master[s.code]; if (m) { s.company = m.company; s.am = m.am; s.block = m.block; } });
  // 理論差（概算）＝仕入率−理論原価率
  stores.concat([area]).forEach(s => { s.gapRate = s.buyRate != null && s.costRate != null ? s.buyRate - s.costRate : null; s.pGapRate = s.pBuyRate != null && s.pCostRate != null ? s.pBuyRate - s.pCostRate : null; });
  const withGap = stores.filter(s => s.gapRate != null);
  area.benchGapRate = withGap.length ? (sum(withGap, s => s.buy) - sum(withGap, s => s.cost)) / sum(withGap, s => s.preSales) : null;
  const withLoss = stores.filter(s => s.lossRate != null);
  area.benchLossRate = withLoss.length ? sum(withLoss, s => s.loss) / sum(withLoss, s => s.preSales) : null;
  // 日別販売実績表から作ったメニュー別には客数がないので、同じ店・同じ期間の売上集計表の客数を入れる
  [menu, pMenu].forEach((ms, j) => { if (!ms.some(r => r.daily)) return; const rows = j ? prev : cur;
    const end = new Date(Math.max(...ms.map(r => +r.date))), from = new Date(Math.min(...ms.map(r => +r.from)));
    const custOf = {}; rows.forEach(r => { if (r.date >= from && r.date <= end) custOf[r.code] = (custOf[r.code] || 0) + (r.cust || 0); });
    ms.forEach(r => { if (r.daily && !r.cust) r.cust = custOf[r.code] || 0; }); });
  if (menu.length) {
    stores.forEach(s => { s.menu = menuAgg(menu.filter(r => r.code === s.code)); s.pMenuAgg = menuAgg(pMenu.filter(r => r.code === s.code)); });
    area.menu = menuAgg(menu); area.pMenuAgg = menuAgg(pMenu);
    const withMenu = stores.filter(s => s.menu && s.costRate != null);
    area.benchCostRate = withMenu.length ? sum(withMenu, s => s.cost) / sum(withMenu, s => s.preSales) : null;
    area.benchUp = withMenu.length ? sum(withMenu, s => s.preSales) / sum(withMenu, s => s.menu.cust) : null;
  }
  const ctx = { isFullMonth, lastDate, commonWeak, areaLabor, scopeLabel, multiArea, area, nStores: stores.length };
  // FC企業ごとの集計（店舗情報があるときだけ）
  let companies = null;
  if (master) {
    const by = {};
    stores.filter(s => s.company && !/^直営/.test(s.fcKind)).forEach(s => { (by[s.company] = by[s.company] || []).push(s); });
    companies = Object.entries(by).map(([name, ss]) => {
      const codes = new Set(ss.map(s => s.code));
      const c = analyzeStore(cur.filter(r => codes.has(r.code)).map(r => ({ ...r, code: 'CO', name })), prev.filter(r => codes.has(r.code)), monthEnd.getDate());
      c.name = name; c.stores = ss; c.codes = [...codes];
      c.ams = [...new Set(ss.map(s => s.am).filter(Boolean))];
      c.fullBudget = sum(ss, s => s.fullBudget || 0);
      c.fc = isFullMonth ? null : forecastSum(ss.map(s => s.fc), c);
      c.okN = ss.filter(s => s.diff >= 0).length;
      c.gapRate = c.buyRate != null && c.costRate != null ? c.buyRate - c.costRate : null;
      const wi = ss.filter(s => s.inv);
      if (wi.length) { const T = f => sum(wi, s => f(s.inv)); const pre = T(i => i.preSales), cost = T(i => i.cost), th = T(i => i.theory * i.preSales); c.inv = { preRate: cost / pre, theory: th / pre, gap: (cost - th) / pre, lossPre: T(i => i.loss) / pre, n: wi.length }; }
      c.lines = companyLines(c, isFullMonth);
      return c;
    }).sort((a, b) => (a.rate || 0) - (b.rate || 0));
  }

  stores.forEach(s => { s.lines = storeComment(s, ctx); s.costLines = costComment(s, ctx); });
  area.hasInv = withInv.length > 0;
  area.lines = areaComment(area, stores, ctx);
  if (labor) {
    // 売上の実績期間と勤怠の確定日が同じときだけ、実績の雑給率を出す
    const sameCut = labor.cut === lastDate.getDate() && labor.mo === lastDate.getMonth();
    const std = extra.std || null, cutTxt = `${labor.mo + 1}/${labor.cut}時点`;
    stores.forEach(s => {
      s.labor = labor.stores.find(x => x.code === s.code) || null;
      s.std = std && s.labor ? analyzeStd(s, s.labor, labor, std, extra.stdSettings, lastDate.getMonth() === labor.mo ? lastDate.getDate() : 0, curRows.filter(r => r.code === s.code)) : null;
      // 予測月（実績のない月）：客数・売上は見込み。雑給率は売上見込みで
      if (extra.planMode && s.std && s.std.days) s.std.days.forEach(d => { d.custEst = true; });
      s.laborLines = s.labor ? laborLines(s.labor, labor, extra.planMode ? { landing: s.sales } : s.fc, sameCut ? s.sales : null, ctx).concat(stdLines(s.std, { cutTxt })) : [{ t: '勤怠データにこの店舗がありません。', d: false }];
    });
    const a = labor.area, fs = stores.filter(s => s.labor);
    area.labor = a;
    const full = fs.length === stores.length && !a.wageless;
    area.laborLines = laborLines(a, labor, full ? (extra.planMode ? { landing: area.sales } : area.fc) : null, sameCut && full ? area.sales : null, ctx);
    // 標準比（対象店の合計）
    const ws = stores.filter(s => s.std && !s.std.missing);
    if (ws.length) {
      const T2 = k => sum(ws, s => s.std[k] || 0);
      const A = { cutRate: T2('cutStd') ? T2('cutHours') / T2('cutStd') : null, landRate: T2('landStd') ? T2('landHours') / T2('landStd') : null, n: ws.length };
      const wb = ws.filter(s => s.std.budgetRate != null);
      A.budgetRate = wb.length ? sum(wb, s => s.std.target) / sum(wb, s => s.std.landStd) : null;
      A.over = wb.length ? sum(wb, s => s.std.over) : null;
      A.planRate = T2('planStd') ? T2('planHours') / T2('planStd') : null; A.planDays = T2('planDays');
      A.overDays = sum(ws, s => (s.std.overDays || []).length);
      const acc = ws.filter(s => s.std.fcAcc); A.fcAcc = acc.length ? sum(acc, s => s.std.fcAcc.ratio * s.std.fcAcc.n) / sum(acc, s => s.std.fcAcc.n) : null;
      area.std = A;
      const rk = [...ws].sort((x, y) => (y.std.landRate - (y.std.budgetRate || 0)) - (x.std.landRate - (x.std.budgetRate || 0)));
      area.laborLines.push({ t: `・標準比：${labor.cut ? `実績${p0(A.cutRate)}（${cutTxt}）${labor.cut < labor.N ? `、着地見込み${p0(A.landRate)}` : ''}` : `勤務予定どおりなら${p0(A.landRate)}`}${A.budgetRate != null ? `（予算標準比${p0(A.budgetRate)}）` : ''}。${ws.length > 1 ? `予算標準比を最も上回る見込みの店：${rk.filter(s => s.std.budgetRate != null && s.std.over > 0).slice(0, 3).map(s => `${s.name}${p0(s.std.landRate)}（予算${p0(s.std.budgetRate)}、${hrs0(s.std.over * 60)}超過）`).join('、') || 'なし'}。` : ''}`, d: false });
      if (A.planRate != null) area.laborLines.push({ t: `・計画標準比（店長の客数予測×計画労働時間）：${p0(A.planRate)}。これからの日で予算標準比を超える計画：${A.overDays}日（${ws.filter(s => (s.std.overDays || []).length).sort((a, b) => b.std.overDays.length - a.std.overDays.length).slice(0, 4).map(s => `${s.name}${s.std.overDays.length}日`).join('、') || 'なし'}）${A.fcAcc != null ? `。客数の実績は予測の${p0(A.fcAcc)}` : ''}。`, d: false });
      if (ws.some(s => !s.std.custom)) area.laborLines.push({ t: `　※店舗設定が初期値のままの店：${ws.filter(s => !s.std.custom).map(s => s.name).join('、')}。`, d: true });
    } else if (std) area.laborLines.push({ t: '・標準比：標準比マスタに対象店の卓数がありません。', d: false });
    if (labor.hasWage && fs.length > 1) {
      const rate = s => s.labor.wageless ? null : s.fc ? s.labor.landCost / s.fc.landing : sameCut ? s.labor.cost / s.sales : null;
      const rk = fs.filter(s => rate(s) != null).sort((x, y) => rate(y) - rate(x));
      const fmt = xs => xs.map(s => `${s.name}${pct(rate(s))}`).join('、');
      if (rk.length >= 5) area.laborLines.push({ t: `・雑給率の高い店（${area.fc ? '着地見込み' : '実績'}）：${fmt(rk.slice(0, 3))}／低い店：${fmt(rk.slice(-2).reverse())}。`, d: false });
      else if (rk.length) area.laborLines.push({ t: `・店舗別の雑給率（${area.fc ? '着地見込み' : '実績'}、高い順）：${fmt(rk)}。`, d: false });
    }
    const est = fs.filter(s => s.labor.estDays).map(s => `${s.name}${s.labor.estDays}日`);
    if (est.length) area.laborLines.push({ t: `・${labor.fromSummary ? '計画の労働時間がない日' : '勤務予定の未作成'}：${est.join('、')}（実績平均で補っています）。`, d: true });
  }
  area.costLines = areaCostComment(area, stores, ctx);
  // 採用・定着（履歴があるとき）
  let tenure = null;
  if (extra.history && Object.keys(extra.history.tc || {}).length) {
    tenure = analyzeTenure(extra.history, labor, { pick });
    if (tenure) {
      const nameOf = c => (stores.find(s => s.code === c) || {}).name || c;
      ['hires', 'leavers', 'risk'].forEach(k => tenure[k].forEach(p => { p.store = nameOf(p.code); }));
      area.tenureLines = tenureLines(tenure, null, stores.length > 1);
      stores.forEach(s => { const st = tenure.stores.find(x => x.code === s.code); s.tenure = st || null; s.tenureLines = st ? tenureLines(tenure, st, false) : [{ t: '当月に勤務したパートの履歴がありません。', d: false }]; });
    }
  }
  return { companies, hasMaster: !!master, tenure, labor, stores, area, firstDate, lastDate, isFullMonth, commonWeak, areaLabor, scopeLabel, titleLabel, multiArea, areas, hasHourly: hourly.length > 0,  hasPays: pays.length > 0 };
}

/* ===== コメント生成 ===== */
const man = v => { const m = v / 10000; return (Math.abs(m) >= 100 ? Math.round(m).toLocaleString('ja-JP') : (Math.round(m * 10) / 10).toLocaleString('ja-JP')) + '万円'; };
const sman = v => (v >= 0 ? '+' : '▲') + man(Math.abs(v));
const pct = v => v == null ? '－' : (v * 100).toFixed(1) + '%';
const yen = v => Math.round(v).toLocaleString('ja-JP') + '円';
const syen = v => (v >= 0 ? '+' : '▲') + Math.round(Math.abs(v)).toLocaleString('ja-JP') + '円';
const md = d => `${d.getMonth() + 1}/${d.getDate()}`;
/** 1万円未満は円で（▲0万円のように見えないように） */
const smanS = v => Math.abs(v) < 10000 ? syen(v) : sman(v);

/** コメント行の配列を返す。d:true は詳細版のみ */
function storeComment(s, ctx) {
  const L = [];
  L.d = t => L.push({ t, d: true });
  const P = L.push.bind(L);
  L.push = t => P(typeof t === 'string' ? { t, d: false } : t);
  const ok = s.diff >= 0;
  const period = ctx.isFullMonth ? '' : `（${md(ctx.lastDate)}時点）`;
  L.push(`【${ok ? '予算達成' : '予算未達'}】売上${man(s.sales)}／予算比${pct(s.rate)}（${sman(s.diff)}）${s.yoy ? `・前年比${pct(s.yoy)}` : ''}${period}`);
  for (const x of forecastLines(s.fc, ctx)) L.push(x);

  // 1) 客数・単価の分解
  const main = Math.abs(s.custEff) >= Math.abs(s.upEff) ? 'cust' : 'up';
  const custTxt = `客数${s.cust.toLocaleString('ja-JP')}人（予算比${pct(s.custRate)}）`;
  const upTxt = `客単価${yen(s.up)}（予算${syen(s.up - s.bUp)}）`;
  let lead;
  if (ok) lead = main === 'cust' ? '客数増が達成の主因' : '客単価の上振れが達成の主因';
  else lead = main === 'cust' ? (s.custEff < 0 ? '客数不足が未達の主因' : '単価減が未達の主因') : (s.upEff < 0 ? '客単価の下振れが未達の主因' : '客数不足が未達の主因');
  L.push(`・要因分解：客数要因${sman(s.custEff)}、単価要因${sman(s.upEff)}。${custTxt}、${upTxt}で、${lead}。`);
  if (s.bUp && s.pUp && Math.abs(s.bUp - s.pUp) / s.pUp >= 0.12)
    L.push(`　※予算の想定客単価${yen(s.bUp)}は前年実績${yen(s.pUp)}と${syen(s.bUp - s.pUp)}乖離しており、客数予算・売上予算の前提確認が必要。`);
  const upMatters = s.upEff < 0 || main === 'up';
  for (const t of unitPriceLines(s)) typeof t === 'object' ? L.push(t) : upMatters ? L.push(t) : L.d(t);

  // 2) 曜日区分
  const types = s.byType.filter(t => t.days > 0 && t.budget > 0);
  const neg = types.filter(t => t.diff < 0).sort((a, b) => a.diff - b.diff);
  const pos = types.filter(t => t.diff >= 0).sort((a, b) => b.diff - a.diff);
  const tTxt = t => `${t.key}（${t.days}日）${pct(t.rate)}・${sman(t.diff)}`;
  if (!ok && neg.length) {
    const share = s.diff < 0 ? neg[0].diff / sum(neg, t => t.diff) : 0;
    let line = `・曜日区分：${neg.slice(0, 2).map(tTxt).join('、')}が未達`;
    if (neg.length > 1 && share >= 0.5) line += `（${neg[0].key}で未達額の${Math.round(share * 100)}%）`;
    if (pos.length) line += `。一方${pos.map(t => `${t.key}は${pct(t.rate)}`).join('、')}と予算クリア`;
    L.push(line + '。');
  } else if (ok && pos.length) {
    let line = `・曜日区分：${pos.slice(0, 2).map(tTxt).join('、')}が牽引`;
    if (neg.length) line += `。${neg.map(t => `${t.key}は${pct(t.rate)}（${sman(t.diff)}）`).join('、')}と課題`;
    L.push(line + '。');
  }
  // 曜日（最も弱い曜日）
  const wds = s.byWd.filter(w => w.budget > 0).sort((a, b) => (a.rate ?? 9) - (b.rate ?? 9));
  if (wds.length >= 3) L.d(`　曜日別では${wds[0].key}曜が最も低く${pct(wds[0].rate)}、最も高いのは${wds[wds.length - 1].key}曜${pct(wds[wds.length - 1].rate)}。`);

  // カレンダー差（前年比較）
  if (s.hasPrev) {
    const cal = types.filter(t => t.pDays !== t.days && t.pDays > 0).map(t => `${t.key}${t.days - t.pDays > 0 ? '+' : ''}${t.days - t.pDays}日`);
    if (cal.length) L.d(`　前年比の暦差：${cal.join('、')}（前年比を見る際は考慮）。`);
  }

  // 3) 時間帯：前年データがあれば前年比で見る（時間帯の予算は按分のため参考程度）
  const hasPY = s.hasPrev && s.byBand.some(b => b.pSales > 0);
  if (hasPY) {
    const bs = s.byBand.filter(b => b.pSales > s.pSales * 0.02 || b.sales > s.sales * 0.02);
    const neg = bs.filter(b => b.pDiff < 0).sort((a, b) => a.pDiff - b.pDiff), pos = bs.filter(b => b.pDiff >= 0).sort((a, b) => b.pDiff - a.pDiff);
    const t = b => `${b.name}(${b.label})${pct(b.yoy)}・${sman(b.pDiff)}`;
    let line = '・時間帯別（前年比）：' + (neg.length ? `${neg.slice(0, 2).map(t).join('、')}が前年割れ` : `全時間帯で前年超え（${pos.slice(0, 2).map(t).join('、')}）`);
    if (neg.length && pos.length) line += `。${pos[0].name}は${pct(pos[0].yoy)}と前年超え`;
    L.push(line + '。');
    // 曜日区分×時間帯で前年から最も落ちているところ（1日あたりの差×日数＝月への影響が最大）
    let w = null;
    s.grid.forEach((row, ti) => row.forEach((c, bi) => { if (c.yoyDay == null || c.dayDiff == null) return; const imp = c.dayDiff * c.d; if (!w || imp < w.imp) w = { imp, ti, bi, c }; }));
    if (w && w.imp < 0 && Math.abs(w.imp) >= Math.abs(Math.min(0, s.sales - s.pSales)) * 0.2)
      L.push(`　前年から最も落ちているのは「${DAYTYPES[w.ti]}×${BANDS[w.bi].name}」で、1日あたり前年比${pct(w.c.yoyDay)}（1日${smanS(w.c.dayDiff)}）。`);
  }
  const bands = hasPY ? [] : s.byBand.filter(b => b.budget > s.budget * 0.02);
  const bNeg = bands.filter(b => b.diff < 0).sort((a, b) => a.diff - b.diff);
  const bPos = bands.filter(b => b.diff >= 0).sort((a, b) => b.diff - a.diff);
  const bTxt = b => `${b.name}(${b.label})${pct(b.rate)}・${sman(b.diff)}${b.yoy ? `／前年比${pct(b.yoy)}` : ''}`;
  if (bands.length) {
    let line = '・時間帯別（予算は前年の時間帯構成比で按分）：';
    if (!ok && bNeg.length) {
      line += `${bNeg.slice(0, 2).map(bTxt).join('、')}が未達`;
      if (bPos.length) line += `。${bPos[0].name}は${pct(bPos[0].rate)}と好調`;
    } else if (ok && bPos.length) {
      line += `${bPos.slice(0, 2).map(bTxt).join('、')}が上振れ`;
      if (bNeg.length) line += `。${bNeg[0].name}は${pct(bNeg[0].rate)}と弱含み`;
    } else {
      line += bands.map(b => `${b.name}${pct(b.rate)}`).join('、');
    }
    L.push(line + '。');
    // 区分×時間帯で最大の未達セル
    var worstCell = null;
    s.grid.forEach((row, ti) => row.forEach((c, bi) => { const d = c.sales - c.budget; if (c.budget > 0 && (!worstCell || d < worstCell.d)) worstCell = { d, ti, bi, r: c.sales / c.budget }; }));
    if (worstCell && worstCell.d < 0 && Math.abs(worstCell.d) >= Math.abs(s.diff) * 0.2)
      L.push(`　最大の未達ポイントは「${DAYTYPES[worstCell.ti]}×${BANDS[worstCell.bi].name}」で${pct(worstCell.r)}・${sman(worstCell.d)}。`);
  }



  // 5) 特異日
  const cw = (ctx.commonWeak || []).map(c => {
    const d = s.daily.find(x => +x.date === +c.date); return d ? { ...c, diff: d.diff, srate: d.rate } : null; }).filter(c => c && c.srate < 0.9);
  if (cw.length) {
    const cwDiff = sum(cw, c => Math.min(0, c.diff));
    if (cwDiff < 0) L.push(`・共通の不振日：${cw.map(c => `${md(c.date)}(${c.wd}${c.dt === '日祝日' && c.wd !== '日' ? '・祝' : ''})自店${pct(c.srate)}`).join('、')}は他店も軒並み低調で外部要因の可能性（自店影響${sman(cwDiff)}${!ok && s.diff < 0 ? `、除くと予算比${pct((s.sales - sum(cw, c => s.daily.find(x => +x.date === +c.date).sales)) / (s.budget - sum(cw, c => s.daily.find(x => +x.date === +c.date).budget)))}` : ''}）。`);
  }
  const cwSet = new Set(cw.map(c => +c.date));
  const worstOwn = s.worst.filter(d => !cwSet.has(+d.date) && d.sales > 0);
  if (!ok && worstOwn.length) L.d(`・未達が大きかった日（店舗固有）：${worstOwn.map(d => `${md(d.date)}(${d.wd})${pct(d.rate)}・${sman(d.diff)}`).join('、')}。`);
  if (ok && s.best.length) L.d(`・上振れが大きかった日：${s.best.map(d => `${md(d.date)}(${d.wd})${pct(d.rate)}・${sman(d.diff)}`).join('、')}。`);
  if (s.closed.length) {
    const cb = sum(s.closed, d => d.budget);
    L.push(`・売上0の日あり：${s.closed.map(d => md(d.date)).join('、')}（予算${man(cb)}が丸ごと未達計上。休業かデータ欠落か要確認。除くと予算比${pct(s.sales / (s.budget - cb))}）。`);
  }

  return L.filter(x => x && x.t);
}

const pct2 = v => v == null ? '－' : (v * 100).toFixed(2) + '%';
/** ロス（廃棄額）の考察 */
function lossLines(s, ctx) {
  if (s.lossRate == null) return [];
  const cmp = [];
  const b = ctx.area && ctx.area.benchLossRate;
  if (ctx.nStores > 1 && b != null) cmp.push(`対象店平均${pct2(b)}`);
  if (s.pLossRate != null) cmp.push(`前年${pct2(s.pLossRate)}`);
  let t = `・ロス：ロス額${man(s.loss)}・ロス率${pct2(s.lossRate)}${cmp.length ? `（${cmp.join('、')}）` : ''}`;
  if (s.lossSpikes.length) t += `。突出日は${s.lossSpikes.map(d => `${md(d.date)}(${d.wd})${man(d.loss)}（売上比${pct(d.rate)}）`).join('、')}で、月間ロスの${Math.round(sum(s.lossSpikes, d => d.loss) / s.loss * 100)}%`;
  // 曜日の偏り：1曜日に30%以上集中
  const top = [...s.lossByWd].sort((a, b) => b.loss - a.loss)[0];
  if (s.loss > 0 && top.loss / s.loss >= 0.3) t += `。曜日別では${top.key}曜に${Math.round(top.loss / s.loss * 100)}%が集中`;
  return [{ t: t + '。', d: false }];
}
const sg = (v, d = 1) => (v >= 0 ? '+' : '▲') + Math.abs(v).toFixed(d);
const pt = v => (v >= 0 ? '+' : '▲') + Math.abs(v * 100).toFixed(1) + 'pt';
/** 原価率・喫食量・皿数の考察（対象店平均・前年と比較） */
function costLines(s, ctx) {
  const m = s.menu, A = ctx.area && ctx.area.menu; if (!m || !A) return [];
  const out = [];
  const bench = ctx.nStores > 1;
  const bRate = ctx.area.benchCostRate, bUp = ctx.area.benchUp;
  if (s.costRate != null) {
    let t = `・原価：理論原価率${pct(s.costRate)}`;
    const cmp = [];
    if (bench && bRate != null) cmp.push(`対象店平均${pct(bRate)}・${pt(s.costRate - bRate)}`);
    if (s.pCostRate != null) cmp.push(`前年${pct(s.pCostRate)}・${pt(s.costRate - s.pCostRate)}`);
    if (cmp.length) t += `（${cmp.join('、')}）`;
    t += `。1人あたり原価${yen(m.costPP)}`;
    // 平均との差を「1人あたり原価（喫食）」と「客単価」に分解
    if (bench && bRate != null && bUp) {
      const up = s.preSales / m.cust, cppA = bRate * bUp;
      const effEat = (m.costPP - cppA) / bUp, effUp = m.costPP / up - m.costPP / bUp;
      if (Math.abs(s.costRate - bRate) >= 0.003) t += `。平均との差は、喫食（1人あたり原価）で${pt(effEat)}、客単価で${pt(effUp)}`;
    }
    // 前年差も同様に分解（売上集計表の前年の原価・客数から）
    if (s.pCostRate != null && s.pCust && s.cust && Math.abs(s.costRate - s.pCostRate) >= 0.003) {
      const cpp = s.cost / s.cust, pCpp = s.pCost / s.pCust;
      const up = s.preSales / s.cust, pUp = s.pPreSales / s.pCust;
      const effEat = (cpp - pCpp) / pUp, effUp = cpp / up - cpp / pUp;
      t += `。前年差は、1人あたり原価${syen(cpp - pCpp)}で${pt(effEat)}、値引前客単価${syen(up - pUp)}で${pt(effUp)}`;
    }
    out.push({ t: t + '。', d: false });
  }
  const P = s.pMenuAgg;
  const f = (v, d) => d ? v.toFixed(d) : Math.round(v).toLocaleString('ja-JP');
  const vs = (v, a, p, unit, d) => { const c = []; if (bench) c.push(`平均${f(a, d)}${unit}`); if (P) c.push(`前年${f(p, d)}${unit}`); return c.length ? `（${c.join('・')}）` : ''; };
  // 前年メニュー別があれば、1人あたり原価の前年差を「量（皿数）」と「原価単価（原価改定など）」に分解
  if (P) {
    let qtyEff = 0, priceEff = 0;
    Object.values(m.items).forEach(it => {
      const p = P.items[it.name];
      const uc = it.qty ? it.cost / it.qty : 0, puc = p && p.qty ? p.cost / p.qty : uc;
      qtyEff += (it.qtyPP - (p ? p.qtyPP : 0)) * puc; priceEff += it.qtyPP * (uc - puc);
    });
    Object.values(P.items).forEach(p => { if (!m.items[p.name]) qtyEff -= p.qtyPP * (p.qty ? p.cost / p.qty : 0); });
    out.push({ t: `　1人あたり原価の前年差${syen(m.costPP - P.costPP)}の内訳：量（皿数・杯数）で${syen(qtyEff)}、原価単価（原価改定など）で${syen(priceEff)}。`, d: false });
  }
  out.push({ t: `　食べ放題フード：1人${m.foodQty.toFixed(1)}皿${vs(m.foodQty, A.foodQty, P && P.foodQty, '皿', 1)}${m.foodG ? `・${Math.round(m.foodG).toLocaleString('ja-JP')}g${vs(m.foodG, A.foodG, P && P.foodG, 'g', 0)}` : ''}。うち精肉${m.meatQty.toFixed(2)}皿${m.meatG ? `・${Math.round(m.meatG)}g` : ''}${vs(m.meatQty, A.meatQty, P && P.meatQty, '皿', 2)}。ドリンク${m.drinkQty.toFixed(1)}杯/人。`, d: false });
  // 分類別：平均より1人あたり原価が高い分類
  if (bench) {
    const diffs = m.cats.map(c => { const a = A.cats.find(x => x.cat === c.cat); return a ? { ...c, dCost: c.costPP - a.costPP, dQty: c.qtyPP - a.qtyPP } : null; }).filter(Boolean)
      .filter(c => Math.abs(c.dCost) >= 5).sort((a, b) => b.dCost - a.dCost);
    const up = diffs.filter(c => c.dCost > 0).slice(0, 3), down = diffs.filter(c => c.dCost < 0).slice(-2).reverse();
    if (up.length || down.length) out.push({ t: `　平均との差（1人あたり原価）：${[...up, ...down].map(c => `${c.cat}${syen(c.dCost)}（${sg(c.dQty, 2)}皿）`).join('、')}。`, d: false });
    // 品目：平均より1人あたり原価が高い上位
    const items = Object.values(m.items).map(it => { const a = A.items[it.name]; return a ? { ...it, d: it.costPP - a.costPP } : null; }).filter(x => x && x.d >= 5).sort((a, b) => b.d - a.d).slice(0, 3);
    if (items.length) out.push({ t: `　平均より出数が多い品目：${items.map(it => `${it.name.replace(/^Ｆ/, '')}${syen(it.d)}/人`).join('、')}。`, d: true });
  }
  if (m.courses.length >= 2) {
    const avgOf = name => { const a = A.courses.find(c => c.name === name); return a ? a.share : 0; };
    const prevOf = name => { const p = P && P.courses.find(c => c.name === name); return p ? p.share : null; };
    const list = m.courses.filter(c => c.share >= 0.01).slice(0, 5).map(c => {
      const cmp = [bench ? `平均${pct(avgOf(c.name))}` : '', prevOf(c.name) != null ? `前年${pct(prevOf(c.name))}` : ''].filter(Boolean).join('・');
      return `${c.name}${pct(c.share)}${cmp ? `（${cmp}）` : ''}`;
    });
    out.push({ t: `　コース構成（年齢区分を合算）：${list.join('、')}。`, d: true });
  }
  if (m.drink) {
    const dcmp = (k) => [bench && A.drink ? `平均${pct(A.drink[k])}` : '', P && P.drink ? `前年${pct(P.drink[k])}` : ''].filter(Boolean).join('・');
    const part = (n, k, note) => m.drink[k] == null ? '' : `${n}${pct(m.drink[k])}${note}${dcmp(k) ? `（${dcmp(k)}）` : ''}`;
    out.push({ t: `・飲み放題比率（コースの人数あたり）：${[part('ソフトドリンク', 'softRate', ''), part('ランチソフト', 'lunchRate', ''), part('アルコール', 'alcRate', '')].filter(Boolean).join('、')}。`, d: false });
  }
  return out;
}

/** 客単価の前年差を「値引前単価」と「値引」に分解し、値引内訳を添える */
function unitPriceLines(s) {
  const out = [];
  if (s.hasPrev && s.pCust && s.cust) {
    const preUp = (s.sales + s.disc) / s.cust, pPreUp = (s.pSales + s.pDisc) / s.pCust;
    const dpc = s.disc / s.cust, pDpc = s.pDisc / s.pCust;
    const discEff = -(dpc - pDpc), preEff = preUp - pPreUp;
    if (Math.abs(discEff) >= 15 || Math.abs(preEff) >= 30)
      out.push(`　客単価は前年比${syen(s.up - s.pUp)}：値引前単価${syen(preEff)}、値引${discEff < 0 ? '増加' : '減少'}で${syen(discEff)}（値引率${pct(s.pDiscRate)}→${pct(s.discRate)}）。`);
  }
  const dd = s.discDetail;
  if (dd && dd.total > 0) {
    const cats = dd.cats.slice(0, 3).map(c => `${c.cat}${Math.round(c.amount / dd.total * 100)}%${c.cat === 'チラシ' ? `（利用${c.count}件）` : ''}`).join('、');
    const top = dd.items[0];
    out.push(`　値引内訳（税込${man(dd.total)}）：${cats}。最大は「${top.name}」${man(top.amount)}・${top.count}件。`);
  }
  return out.concat(discYoyLines(s));
}

/** 値引の前年差（誘客の差と客単価への影響） */
function discYoyLines(s) {
  const d = s.discCmp; if (!d) return [];
  const out = [];
  const movers = [...d.cats].filter(c => c.lure && c.dCount !== 0).sort((a, b) => Math.abs(b.dCount) - Math.abs(a.dCount)).slice(0, 3)
    .map(c => `${c.cat}${c.dCount > 0 ? '+' : '▲'}${Math.abs(c.dCount).toLocaleString('ja-JP')}件`);
  let line = `・値引による誘客（従業員食事を除く）：利用${d.count.toLocaleString('ja-JP')}件（前年${d.pCount.toLocaleString('ja-JP')}件、${d.dCount >= 0 ? '+' : '▲'}${Math.abs(d.dCount).toLocaleString('ja-JP')}件）`;
  if (movers.length) line += `。増減の大きい分類：${movers.join('、')}`;
  if (d.gp && d.dCount) line += `。1組${d.gp.toFixed(1)}人換算で客数${d.dCount > 0 ? '+' : '▲'}${Math.round(Math.abs(d.dCount) * d.gp).toLocaleString('ja-JP')}人相当（推計）`;
  out.push({ t: line + '。', d: false });
  const eff = [...d.cats].filter(c => Math.abs(c.upEff) >= 3).sort((a, b) => Math.abs(b.upEff) - Math.abs(a.upEff)).slice(0, 3);
  if (eff.length) out.push({ t: `　客単価への影響（1人あたり値引の前年差）：${eff.map(c => `${c.cat}${syen(c.upEff)}`).join('、')}。`, d: true });
  return out;
}

function areaComment(a, stores, ctx) {
  const L = [];
  const P = L.push.bind(L);
  L.push = t => P({ t, d: false });
  const okN = stores.filter(s => s.diff >= 0).length;
  const period = ctx.isFullMonth ? '' : `（${md(ctx.lastDate)}時点）`;
  L.push(`【${ctx.scopeLabel}】売上${man(a.sales)}／予算比${pct(a.rate)}（${sman(a.diff)}）${a.yoy ? `・前年比${pct(a.yoy)}` : ''}${period}。予算達成${okN}店／${stores.length}店。`);
  if (a.fc) {
    for (const x of forecastLines(a.fc, ctx)) if (!x.d) L.push(x.t);
    const fs = stores.filter(x => x.fc);
    const okF = fs.filter(x => x.fc.landDiff >= 0).length;
    L.push(`　着地で予算達成見込み${okF}店／${fs.length}店。`);
    // 挽回が必要な店：着地で未達見込み、残り期間に必要な予算比が高い順
    const hard = fs.filter(x => x.fc.landDiff < 0 && x.fc.need > 0).sort((x, y) => y.fc.needRate - x.fc.needRate).slice(0, 3);
    if (hard.length) L.push(`・挽回が必要な店（残り期間に必要な予算比が高い順）：${hard.map(x => `${x.name}（必要${pct(x.fc.needRate)}・1日${man(x.fc.needPerDay)}、着地見込み${sman(x.fc.landDiff)}）`).join('、')}。`);
  }
  L.push(`・要因分解：客数要因${sman(a.custEff)}（客数予算比${pct(a.custRate)}）、単価要因${sman(a.upEff)}（客単価${yen(a.up)}、予算${syen(a.up - a.bUp)}）。`);
  if (ctx.multiArea) {
    const by = {};
    stores.forEach(s => { const k = s.area || '（営業部なし）'; (by[k] = by[k] || { k, sales: 0, budget: 0, n: 0, ok: 0 }); by[k].sales += s.sales; by[k].budget += s.budget; by[k].n++; if (s.diff >= 0) by[k].ok++; });
    const list = Object.values(by).sort((x, y) => x.sales / x.budget - y.sales / y.budget);
    L.push(`・営業部別：${list.map(x => `${x.k}${pct(x.sales / x.budget)}（${sman(x.sales - x.budget)}、達成${x.ok}／${x.n}店）`).join('、')}。`);
  }
  const worst = [...stores].sort((x, y) => x.diff - y.diff).filter(s => s.diff < 0).slice(0, 3);
  const best = [...stores].sort((x, y) => y.diff - x.diff).filter(s => s.diff > 0).slice(0, 3);
  if (worst.length) L.push(`・未達額上位：${worst.map(s => `${s.name}${sman(s.diff)}（${pct(s.rate)}）`).join('、')}。`);
  if (best.length) L.push(`・達成額上位：${best.map(s => `${s.name}${sman(s.diff)}（${pct(s.rate)}）`).join('、')}。`);
  const t = a.byType.filter(x => x.budget > 0).sort((x, y) => x.diff - y.diff);
  if (t.length) L.push(`・曜日区分：${t.map(x => `${x.key}${pct(x.rate)}`).join('、')}。`);
  if (a.hasPrev && a.byBand.some(x => x.pSales > 0)) {
    const b = a.byBand.filter(x => x.pSales > a.pSales * 0.02).sort((x, y) => x.pDiff - y.pDiff);
    if (b.length) L.push(`・時間帯（前年比）：${b.map(x => `${x.name}${pct(x.yoy)}（${sman(x.pDiff)}）`).join('、')}。`);
  } else {
    const b = a.byBand.filter(x => x.budget > a.budget * 0.02).sort((x, y) => x.diff - y.diff);
    if (b.length) L.push(`・時間帯：${b.map(x => `${x.name}${pct(x.rate)}（${sman(x.diff)}）`).join('、')}。`);
  }
  if (a.discDetail) {
    const dd = a.discDetail;
    L.push(`・値引：税込${man(dd.total)}（値引率${pct(a.discRate)}${a.pDiscRate != null ? `、前年${pct(a.pDiscRate)}` : ''}）。内訳は${dd.cats.slice(0, 4).map(c => `${c.cat}${Math.round(c.amount / dd.total * 100)}%`).join('、')}。`);
    for (const x of discYoyLines(a)) if (!x.d) L.push(x.t);
  }
  if (ctx.commonWeak.length) L.push(`・共通の不振日：${ctx.commonWeak.map(c => `${md(c.date)}(${c.wd}${c.dt === '日祝日' && c.wd !== '日' ? '・祝' : ''})${pct(c.rate)}`).join('、')}。対象店の大半が低調のため外部要因の確認を。`);
  return L;
}

/** 着地見込みのコメント（店舗・全体共通） */
function forecastLines(f, ctx) {
  if (!f) return [];
  const out = [];
  const rem = `残り${f.remDays}日（${f.remByType.map(x => `${x.key}${x.days}`).join('・')}）`;
  const how = f.mgrStores != null ? (f.mgrStores ? `${rem}は、店長の売上予測がある店（${f.mgrStores}／${f.nStores}店）はその予測、ない店は予算×ここまでの曜日区分別の予算比で見込んだ値` : `${rem}の予算を、ここまでの曜日区分別の予算比で見込んだ値`)
    : f.mgrDays >= f.remDays ? `${rem}は店長の売上予測で見込んだ値` : f.mgrDays ? `${rem}のうち${f.mgrDays}日は店長の売上予測、残りは予算×ここまでの曜日区分別の予算比で見込んだ値` : `${rem}の予算を、ここまでの曜日区分別の予算比で見込んだ値`;
  out.push({ t: `・着地見込み：${man(f.landing)}（月予算${man(f.monthBudget)}比${pct(f.landRate)}・${sman(f.landDiff)}）。${how}。`, d: false });
  if (f.need <= 0) out.push({ t: `　月予算はすでに達成済み（${sman(-f.need)}）。`, d: false });
  else {
    let t = `　予算達成には残り${f.remDays}日で${man(f.need)}、1日平均${man(f.needPerDay)}が必要（残り期間の予算比${pct(f.needRate)}。ここまでの予算比${pct(f.pace)}`;
    const gap = f.needRate - f.pace;
    if (gap >= 0.005) t += `から${pt(gap)}の上積みが必要`;
    else t += `を維持すれば届く水準`;
    out.push({ t: t + '）。', d: false });
  }
  if (f.estDays) out.push({ t: `　※売上集計表に${f.estDays}日分の予算がないため、その日の予算は同じ曜日区分の平均日予算で補っています（祝日は判別できません）。`, d: false });
  if (f.prevBased != null) out.push({ t: `　参考：前年の残り期間の実績×ここまでの前年比で見込むと${man(f.prevBased)}（予算比${pct(f.prevBasedRate)}）。`, d: true });
  return out;
}

/** 原価分析コメント（店舗） */
function costComment(s, ctx) {
  const L = [];
  const head = [];
  if (s.inv) head.push(`値引前原価率${pct2(s.inv.preRate)}`);
  if (s.costRate != null) head.push(`理論原価率${pct(s.costRate)}`);
  if (s.lossRate != null) head.push(`ロス率${pct2(s.lossRate)}`);
  if (s.inv) head.push(`ロス込原価率理論差${pt2(s.inv.gapExLoss)}`);
  else if (s.buyRate != null) head.push(`仕入率${pct(s.buyRate)}`);
  if (!head.length) return [{ t: '原価のデータがありません（売上集計表の「原価」「仕入高」「廃棄額」列とメニュー別データを読み込むと表示されます）。', d: false }];
  L.push({ t: `【原価】${head.join('／')}`, d: false });
  for (const x of costLines(s, ctx)) L.push(x);
  for (const x of lossLines(s, ctx)) L.push(x);
  for (const x of stockLines(s.stock, ctx)) L.push(x);
  for (const x of (s.inv ? invLines(s, ctx) : gapLines(s, ctx))) L.push(x);
  return L;
}
/** 当月の在庫の動き（月初在庫＝前月の棚卸）と、月末の在庫を月初と同じにするための残りの仕入の目安 */
function stockLines(st, ctx) {
  if (!st) return [];
  const d = ctx.lastDate ? `${ctx.lastDate.getMonth() + 1}/${ctx.lastDate.getDate()}` : '';
  const L = [{ t: `・在庫：月初在庫${man(st.open)}（前月の棚卸）＋当月の仕入${man(st.buy)}−理論の使用額${man(st.use)}−ロス${man(st.loss)}＝今の理論在庫${man(st.now)}${d ? `（${d}時点）` : ''}。`, d: false }];
  const rate = st.open ? st.delta / st.open : 0;
  if (st.delta > 0 && rate >= 0.1) L.push({ t: `　月初より${man(st.delta)}多い（${pct(rate)}増）。仕入が使用を上回っているので、発注量を確認（月末に在庫が多いまま締めると、その分は翌月の原価に回る）。`, d: false });
  else if (st.delta < 0 && rate <= -0.1) L.push({ t: `　月初より${man(-st.delta)}少ない（${pct(-rate)}減）。在庫を使っている状態。月末までに必要な分の仕入を確保。`, d: false });
  else L.push({ t: `　月初とほぼ同じ水準（${sman(st.delta)}）。`, d: true });
  if (st.remainBuy != null && st.remDays > 0) {
    if (st.remainBuy > 0) L.push({ t: `・残りの仕入の目安：月末の在庫を月初と同じ${man(st.open)}で締めるには、残り${st.remDays}日の仕入は${man(st.remainBuy)}まで（1日${man(st.remainBuy / st.remDays)}）。月の理論の使用額${man(st.monthUse)}＋ロス${man(st.monthLoss)}（着地見込みの売上から計算）。`, d: false });
    else L.push({ t: `・残りの仕入の目安：月の理論の使用額＋ロス（${man(st.monthUse + st.monthLoss)}）をすでに${man(-st.remainBuy)}上回る仕入。残り${st.remDays}日は在庫を使い、仕入を最小限に。`, d: false });
  }
  return L;
}
const shortItem = n => n.replace(/\s+/g, ' ').trim();
const unitTxt = (q, u) => `${(Math.round(q * 10) / 10).toLocaleString('ja-JP')}${u}`;
const pt2 = v => Math.abs(v) < 0.00005 ? '±0.00pt' : (v >= 0 ? '+' : '▲') + Math.abs(v * 100).toFixed(2) + 'pt';
const avgOf = a => a.length ? a.reduce((x, y) => x + y, 0) / a.length : null;
/** 原価管理表による原価率理論差（正式） */
function invLines(s, ctx) {
  const i = s.inv, A = ctx.area && ctx.area.inv;
  const out = [];
  const g3 = avgOf(i.trend.slice(-3).map(x => x.gapExLoss).filter(v => v != null));
  const g12 = avgOf(i.trend.slice(0, -1).map(x => x.gapExLoss).filter(v => v != null));
  const cmp = [g3 != null ? `直近3か月平均${pt2(g3)}` : '', ctx.nStores > 1 && A ? `対象店${pt2(A.gapExLoss)}` : ''].filter(Boolean).join('、');
  let t = `・原価率理論差：値引前原価率${pct2(i.preRate)}−理論原価率${pct2(i.theory)}＝${pt2(i.gap)}（${sman(i.gapAmt)}）。ロス率${pct2(i.lossPre)}を除いたロス込原価率理論差は${pt2(i.gapExLoss)}${cmp ? `（${cmp}）` : ''}`;
  if (i.gapExLoss <= -0.005) t += '。理論より使用が少ない（盛り付け不足・仕入や移動の計上漏れ・棚卸の過大計上などの可能性）';
  else if (i.gapExLoss >= 0.005) t += '。理論より使用が多い（盛り付け過多・未計上のロス・棚卸の過少計上などの可能性）';
  out.push({ t: t + '。', d: false });
  // 推移：前年同月から当月までの傾向
  if (i.trend.length >= 6 && g12 != null) {
    const neg = i.trend.filter(x => x.gapExLoss != null && x.gapExLoss <= -0.005).length, posN = i.trend.filter(x => x.gapExLoss != null && x.gapExLoss >= 0.005).length;
    let tr = `　推移：過去${i.trend.length - 1}か月平均${pt2(g12)}`;
    if (neg >= i.trend.length * 0.75) tr += `。${i.trend.length}か月中${neg}か月で▲0.5pt以下と、理論より少ない状態が続いており、単月の棚卸のぶれではなく恒常的な傾向`;
    else if (posN >= i.trend.length * 0.75) tr += `。${i.trend.length}か月中${posN}か月で+0.5pt以上と、理論より多い状態が続いている`;
    else if (Math.abs(i.gapExLoss - g12) >= 0.008) tr += `。今月は平時より${i.gapExLoss < g12 ? 'マイナス' : 'プラス'}に${Math.abs((i.gapExLoss - g12) * 100).toFixed(2)}pt振れている`;
    out.push({ t: tr + '。', d: false });
  }
  if (i.open && i.close) {
    const ch = (i.close - i.open) / i.open;
    out.push({ t: `　棚卸高${man(i.close)}（前月${man(i.open)}・${sman(i.close - i.open)}）${Math.abs(ch) >= 0.1 ? '。棚卸の増減が大きく、単月の理論差がぶれやすい月' : ''}。`, d: Math.abs(ch) < 0.1 });
  }
  if (A && A.catRate && ctx.nStores > 1 && i.preSales) {
    const diffs = i.cats.map(c => ({ cat: c.cat, r: c.cost / i.preSales, a: A.catRate[c.cat] })).filter(c => c.a != null && Math.abs(c.r - c.a) >= 0.003).sort((x, y) => (y.r - y.a) - (x.r - x.a));
    if (diffs.length) out.push({ t: `　大分類別の原価比（値引前売上比）で平均との差が大きい：${diffs.slice(0, 3).map(c => `${c.cat}${pct(c.r)}（平均${pct(c.a)}）`).join('、')}。`, d: true });
  }
  if ((i.lend || 0) !== 0 || (i.borrow || 0) !== 0) out.push({ t: `　店舗間の食材の貸し借り：貸出${man(Math.abs(i.lend || 0))}・借入${man(i.borrow || 0)}。`, d: true });
  return out;
}
/** 理論差（概算）：仕入率−理論原価率。ロスとそれ以外（在庫の増減・使用量の差など）に分ける */
function gapLines(s, ctx) {
  if (s.gapRate == null) return [];
  const cmp = [];
  const b = ctx.area && ctx.area.benchGapRate;
  if (ctx.nStores > 1 && b != null) cmp.push(`対象店平均${pt(b)}`);
  if (s.pGapRate != null) cmp.push(`前年${pt(s.pGapRate)}`);
  const lr = s.lossRate || 0, rest = s.gapRate - lr;
  return [{ t: `・理論差（概算）：仕入率${pct(s.buyRate)}−理論原価率${pct(s.costRate)}＝${pt(s.gapRate)}${cmp.length ? `（${cmp.join('、')}）` : ''}。うちロス${pt(lr)}、在庫の増減・使用量の差など${pt(rest)}。`, d: false }];
}
/** 原価分析コメント（全体） */
function areaCostComment(a, stores, ctx) {
  const L = [];
  const P = L.push.bind(L);
  L.push = t => P({ t, d: false });
  if (a.stock) { const up = stores.filter(x => x.stock && x.stock.open && x.stock.delta / x.stock.open >= 0.1).sort((x, y) => y.stock.delta - x.stock.delta); L.push(`・在庫（${a.stock.n}店・月初在庫＝前月の棚卸）：月初${man(a.stock.open)}→今の理論在庫${man(a.stock.now)}（${sman(a.stock.delta)}）。${up.length ? `月初より10%以上増えている店：${up.slice(0, 3).map(x => `${x.name}${sman(x.stock.delta)}`).join('、')}。` : ''}${a.stock.remainBuy != null && a.stock.remDays > 0 ? `月末を月初の在庫で締める残りの仕入の目安は${man(Math.max(0, a.stock.remainBuy))}（残り${a.stock.remDays}日）。` : ''}`); }
  // ブロック長向け：営業部ごとの原価指標
  if (ctx.multiArea) {
    const by = {};
    stores.forEach(x => { const k = x.area || '（営業部なし）'; (by[k] = by[k] || []).push(x); });
    const parts = Object.entries(by).map(([k, xs]) => {
      const pre = sum(xs, x => x.preSales), f = [];
      if (xs.some(x => x.costRate != null)) f.push(`理論原価率${pct(sum(xs, x => x.cost) / pre)}`);
      if (xs.some(x => x.lossRate != null)) f.push(`ロス率${pct2(sum(xs, x => x.loss) / pre)}`);
      const wi = xs.filter(x => x.inv);
      if (wi.length) { const ip = sum(wi, x => x.inv.preSales); f.push(`ロス込理論差${pt2(sum(wi, x => x.inv.cost - x.inv.theory * x.inv.preSales - x.inv.loss) / ip)}${wi.length < xs.length ? `（原価管理表${wi.length}店）` : ''}`); }
      const wm = xs.filter(x => x.menu);
      if (wm.length) f.push(`フード${(sum(wm, x => x.menu.foodQty * x.menu.cust) / sum(wm, x => x.menu.cust)).toFixed(1)}皿/人`);
      return f.length ? `${k}（${xs.length}店）${f.join('・')}` : null;
    }).filter(Boolean);
    if (parts.length) L.push(`・営業部別：${parts.join('、')}。`);
  }
  if (a.lossRate != null) {
    const hi = stores.filter(x => x.lossRate != null).sort((x, y) => y.lossRate - x.lossRate).slice(0, 3);
    const wd = [...a.lossByWd].sort((x, y) => y.loss - x.loss).slice(0, 2);
    L.push(`・ロス：ロス額${man(a.loss)}・ロス率${pct2(a.lossRate)}${a.pLossRate != null ? `（前年${pct2(a.pLossRate)}）` : ''}。${stores.length > 1 ? `ロス率の高い店：${hi.map(x => `${x.name}${pct2(x.lossRate)}`).join('、')}。` : ''}曜日別の上位は${wd.map(x => `${x.key}曜${Math.round(x.loss / a.loss * 100)}%`).join('、')}。`);
  }
  if (a.menu) {
    const m = a.menu, P = a.pMenuAgg;
    const yoyTxt = a.pCostRate != null && a.pCust ? `（前年${pct(a.pCostRate)}・${pt(a.costRate - a.pCostRate)}。1人あたり原価${syen(a.cost / a.cust - a.pCost / a.pCust)}、値引前客単価${syen(a.preSales / a.cust - a.pPreSales / a.pCust)}）` : '';
    L.push(`・原価：理論原価率${pct(a.costRate)}${yoyTxt}。1人あたり食べ放題フード${m.foodQty.toFixed(1)}皿・${Math.round(m.foodG).toLocaleString('ja-JP')}g${P ? `（前年${P.foodQty.toFixed(1)}皿・${Math.round(P.foodG).toLocaleString('ja-JP')}g）` : ''}、精肉${m.meatQty.toFixed(2)}皿、原価${yen(m.costPP)}。`);
    const hi = stores.filter(x => x.costRate != null).sort((x, y) => y.costRate - x.costRate).slice(0, 3);
    if (stores.length > 1) L.push(`　理論原価率の高い店：${hi.map(x => `${x.name}${pct(x.costRate)}${x.menu ? `（${x.menu.foodQty.toFixed(1)}皿/人）` : ''}`).join('、')}。`);
  }
  if (a.inv) {
    const v = stores.filter(x => x.inv).sort((x, y) => x.inv.gapExLoss - y.inv.gapExLoss);
    L.push(`・原価率理論差（原価管理表）：値引前原価率${pct2(a.inv.preRate)}−理論原価率${pct2(a.inv.theory)}＝${pt2(a.inv.gap)}、ロス込${pt2(a.inv.gapExLoss)}。${v.length > 1 ? `ロス込理論差のマイナスが大きい店：${v.slice(0, 3).map(x => `${x.name}${pt2(x.inv.gapExLoss)}`).join('、')}／プラス寄りの店：${v.slice(-2).reverse().map(x => `${x.name}${pt2(x.inv.gapExLoss)}`).join('、')}。` : ''}`);
    const chronic = v.filter(x => x.inv.trend.length >= 6 && x.inv.trend.filter(t => t.gapExLoss != null && t.gapExLoss <= -0.005).length >= x.inv.trend.length * 0.75);
    if (chronic.length) L.push(`　1年を通して理論より少ない状態が続く店：${chronic.map(x => x.name).join('、')}（理論登録・盛り付け基準・棚卸方法の確認を）。`);
  } else if (a.gapRate != null) {
    const g = stores.filter(x => x.gapRate != null).sort((x, y) => y.gapRate - x.gapRate);
    L.push(`・理論差（概算＝仕入率−理論原価率）：${pt(a.gapRate)}${a.pGapRate != null ? `（前年${pt(a.pGapRate)}）` : ''}。${stores.length > 1 ? `差の大きい店：${g.slice(0, 3).map(x => `${x.name}${pt(x.gapRate)}（うちロス${pt(x.lossRate || 0)}）`).join('、')}。` : ''}`);
  }
  if (a.menu && a.menu.courses.length >= 2 && stores.length > 1) {
    const prem = stores.filter(x => x.menu).map(x => ({ x, p: sum(x.menu.courses.filter(c => /プレミアム/.test(c.name)), c => c.share) })).sort((u, v) => v.p - u.p);
    const drk = stores.filter(x => x.menu && x.menu.drink).sort((u, v) => v.menu.drink.alcRate - u.menu.drink.alcRate);
    if (a.menu.drink && drk.length > 1) L.push(`・飲み放題比率：ソフトドリンク${pct(a.menu.drink.softRate)}${a.menu.drink.lunchRate != null ? `・ランチソフト${pct(a.menu.drink.lunchRate)}` : ''}・アルコール${pct(a.menu.drink.alcRate)}（対象店計）。アルコールの高い店 ${drk.slice(0, 2).map(x => `${x.name}${pct(x.menu.drink.alcRate)}`).join('、')}／低い店 ${drk.slice(-2).reverse().map(x => `${x.name}${pct(x.menu.drink.alcRate)}`).join('、')}。`);
    if (prem.length && prem[0].p > 0) L.push(`・コース構成：プレミアム比率の高い店 ${prem.slice(0, 3).map(u => `${u.x.name}${pct(u.p)}`).join('、')}／低い店 ${prem.slice(-2).reverse().map(u => `${u.x.name}${pct(u.p)}`).join('、')}。`);
  }
  if (!L.length) L.push('原価のデータがありません。');
  return L;
}

/* ===== 前月の課題のその後（連続する2か月を比べる） ===== */
/** s0：前の月の店舗（または全体）、s1：次の月。前の月の課題ごとに、次の月の値と状況を返す */
function followUp(s0, s1) {
  if (!s0 || !s1) return [];
  const out = [];
  const add = (label, r0, r1, lowerBetter, unit) => { if (r0 == null || r1 == null || !isFinite(r0) || !isFinite(r1)) return;
    const d = Math.round((r1 - r0) * (lowerBetter ? -1 : 1) * (unit === 'loss' ? 10000 : 1000)) / (unit === 'loss' ? 10000 : 1000); // 表示の桁で丸めてから判定
    if (unit === 'loss' && Math.abs(d) < 0.0005) return; // ロス率は変化があるときだけ
    const th = unit === 'loss' ? 0.0005 : 0.02;
    const st = (unit === 'rate' || unit === 'yoy') && !lowerBetter && r1 >= 1 ? 'solved' : d >= th ? 'better' : d <= -th ? 'worse' : 'flat';
    out.push({ label, r0, r1, unit, lowerBetter, st }); };
  add('予算比', s0.rate, s1.rate, false, 'rate');
  // 未達だった曜日区分（未達額の大きい順に2つ）
  s0.byType.filter(t => t.budget > 0 && t.diff < 0).sort((a, b) => a.diff - b.diff).slice(0, 2).forEach(t => { const u = s1.byType.find(x => x.key === t.key); if (u && u.budget > 0) add(t.key, t.rate, u.rate, false, 'rate'); });
  if (s0.hasPrev && s1.hasPrev) {
    // 時間帯・曜日区分×時間帯は前年比で（時間帯の予算は按分のため）
    s0.byBand.filter(b => b.pSales > 0 && b.pDiff < 0).sort((a, b) => a.pDiff - b.pDiff).slice(0, 2).forEach(b => { const u = s1.byBand.find(x => x.key === b.key); if (u && u.yoy != null) add(`${b.name}（前年比）`, b.yoy, u.yoy, false, 'yoy'); });
    let w = null; s0.grid.forEach((row, ti) => row.forEach((c, bi) => { if (c.dayDiff == null || c.yoyDay == null) return; const imp = c.dayDiff * c.d; if (imp < 0 && (!w || imp < w.imp)) w = { imp, ti, bi }; }));
    if (w) { const c0 = s0.grid[w.ti][w.bi], c1 = s1.grid[w.ti][w.bi]; if (c1 && c1.yoyDay != null) add(`${DAYTYPES[w.ti]}×${BANDS[w.bi].name}（前年比）`, c0.yoyDay, c1.yoyDay, false, 'yoy'); }
  } else {
    s0.byBand.filter(b => b.budget > s0.budget * 0.02 && b.diff < 0).sort((a, b) => a.diff - b.diff).slice(0, 2).forEach(b => { const u = s1.byBand.find(x => x.key === b.key); if (u && u.budget > 0) add(b.name, b.rate, u.rate, false, 'rate'); });
    let w = null; s0.grid.forEach((row, ti) => row.forEach((c, bi) => { const d = c.sales - c.budget; if (c.budget > 0 && d < 0 && (!w || d < w.d)) w = { d, ti, bi }; }));
    if (w) { const c0 = s0.grid[w.ti][w.bi], c1 = s1.grid[w.ti][w.bi]; if (c1 && c1.budget > 0) add(`${DAYTYPES[w.ti]}×${BANDS[w.bi].name}`, c0.sales / c0.budget, c1.sales / c1.budget, false, 'rate'); }
  }
  if (s0.custRate != null && s0.custRate < 1 && s0.custEff < 0) add('客数予算比', s0.custRate, s1.custRate, false, 'rate');
  if (s0.lossRate != null && s1.lossRate != null) add('ロス率', s0.lossRate, s1.lossRate, true, 'loss');
  // 同じ名前が重なったら最初のものだけ
  return out.filter((x, i) => out.findIndex(y => y.label === x.label) === i);
}
const FU_TXT = { solved: '解消（予算達成・前年超え）', better: '改善中', flat: '横ばい（課題が続いている）', worse: '悪化' };
/** 1行のコメント（状況ごとにまとめる）。L0・L1：月の呼び名（例「10月」「11月途中」） */
function followUpLine(items, L0, L1, viewingLater) {
  if (!items.length) return null;
  const f = (x, v) => x.unit === 'loss' ? pct2(v) : pct(v);
  const head = viewingLater ? `・${L0}の課題のその後（${L0}→${L1}）` : `・${L0}の課題の${L1}の状況（${L0}→${L1}）`;
  const parts = ['worse', 'flat', 'better', 'solved'].map(st => { const xs = items.filter(x => x.st === st); return xs.length ? `${FU_TXT[st]}：${xs.map(x => `${x.label} ${f(x, x.r0)}→${f(x, x.r1)}`).join('、')}` : null; }).filter(Boolean);
  return { t: `${head}：${parts.join('／')}。`, d: false };
}

/** FC企業のコメント：会社合計 → 着地 → 店舗ごと → 客数・単価 → 原価・ロス */
function companyLines(c, isFullMonth) {
  const L = [], n = c.stores.length;
  L.push({ t: `【${c.diff >= 0 ? '予算達成' : '予算未達'}】${c.name}（FC${n}店）売上${man(c.sales)}／予算比${pct(c.rate)}（${sman(c.diff)}）${c.yoy ? `・前年比${pct(c.yoy)}` : ''}。予算達成${c.okN}店／${n}店。`, d: false });
  if (c.fc) L.push({ t: `・着地見込み：${man(c.fc.landing)}（月予算比${pct(c.fc.landRate)}・${sman(c.fc.landDiff)}）。${c.fc.need > 0 ? `達成には残り${c.fc.remDays}日で1日${man(c.fc.needPerDay)}（1店あたり${man(c.fc.needPerDay / n)}）が必要。` : '月予算は達成済み。'}`, d: false });
  if (n > 1) {
    const ng = c.stores.filter(s => s.diff < 0).sort((a, b) => a.diff - b.diff), ok = c.stores.filter(s => s.diff >= 0).sort((a, b) => b.diff - a.diff);
    const nm = s => s.name.replace(/^焼肉きんぐ/, '');
    if (ng.length) L.push({ t: `・店舗別の未達：${ng.slice(0, 4).map(s => `${nm(s)} ${sman(s.diff)}（${pct(s.rate)}）`).join('、')}${ng.length > 4 ? ` ほか${ng.length - 4}店` : ''}。`, d: false });
    if (ok.length) L.push({ t: `・店舗別の達成：${ok.slice(0, 4).map(s => `${nm(s)} ${sman(s.diff)}（${pct(s.rate)}）`).join('、')}${ok.length > 4 ? ` ほか${ok.length - 4}店` : ''}。`, d: false });
    const yo = c.stores.filter(s => s.yoy != null);
    if (yo.length > 1) { const lo = [...yo].sort((a, b) => a.yoy - b.yoy)[0]; if (lo.yoy < 1) L.push({ t: `・店舗別の前年比：いちばん低いのは${nm(lo)}（前年比${pct(lo.yoy)}）。`, d: true }); }
  }
  L.push({ t: `・要因分解：客数要因${sman(c.custEff)}（客数予算比${pct(c.custRate)}${c.custYoy ? `・前年比${pct(c.custYoy)}` : ''}）、単価要因${sman(c.upEff)}（客単価${Math.round(c.up).toLocaleString('ja-JP')}円）。`, d: false });
  if (c.inv) L.push({ t: `・原価（原価管理表${c.inv.n}店）：原価率${pct(c.inv.preRate)}、理論原価率${pct(c.inv.theory)}、理論差${pt(c.inv.gap)}。`, d: false });
  else if (c.gapRate != null) L.push({ t: `・原価（概算）：仕入率${pct(c.buyRate)}、理論原価率${pct(c.costRate)}、理論差${pt(c.gapRate)}。`, d: false });
  if (c.lossRate != null) L.push({ t: `・ロス：廃棄率${pct(c.lossRate)}。`, d: true });
  return L;
}
function commentText(lines, detail) { return lines.filter(l => detail || !l.d).map(l => l.t).join('\n'); }

if (typeof module !== 'undefined') module.exports = { perfReport, extractMenuDaily, qscLines, npsYm, npsSnapshot, npsSeries, isNpsBook, extractNps, NPS_QSC, isMasterSheet, extractMaster, companyLines, followUp, followUpLine, tcSummary, wageSnapshot, mergeHistory, analyzeTenure, tenureLines, combineStd, isStdBook, extractStd, stdHours, analyzeStd, STD_DEFAULT, extractTimecard, extractSchedule, extractWage, analyzeLabor, forecastStore, forecastSum, pt2, extractInventory, isInventorySheet, extractMenu, menuAgg, scopeCounts, listAreas, sheetKind, extractHourly, extractPayments, commentText, SHEET_CUR, SHEET_PREV, findSheet, extractRows, analyze, DAYTYPES, BANDS };
