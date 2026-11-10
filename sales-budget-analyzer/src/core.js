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
  const need = { code: '店舗コード', name: '店舗名', date: '営業日', wd: '曜日', dt: '区分', weather: '天気', sales: '純売上高(合計)', cust: '客数(合計)', groups: '組数(合計)', budget: '売上月次予算', custBudget: '客数月次予算', disc: '税抜値引(合計)', fc: 'FC区分', area: '営業部', cost: '原価', loss: '廃棄額', buy: '仕入高' };
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
      fc: C.fc >= 0 ? String(r[C.fc] || '') : '',
      area: C.area >= 0 ? String(r[C.area] || '').trim() : '',
      cost: C.cost >= 0 ? num(r[C.cost]) : 0,
      loss: C.loss >= 0 ? num(r[C.loss]) : 0,
      buy: C.buy >= 0 ? num(r[C.buy]) : 0,
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
  return {
    cust, cost: sum(rows, r => r.qty * r.unitCost), costPP: pp(r => r.qty * r.unitCost), menuSales: sum(rows, r => r.sales),
    foodQty: sum(food, r => r.qty) / cust, foodG: sum(food, r => r.qty * r.g) / cust, foodCost: sum(food, r => r.qty * r.unitCost) / cust,
    meatQty: sum(meat, r => r.qty) / cust, meatG: sum(meat, r => r.qty * r.g) / cust, meatCost: sum(meat, r => r.qty * r.unitCost) / cust,
    drinkQty: sum(drink, r => r.qty) / cust,
    cats: Object.values(cats).sort((a, b) => b.cost - a.cost), items,
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
  const byBand = BANDS.map((bd, b) => {
    const s = sum(grid, r => r[b].sales), bu = sum(grid, r => r[b].budget), c = sum(grid, r => r[b].cust), p = sum(grid, r => r[b].pSales);
    return { key: bd.key, name: bd.name, label: bd.label, sales: s, budget: bu, diff: s - bu, rate: bu ? s / bu : null, cust: c, pSales: p, yoy: p ? s / p : null };
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
    code: cur[0].code, name: cur[0].name, area: cur[0].area || '',
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
      planMins: sum(fut, d => d.mins), planCost: sum(fut, d => d.cost), estDays: fut.filter(d => d.est).length,
      avgWage: avgW, noWage: noWage.size, wageless: !avgW,
    };
    s.landMins = s.mins + s.planMins; s.landCost = s.cost + s.planCost;
    s.alerts = list.filter(p => p.totalOt >= OT_WARN).sort((a, b) => b.totalOt - a.totalOt);
    return s;
  }).filter(s => s.landMins > 0).sort((a, b) => a.code.localeCompare(b.code));
  const T = k => sum(stores, s => s[k]);
  const wd = stores.filter(s => !s.wageless), TW = k => sum(wd, s => s[k]);
  const area = { mins: T('mins'), staffMins: T('staffMins'), cost: TW('cost'), planMins: T('planMins'), planCost: TW('planCost'), landMins: T('landMins'), landCost: TW('landCost'), wageless: stores.length - wd.length, wagelessNames: stores.filter(s => s.wageless).map(s => s.name),
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
  push(`【労働時間】実績${hrs0(s.mins)}（${cutTxt}。うち社員${hrs0(s.staffMins)}）${land ? `、着地見込み${hrs0(s.landMins)}（残り${L.N - L.cut}日の勤務予定${hrs0(s.planMins)}）` : ''}。`);
  if (s.estDays) push(`　※残り${s.estDays}日は勤務予定が未作成（予定が同じ曜日区分の実績平均の半分未満）のため、実績平均で補っています。`);
  if (L.hasWage && s.wageless === true) push('・雑給：時給データにこの店舗のパートがいないため計算できません（時給データの店舗を確認してください）。');
  else if (L.hasWage) {
    const landSales = f ? f.landing : null;
    let t = `・雑給（パート）：実績${man(s.cost)}${land ? `、着地見込み${man(s.landCost)}` : ''}`;
    if (sales) t += `。雑給率${pct(s.cost / sales)}（実績）${landSales ? `／${pct(s.landCost / landSales)}（売上・雑給とも着地見込み）` : ''}`;
    push(t + `。割増は残業25%・深夜25%（重なる時間は50%）で計算${s.wageless ? `（時給データのない${s.wagelessNames.join('・')}を除く）` : ''}。`);
    if (s.noWage) push(`　時給データにない${s.noWage}名（新規入社など）は店舗のパート平均時給${s.avgWage ? yen(s.avgWage) : ''}で計算しています。`, true);
  }
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
// 店舗設定の初期値（見込君の初期入力と同じ）
const STD_DEFAULT = { lunchWd: false, lunchHd: true, auto: false, t8: 3, hoursWd: 6, hoursHd: 12 };
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
function analyzeStd(s, lab, L, std, settings, salesLast) {
  if (!lab || !std) return null;
  const info = std.info[s.code] || {}, set = settings && settings[s.code];
  const lane = !!info.lane;
  const st = { ...STD_DEFAULT, ...(set || {}) }, tables = (set && set.tables) || info.tables;
  if (lane) st.perCust = info.perCust;
  if (!tables) return { missing: true };
  if (lane ? !(std.laneTables && std.laneTables.length) : !(std.tables && std.tables.length)) return { missing: true, need: lane ? 'lane' : 'normal' };
  const hol = new Set(std.holidays);
  const typeOf = date => hol.has(+date) || date.getDay() === 0 || date.getDay() === 6 ? 2 : 1;
  const actCust = {}; s.daily.forEach(d => { actCust[d.date.getDate()] = d.cust; });
  const fcCust = {}; if (s.fc) s.fc.days.forEach(d => { fcCust[d.date.getDate()] = d.custFc; });
  const days = lab.days.map((d, i) => {
    const n = i + 1, type = typeOf(d.date);
    const useAct = n <= salesLast && actCust[n] != null;
    const cust = useAct ? actCust[n] : fcCust[n] != null ? fcCust[n] : null;
    const sh = cust != null ? stdHours(cust, type, st, tables, std, lane) : null;
    return { date: d.date, type, cust, custEst: !useAct, hours: d.mins / 60, done: d.done, est: d.est, std: sh };
  });
  const ok = days.filter(d => d.std != null);
  const past = ok.filter(d => d.done), all = ok;
  const H = ds => sum(ds, d => d.hours), SH = ds => sum(ds, d => d.std), C = ds => sum(ds, d => d.cust);
  const mo = days[0].date.getMonth() + 1, br = (std.budget[s.code] || {})[mo] ?? null;
  const landStd = SH(all), landH = H(all);
  const r = {
    settings: st, custom: !!set, tables, lane, budgetRate: br, days,
    cutStd: SH(past), cutHours: H(past), cutCust: C(past), cutRate: SH(past) ? H(past) / SH(past) : null,
    landStd, landHours: landH, landCust: C(all), landRate: landStd ? landH / landStd : null,
    remDays: days.filter(d => !d.done).length, complete: ok.length === days.length,
  };
  // 予算標準比との差：目標時間＝標準時間×予算標準比
  if (br != null) { r.target = landStd * br; r.over = landH - r.target; r.overCost = r.over > 0 && lab.avgWage ? r.over * lab.avgWage : null; }
  // 週別（月曜始まり）
  const wk = {}; ok.forEach(d => { const w = Math.floor((d.date.getDate() - 1 + ((new Date(d.date.getFullYear(), d.date.getMonth(), 1).getDay() + 6) % 7)) / 7); (wk[w] = wk[w] || []).push(d); });
  r.weeks = Object.keys(wk).sort((a, b) => a - b).map(w => { const ds = wk[w]; return { from: ds[0].date, to: ds[ds.length - 1].date, n: ds.length, cust: C(ds), std: SH(ds), hours: H(ds), rate: SH(ds) ? H(ds) / SH(ds) : null, done: ds.every(d => d.done), mixed: ds.some(d => d.done) && ds.some(d => !d.done) }; });
  return r;
}
const p0 = v => v == null ? '－' : (v * 100).toFixed(1) + '%';
function stdLines(x, ctx) {
  if (!x) return [];
  if (x.missing) return [{ t: x.need === 'lane' ? '・標準比：特急レーン店のため、レーン店用の標準比マスタを読み込むと計算します。' : x.need === 'normal' ? '・標準比：通常店用の標準比マスタを読み込むと計算します。' : '・標準比：卓数がわからないため計算できません（標準比マスタに店舗がないか、店舗設定で卓数を入れてください）。', d: false }];
  const out = [];
  let t = `・標準比${x.lane ? '（特急レーン店）' : ''}：実績${p0(x.cutRate)}（${ctx.cutTxt}、人時来客${x.cutHours ? (x.cutCust / x.cutHours).toFixed(2) : '－'}）`;
  if (x.remDays) t += `、着地見込み${p0(x.landRate)}（残り${x.remDays}日は客数見込み×勤務予定）`;
  if (x.budgetRate != null) t += `。予算標準比${p0(x.budgetRate)}`;
  out.push({ t: t + '。', d: false });
  if (x.budgetRate != null && x.remDays) {
    if (x.over > 0) out.push({ t: `　予算標準比に収めるには、残り${x.remDays}日で${hrs0(x.over * 60)}（1日${(x.over / x.remDays).toFixed(1)}h）の削減が必要${x.overCost ? `（雑給換算で約${man(x.overCost)}）` : ''}。`, d: false });
    else out.push({ t: `　勤務予定どおりなら予算標準比内（余裕${hrs0(-x.over * 60)}）。`, d: false });
  }
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
  const out = { tc: { ...((H || {}).tc || {}) }, wage: { ...((H || {}).wage || {}) } };
  const key = x => x.ym + (x.area ? '|' + x.area : '');
  ((add || {}).tc || []).forEach(x => { if (x && x.ym) out.tc[key(x)] = x; });
  ((add || {}).wage || []).forEach(x => { if (x && x.ym) out.wage[key(x)] = x; });
  return out;
}
/** 年月ごとにまとめ直す（エリアごとの分をつなげる） */
function byMonth(obj) {
  const out = {};
  Object.values(obj || {}).forEach(x => { const o = (out[x.ym] = out[x.ym] || { ym: x.ym, recs: [], seen: new Set() }); x.recs.forEach(r => { const k = r.e + '@' + r.c; if (!o.seen.has(k)) { o.seen.add(k); o.recs.push(r); } }); });
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
  useMonths.forEach(m => H.tc[m].recs.forEach(r => {
    if (r.s || (pick && !pick.has(r.c))) return;
    const p = (P[r.e] = P[r.e] || { emp: r.e, name: r.n, code: r.c, hrs: {}, days: {} });
    p.hrs[m] = (p.hrs[m] || 0) + r.m / 60; p.days[m] = (p.days[m] || 0) + r.d; p.code = r.c; p.name = r.n;
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
      p.hrs[curYm] = monthEq; p.code = st.code;
      curRate[x.emp] = { act: cur.cut ? x.mins / 60 / cur.cut : 0, plan: covered ? x.planMins / 60 / covered : null, actDays: x.days, planDays: x.planDays, covered };
    });
  });
  const tenureK = (emp, ym) => { const i = info[emp]; if (!i || !i.h) return null; const h = new Date(i.h); return ymIdx(ym) - (h.getFullYear() * 12 + h.getMonth()); };
  const band = k => k == null ? null : TENURE.findIndex(b => b.test(k));
  // 当月の勤務期間別
  const people = Object.values(P).map(p => ({ ...p, k: tenureK(p.emp, curYm), hire: info[p.emp] && info[p.emp].h ? info[p.emp].h : null, wage: info[p.emp] ? info[p.emp].w : null, cur: p.hrs[curYm] || 0 }));
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
    if (prevAvg) c.push({ ratio: p.cur / prevAvg, why: `当月見込み${p.cur.toFixed(0)}h（前${prevMs.length}か月平均${prevAvg.toFixed(0)}h）` });
    if (r && r.plan != null && r.act > 0 && r.actDays >= 3 && r.covered >= 5) c.push({ ratio: r.plan / r.act, why: `今後の予定が1日${r.plan.toFixed(1)}h（これまでの実績は1日${r.act.toFixed(1)}h）` });
    const best = c.sort((a, b) => a.ratio - b.ratio)[0];
    const ratio = best ? best.ratio : null, why = best ? best.why : '';
    if (ratio != null && ratio <= RISK_WARN) risk.push({ ...p, ratio, why, level: ratio <= RISK_HIGH ? 'high' : 'warn' });
  });
  risk.sort((a, b) => a.ratio - b.ratio);
  // 店舗別
  const byStore = {};
  working.forEach(p => { (byStore[p.code] = byStore[p.code] || []).push(p); });
  const stores = Object.entries(byStore).map(([code, xs]) => {
    const b = bandStats(xs), t = sum(xs, p => p.cur); b.forEach(x => { x.share = t ? x.hours / t : null; });
    return { code, n: xs.length, bands: b, hires: hires.filter(h => h.code === code), leavers: leavers.filter(l => l.code === code), risk: risk.filter(r => r.code === code) };
  });
  const order = p => (risk.find(r => r.emp === p.emp) ? 0 : hires.some(h => h.emp === p.emp) ? 1 : 2);
  const peopleOf = code => people.filter(p => p.code === code && Object.values(p.hrs).some(h => h > 0)).sort((a, b) => order(a) - order(b) || (a.k ?? 999) - (b.k ?? 999));
  return { peopleOf, curYm, months: useMonths, wMonths, hasCur: !!Object.keys(curRate).length, bands, total, ramp, hires, leavers, risk, stores, preLeave: pre.length >= 3 ? { n: pre.length, ratio: avgOf(pre) } : null, n: working.length, noHire: working.filter(p => p.k == null).length };
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
  if (b(4).n && b(4).n >= n * 0.2) push(`　※${b(4).n}名は時給データに入社日がありません（社員番号が見つからない、または別の店の所属など）。`, true);
  const hires = st ? st.hires : T.hires, leavers = st ? st.leavers : T.leavers, risk = st ? st.risk : T.risk;
  const nm = p => `${isArea && p.store ? p.store + '・' : ''}${p.name}`;
  if (hires.length) push(`・新規採用（当月・前月入社）：${hires.length}名（${hires.slice(0, 6).map(p => `${nm(p)}${hh(p.cur)}`).join('、')}${hires.length > 6 ? ' ほか' : ''}）。`);
  if (leavers.length) push(`・退社：${leavers.filter(l => l.sure).length}名${leavers.some(l => !l.sure) ? `、退社の可能性${leavers.filter(l => !l.sure).length}名（2か月勤務なし）` : ''}（${leavers.slice(0, 6).map(l => `${nm(l)}${l.k != null ? `・入社${l.k}か月` : ''}`).join('、')}${leavers.length > 6 ? ' ほか' : ''}）。`);
  if (risk.length) push(`・退社リスク注意（入社1年以内で時間が減っている）：${risk.slice(0, 6).map(p => `${nm(p)}（入社${p.k}か月、${p.why}）`).join('、')}${risk.length > 6 ? ` ほか${risk.length - 6}名` : ''}。`);
  else push('・退社リスク注意：入社1年以内で時間が大きく減っている人はいません。');
  if (T.preLeave && !st) push(`　参考：これまでに退社した${T.preLeave.n}名は、最後の月の時間がそれまでの平均の${Math.round(T.preLeave.ratio * 100)}%でした。`, true);
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
  rows.filter(r => r.date > lastDate).forEach(r => { const k = +r.date; (fut[k] = fut[k] || { budget: 0, custBudget: 0, dt: r.dt }); fut[k].budget += r.budget; fut[k].custBudget += r.custBudget || 0; });
  const days = [];
  for (let d = lastDate.getDate() + 1; d <= monthEnd.getDate(); d++) {
    const date = new Date(lastDate.getFullYear(), lastDate.getMonth(), d);
    const f = fut[+date];
    const dt = f ? f.dt : dtOfDate(date);
    const budget = f ? f.budget : avgBudget[dt];
    const cb = f && f.custBudget ? f.custBudget : avgCustB[dt];
    // 客数の見込み：客数予算×ここまでの曜日区分別の客数予算比。客数予算がなければ売上見込み÷ここまでの客単価
    const custFc = cb && custRate[dt] != null ? cb * custRate[dt] : s.up ? budget * typeRate[dt] / s.up : 0;
    days.push({ date, dt, budget, est: !f, fc: budget * typeRate[dt], custFc });
  }
  const remBudget = sum(days, d => d.budget), remFc = sum(days, d => d.fc);
  const monthBudget = s.budget + remBudget, landing = s.sales + remFc, need = monthBudget - s.sales;
  // 参考：前年の残り期間の実績 × ここまでの前年比
  const pRem = prevRows.filter(r => r.date.getDate() > lastDate.getDate() && r.date.getDate() <= monthEnd.getDate());
  const prevBased = s.yoy && pRem.length ? s.sales + sum(pRem, r => r.sales) * s.yoy : null;
  return {
    remDays: n, days, estDays: days.filter(d => d.est).length,
    remByType: DAYTYPES.map(dt => ({ key: dt, days: days.filter(d => d.dt === dt).length })).filter(x => x.days),
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
  const days = f0.days.map((d, i) => ({ date: d.date, dt: d.dt, est: fs.some(f => f.days[i] && f.days[i].est), budget: sum(fs, f => f.days[i] ? f.days[i].budget : 0), fc: sum(fs, f => f.days[i] ? f.days[i].fc : 0), custFc: sum(fs, f => f.days[i] ? f.days[i].custFc : 0) }));
  const pb = fs.every(f => f.prevBased != null) ? sum(fs, f => f.prevBased) : null;
  return {
    remDays: f0.remDays, days, estDays: Math.max(...fs.map(f => f.estDays)), remByType: f0.remByType,
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
  }
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
  const withInv = stores.filter(s => s.inv);
  if (withInv.length) {
    const T = (f) => sum(withInv, s => f(s.inv));
    const pre = T(i => i.preSales), cost = T(i => i.cost), th = T(i => i.theory * i.preSales), loss = T(i => i.loss);
    area.inv = { preSales: pre, cost, preRate: cost / pre, theory: th / pre, gap: (cost - th) / pre, gapAmt: cost - th, lossPre: loss / pre, gapExLoss: (cost - th - loss) / pre, open: T(i => i.open), close: T(i => i.close), n: withInv.length };
    const catSum = {}; withInv.forEach(s => s.inv.cats.forEach(c => { catSum[c.cat] = (catSum[c.cat] || 0) + c.cost; }));
    area.inv.catRate = Object.fromEntries(Object.entries(catSum).map(([k, v]) => [k, v / pre]));
  }
  // 理論差（概算）＝仕入率−理論原価率
  stores.concat([area]).forEach(s => { s.gapRate = s.buyRate != null && s.costRate != null ? s.buyRate - s.costRate : null; s.pGapRate = s.pBuyRate != null && s.pCostRate != null ? s.pBuyRate - s.pCostRate : null; });
  const withGap = stores.filter(s => s.gapRate != null);
  area.benchGapRate = withGap.length ? (sum(withGap, s => s.buy) - sum(withGap, s => s.cost)) / sum(withGap, s => s.preSales) : null;
  const withLoss = stores.filter(s => s.lossRate != null);
  area.benchLossRate = withLoss.length ? sum(withLoss, s => s.loss) / sum(withLoss, s => s.preSales) : null;
  if (menu.length) {
    stores.forEach(s => { s.menu = menuAgg(menu.filter(r => r.code === s.code)); s.pMenuAgg = menuAgg(pMenu.filter(r => r.code === s.code)); });
    area.menu = menuAgg(menu); area.pMenuAgg = menuAgg(pMenu);
    const withMenu = stores.filter(s => s.menu && s.costRate != null);
    area.benchCostRate = withMenu.length ? sum(withMenu, s => s.cost) / sum(withMenu, s => s.preSales) : null;
    area.benchUp = withMenu.length ? sum(withMenu, s => s.preSales) / sum(withMenu, s => s.menu.cust) : null;
  }
  const ctx = { isFullMonth, lastDate, commonWeak, areaLabor, scopeLabel, multiArea, area, nStores: stores.length };
  stores.forEach(s => { s.lines = storeComment(s, ctx); s.costLines = costComment(s, ctx); });
  area.hasInv = withInv.length > 0;
  area.lines = areaComment(area, stores, ctx);
  if (labor) {
    // 売上の実績期間と勤怠の確定日が同じときだけ、実績の雑給率を出す
    const sameCut = labor.cut === lastDate.getDate() && labor.mo === lastDate.getMonth();
    const std = extra.std || null, cutTxt = `${labor.mo + 1}/${labor.cut}時点`;
    stores.forEach(s => {
      s.labor = labor.stores.find(x => x.code === s.code) || null;
      s.std = std && s.labor ? analyzeStd(s, s.labor, labor, std, extra.stdSettings, lastDate.getMonth() === labor.mo ? lastDate.getDate() : 0) : null;
      s.laborLines = s.labor ? laborLines(s.labor, labor, s.fc, sameCut ? s.sales : null, ctx).concat(stdLines(s.std, { cutTxt })) : [{ t: '勤怠データにこの店舗がありません。', d: false }];
    });
    const a = labor.area, fs = stores.filter(s => s.labor);
    area.labor = a;
    const full = fs.length === stores.length && !a.wageless;
    area.laborLines = laborLines(a, labor, full ? area.fc : null, sameCut && full ? area.sales : null, ctx);
    // 標準比（対象店の合計）
    const ws = stores.filter(s => s.std && !s.std.missing);
    if (ws.length) {
      const T2 = k => sum(ws, s => s.std[k] || 0);
      const A = { cutRate: T2('cutStd') ? T2('cutHours') / T2('cutStd') : null, landRate: T2('landStd') ? T2('landHours') / T2('landStd') : null, n: ws.length };
      const wb = ws.filter(s => s.std.budgetRate != null);
      A.budgetRate = wb.length ? sum(wb, s => s.std.target) / sum(wb, s => s.std.landStd) : null;
      A.over = wb.length ? sum(wb, s => s.std.over) : null;
      area.std = A;
      const rk = [...ws].sort((x, y) => (y.std.landRate - (y.std.budgetRate || 0)) - (x.std.landRate - (x.std.budgetRate || 0)));
      area.laborLines.push({ t: `・標準比：実績${p0(A.cutRate)}（${cutTxt}）${labor.cut < labor.N ? `、着地見込み${p0(A.landRate)}` : ''}${A.budgetRate != null ? `（予算標準比${p0(A.budgetRate)}）` : ''}。${ws.length > 1 ? `予算標準比を最も上回る見込みの店：${rk.filter(s => s.std.budgetRate != null && s.std.over > 0).slice(0, 3).map(s => `${s.name}${p0(s.std.landRate)}（予算${p0(s.std.budgetRate)}、${hrs0(s.std.over * 60)}超過）`).join('、') || 'なし'}。` : ''}`, d: false });
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
    if (est.length) area.laborLines.push({ t: `・勤務予定の未作成：${est.join('、')}（実績平均で補っています）。`, d: true });
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
  return { tenure, labor, stores, area, firstDate, lastDate, isFullMonth, commonWeak, areaLabor, scopeLabel, titleLabel, multiArea, areas, hasHourly: hourly.length > 0,  hasPays: pays.length > 0 };
}

/* ===== コメント生成 ===== */
const man = v => { const m = v / 10000; return (Math.abs(m) >= 100 ? Math.round(m).toLocaleString('ja-JP') : (Math.round(m * 10) / 10).toLocaleString('ja-JP')) + '万円'; };
const sman = v => (v >= 0 ? '+' : '▲') + man(Math.abs(v));
const pct = v => v == null ? '－' : (v * 100).toFixed(1) + '%';
const yen = v => Math.round(v).toLocaleString('ja-JP') + '円';
const syen = v => (v >= 0 ? '+' : '▲') + Math.round(Math.abs(v)).toLocaleString('ja-JP') + '円';
const md = d => `${d.getMonth() + 1}/${d.getDate()}`;

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

  // 3) 時間帯
  const bands = s.byBand.filter(b => b.budget > s.budget * 0.02);
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
  out.push({ t: `　食べ放題フード：1人${m.foodQty.toFixed(1)}皿${vs(m.foodQty, A.foodQty, P && P.foodQty, '皿', 1)}・${Math.round(m.foodG).toLocaleString('ja-JP')}g${vs(m.foodG, A.foodG, P && P.foodG, 'g', 0)}。うち精肉${m.meatQty.toFixed(2)}皿・${Math.round(m.meatG)}g${vs(m.meatQty, A.meatQty, P && P.meatQty, '皿', 2)}。ドリンク${m.drinkQty.toFixed(1)}杯/人。`, d: false });
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
  const b = a.byBand.filter(x => x.budget > a.budget * 0.02).sort((x, y) => x.diff - y.diff);
  if (b.length) L.push(`・時間帯：${b.map(x => `${x.name}${pct(x.rate)}（${sman(x.diff)}）`).join('、')}。`);
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
  out.push({ t: `・着地見込み：${man(f.landing)}（月予算${man(f.monthBudget)}比${pct(f.landRate)}・${sman(f.landDiff)}）。${rem}の予算を、ここまでの曜日区分別の予算比で見込んだ値。`, d: false });
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
  for (const x of (s.inv ? invLines(s, ctx) : gapLines(s, ctx))) L.push(x);
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
    if (prem.length && prem[0].p > 0) L.push(`・コース構成：プレミアム比率の高い店 ${prem.slice(0, 3).map(u => `${u.x.name}${pct(u.p)}`).join('、')}／低い店 ${prem.slice(-2).reverse().map(u => `${u.x.name}${pct(u.p)}`).join('、')}。`);
  }
  if (!L.length) L.push('原価のデータがありません。');
  return L;
}

function commentText(lines, detail) { return lines.filter(l => detail || !l.d).map(l => l.t).join('\n'); }

if (typeof module !== 'undefined') module.exports = { tcSummary, wageSnapshot, mergeHistory, analyzeTenure, tenureLines, combineStd, isStdBook, extractStd, stdHours, analyzeStd, STD_DEFAULT, extractTimecard, extractSchedule, extractWage, analyzeLabor, forecastStore, forecastSum, pt2, extractInventory, isInventorySheet, extractMenu, menuAgg, scopeCounts, listAreas, sheetKind, extractHourly, extractPayments, commentText, SHEET_CUR, SHEET_PREV, findSheet, extractRows, analyze, DAYTYPES, BANDS };
