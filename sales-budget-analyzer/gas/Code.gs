/**
 * 予算差異アナライザー（Apps Script ウェブアプリ）
 *
 * 「予算差異アナライザー」フォルダ（マイドライブ・共有ドライブどちらでも可）の下に、
 *   「共通」フォルダ                      … 時給データ・標準比マスタ（見込君）・店舗情報
 *   ブロックのフォルダ（例：西一ブロック）
 *     └ エリアのフォルダ（例：髙谷エリア） … 売上集計表・タイムカード・勤務スケジュールなど
 * を置く（ブロックを作らず、ルートの直下にエリアのフォルダを置いてもよい）。
 * ウェブアプリは「アクセスしているユーザー」として動くため、各自がドライブで見られるフォルダのデータだけが表示される。
 * ルートを全員に共有しなくてよいように、フォルダの一覧（ブロック・エリアのID）はスクリプト プロパティに覚えておき、
 * 各自は一覧のうち自分が開けるエリアだけを使う。一覧はルートを開ける人（管理者）がウェブアプリを開くたびに更新される。
 *
 * 初回だけ：プロジェクトの設定 → スクリプト プロパティ に
 *   ROOT_FOLDER_ID ＝ 「予算差異アナライザー」フォルダのID（URLの folders/ の後ろ）
 * を登録する。
 */
const COMMON_NAME = '共通';
const SKIP_NAMES = ['過去', '保管', 'old'];       // この名前のフォルダは読まない
const HISTORY_NAME = '履歴データ';                // エリアのフォルダの中に作る、採用・定着の履歴（月ごとの小さなJSON）
const MAX_FILES_PER_FOLDER = 40;
const APP_VERSION = 'f559be8';                       // 版（build.py が書き込む。index.html と同じなら正しく貼り替えられている）
const TREE_KEY = 'folder_tree';                // フォルダの一覧（ブロック・エリア・共通のID）

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('予算差異アナライザー')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function rootFolder_() {
  const id = PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID');
  if (!id) throw new Error('スクリプト プロパティ ROOT_FOLDER_ID が未設定です（設定手順 SETUP.md を参照）');
  return DriveApp.getFolderById(id);
}

/**
 * ルートの下のフォルダの一覧。直下のフォルダのうち、中に（過去・履歴データ以外の）フォルダがあるものはブロック、ないものはエリア
 */
function buildTree_(root) {
  const skip = n => n === COMMON_NAME || n === HISTORY_NAME || SKIP_NAMES.indexOf(n) >= 0;
  const tree = { common: null, areas: [], at: new Date().toISOString() };
  const it = root.getFolders();
  while (it.hasNext()) {
    const f = it.next(), name = f.getName();
    if (name === COMMON_NAME) { tree.common = f.getId(); continue; }
    if (skip(name)) continue;
    const subs = []; const si = f.getFolders();
    while (si.hasNext()) { const g = si.next(); if (!skip(g.getName())) subs.push(g); }
    if (subs.length) subs.forEach(g => tree.areas.push({ id: g.getId(), name: g.getName(), block: name }));
    else tree.areas.push({ id: f.getId(), name: name, block: '' });
  }
  tree.areas.sort((x, y) => (x.block + '\u0000' + x.name).localeCompare(y.block + '\u0000' + y.name, 'ja'));
  return tree;
}
/** フォルダの一覧：ルートを開ける人はその場で作り直して覚える。開けない人は覚えてある一覧を使う */
function folderTree_() {
  const props = PropertiesService.getScriptProperties();
  try {
    const tree = buildTree_(rootFolder_());
    props.setProperty(TREE_KEY, JSON.stringify(tree));
    return tree;
  } catch (e) {
    const saved = props.getProperty(TREE_KEY);
    if (saved) return JSON.parse(saved);
    throw new Error('フォルダの一覧がまだありません。管理者（ルートのフォルダを開ける人）が一度ウェブアプリを開くか、エディタで refreshFolders を実行してください');
  }
}
/** エディタで実行：フォルダの一覧を作り直す（ブロック・エリアのフォルダを足したり名前を変えたりしたとき） */
function refreshFolders() {
  const t = buildTree_(rootFolder_());
  PropertiesService.getScriptProperties().setProperty(TREE_KEY, JSON.stringify(t));
  Logger.log('フォルダの一覧を更新しました：共通 ' + (t.common ? 'あり' : 'なし') + '、エリア ' + t.areas.length + '件\n' + t.areas.map(a => (a.block ? a.block + ' / ' : '') + a.name).join('\n'));
}
const canOpen_ = id => { try { return DriveApp.getFolderById(id).getName() != null; } catch (e) { return false; } };

/** 画面の初期表示に使う：見られるエリアのフォルダ（ブロックつき）、共通フォルダ、店舗設定 */
function getConfig() {
  let tree;
  try { tree = folderTree_(); } catch (e) { return { areas: [], error: e.message, stdSettings: getStdSettings_(), actions: getActions_(), version: APP_VERSION }; }
  let user = '';
  try { user = Session.getActiveUser().getEmail(); } catch (e) { /* 取得できなくても続ける */ }
  const me = String(user).toLowerCase();
  const areas = [];
  tree.areas.forEach(a => {
    let f; try { f = DriveApp.getFolderById(a.id); f.getName(); } catch (e) { return; } // 開けないエリアは出さない
    // 自分が直接「編集者」になっているエリア＝自分のエリア（最初に選ぶ）
    let mine = false;
    if (me) { try { mine = f.getEditors().some(u => String(u.getEmail()).toLowerCase() === me) || String(f.getOwner() && f.getOwner().getEmail()).toLowerCase() === me; } catch (e) { /* 取れないときは false */ } }
    areas.push({ id: a.id, name: a.name, block: a.block || '', mine: mine });
  });
  const common = tree.common && canOpen_(tree.common) ? tree.common : null;
  return { areas: areas, common: common, user: user, stdSettings: getStdSettings_(), actions: getActions_(), version: APP_VERSION };
}

/** エリアのフォルダ（複数可）と共通フォルダのファイル一覧。新しい順 */
function listFiles(areaIds) {
  const tree = folderTree_();
  const folders = (areaIds || []).map(id => DriveApp.getFolderById(id));
  if (tree.common && canOpen_(tree.common)) folders.push(DriveApp.getFolderById(tree.common));
  const out = [];
  folders.forEach(folder => {
    const files = [];
    const it = folder.getFiles();
    while (it.hasNext()) {
      const f = it.next(), name = f.getName(), mime = f.getMimeType();
      if (/^~\$/.test(name)) continue; // Excelの一時ファイル
      if (!/\.(csv|txt|xlsx|xlsm|xls)$/i.test(name) && mime !== MimeType.GOOGLE_SHEETS) continue;
      files.push({ id: f.getId(), name: name, folder: folder.getName(), folderId: folder.getId(), common: folder.getId() === tree.common, updated: f.getLastUpdated().getTime(), size: f.getSize() });
    }
    files.sort((a, b) => b.updated - a.updated);
    files.slice(0, MAX_FILES_PER_FOLDER).forEach(x => out.push(x));
  });
  return out;
}

/** ファイルの中身（base64）。Googleスプレッドシートは .xlsx に変換して返す */
function getFile(id) {
  const f = DriveApp.getFileById(id);
  let name = f.getName(), blob;
  if (f.getMimeType() === MimeType.GOOGLE_SHEETS) {
    const url = 'https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx';
    blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
    if (!/\.xlsx$/i.test(name)) name += '.xlsx';
  } else {
    blob = f.getBlob();
  }
  return { name: name, b64: Utilities.base64Encode(blob.getBytes()) };
}

/* ===== 訪問資料の総評：手で直したアクションポイント（企業×月ごと。全員で共有） ===== */
function getActions_() {
  const all = PropertiesService.getScriptProperties().getProperties(), out = {};
  Object.keys(all).forEach(k => { if (k.indexOf('act_') === 0) { try { out[k.slice(4)] = JSON.parse(all[k]); } catch (e) { /* 壊れた値は無視 */ } } });
  return out;
}
/** key＝「企業名|月」。lines＝文の配列（null で自動の下書きに戻す） */
function saveAction(key, lines) {
  key = String(key || '');
  if (!key || key.length > 200) throw new Error('キーが不正です');
  const props = PropertiesService.getScriptProperties();
  if (!lines) { props.deleteProperty('act_' + key); return true; }
  const keep = (lines || []).map(t => String(t).slice(0, 400)).filter(t => t.trim()).slice(0, 10);
  props.setProperty('act_' + key, JSON.stringify({ lines: keep, by: Session.getActiveUser().getEmail() || '', at: new Date().toISOString() }));
  return true;
}

/* ===== 標準比の店舗設定（ランチ営業・T8台数など）。全員で共有する ===== */
function getStdSettings_() {
  const all = PropertiesService.getScriptProperties().getProperties(), out = {};
  Object.keys(all).forEach(k => { if (k.indexOf('std_') === 0) { try { out[k.slice(4)] = JSON.parse(all[k]); } catch (e) { /* 壊れた値は無視 */ } } });
  return out;
}
function saveStdSetting(code, v) {
  if (!/^[0-9A-Za-z]+$/.test(String(code))) throw new Error('店舗コードが不正です');
  const keep = {};
  ['tables', 't8', 'hoursWd', 'hoursHd'].forEach(k => { if (v[k] !== '' && v[k] != null && isFinite(+v[k])) keep[k] = +v[k]; });
  ['lunchWd', 'lunchHd', 'auto'].forEach(k => { keep[k] = !!v[k]; });
  keep.by = Session.getActiveUser().getEmail() || '';
  keep.at = new Date().toISOString();
  PropertiesService.getScriptProperties().setProperty('std_' + code, JSON.stringify(keep));
  return true;
}

/* ===== 採用・定着の履歴（タイムカード・時給データを月ごとにまとめたもの） ===== */
/** エリアの「履歴データ」フォルダにある JSON を返す（kind: tc＝タイムカード、wage＝時給データ、nps＝NPS） */
function getHistory(areaIds) {
  const out = { tc: [], wage: [], nps: [] };
  (areaIds || []).forEach(id => {
    // 「履歴データ」が2つ以上できてしまっていても、すべて読む
    const it = DriveApp.getFolderById(id).getFoldersByName(HISTORY_NAME);
    while (it.hasNext()) {
    const files = it.next().getFiles();
    while (files.hasNext()) {
      const f = files.next(), m = f.getName().match(/^(tc|wage|nps)_(\d{4}-\d{2})\.json$/);
      if (!m) continue;
      try { const x = JSON.parse(f.getBlob().getDataAsString('UTF-8')); x.area = id; out[m[1]].push(JSON.stringify(x)); } catch (e) { /* 壊れたファイルは無視 */ }
    }
    }
  });
  return out;
}
/** 履歴を保存（同じ月は上書き）。エリアのフォルダの編集権限がない人は保存できない（表示はできる） */
function saveHistory(areaId, kind, ym, json) {
  if (!/^(tc|wage|nps)$/.test(kind) || !/^\d{4}-\d{2}$/.test(ym)) throw new Error('保存する内容が不正です');
  // 同時に保存すると「履歴データ」が2つできることがあるので、1人ずつ順番に保存する
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const area = DriveApp.getFolderById(areaId);
    const it = area.getFoldersByName(HISTORY_NAME);
    const folder = it.hasNext() ? it.next() : area.createFolder(HISTORY_NAME);
    const name = kind + '_' + ym + '.json';
    // 同じ名前のファイルは、ほかの「履歴データ」にあるものも含めて上書き（なければ1つ目に作る）
    let done = false; const all = area.getFoldersByName(HISTORY_NAME);
    while (all.hasNext()) { const fo = all.next(), old = fo.getFilesByName(name); while (old.hasNext()) { old.next().setContent(json); done = true; } }
    if (!done) folder.createFile(name, json, MimeType.PLAIN_TEXT);
  } finally { lock.releaseLock(); }
  return true;
}

/* ===== 初期設定・確認（エディタで1回だけ実行する） ===== */
// 作るブロックとエリアのフォルダ名（必要に応じて書き換えてから setupFolders を実行）。
// '' に書いたエリアはブロックに入れず、ルートの直下（「共通」と同じ階層）に作る
//   西山エリア（東一）・陳エリア（西二）は、それぞれのブロックを作るまでルートの直下に置く。
//   ブロックを作るときは '東一ブロック': ['西山エリア'] のように移して setupFolders を実行すると、中身ごとブロックの下へ移る
const BLOCK_FOLDERS = {
  '西一ブロック': ['河野エリア', '金エリア', '松山エリア', '菅エリア', '名越エリア', '髙谷エリア'],
  '': ['西山エリア', '陳エリア'],
};

/**
 * ROOT_FOLDER_ID のフォルダの下に「共通」と、ブロック／エリアのフォルダを作る（すでにあるものは作らない）。
 * ルートの直下にあるエリアのフォルダは、同じ名前のエリアがブロックの下に指定されていれば、そのブロックへ移す（中身ごと）
 */
function setupFolders() {
  const root = rootFolder_();
  const get = (parent, name) => { const i = parent.getFoldersByName(name); if (i.hasNext()) { Logger.log('あり：' + name); return i.next(); } Logger.log('作成：' + name); return parent.createFolder(name); };
  get(root, COMMON_NAME);
  Object.keys(BLOCK_FOLDERS).forEach(block => {
    const parent = block ? get(root, block) : root;
    BLOCK_FOLDERS[block].forEach(area => {
      if (block) { const old = root.getFoldersByName(area); if (old.hasNext() && !parent.getFoldersByName(area).hasNext()) { old.next().moveTo(parent); Logger.log('移動：' + area + ' → ' + block); return; } }
      get(parent, area);
    });
  });
  refreshFolders();
  Logger.log('完了。エリアのフォルダはAMに「編集者」、ブロックのフォルダはブロック長に「閲覧者」で共有してください（ルートは共有しなくてよい）。「共通」は社内全員を閲覧者にしてよい（時給データは全社に配布済みの資料）。');
}

/** 設定の確認：フォルダが見えるか、各フォルダのファイル数、店舗設定の数をログに出す */
function checkSetup() {
  let root;
  try { root = rootFolder_(); } catch (e) { Logger.log('NG：' + e.message); return; }
  Logger.log('OK：ルートフォルダ「' + root.getName() + '」');
  refreshFolders();
  const cfg = getConfig();
  Logger.log('実行ユーザー：' + (cfg.user || '（取得できません）'));
  Logger.log('共通フォルダ：' + (cfg.common ? 'あり' : 'なし（「共通」という名前のフォルダを作ってください）'));
  if (!cfg.areas.length) Logger.log('NG：エリアのフォルダがありません');
  cfg.areas.forEach(a => {
    const files = listFiles([a.id]).filter(f => !f.common);
    Logger.log((a.block ? a.block + ' / ' : '') + 'エリア「' + a.name + '」：ファイル ' + files.length + '件' + (files.length ? '（' + files.slice(0, 5).map(f => f.name).join('、') + (files.length > 5 ? ' ほか' : '') + '）' : ''));
  });
  const common = cfg.common ? listFiles([]).filter(f => f.common) : [];
  Logger.log('共通のファイル：' + common.length + '件' + (common.length ? '（' + common.map(f => f.name).join('、') + '）' : ''));
  Logger.log('標準比の店舗設定：' + Object.keys(cfg.stdSettings).length + '店');
}

/* ===== 自動コピー（NPS・時給データ）：元のフォルダの最新のファイルを「共通」へ。毎朝7時台（installAutoCopy） ===== */
/* ----- NPS：NPSのフォルダの最新の「月次サマリ」 ----- */
// NPSのフォルダ（毎月「261102_Monthly_焼肉」のような月のフォルダができ、前の月は「アーカイブ」へ移る）。
// 変えるときはスクリプト プロパティ NPS_SOURCE_ID に入れる
const NPS_SOURCE_DEFAULT = '12xXm7FyXTYbwpwNu-8PQpqczLdtkQUvk';
const NPS_PREFIX = 'NPS_';
/** 月のフォルダ（名前の先頭が6桁の日付）ごとの月次サマリ（新しい順。アーカイブの中も見る） */
function npsMonthFiles_(n) {
  const src = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('NPS_SOURCE_ID') || NPS_SOURCE_DEFAULT);
  const months = [];
  const scan = folder => { const it = folder.getFolders(); while (it.hasNext()) { const f = it.next(), m = f.getName().match(/^(\d{6})_/); if (m) months.push({ key: m[1], folder: f }); else if (/アーカイブ|archive/i.test(f.getName())) scan(f); } };
  scan(src);
  months.sort((a, b) => b.key.localeCompare(a.key));
  const out = [];
  for (const mo of months) {
    if (out.length >= (n || 1)) break;
    const it = mo.folder.getFiles(); let best = null;
    while (it.hasNext()) { const f = it.next(), name = f.getName(), mime = f.getMimeType();
      const sheet = /\.xlsx?$/i.test(name) || mime === MimeType.GOOGLE_SHEETS || mime === MimeType.MICROSOFT_EXCEL;
      if (!sheet) continue;
      if (!best || /サマリ/.test(name)) best = f; }
    if (best) out.push({ key: mo.key, folder: mo.folder.getName(), file: best });
  }
  return out;
}
function commonFolder_() { const t = folderTree_(); if (!t.common) throw new Error('「共通」フォルダがありません'); return DriveApp.getFolderById(t.common); }
/** 「共通」にある前のコピー（名前が prefix で始まるもの）を「共通/過去」へ移す */
function moveOldNps_(common, keepIds) { return moveOldPrefix_(common, NPS_PREFIX, keepIds); }
function moveOldPrefix_(common, prefix, keepIds) {
  const it = common.getFiles(); let past = null, n = 0;
  while (it.hasNext()) { const f = it.next(); if (f.getName().indexOf(prefix) !== 0 || (keepIds || []).indexOf(f.getId()) >= 0) continue;
    if (!past) { const p = common.getFoldersByName(ARCHIVE_NAME); past = p.hasNext() ? p.next() : common.createFolder(ARCHIVE_NAME); }
    f.moveTo(past); n++; }
  return n;
}
/** 最新の月次サマリが新しくなっていたら「共通」へコピーする（同じものならなにもしない）。毎朝のトリガーからも呼ばれる */
function copyLatestNps() {
  const list = npsMonthFiles_(1);
  if (!list.length) { Logger.log('NPSのフォルダに月次サマリが見つかりません'); return; }
  copyToCommon_(list[0].file, NPS_PREFIX, 'nps_copied', 'NPS（' + list[0].folder + '）');
}
/** src を「共通」へ prefix＋元の名前でコピー。同じもの（ID・更新日時が同じ）がコピー済みなら何もしない。前のコピーは「共通/過去」へ */
function copyToCommon_(src, prefix, propKey, label) {
  const key = src.getId() + '@' + src.getLastUpdated().getTime();
  const props = PropertiesService.getScriptProperties(), common = commonFolder_();
  const name = prefix + src.getName();
  if (props.getProperty(propKey) === key && common.getFilesByName(name).hasNext()) { Logger.log(label + 'は最新です'); return false; }
  const copy = src.makeCopy(name, common);
  const moved = moveOldPrefix_(common, prefix, [copy.getId()]);
  props.setProperty(propKey, key);
  Logger.log(label + 'をコピーしました：' + src.getName() + ' → 共通/' + name + (moved ? '（前の ' + moved + '件を 共通/過去 へ）' : ''));
  return true;
}

/* ===== 時給データの自動コピー：時給のフォルダの最新のファイルを「共通」へ ===== */
// 時給データ（【店長資料】①有休残日数＆時給・所属・甲乙確認_2026.11≪発信用≫.xlsx など）が毎月入るフォルダ。変えるときは WAGE_SOURCE_ID
const WAGE_SOURCE_DEFAULT = '1Khr9RG-LIw0GVv1gLpy-_zYBFEZ8gx58';
const WAGE_PREFIX = 'WAGE_';
/** フォルダ直下の Excel／スプレッドシートのうち、名前の年月（2026.11 など）がいちばん新しいもの（なければ更新日時の新しいもの）。名前に「時給」があるものを優先 */
function latestWageFile_() {
  const src = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('WAGE_SOURCE_ID') || WAGE_SOURCE_DEFAULT);
  const it = src.getFiles(), list = [];
  while (it.hasNext()) { const f = it.next(), name = f.getName(), mime = f.getMimeType();
    if (!(/\.xls[xm]?$/i.test(name) || mime === MimeType.GOOGLE_SHEETS || mime === MimeType.MICROSOFT_EXCEL)) continue;
    const m = name.match(/(20\d{2})[.\-年\/_]\s*(\d{1,2})/);
    list.push({ f, wage: /時給/.test(name) ? 1 : 0, ym: m ? +m[1] * 12 + +m[2] : 0, up: f.getLastUpdated().getTime() }); }
  list.sort((a, b) => b.wage - a.wage || b.ym - a.ym || b.up - a.up);
  return list.length ? list[0].f : null;
}
/** 最新の時給データが新しくなっていたら「共通」へコピーする */
function copyLatestWage() {
  const f = latestWageFile_();
  if (!f) { Logger.log('時給のフォルダにファイルが見つかりません'); return; }
  copyToCommon_(f, WAGE_PREFIX, 'wage_copied', '時給データ');
}

/** 自動コピーをまとめて動かす（NPS・時給データ）。どちらかが失敗しても、もう片方は動かす */
function autoCopy_() {
  try { copyLatestNps(); } catch (e) { Logger.log('NPSのコピーに失敗：' + e.message); }
  try { copyLatestWage(); } catch (e) { Logger.log('時給データのコピーに失敗：' + e.message); }
}
/**
 * 過去のNPSをまとめて取り込む（最初に1回）：直近 n か月の月次サマリを「共通」へコピーする。
 * そのあとウェブアプリを「すべてのエリア」で1回開くと、各エリアの「履歴データ」に月ごとのNPSが積み上がる。
 * 最後に copyLatestNps を実行すると、最新の月以外は「共通/過去」へ片付く（読み込みが速くなる）
 */
function backfillNps() {
  const n = 6, list = npsMonthFiles_(n), common = commonFolder_();
  moveOldNps_(common, []);
  list.forEach(x => { x.file.makeCopy(NPS_PREFIX + x.file.getName(), common); Logger.log('コピー：' + x.folder + ' の ' + x.file.getName()); });
  PropertiesService.getScriptProperties().deleteProperty('nps_copied');
  Logger.log(list.length + 'か月分を「共通」へコピーしました。ウェブアプリを「すべてのエリア」で1回開いてから、copyLatestNps を実行してください。');
}
/** 自動コピー（NPS・時給データ）を毎朝7時台に動かす（1回だけ実行）。止めるときは uninstallAutoCopy */
function installAutoCopy() {
  ScriptApp.getProjectTriggers().filter(t => ['npsCopy_', 'autoCopy_'].indexOf(t.getHandlerFunction()) >= 0).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('autoCopy_').timeBased().everyDays(1).atHour(7).inTimezone('Asia/Tokyo').create();
  autoCopy_();
  Logger.log('毎朝7時台に、NPSの最新の月次サマリと最新の時給データを確認して、新しければ「共通」へコピーする設定をしました。');
}
function uninstallAutoCopy() {
  const ts = ScriptApp.getProjectTriggers().filter(t => ['npsCopy_', 'autoCopy_'].indexOf(t.getHandlerFunction()) >= 0);
  ts.forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('自動コピーを止めました（' + ts.length + '件）。');
}
// 以前の名前（NPSだけのとき）でも同じものが動くように
function installNpsCopy() { installAutoCopy(); }
function uninstallNpsCopy() { uninstallAutoCopy(); }
function npsCopy_() { autoCopy_(); }

/* ===== 月初の自動整理：古いファイルを各エリアの「過去」フォルダへ移す ===== */
const ARCHIVE_NAME = '過去';
/**
 * 各エリアのフォルダ直下で、最終更新が「前月1日より前」のファイルを「エリア/過去/YYYY-MM」へ移す。
 * 例：12月1日に実行 → 10月以前に更新されたファイルを移す（11月・12月に入れたものは残る）。
 * 「共通」「履歴データ」とフォルダの中のフォルダには触らない。
 * dryRun=true なら移さずに、移す予定のファイルを返す（ログにも出す）。お知らせのメールは送らない（結果は実行ログで確認）。
 */
function archiveOldFiles(dryRun) {
  const tree = buildTree_(rootFolder_()), now = new Date();
  PropertiesService.getScriptProperties().setProperty(TREE_KEY, JSON.stringify(tree)); // ついでに一覧も新しくする
  const limit = new Date(now.getFullYear(), now.getMonth() - 1, 1); // 前月1日
  const moved = [];
  tree.areas.forEach(a => {
    const area = DriveApp.getFolderById(a.id), name = (a.block ? a.block + '／' : '') + a.name;
    let archive = null;
    const files = area.getFiles();
    while (files.hasNext()) {
      const f = files.next(), up = f.getLastUpdated();
      if (up >= limit) continue;
      const ym = Utilities.formatDate(up, 'Asia/Tokyo', 'yyyy-MM');
      moved.push(name + '／' + f.getName() + '（' + Utilities.formatDate(up, 'Asia/Tokyo', 'yyyy/MM/dd') + '更新 → ' + ARCHIVE_NAME + '/' + ym + '）');
      if (dryRun) continue;
      if (!archive) { const a = area.getFoldersByName(ARCHIVE_NAME); archive = a.hasNext() ? a.next() : area.createFolder(ARCHIVE_NAME); }
      const m = archive.getFoldersByName(ym), dest = m.hasNext() ? m.next() : archive.createFolder(ym);
      f.moveTo(dest);
    }
  });
  Logger.log((dryRun ? '【移す予定】' : '【移しました】') + moved.length + '件\n' + moved.join('\n'));
  return moved;
}
/** 移す予定のファイルを確認するだけ（移さない） */
function previewArchive() { return archiveOldFiles(true); }
/** 自動整理を毎月1日の朝5時に動かす（1回だけ実行）。すでにあれば作り直す */
function installMonthlyArchive() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'monthlyArchive_').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('monthlyArchive_').timeBased().onMonthDay(1).atHour(5).inTimezone('Asia/Tokyo').create();
  Logger.log('毎月1日の5時台に自動整理を動かす設定をしました。止めるときは uninstallMonthlyArchive を実行してください。');
}
function uninstallMonthlyArchive() {
  const ts = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'monthlyArchive_');
  ts.forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('自動整理を止めました（' + ts.length + '件）。');
}
function monthlyArchive_() { archiveOldFiles(false); }
