/**
 * 予算差異アナライザー（Apps Script ウェブアプリ）
 *
 * 共有ドライブの「予算差異アナライザー」フォルダの下に、
 *   エリアごとのフォルダ（例：髙谷エリア）… 売上集計表・タイムカード・勤務スケジュールなど
 *   「共通」フォルダ                    … 時給データ・標準比マスタ（見込君）
 * を置く。ウェブアプリは「アクセスしているユーザー」として動くため、
 * 各自がドライブで見られるフォルダのデータだけが表示される。
 *
 * 初回だけ：プロジェクトの設定 → スクリプト プロパティ に
 *   ROOT_FOLDER_ID ＝ 「予算差異アナライザー」フォルダのID（URLの folders/ の後ろ）
 * を登録する。
 */
const COMMON_NAME = '共通';
const SKIP_NAMES = ['過去', '保管', 'old'];       // この名前のフォルダは読まない
const HISTORY_NAME = '履歴データ';                // エリアのフォルダの中に作る、採用・定着の履歴（月ごとの小さなJSON）
const MAX_FILES_PER_FOLDER = 40;

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

/** 画面の初期表示に使う：見られるエリアのフォルダ、共通フォルダ、店舗設定 */
function getConfig() {
  let root;
  try { root = rootFolder_(); } catch (e) { return { areas: [], error: e.message, stdSettings: getStdSettings_() }; }
  const areas = [];
  let common = null;
  const it = root.getFolders();
  while (it.hasNext()) {
    const f = it.next(), name = f.getName();
    if (name === COMMON_NAME) { common = f.getId(); continue; }
    if (SKIP_NAMES.indexOf(name) >= 0) continue;
    areas.push({ id: f.getId(), name: name });
  }
  areas.sort((a, b) => a.name.localeCompare(b.name, 'ja'));
  let user = '';
  try { user = Session.getActiveUser().getEmail(); } catch (e) { /* 取得できなくても続ける */ }
  return { areas: areas, common: common, user: user, stdSettings: getStdSettings_() };
}

/** エリアのフォルダ（複数可）と共通フォルダのファイル一覧。新しい順 */
function listFiles(areaIds) {
  const root = rootFolder_();
  const folders = (areaIds || []).map(id => DriveApp.getFolderById(id));
  const c = root.getFoldersByName(COMMON_NAME);
  if (c.hasNext()) folders.push(c.next());
  const out = [];
  folders.forEach(folder => {
    const files = [];
    const it = folder.getFiles();
    while (it.hasNext()) {
      const f = it.next(), name = f.getName(), mime = f.getMimeType();
      if (/^~\$/.test(name)) continue; // Excelの一時ファイル
      if (!/\.(csv|txt|xlsx|xlsm|xls)$/i.test(name) && mime !== MimeType.GOOGLE_SHEETS) continue;
      files.push({ id: f.getId(), name: name, folder: folder.getName(), folderId: folder.getId(), common: folder.getName() === COMMON_NAME, updated: f.getLastUpdated().getTime(), size: f.getSize() });
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
/** エリアの「履歴データ」フォルダにある JSON を返す（kind: tc＝タイムカード、wage＝時給データ） */
function getHistory(areaIds) {
  const out = { tc: [], wage: [] };
  (areaIds || []).forEach(id => {
    const it = DriveApp.getFolderById(id).getFoldersByName(HISTORY_NAME);
    if (!it.hasNext()) return;
    const files = it.next().getFiles();
    while (files.hasNext()) {
      const f = files.next(), m = f.getName().match(/^(tc|wage)_(\d{4}-\d{2})\.json$/);
      if (!m) continue;
      try { const x = JSON.parse(f.getBlob().getDataAsString('UTF-8')); x.area = id; out[m[1]].push(JSON.stringify(x)); } catch (e) { /* 壊れたファイルは無視 */ }
    }
  });
  return out;
}
/** 履歴を保存（同じ月は上書き）。エリアのフォルダの編集権限がない人は保存できない（表示はできる） */
function saveHistory(areaId, kind, ym, json) {
  if (!/^(tc|wage)$/.test(kind) || !/^\d{4}-\d{2}$/.test(ym)) throw new Error('保存する内容が不正です');
  const area = DriveApp.getFolderById(areaId);
  const it = area.getFoldersByName(HISTORY_NAME);
  const folder = it.hasNext() ? it.next() : area.createFolder(HISTORY_NAME);
  const name = kind + '_' + ym + '.json';
  const old = folder.getFilesByName(name);
  if (old.hasNext()) old.next().setContent(json);
  else folder.createFile(name, json, MimeType.PLAIN_TEXT);
  return true;
}
