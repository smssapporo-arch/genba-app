/**
 * 施工現場管理2026 — 入力アプリ用バックエンド
 *
 * 【設置手順】
 *  1. 対象スプレッドシートを開く → 拡張機能 → Apps Script
 *  2. このファイルの中身を全部貼り付ける
 *  3. 下の TOKEN を自分だけが知っている文字列に書き換える
 *  4. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *       実行ユーザー : 自分
 *       アクセスできるユーザー : 全員
 *  5. 発行されたURL（.../exec）を index.html の GAS_URL に貼る
 *
 * 【方針】
 *  - このスクリプトが書き込むのは JOBシートと、3つのマスタシート
 *    （取引会社／取引先別営業担当／取引先元請営業担当一覧）だけ。
 *  - 「次請求締日」「支払スパン」以外は、アプリが持つ列にしか触れない。
 *  - 更新時は変更された項目だけを書く。数式が入っているセルは書き換えずスキップする。
 *  - マスタへの追加は行の末尾に足すだけで、既存行・既存列の位置は動かさない。
 */

// ── 設定 ──────────────────────────────────────────────
var TOKEN     = 'CHANGE_ME_TOKEN';   // ★必ず変更する
var JOB_GID   = 183661348;           // JOBシートのID（タブ名を変えても動く）
var SYNC_DAYS = 60;                  // アプリに配る直近日数
// ─────────────────────────────────────────────────────

// キー → シートの見出し文字（JOB）
var FIELDS = {
  jobId:     'JOBID',
  planDate:  '作業予定日',
  doneDate:  '完了日',
  companyId: '取引会社ID',
  repId:     '取引先別営業担当ID',
  clientId:  '取引先元請営業担当一覧ID',
  site:      '現場名',
  work:      '内容',
  amount:    '金額',
  parking:   '駐車場代',
  toll:      '高速代',
  note:      '備考',
  assessed:  '査定有無',
  billed:    '請求有無',
  fromIc:    '出発IC',
  toIc:      '到着IC',
  enteredBy: '入力者'
};

var DATE_KEYS = { planDate: 1, doneDate: 1 };
var NUM_KEYS  = { amount: 1, parking: 1, toll: 1 };
var BOOL_KEYS = { assessed: 1, billed: 1 };
var ID_KEYS   = { companyId: 1, repId: 1, clientId: 1 };

// マスタ3シートの列マッピング（見出し文字はシートの実データに合わせてある）
var COMPANY_FIELDS = { id: '取引会社ID', name: '会社名', address: '住所', tel: '電話番号',
                        fax: 'FAX', contact: '担当者', closing: '請求締日', span: '支払スパン', short: '短縮名称' };
var REP_FIELDS     = { id: '取引先別営業担当ID', sort: '並べ替え列', companyId: '取引会社ID',
                        name: '営業担当', title: '役職', updatedAt: '更新日' };
var CLIENT_FIELDS  = { id: '取引先元請営業担当一覧ID', name: '販売先', companyId: '取引会社ID',
                        repId: '取引先別営業担当ID', hidden: '非表示' };

function doGet(e) {
  return respond(function () { return handle(e.parameter || {}, null); });
}

function doPost(e) {
  return respond(function () {
    var body = {};
    try { body = JSON.parse(e.postData.contents); } catch (err) { throw new Error('リクエストの形式が不正です'); }
    return handle(e.parameter || {}, body);
  });
}

function respond(fn) {
  var out;
  try { out = fn(); out.ok = true; }
  catch (err) { out = { ok: false, error: String(err && err.message ? err.message : err) }; }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function handle(params, body) {
  var token = (body && body.token) || params.token;
  if (token !== TOKEN) throw new Error('合言葉が違います');

  if (body && body.ops && body.ops.length) {
    return applyOps(body.ops, body.user || '');
  }
  return bootstrap();
}

// ── マスタシートの特定 ──────────────────────────────

/** 3つのマスタシートを1回のスキャンで見分けて返す */
function findMasterSheets() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var found = { company: null, rep: null, client: null };
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    if (sh.getSheetId() === JOB_GID) continue;
    var map = headerMap(sh);
    if (!found.company && map['会社名'] && map['取引会社ID']) found.company = sh;
    else if (!found.rep && map['取引先別営業担当ID'] && map['営業担当']) found.rep = sh;
    else if (!found.client && map['取引先元請営業担当一覧ID'] && map['販売先']) found.client = sh;
  }
  return found;
}

// ── 読み取り ──────────────────────────────────────────

function bootstrap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var m = findMasterSheets();

  var companies = m.company ? readRows(m.company, headerMap(m.company), tz, COMPANY_FIELDS, 'id') : [];
  var reps      = m.rep     ? readRows(m.rep,     headerMap(m.rep),     tz, REP_FIELDS,     'id') : [];
  var clients   = m.client  ? readRows(m.client,  headerMap(m.client),  tz, CLIENT_FIELDS,  'id') : [];

  return {
    masters: { companies: companies, reps: reps, clients: clients },
    jobs: readJobs(tz),
    syncedAt: Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss")
  };
}

function headerMap(sh) {
  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return {};
  var row = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var map = {};
  for (var c = 0; c < row.length; c++) {
    var key = String(row[c]).trim();
    if (key && !map[key]) map[key] = c + 1;
  }
  return map;
}

function readRows(sh, map, tz, spec, requiredKey) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  var keys = Object.keys(spec);
  var out = [];
  for (var r = 0; r < values.length; r++) {
    var obj = {};
    for (var k = 0; k < keys.length; k++) {
      var col = map[spec[keys[k]]];
      obj[keys[k]] = col ? cellToText(values[r][col - 1], tz) : '';
    }
    if (String(obj[requiredKey]) === '') continue;
    out.push(obj);
  }
  return out;
}

function readJobs(tz) {
  var sh = jobSheet();
  var map = headerMap(sh);
  var last = sh.getLastRow();
  if (last < 2) return [];

  var values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  var limit = new Date();
  limit.setDate(limit.getDate() - SYNC_DAYS);
  var limitText = Utilities.formatDate(limit, tz, 'yyyy-MM-dd');

  var keys = Object.keys(FIELDS);
  var out = [];
  for (var r = 0; r < values.length; r++) {
    var job = {};
    for (var k = 0; k < keys.length; k++) {
      var col = map[FIELDS[keys[k]]];
      job[keys[k]] = col ? cellToText(values[r][col - 1], tz) : '';
    }
    if (!job.jobId) continue;
    var newest = job.doneDate > job.planDate ? job.doneDate : job.planDate;
    if (newest && newest < limitText) continue;
    out.push(job);
  }
  return out;
}

function cellToText(v, tz) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  if (v === true) return 'TRUE';
  if (v === false) return 'FALSE';
  return String(v).trim();
}

function jobSheet() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === JOB_GID) return sheets[i];
  }
  throw new Error('JOBシートが見つかりません（gid=' + JOB_GID + '）');
}

// ── 書き込み ──────────────────────────────────────────

function applyOps(ops, user) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone();
    var sh = jobSheet();
    var map = ensureEnteredByColumn(sh);
    var idCol = map[FIELDS.jobId];
    if (!idCol) throw new Error('JOBID列が見つかりません');

    var index = buildIdIndex(sh, idCol);
    var masters = findMasterSheets();
    var results = [];

    for (var i = 0; i < ops.length; i++) {
      var op = ops[i];
      try {
        if (op.type === 'create')          results.push(doCreate(sh, map, index, op, user, tz));
        else if (op.type === 'update')     results.push(doUpdate(sh, map, index, op, user, tz));
        else if (op.type === 'delete')     results.push(doDelete(sh, map, index, op));
        else if (op.type === 'createCompany') results.push(doCreateMaster(masters.company, COMPANY_FIELDS, op, companyDefaults(op)));
        else if (op.type === 'createRep')     results.push(doCreateMaster(masters.rep, REP_FIELDS, op, repDefaults(op, tz)));
        else if (op.type === 'createClient')  results.push(doCreateMaster(masters.client, CLIENT_FIELDS, op, clientDefaults(op)));
        else throw new Error('不明な操作: ' + op.type);
      } catch (err) {
        results.push({ opId: op.opId, ok: false, error: String(err && err.message ? err.message : err) });
      }
    }
    SpreadsheetApp.flush();
    return { results: results, syncedAt: Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss") };
  } finally {
    lock.releaseLock();
  }
}

/** 「入力者」列が無ければ末尾に作る */
function ensureEnteredByColumn(sh) {
  var map = headerMap(sh);
  if (!map[FIELDS.enteredBy]) {
    var col = sh.getLastColumn() + 1;
    sh.getRange(1, col).setValue(FIELDS.enteredBy);
    map[FIELDS.enteredBy] = col;
  }
  return map;
}

function buildIdIndex(sh, idCol) {
  var last = sh.getLastRow();
  var index = {};
  if (last < 2) return index;
  var ids = sh.getRange(2, idCol, last - 1, 1).getDisplayValues();
  for (var r = 0; r < ids.length; r++) {
    var id = String(ids[r][0]).trim();
    if (id) index[id] = r + 2;
  }
  return index;
}

function doCreate(sh, map, index, op, user, tz) {
  var id = String(op.jobId || '').trim();
  if (!id) throw new Error('JOBIDがありません');

  // 同じJOBIDが既にある＝オフライン再送。重複させず更新にまわす。
  if (index[id]) return doUpdate(sh, map, index, op, user, tz);

  var row = sh.getLastRow() + 1;
  var fields = op.fields || {};
  fields.jobId = id;
  if (!fields.enteredBy) fields.enteredBy = user;
  if (fields.assessed === undefined) fields.assessed = 'FALSE';
  if (fields.billed === undefined)   fields.billed   = 'FALSE';

  writeFields(sh, map, row, fields, false, tz);
  index[id] = row;
  return { opId: op.opId, ok: true, jobId: id, row: row };
}

function doUpdate(sh, map, index, op, user, tz) {
  var id = String(op.jobId || '').trim();
  var row = index[id];
  if (!row) throw new Error('JOBID ' + id + ' が見つかりません');

  var fields = {};
  var src = op.fields || {};
  for (var k in src) { if (k !== 'jobId') fields[k] = src[k]; }
  fields.enteredBy = src.enteredBy || user;

  writeFields(sh, map, row, fields, true, tz);
  return { opId: op.opId, ok: true, jobId: id, row: row };
}

function doDelete(sh, map, index, op) {
  var id = String(op.jobId || '').trim();
  var row = index[id];
  if (!row) return { opId: op.opId, ok: true, jobId: id, note: '既に削除済み' };

  sh.deleteRow(row);
  delete index[id];
  for (var key in index) { if (index[key] > row) index[key] = index[key] - 1; }
  return { opId: op.opId, ok: true, jobId: id, deleted: true };
}

/** 指定された項目だけを書く。数式セルは触らない。 */
function writeFields(sh, map, row, fields, skipFormula, tz) {
  var colVal = {};
  for (var key in fields) {
    var header = FIELDS[key];
    if (!header) continue;
    var col = map[header];
    if (!col) continue;
    if (skipFormula && sh.getRange(row, col).getFormula() !== '') continue;
    colVal[col] = coerce(key, fields[key], tz);
  }

  var cols = Object.keys(colVal).map(Number).sort(function (a, b) { return a - b; });
  var i = 0;
  while (i < cols.length) {
    var j = i;
    while (j + 1 < cols.length && cols[j + 1] === cols[j] + 1) j++;
    var run = [];
    for (var k = i; k <= j; k++) run.push(colVal[cols[k]]);
    sh.getRange(row, cols[i], 1, run.length).setValues([run]);
    i = j + 1;
  }
}

function coerce(key, v, tz) {
  if (v === null || v === undefined) return '';
  var s = String(v).trim();
  if (s === '') return '';

  if (DATE_KEYS[key]) {
    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return s;
    // new Date(y,m,d) は「Apps Scriptプロジェクトの既定タイムゾーン」で真夜中を作ってしまい、
    // スプレッドシート自体のタイムゾーンと食い違っていると、表示上1日ずれて（前日に繰り上がって）しまう。
    // 必ずスプレッドシートのタイムゾーンを基準にして真夜中を作る。
    return Utilities.parseDate(s, tz || Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (NUM_KEYS[key]) {
    var n = Number(s.replace(/[,\s￥¥]/g, ''));
    return isNaN(n) ? s : n;
  }
  if (BOOL_KEYS[key]) return s.toUpperCase() === 'TRUE';
  if (ID_KEYS[key])   return /^\d+$/.test(s) ? Number(s) : s;
  return s;
}

// ── マスタへの追加 ──────────────────────────────────

/**
 * マスタシートの末尾に1行追加する共通処理。
 * fieldsMap: キー→見出し文字, values: そのopで書き込む値一式（IDは呼び出し側で確定済み）
 */
function doCreateMaster(sh, fieldsMap, op, values) {
  if (!sh) throw new Error('対象のマスタシートが見つかりません');
  var map = headerMap(sh);
  var row = sh.getLastRow() + 1;

  var colVal = {};
  for (var key in values) {
    var header = fieldsMap[key];
    if (!header) continue;
    var col = map[header];
    if (!col) continue;
    colVal[col] = values[key];
  }

  var cols = Object.keys(colVal).map(Number).sort(function (a, b) { return a - b; });
  var i = 0;
  while (i < cols.length) {
    var j = i;
    while (j + 1 < cols.length && cols[j + 1] === cols[j] + 1) j++;
    var run = [];
    for (var k = i; k <= j; k++) run.push(colVal[cols[k]]);
    sh.getRange(row, cols[i], 1, run.length).setValues([run]);
    i = j + 1;
  }
  return { opId: op.opId, ok: true, id: values.id, row: row };
}

function companyDefaults(op) {
  var f = op.fields || {};
  var id = String(op.id || '').trim();
  if (!id) throw new Error('会社IDがありません');
  if (!f.name) throw new Error('会社名がありません');
  return {
    id: id,
    name: f.name,
    short: f.short || f.name,
    address: f.address || '',
    tel: f.tel || '',
    closing: f.closing || '末日',
    span: f.span || '翌月'
  };
}

function repDefaults(op, tz) {
  var f = op.fields || {};
  var id = String(op.id || '').trim();
  if (!id) throw new Error('担当者IDがありません');
  if (!f.name) throw new Error('氏名がありません');
  if (!f.companyId) throw new Error('取引会社IDがありません');
  return {
    id: id,
    sort: f.sort || '',
    companyId: /^\d+$/.test(String(f.companyId)) ? Number(f.companyId) : f.companyId,
    name: f.name,
    title: f.title || '',
    updatedAt: Utilities.formatDate(new Date(), tz, 'yyyy/M/d')
  };
}

function clientDefaults(op) {
  var f = op.fields || {};
  var id = String(op.id || '').trim();
  if (!id) throw new Error('販売先IDがありません');
  if (!f.name) throw new Error('販売先名がありません');
  if (!f.companyId) throw new Error('取引会社IDがありません');
  return {
    id: id,
    name: f.name,
    companyId: /^\d+$/.test(String(f.companyId)) ? Number(f.companyId) : f.companyId,
    repId: f.repId ? (/^\d+$/.test(String(f.repId)) ? Number(f.repId) : f.repId) : ''
  };
}
