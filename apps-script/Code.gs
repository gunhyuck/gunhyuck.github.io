/**
 * ============================================================
 *  피클 합동 비서실 DB — Google Apps Script (웹 앱)
 *  카림애비뉴 합동 비서실 / 스타타워 비서실 두 페이지가 함께 사용합니다.
 *
 *  설치
 *   1. 스프레드시트(피클 합동 비서실 DB) → 확장 프로그램 → Apps Script
 *   2. 이 파일 전체를 Code.gs 에 붙여넣고 저장
 *   3. 배포 → 새 배포 → 유형: 웹 앱
 *        실행 계정: 나 / 액세스 권한: 모든 사용자
 *   4. 배포 후 나오는 .../exec 주소를 두 페이지의 ⚙ 설정에 붙여넣기
 *
 *  API
 *   GET  ?action=list&site=karim
 *   GET  ?action=ping
 *   GET  ?action=addTask|updateStatus|addMemo&data=<JSON>     (JSONP 쓰기용)
 *   POST (Content-Type: text/plain)  body = {"action":"addTask", "site":"karim", ...}
 *   GET/POST 모두 &callback=함수명 을 주면 JSONP 로 감싸서 돌려줍니다.
 * ============================================================
 */

var SPREADSHEET_ID = '1bWeBT0SEThBbFqP_87QkqgclWGYzk62wtM5iRYibu6k';

var SITES = {
  karim:     { tasks: 'karim_tasks',     memos: 'karim_memos',     prefix: 'K' },
  startower: { tasks: 'startower_tasks', memos: 'startower_memos', prefix: 'S' }
};

var TASK_HEADERS = ['id', 'title', 'category', 'owner', 'status', 'due', 'summary', 'updated_by', 'updated_at'];
var MEMO_HEADERS = ['id', 'author', 'role', 'content', 'created_at'];
var VALID_STATUS = ['대기', '진행중', '검토중', '완료'];
var TZ = 'Asia/Seoul';

/* ---------------- 진입점 ---------------- */

function doGet(e) {
  return handle(e);
}

function doPost(e) {
  var body = {};
  try {
    if (e && e.postData && e.postData.contents) body = JSON.parse(e.postData.contents);
  } catch (err) {
    body = {};
  }
  return handle(e, body);
}

function handle(e, body) {
  var p = {};
  var q = (e && e.parameter) || {};
  for (var k in q) if (q.hasOwnProperty(k)) p[k] = q[k];
  if (p.data) {
    try {
      var parsed = JSON.parse(p.data);
      for (var k2 in parsed) if (parsed.hasOwnProperty(k2)) p[k2] = parsed[k2];
    } catch (err) { /* data 파라미터가 JSON 이 아니면 무시 */ }
  }
  if (body && typeof body === 'object') {
    for (var k3 in body) if (body.hasOwnProperty(k3)) p[k3] = body[k3];
  }

  var action = String(p.action || 'list');
  var site = resolveSite(p);
  var out;
  try {
    out = dispatch(action, site, p);
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  out.action = action;
  out.site = site;
  return output(out, p.callback);
}

function dispatch(action, site, p) {
  switch (action) {
    case 'ping':
      return {
        ok: true,
        message: '피클 합동 비서실 API 연결됨',
        site: site,
        tasks: countRows(SITES[site].tasks),
        memos: countRows(SITES[site].memos)
      };
    case 'list':
      return listAll(site);
    case 'addTask':
      return addTask(site, p.task || p);
    case 'updateStatus':
      return updateStatus(site, p.id || (p.task && p.task.id), p.status, p.updated_by);
    case 'addMemo':
      return addMemo(site, p.memo || p);
    default:
      throw new Error('알 수 없는 action: ' + action);
  }
}

function resolveSite(p) {
  var s = String(p.site || '').toLowerCase();
  if (SITES[s]) return s;
  var id = String(p.id || (p.task && p.task.id) || '');
  if (/^S-/i.test(id)) return 'startower';
  if (/^K-/i.test(id)) return 'karim';
  return 'karim';
}

function output(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback && /^[A-Za-z_$][A-Za-z0-9_$.]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- 읽기 ---------------- */

function ss() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet(name, headers) {
  var book = ss();
  var sheet = book.getSheetByName(name);
  if (!sheet) sheet = book.insertSheet(name);
  if (headers && sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function readObjects(sheetName, headers, dateFmt) {
  var sheet = getSheet(sheetName, headers);
  var fmt = dateFmt || 'yyyy-MM-dd HH:mm';
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var width = Math.max(sheet.getLastColumn(), headers.length);
  var values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var obj = {};
    var empty = true;
    for (var c = 0; c < headers.length; c++) {
      var v = row[c];
      if (v instanceof Date) v = Utilities.formatDate(v, TZ, fmt);
      obj[headers[c]] = v === null || v === undefined ? '' : String(v).trim();
      if (obj[headers[c]] !== '') empty = false;
    }
    if (!empty) out.push(obj);
  }
  return out;
}

function countRows(sheetName) {
  try {
    var sheet = ss().getSheetByName(sheetName);
    if (!sheet) return 0;
    return Math.max(0, sheet.getLastRow() - 1);
  } catch (err) {
    return 0;
  }
}

function listAll(site) {
  var cfg = SITES[site];
  return {
    ok: true,
    tasks: readObjects(cfg.tasks, TASK_HEADERS, 'yyyy-MM-dd'),
    memos: readObjects(cfg.memos, MEMO_HEADERS, 'yyyy-MM-dd HH:mm'),
    serverTime: nowStamp()
  };
}

/* ---------------- 쓰기 ---------------- */

function withLock(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('다른 저장 작업이 진행 중입니다. 잠시 후 다시 시도해 주세요.');
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function nextTaskId(site, sheet) {
  var prefix = SITES[site].prefix;
  var lastRow = sheet.getLastRow();
  var max = 0;
  if (lastRow >= 2) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      var m = String(ids[i][0]).match(/(\d+)\s*$/);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    }
  }
  return prefix + '-' + (max + 1);
}

function addTask(site, t) {
  t = t || {};
  var title = String(t.title || '').trim();
  if (!title) throw new Error('안건 제목이 없습니다.');

  return withLock(function () {
    var cfg = SITES[site];
    var sheet = getSheet(cfg.tasks, TASK_HEADERS);

    var requested = String(t.id || '').trim();
    var id = requested;
    if (requested) {
      var row = findRow(sheet, requested);
      if (row > 0) {
        // 같은 id + 같은 제목 = 응답을 못 받고 재전송된 경우 → 중복 저장하지 않음
        var existingTitle = String(sheet.getRange(row, 2).getValue()).trim();
        if (existingTitle === title) return { ok: true, id: requested, existed: true };
        // id 만 겹치고 다른 안건이면 새 번호 부여
        id = nextTaskId(site, sheet);
      }
    } else {
      id = nextTaskId(site, sheet);
    }

    var status = String(t.status || '대기').trim();
    if (VALID_STATUS.indexOf(status) < 0) status = '대기';

    var row = [
      id,
      title,
      String(t.category || '').trim(),
      String(t.owner || '').trim(),
      status,
      String(t.due || '-').trim() || '-',
      String(t.summary || '').trim(),
      String(t.updated_by || '웹 비서실').trim(),
      String(t.updated_at || nowStamp().slice(0, 10)).trim()
    ];
    sheet.appendRow(row);
    return { ok: true, id: id, task: rowToTask(row) };
  });
}

function updateStatus(site, id, status, updatedBy) {
  id = String(id || '').trim();
  status = String(status || '').trim();
  if (!id) throw new Error('변경할 안건 id 가 없습니다.');
  if (VALID_STATUS.indexOf(status) < 0) throw new Error('상태 값이 올바르지 않습니다: ' + status);

  return withLock(function () {
    var sheet = getSheet(SITES[site].tasks, TASK_HEADERS);
    var row = findRow(sheet, id);
    if (row < 1) throw new Error('안건을 찾을 수 없습니다: ' + id);
    sheet.getRange(row, 5).setValue(status);                                  // status
    sheet.getRange(row, 8).setValue(String(updatedBy || '웹 비서실'));          // updated_by
    sheet.getRange(row, 9).setValue(nowStamp().slice(0, 10));                 // updated_at
    return { ok: true, id: id, status: status, row: row };
  });
}

function addMemo(site, m) {
  m = m || {};
  var content = String(m.content || '').trim();
  var author = String(m.author || '').trim();
  if (!content) throw new Error('메모 내용이 없습니다.');
  if (!author) author = '익명';

  return withLock(function () {
    var cfg = SITES[site];
    var sheet = getSheet(cfg.memos, MEMO_HEADERS);
    var stamp = nowStamp();
    var requested = String(m.id || '').trim();
    var id = requested || ('M-' + cfg.prefix + Date.now().toString(36));
    if (requested) {
      var found = findRow(sheet, requested);
      if (found > 0) {
        var existingText = String(sheet.getRange(found, 4).getValue()).trim();
        if (existingText === content) return { ok: true, id: requested, existed: true };
        id = 'M-' + cfg.prefix + Date.now().toString(36);
      }
    }
    var row = [
      id,
      author,
      String(m.role || '').trim(),
      content,
      String(m.created_at || stamp).trim()
    ];
    sheet.appendRow(row);
    return { ok: true, id: id, memo: rowToMemo(row) };
  });
}

/* ---------------- 도우미 ---------------- */

function findRow(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === id) return i + 2;
  }
  return -1;
}

function rowToTask(row) {
  var o = {};
  for (var i = 0; i < TASK_HEADERS.length; i++) o[TASK_HEADERS[i]] = row[i] === null || row[i] === undefined ? '' : String(row[i]);
  return o;
}

function rowToMemo(row) {
  var o = {};
  for (var i = 0; i < MEMO_HEADERS.length; i++) o[MEMO_HEADERS[i]] = row[i] === null || row[i] === undefined ? '' : String(row[i]);
  return o;
}

function nowStamp() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
}

/* ---------------- 점검용 (편집기에서 직접 실행) ---------------- */

function testList() {
  var r = listAll('karim');
  Logger.log('karim tasks=%s memos=%s', r.tasks.length, r.memos.length);
  return r;
}

function testAddMemo() {
  var r = addMemo('karim', { author: '테스트', role: '점검', content: '연결 점검용 메모입니다. 시트에서 직접 지우시면 됩니다.' });
  Logger.log(JSON.stringify(r));
  return r;
}
