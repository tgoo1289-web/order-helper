/**
 * 매장 발주 기록 – 발주도우미 → 공유 시트 저장 / 공지 공유
 *
 * POST  { date: '2026-11-22', rows: [ {vendor:'여주농산', name:'레몬', qty:1, unit:'박스'}, ... ] }
 *   → 요청한 날짜 탭(M/D (요일))이 없으면 가장 최근 날짜 탭을 복제해
 *     수량(D)·단위(E) 칸만 비운 새 날짜 탭을 만들고,
 *     각 품목명을 B열에서 찾아 D열에 수량, E열에 단위를 써 넣는다.
 * POST  { action: 'add_notice', text: '내용', author: '이름(선택)' }
 *   → '공지' 탭에 한 줄 추가 (없으면 자동 생성)
 * POST  { action: 'delete_notice', id: '<공지 id>' }
 *   → '공지' 탭에서 해당 줄 삭제
 * GET   → 배포 상태 확인
 * GET   ?action=notices → 공지 목록(최신순) + 배포 상태
 *
 * 작업(할일) 공유 — '작업' 탭 (없으면 자동 생성)
 * GET   ?action=tasks → { ok, version, tasks: [{id, name, qty, unit, checked, updated}] }
 * POST  { action: 'save_tasks', tasks: [...] }   → 목록 전체 교체 (최초 이관·병합용)
 * POST  { action: 'add_task', task: { id, name, qty, unit, checked } }
 * POST  { action: 'update_task', id, name?, qty?, unit?, checked? }
 * POST  { action: 'delete_task', id }
 * POST  { action: 'reset_tasks' }                → 모든 작업의 체크·작업량만 비움 (목록·단위 유지)
 */
var TIMEZONE = 'Asia/Seoul';
var DAYS = ['일', '월', '화', '수', '목', '금', '토'];
var QTY_COL = 4;   // D열 수량
var UNIT_COL = 5;  // E열 단위
var EXTRA_LABEL = '추가 발주';

var NOTICE_SHEET = '공지';
var NOTICE_HEADER = ['id', '시각', '내용', '작성자'];
var NOTICE_MAX_LEN = 1000;

var TASK_SHEET = '작업';
var TASK_HEADER = ['id', '작업명', '작업량', '단위', '체크', '갱신'];
var TASK_NAME_MAX = 60;

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || '';
    if (action === 'notices' || action === 'list_notices') {
      return out_({ ok: true, service: 'order-helper', version: 4, notices: readNotices_() });
    }
    if (action === 'tasks' || action === 'list_tasks' || action === 'getTasks') {
      return out_({ ok: true, service: 'order-helper', version: 4, tasks: readTasks_() });
    }
    return out_({ ok: true, service: 'order-helper', version: 4 });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var action = String(body.action || '');
    if (action === 'add_notice') return out_(addNotice_(body));
    if (action === 'delete_notice') return out_(deleteNotice_(body));
    if (action === 'save_tasks' || action === 'saveTasks') return out_(saveTasksAll_(body));
    if (action === 'add_task') return out_(addTask_(body));
    if (action === 'update_task') return out_(updateTask_(body));
    if (action === 'delete_task') return out_(deleteTask_(body));
    if (action === 'reset_tasks') return out_(resetTasks_(body));
    return out_(saveRows_(body));
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

/* ── 발주/재고 수량을 날짜 탭에 저장 ── */
function saveRows_(body) {
  var rows = Array.isArray(body.rows) ? body.rows : [];
  if (!rows.length) return { ok: false, error: 'rows 없음' };

  var now = new Date();
  var date = parseDate_(body.date, now);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var created = false;
  var sheet = ss.getSheetByName(sheetName_(date));
  if (!sheet) {
    sheet = createDaySheet_(ss, date);
    created = true;
  }

  var idx = nameIndex_(sheet, 1, sheet.getLastRow(), 2);
  var filled = 0;
  var missing = [];

  rows.forEach(function (r) {
    var name = String((r && r.name) || '').trim();
    if (!name) return;
    var qty = Number(r.qty);
    if (!isFinite(qty)) qty = '';
    var unit = String((r && r.unit) || '');
    var row = idx[name];
    if (row) {
      sheet.getRange(row, QTY_COL).setValue(qty);
      sheet.getRange(row, UNIT_COL).setValue(unit);
      filled++;
    } else {
      missing.push({ vendor: String((r && r.vendor) || ''), name: name, qty: qty, unit: unit });
    }
  });

  if (missing.length) appendExtras_(sheet, missing);
  SpreadsheetApp.flush();

  return {
    ok: true,
    sheet: sheet.getName(),
    added: filled + missing.length,
    filled: filled,
    addedExtra: missing.length,
    created: created,
    notFound: missing.map(function (m) { return m.name; })
  };
}

/* ── 공지: 탭 확보(없으면 자동 생성) ── */
function noticeSheet_(create) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(NOTICE_SHEET);
  if (!sh && create) {
    sh = ss.insertSheet(NOTICE_SHEET, 0);
    var head = sh.getRange(1, 1, 1, NOTICE_HEADER.length);
    head.setValues([NOTICE_HEADER])
      .setFontWeight('bold')
      .setBackground('#d1d5db')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160);
    sh.setColumnWidth(2, 130);
    sh.setColumnWidth(3, 520);
    sh.setColumnWidth(4, 110);
    sh.getRange(1, 3).setHorizontalAlignment('left');
    sh.getRange(2, 3, sh.getMaxRows() - 1, 1).setWrap(true);
  }
  return sh;
}

function addNotice_(body) {
  var text = String(body.text || body.content || '').trim();
  if (!text) return { ok: false, error: '내용 없음' };
  if (text.length > NOTICE_MAX_LEN) text = text.slice(0, NOTICE_MAX_LEN);
  var author = String(body.author || '').trim().slice(0, 20);

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = noticeSheet_(true);
    var now = new Date();
    var id = String(now.getTime()) + '-' + Math.floor(Math.random() * 10000);
    var row = Math.max(sh.getLastRow() + 1, 2);
    sh.getRange(row, 1, 1, 4).setNumberFormat('@');
    sh.getRange(row, 1).setValue(id);
    sh.getRange(row, 2).setValue(formatStamp_(now));
    // '=' 로 시작하는 글도 수식이 아니라 글자로 저장
    sh.getRange(row, 3).setValue(/^[=+\-@]/.test(text) ? "'" + text : text);
    sh.getRange(row, 4).setValue(author);
    sh.getRange(row, 1, 1, 4).setVerticalAlignment('top');
    sh.setRowHeight(row, 24);
    SpreadsheetApp.flush();
    return { ok: true, notice: { id: id, time: formatStamp_(now), text: text, author: author } };
  } finally {
    lock.releaseLock();
  }
}

function deleteNotice_(body) {
  var id = String(body.id || '').trim();
  if (!id) return { ok: false, error: 'id 없음' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = noticeSheet_(false);
    if (!sh) return { ok: false, error: '공지 탭 없음' };
    var last = sh.getLastRow();
    if (last < 2) return { ok: false, error: '삭제할 공지 없음' };
    var ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === id) {
        sh.deleteRow(i + 2);
        SpreadsheetApp.flush();
        return { ok: true, deleted: id };
      }
    }
    return { ok: false, error: '해당 공지 없음' };
  } finally {
    lock.releaseLock();
  }
}

function readNotices_() {
  var sh = noticeSheet_(false);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, 4).getDisplayValues();
  var list = [];
  for (var i = 0; i < vals.length; i++) {
    var id = String(vals[i][0]).trim();
    var text = String(vals[i][2]).trim();
    if (!id || !text) continue;
    list.push({
      id: id,
      time: String(vals[i][1]).trim(),
      text: text,
      author: String(vals[i][3]).trim()
    });
  }
  // 최신순 (id 앞자리가 등록 시각)
  list.sort(function (a, b) {
    return (parseFloat(b.id) || 0) - (parseFloat(a.id) || 0);
  });
  return list;
}

/* ── 작업(할일): 탭 확보(없으면 자동 생성) ── */
function taskSheet_(create) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(TASK_SHEET);
  if (!sh && create) {
    sh = ss.insertSheet(TASK_SHEET, 0);
    var head = sh.getRange(1, 1, 1, TASK_HEADER.length);
    head.setValues([TASK_HEADER])
      .setFontWeight('bold')
      .setBackground('#d1d5db')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 170);
    sh.setColumnWidth(2, 260);
    sh.setColumnWidth(3, 76);
    sh.setColumnWidth(4, 66);
    sh.setColumnWidth(5, 56);
    sh.setColumnWidth(6, 132);
    sh.getRange(1, 1).setNumberFormat('@');
  }
  return sh;
}

function readTasks_(sh) {
  sh = sh || taskSheet_(false);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, 6).getValues();
  var list = [];
  for (var i = 0; i < vals.length; i++) {
    var name = String(vals[i][1] == null ? '' : vals[i][1]).trim();
    if (!name) continue;
    var id = String(vals[i][0] == null ? '' : vals[i][0]).trim() || ('row_' + (i + 2));
    var qty = Number(vals[i][2]);
    if (!isFinite(qty) || qty < 0) qty = 0;
    list.push({
      id: id,
      name: name,
      qty: qty,
      unit: String(vals[i][3] == null ? '' : vals[i][3]).trim(),
      checked: toBool_(vals[i][4]),
      updated: String(vals[i][5] == null ? '' : vals[i][5]).trim()
    });
  }
  return list;
}

function toBool_(v) {
  if (v === true) return true;
  var s = String(v == null ? '' : v).trim().toUpperCase();
  return s === 'TRUE' || s === '1' || s === 'O' || s === 'Y' || s === '✓';
}

function taskRow_(t, stamp) {
  return [
    String((t && t.id) || ''),
    String((t && t.name) || '').trim().slice(0, TASK_NAME_MAX),
    Math.max(0, Number(t && t.qty) || 0),
    String((t && t.unit) || ''),
    (t && t.checked) ? 'O' : '',
    stamp || formatStamp_(new Date())
  ];
}

function newTaskId_() {
  return String(new Date().getTime()) + '-' + Math.floor(Math.random() * 10000);
}

/* 목록 전체 교체 — 이 폰에만 있던 작업을 처음 올리거나 합칠 때 */
function saveTasksAll_(body) {
  var tasks = Array.isArray(body.tasks) ? body.tasks : [];
  if (!tasks.length) return { ok: false, error: 'tasks 없음' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = taskSheet_(true);
    var stamp = formatStamp_(new Date());
    var vals = [];
    var seen = {};
    tasks.forEach(function (t) {
      var name = String((t && t.name) || '').trim().slice(0, TASK_NAME_MAX);
      if (!name || seen[name]) return;
      seen[name] = true;
      vals.push(taskRow_({
        id: String((t && t.id) || '').trim() || newTaskId_(),
        name: name,
        qty: t && t.qty,
        unit: t && t.unit,
        checked: t && t.checked
      }, stamp));
    });
    if (!vals.length) return { ok: false, error: '저장할 작업 없음' };

    var last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, 6).clearContent();
    sh.getRange(2, 1, vals.length, 6).setValues(vals);
    sh.getRange(2, 5, Math.max(vals.length, 1), 1).setNumberFormat('@');
    SpreadsheetApp.flush();
    return { ok: true, count: vals.length, tasks: readTasks_(sh) };
  } finally {
    lock.releaseLock();
  }
}

/* 작업 한 건 추가 (이미 같은 이름이 있으면 그것을 돌려준다) */
function addTask_(body) {
  var t = body.task || body;
  var name = String((t && t.name) || '').trim().slice(0, TASK_NAME_MAX);
  if (!name) return { ok: false, error: '작업 이름 없음' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = taskSheet_(true);
    var list = readTasks_(sh);
    for (var i = 0; i < list.length; i++) {
      if (list[i].name === name) return { ok: true, duplicate: true, task: list[i] };
    }
    var id = String((t && t.id) || '').trim() || newTaskId_();
    var row = Math.max(sh.getLastRow() + 1, 2);
    var vals = taskRow_({ id: id, name: name, qty: t && t.qty, unit: t && t.unit, checked: t && t.checked });
    sh.getRange(row, 1, 1, 6).setValues([vals]);
    sh.getRange(row, 5).setNumberFormat('@');
    SpreadsheetApp.flush();
    return {
      ok: true,
      task: { id: id, name: name, qty: Number(vals[2]) || 0, unit: vals[3], checked: !!vals[4], updated: vals[5] }
    };
  } finally {
    lock.releaseLock();
  }
}

/* 작업 한 건 수정 — 이름·작업량·단위·체크 중 넘어온 값만 바꾼다 */
function updateTask_(body) {
  var id = String(body.id || '').trim();
  if (!id) return { ok: false, error: 'id 없음' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = taskSheet_(true);
    var last = sh.getLastRow();
    if (last < 2) return { ok: false, error: '해당 작업 없음', missing: true };
    var ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
    var row = -1;
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === id) { row = i + 2; break; }
    }
    if (row < 0) return { ok: false, error: '해당 작업 없음', missing: true };

    var cur = sh.getRange(row, 1, 1, 6).getValues()[0];
    var name = String(cur[1] == null ? '' : cur[1]).trim();
    var qty = Math.max(0, Number(cur[2]) || 0);
    var unit = String(cur[3] == null ? '' : cur[3]);
    var checked = toBool_(cur[4]);

    if (body.name !== undefined) {
      var nextName = String(body.name || '').trim().slice(0, TASK_NAME_MAX);
      if (nextName && nextName !== name) {
        var list = readTasks_(sh);
        for (var j = 0; j < list.length; j++) {
          if (list[j].id !== id && list[j].name === nextName) {
            return { ok: false, error: '이미 같은 이름의 작업이 있음', duplicate: true };
          }
        }
        name = nextName;
      }
    }
    if (body.qty !== undefined) qty = Math.max(0, Number(body.qty) || 0);
    if (body.unit !== undefined) unit = String(body.unit || '');
    if (body.checked !== undefined) checked = !!body.checked;

    var vals = taskRow_({ id: id, name: name, qty: qty, unit: unit, checked: checked });
    sh.getRange(row, 1, 1, 6).setValues([vals]);
    sh.getRange(row, 5).setNumberFormat('@');
    SpreadsheetApp.flush();
    return {
      ok: true,
      task: { id: id, name: name, qty: qty, unit: unit, checked: checked, updated: vals[5] }
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteTask_(body) {
  var id = String(body.id || '').trim();
  if (!id) return { ok: false, error: 'id 없음' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = taskSheet_(true);
    var last = sh.getLastRow();
    if (last < 2) return { ok: false, error: '삭제할 작업 없음', missing: true };
    var ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === id) {
        sh.deleteRow(i + 2);
        SpreadsheetApp.flush();
        return { ok: true, deleted: id };
      }
    }
    return { ok: false, error: '해당 작업 없음', missing: true };
  } finally {
    lock.releaseLock();
  }
}

/* 오늘치 체크·작업량만 비움 — 작업 목록과 단위는 그대로 */
function resetTasks_(body) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = taskSheet_(true);
    var last = sh.getLastRow();
    if (last < 2) return { ok: true, count: 0 };
    var vals = sh.getRange(2, 1, last - 1, 6).getValues();
    var stamp = formatStamp_(new Date());
    var count = 0;
    for (var i = 0; i < vals.length; i++) {
      if (!String(vals[i][1] == null ? '' : vals[i][1]).trim()) continue;
      vals[i][2] = 0;
      vals[i][4] = '';
      vals[i][5] = stamp;
      count++;
    }
    sh.getRange(2, 1, vals.length, 6).setValues(vals);
    sh.getRange(2, 5, vals.length, 1).setNumberFormat('@');
    SpreadsheetApp.flush();
    return { ok: true, count: count, tasks: readTasks_(sh) };
  } finally {
    lock.releaseLock();
  }
}

/* ── 새 날짜 탭: 가장 최근 날짜 탭을 복제하고 수량·단위만 비운다 ── */
function createDaySheet_(ss, date) {
  var template = pickTemplate_(ss);
  var sheet = template.copyTo(ss);
  sheet.setName(sheetName_(date));
  clearQuantities_(sheet);
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(ss.getNumSheets());
  return sheet;
}

function clearQuantities_(sheet) {
  var last = sheet.getLastRow();
  if (last < 1) return;
  var vals = sheet.getRange(1, 1, last, 6).getValues();
  var extraRow = -1;
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === EXTRA_LABEL) extraRow = i + 1;
  }
  var end = extraRow > 0 ? extraRow - 1 : last;
  for (var r = 1; r <= end; r++) {
    var b = String(vals[r - 1][1]).trim();
    if (!b) continue;                      // 빈 줄
    if (b.indexOf('합계') >= 0) continue;   // 합계 줄의 =SUM 수식은 그대로 둔다
    sheet.getRange(r, QTY_COL, 1, 2).clearContent();
  }
  if (extraRow > 0 && extraRow <= last) {
    sheet.deleteRows(extraRow, last - extraRow + 1);
  }
}

function pickTemplate_(ss) {
  var best = null;
  var bestScore = -1;
  ss.getSheets().forEach(function (s) {
    if (s.getName() === NOTICE_SHEET || s.getName() === TASK_SHEET) return;
    var md = monthDay_(s.getName());
    if (!md) return;
    var score = md.m * 100 + md.d;
    if (score > bestScore) { bestScore = score; best = s; }
  });
  if (best) return best;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() !== NOTICE_SHEET && sheets[i].getName() !== TASK_SHEET) return sheets[i];
  }
  return sheets[0];
}

/* ── 표에 없는 신규 품목은 맨 아래 '추가 발주' 칸에 덧붙인다 ── */
function appendExtras_(sheet, missing) {
  var last = Math.max(sheet.getLastRow(), 1);
  var colA = sheet.getRange(1, 1, last, 1).getValues();
  var labelRow = -1;
  for (var i = 0; i < colA.length; i++) {
    if (String(colA[i][0]).trim() === EXTRA_LABEL) { labelRow = i + 1; break; }
  }
  if (labelRow < 0) {
    labelRow = last + 1;
    var band = sheet.getRange(labelRow, 1, 1, 6);
    band.breakApart();
    band.merge()
      .setValue(EXTRA_LABEL)
      .setBackground('#d1d5db')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(labelRow, 27);
    sheet.getRange(labelRow, 1, 1, 6)
      .setBorder(true, true, true, true, true, true, '#1a1a1a', SpreadsheetApp.BorderStyle.SOLID);
  }

  var startRow = labelRow + 1;
  var span = Math.max(sheet.getLastRow() - startRow + 1, 0);
  var idx = {};
  if (span > 0) idx = nameIndex_(sheet, startRow, sheet.getLastRow(), 2);

  var cur = startRow;
  missing.forEach(function (m) {
    var row = idx[m.name];
    if (!row) {
      row = cur;
      while (true) {
        var taken = String(sheet.getRange(row, 2).getValue()).trim();
        if (!taken) break;
        row++;
      }
      cur = row + 1;
      idx[m.name] = row;
    }
    sheet.getRange(row, 2).setValue(m.name);
    sheet.getRange(row, 3).setValue(m.vendor);
    sheet.getRange(row, QTY_COL).setValue(m.qty);
    sheet.getRange(row, UNIT_COL).setValue(m.unit);
    sheet.setRowHeight(row, 27);
    sheet.getRange(row, 1, 1, 6)
      .setBorder(true, true, true, true, true, true, '#1a1a1a', SpreadsheetApp.BorderStyle.SOLID);
  });
}

/* ── 도우미 ── */
function nameIndex_(sheet, firstRow, lastRow, col) {
  var idx = {};
  var n = lastRow - firstRow + 1;
  if (n < 1) return idx;
  var vals = sheet.getRange(firstRow, col, n, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    var name = String(vals[i][0]).trim();
    if (name && !idx[name]) idx[name] = firstRow + i;
  }
  return idx;
}

function monthDay_(name) {
  var m = /^(\d{1,2})\s*\/\s*(\d{1,2})/.exec(String(name));
  return m ? { m: Number(m[1]), d: Number(m[2]) } : null;
}

function sheetName_(d) {
  return (d.getMonth() + 1) + '/' + d.getDate() + ' (' + DAYS[d.getDay()] + ')';
}

function parseDate_(value, fallback) {
  var s = String(value || '');
  var m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return fallback || new Date();
}

function formatStamp_(d) {
  var md = Utilities.formatDate(d, TIMEZONE, 'M/d');
  var hm = Utilities.formatDate(d, TIMEZONE, 'HH:mm');
  var iso = Number(Utilities.formatDate(d, TIMEZONE, 'u')); // 1=월 ... 7=일
  return md + ' (' + DAYS[iso % 7] + ') ' + hm;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
