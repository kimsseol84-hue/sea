/**
 * 학부모 상담 예약 웹앱 (Google Apps Script + 구글 시트)
 *
 * 원 전체가 시트 하나, 링크 하나로 함께 씁니다.
 *  - 설정     : 페이지 제목, 안내 문구, 신청 받기, 명단 확인, 전체 관리자 코드
 *  - 반       : 반 이름 | 담임 | 상담 시간(분) | 반 관리자 코드
 *  - 일정     : 반 | 날짜 | 시간(쉼표로 구분)
 *  - 명단     : 반 | 아이 이름   (명단에 있는 아이만 신청 가능)
 *  - 예약현황 : 신청 내역 (관리자만 보는 곳)
 *  - 보관     : 지난 상담 기간의 예약 (메뉴 > 지난 예약 보관하기)
 *
 * 선생님들은 이 스프레드시트를 "편집자"로 공유받아 자기 반 줄을 고칩니다.
 * 학부모에게는 웹앱 링크만 보냅니다. 학부모는 다른 사람 이름을 볼 수 없습니다.
 */

var SHEET = { CONFIG: '설정', CLASSES: '반', SLOTS: '일정', ROSTER: '명단', BOOK: '예약현황', ARCHIVE: '보관' };
var BOOK_HEADERS = ['접수시각', '반', '날짜', '시간', '아이 이름', '보호자', '연락처', '요청사항', '확인번호', '상태'];
var COL = { created: 0, cls: 1, date: 2, time: 3, child: 4, guardian: 5, phone: 6, memo: 7, code: 8, status: 9 };
var ACTIVE = '예약';
var CANCELLED = '취소';
var WEEKDAYS = ['일', '월', '화', '수', '목', '금', '토'];

/* ───────── 메뉴 & 초기 설정 ───────── */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('상담예약')
    .addItem('처음 설정하기', 'setup')
    .addItem('지난 예약 보관하기 (새 상담 기간 시작)', 'archiveBookings')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('Asia/Seoul');
  var tz = ss.getSpreadsheetTimeZone();

  if (!ss.getSheetByName(SHEET.CONFIG)) {
    var rows = [
      ['항목', '값', '설명'],
      ['제목', '2학기 학부모 상담 신청', '페이지 맨 위에 보이는 제목'],
      ['안내 문구', '반을 고른 뒤 원하시는 날짜와 시간을 골라 주세요. 한 아이당 한 번 신청할 수 있어요.', ''],
      ['예약 받기', '예', '"아니오"로 바꾸면 모든 반의 신청이 멈춥니다'],
      ['명단 확인', '예', '"예"이면 "명단" 시트에 있는 아이 이름으로만 신청할 수 있어요 (그 반 명단이 비어 있으면 확인하지 않음)'],
      ['연락처 받기', '아니오', '예 / 아니오 — 꼭 필요할 때만 "예"로 바꾸세요'],
      ['전체 관리자 코드', randomDigits_(6), '모든 반의 명단을 볼 수 있는 코드. 원장님·행정 담당만 아세요']
    ];
    var c = ss.insertSheet(SHEET.CONFIG, 0);
    c.getRange(1, 1, rows.length, 3).setNumberFormat('@').setValues(rows);
    styleHeader_(c, 3, [130, 380, 420]);
  }

  if (!ss.getSheetByName(SHEET.CLASSES)) {
    var cl = ss.insertSheet(SHEET.CLASSES, 1);
    cl.getRange(1, 1, 3, 4).setNumberFormat('@').setValues([
      ['반 이름', '담임', '상담 시간(분)', '반 관리자 코드'],
      ['햇살반', '김OO 선생님', '20', randomDigits_(6)],
      ['꽃잎반', '이OO 선생님', '20', randomDigits_(6)]
    ]);
    styleHeader_(cl, 4, [120, 160, 110, 140]);
  }

  var slots = ss.getSheetByName(SHEET.SLOTS);
  if (!slots) {
    slots = ss.insertSheet(SHEET.SLOTS, 2);
    var sample = [['반', '날짜', '시간 (쉼표로 구분)']];
    var d = new Date();
    var made = 0;
    while (made < 3) {
      d.setDate(d.getDate() + 1);
      if (Number(Utilities.formatDate(d, tz, 'u')) > 5) continue; // 주말 건너뜀
      var day = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      sample.push(['햇살반', day, '14:00, 14:30, 15:00, 15:30']);
      sample.push(['꽃잎반', day, '16:00, 16:30']);
      made++;
    }
    slots.getRange(1, 1, sample.length, 3).setNumberFormat('@').setValues(sample);
    styleHeader_(slots, 3, [120, 130, 360]);
  } else if (String(slots.getRange(1, 1).getValue()).trim() === '날짜') {
    // 이전 버전(반 열 없음)에서 넘어온 경우
    slots.insertColumnBefore(1);
    slots.getRange(1, 1).setValue('반').setFontWeight('bold');
  }

  if (!ss.getSheetByName(SHEET.ROSTER)) {
    var r = ss.insertSheet(SHEET.ROSTER, 3);
    r.getRange(1, 1, 1, 2).setValues([['반', '아이 이름']]);
    styleHeader_(r, 2, [120, 160]);
    r.getRange('A2').setNote('예: 햇살반 | 김하늘\n한 줄에 한 명씩. 비워 두면 그 반은 이름 확인 없이 신청을 받아요.');
  }

  ensureBookSheet_(ss);
  if (!ss.getSheetByName(SHEET.ARCHIVE)) {
    var a = ss.insertSheet(SHEET.ARCHIVE);
    a.appendRow(BOOK_HEADERS.concat(['보관일']));
    a.getRange(1, 1, 1, BOOK_HEADERS.length + 1).setFontWeight('bold');
  }

  ss.getSheets().forEach(function (s) {
    if (/^(Sheet1|시트1)$/.test(s.getName()) && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  SpreadsheetApp.getUi().alert('설정 완료! "반", "일정", "명단" 시트를 원에 맞게 고친 뒤 웹앱으로 배포하세요.');
}

function styleHeader_(sheet, cols, widths) {
  sheet.getRange(1, 1, 1, cols).setFontWeight('bold');
  sheet.setFrozenRows(1);
  widths.forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });
}

function ensureBookSheet_(ss) {
  var b = ss.getSheetByName(SHEET.BOOK);
  if (!b) {
    b = ss.insertSheet(SHEET.BOOK, 4);
    b.appendRow(BOOK_HEADERS);
    b.getRange(1, 1, 1, BOOK_HEADERS.length).setFontWeight('bold');
    b.setFrozenRows(1);
  } else if (String(b.getRange(1, 2).getValue()).trim() === '날짜') {
    // 이전 버전(반 열 없음)에서 넘어온 경우
    b.insertColumnBefore(2);
    b.getRange(1, 2).setValue('반').setFontWeight('bold');
  }
  return b;
}

function archiveBookings() {
  var ui = SpreadsheetApp.getUi();
  var ok = ui.alert('지난 예약 보관', '"예약현황"의 모든 내역을 "보관" 시트로 옮기고 비웁니다. 계속할까요?', ui.ButtonSet.YES_NO);
  if (ok !== ui.Button.YES) return;

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.getActive();
    var b = ensureBookSheet_(ss);
    var a = ss.getSheetByName(SHEET.ARCHIVE);
    var n = b.getLastRow() - 1;
    if (n > 0) {
      var rows = b.getRange(2, 1, n, BOOK_HEADERS.length).getValues();
      var stamp = new Date();
      var out = rows.map(function (r) { return r.concat([stamp]); });
      a.getRange(a.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
      b.deleteRows(2, n);
    }
    ui.alert(n > 0 ? n + '건을 보관했어요. "일정" 시트에 새 날짜를 넣으면 바로 다시 쓸 수 있어요.' : '보관할 예약이 없어요.');
  } finally {
    lock.releaseLock();
  }
}

/* ───────── 웹앱 ───────── */

function doGet() {
  var cfg = readConfig_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(cfg.title || '학부모 상담 신청')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 학부모 화면에 필요한 공개 정보. 아이 이름·관리자 코드는 절대 포함하지 않는다. */
function getPublicData() {
  var cfg = readConfig_();
  var classes = readClasses_();
  var days = buildDays_(classes, activeBookings_(classes));
  return {
    title: cfg.title,
    notice: cfg.notice,
    askPhone: cfg.askPhone,
    open: cfg.open,
    classes: classes.map(function (c) {
      return { name: c.name, teacher: c.teacher, minutes: c.minutes, days: days[c.name] || [] };
    })
  };
}

function book(form) {
  var cls = cleanText_(form && form.cls, 30);
  var child = cleanText_(form && form.child, 20);
  var guardian = cleanText_(form && form.guardian, 20);
  var phone = cleanText_(form && form.phone, 20);
  var memo = cleanText_(form && form.memo, 200);
  var date = String(form && form.date || '');
  var time = String(form && form.time || '');
  if (!child) throw new Error('아이 이름을 적어 주세요.');
  if (!guardian) throw new Error('보호자(예: 엄마, 아빠)를 적어 주세요.');

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('신청이 몰리고 있어요. 잠시 뒤 다시 눌러 주세요.');
  try {
    var cfg = readConfig_();
    if (!cfg.open) throw new Error('지금은 신청을 받지 않고 있어요.');
    if (cfg.askPhone && !phone) throw new Error('연락처를 적어 주세요.');

    var classes = readClasses_();
    var klass = classes.filter(function (c) { return c.name === cls; })[0];
    if (!klass) throw new Error('반을 다시 골라 주세요. 화면을 새로고침하면 최신 반 목록이 보여요.');

    if (cfg.checkRoster) {
      var roster = readRoster_(classes)[cls];
      if (roster) {
        var canonical = roster[nameKey_(child)];
        if (!canonical) {
          Utilities.sleep(800);
          throw new Error(cls + ' 명단에서 "' + child + '" 이름을 찾지 못했어요. 아이 이름을 정확히 적어 주세요. 계속 안 되면 담임 선생님께 알려 주세요.');
        }
        child = canonical;
      }
    }

    var bookings = activeBookings_(classes);
    var day = (buildDays_(classes, bookings)[cls] || []).filter(function (d) { return d.date === date; })[0];
    var slot = day && day.times.filter(function (t) { return t.time === time; })[0];
    if (!slot) throw new Error('선택한 시간이 더 이상 없어요. 화면을 새로고침해 주세요.');
    if (slot.taken) throw new Error('방금 다른 분이 이 시간을 신청했어요. 다른 시간을 골라 주세요.');

    var key = nameKey_(child);
    var mine = bookings.filter(function (b) { return b.cls === cls && nameKey_(b.child) === key; })[0];
    if (mine) {
      throw new Error(child + ' 이름으로 이미 ' + dayLabel_(mine.date) + ' ' + mine.time +
        ' 예약이 있어요. 바꾸려면 아래 "내 예약 확인·취소"에서 먼저 취소해 주세요.');
    }

    var code = randomDigits_(4);
    var sheet = ensureBookSheet_(SpreadsheetApp.getActive());
    var row = [new Date(), cls, date, time, safeCell_(child), safeCell_(guardian), safeCell_(phone), safeCell_(memo), code, ACTIVE];
    var r = sheet.getLastRow() + 1;
    sheet.getRange(r, 2, 1, BOOK_HEADERS.length - 1).setNumberFormat('@');
    sheet.getRange(r, 1, 1, row.length).setValues([row]);
    SpreadsheetApp.flush();

    return { cls: cls, date: date, time: time, label: dayLabel_(date), code: code, child: child };
  } finally {
    lock.releaseLock();
  }
}

function findMyBooking(child, code) {
  var b = matchBooking_(child, code);
  return { cls: b.cls, date: b.date, time: b.time, label: dayLabel_(b.date), child: b.child };
}

function cancelMyBooking(child, code) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('잠시 뒤 다시 시도해 주세요.');
  try {
    var b = matchBooking_(child, code);
    ensureBookSheet_(SpreadsheetApp.getActive()).getRange(b.row, COL.status + 1).setValue(CANCELLED);
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 관리자 화면. 전체 관리자 코드면 모든 반, 반 관리자 코드면 그 반만 돌려준다.
 */
function getAdminData(adminCode) {
  var cfg = readConfig_();
  var classes = readClasses_();
  var code = String(adminCode || '').trim();
  var scope = null;
  if (code && code === cfg.adminCode) scope = '*';
  else if (code) classes.forEach(function (c) { if (c.code && c.code === code) scope = c.name; });
  if (!scope) {
    Utilities.sleep(800); // 무작위 대입을 느리게
    throw new Error('관리자 코드가 맞지 않아요.');
  }

  var inScope = function (name) { return scope === '*' || name === scope; };
  var list = activeBookings_(classes)
    .filter(function (b) { return inScope(b.cls); })
    .sort(function (a, b) { return (a.date + a.time + a.cls).localeCompare(b.date + b.time + b.cls); })
    .map(function (b) {
      return { cls: b.cls, label: dayLabel_(b.date), time: b.time, child: b.child, guardian: b.guardian, phone: b.phone, memo: b.memo };
    });

  var days = buildDays_(classes, []);
  var total = 0;
  Object.keys(days).forEach(function (name) {
    if (inScope(name)) days[name].forEach(function (d) { total += d.times.length; });
  });

  // 명단은 있는데 아직 신청하지 않은 아이 (담임이 따로 연락할 때 쓰기)
  var missing = [];
  var roster = readRoster_(classes);
  Object.keys(roster).forEach(function (name) {
    if (!inScope(name)) return;
    var booked = {};
    list.forEach(function (b) { if (b.cls === name) booked[nameKey_(b.child)] = true; });
    Object.keys(roster[name]).forEach(function (k) {
      if (!booked[k]) missing.push({ cls: name, child: roster[name][k] });
    });
  });

  return { scope: scope === '*' ? '전체 반' : scope, list: list, slots: total, missing: missing };
}

/* ───────── 내부 함수 ───────── */

function tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Seoul'; }

function randomDigits_(n) {
  var s = '';
  for (var i = 0; i < n; i++) s += Math.floor(Math.random() * 10);
  return s.charAt(0) === '0' ? '1' + s.slice(1) : s;
}

function sheetRows_(name, cols) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, cols).getValues();
}

function readConfig_() {
  var map = {};
  sheetRows_(SHEET.CONFIG, 2).forEach(function (r) { map[String(r[0]).trim()] = r[1]; });
  var yes = function (v) { return /^(예|네|y|yes|o|true)$/i.test(String(v).trim()); };
  return {
    title: String(map['제목'] || '학부모 상담 신청'),
    notice: String(map['안내 문구'] || ''),
    askPhone: yes(map['연락처 받기']),
    open: map['예약 받기'] === undefined ? true : yes(map['예약 받기']),
    checkRoster: map['명단 확인'] === undefined ? true : yes(map['명단 확인']),
    adminCode: String(map['전체 관리자 코드'] || map['관리자 코드'] || '').trim()
  };
}

function readClasses_() {
  var seen = {};
  return sheetRows_(SHEET.CLASSES, 4).map(function (r) {
    return {
      name: String(r[0]).trim(),
      teacher: String(r[1]).trim(),
      minutes: Number(r[2]) || 0,
      code: String(r[3]).trim()
    };
  }).filter(function (c) {
    if (!c.name || seen[c.name]) return false;
    seen[c.name] = true;
    return true;
  });
}

/** 반 칸이 비어 있어도 반이 하나뿐이면 그 반으로 본다. */
function resolveClass_(raw, classes) {
  var name = String(raw == null ? '' : raw).trim();
  if (!name && classes.length === 1) return classes[0].name;
  return name;
}

/** { 반 이름: { 이름키: 명단에 적힌 이름 } } — 명단이 빈 반은 키가 없다. */
function readRoster_(classes) {
  var out = {};
  var known = {};
  classes.forEach(function (c) { known[c.name] = true; });
  sheetRows_(SHEET.ROSTER, 2).forEach(function (r) {
    var cls = resolveClass_(r[0], classes);
    var name = String(r[1]).trim();
    if (!known[cls] || !name) return;
    out[cls] = out[cls] || {};
    out[cls][nameKey_(name)] = name;
  });
  return out;
}

function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  var m = String(v).trim().match(/^(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

function normTime_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'HH:mm');
  var m = String(v).trim().match(/^(\d{1,2})\s*[:시]\s*(\d{0,2})/);
  if (!m) return '';
  return ('0' + m[1]).slice(-2) + ':' + ('0' + (m[2] || '0')).slice(-2);
}

function dayLabel_(date) {
  var p = date.split('-').map(Number);
  var dow = new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay();
  return p[1] + '월 ' + p[2] + '일 (' + WEEKDAYS[dow] + ')';
}

/** { 반 이름: [{date, label, left, times:[{time, taken}]}] } — 지난 시간은 뺀다. */
function buildDays_(classes, bookings) {
  var now = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm');
  var known = {};
  classes.forEach(function (c) { known[c.name] = true; });
  var taken = {};
  bookings.forEach(function (b) { taken[b.cls + '|' + b.date + ' ' + b.time] = true; });

  var grid = {};
  sheetRows_(SHEET.SLOTS, 3).forEach(function (r) {
    var cls = resolveClass_(r[0], classes);
    var date = normDate_(r[1]);
    if (!known[cls] || !date) return;
    var raw = r[2] instanceof Date ? [r[2]] : String(r[2]).split(/[,，\n]/);
    raw.forEach(function (t) {
      var time = normTime_(t);
      if (!time || date + ' ' + time <= now) return;
      grid[cls] = grid[cls] || {};
      grid[cls][date] = grid[cls][date] || {};
      grid[cls][date][time] = true;
    });
  });

  var out = {};
  Object.keys(grid).forEach(function (cls) {
    out[cls] = Object.keys(grid[cls]).sort().map(function (date) {
      var times = Object.keys(grid[cls][date]).sort().map(function (time) {
        return { time: time, taken: !!taken[cls + '|' + date + ' ' + time] };
      });
      var left = times.filter(function (t) { return !t.taken; }).length;
      return { date: date, label: dayLabel_(date), times: times, left: left };
    });
  });
  return out;
}

function readBookRows_(classes) {
  var sh = ensureBookSheet_(SpreadsheetApp.getActive());
  if (sh.getLastRow() < 2) return [];
  var strip = function (v) { return String(v).replace(/^'/, ''); };
  return sh.getRange(2, 1, sh.getLastRow() - 1, BOOK_HEADERS.length).getValues().map(function (r, i) {
    return {
      row: i + 2,
      cls: resolveClass_(r[COL.cls], classes),
      date: normDate_(r[COL.date]),
      time: normTime_(r[COL.time]),
      child: strip(r[COL.child]),
      guardian: strip(r[COL.guardian]),
      phone: strip(r[COL.phone]),
      memo: strip(r[COL.memo]),
      code: String(r[COL.code]).trim(),
      status: String(r[COL.status]).trim()
    };
  });
}

function activeBookings_(classes) {
  return readBookRows_(classes).filter(function (b) { return b.status === ACTIVE && b.date && b.time; });
}

function matchBooking_(child, code) {
  var key = nameKey_(cleanText_(child, 20));
  var c = String(code || '').trim();
  if (!key || !c) throw new Error('아이 이름과 확인번호를 모두 적어 주세요.');
  var found = activeBookings_(readClasses_()).filter(function (b) {
    return nameKey_(b.child) === key && b.code === c;
  })[0];
  if (!found) {
    Utilities.sleep(800);
    throw new Error('일치하는 예약이 없어요. 이름과 확인번호 4자리를 다시 확인해 주세요.');
  }
  return found;
}

function nameKey_(s) { return String(s || '').replace(/\s+/g, ''); }

function cleanText_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);
}

/** 이름 등이 = + - @ 로 시작하면 시트에서 수식으로 실행되지 않도록 막는다. */
function safeCell_(s) { return /^[=+\-@]/.test(s) ? "'" + s : s; }
