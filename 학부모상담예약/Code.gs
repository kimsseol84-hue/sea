/**
 * 학부모 상담 예약 웹앱 (Google Apps Script + 구글 시트)
 *
 * 시트 구성
 *  - 설정     : 제목, 반 이름, 담임, 상담 시간, 안내 문구, 관리자 코드 등
 *  - 일정     : 날짜 한 줄 + 시간 목록(쉼표로 구분)
 *  - 예약현황 : 신청 내역 (관리자만 보는 곳)
 *  - 보관     : 지난 상담 기간의 예약 (메뉴 > 지난 예약 보관하기)
 *
 * 공동 담임·원장님이 설정을 바꾸게 하려면 이 스프레드시트를 "편집자"로 공유하세요.
 * 학부모에게는 웹앱 링크만 보냅니다. 학부모는 이름을 볼 수 없습니다.
 */

var SHEET = { CONFIG: '설정', SLOTS: '일정', BOOK: '예약현황', ARCHIVE: '보관' };
var BOOK_HEADERS = ['접수시각', '날짜', '시간', '아이 이름', '보호자', '연락처', '요청사항', '확인번호', '상태'];
var COL = { created: 0, date: 1, time: 2, child: 3, guardian: 4, phone: 5, memo: 6, code: 7, status: 8 };
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

  if (!ss.getSheetByName(SHEET.CONFIG)) {
    var c = ss.insertSheet(SHEET.CONFIG, 0);
    var adminCode = String(Math.floor(100000 + Math.random() * 900000));
    var rows = [
      ['항목', '값', '설명'],
      ['제목', '2학기 학부모 상담 신청', '페이지 맨 위에 보이는 제목'],
      ['반 이름', '햇살반', ''],
      ['담임', '', '예: 김OO 선생님'],
      ['상담 시간(분)', 20, '한 번 상담에 걸리는 시간'],
      ['안내 문구', '원하시는 날짜와 시간을 골라 주세요. 한 아이당 한 번 신청할 수 있어요.', ''],
      ['연락처 받기', '아니오', '예 / 아니오 — 꼭 필요할 때만 "예"로 바꾸세요'],
      ['예약 받기', '예', '"아니오"로 바꾸면 신청이 멈춥니다'],
      ['관리자 코드', adminCode, '웹앱 맨 아래 "관리자"에서 신청자 명단을 볼 때 입력. 학부모에게 알려주지 마세요']
    ];
    c.getRange(1, 1, rows.length, 3).setValues(rows);
    c.getRange('B9').setNumberFormat('@').setValue(adminCode);
    c.getRange('A1:C1').setFontWeight('bold');
    c.setColumnWidth(1, 120); c.setColumnWidth(2, 360); c.setColumnWidth(3, 380);
    c.setFrozenRows(1);
  }

  if (!ss.getSheetByName(SHEET.SLOTS)) {
    var s = ss.insertSheet(SHEET.SLOTS, 1);
    var tz = ss.getSpreadsheetTimeZone();
    var sample = [['날짜', '시간 (쉼표로 구분)']];
    var d = new Date();
    while (sample.length < 4) {
      d.setDate(d.getDate() + 1);
      var dow = Number(Utilities.formatDate(d, tz, 'u')); // 1=월 … 7=일
      if (dow <= 5) sample.push([Utilities.formatDate(d, tz, 'yyyy-MM-dd'), '14:00, 14:30, 15:00, 15:30']);
    }
    s.getRange(1, 1, sample.length, 2).setNumberFormat('@').setValues(sample);
    s.getRange('A1:B1').setFontWeight('bold');
    s.setColumnWidth(1, 140); s.setColumnWidth(2, 360);
    s.setFrozenRows(1);
  }

  ensureBookSheet_(ss);
  if (!ss.getSheetByName(SHEET.ARCHIVE)) {
    var a = ss.insertSheet(SHEET.ARCHIVE);
    a.appendRow(BOOK_HEADERS.concat(['보관일']));
    a.getRange(1, 1, 1, BOOK_HEADERS.length + 1).setFontWeight('bold');
  }

  var first = ss.getSheets()[ss.getSheets().length - 1];
  if (first.getName() === 'Sheet1' || first.getName() === '시트1') {
    if (first.getLastRow() === 0) ss.deleteSheet(first);
  }
  SpreadsheetApp.getUi().alert('설정 완료! "설정"과 "일정" 시트를 반에 맞게 고친 뒤 웹앱으로 배포하세요.');
}

function ensureBookSheet_(ss) {
  var b = ss.getSheetByName(SHEET.BOOK);
  if (!b) {
    b = ss.insertSheet(SHEET.BOOK, 2);
    b.appendRow(BOOK_HEADERS);
    b.getRange(1, 1, 1, BOOK_HEADERS.length).setFontWeight('bold');
    b.getRange('B:C').setNumberFormat('@');
    b.getRange('H:H').setNumberFormat('@');
    b.setFrozenRows(1);
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

/** 학부모 화면에 필요한 공개 정보. 이름·관리자 코드는 절대 포함하지 않는다. */
function getPublicData() {
  var cfg = readConfig_();
  return {
    title: cfg.title,
    className: cfg.className,
    teacher: cfg.teacher,
    minutes: cfg.minutes,
    notice: cfg.notice,
    askPhone: cfg.askPhone,
    open: cfg.open,
    days: buildDays_(cfg, activeBookings_())
  };
}

function book(form) {
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

    var bookings = activeBookings_();
    var days = buildDays_(cfg, bookings);
    var slot = null;
    days.forEach(function (d) {
      if (d.date !== date) return;
      d.times.forEach(function (t) { if (t.time === time) slot = { day: d, t: t }; });
    });
    if (!slot) throw new Error('선택한 시간이 더 이상 없어요. 화면을 새로고침해 주세요.');
    if (slot.t.taken) throw new Error('방금 다른 분이 이 시간을 신청했어요. 다른 시간을 골라 주세요.');

    var key = nameKey_(child);
    var mine = bookings.filter(function (b) { return nameKey_(b.child) === key; })[0];
    if (mine) {
      throw new Error(child + ' 이름으로 이미 ' + dayLabel_(mine.date) + ' ' + mine.time +
        ' 예약이 있어요. 바꾸려면 아래 "내 예약 확인·취소"에서 먼저 취소해 주세요.');
    }

    var code = String(Math.floor(1000 + Math.random() * 9000));
    var sheet = ensureBookSheet_(SpreadsheetApp.getActive());
    var row = [new Date(), date, time, safeCell_(child), safeCell_(guardian), safeCell_(phone), safeCell_(memo), code, ACTIVE];
    var r = sheet.getLastRow() + 1;
    sheet.getRange(r, 2, 1, 2).setNumberFormat('@');
    sheet.getRange(r, 8).setNumberFormat('@');
    sheet.getRange(r, 1, 1, row.length).setValues([row]);
    SpreadsheetApp.flush();

    return { date: date, time: time, label: dayLabel_(date), code: code, child: child };
  } finally {
    lock.releaseLock();
  }
}

function findMyBooking(child, code) {
  var b = matchBooking_(child, code);
  return { date: b.date, time: b.time, label: dayLabel_(b.date), child: b.child };
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

/** 관리자 화면: 코드가 맞을 때만 이름이 담긴 명단을 돌려준다. */
function getAdminData(adminCode) {
  var cfg = readConfig_();
  if (!cfg.adminCode || String(adminCode || '').trim() !== cfg.adminCode) {
    Utilities.sleep(800); // 무작위 대입을 느리게
    throw new Error('관리자 코드가 맞지 않아요.');
  }
  var list = activeBookings_()
    .sort(function (a, b) { return (a.date + a.time).localeCompare(b.date + b.time); })
    .map(function (b) {
      return { label: dayLabel_(b.date), time: b.time, child: b.child, guardian: b.guardian, phone: b.phone, memo: b.memo };
    });
  var total = 0;
  buildDays_(cfg, []).forEach(function (d) { total += d.times.length; });
  return { list: list, upcomingSlots: total, sheetUrl: SpreadsheetApp.getActive().getUrl() };
}

/* ───────── 내부 함수 ───────── */

function tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Seoul'; }

function readConfig_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.CONFIG);
  var map = {};
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      map[String(r[0]).trim()] = r[1];
    });
  }
  var yes = function (v) { return /^(예|네|y|yes|o|true)$/i.test(String(v).trim()); };
  return {
    title: String(map['제목'] || '학부모 상담 신청'),
    className: String(map['반 이름'] || ''),
    teacher: String(map['담임'] || ''),
    minutes: Number(map['상담 시간(분)']) || 0,
    notice: String(map['안내 문구'] || ''),
    askPhone: yes(map['연락처 받기']),
    open: map['예약 받기'] === undefined ? true : yes(map['예약 받기']),
    adminCode: String(map['관리자 코드'] || '').trim()
  };
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

function buildDays_(cfg, bookings) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET.SLOTS);
  if (!sh || sh.getLastRow() < 2) return [];
  var now = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm');
  var taken = {};
  bookings.forEach(function (b) { taken[b.date + ' ' + b.time] = true; });

  var byDate = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    var date = normDate_(r[0]);
    if (!date) return;
    var raw = r[1] instanceof Date ? [r[1]] : String(r[1]).split(/[,，\n]/);
    raw.forEach(function (t) {
      var time = normTime_(t);
      if (!time || date + ' ' + time <= now) return;
      byDate[date] = byDate[date] || {};
      byDate[date][time] = true;
    });
  });

  return Object.keys(byDate).sort().map(function (date) {
    var times = Object.keys(byDate[date]).sort().map(function (time) {
      return { time: time, taken: !!taken[date + ' ' + time] };
    });
    var left = times.filter(function (t) { return !t.taken; }).length;
    return { date: date, label: dayLabel_(date), times: times, left: left };
  });
}

function readBookRows_() {
  var sh = ensureBookSheet_(SpreadsheetApp.getActive());
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, BOOK_HEADERS.length).getValues().map(function (r, i) {
    return {
      row: i + 2,
      date: normDate_(r[COL.date]),
      time: normTime_(r[COL.time]),
      child: String(r[COL.child]).replace(/^'/, ''),
      guardian: String(r[COL.guardian]).replace(/^'/, ''),
      phone: String(r[COL.phone]).replace(/^'/, ''),
      memo: String(r[COL.memo]).replace(/^'/, ''),
      code: String(r[COL.code]).trim(),
      status: String(r[COL.status]).trim()
    };
  });
}

function activeBookings_() {
  return readBookRows_().filter(function (b) { return b.status === ACTIVE && b.date && b.time; });
}

function matchBooking_(child, code) {
  var key = nameKey_(cleanText_(child, 20));
  var c = String(code || '').trim();
  if (!key || !c) throw new Error('아이 이름과 확인번호를 모두 적어 주세요.');
  var found = activeBookings_().filter(function (b) { return nameKey_(b.child) === key && b.code === c; })[0];
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
