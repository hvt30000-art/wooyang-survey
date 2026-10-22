/**
 * 서도호전 감상카드 → 구글 시트 답변 받기
 *
 * 구글 시트의 [확장 프로그램 → Apps Script]에 이 파일 내용을 통째로 붙여 넣어 사용합니다.
 * 감상카드 페이지는 학생이 마지막 화면에서 '보낼게요'를 고른 뒤부터 답을 자동으로 보냅니다.
 * 휴대폰 번호표가 같은 줄이 있으면 그 줄을 최신 답으로 덮어쓰고, 없으면 새 줄을 추가합니다.
 * 학생이 나중에 '보내지 않을게요'로 바꾸면 그 학생의 줄을 지웁니다.
 * 이 웹 앱은 밖으로 시트 내용을 돌려주지 않습니다.
 *
 * 페이지가 보내는 내용(JSON, text/plain):
 *   { device: "무작위 번호표", name: "이름", consent: true/false/null,
 *     answers: { c1: "...", ..., c9: "..." } }
 *   또는 { device: "무작위 번호표", withdraw: true }
 */

var SHEET_NAME = '답변';
var CARDS = [
  ['c1', '1. 오늘의 한 장면'],
  ['c2', '2. 나의 씨앗'],
  ['c3', '3. 오래 기억하고 싶은 장소'],
  ['c4', '4. 나만의 완벽한 집'],
  ['c5', '5. 다리를 놓는다면'],
  ['c6', '6. 우리 안의 나'],
  ['c7', '7. 나를 지탱하는 것들'],
  ['c8', '8. 몸에 지니는 기억'],
  ['c9', '9. 나에게 집은']
];
var HEADER = ['받은 시각', '휴대폰 번호표', '이름', '활용 동의', '적은 개수']
  .concat(CARDS.map(function (c) { return c[1]; }));
var MAX_ANSWER = 5000;   // 답 하나 최대 글자 수
var MAX_NAME = 40;
var MAX_DEVICE = 64;

/** 처음 한 번 편집기에서 실행: '답변' 시트와 머리글을 만들고 권한 승인을 받습니다. */
function setup() {
  var sheet = getSheet_();
  SpreadsheetApp.getActiveSpreadsheet().setSpreadsheetTimeZone('Asia/Seoul');
  Logger.log('준비 완료: "' + sheet.getName() + '" 시트');
}

/** 브라우저로 웹 앱 주소를 열었을 때 동작 확인용 */
function doGet() {
  return json_({ ok: true, message: '서도호전 감상카드 답변 받기가 켜져 있습니다.' });
}

function doPost(e) {
  var data;
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '');
  } catch (err) {
    return json_({ ok: false, error: 'bad-json' });
  }
  if (!data || typeof data !== 'object') return json_({ ok: false, error: 'bad-data' });
  var device = clean_(data.device, MAX_DEVICE);
  if (!device) return json_({ ok: false, error: 'no-device' });

  // 학생이 '보내지 않을게요'로 바꾸면 그 학생의 줄을 지웁니다.
  if (data.withdraw === true) {
    return withLock_(function (sheet) {
      var at = findRow_(sheet, device);
      if (at > 0) sheet.deleteRow(at);
    });
  }
  if (!data.answers || typeof data.answers !== 'object') return json_({ ok: false, error: 'bad-data' });
  var answers = CARDS.map(function (c) { return clean_(data.answers[c[0]], MAX_ANSWER); });
  var count = answers.filter(function (a) { return a.trim() !== ''; }).length;

  var row = [
    new Date(),
    device,
    clean_(data.name, MAX_NAME),
    data.consent === true ? '동의' : (data.consent === false ? '동의 안 함' : '아직 안 고름'),
    count
  ].concat(answers);

  return withLock_(function (sheet) {
    var at = findRow_(sheet, device);
    if (at > 0) sheet.getRange(at, 1, 1, row.length).setValues([row]);
    else sheet.appendRow(row);
  });
}

/** 동시에 들어온 요청이 섞이지 않게 잠근 뒤 시트 작업을 합니다. */
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    fn(getSheet_());
  } catch (err) {
    return json_({ ok: false, error: 'busy' });
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

/** 휴대폰 번호표가 같은 줄의 번호(없으면 -1) */
function findRow_(sheet, device) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var ids = sheet.getRange(2, 2, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === device) return i + 2;
  }
  return -1;
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADER.length).setFontWeight('bold');
  }
  return sheet;
}

/** 문자열로 바꾸고 길이를 자르고, 시트가 수식으로 읽지 않게 막습니다. */
function clean_(v, max) {
  var s = (typeof v === 'string') ? v : '';
  s = s.slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
