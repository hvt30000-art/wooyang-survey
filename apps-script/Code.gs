/**
 * 서도호전 감상카드 → 구글 시트 답변 받기
 *
 * 구글 시트의 [확장 프로그램 → Apps Script]에 이 파일 내용을 통째로 붙여 넣어 사용합니다.
 * 감상카드 페이지가 학생 답을 보내면 '답변' 시트에 한 줄씩 추가합니다.
 * 이 웹 앱은 줄을 추가하기만 하고, 시트 내용을 읽거나 지우지 않습니다.
 *
 * 페이지가 보내는 내용(JSON, text/plain):
 *   { device: "무작위 번호표", name: "이름(선택)", consent: true/false,
 *     answers: { c1: "...", ..., c9: "..." } }
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
  if (!data || typeof data !== 'object' || !data.answers || typeof data.answers !== 'object') {
    return json_({ ok: false, error: 'bad-data' });
  }

  var answers = CARDS.map(function (c) { return clean_(data.answers[c[0]], MAX_ANSWER); });
  var count = answers.filter(function (a) { return a.trim() !== ''; }).length;
  if (count === 0) return json_({ ok: false, error: 'empty' });

  var row = [
    new Date(),
    clean_(data.device, MAX_DEVICE),
    clean_(data.name, MAX_NAME),
    data.consent === true ? '동의' : '동의 안 함',
    count
  ].concat(answers);

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    getSheet_().appendRow(row);
  } catch (err) {
    return json_({ ok: false, error: 'busy' });
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
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
