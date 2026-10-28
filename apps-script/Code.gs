/**
 * Kenta Coffee 문의 폼 → Google 스프레드시트 저장 + 메일 알림
 *
 * 설치 방법
 * 1. 새 Google 스프레드시트를 만들고 [확장 프로그램] → [Apps Script]를 엽니다.
 * 2. 기본 Code.gs 내용을 지우고 이 파일 전체를 붙여넣은 뒤 저장합니다.
 * 3. 위쪽 함수 목록에서 setup 을 선택하고 [실행]을 누른 뒤 권한을 허용합니다.
 *    → "문의" 시트가 만들어지고, 알림 메일 권한이 승인됩니다.
 * 4. [배포] → [새 배포] → 유형 "웹 앱"
 *    - 실행 사용자: 나
 *    - 액세스 권한이 있는 사용자: 모든 사용자
 *    배포 후 나오는 웹 앱 URL(https://script.google.com/macros/s/.../exec)을 홈페이지에 넣습니다.
 *
 * 코드를 고친 뒤에는 [배포] → [배포 관리] → 수정(연필) → 버전 "새 버전"으로 다시 배포해야 반영됩니다.
 */

// 알림을 받을 메일 주소. 비워두면 이 스크립트를 배포한 Google 계정으로 보냅니다.
const NOTIFY_EMAIL = '';

const SHEET_NAME = '문의';
const HEADERS = ['접수 시각', '이름', '연락처', '매장명', '문의 내용', '메시지', '처리 상태', '메모'];

// 같은 연락처로 다시 보낼 수 있을 때까지 기다리는 시간(초)
const RESUBMIT_WAIT_SECONDS = 60;

function setup() {
  getSheet_();
  MailApp.getRemainingDailyQuota();
  Logger.log('준비 완료. 알림 메일 주소: ' + notifyAddress_());
}

function doGet() {
  return json_({ ok: true, service: 'kenta-coffee-contact' });
}

function doPost(e) {
  try {
    const p = (e && e.parameter) || {};

    // 사람 눈에는 보이지 않는 칸. 값이 있으면 봇으로 보고 조용히 성공 처리합니다.
    if (p.website) return json_({ ok: true });

    const data = {
      name: clean_(p.name, 50),
      phone: clean_(p.phone, 30),
      shop: clean_(p.shop, 100),
      types: clean_(p.types, 100),
      message: clean_(p.message, 2000),
    };

    if (!data.name || !data.phone) return json_({ ok: false, error: 'missing' });
    if (!/^[0-9+\-() ]{8,20}$/.test(data.phone)) return json_({ ok: false, error: 'phone' });
    if (p.consent !== 'yes') return json_({ ok: false, error: 'consent' });

    const cache = CacheService.getScriptCache();
    const key = 'phone:' + data.phone.replace(/\D/g, '');
    if (cache.get(key)) return json_({ ok: false, error: 'too_many' });
    cache.put(key, '1', RESUBMIT_WAIT_SECONDS);

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      getSheet_().appendRow([
        new Date(),
        safeCell_(data.name),
        "'" + data.phone, // 앞자리 0이 사라지지 않도록 텍스트로 저장
        safeCell_(data.shop),
        safeCell_(data.types),
        safeCell_(data.message),
        '접수',
        '',
      ]);
    } finally {
      lock.releaseLock();
    }

    sendNotice_(data);
    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server' });
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm');
    sheet.setColumnWidth(6, 360);
  }
  return sheet;
}

function sendNotice_(d) {
  const rows = [
    ['이름', d.name],
    ['연락처', d.phone],
    ['매장명', d.shop || '-'],
    ['문의 내용', d.types || '-'],
    ['메시지', d.message || '-'],
  ].map(function (r) {
    return '<tr><th align="left" style="padding:6px 16px 6px 0;color:#6E6A66;font-weight:500;vertical-align:top">' + r[0] +
      '</th><td style="padding:6px 0;white-space:pre-wrap">' + escape_(r[1]) + '</td></tr>';
  }).join('');

  const url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  MailApp.sendEmail({
    to: notifyAddress_(),
    subject: '[Kenta Coffee] 새 문의 - ' + d.name + (d.shop ? ' (' + d.shop + ')' : ''),
    htmlBody:
      '<div style="font-family:sans-serif;font-size:15px;color:#1D1B19">' +
      '<p>홈페이지 문의 폼으로 새 문의가 들어왔습니다.</p>' +
      '<table style="border-collapse:collapse">' + rows + '</table>' +
      '<p style="margin-top:20px"><a href="' + url + '">스프레드시트에서 보기</a></p></div>',
  });
}

function notifyAddress_() {
  return NOTIFY_EMAIL || Session.getEffectiveUser().getEmail();
}

function clean_(v, max) {
  return String(v || '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max);
}

// =, +, -, @ 로 시작하는 값이 수식으로 실행되지 않도록 막습니다.
function safeCell_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function escape_(v) {
  return String(v).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
