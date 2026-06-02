/**
 * 사회심리학 실험 - 데이터 수신용 Google Apps Script
 *
 * [사용 방법]
 * 1) Google Sheets에서 새 스프레드시트 생성
 * 2) 상단 메뉴: 확장 프로그램(Extensions) > Apps Script
 * 3) 기본 코드 지우고 이 파일 내용 전체 붙여넣기 > 저장
 * 4) 상단 "배포(Deploy)" > "새 배포" > 유형: 웹 앱
 *    - 설명: 아무거나 (예: v1)
 *    - 실행: 나
 *    - 액세스 권한: 모든 사용자(Anyone)
 *    - 배포 클릭 > 권한 승인
 * 5) 발급된 웹앱 URL 을 복사 → index.html 의 WEBHOOK_URL 에 붙여넣기
 *
 * 시트 컬럼은 자동 생성됩니다.
 */

const SHEET_NAME = 'Sessions';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();
    appendRow(sheet, data);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput('OK - 사회심리학 실험 백엔드 동작 중')
    .setMimeType(ContentService.MimeType.TEXT);
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    const headers = [
      'serverReceivedAt',
      'sessionId',
      'condition',
      'conditionUserClicked',
      'forcedCondition',
      'choice',
      'choiceName',
      'selectDecisionMs',
      'totalMs',
      'intro_ms',
      'start_ms',
      'select_ms',
      'startedAt',
      'finishedAt',
      'userAgent',
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function appendRow(sheet, d) {
  const t = d.timings || {};
  sheet.appendRow([
    new Date(),
    d.sessionId || '',
    d.condition || '',
    d.conditionUserClicked || '',
    d.forcedCondition || '',
    d.choice || '',
    d.choiceName || '',
    d.selectDecisionMs ?? '',
    d.totalMs ?? '',
    t.intro ?? '',
    t.start ?? '',
    t.select ?? '',
    d.startedAt || '',
    d.finishedAt || '',
    d.userAgent || '',
  ]);
}
