/**
 * 사회심리학 실험 - 데이터 수신용 Google Apps Script
 *
 * 컬럼 배치:
 *   A: serverReceivedAt        ← 서버 도착 시각
 *   B: sessionId
 *   C: condition               ← high(신중) / low(가볍게)
 *   D: conditionUserClicked    ← 참여자가 실제 클릭한 조건
 *   E: forcedCondition         ← URL ?cond= 으로 강제된 조건
 *   F: choice                  ← A / B
 *   G: choiceName              ← varek / lirun
 *   H: selectDecisionMs        ← 선택 화면 진입 → 클릭까지(핵심 종속변수)
 *   I: totalMs
 *   J: intro_ms
 *   K: start_ms
 *   L: select_ms
 *   M: startedAt
 *   N: finishedAt
 *   O: userAgent
 *   P: select_minus_start_ms   ← ARRAYFORMULA (addComputedColumn 실행 시 생성)
 *   Q: participantName         ← 참여자 이름/ID (자동 보장)
 */

const SHEET_NAME = 'Sessions';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();
    ensureParticipantColumn(sheet);
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
      'serverReceivedAt','sessionId','condition','conditionUserClicked',
      'forcedCondition','choice','choiceName','selectDecisionMs','totalMs',
      'intro_ms','start_ms','select_ms','startedAt','finishedAt','userAgent',
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Q 컬럼에 participantName 헤더 자동 보장 (이미 있으면 skip).
 * P는 select_minus_start_ms 수식 영역이라 건드리지 않습니다.
 */
function ensureParticipantColumn(sheet) {
  if (sheet.getRange('Q1').getValue() === 'participantName') return;
  sheet.getRange('Q1').setValue('participantName').setFontWeight('bold');
}

function appendRow(sheet, d) {
  const t = d.timings || {};
  const row = sheet.getLastRow() + 1;
  // A-O 컬럼 (기존 데이터)
  const main = [
    new Date(), d.sessionId || '', d.condition || '',
    d.conditionUserClicked || '', d.forcedCondition || '',
    d.choice || '', d.choiceName || '',
    d.selectDecisionMs ?? '', d.totalMs ?? '',
    t.intro ?? '', t.start ?? '', t.select ?? '',
    d.startedAt || '', d.finishedAt || '', d.userAgent || '',
  ];
  sheet.getRange(row, 1, 1, main.length).setValues([main]);
  // Q 컬럼(=17번째): participantName.
  // P 컬럼(=16번째)은 수식 영역이라 건너뜁니다.
  sheet.getRange(row, 17).setValue(d.participantName || '');
}

/**
 * 한 번만 실행하면 P 컬럼에 select_ms - start_ms 자동 계산 수식이 추가됩니다.
 * Apps Script 편집기 상단 함수 드롭다운에서 'addComputedColumn' 선택 후 ▶ 실행.
 */
function addComputedColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sessions 시트가 없습니다.');
  sheet.getRange('P1').setValue('select_minus_start_ms').setFontWeight('bold');
  sheet.getRange('P2').setFormula('=ARRAYFORMULA(IF(LEN(L2:L), L2:L-K2:K, ""))');
}
