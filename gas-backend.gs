/**
 * 사회심리학 실험 - 데이터 수신용 Google Apps Script
 *
 * 컬럼 배치:
 *   A: serverReceivedAt
 *   B: sessionId
 *   C: condition               (high / low)
 *   D: conditionUserClicked
 *   E: forcedCondition
 *   F: choice                  (A / B)
 *   G: choiceName              (varek / lirun)
 *   H: selectDecisionMs        ← 핵심 종속변수
 *   I: totalMs
 *   J: intro_ms
 *   K: start_ms
 *   L: select_ms
 *   M: startedAt
 *   N: finishedAt
 *   O: userAgent
 *   P: select_minus_start_ms   (Apps Script가 직접 계산)
 *   Q: participantName
 */

const SHEET_NAME = 'Sessions';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();
    ensureExtraColumns(sheet);
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

function ensureExtraColumns(sheet) {
  if (sheet.getRange('P1').getValue() !== 'select_minus_start_ms') {
    sheet.getRange('P1').setValue('select_minus_start_ms').setFontWeight('bold');
  }
  if (sheet.getRange('Q1').getValue() !== 'participantName') {
    sheet.getRange('Q1').setValue('participantName').setFontWeight('bold');
  }
}

/**
 * 다음 비어있는 행 번호를 안전하게 구합니다.
 * (getLastRow는 ARRAYFORMULA 잔여물에 의해 잘못된 값을 줄 수 있으므로
 *  A 컬럼(serverReceivedAt) 기준으로 마지막 실데이터 행을 찾습니다)
 */
function getNextRow(sheet) {
  const colA = sheet.getRange('A:A').getValues();
  let last = 0;
  for (let i = 0; i < colA.length; i++) {
    if (colA[i][0] !== '' && colA[i][0] !== null) last = i + 1;
  }
  return last + 1;
}

function appendRow(sheet, d) {
  const t = d.timings || {};
  const row = getNextRow(sheet);
  const startMs = typeof t.start === 'number' ? t.start : null;
  const selectMs = typeof t.select === 'number' ? t.select : null;
  const selectMinusStart = (startMs !== null && selectMs !== null) ? (selectMs - startMs) : '';

  // A-P 컬럼 (P 포함, 16개)
  const main = [
    new Date(), d.sessionId || '', d.condition || '',
    d.conditionUserClicked || '', d.forcedCondition || '',
    d.choice || '', d.choiceName || '',
    d.selectDecisionMs ?? '', d.totalMs ?? '',
    t.intro ?? '', t.start ?? '', t.select ?? '',
    d.startedAt || '', d.finishedAt || '', d.userAgent || '',
    selectMinusStart,
  ];
  sheet.getRange(row, 1, 1, main.length).setValues([main]);
  // Q 컬럼: participantName
  sheet.getRange(row, 17).setValue(d.participantName || '');
}

/**
 * ==== 한 번만 실행하는 마이그레이션 ====
 * - 이전에 깔아둔 ARRAYFORMULA를 제거
 * - 흩어진 모든 데이터(2~끝 행)를 모아서 2행부터 다시 압축 정렬
 * - select_minus_start_ms 를 각 행에서 직접 계산해 P 컬럼에 정적 값으로 저장
 *
 * 실행 방법: 편집기 상단 함수 드롭다운에서 'migrateData' 선택 → ▶ 실행
 */
function migrateData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sessions 시트가 없습니다.');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // 1) 2행부터 끝까지, A~Q(17개 컬럼) 읽어오기
  const all = sheet.getRange(2, 1, lastRow - 1, 17).getValues();

  // 2) 실제 데이터가 있는 행만 필터링 (A 컬럼이 비어있지 않은 경우)
  const dataRows = all.filter(r => r[0] !== '' && r[0] !== null);

  // 3) 각 행에 대해 P 값(=select_ms - start_ms)을 재계산
  const compacted = dataRows.map(r => {
    const k = r[10]; // start_ms
    const l = r[11]; // select_ms
    const p = (typeof k === 'number' && typeof l === 'number') ? l - k : '';
    const row = r.slice();
    row[15] = p; // P 컬럼
    return row;
  });

  // 4) 2행부터 끝까지 전부 깨끗하게 비우기 (수식 포함)
  sheet.getRange(2, 1, lastRow - 1, 17).clearContent();

  // 5) 2행부터 데이터 다시 쓰기
  if (compacted.length > 0) {
    sheet.getRange(2, 1, compacted.length, 17).setValues(compacted);
  }

  // 6) 헤더 한 번 더 보장
  ensureExtraColumns(sheet);
}
