/**
 * 사회심리학 실험 - 데이터 수신용 Google Apps Script
 *
 * 컬럼 배치 (비교 핵심 4개를 C-F열에 모음):
 *   A: serverReceivedAt
 *   B: sessionId
 *   C: participantName        ← 비교용
 *   D: condition              ← 비교용
 *   E: choiceName             ← 비교용
 *   F: selectDecisionMs       ← 비교용 (핵심 종속변수)
 *   G: conditionUserClicked
 *   H: forcedCondition
 *   I: choice
 *   J: totalMs
 *   K: intro_ms
 *   L: start_ms
 *   M: select_ms
 *   N: startedAt
 *   O: finishedAt
 *   P: userAgent
 *   Q: select_minus_start_ms
 *
 * 동시 쓰기 충돌 방지를 위해 LockService 사용 — 덮어쓰기 발생 안 함.
 */

const SHEET_NAME = 'Sessions';

const HEADERS = [
  'serverReceivedAt',
  'sessionId',
  'participantName',
  'condition',
  'choiceName',
  'selectDecisionMs',
  'conditionUserClicked',
  'forcedCondition',
  'choice',
  'totalMs',
  'intro_ms',
  'start_ms',
  'select_ms',
  'startedAt',
  'finishedAt',
  'userAgent',
  'select_minus_start_ms',
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // 최대 30초 대기 → 동시 요청 줄세움
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
  } finally {
    try { lock.releaseLock(); } catch (e) {}
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
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function appendRow(sheet, d) {
  const t = d.timings || {};
  const startMs = typeof t.start === 'number' ? t.start : null;
  const selectMs = typeof t.select === 'number' ? t.select : null;
  const selectMinusStart = (startMs !== null && selectMs !== null) ? (selectMs - startMs) : '';

  // HEADERS 순서와 1:1 대응
  sheet.appendRow([
    new Date(),                    // A
    d.sessionId || '',             // B
    d.participantName || '',       // C
    d.condition || '',             // D
    d.choiceName || '',            // E
    d.selectDecisionMs ?? '',      // F
    d.conditionUserClicked || '',  // G
    d.forcedCondition || '',       // H
    d.choice || '',                // I
    d.totalMs ?? '',               // J
    t.intro ?? '',                 // K
    t.start ?? '',                 // L
    t.select ?? '',                // M
    d.startedAt || '',             // N
    d.finishedAt || '',            // O
    d.userAgent || '',             // P
    selectMinusStart,              // Q
  ]);
}

/**
 * ==== 한 번 실행하는 마이그레이션 ====
 * 기존 시트의 헤더를 인식해서 새 컬럼 순서로 데이터 재배치합니다.
 * - 이전 ARRAYFORMULA / 빈 행 정리
 * - 컬럼 순서를 새 HEADERS 순서로 변환
 * - 데이터를 2번행부터 압축 정렬
 *
 * 실행 방법: 편집기 상단 함수 드롭다운에서 'migrateData' 선택 → ▶ 실행
 */
function migrateData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sessions 시트가 없습니다.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // 기존 헤더 + 데이터 읽기
  let dataRows = [];
  let oldIdx = {};
  if (lastRow >= 1 && lastCol >= 1) {
    const oldHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    oldHeaders.forEach((h, i) => { if (h) oldIdx[String(h)] = i; });

    if (lastRow >= 2) {
      const all = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      // A 컬럼이 비어있지 않은 실데이터 행만 (단, 기존 A는 serverReceivedAt 위치)
      const aIdx = oldIdx['serverReceivedAt'] ?? 0;
      dataRows = all.filter(r => r[aIdx] !== '' && r[aIdx] !== null);
    }
  }

  // 새 순서로 데이터 변환
  const newRows = dataRows.map(row => HEADERS.map(headerName => {
    if (headerName === 'select_minus_start_ms') {
      // 이전 P 값이 수식이었을 수 있으므로 재계산 시도
      const sIdx = oldIdx['start_ms'];
      const seIdx = oldIdx['select_ms'];
      const s = (sIdx !== undefined) ? row[sIdx] : '';
      const se = (seIdx !== undefined) ? row[seIdx] : '';
      if (typeof s === 'number' && typeof se === 'number') return se - s;
      // 기존 P 값이 있으면 그대로 사용
      const oldI = oldIdx[headerName];
      return (oldI !== undefined) ? row[oldI] : '';
    }
    const oldI = oldIdx[headerName];
    return (oldI !== undefined && oldI < row.length) ? row[oldI] : '';
  }));

  // 시트 전체 클리어
  sheet.clear();

  // 새 헤더
  sheet.appendRow(HEADERS);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  // 데이터 새 순서로 기록
  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, HEADERS.length).setValues(newRows);
  }
}
