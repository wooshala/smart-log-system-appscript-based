function importFebHistoryToDb() {
  const FEB_SS_ID = '1kjmM1-P1Fx3HUyUxtuy-uC9gLAmDiHiu'; // 2월 파일 ID
  const ss    = SpreadsheetApp.openById(FEB_SS_ID);
  const tz    = 'Asia/Seoul';
  const year  = 2026;
  const month = 2;
  const dbSh  = getCustDbSheet_();
  let added   = 0;
  let skipped = 0;

  for (let day = 1; day <= 28; day++) {
    const dt      = new Date(year, month - 1, day);
    const wk      = ['일','월','화','수','목','금','토'];
    const tabName = `${day}(${wk[dt.getDay()]})`;
    const sheet   = ss.getSheetByName(tabName);

    if (!sheet) {
      Logger.log(`탭 없음: ${tabName}`);
      continue;
    }

    const dateStr = Utilities.formatDate(dt, tz, 'yyyy-MM-dd');
    const lastRow = sheet.getLastRow();
    if (lastRow < 4) continue;

    // ── 대실 읽기 (4~35행) ──
    const daeEndRow = Math.min(35, lastRow);
    const daeData   = sheet.getRange(4, 1, daeEndRow - 3, 23).getValues();
    for (const r of daeData) {
      const name = String(r[8] || '').trim(); // I열
      if (!name) { skipped++; continue; }

      dbSh.appendRow([
        name,
        String(r[10] || '').trim(), // B: 차종 (K열)
        dateStr,                     // C: 방문일
        '대실',                      // D: 숙박타입
        String(r[1]  || '').trim(), // E: 객실번호 (B열)
        r[4] || '',                  // F: 금액 (E열)
        String(r[5]  || '').trim(), // G: 결제방식 (F열)
        String(r[11] || '').trim(), // H: 주차위치 (L열)
        String(r[9]  || '').trim(), // I: 특이사항 (J열)
        '숙박일지_이력이전',
      ]);
      added++;
    }

    // ── 숙박 읽기 (4~45행) ──
    const sukEndRow = Math.min(45, lastRow);
    const sukData   = sheet.getRange(4, 1, sukEndRow - 3, 23).getValues();
    for (const r of sukData) {
      const name = String(r[19] || '').trim(); // T열
      if (!name) { skipped++; continue; }

      dbSh.appendRow([
        name,
        String(r[21] || '').trim(), // B: 차종 (V열)
        dateStr,                     // C: 방문일
        '숙박',                      // D: 숙박타입
        String(r[13] || '').trim(), // E: 객실번호 (N열)
        r[15] || '',                 // F: 금액 (P열)
        String(r[16] || '').trim(), // G: 결제방식 (Q열)
        String(r[22] || '').trim(), // H: 주차위치 (W열)
        String(r[20] || '').trim(), // I: 특이사항 (U열)
        '숙박일지_이력이전',
      ]);
      added++;
    }

    Logger.log(`2월 ${tabName} 완료`);
  }

  Logger.log(`완료! 추가: ${added}건 / 빈행 스킵: ${skipped}건`);
}
function attachNotesToHistory() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const year    = 2026;
  const month   = 3;
  const FROM_DAY = 1;
  const TO_DAY   = 7;
  let count = 0;

  for (let day = FROM_DAY; day <= TO_DAY; day++) {
    const dt      = new Date(year, month - 1, day);
    const wk      = ['일','월','화','수','목','금','토'];
    const tabName = `${day}(${wk[dt.getDay()]})`;
    const sheet   = ss.getSheetByName(tabName);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    if (lastRow < 4) continue;

    // ── 대실: I열(9) 이름 셀 ──
    for (let r = 4; r <= Math.min(35, lastRow); r++) {
      const name = String(sheet.getRange(r, 9).getValue() || '').trim();
      if (!name) continue;
      const car  = String(sheet.getRange(r, 11).getValue() || '').trim(); // K열
      attachNote_(sheet, r, 9, name, car);
      count++;
    }

    // ── 숙박: T열(20) 이름 셀 ──
    for (let r = 4; r <= Math.min(45, lastRow); r++) {
      const name = String(sheet.getRange(r, 20).getValue() || '').trim();
      if (!name) continue;
      const car  = String(sheet.getRange(r, 22).getValue() || '').trim(); // V열
      attachNote_(sheet, r, 20, name, car);
      count++;
    }

    Logger.log(`${tabName} 말풍선 완료`);
  }

  Logger.log(`말풍선 총 ${count}개 생성 완료`);
}
function attachNotesToHistory() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const year    = 2026;
  const month   = 3;
  const FROM_DAY = 1;
  const TO_DAY   = 7;
  let count = 0;

  for (let day = FROM_DAY; day <= TO_DAY; day++) {
    const dt      = new Date(year, month - 1, day);
    const wk      = ['일','월','화','수','목','금','토'];
    const tabName = `${day}(${wk[dt.getDay()]})`;
    const sheet   = ss.getSheetByName(tabName);
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    if (lastRow < 4) continue;

    // ── 대실: I열(9) 이름 셀 ──
    for (let r = 4; r <= Math.min(35, lastRow); r++) {
      const name = String(sheet.getRange(r, 9).getValue() || '').trim();
      if (!name) continue;
      const car  = String(sheet.getRange(r, 11).getValue() || '').trim(); // K열
      attachNote_(sheet, r, 9, name, car);
      count++;
    }

    // ── 숙박: T열(20) 이름 셀 ──
    for (let r = 4; r <= Math.min(45, lastRow); r++) {
      const name = String(sheet.getRange(r, 20).getValue() || '').trim();
      if (!name) continue;
      const car  = String(sheet.getRange(r, 22).getValue() || '').trim(); // V열
      attachNote_(sheet, r, 20, name, car);
      count++;
    }

    Logger.log(`${tabName} 말풍선 완료`);
  }

  Logger.log(`말풍선 총 ${count}개 생성 완료`);
}

