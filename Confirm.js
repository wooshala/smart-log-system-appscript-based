// ════════════════════════════════════════════════════════
//  숙박확인서  Confirm.gs  ─  Sheets 기반 PDF 최종본
//  ※ getOrCreateFolder_ → Quote.gs 사용
//  ※ sendSms_           → Code.gs 사용
// ════════════════════════════════════════════════════════

const CONFIRM_DB_SHEET  = '확인서DB';
const RESERVATION_SHEET = 'RES_DB';
const CONFIRM_FOLDER    = '숙박확인서_PDF';
const SIGNATURE_FILE_ID = '1SdjQQ2F_xBtPTEkWJG37I1HDpB9IAROh';

// ── 색상 상수 (전역, Quote.gs 와 이름 다름) ───────────────
const C_NAVY  = '#1B2A4A';
const C_GOLD  = '#C9A84C';
const C_WHITE = '#FFFFFF';
const C_LGRAY = '#F2F2F2';
const C_DGRAY = '#444444';
const C_MGRAY = '#888888';



// ════════════════════════════════════════════════════════
//  2. 확인서 저장  ─  ConfirmSidebar.html 진입점
// ════════════════════════════════════════════════════════
function saveConfirm(payload, lang) {
  try {
    var confirmId = generateConfirmId_();
    var pdfUrl    = createConfirmPdf_(confirmId, payload, lang || 'ko');
    saveToDb_(confirmId, payload, pdfUrl, lang || 'ko');
    return { success: true, confirmId: confirmId, pdfUrl: pdfUrl };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// ════════════════════════════════════════════════════════
//  3. PDF 생성  ─  임시 Sheet 생성 → 그리기 → 내보내기 → 삭제
// ════════════════════════════════════════════════════════
function createConfirmPdf_(confirmId, payload, lang) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = '확인서_임시_' + confirmId;
  var tmp       = ss.insertSheet(sheetName);
  try {
    tmp.setHiddenGridlines(true);
    drawConfirmSheet_(tmp, payload, confirmId, lang);
    SpreadsheetApp.flush();
    return exportConfirmToPdf_(ss, tmp, confirmId, lang);
  } finally {
    ss.deleteSheet(tmp);
  }
}


// ════════════════════════════════════════════════════════
//  4. Sheet 레이아웃 그리기
// ════════════════════════════════════════════════════════
function drawConfirmSheet_(sheet, p, confirmId, lang) {
  var isEn = (lang === 'en');

  // ── 컬럼 너비 (6열, 합계 590 px) ──────────────────────
  // A:레이블/번호(75) B:날짜/값1(120) C:요일/값1끝(60)
  // D:레이블2/객실(100) E:결제방식(115) F:금액(120)
  [75, 120, 60, 100, 115, 120].forEach(function(w, i) {
    sheet.setColumnWidth(i + 1, w);
  });

  // ── G열 이후 모두 숨기기 (PDF 좌편향 방지) ──────────────
  var maxCol = sheet.getMaxColumns();
  if (maxCol > 6) sheet.hideColumns(7, maxCol - 6);

  var r = 1;
  function H(n, h) { sheet.setRowHeight(n, h); }

  // ── 행 1: 문서 제목 ──────────────────────────────────
  H(r, 50);
  applyMergedCell_(sheet, r, 1, 6,
    isEn ? 'ACCOMMODATION CONFIRMATION' : '숙박확인서',
    { bg: C_NAVY, fc: C_WHITE, bold: true, size: 21,
      align: 'center', va: 'middle' });
  r++;

  // ── 행 2: 호텔 정보 ──────────────────────────────────
  H(r, 22);
  applyMergedCell_(sheet, r, 1, 6,
    CFG.HOTEL_ADDRESS + '  |  TEL ' + CFG.HOTEL_TEL +
    '  |  ' + (isEn ? 'Biz. No.: ' : '사업자등록번호: ') + CFG.HOTEL_BIZ_NO,
    { bg: C_GOLD, fc: C_WHITE, size: 9,
      align: 'center', va: 'middle' });
  r++;

  // ── 행 3: 여백 ───────────────────────────────────────
  H(r, 10); r++;

  // ── 행 4: 투숙객명 / 확인서번호 ─────────────────────
  H(r, 30);
  applyInfoRow_(sheet, r,
    isEn ? 'Guest Name'       : '투숙객명',   p.guestName || '',
    isEn ? 'Confirmation No.' : '확인서번호', confirmId);
  r++;

  // ── 행 5: 연락처 / 발급일자 ──────────────────────────
  H(r, 30);
  applyInfoRow_(sheet, r,
    isEn ? 'Contact'    : '연락처',   p.phone || '',
    isEn ? 'Issue Date' : '발급일자', today_());
  r++;

  // ── 행 6: 여백 ───────────────────────────────────────
  H(r, 10); r++;

  // ── 행 7: 섹션 헤더 ──────────────────────────────────
  H(r, 30);
  applyMergedCell_(sheet, r, 1, 6,
    isEn ? '■  ACCOMMODATION DETAILS' : '■  숙박 내역',
    { bg: C_NAVY, fc: C_WHITE, bold: true, size: 11,
      align: 'center', va: 'middle' });
  r++;

  // ── 행 8: 테이블 헤더 ────────────────────────────────
  H(r, 30);
  var hdrs = isEn
    ? ['No.', 'Date', 'Day', 'Room / Type', 'Payment', 'Amount (KRW)']
    : ['번호', '날짜',  '요일', '객실 / 유형', '결제방식',  '금액'];
  hdrs.forEach(function(txt, i) {
    applyCell_(sheet.getRange(r, i + 1), txt,
      { bg: C_GOLD, fc: C_WHITE, bold: true, size: 10,
        align: 'center', va: 'middle' });
  });
  r++;

  // ── 행 9~: 1박별 데이터 행 ───────────────────────────
  var DOW_EN = { '일':'Sun','월':'Mon','화':'Tue','수':'Wed',
                 '목':'Thu','금':'Fri','토':'Sat' };
  var nRows  = buildNightlyRows_(p.checkIn, p.checkOut, p.roomName, p.totalAmt, p.stayType);

  nRows.forEach(function(nr, idx) {
    H(r, 27);
    var rowVals = [
      String(idx + 1),
      nr.date || '',
      isEn ? (DOW_EN[nr.dow] || nr.dow || '') : (nr.dow || ''),
      (p.roomName || '') + (p.stayType ? ' [' + p.stayType + ']' : ''),
      idx === 0 ? (isEn ? translatePayEn_(p.payMethod) : (p.payMethod || '현금')) : '',
      isEn ? formatAmtEn_(nr.amount) : formatAmt_(nr.amount)
    ];
    rowVals.forEach(function(v, i) {
      applyCell_(sheet.getRange(r, i + 1), v,
        { bg: (idx % 2 === 0 ? C_WHITE : C_LGRAY), fc: C_DGRAY, size: 10,
          align: 'center', va: 'middle' });
    });
    r++;
  });

  // ── 빈 행 (합계 위 여백) ─────────────────────────────
  var blankCnt = Math.max(3, 8 - nRows.length);
  for (var k = 0; k < blankCnt; k++) { H(r, 18); r++; }

  // ── 합계 행 ──────────────────────────────────────────
  H(r, 34);
  applyMergedCell_(sheet, r, 1, 5,
    isEn ? 'TOTAL AMOUNT' : '합   계',
    { bg: C_NAVY, fc: C_WHITE, bold: true, size: 13,
      align: 'center', va: 'middle' });
  applyCell_(sheet.getRange(r, 6),
    isEn ? formatAmtEn_(p.totalAmt) : formatAmt_(p.totalAmt),
    { bg: C_NAVY, fc: C_GOLD, bold: true, size: 13,
      align: 'center', va: 'middle' });
  r++;

  // ── 여백 ─────────────────────────────────────────────
  H(r, 14); r++;

  // ── 확인 문구 ─────────────────────────────────────────
  H(r, 36);
  applyMergedCell_(sheet, r, 1, 6,
    isEn
      ? 'This is to confirm that the above accommodation service has been duly provided.'
      : '위와 같이 숙박 서비스 이용 내역을 확인합니다.',
    { fc: C_DGRAY, size: 10, align: 'center', va: 'middle', italic: true });
  r++;

  // ── 여백 ─────────────────────────────────────────────
  H(r, 10); r++;

  // ── 발급처 ───────────────────────────────────────────
  H(r, 24);
  applyMergedCell_(sheet, r, 1, 6,
    (isEn ? 'Issued by : ' : '발급처 : ') + CFG.HOTEL_NAME,
    { fc: C_MGRAY, size: 10, align: 'right', va: 'middle' });
  r++;

  // ── 서명 레이블 ───────────────────────────────────────
  H(r, 22);
  applyMergedCell_(sheet, r, 1, 6,
    isEn ? 'Authorized Signature :' : '서명 :',
    { fc: C_MGRAY, size: 10, align: 'right', va: 'middle' });
  r++;

  // ── 서명 이미지 (base64 → newCellImage) ───────────────
  H(r, 64);
  try {
    var sigBlob = DriveApp.getFileById(SIGNATURE_FILE_ID).getBlob();
    var b64     = Utilities.base64Encode(sigBlob.getBytes());
    var mime    = sigBlob.getContentType() || 'image/png';
    var dataUrl = 'data:' + mime + ';base64,' + b64;
    var cellImg = SpreadsheetApp.newCellImage().setSourceUrl(dataUrl).build();
    var sigRange = sheet.getRange(r, 4, 1, 3);
    sigRange.merge();
    sigRange.setValue(cellImg);
    sigRange.setHorizontalAlignment('right').setVerticalAlignment('middle');
  } catch (e) {
    applyMergedCell_(sheet, r, 1, 6,
      isEn ? '(Signature)' : '(서명)',
      { fc: C_MGRAY, size: 10, align: 'right', va: 'middle' });
  }
  r++;

  // ── 대표자명 ──────────────────────────────────────────
  H(r, 24);
  applyMergedCell_(sheet, r, 1, 6,
    isEn
      ? CFG.HOTEL_NAME + '   CEO : ' + CFG.HOTEL_CEO
      : CFG.HOTEL_NAME + '   대표   ' + CFG.HOTEL_CEO,
    { fc: C_DGRAY, size: 10, align: 'right', va: 'middle' });
  r++;

  // ── 하단 여백 ─────────────────────────────────────────
  H(r, 10);
}


// ════════════════════════════════════════════════════════
//  스타일 헬퍼
// ════════════════════════════════════════════════════════
function applyCell_(range, text, s) {
  range.setValue(text);
  if (s.bg)     range.setBackground(s.bg);
  if (s.fc)     range.setFontColor(s.fc);
  if (s.bold)   range.setFontWeight('bold');
  if (s.italic) range.setFontStyle('italic');
  if (s.size)   range.setFontSize(s.size);
  if (s.align)  range.setHorizontalAlignment(s.align);
  if (s.va)     range.setVerticalAlignment(s.va);
  range.setFontFamily('Arial');
}

function applyMergedCell_(sheet, row, colStart, colSpan, text, s) {
  var range = sheet.getRange(row, colStart, 1, colSpan);
  range.merge();
  applyCell_(range, text, s);
}

function applyInfoRow_(sheet, row, lbl1, val1, lbl2, val2) {
  // A: 레이블1 / B-C 병합: 값1 / D: 레이블2 / E-F 병합: 값2
  applyCell_(sheet.getRange(row, 1), lbl1,
    { bg: C_NAVY, fc: C_WHITE, bold: true, size: 9,
      align: 'center', va: 'middle' });

  var v1 = sheet.getRange(row, 2, 1, 2); v1.merge();
  applyCell_(v1, val1,
    { bg: C_LGRAY, fc: C_DGRAY, size: 9, align: 'left', va: 'middle' });

  applyCell_(sheet.getRange(row, 4), lbl2,
    { bg: C_NAVY, fc: C_WHITE, bold: true, size: 9,
      align: 'center', va: 'middle' });

  var v2 = sheet.getRange(row, 5, 1, 2); v2.merge();
  applyCell_(v2, val2,
    { bg: C_LGRAY, fc: C_DGRAY, size: 9, align: 'left', va: 'middle' });
}


// ════════════════════════════════════════════════════════
//  5. PDF 내보내기  ─  좌우 여백 0.50인치 균일
// ════════════════════════════════════════════════════════
function exportConfirmToPdf_(ss, sheet, confirmId, lang) {
  var ssId     = ss.getId();
  var gid      = sheet.getSheetId();
  var fileName = confirmId + (lang === 'en' ? '_Confirmation' : '_숙박확인서') + '.pdf';

  var url = 'https://docs.google.com/spreadsheets/d/' + ssId
    + '/export?exportFormat=pdf&format=pdf'
    + '&gid='          + gid
    + '&size=A4&portrait=true&fitw=true&gridlines=false'
    + '&top_margin=0.50&bottom_margin=0.50'
    + '&left_margin=0.50&right_margin=0.50'
    + '&sheetnames=false&printtitle=false'
    + '&pagenumbers=false&attachment=true';

  var token  = ScriptApp.getOAuthToken();
  var resp   = UrlFetchApp.fetch(url, {
    headers            : { Authorization: 'Bearer ' + token },
    muteHttpExceptions : true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error('PDF 내보내기 실패: HTTP ' + resp.getResponseCode());
  }

  var blob   = resp.getBlob().setName(fileName);
  var folder = getOrCreateFolder_(CONFIRM_FOLDER);   // ← Quote.gs 의 함수 사용
  var file   = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}


// ════════════════════════════════════════════════════════
//  6. DB 저장 / 조회
// ════════════════════════════════════════════════════════
function saveToDb_(confirmId, p, pdfUrl, lang) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIRM_DB_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIRM_DB_SHEET);
    sheet.appendRow([
      '발급번호','발급일시','투숙자명','연락처',
      '체크인','체크아웃','객실','금액','결제방식',
      'PDF_URL','전송상태','LANG'
    ]);
  }
  sheet.appendRow([
    confirmId,   new Date(),
    p.guestName  || '', p.phone    || '',
    p.checkIn    || '', p.checkOut || '',
    p.roomName   || '', p.totalAmt || 0,
    p.payMethod  || '', pdfUrl,
    '미전송',    lang
  ]);
}

function getConfirmData_(confirmId) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIRM_DB_SHEET);
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  var hdr  = data[0];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === confirmId) {
      var obj = {};
      hdr.forEach(function(h, j) { obj[h] = data[i][j]; });
      return obj;
    }
  }
  return null;
}


// ════════════════════════════════════════════════════════
//  7. SMS 전송  ─  sendSms_ 는 Code.gs 에서 호출
// ════════════════════════════════════════════════════════
function sendConfirmSms(confirmId) {
  var rec = getConfirmData_(confirmId);
  if (!rec) return { success: false, error: '확인서 데이터 없음' };
  var phone  = rec['연락처'] || '';
  var pdfUrl = rec['PDF_URL'] || '';
  if (!phone) return { success: false, error: '연락처 없음' };
  var msg = '[' + CFG.HOTEL_NAME + '] 숙박확인서가 발급되었습니다.\n' + pdfUrl;
  return sendSms_(phone, msg);   // ← Code.gs 의 sendSms_ 호출
}


// ════════════════════════════════════════════════════════
//  8. 유틸리티
//  ※ getOrCreateFolder_ → Quote.gs
//  ※ sendSms_           → Code.gs
// ════════════════════════════════════════════════════════
function today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function generateConfirmId_() {
  var d = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  var r = Math.random().toString(36).substring(2, 6).toUpperCase();
  return 'CF-' + d + '-' + r;
}

function formatAmt_(v) {
  return '₩ ' + (Number(v) || 0).toLocaleString() + ' 원';
}

function formatAmtEn_(v) {
  return 'KRW ' + (Number(v) || 0).toLocaleString();
}

function translatePayEn_(m) {
  return ({
    '법인카드' : 'Corporate Card',
    '개인카드' : 'Personal Card',
    '현금'    : 'Cash',
    '계좌이체' : 'Bank Transfer',
    '무통장'  : 'Bank Transfer'
  })[m] || m || 'Cash';
}
// ════════════════════════════════════════════════════════
//  1박별 요금 계산
//  우선순위: 가격표(loadPriceTable_) → 균등분배 → 단일 행
// ════════════════════════════════════════════════════════
function buildNightlyRows_(checkIn, checkOut, roomName, totalAmt, stayType) {
  var DOW_MAP = { 0:'일', 1:'월', 2:'화', 3:'수', 4:'목', 5:'금', 6:'토' };

  // 대실(당일) 처리
  if (stayType === '대실') {
    return [{ date: String(checkIn).substring(0, 10), dow: '', amount: Number(totalAmt) }];
  }

  // 날짜 파싱
  var ci = new Date(String(checkIn).substring(0, 10));
  var co = new Date(String(checkOut).substring(0, 10));
  if (isNaN(ci.getTime()) || isNaN(co.getTime()) || co <= ci) {
    return [{ date: String(checkIn).substring(0, 10), dow: '', amount: Number(totalAmt) }];
  }

  var nights = Math.round((co - ci) / 86400000);
  if (nights <= 0) {
    return [{ date: String(checkIn).substring(0, 10), dow: '', amount: Number(totalAmt) }];
  }

  // ── 가격표 기반 계산 시도 ──────────────────────────────
  try {
    var pt       = loadPriceTable_();
    var priceRow = getPriceRow_(pt.map, roomName);

    if (priceRow) {
      var rows = [];
      var cur  = new Date(ci);

      while (cur < co) {
        var dateStr = Utilities.formatDate(cur, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        var dow     = DOW_MAP[cur.getDay()];
        // 요금표에서 해당 요일 열 탐색
        var colKey  = Object.keys(priceRow).find(function(k) {
          return String(k).indexOf(dow) >= 0;
        });
        var amt = colKey ? (Number(priceRow[colKey]) || 0) : 0;
        rows.push({ date: dateStr, dow: dow, amount: amt });
        cur.setDate(cur.getDate() + 1);
      }

      // 가격표 합계와 실제 결제액이 다를 경우 → 비례 조정
      var calcTotal = rows.reduce(function(s, r) { return s + r.amount; }, 0);
      if (calcTotal > 0 && calcTotal !== Number(totalAmt)) {
        var ratio = Number(totalAmt) / calcTotal;
        rows.forEach(function(r) {
          r.amount = Math.round(r.amount * ratio);
        });
        // 반올림 오차 → 마지막 행에 보정
        var adjTotal = rows.reduce(function(s, r) { return s + r.amount; }, 0);
        rows[rows.length - 1].amount += (Number(totalAmt) - adjTotal);
      }

      return rows;
    }
  } catch (e) {
    Logger.log('buildNightlyRows_ 가격표 조회 실패: ' + e.message);
  }

  // ── 가격표 없을 경우 → 균등 분배 ──────────────────────
  var amtPerNight = Math.round(Number(totalAmt) / nights);
  var rows = [];
  var cur  = new Date(ci);

  for (var i = 0; i < nights; i++) {
    var dateStr = Utilities.formatDate(cur, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var dow     = DOW_MAP[cur.getDay()];
    var amt     = (i < nights - 1)
      ? amtPerNight
      : (Number(totalAmt) - amtPerNight * (nights - 1));  // 마지막 행 오차 보정
    rows.push({ date: dateStr, dow: dow, amount: amt });
    cur.setDate(cur.getDate() + 1);
  }

  return rows;
}

