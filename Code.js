/*******************************************************
 * Code.gs — Netflix SMS + 예약자동입력 + 견적
 *******************************************************/

const CFG = {
  ROOMS_SHEET:     'ROOMS',
  INBOX_SHEET:     'INBOX',
  PRICE_SHEET:     '가격표',
  DB_SHEET:        '견적DB',
  LOG_SHEET:       'SEND_LOG',
  RES_DB_SHEET:    'RES_DB',
  SENS_HOST:       'https://sens.apigw.ntruss.com',
  DOC_FOLDER_NAME: '견적서_PDF',
  INVOICE_SHEET:   '견적서_템플릿',
  HOTEL_NAME:      'HOTEL LABEL',
  HOTEL_ADDRESS:   '경기도 성남시 중원구 제일로 76',
  HOTEL_BIZ_NO:    '453-01-02738',
  HOTEL_CEO:       '우성열',
  HOTEL_TEL:       '031-757-6680, 010-4657-6680',
};

const USE_CHECKIN_TAB = true;
const CUSTOMER_START_ROW = 4;
const CUSTOMER_END_ROW = 45;

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('▶넷플릭스 ID비번보내기')
    .addItem('객실 계정 문자 발송', 'sendNetflix').addToUi();
  ui.createMenu('▶카톡예약입력하기')
    .addItem('카톡 붙여넣기 → 자동 입력', 'importKakao').addToUi();
  ui.createMenu('▶CSV예약파일 한번에입력하기')
    .addItem('CSV 예약 가져오기', 'importCsv').addToUi();
ui.createMenu('▶견적서 생성/이력관리')
  .addItem('견적서 생성', 'openQuoteSidebar')
  .addItem('이력관리', 'openQuoteHistorySidebar')
  .addToUi();
  ui.createMenu('▶숙박확인서')
    .addItem('숙박확인서 발급', 'openConfirmSidebar').addToUi();
}


let _CUST_DB_VALUES_CACHE = null;

function resetCustDbCache_() {
  _CUST_DB_VALUES_CACHE = null;
  _CUSTOMER_VISIT_MAP = null;
}

/**
 * 고객DB 전체 값을 1회 실행 동안만 메모리에 캐시
 */
function getCustDbValues_() {
  if (_CUST_DB_VALUES_CACHE) return _CUST_DB_VALUES_CACHE;

  const sh = getCustDbSheet_();
  if (!sh || sh.getLastRow() < 2) {
    _CUST_DB_VALUES_CACHE = [];
    return _CUST_DB_VALUES_CACHE;
  }

  _CUST_DB_VALUES_CACHE = sh.getDataRange().getValues();
  return _CUST_DB_VALUES_CACHE;
}



function setupConfig() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('setupConfig 시작');
}

function sendRoomNetflixSmsPrompt() {
  const ui = SpreadsheetApp.getUi();
  const roomRes = ui.prompt('객실번호 입력', '예: 302', ui.ButtonSet.OK_CANCEL);
  if (roomRes.getSelectedButton() !== ui.Button.OK) return;
  const room = roomRes.getResponseText().trim();
  const toRes = ui.prompt('수신자 전화번호 입력', '예: 01012345678', ui.ButtonSet.OK_CANCEL);
  if (toRes.getSelectedButton() !== ui.Button.OK) return;
  const to = toRes.getResponseText().trim().replace(/[^0-9]/g, '');
  const creds = getRoomCreds_(room);
  if (!creds) { ui.alert(`ROOMS 탭에서 객실 ${room}을 찾지 못했습니다.`); return; }
  const msg = `[${room}호 넷플릭스 안내]\nID: ${creds.id}\nPW: ${creds.pw}\n\n 로그인 후 이용가구 업데이트 메시지 나오면 이메일 인증하기 후 프론트로 전화 부탁드립니다.`;
  const result = sensSendSms_(to, msg);
  logSend_(room, to, result);
  if (result.ok) ui.alert(`발송 성공 (requestId: ${result.requestId || '-'})`);
  else ui.alert(`발송 실패: ${result.error || 'unknown'}`);
}

function getRoomCreds_(room) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CFG.ROOMS_SHEET);
  if (!sh) throw new Error(`Sheet not found: ${CFG.ROOMS_SHEET}`);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(room))
      return { id: String(values[i][1]).trim(), pw: String(values[i][2]).trim() };
  }
  return null;
}

function sensSendSms_(to, content) {
  const props = PropertiesService.getScriptProperties();
  const serviceId = props.getProperty('SENS_SERVICE_ID');
  const accessKey = props.getProperty('SENS_ACCESS_KEY');
  const secretKey = props.getProperty('SENS_SECRET_KEY');
  const from      = props.getProperty('SENS_FROM');

  Logger.log('SENS serviceId exists=' + !!serviceId);
  Logger.log('SENS accessKey exists=' + !!accessKey);
  Logger.log('SENS secretKey exists=' + !!secretKey);
  Logger.log('SENS from=' + from);

  if (!serviceId || !accessKey || !secretKey || !from)
    return { ok: false, error: '설정값 없음. 초기설정 먼저 실행하세요.' };
  const method = 'POST';
  const uri = `/sms/v2/services/${serviceId}/messages`;
  const url = `${CFG.SENS_HOST}${uri}`;
  const timestamp = String(Date.now());
  const signature = makeSignature_(method, uri, timestamp, accessKey, secretKey);
  const payload = { type: 'SMS', countryCode: '82', from, content, messages: [{ to }] };
  try {
    const resp = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json; charset=utf-8',
      payload: JSON.stringify(payload), muteHttpExceptions: true,
      headers: {
        'x-ncp-apigw-timestamp': timestamp,
        'x-ncp-iam-access-key': accessKey,
        'x-ncp-apigw-signature-v2': signature
      },
    });
    const code = resp.getResponseCode();
    const json = safeJsonParse_(resp.getContentText());
    if (code >= 200 && code < 300)
      return { ok: true, httpCode: code, requestId: json && json.requestId };
    return { ok: false, httpCode: code, error: (json && json.errorMessage) || resp.getContentText() };
  } catch (e) { return { ok: false, error: String(e) }; }
}

function makeSignature_(method, uri, timestamp, accessKey, secretKey) {
  const message = method + ' ' + uri + '\n' + timestamp + '\n' + accessKey;
  return Utilities.base64Encode(
    Utilities.computeHmacSha256Signature(message, secretKey, Utilities.Charset.UTF_8)
  );
}

function safeJsonParse_(s) { try { return JSON.parse(s); } catch (_) { return null; } }

function logSend_(room, to, result) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.LOG_SHEET) || ss.insertSheet(CFG.LOG_SHEET);
  sh.appendRow([new Date(), 'front', room, to,
    result.ok ? 'OK' : 'FAIL', result.requestId || '', result.ok ? '' : (result.error || '')]);
}


function autoPasteToLedger() {
  const ui = SpreadsheetApp.getUi();
  const pr = ui.prompt('예약 자동 입력', '카톡/예약 내용을 통째로 붙여넣고 OK를 누르세요.', ui.ButtonSet.OK_CANCEL);
  if (pr.getSelectedButton() !== ui.Button.OK) return;
  const text = String(pr.getResponseText() || '').trim();
  if (!text) { ui.alert('내용이 비어 있습니다.'); return; }
  pasteTextCore_(text);
}

function pasteInboxToLedger() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const inbox = ss.getSheetByName(CFG.INBOX_SHEET);
  if (!inbox) { ui.alert('INBOX 탭이 없습니다.'); return; }
  const text = String(inbox.getRange('A1').getValue() || '').trim();
  if (!text) { ui.alert('INBOX!A1이 비어있습니다.'); return; }
  pasteTextCore_(text);
}

function pasteTextCore_(text) {
  const ui = SpreadsheetApp.getUi();
  const blocks = splitReservationBlocks_(text);
  if (!blocks.length) { ui.alert('예약 텍스트 블록을 찾지 못했습니다.'); return; }
  const db = ensureResDb_();
  const START_ROW = 4;
  let inserted = 0, failed = 0;
  const usedTabs = new Set();
  for (const block of blocks) {
    const r = parseReservation_(block);
    const roomType = roomTypeFromBlock_(block);
    if (!r) continue;
    const amountNum = amountFromBlock_(block);
    const tIn  = extractHourOnly_(r.checkin);
    const tOut = extractHourOnly_(r.checkout);
    const ledger = pickLedgerSheet_(r.checkin);
    usedTabs.add(ledger.getName());
    try {
      if (r.stayType === '대실') {
        const row = findInsertRowByCheckin_(ledger, '대실', tIn, START_ROW);
        ledger.getRange(`C${row}`).setValue(tIn || '');
        ledger.getRange(`D${row}`).setValue(tOut || '');
        if (amountNum !== null) ledger.getRange(`E${row}`).setValue(amountNum);
        ledger.getRange(`F${row}`).setValue(channelToPay_(r.channel));
        ledger.getRange(`I${row}`).setValue(r.name || '');
        if (roomType) ledger.getRange(`J${row}`).setValue(decorateRoomTypeWithDates_(roomType, r.checkin, r.checkout));
        if (r.vehicle === '차량') ledger.getRange(`K${row}`).setValue('차');
        markInserted_(ledger.getRange(`B${row}`));
        const _checkinDt = (r.checkin instanceof Date) ? r.checkin : new Date(String(r.checkin));
        const visitDate_ = Utilities.formatDate(_checkinDt, 'Asia/Seoul', 'yyyy-MM-dd');
        const source_ = SpreadsheetApp.getActive().getName().replace('숙박일지_', '');
        addCustomerRecord_(r.name || '', ledger.getRange(`K${row}`).getValue(),
          visitDate_, '대실', ledger.getRange(`B${row}`).getValue(), amountNum,
          channelToPay_(r.channel), ledger.getRange(`L${row}`).getValue(),
          ledger.getRange(`J${row}`).getValue(), source_);
        attachNote_(ledger, row, 9, r.name || '', r.vehicle || '');
      } else {
        const row = findInsertRowByCheckin_(ledger, '숙박', tIn, START_ROW);
        ledger.getRange(`O${row}`).setValue(tIn || '');
        if (amountNum !== null) ledger.getRange(`P${row}`).setValue(amountNum).setNumberFormat('#,##0');
        ledger.getRange(`Q${row}`).setValue(channelToPay_(r.channel));
        ledger.getRange(`T${row}`).setValue(r.name || '');
        if (roomType) ledger.getRange(`U${row}`).setValue(decorateRoomTypeWithDates_(roomType, r.checkin, r.checkout));
        if (r.vehicle === '차량') ledger.getRange(`V${row}`).setValue('차');
        else if (r.vehicle === '도보') ledger.getRange(`V${row}`).setValue('-');
        markInserted_(ledger.getRange(`N${row}`));
        const _checkinDt = (r.checkin instanceof Date) ? r.checkin : new Date(String(r.checkin));
        const visitDate_ = Utilities.formatDate(_checkinDt, 'Asia/Seoul', 'yyyy-MM-dd');
        const source_ = SpreadsheetApp.getActive().getName().replace('숙박일지_', '');
        addCustomerRecord_(r.name || '', ledger.getRange(`V${row}`).getValue(),
          visitDate_, '숙박', ledger.getRange(`N${row}`).getValue(), amountNum,
          channelToPay_(r.channel), ledger.getRange(`W${row}`).getValue(),
          ledger.getRange(`U${row}`).getValue(), source_);
        attachNote_(ledger, row, 20, r.name || '', r.vehicle || '');
      }
      db.appendRow([new Date(), '', r.channel || '', r.stayType || '', r.bookingNo || '',
        r.property || '', r.roomType || '', r.name || '', r.safeNo || '',
        r.checkin || '', r.checkout || '', r.duration || '', r.amount || '',
        r.vehicle || '', block]);
      inserted++;
    } catch (e) {
      failed++;
      db.appendRow([new Date(), '', 'ERROR', 'WRITE_FAIL', '', '', '', '', '', '', '', '', '', '', String(e)]);
      SpreadsheetApp.getActive().toast(`❌ 입력 실패: ${e}`, '예약자동입력', 6);
    }
  }
  ui.alert(`완료!\n입력 탭: ${Array.from(usedTabs).join(', ') || '(없음)'}\n입력 ${inserted}건 / 실패 ${failed}건`);
}

function autoPasteToTodayLedger() { autoPasteToLedger(); }

function pickLedgerSheet_(checkinText) {
  const ss = SpreadsheetApp.getActive();
  if (USE_CHECKIN_TAB) {
    const tabName = ledgerTabNameFromCheckin_(checkinText);
    const sh = ss.getSheetByName(tabName);
    if (sh) return sh;
    throw new Error('숙박일지 탭을 찾을 수 없습니다: ' + tabName);
  }
  return ss.getSheetByName(getTodayTabName_()) || ss.getActiveSheet();
}

function markInserted_(cell) {
  const sh  = cell.getSheet();
  const row = cell.getRow();
  const col = cell.getColumn();
  cell.setNumberFormat('@').setValue('V')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setTextStyle(SpreadsheetApp.newTextStyle().setBold(true).build());
  const isDaeSil = (col === 2);
  const isSukBak = (col === 14);
  let colsToFormat = [];
  if (isDaeSil) colsToFormat = [2, 3, 4, 5, 6, 9, 10, 11];
  if (isSukBak) colsToFormat = [14, 15, 16, 17, 20, 21, 22];
  if (colsToFormat.length === 0) return;
  const ranges = groupConsecutiveCols_(colsToFormat).map(g => sh.getRange(row, g.start, 1, g.len));
  const style = SpreadsheetApp.newTextStyle().setBold(true).build();
  ranges.forEach(rng => {
    rng.setHorizontalAlignment('center').setVerticalAlignment('middle').setTextStyle(style);
  });
}

function groupConsecutiveCols_(cols) {
  const arr = cols.slice().sort((a, b) => a - b);
  const out = [];
  let start = arr[0], prev = arr[0];
  for (let i = 1; i < arr.length; i++) {
    const c = arr[i];
    if (c === prev + 1) { prev = c; }
    else { out.push({ start: start, len: prev - start + 1 }); start = prev = c; }
  }
  out.push({ start: start, len: prev - start + 1 });
  return out;
}

function getTodayTabName_() {
  const tz  = SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Seoul';
  const now = new Date();
  const day = Number(Utilities.formatDate(now, tz, 'd'));
  const dow = Utilities.formatDate(now, tz, 'u');
  const map = { '1': '월', '2': '화', '3': '수', '4': '목', '5': '금', '6': '토', '7': '일' };
  return `${day}(${map[dow] || ''})`;
}

function ensureResDb_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(CFG.RES_DB_SHEET);
  if (!sh) sh = ss.insertSheet(CFG.RES_DB_SHEET);
  if (sh.getLastRow() === 0)
    sh.appendRow(['ts', 'room', 'channel', 'stay_type', 'booking_no', 'property', 'room_type',
      'name', 'safe_no', 'checkin', 'checkout', 'nights_or_hours', 'amount', 'vehicle', 'raw']);
  return sh;
}

function channelToPay_(channel) {
  if (!channel) return '';
  const ch = String(channel).toLowerCase();
  if (ch.includes('nol') || ch.includes('야놀자')) return '야';
  if (ch.includes('여기어때')) return '여';
  if (ch.includes('꿀스테이')) return '꿀';
  return '';
}

function extractHourOnly_(s) {
  const m = String(s || '').match(/(\d{1,2}):(\d{2})/);
  if (!m) return '';
  const h = Number(m[1]);
  const min = Number(m[2]);
  if (!isFinite(h)) return '';
  return min === 30 ? `${h}시반` : `${h}시`;
}

function splitReservationBlocks_(text) {
  const t = String(text || '').replace(/\r/g, '\n').trim();
  if (!t) return [];
  if (t.includes('//')) return t.split('//').map(s => s.trim()).filter(Boolean);
  const markers = ['NOL', '[여기어때]', '여기어때', '착한 숙박앱 꿀스테이', '꿀스테이', '야놀자'];
  let parts = [t];
  markers.forEach(m => {
    const next = [];
    parts.forEach(p => {
      const sp = p.split(new RegExp(`(?=\\b${escapeRegExp_(m)}\\b)`, 'g')).map(s => s.trim()).filter(Boolean);
      next.push(...sp);
    });
    parts = next;
  });
  return parts.length ? parts : [t];
}

function escapeRegExp_(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

function parseReservation_(block) {
  const b = String(block || '').replace(/\s+/g, ' ').trim();
  if (!b) return null;
  const res = {};
  if (b.includes('NOL')) res.channel = 'NOL';
  else if (b.includes('[여기어때]') || b.includes('여기어때')) res.channel = '여기어때';
  else if (b.includes('야놀자')) res.channel = '야놀자';
  else if (b.includes('꿀스테이')) res.channel = '꿀스테이';
  else res.channel = '';
  if (b.includes('대실') || /\(\s*\d+\s*시간\s*\)/.test(b) || (b.includes('퇴실일시') && b.includes('시간'))) res.stayType = '대실';
  else res.stayType = '숙박';
  let m = b.match(/예약번호[:\s]*([0-9A-Z]+)\b/);
  if (m) res.bookingNo = m[1];
  else { m = b.match(/\b(2\d{10,})\b/); if (m) res.bookingNo = m[1]; }
  m = b.match(/(성남\s*호텔\s*레이블)/); if (m) res.property = m[1];
  m = b.match(/객실정보:\s*[^/]+\/\s*([^(]+)\(/);
  if (m) res.roomType = m[1].trim();
  else { m = b.match(/\b(스탠다드|디럭스|스위트|프리미엄|도보특가)\b/); res.roomType = m ? m[1].trim() : ''; }
  m = b.match(/([가-힣A-Za-z][가-힣A-Za-z .'-]{1,30})\s*\/\s*(0\d[\d-]{8,15})/);
  if (m) { res.name = m[1]; res.safeNo = m[2].replace(/[^0-9]/g, ''); }
  if (!res.name) {
    const mn = b.match(/예약자명\s*[:]*\s*([^\n\r]+)/);
    if (mn) res.name = mn[1].split('안심번호')[0].split('방문방법')[0].split('*')[0].trim();
  }
  if (!res.safeNo) { const mp = b.match(/안심번호\s*[:]*\s*(0\d[\d-]{8,15})/); if (mp) res.safeNo = mp[1].replace(/[^0-9]/g, ''); }
  if (!res.safeNo) { m = b.match(/\(\s*(0\d[\d-]{8,15})\s*\)/); if (m) res.safeNo = m[1].replace(/[^0-9]/g, ''); }
  if (!res.name) { m = b.match(/([가-힣A-Za-z][가-힣A-Za-z .'-]{1,30})\s*\(\s*0\d[\d-]{8,15}\s*\)/); if (m) res.name = m[1]; }
  m = b.match(/(\d{4}-\d{2}-\d{2}\([^\)]*\)\s*\d{1,2}:\d{2})\s*~\s*(\d{4}-\d{2}-\d{2}\([^\)]*\)\s*\d{1,2}:\d{2})/);
  if (m) { res.checkin = m[1]; res.checkout = m[2]; }
  else {
    m = b.match(/(\d{1,2}\/\d{1,2}\([^\)]*\)\s*\d{1,2}:\d{2})\s*~\s*(\d{1,2}\/\d{1,2}\([^\)]*\)\s*\d{1,2}:\d{2})/);
    if (m) { res.checkin = addYearIfMissing_(m[1]); res.checkout = addYearIfMissing_(m[2]); }
    const ci = b.match(/입실일시[:\s]*([0-9]{1,2}\/[0-9]{1,2}\s*\([^\)]*\)\s*\d{1,2}:\d{2})/);
    const co = b.match(/퇴실일시[:\s]*([0-9]{1,2}\/[0-9]{1,2}\s*\([^\)]*\)\s*\d{1,2}:\d{2})/);
    if (ci) res.checkin  = addYearIfMissing_(ci[1].replace(/\s+/g, ''));
    if (co) res.checkout = addYearIfMissing_(co[1].replace(/\s+/g, ''));
  }
  m = b.match(/\((\d+\s*(박|시간))\)/); if (m) res.duration = m[1].replace(/\s+/g, '');
  if (b.includes('차량')) res.vehicle = '차량';
  else if (b.includes('도보')) res.vehicle = '도보';
  else res.vehicle = '';
  const amt = amountFromBlock_(block); if (amt !== null) res.amount = String(amt) + '원';
  if (!(res.name || res.checkin || res.bookingNo || res.amount)) return null;
  if (res.name) res.name = String(res.name).replace(/^원\s*/, '').trim();
  return res;
}

function addYearIfMissing_(md) {
  const s = String(md || '').trim().replace(/\s+/g, '');
  const m = s.match(/(\d{1,2})\/(\d{1,2})\(([월화수목금토일])\)(\d{1,2}:\d{2})/);
  if (!m) return md;
  const mm = String(m[1]).padStart(2, '0'), dd = String(m[2]).padStart(2, '0');
  const hhmm = m[4].length === 4 ? ('0' + m[4]) : m[4];
  return `2026-${mm}-${dd}(${m[3]}) ${hhmm}`;
}

function amountFromBlock_(block) {
  const b = String(block || '');
  let m = b.match(/판매금액\s*[:]*\s*([0-9,]+)\s*원?/);
  if (!m) m = b.match(/판매가\s*[:]*\s*([0-9,]+)/);
  if (!m) m = b.match(/\b([0-9]{1,3}(?:,[0-9]{3})+)\s*원?\b/);
  if (!m) m = b.match(/\b([0-9]{4,7})\s*원\b/);
  if (!m) return null;
  return Number(String(m[1]).replace(/,/g, ''));
}

function ledgerTabNameFromCheckin_(checkinText) {
  const s = String(checkinText || '').trim();
  let m = s.match(/(\d{4})-(\d{2})-(\d{2})\(([월화수목금토일])\)/);
  if (m) return `${Number(m[3])}(${m[4]})`;
  m = s.match(/(\d{1,2})\/(\d{1,2})\(([월화수목금토일])\)/);
  if (m) return `${Number(m[2])}(${m[3]})`;
  m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (m) {
    const dt = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    const wk = ['일', '월', '화', '수', '목', '금', '토'];
    return `${Number(m[3])}(${wk[dt.getDay()]})`;
  }
  return getTodayTabName_();
}

function roomTypeFromBlock_(block) {
  const raw = String(block || '').trim();
  if (!raw) return '';
  const lines = raw.replace(/\r/g, '\n').split('\n').map(s => s.trim()).filter(Boolean);
  if (lines.length >= 2) {
    const priceIdx = lines.findIndex(l => /[0-9]{1,3}(?:,[0-9]{3})+\s*원$/.test(l) || /[0-9]{4,7}\s*원$/.test(l));
    if (priceIdx > 0) {
      const c = lines[priceIdx - 1];
      if (!/^NOL/.test(c) && !/^\d{10,}$/.test(c) && !c.includes('호텔')) return c.trim();
    }
  }
  const one = raw.replace(/\s+/g, ' ');
  let mi = one.match(/객실정보\s*:\s*[^/]+\/\s*([^(]+)/);
  if (mi) return normalizeRoomType_(mi[1].trim());
  let m = one.match(/([가-힣A-Za-z0-9\s\-\&\/]+?)\s*(?:\([^)]*\)\s*)+[0-9]{1,3}(?:,[0-9]{3})+\s*원/);
  if (m) return normalizeRoomType_(m[1].trim());
  m = one.match(/([가-힣A-Za-z0-9\s\-\&\/]+?)\s*[0-9]{1,3}(?:,[0-9]{3})+\s*원/);
  if (m) return cleanupRoomType_(m[1].trim());
  const keywords = ['그랜드 스위트', '시그니처 스위트', '스위트 - 투베드', '스위트 - 킹베드',
    '디럭스 - 빠른입실 늦은퇴실', '스탠다드 - 빠른입실 늦은퇴실', '도보특가 - 주차불가', '스위트', '디럭스', '스탠다드'];
  for (const k of keywords) if (one.includes(k)) return k;
  return '';
}

function cleanupRoomType_(s) {
  let t = String(s || '').replace(/\s+/g, ' ').trim()
    .replace(/^NOL\s*미리예약\s*/i, '').replace(/^NOL\s*당일예약\s*/i, '')
    .replace(/<\s*연박\s*>/g, '').replace(/<\s*숙박\s*>/g, '')
    .replace(/성남\s*호텔\s*레이블/g, '').replace(/\b2\d{10,}\b/g, '')
    .replace(/\s+/g, ' ').trim();
  return t.length > 60 ? t.slice(0, 60).trim() : t;
}

function dayWkFromDateText_(s) {
  const str = String(s || '');
  let m = str.match(/\d{4}-\d{2}-(\d{2})\(([월화수목금토일])\)/);
  if (m) return `${Number(m[1])}(${m[2]})`;
  m = str.match(/\d{1,2}\/(\d{1,2})\(([월화수목금토일])\)/);
  if (m) return `${Number(m[1])}(${m[2]})`;
  return '';
}

function decorateRoomTypeWithDates_(roomType, checkinText, checkoutText) {
  const base = String(roomType || '').trim();
  if (!base) return '';
  const inDW  = dayWkFromDateText_(checkinText);
  const outDW = dayWkFromDateText_(checkoutText);
  if (inDW && outDW) return `${base}-${inDW}입실/${outDW}퇴실`;
  if (inDW) return `${base}-${inDW}입실`;
  return base;
}

function normalizeRoomType_(name) {
  if (name.includes('시그니처')) return '시그니처 스위트';
  if (name.includes('그랜드'))   return '그랜드 스위트';
  if (name.includes('디럭스'))   return '디럭스';
  if (name.includes('스탠다드')) return '스탠다드';
  return name;
}

function openQuoteSidebar() {
  const tpl = HtmlService.createTemplateFromFile('Sidebar');
  tpl.initialQuoteData = 'null';

  const html = tpl.evaluate()
    .setWidth(960)
    .setHeight(820);

  SpreadsheetApp.getUi().showModalDialog(html, '견적서 작성');
}

function openConfirmSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ConfirmSidebar').setTitle('숙박확인서 발급').setWidth(700);
  SpreadsheetApp.getUi().showSidebar(html);
}

function initSensConfig() {
  PropertiesService.getScriptProperties().setProperties({
    'SENS_SERVICE_ID': 'ncp:sms:kr:367091155488:neflix_sms',
    'SENS_ACCESS_KEY': '...',
    'SENS_SECRET_KEY': '...',
    'SENS_FROM': '01046576680'
  });
  Logger.log('✅ SENS 설정 저장 완료');
}

function checkSensProps() {
  const props = PropertiesService.getScriptProperties().getProperties();
  Logger.log(JSON.stringify(props));
}

function sendNetflix()    { sendRoomNetflixSmsPrompt(); }
function initNetflixKey() { setupConfig(); }
function importKakao()    { autoPasteToLedger(); }
function importInbox()    { pasteInboxToLedger(); }

function importCsv() {
  const html = HtmlService.createHtmlOutputFromFile('CsvPicker').setWidth(500).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'CSV 예약 가져오기');
}

function importCsvText(csvText, fileName) {
  const ss   = SpreadsheetApp.getActive();
  const tz   = ss.getSpreadsheetTimeZone() || 'Asia/Seoul';
  const rows = Utilities.parseCsv(String(csvText || ''));
  if (!rows || rows.length < 2) throw new Error('CSV에 데이터가 없습니다.');
  const header = rows[0].map(h => String(h || '').trim());
  const idx    = makeHeaderIndex_(header);
  const resDb  = ensureResDb_();
  const startRow = resDb.getLastRow() + 1;
  const rawDump  = rows.slice(1).map(r => [JSON.stringify(rowToObj_(r, idx))]);
  if (rawDump.length) resDb.getRange(startRow, 1, rawDump.length, 1).setValues(rawDump);
  let inserted = 0, failed = 0;
  const usedTabs = new Set();
  const validRows = [];
  for (let i = 1; i < rows.length; i++) {
    const r0 = rows[i];
    if (!r0 || r0.join('').trim() === '') continue;
    const obj = rowToObj_(r0, idx);
    if (!obj.checkin || !obj.checkout) continue;
    if (obj.payStatus === '취소됨') continue;
    validRows.push(obj);
  }
  validRows.sort((a, b) => String(a.checkin).localeCompare(String(b.checkin)));
  for (const obj of validRows) {
    try {
      const stayType    = normalizeStayType_(obj.resType);
      const ledger      = pickLedgerSheetByIso_(ss, tz, obj.checkin);
      usedTabs.add(ledger.getName());
      const tIn         = extractHourOnly_(obj.checkin);
      const tOut        = extractHourOnly_(obj.checkout);
      const amountNum   = toNumber_(obj.payAmount || obj.totalAmount);
      const pay         = channelToPayFromCsv_(obj.payRoute) || channelToPayFromCsv_(obj.payMethod);
      const name        = obj.payerName || '';
      const roomTypeClean = normalizeRoomType_(obj.roomType || obj.reserveRoomType || '');
      const START_ROW   = 4;
      let row;
      if (stayType === '대실') {
        row = findInsertRowByCheckin_(ledger, '대실', tIn, START_ROW);
        ledger.getRange(`C${row}`).setValue(tIn || '');
        ledger.getRange(`D${row}`).setValue(tOut || '');
        if (amountNum !== null) ledger.getRange(`E${row}`).setValue(amountNum);
        ledger.getRange(`F${row}`).setValue(pay);
        ledger.getRange(`I${row}`).setValue(name);
        if (roomTypeClean) ledger.getRange(`J${row}`).setValue(roomTypeClean);
        markInserted_(ledger.getRange(`B${row}`));
      } else {
        row = findInsertRowByCheckin_(ledger, '숙박', tIn, START_ROW);
        ledger.getRange(`O${row}`).setValue(tIn || '');
        if (amountNum !== null) ledger.getRange(`P${row}`).setValue(amountNum).setNumberFormat('#,##0');
        ledger.getRange(`Q${row}`).setValue(pay);
        ledger.getRange(`T${row}`).setValue(name);
        if (roomTypeClean) ledger.getRange(`U${row}`).setValue(roomTypeClean);
        markInserted_(ledger.getRange(`N${row}`));
      }
      inserted++;
    } catch (e) { failed++; }
  }
  SpreadsheetApp.getActive().toast(`CSV 입력 완료: ${inserted}행 (실패 ${failed})`, 'CSV예약파일 입력', 5);
  return { ok: true, inserted, failed, tabs: Array.from(usedTabs), resDbA1: `A${startRow}` };
}

function makeHeaderIndex_(headerArr) {
  const m = {};
  headerArr.forEach((h, i) => { m[h] = i; });
  return m;
}

function rowToObj_(row, idx) {
  const v = (name) => { const j = idx[name]; return (j === undefined) ? '' : String(row[j] || '').trim(); };
  return {
    checkin:         v('입실일시'),
    checkout:        v('퇴실일시'),
    payAt:           v('결제일시'),
    resType:         v('예약타입'),
    reserveRoomType: v('예약 객실 타입'),
    roomType:        v('객실타입'),
    assignedRoom:    v('배정호실'),
    payerName:       v('결제자명'),
    totalAmount:     v('총결제금액'),
    payAmount:       v('결제금액'),
    unpaid:          v('미수금'),
    payRoute:        v('결제경로'),
    payMethod:       v('결제방식'),
    payStatus:       v('결제상태'),
    txId:            v('거래번호'),
    extraPay:        v('부속결제'),
  };
}

function normalizeStayType_(s) {
  const t = String(s || '');
  if (t.includes('대실')) return '대실';
  return '숙박';
}

function pickLedgerSheetByIso_(ss, tz, checkinText) {
  const m = String(checkinText || '').match(/(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return ss.getSheetByName(getTodayTabName_()) || ss.getActiveSheet();
  const y = Number(m[1]), mo = Number(m[2]) - 1, day = Number(m[3]);
  const dt = new Date(y, mo, day);
  const wk = ['일', '월', '화', '수', '목', '금', '토'];
  const name = `${day}(${wk[dt.getDay()]})`;
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('숙박일지 탭 없음: ' + name);
  return sh;
}

function parseIsoDate_(s) {
  const m = String(s || '').match(/(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  const y = Number(m[1]), mo = Number(m[2]) - 1, d = Number(m[3]);
  const dt = new Date(y, mo, d);
  return isNaN(dt.getTime()) ? null : dt;
}

function toNumber_(s) {
  const n = Number(String(s || '').replace(/[^0-9]/g, ''));
  return isFinite(n) && n > 0 ? n : null;
}

function channelToPayFromCsv_(route) {
  const ch = String(route || '').toLowerCase();
  if (ch.includes('야놀자') || ch.includes('nol')) return '야';
  if (ch.includes('여기어때')) return '여';
  if (ch.includes('꿀스테이')) return '꿀';
  if (ch.includes('아고다'))   return '아고다';
  if (ch.includes('booking') || ch.includes('부킹')) return '부킹';
  if (ch.includes('expedia') || ch.includes('익스')) return '익스';
  return '';
}

function parseTimeToMin_(t) {
  const s = String(t || '').trim();
  const m1 = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m1) return Number(m1[1]) * 60 + Number(m1[2]);
  const m2 = s.match(/^(\d{1,2})시(\d{0,2})분?$/);
  if (m2) return Number(m2[1]) * 60 + (Number(m2[2]) || 0);
  return -1;
}

function debugTabName() {
  const testDate = '2026-03-09';
  Logger.log('탭 이름 결과: ' + csvTabNameFromCheckin_(testDate));
}

function csvTabNameFromCheckin_(checkinText) {
  const s = String(checkinText || '').trim();
  const m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (!m) throw new Error('입실일시 해석 실패: ' + s);
  const dt = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const wk = ['일', '월', '화', '수', '목', '금', '토'];
  return Number(m[3]) + '(' + wk[dt.getDay()] + ')';
}

// ── 행 삽입 없이 시간 기반으로 올바른 행 탐색 ──
function findInsertRowByCheckin_(sheet, stayType, newTime, startRow) {
  const vals = sheet.getDataRange().getValues();
  if (stayType === '대실') {
    for (let i = startRow - 1; i < vals.length; i++) {
      const c = String(vals[i][2] || '').trim();
      const iName = String(vals[i][8] || '').trim();
      if (c === newTime && iName === '') return i + 1;
    }
    for (let i = startRow - 1; i < vals.length; i++) {
      const c = String(vals[i][2] || '').trim();
      const iName = String(vals[i][8] || '').trim();
      if (c === '' && iName === '') return i + 1;
    }
  } else {
    for (let i = startRow - 1; i < vals.length; i++) {
      const o = String(vals[i][14] || '').trim();
      const tName = String(vals[i][19] || '').trim();
      if (o === newTime && tName === '') return i + 1;
    }
    for (let i = startRow - 1; i < vals.length; i++) {
      const o = String(vals[i][14] || '').trim();
      const tName = String(vals[i][19] || '').trim();
      if (o === '' && tName === '') return i + 1;
    }
  }
  return vals.length + 1;
}

// ══════════════════════════════════════════════════
//  고객DB
// ══════════════════════════════════════════════════
const CUST_DB_SS_ID = '10LHCt-t7JBxjLHVntNDw6vQsMoRx5EEs1Wm22K2ybFU';
const CUST_DB_TAB   = '고객기록';

function getCustDbSheet_() {
  const ss = SpreadsheetApp.openById(CUST_DB_SS_ID);
  const sh = ss.getSheetByName(CUST_DB_TAB);
  if (!sh) throw new Error('고객DB 탭 없음: ' + CUST_DB_TAB);
  return sh;
}

function addCustomerRecord_(name, vehicle, visitDate, stayType,
                             markerVal, amount, channel,
                             assignedRoom, extraInfo, source) {
  if (!name) return;

  try {
    const sh = getCustDbSheet_();
    const dateStr = String(visitDate || '').slice(0, 10);
    const cleanName = String(name || '').trim();
    const roomStr = String(assignedRoom || '').trim();
    const extraStr = String(extraInfo || '').trim();

    const values = sh.getDataRange().getValues();

    // 고객명 + 방문일 = 1건
    for (let i = 1; i < values.length; i++) {
      const rowName = String(values[i][0] || '').trim();
      const rowDateRaw = values[i][2];
      const rowDate = rowDateRaw instanceof Date
        ? Utilities.formatDate(rowDateRaw, 'Asia/Seoul', 'yyyy-MM-dd')
        : String(rowDateRaw || '').slice(0, 10);

      if (rowName === cleanName && rowDate === dateStr) {
        sh.getRange(i + 1, 1, 1, 10).setValues([[
          cleanName,
          vehicle || '',
          dateStr,
          stayType || '',
          markerVal || '',
          amount || '',
          channel || '',
          roomStr,
          extraStr,
          source || ''
        ]]);
        return;
      }
    }

    sh.appendRow([
      cleanName,
      vehicle || '',
      dateStr,
      stayType || '',
      markerVal || '',
      amount || '',
      channel || '',
      roomStr,
      extraStr,
      source || ''
    ]);

  } catch (e) {
    Logger.log('addCustomerRecord_ 오류: ' + e);
  }
}

function attachNote_(ledger, row, col, name, vehicle) {
  if (!name) return;

  try {
    const cell = ledger.getRange(row, col);
    const note = buildCustomerNote_(name);
    cell.setNote(note);

    // 색상 우선순위: 주의 > 단골/VIP > 재방문 > 첫방문
    if (note && note.includes('⚠️ 주의 고객')) {
      cell.setBackground('#F4B183'); // 주황
    } else if (note && (note.includes('⭐ 단골 고객') || note.includes('👑 VIP 고객'))) {
      cell.setBackground('#FFD700'); // 금색
    } else if (note && note.includes('재방문 고객')) {
      cell.setBackground('#FFF9C4'); // 연노랑
    } else {
      cell.setBackground(null);
    }

  } catch (e) {
    Logger.log('attachNote_ 오류: ' + e);
  }
}


function buildCustomerNote_(name) {
  if (!name) return '';

  try {
    const KST = 'Asia/Seoul';
    const map = getCustomerVisitMap_();
    if (!map) return '';

    const todayStr = Utilities.formatDate(new Date(), KST, 'yyyy-MM-dd');
    const todayDate = parseIsoDate_(todayStr);

    const rows = map[String(name).trim()] || [];

    const mapped = rows
      .filter(r => String(r[3] || '').trim() !== '취소')
      .map(r => {
        const raw = r[2];
        const dateStr = raw instanceof Date
          ? Utilities.formatDate(raw, KST, 'yyyy-MM-dd')
          : String(raw || '').slice(0, 10);

        return {
          dateStr: dateStr,
          date: parseIsoDate_(dateStr),
          vehicle: String(r[1] || '').trim(), // 고객DB 2열: 차번호/차량정보
          room: String(r[7] || '').trim(),    // 고객DB 8열: 객실
          type: String(r[3] || '').trim(),    // 고객DB 4열: 숙박/대실
          amount: Number(r[5]) || 0,          // 고객DB 6열: 금액
          note: String(r[8] || '').trim()     // 고객DB 9열: 특이사항/차종
        };
      })
      .filter(x => x.date);

    // 주의 고객 판정은 전체 이력 기준
    const warningHit = mapped.find(x =>
      /주의|블랙|블랙리스트|진상/.test(String(x.note || ''))
    );
    const warningBadge = warningHit ? '⚠️ 주의 고객' : '';

    // 오늘 제외한 과거 방문만 표시용으로 사용
    const pastOnly = mapped.filter(x => x.dateStr < todayStr);

    // 최신순
    pastOnly.sort((a, b) => b.date - a.date);

    // 같은 날짜 중복 제거
    const seen = new Set();
    const deduped = [];
    pastOnly.forEach(x => {
      if (seen.has(x.dateStr)) return;
      seen.add(x.dateStr);
      deduped.push(x);
    });

    // 과거 방문 없으면 첫 방문
    if (deduped.length === 0) {
      return [
        '👤 ' + name,
        '─────────────────',
        warningBadge,
        '🎉 오늘이 첫 방문입니다!'
      ].filter(Boolean).join('\n');
    }

    const total = deduped.length;
    const lastVisit = deduped[0].dateStr;
    const lastVisitDate = parseIsoDate_(lastVisit);

    // 등급
    let gradeLabel = '재방문 고객';
    if (total >= 10) {
      gradeLabel = '⭐ 단골 고객';
    } else if (total >= 5) {
      gradeLabel = '👑 VIP 고객';
    }

    // 최근 방문 배지
    let recentBadge = '';
    if (todayDate && lastVisitDate) {
      const diffDays = Math.floor((todayDate - lastVisitDate) / (1000 * 60 * 60 * 24));
      if (diffDays <= 7) {
        recentBadge = '🟢 최근 7일 내 방문';
      } else if (diffDays <= 30) {
        recentBadge = '🟢 최근 30일 내 방문';
      }
    }

    // 평균 결제금액
    const avgAmount = total > 0
      ? Math.round(deduped.reduce((sum, v) => sum + (v.amount || 0), 0) / total)
      : 0;
    const avgAmountLine = avgAmount > 0 ? ('평균 결제: ' + avgAmount.toLocaleString() + '원') : '';

    // 최근 3회 표시
const recentLines = deduped.slice(0, 3).map(v => {
  const parts = [v.dateStr];

  if (v.room) parts.push(v.room + '호');
  if (v.type) parts.push(v.type);
  if (v.vehicle) parts.push('차:' + v.vehicle);

  if (v.amount) {
    parts.push(v.amount.toLocaleString() + '원');
  }

  if (v.note) parts.push('메모:' + v.note);

  return parts.join(' / ');
});

    return [
      '👤 ' + name,
      '─────────────────',
      warningBadge,
      gradeLabel,
      '총 ' + total + '회',
      '마지막 방문: ' + lastVisit,
      recentBadge,
      avgAmountLine,
      '최근 이력:',
      recentLines.join('\n')
    ].filter(Boolean).join('\n');

  } catch (e) {
    Logger.log('buildCustomerNote_ 오류: ' + e);
    return '';
  }
}


function formatKorDate_(d) {
  if (!(d instanceof Date) || isNaN(d)) return '-';
  return `${d.getFullYear()} ${d.getMonth()+1}월 ${d.getDate()}일`;
}

function onEdit(e) {
  const range = e.range;
  const col   = range.getColumn();
  if (col === 10 || col === 11 || col === 21 || col === 22) {
    const newVal = String(e.range.getValue()||'').trim();
    if (!newVal) return;
    const sheet = range.getSheet();
    const row   = range.getRow();
    const targetDateStr = getDateStrFromSheet_(sheet.getName());
    if (!targetDateStr) return;
    const nameCol = (col===10||col===11) ? 9 : 20;
    const name    = String(sheet.getRange(row, nameCol).getValue()||'').trim();
    if (!name) return;
    const dbCol = (col===11||col===22) ? 2 : 9;
    try {
      updateFieldInDb_(name, targetDateStr, dbCol, newVal);
      const carCol = (col===10||col===11) ? 11 : 22;
      const car    = String(sheet.getRange(row, carCol).getValue()||'').trim();
      attachNote_(sheet, row, nameCol, name, car);
    } catch (err) { Logger.log('[onEdit] 오류: ' + err); }
    return;
  }
if (col === 9 || col === 20) {
  const newVal = String(e.range.getValue() || '').trim();
  const oldVal = String(e.oldValue || '').trim();
  const sheet = range.getSheet();
  const row   = range.getRow();

  if (row < CUSTOMER_START_ROW || row > CUSTOMER_END_ROW) return;

  if (!newVal) {
    e.range.clearNote();
    e.range.setBackground(null);

    if (oldVal) {
      cancelCustomerRecordByKey_(oldVal, getDateStrFromSheet_(sheet.getName()));
    }
    return;
  }

  e.range.clearNote();
  e.range.setBackground(null);

  syncCustomerRecordFromRow_(sheet, row, col);
  attachNote_(sheet, row, col, newVal);
  return;
}

  if (col !== 2 && col !== 14) return;
  const newVal = String(e.range.getValue()||'').trim();
  if (!newVal) {
    const sheet    = range.getSheet();
    const row      = range.getRow();
    const nameCol  = (col===2) ? 9 : 20;
    const nameCell = sheet.getRange(row, nameCol);
    nameCell.clearNote(); nameCell.setBackground(null);
    return;
  }
  if (newVal === 'V') return;
  const sheet   = range.getSheet();
  const row     = range.getRow();
  const nameCol = (col===2) ? 9 : 20;
  const name    = String(sheet.getRange(row, nameCol).getValue()||'').trim();
  if (!name) return;
  const targetDateStr = getDateStrFromSheet_(sheet.getName());
  if (!targetDateStr) return;
  try {
    updateRoomInDb_(name, targetDateStr, newVal);
    attachNote_(sheet, row, nameCol, name);
  } catch (err) { Logger.log('[onEdit] 오류: ' + err); }
}

function updateRoomInDb_(name, targetDateStr, newRoom) {
  if (!name || !newRoom) return;
  const sh   = getCustDbSheet_();
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return;
  for (let i = 1; i < data.length; i++) {
    const dbName = String(data[i][0]||'').trim();
    const rawDate = data[i][2];
    const dbDate = (rawDate instanceof Date)
      ? Utilities.formatDate(rawDate, 'Asia/Seoul', 'yyyy-MM-dd')
      : String(rawDate||'').trim();
    if (dbName !== name) continue;
    if (!dbDate.startsWith(targetDateStr)) continue;
    if (newRoom === 'V') continue;
    sh.getRange(i+1, 5).setValue(newRoom);
  }
}

function cancelCustomerRecord_(sheet, row, col) {
  try {
    const sh = getCustDbSheet_();

    const name = String(sheet.getRange(row, col).getValue() || '').trim();
    if (!name) return;

    const visitDate = getDateStrFromSheet_(sheet.getName());
    if (!visitDate) return;

    const values = sh.getDataRange().getValues();

    for (let i = values.length - 1; i >= 1; i--) {
      const rowName = String(values[i][0] || '').trim();
      const rowDateRaw = values[i][2];
      const rowDate = rowDateRaw instanceof Date
        ? Utilities.formatDate(rowDateRaw, 'Asia/Seoul', 'yyyy-MM-dd')
        : String(rowDateRaw || '').slice(0, 10);

      if (rowName === name && rowDate === visitDate) {
        sh.deleteRow(i + 1);
        return;
      }
    }

  } catch (e) {
    Logger.log('cancelCustomerRecord_ 오류: ' + e);
  }
}

function cancelCustomerRecordByKey_(name, visitDate) {
  try {
    if (!name || !visitDate) return;

    const sh = getCustDbSheet_();
    const values = sh.getDataRange().getValues();

    for (let i = values.length - 1; i >= 1; i--) {
      const rowName = String(values[i][0] || '').trim();
      const rowDateRaw = values[i][2];
      const rowDate = rowDateRaw instanceof Date
        ? Utilities.formatDate(rowDateRaw, 'Asia/Seoul', 'yyyy-MM-dd')
        : String(rowDateRaw || '').slice(0, 10);

      if (rowName === name && rowDate === visitDate) {
        sh.deleteRow(i + 1);
        return;
      }
    }
  } catch (e) {
    Logger.log('cancelCustomerRecordByKey_ 오류: ' + e);
  }
}

function updateFieldInDb_(name, targetDateStr, dbCol, newVal) {
  if (!name || !newVal) return;
  const sh   = getCustDbSheet_();
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return;
  for (let i = 1; i < data.length; i++) {
    const dbName  = String(data[i][0]||'').trim();
    const rawDate = data[i][2];
    const dbDate  = (rawDate instanceof Date)
      ? Utilities.formatDate(rawDate, 'Asia/Seoul', 'yyyy-MM-dd')
      : String(rawDate||'').trim();
    if (dbName !== name) continue;
    if (!dbDate.startsWith(targetDateStr)) continue;
    sh.getRange(i+1, dbCol).setValue(newVal);
  }
}

function getDateStrFromSheet_(sheetName) {
  const mA = sheetName.match(/(\d+)월(\d+)일/);
  if (mA) {
    const today = new Date();
    const d = new Date(today.getFullYear(), Number(mA[1])-1, Number(mA[2]));
    return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  const mB = sheetName.match(/^(\d+)\(/);
  if (mB) {
    const today = new Date();
    const d = new Date(today.getFullYear(), today.getMonth(), Number(mB[1]));
    return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  return null;
}

function restoreCustomerRecord_(sheet, row, col, name) {
  try {
    const targetDateStr = getDateStrFromSheet_(sheet.getName());
    if (!targetDateStr) return;
    const stayType = (col===9) ? '대실' : '숙박';
    const sh       = getCustDbSheet_();
    const data     = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const dbName  = String(data[i][0]||'').trim();
      const rawDate = data[i][2];
      const dbDate  = (rawDate instanceof Date)
        ? Utilities.formatDate(rawDate, 'Asia/Seoul', 'yyyy-MM-dd')
        : String(rawDate||'').trim();
      const dbType = String(data[i][3]||'').trim();
      if (dbName !== name) continue;
      if (!dbDate.startsWith(targetDateStr)) continue;
      if (dbType !== '취소') continue;
      sh.getRange(i+1, 4).setValue(stayType);
      return;
    }
  } catch (e) { Logger.log('[restoreCustomerRecord_] 오류: ' + e); }
}

// ══════════════════════════════════════════════════
//  findReservationByName — 숙박확인서 사이드바용
//  ✅ 파일명 "숙박일지(2603)" 형식 지원
// ══════════════════════════════════════════════════
// ══════════════════════════════════════════════════
//  findReservationByName — 숙박확인서 사이드바용
//  ✅ 단일 루프로 최적화: getValues() 호출 횟수 절반으로 감소
// ══════════════════════════════════════════════════
function findReservationForConfirm(name, fromDateStr, toDateStr) {

  Logger.log('findReservationByName 호출됨');
  Logger.log('name=[' + name + '] from=[' + fromDateStr + '] to=[' + toDateStr + ']');

  if (!name) return null;

  name = String(name).trim();
  Logger.log('trim 후 name=[' + name + ']');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssName = ss.getName();

const mmMatch = ssName.match(/(\d{2})(\d{2})/);
  if (!mmMatch) return null;
  const baseYear  = 2000 + Number(mmMatch[1]);
  const baseMonth = Number(mmMatch[2]);

  // ← 기존 today ±60일 대신 수동 날짜 or 현재 월 전체
function parseLocalDate_(ymd) {
  const [y, m, d] = String(ymd).split('-').map(Number);
  return new Date(y, m - 1, d);   // ← 로컬 자정 (UTC 파싱 금지)
}

const rangeMin = fromDateStr
  ? parseLocalDate_(fromDateStr)   // ← 로컬 자정으로 파싱
  : new Date(baseYear, baseMonth - 1, 1);

const rangeMax = toDateStr
  ? parseLocalDate_(toDateStr)     // ← 로컬 자정으로 파싱
  : new Date(baseYear, baseMonth, 0);


  function tabToDate(tabName) {
    const m = String(tabName).match(/^(\d+)\(/);
    if (!m) return null;
    return new Date(baseYear, baseMonth - 1, Number(m[1]));
  }

  const stayGroups = {};   // 숙박 누적
  const results    = [];   // 대실 + 최종 숙박 결과

  // ✅ 핵심 수정: 탭당 getValues() 1회만 호출 (기존 2회 → 1회)
  const sheets = ss.getSheets();
  for (const sh of sheets) {
    const tabDate = tabToDate(sh.getName());
    if (!tabDate || tabDate < rangeMin || tabDate > rangeMax) continue;

    const tabDateStr = formatDateStr_(tabDate);
    const data = sh.getDataRange().getValues(); // ← 탭당 딱 1번

    for (let i = 3; i < data.length; i++) {
  // 🔎 디버그: 1(일) 시트 T12 비교 확인
  if (sh.getName() === '1(일)' && i === 11) {
    Logger.log('1(일) T12 실제값=[' + String(data[i][19] || '') + ']');
    Logger.log('현재 검색 name=[' + name + ']');
    Logger.log('같은가? ' + (String(data[i][19] || '').trim() === name));
  }
      // ── 숙박 검색: 열 T (index 19) ──
      if (String(data[i][19]||'').trim() === name) {
        const rawAmt    = data[i][15];
        const amount    = (typeof rawAmt === 'number') ? rawAmt
                        : Number(String(rawAmt||'').replace(/[^0-9.-]/g,''))||0;
        const roomType  = extractBaseRoomType_(String(data[i][20]||'').trim());
        const payMethod = String(data[i][16]||'').trim();
        const assignedRoom = String(data[i][22]||'').trim();

        if (!stayGroups[name]) {
          stayGroups[name] = { payMethod, assignedRoom, nights: [] };
        }
        if (assignedRoom && !stayGroups[name].assignedRoom) {
          stayGroups[name].assignedRoom = assignedRoom;
        }
        stayGroups[name].nights.push({ date: tabDateStr, roomType, amount });
      }

      // ── 대실 검색: 열 I (index 8) ──
      if (String(data[i][8]||'').trim() === name) {
        const rawAmt   = data[i][4];
        const amount   = (typeof rawAmt === 'number') ? rawAmt
                       : Number(String(rawAmt||'').replace(/[^0-9.-]/g,''))||0;
        const roomType  = extractBaseRoomType_(String(data[i][9]||'').trim());
        const payMethod = String(data[i][5]||'').trim();
        results.push({
  guestName: name,   // ← 추가
  stayType: '대실',
          checkIn: tabDateStr, checkOut: tabDateStr,
          roomNo: roomType, roomType, amount, totalAmount: amount,
          payMethod: cleanPayMethod_(payMethod), nights: 0,
          nightlyData: [{ date: tabDateStr, roomType, amount }],
          tabDate: tabDate.getTime(),
        });
      }
    }
  }

  // ── 숙박 결과 조합 ──
  Object.values(stayGroups).forEach(g => {
    if (!g.nights.length) return;
    g.nights.sort((a, b) => a.date.localeCompare(b.date));

    const first       = g.nights[0];
    const last        = g.nights[g.nights.length - 1];
    const lastDateObj = new Date(last.date);
    lastDateObj.setDate(lastDateObj.getDate() + 1);
    const checkOut    = formatDateStr_(lastDateObj);

    let totalAmount = g.nights.reduce((s, r) => s + r.amount, 0);
    let nightlyData = g.nights.map(r => ({ date: r.date, roomType: r.roomType, amount: r.amount }));

    const nonZero = g.nights.filter(r => r.amount > 0);
    if (nonZero.length === 1 && g.nights.length > 1) {
      const perNight = Math.round(nonZero[0].amount / g.nights.length);
      totalAmount    = nonZero[0].amount;
      nightlyData    = g.nights.map(r => ({ date: r.date, roomType: r.roomType, amount: perNight }));
    }

    const cleanType = g.nights.map(n => n.roomType).find(t => t) || '';
results.push({
  guestName: name,   // ← 이 줄 추가
  stayType: '숙박',
  checkIn: first.date,
  checkOut,
      roomNo: g.assignedRoom || cleanType,
      roomType: cleanType,
      amount: totalAmount, totalAmount,
      payMethod: cleanPayMethod_(g.payMethod),
      nights: g.nights.length, nightlyData,
      tabDate: new Date(first.date).getTime(),
    });
  });

  if (results.length === 0) return null;
  results.sort((a, b) => b.tabDate - a.tabDate);
  return results[0];
}


function formatDateStr_(dt) {
  if (!(dt instanceof Date)) dt = new Date(dt);
  const y  = dt.getFullYear();
  const mo = String(dt.getMonth()+1).padStart(2,'0');
  const d  = String(dt.getDate()).padStart(2,'0');
  return `${y}-${mo}-${d}`;
}

function cleanPayMethod_(val) {
  const s = String(val||'').trim();
  if (!s) return '법인카드';
  if (/^\d{1,2}(시|:\d{2})?$/.test(s)) return '법인카드';
  return s;
}

function extractBaseRoomType_(raw) {
  const s = String(raw || '').trim();
  if (!s) return '';
  const keywords = [
    '그랜드 스위트', '시그니처 스위트', '스위트 - 투베드', '스위트 - 킹베드',
    '디럭스 - 빠른입실 늦은퇴실', '스탠다드 - 빠른입실 늦은퇴실',
    '도보특가 - 주차불가', '스위트', '디럭스', '스탠다드', '도보특가'
  ];
  for (const k of keywords) { if (s.includes(k)) return k; }
  const m = s.match(/^(.+?)-\d+\([월화수목금토일]\)(입실|퇴실)/);
  if (m) return m[1].trim();
  if (/퇴실|예약자|입실일|아버지|어머니|남편|부인|아내/.test(s)) return '';
  return s;
}


function debugFindReservationEvidence() {
  const targetName = '실제이름'; // ← 여기만 실제 이름으로 바꾸기

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssName = ss.getName();
  Logger.log('파일명 = [' + ssName + ']');

  const mmMatch = ssName.match(/(?:\(|_|^)(\d{2})(\d{2})(?:\)|\b)/);
  Logger.log('mmMatch = ' + JSON.stringify(mmMatch));
  if (!mmMatch) {
    Logger.log('연월 추출 실패');
    return;
  }

  const baseYear  = 2000 + Number(mmMatch[1]);
  const baseMonth = Number(mmMatch[2]);
  Logger.log('baseYear=' + baseYear + ', baseMonth=' + baseMonth);

  const rangeMin = new Date(baseYear, baseMonth - 1, 1);
  const rangeMax = new Date(baseYear, baseMonth, 0);
  Logger.log('rangeMin=' + formatDateStr_(rangeMin) + ', rangeMax=' + formatDateStr_(rangeMax));

  function tabToDate(tabName) {
    const m = String(tabName).match(/^(\d+)\(/);
    if (!m) return null;
    return new Date(baseYear, baseMonth - 1, Number(m[1]));
  }

  const sheets = ss.getSheets();
  let hitCount = 0;

  for (const sh of sheets) {
    const shName = sh.getName();
    const tabDate = tabToDate(shName);

    Logger.log('--- 시트 [' + shName + '] / tabDate=' + (tabDate ? formatDateStr_(tabDate) : 'null'));

    if (!tabDate || tabDate < rangeMin || tabDate > rangeMax) {
      Logger.log('   -> 스킵');
      continue;
    }

    const data = sh.getDataRange().getValues();
    Logger.log('   -> 검색대상, rows=' + data.length);

    for (let i = 3; i < data.length; i++) {
      const tName = String(data[i][19] || '').trim();
      const iName = String(data[i][8] || '').trim();

      if (tName === targetName) {
        hitCount++;
        Logger.log('   [숙박 HIT] 시트=' + shName + ', row=' + (i + 1) +
                   ', T=' + tName +
                   ', P=' + data[i][15] +
                   ', Q=' + data[i][16] +
                   ', U=' + data[i][20] +
                   ', W=' + data[i][22]);
      }

      if (iName === targetName) {
        hitCount++;
        Logger.log('   [대실 HIT] 시트=' + shName + ', row=' + (i + 1) +
                   ', I=' + iName +
                   ', E=' + data[i][4] +
                   ', F=' + data[i][5] +
                   ', J=' + data[i][9]);
      }
    }
  }

  Logger.log('총 HIT 수 = ' + hitCount);
}
function syncCustomerRecordFromRow_(sheet, row, nameCol) {
  const name = String(sheet.getRange(row, nameCol).getValue() || '').trim();
  if (!name) return;

  const isLeft = (nameCol === 9);

  const roomCol   = isLeft ? 2  : 14;
  const amountCol = isLeft ? 6  : 17;   // ← 선생님 시트 금액열에 맞게 필요시 조정
  const carLocCol = isLeft ? 11 : 22;
  const noteCol   = isLeft ? 12 : 23;   // ← 특이사항 열번호 맞게 필요시 조정

  const visitDate = getDateStrFromSheet_(sheet.getName());
  const roomNo    = String(sheet.getRange(row, roomCol).getValue() || '').trim();
  const amount    = sheet.getRange(row, amountCol).getValue();
  const carLoc    = String(sheet.getRange(row, carLocCol).getValue() || '').trim();
  const extraInfo = String(sheet.getRange(row, noteCol).getValue() || '').trim();

  addCustomerRecord_(
    name,
    '',
    visitDate,
    '숙박',
    '',
    amount || '',
    '',
    roomNo,
    extraInfo || carLoc,
    'manual'
  );
}

function cancelCustomerRecordByKey_(name, visitDate) {
  try {
    if (!name || !visitDate) return;

    const sh = getCustDbSheet_();
    const values = sh.getDataRange().getValues();

    for (let i = values.length - 1; i >= 1; i--) {
      const rowName = String(values[i][0] || '').trim();
      const rowDateRaw = values[i][2];
      const rowDate = rowDateRaw instanceof Date
        ? Utilities.formatDate(rowDateRaw, 'Asia/Seoul', 'yyyy-MM-dd')
        : String(rowDateRaw || '').slice(0, 10);

      if (rowName === name && rowDate === visitDate) {
        sh.deleteRow(i + 1);
        return;
      }
    }
  } catch (e) {
    Logger.log('cancelCustomerRecordByKey_ 오류: ' + e);
  }
}

function syncCustomerRecordFromRow_(sheet, row, nameCol) {
  const name = String(sheet.getRange(row, nameCol).getValue() || '').trim();
  if (!name) return;

  const isLeft = (nameCol === 9);

  const roomCol    = isLeft ? 2  : 14;
  const vehicleCol = isLeft ? 10 : 21;
  const carLocCol  = isLeft ? 11 : 22;

  // ※ 아래 2개 열번호는 선생님 시트 기준으로 맞춰주세요.
  const amountCol  = isLeft ? 6  : 17;
  const noteCol    = isLeft ? 12 : 23;

  const visitDate  = getDateStrFromSheet_(sheet.getName());
  const roomNo     = String(sheet.getRange(row, roomCol).getValue() || '').trim();
  const vehicle    = String(sheet.getRange(row, vehicleCol).getValue() || '').trim();
  const carLoc     = String(sheet.getRange(row, carLocCol).getValue() || '').trim();
  const amount     = sheet.getRange(row, amountCol).getValue();
  const note       = String(sheet.getRange(row, noteCol).getValue() || '').trim();

  addCustomerRecord_(
    name,
    vehicle,
    visitDate,
    '숙박',
    '',
    amount || '',
    '',
    roomNo,
    note || carLoc,
    'manual'
  );
}

function installDailyCustomerDbRefresh_() {
  const fn = 'rebuildCustomerDb_';

  // 기존 같은 트리거 삭제
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === fn) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 매일 새벽 5시 자동 실행
  ScriptApp.newTrigger(fn)
    .timeBased()
    .everyDays(1)
    .atHour(5)
    .create();
}

function rebuildCustomerDb_() {
  resetCustDbCache_();

  const db = getCustDbSheet_();

  if (db.getLastRow() > 1) {
    db.getRange(2, 1, db.getLastRow() - 1, 10).clearContent();
  }

const files = [
  { id: '10z3CRqgf6cKBNCNneqCuueWlRTuNzJS0VTENuaBeiIQ', month: 2 }, // 2월
  { id: '1X8nuLsO3OjyCA5UcCinj7N7IXuL0ha3g_aLpJcIhAwM', month: 3 }  // 3월
];

const rows = [];

files.forEach(file => {
  const ss = SpreadsheetApp.openById(file.id);
  const sheets = ss.getSheets();

  sheets.forEach(sh => {
    const name = sh.getName();
    if (!name.match(/^\d+\(/)) return;

    const m = name.match(/^(\d+)\(/);
    if (!m) return;

    const day = Number(m[1]);
    const dateStr = Utilities.formatDate(
      new Date(2026, file.month - 1, day),
      'Asia/Seoul',
      'yyyy-MM-dd'
    );
const values = sh.getDataRange().getValues();
const endRow = Math.min(values.length, CUSTOMER_END_ROW);

for (let r = CUSTOMER_START_ROW - 1; r < endRow; r++) {
  const row = values[r];

  // =========================
  // 1) 대실 블록 (I/J/K + B/E)
  // =========================
  const guestDay = String(row[8] || '').trim();    // I열 이름
  const noteDay  = String(row[9] || '').trim();    // J열 특이사항
  const carDay   = String(row[10] || '').trim();   // K열 차량/차번호
  let roomDay    = String(row[1] || '').trim();    // B열 객실
  const amountDay = row[4];                        // E열 금액

  if (guestDay && !/^\d+$/.test(guestDay) && !/^(숙박비품|대실비품|합)$/.test(noteDay)) {
    if (!roomDay) {
      const m = noteDay.match(/(\d{3})/);
      if (m) roomDay = m[1];
    }

    rows.push([
      guestDay,
      carDay || '',
      dateStr,
      '대실',
      '',
      amountDay || '',
      '',
      roomDay || '',
      noteDay || '',
      '숙박일지'
    ]);
  }

  // =========================
  // 2) 숙박 블록 (T/U/V/W + N/P)
  // =========================
  const guestStay = String(row[19] || '').trim();   // T열 이름
  const noteStay  = String(row[20] || '').trim();   // U열 특이사항/차종
  const carStay   = String(row[21] || '').trim();   // V열 차번호
  let roomStay    = String(row[13] || '').trim();   // N열 객실
  const amountStay = row[15];                       // P열 금액

  if (guestStay && !/^\d+$/.test(guestStay) && !/^(숙박비품|대실비품|합)$/.test(noteStay)) {
    if (!roomStay) {
      const m = noteStay.match(/(\d{3})/);
      if (m) roomStay = m[1];
    }

    rows.push([
      guestStay,
      carStay || '',
      dateStr,
      '숙박',
      '',
      amountStay || '',
      '',
      roomStay || '',
      noteStay || '',
      '숙박일지'
    ]);
  }
}

    });
  });

  if (rows.length) {
    db.getRange(2, 1, rows.length, 10).setValues(rows);
  }
}

/**
 * 날짜를 yyyy-MM-dd 문자열로 통일
 */
function toDateKey_(v) {
  if (!v) return '';
  const d = (v instanceof Date) ? v : new Date(v);
  if (isNaN(d)) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * 고객 식별키 생성
 * 1차: 이름 + 전화번호 뒤 4자리
 * 전화번호가 없으면 이름만
 */
function makeCustomerKey_(name, phone) {
  const nm = String(name || '').trim();
  const ph = String(phone || '').replace(/\D/g, '');
  const last4 = ph ? ph.slice(-4) : '';
  return nm + '|' + last4;
}

/**
 * 과거 방문 이력만 추출
 * - 기준일 당일 제외
 * - 같은 날짜 중복 제거
 * - 최신순 정렬
 *
 * rows: [{name, phone, visitDate, roomNo, stayType, note}, ...]
 */
function buildPastVisitSummary_(rows, targetName, targetPhone, baseDate) {
  const customerKey = makeCustomerKey_(targetName, targetPhone);
  const baseKey = toDateKey_(baseDate);

  const filtered = [];
  const seenDate = new Set();

  rows.forEach(row => {
    const rowKey = makeCustomerKey_(row.name, row.phone);
    if (rowKey !== customerKey) return;

    const visitKey = toDateKey_(row.visitDate);
    if (!visitKey) return;

    // 기준일 당일 제외, 미래도 제외
    if (visitKey >= baseKey) return;

    // 같은 날짜 중복 제거
    if (seenDate.has(visitKey)) return;
    seenDate.add(visitKey);

    filtered.push({
      visitDate: visitKey,
      roomNo: row.roomNo || '',
      stayType: row.stayType || '',
      note: row.note || ''
    });
  });

  filtered.sort((a, b) => b.visitDate.localeCompare(a.visitDate));
  return filtered;
}



/**
 * 말풍선 표시용 텍스트 생성
 * 직원이 1초 안에 읽을 수 있도록 압축
 */
function makeVisitBubbleText_(visitList) {

  if (!visitList || !visitList.length) {
    return '첫 방문';
  }

  const total = visitList.length;
  const lastVisit = visitList[0].visitDate;

  const recentLines = visitList.slice(0, 3).map(v => {
    let line = v.visitDate;

    if (v.roomNo) line += ' / ' + v.roomNo + '호';
    if (v.stayType) line += ' / ' + v.stayType;

    return line;
  });

  return [
    '재방문 고객',
    '총 ' + total + '회',
    '마지막 방문: ' + lastVisit,
    '최근 이력:',
    recentLines.join('\n')
  ].join('\n');
}

/**
 * 고객DB rows 배열에서 특정 고객의 말풍선 정보를 생성
 * - 기준일 당일 제외
 * - 같은 날짜 중복 제거
 */
function getCustomerBubbleInfoFromRows_(rows, targetName, targetPhone, baseDate) {
  const pastVisits = buildPastVisitSummary_(rows, targetName, targetPhone, baseDate);
  const bubbleText = makeVisitBubbleText_(pastVisits);

  return {
    count: pastVisits.length,
    visits: pastVisits,
    text: bubbleText
  };
}

let _CUSTOMER_VISIT_MAP = null;

/**
 * 고객 이름 기준 방문기록 인덱스 생성
 * { 이름 : [row,row,row] }
 */
function getCustomerVisitMap_() {

  if (_CUSTOMER_VISIT_MAP) return _CUSTOMER_VISIT_MAP;

  const data = getCustDbValues_();
  const map = {};

  data.slice(1).forEach(r => {

    const name = String(r[0] || '').trim();
    if (!name) return;

    if (!map[name]) map[name] = [];
    map[name].push(r);

  });

  _CUSTOMER_VISIT_MAP = map;

  return map;
}

function removeDailyCustomerDbRefresh_() {
  const fn = 'rebuildCustomerDb_';
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === fn) {
      ScriptApp.deleteTrigger(t);
    }
  });
}


function makeCustomerKey_(name, vehicle, note) {
  const cleanName = String(name || '').trim();
  const car = String(vehicle || '').trim();

  if (car) return cleanName + '|' + car;

  return cleanName + '|도보';
}

function refreshAllCustomerNotes() {
  const sh = SpreadsheetApp.getActiveSheet();
  const lastRow = sh.getLastRow();
  const values = sh.getDataRange().getValues();

  // 기존 말풍선 먼저 제거
  if (lastRow >= CUSTOMER_START_ROW) {
    sh.getRange(CUSTOMER_START_ROW, 9,  lastRow - CUSTOMER_START_ROW + 1, 1).clearNote().setBackground(null);
    sh.getRange(CUSTOMER_START_ROW, 20, lastRow - CUSTOMER_START_ROW + 1, 1).clearNote().setBackground(null);
  }

  for (let r = CUSTOMER_START_ROW - 1; r < Math.min(values.length, CUSTOMER_END_ROW); r++) {
    const row = r + 1;

    const nameDay = String(values[r][8] || '').trim();
    if (nameDay) {
      attachNote_(sh, row, 9, nameDay);
    }

    const nameStay = String(values[r][19] || '').trim();
    if (nameStay) {
      attachNote_(sh, row, 20, nameStay);
    }
  }

  SpreadsheetApp.getActive().toast('말풍선 전체 재생성 완료', 'Customer Note', 3);
}

function debugSensProps() {
  const props = PropertiesService.getScriptProperties().getProperties();

  Logger.log('ALL PROPS = ' + JSON.stringify(props));
  Logger.log('SENS_SERVICE_ID = ' + (props.SENS_SERVICE_ID || ''));
  Logger.log('SENS_ACCESS_KEY = ' + (props.SENS_ACCESS_KEY ? 'exists' : ''));
  Logger.log('SENS_SECRET_KEY = ' + (props.SENS_SECRET_KEY ? 'exists' : ''));
  Logger.log('SENS_FROM = ' + (props.SENS_FROM || ''));
}
function pingTest() {
  Logger.log('PING OK');
}
function propTest() {
  const props = PropertiesService.getScriptProperties().getProperties();
  Logger.log('PROPS=' + JSON.stringify(props));
}
function fixSensFromKey() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();

  const sender = all.SENS_SENDER || '';
  if (!sender) {
    Logger.log('SENS_SENDER 값이 없습니다.');
    return;
  }

  props.setProperty('SENS_FROM', sender);
  Logger.log('SENS_FROM 복사 완료: ' + sender);
}

function runRebuildCustomerDb() {
  rebuildCustomerDb_();
}

function clearInvalidCustomerNotes() {
  const sh = SpreadsheetApp.getActiveSheet();
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const startRow = CUSTOMER_END_ROW + 1;

  if (lastRow < startRow) {
    SpreadsheetApp.getActive().toast('삭제할 말풍선이 없습니다.', 'Customer Note', 3);
    return;
  }

  sh.getRange(startRow, 1, lastRow - startRow + 1, lastCol)
    .clearNote()
    .setBackground(null);

  SpreadsheetApp.getActive().toast('45행 이하 전체 말풍선 삭제 완료', 'Customer Note', 3);
}