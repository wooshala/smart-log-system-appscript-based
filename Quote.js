/* ═══════════════════════════════════════════════════════════
   Quote.gs  –  견적 저장 · PDF 생성 · SMS 발송
   DB 스키마(영어 헤더) 기준
═══════════════════════════════════════════════════════════ */

// ── 상수 ──────────────────────────────────────────────────
const Q_CFG = {
  PRICE_SHEET : '가격표',
  DB_SHEET    : '견적DB',
  TMPL_SHEET  : '견적서양식',
DB_COLS     : [
  'quoteId',
  'createdAt',
  'updatedAt',
  'customerName',
  'customerPhone',
  'itemsJson',
  'dateSummaryJson',

  'roomTotalAmount',
  'extraLabel1',
  'extraAmount1',
  'extraLabel2',
  'extraAmount2',
  'extraLabel3',
  'extraAmount3',
  'extraAmountTotal',
  'finalTotalAmount',
  'priceDisplayMode',

  'pdfFileId',
  'pdfUrl',
  'vehicleType',
  'vehicleCount',
  'employeeName',
  'extraNote',
  'lastSentAt',
  'lastSendTo',
  'lastSendResult'
]
};

// ── DB 유틸 ───────────────────────────────────────────────
function getDbSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Q_CFG.DB_SHEET);
}

function ensureDbSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(Q_CFG.DB_SHEET);
  if (!sh) {
    sh = ss.insertSheet(Q_CFG.DB_SHEET);
    sh.getRange(1, 1, 1, Q_CFG.DB_COLS.length).setValues([Q_CFG.DB_COLS]);
  }
  return sh;
}

function getNextQuoteId_() {
  const sh   = ensureDbSheet_();
  const last = sh.getLastRow();
  if (last < 2) return 'Q-' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd') + '01';
  const ids = sh.getRange(2, 1, last - 1, 1).getValues().flat()
                .map(String).filter(v => v.startsWith('Q-'));
  if (!ids.length) return 'Q-' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd') + '01';
  const today   = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd');
  const todayIds = ids.filter(id => id.startsWith('Q-' + today));
  const seq = todayIds.length + 1;
  return 'Q-' + today + String(seq).padStart(2, '0');
}

// ── 요금표 로더 ───────────────────────────────────────────
function loadPriceTable_() {
  const sh   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(Q_CFG.PRICE_SHEET);
  if (!sh) throw new Error('요금표 시트 없음');
  const data = sh.getDataRange().getValues();
  // 헤더 행 탐색
  let hRow = -1;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).includes('객실') || String(data[i][0]).includes('타입') || String(data[i][0]).includes('룸')) {
      hRow = i; break;
    }
  }
  if (hRow < 0) hRow = 0;
  const headers = data[hRow].map(String);
  const map = {};
  for (let r = hRow + 1; r < data.length; r++) {
    const roomRaw = String(data[r][0]).trim();
    if (!roomRaw) continue;
    const key = normalizeRoomKey_(roomRaw);
    const row = {};
    headers.forEach((h, c) => { row[h] = data[r][c]; });
    map[key] = row;
  }
  return { headers, map };
}

function normalizeRoomKey_(raw) {
  return String(raw).trim().replace(/\s+/g, ' ');
}

function getPriceRow_(priceMap, roomType) {
  const key = normalizeRoomKey_(roomType);
  return priceMap[key] || priceMap[Object.keys(priceMap).find(k => k.includes(key) || key.includes(k)) || ''] || null;
}

// ── 날짜 파싱 ─────────────────────────────────────────────
function parseYmd_(str) {
  const s = String(str).replace(/\./g, '-').replace(/\//g, '-').trim();
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// ── 요일별 금액 계산 ──────────────────────────────────────
function calcStayAmountByDow_(priceRow, checkIn, checkOut) {
  const dow_map = { 0:'일', 1:'월', 2:'화', 3:'수', 4:'목', 5:'금', 6:'토' };
  let total = 0;
  const cur = new Date(checkIn);
  while (cur < checkOut) {
    const dow = dow_map[cur.getDay()];
    const col = Object.keys(priceRow).find(k => String(k).includes(dow));
    const price = col ? (Number(priceRow[col]) || 0) : 0;
    total += price;
    cur.setDate(cur.getDate() + 1);
  }
  return total;
}

// ── 사이드바용 calcItemPreview ────────────────────────────
function calcItemPreview(roomType, checkIn, checkOut) {
  try {
    const { map } = loadPriceTable_();
    const priceRow = getPriceRow_(map, roomType);
    if (!priceRow) return { ok: false, error: '요금 정보 없음: ' + roomType };
    const ci = parseYmd_(checkIn);
    const co = parseYmd_(checkOut);
    if (!ci || !co || co <= ci) return { ok: false, error: '날짜 오류' };
    const nights = Math.round((co - ci) / 86400000);
    const subtotal = calcStayAmountByDow_(priceRow, ci, co);
    const unitPrice = nights > 0 ? Math.round(subtotal / nights) : 0;
    return { ok: true, nights, unitPrice, subtotal };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ── 사이드바용 객실 목록 ──────────────────────────────────
function getRoomsForSidebar() {
  try {
    const { headers, map } = loadPriceTable_();

    // 요금표 한국어 요일 헤더 → 사이드바 영어 키 매핑
    const KO_TO_EN = {
      '일': 'sun', '월': 'mon', '화': 'tue',
      '수': 'wed', '목': 'thu', '금': 'fri', '토': 'sat'
    };

    return Object.keys(map).map(key => {
      const priceRow = map[key];
      const room = { type: key };   // ← 핵심: name → type

      // 요금표 열 이름(한국어 요일 포함)을 영어 키로 변환
      Object.keys(priceRow).forEach(col => {
        Object.entries(KO_TO_EN).forEach(([ko, en]) => {
          if (String(col).includes(ko)) {
            room[en] = Number(priceRow[col]) || 0;
          }
        });
      });

      // fallback: 평일(월) / 주말(토)
      room.weekday = room.mon || 0;
      room.weekend = room.sat || 0;

      return room;
    });
  } catch (e) {
    Logger.log('getRoomsForSidebar 오류: ' + e.message);
    return [];
  }
}


// ── 견적 저장 ─────────────────────────────────────────────
function saveQuote(payload) {
  try {
    const sh = ensureDbSheet_();
    syncQuoteDbHeader_();

    const now    = new Date();
    const nowStr = Utilities.formatDate(now, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

    let qId    = payload.quoteId;
    let rowIdx = -1;

    if (qId) {
      const data = sh.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(qId)) { rowIdx = i + 1; break; }
      }
    }

    const isNew = (rowIdx < 0);
    if (isNew) {
      qId    = getNextQuoteId_();
      rowIdx = sh.getLastRow() + 1;
    }

    let existing = new Array(Q_CFG.DB_COLS.length).fill('');
    if (!isNew) {
      existing = sh.getRange(rowIdx, 1, 1, Q_CFG.DB_COLS.length).getValues()[0];
    }

const sourceQuoteId = String(payload.sourceQuoteId || '').trim();
const employeeName  = String(payload.employeeName || '').trim();

const extraLabel1 = String(payload.extraLabel1 || '').trim();
const extraAmount1 = parseMoney_(payload.extraAmount1);

const extraLabel2 = String(payload.extraLabel2 || '').trim();
const extraAmount2 = parseMoney_(payload.extraAmount2);

const extraLabel3 = String(payload.extraLabel3 || '').trim();
const extraAmount3 = parseMoney_(payload.extraAmount3);

const extraAmountTotal = extraAmount1 + extraAmount2 + extraAmount3;

const rawItems = payload.items || [];
const recalculatedItems = [];
let recalculatedTotal = 0;

    rawItems.forEach(function(it) {
      const roomType = String(it.roomType || '').trim();
      const checkIn  = String(it.checkIn || '').trim();
      const checkOut = String(it.checkOut || '').trim();
      const rooms    = Number(it.rooms || 1);

      const preview = calcItemPreview(roomType, checkIn, checkOut);

      let nights = Number(it.nights || 1);
      let unitPrice = Number(it.unitPrice || 0);
      let subtotal = nights * unitPrice;

      if (preview && preview.ok) {
        nights = Number(preview.nights || nights);
        unitPrice = Number(preview.unitPrice || unitPrice);
        subtotal = Number(preview.subtotal || 0);
      }

      const finalSubtotal = subtotal * rooms;

      recalculatedItems.push({
        roomType: roomType,
        checkIn: checkIn,
        checkOut: checkOut,
        nights: nights,
        rooms: rooms,
        unitPrice: unitPrice,
        subtotal: finalSubtotal
      });

      recalculatedTotal += finalSubtotal;
    });

const isEditedFromPrevious = !!sourceQuoteId;

const roomTotalAmount = recalculatedTotal;
const finalTotalAmount = roomTotalAmount + extraAmountTotal;

const row = [
  qId,
  isNew ? nowStr : existing[1],
  nowStr,
  payload.customerName || payload.guestName || '',
  preservePhoneText_(payload.customerPhone || payload.guestPhone || ''),
  JSON.stringify(recalculatedItems),
  JSON.stringify(payload.dateSummary || {}),

  roomTotalAmount,
  extraLabel1,
  extraAmount1,
  extraLabel2,
  extraAmount2,
  extraLabel3,
  extraAmount3,
  extraAmountTotal,
  finalTotalAmount,
  String(payload.priceDisplayMode || 'avg'),

  isEditedFromPrevious ? '' : (existing[17] || ''),
  isEditedFromPrevious ? '' : (existing[18] || ''),
  payload.vehicleType || '',
  payload.vehicleCount || '',
  payload.employeeName || '',
  payload.extraNote || '',
  isEditedFromPrevious ? '' : (existing[23] || ''),
  isEditedFromPrevious ? '' : (existing[24] || ''),
  isEditedFromPrevious ? '' : (existing[25] || '')
];

sh.getRange(rowIdx, 1, 1, row.length).setValues([row]);

    let oldDetail = null;
    if (sourceQuoteId) {
      const oldRes = getQuoteHistoryDetail(sourceQuoteId);
      if (oldRes && oldRes.ok) oldDetail = oldRes.detail;
    }

    const changeSummary = buildChangeSummary_(
      oldDetail,
      payload,
      recalculatedItems,
      recalculatedTotal
    );

    appendQuoteLog_(
      qId,
      sourceQuoteId,
      sourceQuoteId ? 'edit_from_previous' : 'create',
      employeeName,
      payload.customerName || payload.guestName || '',
      changeSummary,
      {
        items: recalculatedItems,
        totalAmount: recalculatedTotal,
        vehicleType: payload.vehicleType || '',
        vehicleCount: payload.vehicleCount || '',
        extraNote: payload.extraNote || ''
      }
    );

  return {
  ok: true,
  quoteId: qId,
  totalAmount: finalTotalAmount,
  roomTotalAmount: roomTotalAmount,
  extraAmountTotal: extraAmountTotal,
  finalTotalAmount: finalTotalAmount
};

  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ── PDF 생성 ──────────────────────────────────────────────
function generatePdfForQuote(quoteId) {
  try {
    const sh   = ensureDbSheet_();
    const data = sh.getDataRange().getValues();
    const hdr  = data[0];
    const rowArr = data.slice(1).find(r => String(r[0]) === String(quoteId));
    if (!rowArr) return { ok: false, error: '견적 ID 없음: ' + quoteId };

    const rec = {};
    hdr.forEach((h, i) => { rec[h] = rowArr[i]; });

    // items 파싱
    let items = [];
    try { items = JSON.parse(rec['itemsJson'] || '[]'); } catch(e) {}

    // 인보이스 시트 생성
    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const shName   = getInvoiceSheetName_(quoteId);
    let   invSh    = ss.getSheetByName(shName);
    if (invSh) ss.deleteSheet(invSh);
    invSh = ss.insertSheet(shName);

drawInvoiceToSheet_(invSh, {
  quoteId         : rec['quoteId'],
  createdAt       : rec['createdAt'],
  customerName    : rec['customerName'],
  customerPhone   : formatPhoneDisplay_(String(rec['customerPhone'] || '').replace(/^'/, '')),
  items,
  roomTotalAmount : Number(rec['roomTotalAmount'] || 0),
  extraLabel1     : rec['extraLabel1'] || '',
  extraAmount1    : Number(rec['extraAmount1'] || 0),
  extraLabel2     : rec['extraLabel2'] || '',
  extraAmount2    : Number(rec['extraAmount2'] || 0),
  extraLabel3     : rec['extraLabel3'] || '',
  extraAmount3    : Number(rec['extraAmount3'] || 0),
  extraAmountTotal: Number(rec['extraAmountTotal'] || 0),
  finalTotalAmount: Number(rec['finalTotalAmount'] || 0),
  totalAmount     : Number(rec['finalTotalAmount'] || rec['roomTotalAmount'] || 0),
  priceDisplayMode: String(rec['priceDisplayMode'] || 'avg'),
  vehicleType     : rec['vehicleType'],
  vehicleCount    : rec['vehicleCount'],
  extraNote       : rec['extraNote']
});

    SpreadsheetApp.flush(); // ✅ 데이터 반영 대기

    // PDF 내보내기
    const pdfBlob = exportSheetToPdf_(ss.getId(), invSh.getSheetId());


    const folder  = getOrCreateFolder_('견적서PDF');
    const fname   = quoteId + '_' + rec['customerName'] + '.pdf';

    // 기존 파일 삭제
    if (rec['pdfFileId']) {
      try { DriveApp.getFileById(rec['pdfFileId']).setTrashed(true); } catch(e) {}
    }

    const file    = folder.createFile(pdfBlob.setName(fname));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileId  = file.getId();
    const fileUrl = 'https://drive.google.com/file/d/' + fileId + '/view';

    // DB 업데이트 (pdfFileId, pdfUrl)
    const rowIdx = data.findIndex(r => String(r[0]) === String(quoteId)) + 1;
    const colPdfId  = Q_CFG.DB_COLS.indexOf('pdfFileId')  + 1;  // 9
    const colPdfUrl = Q_CFG.DB_COLS.indexOf('pdfUrl')     + 1;  // 10
    sh.getRange(rowIdx, colPdfId).setValue(fileId);
    sh.getRange(rowIdx, colPdfUrl).setValue(fileUrl);

    // 인보이스 시트 삭제
    ss.deleteSheet(invSh);

    return { ok: true, pdfUrl: fileUrl };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ── 인보이스 시트 그리기 ──────────────────────────────────
// ── 인보이스 시트 그리기 ──────────────────────────────────
function drawInvoiceToSheet_(sh, d) {
  sh.clear();
  sh.setHiddenGridlines(true);

  const NAVY    = '#1a3c5e';
  const NAVY_LT = '#dce6f1';
  const WHITE   = '#ffffff';
  const GRAY    = '#f5f5f5';

  const START_COL = 2; // B열부터 시작

  // A열은 왼쪽 여백
  sh.setColumnWidth(1, 24);

  // B:H, 7열
  const colWidths = [100, 130, 75, 80, 80, 40, 85];
  colWidths.forEach((w, i) => sh.setColumnWidth(START_COL + i, w));

  function realCol_(c) {
    return START_COL + c - 1; // 1→B, 2→C ... 7→H
  }

  function rng(r, c, rr, cc) {
    return (rr && cc)
      ? sh.getRange(r, realCol_(c), rr - r + 1, cc - c + 1)
      : sh.getRange(r, realCol_(c));
  }

  function merge(r1, c1, r2, c2) { rng(r1, c1, r2, c2).merge(); }
  function val(r, c, v)          { sh.getRange(r, realCol_(c)).setValue(v); }
  function fmt(r, c, f)          { sh.getRange(r, realCol_(c)).setNumberFormat(f); }

  function applyStyle(r, c, r2, c2, opts) {
    const g = rng(r, c, r2, c2);
    if (opts.bg)     g.setBackground(opts.bg);
    if (opts.color)  g.setFontColor(opts.color);
    if (opts.bold)   g.setFontWeight('bold');
    if (opts.size)   g.setFontSize(opts.size);
    if (opts.halign) g.setHorizontalAlignment(opts.halign);
    if (opts.valign) g.setVerticalAlignment(opts.valign);
    if (opts.wrap)   g.setWrap(true);
    if (opts.border) g.setBorder(true, true, true, true, false, false,
      '#aaaaaa', SpreadsheetApp.BorderStyle.SOLID);
  }

  // ── 행 높이 기본값 ──
  for (let r = 1; r <= 50; r++) sh.setRowHeight(r, 22);

  // ══════════════════════════════════════
  // 행 1: 호텔명 헤더
  // ══════════════════════════════════════
  sh.setRowHeight(1, 38);
  merge(1,1, 1,7);
  val(1, 1, 'HOTEL LABEL  견적서');
  applyStyle(1,1,1,7, {bg:NAVY, color:WHITE, bold:true, size:16,
                        halign:'center', valign:'middle'});

  // ── 행 2: 주소/연락처 ──
  sh.setRowHeight(2, 20);
  merge(2,1, 2,7);
  val(2, 1, CFG.HOTEL_ADDRESS + '  |  TEL: ' + CFG.HOTEL_TEL +
            '  |  사업자번호: ' + CFG.HOTEL_BIZ_NO);
  applyStyle(2,1,2,7, {size:9, halign:'center', valign:'middle', border:true});

  // ══════════════════════════════════════
  // 행 3~4: 고객정보 (2열 구성)
  // ══════════════════════════════════════
sh.setRowHeight(3, 26);
sh.setRowHeight(4, 26);


  const dateStr = d.createdAt instanceof Date
    ? Utilities.formatDate(d.createdAt, 'Asia/Seoul', 'yyyy-MM-dd')
    : String(d.createdAt || '').substring(0, 10);

  // 왼쪽: 고객명 / 연락처
  const infoLeft = [['고객명', d.customerName || ''], ['연락처', d.customerPhone || '']];
  infoLeft.forEach(([label, value], i) => {
    const r = 3 + i;
    val(r, 1, label);
    applyStyle(r,1,r,1, {bg:NAVY_LT, bold:true, halign:'center', valign:'middle', border:true});
    merge(r, 2, r, 4);
    val(r, 2, value);
    applyStyle(r,2,r,4, {halign:'center', valign:'middle', border:true});
  });

  // 오른쪽: 문서번호 / 작성일
  const infoRight = [['문서번호', d.quoteId || ''], ['작성일', dateStr]];
  infoRight.forEach(([label, value], i) => {
    const r = 3 + i;
    val(r, 5, label);
    applyStyle(r,5,r,5, {bg:NAVY_LT, bold:true, halign:'center', valign:'middle', border:true});
    merge(r, 6, r, 7);
    val(r, 6, value);
    applyStyle(r,6,r,7, {halign:'center', valign:'middle', border:true});
  });

  // ══════════════════════════════════════
  // 행 5: 견적내역 섹션 헤더
  // ══════════════════════════════════════
  sh.setRowHeight(5, 26);
  merge(5,1, 5,7);
  val(5, 1, '■ 견 적 내 역');
  applyStyle(5,1,5,7, {bg:NAVY, color:WHITE, bold:true, size:11,
                        halign:'left', valign:'middle'});

  // ── 행 6: 테이블 헤더 ──
  sh.setRowHeight(6, 24);
  const tblHdr = ['구  분', '객실수', '단  가', '체크인', '체크아웃', '박수', '금  액'];
  tblHdr.forEach((h, i) => {
    val(6, i + 1, h);
    applyStyle(6, i+1, 6, i+1, {bg:NAVY, color:WHITE, bold:true,
                                  halign:'center', valign:'middle', border:true});
  });

  // ── 행 7~: 객실 항목 ──
  const items    = d.items || [];
  const ITEM_START = 7;
  const MAX_ITEMS  = 15;

  for (let i = 0; i < MAX_ITEMS; i++) {
    const r   = ITEM_START + i;
    const itm = items[i];
    sh.setRowHeight(r, 22);

    if (itm) {
      val(r, 1, itm.roomType  || '');
      val(r, 2, itm.rooms     || 1);
      val(r, 3, itm.unitPrice || 0);
      fmt(r, 3, '#,##0"원"');
      val(r, 4, itm.checkIn   || '');
      val(r, 5, itm.checkOut  || '');
      val(r, 6, itm.nights    || '');
      val(r, 7, itm.subtotal  || 0);
      fmt(r, 7, '#,##0"원"');

applyStyle(r,1,r,1, {valign:'middle', border:true});
      applyStyle(r,2,r,2, {halign:'center', valign:'middle', border:true});
      applyStyle(r,3,r,3, {halign:'right',  valign:'middle', border:true});
      applyStyle(r,4,r,4, {halign:'center', valign:'middle', border:true});
      applyStyle(r,5,r,5, {halign:'center', valign:'middle', border:true});
      applyStyle(r,6,r,6, {halign:'center', valign:'middle', border:true});
      applyStyle(r,7,r,7, {halign:'right',  valign:'middle', border:true});
    } else {
      for (let c = 1; c <= 7; c++) {
        applyStyle(r,c,r,c, {border:true});
      }
    }
  }


// ── 합계 행 ──
const totalRow = ITEM_START + MAX_ITEMS;
sh.setRowHeight(totalRow, 28);
merge(totalRow, 1, totalRow, 6);
val(totalRow, 1, '객실 합계');
applyStyle(totalRow,1,totalRow,6, {
  bg:NAVY_LT, bold:true, halign:'center', valign:'middle', border:true
});
val(totalRow, 7, Number(d.roomTotalAmount || d.totalAmount || 0));
fmt(totalRow, 7, '#,##0"원"');
applyStyle(totalRow,7,totalRow,7, {
  bold:true, halign:'right', valign:'middle', border:true, bg:NAVY_LT
});

let cursorRow = totalRow;

if (String(d.priceDisplayMode || 'avg') === 'daily') {
  const itemGroups = (d.items || []).map(function(it) {
    return {
      roomType: it.roomType || '',
      lines: buildDailyPriceLines_(it)
    };
  }).filter(function(group) {
    return group.lines && group.lines.length;
  });

  if (itemGroups.length) {
    cursorRow += 1;
    sh.setRowHeight(cursorRow, 24);
    merge(cursorRow, 1, cursorRow, 7);
    val(cursorRow, 1, '■ 일 자 별 요 금');
    applyStyle(cursorRow,1,cursorRow,7, {
      bg:NAVY, color:WHITE, bold:true, size:11, halign:'left', valign:'middle'
    });

    itemGroups.forEach(function(group) {
      // 객실명 1회만 표시
      cursorRow += 1;
      sh.setRowHeight(cursorRow, 22);
      merge(cursorRow, 1, cursorRow, 7);
      val(cursorRow, 1, '객실명: ' + group.roomType);
      applyStyle(cursorRow,1,cursorRow,7, {
        bg:NAVY_LT, bold:true, halign:'left', valign:'middle', border:true
      });

      // 날짜 / 요일 / 금액 헤더
      cursorRow += 1;
      sh.setRowHeight(cursorRow, 22);
      merge(cursorRow, 1, cursorRow, 3);
      val(cursorRow, 1, '날짜');
      applyStyle(cursorRow,1,cursorRow,3, {
        bg:GRAY, bold:true, halign:'center', valign:'middle', border:true
      });

      merge(cursorRow, 4, cursorRow, 5);
      val(cursorRow, 4, '요일');
      applyStyle(cursorRow,4,cursorRow,5, {
        bg:GRAY, bold:true, halign:'center', valign:'middle', border:true
      });

      merge(cursorRow, 6, cursorRow, 7);
      val(cursorRow, 6, '금액');
      applyStyle(cursorRow,6,cursorRow,7, {
        bg:GRAY, bold:true, halign:'center', valign:'middle', border:true
      });

      // 날짜별 금액 행
      group.lines.forEach(function(line) {
        cursorRow += 1;
        sh.setRowHeight(cursorRow, 22);

        merge(cursorRow, 1, cursorRow, 3);
        val(cursorRow, 1, line.date);
        applyStyle(cursorRow,1,cursorRow,3, {
          border:true, halign:'center', valign:'middle'
        });

        merge(cursorRow, 4, cursorRow, 5);
        val(cursorRow, 4, line.dow);
        applyStyle(cursorRow,4,cursorRow,5, {
          border:true, halign:'center', valign:'middle'
        });

        merge(cursorRow, 6, cursorRow, 7);
        val(cursorRow, 6, line.amount);
        fmt(cursorRow, 6, '#,##0"원"');
        applyStyle(cursorRow,6,cursorRow,7, {
          border:true, halign:'right', valign:'middle'
        });
      });
    });
  }
}

let extraRow = cursorRow + 1;

const extraItems = [
  { label: d.extraLabel1 || '', amount: Number(d.extraAmount1 || 0) },
  { label: d.extraLabel2 || '', amount: Number(d.extraAmount2 || 0) },
  { label: d.extraLabel3 || '', amount: Number(d.extraAmount3 || 0) }
].filter(function(x) {
  return x.label || x.amount > 0;
});

if (extraItems.length) {
  sh.setRowHeight(extraRow, 24);
  merge(extraRow, 1, extraRow, 7);
  val(extraRow, 1, '■ 추 가 비 용');
  applyStyle(extraRow,1,extraRow,7, {
    bg:NAVY, color:WHITE, bold:true, size:11, halign:'left', valign:'middle'
  });

  extraItems.forEach(function(ex, idx) {
    const r = extraRow + 1 + idx;
    sh.setRowHeight(r, 22);

    merge(r, 1, r, 5);
    val(r, 1, ex.label);
    applyStyle(r,1,r,5, {border:true, valign:'middle'});

    merge(r, 6, r, 7);
    val(r, 6, ex.amount);
    fmt(r, 6, '#,##0"원"');
    applyStyle(r,6,r,7, {border:true, halign:'right', valign:'middle'});
  });

  const extraTotalRow = extraRow + 1 + extraItems.length;
  merge(extraTotalRow, 1, extraTotalRow, 5);
  val(extraTotalRow, 1, '추가비용 합계');
  applyStyle(extraTotalRow,1,extraTotalRow,5, {
    bg:NAVY_LT, bold:true, halign:'center', valign:'middle', border:true
  });

  merge(extraTotalRow, 6, extraTotalRow, 7);
  val(extraTotalRow, 6, Number(d.extraAmountTotal || 0));
  fmt(extraTotalRow, 6, '#,##0"원"');
  applyStyle(extraTotalRow,6,extraTotalRow,7, {
    bg:NAVY_LT, bold:true, halign:'right', valign:'middle', border:true
  });

  const finalTotalRow = extraTotalRow + 1;
  merge(finalTotalRow, 1, finalTotalRow, 5);
  val(finalTotalRow, 1, '최 종 합 계');
  applyStyle(finalTotalRow,1,finalTotalRow,5, {
    bg:NAVY, color:WHITE, bold:true, halign:'center', valign:'middle', border:true
  });

  merge(finalTotalRow, 6, finalTotalRow, 7);
  val(finalTotalRow, 6, Number(d.finalTotalAmount || d.totalAmount || 0));
  fmt(finalTotalRow, 6, '#,##0"원"');
  applyStyle(finalTotalRow,6,finalTotalRow,7, {
    bg:NAVY, color:WHITE, bold:true, halign:'right', valign:'middle', border:true
  });

  cursorRow = finalTotalRow;
} else {
  const finalTotalRow = extraRow;
  merge(finalTotalRow, 1, finalTotalRow, 5);
  val(finalTotalRow, 1, '최 종 합 계');
  applyStyle(finalTotalRow,1,finalTotalRow,5, {
    bg:NAVY, color:WHITE, bold:true, halign:'center', valign:'middle', border:true
  });

  merge(finalTotalRow, 6, finalTotalRow, 7);
  val(finalTotalRow, 6, Number(d.finalTotalAmount || d.totalAmount || 0));
  fmt(finalTotalRow, 6, '#,##0"원"');
  applyStyle(finalTotalRow,6,finalTotalRow,7, {
    bg:NAVY, color:WHITE, bold:true, halign:'right', valign:'middle', border:true
  });

  cursorRow = finalTotalRow;
}

const termStart = cursorRow + 1;

// ══════════════════════════════════════
// 특약사항 섹션
// ══════════════════════════════════════
  sh.setRowHeight(termStart, 26);
  merge(termStart, 1, termStart, 7);
  val(termStart, 1, '■ 특 약 사 항');
  applyStyle(termStart,1,termStart,7, {bg:NAVY, color:WHITE, bold:true,
                                        size:11, halign:'left', valign:'middle'});

  const terms = [
    '본 견적서의 유효기간은 발행일로부터 30일입니다.',
    '체크인 18:00 / 체크아웃 12:00 입니다.',
    '예약 확정 후 취소 시 숙박일 기준 7일 전까지 전액 환불 가능합니다.',
    '부가세(VAT 10%)가 포함된 금액입니다.'
  ];

  // 추가 특약 (extraNote)
  const extraLines = (d.extraNote || '').split('\n').filter(l => l.trim());
  const allTerms   = [...terms, ...extraLines];

  allTerms.forEach((t, i) => {
    const r = termStart + 1 + i;
    sh.setRowHeight(r, 22);
    merge(r, 1, r, 7);
    val(r, 1, '• ' + t);
    applyStyle(r,1,r,7, {valign:'middle', border:true, wrap:true});
  });

  // 빈 줄 3개 (여백)
  for (let i = 0; i < 3; i++) {
    const r = termStart + 1 + allTerms.length + i;
    sh.setRowHeight(r, 18);
    merge(r, 1, r, 7);
    applyStyle(r,1,r,7, {border:true});
  }
}


// ── PDF 내보내기 URL ──────────────────────────────────────
function exportSheetToPdf_(ssId, shId) {
  const url =
    'https://docs.google.com/spreadsheets/d/' + ssId +
    '/export?exportFormat=pdf' +
    '&format=pdf' +
    '&size=A4' +
    '&portrait=true' +
    '&fitw=true' +
    '&sheetnames=false' +
    '&printtitle=false' +
    '&pagenumbers=false' +
    '&gridlines=false' +
    '&fzr=false' +
    '&gid=' + shId +
    '&ir=false' +
    '&ic=false' +
'&top_margin=0.60' +
'&bottom_margin=0.60' +
'&left_margin=0.75' +
'&right_margin=0.75' +
'&scale=2';

  const token = ScriptApp.getOAuthToken();
  const res   = UrlFetchApp.fetch(url, {
    headers : { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('PDF 내보내기 실패: HTTP ' + res.getResponseCode());
  }
  return res.getBlob().setContentType('application/pdf');
}

// ── Drive 폴더 ────────────────────────────────────────────
function getOrCreateFolder_(name) {
  const iter = DriveApp.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : DriveApp.createFolder(name);
}

// ── 인보이스 시트명 ───────────────────────────────────────
function getInvoiceSheetName_(quoteId) {
  return ('INV_' + quoteId).replace(/[:\\\/\?\*\[\]]/g, '_').substring(0, 100);
}

// ── 인보이스 템플릿 초기화 ────────────────────────────────
function initInvoiceTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName(Q_CFG.TMPL_SHEET);
  if (!sh) sh = ss.insertSheet(Q_CFG.TMPL_SHEET);
  Logger.log('인보이스 템플릿 시트 준비 완료');
}

// ── SMS 전송 ──────────────────────────────────────────────
function sendQuoteSms(quoteId, phoneOverride) {
  try {
    const sh   = ensureDbSheet_();
    const data = sh.getDataRange().getValues();
    const hdr  = data[0];
    const rowArr = data.slice(1).find(r => String(r[0]) === String(quoteId));

    if (!rowArr) {
      return {
        ok: false,
        success: false,
        message: '해당 견적을 찾을 수 없습니다. ID: ' + quoteId,
        error: '해당 견적을 찾을 수 없습니다. ID: ' + quoteId
      };
    }

    const rec = {};
    hdr.forEach((h, i) => { rec[h] = rowArr[i]; });

    const phone = String(phoneOverride || rec['customerPhone'] || '').trim();
    if (!phone) {
      return {
        ok: false,
        success: false,
        message: '전화번호가 없습니다.',
        error: '전화번호가 없습니다.'
      };
    }

    const msg = buildLmsMessage_(rec);
    const result = sendSms_(phone, msg);

    const rowIdx     = data.findIndex(r => String(r[0]) === String(quoteId)) + 1;
    const nowStr     = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    const colSentAt  = Q_CFG.DB_COLS.indexOf('lastSentAt')     + 1;
    const colSendTo  = Q_CFG.DB_COLS.indexOf('lastSendTo')     + 1;
    const colSendRes = Q_CFG.DB_COLS.indexOf('lastSendResult') + 1;

    sh.getRange(rowIdx, colSentAt).setValue(nowStr);
    sh.getRange(rowIdx, colSendTo).setValue("'" + phone);
    sh.getRange(rowIdx, colSendRes).setValue(result.ok ? '성공' : (result.error || '실패'));

    if (result.ok || result.success) {
      return {
        ok: true,
        success: true,
        message: 'SMS 발송이 완료되었습니다.',
        requestId: result.requestId || ''
      };
    }

    return {
      ok: false,
      success: false,
      message: result.error || 'SMS 발송 실패',
      error: result.error || 'SMS 발송 실패'
    };

  } catch (e) {
    return {
      ok: false,
      success: false,
      message: e.message,
      error: e.message
    };
  }
}

// ── LMS 메시지 빌더 ───────────────────────────────────────
function buildLmsMessage_(rec) {
  let items = [];
  try { items = JSON.parse(rec['itemsJson'] || '[]'); } catch(e) {}

  const lines = ['[호텔 레이블 성남 숙박 견적서]', ''];
  lines.push('■ 고객명: ' + (rec['customerName'] || ''));
  lines.push('■ 연락처: ' + (rec['customerPhone'] || ''));
  if (rec['vehicleType']) {
    lines.push('■ 차량: ' + rec['vehicleType'] + (rec['vehicleCount'] ? ' ' + rec['vehicleCount'] + '대' : ''));
  }
  lines.push('');
  lines.push('【 예약 내역 】');
  items.forEach((itm, i) => {
    lines.push((i + 1) + '. ' + (itm.roomType || '') +
               ' | ' + (itm.checkIn || '') + ' ~ ' + (itm.checkOut || '') +
               ' | ' + (itm.nights || 0) + '박' +
               ' | ' + Number(itm.subtotal || 0).toLocaleString() + '원');
  });
  lines.push('');
lines.push('■ 합계: ' + Number(rec['finalTotalAmount'] || rec['roomTotalAmount'] || 0).toLocaleString() + '원');
  if (rec['extraNote']) lines.push('■ 비고: ' + rec['extraNote']);
  lines.push('');
  if (rec['pdfUrl']) lines.push('■ 견적서 PDF: ' + rec['pdfUrl']);
  lines.push('');
  lines.push('※ 예약문의: 010-4657-6680,');
  lines.push('※ 본 견적은 7일간 유효합니다.');


  return lines.join('\n');
}

// ── 요금표 디버그 ─────────────────────────────────────────
function debugPriceTable() {
  const { headers, map } = loadPriceTable_();
  Logger.log('Headers: ' + JSON.stringify(headers));
  Logger.log('Rooms: ' + JSON.stringify(Object.keys(map)));
  Object.entries(map).forEach(([k, v]) => Logger.log(k + ' → ' + JSON.stringify(v)));
}
function debugInvoice() {
  const sh   = ensureDbSheet_();
  const data = sh.getDataRange().getValues();
  const hdr  = data[0];
  const row  = data[1]; // 가장 최근 저장된 견적 (2번째 행)
  
  const rec = {};
  hdr.forEach((h, i) => { rec[h] = row[i]; });
  
  Logger.log('quoteId: '      + rec['quoteId']);
  Logger.log('customerName: ' + rec['customerName']);
  Logger.log('customerPhone: '+ rec['customerPhone']);
  Logger.log('finalTotalAmount: ' + rec['finalTotalAmount']);
Logger.log('roomTotalAmount: ' + rec['roomTotalAmount']);
  Logger.log('itemsJson: '    + rec['itemsJson']);
}

/* ═══════════════════════════════════════════════════════════
   견적 이력관리 추가
   - 견적DB 조회
   - 고객명 / 전화번호 / 견적ID 검색
   - 상세조회
═══════════════════════════════════════════════════════════ */

// ── 이력관리 사이드바 열기 ────────────────────────────────
function openQuoteHistorySidebar() {
  const html = HtmlService.createHtmlOutputFromFile('QuoteHistory')
    .setTitle('견적 이력관리');
  SpreadsheetApp.getUi().showSidebar(html);
}

// ── DB 헤더 인덱스 맵 ────────────────────────────────────
function getQuoteDbHeaderMap_() {
  const sh = ensureDbSheet_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach(function(h, i) {
    map[String(h).trim()] = i;
  });
  return map;
}

// ── 전화번호 표시용 포맷 ─────────────────────────────────
function formatPhoneDisplay_(raw) {
  const num = String(raw || '').replace(/\D/g, '');
  if (!num) return '';

  if (num.length === 11) {
    return num.replace(/(\d{3})(\d{4})(\d{4})/, '$1-$2-$3');
  }
  if (num.length === 10) {
    if (num.startsWith('02')) {
      return num.replace(/(\d{2})(\d{4})(\d{4})/, '$1-$2-$3');
    }
    return num.replace(/(\d{3})(\d{3})(\d{4})/, '$1-$2-$3');
  }
  return String(raw || '');
}

// ── items 요약 문자열 ────────────────────────────────────
function buildItemsSummary_(items) {
  if (!Array.isArray(items) || !items.length) return '';
  return items.map(function(it) {
    const roomType = it.roomType || '';
    const rooms = Number(it.rooms || 1);
    return roomType + ' ' + rooms + '실';
  }).join(', ');
}

// ── 날짜 표시 정리 ───────────────────────────────────────
function formatDateTimeDisplay_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  }
  return String(v);
}

// ── 견적 목록 조회 ───────────────────────────────────────
function getQuoteHistoryList(params) {
  try {
    const keyword = String((params && params.keyword) || '').trim();
    const field   = String((params && params.field) || 'all').trim();

    const sh = ensureDbSheet_();
    const data = sh.getDataRange().getValues();
    if (data.length < 2) {
      return { ok: true, rows: [] };
    }

    const headerMap = getQuoteDbHeaderMap_();
    const rows = [];

    for (let r = 1; r < data.length; r++) {
      const row = data[r];

      const quoteId = String(row[headerMap['quoteId']] || '').trim();
      if (!quoteId) continue;

      const createdAt = row[headerMap['createdAt']] || '';
      const updatedAt = row[headerMap['updatedAt']] || '';
      const customerName = row[headerMap['customerName']] || '';
      const customerPhone = row[headerMap['customerPhone']] || '';

const totalAmountRaw =
  row[headerMap['finalTotalAmount']] != null && row[headerMap['finalTotalAmount']] !== ''
    ? row[headerMap['finalTotalAmount']]
    : row[headerMap['roomTotalAmount']];

let totalAmount = Number(totalAmountRaw);
if (isNaN(totalAmount)) totalAmount = 0;

      const pdfUrl = row[headerMap['pdfUrl']] || '';
      const lastSentAt = row[headerMap['lastSentAt']] || '';
      const lastSendTo = row[headerMap['lastSendTo']] || '';
      const lastSendResult = row[headerMap['lastSendResult']] || '';
      const itemsJson = row[headerMap['itemsJson']] || '[]';

      const items = safeJsonParse_(itemsJson) || [];
      const itemsSummary = buildItemsSummary_(items);

      const nameStr = String(customerName).toLowerCase();
      const phoneNum = String(customerPhone).replace(/\D/g, '');
      const quoteStr = quoteId.toLowerCase();

      let matched = true;

      if (keyword) {
        if (field === 'customerName') {
          matched = nameStr.includes(keyword.toLowerCase());
        } else if (field === 'customerPhone') {
          matched = phoneNum.includes(String(keyword).replace(/\D/g, ''));
        } else if (field === 'quoteId') {
          matched = quoteStr.includes(keyword.toLowerCase());
        } else {
          matched =
            nameStr.includes(keyword.toLowerCase()) ||
            phoneNum.includes(String(keyword).replace(/\D/g, '')) ||
            quoteStr.includes(keyword.toLowerCase());
        }
      }

      if (!matched) continue;

      rows.push({
        quoteId: quoteId,
        createdAt: formatDateTimeDisplay_(createdAt),
        updatedAt: formatDateTimeDisplay_(updatedAt),
        customerName: String(customerName),
        customerPhone: formatPhoneDisplay_(customerPhone),
        customerPhoneRaw: String(customerPhone || ''),
        itemsSummary: itemsSummary,
        totalAmount: totalAmount,
        totalAmountDisplay: totalAmount.toLocaleString() + '원',
        pdfUrl: String(pdfUrl || ''),
        lastSentAt: formatDateTimeDisplay_(lastSentAt),
        lastSendTo: formatPhoneDisplay_(lastSendTo),
        lastSendResult: String(lastSendResult || '')
      });
    }

    rows.sort(function(a, b) {
      const aTime = new Date(a.updatedAt || a.createdAt || 0).getTime();
      const bTime = new Date(b.updatedAt || b.createdAt || 0).getTime();
      return bTime - aTime;
    });

    return { ok: true, rows: rows.slice(0, 200) };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}


function getQuoteBasicDetailNoLogs_(quoteId) {
  const qid = String(quoteId || '').trim();
  if (!qid) return null;

  const sh = ensureDbSheet_();
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return null;

  const headerMap = getQuoteDbHeaderMap_();
  let foundRow = null;

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][headerMap['quoteId']]) === qid) {
      foundRow = data[r];
      break;
    }
  }

  if (!foundRow) return null;

  const rec = {};
  Object.keys(headerMap).forEach(function(key) {
    rec[key] = foundRow[headerMap[key]];
  });

  let items = [];
  try { items = JSON.parse(rec.itemsJson || '[]'); } catch(e) {}

return {
  quoteId: String(rec.quoteId || ''),
  customerName: String(rec.customerName || ''),
  totalAmount: Number(rec.finalTotalAmount || rec.roomTotalAmount || 0),
  totalAmountDisplay: Number(rec.finalTotalAmount || rec.roomTotalAmount || 0).toLocaleString() + '원',
  items: items.map(function(it) {
      return {
        roomType: it.roomType || '',
        rooms: Number(it.rooms || 1),
        unitPrice: Number(it.unitPrice || 0),
        unitPriceDisplay: Number(it.unitPrice || 0).toLocaleString() + '원',
        checkIn: it.checkIn || '',
        checkOut: it.checkOut || '',
        nights: Number(it.nights || 0),
        subtotal: Number(it.subtotal || 0),
        subtotalDisplay: Number(it.subtotal || 0).toLocaleString() + '원'
      };
    })
  };
}


// ── 견적 상세 조회 ───────────────────────────────────────
function getQuoteHistoryDetail(quoteId) {
  try {
    const qid = String(quoteId || '').trim();
    if (!qid) return { ok: false, error: '견적ID가 없습니다.' };

    const sh = ensureDbSheet_();
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok: false, error: '견적 데이터가 없습니다.' };

    const headerMap = getQuoteDbHeaderMap_();
    let foundRow = null;

    for (let r = 1; r < data.length; r++) {
      if (String(data[r][headerMap['quoteId']]) === qid) {
        foundRow = data[r];
        break;
      }
    }

    if (!foundRow) {
      return { ok: false, error: '해당 견적을 찾을 수 없습니다: ' + qid };
    }

    const rec = {};
    Object.keys(headerMap).forEach(function(key) {
      rec[key] = foundRow[headerMap[key]];
    });

    const items = safeJsonParse_(rec.itemsJson || '[]') || [];
    const dateSummary = safeJsonParse_(rec.dateSummaryJson || '{}') || {};

const detailItems = items.map(function(it) {

  let nights = Number(it.nights || 0);
  let unitPrice = Number(it.unitPrice || 0);
  let subtotal = Number(it.subtotal || 0);

  if (!subtotal && it.roomType && it.checkIn && it.checkOut) {
    const preview = calcItemPreview(it.roomType, it.checkIn, it.checkOut);
    if (preview && preview.ok) {
      nights = preview.nights;
      unitPrice = preview.unitPrice;
      subtotal = preview.subtotal;
    }
  }

  return {
    roomType: it.roomType || '',
    rooms: Number(it.rooms || 1),
    unitPrice: unitPrice,
    unitPriceDisplay: unitPrice.toLocaleString() + '원',
    checkIn: it.checkIn || '',
    checkOut: it.checkOut || '',
    nights: nights,
    subtotal: subtotal,
    subtotalDisplay: subtotal.toLocaleString() + '원'
  };
});

    return {
      ok: true,
detail: {
  quoteId: String(rec.quoteId || ''),
  createdAt: formatDateTimeDisplay_(rec.createdAt),
  updatedAt: formatDateTimeDisplay_(rec.updatedAt),
  customerName: String(rec.customerName || ''),
customerPhone: formatPhoneDisplay_(String(rec.customerPhone || '').replace(/^'/, '')),
customerPhoneRaw: String(rec.customerPhone || '').replace(/^'/, ''),
roomTotalAmount: Number(rec.roomTotalAmount || 0),
roomTotalAmountDisplay: Number(rec.roomTotalAmount || 0).toLocaleString() + '원',

extraLabel1: String(rec.extraLabel1 || ''),
extraAmount1: Number(rec.extraAmount1 || 0),
extraAmount1Display: Number(rec.extraAmount1 || 0).toLocaleString() + '원',

extraLabel2: String(rec.extraLabel2 || ''),
extraAmount2: Number(rec.extraAmount2 || 0),
extraAmount2Display: Number(rec.extraAmount2 || 0).toLocaleString() + '원',

extraLabel3: String(rec.extraLabel3 || ''),
extraAmount3: Number(rec.extraAmount3 || 0),
extraAmount3Display: Number(rec.extraAmount3 || 0).toLocaleString() + '원',

extraAmountTotal: Number(rec.extraAmountTotal || 0),
extraAmountTotalDisplay: Number(rec.extraAmountTotal || 0).toLocaleString() + '원',

totalAmount: Number(rec.finalTotalAmount || rec.roomTotalAmount || 0),
totalAmountDisplay: Number(rec.finalTotalAmount || rec.roomTotalAmount || 0).toLocaleString() + '원',

vehicleType: String(rec.vehicleType || ''),
vehicleCount: String(rec.vehicleCount || ''),
employeeName: String(rec.employeeName || ''),
extraNote: String(rec.extraNote || ''),
        pdfUrl: String(rec.pdfUrl || ''),
        pdfFileId: String(rec.pdfFileId || ''),
        lastSentAt: formatDateTimeDisplay_(rec.lastSentAt),
        lastSendTo: formatPhoneDisplay_(rec.lastSendTo),
        lastSendResult: String(rec.lastSendResult || ''),
        items: detailItems,
        dateSummary: dateSummary,
        logs: getQuoteLogs_(qid)
      }
    };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
// ── 견적 이력 삭제 ───────────────────────────────────────
function deleteQuoteHistory(quoteId) {
  try {
    const qid = String(quoteId || '').trim();
    if (!qid) return { ok: false, error: '견적ID가 없습니다.' };

    const sh = ensureDbSheet_();
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return { ok: false, error: '삭제할 데이터가 없습니다.' };

    const headerMap = getQuoteDbHeaderMap_();
    const quoteColIdx = headerMap['quoteId'];

    let rowIdx = -1;
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][quoteColIdx]) === qid) {
        rowIdx = r + 1; // 시트 row 번호
        break;
      }
    }

    if (rowIdx < 0) {
      return { ok: false, error: '해당 견적을 찾을 수 없습니다: ' + qid };
    }

    sh.deleteRow(rowIdx);
    return { ok: true, quoteId: qid };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
// ── 견적 수정용: 기존 견적 데이터로 사이드바 열기 ─────────────────


function getQuoteEditPayload_(quoteId) {
  Logger.log('getQuoteEditPayload_ raw quoteId=' + quoteId);
  Logger.log('getQuoteEditPayload_ type=' + (typeof quoteId));

  const qid = String(quoteId || '').trim();
  Logger.log('getQuoteEditPayload_ trimmed qid=' + qid);

  if (!qid) throw new Error('견적ID가 없습니다.');

  const sh = ensureDbSheet_();
  const data = sh.getDataRange().getValues();
  if (data.length < 2) throw new Error('견적 데이터가 없습니다.');

  const headerMap = getQuoteDbHeaderMap_();
  let foundRow = null;

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][headerMap['quoteId']]) === qid) {
      foundRow = data[r];
      break;
    }
  }

  if (!foundRow) throw new Error('해당 견적을 찾을 수 없습니다: ' + qid);

  const rec = {};
  Object.keys(headerMap).forEach(function(key) {
    rec[key] = foundRow[headerMap[key]];
  });

  let items = [];
  try { items = JSON.parse(rec.itemsJson || '[]'); } catch(e) {}

return {
  quoteId: '',
  sourceQuoteId: String(rec.quoteId || ''),
  guestName: String(rec.customerName || ''),
  guestPhone: String(rec.customerPhone || '').replace(/^'/, ''),
  employeeName: String(rec.employeeName || ''),
  vehicleType: String(rec.vehicleType || ''),
  vehicleCount: Number(rec.vehicleCount || 0),
  extraNote: String(rec.extraNote || ''),

  priceDisplayMode: String(rec.priceDisplayMode || 'avg'),

  extraLabel1: String(rec.extraLabel1 || ''),
  extraAmount1: Number(rec.extraAmount1 || 0),
  extraLabel2: String(rec.extraLabel2 || ''),
  extraAmount2: Number(rec.extraAmount2 || 0),
  extraLabel3: String(rec.extraLabel3 || ''),
  extraAmount3: Number(rec.extraAmount3 || 0),

  totalAmount: Number(rec.finalTotalAmount || rec.roomTotalAmount || 0),

  roomItems: items.map(function(it) {
    return {
      roomType: it.roomType || '',
      checkIn: it.checkIn || '',
      checkOut: it.checkOut || '',
      nights: Number(it.nights || 1),
      rooms: Number(it.rooms || 1),
      unitPrice: Number(it.unitPrice || 0),
      subtotal: Number(it.subtotal || 0)
    };
  })
};
}

function openQuoteEditSidebar(quoteId) {
  const qid = String(quoteId || '').trim();
  if (!qid) {
    throw new Error('견적ID가 없습니다.');
  }

  const payload = getQuoteEditPayload_(qid);
  payload.items = Array.isArray(payload.roomItems) ? payload.roomItems : [];

  const tpl = HtmlService.createTemplateFromFile('Sidebar');
  tpl.initialQuoteData = JSON.stringify(payload);

  const html = tpl.evaluate()
    .setWidth(960)
    .setHeight(820);

  SpreadsheetApp.getUi().showModalDialog(html, '견적서 수정');
}

function openQuoteDetailDialog(quoteId) {
  const res = getQuoteHistoryDetail(quoteId);
  if (!res || !res.ok) {
    throw new Error(res && res.error ? res.error : '견적 상세를 불러오지 못했습니다.');
  }

  const tpl = HtmlService.createTemplateFromFile('QuoteDetail');
  tpl.detailJson = JSON.stringify(res.detail || {});

  const html = tpl.evaluate()
    .setWidth(920)
    .setHeight(760);

  SpreadsheetApp.getUi().showModalDialog(html, '견적 상세');
}
function ensureQuoteLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('견적로그');
  if (!sh) {
    sh = ss.insertSheet('견적로그');
    sh.getRange(1, 1, 1, 8).setValues([[
      'ts',
      'quoteId',
      'sourceQuoteId',
      'action',
      'employeeName',
      'customerName',
      'changeSummary',
      'snapshotJson'
    ]]);
  }
  return sh;
}

function simplifyItemsForCompare_(items) {
  return (items || []).map(function(it) {
    return {
      roomType: it.roomType || '',
      checkIn: it.checkIn || '',
      checkOut: it.checkOut || '',
      nights: Number(it.nights || 0),
      rooms: Number(it.rooms || 0),
      unitPrice: Number(it.unitPrice || 0),
      subtotal: Number(it.subtotal || 0)
    };
  });
}

function buildChangeSummary_(oldDetail, newPayload, recalculatedItems, recalculatedTotal) {
  if (!oldDetail) {
    return '최초 견적 생성';
  }

  const changes = [];

  // 1) 고객 기본정보
  if ((oldDetail.customerName || '') !== (newPayload.guestName || '')) {
    changes.push(
      '고객명이 "' + (oldDetail.customerName || '-') + '" → "' + (newPayload.guestName || '-') + '" 변경'
    );
  }

  if ((oldDetail.customerPhoneRaw || '') !== (newPayload.guestPhone || '')) {
    changes.push('전화번호 변경');
  }

  if ((oldDetail.vehicleType || '') !== (newPayload.vehicleType || '')) {
    changes.push(
      '차량종류가 "' + (oldDetail.vehicleType || '없음') + '" → "' + (newPayload.vehicleType || '없음') + '" 변경'
    );
  }

  if (String(oldDetail.vehicleCount || '') !== String(newPayload.vehicleCount || '')) {
    changes.push(
      '차량대수가 "' + (oldDetail.vehicleCount || 0) + '대" → "' + (newPayload.vehicleCount || 0) + '대" 변경'
    );
  }

  if ((oldDetail.extraNote || '') !== (newPayload.extraNote || '')) {
    changes.push('비고 변경');
  }

  // 2) 객실내역 비교용 문자열
  function itemToText_(it) {
    return [
      it.roomType || '-',
      (it.checkIn || '-') + '~' + (it.checkOut || '-'),
      (Number(it.rooms || 1)) + '실',
      (Number(it.nights || 0)) + '박',
      (Number(it.subtotal || 0)).toLocaleString() + '원'
    ].join(' / ');
  }

  const oldItems = simplifyItemsForCompare_(oldDetail.items || []);
  const newItems = simplifyItemsForCompare_(recalculatedItems || []);

  if (JSON.stringify(oldItems) !== JSON.stringify(newItems)) {
    if (oldItems.length === 1 && newItems.length === 1) {
      changes.push(
        '객실내역이 [' + itemToText_(oldItems[0]) + '] → [' + itemToText_(newItems[0]) + '] 변경'
      );
    } else {
      changes.push(
        '객실내역 변경 (' + oldItems.length + '건 → ' + newItems.length + '건)'
      );
    }
  }

  // 3) 총금액 비교
  const oldTotal = Number(oldDetail.totalAmount || 0);
  const newTotal = Number(recalculatedTotal || 0);

  if (oldTotal !== newTotal) {
    changes.push(
      '총금액이 ' + oldTotal.toLocaleString() + '원 → ' + newTotal.toLocaleString() + '원 변경'
    );
  }

  return changes.length ? changes.join(' / ') : '변경 없음';
}

function appendQuoteLog_(quoteId, sourceQuoteId, action, employeeName, customerName, changeSummary, snapshotObj) {
  const sh = ensureQuoteLogSheet_();
  sh.appendRow([
    Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss'),
    quoteId || '',
    sourceQuoteId || '',
    action || '',
    employeeName || '',
    customerName || '',
    changeSummary || '',
    JSON.stringify(snapshotObj || {})
  ]);
}

function getQuoteLogs_(quoteId) {
  const sh = ensureQuoteLogSheet_();
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const out = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const logTs = row[0];
    const logQuoteId = String(row[1] || '');
    const logSourceQuoteId = String(row[2] || '');
    const logAction = String(row[3] || '');
    const logEmployeeName = String(row[4] || '');
    const logCustomerName = String(row[5] || '');
    const logChangeSummary = String(row[6] || '');
    const logSnapshotJson = String(row[7] || '{}');

    if (logQuoteId === String(quoteId) || logSourceQuoteId === String(quoteId)) {
      let snapshot = {};
      try { snapshot = JSON.parse(logSnapshotJson || '{}'); } catch(e) {}

      const beforeDetail = logSourceQuoteId ? getQuoteBasicDetailNoLogs_(logSourceQuoteId) : null;

      const beforeItemsRaw = snapshot.beforeItems || (beforeDetail ? beforeDetail.items : []);
      const afterItemsRaw  = snapshot.afterItems || snapshot.items || [];

      const beforeItems = (beforeItemsRaw || []).map(function(it) {
        return {
          roomType: it.roomType || '',
          rooms: Number(it.rooms || 1),
          unitPrice: Number(it.unitPrice || 0),
          unitPriceDisplay: Number(it.unitPrice || 0).toLocaleString() + '원',
          checkIn: it.checkIn || '',
          checkOut: it.checkOut || '',
          nights: Number(it.nights || 0),
          subtotal: Number(it.subtotal || 0),
          subtotalDisplay: Number(it.subtotal || 0).toLocaleString() + '원'
        };
      });

      const afterItems = (afterItemsRaw || []).map(function(it) {
        return {
          roomType: it.roomType || '',
          rooms: Number(it.rooms || 1),
          unitPrice: Number(it.unitPrice || 0),
          unitPriceDisplay: Number(it.unitPrice || 0).toLocaleString() + '원',
          checkIn: it.checkIn || '',
          checkOut: it.checkOut || '',
          nights: Number(it.nights || 0),
          subtotal: Number(it.subtotal || 0),
          subtotalDisplay: Number(it.subtotal || 0).toLocaleString() + '원'
        };
      });

      const beforeTotalAmount = Number(
        snapshot.beforeTotalAmount != null
          ? snapshot.beforeTotalAmount
          : (beforeDetail ? beforeDetail.totalAmount : 0)
      );

      const afterTotalAmount = Number(
        snapshot.afterTotalAmount != null
          ? snapshot.afterTotalAmount
          : (snapshot.totalAmount || 0)
      );

      out.push({
        ts: formatDateTimeDisplay_(logTs),
        quoteId: logQuoteId,
        sourceQuoteId: logSourceQuoteId,
        action: logAction,
        employeeName: logEmployeeName,
        customerName: logCustomerName,
        changeSummary: logChangeSummary,

        beforeQuoteId: beforeDetail ? beforeDetail.quoteId : logSourceQuoteId,
        beforeTotalAmount: beforeTotalAmount,
        beforeTotalAmountDisplay: beforeTotalAmount ? beforeTotalAmount.toLocaleString() + '원' : '-',
        beforeItems: beforeItems,

        afterQuoteId: logQuoteId,
        afterTotalAmount: afterTotalAmount,
        afterTotalAmountDisplay: afterTotalAmount ? afterTotalAmount.toLocaleString() + '원' : '-',
        afterItems: afterItems
      });
    }
  }

  return out;
}

function debugQuoteRoomTypes() {
  const sh = ensureDbSheet_();
  const data = sh.getDataRange().getValues();
  const headerMap = getQuoteDbHeaderMap_();

  for (let r = 1; r < data.length; r++) {
    const quoteId = String(data[r][headerMap['quoteId']] || '').trim();
    if (!quoteId) continue;

    const itemsJson = data[r][headerMap['itemsJson']] || '[]';
    let items = [];
    try { items = JSON.parse(itemsJson); } catch (e) {}

    Logger.log('quoteId=' + quoteId);
    Logger.log('items=' + JSON.stringify(items));
  }

  const loaded = loadPriceTable_();
  Logger.log('price keys=' + JSON.stringify(Object.keys(loaded.map)));
}
function debugPreviewDirect() {
  Logger.log('디럭스=' + JSON.stringify(calcItemPreview('디럭스', '2026-03-10', '2026-03-20')));
  Logger.log('스탠다드=' + JSON.stringify(calcItemPreview('스탠다드', '2026-03-09', '2026-03-13')));
}
// ── SMS 발송 (SENS API 기반) ──────────────────────────────────
function sendSms_(phone, message) {
  try {
    const props = PropertiesService.getScriptProperties();
    const serviceId = props.getProperty('SENS_SERVICE_ID');
    const accessKey = props.getProperty('SENS_ACCESS_KEY');
    const secretKey = props.getProperty('SENS_SECRET_KEY');
    const from      = props.getProperty('SENS_FROM');

    if (!serviceId || !accessKey || !secretKey || !from) {
      return { 
        ok: false, 
        error: '설정값 없음. 초기설정을 먼저 실행하세요.' 
      };
    }

    const cleanPhone = String(phone || '').replace(/[^0-9]/g, '');
    if (!cleanPhone) {
      return { ok: false, error: '전화번호가 유효하지 않습니다.' };
    }

    const method = 'POST';
    const uri = '/sms/v2/services/' + serviceId + '/messages';
    const url = 'https://sens.apigw.ntruss.com' + uri;
    const timestamp = String(Date.now());
    const signature = makeSignature_(method, uri, timestamp, accessKey, secretKey);

    const payload = {
      type: 'SMS',
      countryCode: '82',
      from: from,
      content: message,
      messages: [{ to: cleanPhone }]
    };

    const resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      headers: {
        'x-ncp-apigw-timestamp': timestamp,
        'x-ncp-iam-access-key': accessKey,
        'x-ncp-apigw-signature-v2': signature
      }
    });

    const code = resp.getResponseCode();
    const json = safeJsonParse_(resp.getContentText());




    function normalizePhoneText_(raw) {
  return String(raw || '').replace(/[^\d]/g, '').trim();
}

function preservePhoneText_(raw) {
  var phone = normalizePhoneText_(raw);
  return phone ? "'" + phone : '';
}

    if (code >= 200 && code < 300) {
      return {
        ok: true,
        success: true,
        httpCode: code,
        requestId: json && json.requestId
      };
    }

    return {
      ok: false,
      success: false,
      httpCode: code,
      error: (json && json.errorMessage) || resp.getContentText()
    };

  } catch (e) {
    return { ok: false, success: false, error: String(e) };
  }
}

// ── HMAC-SHA256 서명 생성 ───────────────────────────────────
function makeSignature_(method, uri, timestamp, accessKey, secretKey) {
  const message = method + ' ' + uri + '\n' + timestamp + '\n' + accessKey;
  return Utilities.base64Encode(
    Utilities.computeHmacSha256Signature(message, secretKey, Utilities.Charset.UTF_8)
  );
}

// ── 안전한 JSON 파싱 ────────────────────────────────────────
function safeJsonParse_(s) {
  try {
    return JSON.parse(s);
  } catch (_) {
    return null;
  }
}
function syncQuoteDbHeader_() {
  const sh = ensureDbSheet_();
  const currentHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const targetHeaders = Q_CFG.DB_COLS.slice();

  let changed = false;

  targetHeaders.forEach(function(h, idx) {
    if (currentHeaders[idx] !== h) {
      sh.getRange(1, idx + 1).setValue(h);
      changed = true;
    }
  });

  if (sh.getLastColumn() < targetHeaders.length) {
    sh.insertColumnsAfter(sh.getLastColumn(), targetHeaders.length - sh.getLastColumn());
    sh.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
    changed = true;
  }

  return changed;
}
function normalizePhoneText_(raw) {
  return String(raw || '').replace(/[^\d]/g, '').trim();
}

function preservePhoneText_(raw) {
  var phone = normalizePhoneText_(raw);
  return phone ? "'" + phone : '';
}
function parseMoney_(v) {
  return Number(String(v || '').replace(/[^\d.-]/g, '')) || 0;
}

function buildDailyPriceLines_(item) {
  const roomType = String(item.roomType || '').trim();
  const checkIn = parseYmd_(item.checkIn);
  const checkOut = parseYmd_(item.checkOut);
  if (!roomType || !checkIn || !checkOut || checkOut <= checkIn) return [];

  const { map } = loadPriceTable_();
  const priceRow = getPriceRow_(map, roomType);
  if (!priceRow) return [];

  const dowMap = { 0:'일', 1:'월', 2:'화', 3:'수', 4:'목', 5:'금', 6:'토' };
  const cur = new Date(checkIn);
  const lines = [];

  while (cur < checkOut) {
    const dow = dowMap[cur.getDay()];
    const col = Object.keys(priceRow).find(k => String(k).includes(dow));
    const price = col ? (Number(priceRow[col]) || 0) : 0;

    lines.push({
      date: Utilities.formatDate(new Date(cur), 'Asia/Seoul', 'yyyy-MM-dd'),
      dow: dow,
      amount: price
    });

    cur.setDate(cur.getDate() + 1);
  }

  return lines;
}
function testQuoteHistoryList() {
  const res = getQuoteHistoryList({ keyword: '', field: 'all' });
  Logger.log(JSON.stringify(res));
}
