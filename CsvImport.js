/* =========================================================
   ✅ CSV → 숙박일지 자동 입력 (덮어쓰기 방지)
   - 취소건 제외
   - 자동입력 표시: 대실(B열)/숙박(N열)에 'V'
   - 객실타입은 특이사항/차종(J/U)에 기록 (의도)
   - 직원 수동 입력(호실/차번호 등)은 절대 건드리지 않음:
     → 해당 섹션에서 "한 칸이라도 값이 있는 행"은 자동이 피함
   - 정렬은 "자동(V) 행만" 시간순 정렬, 수동 행 위치는 고정 (요청 B)
========================================================= */

var CHANNEL_MAP = {
  "아고다":"아","야놀자":"야","여기어때":"여","꿀스테이":"꿀",
  "트립닷컴":"트립닷컴","부킹닷컴":"부킹","익스피디아":"익스",
  "신용카드":"카","현장결제":"현장"
};

/* =========================
 * UI
 * ========================= */
function openCsvImportDialog() {
  var html = HtmlService.createHtmlOutputFromFile('CsvImportDialog')
    .setWidth(480).setHeight(320).setTitle('CSV');
  SpreadsheetApp.getUi().showModalDialog(html, 'CSV Import');
}

/* =========================
 * 메인
 * ========================= */
function importReservationsFromCsv(csvText) {
  var parsed = parseReservationCsv_(csvText);
  if (!parsed.ok) return parsed; // {ok:false, message:"..."}

  var rows = parsed.rows;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var inserted = 0;
  var noTab = [];
  var touchedTabs = {};

  // 탭별/섹션별로 모아서 한번에 쓰기
  // bucket[tabName] = { daeSil: [rowData...], sukBak: [rowData...] }
  var bucket = {};

  rows.forEach(function(r) {
    if (r.status === '취소됨') return;

    var nights = getDateRange_(r.checkIn, r.checkOut);
    nights.forEach(function(dateObj, idx) {
      var tabName = getTabName_(ss, dateObj);
      if (!tabName) {
        var label = Utilities.formatDate(dateObj, 'Asia/Seoul', 'M/d');
        if (noTab.indexOf(label) === -1) noTab.push(label);
        return;
      }

      if (!bucket[tabName]) bucket[tabName] = { daeSil: [], sukBak: [] };

      var isFirst = (idx === 0);
      var rowData = buildRowData_(r, isFirst);

      if (r.type === '대실') bucket[tabName].daeSil.push(rowData);
      else bucket[tabName].sukBak.push(rowData);

      touchedTabs[tabName] = true;
    });
  });

  // 탭별 배치 write (덮어쓰기 방지 정책)
  Object.keys(bucket).forEach(function(tabName) {
    var tab = ss.getSheetByName(tabName);
    if (!tab) return;

    var b = bucket[tabName];

    var res1 = writeReservationRowsBatchNoOverwrite_(tab, '대실', b.daeSil);
    inserted += res1.inserted;

    var res2 = writeReservationRowsBatchNoOverwrite_(tab, '숙박', b.sukBak);
    inserted += res2.inserted;
  });

  // ✅ 정렬: 자동(V) 행만 시간순 정렬, 수동 행은 위치 고정
  Object.keys(touchedTabs).forEach(function(name) {
    var t = ss.getSheetByName(name);
    if (!t) return;

    // 대실: B~L (11칸), 시간은 C(섹션 index 1)
    sortSectionAutoOnly_(t, 4, 2, 11, 1);

    // 숙박: N~X (11칸), 시간은 O(섹션 index 1)
    sortSectionAutoOnly_(t, 4, 14, 11, 1);
  });

  var msg = 'OK: ' + inserted;
  if (noTab.length > 0) msg += '\nNO TAB: ' + noTab.join(', ');
  return { ok: true, message: msg };
}

/* =========================
 * CSV 파싱 (표준 CSV)
 * ========================= */
function parseReservationCsv_(csvText) {
  if (!csvText || !String(csvText).trim()) {
    return { ok: false, message: 'CSV 내용이 비었습니다.' };
  }

  // BOM 제거
  csvText = String(csvText).replace(/^\uFEFF/, '');

  var table;
  try {
    table = Utilities.parseCsv(csvText);
  } catch (e) {
    return { ok: false, message: 'CSV 파싱 실패: ' + e };
  }

  // 빈 줄 제거
  table = table.filter(function(row) {
    return row && row.some(function(c) { return String(c || '').trim() !== ''; });
  });

  if (!table.length || table.length < 2) {
    return { ok: false, message: 'CSV에 헤더/데이터가 부족합니다.' };
  }

  var header = table[0].map(function(h) { return normalizeHeader_(h); });

  var col = {
    checkIn:  header.indexOf(normalizeHeader_('입실일시')),
    checkOut: header.indexOf(normalizeHeader_('퇴실일시')),
    type:     header.indexOf(normalizeHeader_('예약타입')),
    roomType: header.indexOf(normalizeHeader_('객실타입')),
    name:     header.indexOf(normalizeHeader_('결제자명')),
    amount:   header.indexOf(normalizeHeader_('총결제금액')),
    status:   header.indexOf(normalizeHeader_('결제상태')),
    channel:  header.indexOf(normalizeHeader_('결제방식'))
  };

  // 필수 컬럼 검증
  var required = [
    ['입실일시','checkIn'],
    ['퇴실일시','checkOut'],
    ['예약타입','type'],
    ['결제자명','name'],
    ['총결제금액','amount'],
    ['결제상태','status'],
    ['결제방식','channel']
  ];
  var missing = required.filter(function(pair){
    return col[pair[1]] === -1;
  }).map(function(pair){ return pair[0]; });

  if (missing.length) {
    return { ok: false, message: 'CSV 헤더에 다음 컬럼이 없습니다: ' + missing.join(', ') };
  }

  var rows = table.slice(1).map(function(cells) {
    var chRaw = String(cells[col.channel] || '');
    var checkInStr  = String(cells[col.checkIn]  || '');
    var checkOutStr = String(cells[col.checkOut] || '');

    var checkIn  = parseDateTime_(checkInStr);
    var checkOut = parseDateTime_(checkOutStr);

    return {
      checkIn:  checkIn,
      checkOut: checkOut,
      type:     String(cells[col.type] || '').trim(),
      roomType: String(cells[col.roomType] || '').trim(),
      name:     String(cells[col.name] || '').trim(),
      amount:   parseAmount_(cells[col.amount]),
      status:   String(cells[col.status] || '').trim(),
      channel:  extractChannel_(chRaw)
    };
  }).filter(function(r) {
    var validType = (r.type === '대실' || r.type === '숙박');
    var validDate = r.checkIn instanceof Date && !isNaN(r.checkIn.getTime());
    var validName = (r.name || '').trim() !== '';
    var validAmt  = r.amount > 0;
    return validType && validDate && validName && validAmt;
  });

  return { ok: true, rows: rows };
}

function normalizeHeader_(h) {
  return String(h || '')
    .replace(/^\uFEFF/, '')
    .replace(/"/g, '')
    .trim()
    .toLowerCase();
}

function parseAmount_(v) {
  var s = String(v || '').trim();
  if (!s) return 0;
  s = s.replace(/,/g, '');
  var n = parseFloat(s);
  return isFinite(n) ? n : 0;
}

/**
 * 날짜/시간 파서
 * 지원 예:
 * - 2026-03-05 14:00
 * - 2026/03/05 14:00:00
 * - 2026.03.05 14:00
 * - 2026-03-05T14:00:00
 */
function parseDateTime_(s) {
  s = String(s || '').trim();
  if (!s) return new Date(''); // Invalid

  // ISO 계열
  if (s.indexOf('T') !== -1) {
    var dIso = new Date(s);
    if (!isNaN(dIso.getTime())) return dIso;
  }

  var t = s.replace(/\./g, '-').replace(/\//g, '-');

  var m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    var y = parseInt(m[1], 10);
    var mo = parseInt(m[2], 10) - 1;
    var da = parseInt(m[3], 10);
    var hh = m[4] ? parseInt(m[4], 10) : 0;
    var mm = m[5] ? parseInt(m[5], 10) : 0;
    var ss = m[6] ? parseInt(m[6], 10) : 0;
    return new Date(y, mo, da, hh, mm, ss);
  }

  // fallback
  return new Date(s);
}

/* =========================
 * 행 데이터 구성 (의도 유지)
 * ========================= */
function buildRowData_(r, isFirst) {
  if (r.type === '대실') {
    return {
      type: '대실',
      checkIn: formatHour_(r.checkIn),
      checkOut: formatHour_(r.checkOut),
      amount: r.amount,
      channel: r.channel,
      name: r.name,
      roomType: r.roomType
    };
  }

  // 숙박
  if (isFirst) {
    return {
      type: '숙박',
      checkIn: formatHour_(r.checkIn),
      amount: r.amount,
      channel: r.channel,
      name: r.name,
      roomType: r.roomType
    };
  }

  // 연박: 둘째날부터는 '-' (의도)
  return {
    type: '숙박',
    checkIn: ' - ',
    amount: ' - ',
    channel: ' - ',
    name: r.name,
    roomType: '' // 둘째날 이후 객실타입은 비움
  };
}

/* =========================================================
 * ✅ 덮어쓰기 방지 배치 입력
 * - 섹션 범위 내에서 "한 칸이라도 값 있으면" 그 행은 자동이 피함
 * - 완전 빈 행에만 기록
 * - 빈 행이 없으면 아래로 append
 *
 * 템플릿 매핑(확정):
 * [대실] B~L(11칸)
 *   B: V표시 / C:입실 / D:퇴실 / E:금액 / F:결제(채널) / I:이름 / J:특이사항/차종(객실타입)
 * [숙박] N~X(11칸)
 *   N: V표시 / O:시간 / P:금액 / Q:결제(채널) / T:이름 / U:특이사항/차종(객실타입)
========================================================= */
function writeReservationRowsBatchNoOverwrite_(tab, type, rowDataList) {
  if (!rowDataList || rowDataList.length === 0) return { inserted: 0 };

  var cfg = (type === '대실')
    ? { startRow: 4, startCol: 2,  numCols: 11, // B~L
        idx: { marker:0, checkIn:1, checkOut:2, amount:3, pay:4, name:7, note:8 } }
    : { startRow: 4, startCol: 14, numCols: 11, // N~X
        idx: { marker:0, time:1, amount:2, pay:3, name:6, note:7 } };

  var lastRow = Math.max(tab.getLastRow(), cfg.startRow - 1);
  var numRows = Math.max(0, lastRow - cfg.startRow + 1);

  var data = (numRows > 0)
    ? tab.getRange(cfg.startRow, cfg.startCol, numRows, cfg.numCols).getValues()
    : [];

  function isRowUsed_(row) {
    return row.some(function(v) { return v !== '' && v !== null; });
  }

  // 섹션 내 "완전 빈 행"들
  var emptyIdx = [];
  for (var i = 0; i < data.length; i++) {
    if (!isRowUsed_(data[i])) emptyIdx.push(i);
  }

  var writes = [];
  var appended = [];
  var inserted = 0;

  rowDataList.forEach(function(rd) {
    var values = new Array(cfg.numCols).fill('');

    if (type === '대실') {
      values[cfg.idx.marker]  = 'V';
      values[cfg.idx.checkIn] = rd.checkIn || '';
      values[cfg.idx.checkOut]= rd.checkOut || '';
      values[cfg.idx.amount]  = rd.amount || '';
      values[cfg.idx.pay]     = rd.channel || '';
      values[cfg.idx.name]    = rd.name || '';
      values[cfg.idx.note]    = rd.roomType || ''; // 객실타입 → 특이사항/차종(J)
    } else {
      values[cfg.idx.marker]  = 'V';
      values[cfg.idx.time]    = rd.checkIn || '';  // 숙박은 checkIn에 '18시' 또는 ' - '
      values[cfg.idx.amount]  = rd.amount || '';
      values[cfg.idx.pay]     = rd.channel || '';
      values[cfg.idx.name]    = rd.name || '';
      values[cfg.idx.note]    = rd.roomType || ''; // 객실타입 → 특이사항/차종(U)
    }

    if (emptyIdx.length > 0) {
      writes.push({ rowOffset: emptyIdx.shift(), values: values });
    } else {
      appended.push(values);
    }
    inserted++;
  });

  // 섹션 내부 빈 행 채우기 (연속 구간 단위 setValues)
  if (writes.length > 0) {
    writes.sort(function(a,b){ return a.rowOffset - b.rowOffset; });

    var k = 0;
    while (k < writes.length) {
      var startOff = writes[k].rowOffset;
      var chunk = [writes[k].values];
      var j = k + 1;

      while (j < writes.length && writes[j].rowOffset === writes[j-1].rowOffset + 1) {
        chunk.push(writes[j].values);
        j++;
      }

      tab.getRange(cfg.startRow + startOff, cfg.startCol, chunk.length, cfg.numCols).setValues(chunk);
      k = j;
    }
  }

  // 아래로 append
  if (appended.length > 0) {
    var appendRow = cfg.startRow + data.length;
    tab.getRange(appendRow, cfg.startCol, appended.length, cfg.numCols).setValues(appended);
  }

  return { inserted: inserted };
}

/* =========================================================
 * ✅ 정렬: 자동(V) 행만 시간순 정렬, 수동 행 위치 고정
 * - startCol 첫 칸(대실=B / 숙박=N)이 'V'인 행만 정렬 대상
 * - 정렬 결과는 "V행들끼리"만 섞고, 그 외 행은 그대로 둠
========================================================= */
function sortSectionAutoOnly_(tab, startRow, startCol, numCols, timeOffset) {
  var lastRow = tab.getLastRow();
  if (lastRow < startRow) return;

  var numRows = lastRow - startRow + 1;
  var rng = tab.getRange(startRow, startCol, numRows, numCols);
  var data = rng.getValues();

  // V행 인덱스와 행데이터만 분리
  var vIndexes = [];
  var vRows = [];

  for (var i = 0; i < data.length; i++) {
    var marker = String(data[i][0] || '').trim(); // 섹션 첫 칸(B/N)
    if (marker === 'V') {
      vIndexes.push(i);
      vRows.push(data[i]);
    }
  }

  if (vRows.length <= 1) return;

  // V행들만 시간 기준 정렬
  vRows.sort(function(a, b) {
    return parseHour_(String(a[timeOffset])) - parseHour_(String(b[timeOffset]));
  });

  // 정렬된 V행을 원래 V자리(인덱스)에만 다시 꽂기
  for (var k = 0; k < vIndexes.length; k++) {
    data[vIndexes[k]] = vRows[k];
  }

  // 전체를 다시 쓰되, 수동행은 data에서 그대로 유지되므로 위치 고정됨
  rng.setValues(data);
}

/* =========================
 * 날짜 → 탭
 * ========================= */
function getDateRange_(checkIn, checkOut) {
  var dates = [];
  var cur = new Date(checkIn);
  cur.setHours(0, 0, 0, 0);
  var end = new Date(checkOut);
  end.setHours(0, 0, 0, 0);

  while (cur < end) {
    dates.push(new Date(cur));
    cur.setDate(cur.getDate() + 1);
  }
  return dates.length ? dates : [new Date(checkIn)];
}

function getTabName_(ss, dateObj) {
  var day = dateObj.getDate();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (/^\d+\(/.test(name) && parseInt(name, 10) === day) return name;
  }
  return null;
}

function formatHour_(dateObj) {
  if (!dateObj || !(dateObj instanceof Date) || isNaN(dateObj.getTime())) return '';
  return Utilities.formatDate(dateObj, 'Asia/Seoul', 'H') + '시';
}

/* =========================
 * 채널 추출 (의도 유지)
 * ========================= */
function extractChannel_(raw) {
  if (!raw) return '';
  raw = String(raw).trim().replace(/"/g, '');
  var m = raw.match(/[(（]([^）)]+)[)）]/);
  if (m) {
    var inner = m[1].trim();
    return CHANNEL_MAP[inner] || inner;
  }
  return CHANNEL_MAP[raw] || raw;
}

/* =========================
 * 시간 파싱 (정렬용)
 * ========================= */
function parseHour_(s) {
  var m = String(s).match(/^(\d+)/);
  if (!m) return 999;
  var h = parseInt(m[1], 10);
  return (h >= 1 && h <= 23) ? h : 999; // 0시/비정상 → 맨 뒤
}