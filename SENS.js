// =====================================================
// SENS.gs — 네이버 SENS SMS 발송
// =====================================================

function getSensConfig_() {
  var props = PropertiesService.getScriptProperties();
  return {
    ACCESS_KEY: props.getProperty('SENS_ACCESS_KEY') || '',
    SECRET_KEY: props.getProperty('SENS_SECRET_KEY') || '',
    SERVICE_ID: props.getProperty('SENS_SERVICE_ID') || '',
    SENDER: props.getProperty('SENS_SENDER') || ''
  };
}

// ── 외부에서 호출하는 메인 함수 ──────────────────────
function sendSms_(to, content) {
  try {
    var cfg = getSensConfig_();
    var phone = String(to || '').replace(/[^0-9]/g, '');

    if (!phone) {
      return { ok: false, error: '전화번호 없음' };
    }

    if (!cfg.ACCESS_KEY || !cfg.SECRET_KEY || !cfg.SERVICE_ID || !cfg.SENDER) {
      return { ok: false, error: 'SENS Script Properties 설정값 누락' };
    }

    var url = 'https://sens.apigw.ntruss.com/sms/v2/services/' +
              cfg.SERVICE_ID + '/messages';

    var timestamp = String(Date.now());
    var signature = makeSensSignature_(timestamp);

    var body = JSON.stringify({
      type: 'LMS',
      contentType: 'COMM',
      countryCode: '82',
      from: cfg.SENDER,
      content: String(content || ''),
      messages: [{ to: phone }]
    });

    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      headers: {
        'x-ncp-apigw-timestamp': timestamp,
        'x-ncp-iam-access-key': cfg.ACCESS_KEY,
        'x-ncp-apigw-signature-v2': signature
      },
      payload: body,
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    var text = response.getContentText();
    var result;

    if (text) {
      result = JSON.parse(text);
    } else {
      result = {};
    }

    Logger.log('[sendSms_] HTTP ' + code + ' -> ' + JSON.stringify(result));

    if (code === 202) {
      return { ok: true, requestId: result.requestId };
    }

    return {
      ok: false,
      error: result.errorMessage || ('HTTP ' + code)
    };

  } catch (e) {
    Logger.log('[sendSms_] error: ' + e.message);
    return { ok: false, error: e.message };
  }
}

// ── SENS HMAC-SHA256 서명 생성 ────────────────────────
function makeSensSignature_(timestamp) {
  var cfg = getSensConfig_();
  var method = 'POST';
  var url = '/sms/v2/services/' + cfg.SERVICE_ID + '/messages';
  var message = method + ' ' + url + '\n' + timestamp + '\n' + cfg.ACCESS_KEY;

  var signature = Utilities.computeHmacSha256Signature(
    message,
    cfg.SECRET_KEY,
    Utilities.Charset.UTF_8
  );

  return Utilities.base64Encode(signature);
}

// 사이드바에서 직접 전화번호+내용으로 SMS 발송
function sendSmsDirectly(phone, quoteId) {
  try {
    var dbSh = ensureDbSheet_();
    var data = dbSh.getDataRange().getValues();
    var header = data[0];
    var row = null;
    var i;
    var rec = {};
    var msg;

    for (i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(quoteId)) {
        row = data[i];
        break;
      }
    }

    if (!row) {
      return { ok: false, error: '견적 ID 없음. 먼저 저장하세요.' };
    }

    for (i = 0; i < header.length; i++) {
      rec[header[i]] = row[i];
    }

    rec['전화번호'] = phone;
    msg = buildLmsMessage_(rec);

    return sendSms_(phone, msg);

  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ── 테스트용 ─────────────────────────────────────────
function testSens() {
  var result = sendSms_('01036636680', '[라벨호텔] SENS 테스트 메시지입니다.');
  Logger.log('테스트 결과: ' + JSON.stringify(result));
}

function checkSensConfig() {
  var cfg = getSensConfig_();

  Logger.log('ACCESS_KEY length: ' + cfg.ACCESS_KEY.length);
  Logger.log('SECRET_KEY length: ' + cfg.SECRET_KEY.length);
  Logger.log('ACCESS contains space: ' + (cfg.ACCESS_KEY.indexOf(' ') > -1));
  Logger.log('SECRET contains space: ' + (cfg.SECRET_KEY.indexOf(' ') > -1));
  Logger.log('ACCESS_KEY last 5: [' + cfg.ACCESS_KEY.slice(-5) + ']');
  Logger.log('SECRET_KEY last 5: [' + cfg.SECRET_KEY.slice(-5) + ']');
}