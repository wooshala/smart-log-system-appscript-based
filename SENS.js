// =====================================================
// Sens.gs — 네이버 SENS SMS 발송
// =====================================================

function getSensConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    ACCESS_KEY: props.getProperty('SENS_ACCESS_KEY'),
    SECRET_KEY: props.getProperty('SENS_SECRET_KEY'),
    SERVICE_ID: props.getProperty('SENS_SERVICE_ID'),
    SENDER: props.getProperty('SENS_SENDER')
  };
}

function sendSms_(to, content) {

  const cfg = getSensConfig_();

  const phone = String(to).replace(/[^0-9]/g, '');

  const url =
    'https://sens.apigw.ntruss.com/sms/v2/services/' +
    cfg.SERVICE_ID +
    '/messages';

    const timestamp = String(Date.now());
    const signature = makeSensSignature_(timestamp);

    const body = JSON.stringify({
      type        : 'LMS',
      contentType : 'COMM',
      countryCode : '82',
      from        : SENS_CONFIG.SENDER,
      content     : content,
      messages    : [{ to: phone }],
    });

    const response = UrlFetchApp.fetch(url, {
      method            : 'post',
      contentType       : 'application/json; charset=utf-8',
      headers           : {
        'x-ncp-apigw-timestamp'  : timestamp,
        'x-ncp-iam-access-key'   : SENS_CONFIG.ACCESS_KEY,
        'x-ncp-apigw-signature-v2': signature,
      },
      payload           : body,
      muteHttpExceptions: true,
    });

    const code   = response.getResponseCode();
    const result = JSON.parse(response.getContentText());

    Logger.log('[sendSms_] HTTP %s → %s', code, JSON.stringify(result));

    if (code === 202) {
      return { ok: true, requestId: result.requestId };
    } else {
      return { ok: false, error: result.errorMessage || ('HTTP ' + code) };
    }

  } catch (e) {
    Logger.log('[sendSms_] 오류: ' + e.message);
    return { ok: false, error: e.message };
  }
}

// ── SENS HMAC-SHA256 서명 생성 ────────────────────────
function makeSensSignature_(timestamp) {
  const method  = 'POST';
  const url     = '/sms/v2/services/' + SENS_CONFIG.SERVICE_ID + '/messages';

  const message = method + ' ' + url + '\n' + timestamp + '\n' + SENS_CONFIG.ACCESS_KEY;

  // UTF-8 명시적 지정
  const signature = Utilities.computeHmacSha256Signature(
    message,
    SENS_CONFIG.SECRET_KEY,
    Utilities.Charset.UTF_8   // ← 이 줄 추가가 핵심!
  );
  return Utilities.base64Encode(signature);
}


// 사이드바에서 직접 전화번호+내용으로 SMS 발송
function sendSmsDirectly(phone, quoteId) {
  try {
    const dbSh  = ensureDbSheet_();
    const data  = dbSh.getDataRange().getValues();
    const header = data[0];
    const row   = data.slice(1).find(r => String(r[0]) === String(quoteId));
    if (!row) return { ok: false, error: '견적 ID 없음. 먼저 저장하세요.' };

    const rec = {};
    header.forEach((h, i) => { rec[h] = row[i]; });
    rec['전화번호'] = phone; // 사이드바 입력값 우선 사용

    const msg = buildLmsMessage_(rec);
    return sendSms_(phone, msg);

  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ── 테스트용 (편집기에서 직접 실행 가능) ─────────────
function testSens() {
  const result = sendSms_('01036636680', '[라벨호텔] SENS 테스트 메시지입니다.');
  Logger.log('테스트 결과: ' + JSON.stringify(result));
}
function checkSensConfig() {
  Logger.log('ACCESS_KEY 길이: ' + SENS_CONFIG.ACCESS_KEY.length);
  Logger.log('SECRET_KEY 길이: ' + SENS_CONFIG.SECRET_KEY.length);
  Logger.log('공백 포함(ACCESS): ' + SENS_CONFIG.ACCESS_KEY.includes(' '));
  Logger.log('공백 포함(SECRET): ' + SENS_CONFIG.SECRET_KEY.includes(' '));
  Logger.log('ACCESS_KEY 끝 5자: [' + SENS_CONFIG.ACCESS_KEY.slice(-5) + ']');
  Logger.log('SECRET_KEY 끝 5자: [' + SENS_CONFIG.SECRET_KEY.slice(-5) + ']');
}

