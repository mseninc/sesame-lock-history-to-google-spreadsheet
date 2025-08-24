/*
 * SESAME ロック履歴を毎日取得してスプレッドシートへ追記する GAS
 * - Sesame API: https://doc.candyhouse.co/ja/SesameAPI
 *
 * 事前に Script Properties に以下を設定してください:
 *  - SESAME_識別名_API_KEY
 *  - SESAME_識別名_DEVICE_ID
 */

/* =========================
 * 定数
 * ========================= */

/** SESAME API base URL */
const API_BASE = 'https://app.candyhouse.co/api';

/** 1ページの取得件数（最新から50件ずつ） */
const PAGE_LIMIT = 50;

/** type 対応表（仕様に基づく） */
const TYPE_MAP = {
  0:  { typeKey: 'none',                   description: 'なし' },
  1:  { typeKey: 'bleLock',                description: '施錠 (BLE)' },
  2:  { typeKey: 'bleUnLock',              description: '解錠 (BLE)' },
  3:  { typeKey: 'timeChanged',            description: '内部時計校正' },
  4:  { typeKey: 'autoLockUpdated',        description: 'オートロック設定変更' },
  5:  { typeKey: 'mechSettingUpdated',     description: '施解錠角度設定変更' },
  6:  { typeKey: 'autoLock',               description: 'オートロック' },
  7:  { typeKey: 'manualLocked',           description: '施錠 (手動)' },
  8:  { typeKey: 'manualUnlocked',         description: '解錠 (手動)' },
  9:  { typeKey: 'manualElse',             description: '手動操作' },
  10: { typeKey: 'driveLocked',            description: 'モーターが確実に施錠' },
  11: { typeKey: 'driveUnlocked',          description: 'モーターが確実に解錠' },
  12: { typeKey: 'driveFailed',            description: 'モーターが施解錠の途中に失敗' },
  13: { typeKey: 'bleAdvParameterUpdated', description: 'BLEアドバタイズ設定変更' },
  14: { typeKey: 'wm2Lock',                description: '施錠 (WiFi Module2)' },
  15: { typeKey: 'wm2Unlock',              description: '解錠 (WiFi Module2)' },
  16: { typeKey: 'webLock',                description: '施錠 (Web API)' },
  17: { typeKey: 'webUnlock',              description: '解錠 (Web API)' },
  90: { typeKey: 'sensorOpen',             description: 'オープン' },
  91: { typeKey: 'sensorClose',            description: 'クローズ' },
};

/** 対象のロック一覧（環境変数名、シート名の識別子） */
const lockNames = ['4F', '5F'];

/* =========================
 * エントリーポイント
 * ========================= */

function runDaily() {
  const sp = SpreadsheetApp.getActiveSpreadsheet();

  const props = PropertiesService.getScriptProperties();


  lockNames.forEach((lockName) => {
    const config = {
      apiKey: props.getProperty(`SESAME_${lockName}_API_KEY`),
      deviceId: props.getProperty(`SESAME_${lockName}_DEVICE_ID`),
    }
    validateConfig(lockName, config);
    const sheet = ensureSheet(sp, lockName);
    const existingKeySet = buildExistingKeySet(sheet);
    const newRows = collectNewHistoryRows(config, existingKeySet);
    appendRows(sheet, newRows.reverse());
  });
}

function createDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runDaily' && t.getEventType() === ScriptApp.EventType.CLOCK);

  if (triggers.length > 0) return;

  ScriptApp.newTrigger('runDaily')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
}

/* =========================
 * コアロジック
 * ========================= */

function collectNewHistoryRows(cfg, existingKeySet) {
  const rows = [];
  let pageIndex = 0;
  let hitExisting = false;

  while (!hitExisting) {
    const page = fetchHistoryPage(cfg.apiKey, cfg.deviceId, PAGE_LIMIT, pageIndex);
    if (!Array.isArray(page) || page.length === 0) break;

    for (const item of page) {
      const decodedTag = decodeHistoryTag(item.historyTag);
      const key = item.recordID;
      if (existingKeySet.has(key)) {
        hitExisting = true;
        break;
      }
      rows.push(historyToRow(item, decodedTag));
    }

    pageIndex += 1;
    if (hitExisting || page.length < PAGE_LIMIT) break;
  }

  return rows;
}

function fetchHistoryPage(apiKey, deviceId, limit, page) {
  const url = `${API_BASE}/sesame2/${encodeURIComponent(deviceId)}/history?page=${page}&lg=${limit}`;
  const opts = {
    method: 'get',
    headers: { 'x-api-key': apiKey },
    muteHttpExceptions: true,
  };

  let lastError = null;
  for (let i = 0; i < 3; i++) {
    const attempt = i + 1;
    try {
      const res = UrlFetchApp.fetch(url, opts);
      const code = res.getResponseCode();
      if (code >= 200 && code < 300) {
        const text = res.getContentText();
        try {
          const json = JSON.parse(text || '[]');
          console.log({ message: 'SesameAPI response', deviceId, attempt, page, data: json });
          return Array.isArray(json) ? json : [];
        } catch (parseErr) {
          logFetchException(url, attempt, parseErr, { snippet: text && text.slice(0, 500) });
          lastError = parseErr;
        }
      } else {
        logHttpError(url, attempt, code, res);
        lastError = new Error(`HTTP ${code}`);
      }
    } catch (err) {
      logFetchException(url, attempt, err);
      lastError = err;
    }
    Utilities.sleep(300 * attempt);
  }
  throw new Error(`Sesame API 呼び出し失敗: ${url} | lastError=${lastError && lastError.message ? lastError.message : 'Unknown'}`);
}

function historyToRow(item, decodedTag) {
  const tsNum = Number(item.timeStamp);
  const tsMs = tsNum < 1e12 ? tsNum * 1000 : tsNum;
  const date = new Date(tsMs);
  const mapped = TYPE_MAP[item.type] || { description: String(item.type) };
  const rawStr = JSON.stringify(item);
  return [date, mapped.description, decodedTag, rawStr];
}

/* =========================
 * ヘルパー
 * ========================= */

function ensureSheet(sp, name) {
  const sheet = sp.getSheetByName(name) || sp.insertSheet(name);
  const hasHeader = sheet.getLastRow() > 0;
  if (!hasHeader) {
    sheet.appendRow(['timeStamp', 'type', 'historyTag', 'raw']);
    sheet.getRange(2, 1, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
  return sheet;
}

function buildExistingKeySet(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return new Set();

  const values = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
  return values.reduce((set, [rawJson]) => {
    try {
      const raw = JSON.parse(rawJson);
      if (raw && raw.recordID) set.add(raw.recordID);
    } catch (e) {
      // skip
    }
    return set;
  }, new Set());
}

function appendRows(sheet, rows) {
  if (!rows || rows.length === 0) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

/* =========================
 * ユーティリティ
 * ========================= */

function decodeHistoryTag(b64) {
  if (!b64) return '';
  const bytes = Utilities.base64Decode(b64);
  return Utilities.newBlob(bytes).getDataAsString('UTF-8');
}

function logHttpError(url, attempt, status, res) {
  const headers = (typeof res.getAllHeaders === 'function') ? res.getAllHeaders() : {};
  const reqId = headers['x-request-id'] || headers['X-Request-Id'] || headers['x-amzn-RequestId'] || headers['x-amz-request-id'] || '';
  let bodySnippet = '';
  try {
    const text = res.getContentText();
    bodySnippet = text ? text.substring(0, 500) : '';
  } catch (e) {
    try {
      bodySnippet = `<binary ${res.getContent().length} bytes>`;
    } catch (_) {
      bodySnippet = '<no-body>';
    }
  }
  console.error('[SesameAPI][HTTP_ERROR]', { attempt, status, url, requestId: String(reqId), headers, body: bodySnippet });
}

function logFetchException(url, attempt, err, extra) {
  const name = err && err.name ? err.name : 'Error';
  const message = err && err.message ? err.message : String(err);
  const stack = err && err.stack ? err.stack : '';
  const snippet = extra && extra.snippet ? extra.snippet : '';
  console.error('[SesameAPI][EXCEPTION]', { attempt, url, name, message, stack, snippet });
}

function validateConfig(lockName, cfg) {
  const missing = ['apiKey', 'deviceId'].filter(k => !cfg[k]);
  if (missing.length) {
    throw new Error(`${lockName} の設定不足: ${missing.join(', ')} を Script Properties に設定してください`);
  }
}
