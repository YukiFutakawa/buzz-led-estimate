// ===== LED見積フィードバック API =====
// このコードを Google Apps Script に貼り付けてデプロイしてください。
// スプレッドシート ID: 1nCTobFRPKY794NwvUzxkNIjzj540IKBHej6TWfWuUlI

const SPREADSHEET_ID = "1nCTobFRPKY794NwvUzxkNIjzj540IKBHej6TWfWuUlI";

const FEEDBACK_HEADERS = [
  "id", "timestamp", "property_name",
  "total_diffs", "fixture_match_rate", "led_match_rate",
  "comment_reading", "comment_selection",
  "fixture_diffs_json", "selection_diffs_json", "header_diffs_json",
  "submitter", "synced"
];

const LOG_HEADERS = ["timestamp", "action", "details"];

// ---- シート取得（なければ自動作成） ----

function getOrCreateSheet(title, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(title);
  if (!sheet) {
    sheet = ss.insertSheet(title);
    sheet.appendRow(headers);
  }
  return sheet;
}

// ---- GET: フィードバック取得 ----

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || "get_unsynced";

    if (action === "get_unsynced") {
      return jsonResponse(getUnsyncedFeedback());
    } else if (action === "get_stats") {
      return jsonResponse(getStats());
    } else if (action === "get_all") {
      return jsonResponse(getAllFeedback());
    } else {
      return jsonResponse({ error: "Unknown action: " + action });
    }
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ---- POST: フィードバック送信・同期マーク ----

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || "submit";

    if (action === "submit") {
      const result = submitFeedback(body);
      return jsonResponse(result);
    } else if (action === "mark_synced") {
      markSynced(body.feedback_ids || []);
      return jsonResponse({ success: true });
    } else {
      return jsonResponse({ error: "Unknown action: " + action });
    }
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ---- フィードバック送信 ----

function submitFeedback(data) {
  const sheet = getOrCreateSheet("feedback_raw", FEEDBACK_HEADERS);
  const id = Utilities.getUuid().substring(0, 8);
  const now = new Date().toISOString();
  const summary = data.summary || {};

  const row = [
    id,
    now,
    data.property_name || "",
    summary.total_diffs || 0,
    Math.round((summary.fixture_match_rate || 0) * 1000) / 1000,
    Math.round((summary.led_selection_match_rate || 0) * 1000) / 1000,
    data.comment_reading || "",
    data.comment_selection || "",
    JSON.stringify(data.fixture_diffs || []),
    JSON.stringify(data.selection_diffs || []),
    JSON.stringify(data.header_diffs || []),
    data.submitter || "",
    ""  // synced: 空 = 未同期
  ];

  sheet.appendRow(row);
  writeLog("feedback_submitted", (data.property_name || "?") + ": diffs=" + (summary.total_diffs || 0));

  return { success: true, feedback_id: id };
}

// ---- 未同期フィードバック取得 ----

function getUnsyncedFeedback() {
  const sheet = getOrCreateSheet("feedback_raw", FEEDBACK_HEADERS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { records: [] };

  const headers = data[0];
  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const syncedIdx = headers.indexOf("synced");
    if (!row[syncedIdx]) {
      const record = {};
      for (let j = 0; j < headers.length; j++) {
        record[headers[j]] = row[j];
      }
      records.push(record);
    }
  }
  return { records: records };
}

// ---- 全フィードバック取得 ----

function getAllFeedback() {
  const sheet = getOrCreateSheet("feedback_raw", FEEDBACK_HEADERS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { records: [] };

  const headers = data[0];
  const records = [];
  for (let i = 1; i < data.length; i++) {
    const record = {};
    for (let j = 0; j < headers.length; j++) {
      record[headers[j]] = data[i][j];
    }
    records.push(record);
  }
  return { records: records };
}

// ---- 同期済みマーク ----

function markSynced(feedbackIds) {
  if (!feedbackIds || feedbackIds.length === 0) return;

  const sheet = getOrCreateSheet("feedback_raw", FEEDBACK_HEADERS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const headers = data[0];
  const idCol = headers.indexOf("id");
  const syncedCol = headers.indexOf("synced");
  const idsSet = new Set(feedbackIds.map(String));
  const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm");

  for (let i = 1; i < data.length; i++) {
    if (idsSet.has(String(data[i][idCol]))) {
      sheet.getRange(i + 1, syncedCol + 1).setValue(now);
    }
  }
}

// ---- 統計 ----

function getStats() {
  const sheet = getOrCreateSheet("feedback_raw", FEEDBACK_HEADERS);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { total_feedback: 0, avg_led_match_rate: 0 };

  const headers = data[0];
  const rateCol = headers.indexOf("led_match_rate");
  let total = 0;
  let sum = 0;

  for (let i = 1; i < data.length; i++) {
    total++;
    sum += Number(data[i][rateCol]) || 0;
  }

  return {
    total_feedback: total,
    avg_led_match_rate: total > 0 ? Math.round((sum / total) * 1000) / 1000 : 0
  };
}

// ---- ログ ----

function writeLog(action, details) {
  try {
    const sheet = getOrCreateSheet("system_log", LOG_HEADERS);
    sheet.appendRow([new Date().toISOString(), action, details]);
  } catch (e) {
    // ログエラーは無視
  }
}

// ---- ユーティリティ ----

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
