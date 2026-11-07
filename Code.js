// =========================================================================
// Code.gs â€” V5: Entry Point + Setup + Auth
// FIX Bug 5: Password case sensitivity benar-benar pakai config
// BARU: Key "User ID Prefix" ditambahkan di AppConfig
// =========================================================================

function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var requestedPage = String(params.page || params.view || "").trim().toLowerCase();
    var v2SessionId = String(params.handoverV2 || params.sessionId || params.session || params.sid || "").trim();
    if (requestedPage === "handoverv2" || requestedPage === "handover-v2" || params.handoverV2) {
      if (!v2SessionId) {
        throw new Error("Session ID diperlukan. Tambahkan ?page=handoverV2&sessionId=SESSION_ID pada URL.");
      }
      return HtmlService.createHtmlOutput(renderHandoverHtmlV2_(v2SessionId, "handover-v2-url-test"))
        .setTitle("Handover V2 - " + v2SessionId)
        .addMetaTag("viewport","width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    var title = "InLens";
    try { title = cfg("App Title") || title; } catch(e) {}
    return HtmlService.createTemplateFromFile("Index")
      .evaluate()
      .setTitle(title)
      .addMetaTag("viewport","width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    // Tampilkan error page yang jelas (bukan blank page) untuk user eksternal
    return HtmlService.createHtmlOutput(
      '<div style="padding:40px;text-align:center;font-family:Poppins,sans-serif">'
      + '<h2 style="color:#dc2626">âš ï¸ Terjadi Kesalahan</h2>'
      + '<p style="color:#475569">' + err.toString() + '</p>'
      + '<p style="color:#94a3b8;font-size:13px">Hubungi admin (pemilik app) untuk bantuan.</p>'
      + '<button onclick="location.reload()" style="margin-top:16px;padding:10px 24px;'
      + 'background:#1d4ed8;color:white;border:none;border-radius:8px;font-size:14px;'
      + 'cursor:pointer">Coba Lagi</button></div>'
    ).setTitle('InLens - Error');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Server-side render Handover Report HTML resmi menggunakan template HandoverV2.html.
 */
function renderHandoverReport(sessionId, traceId) {
  return renderHandoverHtml_(sessionId, traceId);
}

function renderHandoverHtml_(sessionId, traceId) {
  perfTrace_("printHandover", traceId, "server_start", { sessionId: sessionId });
  var result = getHandOverData(sessionId, traceId);
  if (!result || !result.success || !result.data) {
    perfTrace_("printHandover", traceId, "server_error", { message: result ? result.message : "Data tidak ditemukan" });
    throw new Error(result ? result.message : "Data tidak ditemukan");
  }
  perfTrace_("printHandover", traceId, "template_render_start");
  var template = HtmlService.createTemplateFromFile("HandoverV2");
  template.data = result.data;
  var html = template.evaluate().getContent();
  perfTrace_("printHandover", traceId, "server_end", { success: true, htmlLength: html.length });
  return html;
}

/**
 * Server-side render Handover Report V2 untuk test manual.
 * Saat ini sama dengan template resmi.
 */
function renderHandoverReportV2(sessionId, traceId) {
  return renderHandoverHtmlV2_(sessionId, traceId);
}

function renderHandoverHtmlV2_(sessionId, traceId) {
  perfTrace_("printHandoverV2", traceId, "server_start", { sessionId: sessionId });
  var result = getHandOverData(sessionId, traceId);
  if (!result || !result.success || !result.data) {
    perfTrace_("printHandoverV2", traceId, "server_error", { message: result ? result.message : "Data tidak ditemukan" });
    throw new Error(result ? result.message : "Data tidak ditemukan");
  }
  perfTrace_("printHandoverV2", traceId, "template_render_start");
  var template = HtmlService.createTemplateFromFile("HandoverV2");
  template.data = result.data;
  var html = template.evaluate().getContent();
  perfTrace_("printHandoverV2", traceId, "server_end", { success: true, htmlLength: html.length });
  return html;
}

// =========================================================================
// setupDatabase()
// =========================================================================
function setupDatabase() {
  var ss = getSpreadsheet_();
  createConfigSheet_(ss);
  var DB = getSheetNames_();

  var tables = [
    { name: DB.SESSIONS, headers: ["Session ID","Date","Brand","Shipment","PIC","Status","Timestamp","Folder URL","Created By","Forwarder","Loss Prevention Name","Loss Prevention Filled At","Loss Prevention Filled By","Loss Prevention Signature Mode","Loss Prevention Signature File ID","Loss Prevention Signature File URL","Loss Prevention Signature Signed At","Handover File ID","Handover File URL","Handover Current Revision ID","Handover Version No","Handover Generated At","Rekap Status","Rekap Error","Handover PDF Status","Handover PDF Error","WA Status","WA Error","Artifact Updated At"] },
    { name: DB.FLEET,    headers: ["Container ID","Session ID","Number","Type","Driver","Plate","Timestamp","Folder URL","Status","Created By","Qty Koli","Work Status","Started At","Started By","Completed At","Completed By","Reopened At","Reopened By"] },
    { name: DB.DATA,     headers: ["Brand List","PIC List","Vehicle Type"] },
    { name: DB.AUDIT,    headers: ["Timestamp","Action","Status","Message","User/Context"] },
    { name: DB.PHOTOS,   headers: ["Photo ID","Container ID","Type","Drive URL","Timestamp","Uploaded By","Nama Palet","Koli Palet","Nama File","Local ID","Job ID","Log Status"] },
    { name: DB.USERS,    headers: ["User ID","Session ID","User Name","Status","Last Active","Join Time"] },
    { name: DB.ACCOUNTS, headers: ["Username","Password","NamaLengkap","Active","Signature File ID"] },
    { name: DB.REKAP,    headers: cfgList("Rekap Column Order") },
    { name: DB.SJ_QTY,   headers: SJ_QTY_HEADERS },
    { name: DB.HANDOVER_REVISION_LOG, headers: ["Log ID","Session ID","File ID","Revision ID","Version No","Generated At","Generated By","Keep Forever","Notes"] }
  ];

  if (!tables[7].headers || tables[7].headers.length === 0) {
    tables[7].headers = ["Tanggal","Brand","Shipment","Sesi ID","Armada Ke","Tipe","Nama Supir","Plat Nomor","Jml Dokumen","Jml Palet","Jml Surat Jalan","Qty Koli","Link Folder Armada"];
  }

  tables.forEach(function(t) {
    if (!t.name) return;
    if (!ss.getSheetByName(t.name)) {
      var s = ss.insertSheet(t.name);
      s.appendRow(t.headers);
      s.getRange(1,1,1,t.headers.length).setFontWeight("bold").setBackground("#f1f5f9");
      s.setFrozenRows(1);
    }
  });

  var sessionSheet = ss.getSheetByName(DB.SESSIONS);
  if (sessionSheet) ensureSessionColumns_(sessionSheet);

  // Pastikan kolom lifecycle armada tersedia di Fleet sheet lama.
  var fleetSheet = ss.getSheetByName(DB.FLEET);
  if (fleetSheet) ensureFleetWorkColumns_(fleetSheet);

  var photoSheet = ss.getSheetByName(DB.PHOTOS);
  if (photoSheet) ensurePhotoLogColumns_(photoSheet);

  var accountSheet = ss.getSheetByName(DB.ACCOUNTS);
  if (accountSheet) ensureAccountColumns_(accountSheet);

  if (typeof ensureSjQtySheet_ === "function") ensureSjQtySheet_();
  if (typeof ensureHandoverRevisionLogSheet_ === "function") ensureHandoverRevisionLogSheet_();

  return "Setup selesai!";
}

function ensureSessionColumns_(sheet) {
  if (!sheet) return;
  var required = [
    "Forwarder",
    "Loss Prevention Name",
    "Loss Prevention Filled At",
    "Loss Prevention Filled By",
    "Loss Prevention Signature Mode",
    "Loss Prevention Signature File ID",
    "Loss Prevention Signature File URL",
    "Loss Prevention Signature Signed At",
    "Handover File ID",
    "Handover File URL",
    "Handover Current Revision ID",
    "Handover Version No",
    "Handover Generated At",
    "Rekap Status",
    "Rekap Error",
    "Handover PDF Status",
    "Handover PDF Error",
    "WA Status",
    "WA Error",
    "Artifact Updated At"
  ];
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var missing = [];
  for (var i = 0; i < required.length; i++) {
    if (headers.indexOf(required[i]) < 0) missing.push(required[i]);
  }
  if (missing.length > 0) {
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    sheet.getRange(1, lastCol + 1, 1, missing.length).setFontWeight("bold").setBackground("#f1f5f9");
  }
}

function ensurePhotoLogColumns_(sheet) {
  var required = ["Nama Palet","Koli Palet","Nama File","Local ID","Job ID","Log Status"];
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var missing = [];
  for (var i = 0; i < required.length; i++) {
    if (headers.indexOf(required[i]) < 0) missing.push(required[i]);
  }
  if (missing.length > 0) {
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    sheet.getRange(1, lastCol + 1, 1, missing.length).setFontWeight("bold").setBackground("#f1f5f9");
  }
}

function getSheetHeaderMap_(sheet) {
  var headers = sheet && sheet.getLastColumn() > 0
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    : [];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var key = String(headers[i] || "").trim();
    if (key) map[key] = i;
  }
  return map;
}

function ensureFleetWorkColumns_(sheet) {
  if (!sheet) return {};
  var required = [
    "Qty Koli",
    "Work Status",
    "Started At",
    "Started By",
    "Completed At",
    "Completed By",
    "Reopened At",
    "Reopened By",
    "Accumulated Work Ms"
  ];
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var missing = [];
  for (var i = 0; i < required.length; i++) {
    if (headers.indexOf(required[i]) < 0) missing.push(required[i]);
  }
  if (missing.length > 0) {
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    sheet.getRange(1, lastCol + 1, 1, missing.length).setFontWeight("bold").setBackground("#f1f5f9");
  }

  var map = getSheetHeaderMap_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return map;

  var qtyCol = map["Qty Koli"];
  if (qtyCol !== undefined) {
    var qtyValues = sheet.getRange(2, qtyCol + 1, lastRow - 1, 1).getValues();
    var qtyChanged = false;
    for (var q = 0; q < qtyValues.length; q++) {
      if (qtyValues[q][0] === "" || qtyValues[q][0] === null) {
        qtyValues[q][0] = 0;
        qtyChanged = true;
      }
    }
    if (qtyChanged) sheet.getRange(2, qtyCol + 1, lastRow - 1, 1).setValues(qtyValues);
  }

  var statusCol = map["Work Status"];
  if (statusCol !== undefined) {
    var rowValues = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    var statusValues = sheet.getRange(2, statusCol + 1, lastRow - 1, 1).getValues();
    var statusChanged = false;
    for (var s = 0; s < rowValues.length; s++) {
      var fleetStatus = String(rowValues[s][8] || "").trim().toUpperCase();
      var workStatus = String(statusValues[s][0] || "").trim();
      if (fleetStatus !== "DELETED" && !workStatus) {
        statusValues[s][0] = "READY";
        statusChanged = true;
      }
    }
    if (statusChanged) sheet.getRange(2, statusCol + 1, lastRow - 1, 1).setValues(statusValues);
  }

  return getSheetHeaderMap_(sheet);
}

function getFleetCell_(row, map, name, fallbackIndex) {
  if (map && map[name] !== undefined) return row[map[name]];
  return fallbackIndex !== undefined ? row[fallbackIndex] : "";
}

function setFleetCell_(row, map, name, value, fallbackIndex) {
  var idx = map && map[name] !== undefined ? map[name] : fallbackIndex;
  if (idx === undefined || idx < 0) return;
  while (row.length <= idx) row.push("");
  row[idx] = value;
}

function normalizeFleetWorkStatus_(value) {
    var s = String(value || "").trim().toUpperCase().replace(/\s+/g, "_");
    if (s === "DONE" || s === "FINISH" || s === "FINISHED" || s === "COMPLETED") return "DONE";
    if (s === "IN_PROGRESS" || s === "STARTED" || s === "WORKING" || s === "ON_PROGRESS") return "IN_PROGRESS";
    if (s === "UPLOAD_ONLY") return "UPLOAD_ONLY";
    return "READY";
  }

function getFleetWorkStatusText_(status) {
  status = normalizeFleetWorkStatus_(status);
  if (status === "DONE") return "Selesai";
  if (status === "IN_PROGRESS") return "Berjalan";
    if (status === "UPLOAD_ONLY") return "Break (Bisa Upload)";
    return "Belum mulai";
}

function formatFleetWorkTimestamp_(value) {
  var ms = toTimestampMs_(value);
  if (!ms) return "";
  return Utilities.formatDate(new Date(ms), cfg("Timezone") || "Asia/Jakarta", "dd MMM yyyy HH:mm");
}

function buildFleetWorkSnapshot_(row, map) {
  var status = normalizeFleetWorkStatus_(getFleetCell_(row, map, "Work Status"));
  var startedMs = toTimestampMs_(getFleetCell_(row, map, "Started At"));
  var completedMs = toTimestampMs_(getFleetCell_(row, map, "Completed At"));
  var reopenedMs = toTimestampMs_(getFleetCell_(row, map, "Reopened At"));
  var accMs = Number(getFleetCell_(row, map, "Accumulated Work Ms")) || 0;
  var durationMs = 0;
  if (status === "DONE" || status === "UPLOAD_ONLY") {
    if (accMs > 0) {
      durationMs = accMs;
    } else if (startedMs && completedMs && completedMs >= startedMs) {
      durationMs = completedMs - startedMs;
    }
  } else if (status === "IN_PROGRESS") {
    var activeSegmentStartMs = reopenedMs || startedMs;
    if (activeSegmentStartMs) durationMs = accMs + Math.max(0, Date.now() - activeSegmentStartMs);
  }
  return {
    workStatus: status,
    workStatusText: getFleetWorkStatusText_(status),
    startedAt: startedMs ? new Date(startedMs).toISOString() : "",
    startedAtMs: startedMs || 0,
    startedAtText: formatFleetWorkTimestamp_(getFleetCell_(row, map, "Started At")),
    startedBy: String(getFleetCell_(row, map, "Started By") || "").trim(),
    completedAt: completedMs ? new Date(completedMs).toISOString() : "",
    completedAtMs: completedMs || 0,
    completedAtText: formatFleetWorkTimestamp_(getFleetCell_(row, map, "Completed At")),
    completedBy: String(getFleetCell_(row, map, "Completed By") || "").trim(),
    reopenedAt: reopenedMs ? new Date(reopenedMs).toISOString() : "",
    reopenedAtMs: reopenedMs || 0,
    reopenedAtText: formatFleetWorkTimestamp_(getFleetCell_(row, map, "Reopened At")),
    reopenedBy: String(getFleetCell_(row, map, "Reopened By") || "").trim(),
    durationMs: durationMs,
    durationText: durationMs > 0
      ? formatDurationRange_(0, durationMs, "Kurang dari 1 menit")
      : ""
  };
}

function ensureAccountColumns_(sheet) {
  if (!sheet) return;
  var required = ["Signature File ID"];
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var missing = [];
  for (var i = 0; i < required.length; i++) {
    if (headers.indexOf(required[i]) < 0) missing.push(required[i]);
  }
  if (missing.length > 0) {
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    sheet.getRange(1, lastCol + 1, 1, missing.length).setFontWeight("bold").setBackground("#f1f5f9");
  }
}

// =========================================================================
// createConfigSheet_() â€” Sheet AppConfig dengan SEMUA config, tidak ada hardcode
// =========================================================================
function getAppBindingConfigRows_(ss) {
  return [
    ["# APP BINDING","","",""],
    ["Spreadsheet ID",ss.getId(),"ID Google Sheet database yang dipakai aplikasi. Edit value ini untuk pindah database, lalu clear cache / reload app.","Contoh: 1Zo_-... atau URL Google Sheet"],
    ["Apps Script ID",typeof getCurrentScriptId_ === "function" ? getCurrentScriptId_() : "","ID project GAS ini untuk audit duplicate. Informasi saja.",""],
    ["Spreadsheet ID Property Key","SPREADSHEET_ID","Nama cache/fallback internal untuk project standalone. Konfigurasi utama tetap baris Spreadsheet ID di AppConfig.","Opsional"]
  ];
}

function moveAppBindingConfigToTop_(sheet, ss) {
  if (!sheet || !ss || sheet.getLastRow() < 1) return;
  var desiredRows = getAppBindingConfigRows_(ss);
  var desiredKeys = {};
  for (var d = 0; d < desiredRows.length; d++) desiredKeys[String(desiredRows[d][0]).trim()] = d;

  var values = sheet.getDataRange().getValues();
  var foundByKey = {};
  var deleteRows = [];
  for (var i = 1; i < values.length; i++) {
    var key = String(values[i][0] || "").trim();
    if (desiredKeys[key] !== undefined) {
      foundByKey[key] = values[i];
      deleteRows.push(i + 1);
    }
  }

  for (var r = deleteRows.length - 1; r >= 0; r--) {
    sheet.deleteRow(deleteRows[r]);
  }

  sheet.insertRowsAfter(1, desiredRows.length);
  var outputRows = desiredRows.map(function(row) {
    var key = String(row[0]).trim();
    var existing = foundByKey[key] || [];
    var merged = row.slice();
    if (key === "Spreadsheet ID") merged[1] = existing[1] || row[1];
    if (key === "Apps Script ID") merged[1] = row[1] || existing[1] || "";
    if (key === "Spreadsheet ID Property Key") merged[1] = existing[1] || row[1];
    if (existing[2]) merged[2] = existing[2];
    if (existing[3]) merged[3] = existing[3];
    return merged;
  });
  sheet.getRange(2, 1, outputRows.length, 4).setValues(outputRows);
  sheet.getRange(2, 1, 1, 4).setFontWeight("bold").setBackground("#D7E9FF");
  sheet.getRange(3, 2, 3, 1).setBackground("#FFFDE7");
}

function createConfigSheet_(ss) {
  var s = ss.getSheetByName(CONFIG_SHEET_NAME);
  var isNew = false;
  if (!s) {
    s = ss.insertSheet(CONFIG_SHEET_NAME);
    var headers = ["Setting (KEY â€” JANGAN DIUBAH)","Value  (âœï¸ edit kolom ini)","Keterangan","Contoh / Opsi"];
    s.appendRow(headers);
    s.getRange(1,1,1,4).setFontWeight("bold").setBackground("#1E3A5F").setFontColor("#FFFFFF");
    s.setColumnWidth(1,250); s.setColumnWidth(2,350); s.setColumnWidth(3,300); s.setColumnWidth(4,300);
    s.setFrozenRows(1);
    isNew = true;
  }

  var existingKeys = {};
  if (!isNew) {
    var data = s.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      existingKeys[String(data[i][0]).trim()] = true;
    }
  }

  var rows = [
    // APP BINDING
    ["# APP BINDING","","",""],
    ["Spreadsheet ID",ss.getId(),"ID Google Sheet database yang dipakai aplikasi. Edit value ini untuk pindah database, lalu clear cache / reload app.","Contoh: 1Zo_-... atau URL Google Sheet"],
    ["Apps Script ID",typeof getCurrentScriptId_ === "function" ? getCurrentScriptId_() : "","ID project GAS ini untuk audit duplicate. Informasi saja.",""],
    ["Spreadsheet ID Property Key","SPREADSHEET_ID","Nama cache/fallback internal untuk project standalone. Konfigurasi utama tetap baris Spreadsheet ID di AppConfig.","Opsional"],

    // GOOGLE DRIVE
    ["  â–¸  GOOGLE DRIVE","","",""],
    ["Root Folder ID","0AIOQcXjlIoNcUk9PVA","ID folder root Drive / Shared Drive","Ambil dari URL Drive"],
    ["Use Shared Drive","Ya","Root Folder = Shared Drive?","Ya / Tidak"],

    // SECRETS & CREDENTIALS
    ["  â–¸  SECRETS & API KEYS","","",""],
    ["SA Email","name-mcp-access@mcp-gsheet-489510.iam.gserviceaccount.com","Email Google Service Account","email@namaproject.iam.gserviceaccount.com"],
    ["SA Private Key","-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQCvMGfNSpMvRKaw\nQr74GAeLOCToBdGu2mPVdRJGtK8qv0lF66SWKeQQYirX5P07Rn4FXlzoI88uFbVN\n0JNysLF2v07JZFY7CsmfxJxGoFPvxnZY5SNklNETJgUJHn6CIlW5CikbEQEwbTka\nmItj7lfUECzGIuPQ7jm7zq+jKc2gz7tYt08g3E/esycG9wHtDgvQo7VLDhT+kcm4\nfg9QpvPuRbsn4a9pNlH0J+oS1ptmRM1HPUq+Q/d3ixiUEJbQVPERP/RTAWkSIqBR\nHVKWORu+CE0jUDWoPNEqDjH1V9dalpJY56GbbLeYT7Susr972DJP9dMpTt+y2SaW\nDOmGuryZAgMBAAECggEAAYl8L6nq1qxUhhFf2Y9kaY1JjZuXdmBGJfZL/cDSFJ7v\nBqUFoaemb5pc494HdsCslGF6kTLXO4FHQPKGiKgSNoeBKnUt6IC0uikgceIwOldU\nbH0+LUC9AN0BiE3zMDMa62+yYqFokEM1EA2sk2jg7mRh8C8YLpoYeZRbqCPSwQlt\n1Osrke+sWtXhE04J8NjMfRV/sD5+RbNOME5NXBa+7CsNXBB6AUAsH+79H6DzOARo\nHFBMt3NLzZ2+C1afTG5/OuCikRqz8W1+Ah26T0RlNLs4xlNN7f6f5v/csAvQMrtb\nH7s4q1yKOo6pEV+sx0nxIIvPUzafySzy5xhXfVzs3QKBgQDXKZ9wf1C0EMR/HJwJ\nigBQhtfMR/dkXaMyn/6oOYP+ilzCl/hz9h28FEQkxnkcoOWUi0SA0HOLKim/yAjK\nQPjR82UHheDJy1HFG8DJKRJ5vgRt3raeJdzEBZleRNkPlkVG/oX6Y6JxPPh+uUR6\nI7UrCDCVrUV2PfXah7FtCoZwiwKBgQDQcIoSrJFEULjTBQR9HNNCPcgBKxPJ63io\nL+Ap6+dLWwRW5Xc5ykecgXP+5tKnvtxFUym/zNzGuG9MvXybj+VmL0dcOL/9KMrf\n1cHbWT9atXkvKdoI6BVM+dIEQk2ymVR5neaYRCm72dTt0XuPRDn+mpvkbkH0dB2V\n8U80fczn6wKBgBn/potiZu8LI6M0YspoSXdxg7vkaklPYzvnx06ufHuSXnkJAfWJ\nJLioPj+vV80UyAj19xYVmhCTzK2mzyEZTreX0HpjO/FGRyApJVdpRhk9060u/yQG\n8TEieBDDAvN2NfLJKIlD5soGEwTFADg9/lMvzycW49DQ1ddNFCxX3DzXAoGAQdRp\nJwCtWGAOH0ckZccmyXLtqcDLHShrEOMd7t98SwZR1ofeIbhK6ZjQ455OxX+33YUo\nWFUBHRm+vQPTSOYnUyB3bx7VM8UgXZULBpWnNqS9wzF6lPmHUKpNs96k31Uhos9F\ntq1AaP2hDALu/Q5VKp31wi04U1YgyexOUYuxnUUCgYBu3zpXirk5PnVt6gfdtiKV\ndkyrPDGRPVMUeBMXab4DXJbFIdwhxOPClJJ4oCVYniMj866KrFu6YsVu9K+8fgZ4\npU55flTbQyf291EKcdoDaxPCrys42NN6RpMMk2aKrBhARgB4rMBELRMha7HDKVtJ\nHYKDWKWDMRMEfcnS2xgUQg==\n-----END PRIVATE KEY-----\n","Private Key (mulai dari -----BEGIN PRIVATE KEY-----)","-----BEGIN PRIVATE KEY-----\\nMIIEv..."],
    ["Fonnte Token","aV84aCSxxvHNY32oupWP","Token Fonnte untuk notifikasi WA","(Opsional) Kosongkan jika tidak pakai Fonnte"],

    // WHATSAPP
    ["  â–¸  WHATSAPP","","",""],
    ["Target Group WA","120363423657558569@g.us","ID grup WA tujuan laporan","Dari Fonnte dashboard"],
    ["WA Report Enabled","Ya","Kirim laporan WA otomatis?","Ya / Tidak"],
    ["WA Report Template","*LAPORAN UNLOADING INBOUND B2B*\n\n*Detail Sesi*\nBrand: [Brand]\nShipment: [Shipment]\nTanggal: [Tanggal]\nPIC Gudang: [PIC]\nLoss Prevention: [LossPreventionName]\n\n*Evaluasi Kerja*\nDurasi Total Sesi: [Duration]\nTotal Dokumentasi: [PhotoCount] Foto\nTotal Koli: [TotalKoli]\n\n*Detail Kendaraan ([FleetCount] Unit):*\n[FleetList]\n\n*Akses Folder Drive:*\n[FolderURL]\n\n_InLens_","Template pesan WA","[Brand] [Shipment] [Tanggal] [PIC] [LossPreventionName] [Duration] [PhotoCount] [TotalKoli] [FleetCount] [FleetList] [FolderURL]"],

    // NAMA SHEET
    ["  â–¸  NAMA SHEET","","",""],
    ["Sessions Sheet","Sessions","Sheet daftar sesi",""],
    ["Fleet Sheet","FleetLogs","Sheet daftar armada",""],
    ["Data Sheet","Data","Sheet master data dropdown",""],
    ["Audit Sheet","AuditLog","Sheet log aktivitas",""],
    ["Photos Sheet","PhotoLogs","Sheet catatan foto",""],
    ["Users Sheet","UserSessions","Sheet user aktif per sesi",""],
    ["Accounts Sheet","Accounts","Sheet login user",""],
    ["Rekap Sheet","Rekap_Sessions","Sheet rekap otomatis",""],
    ["SJ Qty Sheet","SJ_QTY","Sheet SJ Qty handover matching",""],
    ["Handover Revision Log Sheet","HandoverRevisionLog","Sheet log versi PDF handover",""],
    ["Config Sheet","AppConfig","Sheet config ini sendiri",""],

    // MASTER DATA RANGES
    ["  â–¸  MASTER DATA RANGES","","",""],
    ["Brand Column","A","Kolom Brand di sheet Data","A / B / C dst"],
    ["Brand Start Row","2","Baris awal data Brand","Default: 2"],
    ["PIC Column","B","Kolom PIC di sheet Data","A / B / C dst"],
    ["PIC Start Row","2","Baris awal data PIC","Default: 2"],
    ["Vehicle Column","C","Kolom Jenis Kendaraan di sheet Data","A / B / C dst"],
    ["Vehicle Start Row","2","Baris awal data Kendaraan","Default: 2"],
    ["Data Max Row","100","Baris maks scan data master","Naikkan jika data > 100"],

    // UPLOAD
    ["  â–¸  UPLOAD & KOMPRESI","","",""],
    ["Max Image Pixel","1000","Resolusi max sisi terpanjang (px)","800 / 1000 / 1200"],
    ["JPEG Quality","0.65","Kualitas kompresi 0.0-1.0","0.5=kecil 0.65=seimbang 0.8=jernih"],
    ["Compress Concurrency","2","Kompres paralel","1 / 2 / 3"],
    ["Upload Concurrency","3","Upload paralel","2 / 3 / 5"],

    // FOTO URL
    ["  â–¸  FOTO URL","","",""],
    ["Photo URL Template","https://lh3.googleusercontent.com/d/{FILE_ID}=s800","Template URL foto","{FILE_ID} diganti otomatis"],

    // CACHE
    ["  â–¸  CACHE DURATION","","",""],
    ["Cache Folder (detik)","3600","Cache folder ID","3600 = 1 jam"],
    ["Cache Token (detik)","2700","Cache OAuth token","2700 = 45 menit"],
    ["Cache Sync (detik)","20","Cache data sync","20 = balance, 60 = hemat"],

    // SESI & ARMADA
    ["  â–¸  SESI & ARMADA","","",""],
    ["Session ID Prefix","SESS","Awalan ID sesi","SESS / SES / S"],
    ["Container ID Prefix","C","Awalan ID armada","C / ARM / F"],
    ["User ID Prefix","U","Awalan ID user session","U / USR"],
    ["Auto Join On Create","Ya","Auto join saat buat sesi?","Ya / Tidak"],
    ["Require All Done To Finish","Ya","Semua user harus DONE?","Ya / Tidak"],
    ["Require Photo To Finish","Ya","Semua armada harus ada foto?","Ya / Tidak"],

    // STRUKTUR FOLDER
    ["  â–¸  STRUKTUR FOLDER DRIVE","","",""],
    ["Brand Folder Name","[Brand]","Format nama folder brand","[Brand]"],
    ["Session Folder Name","[Shipment] - [Date]","Format nama folder sesi","[Shipment] [Brand] [Date] [PIC]"],
    ["Documentation Folder Name","Dokumentasi","Nama folder induk dokumentasi armada","Dokumentasi"],
    ["Handover Folder Name","Handover","Nama folder PDF handover","Handover"],
    ["Fleet Folder Name","[Number] - [Driver] ([Plate])","Format nama folder armada","[Number] [Driver] [Plate] [Type]"],
    ["Doc Subfolder Name","Unloading","Nama subfolder dokumen","Foto Unloading / Dokumen"],
    ["Palet Subfolder Name","Palet","Nama subfolder palet","Kondisi Palet / Foto Palet"],
    ["Surat Subfolder Name","Surat Jalan","Nama subfolder surat jalan","Surat Jalan / Delivery Note"],

    // LOGIN
    ["  â–¸  LOGIN & SECURITY","","",""],
    ["Login Enabled","Ya","Aktifkan login?","Ya / Tidak"],
    ["Password Case Sensitive","Ya","Password case-sensitive?","Ya / Tidak"],
    ["Username Case Sensitive","Tidak","Username case-sensitive?","Ya / Tidak"],
    ["Max Input Length","50","Maks karakter input","50"],

    // UI
    ["  â–¸  UI / UX","","",""],
    ["App Title","InLens","Judul di browser tab",""],
    ["Sidebar Brand","InLens","Teks brand sidebar",""],
    ["App Version","v5.0.0","Versi di sidebar footer",""],
    ["Default Dark Mode","Tidak","Dark mode by default?","Ya / Tidak"],
    ["Debug Performance","Tidak","Aktifkan log [CRUD_TRACE] untuk audit delay","Ya / Tidak"],
    ["Sync Interval (ms)","15000","Auto-sync interval","10000 / 15000 / 30000"],
    ["Toast Duration (ms)","2500","Durasi notif toast","2000 / 2500 / 3000"],
    ["Dashboard Title","Sesi Aktif","Judul halaman dashboard","Sesi Aktif / Dashboard / Beranda"],

    // HANDOVER REPORT
    ["  â–¸  HANDOVER REPORT","","",""],
    ["Handover Report Title","UNLOADING REPORT B2B","Judul di atas laporan",""],
    ["Handover Logo File ID","1HK8AMDGD65CZ4J6dREN1AgiLBWnEeTrU","ID file gambar logo di Drive",""],
    ["Handover Signature Left","Pic Unloading","Label tanda tangan kiri",""],
    ["Handover Signature Right","Loss Prevention","Label tanda tangan kanan",""],

    // REKAP
    ["  â–¸  REKAP OTOMATIS","","",""],
    ["Rekap Enabled","Ya","Generate rekap saat finish?","Ya / Tidak"],
    ["Rekap Column Order","Tanggal, Brand, Shipment, Sesi ID, Armada Ke, Tipe, Nama Supir, Plat Nomor, Jml Dokumen, Jml Palet, Jml Surat Jalan, Qty Koli, Link Folder Armada","Urutan kolom Rekap (pisah koma)","Tambah/hapus kolom tanpa edit kode"],

    // SYSTEM
    ["  â–¸  SYSTEM","","",""],
    ["Timezone","Asia/Jakarta","Timezone format tanggal","Asia/Jakarta"],
    ["Date Format","dd MMMM yyyy","Format tanggal di laporan","dd-MM-yyyy / dd MMMM yyyy"],
    ["Lock Timeout (ms)","10000","Batas waktu script lock","10000 / 15000 / 30000"],
    ["Archive Cutoff Days","30","Sesi COMPLETED berapa hari sebelum diarsip","30 / 60 / 90"],
    ["Test Mode","Tidak","Mode testing?","Ya / Tidak"],

    // NOTES
    ["","","",""],
    ["ðŸ’¡  Petunjuk Penggunaan","","",""],
    ["â€¢ Ubah nilai di kolom B saja. Kolom A, C, D jangan diubah.","","",""],
    ["â€¢ Baris â–¸ (biru) = pemisah section, tidak dibaca script.","","",""],
    ["â€¢ Perubahan berlaku otomatis (tidak perlu deploy ulang).","","",""],
    ["â€¢ Placeholder folder: [Brand] [Shipment] [Date] [PIC] [Number] [Driver] [Plate] [Type]","","",""],
    ["â€¢ Placeholder WA: [Brand] [Shipment] [Tanggal] [PIC] [FleetCount] [FolderURL]","","",""],
    ["â€¢ Untuk clear cache manual: jalankan fungsi clearConfigCache() dari GAS Editor.","","",""]
  ];

  if (isNew) {
    if (rows.length > 0) s.getRange(2,1,rows.length,4).setValues(rows);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).indexOf("â–¸") >= 0) {
        s.getRange(i+2,1,1,4).setBackground("#DBEAFE").setFontWeight("bold").setFontColor("#1E3A5F");
      }
    }
    for (var i = 0; i < rows.length; i++) {
      var key = String(rows[i][0]).trim();
      if (key && key.indexOf("â–¸") < 0 && key.indexOf("ðŸ’¡") < 0 && key.indexOf("â€¢") < 0 && rows[i][1]) {
        s.getRange(i+2,2).setBackground("#FFFDE7");
      }
    }
  } else {
    // Append missing rows to existing sheet
    var missingRows = [];
    for (var i = 0; i < rows.length; i++) {
      var key = String(rows[i][0]).trim();
      if (!key || key.indexOf("â–¸") >= 0 || key.indexOf("ðŸ’¡") >= 0 || key.indexOf("â€¢") >= 0) continue;
      if (!existingKeys[key]) missingRows.push(rows[i]);
    }
    
    if (missingRows.length > 0) {
      var lastRow = s.getLastRow();
      s.getRange(lastRow + 1, 1, missingRows.length, 4).setValues(missingRows);
      for (var i = 0; i < missingRows.length; i++) {
        s.getRange(lastRow + 1 + i, 2).setBackground("#FFFDE7");
      }
    }
  }

  var protection = s.protect().setDescription("Config â€” hanya edit kolom B");
  moveAppBindingConfigToTop_(s, ss);
  protection.setUnprotectedRanges([s.getRange("B:B")]);
  protection.setWarningOnly(true);
}

// =========================================================================
// verifyLogin() â€” DIHAPUS dari file ini!
// Sekarang HANYA ada di AuthService.gs (satu versi, tidak ada duplikat)
// Alasan: duplikat fungsi di GAS = urutan load random = bug login acak
// =========================================================================
