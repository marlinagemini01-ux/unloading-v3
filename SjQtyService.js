// =========================================================================
// SjQtyService.gs — SJ Qty Handover Matching
// =========================================================================

var SJ_QTY_HEADERS = [
  "sessionId",
  "palletNo",
  "palletKey",
  "sjQty",
  "source",
  "createdBy",
  "updatedAt",
  "shipment",
  "brand",
  "note"
];

function cleanSjPalletName_(value) {
  if (!value && value !== 0) return "";
  return String(value)
    .toUpperCase()
    .replace(/[^A-Z0-9\s\-]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function getSjPalletKey_(value) {
  return cleanSjPalletName_(value);
}

function compareSjPalletNatural_(a, b) {
  var left = cleanSjPalletName_(typeof a === "object" ? (a.palletNo || a.pallet || a.palletKey || "") : a);
  var right = cleanSjPalletName_(typeof b === "object" ? (b.palletNo || b.pallet || b.palletKey || "") : b);
  var leftParts = left.match(/\d+|\D+/g) || [left];
  var rightParts = right.match(/\d+|\D+/g) || [right];
  var len = Math.min(leftParts.length, rightParts.length);
  for (var i = 0; i < len; i++) {
    var lp = leftParts[i];
    var rp = rightParts[i];
    var ln = /^\d+$/.test(lp);
    var rn = /^\d+$/.test(rp);
    if (ln && rn) {
      var diff = Number(lp) - Number(rp);
      if (diff !== 0) return diff;
    } else {
      var cmp = lp.localeCompare(rp);
      if (cmp !== 0) return cmp;
    }
  }
  return leftParts.length - rightParts.length || left.localeCompare(right);
}

function sortSjRowsNatural_(rows) {
  return (rows || []).slice().sort(compareSjPalletNatural_);
}

function getSjQtySheetName_() {
  try {
    var DB = getSheetNames_();
    return DB.SJ_QTY || "SJ_QTY";
  } catch (e) {
    return "SJ_QTY";
  }
}

function ensureSjQtySheet_() {
  var ss = getSpreadsheet_();
  var sheetName = getSjQtySheetName_();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, SJ_QTY_HEADERS.length).setValues([SJ_QTY_HEADERS]);
    sheet.getRange(1, 1, 1, SJ_QTY_HEADERS.length).setFontWeight("bold").setBackground("#f1f5f9");
    sheet.setFrozenRows(1);
    return sheet;
  }

  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  if (headers.length === 0 || !headers[0]) {
    sheet.getRange(1, 1, 1, SJ_QTY_HEADERS.length).setValues([SJ_QTY_HEADERS]);
    sheet.getRange(1, 1, 1, SJ_QTY_HEADERS.length).setFontWeight("bold").setBackground("#f1f5f9");
    sheet.setFrozenRows(1);
    return sheet;
  }

  var missing = [];
  for (var i = 0; i < SJ_QTY_HEADERS.length; i++) {
    if (headers.indexOf(SJ_QTY_HEADERS[i]) < 0) missing.push(SJ_QTY_HEADERS[i]);
  }
  if (missing.length > 0) {
    sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    sheet.getRange(1, lastCol + 1, 1, missing.length).setFontWeight("bold").setBackground("#f1f5f9");
  }
  return sheet;
}

function getSjHeaderMap_(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    map[String(headers[i] || "").trim()] = i;
  }
  return map;
}

function getSjCell_(row, map, name) {
  var idx = map[name];
  return idx === undefined ? "" : row[idx];
}

function sanitizeSjText_(value, maxLen) {
  var clean = "";
  try {
    clean = sanitize_(String(value == null ? "" : value), null);
  } catch (e) {
    clean = String(value == null ? "" : value).trim().replace(/^[=+\-@]+/, "");
  }
  maxLen = maxLen || 100;
  if (clean.length > maxLen) clean = clean.substring(0, maxLen);
  return clean;
}

function parseSjQtyNumber_(value) {
  if (typeof value === "number") {
    if (!isFinite(value) || value < 0) return null;
    return value;
  }
  var raw = String(value == null ? "" : value).trim();
  if (!raw) return null;
  if (!/^\d+([.,]\d+)?$/.test(raw)) return null;
  var num = Number(raw.replace(",", "."));
  if (!isFinite(num) || num < 0) return null;
  return num;
}

function getSjQtySessionMeta_(sessionId) {
  var meta = { brand: "", shipment: "" };
  try {
    var DB = getSheetNames_();
    var sheet = getSpreadsheet_().getSheetByName(DB.SESSIONS);
    if (!sheet) return meta;
    var rows = sheet.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) {
      if (rows[i][0] === sessionId) {
        meta.brand = rows[i][2] || "";
        meta.shipment = rows[i][3] || "";
        break;
      }
    }
  } catch (e) {}
  return meta;
}

function getActualPalletRowsForSj_(sessionId) {
  var DB = getSheetNames_();
  var ss = getSpreadsheet_();
  var fleetSheet = ss.getSheetByName(DB.FLEET);
  var photoSheet = ss.getSheetByName(DB.PHOTOS);
  if (!fleetSheet || !photoSheet) return [];

  var fleetRows = fleetSheet.getDataRange().getValues();
  var activeFleetIds = {};
  for (var i = 1; i < fleetRows.length; i++) {
    if (fleetRows[i][1] === sessionId && fleetRows[i][8] !== "DELETED") {
      activeFleetIds[fleetRows[i][0]] = {
        containerId: fleetRows[i][0],
        containerNo: fleetRows[i][2] || "",
        driver: fleetRows[i][4] || "",
        plate: fleetRows[i][5] || ""
      };
    }
  }

  var rows = [];
  var photoRows = photoSheet.getDataRange().getValues();
  var seenPhotoIds = {};
  for (var p = 1; p < photoRows.length; p++) {
    var photoId = String(photoRows[p][0] || "").trim();
    if (!photoId) continue;
    if (seenPhotoIds[photoId]) continue;
    seenPhotoIds[photoId] = true;
    var container = activeFleetIds[photoRows[p][1]];
    if (!container) continue;
    if (normalizePhotoType_(photoRows[p][2]) !== "palet") continue;

    var palletNo = sanitizeSjText_(photoRows[p][6] || "", 100);
    var palletKey = getSjPalletKey_(palletNo);
    var qty = parseSjQtyNumber_(photoRows[p][7]);
    rows.push({
      palletNo: palletNo,
      palletKey: palletKey,
      actualQty: qty,
      sourceRow: p + 1,
      containerId: container.containerId,
      containerNo: container.containerNo,
      driver: container.driver,
      plate: container.plate,
      valid: !!palletKey && qty !== null && qty > 0
    });
  }
  return rows;
}

function readSjQtyRows_(sessionId) {
  var sheet = ensureSjQtySheet_();
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), SJ_QTY_HEADERS.length);
  if (lastRow < 2) return [];

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0];
  var map = getSjHeaderMap_(headers);
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var rowSessionId = String(getSjCell_(values[i], map, "sessionId") || "").trim();
    if (rowSessionId !== sessionId) continue;
    var palletNo = sanitizeSjText_(getSjCell_(values[i], map, "palletNo") || "", 100);
    var palletKey = String(getSjCell_(values[i], map, "palletKey") || "").trim() || getSjPalletKey_(palletNo);
    var sjQty = parseSjQtyNumber_(getSjCell_(values[i], map, "sjQty"));
    rows.push({
      sessionId: rowSessionId,
      palletNo: palletNo,
      palletKey: palletKey,
      sjQty: sjQty,
      source: getSjCell_(values[i], map, "source") || "",
      createdBy: getSjCell_(values[i], map, "createdBy") || "",
      updatedAt: getSjCell_(values[i], map, "updatedAt") || "",
      shipment: getSjCell_(values[i], map, "shipment") || "",
      brand: getSjCell_(values[i], map, "brand") || "",
      note: getSjCell_(values[i], map, "note") || "",
      sourceRow: i + 1,
      valid: !!palletKey && sjQty !== null
    });
  }
  return rows;
}

function findDuplicateKeys_(rows) {
  var counts = {};
  var labels = {};
  rows.forEach(function(row) {
    var key = row && row.palletKey;
    if (!key) return;
    counts[key] = (counts[key] || 0) + 1;
    if (!labels[key]) labels[key] = row.palletNo || key;
  });

  var result = [];
  for (var key in counts) {
    if (counts[key] > 1) result.push({ palletKey: key, palletNo: labels[key], count: counts[key] });
  }
  return result;
}

function indexSjRowsByKey_(rows) {
  var map = {};
  rows.forEach(function(row) {
    if (row && row.palletKey && !map[row.palletKey]) map[row.palletKey] = row;
  });
  return map;
}

function buildSjValidationDetailRows_(actualRows, sjRows) {
  var actualMap = indexSjRowsByKey_(actualRows || []);
  var sjMap = indexSjRowsByKey_(sjRows || []);
  var keys = {};
  (actualRows || []).forEach(function(row) { if (row && row.palletKey) keys[row.palletKey] = true; });
  (sjRows || []).forEach(function(row) { if (row && row.palletKey) keys[row.palletKey] = true; });

  var detailRows = [];
  for (var key in keys) {
    var actual = actualMap[key] || null;
    var sj = sjMap[key] || null;
    var actualQty = actual ? Number(actual.actualQty || 0) : null;
    var sjQty = sj ? Number(sj.sjQty || 0) : null;
    var status = "Cocok";
    var variance = "";
    if (actual && sj) {
      variance = actualQty - sjQty;
      status = variance === 0 ? "Cocok" : "Selisih";
    } else if (sj) {
      status = "SJ saja";
    } else {
      status = "Actual saja";
    }
    detailRows.push({
      pallet: (actual && actual.palletNo) || (sj && sj.palletNo) || key,
      palletKey: key,
      actual: actual ? actualQty : "",
      sj: sj ? sjQty : "",
      variance: variance,
      status: status
    });
  }
  return detailRows.sort(function(a, b) { return compareSjPalletNatural_(a.pallet || a.palletKey, b.pallet || b.palletKey); });
}

function validateSjQtyRows_(rows) {
  var errors = [];
  var normalized = [];
  var seen = {};
  for (var i = 0; i < rows.length; i++) {
    var raw = rows[i] || {};
    var rowNo = raw.rowNumber || i + 2;
    var palletNo = sanitizeSjText_(raw.palletNo, 100);
    var palletKey = getSjPalletKey_(palletNo);
    var sjQty = parseSjQtyNumber_(raw.sjQty);

    if (!palletNo) errors.push("Baris " + rowNo + ": Pallet No wajib diisi.");
    if (raw.sjQty === "" || raw.sjQty === null || raw.sjQty === undefined) errors.push("Baris " + rowNo + ": SJ Qty wajib diisi.");
    else if (sjQty === null) errors.push("Baris " + rowNo + ": SJ Qty harus angka.");
    if (palletKey && seen[palletKey]) errors.push("Pallet No duplicate di file SJ: " + palletNo + ".");
    if (palletKey) seen[palletKey] = true;

    normalized.push({
      palletNo: palletNo,
      palletKey: palletKey,
      sjQty: sjQty,
      rowNumber: rowNo
    });
  }

  if (normalized.length === 0) errors.push("Minimal ada 1 baris data SJ Qty.");
  return { success: errors.length === 0, errors: errors, rows: normalized };
}

function parseSjQtyImport(rawData) {
  try {
    var matrix = [];
    if (rawData && rawData.length && Object.prototype.toString.call(rawData) === "[object Array]") {
      matrix = rawData;
    } else {
      var raw = String(rawData == null ? "" : rawData).replace(/^\uFEFF/, "").trim();
      if (!raw) return { success: false, message: "File SJ Qty kosong.", data: { errors: ["File SJ Qty kosong."] } };
      var firstLine = raw.split(/\r?\n/).filter(function(line) { return String(line || "").trim(); })[0] || "";
      var delimiter = ",";
      if ((firstLine.split(";").length > firstLine.split(",").length) && firstLine.indexOf(";") >= 0) delimiter = ";";
      if (firstLine.indexOf("\t") >= 0) delimiter = "\t";
      if ((firstLine.split("|").length > firstLine.split(delimiter).length) && firstLine.indexOf("|") >= 0) delimiter = "|";
      matrix = Utilities.parseCsv(raw, delimiter);
    }

    while (matrix.length && matrix[0].every(function(cell) { return String(cell || "").trim() === ""; })) {
      matrix.shift();
    }
    if (matrix.length < 2) {
      return { success: false, message: "Minimal ada 1 baris data SJ Qty.", data: { errors: ["Minimal ada 1 baris data SJ Qty."] } };
    }

    var headers = matrix[0].map(function(h) {
      return String(h || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
    });
    var palletCol = -1;
    var sjCol = -1;
    for (var i = 0; i < headers.length; i++) {
      if (["palletno", "paletno", "palletnumber", "namapalet"].indexOf(headers[i]) >= 0) palletCol = i;
      if (["sjqty", "qtysj", "sjkoli", "kolisj"].indexOf(headers[i]) >= 0) sjCol = i;
    }
    if (palletCol < 0 || sjCol < 0) {
      return { success: false, message: "Header wajib ada: Pallet No dan SJ Qty.", data: { errors: ["Header wajib ada: Pallet No dan SJ Qty."] } };
    }

    var rows = [];
    for (var r = 1; r < matrix.length; r++) {
      var row = matrix[r] || [];
      var empty = row.every(function(cell) { return String(cell || "").trim() === ""; });
      if (empty) continue;
      var sjRaw = row[sjCol];
      if (String(sjRaw == null ? "" : sjRaw).trim() === "") continue;
      rows.push({
        palletNo: row[palletCol],
        sjQty: sjRaw,
        rowNumber: r + 1
      });
    }

    var validation = validateSjQtyRows_(rows);
    if (!validation.success) {
      return { success: false, message: validation.errors[0], data: { errors: validation.errors } };
    }
    return { success: true, message: "File SJ Qty valid.", data: { rows: sortSjRowsNatural_(validation.rows) } };
  } catch (e) {
    return { success: false, message: e.toString(), data: { errors: [e.toString()] } };
  }
}

function replaceSjQtyRows_(sessionId, rows, source, createdBy, traceId) {
  var lock = LockService.getScriptLock();
  perfTrace_("saveSjQty", traceId, "lock_wait_start");
  lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);
  perfTrace_("saveSjQty", traceId, "lock_wait_end");
  try {
    var sheet = ensureSjQtySheet_();
    var lastRow = sheet.getLastRow();
    var lastCol = Math.max(sheet.getLastColumn(), SJ_QTY_HEADERS.length);
    perfTrace_("saveSjQty", traceId, "sheet_read_start");
    var values = lastRow > 0 ? sheet.getRange(1, 1, lastRow, lastCol).getValues() : [SJ_QTY_HEADERS];
    perfTrace_("saveSjQty", traceId, "sheet_read_end", { rows: values.length });
    var headers = values[0] && values[0].length ? values[0] : SJ_QTY_HEADERS;
    var map = getSjHeaderMap_(headers);
    var keptRows = [];
    for (var i = 1; i < values.length; i++) {
      if (String(getSjCell_(values[i], map, "sessionId") || "").trim() !== sessionId) {
        keptRows.push(values[i]);
      }
    }

    var now = new Date();
    var meta = getSjQtySessionMeta_(sessionId);
    rows.forEach(function(row) {
      var output = new Array(lastCol);
      for (var c = 0; c < lastCol; c++) output[c] = "";
      output[map.sessionId] = sessionId;
      output[map.palletNo] = row.palletNo;
      output[map.palletKey] = row.palletKey || getSjPalletKey_(row.palletNo);
      output[map.sjQty] = row.sjQty;
      output[map.source] = source;
      output[map.createdBy] = createdBy || "";
      output[map.updatedAt] = now;
      output[map.shipment] = meta.shipment;
      output[map.brand] = meta.brand;
      output[map.note] = row.note || "";
      keptRows.push(output);
    });

    perfTrace_("saveSjQty", traceId, "sheet_write_start", { rows: keptRows.length });
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    if (keptRows.length > 0) sheet.getRange(2, 1, keptRows.length, lastCol).setValues(keptRows);
    SpreadsheetApp.flush();
    perfTrace_("saveSjQty", traceId, "sheet_write_end");
  } finally {
    lock.releaseLock();
  }

  if (typeof invalidateSyncCache_ === "function") invalidateSyncCache_(sessionId);
}

function saveSjQtyImport(payload) {
  try {
    payload = payload || {};
    var traceId = payload.traceId;
    perfTrace_("saveSjQtyImport", traceId, "server_start", { sessionId: payload.sessionId });
    var sessionId = String(payload.sessionId || "").trim();
    if (!sessionId) return { success: false, message: "Session ID diperlukan." };

    var parsed = payload.rows ? { success: true, data: { rows: payload.rows } } : parseSjQtyImport(payload.rawData);
    if (!parsed.success) return parsed;

    var validation = validateSjQtyRows_(parsed.data.rows || []);
    if (!validation.success) {
      return { success: false, message: validation.errors[0], data: { errors: validation.errors } };
    }

    replaceSjQtyRows_(sessionId, sortSjRowsNatural_(validation.rows), "IMPORT", payload.createdBy || payload.userName || "", traceId);
    perfTrace_("saveSjQtyImport", traceId, "validation_start");
    var statusData = getSjQtyStatus(sessionId).data;
    perfTrace_("saveSjQtyImport", traceId, "validation_end", { status: statusData && statusData.status, actualCount: statusData && statusData.actualCount, sjCount: statusData && statusData.sjCount });
    perfTrace_("saveSjQtyImport", traceId, "server_end", { success: true });
    return {
      success: true,
      message: "SJ Qty berhasil diimport.",
      data: statusData
    };
  } catch (e) {
    perfTrace_("saveSjQtyImport", payload && payload.traceId, "server_error", { message: e.toString() });
    return { success: false, message: e.toString() };
  }
}

function saveSjQtyManual(payload) {
  try {
    payload = payload || {};
    var traceId = payload.traceId;
    perfTrace_("saveSjQtyManual", traceId, "server_start", { sessionId: payload.sessionId });
    var sessionId = String(payload.sessionId || "").trim();
    if (!sessionId) return { success: false, message: "Session ID diperlukan." };

    var actualRows = sortSjRowsNatural_(getActualPalletRowsForSj_(sessionId));
    var validActual = actualRows.filter(function(row) { return row.valid; });
    if (validActual.length === 0) {
      return {
        success: false,
        message: "Input Manual belum tersedia. Silakan input data Palet terlebih dahulu agar sistem memiliki daftar Pallet No dan Actual Qty."
      };
    }

    var duplicateActual = findDuplicateKeys_(validActual);
    if (duplicateActual.length > 0) {
      return { success: false, message: "Ada Pallet No duplicate di data Actual.", data: { duplicateActual: duplicateActual } };
    }

    var inputRows = payload.rows || [];
    var validation = validateSjQtyRows_(inputRows);
    if (!validation.success) {
      return { success: false, message: validation.errors[0], data: { errors: validation.errors } };
    }

    var actualMap = indexSjRowsByKey_(validActual);
    var inputMap = indexSjRowsByKey_(validation.rows);
    var missing = [];
    var extra = [];
    for (var key in actualMap) {
      if (!inputMap[key]) missing.push(actualMap[key].palletNo);
    }
    for (var inputKey in inputMap) {
      if (!actualMap[inputKey]) extra.push(inputMap[inputKey].palletNo);
    }
    if (missing.length || extra.length) {
      return {
        success: false,
        message: "Daftar Pallet No manual harus mengikuti data Actual.",
        data: { missingInManual: missing, extraInManual: extra }
      };
    }

    replaceSjQtyRows_(sessionId, sortSjRowsNatural_(validation.rows), "MANUAL", payload.createdBy || payload.userName || "", traceId);
    perfTrace_("saveSjQtyManual", traceId, "validation_start");
    var statusData = getSjQtyStatus(sessionId).data;
    perfTrace_("saveSjQtyManual", traceId, "validation_end", { status: statusData && statusData.status, actualCount: statusData && statusData.actualCount, sjCount: statusData && statusData.sjCount });
    perfTrace_("saveSjQtyManual", traceId, "server_end", { success: true });
    return {
      success: true,
      message: "SJ Qty manual berhasil disimpan.",
      data: statusData
    };
  } catch (e) {
    perfTrace_("saveSjQtyManual", payload && payload.traceId, "server_error", { message: e.toString() });
    return { success: false, message: e.toString() };
  }
}

function logSjQtyServerDebug_(stage, payload) {
  try {
    console.log("[SJ_QTY_GATE_SERVER] " + stage + " " + JSON.stringify(payload || {}));
  } catch (e) {
    console.log("[SJ_QTY_GATE_SERVER] " + stage);
  }
}

function buildSjQtyValidationErrorData_(message, reasonCode) {
  return {
    status: "ERROR",
    statusText: "SJ Qty gagal dibaca",
    subText: message || "Validasi SJ Qty gagal dibaca dari server.",
    reasonCode: reasonCode || "SERVER_ERROR",
    canHandover: false,
    validationFinal: false,
    actualCount: 0,
    sjCount: 0,
    matchedCount: 0,
    totalActualQty: 0,
    totalSjQty: 0,
    totalVariance: 0,
    missingInActual: [],
    missingInSj: [],
    qtyMismatch: [],
    duplicateActual: [],
    duplicateSj: [],
    invalidActual: [],
    invalidSj: [],
    detailRows: []
  };
}

function makeSjQtyClientSafe_(value) {
  if (value instanceof Date) {
    try {
      return value.toISOString();
    } catch (e) {
      return String(value);
    }
  }
  if (Array.isArray(value)) {
    return value.map(function(item) { return makeSjQtyClientSafe_(item); });
  }
  if (value && typeof value === "object") {
    var out = {};
    for (var key in value) {
      if (!Object.prototype.hasOwnProperty.call(value, key)) continue;
      if (key === "rows" || key === "actualRows") continue;
      if (value[key] === undefined) continue;
      out[key] = makeSjQtyClientSafe_(value[key]);
    }
    return out;
  }
  return value;
}

function validateSjQtyMatching(sessionId, traceId) {
  try {
    perfTrace_("sjQtyValidation", traceId, "server_start", { sessionId: sessionId });
    sessionId = String(sessionId || "").trim();
    logSjQtyServerDebug_("validateSjQtyMatching:start", { requestId: traceId || "", sessionId: sessionId });
    if (!sessionId) {
      logSjQtyServerDebug_("validateSjQtyMatching:block", { requestId: traceId || "", sessionId: sessionId, reasonCode: "EMPTY_SESSION_ID" });
      return { success: false, message: "Session ID diperlukan.", data: buildSjQtyValidationErrorData_("Session ID diperlukan.", "EMPTY_SESSION_ID") };
    }

    perfTrace_("sjQtyValidation", traceId, "actual_read_start");
    var actualRowsAll = getActualPalletRowsForSj_(sessionId);
    perfTrace_("sjQtyValidation", traceId, "actual_read_end", { rows: actualRowsAll.length });
    var actualRows = actualRowsAll.filter(function(row) { return row.valid; });
    perfTrace_("sjQtyValidation", traceId, "sj_read_start");
    var sjRowsAll = readSjQtyRows_(sessionId);
    perfTrace_("sjQtyValidation", traceId, "sj_read_end", { rows: sjRowsAll.length });
    var sjRows = sjRowsAll.filter(function(row) { return row.valid; });

    var duplicateActual = findDuplicateKeys_(actualRows);
    var duplicateSj = findDuplicateKeys_(sjRows);
    var invalidActual = actualRowsAll.filter(function(row) { return !row.valid; });
    var invalidSj = sjRowsAll.filter(function(row) { return !row.valid; });
    var detailRows = buildSjValidationDetailRows_(actualRows, sjRows);
    var qtyMismatch = detailRows.filter(function(row) { return row.status === "Selisih"; }).map(function(row) { return row.pallet; });

    var actualMap = indexSjRowsByKey_(actualRows);
    var sjMap = indexSjRowsByKey_(sjRows);
    var missingInActual = [];
    var missingInSj = [];
    var matched = 0;
    var totalActualQty = 0;
    var totalSjQty = 0;

    for (var aKey in actualMap) {
      totalActualQty += Number(actualMap[aKey].actualQty || 0);
      if (sjMap[aKey]) matched++;
      else missingInSj.push(actualMap[aKey].palletNo || aKey);
    }
    for (var sKey in sjMap) {
      totalSjQty += Number(sjMap[sKey].sjQty || 0);
      if (!actualMap[sKey]) missingInActual.push(sjMap[sKey].palletNo || sKey);
    }

    var status = "EMPTY";
    var statusText = "SJ Qty belum diisi";
    var subText = "";
    var reasonCode = "NO_SJ_QTY";
    var canHandover = false;

    if (sjRowsAll.length > 0 && sjRows.length === 0) {
      status = "MISMATCH";
      statusText = "Perlu perbaikan";
      subText = "Ada Pallet No yang tidak cocok antara Actual dan SJ";
      reasonCode = "INVALID_SJ_ROWS";
    } else if (sjRows.length === 0) {
      status = "EMPTY";
      statusText = "SJ Qty belum diisi";
      subText = "Import atau isi manual sebelum handover.";
      reasonCode = "NO_SJ_QTY";
    } else if (actualRows.length === 0) {
      status = "WAITING_ACTUAL";
      statusText = "SJ Qty sudah diimport";
      subText = "Menunggu data actual";
      reasonCode = "WAITING_ACTUAL";
    } else if (duplicateActual.length || duplicateSj.length || invalidActual.length || invalidSj.length || missingInSj.length > 0) {
      status = "MISMATCH";
      statusText = "Perlu perbaikan";
      subText = "Ada Pallet No yang tidak cocok antara Actual dan SJ";
      reasonCode = "MISMATCH";
    } else if (missingInActual.length > 0) {
      status = "MATCHING_TEMPORARY";
      statusText = "Matching sementara";
      subText = matched + "/" + sjRows.length + " pallet cocok";
      reasonCode = "SJ_EXTRA";
    } else {
      status = "VALID";
      statusText = qtyMismatch.length > 0 ? "Valid dengan selisih qty" : "Valid";
      subText = qtyMismatch.length > 0 ? "Qty Actual dan SJ selisih, tetapi tetap boleh handover." : "Data siap handover";
      reasonCode = qtyMismatch.length > 0 ? "VALID_QTY_VARIANCE" : "VALID";
      canHandover = true;
    }

    var data = {
      status: status,
      statusText: statusText,
      subText: subText,
      reasonCode: reasonCode,
      canHandover: canHandover,
      validationFinal: true,
      actualCount: actualRows.length,
      sjCount: sjRows.length,
      matchedCount: matched,
      totalActualQty: totalActualQty,
      totalSjQty: totalSjQty,
      totalVariance: totalActualQty - totalSjQty,
      missingInActual: sortSjRowsNatural_(missingInActual),
      missingInSj: sortSjRowsNatural_(missingInSj),
      qtyMismatch: sortSjRowsNatural_(qtyMismatch),
      duplicateActual: sortSjRowsNatural_(duplicateActual),
      duplicateSj: sortSjRowsNatural_(duplicateSj),
      invalidActual: invalidActual,
      invalidSj: invalidSj,
      detailRows: detailRows,
      rows: sjRows,
      actualRows: actualRows
    };

    logSjQtyServerDebug_("validateSjQtyMatching:end", {
      requestId: traceId || "",
      sessionId: sessionId,
      actualCount: data.actualCount,
      sjCount: data.sjCount,
      matchedCount: data.matchedCount,
      qtyMismatch: data.qtyMismatch,
      canHandover: data.canHandover,
      validationFinal: data.validationFinal,
      reasonCode: data.reasonCode
    });
    perfTrace_("sjQtyValidation", traceId, "server_end", {
      status: status,
      actualCount: data.actualCount,
      sjCount: data.sjCount,
      matchedCount: data.matchedCount,
      qtyMismatch: data.qtyMismatch,
      canHandover: canHandover,
      validationFinal: data.validationFinal,
      reasonCode: data.reasonCode
    });
    return { success: true, message: statusText, data: makeSjQtyClientSafe_(data) };
  } catch (e) {
    var message = e && e.toString ? e.toString() : String(e);
    logSjQtyServerDebug_("validateSjQtyMatching:error", { requestId: traceId || "", sessionId: String(sessionId || "").trim(), error: message });
    perfTrace_("sjQtyValidation", traceId, "server_error", { message: message });
    return { success: false, message: message, data: buildSjQtyValidationErrorData_(message, "SERVER_ERROR") };
  }
}

function buildSjQtyHandoverBlockMessage_(validation) {
  if (!validation || !validation.success) {
    return validation && validation.message ? validation.message : "Validasi SJ Qty gagal dibaca dari server.";
  }
  var data = validation.data || {};
  var status = String(data.status || "").toUpperCase();
  var actualCount = Number(data.actualCount || 0);
  var sjCount = Number(data.sjCount || 0);
  if (actualCount === 0 && sjCount === 0) {
    return "Data validasi belum terbaca oleh server. Pastikan upload/log selesai, SJ Qty sudah tersimpan, lalu sinkronkan ulang.";
  }
  if (status === "EMPTY" || sjCount === 0) {
    return "SJ Qty belum tersimpan atau belum terbaca oleh server. Import atau input manual SJ Qty sebelum handover.";
  }
  if (status === "WAITING_ACTUAL" || actualCount === 0) {
    return "SJ Qty sudah terbaca, tetapi data Actual belum terbaca oleh server. Pastikan upload foto Palet dan log Sheet sudah selesai.";
  }
  if (status === "MATCHING_TEMPORARY") {
    return data.subText || "Sebagian data cocok, tetapi masih ada data SJ yang belum punya Actual.";
  }
  if ((data.detailRows || []).length > 0 && (data.detailRows || []).every(function(row) { return String(row.status || "") === "Cocok"; })) {
    return "Detail baris terlihat cocok, tetapi server belum mengizinkan handover. Status final: " + (data.statusText || status) + ". Sinkronkan ulang dan cek pending upload/log.";
  }
  return data.subText || "Masih ada Pallet No atau Qty yang tidak cocok antara Actual dan SJ.";
}

function isSjQtyFinalReadyForHandoverServer_(data) {
  return !!data
    && String(data.status || "").toUpperCase() === "VALID"
    && data.canHandover === true
    && data.validationFinal === true;
}

function ensureSjQtyReadyForHandover_(sessionId, traceId) {
  var validation = validateSjQtyMatching(sessionId, traceId);
  if (validation && validation.success && isSjQtyFinalReadyForHandoverServer_(validation.data)) {
    return { success: true, message: "SJ Qty valid.", data: validation.data };
  }
  var data = validation && validation.data ? validation.data : null;
  return {
    success: false,
    code: "SJ_QTY_INVALID",
    message: buildSjQtyHandoverBlockMessage_(validation),
    data: data,
    sjQty: data
  };
}

function getSjQtyStatus(sessionId) {
  var res = validateSjQtyMatching(sessionId);
  if (res && res.success && res.data) {
    delete res.data.rows;
    delete res.data.actualRows;
  }
  return res;
}

function getSjQtyRows(sessionId) {
  try {
    sessionId = String(sessionId || "").trim();
    if (!sessionId) return { success: false, message: "Session ID diperlukan." };
    return { success: true, message: "OK", data: { rows: makeSjQtyClientSafe_(sortSjRowsNatural_(readSjQtyRows_(sessionId))) } };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getSjQtyMap_(sessionId) {
  return indexSjRowsByKey_(readSjQtyRows_(sessionId).filter(function(row) { return row.valid; }));
}

function buildCsvCell_(value) {
  var text = String(value == null ? "" : value);
  if (/[",\r\n]/.test(text)) text = '"' + text.replace(/"/g, '""') + '"';
  return text;
}

function buildSjQtyTemplate(sessionId) {
  try {
    sessionId = String(sessionId || "").trim();
    if (!sessionId) return { success: false, message: "Session ID diperlukan." };

    var actualRows = sortSjRowsNatural_(getActualPalletRowsForSj_(sessionId).filter(function(row) { return row.valid; }));
    var sjMap = getSjQtyMap_(sessionId);
    var lines = [];
    if (actualRows.length > 0) {
      lines.push(["Pallet No", "Actual Qty", "SJ Qty"].map(buildCsvCell_).join(","));
      actualRows.forEach(function(row) {
        var sj = sjMap[row.palletKey] ? sjMap[row.palletKey].sjQty : "";
        lines.push([row.palletNo, row.actualQty, sj].map(buildCsvCell_).join(","));
      });
    } else {
      lines.push(["Pallet No", "SJ Qty"].map(buildCsvCell_).join(","));
      for (var i = 1; i <= 10; i++) {
        lines.push([i, ""].map(buildCsvCell_).join(","));
      }
    }

    return {
      success: true,
      message: "Template SJ Qty siap.",
      data: {
        fileName: "SJ_Qty_Template_" + sessionId + ".csv",
        mimeType: "text/csv;charset=utf-8",
        content: lines.join("\r\n")
      }
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
