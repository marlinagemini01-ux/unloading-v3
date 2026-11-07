// =========================================================================
// addContainer() — OPTIMASI: Lazy folder (skip Drive API = 2-5 detik hemat)
// Folder armada dibuat saat user pertama kali upload foto, bukan saat add.
// =========================================================================
function addContainer(data) {
  var traceId = data && data.traceId;
  perfTrace_("createContainer", traceId, "server_start", { sessionId: data && data.sessionId });
  var lock = LockService.getScriptLock();
  try {
    perfTrace_("createContainer", traceId, "lock_wait_start");
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);
    perfTrace_("createContainer", traceId, "lock_wait_end");

    var driver = titleCaseText_(data.driver);
    var plate = sanitize_(data.plate);
    var type = sanitize_(data.type);
    validateRequired_({ "Nama supir": driver, "Plat nomor": plate });

    var containerId = generateUniqueId_("Container ID Prefix");
    // Qty Koli adalah nilai turunan dari foto Palet, bukan input manual armada.
    var qtyKoli = 0;
    var DB = getSheetNames_();
    var sheet = getSpreadsheet_().getSheetByName(DB.FLEET);
    var fleetMap = ensureFleetWorkColumns_(sheet);
    perfTrace_("createContainer", traceId, "sheet_read_start");
    var dataRange = sheet.getDataRange().getValues();
    perfTrace_("createContainer", traceId, "sheet_read_end", { rows: dataRange.length });

    var fleetNumber = 1;
    for (var i = 1; i < dataRange.length; i++) {
      if (dataRange[i][1] === data.sessionId && dataRange[i][8] !== "DELETED")
        fleetNumber++;
    }

    // OPTIMASI: Folder armada TIDAK dibuat di sini
    // Folder akan dibuat LAZY saat requestUploadCredentials() — hemat 2-5 detik
    perfTrace_("createContainer", traceId, "sheet_write_start");
    var rowData = [
      containerId,
      data.sessionId,
      fleetNumber,
      type,
      driver,
      plate.toUpperCase(),
      new Date(),
      "",
      "ACTIVE",
      sanitize_(data.createdBy) || "unknown",
      qtyKoli
    ];
    while (rowData.length < sheet.getLastColumn()) rowData.push("");
    setFleetCell_(rowData, fleetMap, "Work Status", "READY");
    setFleetCell_(rowData, fleetMap, "Started At", "");
    setFleetCell_(rowData, fleetMap, "Started By", "");
    setFleetCell_(rowData, fleetMap, "Completed At", "");
    setFleetCell_(rowData, fleetMap, "Completed By", "");
    setFleetCell_(rowData, fleetMap, "Reopened At", "");
    setFleetCell_(rowData, fleetMap, "Reopened By", "");
    sheet.appendRow(rowData);
    perfTrace_("createContainer", traceId, "sheet_write_end", { containerId: containerId, number: fleetNumber });

    rebuildAndCacheSync_(data.sessionId);
    perfTrace_("createContainer", traceId, "server_end", { success: true });
    return {
      success: true,
      data: { containerId: containerId, number: fleetNumber },
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// updateKoli() — kompatibilitas lama; Qty Koli tetap dihitung dari foto Palet
// =========================================================================
function updateKoli(containerId, qtyKoli) {
  try {
    var DB = getSheetNames_();
    var sheet = getSpreadsheet_().getSheetByName(DB.FLEET);
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === containerId && rows[i][8] !== "DELETED") {
        var sessionId = rows[i][1];
        syncFleetKoliFromPhotos_(sessionId);
        rows = sheet.getDataRange().getValues();
        var autoQty = 0;
        for (var j = 1; j < rows.length; j++) {
          if (rows[j][0] === containerId && rows[j][8] !== "DELETED") {
            autoQty = Math.max(0, Math.floor(Number(rows[j][10]) || 0));
            break;
          }
        }
        rebuildAndCacheSync_(sessionId);
        return { success: true, data: { qtyKoli: autoQty, autoCalculated: true } };
      }
    }
    return { success: false, message: "Armada tidak ditemukan." };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// =========================================================================
// updateContainer() — OPTIMASI: 1x setValues() bukan 3x setValue()
// + PERBAIKAN: Rename folder Drive saat data armada berubah
// + PERBAIKAN: Invalidasi cache folder lama
// =========================================================================
function normalizeFleetActorData_(actor, traceId) {
  var data = {};
  if (actor && typeof actor === "object") {
    data = actor;
  } else {
    data = { requestedBy: actor, userName: actor, username: actor };
  }
  if (traceId && !data.traceId) data.traceId = traceId;
  return data;
}

function readFleetWorkRecord_(containerId, ss) {
  var cleanId = String(containerId || "").trim();
  if (!cleanId) throw new Error("Container ID diperlukan.");
  var DB = getSheetNames_();
  ss = ss || getSpreadsheet_();
  var sheet = ss.getSheetByName(DB.FLEET);
  var map = ensureFleetWorkColumns_(sheet);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0] || "").trim() === cleanId && String(rows[i][8] || "").trim().toUpperCase() !== "DELETED") {
      return {
        sheet: sheet,
        map: map,
        row: rows[i],
        rowIndex: i + 1,
        containerId: cleanId,
        sessionId: rows[i][1],
        number: rows[i][2],
        driver: rows[i][4],
        plate: rows[i][5],
        status: normalizeFleetWorkStatus_(getFleetCell_(rows[i], map, "Work Status"))
      };
    }
  }
  throw new Error("Armada tidak ditemukan atau sudah dihapus.");
}

function assertSessionIsActive_(sessionId, ss) {
  var DB = getSheetNames_();
  ss = ss || getSpreadsheet_();
  var sheet = ss.getSheetByName(DB.SESSIONS);
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === sessionId) {
      if (String(rows[i][5] || "").trim().toUpperCase() !== "ACTIVE") {
        throw new Error("Sesi sudah selesai, data armada tidak bisa diubah.");
      }
      return;
    }
  }
  throw new Error("Sesi armada tidak ditemukan.");
}

function assertWorkingUserForFleet_(sessionId, actorData, ss) {
  var DB = getSheetNames_();
  ss = ss || getSpreadsheet_();
  var sheet = ss.getSheetByName(DB.USERS);
  var fallbackName = getDisplayUserName_(actorData);
  if (!sheet) return fallbackName;
  var rows = sheet.getDataRange().getValues();
  var candidates = buildUserNameCandidates_(actorData);
  var sessionUserCount = 0;
  var matchedDone = "";

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][1] !== sessionId) continue;
    sessionUserCount++;
    var name = String(rows[i][2] || "").trim();
    var key = name.toLowerCase();
    if (!candidates[key]) continue;
    var userStatus = String(rows[i][3] || "").trim().toUpperCase();
    if (userStatus === "WORKING") return name || fallbackName;
    if (userStatus === "DONE") matchedDone = name || fallbackName;
  }

  if (matchedDone) throw new Error("Status Anda sudah selesai. Klik Batal dulu sebelum mengubah armada.");
  if (sessionUserCount === 0) return fallbackName;
  throw new Error("Anda belum berstatus WORKING di sesi ini. Gabung atau klik Batal dulu sebelum mengubah armada.");
}

function assertFleetWorkStatusAllowed_(containerId, allowedStatuses, actionLabel) {
  // PATCH: ijinkan UPLOAD_ONLY bertindak seperti IN_PROGRESS untuk aksi terkait foto/palet
  if (allowedStatuses && allowedStatuses.indexOf("IN_PROGRESS") >= 0 && allowedStatuses.indexOf("UPLOAD_ONLY") < 0) {
    allowedStatuses.push("UPLOAD_ONLY");
  }
  var record = readFleetWorkRecord_(containerId);
  var allowed = {};
  for (var i = 0; i < (allowedStatuses || []).length; i++) {
    allowed[normalizeFleetWorkStatus_(allowedStatuses[i])] = true;
  }
  if (!allowed[record.status]) {
    var action = actionLabel || "aksi ini";
    if (record.status === "READY") throw new Error("Klik Mulai Armada dulu sebelum " + action + ".");
    if (record.status === "DONE") throw new Error("Armada sudah di-finish. Klik Batal Finish dulu sebelum " + action + ".");
    throw new Error("Status armada tidak valid untuk " + action + ".");
  }
  return record;
}

function assertFleetUploadAllowed_(containerId) {
  return assertFleetWorkStatusAllowed_(containerId, ["IN_PROGRESS"], "upload foto");
}

function getUnfinishedFleetWorkForSession_(sessionId, ss) {
  var cleanSessionId = String(sessionId || "").trim();
  if (!cleanSessionId) return [];
  var DB = getSheetNames_();
  ss = ss || getSpreadsheet_();
  var sheet = ss.getSheetByName(DB.FLEET);
  if (!sheet) return [];
  var map = ensureFleetWorkColumns_(sheet);
  var rows = sheet.getDataRange().getValues();
  var pending = [];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1] || "").trim() !== cleanSessionId) continue;
    if (String(rows[i][8] || "").trim().toUpperCase() === "DELETED") continue;
    var fleetWork = buildFleetWorkSnapshot_(rows[i], map);
    if (fleetWork.workStatus !== "DONE" || !fleetWork.startedAtMs || !fleetWork.completedAtMs) {
      pending.push({
        id: rows[i][0],
        number: rows[i][2],
        driver: rows[i][4],
        plate: rows[i][5],
        workStatus: fleetWork.workStatus,
        startedAtMs: fleetWork.startedAtMs,
        completedAtMs: fleetWork.completedAtMs
      });
    }
  }
  return pending;
}

function assertUploadItemsWorkAllowed_(items, fallbackContainerId) {
  var checked = {};
  var list = items || [];
  if (list.length === 0 && fallbackContainerId) list = [{ containerId: fallbackContainerId }];
  for (var i = 0; i < list.length; i++) {
    var cid = String((list[i] && list[i].containerId) || fallbackContainerId || "").trim();
    if (!cid || checked[cid]) continue;
    assertFleetUploadAllowed_(cid);
    checked[cid] = true;
  }
}

function validateFleetReadyToFinish_(containerId) {
  var DB = getSheetNames_();
  var photoSheet = getSpreadsheet_().getSheetByName(DB.PHOTOS);
  var stats = { dokumen: 0, palet: 0, suratjalan: 0 };
  var paletValid = 0;
  var paletIncomplete = 0;
  var rows = photoSheet ? photoSheet.getDataRange().getValues() : [];
  var seen = {};
  for (var i = 1; i < rows.length; i++) {
    var photoId = String(rows[i][0] || "").trim();
    if (!photoId || seen[photoId]) continue;
    seen[photoId] = true;
    if (rows[i][1] !== containerId) continue;
    var type = normalizePhotoType_(rows[i][2]);
    if (stats[type] !== undefined) stats[type]++;
    if (type === "palet") {
      var paletName = String(rows[i][6] || "").trim();
      var paletKoli = Number(rows[i][7]) || 0;
      if (paletName && paletKoli > 0) paletValid++;
      else paletIncomplete++;
    }
  }

  var missingTypes = [];
  if (stats.dokumen === 0) missingTypes.push(cfg("Doc Subfolder Name") || "Unloading");
  if (stats.palet === 0) missingTypes.push(cfg("Palet Subfolder Name") || "Palet");
  if (stats.suratjalan === 0) missingTypes.push(cfg("Surat Subfolder Name") || "Surat Jalan");
  if (missingTypes.length > 0) {
    return { success: false, code: "PHOTO_TYPES_REQUIRED", message: "Armada belum punya foto wajib: " + missingTypes.join(", ") + ".", missingTypes: missingTypes };
  }
  if (paletValid === 0) {
    return { success: false, code: "PALET_REQUIRED", message: "Armada belum punya foto Palet valid dengan Nama Palet dan Jumlah Koli." };
  }
  if (paletIncomplete > 0) {
    return { success: false, code: "PALET_INCOMPLETE", message: "Masih ada foto Palet yang belum memiliki Nama Palet dan Jumlah Koli.", incompleteCount: paletIncomplete };
  }
  return { success: true };
}

function writeFleetWorkSnapshot_(record, newStatus, extra) {
  extra = extra || {};
  var row = record.row.slice();
  while (row.length < record.sheet.getLastColumn()) row.push("");
  setFleetCell_(row, record.map, "Work Status", newStatus);
  if (extra.startedAt !== undefined) setFleetCell_(row, record.map, "Started At", extra.startedAt);
  if (extra.startedBy !== undefined) setFleetCell_(row, record.map, "Started By", extra.startedBy);
  if (extra.completedAt !== undefined) setFleetCell_(row, record.map, "Completed At", extra.completedAt);
  if (extra.completedBy !== undefined) setFleetCell_(row, record.map, "Completed By", extra.completedBy);
  if (extra.reopenedAt !== undefined) setFleetCell_(row, record.map, "Reopened At", extra.reopenedAt);
  if (extra.reopenedBy !== undefined) setFleetCell_(row, record.map, "Reopened By", extra.reopenedBy);
  if (extra.accumulatedWorkMs !== undefined) setFleetCell_(row, record.map, "Accumulated Work Ms", extra.accumulatedWorkMs);
  record.sheet.getRange(record.rowIndex, 1, 1, row.length).setValues([row]);
  record.row = row;
  record.status = normalizeFleetWorkStatus_(newStatus);
  return buildFleetWorkSnapshot_(row, record.map);
}

function startFleetWork(containerId, actor, traceId) {
  var actorData = normalizeFleetActorData_(actor, traceId);
  traceId = actorData.traceId || traceId;
  perfTrace_("startFleetWork", traceId, "server_start", { containerId: containerId });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);
    var ss = getSpreadsheet_();
    var record = readFleetWorkRecord_(containerId, ss);
    assertSessionIsActive_(record.sessionId, ss);
    var actorName = assertWorkingUserForFleet_(record.sessionId, actorData, ss);
    if (record.status === "DONE") {
      return { success: false, code: "FLEET_ALREADY_DONE", message: "Armada sudah di-finish. Klik Batal Finish dulu jika perlu upload/edit lagi." };
    }
    var now = new Date();
    var snapshot = writeFleetWorkSnapshot_(record, "IN_PROGRESS", {
      startedAt: getFleetCell_(record.row, record.map, "Started At") || now,
      startedBy: getFleetCell_(record.row, record.map, "Started By") || actorName,
      completedAt: "",
      completedBy: ""
    });
    rebuildAndCacheSync_(record.sessionId);
    return { success: true, data: { containerId: record.containerId, sessionId: record.sessionId, workStatus: "IN_PROGRESS", actor: actorName, lifecycle: snapshot } };
  } catch (e) {
    perfTrace_("startFleetWork", traceId, "server_error", { message: e.toString() });
    return { success: false, message: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (releaseErr) {}
  }
}

function finishFleetWork(containerId, actor, traceId) {
  var actorData = normalizeFleetActorData_(actor, traceId);
  traceId = actorData.traceId || traceId;
  perfTrace_("finishFleetWork", traceId, "server_start", { containerId: containerId });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);
    var ss = getSpreadsheet_();
    var record = readFleetWorkRecord_(containerId, ss);
    assertSessionIsActive_(record.sessionId, ss);
    var actorName = assertWorkingUserForFleet_(record.sessionId, actorData, ss);
    if (record.status === "READY") return { success: false, code: "FLEET_NOT_STARTED", message: "Klik Mulai Armada dulu sebelum Finish Armada." };
    if (record.status === "DONE") {
      return { success: true, data: { containerId: record.containerId, sessionId: record.sessionId, workStatus: "DONE", actor: actorName, lifecycle: buildFleetWorkSnapshot_(record.row, record.map) } };
    }
    var validation = validateFleetReadyToFinish_(record.containerId);
    if (!validation.success) {
      validation.container = { id: record.containerId, number: record.number, driver: record.driver, plate: record.plate };
      return validation;
    }
    var now = new Date();
    var nowMs = now.getTime();
    var startedMs = toTimestampMs_(getFleetCell_(record.row, record.map, "Started At"));
    var reopenedMs = toTimestampMs_(getFleetCell_(record.row, record.map, "Reopened At"));
    var segmentStartMs = reopenedMs > 0 ? reopenedMs : startedMs;
    var segmentDuration = (record.status === "IN_PROGRESS" && segmentStartMs > 0) ? Math.max(0, nowMs - segmentStartMs) : 0;
    var accMs = Number(getFleetCell_(record.row, record.map, "Accumulated Work Ms")) || 0;
    var snapshot = writeFleetWorkSnapshot_(record, "DONE", {
      completedAt: now,
      completedBy: actorName,
      accumulatedWorkMs: accMs + segmentDuration
    });
    rebuildAndCacheSync_(record.sessionId);
    syncFleetKoliFromPhotos_(record.sessionId, { skipLock: true });
    return { success: true, data: { containerId: record.containerId, sessionId: record.sessionId, workStatus: "DONE", actor: actorName, lifecycle: snapshot } };
  } catch (e) {
    perfTrace_("finishFleetWork", traceId, "server_error", { message: e.toString() });
    return { success: false, message: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (releaseErr) {}
  }
}

function reopenFleetWork(containerId, actor, traceId) {
  var actorData = normalizeFleetActorData_(actor, traceId);
  traceId = actorData.traceId || traceId;
  perfTrace_("reopenFleetWork", traceId, "server_start", { containerId: containerId });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);
    var ss = getSpreadsheet_();
    var record = readFleetWorkRecord_(containerId, ss);
    assertSessionIsActive_(record.sessionId, ss);
    var actorName = assertWorkingUserForFleet_(record.sessionId, actorData, ss);
    if (record.status === "READY") return { success: false, code: "FLEET_NOT_DONE", message: "Armada belum di-finish." };
    if (record.status === "IN_PROGRESS") {
      return { success: true, data: { containerId: record.containerId, sessionId: record.sessionId, workStatus: "IN_PROGRESS", actor: actorName, lifecycle: buildFleetWorkSnapshot_(record.row, record.map) } };
    }
    var now = new Date();
    var snapshot = writeFleetWorkSnapshot_(record, "IN_PROGRESS", {
      startedAt: getFleetCell_(record.row, record.map, "Started At") || now,
      startedBy: getFleetCell_(record.row, record.map, "Started By") || actorName,
      completedAt: "",
      completedBy: "",
      reopenedAt: now,
      reopenedBy: actorName
    });
    rebuildAndCacheSync_(record.sessionId);
    return { success: true, data: { containerId: record.containerId, sessionId: record.sessionId, workStatus: "IN_PROGRESS", actor: actorName, lifecycle: snapshot } };
  } catch (e) {
    perfTrace_("reopenFleetWork", traceId, "server_error", { message: e.toString() });
    return { success: false, message: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (releaseErr) {}
  }
}

function updateContainer(containerId, updates) {
  updates = updates || {};
  var traceId = updates.traceId;
  perfTrace_("updateContainer", traceId, "server_start", { containerId: containerId });
  var DB = getSheetNames_();
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(DB.FLEET);
  var fleetMap = ensureFleetWorkColumns_(sheet);
  var lock = LockService.getScriptLock();
  var found = false;
  var sessionId = "";
  var renameContext = null;

  try {
    perfTrace_("updateContainer", traceId, "lock_wait_start");
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);
    perfTrace_("updateContainer", traceId, "lock_wait_end");

    perfTrace_("updateContainer", traceId, "sheet_read_start");
    var data = sheet.getDataRange().getValues();
    perfTrace_("updateContainer", traceId, "sheet_read_end", { rows: data.length });

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === containerId && data[i][8] !== "DELETED") {
        found = true;
        if (normalizeFleetWorkStatus_(getFleetCell_(data[i], fleetMap, "Work Status")) === "DONE") {
          return { success: false, message: "Armada sudah di-finish. Klik Batal Finish dulu sebelum mengedit armada." };
        }
        var row = i + 1;
        sessionId = data[i][1];
        var fleetNumber = data[i][2];
        var oldDriver = String(data[i][4] || "").trim();
        var oldPlate = String(data[i][5] || "").trim();
        var oldType = String(data[i][3] || "").trim();

        var newType = updates.type || data[i][3];
        var newDriver = updates.driver ? titleCaseText_(updates.driver) : data[i][4];
        var newPlate = updates.plate ? updates.plate.toUpperCase() : data[i][5];

        perfTrace_("updateContainer", traceId, "sheet_write_start");
        sheet.getRange(row, 4, 1, 3).setValues([[newType, newDriver, newPlate]]);
        perfTrace_("updateContainer", traceId, "sheet_write_end");

        if (newDriver.toLowerCase() !== oldDriver.toLowerCase()
            || newPlate.toLowerCase() !== oldPlate.toLowerCase()
            || newType.toLowerCase() !== oldType.toLowerCase()) {
          renameContext = {
            sessionId: sessionId,
            fleetNumber: fleetNumber,
            oldDriver: oldDriver,
            oldPlate: oldPlate,
            oldType: oldType,
            newDriver: newDriver,
            newPlate: newPlate,
            newType: newType
          };
        }
        break;
      }
    }
  } catch (e) {
    perfTrace_("updateContainer", traceId, "server_error", { message: e.toString() });
    return { success: false, message: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (releaseErr) {}
  }

  if (!found) {
    perfTrace_("updateContainer", traceId, "server_end", { success: false, reason: "not_found" });
    return { success: false, message: "Armada tidak ditemukan." };
  }

  rebuildAndCacheSync_(sessionId);

  var driveWarning = "";
  if (renameContext) {
    perfTrace_("updateContainer", traceId, "drive_rename_start");
    try {
      var sSession = ss.getSheetByName(DB.SESSIONS);
      var sessRows = sSession.getDataRange().getValues();
      var sessData = null;
      for (var s = 1; s < sessRows.length; s++) {
        if (sessRows[s][0] === renameContext.sessionId) { sessData = sessRows[s]; break; }
      }

      if (sessData) {
        var tz = cfg("Timezone") || "Asia/Jakarta";
        var rawDate = sessData[1];
        var sessDate = rawDate instanceof Date
          ? Utilities.formatDate(rawDate, tz, "yyyy-MM-dd")
          : String(rawDate).trim();

        var rootFolder = getRootFolder_();
        var rootId = rootFolder.getId();
        var oldPlaceholders = {
          Brand: String(sessData[2]).trim(),
          Shipment: String(sessData[3]).trim(),
          Date: sessDate,
          PIC: String(sessData[4]).trim(),
          Number: renameContext.fleetNumber,
          Driver: renameContext.oldDriver,
          Plate: renameContext.oldPlate,
          Type: renameContext.oldType
        };
        var brandName = safeFolderName_(replacePlaceholders_(cfg("Brand Folder Name") || "[Brand]", oldPlaceholders));
        var sessionName = safeFolderName_(replacePlaceholders_(cfg("Session Folder Name") || "[Shipment] - [Date]", oldPlaceholders));
        var sKey = "f_" + rootId + "_" + brandName + "_" + sessionName;
        var oldFleetName = safeFolderName_(replacePlaceholders_(cfg("Fleet Folder Name") || "[Number] - [Driver] ([Plate])", oldPlaceholders));
        var oldAKey = sKey + "_a" + renameContext.fleetNumber;
        var oldDocumentationAKey = sKey + "_documentation_a" + renameContext.fleetNumber;

        var newPlaceholders = {
          Brand: String(sessData[2]).trim(),
          Shipment: String(sessData[3]).trim(),
          Date: sessDate,
          PIC: String(sessData[4]).trim(),
          Number: renameContext.fleetNumber,
          Driver: renameContext.newDriver,
          Plate: renameContext.newPlate,
          Type: renameContext.newType
        };
        var newFleetName = safeFolderName_(replacePlaceholders_(cfg("Fleet Folder Name") || "[Number] - [Driver] ([Plate])", newPlaceholders));
        if (oldFleetName !== newFleetName) {
          var sessionFolder = getOrCreateFolderCached_(
            getOrCreateFolderCached_(rootFolder, brandName, "f_" + rootId + "_" + brandName),
            sessionName, sKey
          );
          var fleetFolder = null;
          var documentationName = safeFolderName_(cfg("Documentation Folder Name") || "Dokumentasi");
          var documentationFolders = sessionFolder.getFoldersByName(documentationName);
          if (documentationFolders.hasNext()) {
            var documentationFolder = documentationFolders.next();
            var documentationFleetFolders = documentationFolder.getFoldersByName(oldFleetName);
            if (documentationFleetFolders.hasNext()) fleetFolder = documentationFleetFolders.next();
          }
          if (!fleetFolder) {
            var fleetFolders = sessionFolder.getFoldersByName(oldFleetName);
            if (fleetFolders.hasNext()) fleetFolder = fleetFolders.next();
          }
          if (fleetFolder) fleetFolder.setName(newFleetName);
          var cache = CacheService.getScriptCache();
          cache.remove(oldAKey);
          cache.remove(oldAKey + "_dok");
          cache.remove(oldAKey + "_pal");
          cache.remove(oldAKey + "_sj");
          cache.remove(oldDocumentationAKey);
          cache.remove(oldDocumentationAKey + "_dok");
          cache.remove(oldDocumentationAKey + "_pal");
          cache.remove(oldDocumentationAKey + "_sj");
        }
      }
      perfTrace_("updateContainer", traceId, "drive_rename_end");
    } catch (driveErr) {
      driveWarning = "Folder Drive gagal di-rename, tetapi data armada sudah tersimpan.";
      console.warn("updateContainer: rename folder gagal (non-critical): " + driveErr);
      perfTrace_("updateContainer", traceId, "drive_rename_error", { message: driveErr.toString() });
    }
  }

  perfTrace_("updateContainer", traceId, "server_end", { success: true });
  return { success: true, driveWarning: driveWarning };
}

// =========================================================================
// deleteContainer()
// V6 FIX: Drive trash dijalankan SETELAH lock dilepas.
// PATCH: Foto dihapus dengan clearContent, bukan deleteRow → tidak ada index shift.
// =========================================================================
function deleteContainer(containerId, traceId) {
  perfTrace_("deleteContainer", traceId, "server_start", { containerId: containerId });
  var DB = getSheetNames_();
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(DB.FLEET);
  var fleetMap = ensureFleetWorkColumns_(sheet);
  var photoSheet = ss.getSheetByName(DB.PHOTOS);
  var sessionId = null;
  var fleetFolderIdToTrash = "";
  var fileIdsToTrash = [];

  // === LOCK: hanya untuk modifikasi sheet ===
  var lock = LockService.getScriptLock();
  try {
    perfTrace_("deleteContainer", traceId, "lock_wait_start");
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);
    perfTrace_("deleteContainer", traceId, "lock_wait_end");
    perfTrace_("deleteContainer", traceId, "sheet_read_start");
    var data = sheet.getDataRange().getValues();
    perfTrace_("deleteContainer", traceId, "sheet_read_end", { rows: data.length });

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === containerId) {
        if (data[i][8] !== "DELETED" && normalizeFleetWorkStatus_(getFleetCell_(data[i], fleetMap, "Work Status")) === "DONE") {
          return { success: false, message: "Armada sudah di-finish. Klik Batal Finish dulu sebelum menghapus armada." };
        }
        sessionId = data[i][1];
        fleetFolderIdToTrash = extractDriveFolderId_(data[i][7]);
        perfTrace_("deleteContainer", traceId, "sheet_write_start", { row: i + 1 });
        sheet.getRange(i + 1, 9).setValue("DELETED");
        perfTrace_("deleteContainer", traceId, "sheet_write_end");
        break;
      }
    }

    perfTrace_("deleteContainer", traceId, "photo_sheet_read_start");
    var photos  = photoSheet.getDataRange().getValues();
    perfTrace_("deleteContainer", traceId, "photo_sheet_read_end", { rows: photos.length });
    var lastCol = photoSheet.getLastColumn();

    for (var p = 1; p < photos.length; p++) {
      // PATCH: skip baris kosong hasil clearContent sebelumnya
      if (!String(photos[p][0] || "").trim()) continue;

      if (photos[p][1] === containerId) {
        fileIdsToTrash.push(photos[p][0]);
        // PATCH: clearContent — baris tetap, isi dikosongkan (ganti deleteRow)
        perfTrace_("deleteContainer", traceId, "photo_sheet_clear_start", { row: p + 1 });
        photoSheet.getRange(p + 1, 1, 1, lastCol).clearContent();
        perfTrace_("deleteContainer", traceId, "photo_sheet_clear_end");
      }
    }
  } catch (e) {
    perfTrace_("deleteContainer", traceId, "server_error", { message: e.toString() });
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }

  // === DI LUAR LOCK: Drive trash (non-critical, tidak blokir user lain) ===
  perfTrace_("deleteContainer", traceId, "drive_trash_start", { files: fileIdsToTrash.length, folder: !!fleetFolderIdToTrash });
  var driveFailures = [];
  for (var f = 0; f < fileIdsToTrash.length; f++) {
    try {
      DriveApp.getFileById(fileIdsToTrash[f]).setTrashed(true);
    } catch (e) {
      driveFailures.push(fileIdsToTrash[f]);
    }
  }
  var folderTrashSuccess = false;
  var folderTrashError = "";
  if (fleetFolderIdToTrash) {
    try {
      DriveApp.getFolderById(fleetFolderIdToTrash).setTrashed(true);
      folderTrashSuccess = true;
    } catch (folderErr) {
      try {
        Drive.Files.update(
          { trashed: true },
          fleetFolderIdToTrash,
          null,
          { supportsAllDrives: true }
        );
        folderTrashSuccess = true;
      } catch (advancedFolderErr) {
        folderTrashError = advancedFolderErr.toString() + " | " + folderErr.toString();
      }
    }
  }
  if (driveFailures.length > 0) {
    console.warn(
      "deleteContainer: " +
        driveFailures.length +
        " file gagal di-trash: " +
        driveFailures.join(", "),
    );
  }
  if (folderTrashError) console.warn("deleteContainer: folder armada gagal di-trash: " + folderTrashError);
  perfTrace_("deleteContainer", traceId, "drive_trash_end", { failures: driveFailures.length, folderTrashed: folderTrashSuccess, folderError: folderTrashError });

  rebuildAndCacheSync_(sessionId);
  perfTrace_("deleteContainer", traceId, "server_end", { success: true });
  return {
    success: true,
    driveTrashFailures: driveFailures.length,
    folderTrashed: folderTrashSuccess,
    folderTrashError: folderTrashError
  };
}


function pauseFleetWork(containerId, actor, traceId) {
  var actorData = normalizeFleetActorData_(actor, traceId);
  traceId = actorData.traceId || traceId;
  perfTrace_("pauseFleetWork", traceId, "server_start", { containerId: containerId });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);
    var ss = getSpreadsheet_();
    var record = readFleetWorkRecord_(containerId, ss);
    assertSessionIsActive_(record.sessionId, ss);
    var actorName = assertWorkingUserForFleet_(record.sessionId, actorData, ss);
    if (record.status === "READY") return { success: false, code: "FLEET_NOT_STARTED", message: "Klik Mulai Armada dulu." };
    if (record.status === "DONE") return { success: false, message: "Armada sudah di-finish." };
    if (record.status === "UPLOAD_ONLY") {
      return { success: true, data: { containerId: record.containerId, sessionId: record.sessionId, workStatus: "UPLOAD_ONLY", actor: actorName, lifecycle: buildFleetWorkSnapshot_(record.row, record.map) } };
    }
    
    // Berbeda dengan finish, pause TIDAK memerlukan validasi kelengkapan foto
    var now = new Date();
    var nowMs = now.getTime();
    var startedMs = toTimestampMs_(getFleetCell_(record.row, record.map, "Started At"));
    var reopenedMs = toTimestampMs_(getFleetCell_(record.row, record.map, "Reopened At"));
    var segmentStartMs = reopenedMs > 0 ? reopenedMs : startedMs;
    var segmentDuration = segmentStartMs > 0 ? Math.max(0, nowMs - segmentStartMs) : 0;
    var accMs = Number(getFleetCell_(record.row, record.map, "Accumulated Work Ms")) || 0;
    
    var snapshot = writeFleetWorkSnapshot_(record, "UPLOAD_ONLY", {
      completedAt: now,
      completedBy: actorName,
      accumulatedWorkMs: accMs + segmentDuration
    });
    rebuildAndCacheSync_(record.sessionId);
    return { success: true, data: { containerId: record.containerId, sessionId: record.sessionId, workStatus: "UPLOAD_ONLY", actor: actorName, lifecycle: snapshot } };
  } catch (e) {
    perfTrace_("pauseFleetWork", traceId, "server_error", { message: e.toString() });
    return { success: false, message: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (releaseErr) {}
  }
}
