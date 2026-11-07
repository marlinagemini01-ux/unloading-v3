// =========================================================================
// SessionService.gs — V6 PERFORMANCE EDITION
// =========================================================================
// OPTIMASI:
//   - Semua setValue() → batch setValues() (60-70% lebih cepat)
//   - addContainer() → lazy folder (skip Drive API = 2-5 detik lebih cepat)
//   - markUserDone() → single-loop (baca sheet 1x bukan 2x)
//   - startSession() duplicate check → lightweight (skip photo count)
//   - deleteContainer() → batch operations
//   - finishSession() → return rekap warning ke frontend
// =========================================================================

// =========================================================================
// startSession()
// V6 FIX: buildFolderPath_() (2-5 detik Drive API) dijalankan SEBELUM lock.
// Sebelumnya: lock ditahan 4-8 detik (Drive API + baca sheet + tulis).
// Sesudah: lock ~500ms (cek duplikat + appendRow saja).
// =========================================================================
function startSession(data) {
  data = data || {};
  var traceId = data.traceId;
  perfTrace_("createSession", traceId, "server_start");
  var brand = sanitize_(data.brand);
  var shipment = sanitize_(data.shipment);
  var pic = titleCaseText_(data.pic);
  var date = sanitize_(data.date);
  var createdBy = sanitize_(data.createdBy || data.pic);
  var forwarder = titleCaseText_(data.forwarder || "");

  try {
    validateRequired_({
      Brand: brand,
      Shipment: shipment,
      PIC: pic,
      Tanggal: date,
    });
  } catch (e) {
    return { success: false, message: e.toString() };
  }

  var DB = getSheetNames_();
  var ss = getSpreadsheet_();
  var tz = cfg("Timezone") || "Asia/Jakarta";

  // === Cek duplikat SEBELUM lock (read-only, tidak perlu lock) ===
  var sessionSheet = ss.getSheetByName(DB.SESSIONS);
  ensureSessionColumns_(sessionSheet);
  perfTrace_("createSession", traceId, "duplicate_read_start");
  var existing = sessionSheet.getDataRange().getValues();
  perfTrace_("createSession", traceId, "duplicate_read_end", { rows: existing.length });
  for (var i = 1; i < existing.length; i++) {
    if (
      existing[i][5] === "ACTIVE" &&
      String(existing[i][2]).trim().toLowerCase() === brand.toLowerCase() &&
      String(existing[i][3]).trim().toLowerCase() === shipment.toLowerCase()
    ) {
      var rawDate = existing[i][1];
      var existDate =
        rawDate instanceof Date
          ? Utilities.formatDate(rawDate, tz, "yyyy-MM-dd")
          : String(rawDate).trim();

      if (existDate === date) {
        var existingId = existing[i][0];
        var existingCreatedBy = String(
          existing[i][8] || existing[i][4] || "",
        ).trim();

        var fleetSheet = ss.getSheetByName(DB.FLEET);
        perfTrace_("createSession", traceId, "duplicate_fleet_read_start");
        var fRows = fleetSheet ? fleetSheet.getDataRange().getValues() : [];
        perfTrace_("createSession", traceId, "duplicate_fleet_read_end", { rows: fRows.length });
        var eFleetCount = 0;
        for (var f = 1; f < fRows.length; f++) {
          if (fRows[f][1] === existingId && fRows[f][8] !== "DELETED")
            eFleetCount++;
        }

        var creatorFullName = existingCreatedBy;
        var accSheet = ss.getSheetByName(DB.ACCOUNTS);
        if (accSheet) {
          perfTrace_("createSession", traceId, "duplicate_account_read_start");
          var accRows = accSheet.getDataRange().getValues();
          perfTrace_("createSession", traceId, "duplicate_account_read_end", { rows: accRows.length });
          for (var a = 1; a < accRows.length; a++) {
            if (
              String(accRows[a][0]).trim().toLowerCase() ===
              existingCreatedBy.toLowerCase()
            ) {
              creatorFullName = String(accRows[a][2] || accRows[a][0]).trim();
              break;
            }
          }
        }

        return {
          success: false,
          duplicate: true,
          existingSession: {
            sessionId: existingId,
            brand: String(existing[i][2]).trim(),
            shipment: String(existing[i][3]).trim(),
            date: existDate,
            createdBy: creatorFullName,
            createdByUsername: existingCreatedBy,
            fleetCount: eFleetCount,
            photoCount: 0,
          },
        };
      }
    }
  }

  // === Buat folder Drive SEBELUM lock (2-5 detik, tidak perlu lock) ===
  var folders;
  try {
    perfTrace_("createSession", traceId, "drive_folder_start");
    folders = buildFolderPath_({
      brand: brand,
      shipment: shipment,
      date: date,
      pic: pic,
    });
    perfTrace_("createSession", traceId, "drive_folder_end");
  } catch (e) {
    return {
      success: false,
      message: "Gagal buat folder Drive: " + e.toString(),
    };
  }

  var sessionId = generateUniqueId_("Session ID Prefix");

  // === LOCK: appendRow + DOUBLE-CHECK duplikat di dalam lock ===
  var lock = LockService.getScriptLock();
  try {
    perfTrace_("createSession", traceId, "lock_wait_start");
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);
    perfTrace_("createSession", traceId, "lock_wait_end");

    // === DOUBLE-CHECK: Cek duplikat LAGI di dalam lock ===
    // Kenapa? Karena antara cek duplikat pertama (line 41) dan lock ini,
    // user lain bisa saja sudah appendRow sesi yang sama.
    perfTrace_("createSession", traceId, "double_check_read_start");
    var freshData = sessionSheet.getDataRange().getValues();
    perfTrace_("createSession", traceId, "double_check_read_end", { rows: freshData.length });
    for (var dc = 1; dc < freshData.length; dc++) {
      if (
        freshData[dc][5] === "ACTIVE" &&
        String(freshData[dc][2]).trim().toLowerCase() === brand.toLowerCase() &&
        String(freshData[dc][3]).trim().toLowerCase() === shipment.toLowerCase()
      ) {
        var dcRaw = freshData[dc][1];
        var dcDate = dcRaw instanceof Date
          ? Utilities.formatDate(dcRaw, tz, "yyyy-MM-dd")
          : String(dcRaw).trim();
        if (dcDate === date) {
          // Duplikat terdeteksi di dalam lock — TOLAK
          // Folder Drive sudah terlanjur dibuat, tapi tidak ada sesi duplikat di DB
          return {
            success: false,
            duplicate: true,
            existingSession: {
              sessionId: freshData[dc][0],
              brand: String(freshData[dc][2]).trim(),
              shipment: String(freshData[dc][3]).trim(),
              date: dcDate,
              createdBy: String(freshData[dc][8] || freshData[dc][4] || "").trim(),
              createdByUsername: String(freshData[dc][8] || "").trim(),
              fleetCount: 0,
              photoCount: 0,
            },
          };
        }
      }
    }

    perfTrace_("createSession", traceId, "session_write_start");
    sessionSheet.appendRow([
      sessionId,
      date,
      brand,
      shipment,
      pic,
      "ACTIVE",
      new Date(),
      folders.sessionFolder.getUrl(),
      createdBy,
      forwarder,
      "",
      "",
      "",
    ]);
    perfTrace_("createSession", traceId, "session_write_end", { sessionId: sessionId });

    if (cfgBool("Auto Join On Create")) {
      var userSheet = ss.getSheetByName(DB.USERS);
      if (userSheet) {
        perfTrace_("createSession", traceId, "user_write_start");
        userSheet.appendRow([
          generateUniqueId_("User ID Prefix"),
          sessionId,
          data.fullName || createdBy,
          "WORKING",
          new Date(),
          new Date(),
        ]);
        perfTrace_("createSession", traceId, "user_write_end");
      }
    }
  } catch (e) {
    perfTrace_("createSession", traceId, "server_error", { message: e.toString() });
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }

  rebuildAndCacheSync_(sessionId);
  perfTrace_("createSession", traceId, "server_end", { success: true, sessionId: sessionId });
  return {
    success: true,
    data: { sessionId: sessionId, folderUrl: folders.sessionFolder.getUrl() },
  };
}

// =========================================================================
// updateSession()
// Hanya pembuat sesi (createdBy) yang boleh edit brand/shipment/tanggal.
// Hanya sesi berstatus ACTIVE yang bisa diedit.
// Pola lock: DB write di dalam lock, Drive rename SETELAH lock dilepas.
// =========================================================================
function updateSession(sessionId, updates, requestedBy) {
  var newBrand = updates.brand ? sanitize_(updates.brand) : "";
  var newShipment = updates.shipment ? sanitize_(updates.shipment) : "";
  var newDate = updates.date ? sanitize_(updates.date) : "";
  var newForwarder = updates && Object.prototype.hasOwnProperty.call(updates, "forwarder")
    ? titleCaseText_(updates.forwarder || "")
    : null;

  var DB = getSheetNames_();
  var ss = getSpreadsheet_();
  var tz = cfg("Timezone") || "Asia/Jakarta";
  var df = cfg("Date Format") || "dd MMMM yyyy";
  var sSession = ss.getSheetByName(DB.SESSIONS);
  ensureSessionColumns_(sSession);
  var tz = cfg("Timezone") || "Asia/Jakarta";

  // Baca data SEBELUM lock (read-only)
  var sessionRows = sSession.getDataRange().getValues();
  var sessionData = null,
    rowIndex = -1;
  for (var i = 1; i < sessionRows.length; i++) {
    if (sessionRows[i][0] === sessionId) {
      sessionData = sessionRows[i];
      rowIndex = i + 1;
      break;
    }
  }
  if (!sessionData) return { success: false, message: "Sesi tidak ditemukan." };
  if (sessionData[5] !== "ACTIVE")
    return {
      success: false,
      message: "Sesi sudah COMPLETED, tidak bisa diedit.",
    };

  var admin = String(sessionData[8] || "")
    .trim()
    .toLowerCase();
  if (
    admin &&
    admin !==
      String(requestedBy || "")
        .trim()
        .toLowerCase()
  ) {
    return {
      success: false,
      message:
        "Hanya pembuat sesi (" + sessionData[8] + ") yang bisa mengedit.",
    };
  }

  // Resolve nilai final (gunakan nilai lama jika tidak diubah)
  var finalBrand = newBrand || String(sessionData[2]).trim();
  var finalShipment = newShipment || String(sessionData[3]).trim();
  var rawDate = sessionData[1];
  var currentDate =
    rawDate instanceof Date
      ? Utilities.formatDate(rawDate, tz, "yyyy-MM-dd")
      : String(rawDate).trim();
  var finalDate = newDate || currentDate;
  var currentForwarder = String(sessionData[9] || "").trim();
  var finalForwarder = newForwarder !== null ? newForwarder : currentForwarder;

  var brandChanged =
    finalBrand.toLowerCase() !== String(sessionData[2]).trim().toLowerCase();
  var shipmentChanged =
    finalShipment.toLowerCase() !== String(sessionData[3]).trim().toLowerCase();
  var dateChanged = finalDate !== currentDate;
  var forwarderChanged = finalForwarder !== currentForwarder;
  var sessionMetaChanged = brandChanged || shipmentChanged || dateChanged;

  if (!sessionMetaChanged && !forwarderChanged) {
    return { success: false, message: "Tidak ada perubahan." };
  }

  if (sessionMetaChanged) {
    // Cek duplikat: brand+shipment+tanggal harus unik di antara sesi ACTIVE lain
    for (var j = 1; j < sessionRows.length; j++) {
      if (sessionRows[j][0] === sessionId) continue;
      if (sessionRows[j][5] !== "ACTIVE") continue;
      var sB = String(sessionRows[j][2]).trim().toLowerCase();
      var sS = String(sessionRows[j][3]).trim().toLowerCase();
      var sRaw = sessionRows[j][1];
      var sD =
        sRaw instanceof Date
          ? Utilities.formatDate(sRaw, tz, "yyyy-MM-dd")
          : String(sRaw).trim();
      if (
        sB === finalBrand.toLowerCase() &&
        sS === finalShipment.toLowerCase() &&
        sD === finalDate
      ) {
        return {
          success: false,
          message:
            "Sesi " +
            finalBrand +
            " / " +
            finalShipment +
            " pada tanggal ini sudah ada.",
        };
      }
    }
  }

  // === LOCK: hanya untuk write sheet ===
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);
    var rowData = sessionData.slice();
    while (rowData.length < 13) rowData.push("");
    rowData[1] = finalDate;
    rowData[2] = finalBrand;
    rowData[3] = finalShipment;
    rowData[9] = finalForwarder;
    sSession.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } catch (e) {
    perfTrace_("createContainer", traceId, "server_error", { message: e.toString() });
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock(); // ← Lepas SEGERA setelah write
  }

  rebuildAndCacheSync_(sessionId);

  // === LUAR LOCK: Rename/pindah folder Drive (non-critical, tidak blokir user lain) ===
  var driveWarning = null;
  var folderUrl = String(sessionData[7] || "");
  var m = folderUrl.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  var sessionFolderId = m ? m[1] : null;

  if (sessionFolderId && sessionMetaChanged) {
    try {
      var placeholders = {
        Brand: finalBrand,
        Shipment: finalShipment,
        Date: finalDate,
      };
      var newSessionName = safeFolderName_(
        replacePlaceholders_(
          cfg("Session Folder Name") || "[Shipment] - [Date]",
          placeholders,
        ),
      );

      if (brandChanged) {
        // Brand berubah: pindah session folder ke brand folder baru + rename
        var oldBrandName = safeFolderName_(
          replacePlaceholders_(
            cfg("Brand Folder Name") || "[Brand]",
            { Brand: String(sessionData[2]).trim() },
          ),
        );
        var newBrandName = safeFolderName_(
          replacePlaceholders_(
            cfg("Brand Folder Name") || "[Brand]",
            placeholders,
          ),
        );
        var _fixRootFolder = getRootFolder_();
        var rootId = _fixRootFolder.getId();
        var newBrandFolder = getOrCreateFolderCached_(
          _fixRootFolder,
          newBrandName,
          "f_" + rootId + "_" + newBrandName,
        );
        var sessionFolder = DriveApp.getFolderById(sessionFolderId);

        // Ambil parent lama SEBELUM dipindah
        var oldParents = sessionFolder.getParents();
        var oldBrandFolder = oldParents.hasNext() ? oldParents.next() : null;

        // Step 1: Rename session folder
        sessionFolder.setName(newSessionName);

        // Step 2: Tambahkan ke brand folder baru
        newBrandFolder.addFolder(sessionFolder);

        // Step 3: Hapus dari brand folder lama (jika berbeda)
        if (
          oldBrandFolder &&
          oldBrandFolder.getId() !== newBrandFolder.getId()
        ) {
          oldBrandFolder.removeFolder(sessionFolder);

          // Cleanup: trash folder brand lama jika sudah kosong
          try {
            var stillHasContent =
              oldBrandFolder.getFolders().hasNext() ||
              oldBrandFolder.getFiles().hasNext();
            if (!stillHasContent) {
              oldBrandFolder.setTrashed(true);
            }
          } catch (cleanupErr) {
            console.warn(
              "Cleanup folder brand lama gagal (non-critical):",
              cleanupErr,
            );
          }
        }

        // Step 4: Invalidasi cache folder LAMA agar tidak stale
        // Tanpa ini, cache masih menyimpan ID folder brand lama
        // yang menyebabkan data lama "tidak terhapus" dari sisi sistem
        try {
          var cache = CacheService.getScriptCache();
          var oldBrandCacheKey = "f_" + rootId + "_" + oldBrandName;
          var oldSessionCacheKey = "f_" + rootId + "_" + oldBrandName + "_" + safeFolderName_(
            replacePlaceholders_(
              cfg("Session Folder Name") || "[Shipment] - [Date]",
              { Brand: String(sessionData[2]).trim(), Shipment: String(sessionData[3]).trim(), Date: currentDate },
            ),
          );
          cache.remove(oldBrandCacheKey);
          cache.remove(oldSessionCacheKey);
          // Juga hapus subfolder cache (armada + dok/palet)
          var fleetData = ss.getSheetByName(DB.FLEET).getDataRange().getValues();
          for (var fc = 1; fc < fleetData.length; fc++) {
            if (fleetData[fc][1] === sessionId && fleetData[fc][8] !== "DELETED") {
              var fleetNo = fleetData[fc][2];
              var oldArmadaKey = oldSessionCacheKey + "_a" + fleetNo;
              var oldDocumentationArmadaKey = oldSessionCacheKey + "_documentation_a" + fleetNo;
              cache.remove(oldArmadaKey);
              cache.remove(oldArmadaKey + "_dok");
              cache.remove(oldArmadaKey + "_pal");
              cache.remove(oldArmadaKey + "_sj");
              cache.remove(oldDocumentationArmadaKey);
              cache.remove(oldDocumentationArmadaKey + "_dok");
              cache.remove(oldDocumentationArmadaKey + "_pal");
              cache.remove(oldDocumentationArmadaKey + "_sj");
            }
          }
        } catch (cacheErr) {
          console.warn("Cache invalidation setelah rename gagal (non-critical):", cacheErr);
        }
      } else {
        // Hanya shipment/tanggal berubah: cukup rename
        DriveApp.getFolderById(sessionFolderId).setName(newSessionName);
      }
    } catch (e) {
      driveWarning =
        "Data berhasil diperbarui, tapi rename folder Drive gagal: " +
        e.toString();
      console.warn("updateSession Drive rename gagal:", e);
    }
  }

  return {
    success: true,
    data: {
      brand: finalBrand,
      shipment: finalShipment,
      date: finalDate,
      forwarder: finalForwarder,
    },
    driveWarning: driveWarning,
  };
}

function setLossPreventionInfo(sessionId, lpName, requestedBy, signatureMode, signatureBase64) {
  var finalName = titleCaseText_(lpName || "");
  var actor = sanitize_(requestedBy || "");
  var mode = String(signatureMode || "MANUAL").toUpperCase() === "DIGITAL" ? "DIGITAL" : "MANUAL";
  var tz = cfg("Timezone") || "Asia/Jakarta";
  var df = cfg("Date Format") || "dd MMMM yyyy";
  try {
    validateRequired_({
      "Nama Loss Prevention": finalName
    });
  } catch (e) {
    return { success: false, message: e.toString() };
  }

  var DB = getSheetNames_();
  var ss = getSpreadsheet_();
  var sSession = ss.getSheetByName(DB.SESSIONS);
  ensureSessionColumns_(sSession);
  var rows = sSession.getDataRange().getValues();
  var map = getHandoverHeaderMap_(sSession);
  var rowIndex = -1;
  var rowData = null;

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === sessionId) {
      rowIndex = i + 1;
      rowData = rows[i].slice();
      break;
    }
  }

  if (!rowData) return { success: false, message: "Sesi tidak ditemukan." };

  while (rowData.length < sSession.getLastColumn()) rowData.push("");
  var existingSignatureFileId = map["Loss Prevention Signature File ID"] !== undefined
    ? String(rowData[map["Loss Prevention Signature File ID"]] || "").trim()
    : "";
  var signatureFile = null;
  var signedAt = new Date();
  if (mode === "DIGITAL") {
    if (!signatureBase64) return { success: false, message: "Tanda tangan digital LP wajib diisi." };
    try {
      signatureFile = createOrUpdateLossPreventionSignatureFile_(sessionId, finalName, signatureBase64, existingSignatureFileId);
    } catch (sigErr) {
      return { success: false, message: "Gagal menyimpan tanda tangan LP: " + sigErr.toString() };
    }
  }

  rowData[10] = finalName;
  rowData[11] = signedAt;
  rowData[12] = actor || rowData[12] || "";
  if (map["Loss Prevention Signature Mode"] !== undefined) rowData[map["Loss Prevention Signature Mode"]] = mode;
  if (map["Loss Prevention Signature File ID"] !== undefined) rowData[map["Loss Prevention Signature File ID"]] = signatureFile ? signatureFile.fileId : "";
  if (map["Loss Prevention Signature File URL"] !== undefined) rowData[map["Loss Prevention Signature File URL"]] = signatureFile ? signatureFile.fileUrl : "";
  if (map["Loss Prevention Signature Signed At"] !== undefined) rowData[map["Loss Prevention Signature Signed At"]] = mode === "DIGITAL" ? signedAt : "";

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);
    sSession.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }

  rebuildAndCacheSync_(sessionId);
  return {
    success: true,
    data: {
      lossPreventionName: rowData[10],
      lossPreventionFilledAt: rowData[11] instanceof Date ? Utilities.formatDate(rowData[11], tz, df) : String(rowData[11] || "").trim(),
      lossPreventionFilledBy: rowData[12],
      lossPreventionSignatureMode: mode,
      lossPreventionSignatureFileId: signatureFile ? signatureFile.fileId : "",
      lossPreventionSignatureFileUrl: signatureFile ? signatureFile.fileUrl : "",
      lossPreventionSignatureSignedAt: mode === "DIGITAL" ? Utilities.formatDate(signedAt, tz, df) : ""
    }
  };
}

function getLossPreventionInfo(sessionId) {
  if (!sessionId) return { success: false, message: "Session ID diperlukan." };
  var DB = getSheetNames_();
  var ss = getSpreadsheet_();
  var tz = cfg("Timezone") || "Asia/Jakarta";
  var df = cfg("Date Format") || "dd MMMM yyyy";
  var sSession = ss.getSheetByName(DB.SESSIONS);
  ensureSessionColumns_(sSession);
  var rows = sSession.getDataRange().getValues();
  var map = getHandoverHeaderMap_(sSession);
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] === sessionId) {
      var rawDate = rows[i][11];
      var rawSignedAt = map["Loss Prevention Signature Signed At"] !== undefined ? rows[i][map["Loss Prevention Signature Signed At"]] : "";
      return {
        success: true,
        data: {
          lossPreventionName: String(rows[i][10] || "").trim(),
          lossPreventionFilledAt: rawDate instanceof Date ? Utilities.formatDate(rawDate, tz, df) : String(rawDate || "").trim(),
          lossPreventionFilledBy: String(rows[i][12] || "").trim(),
          lossPreventionSignatureMode: map["Loss Prevention Signature Mode"] !== undefined ? String(rows[i][map["Loss Prevention Signature Mode"]] || "").trim() : "",
          lossPreventionSignatureFileId: map["Loss Prevention Signature File ID"] !== undefined ? String(rows[i][map["Loss Prevention Signature File ID"]] || "").trim() : "",
          lossPreventionSignatureFileUrl: map["Loss Prevention Signature File URL"] !== undefined ? String(rows[i][map["Loss Prevention Signature File URL"]] || "").trim() : "",
          lossPreventionSignatureSignedAt: rawSignedAt instanceof Date ? Utilities.formatDate(rawSignedAt, tz, df) : String(rawSignedAt || "").trim()
        }
      };
    }
  }
  return { success: false, message: "Sesi tidak ditemukan." };
}

// =========================================================================
// finishSession()
// V6 FIX: Lock HANYA untuk setValue("COMPLETED") ~300ms.
// generateRekap_ + kirim WA dijalankan SETELAH lock dilepas.
// Sebelumnya: lock ditahan 10-60 detik (3x baca sheet + rekap + WA API call).
// Sesudah: lock ~500ms — user lain tidak perlu antri sama sekali.
// =========================================================================
function getSessionArtifactHeaderMap_(sheet) {
  ensureSessionColumns_(sheet);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    map[String(headers[i] || "").trim()] = i;
  }
  return map;
}

function writeSessionArtifactStatus_(sheet, rowIndex, values) {
  if (!sheet || !rowIndex || rowIndex < 2) return;
  var map = getSessionArtifactHeaderMap_(sheet);
  var payload = values || {};
  Object.keys(payload).forEach(function(name) {
    if (map[name] === undefined) return;
    sheet.getRange(rowIndex, map[name] + 1).setValue(payload[name]);
  });
  if (map["Artifact Updated At"] !== undefined) {
    sheet.getRange(rowIndex, map["Artifact Updated At"] + 1).setValue(new Date());
  }
}

function finishSession(sessionId, requestedBy, traceId, lpSignaturePayload) {
  perfTrace_("finishSession", traceId, "server_start", { sessionId: sessionId });
  var DB = getSheetNames_();
  var ss = getSpreadsheet_();
  var sSession = ss.getSheetByName(DB.SESSIONS);
  var sFleet = ss.getSheetByName(DB.FLEET);
  var userSheet = ss.getSheetByName(DB.USERS);
  ensureSessionColumns_(sSession);

  // Baca data validasi SEBELUM lock (read-only, tidak perlu lock)
  var sessionData = null,
    rowIndex = -1;
  perfTrace_("finishSession", traceId, "session_sheet_read_start");
  var sessionRows = sSession.getDataRange().getValues();
  perfTrace_("finishSession", traceId, "session_sheet_read_end", { rows: sessionRows.length });
  for (var i = 1; i < sessionRows.length; i++) {
    if (sessionRows[i][0] === sessionId) {
      sessionData = sessionRows[i];
      rowIndex = i + 1;
      break;
    }
  }
  if (!sessionData) return { success: false, message: "Sesi tidak ditemukan." };

  if (lpSignaturePayload && String(lpSignaturePayload.signatureMode || "").toUpperCase() === "DIGITAL") {
    perfTrace_("finishSession", traceId, "lp_signature_save_start");
    var lpSignatureRes = setLossPreventionInfo(
      sessionId,
      lpSignaturePayload.lpName || sessionData[10] || "",
      requestedBy,
      "DIGITAL",
      lpSignaturePayload.signatureDataUrl || ""
    );
    if (!lpSignatureRes || !lpSignatureRes.success) {
      perfTrace_("finishSession", traceId, "lp_signature_save_error", { message: lpSignatureRes && lpSignatureRes.message ? lpSignatureRes.message : "Unknown" });
      return {
        success: false,
        message: lpSignatureRes && lpSignatureRes.message ? lpSignatureRes.message : "Tanda tangan LP gagal disimpan."
      };
    }
    if (!lpSignatureRes.data || !lpSignatureRes.data.lossPreventionSignatureFileId) {
      perfTrace_("finishSession", traceId, "lp_signature_missing_file_id");
      return {
        success: false,
        message: "Tanda tangan LP belum menghasilkan file Drive, PDF handover belum dibuat."
      };
    }
    perfTrace_("finishSession", traceId, "lp_signature_save_end", { fileId: lpSignatureRes.data.lossPreventionSignatureFileId });

    perfTrace_("finishSession", traceId, "session_sheet_reread_after_signature");
    sessionRows = sSession.getDataRange().getValues();
    for (var sr = 1; sr < sessionRows.length; sr++) {
      if (sessionRows[sr][0] === sessionId) {
        sessionData = sessionRows[sr];
        rowIndex = sr + 1;
        break;
      }
    }
  }

  var admin = String(sessionData[8] || "")
    .trim()
    .toLowerCase();
  if (
    requestedBy &&
    admin &&
    admin !== String(requestedBy).trim().toLowerCase()
  ) {
    return {
      success: false,
      message: "Hanya admin sesi (" + sessionData[8] + ") yang bisa finish.",
    };
  }

  if (cfgBool("Require All Done To Finish") && userSheet) {
    perfTrace_("finishSession", traceId, "users_sheet_read_start");
    var uRows = userSheet.getDataRange().getValues();
    perfTrace_("finishSession", traceId, "users_sheet_read_end", { rows: uRows.length });
    var working = [];
    for (var k = 1; k < uRows.length; k++) {
      if (uRows[k][1] === sessionId && uRows[k][3] === "WORKING")
        working.push(uRows[k][2]);
    }
    if (working.length > 0)
      return {
        success: false,
        message: "Masih ada yang belum selesai: " + working.join(", "),
      };
  }

  var fleetCount = 0;
  var activeFleets = [];
  perfTrace_("finishSession", traceId, "fleet_sheet_read_start");
  var fleetMap = ensureFleetWorkColumns_(sFleet);
  var fleetRows = sFleet.getDataRange().getValues();
  perfTrace_("finishSession", traceId, "fleet_sheet_read_end", { rows: fleetRows.length });
  for (var j = 1; j < fleetRows.length; j++) {
    if (fleetRows[j][1] === sessionId && fleetRows[j][8] !== "DELETED") {
      var fleetWork = buildFleetWorkSnapshot_(fleetRows[j], fleetMap);
      fleetCount++;
      activeFleets.push({
        id: fleetRows[j][0],
        number: fleetRows[j][2],
        driver: fleetRows[j][4],
        plate: fleetRows[j][5],
        workStatus: fleetWork.workStatus,
        startedAtMs: fleetWork.startedAtMs,
        completedAtMs: fleetWork.completedAtMs,
      });
    }
  }

  var unfinishedFleetWork = [];
  activeFleets.forEach(function (f) {
    if (f.workStatus !== "DONE" || !f.startedAtMs || !f.completedAtMs) {
      unfinishedFleetWork.push(f);
    }
  });
  if (unfinishedFleetWork.length > 0) {
    return {
      success: false,
      code: "FLEET_WORK_NOT_DONE",
      message: "Masih ada armada yang belum di-finish.",
      pendingFleets: unfinishedFleetWork,
    };
  }

  var photoTypeStats = {};
  var paletStats = {};
  activeFleets.forEach(function (f) {
    photoTypeStats[f.id] = { dokumen: 0, palet: 0, suratjalan: 0 };
    paletStats[f.id] = { valid: 0, incomplete: 0 };
  });
  var sPhotos = ss.getSheetByName(DB.PHOTOS);
  if (sPhotos && activeFleets.length > 0) {
    perfTrace_("finishSession", traceId, "photo_sheet_read_start");
    var photoRows = sPhotos.getDataRange().getValues();
    perfTrace_("finishSession", traceId, "photo_sheet_read_end", { rows: photoRows.length });
    var seenFinishPhotoIds = {};
    for (var p = 1; p < photoRows.length; p++) {
      var finishPhotoId = String(photoRows[p][0] || "").trim();
      if (!finishPhotoId) continue;
      if (seenFinishPhotoIds[finishPhotoId]) continue;
      seenFinishPhotoIds[finishPhotoId] = true;
      var cid = photoRows[p][1];
      if (!paletStats[cid]) continue;
      var normalizedType = normalizePhotoType_(photoRows[p][2]);
      if (photoTypeStats[cid] && photoTypeStats[cid][normalizedType] !== undefined) {
        photoTypeStats[cid][normalizedType]++;
      }
      if (normalizedType !== "palet") continue;
      var paletName = String(photoRows[p][6] || "").trim();
      var paletKoli = Number(photoRows[p][7]) || 0;
      if (paletName && paletKoli > 0) paletStats[cid].valid++;
      else paletStats[cid].incomplete++;
    }
  }
  var missingPhotoTypeContainers = [];
  activeFleets.forEach(function (f) {
    var typeStat = photoTypeStats[f.id] || { dokumen: 0, palet: 0, suratjalan: 0 };
    var missingTypes = [];
    if (typeStat.dokumen === 0) missingTypes.push(cfg("Doc Subfolder Name") || "Unloading");
    if (typeStat.palet === 0) missingTypes.push(cfg("Palet Subfolder Name") || "Palet");
    if (typeStat.suratjalan === 0) missingTypes.push(cfg("Surat Subfolder Name") || "Surat Jalan");
    if (missingTypes.length > 0) {
      missingPhotoTypeContainers.push({
        id: f.id,
        number: f.number,
        driver: f.driver,
        plate: f.plate,
        missingTypes: missingTypes,
      });
    }
  });
  if (missingPhotoTypeContainers.length > 0) {
    return {
      success: false,
      code: "PHOTO_TYPES_REQUIRED",
      message: "Setiap armada wajib memiliki minimal 1 foto Unloading, Palet, dan Surat Jalan.",
      missingPhotoTypeContainers: missingPhotoTypeContainers,
    };
  }
  var missingPaletContainers = [];
  var incompletePaletContainers = [];
  activeFleets.forEach(function (f) {
    var st = paletStats[f.id] || { valid: 0, incomplete: 0 };
    if (st.valid === 0) missingPaletContainers.push(f);
    else if (st.incomplete > 0) incompletePaletContainers.push(f);
  });
  if (missingPaletContainers.length > 0) {
    return {
      success: false,
      code: "PALET_REQUIRED",
      message: "Ada armada yang belum punya foto Palet valid.",
      missingPaletContainers: missingPaletContainers,
    };
  }
  if (incompletePaletContainers.length > 0) {
    return {
      success: false,
      code: "PALET_INCOMPLETE",
      message: "Ada foto Palet yang belum memiliki Nama Palet dan Jumlah Koli.",
      incompletePaletContainers: incompletePaletContainers,
    };
  }

  if (typeof ensureSjQtyReadyForHandover_ === "function" || typeof validateSjQtyMatching === "function") {
    var sjReady = typeof ensureSjQtyReadyForHandover_ === "function"
      ? ensureSjQtyReadyForHandover_(sessionId, traceId)
      : (function() {
          var sjValidation = validateSjQtyMatching(sessionId);
          return sjValidation && sjValidation.success && sjValidation.data && sjValidation.data.canHandover
            ? { success: true, data: sjValidation.data }
            : { success: false, message: "SJ Qty belum valid.", data: sjValidation && sjValidation.data ? sjValidation.data : null };
        })();
    if (!sjReady || !sjReady.success) {
      return {
        success: false,
        code: "SJ_QTY_INVALID",
        message: sjReady && sjReady.message ? sjReady.message : "SJ Qty belum valid.",
        sjQty: sjReady && (sjReady.sjQty || sjReady.data) ? (sjReady.sjQty || sjReady.data) : null,
      };
    }
  }

  var tz = cfg("Timezone") || "Asia/Jakarta";
  var df = cfg("Date Format") || "dd MMMM yyyy";
  var tgl =
    sessionData[1] instanceof Date
      ? Utilities.formatDate(sessionData[1], tz, df)
      : sessionData[1];

  // === LOCK: HANYA untuk ubah status menjadi COMPLETED ===
  var lock = LockService.getScriptLock();
  try {
    perfTrace_("finishSession", traceId, "lock_wait_start");
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);
    perfTrace_("finishSession", traceId, "lock_wait_end");
    perfTrace_("finishSession", traceId, "sheet_write_start", { row: rowIndex });
    sSession.getRange(rowIndex, 6).setValue("COMPLETED");
    writeSessionArtifactStatus_(sSession, rowIndex, {
      "Rekap Status": cfgBool("Rekap Enabled") ? "PENDING" : "SKIPPED",
      "Rekap Error": "",
      "Handover PDF Status": "PENDING",
      "Handover PDF Error": "",
      "WA Status": cfgBool("WA Report Enabled") ? "PENDING" : "SKIPPED",
      "WA Error": ""
    });
    perfTrace_("finishSession", traceId, "sheet_write_end");
  } catch (e) {
    perfTrace_("finishSession", traceId, "server_error", { message: e.toString() });
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock(); // ← Lepas lock SEGERA
  }

  // === DI LUAR LOCK: operasi berat (tidak blokir user lain) ===
  rebuildAndCacheSync_(sessionId);
  perfTrace_("finishSession", traceId, "sync_koli_start");
  syncFleetKoliFromPhotos_(sessionId);
  perfTrace_("finishSession", traceId, "sync_koli_end");

  // PHASE 1 FIX: Masukkan tugas berat ke antrean background untuk diproses secara asynchronous
  try {
    enqueueFinishSessionTask(sessionId, requestedBy, traceId);
    console.log("Tugas finishSession dimasukkan ke background queue untuk sesi: " + sessionId);
  } catch (e) {
    console.warn("Gagal memasukkan tugas ke queue:", e);
  }

  perfTrace_("finishSession", traceId, "server_end", { success: true, backgroundQueued: true });
  return {
    success: true,
    message: "Sesi berhasil diselesaikan. Pembuatan PDF, Rekap, dan Notifikasi WA sedang diproses di latar belakang.",
    waSuccess: true, // Flagged true so frontend shows success instantly
    rekapSuccess: true,
    handoverPdfSuccess: true,
    handoverPdf: null,
    artifactStatus: {
      rekapStatus: "PROCESSING",
      handoverPdfStatus: "PROCESSING",
      waStatus: "PROCESSING"
    }
  };
}

// =========================================================================
// reopenSession() — Buka kembali sesi COMPLETED menjadi ACTIVE
// Use case: Admin salah finish, atau ada armada yang belum selesai
// Guard: hanya admin sesi (createdBy) yang bisa reopen
// =========================================================================
function reopenSession(sessionId, requestedBy) {
  var DB = getSheetNames_();
  var ss = getSpreadsheet_();
  var sSession = ss.getSheetByName(DB.SESSIONS);

  // Cari sesi
  var sessionRows = sSession.getDataRange().getValues();
  var rowIndex = -1, sessionData = null;
  for (var i = 1; i < sessionRows.length; i++) {
    if (sessionRows[i][0] === sessionId) {
      sessionData = sessionRows[i];
      rowIndex = i + 1;
      break;
    }
  }

  if (!sessionData) return { success: false, message: "Sesi tidak ditemukan." };
  if (sessionData[5] !== "COMPLETED") {
    return { success: false, message: "Sesi ini masih ACTIVE, tidak perlu di-reopen." };
  }

  // Guard: hanya admin sesi
  var admin = String(sessionData[8] || "").trim().toLowerCase();
  if (admin && admin !== String(requestedBy || "").trim().toLowerCase()) {
    return {
      success: false,
      message: "Hanya admin sesi (" + sessionData[8] + ") yang bisa membuka kembali."
    };
  }

  // === LOCK: Ubah status kembali ke ACTIVE ===
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);
    sSession.getRange(rowIndex, 6).setValue("ACTIVE");
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }

  // Kembalikan semua user di sesi ini ke WORKING
  try {
    var userSheet = ss.getSheetByName(DB.USERS);
    if (userSheet) {
      var uRows = userSheet.getDataRange().getValues();
      for (var u = 1; u < uRows.length; u++) {
        if (uRows[u][1] === sessionId && uRows[u][3] === "DONE") {
          userSheet.getRange(u + 1, 4, 1, 2).setValues([["WORKING", new Date()]]);
        }
      }
    }
  } catch (e) {
    console.warn("reopenSession: reset user status gagal: " + e);
  }

  // Hapus data rekap lama (cegah duplikat saat finish ulang nanti)
  try {
    var rekapSheet = ss.getSheetByName(DB.REKAP);
    if (rekapSheet && rekapSheet.getLastRow() > 1) {
      var rekapData = rekapSheet.getDataRange().getValues();
      var rowsToDelete = [];
      for (var r = 1; r < rekapData.length; r++) {
        // Cari sessionId di semua kolom baris ini
        for (var c = 0; c < rekapData[r].length; c++) {
          if (String(rekapData[r][c]).trim() === sessionId) {
            rowsToDelete.push(r + 1);
            break;
          }
        }
      }
      // Delete bottom-up agar index tidak shift
      rowsToDelete.sort(function(a, b) { return b - a; });
      for (var d = 0; d < rowsToDelete.length; d++) {
        rekapSheet.deleteRow(rowsToDelete[d]);
      }
    }
  } catch (e) {
    console.warn("reopenSession: cleanup rekap gagal (non-critical): " + e);
  }

  // Invalidasi cache dashboard agar sesi muncul lagi di dashboard
  try {
    var cache = CacheService.getScriptCache();
    cache.put("DASH_INVALIDATED", String(Date.now()), 120);
  } catch (e) {}

  rebuildAndCacheSync_(sessionId);

  return {
    success: true,
    message: "Sesi berhasil dibuka kembali. Semua user dikembalikan ke status WORKING."
  };
}
