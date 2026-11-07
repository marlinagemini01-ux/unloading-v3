// =========================================================================
// DriveService.gs — V6 PERFORMANCE EDITION
// =========================================================================
// OPTIMASI:
//   - getOrCreateFolder_() → cache check SEBELUM lock (non-blocking read)
//   - Lock hanya dipegang saat createFolder() (bukan seluruh operasi)
//   - Concurrency: multiple user bisa read cache tanpa saling block
// =========================================================================
// [PERBAIKAN] buildFolderPath_():
//   Cache key folder sekarang menyertakan Root Folder ID.
//   Sebelumnya: "f_BrandA"  → tidak peduli root folder mana
//   Sesudah   : "f_<rootId>_BrandA" → unik per root folder
//   Efek: kalau Root Folder ID pernah berubah / belum diset saat testing,
//         cache lama otomatis diabaikan (beda key) dan folder baru dibuat
//         di lokasi yang benar sesuai AppConfig.
// =========================================================================

// =========================================================================
// getOrCreateFolder_()
// OPTIMASI: 2-phase approach untuk reduce lock contention
//   Phase 1: Cek cache (tanpa lock) → instant return jika ada
//   Phase 2: Lock → cek folder → create jika perlu
// =========================================================================
function getOrCreateFolder_(parentFolder, folderName, cacheKey) {
  var cache = CacheService.getScriptCache();
  var cacheTTL = cfgNum("Cache Folder (detik)") || 3600;

  // === PHASE 1: Non-blocking cache read (tidak perlu lock) ===
  if (cacheKey) {
    var cachedId = cache.get(cacheKey);
    if (cachedId) {
      try { return DriveApp.getFolderById(cachedId); }
      catch(e) { cache.remove(cacheKey); }
    }
  }

  // === PHASE 2: Lock hanya untuk create (reduce contention) ===
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    // Re-check cache setelah dapat lock (mungkin user lain sudah buat)
    if (cacheKey) {
      var cachedId2 = cache.get(cacheKey);
      if (cachedId2) {
        try { return DriveApp.getFolderById(cachedId2); }
        catch(e) { cache.remove(cacheKey); }
      }
    }

    // Cek folder di Drive
    var folders = parentFolder.getFoldersByName(folderName);
    if (folders.hasNext()) {
      var f = folders.next();
      if (cacheKey) cache.put(cacheKey, f.getId(), cacheTTL);
      return f;
    }

    // Propagation wait — 300ms
    Utilities.sleep(300);

    // Cek ulang setelah sleep
    folders = parentFolder.getFoldersByName(folderName);
    if (folders.hasNext()) {
      var f2 = folders.next();
      if (cacheKey) cache.put(cacheKey, f2.getId(), cacheTTL);
      return f2;
    }

    // Tidak ada, baru buat
    var newFolder = parentFolder.createFolder(folderName);
    if (cacheKey) cache.put(cacheKey, newFolder.getId(), cacheTTL);
    return newFolder;

  } finally {
    lock.releaseLock();
  }
}

function getOrCreateFolderCached_(parentFolder, folderName, cacheKey) {
  var cache = CacheService.getScriptCache();
  var cachedId = cache.get(cacheKey);
  if (cachedId) {
    try { return DriveApp.getFolderById(cachedId); } catch(e) { cache.remove(cacheKey); }
  }
  return getOrCreateFolder_(parentFolder, folderName, cacheKey);
}

function setFilePublic_(fileId) {
  try {
    Drive.Permissions.create(
      { role: "reader", type: "anyone" },
      fileId,
      { supportsAllDrives: true }
    );
  } catch (e) {
    console.warn("setFilePublic_ gagal: " + fileId + " — " + e.toString());
  }
}

/**
 * Batch set public permission untuk banyak file sekaligus.
 * Lebih efisien dari loop setFilePublic_ karena mengurangi overhead per call.
 */
function setFilesPublicBatch_(fileIds) {
  if (!fileIds || fileIds.length === 0) return;
  for (var i = 0; i < fileIds.length; i++) {
    try {
      Drive.Permissions.create(
        { role: "reader", type: "anyone" },
        fileIds[i],
        { supportsAllDrives: true }
      );
    } catch (e) {
      // Non-critical — lanjutkan ke file berikutnya
    }
  }
}

/**
 * Sanitize folder name — hapus karakter ilegal dan cegah path traversal.
 * @param {string} str — Nama folder mentah
 * @returns {string} Nama folder yang aman untuk Google Drive
 */
function safeFolderName_(str) {
  if (!str || typeof str !== "string") return "Unnamed";
  return String(str)
    .replace(/[\\/:*?"<>|]/g, "_")  // karakter ilegal di Drive
    .replace(/\.\./g, "_")           // cegah path traversal
    .replace(/^\s+|\s+$/g, "")       // trim whitespace
    .substring(0, 100)               // max 100 karakter
    || "Unnamed";
}

function cleanPaletKoli_(value) {
  var qty = Number(value);
  if (isNaN(qty) || qty < 0) return 0;
  return Math.floor(qty);
}

function getDeployerToken_() {
  return ScriptApp.getOAuthToken();
}

function getDriveAccessToken_() {
  var cache  = CacheService.getScriptCache();
  var ttl    = cfgNum("Cache Token (detik)") || 2700;
  var cached = cache.get("SA_DRIVE_TOKEN");
  if (cached) return cached;

  var secrets = getSecretProps_();
  if (!secrets.SA_EMAIL || !secrets.SA_PRIVATE_KEY) {
    throw new Error("SA_EMAIL atau SA_PRIVATE_KEY belum diset! Tambahkan di Sheet AppConfig (key: 'SA Email' dan 'SA Private Key') atau Script Properties.");
  }

  var privateKey = secrets.SA_PRIVATE_KEY.replace(/\\n/g, "\n").replace(/"/g, "").replace(/\\r/g, "");
  var keyMatch   = privateKey.match(/-----BEGIN PRIVATE KEY-----\s*([\s\S]*?)\s*-----END PRIVATE KEY-----/);
  if (keyMatch && keyMatch[1]) {
    var cleanKey  = keyMatch[1].replace(/\s+/g, "");
    var formatted = "-----BEGIN PRIVATE KEY-----\n";
    for (var i = 0; i < cleanKey.length; i += 64) formatted += cleanKey.substring(i, i + 64) + "\n";
    formatted += "-----END PRIVATE KEY-----\n";
    privateKey = formatted;
  } else {
    throw new Error("Format Private Key tidak valid!");
  }

  var now = Math.floor(Date.now() / 1000);
  function b64safe(str) { return Utilities.base64EncodeWebSafe(str).replace(/=+$/, ""); }

  var toSign = b64safe(JSON.stringify({ alg: "RS256", typ: "JWT" })) + "." +
    b64safe(JSON.stringify({
      iss: secrets.SA_EMAIL.trim(),
      scope: "https://www.googleapis.com/auth/drive",
      aud: "https://oauth2.googleapis.com/token",
      exp: now + 3600, iat: now
    }));

  var sig;
  try { sig = Utilities.computeRsaSha256Signature(toSign, privateKey); }
  catch(e) { throw new Error("Gagal sign JWT: " + e); }

  var jwt = toSign + "." + b64safe(sig);
  var res = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", {
    method: "post",
    payload: { grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", assertion: jwt },
    muteHttpExceptions: true
  });
  var json = JSON.parse(res.getContentText());
  if (json.error) throw new Error("Token SA gagal: " + json.error);
  cache.put("SA_DRIVE_TOKEN", json.access_token, ttl);
  return json.access_token;
}

function extractDriveFolderId_(value) {
  var text = String(value || "").trim();
  if (!text) return "";
  var match = text.match(/[-\w]{25,}/);
  return match ? match[0] : "";
}

function getExistingFleetFolderFromData_(data) {
  data = data || {};
  var directId = extractDriveFolderId_(data.fleetFolderUrl || data.folderUrl);
  if (directId) {
    try { return DriveApp.getFolderById(directId); } catch (e) {}
  }

  var containerId = String(data.containerId || data.fleetId || "").trim();
  var sessionId = String(data.sessionId || "").trim();
  var fleetNo = String(data.fleetNo || "").trim();
  if (!containerId && (!sessionId || !fleetNo)) return null;

  try {
    var DB = getSheetNames_();
    var sheet = getSpreadsheet_().getSheetByName(DB.FLEET);
    if (!sheet) return null;
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      var matchesContainer = containerId && String(rows[i][0] || "").trim() === containerId;
      var matchesSessionFleet = !containerId
        && sessionId
        && fleetNo
        && String(rows[i][1] || "").trim() === sessionId
        && String(rows[i][2] || "").trim() === fleetNo;
      if (!matchesContainer && !matchesSessionFleet) continue;
      var folderId = extractDriveFolderId_(rows[i][7]);
      if (!folderId) return null;
      try { return DriveApp.getFolderById(folderId); } catch (e2) { return null; }
    }
  } catch (e3) {}
  return null;
}

function updateFleetFolderUrlIfBlank_(containerId, folderUrl) {
  containerId = String(containerId || "").trim();
  folderUrl = String(folderUrl || "").trim();
  if (!containerId || !folderUrl) return;

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);
    var DB = getSheetNames_();
    var sheet = getSpreadsheet_().getSheetByName(DB.FLEET);
    if (!sheet) return;
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0] || "").trim() !== containerId) continue;
      if (!String(rows[i][7] || "").trim()) {
        sheet.getRange(i + 1, 8).setValue(folderUrl);
      }
      return;
    }
  } catch (e) {
    console.warn("updateFleetFolderUrlIfBlank_ gagal: " + e);
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

// =========================================================================
// buildFolderPath_()
// BUG 3 FIX: uploadType = null → hanya buat brand+session+armada folder
//            Subfolder Unloading/Palet hanya dibuat saat uploadType diisi
// =========================================================================
// [PERBAIKAN] Cache key sekarang menyertakan rootId:
//   Sebelumnya: bKey = "f_BrandA"
//   Sesudah   : bKey = "f_<rootFolderId>_BrandA"
//   Dengan ini, kalau Root Folder ID berubah di AppConfig, cache lama
//   otomatis tidak akan dipakai (key berbeda), sehingga folder baru
//   selalu dibuat di dalam root folder yang benar.
// =========================================================================
function buildFolderPath_(data, uploadType) {
  var rootFolder = getRootFolder_();
  var rootId     = rootFolder.getId(); // [PERBAIKAN] ambil ID root folder

  var placeholders = {
    Brand   : data.brand    || "",
    Shipment: data.shipment || "",
    Date    : data.date     || "",
    PIC     : data.pic      || "",
    Number  : data.fleetNo  || "",
    Driver  : data.driver   || "",
    Plate   : data.plate    || "",
    Type    : data.type     || ""
  };

  var brandName   = safeFolderName_(replacePlaceholders_(cfg("Brand Folder Name")   || "[Brand]",            placeholders));
  var sessionName = safeFolderName_(replacePlaceholders_(cfg("Session Folder Name") || "[Shipment] - [Date]", placeholders));

  // [PERBAIKAN] Sertakan rootId di cache key agar tidak terjadi cache stale
  // ketika Root Folder ID berubah di AppConfig
  var bKey = "f_" + rootId + "_" + brandName;
  var sKey = "f_" + rootId + "_" + brandName + "_" + sessionName;

  var brandFolder   = getOrCreateFolderCached_(rootFolder,   brandName,   bKey);
  var sessionFolder = getOrCreateFolderCached_(brandFolder,  sessionName, sKey);

  var result = { brandFolder: brandFolder, sessionFolder: sessionFolder };

  var documentationName = safeFolderName_(cfg("Documentation Folder Name") || "Dokumentasi");
  var documentationKey = sKey + "_documentation";
  if (!data.fleetNo) {
    result.documentationFolder = getOrCreateFolderCached_(sessionFolder, documentationName, documentationKey);
  }

  if (data.fleetNo) {
    var fleetName    = safeFolderName_(replacePlaceholders_(cfg("Fleet Folder Name") || "[Number] - [Driver] ([Plate])", placeholders));
    var aKey         = sKey + "_documentation_a" + data.fleetNo;
    var armadaFolder = getExistingFleetFolderFromData_(data);
    if (!armadaFolder) {
      var documentationFolder = getOrCreateFolderCached_(sessionFolder, documentationName, documentationKey);
      result.documentationFolder = documentationFolder;
      armadaFolder = getOrCreateFolderCached_(documentationFolder, fleetName, aKey);
    }
    result.armadaFolder = armadaFolder;

    // BUG 3 FIX: Hanya buat subfolder jika uploadType secara eksplisit diminta
    // uploadType = null / undefined → skip subfolder (saat addContainer)
    // uploadType = "dokumen" → hanya buat Unloading
    // uploadType = "palet"   → hanya buat Palet
    // uploadType = "ALL"     → buat keduanya (hanya dipakai oleh uploadPhotoBatch fallback)
    var docName = cfg("Doc Subfolder Name")  || "Unloading";
    var palName = cfg("Palet Subfolder Name") || "Palet";
    var suratName = cfg("Surat Subfolder Name") || "Surat Jalan";

    if (uploadType === "dokumen" || uploadType === "ALL") {
      result.docFolder = getOrCreateFolderCached_(armadaFolder, docName, aKey + "_dok");
    }
    if (uploadType === "palet" || uploadType === "ALL") {
      result.palFolder = getOrCreateFolderCached_(armadaFolder, palName, aKey + "_pal");
    }
    if (uploadType === "suratjalan" || uploadType === "ALL") {
      result.suratFolder = getOrCreateFolderCached_(armadaFolder, suratName, aKey + "_sj");
    }
  }

  return result;
}

// =========================================================================
// requestUploadCredentials()
// =========================================================================
function requestUploadCredentials(data) {
  try {
    var uploadType = data.uploadType; // "dokumen", "palet", atau "suratjalan"
    if (data && data.containerId) assertFleetUploadAllowed_(data.containerId);
    var folders    = buildFolderPath_(data, uploadType);
    var token      = getDeployerToken_();
    if (folders.armadaFolder) {
      updateFleetFolderUrlIfBlank_(data.containerId, folders.armadaFolder.getUrl());
    }

    var folderIds = {};
    if (folders.docFolder) folderIds.dokumen = folders.docFolder.getId();
    if (folders.palFolder) folderIds.palet   = folders.palFolder.getId();
    if (folders.suratFolder) folderIds.suratjalan = folders.suratFolder.getId();

    if (uploadType === "dokumen" && !folderIds.dokumen) {
      var docF = getOrCreateFolder_(folders.armadaFolder, cfg("Doc Subfolder Name") || "Unloading");
      folderIds.dokumen = docF.getId();
    }
    if (uploadType === "palet" && !folderIds.palet) {
      var palF = getOrCreateFolder_(folders.armadaFolder, cfg("Palet Subfolder Name") || "Palet");
      folderIds.palet = palF.getId();
    }
    if (uploadType === "suratjalan" && !folderIds.suratjalan) {
      var sjF = getOrCreateFolder_(folders.armadaFolder, cfg("Surat Subfolder Name") || "Surat Jalan");
      folderIds.suratjalan = sjF.getId();
    }

    return { success: true, data: { token: token, folderIds: folderIds } };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// =========================================================================
// logDirectUploads()
// V6 FIX: Lock HANYA melindungi sheet write (~300ms).
// Rebuild cache + setFilePublic dijalankan SETELAH lock dilepas.
// Sebelumnya: lock ditahan 5-15 detik (rebuild 2-3s + setPublic N×1-2s).
// Sesudah: lock ~300ms — user lain tidak perlu antri.
// =========================================================================
function getPhotoLogHeaderMap_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    map[String(headers[i] || "").trim()] = i;
  }
  return map;
}

function buildExistingPhotoIdMap_(sheet, headerMap) {
  var result = {};
  var photoIdCol = headerMap["Photo ID"] !== undefined ? headerMap["Photo ID"] : 0;
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return result;
  var values = sheet.getRange(2, photoIdCol + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    var photoId = String(values[i][0] || "").trim();
    if (photoId) result[photoId] = true;
  }
  return result;
}

function setPhotoLogCell_(row, headerMap, name, value) {
  if (headerMap[name] === undefined) return;
  row[headerMap[name]] = value;
}

function buildPhotoLogRow_(headerMap, item, meta, uploadedBy, now) {
  item = item || {};
  meta = meta || {};
  var fileId = String(item.fileId || "").trim();
  var rowLength = 0;
  Object.keys(headerMap).forEach(function(name) {
    if (headerMap[name] !== undefined && headerMap[name] + 1 > rowLength) rowLength = headerMap[name] + 1;
  });
  var row = new Array(rowLength);
  for (var i = 0; i < row.length; i++) row[i] = "";
  setPhotoLogCell_(row, headerMap, "Photo ID", fileId);
  setPhotoLogCell_(row, headerMap, "Container ID", item.containerId || "");
  setPhotoLogCell_(row, headerMap, "Type", item.type || "");
  setPhotoLogCell_(row, headerMap, "Drive URL", buildPhotoUrl_(fileId));
  setPhotoLogCell_(row, headerMap, "Timestamp", now);
  setPhotoLogCell_(row, headerMap, "Uploaded By", uploadedBy || "unknown");
  setPhotoLogCell_(row, headerMap, "Nama Palet", sanitize_(item.paletName || ""));
  setPhotoLogCell_(row, headerMap, "Koli Palet", cleanPaletKoli_(item.paletKoli));
  setPhotoLogCell_(row, headerMap, "Nama File", sanitize_(item.fileName || "", "Max Input Length"));
  setPhotoLogCell_(row, headerMap, "Local ID", item.localId || "");
  setPhotoLogCell_(row, headerMap, "Job ID", item.jobId || meta.jobId || "");
  setPhotoLogCell_(row, headerMap, "Log Status", "LOGGED");
  return row;
}

function appendPhotoLogsIdempotent_(sheet, items, meta, uploadedBy, now) {
  ensurePhotoLogColumns_(sheet);
  var headerMap = getPhotoLogHeaderMap_(sheet);
  var existing = buildExistingPhotoIdMap_(sheet, headerMap);
  var rows = [];
  var fileIdsToPublic = [];
  var duplicateFileIds = [];
  var invalidCount = 0;

  for (var i = 0; i < (items || []).length; i++) {
    var item = items[i] || {};
    var fileId = String(item.fileId || "").trim();
    if (!fileId) {
      invalidCount++;
      continue;
    }
    fileIdsToPublic.push(fileId);
    if (existing[fileId]) {
      duplicateFileIds.push(fileId);
      continue;
    }
    rows.push(buildPhotoLogRow_(headerMap, item, meta, uploadedBy, now));
    existing[fileId] = true;
  }

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    SpreadsheetApp.flush();
  }

  return {
    rowsWritten: rows.length,
    duplicateCount: duplicateFileIds.length,
    duplicateFileIds: duplicateFileIds,
    invalidCount: invalidCount,
    fileIdsToPublic: fileIdsToPublic
  };
}

function logDirectUploads(payload) {
  payload = payload || {};
  payload.meta = payload.meta || {};
  var traceId = payload.meta && payload.meta.traceId;
  perfTrace_("uploadPhoto", traceId, "log_sheet_server_start", { items: payload.items ? payload.items.length : 0 });
  var items = payload.items;
  if (!items || items.length === 0) return { success: true };
  assertUploadItemsWorkAllowed_(items, payload.meta.containerId);

  var DB   = getSheetNames_();
  var sheet = getSpreadsheet_().getSheetByName(DB.PHOTOS);
  ensurePhotoLogColumns_(sheet);
  var now = new Date();
  var uploadedBy = payload.meta.uploadedBy || "unknown";
  var sessionIdForSync = payload.meta.sessionId;
  var writeResult = {
    rowsWritten: 0,
    duplicateCount: 0,
    duplicateFileIds: [],
    invalidCount: 0,
    fileIdsToPublic: []
  };

  // === LOCK: hanya untuk write ke sheet ===
  var lock = LockService.getScriptLock();
  try {
    perfTrace_("uploadPhoto", traceId, "log_sheet_lock_wait_start");
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);
    perfTrace_("uploadPhoto", traceId, "log_sheet_lock_wait_end");
    perfTrace_("uploadPhoto", traceId, "log_sheet_write_start", { rows: items.length });
    writeResult = appendPhotoLogsIdempotent_(sheet, items, payload.meta, uploadedBy, now);
    perfTrace_("uploadPhoto", traceId, "log_sheet_write_end", {
      rowsWritten: writeResult.rowsWritten,
      duplicateCount: writeResult.duplicateCount,
      invalidCount: writeResult.invalidCount
    });
  } catch (e) {
    perfTrace_("uploadPhoto", traceId, "log_sheet_server_error", { message: e.toString() });
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock(); // ← Lepas lock SEGERA setelah sheet write
  }

  // === DI LUAR LOCK: operasi non-critical (tidak blokir user lain) ===
  // Invalidate cache supaya frontend sync ambil data terbaru
  rebuildAndCacheSync_(sessionIdForSync);
  syncFleetKoliFromPhotos_(sessionIdForSync);
  // Set file public (Drive API — non-critical, boleh gagal)
  var publicWarning = "";
  try {
    setFilesPublicBatch_(writeResult.fileIdsToPublic);
  } catch (publicErr) {
    publicWarning = publicErr.toString();
    console.warn("setFilesPublicBatch_ gagal setelah log Sheet sukses:", publicErr);
  }

  perfTrace_("uploadPhoto", traceId, "log_sheet_server_end", {
    count: items.length,
    rowsWritten: writeResult.rowsWritten,
    duplicateCount: writeResult.duplicateCount
  });
  return {
    success: true,
    count: items.length,
    rowsWritten: writeResult.rowsWritten,
    duplicateCount: writeResult.duplicateCount,
    duplicateFileIds: writeResult.duplicateFileIds,
    invalidCount: writeResult.invalidCount,
    publicWarning: publicWarning
  };
}

// =========================================================================
// uploadPhotoBatch() — Fallback GAS server upload
// BUG 4 FIX: Nama file pakai format yang readable, bukan DOKUMEN_timestamp_i.jpg
// =========================================================================
function uploadPhotoBatch(batchPayload) {
  try {
    batchPayload = batchPayload || {};
    var meta  = batchPayload.meta || {};
    var items = batchPayload.items || [];
    assertUploadItemsWorkAllowed_(items, meta.containerId);

    // Gunakan "ALL" agar kedua subfolder dibuat jika ada foto kedua jenis
    var folders = buildFolderPath_(meta, "ALL");
    if (folders.armadaFolder) {
      updateFleetFolderUrlIfBlank_(meta.containerId, folders.armadaFolder.getUrl());
    }

    var DB    = getSheetNames_();
    var sheet = getSpreadsheet_().getSheetByName(DB.PHOTOS);
    ensurePhotoLogColumns_(sheet);
    var results = [], failed = [], logItems = [];
    var uploadedBy = meta.uploadedBy || "unknown";

    // Ambil nama subfolder dari config untuk penamaan file
    var docSubName = cfg("Doc Subfolder Name")  || "Unloading";
    var palSubName = cfg("Palet Subfolder Name") || "Palet";
    var suratSubName = cfg("Surat Subfolder Name") || "Surat Jalan";

    var typeFolders = {};
    for (var i = 0; i < items.length; i++) {
      try {
        var item = items[i];

        if (!typeFolders[item.type]) {
          if (item.type === "dokumen") typeFolders[item.type] = folders.docFolder;
          else if (item.type === "palet") typeFolders[item.type] = folders.palFolder;
          else if (item.type === "suratjalan") typeFolders[item.type] = folders.suratFolder;
          
          if (!typeFolders[item.type]) {
            var label = docSubName;
            if (item.type === "palet") label = palSubName;
            else if (item.type === "suratjalan") label = suratSubName;
            typeFolders[item.type] = getOrCreateFolder_(folders.armadaFolder, label);
          }
        }

        // BUG 4 FIX: Nama file = [NamaSubfolder]_[Timestamp]_[Urutan].jpg
        var subName  = docSubName;
        if (item.type === "palet") subName = palSubName;
        else if (item.type === "suratjalan") subName = suratSubName;
        
        var tz       = cfg("Timezone") || "Asia/Jakarta";
        var dateStr  = Utilities.formatDate(new Date(), tz, "yyyyMMdd_HHmmss");
        var fileName = item.fileName || (subName + "_" + dateStr + "_" + (i + 1) + ".jpg");

        var blob = Utilities.newBlob(
          Utilities.base64Decode(item.base64Data), "image/jpeg", fileName
        );
        var file = typeFolders[item.type].createFile(blob);
        setFilePublic_(file.getId());

        var url = buildPhotoUrl_(file.getId());
        logItems.push({
          localId: item.localId || "",
          jobId: item.jobId || (meta && meta.jobId) || "",
          fileId: file.getId(),
          containerId: item.containerId,
          type: item.type,
          fileName: fileName,
          paletName: item.paletName || "",
          paletKoli: item.paletKoli || 0
        });
        results.push({ id: file.getId(), url: url, paletName: item.paletName || "", paletKoli: item.paletKoli || 0, fileName: fileName });

      } catch (e) {
        failed.push({ index: i, error: e.toString() });
      }
    }

    var writeResult = { rowsWritten: 0, duplicateCount: 0, duplicateFileIds: [] };
    if (logItems.length > 0) {
      var lock = LockService.getScriptLock();
      try {
        lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);
        writeResult = appendPhotoLogsIdempotent_(sheet, logItems, meta, uploadedBy, new Date());
      } finally {
        try { lock.releaseLock(); } catch (releaseErr) {}
      }
    }
    rebuildAndCacheSync_(meta.sessionId);
    syncFleetKoliFromPhotos_(meta.sessionId);
    return {
      success: true, data: results, failed: failed,
      totalUploaded: results.length, totalFailed: failed.length,
      rowsWritten: writeResult.rowsWritten,
      duplicateCount: writeResult.duplicateCount,
      duplicateFileIds: writeResult.duplicateFileIds || []
    };

  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// =========================================================================
// deletePhoto() — V7 PATCH: clearContent ganti deleteRow
// PATCH: Baris tidak dihapus, hanya dikosongkan → tidak ada index shift
// =========================================================================
function updatePhotoMetaBatch(updates) {
  if (!updates || updates.length === 0) return { success: false, message: "Tidak ada data palet yang dikirim." };

  var normalized = [];
  for (var i = 0; i < updates.length; i++) {
    var u = updates[i] || {};
    var photoId = String(u.photoId || "").trim();
    var paletName = sanitize_(u.paletName || "");
    var paletKoli = cleanPaletKoli_(u.paletKoli);
    if (!photoId) continue;
    var fileName = u.fileName || (paletName && paletKoli > 0 ? ("Palet " + paletName + " - " + paletKoli + " Koli.jpg") : "");
    normalized.push({
      photoId: photoId,
      paletName: paletName,
      paletKoli: paletKoli,
      fileName: fileName ? safeFolderName_(fileName) : ""
    });
  }
  if (normalized.length === 0) return { success: false, message: "Foto palet tidak ditemukan." };

  var DB = getSheetNames_();
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(DB.PHOTOS);
  ensurePhotoLogColumns_(sheet);

  var byId = {};
  normalized.forEach(function(u){ byId[u.photoId] = u; });

  var found = {};
  var containerIds = {};
  var editChecked = {};
  var updated = 0;
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);
    var data = sheet.getDataRange().getValues();
    for (var r = 1; r < data.length; r++) {
      var id = String(data[r][0] || "").trim();
      if (!id || !byId[id]) continue;
      var rowType = String(data[r][2] || "").toLowerCase();
      if (rowType !== "palet") continue;
      var editContainerId = String(data[r][1] || "").trim();
      if (editContainerId && !editChecked[editContainerId]) {
        assertFleetWorkStatusAllowed_(editContainerId, ["IN_PROGRESS"], "mengedit data palet");
        editChecked[editContainerId] = true;
      }
      var item = byId[id];
      var finalFileName = item.fileName || String(data[r][8] || "");
      sheet.getRange(r + 1, 7, 1, 3).setValues([[item.paletName, item.paletKoli, finalFileName]]);
      item.fileName = finalFileName;
      found[id] = item;
      if (data[r][1]) containerIds[data[r][1]] = true;
      updated++;
    }
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }

  var driveFailed = 0;
  for (var fid in found) {
    if (found[fid].fileName) {
      try { DriveApp.getFileById(fid).setName(found[fid].fileName); }
      catch(e) { driveFailed++; }
    }
  }

  var cids = Object.keys(containerIds);
  if (cids.length > 0) {
    try {
      var fleetData = ss.getSheetByName(DB.FLEET).getDataRange().getValues();
      var synced = {};
      for (var f = 1; f < fleetData.length; f++) {
        var cid = fleetData[f][0], sid = fleetData[f][1];
        if (containerIds[cid] && !synced[sid]) {
          synced[sid] = true;
          rebuildAndCacheSync_(sid);
          syncFleetKoliFromPhotos_(sid);
        }
      }
    } catch(e) {}
  }

  return {
    success: true,
    updated: updated,
    failed: normalized.length - updated,
    driveFailed: driveFailed
  };
}

function deletePhoto(fileId, traceId) {
  perfTrace_("deletePhoto", traceId, "server_start", { photoId: fileId });
  var targetId = String(fileId || "").trim();
  if (!targetId) return { success: false, message: "Photo ID kosong." };

  var lock = LockService.getScriptLock();
  var found = false;
  var containerId = null;
  var rowToClear = -1;
  try {
    perfTrace_("deletePhoto", traceId, "lock_wait_start");
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);
    perfTrace_("deletePhoto", traceId, "lock_wait_end");

    var DB    = getSheetNames_();
    var ss    = getSpreadsheet_();
    var sheet = ss.getSheetByName(DB.PHOTOS);
    perfTrace_("deletePhoto", traceId, "sheet_read_start");
    var data  = sheet.getDataRange().getValues();
    perfTrace_("deletePhoto", traceId, "sheet_read_end", { rows: data.length });

    for (var i = 1; i < data.length; i++) {
      // PATCH: skip baris kosong hasil clearContent sebelumnya
      if (!String(data[i][0] || "").trim()) continue;

      if (String(data[i][0]).trim() === targetId) {
        containerId = data[i][1];
        rowToClear  = i + 1;
        found = true;
        break;
      }
    }

    if (found && rowToClear > 0) {
      assertFleetWorkStatusAllowed_(containerId, ["IN_PROGRESS"], "menghapus foto");
      // PATCH: clearContent — baris tetap, isi dikosongkan (ganti deleteRow)
      perfTrace_("deletePhoto", traceId, "sheet_write_start", { row: rowToClear });
      sheet.getRange(rowToClear, 1, 1, sheet.getLastColumn()).clearContent();
      perfTrace_("deletePhoto", traceId, "sheet_write_end");
    }

  } catch (e) {
    perfTrace_("deletePhoto", traceId, "server_error", { message: e.toString() });
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }

  if (!found) {
    return { success: false, message: "Foto tidak ditemukan di database. Coba refresh halaman." };
  }

  // Drive trash di luar lock
  var driveOk = true;
  perfTrace_("deletePhoto", traceId, "drive_delete_start");
  try { Drive.Files.remove(targetId, { supportsAllDrives: true }); }
  catch(e) { driveOk = false; console.warn("deletePhoto: Drive hapus permanen gagal: " + e); }
  perfTrace_("deletePhoto", traceId, "drive_delete_end", { success: driveOk });

  if (containerId) {
    try {
      var ss2 = getSpreadsheet_();
      var fData = ss2.getSheetByName(getSheetNames_().FLEET).getDataRange().getValues();
      for (var j = 1; j < fData.length; j++) {
        if (fData[j][0] === containerId) {
          var sid = fData[j][1];
          rebuildAndCacheSync_(sid);
          syncFleetKoliFromPhotos_(sid);
          break;
        }
      }
    } catch(e) {}
  }

  perfTrace_("deletePhoto", traceId, "server_end", { success: true });
  return { success: true, driveTrashed: driveOk };
}

// =========================================================================
// deletePhotoBatch() — Hapus banyak foto dalam 1 call
// V6 FIX: Drive delete dijalankan SETELAH lock dilepas.
// PATCH: clearContent ganti deleteRow → tidak ada index shift → fix bug "tidak ketemu"
// =========================================================================
function deletePhotoBatch(fileIds) {
  if (!fileIds || fileIds.length === 0) return { success: true, deleted: 0, failed: 0 };

  console.log("deletePhotoBatch: input " + fileIds.length + " ids: " + JSON.stringify(fileIds).substring(0, 500));

  var DB           = getSheetNames_();
  var ss           = getSpreadsheet_();
  var containerIds = {};
  var matchedIds   = {};
  var sheetCleared = 0;

  // Normalisasi ID dari client
  var idSet = {};
  var normalizedInput = [];
  for (var k = 0; k < fileIds.length; k++) {
    var key = String(fileIds[k] || "").trim();
    if (key) { idSet[key] = true; normalizedInput.push(key); }
  }
  if (normalizedInput.length === 0) {
    return { success: false, message: "Tidak ada ID foto valid yang dikirim." };
  }

  // === LOCK: hanya untuk modifikasi sheet ===
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 30000);

    var sheet   = ss.getSheetByName(DB.PHOTOS);
    var data    = sheet.getDataRange().getValues();
    var lastCol = sheet.getLastColumn();

    // Kumpulkan baris yang perlu di-clear
    var rowsToClear = [];
    for (var i = 1; i < data.length; i++) {
      var cell = String(data[i][0] || "").trim();
      // PATCH: skip baris kosong hasil clearContent sebelumnya
      if (!cell) continue;
      if (idSet[cell]) {
        rowsToClear.push({ row: i + 1, containerId: data[i][1] });
        matchedIds[cell] = true;
        if (data[i][1]) containerIds[data[i][1]] = true;
      }
    }

    // PATCH: clearContent per baris — tidak ada index shift, tidak perlu sort bottom-up
    for (var guardCid in containerIds) {
      assertFleetWorkStatusAllowed_(guardCid, ["IN_PROGRESS"], "menghapus foto");
    }

    for (var r = 0; r < rowsToClear.length; r++) {
      try {
        sheet.getRange(rowsToClear[r].row, 1, 1, lastCol).clearContent();
        sheetCleared++;
      } catch(e) {
        console.warn("deletePhotoBatch: clearContent baris " + rowsToClear[r].row + " gagal: " + e);
      }
    }

  } catch (e) {
    return { success: false, message: "Gagal akses database: " + e.toString() };
  } finally {
    lock.releaseLock();
  }

  // Log ID yang tidak ketemu
  var notFoundIds = [];
  for (var ni = 0; ni < normalizedInput.length; ni++) {
    if (!matchedIds[normalizedInput[ni]]) notFoundIds.push(normalizedInput[ni]);
  }
  if (notFoundIds.length > 0) {
    console.warn("deletePhotoBatch: " + notFoundIds.length + " ID tidak ketemu di sheet: " + JSON.stringify(notFoundIds).substring(0, 500));
  }
  console.log("deletePhotoBatch: sheet rows cleared = " + sheetCleared + " / " + normalizedInput.length);

  if (sheetCleared === 0) {
    return {
      success: false,
      deleted: 0,
      failed: normalizedInput.length,
      message: "Tidak ada foto yang cocok di database (" + normalizedInput.length + " dikirim). Coba refresh halaman lalu ulangi."
    };
  }

  // === DI LUAR LOCK: Drive delete — HANYA untuk ID yang match di sheet ===
  var driveOk = 0, driveFail = 0;
  var failedDriveIds = [];
  for (var mid in matchedIds) {
    try { Drive.Files.remove(mid, { supportsAllDrives: true }); driveOk++; }
    catch(e) { driveFail++; failedDriveIds.push(mid); }
  }
  if (driveFail > 0) {
    console.warn("deletePhotoBatch: " + driveFail + " Drive delete gagal: " + failedDriveIds.join(","));
  }

  // Invalidate cache untuk semua sesi yang terdampak
  var cids = Object.keys(containerIds);
  if (cids.length > 0) {
    try {
      var fData = ss.getSheetByName(DB.FLEET).getDataRange().getValues();
      var synced = {};
      for (var j = 1; j < fData.length; j++) {
        var cid = fData[j][0], sid = fData[j][1];
        if (containerIds[cid] && !synced[sid]) {
          synced[sid] = true;
          rebuildAndCacheSync_(sid);
          syncFleetKoliFromPhotos_(sid);
        }
      }
    } catch(e) { console.warn("deletePhotoBatch: cache invalidate error: " + e); }
  }

  // `deleted` = jumlah row yang BENAR-benar di-clear dari sheet.
  // `failed` = ID yang tidak ketemu di sheet (bukan error Drive).
  // `driveFailed` = informasi tambahan kalau sheet ok tapi Drive delete gagal.
  return {
    success: true,
    deleted: sheetCleared,
    failed: normalizedInput.length - sheetCleared,
    driveFailed: driveFail
  };
}

/**
 * V7 DIAGNOSTIC: Scan PhotoLogs sheet for orphan/invalid rows.
 * Jalankan manual dari Apps Script editor: diagnosePhotoSheet() atau diagnosePhotoSheet('SESSION-ID')
 * Mode cleanup: diagnosePhotoSheet(sessionId, true) — akan hapus row orphan.
 */
function diagnosePhotoSheet(sessionId, cleanup) {
  var DB = getSheetNames_();
  var ss = getSpreadsheet_();
  var sPhotos = ss.getSheetByName(DB.PHOTOS);
  var sFleet  = ss.getSheetByName(DB.FLEET);
  if (!sPhotos) { Logger.log("PhotoLogs sheet tidak ditemukan."); return; }

  var containerToSession = {};
  if (sFleet) {
    var fData = sFleet.getDataRange().getValues();
    for (var i = 1; i < fData.length; i++) {
      containerToSession[fData[i][0]] = fData[i][1];
    }
  }

  var data = sPhotos.getDataRange().getValues();
  var emptyIdRows = [];
  var orphanContainerRows = [];
  var driveMissingRows = [];
  var scanned = 0;

  for (var r = 1; r < data.length; r++) {
    var fileId = String(data[r][0] || "").trim();
    var cid = data[r][1];
    var sid = containerToSession[cid];
    if (sessionId && sid !== sessionId) continue;
    scanned++;

    if (!fileId) { emptyIdRows.push(r + 1); continue; }
    if (cid && !containerToSession[cid]) { orphanContainerRows.push({row: r + 1, id: fileId}); continue; }
    try { DriveApp.getFileById(fileId); }
    catch(e) { driveMissingRows.push({row: r + 1, id: fileId}); }
  }

  Logger.log("=== diagnosePhotoSheet ===");
  Logger.log("Session filter: " + (sessionId || "(all)"));
  Logger.log("Rows scanned: " + scanned);
  Logger.log("Empty/invalid fileId rows: " + emptyIdRows.length + " -> " + JSON.stringify(emptyIdRows));
  Logger.log("Orphan container rows: " + orphanContainerRows.length + " -> " + JSON.stringify(orphanContainerRows).substring(0, 500));
  Logger.log("Drive file missing rows: " + driveMissingRows.length + " -> " + JSON.stringify(driveMissingRows).substring(0, 500));

  if (cleanup) {
    var toDelete = [].concat(emptyIdRows, orphanContainerRows.map(function(x){return x.row;}), driveMissingRows.map(function(x){return x.row;}));
    toDelete.sort(function(a,b){return b-a;});
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(15000);
      for (var d = 0; d < toDelete.length; d++) {
        try { sPhotos.deleteRow(toDelete[d]); } catch(e) {}
      }
    } finally { lock.releaseLock(); }
    Logger.log("Cleanup: deleted " + toDelete.length + " orphan rows.");
    if (sessionId) { try { rebuildAndCacheSync_(sessionId); } catch(e) {} }
  }

  return {
    scanned: scanned,
    emptyIdRows: emptyIdRows,
    orphanContainerRows: orphanContainerRows,
    driveMissingRows: driveMissingRows
  };
}

// =========================================================================
// movePhoto() — Pindah foto dari 1 container ke container lain
// Use case: User salah upload foto ke armada yang salah
// Proses: Update sheet → pindah file di Drive → invalidasi cache
// =========================================================================
function movePhoto(photoId, toContainerId) {
  var targetPhotoId = String(photoId || "").trim();
  var targetContId  = String(toContainerId || "").trim();
  if (!targetPhotoId || !targetContId) {
    return { success: false, message: "Photo ID dan armada tujuan wajib diisi." };
  }

  var DB = getSheetNames_();
  var ss = getSpreadsheet_();
  var photoSheet = ss.getSheetByName(DB.PHOTOS);
  var fleetSheet = ss.getSheetByName(DB.FLEET);

  var photoRow = -1, photoData = null, fromContainerId = null, photoType = null;
  var fromFleet = null, toFleet = null;

  // === LOCK: Update containerId di sheet ===
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);

    // Cari foto
    var photos = photoSheet.getDataRange().getValues();
    for (var i = 1; i < photos.length; i++) {
      if (!String(photos[i][0] || "").trim()) continue;
      if (String(photos[i][0]).trim() === targetPhotoId) {
        photoRow = i + 1;
        photoData = photos[i];
        fromContainerId = photos[i][1];
        photoType = photos[i][2] || "dokumen";
        break;
      }
    }
    if (photoRow < 0) return { success: false, message: "Foto tidak ditemukan." };

    if (fromContainerId === targetContId) {
      return { success: false, message: "Foto sudah di armada ini." };
    }

    // Validasi container tujuan & asal
    var fleetData = fleetSheet.getDataRange().getValues();
    for (var f = 1; f < fleetData.length; f++) {
      if (fleetData[f][0] === targetContId && fleetData[f][8] !== "DELETED") toFleet = fleetData[f];
      if (fleetData[f][0] === fromContainerId && fleetData[f][8] !== "DELETED") fromFleet = fleetData[f];
    }
    if (!toFleet) return { success: false, message: "Armada tujuan tidak ditemukan atau sudah dihapus." };
    assertFleetWorkStatusAllowed_(fromContainerId, ["IN_PROGRESS"], "memindahkan foto");
    assertFleetWorkStatusAllowed_(targetContId, ["IN_PROGRESS"], "menerima foto pindahan");

    // Update containerId di PhotoLogs
    photoSheet.getRange(photoRow, 2).setValue(targetContId);

  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }

  // === DI LUAR LOCK: Pindah file di Drive (non-critical) ===
  var driveOk = true;
  try {
    if (toFleet) {
      var sessionId = toFleet[1];
      var sessRows = ss.getSheetByName(DB.SESSIONS).getDataRange().getValues();
      var sessData = null;
      for (var s = 1; s < sessRows.length; s++) {
        if (sessRows[s][0] === sessionId) { sessData = sessRows[s]; break; }
      }

      if (sessData) {
        var tz = cfg("Timezone") || "Asia/Jakarta";
        var rawDate = sessData[1];
        var sessDate = rawDate instanceof Date
          ? Utilities.formatDate(rawDate, tz, "yyyy-MM-dd") : String(rawDate).trim();

        // Bangun folder tujuan
        var destFolders = buildFolderPath_({
          brand: String(sessData[2]).trim(),
          shipment: String(sessData[3]).trim(),
          date: sessDate,
          pic: String(sessData[4]).trim(),
          containerId: toFleet[0],
          fleetNo: toFleet[2],
          driver: String(toFleet[4]).trim(),
          plate: String(toFleet[5]).trim(),
          type: String(toFleet[3]).trim()
        }, photoType);

        var normalizedPhotoType = normalizePhotoType_(photoType);
        var destFolder = normalizedPhotoType === "palet"
          ? destFolders.palFolder
          : (normalizedPhotoType === "suratjalan" ? destFolders.suratFolder : destFolders.docFolder);
        if (destFolder) {
          var file = DriveApp.getFileById(targetPhotoId);

          // Pindah: tambah ke folder baru, hapus dari folder lama
          destFolder.addFile(file);
          var oldParents = file.getParents();
          while (oldParents.hasNext()) {
            var parent = oldParents.next();
            if (parent.getId() !== destFolder.getId()) {
              parent.removeFile(file);
            }
          }
        }
      }
    }
  } catch (driveErr) {
    driveOk = false;
    console.warn("movePhoto: pindah file Drive gagal (non-critical): " + driveErr);
  }

  // Invalidasi cache & sinkronisasi koli
  if (fromFleet) {
    rebuildAndCacheSync_(fromFleet[1]);
    syncFleetKoliFromPhotos_(fromFleet[1]);
  }
  if (toFleet && toFleet[1] !== (fromFleet ? fromFleet[1] : null)) {
    rebuildAndCacheSync_(toFleet[1]);
    syncFleetKoliFromPhotos_(toFleet[1]);
  }

  return {
    success: true,
    message: "Foto berhasil dipindah ke armada " + (toFleet ? toFleet[2] : "tujuan") + ".",
    driveMoved: driveOk
  };
}

// =========================================================================
// movePhotoBatch() — Pindah banyak foto sekaligus
// =========================================================================
function movePhotoBatch(fileIds, toContainerId) {
  var targetContId  = String(toContainerId || "").trim();
  if (!fileIds || fileIds.length === 0 || !targetContId) {
    return { success: false, message: "ID foto dan armada tujuan wajib diisi." };
  }

  // Normalisasi ID dari client
  var idSet = {};
  var normalizedInput = [];
  for (var k = 0; k < fileIds.length; k++) {
    var key = String(fileIds[k] || "").trim();
    if (key) { idSet[key] = true; normalizedInput.push(key); }
  }
  if (normalizedInput.length === 0) {
    return { success: false, message: "Tidak ada ID foto valid yang dikirim." };
  }

  var DB = getSheetNames_();
  var ss = getSpreadsheet_();
  var photoSheet = ss.getSheetByName(DB.PHOTOS);
  var fleetSheet = ss.getSheetByName(DB.FLEET);

  var matchedPhotos = []; // { id, row, fromContainerId, type }
  var fromContainers = {}; // untuk invalidasi cache
  var toFleet = null;
  var sheetUpdated = 0;

  // === LOCK: Update containerId di sheet ===
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 30000);

    // Cari armada tujuan
    var fleetData = fleetSheet.getDataRange().getValues();
    for (var f = 1; f < fleetData.length; f++) {
      if (fleetData[f][0] === targetContId && fleetData[f][8] !== "DELETED") {
        toFleet = fleetData[f];
        break;
      }
    }
    if (!toFleet) return { success: false, message: "Armada tujuan tidak ditemukan atau sudah dihapus." };

    // Cari foto yang match
    var photos = photoSheet.getDataRange().getValues();
    for (var i = 1; i < photos.length; i++) {
      var cell = String(photos[i][0] || "").trim();
      if (!cell) continue;
      
      if (idSet[cell]) {
        var fromContainerId = photos[i][1];
        if (fromContainerId !== targetContId) {
          matchedPhotos.push({
            id: cell,
            row: i + 1,
            fromContainerId: fromContainerId,
            type: photos[i][2] || "dokumen"
          });
          fromContainers[fromContainerId] = true;
        }
      }
    }

    assertFleetWorkStatusAllowed_(targetContId, ["IN_PROGRESS"], "menerima foto pindahan");
    for (var fromCid in fromContainers) {
      assertFleetWorkStatusAllowed_(fromCid, ["IN_PROGRESS"], "memindahkan foto");
    }

    // Update containerId di PhotoLogs secara batch
    for (var m = 0; m < matchedPhotos.length; m++) {
      try {
        photoSheet.getRange(matchedPhotos[m].row, 2).setValue(targetContId);
        sheetUpdated++;
      } catch(e) {
        console.warn("movePhotoBatch: setValue baris " + matchedPhotos[m].row + " gagal: " + e);
      }
    }

  } catch (e) {
    return { success: false, message: "Gagal akses database: " + e.toString() };
  } finally {
    lock.releaseLock();
  }

  if (sheetUpdated === 0) {
    return {
      success: false,
      moved: 0,
      message: "Tidak ada foto yang perlu dipindah (mungkin sudah ada di armada tujuan)."
    };
  }

  // === DI LUAR LOCK: Pindah file di Drive (non-critical) ===
  var driveOk = 0, driveFail = 0;
  try {
    if (toFleet) {
      var sessionId = toFleet[1];
      var sessRows = ss.getSheetByName(DB.SESSIONS).getDataRange().getValues();
      var sessData = null;
      for (var s = 1; s < sessRows.length; s++) {
        if (sessRows[s][0] === sessionId) { sessData = sessRows[s]; break; }
      }

      if (sessData) {
        var tz = cfg("Timezone") || "Asia/Jakarta";
        var rawDate = sessData[1];
        var sessDate = rawDate instanceof Date
          ? Utilities.formatDate(rawDate, tz, "yyyy-MM-dd") : String(rawDate).trim();

        // Bangun folder tujuan.
        var destFolders = buildFolderPath_({
          brand: String(sessData[2]).trim(),
          shipment: String(sessData[3]).trim(),
          date: sessDate,
          pic: String(sessData[4]).trim(),
          containerId: toFleet[0],
          fleetNo: toFleet[2],
          driver: String(toFleet[4]).trim(),
          plate: String(toFleet[5]).trim(),
          type: String(toFleet[3]).trim()
        }, "ALL");

        var docFolder = destFolders.docFolder;
        var palFolder = destFolders.palFolder;
        var suratFolder = destFolders.suratFolder;

        for (var m = 0; m < matchedPhotos.length; m++) {
          try {
            var normalizedType = normalizePhotoType_(matchedPhotos[m].type);
            var destFolder = normalizedType === "palet"
              ? palFolder
              : (normalizedType === "suratjalan" ? suratFolder : docFolder);
            if (destFolder) {
              var file = DriveApp.getFileById(matchedPhotos[m].id);
              // Pindah: tambah ke folder baru, hapus dari folder lama
              destFolder.addFile(file);
              var oldParents = file.getParents();
              while (oldParents.hasNext()) {
                var parent = oldParents.next();
                if (parent.getId() !== destFolder.getId()) {
                  parent.removeFile(file);
                }
              }
              driveOk++;
            }
          } catch(e) {
            driveFail++;
            console.warn("movePhotoBatch: gagal pindah file " + matchedPhotos[m].id + " di Drive: " + e);
          }
        }
      }
    }
  } catch (driveErr) {
    console.warn("movePhotoBatch: operasi Drive gagal (non-critical): " + driveErr);
  }

  // Invalidasi cache untuk semua sesi yang terdampak
  fromContainers[targetContId] = true; // container tujuan juga perlu di-invalidasi
  var cids = Object.keys(fromContainers);
  if (cids.length > 0) {
    try {
      var fData = ss.getSheetByName(DB.FLEET).getDataRange().getValues();
      var synced = {};
      for (var j = 1; j < fData.length; j++) {
        var cid = fData[j][0], sid = fData[j][1];
        if (fromContainers[cid] && !synced[sid]) {
          synced[sid] = true;
          rebuildAndCacheSync_(sid);
          syncFleetKoliFromPhotos_(sid);
        }
      }
    } catch(e) { console.warn("movePhotoBatch: cache invalidate error: " + e); }
  }

  return {
    success: true,
    moved: sheetUpdated,
    failed: normalizedInput.length - sheetUpdated,
    driveFailed: driveFail,
    message: sheetUpdated + " foto berhasil dipindah."
  };
}
