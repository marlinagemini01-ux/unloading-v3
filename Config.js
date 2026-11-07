// =========================================================================
// Config.gs — V6 PERFORMANCE EDITION
// OPTIMASI: onEdit auto-clear cache agar perubahan config instan
// =========================================================================
// [PERBAIKAN] clearConfigCache() sekarang juga hapus SA_DRIVE_TOKEN
// [PERBAIKAN] Tambah fungsi clearAllCache() untuk reset semua cache sekaligus
// =========================================================================

var CONFIG_SHEET_NAME = "AppConfig";
var CONFIG_CACHE_KEY  = "APP_CONFIG_V5";
var CONFIG_CACHE_TTL  = 300;
var PERF_DEBUG_ENABLED_CACHE_ = null;

// =========================================================================
// Spreadsheet binding
// - Prioritas utama tetap AppConfig key "Spreadsheet ID".
// - Untuk bootstrap awal, project bound-sheet bisa otomatis memakai
//   SpreadsheetApp.getActiveSpreadsheet().
// - Script Properties SPREADSHEET_ID tetap dipakai sebagai cache/bantuan untuk
//   project standalone, tetapi bukan lagi satu-satunya jalur.
// - DEFAULT_SPREADSHEET_ID_FOR_SETUP boleh diisi manual jika ingin setupDatabase()
//   langsung memakai database tertentu tanpa isi Script Properties dulu.
// =========================================================================
var SPREADSHEET_ID_PROPERTY_KEY = "SPREADSHEET_ID";
var SPREADSHEET_ID_CONFIG_KEY = "Spreadsheet ID";
var SCRIPT_ID_CONFIG_KEY = "Apps Script ID";
var DEFAULT_SPREADSHEET_ID_FOR_SETUP = "1Zo_-VgwdlCM7sqHUfdXk8rqdtMVuVXroLHTcROaVPrM";

function normalizeSpreadsheetId_(value) {
  var raw = String(value || "").trim();
  if (!raw) return "";

  var urlMatch = raw.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (urlMatch && urlMatch[1]) return urlMatch[1];

  var queryMatch = raw.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (queryMatch && queryMatch[1]) return queryMatch[1];

  return raw;
}

function getCurrentScriptId_() {
  try {
    return ScriptApp && ScriptApp.getScriptId ? String(ScriptApp.getScriptId() || "") : "";
  } catch (e) {
    return "";
  }
}

function getScriptPropertySpreadsheetId_() {
  try {
    return normalizeSpreadsheetId_(PropertiesService.getScriptProperties().getProperty(SPREADSHEET_ID_PROPERTY_KEY));
  } catch (e) {
    return "";
  }
}

function getActiveBoundSpreadsheetId_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    return ss ? normalizeSpreadsheetId_(ss.getId()) : "";
  } catch (e) {
    return "";
  }
}

function getConfigValueFromSpreadsheet_(spreadsheetId, key) {
  if (!spreadsheetId || !key) return "";
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) return "";
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === key) {
      return values[i][1];
    }
  }
  return "";
}

function syncSpreadsheetPropertyFromAppConfig_(bootstrapId, configuredId) {
  configuredId = normalizeSpreadsheetId_(configuredId);
  if (!configuredId || configuredId === bootstrapId) return configuredId || bootstrapId;
  PropertiesService.getScriptProperties().setProperty(SPREADSHEET_ID_PROPERTY_KEY, configuredId);
  clearConfigCache();
  return configuredId;
}

function upsertAppConfigValueBySpreadsheetId_(spreadsheetId, key, value, description, example) {
  if (!spreadsheetId || !key) return false;
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) return false;
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      if (description !== undefined) sheet.getRange(i + 1, 3).setValue(description);
      if (example !== undefined) sheet.getRange(i + 1, 4).setValue(example);
      return true;
    }
  }
  sheet.appendRow([key, value, description || "", example || ""]);
  return true;
}

function getSpreadsheetId_() {
  var propId = getScriptPropertySpreadsheetId_();
  var activeId = getActiveBoundSpreadsheetId_();
  var defaultId = normalizeSpreadsheetId_(DEFAULT_SPREADSHEET_ID_FOR_SETUP);
  var bootstrapId = propId || activeId || defaultId;
  if (!bootstrapId) {
    throw new Error(
      "Spreadsheet database belum bisa ditemukan. " +
      "Jika project GAS ini standalone, jalankan setSpreadsheetIdForThisProject(ID/URL Sheet) sekali. " +
      "Atau isi DEFAULT_SPREADSHEET_ID_FOR_SETUP di Config.js."
    );
  }

  var configId = "";
  try {
    configId = normalizeSpreadsheetId_(getConfigValueFromSpreadsheet_(bootstrapId, SPREADSHEET_ID_CONFIG_KEY));
  } catch (e) {}
  var effectiveId = syncSpreadsheetPropertyFromAppConfig_(bootstrapId, configId);
  try {
    if (!propId || propId !== effectiveId) {
      PropertiesService.getScriptProperties().setProperty(SPREADSHEET_ID_PROPERTY_KEY, effectiveId);
    }
    if (!configId) {
      upsertAppConfigValueBySpreadsheetId_(
        effectiveId,
        SPREADSHEET_ID_CONFIG_KEY,
        effectiveId,
        "ID Google Sheet database yang dipakai aplikasi.",
        "Contoh: 1Zo_-... atau URL Google Sheet"
      );
    }
  } catch (e) {}
  return effectiveId;
}

function setSpreadsheetIdForThisProject(spreadsheetIdOrUrl) {
  var spreadsheetId = normalizeSpreadsheetId_(spreadsheetIdOrUrl);
  if (!spreadsheetId) throw new Error("Spreadsheet ID / URL wajib diisi.");
  PropertiesService.getScriptProperties().setProperty(SPREADSHEET_ID_PROPERTY_KEY, spreadsheetId);
  try {
    upsertAppConfigValueBySpreadsheetId_(
      spreadsheetId,
      SPREADSHEET_ID_CONFIG_KEY,
      spreadsheetId,
      "ID Google Sheet database yang dipakai aplikasi.",
      "Contoh: 1Zo_-... atau URL Google Sheet"
    );
  } catch (e) {}
  clearAllCache();
  return {
    success: true,
    spreadsheetId: spreadsheetId,
    scriptId: getCurrentScriptId_()
  };
}

function getSpreadsheetBindingInfo() {
  var propId = getScriptPropertySpreadsheetId_();
  var activeId = getActiveBoundSpreadsheetId_();
  var defaultId = normalizeSpreadsheetId_(DEFAULT_SPREADSHEET_ID_FOR_SETUP);
  var configId = "";
  var effectiveId = "";
  var error = "";
  try {
    var bootstrapId = propId || activeId || defaultId;
    if (bootstrapId) configId = normalizeSpreadsheetId_(getConfigValueFromSpreadsheet_(bootstrapId, SPREADSHEET_ID_CONFIG_KEY));
    effectiveId = getSpreadsheetId_();
  } catch (idErr) {
    error = idErr.toString();
  }
  return {
    scriptId: getCurrentScriptId_(),
    propertyKey: SPREADSHEET_ID_PROPERTY_KEY,
    scriptPropertySpreadsheetId: propId,
    activeSpreadsheetId: activeId,
    defaultSpreadsheetIdForSetup: defaultId,
    appConfigSpreadsheetId: configId,
    spreadsheetId: effectiveId,
    source: configId ? "APP_CONFIG" : (propId ? "SCRIPT_PROPERTIES" : (activeId ? "ACTIVE_SPREADSHEET" : (defaultId ? "CODE_DEFAULT" : "NOT_SET"))),
    error: error
  };
}

function getSpreadsheet_() {
  return SpreadsheetApp.openById(getSpreadsheetId_());
}

function getConfig_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(CONFIG_CACHE_KEY);
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }

  var ss    = getSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) throw new Error("Sheet '" + CONFIG_SHEET_NAME + "' tidak ditemukan! Jalankan setupDatabase() dulu.");

  var data   = sheet.getDataRange().getValues();
  var config = {};
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0] || "").trim();
    var val = data[i][1];
    if (!key || key.charAt(0) === "#" || key.indexOf("▸") >= 0 || key.indexOf("💡") >= 0 || key.indexOf("•") >= 0) continue;
    config[key] = (val === null || val === undefined) ? "" : String(val).trim();
  }

  try { cache.put(CONFIG_CACHE_KEY, JSON.stringify(config), CONFIG_CACHE_TTL); } catch(e) {}
  return config;
}

function cfg(key) {
  var config = getConfig_();
  var val    = config[key];
  if (val === undefined) { console.warn("Config key tidak ditemukan: '" + key + "'"); return ""; }
  return val;
}

function cfgBool(key) { var v = cfg(key).toLowerCase(); return v === "ya" || v === "true" || v === "1" || v === "yes"; }
function cfgNum(key)  { var v = parseFloat(cfg(key)); return isNaN(v) ? 0 : v; }
function cfgList(key) { var v = cfg(key); if (!v) return []; return v.split(",").map(function(s){return s.trim();}).filter(function(s){return s.length>0;}); }

function isDebugPerformanceEnabled_() {
  if (PERF_DEBUG_ENABLED_CACHE_ !== null) return PERF_DEBUG_ENABLED_CACHE_;
  try {
    var config = getConfig_();
    var v = String(config["Debug Performance"] || "").trim().toLowerCase();
    PERF_DEBUG_ENABLED_CACHE_ = v === "ya" || v === "true" || v === "1" || v === "yes";
    return PERF_DEBUG_ENABLED_CACHE_;
  } catch (e) {
    PERF_DEBUG_ENABLED_CACHE_ = false;
    return false;
  }
}

function perfTrace_(action, requestId, step, extra) {
  if (!isDebugPerformanceEnabled_()) return;
  var msg = "[CRUD_TRACE] action=" + (action || "-")
    + " requestId=" + (requestId || "-")
    + " step=" + (step || "-")
    + " time=" + new Date().toISOString();
  if (extra !== undefined && extra !== null) {
    try { msg += " data=" + JSON.stringify(extra); }
    catch(e) { msg += " data=" + String(extra); }
  }
  console.log(msg);
}

// =========================================================================
// [PERBAIKAN] clearConfigCache()
// Sebelumnya: hanya hapus config cache.
// Sesudah   : hapus config cache + SA Drive Token agar keduanya fresh.
// =========================================================================
function clearConfigCache() {
  var cache = CacheService.getScriptCache();
  cache.remove(CONFIG_CACHE_KEY);
  cache.remove("SA_DRIVE_TOKEN"); // [PERBAIKAN] token ikut di-refresh
  return "Config cache cleared!";
}

// =========================================================================
// [BARU] clearAllCache()
// Fungsi darurat — jalankan dari Apps Script Editor jika ada masalah cache.
// Menghapus: config cache + SA token cache.
// Catatan: folder cache (f_*) tidak bisa di-list, akan expire otomatis
//          dalam 1 jam ATAU langsung terabaikan karena cache key sekarang
//          sudah menyertakan Root Folder ID (lihat DriveService.gs).
// =========================================================================
function clearAllCache() {
  var cache = CacheService.getScriptCache();
  cache.remove(CONFIG_CACHE_KEY);
  cache.remove("SA_DRIVE_TOKEN");
  Logger.log("✅ Semua cache utama berhasil dihapus.");
  Logger.log("ℹ️  Folder cache (f_*) akan expire otomatis dalam 1 jam.");
  Logger.log("ℹ️  Dengan fix DriveService.gs, folder cache lama otomatis diabaikan.");
  return "All cache cleared!";
}

/**
 * onEdit trigger — auto-clear config cache saat sheet AppConfig diubah.
 * Install: bisa sebagai simple trigger (otomatis) atau installable trigger.
 * Dengan ini, perubahan config langsung efektif tanpa tunggu 5 menit.
 */
function onEdit(e) {
  try {
    if (e && e.range && e.range.getSheet().getName() === CONFIG_SHEET_NAME) {
      clearConfigCache();
    }
  } catch(err) {
    // Non-critical — jangan crash app
  }
}

function getSheetNames_() {
  return {
    SESSIONS : cfg("Sessions Sheet"),
    FLEET    : cfg("Fleet Sheet"),
    DATA     : cfg("Data Sheet"),
    AUDIT    : cfg("Audit Sheet"),
    PHOTOS   : cfg("Photos Sheet"),
    USERS    : cfg("Users Sheet"),
    ACCOUNTS : cfg("Accounts Sheet"),
    REKAP    : cfg("Rekap Sheet"),
    SJ_QTY   : cfg("SJ Qty Sheet") || "SJ_QTY",
    HANDOVER_REVISION_LOG : cfg("Handover Revision Log Sheet") || "HandoverRevisionLog",
    CONFIG   : cfg("Config Sheet") || CONFIG_SHEET_NAME
  };
}

/**
 * Baca secrets — Sheet AppConfig dulu, fallback Script Properties.
 * Di Sheet AppConfig, tambahkan row:
 *   Key: "SA Email"        | Value: xxx@xxx.iam.gserviceaccount.com
 *   Key: "SA Private Key"  | Value: -----BEGIN PRIVATE KEY-----\n...
 *   Key: "Fontte Token"    | Value: token_fonnte_kamu
 */
function getSecretProps_() {
  var config = {};
  try { config = getConfig_(); } catch(e) {}
  var props = PropertiesService.getScriptProperties();
  
  return {
    SA_EMAIL      : config["SA Email"]       || props.getProperty("SA_EMAIL")       || "",
    SA_PRIVATE_KEY: config["SA Private Key"]  || props.getProperty("SA_PRIVATE_KEY") || "",
    FONNTE_TOKEN  : config["Fonnte Token"]    || props.getProperty("SA_FONNTE_TOKEN") || ""
  };
}

function buildPhotoUrl_(fileId) {
  var template = cfg("Photo URL Template") || "https://lh3.googleusercontent.com/d/{FILE_ID}=s800";
  return template.replace("{FILE_ID}", fileId);
}

function generateUniqueId_(prefixKey) {
  var prefix = cfg(prefixKey) || prefixKey || "ID";
  return prefix + "-" + new Date().getTime() + "-" + Math.random().toString(36).substring(2, 7);
}

function getRootFolder_() {
  var folderId = cfg("Root Folder ID");
  if (!folderId) throw new Error("Root Folder ID belum diset di sheet " + CONFIG_SHEET_NAME);
  try { return DriveApp.getFolderById(folderId); } catch(e) {
    if (cfgBool("Use Shared Drive")) {
      try { Drive.Drives.get(folderId); return DriveApp.getFolderById(folderId); }
      catch(e2) { throw new Error("Gagal akses Shared Drive: " + folderId + ". Error: " + e2.toString()); }
    }
    throw new Error("Gagal akses folder: " + folderId + ". Error: " + e.toString());
  }
}

function sanitize_(text, maxLengthKey) {
  if (!text || typeof text !== "string") return "";
  var maxLen = maxLengthKey ? cfgNum(maxLengthKey) : cfgNum("Max Input Length") || 50;
  var clean  = String(text).trim().replace(/^[=+\-@]+/, "");
  if (maxLen && clean.length > maxLen) clean = clean.substring(0, maxLen);
  return clean;
}

function titleCaseText_(text, maxLengthKey) {
  var clean = sanitize_(text, maxLengthKey)
    .replace(/\s+/g, " ")
    .toLowerCase();
  return clean.replace(/(^|[\s.'"-])([a-z])/g, function(match, prefix, letter) {
    return prefix + letter.toUpperCase();
  });
}

function validateRequired_(fields) {
  for (var name in fields) {
    if (!fields[name] || !String(fields[name]).trim()) throw new Error(name + " wajib diisi.");
  }
}

function replacePlaceholders_(template, data) {
  if (!template) return "";
  var result = template;
  for (var key in data) { result = result.split("[" + key + "]").join(String(data[key] || "")); }
  return result;
}

function toTimestampMs_(value) {
  if (!value) return 0;
  if (value instanceof Date) return value.getTime();
  var parsed = new Date(value).getTime();
  return isNaN(parsed) ? 0 : parsed;
}

function formatDurationRange_(startMs, endMs, emptyText) {
  var fallback = emptyText || "-";
  startMs = Number(startMs);
  endMs = Number(endMs);
  if (isNaN(startMs) || isNaN(endMs) || endMs <= startMs) return fallback;
  var diffMins = Math.floor((endMs - startMs) / 60000);
  var hours = Math.floor(diffMins / 60);
  var mins = diffMins % 60;
  var durationStr = "";
  if (hours > 0) durationStr += hours + " Jam ";
  if (mins > 0) durationStr += mins + " Menit";
  return durationStr.trim() || fallback;
}

// Kirim ke frontend — JANGAN kirim secrets
function getAppConfig() {
  try {
    return {
      success: true,
      data: {
        appTitle              : cfg("App Title")              || "InLens",
        sidebarBrand          : cfg("Sidebar Brand")          || "InLens",
        appVersion            : cfg("App Version")            || "v5.0.0",
        defaultDarkMode       : cfgBool("Default Dark Mode"),
        syncIntervalMs        : cfgNum("Sync Interval (ms)")  || 15000,
        toastDurationMs       : cfgNum("Toast Duration (ms)") || 2500,
        loginEnabled          : cfgBool("Login Enabled"),
        maxImagePx            : cfgNum("Max Image Pixel")     || 1000,
        jpegQuality           : parseFloat(cfg("JPEG Quality")) || 0.65,
        compressConcurrency   : cfgNum("Compress Concurrency") || 2,
        uploadConcurrency     : cfgNum("Upload Concurrency")   || 3,
        requirePhotoToFinish  : cfgBool("Require Photo To Finish"),
        requireAllDoneToFinish: cfgBool("Require All Done To Finish"),
        docSubfolderName      : cfg("Doc Subfolder Name")     || "Unloading",
        paletSubfolderName    : cfg("Palet Subfolder Name")   || "Palet",
        suratSubfolderName    : cfg("Surat Subfolder Name")   || "Surat Jalan",
        debugPerformance      : isDebugPerformanceEnabled_()
      }
    };
  } catch(e) { return { success: false, message: e.toString() }; }
}
