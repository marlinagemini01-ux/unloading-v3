function fixCacheDarurat() {
  var cache = CacheService.getScriptCache();
  cache.remove("APP_CONFIG_V5");
  cache.remove("SA_DRIVE_TOKEN");
  Logger.log("Cache lama berhasil dihapus.");
  Logger.log("Folder cache akan expire sendiri dalam max 1 jam.");
  Logger.log("Setelah ganti DriveService.gs, folder baru langsung benar.");
}

// Reset data transaksi untuk mode test.
// Header sheet tetap dipertahankan.
// Accounts, AppConfig, dan Data tidak disentuh.
function resetTrialDatabaseKeepAccounts() {
  var ss = getSpreadsheet_();
  var DB = getSheetNames_();
  var sheetNames = [
    DB.SESSIONS,
    DB.FLEET,
    DB.PHOTOS,
    DB.USERS,
    DB.REKAP,
    DB.AUDIT
  ];
  var protectedNames = {};
  protectedNames[DB.ACCOUNTS] = true;
  protectedNames[DB.CONFIG] = true;
  protectedNames[DB.DATA] = true;

  var result = [];
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(cfgNum("Lock Timeout (ms)") || 10000);

    sheetNames.forEach(function(name) {
      if (!name || protectedNames[name]) return;

      var sheet = ss.getSheetByName(name);
      if (!sheet) {
        result.push({ sheet: name, status: "missing", clearedRows: 0 });
        return;
      }

      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      var rowsToClear = Math.max(0, lastRow - 1);

      if (rowsToClear > 0 && lastCol > 0) {
        sheet.getRange(2, 1, rowsToClear, lastCol).clearContent();
      }

      result.push({ sheet: name, status: "cleared", clearedRows: rowsToClear });
    });

    clearAllCache();
    return {
      success: true,
      message: "Database transaksi berhasil dibersihkan. Header, Accounts, AppConfig, dan Data tetap aman.",
      cleared: result
    };
  } catch (e) {
    return {
      success: false,
      message: e.toString(),
      cleared: result
    };
  } finally {
    lock.releaseLock();
  }
}
