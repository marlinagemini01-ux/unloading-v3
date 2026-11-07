// =========================================================================
// DebugSjGateLive.gs — Debug khusus kasus "SJ Qty cocok tapi Belum Final"
// Paste ke paling bawah debug.gs, lalu Run: debugSjGateLive()
// Aman: read-only, tidak mengubah sheet/data.
// =========================================================================

var DEBUG_SJ_GATE_SESSION_ID = ""; // isi manual jika mau. Kosong = ambil ACTIVE terbaru.

function debugSjGateLive() {
  var sessionId = String(DEBUG_SJ_GATE_SESSION_ID || "").trim() || debugSjGateFindLatestActiveSessionId_();
  return debugSjGateById_(sessionId);
}

function debugSjGateById_(sessionId) {
  var out = [];
  var started = Date.now();
  var traceId = "DEBUG_GATE_" + started;

  debugSjGateLine_(out, "════════════════════════════════════════════");
  debugSjGateLine_(out, "DEBUG SJ QTY HANDOVER GATE");
  debugSjGateLine_(out, "Waktu      : " + new Date().toISOString());
  debugSjGateLine_(out, "Session ID : " + (sessionId || "TIDAK DITEMUKAN"));
  debugSjGateLine_(out, "Trace ID   : " + traceId);
  debugSjGateLine_(out, "════════════════════════════════════════════");

  if (!sessionId) {
    debugSjGateLine_(out, "❌ Tidak ada session ACTIVE. Isi DEBUG_SJ_GATE_SESSION_ID manual.");
    return debugSjGateFinish_(out);
  }

  debugSjGateStep_(out, "1. Function exists", function() {
    return {
      validateSjQtyMatching: typeof validateSjQtyMatching,
      getActualPalletRowsForSj_: typeof getActualPalletRowsForSj_,
      readSjQtyRows_: typeof readSjQtyRows_,
      ensureSjQtyReadyForHandover_: typeof ensureSjQtyReadyForHandover_,
      getHandOverData: typeof getHandOverData,
      syncSessionData: typeof syncSessionData
    };
  });

  debugSjGateStep_(out, "2. Session row", function() {
    return debugSjGateGetSessionRow_(sessionId);
  });

  debugSjGateStep_(out, "3. Raw actual pallet rows dari Photos/Fleet", function() {
    var rows = getActualPalletRowsForSj_(sessionId) || [];
    return {
      countAll: rows.length,
      countValid: rows.filter(function(r) { return r.valid; }).length,
      countInvalid: rows.filter(function(r) { return !r.valid; }).length,
      preview: rows.slice(0, 20)
    };
  });

  debugSjGateStep_(out, "4. Raw SJ Qty rows dari SJ_QTY", function() {
    var rows = readSjQtyRows_(sessionId) || [];
    return {
      countAll: rows.length,
      countValid: rows.filter(function(r) { return r.valid; }).length,
      countInvalid: rows.filter(function(r) { return !r.valid; }).length,
      preview: rows.slice(0, 20)
    };
  });

  debugSjGateStep_(out, "5. validateSjQtyMatching(sessionId) — INI YANG DIPAKAI PANEL/DETAIL", function() {
    var res = validateSjQtyMatching(sessionId, traceId);
    var d = res && res.data ? res.data : null;
    if (!d) return res;
    return {
      success: res.success,
      message: res.message,
      status: d.status,
      statusText: d.statusText,
      reasonCode: d.reasonCode,
      canHandover: d.canHandover,
      validationFinal: d.validationFinal,
      validationFinalPresent: Object.prototype.hasOwnProperty.call(d, "validationFinal"),
      actualCount: d.actualCount,
      sjCount: d.sjCount,
      matchedCount: d.matchedCount,
      totalActualQty: d.totalActualQty,
      totalSjQty: d.totalSjQty,
      totalVariance: d.totalVariance,
      missingInActual: d.missingInActual,
      missingInSj: d.missingInSj,
      qtyMismatch: d.qtyMismatch,
      duplicateActual: d.duplicateActual,
      duplicateSj: d.duplicateSj,
      invalidActual: d.invalidActual,
      invalidSj: d.invalidSj,
      detailRows: d.detailRows
    };
  });

  debugSjGateStep_(out, "6. ensureSjQtyReadyForHandover_(sessionId) — INI YANG DIPAKAI SERVER PRINT/FINISH", function() {
    return ensureSjQtyReadyForHandover_(sessionId, traceId);
  });

  debugSjGateStep_(out, "7. getHandOverData(sessionId)", function() {
    var res = getHandOverData(sessionId, traceId);
    if (!res || !res.success) return res;
    return {
      success: res.success,
      message: res.message,
      rowCount: res.data && res.data.rows ? res.data.rows.length : 0,
      totals: res.data && res.data.totals,
      session: res.data && res.data.session
    };
  });

  debugSjGateLine_(out, "");
  debugSjGateLine_(out, "KESIMPULAN CEPAT:");
  debugSjGateLine_(out, "- Kalau step 5 success=false: penyebab ada di message/ERROR step 5.");
  debugSjGateLine_(out, "- Kalau validationFinalPresent=false: SjQtyService.gs patch belum ke-push/deploy.");
  debugSjGateLine_(out, "- Kalau step 5 canHandover=true tapi UI tetap Belum Final: Index.html/deployment/cache browser yang bermasalah.");
  debugSjGateLine_(out, "- Kalau step 5 canHandover=false: lihat reasonCode, missing, duplicate, qtyMismatch.");
  debugSjGateLine_(out, "Durasi: " + (Date.now() - started) + " ms");

  return debugSjGateFinish_(out);
}

function debugSjGateStep_(out, title, fn) {
  debugSjGateLine_(out, "");
  debugSjGateLine_(out, "▶ " + title);
  try {
    var result = fn();
    debugSjGateLine_(out, "✅ OK");
    debugSjGateLine_(out, debugSjGateStringify_(result));
  } catch (e) {
    debugSjGateLine_(out, "❌ ERROR");
    debugSjGateLine_(out, String(e && e.message ? e.message : e));
    if (e && e.stack) debugSjGateLine_(out, String(e.stack));
  }
}

function debugSjGateFindLatestActiveSessionId_() {
  try {
    var res = getActiveSessions();
    if (res && res.success && res.data && res.data.length) return res.data[0].id;
  } catch (e) {}

  try {
    var DB = getSheetNames_();
    var sheet = getSpreadsheet_().getSheetByName(DB.SESSIONS);
    if (!sheet) return "";
    var rows = sheet.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][5] || "").trim().toUpperCase() === "ACTIVE") return String(rows[i][0] || "").trim();
    }
  } catch (err) {}
  return "";
}

function debugSjGateGetSessionRow_(sessionId) {
  var DB = getSheetNames_();
  var sheet = getSpreadsheet_().getSheetByName(DB.SESSIONS);
  if (!sheet) return { found: false, message: "Sheet Sessions tidak ditemukan" };
  var rows = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0] || "").trim() === String(sessionId || "").trim()) {
      return {
        found: true,
        rowNumber: i + 1,
        id: rows[i][0],
        date: rows[i][1],
        brand: rows[i][2],
        shipment: rows[i][3],
        pic: rows[i][4],
        status: rows[i][5],
        createdBy: rows[i][8] || rows[i][4]
      };
    }
  }
  return { found: false, sessionId: sessionId };
}

function debugSjGateStringify_(value) {
  try {
    return JSON.stringify(value, function(key, val) {
      if (val instanceof Date) return val.toISOString();
      return val;
    }, 2);
  } catch (e) {
    return String(value);
  }
}

function debugSjGateLine_(out, text) {
  out.push(String(text == null ? "" : text));
}

function debugSjGateFinish_(out) {
  var text = out.join("\n");
  Logger.log(text);
  return text;
}



function testLogin() {
  var result = verifyLogin("Admin", "123");
  Logger.log(JSON.stringify(result));
}

// =========================================================================
// TEST: updateSession — rename + pindah folder Drive
// Cara pakai:
//   1. Isi SESSION_ID_TEST di bawah dengan ID sesi yang mau diedit
//      (ambil dari sheet Sessions kolom pertama)
//   2. Isi data baru (brand, shipment, tanggal)
//   3. Klik Run → testUpdateSession()
//   4. Lihat hasil di Execution Log (Ctrl+Enter)
// =========================================================================
function testUpdateSession() {
  // ============================================================
  // GANTI NILAI INI SESUAI DATA YANG MAU DI-TEST
  // ============================================================
  var SESSION_ID_TEST = "SESS-1775851321557-rnwnb"; // dari sheet Sessions kolom A
  var NEW_BRAND = "Acwell"; // brand baru
  var NEW_SHIPMENT = "PO-TEST-99"; // shipment baru
  var NEW_DATE = "2026-04-11"; // tanggal (format yyyy-MM-dd)
  var REQUESTED_BY = "admin"; // username pembuat sesi
  // ============================================================

  Logger.log("═══════════════════════════════════════");
  Logger.log("  TEST: updateSession");
  Logger.log("  Session ID : " + SESSION_ID_TEST);
  Logger.log("  Brand baru : " + NEW_BRAND);
  Logger.log("  Shipment   : " + NEW_SHIPMENT);
  Logger.log("  Tanggal    : " + NEW_DATE);
  Logger.log("═══════════════════════════════════════");

  var result = updateSession(
    SESSION_ID_TEST,
    { brand: NEW_BRAND, shipment: NEW_SHIPMENT, date: NEW_DATE },
    REQUESTED_BY,
  );

  Logger.log("RESULT: " + JSON.stringify(result, null, 2));

  if (result.success) {
    Logger.log("✅ SUKSES!");
    Logger.log("   Brand    : " + result.data.brand);
    Logger.log("   Shipment : " + result.data.shipment);
    Logger.log("   Tanggal  : " + result.data.date);
    if (result.driveWarning) {
      Logger.log("⚠️  Drive warning: " + result.driveWarning);
    } else {
      Logger.log("   Drive    : folder berhasil di-rename/pindah");
    }
  } else {
    Logger.log("❌ GAGAL: " + result.message);
  }
}

// =========================================================================
// TEST: cek apakah folder Drive bisa di-rename dan dipindah
// Isi SESSION_ID dengan ID sesi dari sheet Sessions kolom A
// Contoh: "SESS-1775851321557-rnwnb"
// Fungsi ini otomatis cari folder URL-nya dari sheet.
// =========================================================================
function testDriveFolderMove() {
  // ← Isi dengan Session ID dari sheet Sessions kolom A
  var SESSION_ID = "SESS-1775851321557-rnwnb"; // contoh: SESS-1775851321557-rnwnb
  var NEW_BRAND_NAME = "TEST_BRAND_" + Date.now(); // brand folder baru (sementara, untuk test)

  Logger.log("═══════════════════════════════════════");
  Logger.log("TEST: Drive folder rename + move");
  Logger.log("Session ID: " + SESSION_ID);
  Logger.log("═══════════════════════════════════════");

  // Step 1: Cari folder URL dari sheet Sessions
  try {
    var DB = getSheetNames_();
    var ss = getSpreadsheet_();
    var sSession = ss.getSheetByName(DB.SESSIONS);
    var rows = sSession.getDataRange().getValues();
    var folderUrl = null;
    var sessionBrand = null;

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === SESSION_ID) {
        folderUrl = String(rows[i][7] || "");
        sessionBrand = String(rows[i][2] || "");
        Logger.log("✅ Sesi ditemukan di sheet");
        Logger.log("   Brand    : " + sessionBrand);
        Logger.log("   Folder URL: " + folderUrl);
        break;
      }
    }

    if (!folderUrl) {
      Logger.log("❌ Session ID tidak ditemukan di sheet Sessions.");
      Logger.log(
        "   Pastikan SESSION_ID diisi dengan benar (dari kolom A sheet Sessions).",
      );
      return;
    }

    // Step 2: Extract Drive Folder ID dari URL
    var m = folderUrl.match(/\/folders\/([a-zA-Z0-9_-]+)/);
    if (!m) {
      Logger.log("❌ Folder URL tidak valid atau kosong: " + folderUrl);
      return;
    }
    var driveFolderId = m[1];
    Logger.log("   Drive Folder ID: " + driveFolderId);

    // Step 3: Coba akses folder
    var sessionFolder = DriveApp.getFolderById(driveFolderId);
    Logger.log("✅ Folder Drive ditemukan: '" + sessionFolder.getName() + "'");

    // Step 4: Cek parent (brand folder sekarang)
    var oldParents = sessionFolder.getParents();
    var oldBrandFolder = oldParents.hasNext() ? oldParents.next() : null;
    Logger.log(
      "   Parent sekarang: " +
        (oldBrandFolder
          ? "'" +
            oldBrandFolder.getName() +
            "' (" +
            oldBrandFolder.getId() +
            ")"
          : "tidak ada"),
    );

    // Step 5: Coba rename (pakai nama test sementara)
    var originalName = sessionFolder.getName();
    var testName = "TEST_RENAME_" + Date.now();
    sessionFolder.setName(testName);
    Logger.log("✅ setName() berhasil → '" + testName + "'");

    // Step 6: Kembalikan nama asli
    sessionFolder.setName(originalName);
    Logger.log("✅ Nama dikembalikan → '" + originalName + "'");

    // Step 7: Coba addFolder ke folder brand baru (test)
    var rootFolder = getRootFolder_();
    var newBrandFolder = getOrCreateFolderCached_(
      rootFolder,
      NEW_BRAND_NAME,
      "f_test_" + NEW_BRAND_NAME,
    );
    Logger.log(
      "✅ Brand folder test dibuat: '" + newBrandFolder.getName() + "'",
    );

    newBrandFolder.addFolder(sessionFolder);
    Logger.log("✅ addFolder() ke brand folder baru berhasil");

    if (oldBrandFolder && oldBrandFolder.getId() !== newBrandFolder.getId()) {
      oldBrandFolder.removeFolder(sessionFolder);
      Logger.log("✅ removeFolder() dari parent lama berhasil");
    }

    // Step 8: Kembalikan ke brand folder asli
    if (oldBrandFolder) {
      oldBrandFolder.addFolder(sessionFolder);
      newBrandFolder.removeFolder(sessionFolder);
      Logger.log(
        "✅ Dikembalikan ke parent asli '" + oldBrandFolder.getName() + "'",
      );
    }

    // Hapus test brand folder (cleanup)
    try {
      newBrandFolder.setTrashed(true);
      Logger.log("✅ Folder test dihapus");
    } catch (e) {}

    Logger.log("");
    Logger.log("═══════ SEMUA LANGKAH BERHASIL ═══════");
    Logger.log("DriveApp.addFolder() + removeFolder() bekerja normal.");
    Logger.log("Fungsi updateSession() siap dipakai.");
  } catch (e) {
    Logger.log("❌ ERROR: " + e.toString());
    Logger.log("   Stack: " + e.stack);
  }
}

// =========================================================================
// DebugAccess.gs — Diagnosa akses lintas domain (perusahaan vs Gmail)
// Paste di GAS Editor, jalankan debugFullAccess() dari editor,
// atau buka URL deploy untuk test dari browser akun Gmail
// =========================================================================

/**
 * TEST 1: Spreadsheet bisa dibuka pakai openById?
 */
function testSpreadsheetAccess_() {
  var result = {
    test: "Spreadsheet Access (openById)",
    status: "FAIL",
    detail: "",
  };
  try {
    var ss = getSpreadsheet_();
    result.status = "PASS";
    result.detail =
      "Berhasil buka: " + ss.getName() + " (ID: " + ss.getId() + ")";
  } catch (e) {
    result.detail = e.toString();
  }
  return result;
}

/**
 * TEST 2: Sheet AppConfig bisa dibaca?
 */
function testConfigSheet_() {
  var result = { test: "Config Sheet (AppConfig)", status: "FAIL", detail: "" };
  try {
    var ss = getSpreadsheet_();
    var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    if (!sheet) {
      result.detail = "Sheet '" + CONFIG_SHEET_NAME + "' TIDAK DITEMUKAN";
      return result;
    }
    var rows = sheet.getLastRow();
    result.status = "PASS";
    result.detail = "Sheet ditemukan, " + rows + " baris data";
  } catch (e) {
    result.detail = e.toString();
  }
  return result;
}

/**
 * TEST 3: cfg() bisa baca config?
 */
function testCfgFunction_() {
  var result = { test: "cfg() Function", status: "FAIL", detail: "" };
  try {
    var appTitle = cfg("App Title");
    var rootFolder = cfg("Root Folder ID");
    result.status = "PASS";
    result.detail =
      "App Title = '" +
      appTitle +
      "', Root Folder ID = '" +
      (rootFolder ? rootFolder.substring(0, 10) + "..." : "KOSONG") +
      "'";
  } catch (e) {
    result.detail = e.toString();
  }
  return result;
}

/**
 * TEST 4: Semua sheet database bisa diakses?
 */
function testAllSheets_() {
  var result = { test: "All Database Sheets", status: "FAIL", detail: "" };
  try {
    var DB = getSheetNames_();
    var ss = getSpreadsheet_();
    var found = [];
    var missing = [];

    var names = {
      Sessions: DB.SESSIONS,
      Fleet: DB.FLEET,
      Data: DB.DATA,
      Audit: DB.AUDIT,
      Photos: DB.PHOTOS,
      Users: DB.USERS,
      Accounts: DB.ACCOUNTS,
      Rekap: DB.REKAP,
    };

    for (var label in names) {
      var sheetName = names[label];
      if (!sheetName) {
        missing.push(label + " (nama kosong di config)");
        continue;
      }
      var sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        found.push(label + " ✓");
      } else {
        missing.push(label + " (" + sheetName + ") ✗");
      }
    }

    if (missing.length === 0) {
      result.status = "PASS";
      result.detail = "Semua " + found.length + " sheet ditemukan";
    } else {
      result.status = "PARTIAL";
      result.detail =
        "Ditemukan: " + found.length + " | Missing: " + missing.join(", ");
    }
  } catch (e) {
    result.detail = e.toString();
  }
  return result;
}

/**
 * TEST 5: Drive root folder bisa diakses?
 */
function testDriveAccess_() {
  var result = { test: "Drive Root Folder", status: "FAIL", detail: "" };
  try {
    var folder = getRootFolder_();
    result.status = "PASS";
    result.detail =
      "Folder OK: " + folder.getName() + " (ID: " + folder.getId() + ")";
  } catch (e) {
    result.detail = e.toString();
  }
  return result;
}

/**
 * TEST 6: OAuth token bisa didapat? (untuk upload foto)
 */
function testOAuthToken_() {
  var result = { test: "OAuth Token (Deployer)", status: "FAIL", detail: "" };
  try {
    var token = ScriptApp.getOAuthToken();
    if (token && token.length > 20) {
      result.status = "PASS";
      result.detail =
        "Token didapat, panjang " + token.length + " karakter (ya...xxx)";
    } else {
      result.detail = "Token kosong atau terlalu pendek";
    }
  } catch (e) {
    result.detail = e.toString();
  }
  return result;
}

/**
 * TEST 7: Siapa yang menjalankan script ini?
 */
function testEffectiveUser_() {
  var result = { test: "User Identity", status: "INFO", detail: "" };
  try {
    var effective = Session.getEffectiveUser().getEmail();
    var active = "N/A";
    try {
      active = Session.getActiveUser().getEmail();
    } catch (e) {
      active = "(tidak bisa dibaca — normal untuk user luar domain)";
    }
    result.detail =
      "Effective (script owner): " +
      effective +
      " | Active (pengakses): " +
      active;
  } catch (e) {
    result.detail = e.toString();
  }
  return result;
}

/**
 * TEST 8: Login flow bisa jalan?
 */
function testLoginFlow_() {
  var result = { test: "Login System", status: "FAIL", detail: "" };
  try {
    var loginEnabled = cfgBool("Login Enabled");
    if (!loginEnabled) {
      result.status = "PASS";
      result.detail =
        "Login DISABLED — semua user masuk sebagai Guest (ini OK)";
      return result;
    }

    var DB = getSheetNames_();
    var sheet = getSpreadsheet_().getSheetByName(DB.ACCOUNTS);
    if (!sheet) {
      result.detail = "Sheet '" + DB.ACCOUNTS + "' tidak ditemukan!";
      return result;
    }

    var rows = sheet.getLastRow() - 1;
    result.status = "PASS";
    result.detail =
      "Login ENABLED, " + rows + " akun terdaftar di sheet " + DB.ACCOUNTS;
  } catch (e) {
    result.detail = e.toString();
  }
  return result;
}

// =========================================================================
// MAIN: Jalankan semua test
// =========================================================================
function debugFullAccess() {
  var tests = [
    testSpreadsheetAccess_(),
    testConfigSheet_(),
    testCfgFunction_(),
    testAllSheets_(),
    testDriveAccess_(),
    testOAuthToken_(),
    testEffectiveUser_(),
    testLoginFlow_(),
  ];

  var lines = [];
  lines.push("═══════════════════════════════════════════");
  lines.push("  InLens — Debug Akses Lintas Domain");
  lines.push(
    "  Waktu: " +
      new Date().toLocaleString("id-ID", { timeZone: "Asia/Jakarta" }),
  );
  lines.push("═══════════════════════════════════════════");
  lines.push("");

  var passCount = 0,
    failCount = 0;

  for (var i = 0; i < tests.length; i++) {
    var t = tests[i];
    var icon =
      t.status === "PASS"
        ? "✅"
        : t.status === "FAIL"
          ? "❌"
          : t.status === "PARTIAL"
            ? "⚠️"
            : "ℹ️";
    lines.push(icon + " [" + t.status + "] " + t.test);
    lines.push("   → " + t.detail);
    lines.push("");
    if (t.status === "PASS") passCount++;
    if (t.status === "FAIL") failCount++;
  }

  lines.push("───────────────────────────────────────────");
  lines.push(
    "HASIL: " +
      passCount +
      " PASS, " +
      failCount +
      " FAIL, " +
      (tests.length - passCount - failCount) +
      " INFO/PARTIAL",
  );

  if (failCount === 0) {
    lines.push("");
    lines.push("🎉 SEMUA TEST PASS! App siap diakses publik.");
  } else {
    lines.push("");
    lines.push(
      "🔧 Ada " + failCount + " test gagal. Perbaiki dulu sebelum deploy.",
    );
  }

  var output = lines.join("\n");
  console.log(output);
  return output;
}

// =========================================================================
// Endpoint untuk test dari browser (akun Gmail)
// Akses: Deploy → Test deployments → buka URL dari akun Gmail
// =========================================================================
function debugFromBrowser() {
  var result = debugFullAccess();
  return ContentService.createTextOutput(result).setMimeType(
    ContentService.MimeType.PLAIN_TEXT,
  );
}
