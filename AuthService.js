// =========================================================================
// AuthService.gs — V5.2: Plain text password, no hashing
// - Satu verifyLogin() (tidak ada duplikat!)
// - Password langsung dibaca dari sheet, TIDAK di-hash/ubah
// - Rate limiting (max 5 percobaan per 5 menit)
// - Audit logging setiap login attempt
// - Case sensitivity pakai config
// - Safe column access (tidak crash jika kolom kurang)
// =========================================================================

/**
 * Safe audit log — panggil logAudit_ jika ada, jangan crash kalau tidak ada.
 */
function safeAuditLog_(action, status, message, context) {
  try {
    // Coba pakai logAudit_ dari Constants.gs
    logAudit_(action, status, message, context);
  } catch(e) {
    // logAudit_ belum ada — tulis langsung ke sheet AuditLog
    try {
      var DB = getSheetNames_();
      var ss = getSpreadsheet_();
      var sheet = ss.getSheetByName(DB.AUDIT);
      if (sheet) sheet.appendRow([new Date(), action, status, message, context || ""]);
    } catch(e2) {
      console.warn("Audit log gagal:", e2.toString());
    }
  }
}

/**
 * Verifikasi password — langsung bandingkan dengan nilai di sheet.
 * @param {string} inputPassword — Password yang diinput user
 * @param {string} storedPassword — Password dari sheet
 * @param {boolean} caseSensitive — Apakah case-sensitive
 * @returns {boolean}
 */
function verifyPassword_(inputPassword, storedPassword, caseSensitive) {
  if (!inputPassword || !storedPassword) return false;
  if (caseSensitive === false) {
    return inputPassword.toLowerCase() === storedPassword.toLowerCase();
  }
  return inputPassword === storedPassword;
}

// =========================================================================
// Rate Limiting — max 5 percobaan per 5 menit per username
// =========================================================================
function checkLoginThrottle_(username) {
  var cache = CacheService.getScriptCache();
  var key = "login_attempts_" + username.toLowerCase();
  var attempts = Number(cache.get(key) || "0");
  if (attempts >= 5) {
    return { blocked: true, message: "Terlalu banyak percobaan login. Coba lagi dalam 5 menit." };
  }
  return { blocked: false, attempts: attempts };
}

function incrementLoginAttempt_(username) {
  var cache = CacheService.getScriptCache();
  var key = "login_attempts_" + username.toLowerCase();
  var attempts = Number(cache.get(key) || "0") + 1;
  cache.put(key, String(attempts), 300);
}

function clearLoginAttempts_(username) {
  CacheService.getScriptCache().remove("login_attempts_" + username.toLowerCase());
}

// =========================================================================
// verifyLogin() — SATU-SATUNYA versi
// Sheet Accounts format: Username | Password | NamaLengkap | Active | Signature File ID (opsional)
// =========================================================================
function verifyLogin(username, password) {
  try {
    if (!cfgBool("Login Enabled")) {
      return { success: true, data: { username: "guest", fullName: "Guest", role: "staff" } };
    }

    var throttle = checkLoginThrottle_(username);
    if (throttle.blocked) {
      safeAuditLog_("LOGIN_BLOCKED", "FAIL", throttle.message, username);
      return { success: false, message: throttle.message };
    }

    var DB    = getSheetNames_();
    var sheet = getSpreadsheet_().getSheetByName(DB.ACCOUNTS);
    if (!sheet) return { success: false, message: "Sheet " + DB.ACCOUNTS + " belum dibuat. Jalankan setupDatabase()." };

    var data = sheet.getDataRange().getValues();
    var usernameCaseSensitive = cfgBool("Username Case Sensitive");
    var passwordCaseSensitive = cfgBool("Password Case Sensitive");

    for (var i = 1; i < data.length; i++) {
      var row      = data[i];
      var dbUser   = String(row[0] || "").trim();
      var dbPass   = String(row[1] || "").trim();
      var fullName = String(row[2] || "").trim();
      // Kolom 4 (index 3) = Active — safe access jika kolom tidak ada
      var activeRaw = row.length > 3 ? String(row[3] || "").trim().toUpperCase() : "TRUE";

      // Username matching
      var matchUser = usernameCaseSensitive
        ? (dbUser === username.trim())
        : (dbUser.toLowerCase() === username.trim().toLowerCase());

      if (!matchUser) continue;

      // Cek aktif
      // "" (kosong) dianggap aktif — supaya user yang belum isi kolom Active tetap bisa login
      var isActive = (activeRaw === "" || activeRaw === "TRUE" || activeRaw === "YA" 
                   || activeRaw === "1" || activeRaw === "ACTIVE" || activeRaw === "YES");

      if (!isActive) {
        safeAuditLog_("LOGIN_INACTIVE", "FAIL", "Akun nonaktif: " + dbUser, username);
        return { success: false, message: "Akun nonaktif. Hubungi admin." };
      }

      // Password matching — langsung bandingkan, tidak ada hash
      var matchPass = verifyPassword_(password, dbPass, passwordCaseSensitive);

      if (!matchPass) {
        incrementLoginAttempt_(username);
        safeAuditLog_("LOGIN_FAIL", "FAIL", "Password salah untuk: " + dbUser, username);
        return { success: false, message: "Password salah." };
      }

      // === LOGIN BERHASIL ===
      clearLoginAttempts_(username);
      safeAuditLog_("LOGIN_SUCCESS", "OK", "Login berhasil: " + dbUser, username);

      return {
        success: true,
        data: {
          username: dbUser,
          fullName: fullName || dbUser,
          role: "staff"
        }
      };
    }

    // Username tidak ditemukan
    incrementLoginAttempt_(username);
    safeAuditLog_("LOGIN_FAIL", "FAIL", "Username tidak ditemukan: " + username, username);
    return { success: false, message: "Username tidak ditemukan." };

  } catch (e) {
    safeAuditLog_("LOGIN_ERROR", "ERROR", e.toString(), username);
    return { success: false, message: "Terjadi kesalahan sistem. Coba lagi." };
  }
}
