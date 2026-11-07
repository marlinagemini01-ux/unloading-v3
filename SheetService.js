// =========================================================================
// SheetService.gs — V6 PERFORMANCE EDITION
// =========================================================================
// OPTIMASI:
//   - getDashboardData() di-cache 15 detik (hemat 5-6x getDataRange)
//   - syncSessionData() cache nai dari 10s → 20s
//   - rebuildAndCacheSync_() invalidate dashboard cache juga
//   - getDropdownData() — unchanged (sudah OK)
//   - getCompletedSessions() — unchanged (sudah pagination)
// =========================================================================

var DASHBOARD_CACHE_KEY = "DASH_V6_";
var DASHBOARD_CACHE_TTL = 30; // detik — V7: naik dari 15, aman karena manual invalidation via DASH_INVALIDATED

// =========================================================================
// prewarmDashboardCache_() — V7: dipanggil via time trigger tiap 5 menit
// Tujuan: user pertama di tiap window cache 30s tidak kena cold read sheet.
// Aman kalau gagal (try/catch) — ini best-effort optimization.
// =========================================================================
function prewarmDashboardCache_() {
  try {
    var ss = getSpreadsheet_();
    var DB = getSheetNames_();
    var sSession = ss.getSheetByName(DB.SESSIONS);
    if (!sSession) return;
    var rows = sSession.getDataRange().getValues();
    var users = {};
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][5] === "ACTIVE") {
        var u = rows[i][8] || rows[i][4];
        if (u && !users[u]) { users[u] = true; try { getDashboardData(u); } catch(e) {} }
      }
    }
  } catch (e) { /* best-effort */ }
}

function setupPrewarmTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === "prewarmDashboardCache_") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("prewarmDashboardCache_")
    .timeBased()
    .everyMinutes(5)
    .create();
  return "Trigger pre-warm cache berhasil dibuat (tiap 5 menit).";
}

function rebuildAndCacheSync_(sessionId) {
  if (!sessionId) return;
  var cache = CacheService.getScriptCache();
  // Invalidate sync cache — LAZY: tidak rebuild di sini
  // Frontend sync berikutnya yang rebuild (mencegah double-read saat banyak user)
  cache.remove("sync_" + sessionId);
  // Invalidate ALL dashboard caches via timestamp pattern
  cache.put("DASH_INVALIDATED", String(Date.now()), 60);
  // V6 FIX: TIDAK panggil syncSessionData() di sini.
  // Dulu: setiap write → rebuild 4 sheet = 2-3 detik tambahan di dalam lock.
  // Sekarang: invalidate saja, next sync dari frontend yang rebuild secara alami.
}

/**
 * Invalidate sync cache for a session (dipanggil saat data berubah)
 */
function invalidateSyncCache_(sessionId) {
  if (sessionId) {
    var cache = CacheService.getScriptCache();
    cache.remove("sync_" + sessionId);
    cache.put("DASH_INVALIDATED", String(Date.now()), 60);
  }
}

// =========================================================================
// getDropdownData() — Tidak berubah, sudah efisien
// =========================================================================
function getDropdownData() {
  try {
    var DB    = getSheetNames_();
    var sheet = getSpreadsheet_().getSheetByName(DB.DATA);
    if (!sheet) return { success: false, message: "Sheet " + DB.DATA + " tidak ditemukan." };

    var maxRow     = cfgNum("Data Max Row") || 100;
    var brandCol   = cfg("Brand Column")   || "A";
    var brandStart = cfgNum("Brand Start Row") || 2;
    var picCol     = cfg("PIC Column")     || "B";
    var picStart   = cfgNum("PIC Start Row")   || 2;
    var vehCol     = cfg("Vehicle Column") || "C";
    var vehStart   = cfgNum("Vehicle Start Row") || 2;

    function readColumn(col, startRow) {
      var values = sheet.getRange(col + startRow + ":" + col + maxRow).getValues();
      var result = [];
      for (var i = 0; i < values.length; i++) {
        var v = values[i][0];
        if (v !== null && v !== undefined && String(v).trim() !== "") {
          result.push(String(v).trim());
        }
      }
      return result;
    }

    return {
      success: true,
      data: {
        brands  : readColumn(brandCol,  brandStart),
        pics    : readColumn(picCol,    picStart),
        vehicles: readColumn(vehCol,    vehStart)
      }
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// =========================================================================
// syncSessionData() — OPTIMASI: Cache TTL 10 → 20 detik
// Cache sudah di-invalidate manual via rebuildAndCacheSync_()
// forceFresh=true dipakai saat UI perlu benar-benar menimpa state lokal
// dengan data terbaru dari sheet, tanpa memakai cache script.
// =========================================================================
function syncSessionData(sessionId, forceFresh, traceId) {
  try {
    perfTrace_("syncSessionData", traceId, "server_start", { sessionId: sessionId, forceFresh: forceFresh === true });
    var cache    = CacheService.getScriptCache();
    var cacheKey = "sync_" + sessionId;
    if (forceFresh !== true) {
      var cached = cache.get(cacheKey);
      if (cached) {
        try { perfTrace_("syncSessionData", traceId, "cache_hit"); return JSON.parse(cached); } catch(e) { cache.remove(cacheKey); }
      }
    }

    var DB  = getSheetNames_();
    var ss  = getSpreadsheet_();
    var tz  = cfg("Timezone") || "Asia/Jakarta";

    var sSession = ss.getSheetByName(DB.SESSIONS);
    var sFleet   = ss.getSheetByName(DB.FLEET);
    var sPhotos  = ss.getSheetByName(DB.PHOTOS);

    var sessionData = null;
    ensureSessionColumns_(sSession);
    perfTrace_("syncSessionData", traceId, "sessions_read_start");
    var sessionRows = sSession.getDataRange().getValues();
    var sessionCol = {};
    if (sessionRows.length > 0) {
      for (var sc = 0; sc < sessionRows[0].length; sc++) {
        sessionCol[String(sessionRows[0][sc] || "").trim()] = sc;
      }
    }
    function sessionValue_(row, name, fallbackIndex) {
      var idx = sessionCol[name];
      if (idx !== undefined) return row[idx];
      return fallbackIndex !== undefined ? row[fallbackIndex] : "";
    }
    perfTrace_("syncSessionData", traceId, "sessions_read_end", { rows: sessionRows.length });
    for (var i = sessionRows.length - 1; i >= 1; i--) {
      if (sessionRows[i][0] === sessionId && sessionRows[i][5] === "ACTIVE") {
        var d = sessionRows[i][1];
        sessionData = {
          id       : sessionId,
          brand    : sessionRows[i][2],
          shipment : sessionRows[i][3],
          pic      : sessionRows[i][4],
          date     : d instanceof Date ? Utilities.formatDate(d, tz, "yyyy-MM-dd") : d,
          createdBy: sessionRows[i][8] || sessionRows[i][4],
          forwarder: sessionRows[i][9] || "",
          lossPreventionName: sessionRows[i][10] || "",
          lossPreventionDate: sessionRows[i][11] instanceof Date ? Utilities.formatDate(sessionRows[i][11], tz, cfg("Date Format") || "dd MMMM yyyy") : String(sessionRows[i][11] || "").trim(),
          lossPreventionFilledBy: sessionRows[i][12] || "",
          lossPreventionSignatureMode: String(sessionValue_(sessionRows[i], "Loss Prevention Signature Mode", 13) || "").trim(),
          lossPreventionSignatureFileId: String(sessionValue_(sessionRows[i], "Loss Prevention Signature File ID", 14) || "").trim(),
          lossPreventionSignatureFileUrl: String(sessionValue_(sessionRows[i], "Loss Prevention Signature File URL", 15) || "").trim(),
          lossPreventionSignatureSignedAt: sessionValue_(sessionRows[i], "Loss Prevention Signature Signed At", 16) instanceof Date
            ? Utilities.formatDate(sessionValue_(sessionRows[i], "Loss Prevention Signature Signed At", 16), tz, cfg("Date Format") || "dd MMMM yyyy")
            : String(sessionValue_(sessionRows[i], "Loss Prevention Signature Signed At", 16) || "").trim()
        };
        break;
      }
    }
    if (!sessionData) return { success: false, message: "Sesi selesai atau tidak ditemukan." };

    var containerIds = {};
    var containers   = [];
    perfTrace_("syncSessionData", traceId, "fleet_read_start");
    var fleetMap = ensureFleetWorkColumns_(sFleet);
    var fleetRows    = sFleet.getDataRange().getValues();
    perfTrace_("syncSessionData", traceId, "fleet_read_end", { rows: fleetRows.length });
    for (var i = 1; i < fleetRows.length; i++) {
      if (fleetRows[i][1] === sessionId && fleetRows[i][8] !== "DELETED") {
        var cId = fleetRows[i][0];
        var lifecycle = buildFleetWorkSnapshot_(fleetRows[i], fleetMap);
        containerIds[cId] = true;
        containers.push({
          id      : cId,
          number  : fleetRows[i][2],
          type    : fleetRows[i][3],
          driver  : fleetRows[i][4],
          plate   : fleetRows[i][5],
          qtyKoli : 0, // Akan dihitung otomatis dari foto palet
          photos  : [],
          createdBy: fleetRows[i][9] || "unknown",
          workStatus: lifecycle.workStatus,
          workStatusText: lifecycle.workStatusText,
          startedAt: lifecycle.startedAt,
          startedAtMs: lifecycle.startedAtMs,
          startedAtText: lifecycle.startedAtText,
          startedBy: lifecycle.startedBy,
          completedAt: lifecycle.completedAt,
          completedAtMs: lifecycle.completedAtMs,
          completedAtText: lifecycle.completedAtText,
          completedBy: lifecycle.completedBy,
          reopenedAt: lifecycle.reopenedAt,
          reopenedAtMs: lifecycle.reopenedAtMs,
          reopenedAtText: lifecycle.reopenedAtText,
          reopenedBy: lifecycle.reopenedBy,
          durationMs: lifecycle.durationMs,
          durationText: lifecycle.durationText
        });
      }
    }

    if (containers.length > 0) {
      perfTrace_("syncSessionData", traceId, "photos_read_start");
      var allPhotos = sPhotos.getDataRange().getValues();
      perfTrace_("syncSessionData", traceId, "photos_read_end", { rows: allPhotos.length });
      var photoMap  = {};
      var seenSyncPhotoIds = {};
      for (var i = 1; i < allPhotos.length; i++) {
        // PATCH: skip baris kosong hasil clearContent
        var syncPhotoId = String(allPhotos[i][0] || "").trim();
        if (!syncPhotoId) continue;
        if (seenSyncPhotoIds[syncPhotoId]) continue;
        seenSyncPhotoIds[syncPhotoId] = true;
        var cId = allPhotos[i][1];
        if (containerIds[cId]) {
          if (!photoMap[cId]) photoMap[cId] = [];
          var photoUrl = allPhotos[i][3];
          var oldMatch = String(photoUrl).match(/[?&]id=([^&]+)/);
          if (oldMatch && oldMatch[1] && String(photoUrl).indexOf("uc?export=view") >= 0) {
            photoUrl = buildPhotoUrl_(oldMatch[1]);
          }
          photoMap[cId].push({
            id: syncPhotoId,
            type: allPhotos[i][2],
            dataUrl: photoUrl,
            paletName: allPhotos[i][6] || "",
            paletKoli: Number(allPhotos[i][7]) || 0,
            fileName: allPhotos[i][8] || ""
          });
        }
      }
      
      // OTOMATIS: Hitung qtyKoli dari foto palet per armada
      var totalKoli = 0;
      containers.forEach(function(c) { 
        c.photos = photoMap[c.id] || []; 
        var cKoli = 0;
        c.photos.forEach(function(p) {
          var pName = String((p && p.paletName) || "").trim();
          var pKoli = Number((p && p.paletKoli) || 0) || 0;
          if (normalizePhotoType_(p && p.type) === "palet" && pName && pKoli > 0) {
            cKoli += pKoli;
          }
        });
        c.qtyKoli = cKoli;
        totalKoli += cKoli;
      });
      sessionData.totalKoli = totalKoli;
    }

    var creatorAliases = {};
    function addCreatorAlias_(value) {
      var key = String(value || "").trim().toLowerCase();
      if (key) creatorAliases[key] = true;
    }

    addCreatorAlias_(sessionData.createdBy);
    addCreatorAlias_(sessionData.pic);
    sessionData.createdByFullName = sessionData.pic || sessionData.createdBy;

    var accountSheet = ss.getSheetByName(DB.ACCOUNTS);
    if (accountSheet) {
      try {
        var accountRows = accountSheet.getDataRange().getValues();
        for (var a = 1; a < accountRows.length; a++) {
          var username = String(accountRows[a][0] || "").trim();
          var fullName = String(accountRows[a][2] || "").trim();
          if (!username && !fullName) continue;
          var userKey = username.toLowerCase();
          var fullKey = fullName.toLowerCase();
          if ((userKey && creatorAliases[userKey]) || (fullKey && creatorAliases[fullKey])) {
            addCreatorAlias_(username);
            addCreatorAlias_(fullName);
            if (fullName) sessionData.createdByFullName = fullName;
          }
        }
      } catch (accountErr) {}
    }

    var activeUsers = [];
    var userSheet   = ss.getSheetByName(DB.USERS);
    if (userSheet) {
      var userRows = userSheet.getDataRange().getValues();
      for (var i = 1; i < userRows.length; i++) {
        if (userRows[i][1] === sessionId) {
          var userName = String(userRows[i][2] || "").trim();
          activeUsers.push({
            name: userName,
            status: userRows[i][3],
            isAdmin: !!creatorAliases[userName.toLowerCase()]
          });
        }
      }
    }

    if (activeUsers.length > 0 && !activeUsers.some(function(u) { return u.isAdmin; })) {
      activeUsers[0].isAdmin = true;
    }

    var sjQtyStatus = { status: "EMPTY", statusText: "SJ Qty belum diisi", subText: "", canHandover: false, sjCount: 0, actualCount: 0 };
    try {
      if (typeof getSjQtyStatus === "function") {
        var sjRes = getSjQtyStatus(sessionId);
        if (sjRes && sjRes.success && sjRes.data) sjQtyStatus = sjRes.data;
      }
    } catch (sjErr) {
      sjQtyStatus = { status: "ERROR", statusText: "SJ Qty gagal dibaca", subText: sjErr.toString(), canHandover: false, sjCount: 0, actualCount: 0 };
    }

    var result   = { success: true, data: { session: sessionData, containers: containers, users: activeUsers, sjQty: sjQtyStatus } };
    // OPTIMASI: Cache 20 detik (naik dari 10). Aman karena invalidation manual
    var syncTTL  = cfgNum("Cache Sync (detik)") || 20;
    try { cache.put(cacheKey, JSON.stringify(result), syncTTL); } catch(e) {}
    perfTrace_("syncSessionData", traceId, "server_end", { success: true, containers: containers.length });
    return result;

  } catch (e) {
    perfTrace_("syncSessionData", traceId, "server_error", { message: e.toString() });
    return { success: false, message: e.toString() };
  }
}

// =========================================================================
// getActiveSessions() — Tidak berubah
// =========================================================================
function getActiveSessions() {
  try {
    var DB    = getSheetNames_();
    var sheet = getSpreadsheet_().getSheetByName(DB.SESSIONS);
    var data  = sheet.getDataRange().getValues();
    var tz    = cfg("Timezone") || "Asia/Jakarta";
    var sessions = [];

    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][5] === "ACTIVE") {
        var d = data[i][1];
        sessions.push({
          id       : data[i][0],
          date     : d instanceof Date ? Utilities.formatDate(d, tz, "yyyy-MM-dd") : d,
          brand    : data[i][2],
          shipment : data[i][3],
          pic      : data[i][4],
          createdBy: data[i][8] || data[i][4]
        });
      }
    }
    return { success: true, data: sessions };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// =========================================================================
// getDashboardData() — OPTIMASI: Cached 15 detik + efficient loops
// Sebelumnya: 5-6x getDataRange setiap navigasi = 2-4 detik
// Sekarang: dari cache = <50ms (jika belum expired/invalidated)
// =========================================================================
function getDashboardData(username, forceFresh, traceId) {
  try {
    perfTrace_("loadDashboard", traceId, "server_start", { username: username, forceFresh: forceFresh === true });
    var cache = CacheService.getScriptCache();
    var cacheKey = DASHBOARD_CACHE_KEY + (username || "").toLowerCase();

    // Cek apakah ada data valid di cache
    var cached = cache.get(cacheKey);
    var invalidatedTs = cache.get("DASH_INVALIDATED");
    var cachedTs = cache.get(cacheKey + "_ts");

    if (forceFresh !== true && cached && (!invalidatedTs || (cachedTs && Number(cachedTs) > Number(invalidatedTs)))) {
      try {
        perfTrace_("loadDashboard", traceId, "cache_hit");
        return JSON.parse(cached);
      } catch(e) { cache.remove(cacheKey); }
    }

    var DB  = getSheetNames_();
    var ss  = getSpreadsheet_();
    var tz  = cfg("Timezone") || "Asia/Jakarta";

    var sessionSheet = ss.getSheetByName(DB.SESSIONS);
    var fleetSheet   = ss.getSheetByName(DB.FLEET);
    var photoSheet   = ss.getSheetByName(DB.PHOTOS);
    var userSheet    = ss.getSheetByName(DB.USERS);
    var accountSheet = ss.getSheetByName(DB.ACCOUNTS);

    var sessionRows = [], fleetRows = [], userRows = [], photoRows = [], fullNameMap = {};

    // Wrap setiap sheet read dalam try-catch
    try { perfTrace_("loadDashboard", traceId, "sessions_read_start"); sessionRows = sessionSheet.getDataRange().getValues(); perfTrace_("loadDashboard", traceId, "sessions_read_end", { rows: sessionRows.length }); } catch(e) { console.warn("Gagal baca Sessions:", e); }
    try { perfTrace_("loadDashboard", traceId, "fleet_read_start"); fleetRows = fleetSheet.getDataRange().getValues(); perfTrace_("loadDashboard", traceId, "fleet_read_end", { rows: fleetRows.length }); } catch(e) { console.warn("Gagal baca Fleet:", e); }
    try { if (userSheet) { perfTrace_("loadDashboard", traceId, "users_read_start"); userRows = userSheet.getDataRange().getValues(); perfTrace_("loadDashboard", traceId, "users_read_end", { rows: userRows.length }); } } catch(e) { console.warn("Gagal baca Users:", e); }
    try { perfTrace_("loadDashboard", traceId, "photos_read_start"); photoRows = photoSheet.getDataRange().getValues(); perfTrace_("loadDashboard", traceId, "photos_read_end", { rows: photoRows.length }); } catch(e) { console.warn("Gagal baca Photos:", e); }
    try {
      if (accountSheet) {
        perfTrace_("loadDashboard", traceId, "accounts_read_start");
        var accRows = accountSheet.getDataRange().getValues();
        perfTrace_("loadDashboard", traceId, "accounts_read_end", { rows: accRows.length });
        for (var i = 1; i < accRows.length; i++) {
          var uname = String(accRows[i][0]).trim().toLowerCase();
          var fname = String(accRows[i][2] || "").trim();
          if (uname && fname) fullNameMap[uname] = fname;
        }
      }
    } catch(e) { console.warn("Gagal baca Accounts:", e); }

    // Cari semua sesi yang user ini ikut
    var mySessionIds = {};
    var myStatusMap  = {};
    for (var i = userRows.length - 1; i >= 1; i--) {
      var uName = String(userRows[i][2]).trim().toLowerCase();
      var sid   = userRows[i][1];
      var uStat = userRows[i][3];
      if (uName === username.toLowerCase() && (uStat === "WORKING" || uStat === "DONE")) {
        if (!mySessionIds[sid]) {
          mySessionIds[sid]  = true;
          myStatusMap[sid]   = uStat;
        }
      }
    }

    // Cek createdBy di Sessions
    var createdBySessionIds = {};
    for (var i = 1; i < sessionRows.length; i++) {
      if (sessionRows[i][5] === "ACTIVE") {
        var creator = String(sessionRows[i][8] || "").trim().toLowerCase();
        if (creator === username.toLowerCase()) {
          createdBySessionIds[sessionRows[i][0]] = true;
          if (!mySessionIds[sessionRows[i][0]]) {
            mySessionIds[sessionRows[i][0]] = true;
            myStatusMap[sessionRows[i][0]] = "WORKING";
          }
        }
      }
    }

    // OPTIMASI: Single-pass fleet count + container→session map
    var fleetCountMap = {};
    var totalKoliMap = {};
    var containerSessionMap = {};
    for (var i = 1; i < fleetRows.length; i++) {
      if (fleetRows[i][8] !== "DELETED") {
        var sid = fleetRows[i][1];
        fleetCountMap[sid] = (fleetCountMap[sid] || 0) + 1;
        var qtyKoli = Number(fleetRows[i][10]);
        if (isNaN(qtyKoli) || qtyKoli < 0) qtyKoli = 0;
        totalKoliMap[sid] = (totalKoliMap[sid] || 0) + Math.floor(qtyKoli);
        containerSessionMap[fleetRows[i][0]] = sid;
      }
    }

    // Foto per sesi
    var photoCountMap = {};
    var seenDashboardPhotoIds = {};
    for (var i = 1; i < photoRows.length; i++) {
      // PATCH: skip baris kosong hasil clearContent
      var dashboardPhotoId = String(photoRows[i][0] || "").trim();
      if (!dashboardPhotoId) continue;
      if (seenDashboardPhotoIds[dashboardPhotoId]) continue;
      seenDashboardPhotoIds[dashboardPhotoId] = true;
      var cid = photoRows[i][1];
      var sid = containerSessionMap[cid];
      if (sid) photoCountMap[sid] = (photoCountMap[sid] || 0) + 1;
    }

    // User per sesi
    var userCountMap = {};
    var userNamesMap = {};
    for (var i = 1; i < userRows.length; i++) {
      var sid = userRows[i][1];
      var uStatus = userRows[i][3];
      if (uStatus === "WORKING" || uStatus === "DONE") {
        userCountMap[sid] = (userCountMap[sid] || 0) + 1;
        if (!userNamesMap[sid]) userNamesMap[sid] = [];
        var rawName = String(userRows[i][2]).trim();
        userNamesMap[sid].push({ name: rawName, status: uStatus });
      }
    }

    // Foto hari ini (timezone-safe)
    var todayStr   = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    var todayPhotos = 0;
    var seenTodayPhotoIds = {};
    for (var i = 1; i < photoRows.length; i++) {
      // PATCH: skip baris kosong hasil clearContent
      var todayPhotoId = String(photoRows[i][0] || "").trim();
      if (!todayPhotoId) continue;
      if (seenTodayPhotoIds[todayPhotoId]) continue;
      seenTodayPhotoIds[todayPhotoId] = true;
      var photoDate = photoRows[i][4];
      if (photoDate instanceof Date) {
        var photoDateStr = Utilities.formatDate(photoDate, tz, "yyyy-MM-dd");
        if (photoDateStr === todayStr) todayPhotos++;
      }
    }

    // Build session lists
    var mySessions    = [];
    var otherSessions = [];

    for (var i = sessionRows.length - 1; i >= 1; i--) {
      if (sessionRows[i][5] !== "ACTIVE") continue;

      var sid      = sessionRows[i][0];
      var rawDate  = sessionRows[i][1];
      var picUser  = String(sessionRows[i][4] || "").trim();
      var picFullName = fullNameMap[picUser.toLowerCase()] || picUser;
      var createdByUser = String(sessionRows[i][8] || picUser).trim();
      var createdByFullName = fullNameMap[createdByUser.toLowerCase()] || createdByUser;
      var isCreator = createdBySessionIds[sid] || false;

      var sessObj = {
        id              : sid,
        date            : rawDate instanceof Date ? Utilities.formatDate(rawDate, tz, "yyyy-MM-dd") : rawDate,
        brand           : sessionRows[i][2],
        shipment        : sessionRows[i][3],
        pic             : picFullName,
        picUsername      : picUser,
        createdBy       : createdByUser,
        createdByFullName: createdByFullName,
        isCreator       : isCreator,
        fleetCount      : fleetCountMap[sid]  || 0,
        totalKoli       : totalKoliMap[sid]   || 0,
        photoCount      : photoCountMap[sid]  || 0,
        userCount       : userCountMap[sid]   || 0,
        userNames       : userNamesMap[sid]   || []
      };

      if (mySessionIds[sid]) {
        sessObj.myStatus = myStatusMap[sid] || "WORKING";
        mySessions.push(sessObj);
      } else {
        otherSessions.push(sessObj);
      }
    }

    var totalActiveSessions = mySessions.length + otherSessions.length;
    var totalFleet = 0;
    var grandTotalKoli = 0;
    for (var sid in fleetCountMap) { totalFleet += fleetCountMap[sid]; }
    for (var sid2 in totalKoliMap) { grandTotalKoli += totalKoliMap[sid2]; }

    var result = {
      success: true,
      data: {
        mySessions   : mySessions,
        otherSessions: otherSessions,
        mySession    : mySessions.length > 0 ? mySessions[0] : null,
        stats: {
          activeSessions: totalActiveSessions,
          totalFleet    : totalFleet,
          todayPhotos   : todayPhotos,
          totalKoli     : grandTotalKoli
        }
      }
    };

    // OPTIMASI: Cache dashboard result 15 detik
    try {
      var resultStr = JSON.stringify(result);
      if (resultStr.length < 90000) { // CacheService limit ~100KB
        cache.put(cacheKey, resultStr, DASHBOARD_CACHE_TTL);
        cache.put(cacheKey + "_ts", String(Date.now()), DASHBOARD_CACHE_TTL);
      }
    } catch(e) {}

    perfTrace_("loadDashboard", traceId, "server_end", { success: true, activeSessions: totalActiveSessions });
    return result;
  } catch (e) {
    perfTrace_("loadDashboard", traceId, "server_error", { message: e.toString() });
    return { success: false, message: e.toString() };
  }
}

// =========================================================================
// getCompletedSessions() — Tidak berubah, sudah pakai pagination
// =========================================================================
function getCompletedSessions(page, pageSize, traceId) {
  try {
    perfTrace_("loadHistory", traceId, "server_start", { page: page || 1 });
    page     = page     || 1;
    pageSize = pageSize || 20;

    var DB    = getSheetNames_();
    var ss    = getSpreadsheet_();
    var sheet = ss.getSheetByName(DB.REKAP);
    if (!sheet) return { success: true, data: [], hasMore: false, total: 0 };

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, data: [], hasMore: false, total: 0 };

    perfTrace_("loadHistory", traceId, "header_read_start");
    var header     = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    perfTrace_("loadHistory", traceId, "header_read_end");
    var colIdx     = {};
    for (var c = 0; c < header.length; c++) {
      colIdx[String(header[c]).trim()] = c;
    }

    var iDate     = colIdx["Tanggal"]            !== undefined ? colIdx["Tanggal"]            : 0;
    var iBrand    = colIdx["Brand"]              !== undefined ? colIdx["Brand"]              : 1;
    var iShipment = colIdx["Shipment"]           !== undefined ? colIdx["Shipment"]           : 2;
    var iSesiId   = colIdx["Sesi ID"]            !== undefined ? colIdx["Sesi ID"]            : 3;
    var iFolder   = colIdx["Link Folder Armada"] !== undefined ? colIdx["Link Folder Armada"] : -1;

    var totalDataRows = lastRow - 1;
    perfTrace_("loadHistory", traceId, "sheet_read_start", { rows: totalDataRows });
    var data = sheet.getRange(2, 1, totalDataRows, sheet.getLastColumn()).getValues();
    perfTrace_("loadHistory", traceId, "sheet_read_end", { rows: data.length });

    var sessionMap   = {};
    var sessionOrder = [];

    for (var i = data.length - 1; i >= 0; i--) {
      var row = data[i];
      var sid = iSesiId >= 0 ? row[iSesiId] : null;
      if (!sid) continue;

      if (!sessionMap[sid]) {
        var d = row[iDate];
        sessionMap[sid] = {
          id        : sid,
          date      : d instanceof Date ? d.toISOString() : d,
          brand     : row[iBrand]    || "",
          shipment  : row[iShipment] || "",
          fleetCount: 0,
          folderUrl : iFolder >= 0 ? (row[iFolder] || "") : ""
        };
        sessionOrder.push(sid);
      }
      sessionMap[sid].fleetCount++;
    }

    try {
      function mergeSessionMetadataFromSheet_(sessionSheet, sourceName) {
        if (!sessionSheet) return;
        ensureSessionColumns_(sessionSheet);
        var sessionRows = sessionSheet.getDataRange().getValues();
        var sessionHeaders = sessionRows.length > 0 ? sessionRows[0] : [];
        var sessionCol = {};
        for (var sc = 0; sc < sessionHeaders.length; sc++) {
          sessionCol[String(sessionHeaders[sc] || "").trim()] = sc;
        }
        function cell_(row, name) {
          var idx = sessionCol[name];
          return idx !== undefined ? row[idx] : "";
        }
        function dateCell_(row, name) {
          var value = cell_(row, name);
          return value instanceof Date ? value.toISOString() : (value || "");
        }
        for (var sr = 1; sr < sessionRows.length; sr++) {
          var historySessionId = String(sessionRows[sr][0] || "").trim();
          if (!historySessionId || !sessionMap[historySessionId]) continue;
          var sessionFolderUrl = cell_(sessionRows[sr], "Folder URL") || "";
          sessionMap[historySessionId].sessionFolderUrl = sessionFolderUrl;
          if (!sessionMap[historySessionId].folderUrl && sessionFolderUrl) {
            sessionMap[historySessionId].folderUrl = sessionFolderUrl;
          }
          sessionMap[historySessionId].handoverFileId = cell_(sessionRows[sr], "Handover File ID") || "";
          sessionMap[historySessionId].handoverFileUrl = cell_(sessionRows[sr], "Handover File URL") || "";
          sessionMap[historySessionId].handoverRevisionId = cell_(sessionRows[sr], "Handover Current Revision ID") || "";
          sessionMap[historySessionId].handoverVersionNo = Number(cell_(sessionRows[sr], "Handover Version No") || 0) || 0;
          sessionMap[historySessionId].handoverGeneratedAt = dateCell_(sessionRows[sr], "Handover Generated At");
          sessionMap[historySessionId].rekapStatus = cell_(sessionRows[sr], "Rekap Status") || "";
          sessionMap[historySessionId].rekapError = cell_(sessionRows[sr], "Rekap Error") || "";
          sessionMap[historySessionId].handoverPdfStatus = cell_(sessionRows[sr], "Handover PDF Status") || "";
          sessionMap[historySessionId].handoverPdfError = cell_(sessionRows[sr], "Handover PDF Error") || "";
          sessionMap[historySessionId].waStatus = cell_(sessionRows[sr], "WA Status") || "";
          sessionMap[historySessionId].waError = cell_(sessionRows[sr], "WA Error") || "";
          sessionMap[historySessionId].artifactUpdatedAt = dateCell_(sessionRows[sr], "Artifact Updated At");
          sessionMap[historySessionId].isArchived = sourceName === "archive";
        }
      }

      mergeSessionMetadataFromSheet_(ss.getSheetByName(DB.SESSIONS + "_Archive"), "archive");
      mergeSessionMetadataFromSheet_(ss.getSheetByName(DB.SESSIONS), "active");
    } catch (metaErr) {
      console.warn("getCompletedSessions: handover metadata gagal dibaca: " + metaErr);
    }

    var total  = sessionOrder.length;
    var start  = (page - 1) * pageSize;
    var end    = Math.min(start + pageSize, total);
    var sessions = [];
    for (var i = start; i < end; i++) sessions.push(sessionMap[sessionOrder[i]]);

    perfTrace_("loadHistory", traceId, "server_end", { success: true, count: sessions.length, total: total });
    return { success: true, data: sessions, hasMore: end < total, total: total, page: page };

  } catch (e) {
    perfTrace_("loadHistory", traceId, "server_error", { message: e.toString() });
    return { success: false, message: e.toString() };
  }
}

// =========================================================================
// getHandOverData() — Data lengkap untuk Print Hand Over
// Menggabungkan data dari Sessions, Fleet, Photos, Users, dan Rekap
// untuk menghasilkan laporan hand over yang komprehensif.
// =========================================================================
function getHandOverData(sessionId, traceId) {
  try {
    perfTrace_("printHandover", traceId, "handover_data_start", { sessionId: sessionId });
    if (!sessionId) return { success: false, message: "Session ID diperlukan." };

    var DB  = getSheetNames_();
    var ss  = getSpreadsheet_();
    var tz  = cfg("Timezone") || "Asia/Jakarta";
    var df  = cfg("Date Format") || "dd MMMM yyyy";

    // --- Session data ---
    var sessionRows = [];
    var sessionCol = {};
    var sessionRow = null;
    var sessionFromArchive = false;

    function loadHandoverSessionCandidate_(sheetName, archived) {
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) return false;
      ensureSessionColumns_(sheet);
      perfTrace_("printHandover", traceId, "session_sheet_read_start", { sheet: sheetName });
      var rows = sheet.getDataRange().getValues();
      perfTrace_("printHandover", traceId, "session_sheet_read_end", { sheet: sheetName, rows: rows.length });
      var cols = {};
      if (rows.length > 0) {
        for (var sc = 0; sc < rows[0].length; sc++) {
          cols[String(rows[0][sc] || "").trim()] = sc;
        }
      }
      for (var ri = 1; ri < rows.length; ri++) {
        if (String(rows[ri][0] || "").trim() === String(sessionId || "").trim()) {
          sessionRows = rows;
          sessionCol = cols;
          sessionRow = rows[ri];
          sessionFromArchive = archived === true;
          return true;
        }
      }
      return false;
    }

    loadHandoverSessionCandidate_(DB.SESSIONS, false);
    if (!sessionRow) loadHandoverSessionCandidate_(DB.SESSIONS + "_Archive", true);
    if (!sessionRow) return { success: false, message: "Sesi tidak ditemukan." };

    if (!sessionFromArchive) {
      var sjReady = typeof ensureSjQtyReadyForHandover_ === "function"
        ? ensureSjQtyReadyForHandover_(sessionId, traceId)
        : (typeof validateSjQtyMatching === "function" ? (function() {
            var sjValidation = validateSjQtyMatching(sessionId);
            return sjValidation && sjValidation.success && sjValidation.data && sjValidation.data.canHandover
              ? { success: true, data: sjValidation.data }
              : { success: false, message: "SJ Qty belum valid.", data: sjValidation && sjValidation.data ? sjValidation.data : null };
          })() : null);
      if (!sjReady || !sjReady.success) {
        return {
          success: false,
          code: "SJ_QTY_INVALID",
          message: sjReady && sjReady.message ? sjReady.message : "SJ Qty belum valid.",
          data: sjReady && (sjReady.sjQty || sjReady.data) ? (sjReady.sjQty || sjReady.data) : null
        };
      }
    }

    var rawDate = sessionRow[1];
    function sessionValue_(name, fallbackIndex) {
      var idx = sessionCol[name];
      if (idx !== undefined) return sessionRow[idx];
      return fallbackIndex !== undefined ? sessionRow[fallbackIndex] : "";
    }
    var dateStr = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, tz, df)
      : String(rawDate);
    var inboundDateStr = "";
    if (rawDate instanceof Date) {
      inboundDateStr = Utilities.formatDate(rawDate, tz, "yyyy-MM-dd");
    } else {
      var rawDateText = String(rawDate || "").replace(/\s+/g, " ").trim();
      var parsedInboundDate = rawDateText ? new Date(rawDateText) : null;
      inboundDateStr = parsedInboundDate && !isNaN(parsedInboundDate.getTime())
        ? Utilities.formatDate(parsedInboundDate, tz, "yyyy-MM-dd")
        : rawDateText;
    }
    var createdAt = sessionRow[6] instanceof Date
      ? Utilities.formatDate(sessionRow[6], tz, "dd MMM yyyy HH:mm")
      : String(sessionRow[6] || "");

    var session = {
      id       : sessionId,
      brand    : sessionRow[2] || "",
      shipment : sessionRow[3] || "",
      pic      : sessionRow[4] || "",
      date     : dateStr,
      inboundDate: inboundDateStr,
      status   : sessionRow[5] || "COMPLETED",
      isArchived: sessionFromArchive,
      createdBy: sessionRow[8] || sessionRow[4] || "",
      createdAt: createdAt,
      folderUrl: sessionRow[7] || "",
      forwarder: sessionRow[9] || "-",
      lossPreventionName: sessionRow[10] || "",
      lossPreventionDate: sessionRow[11] instanceof Date ? Utilities.formatDate(sessionRow[11], tz, df) : String(sessionRow[11] || "").trim(),
      lossPreventionFilledBy: sessionRow[12] || "",
      lossPreventionSignatureMode: String(sessionValue_("Loss Prevention Signature Mode", 13) || "").trim(),
      lossPreventionSignatureFileId: String(sessionValue_("Loss Prevention Signature File ID", 14) || "").trim(),
      lossPreventionSignatureFileUrl: String(sessionValue_("Loss Prevention Signature File URL", 15) || "").trim(),
      lossPreventionSignatureSignedAt: sessionValue_("Loss Prevention Signature Signed At", 16) instanceof Date
        ? Utilities.formatDate(sessionValue_("Loss Prevention Signature Signed At", 16), tz, df)
        : String(sessionValue_("Loss Prevention Signature Signed At", 16) || "").trim()
    };

    var accountSheet = ss.getSheetByName(DB.ACCOUNTS);
    if (accountSheet && typeof ensureAccountColumns_ === "function") {
      ensureAccountColumns_(accountSheet);
    }

    session.picSignatureFileId = getAccountSignatureFileId_(accountSheet, [
      session.pic,
      session.createdBy
    ]);
    session.picSignatureDataUri = getDriveImageDataUri_(session.picSignatureFileId);
    session.lossPreventionSignatureDataUri = getDriveImageDataUri_(session.lossPreventionSignatureFileId);
    if (
      String(session.lossPreventionSignatureMode || "").toUpperCase() === "DIGITAL" &&
      !session.lossPreventionSignatureDataUri
    ) {
      return {
        success: false,
        message: "Tanda tangan LP digital belum bisa dibaca dari Drive. Simpan ulang tanda tangan LP sebelum cetak/update PDF handover."
      };
    }

    var logoFileId = cfg("Handover Logo File ID") || "";
    var logoDataUri = getDriveImageDataUri_(logoFileId);
    var archiveSuffix = sessionFromArchive ? "_Archive" : "";

    // --- Fleet data ---
    var sFleet = ss.getSheetByName(DB.FLEET + archiveSuffix) || ss.getSheetByName(DB.FLEET);
    if (!sFleet) return { success: false, message: "Sheet Fleet tidak ditemukan." };
    perfTrace_("printHandover", traceId, "fleet_sheet_read_start");
    var fleetRows = sFleet.getDataRange().getValues();
    perfTrace_("printHandover", traceId, "fleet_sheet_read_end", { rows: fleetRows.length });
    var fleets = [];
    var totalKoli = 0;
    var sessionStartMs = 0;
    var sessionEndMs = 0;
    for (var i = 1; i < fleetRows.length; i++) {
      if (fleetRows[i][1] === sessionId && fleetRows[i][8] !== "DELETED") {
        var qtyKoli = Number(fleetRows[i][10]);
        if (isNaN(qtyKoli) || qtyKoli < 0) qtyKoli = 0;
        qtyKoli = Math.floor(qtyKoli);
        totalKoli += qtyKoli;
        var fleetStartMs = toTimestampMs_(fleetRows[i][6]);
        if (fleetStartMs > 0 && (sessionStartMs === 0 || fleetStartMs < sessionStartMs)) {
          sessionStartMs = fleetStartMs;
        }
        if (fleetStartMs > sessionEndMs) sessionEndMs = fleetStartMs;
        fleets.push({
          id       : fleetRows[i][0],
          number   : fleetRows[i][2],
          type     : fleetRows[i][3] || "",
          driver   : fleetRows[i][4] || "",
          plate    : fleetRows[i][5] || "",
          qtyKoli  : qtyKoli,
          createdAt: fleetRows[i][6] instanceof Date
            ? Utilities.formatDate(fleetRows[i][6], tz, "HH:mm")
            : ""
        });
      }
    }

    // --- Photo data per fleet ---
    var sPhotos = ss.getSheetByName(DB.PHOTOS + archiveSuffix) || ss.getSheetByName(DB.PHOTOS);
    if (!sPhotos) return { success: false, message: "Sheet PhotoLogs tidak ditemukan." };
    perfTrace_("printHandover", traceId, "photo_sheet_read_start");
    var photoRows = sPhotos.getDataRange().getValues();
    perfTrace_("printHandover", traceId, "photo_sheet_read_end", { rows: photoRows.length });
    var fleetIdSet = {};
    fleets.forEach(function(f) { fleetIdSet[f.id] = true; });

    var photoCounts = {};
    var paletDetails = {};
    var sjQtyMap = typeof getSjQtyMap_ === "function" ? getSjQtyMap_(sessionId) : {};

    var seenHandoverPhotoIds = {};
    for (var i = 1; i < photoRows.length; i++) {
      var handoverPhotoId = String(photoRows[i][0] || "").trim();
      if (!handoverPhotoId) continue;
      if (seenHandoverPhotoIds[handoverPhotoId]) continue;
      seenHandoverPhotoIds[handoverPhotoId] = true;
      var cId = photoRows[i][1];
      if (!fleetIdSet[cId]) continue;

      if (!photoCounts[cId]) photoCounts[cId] = { dokumen: 0, palet: 0, suratjalan: 0 };
      var pType = String(photoRows[i][2] || "").toLowerCase().trim();
      if (pType === "dokumen") photoCounts[cId].dokumen++;
      else if (pType === "palet") {
        photoCounts[cId].palet++;
        var pName = String(photoRows[i][6] || "").trim();
        var pKoli = Number(photoRows[i][7]) || 0;
        if (pName && pKoli > 0) {
          var pKey = typeof getSjPalletKey_ === "function" ? getSjPalletKey_(pName) : String(pName || "").trim().toUpperCase();
          var sjRow = sjQtyMap[pKey] || null;
          var sjQty = sjRow ? Number(sjRow.sjQty || 0) : "";
          var varianceQty = sjRow ? (pKoli - sjQty) : "";
          if (!paletDetails[cId]) paletDetails[cId] = [];
          paletDetails[cId].push({
            name: pName,
            displayName: formatHandoverPaletLabel_(pName),
            palletKey: pKey,
            koli: pKoli,
            actualQty: pKoli,
            sjQty: sjQty,
            variance: varianceQty
          });
        }
      }
      else if (pType === "suratjalan") photoCounts[cId].suratjalan++;

      var photoTimeMs = toTimestampMs_(photoRows[i][4]);
      if (photoTimeMs > sessionEndMs) sessionEndMs = photoTimeMs;
    }

    // Attach photo counts & palet details to fleets
    var totalDokumen = 0, totalPalet = 0, totalSuratJalan = 0;
    var calculatedTotalKoli = 0;
    var totalSjQty = 0;
    var totalVariance = 0;
    fleets.forEach(function(f) {
      var pc = photoCounts[f.id] || { dokumen: 0, palet: 0, suratjalan: 0 };
      f.dokumenCount = pc.dokumen;
      f.paletCount   = pc.palet;
      f.suratJalanCount = pc.suratjalan;
      f.paletDetails = paletDetails[f.id] || [];
      
      if (f.paletDetails.length > 0) {
        f.paletDetails.forEach(function(p) {
          calculatedTotalKoli += p.koli;
          totalSjQty += Number(p.sjQty || 0);
          totalVariance += Number(p.variance || 0);
        });
      } else {
        calculatedTotalKoli += f.qtyKoli;
      }

      totalDokumen   += pc.dokumen;
      totalPalet     += pc.palet;
      totalSuratJalan += pc.suratjalan;
    });

    // --- Users / Team ---
    var userSheet = ss.getSheetByName(DB.USERS + archiveSuffix) || ss.getSheetByName(DB.USERS);
    var team = [];
    if (userSheet) {
      var uRows = userSheet.getDataRange().getValues();
      for (var i = 1; i < uRows.length; i++) {
        if (uRows[i][1] === sessionId) {
          team.push({
            name  : uRows[i][2] || "",
            status: uRows[i][3] || ""
          });
        }
      }
    }

    // --- Duration ---
    var startMs = sessionStartMs || toTimestampMs_(sessionRow[6]);
    var completedAt = "";
    var endMs = sessionEndMs || startMs || Date.now();
    var durationStr = formatDurationRange_(startMs, endMs, "-");
    completedAt = Utilities.formatDate(new Date(endMs), tz, "dd MMM yyyy HH:mm");

    return {
      success: true,
      data: {
        session     : session,
        fleets      : fleets,
        team        : team,
        summary     : {
          totalFleet     : fleets.length,
          totalKoli      : calculatedTotalKoli,
          totalActualQty : calculatedTotalKoli,
          totalSjQty     : totalSjQty,
          totalVariance  : totalVariance,
          totalDokumen   : totalDokumen,
          totalPalet     : totalPalet,
          totalSuratJalan: totalSuratJalan,
          totalPhoto     : totalDokumen + totalPalet + totalSuratJalan,
          duration       : durationStr.trim(),
          completedAt    : completedAt
        },
        config      : {
          title         : cfg("Handover Report Title") || "UNLOADING REPORT B2B",
          logoFileId    : logoFileId,
          logoDataUri   : logoDataUri,
          signatureLeft : cfg("Handover Signature Left") || "Pic Unloading",
          signatureRight: cfg("Handover Signature Right") || "Loss Prevention"
        }
      }
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getAccountSignatureFileId_(accountSheet, aliases) {
  if (!accountSheet) return "";
  var aliasMap = {};
  for (var i = 0; i < (aliases || []).length; i++) {
    var key = String(aliases[i] || "").trim().toLowerCase();
    if (key) aliasMap[key] = true;
  }
  if (Object.keys(aliasMap).length === 0) return "";

  var rows = accountSheet.getDataRange().getValues();
  for (var r = 1; r < rows.length; r++) {
    var username = String(rows[r][0] || "").trim().toLowerCase();
    var fullName = String(rows[r][2] || "").trim().toLowerCase();
    if (!aliasMap[username] && !aliasMap[fullName]) continue;
    return normalizeDriveFileId_(rows[r][4]);
  }
  return "";
}

function getDriveImageDataUri_(fileId) {
  fileId = normalizeDriveFileId_(fileId);
  if (!fileId) return "";
  try {
    var blob = DriveApp.getFileById(fileId).getBlob();
    var contentType = blob.getContentType() || "image/png";
    return "data:" + contentType + ";base64," + Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    console.warn("getDriveImageDataUri_ gagal untuk " + fileId + ": " + e);
    return "";
  }
}

function normalizeDriveFileId_(value) {
  var text = String(value || "").trim();
  if (!text) return "";
  var match = text.match(/[-\w]{25,}/);
  return match ? match[0] : text;
}

/**
 * Normalisasi tipe foto untuk pengecekan kategori (dokumen/palet/suratjalan)
 */
function normalizePhotoType_(type) {
  if (!type) return "dokumen";
  var t = String(type).toLowerCase().trim();
  if (t.indexOf("dokumen") >= 0) return "dokumen";
  if (t.indexOf("palet") >= 0) return "palet";
  if (t.indexOf("surat") >= 0) return "suratjalan";
  return "dokumen";
}

function formatHandoverPaletLabel_(name) {
  var clean = String(name || "").trim();
  if (!clean) return "-";
  if (/^palet\b/i.test(clean) || /^pallet\b/i.test(clean)) return clean;
  return "Palet " + clean;
}

/**
 * Mengkalkulasi total Koli dari sheet Photos untuk setiap armada dalam satu sesi,
 * lalu menuliskan nilainya ke kolom Qty Koli (kolom 11) pada sheet Fleet.
 */
function syncFleetKoliFromPhotos_(sessionId, options) {
  if (!sessionId) return;
  var lock = LockService.getScriptLock();
  var hasLock = false;
  try {
    if (!options || !options.skipLock) {
      lock.waitLock(15000);
      hasLock = true;
    }
    var DB = getSheetNames_();
    var ss = getSpreadsheet_();
    var sFleet = ss.getSheetByName(DB.FLEET);
    var sPhotos = ss.getSheetByName(DB.PHOTOS);
    if (!sFleet || !sPhotos) return;

    var fleetData = sFleet.getDataRange().getValues();
    var photoData = sPhotos.getDataRange().getValues();

    var containerIds = {};
    var fleetRowIndices = {};
    var currentKoli = {};

    for (var i = 1; i < fleetData.length; i++) {
      if (fleetData[i][1] === sessionId && fleetData[i][8] !== "DELETED") {
        var cid = fleetData[i][0];
        containerIds[cid] = true;
        fleetRowIndices[cid] = i + 1;
        currentKoli[cid] = Number(fleetData[i][10]) || 0;
      }
    }

    var calculatedKoli = {};
    for (var cid in containerIds) {
      calculatedKoli[cid] = 0;
    }

    var seenPhotoIds = {};
    for (var j = 1; j < photoData.length; j++) {
      var photoId = String(photoData[j][0] || "").trim();
      if (!photoId) continue;
      if (seenPhotoIds[photoId]) continue;
      seenPhotoIds[photoId] = true;
      var photoCid = photoData[j][1];
      var type = normalizePhotoType_(photoData[j][2]);
      if (containerIds[photoCid] && type === "palet") {
        var paletName = String(photoData[j][6] || "").trim();
        var paletKoli = Number(photoData[j][7]) || 0;
        if (paletName && paletKoli > 0) calculatedKoli[photoCid] += paletKoli;
      }
    }

    for (var cid in calculatedKoli) {
      var newVal = calculatedKoli[cid];
      var oldVal = currentKoli[cid];
      if (newVal !== oldVal) {
        sFleet.getRange(fleetRowIndices[cid], 11).setValue(newVal);
      }
    }
  } catch (e) {
    console.warn("syncFleetKoliFromPhotos_ error:", e.toString());
  } finally {
    if (hasLock) lock.releaseLock();
  }
}
