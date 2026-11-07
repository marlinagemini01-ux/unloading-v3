// =========================================================================
// joinUserSession() — OPTIMASI: batch setValues (2 kolom dalam 1 call)
// =========================================================================
function buildUserNameCandidates_(data) {
  var map = {};
  [data && data.userName, data && data.fullName, data && data.username, data && data.requestedBy].forEach(function (v) {
    var clean = String(v || "").trim();
    if (clean) map[clean.toLowerCase()] = true;
  });

  var DB = getSheetNames_();
  var accountSheet = getSpreadsheet_().getSheetByName(DB.ACCOUNTS);
  if (accountSheet) {
    try {
      var rows = accountSheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        var username = String(rows[i][0] || "").trim();
        var fullName = String(rows[i][2] || "").trim();
        if (username && map[username.toLowerCase()] && fullName) map[fullName.toLowerCase()] = true;
        if (fullName && map[fullName.toLowerCase()] && username) map[username.toLowerCase()] = true;
      }
    } catch (e) {}
  }
  return map;
}

function getDisplayUserName_(data) {
  return String((data && data.fullName) || (data && data.userName) || (data && data.username) || (data && data.requestedBy) || "Anon").trim() || "Anon";
}

function joinUserSession(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    var DB = getSheetNames_();
    var sheet = getSpreadsheet_().getSheetByName(DB.USERS);
    var rows = sheet.getDataRange().getValues();
    var candidates = buildUserNameCandidates_(data);

    for (var i = 1; i < rows.length; i++) {
      if (
        candidates[String(rows[i][2]).trim().toLowerCase()] &&
        rows[i][1] === data.sessionId
      ) {
        // OPTIMASI: 1x setValues batch (bukan 2x setValue)
        sheet.getRange(i + 1, 4, 1, 2).setValues([["WORKING", new Date()]]);
        rebuildAndCacheSync_(data.sessionId);
        return { success: true, alreadyJoined: true };
      }
    }

    sheet.appendRow([
      generateUniqueId_("User ID Prefix"),
      data.sessionId,
      getDisplayUserName_(data),
      "WORKING",
      new Date(),
      new Date(),
    ]);
    rebuildAndCacheSync_(data.sessionId);
    return { success: true, alreadyJoined: false };
  } catch (e) {
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// markUserDone() — OPTIMASI: Single-loop (update + count dalam 1 pass)
// Sebelumnya: baca sheet 2x. Sekarang: 1x baca + update in-memory.
// =========================================================================
function markUserDone(data) {
  var traceId = data && data.traceId;
  perfTrace_("markDone", traceId, "server_start", { sessionId: data && data.sessionId });
  var lock = LockService.getScriptLock();
  try {
    perfTrace_("markDone", traceId, "lock_wait_start");
    lock.waitLock(5000);
    perfTrace_("markDone", traceId, "lock_wait_end");
    var DB = getSheetNames_();
    var ss = getSpreadsheet_();
    var pendingFleetWork = getUnfinishedFleetWorkForSession_(data.sessionId, ss);
    if (pendingFleetWork.length > 0) {
      perfTrace_("markDone", traceId, "blocked_fleet_work_not_done", { pending: pendingFleetWork.length });
      return {
        success: false,
        code: "FLEET_WORK_NOT_DONE",
        message: "Masih ada armada yang belum di-finish.",
        pendingFleets: pendingFleetWork
      };
    }
    var sheet = ss.getSheetByName(DB.USERS);
    perfTrace_("markDone", traceId, "sheet_read_start");
    var rows = sheet.getDataRange().getValues();
    perfTrace_("markDone", traceId, "sheet_read_end", { rows: rows.length });

    var updated = false;
    var total = 0,
      done = 0,
      working = [];
    var candidates = buildUserNameCandidates_(data);

    for (var i = 1; i < rows.length; i++) {
      if (rows[i][1] !== data.sessionId) continue;

      var isMe = candidates[String(rows[i][2]).trim().toLowerCase()];

      if (isMe && !updated) {
        // OPTIMASI: 1x setValues batch (bukan 2x setValue)
        perfTrace_("markDone", traceId, "sheet_write_start", { row: i + 1 });
        sheet.getRange(i + 1, 4, 1, 2).setValues([["DONE", new Date()]]);
        perfTrace_("markDone", traceId, "sheet_write_end");
        rows[i][3] = "DONE"; // update in-memory juga
        updated = true;
      }

      total++;
      if (rows[i][3] === "DONE") done++;
      else working.push(rows[i][2]);
    }

    if (!updated) {
      var displayName = getDisplayUserName_(data);
      perfTrace_("markDone", traceId, "sheet_write_start", { append: true });
      sheet.appendRow([
        generateUniqueId_("User ID Prefix"),
        data.sessionId,
        displayName,
        "DONE",
        new Date(),
        new Date(),
      ]);
      perfTrace_("markDone", traceId, "sheet_write_end");
      total++;
      done++;
    }

    rebuildAndCacheSync_(data.sessionId);
    perfTrace_("markDone", traceId, "server_end", { success: true, doneUsers: done, totalUsers: total });
    return {
      success: true,
      allDone: total > 0 && done === total,
      totalUsers: total,
      doneUsers: done,
      workingUsers: working,
    };
  } catch (e) {
    perfTrace_("markDone", traceId, "server_error", { message: e.toString() });
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// resumeUserWork() — OPTIMASI: batch setValues
// =========================================================================
function resumeUserWork(data) {
  var traceId = data && data.traceId;
  perfTrace_("resumeWork", traceId, "server_start", { sessionId: data && data.sessionId });
  var lock = LockService.getScriptLock();
  try {
    perfTrace_("resumeWork", traceId, "lock_wait_start");
    lock.waitLock(5000);
    perfTrace_("resumeWork", traceId, "lock_wait_end");
    var DB = getSheetNames_();
    var sheet = getSpreadsheet_().getSheetByName(DB.USERS);
    perfTrace_("resumeWork", traceId, "sheet_read_start");
    var rows = sheet.getDataRange().getValues();
    perfTrace_("resumeWork", traceId, "sheet_read_end", { rows: rows.length });
    var candidates = buildUserNameCandidates_(data);

    for (var i = 1; i < rows.length; i++) {
      if (
        candidates[String(rows[i][2]).trim().toLowerCase()] &&
        rows[i][1] === data.sessionId
      ) {
        // OPTIMASI: 1x setValues batch (bukan 2x setValue)
        perfTrace_("resumeWork", traceId, "sheet_write_start", { row: i + 1 });
        sheet.getRange(i + 1, 4, 1, 2).setValues([["WORKING", new Date()]]);
        perfTrace_("resumeWork", traceId, "sheet_write_end");
        rebuildAndCacheSync_(data.sessionId);
        perfTrace_("resumeWork", traceId, "server_end", { success: true });
        return { success: true };
      }
    }
    perfTrace_("resumeWork", traceId, "sheet_write_start", { append: true });
    sheet.appendRow([
      generateUniqueId_("User ID Prefix"),
      data.sessionId,
      getDisplayUserName_(data),
      "WORKING",
      new Date(),
      new Date(),
    ]);
    perfTrace_("resumeWork", traceId, "sheet_write_end");
    rebuildAndCacheSync_(data.sessionId);
    perfTrace_("resumeWork", traceId, "server_end", { success: true, recreated: true });
    return { success: true, recreated: true };
  } catch (e) {
    perfTrace_("resumeWork", traceId, "server_error", { message: e.toString() });
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}
