// =========================================================================
// ArchiveService.gs — D1: Arsip otomatis sesi lama
// =========================================================================
// Time-driven trigger: setiap hari jam 01:00 WIB
// Memindahkan data COMPLETED > 30 hari ke sheet arsip

function archiveOldSessions() {
  var DB = getSheetNames_();
  var ss = getSpreadsheet_();
  var cutoffDays = cfgNum("Archive Cutoff Days") || 30;
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - cutoffDays);

  var sheetsToArchive = [
    { name: DB.SESSIONS, idCol: 0, statusCol: 5, dateCol: 6 },
    { name: DB.FLEET,    idCol: 1, statusCol: null, dateCol: 6 },
    { name: DB.PHOTOS,   idCol: 1, statusCol: null, dateCol: 4 },
    { name: DB.USERS,    idCol: 1, statusCol: null, dateCol: 5 }
  ];

  // Step 1: Find completed session IDs older than cutoff
  var sessionSheet = ss.getSheetByName(DB.SESSIONS);
  if (!sessionSheet) return "No Sessions sheet found.";
  var sessionData = sessionSheet.getDataRange().getValues();
  var archiveSessionIds = {};
  for (var i = 1; i < sessionData.length; i++) {
    if (sessionData[i][5] === "COMPLETED") {
      var ts = sessionData[i][6];
      if (ts instanceof Date && ts < cutoff) {
        archiveSessionIds[sessionData[i][0]] = true;
      }
    }
  }

  var archivedCount = Object.keys(archiveSessionIds).length;
  if (archivedCount === 0) return "Tidak ada sesi untuk diarsip.";

  // Step 2: For each sheet, move matching rows to archive sheet
  sheetsToArchive.forEach(function(config) {
    var sheet = ss.getSheetByName(config.name);
    if (!sheet) return;

    var archiveName = config.name + "_Archive";
    var archiveSheet = ss.getSheetByName(archiveName);
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet(archiveName);
      // Copy headers
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
      archiveSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
      archiveSheet.getRange(1, 1, 1, headers[0].length).setFontWeight("bold").setBackground("#f1f5f9");
      archiveSheet.setFrozenRows(1);
    }

    var data = sheet.getDataRange().getValues();
    var rowsToArchive = [];
    var rowsToDelete = []; // 1-indexed row numbers

    for (var i = 1; i < data.length; i++) {
      var sessionId = config.idCol === 0 ? data[i][0] : data[i][config.idCol];
      if (archiveSessionIds[sessionId]) {
        rowsToArchive.push(data[i]);
        rowsToDelete.push(i + 1);
      }
    }

    // Append to archive
    if (rowsToArchive.length > 0) {
      archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive.length, rowsToArchive[0].length).setValues(rowsToArchive);
    }

    // Delete from source using batch API to prevent timeout
    if (rowsToDelete.length > 0) {
      batchDeleteRows_(sheet, rowsToDelete);
    }
  });

  SpreadsheetApp.flush();

  // Log to audit
  try {
    var auditSheet = ss.getSheetByName(DB.AUDIT);
    if (auditSheet) {
      auditSheet.appendRow([new Date(), "ARCHIVE", "SUCCESS", archivedCount + " sesi diarsip (>" + cutoffDays + " hari)", "System/Trigger"]);
    }
  } catch(e) {}

  return archivedCount + " sesi berhasil diarsip.";
}

// =========================================================================
// setupArchiveTrigger() — Panggil sekali untuk setup trigger harian
// =========================================================================
function setupArchiveTrigger() {
  // Hapus trigger lama jika ada
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === "archiveOldSessions") ScriptApp.deleteTrigger(t);
  });

  // Buat trigger baru: setiap hari jam 01:00
  ScriptApp.newTrigger("archiveOldSessions")
    .timeBased()
    .atHour(1)
    .everyDays(1)
    .inTimezone(cfg("Timezone") || "Asia/Jakarta")
    .create();

  return "Trigger arsip harian berhasil dibuat (jam 01:00).";
}