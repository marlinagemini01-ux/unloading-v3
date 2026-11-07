/**
 * =========================================================================
 * BatchService.gs
 * Modul untuk melakukan operasi skala besar (Bulk/Batch) menggunakan 
 * Google Sheets Advanced Service (API v4) agar terhindar dari Timeout.
 * =========================================================================
 */

/**
 * Menghapus banyak baris secara instan dalam 1 API call.
 * Cocok untuk mengarsip data lama atau menghapus baris kosong (Ghost Rows).
 * Menggunakan pendekatan Descending (bawah ke atas) via DeleteDimensionRequest.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Objek Sheet yang barisnya akan dihapus.
 * @param {number[]} rowIndices Array berisi nomor baris (1-indexed) yang ingin dihapus.
 * @returns {object} Status operasi
 */
function batchDeleteRows_(sheet, rowIndices) {
  if (!sheet || !rowIndices || rowIndices.length === 0) {
    return { success: true, message: "Tidak ada baris yang perlu dihapus." };
  }

  // 1. Dapatkan Spreadsheet ID dan Sheet ID
  var ss = sheet.getParent();
  var spreadsheetId = ss.getId();
  var sheetId = sheet.getSheetId();

  // 2. Filter validasi & urutkan secara descending (dari paling bawah ke atas)
  // Menghapus dari bawah mencegah baris bergeser (index shift) saat request dieksekusi.
  var sortedRows = rowIndices
    .filter(function(r) { return typeof r === 'number' && r > 0; })
    .sort(function(a, b) { return b - a; }); // Descending

  if (sortedRows.length === 0) {
    return { success: true, message: "Tidak ada nomor baris yang valid." };
  }

  // 3. Bangun payload request untuk batchUpdate
  var requests = [];
  
  // Karena kita menghapus per baris satu per satu tapi dalam 1 payload request:
  for (var i = 0; i < sortedRows.length; i++) {
    var rowIndex0Indexed = sortedRows[i] - 1; // API v4 menggunakan 0-index
    
    requests.push({
      deleteDimension: {
        range: {
          sheetId: sheetId,
          dimension: "ROWS",
          startIndex: rowIndex0Indexed,
          endIndex: rowIndex0Indexed + 1 // Exclusive end
        }
      }
    });
  }

  // 4. Eksekusi request menggunakan Sheets API (Advanced Services)
  try {
    Sheets.Spreadsheets.batchUpdate({ requests: requests }, spreadsheetId);
    return { 
      success: true, 
      deletedCount: sortedRows.length,
      message: sortedRows.length + " baris berhasil dihapus secara batch." 
    };
  } catch (e) {
    console.error("batchDeleteRows_ Gagal:", e.toString());
    // Fallback? Jika Sheets API belum diaktifkan, ini akan gagal.
    return { 
      success: false, 
      error: e.toString(),
      message: "Gagal menghapus baris. Pastikan Sheets Advanced Service aktif." 
    };
  }
}

/**
 * Garbage Collector untuk membersihkan Ghost Rows (baris hasil clearContent).
 * Dipanggil via Time-driven trigger (misal jam 2 pagi) atau manual.
 * Mengurangi ukuran sheet yang membengkak karena baris kosong.
 */
function runGhostRowGarbageCollector() {
  var started = Date.now();
  console.log("Memulai Ghost Row Garbage Collector...");
  
  var DB = getSheetNames_();
  var ss = getSpreadsheet_();
  
  // Daftar sheet yang sering kena clearContent (saat hapus armada/foto)
  var targetSheets = [DB.PHOTOS, DB.FLEET];
  var resultLog = [];
  
  var lock = LockService.getScriptLock();
  try {
    // Kunci selama 20 detik agar tidak ada yang insert data saat kita hapus
    lock.waitLock(20000);
    
    for (var i = 0; i < targetSheets.length; i++) {
      var sheetName = targetSheets[i];
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;
      
      var data = sheet.getDataRange().getValues();
      var emptyRowIndices = [];
      
      // Deteksi baris yang ID-nya kosong (Kolom A kosong)
      // Melewati baris 1 (Header)
      for (var r = 1; r < data.length; r++) {
        var firstCell = String(data[r][0] || "").trim();
        // Kalau kolom A kosong, dan kolom B juga kosong, asumsikan ini Ghost Row
        if (!firstCell) {
          emptyRowIndices.push(r + 1); // r + 1 karena data 0-indexed dan sheet 1-indexed
        }
      }
      
      if (emptyRowIndices.length > 0) {
        var res = batchDeleteRows_(sheet, emptyRowIndices);
        resultLog.push(sheetName + ": Dihapus " + (res.deletedCount || 0) + " baris kosong.");
      } else {
        resultLog.push(sheetName + ": Bersih.");
      }
    }
  } catch (e) {
    console.error("GC Error: " + e.toString());
    resultLog.push("ERROR: " + e.toString());
  } finally {
    lock.releaseLock();
  }
  
  var elapsed = Date.now() - started;
  console.log("Garbage Collector selesai dalam " + elapsed + "ms. Hasil:\n" + resultLog.join("\n"));
  
  return {
    success: true,
    elapsedMs: elapsed,
    results: resultLog
  };
}
