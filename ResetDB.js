/**
 * =========================================================================
 * ResetDB.gs
 * Skrip bantuan untuk MENGHAPUS SEMUA DATA saat testing, 
 * TAPI mempertahankan baris ke-1 (Header).
 * 
 * CARA PAKAI:
 * 1. Buka file ini di Editor Google Apps Script
 * 2. Pilih fungsi DANGER_RESET_ALL_SHEETS di menu atas
 * 3. Klik tombol 'Run' (Jalankan)
 * =========================================================================
 */

function DANGER_RESET_ALL_SHEETS() {
  var ss = getSpreadsheet_();
  var DB = getSheetNames_();
  
  console.log("Mencoba mereset Spreadsheet: " + ss.getName());
  console.log("URL Spreadsheet: " + ss.getUrl());
  
  // Daftar sheet yang akan dibersihkan
  var sheetsToReset = [
    DB.SESSIONS,
    DB.FLEET,
    DB.PHOTOS,
    DB.ARCHIVE,
    DB.USERS,
    DB.SJ_QTY
  ];
  
  var totalRowsCleared = 0;
  
  for (var i = 0; i < sheetsToReset.length; i++) {
    var sheetName = sheetsToReset[i];
    if (!sheetName) continue;
    
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      var maxRows = sheet.getMaxRows();
      var maxCols = sheet.getMaxColumns();
      
      // Jika sheet punya baris lebih dari 1
      if (maxRows > 1) {
        // Hapus SEMUA nilai dan format dari baris ke-2 sampai mentok bawah
        // Gunakan clearContent() agar header di baris 1 tetap utuh.
        sheet.getRange(2, 1, maxRows - 1, maxCols).clearContent();
        totalRowsCleared += (maxRows - 1);
      }
    }
  }
  
  // Kosongkan Cache (agar web app tidak memunculkan data lama yang nyangkut di memori)
  try {
    CacheService.getScriptCache().removeAll([
      "SESSION_INDEX",
      "FLEET_INDEX",
      "PHOTO_INDEX",
      "USER_INDEX"
    ]);
  } catch(e) {}
  
  console.log("✅ RESET BERHASIL! Total " + totalRowsCleared + " baris data telah dikosongkan. Header tetap aman.");
}
