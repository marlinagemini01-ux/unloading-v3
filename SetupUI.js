/**
 * =========================================================================
 * SetupUI.gs
 * Menambahkan menu kustom di Google Sheets (khusus untuk admin)
 * agar mudah mengatur Spreadsheet ID saat duplikasi/pindah database.
 * =========================================================================
 */

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ InLens Setup')
    .addItem('1. 🔍 Cek Status Koneksi Saat Ini', 'showSystemStatus_')
    .addSeparator()
    .addItem('2. 🔗 Hubungkan Spreadsheet ID', 'promptSpreadsheetSetup_')
    .addItem('3. 🛠️ Setup Database Awal (Buat Sheet)', 'setupDatabase')
    .addSeparator()
    .addItem('4. 🧹 Reset Data (Kosongkan Data Uji Coba)', 'DANGER_RESET_ALL_SHEETS')
    .addItem('5. ♻️ Hapus Cache Server', 'clearAllCache')
    .addToUi();
}

function showSystemStatus_() {
  var ui = SpreadsheetApp.getUi();
  
  // 1. Spreadsheet Binding Info
  var propId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID") || "Tidak diset (Kosong)";
  var activeId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var scriptId = "Tidak diketahui";
  var webAppUrl = "Belum di-deploy / Tidak diketahui";
  
  try { scriptId = ScriptApp.getScriptId(); } catch(e) {}
  try { webAppUrl = ScriptApp.getService().getUrl(); } catch(e) {}
  
  // 2. Baca AppConfig (Abaikan error jika belum ada)
  var configSsId = "Gagal membaca AppConfig";
  var rootFolderId = "Kosong";
  var rootFolderName = "Tidak diketahui";
  var saEmail = "Kosong";
  
  try {
    var config = getConfig_();
    configSsId = config["Spreadsheet ID"] || "Kosong";
    rootFolderId = config["Root Folder ID"] || "Kosong";
    saEmail = config["SA Email"] || "Kosong";
    
    if (rootFolderId !== "Kosong") {
      try {
        rootFolderName = DriveApp.getFolderById(rootFolderId).getName();
      } catch (fErr) {
        rootFolderName = "⚠️ Gagal diakses (Pastikan ID Benar & Anda punya akses)";
      }
    }
  } catch (err) {
    // Abaikan jika belum disetup
  }
  
  var statusMsg = "Informasi ini membantu melacak apakah aplikasi Anda 'nyangkut' ke file lama atau konfigurasi salah.\n\n";
  statusMsg += "=== 📊 SPREADSHEET (DATABASE) ===\n";
  statusMsg += "🔸 SPREADSHEET INI (Sedang Buka):\n" + activeId + "\n\n";
  statusMsg += "🔸 MEMORI INTERNAL (Terhubung ke):\n" + propId + "\n\n";
  statusMsg += "🔸 DI DALAM APPCONFIG (Tertulis):\n" + configSsId + "\n\n";
  
  statusMsg += "=== 🌐 APLIKASI WEB ===\n";
  statusMsg += "🔸 SCRIPT ID:\n" + scriptId + "\n\n";
  statusMsg += "🔸 WEB APP URL:\n" + webAppUrl + "\n\n";
  
  statusMsg += "=== 📂 GOOGLE DRIVE (FOTO) ===\n";
  statusMsg += "🔸 ROOT FOLDER ID:\n" + rootFolderId + "\n\n";
  statusMsg += "🔸 NAMA FOLDER:\n" + rootFolderName + "\n\n";
  
  statusMsg += "=== 🤖 SERVICE ACCOUNT (SA) ===\n";
  statusMsg += "🔸 SA EMAIL:\n" + saEmail + "\n\n";
  
  if (propId !== "Tidak diset (Kosong)" && propId !== activeId) {
    statusMsg += "⚠️ PERINGATAN KRITIS:\nMemori internal berbeda dengan Spreadsheet ini!\nAplikasi Web Anda kemungkinan besar membaca dan menulis data ke Spreadsheet LAMA.\nSilakan gunakan menu 'Hubungkan Spreadsheet ID' untuk memperbaikinya.";
  } else {
    statusMsg += "✅ STATUS SPREADSHEET AMAN: Aplikasi terhubung sinkron pada Spreadsheet ini.";
  }
  
  ui.alert('Status Sistem InLens (Super Admin)', statusMsg, ui.ButtonSet.OK);
}

function promptSpreadsheetSetup_() {
  var ui = SpreadsheetApp.getUi();
  
  // Ambil ID yang saat ini menempel di memori Script Properties
  var propId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  var activeId = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  if (propId) {
    var response = ui.alert(
      'Konfirmasi Database Saat Ini', 
      'Aplikasi saat ini TERHUBUNG dengan Spreadsheet ID berikut (di memori internal):\n\n' + 
      propId + 
      '\n\nApakah Anda ingin TETAP menggunakan ID ini?\n(Pilih "No" jika Anda ingin mereset/mengganti dengan ID baru)', 
      ui.ButtonSet.YES_NO
    );
    
    if (response == ui.Button.YES) {
      ui.alert('Info', 'Tidak ada perubahan dilakukan. Aplikasi tetap menggunakan database lama.', ui.ButtonSet.OK);
      return;
    }
  }
  
  // Jika memilih No, atau belum pernah diset sama sekali:
  var promptRes = ui.prompt(
    'Masukkan Spreadsheet ID Baru', 
    'Kosongkan kolom ini lalu klik OK jika Anda ingin OTOMATIS menggunakan Spreadsheet yang sedang Anda buka ini.\n\nAtau paste URL Spreadsheet lain jika Anda ingin menghubungkannya ke tempat lain:', 
    ui.ButtonSet.OK_CANCEL
  );
  
  if (promptRes.getSelectedButton() == ui.Button.OK) {
    var input = promptRes.getResponseText().trim();
    var newId = input ? normalizeSpreadsheetId_(input) : activeId;
    
    // Hapus memori lama dan set yang baru
    PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", newId);
    
    // Hapus cache agar web app langsung refresh
    try {
      clearAllCache();
    } catch(e) {}
    
    ui.alert(
      '✅ Sukses Terhubung!', 
      'Spreadsheet ID berhasil diperbarui menjadi:\n\n' + newId + '\n\nSilakan refresh / buka ulang Web App Anda.', 
      ui.ButtonSet.OK
    );
  }
}
