/**
 * =========================================================================
 * Repository.gs
 * Data Access Layer (DAL) untuk memisahkan logika akses Spreadsheet dari 
 * business logic. Memudahkan maintenance jika struktur kolom berubah.
 * =========================================================================
 */

var AppRepository = (function() {
  
  // Konfigurasi Index Kolom (0-indexed untuk array mapping)
  // Ubah di sini jika struktur kolom di Spreadsheet bergeser.
  var COL = {
    SESSION: {
      ID: 0,
      DATE: 1,
      BRAND: 2,
      SHIPMENT: 3,
      PIC: 4,
      STATUS: 5,
      CREATED_AT: 6,
      DRIVE_FOLDER: 7,
      ADMIN: 8,
      LP_NAME: 10
    },
    FLEET: {
      ID: 0,
      SESSION_ID: 1,
      NUMBER: 2,
      POLICE_NUMBER: 3,
      DRIVER: 4,
      TYPE: 5,
      STARTED_AT: 6,
      COMPLETED_AT: 7,
      STATUS: 8,
      REOPENED_AT: 9,
      QTY_KOLI: 10,
      REOPENED_BY: 11
    }
  };

  /**
   * Helper untuk mencari satu baris berdasarkan ID.
   * @param {string} sheetName Nama sheet (misal: "Sessions")
   * @param {string} id Nilai ID yang dicari
   * @param {number} idColIndex Indeks kolom tempat ID berada (default 0)
   * @returns {object|null} { rowIndex: number, data: array, sheet: sheet } atau null
   */
  function findById(sheetName, id, idColIndex) {
    if (idColIndex === undefined) idColIndex = 0;
    var sheet = getSpreadsheet_().getSheetByName(sheetName);
    if (!sheet) return null;
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][idColIndex] === id) {
        return {
          rowIndex: i + 1, // 1-indexed untuk getRange()
          data: data[i],
          sheet: sheet
        };
      }
    }
    return null;
  }

  return {
    COL: COL,
    
    // --- SESSIONS ---
    Sessions: {
      /**
       * Mendapatkan data Sesi berdasarkan ID.
       */
      getById: function(sessionId) {
        var DB = getSheetNames_();
        return findById(DB.SESSIONS, sessionId, COL.SESSION.ID);
      },
      
      /**
       * Mengubah status sesi.
       */
      updateStatus: function(sessionId, newStatus) {
        var record = this.getById(sessionId);
        if (record) {
          record.sheet.getRange(record.rowIndex, COL.SESSION.STATUS + 1).setValue(newStatus);
          return true;
        }
        return false;
      }
    },
    
    // --- FLEET ---
    Fleet: {
      /**
       * Mendapatkan data Armada berdasarkan Container ID.
       */
      getById: function(containerId) {
        var DB = getSheetNames_();
        return findById(DB.FLEET, containerId, COL.FLEET.ID);
      }
    }
    
  };
})();
