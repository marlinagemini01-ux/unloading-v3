// =========================================================================
// generateRekap_() — BUG 6 FIX tetap: Validasi & update header
// =========================================================================
function generateRekap_(sessionId) {
  var DB = getSheetNames_();
  var ss = getSpreadsheet_();

  var sessionData = ss.getSheetByName(DB.SESSIONS).getDataRange().getValues();
  var sessionRow = null;
  for (var i = sessionData.length - 1; i >= 1; i--) {
    if (sessionData[i][0] === sessionId) {
      sessionRow = sessionData[i];
      break;
    }
  }
  if (!sessionRow) return;

  var tz = cfg("Timezone") || "Asia/Jakarta";
  var df = cfg("Date Format") || "dd MMMM yyyy";
  var date = sessionRow[1];
  var dateStr =
    date instanceof Date ? Utilities.formatDate(date, tz, df) : date;

  var fleetData = ss.getSheetByName(DB.FLEET).getDataRange().getValues();
  var activeFleets = [];
  for (var i = 1; i < fleetData.length; i++) {
    if (fleetData[i][1] === sessionId && fleetData[i][8] !== "DELETED")
      activeFleets.push(fleetData[i]);
  }

  var photoData = ss.getSheetByName(DB.PHOTOS).getDataRange().getValues();
  var photoCounts = {};
  activeFleets.forEach(function (f) {
    photoCounts[f[0]] = { dokumen: 0, palet: 0, suratjalan: 0 };
  });
  var seenRekapPhotoIds = {};
  for (var i = 1; i < photoData.length; i++) {
    // PATCH: skip baris kosong hasil clearContent
    var rekapPhotoId = String(photoData[i][0] || "").trim();
    if (!rekapPhotoId) continue;
    if (seenRekapPhotoIds[rekapPhotoId]) continue;
    seenRekapPhotoIds[rekapPhotoId] = true;
    var cid = photoData[i][1],
      pType = photoData[i][2];
    if (photoCounts[cid]) {
      if (pType === "dokumen") photoCounts[cid].dokumen++;
      else if (pType === "palet") photoCounts[cid].palet++;
      else if (pType === "suratjalan") photoCounts[cid].suratjalan++;
    }
  }

  var rekapSheet = ss.getSheetByName(DB.REKAP);
  if (!rekapSheet) return;

  var colOrder = cfgList("Rekap Column Order");
  if (!colOrder.length) {
    colOrder = [
      "Tanggal",
      "Brand",
      "Shipment",
      "Sesi ID",
      "Armada Ke",
      "Tipe",
      "Nama Supir",
      "Plat Nomor",
      "Jml Dokumen",
      "Jml Palet",
      "Jml Surat Jalan",
      "Qty Koli",
      "Link Folder Armada",
    ];
  }

  // BUG 6 FIX: Bandingkan header sheet dengan colOrder
  var currentHeader = rekapSheet
    .getRange(1, 1, 1, colOrder.length)
    .getValues()[0];
  var headerMatch = colOrder.every(function (col, idx) {
    return currentHeader[idx] === col;
  });
  if (!headerMatch) {
    try {
      rekapSheet.getRange(1, 1, 1, colOrder.length).setValues([colOrder]);
      rekapSheet
        .getRange(1, 1, 1, colOrder.length)
        .setFontWeight("bold")
        .setBackground("#f1f5f9");
      SpreadsheetApp.flush();
    } catch (headerErr) {
      console.error("Rekap header update GAGAL:", headerErr.toString());
      logAudit_("REKAP_HEADER_FAIL", "ERROR", headerErr.toString(), sessionId);
      return;
    }
  }

  var rowsToWrite = activeFleets.map(function (f) {
    var lookup = {
      Tanggal: dateStr,
      Brand: sessionRow[2],
      Shipment: sessionRow[3],
      "Sesi ID": sessionId,
      "Armada Ke": f[2],
      Tipe: f[3],
      "Nama Supir": f[4],
      "Plat Nomor": f[5],
      "Jml Dokumen": photoCounts[f[0]].dokumen,
      "Jml Palet": photoCounts[f[0]].palet,
      "Jml Surat Jalan": photoCounts[f[0]].suratjalan,
      "Qty Koli": Number(f[10]) > 0 ? Math.floor(Number(f[10])) : 0,
      "Link Folder Armada": f[7],
      PIC: sessionRow[4],
      "Created By": sessionRow[8],
    };
    return colOrder.map(function (col) {
      return lookup[col] !== undefined ? lookup[col] : "";
    });
  });

  if (rowsToWrite.length > 0) {
    rekapSheet
      .getRange(
        rekapSheet.getLastRow() + 1,
        1,
        rowsToWrite.length,
        rowsToWrite[0].length,
      )
      .setValues(rowsToWrite);
    SpreadsheetApp.flush();
  }
}
