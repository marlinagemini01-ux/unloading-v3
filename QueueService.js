/**
 * Menambahkan task ke queue untuk diproses di background.
 * Digunakan agar finishSession (rekap + PDF + WA) tidak nahan browser terlalu lama.
 */
function enqueueFinishSessionTask(sessionId, requestedBy, traceId) {
  var props = PropertiesService.getScriptProperties();
  var queueStr = props.getProperty("FINISH_SESSION_QUEUE");
  
  var queue = [];
  if (queueStr) {
    try {
      queue = JSON.parse(queueStr);
    } catch (e) {
      queue = [];
    }
  }

  // Cek apakah sudah ada di antrean
  var exists = queue.some(function(item) {
    return item.sessionId === sessionId;
  });

  if (!exists) {
    queue.push({
      sessionId: sessionId,
      requestedBy: requestedBy,
      traceId: traceId || "BG_" + Date.now(),
      timestamp: Date.now()
    });
    props.setProperty("FINISH_SESSION_QUEUE", JSON.stringify(queue));
    
    // Trigger akan jalan 1 menit lagi
    ScriptApp.newTrigger("processBackgroundTasks_")
      .timeBased()
      .after(1) 
      .create();
  }
}

/**
 * Fungsi yang dipanggil oleh Trigger Background.
 * Menyapu antrean dan memproses pembuatan Rekap, PDF, dan WA.
 */
function processBackgroundTasks_() {
  var props = PropertiesService.getScriptProperties();
  var queueStr = props.getProperty("FINISH_SESSION_QUEUE");
  if (!queueStr) return;
  
  var queue = [];
  try {
    queue = JSON.parse(queueStr);
  } catch(e) {
    return;
  }
  
  if (queue.length === 0) return;
  
  // Ambil 1 task paling depan
  var task = queue.shift();
  props.setProperty("FINISH_SESSION_QUEUE", JSON.stringify(queue));
  
  var sessionId = task.sessionId;
  var requestedBy = task.requestedBy;
  var traceId = task.traceId;
  
  console.log("Background Task Processing:", sessionId);
  
  var ss = getSpreadsheet_();
  var DB = getSheetNames_();
  var sSession = ss.getSheetByName(DB.SESSIONS);
  
  var sessionRows = sSession.getDataRange().getValues();
  var sessionData = null;
  var rowIndex = -1;
  for (var i = 1; i < sessionRows.length; i++) {
    if (sessionRows[i][0] === sessionId) {
      sessionData = sessionRows[i];
      rowIndex = i + 1;
      break;
    }
  }
  
  if (!sessionData) return;
  
  // 1. REKAP
  var rekapSuccess = true, rekapError = "";
  if (cfgBool("Rekap Enabled")) {
    try {
      generateRekap_(sessionId);
    } catch (e) {
      rekapSuccess = false;
      rekapError = e.toString();
    }
    try {
      writeSessionArtifactStatus_(sSession, rowIndex, {
        "Rekap Status": rekapSuccess ? "SUCCESS" : "FAILED",
        "Rekap Error": rekapError || "",
        "Artifact Updated At": new Date()
      });
    } catch(e) { console.error("Gagal tulis rekap status:", e); }
  } else {
    try {
      writeSessionArtifactStatus_(sSession, rowIndex, {
        "Rekap Status": "SKIPPED",
        "Artifact Updated At": new Date()
      });
    } catch(e) {}
  }
  
  // 2. HANDOVER PDF
  var handoverPdfSuccess = true, handoverPdfError = "", handoverPdf = null;
  try {
    var pdfRes = createOrUpdateHandoverPdf(sessionId, requestedBy, "Finish session (Background)", traceId);
    if (pdfRes && pdfRes.success) {
      handoverPdf = pdfRes.data || null;
    } else {
      handoverPdfSuccess = false;
      handoverPdfError = pdfRes && pdfRes.message ? pdfRes.message : "Gagal";
    }
  } catch (pdfErr) {
    handoverPdfSuccess = false;
    handoverPdfError = pdfErr.toString();
  }
  
  try {
    writeSessionArtifactStatus_(sSession, rowIndex, {
      "Handover PDF Status": handoverPdfSuccess ? "SUCCESS" : "FAILED",
      "Handover PDF Error": handoverPdfError || "",
      "Artifact Updated At": new Date()
    });
  } catch(e) { console.error("Gagal tulis handover status:", e); }
  
  // 3. WHATSAPP NOTIFICATION
  var waSuccess = true, waError = "";
  if (cfgBool("WA Report Enabled")) {
    try {
      var waRes = sendWaNotification_(sessionId, sessionData);
      if (!waRes.success) {
        waSuccess = false;
        waError = waRes.message || "Gagal kirim WA";
      }
    } catch (e) {
      waSuccess = false;
      waError = e.toString();
    }
    
    try {
      writeSessionArtifactStatus_(sSession, rowIndex, {
        "WA Status": waSuccess ? "SUCCESS" : "FAILED",
        "WA Error": waError || "",
        "Artifact Updated At": new Date()
      });
    } catch(e) { console.error("Gagal tulis WA status:", e); }
  } else {
    try {
      writeSessionArtifactStatus_(sSession, rowIndex, {
        "WA Status": "SKIPPED",
        "Artifact Updated At": new Date()
      });
    } catch(e) {}
  }

  // Bersihkan trigger sendiri jika antrean sudah kosong
  if (queue.length === 0) {
    var triggers = ScriptApp.getProjectTriggers();
    for (var j = 0; j < triggers.length; j++) {
      if (triggers[j].getHandlerFunction() === "processBackgroundTasks_") {
        ScriptApp.deleteTrigger(triggers[j]);
      }
    }
  } else {
    // Jika masih ada, bikin trigger lagi untuk 1 menit ke depan
    ScriptApp.newTrigger("processBackgroundTasks_").timeBased().after(1).create();
  }
}

/**
 * Mengirim notifikasi WhatsApp via Fonnte
 */
function sendWaNotification_(sessionId, sessionData) {
  var fonnteToken = getSecretProps_().FONNTE_TOKEN;
  if (!fonnteToken) return { success: false, message: "Fonnte Token belum di-set" };
  
  var targetWa = cfg("Target Group WA");
  if (!targetWa) return { success: false, message: "Target Group WA belum di-set" };
  
  var ss = getSpreadsheet_();
  var DB = getSheetNames_();
  var sFleet = ss.getSheetByName(DB.FLEET);
  var fleetRows = sFleet.getDataRange().getValues();
  
  var fleetCount = 0;
  var totalKoli = 0;
  var fleetListText = "";
  
  for (var i = 1; i < fleetRows.length; i++) {
    if (fleetRows[i][1] === sessionId) {
      fleetCount++;
      var koli = Number(fleetRows[i][10]) || 0;
      totalKoli += koli;
      fleetListText += "- " + fleetRows[i][3] + " / " + fleetRows[i][5] + " / " + fleetRows[i][4] + " (" + koli + " Koli)\n";
    }
  }
  
  var template = cfg("WA Report Template");
  if (!template) return { success: false, message: "Template WA belum di-set" };
  
  // Hitung durasi (Timestamp ke Handover Generated At, atau Date.now)
  var startTime = new Date(sessionData[6]).getTime();
  var endTime = sessionData[21] ? new Date(sessionData[21]).getTime() : Date.now();
  var durationStr = formatDurationRange_(startTime, endTime, "0 Jam 0 Menit");
  
  // Hitung photo count
  var sPhotos = ss.getSheetByName(DB.PHOTOS);
  var photoRows = sPhotos.getDataRange().getValues();
  var photoCount = 0;
  for (var p = 1; p < photoRows.length; p++) {
    if (photoRows[p][1] && photoRows[p][1].toString().indexOf(sessionId) === 0) {
      photoCount++;
    }
  }

  var tz = cfg("Timezone") || "Asia/Jakarta";
  var df = cfg("Date Format") || "dd MMMM yyyy";
  var formattedTanggal = sessionData[1] instanceof Date ? Utilities.formatDate(sessionData[1], tz, df) : (sessionData[1] || "-");

  var message = replacePlaceholders_(template, {
    Brand: sessionData[2] || "-",
    Shipment: sessionData[3] || "-",
    Tanggal: formattedTanggal,
    PIC: sessionData[4] || "-",
    LossPreventionName: sessionData[10] || "-",
    Duration: durationStr,
    PhotoCount: photoCount,
    TotalKoli: totalKoli,
    FleetCount: fleetCount,
    FleetList: fleetListText.trim(),
    FolderURL: sessionData[7] || "-"
  });
  
  var options = {
    method: 'post',
    headers: { 'Authorization': fonnteToken },
    payload: { target: targetWa, message: message },
    muteHttpExceptions: true
  };
  
  try {
    var response = UrlFetchApp.fetch("https://api.fonnte.com/send", options);
    var result = JSON.parse(response.getContentText());
    if (result.status === true) {
      return { success: true, data: result };
    } else {
      return { success: false, message: result.reason || "Fonnte error" };
    }
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Dipanggil dari Frontend untuk retry kirim WA
 */
function retryWaNotification(sessionId) {
  try {
    var ss = getSpreadsheet_();
    var DB = getSheetNames_();
    var sSession = ss.getSheetByName(DB.SESSIONS);
    
    var sessionRows = sSession.getDataRange().getValues();
    var sessionData = null;
    var rowIndex = -1;
    for (var i = 1; i < sessionRows.length; i++) {
      if (sessionRows[i][0] === sessionId) {
        sessionData = sessionRows[i];
        rowIndex = i + 1;
        break;
      }
    }
    
    if (!sessionData) return { success: false, message: "Sesi tidak ditemukan" };
    
    var waRes = sendWaNotification_(sessionId, sessionData);
    
    writeSessionArtifactStatus_(sSession, rowIndex, {
      "WA Status": waRes.success ? "SUCCESS" : "FAILED",
      "WA Error": waRes.success ? "" : waRes.message,
      "Artifact Updated At": new Date()
    });
    
    return waRes;
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Test ulang WA untuk sesi yang WA Status-nya PENDING
 */
function testBackgroundWaFix() {
  var ss = getSpreadsheet_();
  var DB = getSheetNames_();
  var sSession = ss.getSheetByName(DB.SESSIONS);
  var sessionRows = sSession.getDataRange().getValues();
  var sessionId = ''; // Akan dicari dari history
  
  // Cari 1 sesi terbaru yang punya WA Error atau WA Status = PENDING/FAILED
  for(var i = sessionRows.length - 1; i > 0; i--) {
    if(sessionRows[i][26] === 'PENDING' || sessionRows[i][26] === 'FAILED') {
      sessionId = sessionRows[i][0];
      break;
    }
  }
  
  if(!sessionId) {
    // Ambil sesi paling akhir saja
    sessionId = sessionRows[sessionRows.length - 1][0];
  }
  
  if(!sessionId) {
    console.log('Tidak ada sesi untuk dites');
    return;
  }
  
  console.log('Mengetes ulang WA untuk Session ID:', sessionId);
  var res = retryWaNotification(sessionId);
  console.log('Hasil Test:', res);
}
