// ═══════════════════════════════════════════════════════════════
//  DC CLEAN — CUSTOMER ENGAGEMENT CHALLENGE 2026
//  Code.gs — paste dalam Extensions > Apps Script
//  Deploy: Execute as Me, Access: Anyone
// ═══════════════════════════════════════════════════════════════

const SHEET_NAME = "Data Pekerja";
const FIRST_ROW  = 5;
const LAST_ROW   = 24;

function doGet(e) {
  var p = e.parameter || {};

  if (p.action === 'getData') {
    var data = getData();
    var json = JSON.stringify(data);
    // JSONP — bypass CORS
    if (p.callback) {
      return ContentService
        .createTextOutput(p.callback + '(' + json + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('DC Clean — Customer Engagement Challenge')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}
  // ── Web App biasa (existing) ──
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('DC Clean — Customer Engagement Challenge')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ── Ambil File ID dari pelbagai format Google Drive URL ──────────
function getFileIdFromUrl(url) {
  if (!url) return null;
  var s = url.toString().trim();
  // Format: /d/FILE_ID/
  var m = s.match(/\/d\/([a-zA-Z0-9_-]{10,})/);
  if (m) return m[1];
  // Format: id=FILE_ID
  var m2 = s.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  if (m2) return m2[1];
  // Format: /open?id= atau /uc?id=
  var m3 = s.match(/\/(?:open|uc)\?(?:.*&)?id=([a-zA-Z0-9_-]{10,})/);
  if (m3) return m3[1];
  return null;
}

// ── Tukar Google Drive URL ke format direct (untuk img src) ──────
function getDriveDirectUrl(url) {
  if (!url) return '';
  var s = url.toString().trim();
  var fileId = getFileIdFromUrl(s);
  if (fileId) {
    // Guna thumbnail API — laju & tak perlu base64
    return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w200';
  }
  // Bukan Drive URL — return as-is (cth: http/https image terus)
  if (s.match(/^https?:\/\//)) return s;
  return '';
}

// ── Fetch photo sebagai base64 (fallback jika direct URL tak work) ──
function getPhotoBase64(url) {
  try {
    var fileId = getFileIdFromUrl(url);
    if (!fileId) return '';
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    var base64 = Utilities.base64Encode(blob.getBytes());
    return 'data:' + blob.getContentType() + ';base64,' + base64;
  } catch(e) {
    Logger.log('Photo base64 error: ' + e.message);
    return '';
  }
}

function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(SHEET_NAME);

  if (!ws) {
    Logger.log('RALAT: Sheet "' + SHEET_NAME + '" tidak dijumpai!');
    return { ok: false, error: 'Sheet tidak dijumpai', workers: [], targets: {} };
  }

  // ── DEBUG: Log header row untuk verify kolum ─────────────────
  var headerRow = ws.getRange(4, 1, 1, 20).getValues()[0];
  Logger.log('HEADER ROW 4: ' + JSON.stringify(headerRow));

  const lastCol = 16; // kolum A hingga P
  const numRows = LAST_ROW - FIRST_ROW + 1;
  const allData = ws.getRange(FIRST_ROW, 1, numRows, lastCol).getValues();

  Logger.log('Jumlah row data: ' + allData.length);

  const workers = [];
  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const name = row[1]; // kolum B
    if (!name || name.toString().trim() === '') continue;

    // ── Photo: cuba direct URL dulu (laju), fallback ke base64 ──
    const photoUrl = row[15] ? row[15].toString().trim() : ''; // kolum P
    var photo = '';
    if (photoUrl) {
      // Cuba guna Drive thumbnail URL (tak perlu base64, laju)
      var directUrl = getDriveDirectUrl(photoUrl);
      if (directUrl) {
        photo = directUrl; // hantar URL terus ke frontend
      } else {
        // Last resort: base64 encode
        photo = getPhotoBase64(photoUrl);
      }
    }

    // ── DEBUG: Log setiap pekerja ────────────────────────────────
    Logger.log('Row ' + (FIRST_ROW + i) + ' | Nama: ' + name + 
               ' | Review(C): ' + row[2] + 
               ' | Video(D): ' + row[3] + 
               ' | Pts(J): ' + row[9] + 
               ' | Ranking(L): ' + row[11] + 
               ' | Photo(P): ' + (photo ? 'ADA' : 'TIADA'));

    workers.push({
      name:      name.toString().trim(),
      review:    Number(row[2])  || 0,  // col C
      video:     Number(row[3])  || 0,  // col D
      mattPro:   Number(row[4])  || 0,  // col E
      mattSel:   Number(row[5])  || 0,  // col F
      recurring: Number(row[6])  || 0,  // col G
      daily:     Number(row[7])  || 0,  // col H
      newSvc:    Number(row[8])  || 0,  // col I
      pts:       Number(row[9])  || 0,  // col J
      ranking:   row[11] ? row[11].toString().trim() : '—', // col L
      actReward: Number(row[12]) || 0,  // col M
      bonus:     Number(row[13]) || 0,  // col N
      total:     Number(row[14]) || 0,  // col O
      photo:     photo
    });
  }

  Logger.log('Pekerja berjaya diload: ' + workers.length);

  // ── Targets dari sheet "Sasaran Kempen" ──────────────────────
  var targets = {};
  try {
    const ws2 = ss.getSheetByName('Sasaran Kempen');
    if (ws2) {
      const tData = ws2.getRange(3, 1, 7, 3).getValues();
      targets = {
        review:    { target: Number(tData[0][1]) || 1, semasa: Number(tData[0][2]) || 0 },
        video:     { target: Number(tData[1][1]) || 1, semasa: Number(tData[1][2]) || 0 },
        mattPro:   { target: Number(tData[2][1]) || 1, semasa: Number(tData[2][2]) || 0 },
        mattSel:   { target: Number(tData[3][1]) || 1, semasa: Number(tData[3][2]) || 0 },
        recurring: { target: Number(tData[4][1]) || 1, semasa: Number(tData[4][2]) || 0 },
        daily:     { target: Number(tData[5][1]) || 1, semasa: Number(tData[5][2]) || 0 },
        newSvc:    { target: Number(tData[6][1]) || 1, semasa: Number(tData[6][2]) || 0 },
      };
      Logger.log('Targets loaded: ' + JSON.stringify(targets));
    } else {
      Logger.log('AMARAN: Sheet "Sasaran Kempen" tidak dijumpai!');
      // Default targets supaya dashboard tak crash
      targets = {
        review:    { target: 100, semasa: 0 },
        video:     { target: 50,  semasa: 0 },
        mattPro:   { target: 30,  semasa: 0 },
        mattSel:   { target: 30,  semasa: 0 },
        recurring: { target: 50,  semasa: 0 },
        daily:     { target: 50,  semasa: 0 },
        newSvc:    { target: 20,  semasa: 0 },
      };
    }
  } catch(err) {
    Logger.log('Targets error: ' + err.message);
  }

  return { ok: true, lastUpdated: new Date().toISOString(), workers, targets };
}

// ── Jalankan ini dalam Apps Script Editor untuk debug ───────────
function testGetData() {
  var result = getData();
  Logger.log('=== RESULT ===');
  Logger.log('Workers: ' + result.workers.length);
  Logger.log('OK: ' + result.ok);
  result.workers.forEach(function(w, i) {
    Logger.log((i+1) + '. ' + w.name + 
               ' | R:' + w.review + 
               ' V:' + w.video + 
               ' MP:' + w.mattPro + 
               ' MS:' + w.mattSel + 
               ' Rec:' + w.recurring + 
               ' D:' + w.daily + 
               ' NS:' + w.newSvc + 
               ' Pts:' + w.pts + 
               ' | Photo: ' + (w.photo ? 'ADA (' + (w.photo.startsWith('data:') ? 'base64' : 'URL') + ')' : 'TIADA'));
  });
}

// ── Check kolum dalam sheet ──────────────────────────────────────
function debugColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(SHEET_NAME);
  
  Logger.log('=== DEBUG KOLUM ===');
  Logger.log('Sheet name: ' + ws.getName());
  
  // Header rows
  for (var r = 1; r <= 5; r++) {
    var rowData = ws.getRange(r, 1, 1, 16).getValues()[0];
    Logger.log('Row ' + r + ': ' + rowData.join(' | '));
  }
  
  // First data row
  Logger.log('=== ROW DATA PERTAMA (row ' + FIRST_ROW + ') ===');
  var firstData = ws.getRange(FIRST_ROW, 1, 1, 16).getValues()[0];
  'ABCDEFGHIJKLMNOP'.split('').forEach(function(col, idx) {
    Logger.log('Col ' + col + ' (' + idx + '): ' + firstData[idx]);
  });
}

// ── Log Review dari dashboard ke Google Sheets ───────────────────
function logReview(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Reviews');

    // Buat sheet baru kalau belum ada
    if (!sheet) {
      sheet = ss.insertSheet('Reviews');
      sheet.appendRow(['ID','Date','Worker','Branch','Customer','Address','Category','Status','Notes']);
      sheet.getRange(1, 1, 1, 9)
        .setFontWeight('bold')
        .setBackground('#0A1628')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      data.id          || '',
      data.date        || '',
      data.workerName  || '',
      data.branch      || '',
      data.customerName|| '',
      data.customerAddr|| '',
      data.category    || '',
      data.status      || '',
      data.notes       || ''
    ]);

    Logger.log('Review saved: ' + data.customerName + ' | ' + data.workerName);
    return { ok: true };

  } catch(e) {
    Logger.log('logReview error: ' + e.message);
    return { ok: false, error: e.message };
  }
  function doPost(e) {
  try {
    var action = e.parameter.action;
    if (action === 'logReview') {
      var data = JSON.parse(e.parameter.data);
      var result = logReview(data);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
}
