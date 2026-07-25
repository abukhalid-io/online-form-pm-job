/**
 * PM CONTROL PANEL - Backend API (versi Netlify + Apps Script)
 * Sheet target : ControlPanel
 * Spreadsheet  : https://docs.google.com/spreadsheets/d/1eZ0rqnA9dg97F4xPR8bJeL3E1qP6-qzw9tevflx2a_A/edit
 *
 * Versi ini TIDAK lagi merender HTML (doGet tidak return HtmlService).
 * Frontend (index.html) di-host terpisah di Netlify, dan memanggil Web App
 * ini lewat fetch() sebagai API JSON biasa:
 *
 *   GET  ?action=list                -> daftar kode panel yang sudah ada
 *   GET  ?action=template&panel=XXX  -> template komponen utk 1 kode panel
 *   GET  ?action=records             -> semua baris data (utk halaman Data)
 *   POST { action: 'submit', data }  -> tambah baris baru
 *   POST { action: 'update', rowIndex, data } -> update baris tertentu
 *   POST { action: 'delete', rowIndex }       -> hapus baris tertentu
 *
 * PENTING soal CORS: request POST dari Netlify harus dikirim dengan
 * Content-Type: text/plain (BUKAN application/json), supaya browser tidak
 * mengirim preflight OPTIONS (Apps Script tidak bisa menjawab preflight).
 * Isinya tetap teks JSON biasa, nanti di-parse manual lewat
 * e.postData.contents. Lihat index.html untuk contoh pemanggilannya.
 *
 * Cara pasang ulang:
 * 1. Buka project Apps Script yang sudah ada, ganti seluruh isi Code.gs
 *    dengan isi file ini.
 * 2. Tambahkan 9 kolom baru di sheet ControlPanel (lihat komentar HEADERS
 *    di bawah) supaya urutannya sama persis.
 * 3. Deploy > Manage deployments > edit (pensil) > Version: New version >
 *    Deploy. URL /exec akan TETAP SAMA, tidak perlu update di index.html.
 * 4. Pastikan "Who has access" = Anyone (bukan "Anyone with Google
 *    account"), supaya bisa dipanggil dari domain Netlify tanpa login.
 */

var SPREADSHEET_ID = '1eZ0rqnA9dg97F4xPR8bJeL3E1qP6-qzw9tevflx2a_A';
var SHEET_NAME = 'ControlPanel';

// Urutan header persis harus sama dengan urutan kolom di sheet (mulai kolom A).
// 29 kolom pertama = struktur lama (sudah ada). 9 kolom terakhir = BARU,
// perlu ditambahkan manual di sheet (PANEL DESC s/d CAB FAN RUNNING).
var HEADERS = [
  'DATE', 'PANEL', 'BREAKER', 'PSU 24 VDC', 'SWITCH HUB', 'FUSE BOARD',           // 0-5
  'RESISTOR BOARD', 'BARIER', 'C1', 'C1 STATUS', 'C2', 'C2 STATUS',               // 6-11
  'C3', 'C3 STATUS', 'C4', 'C4-MA', 'C4-MA STATUS', 'C4-MB', 'C4-MB STATUS',      // 12-18
  'C4-MC', 'C4-MC STATUS', 'LAMP1', 'LAMP2',                                      // 19-22
  'FAN1 STATUS', 'FAN1 AMPERE', 'FAN2 STATUS', 'FAN2 AMPERE',                     // 23-26
  'PANEL CONDITION', 'NOTE',                                                      // 27-28
  'PANEL DESC', 'REPORT TITLE',                                                   // 29-30 (BARU)
  'CAB INSPECTION', 'CAB DUST FILTERS', 'CAB DOOR SWITCHES', 'CAB FAN FUNCTION',  // 31-34 (BARU)
  'CAB CABLE ROUTING', 'CAB DOOR CLOSURE', 'CAB FAN RUNNING'                      // 35-37 (BARU)
];

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'list';
  var result;
  try {
    if (action === 'list') {
      result = { success: true, data: getPanelList() };
    } else if (action === 'template') {
      result = { success: true, data: getPanelTemplate(e.parameter.panel) };
    } else if (action === 'records') {
      result = { success: true, data: getAllRecords() };
    } else {
      result = { success: false, message: 'Action tidak dikenal: ' + action };
    }
  } catch (err) {
    result = { success: false, message: err.message };
  }
  return jsonOutput_(result);
}

function doPost(e) {
  var result;
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    if (action === 'submit') {
      result = submitPmData(body.data);
    } else if (action === 'update') {
      result = updateRecord(body.rowIndex, body.data);
    } else if (action === 'delete') {
      result = deleteRecord(body.rowIndex);
    } else {
      result = { success: false, message: 'Action tidak dikenal: ' + action };
    }
  } catch (err) {
    result = { success: false, message: err.message };
  }
  return jsonOutput_(result);
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" tidak ditemukan di spreadsheet.');
  return sheet;
}

function isPresent_(v) {
  if (v === undefined || v === null) return false;
  var s = String(v).trim();
  return s !== '' && s !== '-';
}

function formatDate_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v || '');
}

/** Daftar kode PANEL unik yang sudah pernah diinput (utk grid pilihan panel). */
function getPanelList() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  var values = sheet.getRange(3, 2, lastRow - 2, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (v) { return v; });
  var unique = Array.from(new Set(values));
  unique.sort();
  return unique;
}

/**
 * Cari histori PALING BARU untuk sebuah tag PANEL, tentukan komponen mana
 * yang ada secara fisik (isPresent_), plus nilai identitas tetap
 * (tipe controller, deskripsi panel, checklist kabinet terakhir) sebagai
 * bahan auto-fill form.
 */
function getPanelTemplate(panelTag) {
  if (!panelTag) return { found: false };
  var target = String(panelTag).trim().toLowerCase();
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return { found: false };
  var values = sheet.getRange(3, 1, lastRow - 2, HEADERS.length).getValues();

  for (var i = values.length - 1; i >= 0; i--) {
    var row = values[i];
    if (String(row[1]).trim().toLowerCase() === target) {
      return {
        found: true,
        panel: String(row[1]).trim(),
        desc: String(row[29] || ''),
        reportTitle: String(row[30] || ''),
        breaker: isPresent_(row[2]),
        psu24vdc: isPresent_(row[3]),
        switchHub: isPresent_(row[4]),
        fuseBoard: isPresent_(row[5]),
        resistorBoard: isPresent_(row[6]),
        barier: isPresent_(row[7]),
        c1: { exists: isPresent_(row[8]), type: String(row[8] || '') },
        c2: { exists: isPresent_(row[10]), type: String(row[10] || '') },
        c3: { exists: isPresent_(row[12]), type: String(row[12] || '') },
        c4: { exists: isPresent_(row[14]), type: String(row[14] || '') },
        c4ma: { exists: isPresent_(row[15]), id: String(row[15] || '') },
        c4mb: { exists: isPresent_(row[17]), id: String(row[17] || '') },
        c4mc: { exists: isPresent_(row[19]), id: String(row[19] || '') },
        lamp1: isPresent_(row[21]),
        lamp2: isPresent_(row[22]),
        fan1: isPresent_(row[23]) || isPresent_(row[24]),
        fan2: isPresent_(row[25]) || isPresent_(row[26]),
        fan1Last: String(row[24] || ''),
        fan2Last: String(row[26] || ''),
        cab: {
          cabInspection: String(row[31] || 'Good'),
          cabDustFilters: String(row[32] || 'Clean'),
          cabDoorSwitches: String(row[33] || 'Functioning'),
          cabFanFunction: String(row[34] || 'Running'),
          cabCableRouting: String(row[35] || 'Proper'),
          cabDoorClosure: String(row[36] || 'Secure'),
          cabFanRunning: String(row[37] || 'Running')
        }
      };
    }
  }
  return { found: false };
}

/** Susun 1 baris array (38 kolom) dari object data yang dikirim frontend. */
function buildRow_(data) {
  var dash = '-';
  return [
    data.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    data.panel || dash,
    data.breaker || dash,
    data.psu24vdc || dash,
    data.switchHub || dash,
    data.fuseBoard || dash,
    data.resistorBoard || dash,
    data.barier || dash,
    data.c1 || dash,
    data.c1Status || dash,
    data.c2 || dash,
    data.c2Status || dash,
    data.c3 || dash,
    data.c3Status || dash,
    data.c4 || dash,
    data.c4ma || dash,
    data.c4maStatus || dash,
    data.c4mb || dash,
    data.c4mbStatus || dash,
    data.c4mc || dash,
    data.c4mcStatus || dash,
    data.lamp1 || dash,
    data.lamp2 || dash,
    data.fan1Status || dash,
    data.fan1Ampere || dash,
    data.fan2Status || dash,
    data.fan2Ampere || dash,
    data.panelCondition || dash,
    data.note || '',
    data.desc || dash,
    data.reportTitle || dash,
    data.cabInspection || dash,
    data.cabDustFilters || dash,
    data.cabDoorSwitches || dash,
    data.cabFanFunction || dash,
    data.cabCableRouting || dash,
    data.cabDoorClosure || dash,
    data.cabFanRunning || dash
  ];
}

function submitPmData(data) {
  if (!data || !data.panel) throw new Error('PANEL wajib diisi.');
  var sheet = getSheet_();
  sheet.appendRow(buildRow_(data));
  return { success: true, row: sheet.getLastRow(), message: 'Data panel ' + data.panel + ' berhasil disimpan.' };
}

function updateRecord(rowIndex, data) {
  if (!rowIndex) throw new Error('rowIndex wajib diisi.');
  if (!data || !data.panel) throw new Error('PANEL wajib diisi.');
  var sheet = getSheet_();
  sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([buildRow_(data)]);
  return { success: true, row: rowIndex, message: 'Data panel ' + data.panel + ' berhasil diperbarui.' };
}

function deleteRecord(rowIndex) {
  if (!rowIndex) throw new Error('rowIndex wajib diisi.');
  var sheet = getSheet_();
  sheet.deleteRow(Number(rowIndex));
  return { success: true, message: 'Baris ' + rowIndex + ' berhasil dihapus.' };
}

/**
 * Ubah 1 baris sheet jadi object record siap pakai frontend: daftar
 * komponen yang benar-benar ada (bukan '-'), checklist kabinet, dan
 * kondisi/catatan. Struktur ini sama persis dengan yang dipakai fungsi
 * buildReportText() di index.html supaya logic report tidak perlu diduplikasi.
 */
function parseRowToRecord_(row, rowIndex) {
  var components = [];
  function pushStatus(idx, label) {
    if (isPresent_(row[idx])) components.push({ kind: 'status', label: label, value: String(row[idx]) });
  }
  function pushPair(statusIdx, valueIdx, label, unit, suffix) {
    if (isPresent_(row[statusIdx]) || isPresent_(row[valueIdx])) {
      components.push({
        kind: 'pair', label: label, unit: unit, suffix: suffix || '',
        value: String(row[statusIdx] || ''), reading: String(row[valueIdx] || '')
      });
    }
  }
  pushStatus(2, 'Breaker');
  pushStatus(3, 'PSU 24 VDC');
  pushStatus(4, 'Switch Hub');
  pushStatus(5, 'Fuse Board');
  pushStatus(6, 'Resistor Board');
  pushStatus(7, 'Barier');
  if (isPresent_(row[8])) pushStatus(9, 'Controller C1 (' + row[8] + ')');
  if (isPresent_(row[10])) pushStatus(11, 'Controller C2 (' + row[10] + ')');
  if (isPresent_(row[12])) pushStatus(13, 'Controller C3 (' + row[12] + ')');
  if (isPresent_(row[15])) pushPair(16, 15, 'Woodward Module A', 'RPM');
  if (isPresent_(row[17])) pushPair(18, 17, 'Woodward Module B', 'RPM');
  if (isPresent_(row[19])) pushPair(20, 19, 'Woodward Module C', 'RPM');
  pushStatus(21, 'Lamp 1');
  pushStatus(22, 'Lamp 2');
  pushPair(23, 24, 'Fan 1', 'Ampere', ' A');
  pushPair(25, 26, 'Fan 2', 'Ampere', ' A');

  return {
    id: rowIndex,
    rowIndex: rowIndex,
    date: formatDate_(row[0]),
    panel: String(row[1] || ''),
    data: {
      desc: String(row[29] || ''),
      title: String(row[30] || ''),
      cab: {
        cabInspection: String(row[31] || 'Good'),
        cabDustFilters: String(row[32] || 'Clean'),
        cabDoorSwitches: String(row[33] || 'Functioning'),
        cabFanFunction: String(row[34] || 'Running'),
        cabCableRouting: String(row[35] || 'Proper'),
        cabDoorClosure: String(row[36] || 'Secure'),
        cabFanRunning: String(row[37] || 'Running')
      },
      components: components,
      condition: String(row[27] || 'CLEAN'),
      note: String(row[28] || '')
    }
  };
}

/** Semua baris data, dipakai halaman "Data Tersimpan" di frontend. */
function getAllRecords() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  var values = sheet.getRange(3, 1, lastRow - 2, HEADERS.length).getValues();
  var records = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!row[1]) continue;
    records.push(parseRowToRecord_(row, 3 + i));
  }
  return records;
}
