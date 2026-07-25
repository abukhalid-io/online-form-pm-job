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
    } else if (action === 'gasHeaders') {
      result = { success: true, data: gasGetHeaders() };
    } else if (action === 'gasDetails') {
      result = { success: true, data: gasGetDetails(e.parameter.id_pm) };
    } else if (action === 'gasSchema') {
      result = { success: true, data: gasInspectSchema() };
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
    } else if (action === 'gasHeaderSubmit') {
      result = gasSubmitHeader(body.data);
    } else if (action === 'gasDetailSubmit') {
      result = gasSubmitDetail(body.data);
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
  var s = String(v || '').trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  // Sel tanggal tidak selalu lolos "instanceof Date" (mis. beda konteks eksekusi),
  // sehingga tanpa fallback ini frontend menerima Date.toString() penuh.
  var parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return s;
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

/* ============================================================================
 * PM GAS ANALYZER (CO / CO2 / O2)
 * ----------------------------------------------------------------------------
 * Spreadsheet TERPISAH dari ControlPanel. Isi ID-nya di GAS_SPREADSHEET_ID
 * di bawah, ambil dari URL sheet PM Analyzer:
 *   https://docs.google.com/spreadsheets/d/<INI_ID_NYA>/edit
 *
 * Struktur (data mulai baris 5, baris 1-3 judul, baris 4 header):
 *   PM_Header         A..M  -> D & J adalah FORMULA (Jenis Gas, Jumlah Iterasi)
 *   PM_Detail_Iterasi A..Q  -> H,I,L,M,N,P adalah FORMULA
 *                              (Standar Zero/Span, Deviasi Zero/Span, Status, Peringatan)
 *
 * Penulisan sengaja dipecah per blok kolom supaya sel formula TIDAK PERNAH
 * ditimpa. Menulis satu baris penuh akan menghapus rumusnya.
 * ==========================================================================*/

// Sheet GAS berada di spreadsheet yang SAMA dengan ControlPanel.
var GAS_SPREADSHEET_ID = SPREADSHEET_ID;
var GAS_SHEET = 'GAS';
var GAS_HEADER_SHEET = 'PM_Header';        // dipakai hanya kalau struktur 2-sheet dipilih
var GAS_DETAIL_SHEET = 'PM_Detail_Iterasi';
var GAS_DATA_START_ROW = 5;

/**
 * DIAGNOSTIK - baca struktur sheet GAS apa adanya supaya implementasi bisa
 * dibuat sesuai kolom yang benar-benar ada, bukan ditebak.
 * Panggil: ?action=gasSchema
 */
function gasInspectSchema() {
  var ss = SpreadsheetApp.openById(GAS_SPREADSHEET_ID);
  var sheet = ss.getSheetByName(GAS_SHEET);
  if (!sheet) {
    return {
      found: false,
      sheetTersedia: ss.getSheets().map(function (s) { return s.getName(); })
    };
  }
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var ambil = Math.min(lastRow, 10);
  var out = {
    found: true, sheet: GAS_SHEET, lastRow: lastRow, lastColumn: lastCol,
    barisAwal: [], formulaPerKolom: []
  };
  if (ambil > 0) out.barisAwal = sheet.getRange(1, 1, ambil, lastCol).getDisplayValues();
  // Rumus di baris data pertama -> menandai kolom mana yang tidak boleh ditulis app.
  if (lastRow >= 2) {
    var f = sheet.getRange(2, 1, Math.min(lastRow - 1, 5), lastCol).getFormulas();
    for (var c = 0; c < lastCol; c++) {
      for (var r = 0; r < f.length; r++) {
        if (f[r][c]) { out.formulaPerKolom.push({ kolom: c + 1, contoh: f[r][c] }); break; }
      }
    }
  }
  return out;
}

// Nomor kolom (A=1). Dipakai untuk menurunkan rumus ke baris baru.
var GAS_HEADER_FORMULA_COLS = [4, 10];              // D, J
var GAS_DETAIL_FORMULA_COLS = [8, 9, 12, 13, 14, 16]; // H, I, L, M, N, P

function getGasSheet_(name) {
  if (!GAS_SPREADSHEET_ID) {
    throw new Error('GAS_SPREADSHEET_ID belum diisi di Code.gs - modul Gas Analyzer belum aktif.');
  }
  var sheet = SpreadsheetApp.openById(GAS_SPREADSHEET_ID).getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" tidak ditemukan di spreadsheet Gas Analyzer.');
  return sheet;
}

/**
 * Baca baris data sebagai teks tampilan. getDisplayValues dipakai (bukan
 * getValues) supaya hasil formula ikut terbaca dan tanggal keluar sesuai format
 * sheet - bukan Date.toString() yang bikin frontend harus menormalkan lagi.
 */
function gasReadRows_(sheetName, lastCol) {
  var sheet = getGasSheet_(sheetName);
  var lastRow = sheet.getLastRow();
  if (lastRow < GAS_DATA_START_ROW) return [];
  return sheet.getRange(GAS_DATA_START_ROW, 1, lastRow - GAS_DATA_START_ROW + 1, lastCol)
    .getDisplayValues()
    .filter(function (r) { return String(r[0]).trim() !== ''; });
}

/**
 * Tambah 1 baris, hanya menulis blok kolom yang boleh diisi app.
 * chunks: [{ start: <kolom awal>, values: [...] }]
 */
function gasAppendRow_(sheetName, chunks, formulaCols) {
  var sheet = getGasSheet_(sheetName);
  var newRow = Math.max(sheet.getLastRow(), GAS_DATA_START_ROW - 1) + 1;

  chunks.forEach(function (c) {
    sheet.getRange(newRow, c.start, 1, c.values.length).setValues([c.values]);
  });

  // Turunkan rumus dari baris di atas HANYA kalau sel target masih kosong.
  // Pola fill-down: sel kosong -> rumus disalin. Pola ARRAYFORMULA: sel sudah
  // terisi otomatis -> dilewati, jadi ARRAYFORMULA-nya tidak dirusak.
  if (newRow > GAS_DATA_START_ROW) {
    formulaCols.forEach(function (col) {
      var target = sheet.getRange(newRow, col);
      if (target.getFormula() === '' && String(target.getValue()) === '') {
        var src = sheet.getRange(newRow - 1, col);
        if (src.getFormula() !== '') src.copyTo(target);
      }
    });
  }

  SpreadsheetApp.flush();
  return newRow;
}

function gasTimestampId_(prefix) {
  return prefix + '-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
}

/** Semua sesi PM (PM_Header) sebagai object siap pakai frontend. */
function gasGetHeaders() {
  return gasReadRows_(GAS_HEADER_SHEET, 13).map(function (r) {
    return {
      id_pm: r[0], tanggal: r[1], id_analyzer: r[2], jenis_gas: r[3], teknisi: r[4],
      id_tabung_zero: r[5], id_tabung_span: r[6], tekanan_zero: r[7], tekanan_span: r[8],
      jumlah_iterasi: r[9], status_akhir: r[10], waktu_selesai: r[11], catatan: r[12]
    };
  });
}

/** Iterasi verifikasi; kalau id_pm diisi, hanya untuk sesi tersebut. */
function gasGetDetails(idPm) {
  var rows = gasReadRows_(GAS_DETAIL_SHEET, 17);
  if (idPm) rows = rows.filter(function (r) { return r[1] === idPm; });
  return rows.map(function (r) {
    return {
      id_detail: r[0], id_pm: r[1], no_iterasi: r[2], tahap: r[3], waktu: r[4],
      pembacaan_zero: r[5], pembacaan_span: r[6], standar_zero: r[7], standar_span: r[8],
      toleransi_zero: r[9], toleransi_span: r[10], deviasi_zero: r[11], deviasi_span: r[12],
      status: r[13], tindakan_kalibrasi: r[14], peringatan: r[15], catatan: r[16]
    };
  });
}

function gasSubmitHeader(data) {
  if (!data) throw new Error('Data sesi kosong.');
  if (!data.id_analyzer) throw new Error('ID Analyzer wajib diisi.');
  var id = data.id_pm || gasTimestampId_('PM');

  gasAppendRow_(GAS_HEADER_SHEET, [
    { start: 1,  values: [id, data.tanggal || '', data.id_analyzer] },                     // A-C
    { start: 5,  values: [data.teknisi || '-', data.id_tabung_zero || '-',                 // E-I
                          data.id_tabung_span || '-', data.tekanan_zero || '-',
                          data.tekanan_span || '-'] },
    { start: 11, values: [data.status_akhir || 'Berjalan',                                 // K-M
                          data.waktu_selesai || '', data.catatan || ''] }
  ], GAS_HEADER_FORMULA_COLS);

  return { success: true, id_pm: id, message: 'Sesi PM ' + id + ' dibuat.' };
}

function gasSubmitDetail(data) {
  if (!data || !data.id_pm) throw new Error('id_pm wajib diisi.');
  var iterasi = Number(data.no_iterasi || 0);
  if (!iterasi || iterasi < 1) throw new Error('Nomor iterasi tidak valid.');
  if (iterasi > 10) throw new Error('Iterasi maksimal 10 per sesi PM.');
  var id = data.id_detail || gasTimestampId_('DET');

  gasAppendRow_(GAS_DETAIL_SHEET, [
    { start: 1,  values: [id, data.id_pm, iterasi, data.tahap || '', data.waktu || '',     // A-G
                          data.pembacaan_zero === undefined ? '' : data.pembacaan_zero,
                          data.pembacaan_span === undefined ? '' : data.pembacaan_span] },
    { start: 10, values: [data.toleransi_zero === undefined ? '' : data.toleransi_zero,    // J-K
                          data.toleransi_span === undefined ? '' : data.toleransi_span] },
    { start: 15, values: [data.tindakan_kalibrasi || ''] },                                // O
    { start: 17, values: [data.catatan || ''] }                                            // Q
  ], GAS_DETAIL_FORMULA_COLS);

  return { success: true, id_detail: id, message: 'Iterasi ke-' + iterasi + ' tersimpan.' };
}
