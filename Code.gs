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
    } else if (action === 'gasAnalyzers') {
      result = { success: true, data: gasGetAnalyzers() };
    } else if (action === 'gasCylinders') {
      result = { success: true, data: gasGetCylinders() };
    } else if (action === 'gasSessions') {
      result = { success: true, data: gasGetSessions() };
    } else if (action === 'gasIterations') {
      result = { success: true, data: gasGetIterations(e.parameter.id_pm) };
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
    } else if (action === 'gasSessionSubmit') {
      result = gasSubmitSession(body.data);
    } else if (action === 'gasSessionStatus') {
      result = gasUpdateSessionStatus(body.id_pm, body.status_akhir, body.catatan);
    } else if (action === 'gasIterationSubmit') {
      result = gasSubmitIteration(body.data);
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
 * 4 sheet relasional di spreadsheet yang SAMA dengan ControlPanel:
 *   Master_Analyzer         - daftar analyzer (diisi manual sekali)
 *   Master_Cylinder_Gas     - daftar tabung gas standar zero/span
 *   PM_Gas_Header           - 1 baris = 1 sesi PM (ringkasan)
 *   PM_Gas_Detail_Iterasi   - 1 baris = 1 kali verifikasi/kalibrasi
 *
 * Semua 4 sheet memakai pola yang sama: baris 1-2 judul (merged), baris 3
 * kosong, baris 4 header, data mulai baris 5.
 *
 * Kolom hijau ("(auto)" di judul) berisi RUMUS dan TIDAK PERNAH ditulis
 * langsung oleh app:
 *   PM_Gas_Header:         D (Jenis Gas), J (Jumlah Iterasi)
 *   PM_Gas_Detail_Iterasi: H (Standar Zero), I (Standar Span),
 *                          L (Deviasi Zero), M (Deviasi Span),
 *                          N (Status), P (Peringatan)
 *
 * SETUP: jalankan gasSetupFormulas() SEKALI dari editor Apps Script untuk
 * memasang rumus tsb ke seluruh baris data yang sudah ada, dan menyiapkan
 * baris pertama sebagai sumber salinan rumus untuk baris berikutnya.
 * ==========================================================================*/

var GAS_DATA_START_ROW = 5;
var GAS_MAX_ITERASI = 10;

var GAS_ANALYZER_SHEET = 'Master_Analyzer';
var GAS_CYLINDER_SHEET = 'Master_Cylinder_Gas';
var GAS_HEADER_SHEET = 'PM_Gas_Header';
var GAS_DETAIL_SHEET = 'PM_Gas_Detail_Iterasi';

// Kolom rumus (1-indexed, A=1).
var GAS_HEADER_FORMULA_COLS = [4, 10];                // D, J
var GAS_DETAIL_FORMULA_COLS = [8, 9, 12, 13, 14, 16]; // H, I, L, M, N, P

function getGasSheetNamed_(name) {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" tidak ditemukan.');
  return sheet;
}

/**
 * Baris kosong berikutnya, dicari lewat kolom A - BUKAN getLastRow(), karena
 * sel berisi rumus tetap dihitung "terisi" oleh getLastRow() walau hasilnya
 * tampil kosong. Tanpa ini, baris data bisa terlewat atau tertimpa.
 */
function gasNextRow_(sheet) {
  var last = sheet.getLastRow();
  if (last < GAS_DATA_START_ROW) return GAS_DATA_START_ROW;
  var vals = sheet.getRange(GAS_DATA_START_ROW, 1, last - GAS_DATA_START_ROW + 1, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).trim() !== '') return GAS_DATA_START_ROW + i + 1;
  }
  return GAS_DATA_START_ROW;
}

/** ID berurutan bergaya "PM-0001", dicari dari angka terbesar yang sudah ada + 1. */
function gasNextSeqId_(sheet, idCol, prefix) {
  var last = sheet.getLastRow();
  var maxNum = 0;
  if (last >= GAS_DATA_START_ROW) {
    var vals = sheet.getRange(GAS_DATA_START_ROW, idCol, last - GAS_DATA_START_ROW + 1, 1).getValues();
    vals.forEach(function (r) {
      var m = String(r[0]).match(new RegExp('^' + prefix + '-(\\d+)$'));
      if (m) maxNum = Math.max(maxNum, Number(m[1]));
    });
  }
  return prefix + '-' + String(maxNum + 1).padStart(4, '0');
}

// ---------- rumus ----------
function gasHeaderFormulasFor_(r) {
  return [
    '=IFERROR(VLOOKUP($C' + r + ',' + GAS_ANALYZER_SHEET + '!$A:$C,3,FALSE),"")',
    '=COUNTIF(' + GAS_DETAIL_SHEET + '!$B:$B,$A' + r + ')'
  ];
}
function gasDetailFormulasFor_(r) {
  var tabungZero = 'IFERROR(VLOOKUP($B' + r + ',' + GAS_HEADER_SHEET + '!$A:$G,6,FALSE),"")';
  var tabungSpan = 'IFERROR(VLOOKUP($B' + r + ',' + GAS_HEADER_SHEET + '!$A:$G,7,FALSE),"")';
  var H = '=IFERROR(VLOOKUP(' + tabungZero + ',' + GAS_CYLINDER_SHEET + '!$A:$G,7,FALSE),"")';
  var I = '=IFERROR(VLOOKUP(' + tabungSpan + ',' + GAS_CYLINDER_SHEET + '!$A:$H,8,FALSE),"")';
  var L = '=IF(OR($F' + r + '="",$H' + r + '=""),"",ABS($F' + r + '-$H' + r + '))';
  var M = '=IF(OR($G' + r + '="",$I' + r + '="",$I' + r + '=0),"",ABS($G' + r + '-$I' + r + ')/$I' + r + '*100)';
  var N = '=IF(AND($L' + r + '="",$M' + r + '=""),"",' +
    'IF(AND(IF($L' + r + '="",TRUE,$L' + r + '<=$J' + r + '),IF($M' + r + '="",TRUE,$M' + r + '<=$K' + r + ')),' +
    '"Lolos","Tidak Lolos"))';
  var P = '=IF($C' + r + '>=10,IF($N' + r + '="Tidak Lolos","STOP - Sudah 10x, ESKALASI ke Supervisor/Ganti Alat",""),' +
    'IF($C' + r + '>=8,"Perhatian - mendekati batas maks 10x",""))';
  return [H, I, L, M, N, P];
}
function gasEnsureFormulas_(sheet, r, cols, formulas) {
  for (var i = 0; i < cols.length; i++) {
    var cell = sheet.getRange(r, cols[i]);
    if (cell.getFormula() === '') cell.setFormula(formulas[i]);
  }
}

/**
 * Jalankan SEKALI dari editor Apps Script. Memasang rumus ke SEMUA baris data
 * yang sudah ada di PM_Gas_Header dan PM_Gas_Detail_Iterasi (kolom hijau saat
 * ini berisi nilai statis hasil isian manual, bukan rumus hidup), tanpa
 * menimpa kolom lain.
 */
function gasSetupFormulas() {
  var hdr = getGasSheetNamed_(GAS_HEADER_SHEET);
  var lastH = hdr.getLastRow();
  var nH = 0;
  for (var r = GAS_DATA_START_ROW; r <= lastH; r++) {
    if (String(hdr.getRange(r, 1).getValue()).trim() === '') continue;
    gasEnsureFormulas_(hdr, r, GAS_HEADER_FORMULA_COLS, gasHeaderFormulasFor_(r));
    nH++;
  }

  var det = getGasSheetNamed_(GAS_DETAIL_SHEET);
  var lastD = det.getLastRow();
  var nD = 0;
  for (var r2 = GAS_DATA_START_ROW; r2 <= lastD; r2++) {
    if (String(det.getRange(r2, 1).getValue()).trim() === '') continue;
    gasEnsureFormulas_(det, r2, GAS_DETAIL_FORMULA_COLS, gasDetailFormulasFor_(r2));
    nD++;
  }

  SpreadsheetApp.flush();
  return 'Rumus terpasang: ' + nH + ' baris PM_Gas_Header, ' + nD + ' baris PM_Gas_Detail_Iterasi.';
}

// ---------- baca data ----------
function gasGetAnalyzers() {
  var sheet = getGasSheetNamed_(GAS_ANALYZER_SHEET);
  var last = sheet.getLastRow();
  if (last < GAS_DATA_START_ROW) return [];
  return sheet.getRange(GAS_DATA_START_ROW, 1, last - GAS_DATA_START_ROW + 1, 11).getDisplayValues()
    .filter(function (r) { return String(r[0]).trim() !== ''; })
    .map(function (r) {
      return {
        id: r[0], nama: r[1], jenis_gas: r[2], lokasi: r[3], range_ukur: r[4], satuan: r[5],
        merk: r[6], no_seri: r[7], tgl_instalasi: r[8], status: r[9], keterangan: r[10]
      };
    });
}
function gasGetCylinders() {
  var sheet = getGasSheetNamed_(GAS_CYLINDER_SHEET);
  var last = sheet.getLastRow();
  if (last < GAS_DATA_START_ROW) return [];
  return sheet.getRange(GAS_DATA_START_ROW, 1, last - GAS_DATA_START_ROW + 1, 14).getDisplayValues()
    .filter(function (r) { return String(r[0]).trim() !== ''; })
    .map(function (r) {
      return {
        id: r[0], jenis_gas: r[1], kandungan: r[2], konsentrasi: r[3], satuan_konsentrasi: r[4],
        fungsi: r[5], standar_zero: r[6], standar_span: r[7], tekanan_awal: r[8],
        tekanan_terakhir: r[9], tgl_kadaluarsa: r[10], status: r[11], supplier: r[12], keterangan: r[13]
      };
    });
}
function gasGetSessions() {
  var sheet = getGasSheetNamed_(GAS_HEADER_SHEET);
  var last = sheet.getLastRow();
  if (last < GAS_DATA_START_ROW) return [];
  var vals = sheet.getRange(GAS_DATA_START_ROW, 1, last - GAS_DATA_START_ROW + 1, 13).getDisplayValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (String(r[0]).trim() === '') continue;
    out.push({
      rowIndex: GAS_DATA_START_ROW + i,
      id_pm: r[0], tanggal: r[1], id_analyzer: r[2], jenis_gas: r[3], teknisi: r[4],
      id_tabung_zero: r[5], id_tabung_span: r[6], tekanan_zero: r[7], tekanan_span: r[8],
      jumlah_iterasi: r[9], status_akhir: r[10], waktu_selesai: r[11], catatan: r[12]
    });
  }
  return out;
}
function gasGetIterations(idPm) {
  var sheet = getGasSheetNamed_(GAS_DETAIL_SHEET);
  var last = sheet.getLastRow();
  if (last < GAS_DATA_START_ROW) return [];
  var vals = sheet.getRange(GAS_DATA_START_ROW, 1, last - GAS_DATA_START_ROW + 1, 17).getDisplayValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (String(r[0]).trim() === '') continue;
    if (idPm && r[1] !== idPm) continue;
    out.push({
      rowIndex: GAS_DATA_START_ROW + i,
      id_detail: r[0], id_pm: r[1], no_iterasi: r[2], tahap: r[3], waktu: r[4],
      baca_zero: r[5], baca_span: r[6], standar_zero: r[7], standar_span: r[8],
      tol_zero: r[9], tol_span: r[10], deviasi_zero: r[11], deviasi_span: r[12],
      status: r[13], tindakan: r[14], peringatan: r[15], catatan: r[16]
    });
  }
  return out;
}

// ---------- tulis data ----------
function gasSubmitSession(d) {
  if (!d) throw new Error('Data sesi kosong.');
  if (!d.id_analyzer) throw new Error('Analyzer wajib dipilih.');
  var sheet = getGasSheetNamed_(GAS_HEADER_SHEET);
  var row = gasNextRow_(sheet);
  var id = gasNextSeqId_(sheet, 1, 'PM');
  var tz = Session.getScriptTimeZone();
  var dash = '-';

  // A-C (identitas), lompati D (rumus), E-I, lompati J (rumus), K-M.
  sheet.getRange(row, 1, 1, 3).setValues([[
    id, d.tanggal || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), d.id_analyzer
  ]]);
  sheet.getRange(row, 5, 1, 5).setValues([[
    d.teknisi || dash, d.id_tabung_zero || dash, d.id_tabung_span || dash,
    d.tekanan_zero === undefined || d.tekanan_zero === '' ? dash : d.tekanan_zero,
    d.tekanan_span === undefined || d.tekanan_span === '' ? dash : d.tekanan_span
  ]]);
  sheet.getRange(row, 11, 1, 3).setValues([[d.status_akhir || 'Berjalan', d.waktu_selesai || '', d.catatan || '']]);

  gasEnsureFormulas_(sheet, row, GAS_HEADER_FORMULA_COLS, gasHeaderFormulasFor_(row));
  SpreadsheetApp.flush();
  return { success: true, id_pm: id, row: row, message: 'Sesi PM ' + id + ' dibuat.' };
}

function gasUpdateSessionStatus(idPm, statusAkhir, catatan) {
  if (!idPm) throw new Error('id_pm wajib diisi.');
  var sheet = getGasSheetNamed_(GAS_HEADER_SHEET);
  var last = sheet.getLastRow();
  var col = sheet.getRange(GAS_DATA_START_ROW, 1, Math.max(last - GAS_DATA_START_ROW + 1, 0), 1).getValues();
  var row = -1;
  for (var i = 0; i < col.length; i++) { if (col[i][0] === idPm) { row = GAS_DATA_START_ROW + i; break; } }
  if (row === -1) throw new Error('Sesi ' + idPm + ' tidak ditemukan.');

  var tz = Session.getScriptTimeZone();
  sheet.getRange(row, 11).setValue(statusAkhir);                                    // K
  sheet.getRange(row, 12).setValue(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss')); // L
  if (catatan) sheet.getRange(row, 13).setValue(catatan);                           // M
  SpreadsheetApp.flush();
  return { success: true, message: 'Status sesi ' + idPm + ' diperbarui: ' + statusAkhir + '.' };
}

function gasNum_(v) {
  return (v === undefined || v === null || v === '') ? '' : Number(v);
}

function gasSubmitIteration(d) {
  if (!d || !d.id_pm) throw new Error('id_pm wajib diisi.');
  var no = Number(d.no_iterasi || 0);
  if (!no || no < 1) throw new Error('Nomor iterasi tidak valid.');
  if (no > GAS_MAX_ITERASI) throw new Error('Iterasi maksimal ' + GAS_MAX_ITERASI + ' per sesi PM (SOP).');

  var sheet = getGasSheetNamed_(GAS_DETAIL_SHEET);
  var row = gasNextRow_(sheet);
  var id = gasNextSeqId_(sheet, 1, 'DET');
  var tz = Session.getScriptTimeZone();

  // A-G (identitas + pembacaan), lompati H-I (rumus), J-K (toleransi manual).
  sheet.getRange(row, 1, 1, 7).setValues([[
    id, d.id_pm, no, d.tahap || '-', d.waktu || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss'),
    gasNum_(d.baca_zero), gasNum_(d.baca_span)
  ]]);
  sheet.getRange(row, 10, 1, 2).setValues([[gasNum_(d.tol_zero), gasNum_(d.tol_span)]]);
  // O, Q - lompati L,M,N (rumus) dan P (rumus, peringatan).
  sheet.getRange(row, 15).setValue(d.tindakan || 'Tidak');
  sheet.getRange(row, 17).setValue(d.catatan || '');

  gasEnsureFormulas_(sheet, row, GAS_DETAIL_FORMULA_COLS, gasDetailFormulasFor_(row));
  SpreadsheetApp.flush();

  // Baca kembali hasil rumus (status, peringatan) supaya frontend bisa langsung
  // menuntun langkah SOP berikutnya tanpa request tambahan.
  var hasil = sheet.getRange(row, 14).getDisplayValue();     // N = Status
  var peringatan = sheet.getRange(row, 16).getDisplayValue(); // P
  return {
    success: true, id_detail: id, row: row, status: hasil, peringatan: peringatan,
    message: 'Iterasi ke-' + no + ' tersimpan (' + (hasil || 'menunggu pembacaan') + ').'
  };
}
