/**
 * INA BASIC STORE - Warehouse & Production System V6.6 (100% Made-to-Order / Gudang Polos)
 * Arsitektur Final:
 * - Keamanan & Validasi: Sanitasi anti-formula injection (sanitizeSheetText_), validasi bilangan bulat, & format Rupiah
 * - Penguncian Harga Jual: Harga jual & HPP pada Barang Keluar diambil 100% dari VARIASI_PRODUK
 * - Konsolidasi Saldo Gudang Polos: Otomatis menggabungkan & menolkan baris duplikat GUDANG POLOS pada STOK_TAILOR
 * - Stock Opname Gudang Polos: Perhitungan & penyesuaian fisik langsung terhadap stok GUDANG POLOS (STOK_TAILOR) dengan atomic rollback
 * - Standardisasi Nama Upah: Seluruh nama penjahit & pelapor/admin upah diseragamkan ke HURUF BESAR (UPPERCASE)
 */

var IBS = {
  SHEET_MASTER: 'MASTER_PRODUK',
  SHEET_ATTRIBUTES: 'ATRIBUT_PRODUK',
  SHEET_VARIANTS: 'VARIASI_PRODUK',
  SHEET_STOCK: 'STOK_VARIASI',
  SHEET_LOG: 'OPERATION_LOG',
  SHEET_LOCATIONS: 'MASTER_LOKASI',
  SHEET_OPNAME_SESI: 'OPNAME_SESI',
  SHEET_OPNAME_DETAIL: 'OPNAME_DETAIL',
  SHEET_MASTER_TAILOR: 'MASTER_TAILOR',
  SHEET_STOK_TAILOR: 'STOK_TAILOR',
  SHEET_LOG_TAILOR: 'LOG_TAILOR',
  SHEET_MAPPING_POLOS: 'MAPPING_POLOS_PRODUK',
  TAILOR_GUDANG_POLOS_NAME: 'GUDANG POLOS',
  DEPRECATED_SHEETS: [
    'MASTER_BAHAN',
    'PENGIRIMAN_BAHAN',
    'STOK_BAHAN',
    'LOG_BAHAN',
    'STOK_LOKASI',
    'MIGRASI_ALOKASI',
    'WO_JAHIT',
    'WO_JAHIT_ITEM',
    'LAPORAN_JAHIT'
  ],
  LOCK_TIMEOUT_MS: 15000,
  CACHE_TTL_SEC: 90,
  HEADERS: {
    MASTER_PRODUK: ['Barcode', 'Nama', 'Status_Aktif', 'Created_At', 'Updated_At'],
    ATRIBUT_PRODUK: ['Barcode_Induk', 'Nama_Atribut', 'Tipe_Atribut', 'Nilai', 'Kode_Warna', 'Urutan_Atribut', 'Urutan_Nilai', 'Status_Aktif'],
    VARIASI_PRODUK: ['ID_Variasi', 'Barcode_Induk', 'Label_Variasi', 'Atribut_JSON', 'Status_Aktif', 'Created_At', 'Updated_At', 'HPP', 'Harga_Jual', 'Margin_Rp', 'Margin_Persen'],
    STOK_VARIASI: ['ID_Variasi', 'Barcode_Induk', 'Lokasi_Rak', 'Kuantitas', 'Updated_At'],
    OPERATION_LOG: ['Timestamp', 'ID_Transaksi', 'Aktivitas', 'Barcode', 'Lokasi_Asal', 'Lokasi_Tujuan', 'Qty', 'Status', 'ID_Variasi', 'Nama_Produk', 'Variasi_JSON', 'Label_Variasi', 'Produk_Tampilan', 'Nama_Pelapor', 'HPP_Satuan', 'Harga_Jual_Satuan', 'Total_Omzet', 'Total_Profit'],
    MASTER_LOKASI: ['Kode_Lokasi', 'Nama_Lokasi', 'Jenis_Lokasi', 'Status_Aktif', 'Created_At'],
    OPNAME_SESI: ['Session_ID', 'Nama_Sesi', 'Lokasi', 'Status', 'Dibuat_Oleh', 'Waktu_Mulai', 'Waktu_Selesai'],
    OPNAME_DETAIL: ['Detail_ID', 'Session_ID', 'Barcode', 'ID_Variasi', 'Label_Variasi', 'Qty_Sistem_Snapshot', 'Qty_Fisik', 'Nama_Pelapor', 'Waktu', 'Status_Hitung'],
    MASTER_TAILOR: ['ID_Tailor', 'Barcode_Kategori', 'Nama_Kategori', 'Warna', 'Ukuran', 'Upah_Per_Pcs', 'Barcode_Produk_Tujuan', 'ID_Variasi_Tujuan', 'Status_Aktif'],
    STOK_TAILOR: ['ID_Tailor', 'Nama_Penjahit', 'Barcode_Kategori', 'Nama_Kategori', 'Warna', 'Ukuran', 'Stok_Ready', 'Updated_At'],
    LOG_TAILOR: ['Log_ID', 'Waktu', 'Jenis_Aktivitas', 'Nama_Penjahit', 'ID_Tailor', 'Detail_Item', 'Qty_Bagus', 'Qty_Reject', 'Upah_Satuan', 'Total_Upah', 'Status_Upah', 'ID_Slip'],
    MAPPING_POLOS_PRODUK: ['Mapping_ID', 'Barcode_Produk', 'ID_Variasi', 'Label_Variasi', 'ID_Tailor', 'Detail_Polos', 'Qty_Per_Unit', 'Updated_At']
  }
};

// ==========================================
// HELPER KEAMANAN, SANITASI & VALIDASI (V6.6)
// ==========================================
function sanitizeSheetText_(val) {
  var str = textV2_(val);
  if (!str) return '';
  if (/^[=+\-@]/.test(str)) {
    return "'" + str;
  }
  return str;
}

function assertAuthorizedRequest_(tokenFromClient) {
  var expectedToken = PropertiesService.getScriptProperties().getProperty('IBS_API_TOKEN');
  if (!expectedToken) return;
  if (textV2_(tokenFromClient) !== textV2_(expectedToken)) {
    throw new Error('Akses ditolak: Token API tidak valid.');
  }
}

function validateOpnameCountInput_(variant, barcode, qty) {
  if (!isFinite(qty) || qty < 0 || Math.floor(qty) !== qty) {
    throw new Error('Hasil hitung fisik harus berupa bilangan bulat >= 0.');
  }
  if (!variant || upperV2_(variant.barcode) !== upperV2_(barcode)) {
    throw new Error('Variasi tidak cocok dengan barcode produk.');
  }
  if (variant.active === false) {
    throw new Error('Variasi sedang nonaktif.');
  }
}

function isGudangPolosName_(namaPenjahit) {
  return upperV2_(namaPenjahit) === upperV2_(IBS.TAILOR_GUDANG_POLOS_NAME);
}

// ==========================================
// HELPER NORMALISASI WARNA & UKURAN (AUTO-MATCH)
// ==========================================
function normalizeMatchToken_(val) {
  var str = upperV2_(val).replace(/[_-]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!str || str === '-' || str === 'STANDAR' || str === 'DEFAULT') return '-';
  if (str === 'ALL SIZE' || str === 'ALLSIZE' || str === 'FREE SIZE' || str === 'FREESIZE' || str === 'ONE SIZE' || str === 'ONESIZE') {
    return 'ALLSIZE';
  }
  return str;
}

function extractColorSizeFromVariant_(attrsObj, labelVariasi) {
  var warna = '';
  var ukuran = '';
  var attrs = (attrsObj && typeof attrsObj === 'object') ? attrsObj : {};
  var keys = Object.keys(attrs);

  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    var kLow = textV2_(k).toLowerCase();
    var val = textV2_(attrs[k]);
    if (!val) continue;
    if (!warna && (kLow.indexOf('warna') !== -1 || kLow.indexOf('color') !== -1 || kLow.indexOf('colour') !== -1)) {
      warna = val;
    } else if (!ukuran && (kLow.indexOf('ukuran') !== -1 || kLow.indexOf('size') !== -1)) {
      ukuran = val;
    }
  }

  // Fallback jika nama atribut berbeda atau hanya label yang tersedia
  if (!warna && !ukuran && keys.length > 0) {
    if (keys.length >= 1) warna = textV2_(attrs[keys[0]]);
    if (keys.length >= 2) ukuran = textV2_(attrs[keys[1]]);
  } else if (!warna && !ukuran && labelVariasi) {
    var lbl = textV2_(labelVariasi);
    if (lbl && upperV2_(lbl) !== 'STANDAR') {
      var parts = lbl.split(/\s*(?:·|\/|\|)\s*/).map(function (p) { return textV2_(p); }).filter(Boolean);
      if (parts.length >= 1) warna = parts[0];
      if (parts.length >= 2) ukuran = parts[1];
    }
  }

  return {
    warna: warna || '-',
    ukuran: ukuran || '-'
  };
}

// ==========================================
// HELPER CACHE SERVICE & SINKRONISASI SHEET
// ==========================================
function getCachedV5_(key, builderFn) {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(key);
  if (hit) {
    try {
      return JSON.parse(hit);
    } catch (e) {}
  }
  var fresh = builderFn();
  try {
    var str = JSON.stringify(fresh);
    if (str.length < 95000) {
      cache.put(key, str, IBS.CACHE_TTL_SEC);
    }
  } catch (e) {}
  return fresh;
}

function invalidateAppCache_() {
  try {
    CacheService.getScriptCache().removeAll([
      'IBS_BOOTSTRAP_V6_6',
      'IBS_MASTER_TAILOR_V6_6',
      'IBS_STOK_TAILOR_V6_6',
      'IBS_GUDANG_POLOS_V6_6',
      'IBS_POLOS_MAPPING_V6_6',
      'IBS_UNPAID_UPAH_V6_6',
      'IBS_OPNAME_SESSIONS_V6_6'
    ]);
  } catch (e) {}
}

/**
 * Membersihkan cache ketika sheet diedit langsung.
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    var sheetName = sheet.getName();
    var watchedSheets = [
      IBS.SHEET_MASTER,
      IBS.SHEET_ATTRIBUTES,
      IBS.SHEET_VARIANTS,
      IBS.SHEET_STOCK,
      IBS.SHEET_MASTER_TAILOR,
      IBS.SHEET_STOK_TAILOR,
      IBS.SHEET_LOG_TAILOR,
      IBS.SHEET_OPNAME_SESI,
      IBS.SHEET_MAPPING_POLOS
    ];

    if (watchedSheets.indexOf(sheetName) !== -1) {
      invalidateAppCache_();
      if (typeof markDashboardDirty_ === 'function') {
        markDashboardDirty_();
      }
    }
  } catch (err) {
    console.warn('onEdit sync:', err);
  }
}

function assertEditableStatus_(currentStatus, labelEntity) {
  var st = upperV2_(currentStatus);
  var locked = ['SELESAI', 'APPROVED', 'REJECTED', 'COMPLETED', 'BATAL', 'DIBATALKAN', 'CLOSED', 'LUNAS'];
  if (locked.indexOf(st) !== -1) {
    throw new Error(labelEntity + ' sudah berstatus ' + st + ' dan terkunci permanen (tidak dapat diubah lagi).');
  }
}

/**
 * FUNGSI MIGRASI AMAN KE SPREADSHEET PRODUKSI (BUKAN TESTING)
 */
function migrasiKeProduksi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureAllOperationalSheets_(ss);
  var cleanedTestCount = cleanupSimulationData_(ss);
  var res = setup();

  if (typeof installDashboardAutoTrigger_ === 'function') {
    installDashboardAutoTrigger_(false);
  }

  if (typeof buildDashboardSummary === 'function') {
    buildDashboardSummary();
  }

  invalidateAppCache_();

  var summary = 'Migrasi Produksi V6.6 (Keamanan, Harga Otomatis, Opname Gudang Polos & Auto-Trigger) Berhasil! ' +
    'Data MASTER_PRODUK & STOK_TAILOR aman dipertahankan.' +
    (cleanedTestCount > 0 ? ' Baris simulasi testing dibersihkan: ' + cleanedTestCount + '.' : '');

  Logger.log(summary);
  try {
    ss.toast(summary, 'Migrasi Produksi Selesai', 8);
  } catch (e) {}

  return {
    success: true,
    message: summary,
    cleaned_test_rows: cleanedTestCount,
    deleted_deprecated_sheets: res.deleted_sheets || []
  };
}

function cleanupSimulationData_(ss) {
  var testSku = 'TEST-SKU-V6';
  var testProdBarcode = 'TEST-PROD-V6';
  var testTailorName = 'PENJAHIT UJI V6';
  var deletedCount = 0;

  var targets = [
    { sheet: IBS.SHEET_MASTER_TAILOR, cols: [2] },
    { sheet: IBS.SHEET_STOK_TAILOR, cols: [2, 3] },
    { sheet: IBS.SHEET_MAPPING_POLOS, cols: [2] },
    { sheet: IBS.SHEET_LOG_TAILOR, cols: [4] },
    { sheet: IBS.SHEET_MASTER, cols: [1] },
    { sheet: IBS.SHEET_ATTRIBUTES, cols: [1] },
    { sheet: IBS.SHEET_VARIANTS, cols: [2] },
    { sheet: IBS.SHEET_STOCK, cols: [2] },
    { sheet: IBS.SHEET_LOG, cols: [4] }
  ];

  targets.forEach(function (t) {
    var sh = ss.getSheetByName(t.sheet);
    if (!sh || sh.getLastRow() < 2) return;
    var vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();
    for (var r = vals.length - 1; r >= 0; r--) {
      var match = t.cols.some(function (cIdx) {
        var v = upperV2_(vals[r][cIdx - 1]);
        return v === testSku || v === testProdBarcode || v === testTailorName;
      });
      if (match) {
        sh.deleteRow(r + 2);
        deletedCount++;
      }
    }
  });

  return deletedCount;
}

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAllOperationalSheets_(ss);

  var deletedSheets = cleanupUnusedSheets_(ss);

  var sheets = [
    IBS.SHEET_MASTER, IBS.SHEET_ATTRIBUTES, IBS.SHEET_VARIANTS, IBS.SHEET_STOCK,
    IBS.SHEET_LOG, IBS.SHEET_LOCATIONS, IBS.SHEET_OPNAME_SESI,
    IBS.SHEET_OPNAME_DETAIL,
    IBS.SHEET_MASTER_TAILOR, IBS.SHEET_STOK_TAILOR, IBS.SHEET_LOG_TAILOR,
    IBS.SHEET_MAPPING_POLOS
  ];

  sheets.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh) {
      sh.setFrozenRows(1);
      var width = Math.max(sh.getLastColumn(), 1);
      sh.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#d9eaf7');
    }
  });

  ss.getSheetByName(IBS.SHEET_MASTER).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_ATTRIBUTES).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_VARIANTS).getRange('A:B').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_STOCK).getRange('A:B').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_LOG).getRange('B:B').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_LOG).getRange('D:F').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_LOG).getRange('I:I').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_LOCATIONS).getRange('A:B').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_OPNAME_DETAIL).getRange('B:D').setNumberFormat('@');

  var shMasterTailor = ss.getSheetByName(IBS.SHEET_MASTER_TAILOR);
  shMasterTailor.getRange('A:B').setNumberFormat('@');
  shMasterTailor.getRange('G:H').setNumberFormat('@');
  shMasterTailor.getRange('F2:F').setNumberFormat('#,##0');

  var shStokTailor = ss.getSheetByName(IBS.SHEET_STOK_TAILOR);
  shStokTailor.getRange('A:A').setNumberFormat('@');
  shStokTailor.getRange('C:C').setNumberFormat('@');
  shStokTailor.getRange('G2:G').setNumberFormat('#,##0');

  var shLogTailor = ss.getSheetByName(IBS.SHEET_LOG_TAILOR);
  shLogTailor.getRange('A:A').setNumberFormat('@');
  shLogTailor.getRange('E:E').setNumberFormat('@');
  shLogTailor.getRange('L:L').setNumberFormat('@');
  shLogTailor.getRange('G2:J').setNumberFormat('#,##0');

  var shMappingPolos = ss.getSheetByName(IBS.SHEET_MAPPING_POLOS);
  shMappingPolos.getRange('A:C').setNumberFormat('@');
  shMappingPolos.getRange('E:E').setNumberFormat('@');
  shMappingPolos.getRange('G2:G').setNumberFormat('#,##0');

  var shVar = ss.getSheetByName(IBS.SHEET_VARIANTS);
  repairVariasiProdukSheet_(ss);
  shVar.getRange('F2:G').setNumberFormat('dd/mm/yyyy HH:mm');
  shVar.getRange('H2:J').setNumberFormat('"Rp "#,##0');
  shVar.getRange('K2:K').setNumberFormat('0.0%');
  syncVariantMarginFormulas_(shVar);

  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  shLog.getRange('O2:R').setNumberFormat('"Rp "#,##0');

  var locSheet = ss.getSheetByName(IBS.SHEET_LOCATIONS);

  cleanAndApplyCheckboxV2_(ss.getSheetByName(IBS.SHEET_MASTER), 3);
  cleanAndApplyCheckboxV2_(ss.getSheetByName(IBS.SHEET_ATTRIBUTES), 8);
  cleanAndApplyCheckboxV2_(shVar, 5);
  cleanAndApplyCheckboxV2_(locSheet, 4);
  cleanAndApplyCheckboxV2_(shMasterTailor, 9);

  if (typeof installDashboardAutoTrigger_ === 'function') {
    installDashboardAutoTrigger_(false);
  }

  if (typeof buildDashboardSummary === 'function') {
    buildDashboardSummary();
  }

  invalidateAppCache_();

  var msg = 'Setup V6.6 (Made-to-Order, Keamanan Input, Opname Gudang Polos) selesai.';
  if (deletedSheets.length > 0) {
    msg += ' Tab lama dihapus: ' + deletedSheets.join(', ');
  }
  Logger.log(msg);
  try {
    ss.toast(msg, 'Ina Basic Store V6.6', 6);
  } catch (e) {}

  return { success: true, message: msg, deleted_sheets: deletedSheets };
}

function hapusTabTidakTerpakai() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var deleted = cleanupUnusedSheets_(ss);
  var msg = deleted.length
    ? 'Berhasil menghapus tab tidak terpakai: ' + deleted.join(', ')
    : 'Tidak ada tab lama yang perlu dihapus.';
  Logger.log(msg);
  try {
    ss.toast(msg, 'Pembersihan Tab', 6);
  } catch (e) {}
  return deleted;
}

function cleanupUnusedSheets_(ss) {
  var deleted = [];
  IBS.DEPRECATED_SHEETS.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh && ss.getSheets().length > 1) {
      ss.deleteSheet(sh);
      deleted.push(name);
    }
  });

  var defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && defaultSheet.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
    deleted.push('Sheet1');
  }

  return deleted;
}

function ensureAllOperationalSheets_(ss) {
  ensureSheetHeadersV2_(ss, IBS.SHEET_MASTER, IBS.HEADERS.MASTER_PRODUK);
  ensureSheetHeadersV2_(ss, IBS.SHEET_ATTRIBUTES, IBS.HEADERS.ATRIBUT_PRODUK);
  ensureSheetHeadersV2_(ss, IBS.SHEET_VARIANTS, IBS.HEADERS.VARIASI_PRODUK);
  repairVariasiProdukSheet_(ss);
  ensureSheetHeadersV2_(ss, IBS.SHEET_STOCK, IBS.HEADERS.STOK_VARIASI);
  ensureSheetHeadersV2_(ss, IBS.SHEET_LOG, IBS.HEADERS.OPERATION_LOG);

  var locSheet = ensureSheetHeadersV2_(ss, IBS.SHEET_LOCATIONS, IBS.HEADERS.MASTER_LOKASI);
  if (locSheet.getLastRow() === 1) {
    locSheet.appendRow(['RAK-UTAMA', 'Rak Utama', 'warehouse', true, new Date()]);
  }

  ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_SESI, IBS.HEADERS.OPNAME_SESI);
  ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_DETAIL, IBS.HEADERS.OPNAME_DETAIL);
  ensureSheetHeadersV2_(ss, IBS.SHEET_MASTER_TAILOR, IBS.HEADERS.MASTER_TAILOR);
  ensureSheetHeadersV2_(ss, IBS.SHEET_STOK_TAILOR, IBS.HEADERS.STOK_TAILOR);
  ensureSheetHeadersV2_(ss, IBS.SHEET_LOG_TAILOR, IBS.HEADERS.LOG_TAILOR);
  ensureSheetHeadersV2_(ss, IBS.SHEET_MAPPING_POLOS, IBS.HEADERS.MAPPING_POLOS_PRODUK);
}

function syncVariantMarginFormulas_(shVar) {
  var lastRow = shVar.getLastRow();
  if (lastRow < 2) return;
  var count = lastRow - 1;
  var formulasJ = [];
  var formulasK = [];
  for (var i = 0; i < count; i++) {
    var r = i + 2;
    formulasJ.push(['=IF(I' + r + '>0, I' + r + '-H' + r + ', 0)']);
    formulasK.push(['=IF(I' + r + '>0, (I' + r + '-H' + r + ')/I' + r + ', 0)']);
  }
  shVar.getRange(2, 10, count, 1).setFormulas(formulasJ);
  shVar.getRange(2, 11, count, 1).setFormulas(formulasK);
}

function cleanAndApplyCheckboxV2_(sheet, checkboxCol) {
  if (!sheet) return;
  var maxRows = sheet.getMaxRows();
  var values = sheet.getRange(1, 1, maxRows, Math.max(sheet.getLastColumn(), 2)).getDisplayValues();
  var lastRealRow = 1;
  for (var i = values.length - 1; i >= 1; i--) {
    if (textV2_(values[i][0]) !== '' || textV2_(values[i][1]) !== '') {
      lastRealRow = i + 1;
      break;
    }
  }
  if (maxRows > lastRealRow) {
    var emptyRange = sheet.getRange(lastRealRow + 1, 1, maxRows - lastRealRow, Math.max(sheet.getLastColumn(), 1));
    emptyRange.clearDataValidations();
    emptyRange.clearContent();
  }
  if (lastRealRow >= 2) {
    var dataRange = sheet.getRange(2, checkboxCol, lastRealRow - 1, 1);
    dataRange.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build());
  }
}

function applyRowCheckboxV2_(sheet, startRow, numRows, checkboxCol) {
  if (!sheet || numRows < 1) return;
  var range = sheet.getRange(startRow, checkboxCol, numRows, 1);
  range.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build());
}

function ensureSheetHeadersV2_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0 || sh.getLastColumn() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sh;
  }
  var currentWidth = Math.max(sh.getLastColumn(), 1);
  var current = sh.getRange(1, 1, 1, currentWidth).getDisplayValues()[0].map(function (v) { return String(v).trim(); });
  headers.forEach(function (header) {
    if (current.indexOf(header) === -1) {
      current.push(header);
      sh.getRange(1, current.length).setValue(header);
    }
  });
  return sh;
}

function repairVariasiProdukSheet_(ss) {
  var shVar = ss.getSheetByName(IBS.SHEET_VARIANTS);
  if (!shVar) return;
  var expected = IBS.HEADERS.VARIASI_PRODUK;
  var lastCol = Math.max(shVar.getLastColumn(), expected.length);
  var headerRow = shVar.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (v) { return textV2_(v); });

  var lastRow = shVar.getLastRow();
  if (lastRow >= 2 && lastCol > expected.length) {
    var strayHppIdx = headerRow.indexOf('HPP', expected.length);
    var strayHjIdx = headerRow.indexOf('Harga_Jual', expected.length);
    if (strayHppIdx !== -1 || strayHjIdx !== -1) {
      var allVals = shVar.getRange(2, 1, lastRow - 1, lastCol).getValues();
      var fixedHJ = [];
      for (var i = 0; i < allVals.length; i++) {
        var curHpp = Number(allVals[i][7]) || 0;
        var curHj = Number(allVals[i][8]) || 0;
        if (curHpp === 0 && strayHppIdx !== -1) curHpp = Number(allVals[i][strayHppIdx]) || 0;
        if (curHj === 0 && strayHjIdx !== -1) curHj = Number(allVals[i][strayHjIdx]) || 0;
        fixedHJ.push([curHpp, curHj]);
      }
      shVar.getRange(2, 8, fixedHJ.length, 2).setValues(fixedHJ);
    }
  }

  shVar.getRange(1, 1, 1, expected.length).setValues([expected]).setFontWeight('bold').setBackground('#d9eaf7');

  if (shVar.getLastColumn() > expected.length) {
    shVar.deleteColumns(expected.length + 1, shVar.getLastColumn() - expected.length);
  }
}

function assertV2SheetsReady_(ss) {
  var required = [
    IBS.SHEET_MASTER, IBS.SHEET_ATTRIBUTES, IBS.SHEET_VARIANTS, IBS.SHEET_STOCK,
    IBS.SHEET_LOG, IBS.SHEET_LOCATIONS, IBS.SHEET_OPNAME_SESI,
    IBS.SHEET_OPNAME_DETAIL,
    IBS.SHEET_MASTER_TAILOR, IBS.SHEET_STOK_TAILOR, IBS.SHEET_LOG_TAILOR,
    IBS.SHEET_MAPPING_POLOS
  ];
  var missing = required.some(function (name) { return !ss.getSheetByName(name); });
  if (missing) {
    ensureAllOperationalSheets_(ss);
  }
}

function getActiveLocationsV2_(ss) {
  var sh = ss.getSheetByName(IBS.SHEET_LOCATIONS);
  if (!sh || sh.getLastRow() < 2) {
    return [{ kode: 'RAK-UTAMA', value: 'Rak Utama', label: 'Rak Utama', kind: 'warehouse' }];
  }
  var rows = readObjectsV2_(sh);
  var active = [];
  var seen = {};
  rows.forEach(function (r) {
    var nama = textV2_(r.Nama_Lokasi || r['Nama Lokasi'] || r.Lokasi);
    if (!nama) return;
    var rawStatus = r.Status_Aktif !== undefined ? r.Status_Aktif : (r['Status Aktif'] !== undefined ? r['Status Aktif'] : r.Aktif);
    var isAct = (rawStatus === '' || rawStatus === null || rawStatus === undefined)
      ? true
      : booleanV2_(rawStatus, true);

    if (isAct && !seen[upperV2_(nama)]) {
      seen[upperV2_(nama)] = true;
      active.push({
        kode: textV2_(r.Kode_Lokasi || r['Kode Lokasi'] || upperV2_(nama).replace(/[^A-Z0-9]+/g, '-')),
        value: nama,
        label: nama,
        kind: textV2_(r.Jenis_Lokasi || r['Jenis Lokasi']) || 'warehouse'
      });
    }
  });
  return active.length ? active : [{ kode: 'RAK-UTAMA', value: 'Rak Utama', label: 'Rak Utama', kind: 'warehouse' }];
}

function isValidLocationV2_(ss, value) {
  var locs = getActiveLocationsV2_(ss);
  return locs.some(function (l) { return upperV2_(l.value) === upperV2_(value); });
}

function getBootstrapBundle_(ss) {
  assertV2SheetsReady_(ss);
  var cachedPart = getCachedV5_('IBS_BOOTSTRAP_V6_6', function () {
    var unpaid = getUnpaidUpah_(ss);
    return {
      master_tailor: getMasterTailorList_(ss, false).items,
      stok_tailor: getStokTailorList_(ss, false).items,
      gudang_polos: getGudangPolosList_(ss).items,
      polos_mappings: getPolosMappingsList_(ss).items,
      opname_sessions: getOpnameSessions_(ss).data.sessions,
      unpaid_summary: {
        grand_total_upah: unpaid.grand_total_upah || 0,
        total_penjahit: unpaid.total_penjahit || 0
      }
    };
  });
  return {
    locations: getActiveLocationsV2_(ss),
    master_tailor: cachedPart.master_tailor,
    stok_tailor: cachedPart.stok_tailor,
    gudang_polos: cachedPart.gudang_polos || [],
    polos_mappings: cachedPart.polos_mappings || [],
    opname_sessions: cachedPart.opname_sessions,
    unpaid_summary: cachedPart.unpaid_summary || { grand_total_upah: 0, total_penjahit: 0 },
    server_time: new Date().toISOString()
  };
}

function doGet(e) {
  var action = e && e.parameter && e.parameter.action ? upperV2_(e.parameter.action) : '';
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    assertAuthorizedRequest_(e && e.parameter ? e.parameter.token : '');
  } catch (authErr) {
    return jsonResponseV2_(false, authErr.message);
  }

  if (action === 'GET_BOOTSTRAP') {
    return jsonResponseV2_(true, 'Bootstrap data berhasil dimuat.', getBootstrapBundle_(ss));
  }
  if (action === 'GET_LOCATIONS') {
    assertV2SheetsReady_(ss);
    return jsonResponseV2_(true, 'Daftar lokasi aktif.', { locations: getActiveLocationsV2_(ss) });
  }
  if (action === 'GET_MASTER_TAILOR') {
    assertV2SheetsReady_(ss);
    var includeInactive = e.parameter && upperV2_(e.parameter.include_inactive) === 'TRUE';
    return jsonResponseV2_(true, 'Daftar kategori & tarif penjahit.', getMasterTailorList_(ss, includeInactive));
  }
  if (action === 'GET_STOK_TAILOR') {
    assertV2SheetsReady_(ss);
    var includeZero = e.parameter && upperV2_(e.parameter.include_zero) === 'TRUE';
    return jsonResponseV2_(true, 'Daftar stok di gudang penjahit.', getStokTailorList_(ss, includeZero));
  }
  if (action === 'GET_GUDANG_POLOS') {
    assertV2SheetsReady_(ss);
    return jsonResponseV2_(true, 'Daftar stok gudang polos & aturan mapping.', {
      gudang_polos: getGudangPolosList_(ss).items,
      polos_mappings: getPolosMappingsList_(ss).items
    });
  }
  if (action === 'GET_TAILOR_SKU') {
    assertV2SheetsReady_(ss);
    var skuParam = e.parameter ? (e.parameter.sku || e.parameter.barcode || e.parameter.barcode_kategori) : '';
    var skuRes = getTailorSkuAction_(ss, { sku: skuParam });
    return jsonResponseV2_(skuRes.success, skuRes.message, skuRes.data);
  }
  if (action === 'GET_UNPAID_UPAH') {
    assertV2SheetsReady_(ss);
    return jsonResponseV2_(true, 'Rekap upah penjahit belum dibayar.', getUnpaidUpah_(ss));
  }
  return jsonResponseV2_(true, 'Ina Basic Store Warehouse API V6.6 (Made-to-Order) aktif.', {
    version: '6.6',
    supported_actions: [
      'GET_BOOTSTRAP', 'GET_PRODUCT', 'GET_LOCATIONS', 'GET_OPNAME_SESSIONS',
      'GET_MASTER_TAILOR', 'GET_STOK_TAILOR', 'GET_GUDANG_POLOS', 'GET_TAILOR_SKU', 'GET_UNPAID_UPAH',
      'CREATE_PRODUCT', 'ADD_VARIANTS', 'UPDATE_PRODUCT', 'SET_VARIANT_STATUS',
      'UPDATE_VARIANT_PRICING', 'SET_VARIANT_PRICING', 'SET_VARIANT_PRICING_BULK',
      'SAVE_TRANSACTION', 'SAVE_BULK_KELUAR', 'SAVE_POLOS_MAPPING', 'AUTO_MAP_POLOS_CATEGORY',
      'DEACTIVATE_PRODUCT', 'CREATE_OPNAME_SESSION', 'SUBMIT_OPNAME_COUNT', 'APPROVE_OPNAME_SESSION',
      'SAVE_MASTER_TAILOR', 'SUBMIT_STOK_TAILOR', 'TRANSFER_TAILOR_TO_WAREHOUSE', 'GENERATE_PAY_SLIP'
    ]
  });
}

function doPost(e) {
  var payload;
  try {
    payload = JSON.parse(e && e.postData && e.postData.contents ? e.postData.contents : '{}');
  } catch (parseError) {
    return jsonResponseV2_(false, 'Format permintaan tidak valid.');
  }

  try {
    assertAuthorizedRequest_(payload.token);
  } catch (authErr) {
    return jsonResponseV2_(false, authErr.message);
  }

  var action = textV2_(payload.action).toUpperCase();
  if (!action && payload.activity) {
    var legacy = textV2_(payload.activity).toLowerCase();
    if (legacy === 'cekproduk') action = 'GET_PRODUCT';
    else if (legacy === 'editproduk') action = 'UPDATE_PRODUCT';
    else if (legacy === 'hapusproduk') action = 'DEACTIVATE_PRODUCT';
    else return jsonResponseV2_(false, 'Aplikasi lama tidak kompatibel. Perbarui aplikasi ke versi terbaru.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // READ-ONLY ACTIONS
  if (action === 'GET_BOOTSTRAP') {
    try {
      return jsonResponseV2_(true, 'Bootstrap data berhasil dimuat.', getBootstrapBundle_(ss));
    } catch (bootErr) {
      return jsonResponseV2_(false, 'Kesalahan server: ' + (bootErr.message || String(bootErr)));
    }
  }

  if (action === 'GET_PRODUCT') {
    try {
      assertV2SheetsReady_(ss);
      var readResult = getProductActionV2_(ss, payload);
      return jsonResponseV2_(readResult.success, readResult.message, readResult.data);
    } catch (readError) {
      return jsonResponseV2_(false, 'Kesalahan server: ' + (readError.message || String(readError)));
    }
  }

  if (action === 'GET_LOCATIONS') {
    try {
      assertV2SheetsReady_(ss);
      return jsonResponseV2_(true, 'Daftar lokasi aktif.', { locations: getActiveLocationsV2_(ss) });
    } catch (locError) {
      return jsonResponseV2_(false, 'Kesalahan server: ' + (locError.message || String(locError)));
    }
  }

  if (action === 'GET_OPNAME_SESSIONS') {
    try {
      assertV2SheetsReady_(ss);
      var opnameRes = getOpnameSessions_(ss);
      return jsonResponseV2_(opnameRes.success, opnameRes.message || 'Daftar sesi opname.', opnameRes.data);
    } catch (err) {
      return jsonResponseV2_(false, 'Kesalahan server: ' + err.message);
    }
  }

  if (action === 'GET_MASTER_TAILOR') {
    try {
      assertV2SheetsReady_(ss);
      return jsonResponseV2_(true, 'Daftar kategori & tarif penjahit.', getMasterTailorList_(ss, payload.include_inactive === true));
    } catch (err) {
      return jsonResponseV2_(false, 'Kesalahan server: ' + err.message);
    }
  }

  if (action === 'GET_STOK_TAILOR') {
    try {
      assertV2SheetsReady_(ss);
      return jsonResponseV2_(true, 'Daftar stok di gudang penjahit.', getStokTailorList_(ss, payload.include_zero === true));
    } catch (err) {
      return jsonResponseV2_(false, 'Kesalahan server: ' + err.message);
    }
  }

  if (action === 'GET_GUDANG_POLOS') {
    try {
      assertV2SheetsReady_(ss);
      return jsonResponseV2_(true, 'Daftar stok gudang polos & aturan mapping.', {
        gudang_polos: getGudangPolosList_(ss).items,
        polos_mappings: getPolosMappingsList_(ss).items
      });
    } catch (err) {
      return jsonResponseV2_(false, 'Kesalahan server: ' + err.message);
    }
  }

  if (action === 'GET_TAILOR_SKU') {
    try {
      assertV2SheetsReady_(ss);
      var skuResPost = getTailorSkuAction_(ss, payload);
      return jsonResponseV2_(skuResPost.success, skuResPost.message, skuResPost.data);
    } catch (err) {
      return jsonResponseV2_(false, 'Kesalahan server: ' + err.message);
    }
  }

  if (action === 'GET_UNPAID_UPAH') {
    try {
      assertV2SheetsReady_(ss);
      return jsonResponseV2_(true, 'Rekap upah penjahit belum dibayar.', getUnpaidUpah_(ss));
    } catch (err) {
      return jsonResponseV2_(false, 'Kesalahan server: ' + err.message);
    }
  }

  // WRITE / MUTATION ACTIONS
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(IBS.LOCK_TIMEOUT_MS)) return jsonResponseV2_(false, 'Server sedang memproses transaksi lain. Coba lagi.');
  try {
    assertV2SheetsReady_(ss);
    var result;
    switch (action) {
      case 'CREATE_PRODUCT':
        result = createProductActionV2_(ss, payload);
        break;
      case 'ADD_VARIANTS':
        result = addVariantsActionV2_(ss, payload);
        break;
      case 'UPDATE_PRODUCT':
        result = updateProductActionV2_(ss, payload);
        break;
      case 'SET_VARIANT_STATUS':
        result = setVariantStatusActionV2_(ss, payload);
        break;
      case 'UPDATE_VARIANT_PRICING':
        result = updateVariantPricingActionV5_(ss, payload);
        break;
      case 'SET_VARIANT_PRICING':
        result = updateVariantPricingActionV5_(ss, {
          barcode: payload.barcode,
          items: [{ variant_id: payload.variant_id, hpp: payload.hpp, harga_jual: payload.harga_jual }]
        });
        break;
      case 'SET_VARIANT_PRICING_BULK':
        result = setVariantPricingBulkV5_(ss, payload);
        break;
      case 'SAVE_TRANSACTION':
        result = saveStockTransactionV2_(ss, payload);
        break;
      case 'SAVE_BULK_KELUAR':
        result = saveBulkKeluarV2_(ss, payload);
        break;
      case 'SAVE_POLOS_MAPPING':
        result = savePolosMappingV2_(ss, payload);
        break;
      case 'AUTO_MAP_POLOS_CATEGORY':
        result = autoMapPolosCategoryV2_(ss, payload);
        break;
      case 'DEACTIVATE_PRODUCT':
        result = updateProductActionV2_(ss, { barcode: payload.barcode, status_aktif: false, transaction_id: payload.transaction_id });
        break;
      case 'CREATE_OPNAME_SESSION':
        result = createOpnameSession_(ss, payload);
        break;
      case 'SUBMIT_OPNAME_COUNT':
        result = submitOpnameCount_(ss, payload);
        break;
      case 'APPROVE_OPNAME_SESSION':
        result = approveOpnameSession_(ss, payload);
        break;
      case 'SAVE_MASTER_TAILOR':
        result = saveMasterTailor_(ss, payload);
        break;
      case 'SUBMIT_STOK_TAILOR':
        result = submitStokTailor_(ss, payload);
        break;
      case 'TRANSFER_TAILOR_TO_WAREHOUSE':
        result = transferTailorToWarehouse_(ss, payload);
        break;
      case 'GENERATE_PAY_SLIP':
        result = generatePaySlip_(ss, payload);
        break;
      default:
        result = { success: false, message: 'Aksi tidak dikenal: ' + action };
    }
    if (result && result.success) {
      invalidateAppCache_();
      if (typeof markDashboardDirty_ === 'function') {
        markDashboardDirty_();
      }
    }
    return jsonResponseV2_(result.success, result.message, result.data);
  } catch (err) {
    console.error(err.stack || err);
    return jsonResponseV2_(false, 'Kesalahan server: ' + (err.message || String(err)));
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// MODUL INTI: PRODUK, VARIASI, HPP, HARGA JUAL & TRANSAKSI
// ==========================================
function getProductActionV2_(ss, payload) {
  var barcode = textV2_(payload.barcode);
  if (!barcode) return { success: false, message: 'Barcode wajib diisi.' };
  var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode);
  if (!product) return { success: false, message: 'Produk tidak ditemukan.' };
  var includeInactive = payload.include_inactive === true;
  var data = composeProductDataV2_(ss, product, includeInactive);
  if (!includeInactive && !product.active) return { success: false, message: 'Produk sedang nonaktif.' };
  if (!data.variants.length && !includeInactive) return { success: false, message: 'Produk belum memiliki variasi. Buka menu Kelola Data untuk menambahkan variasi.' };
  return { success: true, message: 'Data produk ditemukan.', data: data };
}

function composeProductDataV2_(ss, product, includeInactive) {
  var attrRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_ATTRIBUTES));
  var variantRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_VARIANTS));
  var stockRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_STOCK));
  var allMappings = getPolosMappingsList_(ss).items;
  var productMappings = allMappings.filter(function (m) {
    return upperV2_(m.barcode_produk) === upperV2_(product.barcode);
  });

  var definitions = {};
  attrRows.forEach(function (r) {
    if (upperV2_(r.Barcode_Induk) !== upperV2_(product.barcode)) return;
    var name = textV2_(r.Nama_Atribut);
    if (!name) return;
    if (!definitions[name]) definitions[name] = { name: name, type: textV2_(r.Tipe_Atribut) || 'text', order: Number(r.Urutan_Atribut) || 0, values: [] };
    definitions[name].values.push({
      value: textV2_(r.Nilai),
      hex: textV2_(r.Kode_Warna),
      order: Number(r.Urutan_Nilai) || 0,
      active: booleanV2_(r.Status_Aktif, true)
    });
  });

  var variantList = [];
  variantRows.forEach(function (r) {
    if (upperV2_(r.Barcode_Induk) !== upperV2_(product.barcode)) return;
    var active = booleanV2_(r.Status_Aktif, false);
    if (!includeInactive && !active) return;
    var attrs = parseJsonObjectV2_(r.Atribut_JSON);
    var stocks = {};
    stockRows.forEach(function (s) {
      if (textV2_(s.ID_Variasi) === textV2_(r.ID_Variasi)) stocks[textV2_(s.Lokasi_Rak)] = Number(s.Kuantitas) || 0;
    });
    var hpp = Number(r.HPP) || 0;
    var hargaJual = Number(r.Harga_Jual) || 0;
    var marginRp = hargaJual > 0 ? (hargaJual - hpp) : 0;
    var marginPersen = hargaJual > 0 ? Math.round(((hargaJual - hpp) / hargaJual) * 1000) / 10 : 0;

    variantList.push({
      id: textV2_(r.ID_Variasi),
      label: textV2_(r.Label_Variasi),
      attributes: attrs,
      active: active,
      stocks: stocks,
      hpp: hpp,
      harga_jual: hargaJual,
      margin_rp: marginRp,
      margin_persen: marginPersen
    });
  });

  var fullDefinitions = Object.keys(definitions).map(function (name) {
    var def = definitions[name];
    def.values.sort(function (a, b) { return a.order - b.order; });
    return { name: def.name, type: def.type, order: def.order, values: def.values.slice() };
  }).sort(function (a, b) { return a.order - b.order; });

  var attrList = fullDefinitions.map(function (def) {
    var allowed = {};
    variantList.forEach(function (v) {
      if (!v.active) return;
      if (Object.prototype.hasOwnProperty.call(v.attributes, def.name)) allowed[textV2_(v.attributes[def.name])] = true;
    });
    return {
      name: def.name,
      type: def.type,
      order: def.order,
      values: def.values.filter(function (v) { return allowed[v.value] && v.active; })
    };
  });

  return {
    barcode: product.barcode,
    nama: product.name,
    aktif: product.active,
    attributes: attrList,
    attributeDefinitions: fullDefinitions,
    variants: variantList,
    polos_mappings: productMappings,
    locations: getActiveLocationsV2_(ss)
  };
}

function createProductActionV2_(ss, payload) {
  var barcode = sanitizeSheetText_(payload.barcode);
  var name = sanitizeSheetText_(payload.nama);
  var pelapor = sanitizeSheetText_(payload.nama_pelapor);
  var trxId = textV2_(payload.transaction_id);
  var attributes = normalizeAttributesV2_(payload.attributes || []);
  var variants = normalizeVariantsV2_(payload.variants || [], attributes);
  var productActive = booleanV2_(payload.status_aktif, true);
  var defaultHpp = payload.default_hpp !== undefined && payload.default_hpp !== '' ? Number(payload.default_hpp) : 0;
  var defaultHargaJual = payload.default_harga_jual !== undefined && payload.default_harga_jual !== '' ? Number(payload.default_harga_jual) : 0;
  var barcodeKategoriTailor = textV2_(payload.barcode_kategori_tailor || payload.barcode_kategori);
  var autoCreateMissingTailor = booleanV2_(payload.auto_create_missing_tailor, true);

  if (!trxId) return { success: false, message: 'ID transaksi wajib diisi.' };
  if (!barcode) return { success: false, message: 'Barcode wajib diisi.' };
  if (!name || name.length < 2) return { success: false, message: 'Nama produk minimal 2 karakter.' };
  if (!isFinite(defaultHpp) || defaultHpp < 0 || Math.floor(defaultHpp) !== defaultHpp) {
    return { success: false, message: 'Modal (HPP) awal harus berupa bilangan bulat >= 0.' };
  }
  if (!isFinite(defaultHargaJual) || defaultHargaJual < 0 || Math.floor(defaultHargaJual) !== defaultHargaJual) {
    return { success: false, message: 'Harga jual awal harus berupa bilangan bulat >= 0.' };
  }
  if (findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode)) return { success: false, message: 'Barcode sudah terdaftar.' };
  if (trxExistsV2_(ss.getSheetByName(IBS.SHEET_LOG), trxId)) return { success: false, message: 'ID transaksi sudah pernah diproses.' };
  if (!variants.length) return { success: false, message: 'Minimal satu kombinasi variasi harus dipilih.' };
  if (!variants.some(function (v) { return v.active; })) return { success: false, message: 'Minimal satu variasi harus berstatus aktif.' };

  var shMaster = ss.getSheetByName(IBS.SHEET_MASTER);
  var shAttrs = ss.getSheetByName(IBS.SHEET_ATTRIBUTES);
  var shVariants = ss.getSheetByName(IBS.SHEET_VARIANTS);
  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  var now = new Date();
  var starts = [];
  try {
    var mStart = appendRowsV2_(shMaster, [[barcode, name, productActive, now, now]], starts);
    applyRowCheckboxV2_(shMaster, mStart, 1, 3);

    var attrRows = [];
    attributes.forEach(function (a, ai) {
      a.values.forEach(function (v, vi) {
        attrRows.push([barcode, a.name, a.type, v.value, v.hex || '', ai + 1, vi + 1, true]);
      });
    });
    if (attrRows.length) {
      var aStart = appendRowsV2_(shAttrs, attrRows, starts);
      applyRowCheckboxV2_(shAttrs, aStart, attrRows.length, 8);
    }

    // Arsitektur 100% Made-to-Order: Tidak menulis baris bersaldo 0 ke STOK_VARIASI
    var variantRows = [];
    var baseVarRow = shVariants.getLastRow() + 1;
    variants.forEach(function (v, idx) {
      var id = newVariantIdV2_();
      var label = makeVariantLabelV2_(v.attributes, attributes);
      var hppVal = v.hpp !== undefined ? Math.max(0, Number(v.hpp) || 0) : defaultHpp;
      var hjVal = v.harga_jual !== undefined ? Math.max(0, Number(v.harga_jual) || 0) : defaultHargaJual;
      var rNum = baseVarRow + idx;
      variantRows.push([
        id, barcode, label, JSON.stringify(v.attributes), v.active, now, now,
        hppVal, hjVal,
        '=IF(I' + rNum + '>0, I' + rNum + '-H' + rNum + ', 0)',
        '=IF(I' + rNum + '>0, (I' + rNum + '-H' + rNum + ')/I' + rNum + ', 0)'
      ]);
    });
    var vStart = appendRowsV2_(shVariants, variantRows, starts);
    applyRowCheckboxV2_(shVariants, vStart, variantRows.length, 5);
    shVariants.getRange(vStart, 8, variantRows.length, 3).setNumberFormat('"Rp "#,##0');
    shVariants.getRange(vStart, 11, variantRows.length, 1).setNumberFormat('0.0%');

    var logNewStart = appendRowsV2_(shLog, [[now, trxId, 'ProdukBaru', barcode, '', '', 0, 'SUCCESS', '', name, '', '', name, pelapor, defaultHpp, defaultHargaJual, 0, 0]], starts);
    shLog.getRange(logNewStart, 15, 1, 4).setNumberFormat('"Rp "#,##0');
    SpreadsheetApp.flush();

    var autoMapInfo = null;
    var extraMsg = '';
    if (barcodeKategoriTailor) {
      var mapRes = autoMapPolosCategoryV2_(ss, {
        barcode_produk: barcode,
        barcode_kategori_tailor: barcodeKategoriTailor,
        qty_per_unit: payload.qty_per_unit || 1,
        auto_create_missing_tailor: autoCreateMissingTailor
      });
      if (mapRes && mapRes.success) {
        autoMapInfo = mapRes.data;
        extraMsg = ' & otomatis ditautkan ke Gudang Polos (' + mapRes.data.matched_count + ' variasi).';
      } else if (mapRes && !mapRes.success) {
        extraMsg = ' (Peringatan Auto-Match Polos: ' + (mapRes.message || 'gagal menautkan otomatis') + ')';
      }
    }

    return {
      success: true,
      message: 'Produk dan ' + variants.length + ' variasi berhasil didaftarkan' + (extraMsg || '.'),
      data: { barcode: barcode, nama: name, jumlah_variasi: variants.length, auto_map: autoMapInfo }
    };
  } catch (err) {
    rollbackAppendsV2_(starts);
    throw err;
  }
}

function addVariantsActionV2_(ss, payload) {
  var barcode = textV2_(payload.barcode);
  var trxId = textV2_(payload.transaction_id);
  var pelapor = sanitizeSheetText_(payload.nama_pelapor);
  var defaultHpp = payload.default_hpp !== undefined && payload.default_hpp !== '' ? Number(payload.default_hpp) : 0;
  var defaultHargaJual = payload.default_harga_jual !== undefined && payload.default_harga_jual !== '' ? Number(payload.default_harga_jual) : 0;
  var requestedKategoriTailor = textV2_(payload.barcode_kategori_tailor || payload.barcode_kategori);
  var autoCreateMissingTailor = booleanV2_(payload.auto_create_missing_tailor, true);

  if (!barcode || !trxId) return { success: false, message: 'Barcode dan ID transaksi wajib diisi.' };
  if (!isFinite(defaultHpp) || defaultHpp < 0 || Math.floor(defaultHpp) !== defaultHpp) {
    return { success: false, message: 'Modal (HPP) harus berupa bilangan bulat >= 0.' };
  }
  if (!isFinite(defaultHargaJual) || defaultHargaJual < 0 || Math.floor(defaultHargaJual) !== defaultHargaJual) {
    return { success: false, message: 'Harga jual harus berupa bilangan bulat >= 0.' };
  }
  if (trxExistsV2_(ss.getSheetByName(IBS.SHEET_LOG), trxId)) return { success: false, message: 'ID transaksi sudah pernah diproses.' };
  var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode);
  if (!product) return { success: false, message: 'Produk tidak ditemukan.' };
  if (!product.active) return { success: false, message: 'Aktifkan produk sebelum menambahkan variasi.' };

  var current = composeProductDataV2_(ss, product, true);
  var existingDefs = (current.attributeDefinitions || []).map(function (a) { return { name: a.name, type: a.type, values: a.values.map(function (v) { return { value: v.value, hex: v.hex }; }) }; });
  var incomingDefs = normalizeAttributesV2_(payload.attributes || []);
  var mergedDefs = [];
  if (!existingDefs.length) {
    mergedDefs = incomingDefs;
  } else {
    if (incomingDefs.length !== existingDefs.length) return { success: false, message: 'Jumlah atribut harus sama dengan definisi produk yang sudah ada.' };
    existingDefs.forEach(function (oldDef) {
      var incoming = incomingDefs.filter(function (d) { return d.name.toLowerCase() === oldDef.name.toLowerCase(); })[0];
      if (!incoming || incoming.type !== oldDef.type) throw new Error('Definisi atribut ' + oldDef.name + ' tidak cocok.');
      var values = oldDef.values.slice();
      incoming.values.forEach(function (v) {
        if (!values.some(function (x) { return x.value.toLowerCase() === v.value.toLowerCase(); })) values.push(v);
      });
      mergedDefs.push({ name: oldDef.name, type: oldDef.type, values: values });
    });
  }

  var variants = normalizeVariantsV2_(payload.variants || [], mergedDefs);
  if (!variants.length) return { success: false, message: 'Pilih minimal satu kombinasi variasi baru.' };
  var oldKeys = {};
  current.variants.forEach(function (v) { oldKeys[variantKeyV2_(v.attributes, mergedDefs)] = true; });
  variants.forEach(function (v) {
    var key = variantKeyV2_(v.attributes, mergedDefs);
    if (oldKeys[key]) throw new Error('Kombinasi sudah ada: ' + makeVariantLabelV2_(v.attributes, mergedDefs));
    oldKeys[key] = true;
  });

  var shAttrs = ss.getSheetByName(IBS.SHEET_ATTRIBUTES);
  var shVariants = ss.getSheetByName(IBS.SHEET_VARIANTS);
  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  var starts = [];
  var changedVariantRows = [];
  var now = new Date();
  try {
    if (!existingDefs.length && mergedDefs.length) {
      current.variants.forEach(function (oldVariant) {
        if (oldVariant.active && Object.keys(oldVariant.attributes || {}).length === 0) {
          var oldRow = findVariantRowV2_(shVariants, oldVariant.id);
          if (oldRow) { changedVariantRows.push({ row: oldRow.row, oldActive: true, oldUpdated: shVariants.getRange(oldRow.row, 7).getValue() }); shVariants.getRange(oldRow.row, 5).setValue(false); shVariants.getRange(oldRow.row, 7).setValue(now); }
        }
      });
    }
    var newAttrRows = [];
    mergedDefs.forEach(function (def, ai) {
      var oldDef = existingDefs.filter(function (x) { return x.name === def.name; })[0];
      def.values.forEach(function (v, vi) {
        var alreadyExists = oldDef && oldDef.values.some(function (x) { return x.value.toLowerCase() === v.value.toLowerCase(); });
        if (!alreadyExists) newAttrRows.push([barcode, def.name, def.type, v.value, v.hex || '', ai + 1, vi + 1, true]);
      });
    });
    if (newAttrRows.length) {
      var aStart = appendRowsV2_(shAttrs, newAttrRows, starts);
      applyRowCheckboxV2_(shAttrs, aStart, newAttrRows.length, 8);
    }

    // Arsitektur 100% Made-to-Order: Tidak menulis baris bersaldo 0 ke STOK_VARIASI
    var vr = [];
    var baseVarRow = shVariants.getLastRow() + 1;
    variants.forEach(function (v, idx) {
      var id = newVariantIdV2_();
      var rNum = baseVarRow + idx;
      vr.push([
        id, barcode, makeVariantLabelV2_(v.attributes, mergedDefs), JSON.stringify(v.attributes), v.active, now, now,
        defaultHpp, defaultHargaJual,
        '=IF(I' + rNum + '>0, I' + rNum + '-H' + rNum + ', 0)',
        '=IF(I' + rNum + '>0, (I' + rNum + '-H' + rNum + ')/I' + rNum + ', 0)'
      ]);
    });
    var vStart = appendRowsV2_(shVariants, vr, starts);
    applyRowCheckboxV2_(shVariants, vStart, vr.length, 5);
    shVariants.getRange(vStart, 8, vr.length, 3).setNumberFormat('"Rp "#,##0');
    shVariants.getRange(vStart, 11, vr.length, 1).setNumberFormat('0.0%');

    var logAddVarStart = appendRowsV2_(shLog, [[now, trxId, 'TambahVariasi', barcode, '', '', 0, 'SUCCESS', '', product.name, '', '', product.name, pelapor, defaultHpp, defaultHargaJual, 0, 0]], starts);
    shLog.getRange(logAddVarStart, 15, 1, 4).setNumberFormat('"Rp "#,##0');
    SpreadsheetApp.flush();

    // Otomatis tautkan varian baru jika kategori polos dipilih atau produk induk sudah punya kategori tertaut
    var resolvedKategoriTailor = requestedKategoriTailor;
    if (!resolvedKategoriTailor && current.polos_mappings && current.polos_mappings.length > 0) {
      var existingTailorId = textV2_(current.polos_mappings[0].id_tailor);
      var shMasterTailor = ensureSheetHeadersV2_(ss, IBS.SHEET_MASTER_TAILOR, IBS.HEADERS.MASTER_TAILOR);
      if (existingTailorId && shMasterTailor.getLastRow() >= 2) {
        var mtMatch = readObjectsV2_(shMasterTailor).filter(function (m) {
          return textV2_(m.ID_Tailor) === existingTailorId;
        })[0];
        if (mtMatch) {
          resolvedKategoriTailor = textV2_(mtMatch.Barcode_Kategori);
        }
      }
    }

    var extraMsg = '';
    if (resolvedKategoriTailor) {
      var mapRes = autoMapPolosCategoryV2_(ss, {
        barcode_produk: barcode,
        barcode_kategori_tailor: resolvedKategoriTailor,
        qty_per_unit: payload.qty_per_unit || 1,
        auto_create_missing_tailor: autoCreateMissingTailor
      });
      if (mapRes && mapRes.success) {
        extraMsg = ' & otomatis ditautkan ke Gudang Polos (' + mapRes.data.matched_count + ' variasi).';
      } else if (mapRes && !mapRes.success) {
        extraMsg = ' (Peringatan Auto-Match Polos: ' + (mapRes.message || 'gagal menautkan otomatis') + ')';
      }
    }

    return {
      success: true,
      message: variants.length + ' variasi baru berhasil ditambahkan' + (extraMsg || '.')
    };
  } catch (err) {
    rollbackAppendsV2_(starts);
    changedVariantRows.forEach(function (x) { try { shVariants.getRange(x.row, 5).setValue(x.oldActive); shVariants.getRange(x.row, 7).setValue(x.oldUpdated); } catch (restoreError) {} });
    throw err;
  }
}

function updateVariantPricingActionV5_(ss, payload) {
  var barcode = textV2_(payload.barcode);
  var items = payload.items || [];
  var pelapor = sanitizeSheetText_(payload.nama_pelapor) || 'Admin';
  var trxId = textV2_(payload.transaction_id) || ('PRICE-' + Utilities.getUuid().substring(0, 8).toUpperCase());

  if (!Array.isArray(items) || !items.length) return { success: false, message: 'Tidak ada data harga/HPP yang dikirim.' };

  for (var k = 0; k < items.length; k++) {
    var chkItem = items[k] || {};
    if (chkItem.hpp !== '' && chkItem.hpp !== undefined) {
      var hVal = Number(chkItem.hpp);
      if (!isFinite(hVal) || hVal < 0 || Math.floor(hVal) !== hVal) {
        return { success: false, message: 'Modal (HPP) harus berupa bilangan bulat >= 0.' };
      }
    }
    if (chkItem.harga_jual !== '' && chkItem.harga_jual !== undefined) {
      var jVal = Number(chkItem.harga_jual);
      if (!isFinite(jVal) || jVal < 0 || Math.floor(jVal) !== jVal) {
        return { success: false, message: 'Harga jual harus berupa bilangan bulat >= 0.' };
      }
    }
  }

  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  if (payload.transaction_id && trxExistsV2_(shLog, trxId)) {
    return { success: false, message: 'ID transaksi sudah pernah diproses.' };
  }

  var shVar = ss.getSheetByName(IBS.SHEET_VARIANTS);
  var varRows = readObjectsV2_(shVar);
  var rowMap = {};
  varRows.forEach(function (r) {
    rowMap[textV2_(r.ID_Variasi)] = r;
    if (!barcode && items[0] && textV2_(items[0].variant_id) === textV2_(r.ID_Variasi)) {
      barcode = textV2_(r.Barcode_Induk);
    }
  });

  var product = barcode ? findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode) : null;
  var now = new Date();
  var updatedCount = 0;
  var lastMarginPersen = 0;

  items.forEach(function (item) {
    var vid = textV2_(item.variant_id);
    var target = rowMap[vid];
    if (!target) return;

    var hpp = (item.hpp !== '' && item.hpp !== undefined) ? Math.max(0, Number(item.hpp) || 0) : (Number(target.HPP) || 0);
    var hargaJual = (item.harga_jual !== '' && item.harga_jual !== undefined) ? Math.max(0, Number(item.harga_jual) || 0) : (Number(target.Harga_Jual) || 0);
    var rNum = target._row;

    shVar.getRange(rNum, 7).setValue(now);
    shVar.getRange(rNum, 8, 1, 4).setValues([[
      hpp,
      hargaJual,
      '=IF(I' + rNum + '>0, I' + rNum + '-H' + rNum + ', 0)',
      '=IF(I' + rNum + '>0, (I' + rNum + '-H' + rNum + ')/I' + rNum + ', 0)'
    ]]);
    shVar.getRange(rNum, 8, 1, 3).setNumberFormat('"Rp "#,##0');
    shVar.getRange(rNum, 11).setNumberFormat('0.0%');
    lastMarginPersen = hargaJual > 0 ? Math.round(((hargaJual - hpp) / hargaJual) * 1000) / 10 : 0;
    updatedCount++;
  });

  if (updatedCount > 0 && product) {
    shLog.appendRow([now, trxId, 'UpdateHargaHPP', barcode, '', '', updatedCount, 'SUCCESS', '', product.name, '', updatedCount + ' variasi', product.name + ' (Update HPP & Harga)', pelapor, 0, 0, 0, 0]);
    shLog.getRange(shLog.getLastRow(), 15, 1, 4).setNumberFormat('"Rp "#,##0');
  }

  SpreadsheetApp.flush();
  return {
    success: true,
    message: 'HPP & Harga Jual berhasil diperbarui untuk ' + updatedCount + ' variasi.',
    data: { updated_count: updatedCount, margin_persen: lastMarginPersen }
  };
}

function setVariantPricingBulkV5_(ss, payload) {
  var barcode = textV2_(payload.barcode);
  if (!barcode) return { success: false, message: 'Barcode produk wajib diisi.' };
  var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode);
  if (!product) return { success: false, message: 'Produk tidak ditemukan.' };

  var shVar = ss.getSheetByName(IBS.SHEET_VARIANTS);
  var rows = readObjectsV2_(shVar).filter(function (r) {
    return upperV2_(r.Barcode_Induk) === upperV2_(barcode) && booleanV2_(r.Status_Aktif, false);
  });
  if (!rows.length) return { success: false, message: 'Tidak ada variasi aktif pada produk ini.' };

  var items = rows.map(function (r) {
    return {
      variant_id: textV2_(r.ID_Variasi),
      hpp: payload.hpp !== '' && payload.hpp !== undefined ? payload.hpp : r.HPP,
      harga_jual: payload.harga_jual !== '' && payload.harga_jual !== undefined ? payload.harga_jual : r.Harga_Jual
    };
  });
  return updateVariantPricingActionV5_(ss, { barcode: barcode, items: items, nama_pelapor: payload.nama_pelapor || 'Admin' });
}

function updateProductActionV2_(ss, payload) {
  var barcode = textV2_(payload.barcode);
  var name = sanitizeSheetText_(payload.nama);
  var pelapor = sanitizeSheetText_(payload.nama_pelapor);
  if (!barcode) return { success: false, message: 'Barcode wajib diisi.' };
  var sh = ss.getSheetByName(IBS.SHEET_MASTER);
  var product = findMasterV2_(sh, barcode);
  if (!product) return { success: false, message: 'Produk tidak ditemukan.' };
  if (name && name.length < 2) return { success: false, message: 'Nama produk minimal 2 karakter.' };
  var active = booleanV2_(payload.status_aktif, product.active);
  var now = new Date();
  var log = ss.getSheetByName(IBS.SHEET_LOG);
  var trxId = textV2_(payload.transaction_id) || ('EDIT-' + Utilities.getUuid());
  if (trxExistsV2_(log, trxId)) return { success: false, message: 'ID transaksi sudah pernah diproses.' };
  var oldName = sh.getRange(product.row, 2).getValue();
  var oldActive = sh.getRange(product.row, 3).getValue();
  var oldUpdated = sh.getRange(product.row, 5).getValue();
  try {
    if (name) sh.getRange(product.row, 2).setValue(name);
    sh.getRange(product.row, 3).setValue(active);
    sh.getRange(product.row, 5).setValue(now);
    log.appendRow([now, trxId, 'EditProduk', barcode, '', '', 0, 'SUCCESS', '', name || product.name, '', '', name || product.name, pelapor, 0, 0, 0, 0]);
    SpreadsheetApp.flush();
    return { success: true, message: 'Data produk berhasil diperbarui.' };
  } catch (err) {
    try { sh.getRange(product.row, 2).setValue(oldName); sh.getRange(product.row, 3).setValue(oldActive); sh.getRange(product.row, 5).setValue(oldUpdated); } catch (restoreError) { console.error('Rollback edit produk: ' + restoreError); }
    throw err;
  }
}

function setVariantStatusActionV2_(ss, payload) {
  var barcode = textV2_(payload.barcode);
  var variantId = textV2_(payload.variant_id);
  if (!barcode || !variantId) return { success: false, message: 'Barcode dan ID variasi wajib diisi.' };
  var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode);
  if (!product) return { success: false, message: 'Produk tidak ditemukan.' };
  var sh = ss.getSheetByName(IBS.SHEET_VARIANTS);
  var row = findVariantRowV2_(sh, variantId);
  if (!row || upperV2_(row.barcode) !== upperV2_(barcode)) return { success: false, message: 'Variasi tidak ditemukan pada produk ini.' };
  var active = booleanV2_(payload.active, false);
  if (active && !product.active) return { success: false, message: 'Aktifkan produk induk terlebih dahulu.' };
  var now = new Date();
  var log = ss.getSheetByName(IBS.SHEET_LOG);
  var trxId = textV2_(payload.transaction_id) || ('VARSTATUS-' + Utilities.getUuid());
  if (trxExistsV2_(log, trxId)) return { success: false, message: 'ID transaksi sudah pernah diproses.' };
  var oldActive = sh.getRange(row.row, 5).getValue();
  var oldUpdated = sh.getRange(row.row, 7).getValue();
  try {
    sh.getRange(row.row, 5).setValue(active);
    sh.getRange(row.row, 7).setValue(now);
    log.appendRow([now, trxId, 'StatusVariasi', barcode, '', '', 0, 'SUCCESS', variantId, product.name, JSON.stringify(row.attributes || {}), row.label, product.name + ' · ' + row.label, sanitizeSheetText_(payload.nama_pelapor), row.hpp, row.harga_jual, 0, 0]);
    SpreadsheetApp.flush();
    return { success: true, message: active ? 'Variasi diaktifkan.' : 'Variasi dinonaktifkan.' };
  } catch (err) {
    try { sh.getRange(row.row, 5).setValue(oldActive); sh.getRange(row.row, 7).setValue(oldUpdated); } catch (restoreError) { console.error('Rollback status variasi: ' + restoreError); }
    throw err;
  }
}

/**
 * Transaksi stok tunggal (100% Made-to-Order):
 * - Hanya mendukung aktivitas "Keluar" yang memotong saldo STOK_TAILOR ("GUDANG POLOS").
 * - Harga jual & HPP dikunci langsung dari VARIASI_PRODUK.
 */
function saveStockTransactionV2_(ss, payload) {
  var trxId = textV2_(payload.transaction_id);
  var action = textV2_(payload.activity);
  var barcode = textV2_(payload.barcode);
  var variantId = textV2_(payload.variant_id);
  var qty = Number(payload.qty);
  var pelapor = sanitizeSheetText_(payload.nama_pelapor);

  if (!trxId || !barcode || !variantId) return { success: false, message: 'ID transaksi, barcode, dan variasi wajib diisi.' };
  if (!pelapor) return { success: false, message: 'Nama pelapor wajib diisi untuk transaksi ini.' };
  if (action !== 'Keluar') {
    return { success: false, message: 'Sistem Made-to-Order hanya memproses Barang Keluar via Gudang Polos.' };
  }
  if (textV2_(payload.qty) === '' || !isFinite(qty) || qty < 1 || Math.floor(qty) !== qty) {
    return { success: false, message: 'Kuantitas harus bilangan bulat >= 1.' };
  }

  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  if (trxExistsV2_(shLog, trxId)) return { success: false, message: 'Transaksi ini sudah pernah diproses.' };
  var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode);
  if (!product) return { success: false, message: 'Barcode belum terdaftar.' };
  if (!product.active) return { success: false, message: 'Produk sedang nonaktif.' };
  var variant = findVariantRowV2_(ss.getSheetByName(IBS.SHEET_VARIANTS), variantId);
  if (!variant || upperV2_(variant.barcode) !== upperV2_(barcode)) return { success: false, message: 'Variasi tidak cocok dengan barcode produk.' };
  if (!variant.active) return { success: false, message: 'Variasi sedang nonaktif.' };
  if ((Number(variant.hpp) || 0) <= 0 || (Number(variant.harga_jual) || 0) <= 0) {
    return { success: false, message: 'Modal (HPP) dan Harga Jual untuk "' + product.name + ' · ' + variant.label + '" belum diatur (> Rp 0).' };
  }

  var shMapping = ensureSheetHeadersV2_(ss, IBS.SHEET_MAPPING_POLOS, IBS.HEADERS.MAPPING_POLOS_PRODUK);
  var mappingRows = shMapping.getLastRow() >= 2 ? readObjectsV2_(shMapping) : [];
  var polosMapping = findPolosMappingForItem_(mappingRows, barcode, variantId);
  if (!polosMapping) {
    return {
      success: false,
      message: 'Produk "' + product.name + ' · ' + variant.label + '" belum ditautkan ke Gudang Polos. Atur tautan kaos polos di menu Kelola Data Produk terlebih dahulu.'
    };
  }

  var shStokTailor = ensureSheetHeadersV2_(ss, IBS.SHEET_STOK_TAILOR, IBS.HEADERS.STOK_TAILOR);
  var polosRec = findGudangPolosRowInStokTailor_(shStokTailor, polosMapping.id_tailor);
  var neededPolos = qty * polosMapping.qty_per_unit;
  var availPolos = polosRec ? polosRec.stok_ready : 0;
  if (availPolos < neededPolos) {
    return {
      success: false,
      message: 'Stok Gudang Polos tidak cukup (' + (polosMapping.detail_polos || polosMapping.id_tailor) + '). Dibutuhkan: ' + neededPolos + ' pcs, Tersedia: ' + availPolos + ' pcs.'
    };
  }

  var hppSatuan = Number(variant.hpp) || 0;
  var hargaJualSatuan = Number(variant.harga_jual) || 0;
  var totalOmzet = qty * hargaJualSatuan;
  var totalProfit = qty * (hargaJualSatuan - hppSatuan);

  var polosChanges = [];
  var appendStarts = [];
  var now = new Date();
  try {
    var rowsToUpdate = polosRec.rows || [polosRec.row];
    for (var rIdx = 0; rIdx < rowsToUpdate.length; rIdx++) {
      var targetRowNum = rowsToUpdate[rIdx];
      var oldRowQty = Number(shStokTailor.getRange(targetRowNum, 7).getValue()) || 0;
      var oldRowUpdated = shStokTailor.getRange(targetRowNum, 8).getValue();
      var nextRowQty = rIdx === 0 ? (availPolos - neededPolos) : 0;
      shStokTailor.getRange(targetRowNum, 7, 1, 2).setValues([[nextRowQty, now]]);
      shStokTailor.getRange(targetRowNum, 7).setNumberFormat('#,##0');
      polosChanges.push({ row: targetRowNum, oldQty: oldRowQty, oldUpdated: oldRowUpdated });
    }

    var shLogTailor = ensureSheetHeadersV2_(ss, IBS.SHEET_LOG_TAILOR, IBS.HEADERS.LOG_TAILOR);
    var displayUsage = (polosMapping.detail_polos || polosMapping.id_tailor) + ' → Pesanan: ' + product.name + ' · ' + variant.label + ' (' + qty + ' pcs) [Petugas: ' + pelapor + ']';
    var ltStart = appendRowsV2_(shLogTailor, [[
      'USE-' + trxId, now, 'PEMAKAIAN_PESANAN', IBS.TAILOR_GUDANG_POLOS_NAME, polosMapping.id_tailor,
      displayUsage, neededPolos, 0, 0, 0, 'CONSUMED', trxId
    ]], appendStarts);
    shLogTailor.getRange(ltStart, 7, 1, 4).setNumberFormat('#,##0');
    shLogTailor.getRange(ltStart, 12).setNumberFormat('@');

    var display = product.name + (variant.label ? ' · ' + variant.label : '');
    var logStart = appendRowsV2_(shLog, [[
      now, trxId, 'Keluar', barcode, 'Gudang Polos', '', qty, 'SUCCESS', variantId,
      product.name, JSON.stringify(variant.attributes || {}), variant.label, display, pelapor,
      hppSatuan, hargaJualSatuan, totalOmzet, totalProfit
    ]], appendStarts);
    shLog.getRange(logStart, 15, 1, 4).setNumberFormat('"Rp "#,##0');

    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Transaksi keluar berhasil disimpan.',
      data: {
        barcode: barcode,
        nama: product.name,
        variant_id: variantId,
        variasi: variant.label,
        total_omzet: totalOmzet,
        total_profit: totalProfit
      }
    };
  } catch (err) {
    polosChanges.forEach(function (pc) {
      try {
        shStokTailor.getRange(pc.row, 7, 1, 2).setValues([[pc.oldQty, pc.oldUpdated]]);
      } catch (rbPolosErr) {}
    });
    rollbackAppendsV2_(appendStarts);
    throw err;
  }
}

/**
 * ENDPOINT: SAVE_BULK_KELUAR (100% Made-to-Order / Gudang Polos)
 * - Harga jual & HPP dikunci 100% dari VARIASI_PRODUK (menolak jika <= 0).
 * - Mengkonsolidasikan baris duplikat GUDANG POLOS pada STOK_TAILOR dan menolkan baris duplikat.
 */
function saveBulkKeluarV2_(ss, payload) {
  var batchId = textV2_(payload.batch_id || payload.transaction_id);
  var pelapor = sanitizeSheetText_(payload.nama_pelapor);
  var items = Array.isArray(payload.items) ? payload.items : [];

  if (!batchId) return { success: false, message: 'ID batch transaksi wajib diisi.' };
  if (!pelapor) return { success: false, message: 'Nama petugas wajib diisi.' };
  if (!items.length) return { success: false, message: 'Keranjang barang keluar masih kosong.' };

  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  if (shLog && shLog.getLastRow() >= 2) {
    var existingTrx = shLog.getRange(2, 2, shLog.getLastRow() - 1, 1).getDisplayValues();
    for (var t = 0; t < existingTrx.length; t++) {
      var idVal = textV2_(existingTrx[t][0]);
      if (idVal === batchId || idVal.indexOf(batchId + '-') === 0) {
        return { success: false, message: 'Batch transaksi "' + batchId + '" sudah pernah diproses.' };
      }
    }
  }

  var masterMap = {};
  readObjectsV2_(ss.getSheetByName(IBS.SHEET_MASTER)).forEach(function (m) {
    var bKey = upperV2_(m.Barcode);
    if (bKey) {
      masterMap[bKey] = {
        barcode: textV2_(m.Barcode),
        name: textV2_(m.Nama),
        active: booleanV2_(m.Status_Aktif, false)
      };
    }
  });

  var variantMap = {};
  readObjectsV2_(ss.getSheetByName(IBS.SHEET_VARIANTS)).forEach(function (v) {
    var vid = textV2_(v.ID_Variasi);
    if (vid) {
      variantMap[vid] = {
        id: vid,
        barcode: textV2_(v.Barcode_Induk),
        label: textV2_(v.Label_Variasi) || 'Standar',
        attributes: parseJsonObjectV2_(v.Atribut_JSON),
        active: booleanV2_(v.Status_Aktif, false),
        hpp: Number(v.HPP) || 0,
        harga_jual: Number(v.Harga_Jual) || 0
      };
    }
  });

  var shMapping = ensureSheetHeadersV2_(ss, IBS.SHEET_MAPPING_POLOS, IBS.HEADERS.MAPPING_POLOS_PRODUK);
  var mappingRows = shMapping.getLastRow() >= 2 ? readObjectsV2_(shMapping) : [];

  var shStokTailor = ensureSheetHeadersV2_(ss, IBS.SHEET_STOK_TAILOR, IBS.HEADERS.STOK_TAILOR);
  var polosMap = {};
  if (shStokTailor.getLastRow() >= 2) {
    readObjectsV2_(shStokTailor).forEach(function (st) {
      if (!isGudangPolosName_(st.Nama_Penjahit)) return;
      var idT = textV2_(st.ID_Tailor);
      if (!idT) return;
      if (!polosMap[idT]) {
        polosMap[idT] = {
          row: st._row,
          rows: [st._row],
          id_tailor: idT,
          barcode_kategori: textV2_(st.Barcode_Kategori),
          nama_kategori: textV2_(st.Nama_Kategori),
          warna: textV2_(st.Warna) || '-',
          ukuran: textV2_(st.Ukuran) || '-',
          stok_ready: Number(st.Stok_Ready) || 0
        };
      } else {
        polosMap[idT].rows.push(st._row);
        polosMap[idT].stok_ready += (Number(st.Stok_Ready) || 0);
      }
    });
  }

  // Tahap 1: Validasi seluruh baris item, pastikan 100% tertaut ke Gudang Polos, & kunci harga dari master
  var resolvedItems = [];
  var polosDemandByTailorId = {};

  for (var i = 0; i < items.length; i++) {
    var raw = items[i] || {};
    var barcode = textV2_(raw.barcode);
    var variantId = textV2_(raw.variant_id);
    var qty = Number(raw.qty);

    if (!barcode || !variantId) {
      return { success: false, message: 'Baris ke-' + (i + 1) + ': Barcode dan variasi wajib diisi.' };
    }
    if (!isFinite(qty) || qty < 1 || Math.floor(qty) !== qty) {
      return { success: false, message: 'Baris ke-' + (i + 1) + ': Kuantitas harus bilangan bulat >= 1.' };
    }

    var product = masterMap[upperV2_(barcode)];
    if (!product) {
      return { success: false, message: 'Baris ke-' + (i + 1) + ': Barcode "' + barcode + '" tidak ditemukan.' };
    }
    if (!product.active) {
      return { success: false, message: 'Baris ke-' + (i + 1) + ': Produk "' + product.name + '" sedang nonaktif.' };
    }

    var variant = variantMap[variantId];
    if (!variant || upperV2_(variant.barcode) !== upperV2_(barcode)) {
      return { success: false, message: 'Baris ke-' + (i + 1) + ': Variasi tidak cocok dengan produk "' + product.name + '".' };
    }
    if (!variant.active) {
      return { success: false, message: 'Baris ke-' + (i + 1) + ': Variasi "' + variant.label + '" sedang nonaktif.' };
    }

    var hppSatuan = Number(variant.hpp) || 0;
    var hargaJualSatuan = Number(variant.harga_jual) || 0;
    if (hppSatuan <= 0 || hargaJualSatuan <= 0) {
      return {
        success: false,
        message: 'Baris ke-' + (i + 1) + ': Modal (HPP) dan Harga Jual untuk "' + product.name + ' · ' + variant.label + '" belum diatur (> Rp 0). Silakan atur harga terlebih dahulu.'
      };
    }

    var mapping = findPolosMappingForItem_(mappingRows, barcode, variantId);
    if (!mapping) {
      return {
        success: false,
        message: 'Baris ke-' + (i + 1) + ': Produk "' + product.name + ' · ' + variant.label + '" belum ditautkan ke Gudang Polos. Atur tautan kaos polos di menu Kelola Data Produk terlebih dahulu.'
      };
    }

    var totalOmzet = qty * hargaJualSatuan;
    var totalProfit = qty * (hargaJualSatuan - hppSatuan);
    var neededPolos = qty * mapping.qty_per_unit;

    if (!polosDemandByTailorId[mapping.id_tailor]) {
      polosDemandByTailorId[mapping.id_tailor] = {
        id_tailor: mapping.id_tailor,
        detail_polos: mapping.detail_polos,
        total_needed: 0
      };
    }
    polosDemandByTailorId[mapping.id_tailor].total_needed += neededPolos;

    resolvedItems.push({
      index: i + 1,
      barcode: barcode,
      product: product,
      variant: variant,
      qty: qty,
      hpp_satuan: hppSatuan,
      harga_jual_satuan: hargaJualSatuan,
      total_omzet: totalOmzet,
      total_profit: totalProfit,
      mapping: mapping,
      needed_polos: neededPolos
    });
  }

  // Validasi akumulasi stok Gudang Polos (STOK_TAILOR: GUDANG POLOS) per ID_Tailor
  var tailorIdsNeeded = Object.keys(polosDemandByTailorId);
  for (var p = 0; p < tailorIdsNeeded.length; p++) {
    var tId = tailorIdsNeeded[p];
    var demand = polosDemandByTailorId[tId];
    var gpRec = polosMap[tId];
    var avail = gpRec ? gpRec.stok_ready : 0;
    var labelPolos = gpRec
      ? (gpRec.nama_kategori + ' · ' + gpRec.warna + ' · ' + gpRec.ukuran)
      : (demand.detail_polos || tId);
    if (avail < demand.total_needed) {
      return {
        success: false,
        message: 'Stok Gudang Polos kurang untuk "' + labelPolos + '". Total dibutuhkan: ' + demand.total_needed + ' pcs, Tersedia: ' + avail + ' pcs.'
      };
    }
  }

  // Tahap 2: Eksekusi atomik dengan rollback penuh (100% Gudang Polos)
  var now = new Date();
  var polosChanges = [];
  var appendStarts = [];
  var shLogTailor = ensureSheetHeadersV2_(ss, IBS.SHEET_LOG_TAILOR, IBS.HEADERS.LOG_TAILOR);

  try {
    var polosDeductedSummary = [];
    tailorIdsNeeded.forEach(function (tId) {
      var demand = polosDemandByTailorId[tId];
      var gpRec = polosMap[tId];
      var newQty = gpRec.stok_ready - demand.total_needed;
      var targetRows = gpRec.rows && gpRec.rows.length ? gpRec.rows : [gpRec.row];

      for (var rIdx = 0; rIdx < targetRows.length; rIdx++) {
        var rNum = targetRows[rIdx];
        var oldQty = Number(shStokTailor.getRange(rNum, 7).getValue()) || 0;
        var oldUpdated = shStokTailor.getRange(rNum, 8).getValue();
        var rowQtyToWrite = rIdx === 0 ? newQty : 0;
        shStokTailor.getRange(rNum, 7, 1, 2).setValues([[rowQtyToWrite, now]]);
        shStokTailor.getRange(rNum, 7).setNumberFormat('#,##0');
        polosChanges.push({ row: rNum, oldQty: oldQty, oldUpdated: oldUpdated });
      }

      polosDeductedSummary.push({
        id_tailor: tId,
        detail_polos: gpRec.nama_kategori + ' · ' + gpRec.warna + ' · ' + gpRec.ukuran,
        qty_potong: demand.total_needed,
        sisa_stok: newQty
      });
    });

    var opLogRows = [];
    var tailorLogRows = [];
    var grandTotalQty = 0;
    var grandTotalOmzet = 0;
    var grandTotalProfit = 0;

    resolvedItems.forEach(function (item) {
      var itemTrxId = batchId + '-' + item.index;
      var display = item.product.name + (item.variant.label ? ' · ' + item.variant.label : '');

      opLogRows.push([
        now, itemTrxId, 'Keluar', item.barcode, 'Gudang Polos', '', item.qty, 'SUCCESS', item.variant.id,
        item.product.name, JSON.stringify(item.variant.attributes || {}), item.variant.label, display, pelapor,
        item.hpp_satuan, item.harga_jual_satuan, item.total_omzet, item.total_profit
      ]);

      var ltId = 'USE-' + batchId + '-' + item.index;
      var detailUsage = (item.mapping.detail_polos || item.mapping.id_tailor) + ' → Pesanan: ' + display + ' (' + item.qty + ' pcs) [Petugas: ' + pelapor + ']';
      tailorLogRows.push([
        ltId, now, 'PEMAKAIAN_PESANAN', IBS.TAILOR_GUDANG_POLOS_NAME, item.mapping.id_tailor,
        detailUsage, item.needed_polos, 0, 0, 0, 'CONSUMED', batchId
      ]);

      grandTotalQty += item.qty;
      grandTotalOmzet += item.total_omzet;
      grandTotalProfit += item.total_profit;
    });

    var opStart = appendRowsV2_(shLog, opLogRows, appendStarts);
    shLog.getRange(opStart, 2, opLogRows.length, 1).setNumberFormat('@');
    shLog.getRange(opStart, 15, opLogRows.length, 4).setNumberFormat('"Rp "#,##0');

    var ltStart = appendRowsV2_(shLogTailor, tailorLogRows, appendStarts);
    shLogTailor.getRange(ltStart, 1, tailorLogRows.length, 1).setNumberFormat('@');
    shLogTailor.getRange(ltStart, 5, tailorLogRows.length, 1).setNumberFormat('@');
    shLogTailor.getRange(ltStart, 7, tailorLogRows.length, 4).setNumberFormat('#,##0');
    shLogTailor.getRange(ltStart, 12, tailorLogRows.length, 1).setNumberFormat('@');

    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Berhasil memproses ' + resolvedItems.length + ' item (' + grandTotalQty + ' pcs) dari Gudang Polos.',
      data: {
        batch_id: batchId,
        total_items: resolvedItems.length,
        total_qty: grandTotalQty,
        total_omzet: grandTotalOmzet,
        total_profit: grandTotalProfit,
        polos_deducted: polosDeductedSummary
      }
    };
  } catch (err) {
    polosChanges.forEach(function (pc) {
      try {
        shStokTailor.getRange(pc.row, 7, 1, 2).setValues([[pc.oldQty, pc.oldUpdated]]);
      } catch (rbPolosErr) {}
    });
    rollbackAppendsV2_(appendStarts);
    throw err;
  }
}

// ==========================================
// MODUL GUDANG POLOS (STOK_TAILOR) & AUTO-MATCH KATEGORI
// ==========================================
function getGudangPolosList_(ss) {
  var items = getCachedV5_('IBS_GUDANG_POLOS_V6_6', function () {
    var shStok = ensureSheetHeadersV2_(ss, IBS.SHEET_STOK_TAILOR, IBS.HEADERS.STOK_TAILOR);
    if (shStok.getLastRow() < 2) return [];
    var consolidated = {};
    readObjectsV2_(shStok).forEach(function (r) {
      if (!isGudangPolosName_(r.Nama_Penjahit)) return;
      var idT = textV2_(r.ID_Tailor);
      if (!idT) return;
      var qty = Number(r.Stok_Ready) || 0;
      if (!consolidated[idT]) {
        consolidated[idT] = {
          id_tailor: idT,
          nama_penjahit: IBS.TAILOR_GUDANG_POLOS_NAME,
          barcode_kategori: textV2_(r.Barcode_Kategori),
          nama_kategori: textV2_(r.Nama_Kategori),
          warna: textV2_(r.Warna) || '-',
          ukuran: textV2_(r.Ukuran) || '-',
          stok_ready: qty,
          stok_tersedia: qty,
          updated_at: r.Updated_At
        };
      } else {
        consolidated[idT].stok_ready += qty;
        consolidated[idT].stok_tersedia += qty;
      }
    });

    var list = Object.keys(consolidated).map(function (k) { return consolidated[k]; });
    list.sort(function (a, b) {
      return a.nama_kategori.localeCompare(b.nama_kategori) ||
        a.warna.localeCompare(b.warna) ||
        a.ukuran.localeCompare(b.ukuran);
    });
    return list;
  });
  return { items: items };
}

function getPolosMappingsList_(ss) {
  var items = getCachedV5_('IBS_POLOS_MAPPING_V6_6', function () {
    var sh = ensureSheetHeadersV2_(ss, IBS.SHEET_MAPPING_POLOS, IBS.HEADERS.MAPPING_POLOS_PRODUK);
    if (sh.getLastRow() < 2) return [];
    return readObjectsV2_(sh).map(function (r) {
      var idVar = textV2_(r.ID_Variasi) || '*';
      return {
        mapping_id: textV2_(r.Mapping_ID),
        barcode_produk: textV2_(r.Barcode_Produk),
        id_variasi: idVar,
        label_variasi: textV2_(r.Label_Variasi) || (idVar === '*' ? 'Semua Variasi (*)' : idVar),
        id_tailor: textV2_(r.ID_Tailor),
        detail_polos: textV2_(r.Detail_Polos),
        qty_per_unit: Math.max(1, Number(r.Qty_Per_Unit) || 1),
        updated_at: r.Updated_At
      };
    }).filter(function (m) {
      return m.barcode_produk !== '' && m.id_tailor !== '';
    });
  });
  return { items: items };
}

function findPolosMappingForItem_(mappingRows, barcode, variantId) {
  var bUpper = upperV2_(barcode);
  var vId = textV2_(variantId);
  var wildcardMatch = null;

  for (var i = 0; i < mappingRows.length; i++) {
    var r = mappingRows[i];
    if (upperV2_(r.Barcode_Produk) !== bUpper) continue;
    var rowVarId = textV2_(r.ID_Variasi) || '*';
    var idTailor = textV2_(r.ID_Tailor);
    if (!idTailor) continue;

    var parsed = {
      mapping_id: textV2_(r.Mapping_ID),
      barcode_produk: textV2_(r.Barcode_Produk),
      id_variasi: rowVarId,
      id_tailor: idTailor,
      detail_polos: textV2_(r.Detail_Polos),
      qty_per_unit: Math.max(1, Number(r.Qty_Per_Unit) || 1)
    };

    if (rowVarId === vId) {
      return parsed;
    }
    if (rowVarId === '*' && !wildcardMatch) {
      wildcardMatch = parsed;
    }
  }
  return wildcardMatch;
}

function findGudangPolosRowInStokTailor_(shStokTailor, idTailor) {
  if (!shStokTailor || shStokTailor.getLastRow() < 2) return null;
  var rows = readObjectsV2_(shStokTailor);
  var found = null;
  for (var i = 0; i < rows.length; i++) {
    if (textV2_(rows[i].ID_Tailor) === textV2_(idTailor) && isGudangPolosName_(rows[i].Nama_Penjahit)) {
      var qty = Number(rows[i].Stok_Ready) || 0;
      if (!found) {
        found = {
          row: rows[i]._row,
          rows: [rows[i]._row],
          id_tailor: textV2_(rows[i].ID_Tailor),
          barcode_kategori: textV2_(rows[i].Barcode_Kategori),
          nama_kategori: textV2_(rows[i].Nama_Kategori),
          warna: textV2_(rows[i].Warna) || '-',
          ukuran: textV2_(rows[i].Ukuran) || '-',
          stok_ready: qty,
          stok_tersedia: qty
        };
      } else {
        found.rows.push(rows[i]._row);
        found.stok_ready += qty;
        found.stok_tersedia += qty;
      }
    }
  }
  return found;
}

/**
 * ENDPOINT: AUTO_MAP_POLOS_CATEGORY
 * Mencocokkan seluruh variasi produk desain (VARIASI_PRODUK) terhadap 1 Kategori Kaos Polos Induk
 * (Barcode_Kategori di MASTER_TAILOR) berdasarkan kesamaan Warna & Ukuran secara otomatis.
 * Mendukung auto_create_missing_tailor untuk membuat varian polos yang belum terdaftar di MASTER_TAILOR.
 */
function autoMapPolosCategoryV2_(ss, payload) {
  var barcodeProduk = textV2_(payload.barcode_produk || payload.barcode);
  var barcodeKategori = textV2_(payload.barcode_kategori_tailor || payload.barcode_kategori);
  var qtyPerUnit = payload.qty_per_unit !== undefined && payload.qty_per_unit !== ''
    ? Number(payload.qty_per_unit)
    : 1;
  var autoCreateMissing = booleanV2_(payload.auto_create_missing_tailor, false);

  if (!barcodeProduk) {
    return { success: false, message: 'Barcode produk induk wajib diisi.' };
  }
  if (!barcodeKategori) {
    return { success: false, message: 'Kategori kaos polos induk wajib dipilih.' };
  }
  if (!isFinite(qtyPerUnit) || qtyPerUnit < 1 || Math.floor(qtyPerUnit) !== qtyPerUnit) {
    return { success: false, message: 'Kebutuhan polos per unit harus bilangan bulat minimal 1.' };
  }

  var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcodeProduk);
  if (!product) {
    return { success: false, message: 'Produk induk "' + barcodeProduk + '" tidak ditemukan di MASTER_PRODUK.' };
  }

  var shVar = ss.getSheetByName(IBS.SHEET_VARIANTS);
  var allProductVars = readObjectsV2_(shVar).filter(function (v) {
    return upperV2_(v.Barcode_Induk) === upperV2_(barcodeProduk);
  });
  var targetVars = allProductVars.filter(function (v) {
    return booleanV2_(v.Status_Aktif, false);
  });
  if (!targetVars.length) targetVars = allProductVars;
  if (!targetVars.length) {
    return { success: false, message: 'Produk "' + product.name + '" belum memiliki variasi.' };
  }

  var shMasterTailor = ensureSheetHeadersV2_(ss, IBS.SHEET_MASTER_TAILOR, IBS.HEADERS.MASTER_TAILOR);
  var tailorRows = shMasterTailor.getLastRow() >= 2 ? readObjectsV2_(shMasterTailor) : [];
  var categoryTailors = tailorRows.filter(function (m) {
    return upperV2_(m.Barcode_Kategori) === upperV2_(barcodeKategori) &&
      booleanV2_(m.Status_Aktif, true);
  });

  if (!categoryTailors.length) {
    return {
      success: false,
      message: 'Kategori kaos polos "' + barcodeKategori + '" tidak ditemukan atau nonaktif di Master Tailor.'
    };
  }

  var canonicalBarcodeKat = textV2_(categoryTailors[0].Barcode_Kategori);
  var namaKategori = textV2_(categoryTailors[0].Nama_Kategori) || canonicalBarcodeKat;
  var defaultUpah = Number(categoryTailors[0].Upah_Per_Pcs) || 0;

  // Indeks MASTER_TAILOR berdasarkan kombinasi normalisasi (Warna || Ukuran)
  var tailorLookup = {};
  categoryTailors.forEach(function (m) {
    var normW = normalizeMatchToken_(m.Warna);
    var normU = normalizeMatchToken_(m.Ukuran);
    var key = normW + '||' + normU;
    if (!tailorLookup[key]) {
      tailorLookup[key] = {
        id_tailor: textV2_(m.ID_Tailor),
        barcode_kategori: canonicalBarcodeKat,
        nama_kategori: textV2_(m.Nama_Kategori) || namaKategori,
        warna: textV2_(m.Warna) || '-',
        ukuran: textV2_(m.Ukuran) || '-',
        upah_per_pcs: Number(m.Upah_Per_Pcs) || 0
      };
    }
  });

  var shMapping = ensureSheetHeadersV2_(ss, IBS.SHEET_MAPPING_POLOS, IBS.HEADERS.MAPPING_POLOS_PRODUK);
  var existingMappings = shMapping.getLastRow() >= 2 ? readObjectsV2_(shMapping) : [];
  var existingByVarId = {};
  var wildcardRowsToDelete = [];

  existingMappings.forEach(function (r) {
    if (upperV2_(r.Barcode_Produk) !== upperV2_(barcodeProduk)) return;
    var vid = textV2_(r.ID_Variasi) || '*';
    if (vid === '*') {
      wildcardRowsToDelete.push(r._row);
    } else {
      existingByVarId[vid] = r;
    }
  });

  var now = new Date();
  var matchedResults = [];
  var unmatchedVariants = [];
  var createdTailorCount = 0;
  var newMappingRows = [];

  targetVars.forEach(function (v) {
    var vid = textV2_(v.ID_Variasi);
    var labelVariasi = textV2_(v.Label_Variasi) || 'Standar';
    var attrs = parseJsonObjectV2_(v.Atribut_JSON);
    var extracted = extractColorSizeFromVariant_(attrs, labelVariasi);
    var normW = normalizeMatchToken_(extracted.warna);
    var normU = normalizeMatchToken_(extracted.ukuran);

    var matchedTailor = tailorLookup[normW + '||' + normU] || null;
    if (!matchedTailor && normU === '-') {
      matchedTailor = tailorLookup[normW + '||ALLSIZE'] || null;
    }
    if (!matchedTailor && normU === 'ALLSIZE') {
      matchedTailor = tailorLookup[normW + '||' + '-'] || null;
    }

    if (!matchedTailor && autoCreateMissing) {
      var newTailorId = 'TLR-' + Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();
      var warnaNew = sanitizeSheetText_(extracted.warna) || '-';
      var ukuranNew = sanitizeSheetText_(extracted.ukuran) || '-';
      var insertTailorRow = shMasterTailor.getLastRow() + 1;
      shMasterTailor.getRange(insertTailorRow, 1, 1, 9).setValues([[
        newTailorId, canonicalBarcodeKat, namaKategori, warnaNew, ukuranNew, defaultUpah, '', '', true
      ]]);
      shMasterTailor.getRange(insertTailorRow, 1, 1, 2).setNumberFormat('@');
      shMasterTailor.getRange(insertTailorRow, 6).setNumberFormat('"Rp "#,##0');
      shMasterTailor.getRange(insertTailorRow, 7, 1, 2).setNumberFormat('@');
      applyRowCheckboxV2_(shMasterTailor, insertTailorRow, 1, 9);

      matchedTailor = {
        id_tailor: newTailorId,
        barcode_kategori: canonicalBarcodeKat,
        nama_kategori: namaKategori,
        warna: warnaNew,
        ukuran: ukuranNew,
        upah_per_pcs: defaultUpah
      };
      tailorLookup[normW + '||' + normU] = matchedTailor;
      createdTailorCount++;
    }

    if (matchedTailor) {
      var detailPolos = matchedTailor.nama_kategori + ' (' + matchedTailor.barcode_kategori + ') · ' +
        matchedTailor.warna + ' · ' + matchedTailor.ukuran;
      var existingRow = existingByVarId[vid];

      if (existingRow) {
        shMapping.getRange(existingRow._row, 2, 1, 7).setValues([[
          product.barcode, vid, labelVariasi, matchedTailor.id_tailor, detailPolos, qtyPerUnit, now
        ]]);
        shMapping.getRange(existingRow._row, 2, 1, 2).setNumberFormat('@');
        shMapping.getRange(existingRow._row, 5).setNumberFormat('@');
        shMapping.getRange(existingRow._row, 7).setNumberFormat('#,##0');
      } else {
        var newMapId = 'MAP-' + Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();
        newMappingRows.push([
          newMapId, product.barcode, vid, labelVariasi, matchedTailor.id_tailor, detailPolos, qtyPerUnit, now
        ]);
      }

      matchedResults.push({
        id_variasi: vid,
        label_variasi: labelVariasi,
        id_tailor: matchedTailor.id_tailor,
        detail_polos: detailPolos
      });
    } else {
      unmatchedVariants.push({
        id_variasi: vid,
        label_variasi: labelVariasi,
        warna: extracted.warna,
        ukuran: extracted.ukuran
      });
    }
  });

  if (newMappingRows.length > 0) {
    var startRow = shMapping.getLastRow() + 1;
    shMapping.getRange(startRow, 1, newMappingRows.length, 8).setValues(newMappingRows);
    shMapping.getRange(startRow, 1, newMappingRows.length, 3).setNumberFormat('@');
    shMapping.getRange(startRow, 5, newMappingRows.length, 1).setNumberFormat('@');
    shMapping.getRange(startRow, 7, newMappingRows.length, 1).setNumberFormat('#,##0');
  }

  // Hapus baris wildcard (*) lama jika auto-match per variasi berhasil disimpan
  if (matchedResults.length > 0 && wildcardRowsToDelete.length > 0) {
    wildcardRowsToDelete.sort(function (a, b) { return b - a; });
    wildcardRowsToDelete.forEach(function (rIdx) {
      shMapping.deleteRow(rIdx);
    });
  }

  SpreadsheetApp.flush();
  invalidateAppCache_();

  var msg = 'Auto-match selesai: ' + matchedResults.length + ' dari ' + targetVars.length + ' variasi berhasil ditautkan ke ' + namaKategori + '.';
  if (createdTailorCount > 0) {
    msg += ' (' + createdTailorCount + ' varian polos baru dibuat otomatis di Master Tailor).';
  }
  if (unmatchedVariants.length > 0) {
    msg += ' ' + unmatchedVariants.length + ' variasi belum ada di Master Tailor.';
  }

  return {
    success: true,
    message: msg,
    data: {
      barcode_produk: product.barcode,
      barcode_kategori_tailor: canonicalBarcodeKat,
      nama_kategori_tailor: namaKategori,
      total_variants: targetVars.length,
      matched_count: matchedResults.length,
      created_tailor_count: createdTailorCount,
      unmatched_count: unmatchedVariants.length,
      unmatched_variants: unmatchedVariants,
      polos_mappings: getPolosMappingsList_(ss).items,
      master_tailor: getMasterTailorList_(ss, false).items
    }
  };
}

function savePolosMappingV2_(ss, payload) {
  var barcodeProduk = textV2_(payload.barcode_produk || payload.barcode);
  var idVariasi = textV2_(payload.id_variasi || payload.variant_id) || '*';
  var idTailor = textV2_(payload.id_tailor);
  var qtyPerUnit = payload.qty_per_unit !== undefined && payload.qty_per_unit !== ''
    ? Number(payload.qty_per_unit)
    : 1;

  if (!barcodeProduk) {
    return { success: false, message: 'Barcode produk induk wajib diisi.' };
  }

  var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcodeProduk);
  if (!product) {
    return { success: false, message: 'Produk induk "' + barcodeProduk + '" tidak ditemukan di MASTER_PRODUK.' };
  }

  var shMapping = ensureSheetHeadersV2_(ss, IBS.SHEET_MAPPING_POLOS, IBS.HEADERS.MAPPING_POLOS_PRODUK);
  var existingRows = shMapping.getLastRow() >= 2 ? readObjectsV2_(shMapping) : [];
  var targetRow = existingRows.filter(function (r) {
    return upperV2_(r.Barcode_Produk) === upperV2_(barcodeProduk) &&
      (textV2_(r.ID_Variasi) || '*') === idVariasi;
  })[0];

  if (!idTailor) {
    if (targetRow) {
      shMapping.deleteRow(targetRow._row);
      SpreadsheetApp.flush();
      invalidateAppCache_();
      return {
        success: true,
        message: 'Aturan tautan kaos polos berhasil dihapus.',
        data: {
          deleted: true,
          barcode_produk: product.barcode,
          id_variasi: idVariasi,
          polos_mappings: getPolosMappingsList_(ss).items
        }
      };
    }
    return { success: false, message: 'Pilih jenis kaos polos atau aturan yang akan dihapus.' };
  }

  if (!isFinite(qtyPerUnit) || qtyPerUnit < 1 || Math.floor(qtyPerUnit) !== qtyPerUnit) {
    return { success: false, message: 'Kebutuhan polos per unit harus bilangan bulat minimal 1.' };
  }

  var labelVariasi = textV2_(payload.label_variasi);
  if (idVariasi === '*') {
    labelVariasi = 'Semua Variasi (*)';
  } else {
    var variant = findVariantRowV2_(ss.getSheetByName(IBS.SHEET_VARIANTS), idVariasi);
    if (!variant || upperV2_(variant.barcode) !== upperV2_(barcodeProduk)) {
      return { success: false, message: 'Variasi produk tidak valid untuk barcode "' + barcodeProduk + '".' };
    }
    if (!labelVariasi) labelVariasi = variant.label;
  }

  var shMasterTailor = ensureSheetHeadersV2_(ss, IBS.SHEET_MASTER_TAILOR, IBS.HEADERS.MASTER_TAILOR);
  var tailorRec = (shMasterTailor.getLastRow() >= 2 ? readObjectsV2_(shMasterTailor) : []).filter(function (r) {
    return textV2_(r.ID_Tailor) === idTailor;
  })[0];

  var detailPolos = textV2_(payload.detail_polos);
  if (tailorRec) {
    detailPolos = textV2_(tailorRec.Nama_Kategori) + ' (' + textV2_(tailorRec.Barcode_Kategori) + ') · ' +
      (textV2_(tailorRec.Warna) || '-') + ' · ' + (textV2_(tailorRec.Ukuran) || '-');
  } else {
    var gpRec = findGudangPolosRowInStokTailor_(ensureSheetHeadersV2_(ss, IBS.SHEET_STOK_TAILOR, IBS.HEADERS.STOK_TAILOR), idTailor);
    if (!gpRec) {
      return { success: false, message: 'Jenis kaos polos (ID_Tailor: ' + idTailor + ') tidak ditemukan.' };
    }
    if (!detailPolos) {
      detailPolos = gpRec.nama_kategori + ' (' + gpRec.barcode_kategori + ') · ' + gpRec.warna + ' · ' + gpRec.ukuran;
    }
  }

  var now = new Date();
  var mappingId = targetRow
    ? textV2_(targetRow.Mapping_ID)
    : ('MAP-' + Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase());

  if (targetRow) {
    shMapping.getRange(targetRow._row, 2, 1, 7).setValues([[
      product.barcode, idVariasi, labelVariasi, idTailor, detailPolos, qtyPerUnit, now
    ]]);
    shMapping.getRange(targetRow._row, 2, 1, 2).setNumberFormat('@');
    shMapping.getRange(targetRow._row, 5).setNumberFormat('@');
    shMapping.getRange(targetRow._row, 7).setNumberFormat('#,##0');
  } else {
    var newRow = shMapping.getLastRow() + 1;
    shMapping.getRange(newRow, 1, 1, 8).setValues([[
      mappingId, product.barcode, idVariasi, labelVariasi, idTailor, detailPolos, qtyPerUnit, now
    ]]);
    shMapping.getRange(newRow, 1, 1, 3).setNumberFormat('@');
    shMapping.getRange(newRow, 5).setNumberFormat('@');
    shMapping.getRange(newRow, 7).setNumberFormat('#,##0');
  }

  SpreadsheetApp.flush();
  invalidateAppCache_();

  return {
    success: true,
    message: 'Tautan kaos polos berhasil disimpan untuk ' + product.name + ' (' + labelVariasi + ').',
    data: {
      mapping_id: mappingId,
      barcode_produk: product.barcode,
      id_variasi: idVariasi,
      label_variasi: labelVariasi,
      id_tailor: idTailor,
      detail_polos: detailPolos,
      qty_per_unit: qtyPerUnit,
      polos_mappings: getPolosMappingsList_(ss).items
    }
  };
}

// ==========================================
// NORMALISASI ATRIBUT, VARIASI & HELPER TRANSAKSI
// ==========================================
function normalizeAttributesV2_(input) {
  if (!Array.isArray(input)) throw new Error('Format atribut produk tidak valid.');
  if (input.length > 6) throw new Error('Maksimal 6 atribut per produk pada versi ini.');
  var seen = {};
  return input.map(function (a, index) {
    var name = textV2_(a.name);
    var type = textV2_(a.type || 'text').toLowerCase();
    if (!name) throw new Error('Nama atribut ke-' + (index + 1) + ' wajib diisi.');
    if (seen[name.toLowerCase()]) throw new Error('Nama atribut tidak boleh duplikat: ' + name);
    seen[name.toLowerCase()] = true;
    if (['text', 'color'].indexOf(type) === -1) throw new Error('Tipe atribut hanya boleh text atau color.');
    if (!Array.isArray(a.values) || !a.values.length) throw new Error('Atribut ' + name + ' harus memiliki minimal satu nilai.');
    var valueSeen = {};
    var values = a.values.map(function (v) {
      var valText = typeof v === 'string' ? textV2_(v) : textV2_(v && v.value);
      var hex = typeof v === 'object' && v ? textV2_(v.hex) : '';
      if (!valText) throw new Error('Nilai pada atribut ' + name + ' tidak boleh kosong.');
      if (valueSeen[valText.toLowerCase()]) throw new Error('Nilai duplikat pada atribut ' + name + ': ' + valText);
      valueSeen[valText.toLowerCase()] = true;
      return { value: valText, hex: hex };
    });
    return { name: name, type: type, values: values };
  });
}

function normalizeVariantsV2_(input, attributes) {
  if (!Array.isArray(input)) throw new Error('Format variasi tidak valid.');
  if (!attributes || !attributes.length) {
    if (!input.length) return [];
    return [{
      attributes: {},
      active: booleanV2_(input[0].active, true),
      hpp: input[0].hpp,
      harga_jual: input[0].harga_jual
    }];
  }

  var keySeen = {};
  var result = [];
  input.forEach(function (v) {
    var rawAttrs = v && v.attributes && typeof v.attributes === 'object' ? v.attributes : {};
    var cleanAttrs = {};
    attributes.forEach(function (def) {
      var foundKey = Object.keys(rawAttrs).filter(function (k) { return k.toLowerCase() === def.name.toLowerCase(); })[0];
      var val = foundKey ? textV2_(rawAttrs[foundKey]) : '';
      var matchedVal = def.values.filter(function (x) { return x.value.toLowerCase() === val.toLowerCase(); })[0];
      if (!matchedVal) {
        throw new Error('Nilai atribut "' + def.name + '" tidak valid pada kombinasi variasi.');
      }
      cleanAttrs[def.name] = matchedVal.value;
    });
    var key = variantKeyV2_(cleanAttrs, attributes);
    if (!keySeen[key]) {
      keySeen[key] = true;
      result.push({
        attributes: cleanAttrs,
        active: booleanV2_(v.active, true),
        hpp: v.hpp,
        harga_jual: v.harga_jual
      });
    }
  });
  return result;
}

function makeVariantLabelV2_(attrs, attributes) {
  if (!attributes || !attributes.length) return 'Standar';
  var parts = [];
  attributes.forEach(function (a) {
    var val = attrs && Object.prototype.hasOwnProperty.call(attrs, a.name) ? textV2_(attrs[a.name]) : '';
    if (val) parts.push(val);
  });
  return parts.length ? parts.join(' · ') : 'Standar';
}

function variantKeyV2_(attrs, attributes) {
  if (!attributes || !attributes.length) return '__STANDARD__';
  return attributes.map(function (a) {
    var val = attrs && Object.prototype.hasOwnProperty.call(attrs, a.name) ? upperV2_(attrs[a.name]) : '';
    return upperV2_(a.name) + '=' + val;
  }).join('|');
}

function newVariantIdV2_() {
  return 'VAR-' + Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();
}

function trxExistsV2_(logSheet, trxId) {
  if (!logSheet || logSheet.getLastRow() < 2 || !trxId) return false;
  var vals = logSheet.getRange(2, 2, logSheet.getLastRow() - 1, 1).getDisplayValues();
  for (var i = 0; i < vals.length; i++) {
    if (textV2_(vals[i][0]) === textV2_(trxId)) return true;
  }
  return false;
}

function appendRowsV2_(sheet, rows, startsTracker) {
  if (!rows || !rows.length) return 0;
  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  if (Array.isArray(startsTracker)) {
    startsTracker.push({ sheet: sheet, startRow: startRow, numRows: rows.length });
  }
  return startRow;
}

function rollbackAppendsV2_(startsTracker) {
  if (!Array.isArray(startsTracker)) return;
  for (var i = startsTracker.length - 1; i >= 0; i--) {
    var item = startsTracker[i];
    try {
      item.sheet.deleteRows(item.startRow, item.numRows);
    } catch (e) {
      console.error('Rollback append error: ' + e);
    }
  }
}

function parseJsonObjectV2_(raw) {
  var str = textV2_(raw);
  if (!str) return {};
  try {
    var obj = JSON.parse(str);
    return obj && typeof obj === 'object' ? obj : {};
  } catch (e) {
    return {};
  }
}

function jsonResponseV2_(success, message, data) {
  var payload = {
    success: Boolean(success),
    message: message || '',
    data: data === undefined ? null : data
  };
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// MODUL GUDANG PENJAHIT (MASTER_TAILOR, STOK_TAILOR, LOG_TAILOR)
// ==========================================
function getMasterTailorList_(ss, includeInactive) {
  var allItems = getCachedV5_('IBS_MASTER_TAILOR_V6_6', function () {
    var shMaster = ensureSheetHeadersV2_(ss, IBS.SHEET_MASTER_TAILOR, IBS.HEADERS.MASTER_TAILOR);
    if (shMaster.getLastRow() < 2) return [];
    return readObjectsV2_(shMaster).map(function (r) {
      return {
        id_tailor: textV2_(r.ID_Tailor),
        barcode_kategori: textV2_(r.Barcode_Kategori),
        nama_kategori: textV2_(r.Nama_Kategori),
        warna: textV2_(r.Warna) || '-',
        ukuran: textV2_(r.Ukuran) || '-',
        upah_per_pcs: Number(r.Upah_Per_Pcs) || 0,
        status_aktif: booleanV2_(r.Status_Aktif, true)
      };
    }).filter(function (it) {
      return it.id_tailor !== '' || it.barcode_kategori !== '';
    });
  });

  var items = includeInactive
    ? allItems
    : allItems.filter(function (it) { return it.status_aktif !== false; });
  return { items: items };
}

function getStokTailorList_(ss, includeZero, includeGudangPolos) {
  var buildAll = function () {
    var shStok = ensureSheetHeadersV2_(ss, IBS.SHEET_STOK_TAILOR, IBS.HEADERS.STOK_TAILOR);
    if (shStok.getLastRow() < 2) return [];

    var shMaster = ensureSheetHeadersV2_(ss, IBS.SHEET_MASTER_TAILOR, IBS.HEADERS.MASTER_TAILOR);
    var masterMap = {};
    var skuDefaultTarif = {};
    if (shMaster.getLastRow() >= 2) {
      readObjectsV2_(shMaster).forEach(function (m) {
        var idT = textV2_(m.ID_Tailor);
        var bKat = upperV2_(m.Barcode_Kategori);
        var vId = textV2_(m.ID_Variasi_Tujuan);
        var upah = Number(m.Upah_Per_Pcs) || 0;
        masterMap[idT] = {
          upah_per_pcs: upah,
          barcode_produk_tujuan: textV2_(m.Barcode_Produk_Tujuan) || textV2_(m.Barcode_Kategori),
          id_variasi_tujuan: vId
        };
        if (bKat && !vId && booleanV2_(m.Status_Aktif, true)) {
          skuDefaultTarif[bKat] = upah;
        }
      });
    }

    return readObjectsV2_(shStok).map(function (r) {
      var idTailor = textV2_(r.ID_Tailor);
      var barcodeKat = textV2_(r.Barcode_Kategori);
      var stokReady = Number(r.Stok_Ready) || 0;
      var meta = masterMap[idTailor] || { upah_per_pcs: 0, barcode_produk_tujuan: barcodeKat, id_variasi_tujuan: '' };
      var resolvedUpah = meta.upah_per_pcs > 0 ? meta.upah_per_pcs : (skuDefaultTarif[upperV2_(barcodeKat)] || 0);
      return {
        id_tailor: idTailor,
        nama_penjahit: textV2_(r.Nama_Penjahit),
        barcode_kategori: barcodeKat,
        nama_kategori: textV2_(r.Nama_Kategori),
        warna: textV2_(r.Warna) || '-',
        ukuran: textV2_(r.Ukuran) || '-',
        stok_ready: stokReady,
        upah_per_pcs: resolvedUpah,
        estimasi_upah: stokReady * resolvedUpah,
        barcode_produk_tujuan: meta.barcode_produk_tujuan || barcodeKat,
        id_variasi_tujuan: meta.id_variasi_tujuan,
        updated_at: r.Updated_At
      };
    });
  };

  var allItems = getCachedV5_('IBS_STOK_TAILOR_V6_6', buildAll);
  var items = allItems.filter(function (it) {
    if (includeGudangPolos !== true && isGudangPolosName_(it.nama_penjahit)) {
      return false;
    }
    if (!includeZero && it.stok_ready <= 0) {
      return false;
    }
    return true;
  });
  return { items: items };
}

function getTailorSkuAction_(ss, payload) {
  var sku = textV2_(payload.sku || payload.barcode || payload.barcode_kategori);
  if (!sku) {
    return { success: false, message: 'Kode SKU jahit wajib diisi.' };
  }

  var shMaster = ensureSheetHeadersV2_(ss, IBS.SHEET_MASTER_TAILOR, IBS.HEADERS.MASTER_TAILOR);
  var masterRows = shMaster.getLastRow() >= 2 ? readObjectsV2_(shMaster) : [];
  var matchedSku = masterRows.filter(function (r) {
    return upperV2_(r.Barcode_Kategori) === upperV2_(sku);
  });

  if (!matchedSku.length) {
    return {
      success: false,
      message: 'Validasi Gagal: SKU "' + sku + '" tidak ditemukan di Master Tailor.'
    };
  }

  var activeVariants = matchedSku.filter(function (r) {
    return booleanV2_(r.Status_Aktif, true);
  });

  if (!activeVariants.length) {
    return {
      success: false,
      message: 'Validasi Gagal: SKU "' + sku + '" di Master Tailor sedang nonaktif.'
    };
  }

  var shStok = ensureSheetHeadersV2_(ss, IBS.SHEET_STOK_TAILOR, IBS.HEADERS.STOK_TAILOR);
  var stokRows = shStok.getLastRow() >= 2 ? readObjectsV2_(shStok) : [];
  var stokByTailorId = {};
  stokRows.forEach(function (st) {
    var idT = textV2_(st.ID_Tailor);
    var qty = Number(st.Stok_Ready) || 0;
    if (!stokByTailorId[idT]) {
      stokByTailorId[idT] = { total_penjahit: 0, gudang_polos: 0, by_tailor: [] };
    }
    if (isGudangPolosName_(st.Nama_Penjahit)) {
      stokByTailorId[idT].gudang_polos += qty;
    } else {
      stokByTailorId[idT].total_penjahit += qty;
      if (qty > 0) {
        stokByTailorId[idT].by_tailor.push({
          nama_penjahit: textV2_(st.Nama_Penjahit),
          stok_ready: qty
        });
      }
    }
  });

  var variants = activeVariants.map(function (r) {
    var idT = textV2_(r.ID_Tailor);
    var stInfo = stokByTailorId[idT] || { total_penjahit: 0, gudang_polos: 0, by_tailor: [] };
    return {
      id_tailor: idT,
      barcode_kategori: textV2_(r.Barcode_Kategori),
      nama_kategori: textV2_(r.Nama_Kategori),
      warna: textV2_(r.Warna) || '-',
      ukuran: textV2_(r.Ukuran) || '-',
      upah_per_pcs: Number(r.Upah_Per_Pcs) || 0,
      stok_ready_total: stInfo.total_penjahit,
      stok_gudang_polos: stInfo.gudang_polos,
      stok_per_penjahit: stInfo.by_tailor,
      status_aktif: true
    };
  });

  return {
    success: true,
    message: 'SKU valid di Master Tailor: ' + textV2_(activeVariants[0].Nama_Kategori),
    data: {
      barcode_kategori: textV2_(activeVariants[0].Barcode_Kategori),
      nama_kategori: textV2_(activeVariants[0].Nama_Kategori),
      variants: variants
    }
  };
}

function saveMasterTailor_(ss, payload) {
  var barcodeKategori = sanitizeSheetText_(payload.barcode_kategori || payload.barcode);
  var namaKategori = sanitizeSheetText_(payload.nama_kategori || payload.nama);
  var defaultUpah = Number(payload.upah_per_pcs);

  if (!barcodeKategori) {
    return { success: false, message: 'Kode SKU jahit wajib diisi.' };
  }
  if (!namaKategori) {
    return { success: false, message: 'Nama produk jahit wajib diisi.' };
  }

  var itemsToSave = [];
  if (Array.isArray(payload.variants) && payload.variants.length > 0) {
    itemsToSave = payload.variants.map(function (v) {
      return {
        id_tailor: textV2_(v.id_tailor),
        warna: sanitizeSheetText_(v.warna) || '-',
        ukuran: sanitizeSheetText_(v.ukuran) || '-',
        upah_per_pcs: v.upah_per_pcs !== undefined ? Number(v.upah_per_pcs) : defaultUpah,
        status_aktif: booleanV2_(v.status_aktif, booleanV2_(payload.status_aktif, true))
      };
    });
  } else {
    itemsToSave = [{
      id_tailor: textV2_(payload.id_tailor),
      warna: sanitizeSheetText_(payload.warna) || '-',
      ukuran: sanitizeSheetText_(payload.ukuran) || '-',
      upah_per_pcs: defaultUpah,
      status_aktif: booleanV2_(payload.status_aktif, true)
    }];
  }

  for (var i = 0; i < itemsToSave.length; i++) {
    if (!isFinite(itemsToSave[i].upah_per_pcs) || itemsToSave[i].upah_per_pcs < 0 || Math.floor(itemsToSave[i].upah_per_pcs) !== itemsToSave[i].upah_per_pcs) {
      return { success: false, message: 'Tarif upah per pcs harus berupa bilangan bulat >= 0.' };
    }
  }

  var shMaster = ensureSheetHeadersV2_(ss, IBS.SHEET_MASTER_TAILOR, IBS.HEADERS.MASTER_TAILOR);
  var shStok = ensureSheetHeadersV2_(ss, IBS.SHEET_STOK_TAILOR, IBS.HEADERS.STOK_TAILOR);
  var rows = shMaster.getLastRow() >= 2 ? readObjectsV2_(shMaster) : [];
  var stokRows = shStok.getLastRow() >= 2 ? readObjectsV2_(shStok) : [];
  var savedResults = [];

  itemsToSave.forEach(function (item) {
    var targetRow = null;
    if (item.id_tailor) {
      targetRow = rows.filter(function (r) { return textV2_(r.ID_Tailor) === item.id_tailor; })[0];
    }
    if (!targetRow) {
      targetRow = rows.filter(function (r) {
        return upperV2_(r.Barcode_Kategori) === upperV2_(barcodeKategori) &&
          upperV2_(textV2_(r.Warna) || '-') === upperV2_(item.warna) &&
          upperV2_(textV2_(r.Ukuran) || '-') === upperV2_(item.ukuran);
      })[0];
    }

    var finalId = targetRow
      ? textV2_(targetRow.ID_Tailor)
      : ('TLR-' + Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase());

    if (targetRow) {
      shMaster.getRange(targetRow._row, 2, 1, 8).setValues([[
        barcodeKategori, namaKategori, item.warna, item.ukuran, item.upah_per_pcs, '', '', item.status_aktif
      ]]);
      shMaster.getRange(targetRow._row, 6).setNumberFormat('"Rp "#,##0');
      applyRowCheckboxV2_(shMaster, targetRow._row, 1, 9);
    } else {
      var newRow = shMaster.getLastRow() + 1;
      shMaster.getRange(newRow, 1, 1, 9).setValues([[
        finalId, barcodeKategori, namaKategori, item.warna, item.ukuran, item.upah_per_pcs, '', '', item.status_aktif
      ]]);
      shMaster.getRange(newRow, 1, 1, 2).setNumberFormat('@');
      shMaster.getRange(newRow, 6).setNumberFormat('"Rp "#,##0');
      shMaster.getRange(newRow, 7, 1, 2).setNumberFormat('@');
      applyRowCheckboxV2_(shMaster, newRow, 1, 9);
      rows.push({
        _row: newRow,
        ID_Tailor: finalId,
        Barcode_Kategori: barcodeKategori,
        Nama_Kategori: namaKategori,
        Warna: item.warna,
        Ukuran: item.ukuran,
        Upah_Per_Pcs: item.upah_per_pcs,
        Status_Aktif: item.status_aktif
      });
    }

    stokRows.forEach(function (st) {
      if (textV2_(st.ID_Tailor) === finalId) {
        shStok.getRange(st._row, 3, 1, 4).setValues([[barcodeKategori, namaKategori, item.warna, item.ukuran]]);
      }
    });

    savedResults.push({
      id_tailor: finalId,
      barcode_kategori: barcodeKategori,
      nama_kategori: namaKategori,
      warna: item.warna,
      ukuran: item.ukuran,
      upah_per_pcs: item.upah_per_pcs,
      status_aktif: item.status_aktif
    });
  });

  SpreadsheetApp.flush();
  var summaryMsg = savedResults.length > 1
    ? 'Berhasil menyimpan ' + savedResults.length + ' variasi untuk ' + namaKategori + ' (' + barcodeKategori + ').'
    : 'Variasi jahit disimpan: ' + namaKategori + ' (' + savedResults[0].warna + ' · ' + savedResults[0].ukuran + ').';

  return {
    success: true,
    message: summaryMsg,
    data: savedResults.length === 1 ? savedResults[0] : { items: savedResults, count: savedResults.length }
  };
}

function submitStokTailor_(ss, payload) {
  var trxId = textV2_(payload.transaction_id);
  var idTailorInput = textV2_(payload.id_tailor);
  var barcodeSku = textV2_(payload.barcode_kategori || payload.barcode);
  var warnaInput = textV2_(payload.warna);
  var ukuranInput = textV2_(payload.ukuran);
  var namaPenjahit = upperV2_(sanitizeSheetText_(payload.nama_penjahit));
  var qty = Number(payload.qty !== undefined ? payload.qty : payload.qty_bagus);

  if (!namaPenjahit) {
    return { success: false, message: 'Nama penjahit wajib diisi.' };
  }
  if (isGudangPolosName_(namaPenjahit)) {
    return { success: false, message: 'Nama penjahit "GUDANG POLOS" adalah penanda sistem dan tidak dapat digunakan sebagai nama penjahit.' };
  }
  if (!barcodeSku && !idTailorInput) {
    return { success: false, message: 'SKU produk jahit wajib dipindai atau dipilih.' };
  }
  if (!isFinite(qty) || qty < 1 || Math.floor(qty) !== qty) {
    return { success: false, message: 'Jumlah hasil jahitan minimal 1 pcs (bilangan bulat).' };
  }

  var shLogTailor = ensureSheetHeadersV2_(ss, IBS.SHEET_LOG_TAILOR, IBS.HEADERS.LOG_TAILOR);
  if (trxId && shLogTailor.getLastRow() >= 2) {
    var dup = readObjectsV2_(shLogTailor).some(function (r) { return textV2_(r.Log_ID) === trxId; });
    if (dup) {
      return { success: false, message: 'Input hasil jahitan ini sudah pernah tercatat.' };
    }
  }

  var shMasterTailor = ensureSheetHeadersV2_(ss, IBS.SHEET_MASTER_TAILOR, IBS.HEADERS.MASTER_TAILOR);
  var tailorRows = shMasterTailor.getLastRow() >= 2 ? readObjectsV2_(shMasterTailor) : [];

  var skuMatches = barcodeSku
    ? tailorRows.filter(function (r) { return upperV2_(r.Barcode_Kategori) === upperV2_(barcodeSku); })
    : tailorRows.filter(function (r) { return textV2_(r.ID_Tailor) === idTailorInput; });

  if (!skuMatches.length) {
    return {
      success: false,
      message: 'Validasi Gagal: SKU "' + (barcodeSku || idTailorInput) + '" tidak ditemukan di Master Tailor.'
    };
  }

  var activeMatches = skuMatches.filter(function (r) {
    return booleanV2_(r.Status_Aktif, true);
  });

  if (!activeMatches.length) {
    return {
      success: false,
      message: 'Validasi Gagal: SKU "' + (barcodeSku || idTailorInput) + '" di Master Tailor sedang nonaktif.'
    };
  }

  var matchedRecord = null;
  if (idTailorInput) {
    matchedRecord = activeMatches.filter(function (r) { return textV2_(r.ID_Tailor) === idTailorInput; })[0];
  }
  if (!matchedRecord && (warnaInput || ukuranInput)) {
    matchedRecord = activeMatches.filter(function (r) {
      return upperV2_(textV2_(r.Warna) || '-') === upperV2_(warnaInput || '-') &&
        upperV2_(textV2_(r.Ukuran) || '-') === upperV2_(ukuranInput || '-');
    })[0];
  }
  if (!matchedRecord && activeMatches.length === 1) {
    matchedRecord = activeMatches[0];
  }

  if (!matchedRecord) {
    return {
      success: false,
      message: 'Validasi Gagal: Variasi (' + (warnaInput || '-') + ' / ' + (ukuranInput || '-') + ') tidak terdaftar di Master Tailor untuk SKU "' + barcodeSku + '".'
    };
  }

  var resolvedId = textV2_(matchedRecord.ID_Tailor);
  var resolvedBarcode = textV2_(matchedRecord.Barcode_Kategori);
  var namaKategori = textV2_(matchedRecord.Nama_Kategori) || resolvedBarcode;
  var warnaLabel = textV2_(matchedRecord.Warna) || '-';
  var ukuranLabel = textV2_(matchedRecord.Ukuran) || '-';
  var upahSatuan = Number(matchedRecord.Upah_Per_Pcs) || 0;

  var now = new Date();
  var shStok = ensureSheetHeadersV2_(ss, IBS.SHEET_STOK_TAILOR, IBS.HEADERS.STOK_TAILOR);
  var stokRows = shStok.getLastRow() >= 2 ? readObjectsV2_(shStok) : [];
  var existingStok = stokRows.filter(function (r) {
    return textV2_(r.ID_Tailor) === resolvedId && upperV2_(r.Nama_Penjahit) === upperV2_(namaPenjahit);
  })[0];

  var newStokReady = qty;
  if (existingStok) {
    newStokReady = (Number(existingStok.Stok_Ready) || 0) + qty;
    shStok.getRange(existingStok._row, 2).setValue(namaPenjahit);
    shStok.getRange(existingStok._row, 7, 1, 2).setValues([[newStokReady, now]]);
    shStok.getRange(existingStok._row, 7).setNumberFormat('#,##0');
  } else {
    var newRow = shStok.getLastRow() + 1;
    shStok.getRange(newRow, 1, 1, 8).setValues([[
      resolvedId, namaPenjahit, resolvedBarcode, namaKategori, warnaLabel, ukuranLabel, qty, now
    ]]);
    shStok.getRange(newRow, 1).setNumberFormat('@');
    shStok.getRange(newRow, 3).setNumberFormat('@');
    shStok.getRange(newRow, 7).setNumberFormat('#,##0');
  }

  var logId = trxId || ('JHT-' + Utilities.formatDate(now, 'Asia/Jakarta', 'yyMMdd') + '-' + Math.floor(Math.random() * 9000 + 1000));
  var detailItem = namaKategori + ' (' + resolvedBarcode + ') · ' + warnaLabel + ' · ' + ukuranLabel;

  shLogTailor.appendRow([
    logId, now, 'INPUT_JAHIT', namaPenjahit, resolvedId, detailItem,
    qty, 0, upahSatuan, 0, 'PENDING_TRANSFER', ''
  ]);
  shLogTailor.getRange(shLogTailor.getLastRow(), 7, 1, 4).setNumberFormat('#,##0');

  SpreadsheetApp.flush();
  return {
    success: true,
    message: 'Berhasil menambah +' + qty + ' pcs ke stok penjahit (' + detailItem + ').',
    data: {
      log_id: logId,
      id_tailor: resolvedId,
      nama_penjahit: namaPenjahit,
      detail_item: detailItem,
      stok_ready: newStokReady
    }
  };
}

function transferTailorToWarehouse_(ss, payload) {
  var trxId = textV2_(payload.transaction_id);
  var idTailor = textV2_(payload.id_tailor);
  var namaPenjahit = upperV2_(sanitizeSheetText_(payload.nama_penjahit));
  var qtyLolos = Number(payload.qty_lolos) || 0;
  var qtyGagal = Number(payload.qty_gagal) || 0;
  var totalKirim = qtyLolos + qtyGagal;
  var alasanGagal = sanitizeSheetText_(payload.alasan_gagal);
  var namaPenerima = upperV2_(sanitizeSheetText_(payload.nama_penerima || payload.nama_qc)) || 'ADMIN';

  if (!idTailor || !namaPenjahit) {
    return { success: false, message: 'ID kategori penjahit dan nama penjahit wajib diisi.' };
  }
  if (isGudangPolosName_(namaPenjahit)) {
    return { success: false, message: 'Saldo GUDANG POLOS sudah berada di gudang polos.' };
  }
  if (!isFinite(qtyLolos) || !isFinite(qtyGagal) || Math.floor(qtyLolos) !== qtyLolos || Math.floor(qtyGagal) !== qtyGagal) {
    return { success: false, message: 'Jumlah barang lolos dan reject harus berupa bilangan bulat.' };
  }
  if (qtyLolos < 0 || qtyGagal < 0 || totalKirim <= 0) {
    return { success: false, message: 'Jumlah barang diterima minimal 1 pcs dan tidak boleh negatif.' };
  }

  var shLogTailor = ensureSheetHeadersV2_(ss, IBS.SHEET_LOG_TAILOR, IBS.HEADERS.LOG_TAILOR);
  if (trxId && shLogTailor.getLastRow() >= 2) {
    var dupTailorLog = readObjectsV2_(shLogTailor).some(function (r) { return textV2_(r.Log_ID) === trxId; });
    if (dupTailorLog) {
      return { success: false, message: 'Transaksi penerimaan setoran ini sudah tercatat di LOG_TAILOR.' };
    }
  }

  var shStokTailor = ensureSheetHeadersV2_(ss, IBS.SHEET_STOK_TAILOR, IBS.HEADERS.STOK_TAILOR);
  var stokRows = shStokTailor.getLastRow() >= 2 ? readObjectsV2_(shStokTailor) : [];
  var stokRec = stokRows.filter(function (r) {
    return textV2_(r.ID_Tailor) === idTailor && upperV2_(r.Nama_Penjahit) === upperV2_(namaPenjahit);
  })[0];

  if (!stokRec) {
    return { success: false, message: 'Saldo stok penjahit tidak ditemukan.' };
  }
  var currentTailorQty = Number(stokRec.Stok_Ready) || 0;
  if (currentTailorQty < totalKirim) {
    return { success: false, message: 'Stok di penjahit tidak cukup. Tersedia: ' + currentTailorQty + ' pcs.' };
  }

  var shMasterTailor = ensureSheetHeadersV2_(ss, IBS.SHEET_MASTER_TAILOR, IBS.HEADERS.MASTER_TAILOR);
  var allMasterTailor = shMasterTailor.getLastRow() >= 2 ? readObjectsV2_(shMasterTailor) : [];
  var masterRec = allMasterTailor.filter(function (r) {
    return textV2_(r.ID_Tailor) === idTailor;
  })[0];

  var upahSatuan = masterRec ? (Number(masterRec.Upah_Per_Pcs) || 0) : 0;
  if (upahSatuan === 0) {
    var skuFallback = allMasterTailor.filter(function (r) {
      return upperV2_(r.Barcode_Kategori) === upperV2_(stokRec.Barcode_Kategori) &&
        booleanV2_(r.Status_Aktif, true);
    })[0];
    if (skuFallback) {
      upahSatuan = Number(skuFallback.Upah_Per_Pcs) || 0;
    }
  }

  var barcodeKat = textV2_(stokRec.Barcode_Kategori) || (masterRec ? textV2_(masterRec.Barcode_Kategori) : '');
  var namaKat = textV2_(stokRec.Nama_Kategori) || (masterRec ? textV2_(masterRec.Nama_Kategori) : barcodeKat);
  var warnaLabel = textV2_(stokRec.Warna) || (masterRec ? textV2_(masterRec.Warna) : '-');
  var ukuranLabel = textV2_(stokRec.Ukuran) || (masterRec ? textV2_(masterRec.Ukuran) : '-');

  var now = new Date();
  // Audit BE-01: Gunakan findGudangPolosRowInStokTailor_ agar seluruh baris duplikat GUDANG POLOS ikut terkonsolidasi
  var existingGudangPolos = findGudangPolosRowInStokTailor_(shStokTailor, idTailor);

  var polosRollbacks = [];
  var appendStarts = [];
  var oldTailorQty = currentTailorQty;
  var oldTailorUpdated = shStokTailor.getRange(stokRec._row, 8).getValue();

  try {
    var sisaStokTailor = currentTailorQty - totalKirim;
    shStokTailor.getRange(stokRec._row, 7, 1, 2).setValues([[sisaStokTailor, now]]);
    shStokTailor.getRange(stokRec._row, 7).setNumberFormat('#,##0');

    var stokGudangPolosBaru = existingGudangPolos ? existingGudangPolos.stok_ready : 0;
    if (qtyLolos > 0) {
      if (existingGudangPolos) {
        stokGudangPolosBaru = existingGudangPolos.stok_ready + qtyLolos;
        var gpRows = existingGudangPolos.rows || [existingGudangPolos.row];
        for (var idx = 0; idx < gpRows.length; idx++) {
          var targetRowNum = gpRows[idx];
          var oldRowQty = Number(shStokTailor.getRange(targetRowNum, 7).getValue()) || 0;
          var oldRowUpdated = shStokTailor.getRange(targetRowNum, 8).getValue();
          var qtyToWrite = idx === 0 ? stokGudangPolosBaru : 0;
          shStokTailor.getRange(targetRowNum, 2, 1, 7).setValues([[
            IBS.TAILOR_GUDANG_POLOS_NAME, barcodeKat, namaKat, warnaLabel, ukuranLabel, qtyToWrite, now
          ]]);
          shStokTailor.getRange(targetRowNum, 7).setNumberFormat('#,##0');
          polosRollbacks.push({ type: 'update', row: targetRowNum, oldQty: oldRowQty, oldUpdated: oldRowUpdated });
        }
      } else {
        stokGudangPolosBaru = qtyLolos;
        var newPolosRow = shStokTailor.getLastRow() + 1;
        shStokTailor.getRange(newPolosRow, 1, 1, 8).setValues([[
          idTailor, IBS.TAILOR_GUDANG_POLOS_NAME, barcodeKat, namaKat, warnaLabel, ukuranLabel, stokGudangPolosBaru, now
        ]]);
        shStokTailor.getRange(newPolosRow, 1).setNumberFormat('@');
        shStokTailor.getRange(newPolosRow, 3).setNumberFormat('@');
        shStokTailor.getRange(newPolosRow, 7).setNumberFormat('#,##0');
        polosRollbacks.push({ type: 'append', row: newPolosRow });
      }
    }

    var totalUpah = qtyLolos * upahSatuan;
    var statusUpah = totalUpah > 0 ? 'UNPAID' : 'NO_PAY';
    var detailItem = namaKat + ' · ' + warnaLabel + ' · ' + ukuranLabel + ' [Penerima: ' + namaPenerima + ']';
    if (alasanGagal) detailItem += ' [Reject: ' + alasanGagal + ']';
    var logTailorId = trxId || ('TRF-' + Utilities.formatDate(now, 'Asia/Jakarta', 'yyMMdd') + '-' + Math.floor(Math.random() * 9000 + 1000));

    var ltStart = appendRowsV2_(shLogTailor, [[
      logTailorId, now, 'TRANSFER_GUDANG', namaPenjahit, idTailor, detailItem,
      qtyLolos, qtyGagal, upahSatuan, totalUpah, statusUpah, ''
    ]], appendStarts);
    shLogTailor.getRange(ltStart, 7, 1, 4).setNumberFormat('#,##0');

    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Berhasil menerima +' + qtyLolos + ' pcs ke Gudang Polos (Upah tercatat: Rp ' + totalUpah.toLocaleString('id-ID') + ').',
      data: {
        log_id: logTailorId,
        id_tailor: idTailor,
        qty_lolos: qtyLolos,
        qty_gagal: qtyGagal,
        sisa_stok_tailor: sisaStokTailor,
        stok_gudang_polos: stokGudangPolosBaru,
        upah_satuan: upahSatuan,
        total_upah: totalUpah,
        status_upah: statusUpah
      }
    };
  } catch (err) {
    for (var rbIdx = polosRollbacks.length - 1; rbIdx >= 0; rbIdx--) {
      var rbItem = polosRollbacks[rbIdx];
      try {
        if (rbItem.type === 'append') shStokTailor.deleteRow(rbItem.row);
        else shStokTailor.getRange(rbItem.row, 7, 1, 2).setValues([[rbItem.oldQty, rbItem.oldUpdated]]);
      } catch (rbPolosErr) {}
    }
    rollbackAppendsV2_(appendStarts);
    try {
      shStokTailor.getRange(stokRec._row, 7, 1, 2).setValues([[oldTailorQty, oldTailorUpdated]]);
    } catch (rbErr) {}
    throw err;
  }
}

function getUnpaidUpah_(ss) {
  return getCachedV5_('IBS_UNPAID_UPAH_V6_6', function () {
    var shLog = ensureSheetHeadersV2_(ss, IBS.SHEET_LOG_TAILOR, IBS.HEADERS.LOG_TAILOR);
    if (shLog.getLastRow() < 2) {
      return { tailors: [], grand_total_upah: 0, total_penjahit: 0 };
    }

    var grouped = {};
    readObjectsV2_(shLog).forEach(function (r) {
      if (upperV2_(r.Status_Upah) !== 'UNPAID') return;
      var totalUpah = Number(r.Total_Upah) || 0;
      if (totalUpah <= 0) return;

      var nama = upperV2_(r.Nama_Penjahit) || 'TANPA NAMA';
      var key = nama;
      if (!grouped[key]) {
        grouped[key] = {
          nama_penjahit: nama,
          total_qty_bagus: 0,
          total_qty_reject: 0,
          total_upah: 0,
          jumlah_transaksi: 0,
          items: []
        };
      }
      var qtyBagus = Number(r.Qty_Bagus) || 0;
      var qtyReject = Number(r.Qty_Reject) || 0;
      grouped[key].total_qty_bagus += qtyBagus;
      grouped[key].total_qty_reject += qtyReject;
      grouped[key].total_upah += totalUpah;
      grouped[key].jumlah_transaksi += 1;
      grouped[key].items.push({
        log_id: textV2_(r.Log_ID),
        waktu: r.Waktu,
        id_tailor: textV2_(r.ID_Tailor),
        detail_item: textV2_(r.Detail_Item),
        qty_bagus: qtyBagus,
        qty_reject: qtyReject,
        upah_satuan: Number(r.Upah_Satuan) || 0,
        total_upah: totalUpah
      });
    });

    var tailors = Object.keys(grouped).map(function (k) { return grouped[k]; });
    tailors.sort(function (a, b) { return b.total_upah - a.total_upah; });
    var grandTotal = tailors.reduce(function (sum, t) { return sum + t.total_upah; }, 0);

    return {
      tailors: tailors,
      grand_total_upah: grandTotal,
      total_penjahit: tailors.length
    };
  });
}

function generatePaySlip_(ss, payload) {
  var namaPenjahit = upperV2_(sanitizeSheetText_(payload.nama_penjahit));
  var requestedIds = Array.isArray(payload.log_ids) ? payload.log_ids.map(function (id) { return textV2_(id); }) : [];
  var namaAdmin = upperV2_(sanitizeSheetText_(payload.nama_admin || payload.nama_pelapor)) || 'ADMIN';

  if (!namaPenjahit) {
    return { success: false, message: 'Nama penjahit wajib dipilih untuk pencetakan slip upah.' };
  }

  var shLog = ensureSheetHeadersV2_(ss, IBS.SHEET_LOG_TAILOR, IBS.HEADERS.LOG_TAILOR);
  if (shLog.getLastRow() < 2) {
    return { success: false, message: 'Belum ada riwayat upah pada LOG_TAILOR.' };
  }

  var rows = readObjectsV2_(shLog);
  var matched = rows.filter(function (r) {
    if (upperV2_(r.Nama_Penjahit) !== namaPenjahit) return false;
    if (upperV2_(r.Status_Upah) !== 'UNPAID') return false;
    if ((Number(r.Total_Upah) || 0) <= 0) return false;
    if (requestedIds.length > 0 && requestedIds.indexOf(textV2_(r.Log_ID)) === -1) return false;
    return true;
  });

  if (!matched.length) {
    return { success: false, message: 'Tidak ada tagihan upah UNPAID yang dapat dibayarkan untuk ' + namaPenjahit + '.' };
  }

  var now = new Date();
  var slipId = 'SLIP-' + Utilities.formatDate(now, 'Asia/Jakarta', 'yyMMdd') + '-' + Math.floor(Math.random() * 9000 + 1000);
  var totalQtyBagus = 0;
  var totalQtyReject = 0;
  var totalUpah = 0;
  var rincian = [];
  var rowRollbacks = [];
  var appendStarts = [];

  try {
    matched.forEach(function (r) {
      rowRollbacks.push({
        row: r._row,
        oldNama: shLog.getRange(r._row, 4).getValue(),
        oldStatus: shLog.getRange(r._row, 11).getValue(),
        oldSlip: shLog.getRange(r._row, 12).getValue()
      });
      shLog.getRange(r._row, 4).setValue(namaPenjahit);
      shLog.getRange(r._row, 11, 1, 2).setValues([['PAID', slipId]]);
      shLog.getRange(r._row, 12).setNumberFormat('@');
      var qBagus = Number(r.Qty_Bagus) || 0;
      var qReject = Number(r.Qty_Reject) || 0;
      var uTotal = Number(r.Total_Upah) || 0;
      totalQtyBagus += qBagus;
      totalQtyReject += qReject;
      totalUpah += uTotal;
      rincian.push({
        log_id: textV2_(r.Log_ID),
        waktu: r.Waktu,
        detail_item: textV2_(r.Detail_Item),
        qty_bagus: qBagus,
        qty_reject: qReject,
        upah_satuan: Number(r.Upah_Satuan) || 0,
        total_upah: uTotal
      });
    });

    var slipStart = appendRowsV2_(shLog, [[
      slipId, now, 'PEMBAYARAN_SLIP', namaPenjahit, '',
      'Pelunasan ' + matched.length + ' transaksi oleh ' + namaAdmin,
      totalQtyBagus, totalQtyReject, 0, totalUpah, 'PAID', slipId
    ]], appendStarts);
    shLog.getRange(slipStart, 7, 1, 4).setNumberFormat('#,##0');

    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Slip Upah ' + slipId + ' berhasil diterbitkan & ditandai LUNAS (PAID).',
      data: {
        slip_id: slipId,
        nama_penjahit: namaPenjahit,
        tanggal_bayar: now.toISOString(),
        dibayar_oleh: namaAdmin,
        jumlah_item: matched.length,
        total_qty_bagus: totalQtyBagus,
        total_qty_reject: totalQtyReject,
        total_upah: totalUpah,
        rincian: rincian
      }
    };
  } catch (slipErr) {
    for (var rbIdx = rowRollbacks.length - 1; rbIdx >= 0; rbIdx--) {
      var rb = rowRollbacks[rbIdx];
      try {
        shLog.getRange(rb.row, 4).setValue(rb.oldNama);
        shLog.getRange(rb.row, 11, 1, 2).setValues([[rb.oldStatus, rb.oldSlip]]);
      } catch (e) {}
    }
    rollbackAppendsV2_(appendStarts);
    throw slipErr;
  }
}

// ==========================================
// MODUL STOCK OPNAME SESI (GUDANG POLOS)
// ==========================================
function getOpnameSessions_(ss) {
  return getCachedV5_('IBS_OPNAME_SESSIONS_V6_6', function () {
    var sh = ss.getSheetByName(IBS.SHEET_OPNAME_SESI);
    if (!sh || sh.getLastRow() < 2) return { success: true, data: { sessions: [] } };

    var detailStats = {};
    var shDetail = ss.getSheetByName(IBS.SHEET_OPNAME_DETAIL);
    if (shDetail && shDetail.getLastRow() >= 2) {
      readObjectsV2_(shDetail).forEach(function (d) {
        var sid = textV2_(d.Session_ID);
        if (!sid) return;
        if (!detailStats[sid]) {
          detailStats[sid] = { counted_variants: 0, total_counted_qty: 0 };
        }
        detailStats[sid].counted_variants += 1;
        detailStats[sid].total_counted_qty += (Number(d.Qty_Fisik) || 0);
      });
    }

    var rows = readObjectsV2_(sh);
    var activeSessions = rows.filter(function (r) { return upperV2_(r.Status) === 'ACTIVE'; }).map(function (r) {
      var sid = textV2_(r.Session_ID);
      var st = detailStats[sid] || { counted_variants: 0, total_counted_qty: 0 };
      return {
        session_id: sid,
        nama_sesi: textV2_(r.Nama_Sesi),
        lokasi: textV2_(r.Lokasi),
        dibuat_oleh: textV2_(r.Dibuat_Oleh),
        waktu_mulai: r.Waktu_Mulai,
        counted_variants: st.counted_variants,
        total_counted_qty: st.total_counted_qty
      };
    });
    return { success: true, data: { sessions: activeSessions } };
  });
}

function createOpnameSession_(ss, payload) {
  var name = sanitizeSheetText_(payload.nama_sesi);
  var loc = upperV2_(payload.lokasi) || IBS.TAILOR_GUDANG_POLOS_NAME;
  var pelapor = upperV2_(sanitizeSheetText_(payload.nama_pelapor));

  if (!name || !pelapor) return { success: false, message: 'Nama sesi dan pelapor wajib diisi.' };
  if (loc !== IBS.TAILOR_GUDANG_POLOS_NAME && !isValidLocationV2_(ss, loc)) {
    loc = IBS.TAILOR_GUDANG_POLOS_NAME;
  }

  var sessionId = 'OPN-' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyMMdd') + '-' + Math.floor(Math.random() * 9000 + 1000);
  var sh = ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_SESI, IBS.HEADERS.OPNAME_SESI);
  sh.appendRow([sessionId, name, loc, 'ACTIVE', pelapor, new Date(), '']);
  SpreadsheetApp.flush();
  return { success: true, message: 'Sesi opname ' + loc + ' berhasil dibuat.', data: { session_id: sessionId, lokasi: loc } };
}

function submitOpnameCount_(ss, payload) {
  var sessionId = textV2_(payload.session_id);
  var barcode = upperV2_(payload.barcode || payload.sku);
  var targetId = textV2_(payload.id_tailor || payload.variant_id);
  var qty = Number(payload.qty);
  var pelapor = upperV2_(sanitizeSheetText_(payload.nama_pelapor));

  if (!sessionId || !barcode || !targetId || !pelapor || isNaN(qty) || qty < 0) {
    return { success: false, message: 'Data opname tidak lengkap atau tidak valid.' };
  }
  if (!isFinite(qty) || Math.floor(qty) !== qty) {
    return { success: false, message: 'Hasil hitung fisik harus berupa bilangan bulat >= 0.' };
  }

  var shSesi = ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_SESI, IBS.HEADERS.OPNAME_SESI);
  var sesiRows = readObjectsV2_(shSesi);
  var sesi = sesiRows.filter(function (r) { return textV2_(r.Session_ID) === sessionId; })[0];
  if (!sesi) return { success: false, message: 'Sesi opname tidak ditemukan.' };
  assertEditableStatus_(sesi.Status, 'Sesi opname ' + sessionId);

  var isPolosSession = isGudangPolosName_(sesi.Lokasi);
  var snapshotQty = 0;
  var labelVariasi = '';

  if (isPolosSession) {
    var shMasterTailor = ensureSheetHeadersV2_(ss, IBS.SHEET_MASTER_TAILOR, IBS.HEADERS.MASTER_TAILOR);
    var mtRec = (shMasterTailor.getLastRow() >= 2 ? readObjectsV2_(shMasterTailor) : []).filter(function (m) {
      return textV2_(m.ID_Tailor) === targetId;
    })[0];
    if (!mtRec) {
      return { success: false, message: 'Variasi kaos polos tidak ditemukan di Master Tailor.' };
    }
    if (upperV2_(mtRec.Barcode_Kategori) !== barcode) {
      return { success: false, message: 'Variasi tidak cocok dengan kode SKU kaos polos.' };
    }
    if (!booleanV2_(mtRec.Status_Aktif, true)) {
      return { success: false, message: 'Variasi kaos polos sedang nonaktif.' };
    }

    var shStokTailor = ensureSheetHeadersV2_(ss, IBS.SHEET_STOK_TAILOR, IBS.HEADERS.STOK_TAILOR);
    var polosRec = findGudangPolosRowInStokTailor_(shStokTailor, targetId);
    snapshotQty = polosRec ? polosRec.stok_ready : 0;
    labelVariasi = (textV2_(mtRec.Nama_Kategori) || barcode) + ' · ' + (textV2_(mtRec.Warna) || '-') + ' · ' + (textV2_(mtRec.Ukuran) || '-');
  } else {
    var variant = findVariantRowV2_(ss.getSheetByName(IBS.SHEET_VARIANTS), targetId);
    validateOpnameCountInput_(variant, barcode, qty);
    var shStock = ss.getSheetByName(IBS.SHEET_STOCK);
    var stockRow = findStockRowV2_(shStock, targetId, textV2_(sesi.Lokasi));
    snapshotQty = stockRow ? stockRow.qty : 0;
    labelVariasi = variant.label;
  }

  var shDetail = ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_DETAIL, IBS.HEADERS.OPNAME_DETAIL);
  var detailRows = readObjectsV2_(shDetail);
  var existing = detailRows.filter(function (r) { return textV2_(r.Session_ID) === sessionId && textV2_(r.ID_Variasi) === targetId; })[0];

  var now = new Date();
  if (existing) {
    shDetail.getRange(existing._row, 5).setValue(labelVariasi);
    shDetail.getRange(existing._row, 7).setValue(qty);
    shDetail.getRange(existing._row, 8).setValue(pelapor);
    shDetail.getRange(existing._row, 9).setValue(now);
    shDetail.getRange(existing._row, 10).setValue('RECOUNT');
  } else {
    var detailId = sessionId + '-' + targetId;
    shDetail.appendRow([detailId, sessionId, barcode, targetId, labelVariasi, snapshotQty, qty, pelapor, now, 'COUNTED']);
  }

  SpreadsheetApp.flush();
  return {
    success: true,
    message: 'Hasil hitung ' + (isPolosSession ? 'Gudang Polos ' : '') + 'disimpan (Sistem: ' + snapshotQty + ', Fisik: ' + qty + ').',
    data: {
      snapshot_qty: snapshotQty,
      qty_fisik: qty,
      label_variasi: labelVariasi
    }
  };
}

function approveOpnameSession_(ss, payload) {
  var sessionId = textV2_(payload.session_id);
  var pelapor = upperV2_(sanitizeSheetText_(payload.nama_pelapor));
  if (!sessionId || !pelapor) return { success: false, message: 'Session ID dan pelapor wajib diisi.' };

  var shSesi = ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_SESI, IBS.HEADERS.OPNAME_SESI);
  var sesiRows = readObjectsV2_(shSesi);
  var sesi = sesiRows.filter(function (r) { return textV2_(r.Session_ID) === sessionId; })[0];
  if (!sesi) return { success: false, message: 'Sesi tidak ditemukan.' };
  assertEditableStatus_(sesi.Status, 'Sesi opname ' + sessionId);

  var loc = textV2_(sesi.Lokasi);
  var shDetail = ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_DETAIL, IBS.HEADERS.OPNAME_DETAIL);
  var detailRows = readObjectsV2_(shDetail);
  var counts = detailRows.filter(function (r) { return textV2_(r.Session_ID) === sessionId; });

  if (!counts.length) {
    return { success: false, message: 'Belum ada barang yang dihitung pada sesi opname ini.' };
  }

  var now = new Date();

  if (isGudangPolosName_(loc)) {
    var shStokTailor = ensureSheetHeadersV2_(ss, IBS.SHEET_STOK_TAILOR, IBS.HEADERS.STOK_TAILOR);
    var shLogTailor = ensureSheetHeadersV2_(ss, IBS.SHEET_LOG_TAILOR, IBS.HEADERS.LOG_TAILOR);
    var shMasterTailor = ensureSheetHeadersV2_(ss, IBS.SHEET_MASTER_TAILOR, IBS.HEADERS.MASTER_TAILOR);
    var mtMap = {};
    if (shMasterTailor.getLastRow() >= 2) {
      readObjectsV2_(shMasterTailor).forEach(function (m) {
        mtMap[textV2_(m.ID_Tailor)] = m;
      });
    }

    // Audit BE-02: Mekanisme atomic rollback penuh pada cabang GUDANG POLOS
    var polosChanges = [];
    var appendedPolosRows = [];
    var appendLogStarts = [];
    var adjCount = 0;

    try {
      counts.forEach(function (c) {
        var idTailor = textV2_(c.ID_Variasi);
        var sku = textV2_(c.Barcode);
        var qtyFisik = Number(c.Qty_Fisik) || 0;
        var qtySistem = Number(c.Qty_Sistem_Snapshot) || 0;
        var diff = qtyFisik - qtySistem;

        var existingPolos = findGudangPolosRowInStokTailor_(shStokTailor, idTailor);
        if (existingPolos) {
          var rList = existingPolos.rows || [existingPolos.row];
          for (var idx = 0; idx < rList.length; idx++) {
            var targetRowNum = rList[idx];
            var oldQty = Number(shStokTailor.getRange(targetRowNum, 7).getValue()) || 0;
            var oldUpdated = shStokTailor.getRange(targetRowNum, 8).getValue();
            shStokTailor.getRange(targetRowNum, 7, 1, 2).setValues([[idx === 0 ? qtyFisik : 0, now]]);
            shStokTailor.getRange(targetRowNum, 7).setNumberFormat('#,##0');
            polosChanges.push({ row: targetRowNum, oldQty: oldQty, oldUpdated: oldUpdated });
          }
        } else if (qtyFisik > 0) {
          var mtInfo = mtMap[idTailor] || {};
          var newRow = shStokTailor.getLastRow() + 1;
          shStokTailor.getRange(newRow, 1, 1, 8).setValues([[
            idTailor,
            IBS.TAILOR_GUDANG_POLOS_NAME,
            textV2_(mtInfo.Barcode_Kategori) || sku,
            textV2_(mtInfo.Nama_Kategori) || sku,
            textV2_(mtInfo.Warna) || '-',
            textV2_(mtInfo.Ukuran) || '-',
            qtyFisik,
            now
          ]]);
          shStokTailor.getRange(newRow, 1).setNumberFormat('@');
          shStokTailor.getRange(newRow, 3).setNumberFormat('@');
          shStokTailor.getRange(newRow, 7).setNumberFormat('#,##0');
          appendedPolosRows.push(newRow);
        }

        if (diff !== 0) {
          adjCount++;
          var logId = 'OPN-' + Utilities.getUuid().substring(0, 8).toUpperCase();
          var detailText = textV2_(c.Label_Variasi) + ' [Opname: ' + qtySistem + ' -> ' + qtyFisik + ' oleh ' + pelapor + ']';
          var ltStart = appendRowsV2_(shLogTailor, [[
            logId, now, 'OPNAME_POLOS', IBS.TAILOR_GUDANG_POLOS_NAME, idTailor,
            detailText, diff, 0, 0, 0, 'OPNAME_ADJ', sessionId
          ]], appendLogStarts);
          shLogTailor.getRange(ltStart, 7, 1, 4).setNumberFormat('#,##0');
          shLogTailor.getRange(ltStart, 12).setNumberFormat('@');
        }
      });

      shSesi.getRange(sesi._row, 4).setValue('COMPLETED');
      shSesi.getRange(sesi._row, 7).setValue(now);

      SpreadsheetApp.flush();
      return { success: true, message: 'Sesi opname Gudang Polos selesai & dikunci. ' + adjCount + ' penyesuaian stok dicatat.' };
    } catch (polosErr) {
      for (var cIdx = polosChanges.length - 1; cIdx >= 0; cIdx--) {
        var ch = polosChanges[cIdx];
        try {
          shStokTailor.getRange(ch.row, 7, 1, 2).setValues([[ch.oldQty, ch.oldUpdated]]);
        } catch (rbUpdateErr) {}
      }
      appendedPolosRows.sort(function (a, b) { return b - a; });
      appendedPolosRows.forEach(function (rDel) {
        try {
          shStokTailor.deleteRow(rDel);
        } catch (rbDelErr) {}
      });
      rollbackAppendsV2_(appendLogStarts);
      throw polosErr;
    }
  }

  var shStock = ss.getSheetByName(IBS.SHEET_STOCK);
  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  var shVar = ss.getSheetByName(IBS.SHEET_VARIANTS);
  var shMaster = ss.getSheetByName(IBS.SHEET_MASTER);

  var masterMap = {};
  readObjectsV2_(shMaster).forEach(function (m) {
    var bKey = upperV2_(m.Barcode);
    if (bKey) {
      masterMap[bKey] = {
        barcode: textV2_(m.Barcode),
        name: textV2_(m.Nama)
      };
    }
  });

  var variantMap = {};
  readObjectsV2_(shVar).forEach(function (v) {
    var vId = textV2_(v.ID_Variasi);
    if (vId) {
      variantMap[vId] = {
        id: vId,
        hpp: Number(v.HPP) || 0
      };
    }
  });

  var changes = [];
  var logRows = [];

  try {
    counts.forEach(function (c) {
      var variantId = textV2_(c.ID_Variasi);
      var barcode = textV2_(c.Barcode);
      var qtyFisik = Number(c.Qty_Fisik);
      var qtySistem = Number(c.Qty_Sistem_Snapshot);
      var diff = qtyFisik - qtySistem;

      if (diff !== 0) {
        changes.push(writeStockQuantityV2_(shStock, variantId, barcode, loc, qtyFisik, now));
        var trxId = 'OPN-ADJ-' + Utilities.getUuid().substring(0, 8);
        var product = masterMap[upperV2_(barcode)] || null;
        var variant = variantMap[variantId] || null;
        var display = (product ? product.name : barcode) + ' · ' + textV2_(c.Label_Variasi);
        var hppSatuan = variant ? variant.hpp : 0;
        logRows.push([now, trxId, 'Opname', barcode, loc, loc, diff, 'SUCCESS', variantId, product ? product.name : barcode, '', textV2_(c.Label_Variasi), display, pelapor, hppSatuan, 0, 0, 0]);
      }
    });

    if (logRows.length) {
      var start = shLog.getLastRow() + 1;
      shLog.getRange(start, 1, logRows.length, logRows[0].length).setValues(logRows);
      shLog.getRange(start, 15, logRows.length, 4).setNumberFormat('"Rp "#,##0');
    }

    shSesi.getRange(sesi._row, 4).setValue('COMPLETED');
    shSesi.getRange(sesi._row, 7).setValue(now);

    SpreadsheetApp.flush();
    return { success: true, message: 'Sesi opname selesai & dikunci. ' + logRows.length + ' penyesuaian stok dicatat.' };
  } catch (err) {
    rollbackStockChangesV2_(shStock, changes);
    throw err;
  }
}

// ==========================================
// DATA HELPER FUNCTIONS
// ==========================================
function findMasterV2_(sheet, barcode) {
  var rows = readObjectsV2_(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (upperV2_(rows[i].Barcode) === upperV2_(barcode)) {
      return { row: rows[i]._row, barcode: textV2_(rows[i].Barcode), name: textV2_(rows[i].Nama), active: booleanV2_(rows[i].Status_Aktif, false) };
    }
  }
  return null;
}

function findVariantRowV2_(sheet, variantId) {
  var rows = readObjectsV2_(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (textV2_(rows[i].ID_Variasi) === textV2_(variantId)) {
      var hpp = Number(rows[i].HPP) || 0;
      var hargaJual = Number(rows[i].Harga_Jual) || 0;
      return {
        row: rows[i]._row,
        id: textV2_(rows[i].ID_Variasi),
        barcode: textV2_(rows[i].Barcode_Induk),
        label: textV2_(rows[i].Label_Variasi),
        attributes: parseJsonObjectV2_(rows[i].Atribut_JSON),
        active: booleanV2_(rows[i].Status_Aktif, false),
        hpp: hpp,
        harga_jual: hargaJual
      };
    }
  }
  return null;
}

function findStockRowV2_(sheet, variantId, location) {
  var rows = readObjectsV2_(sheet);
  var found = null;
  rows.forEach(function (r) {
    if (textV2_(r.ID_Variasi) === textV2_(variantId) && upperV2_(r.Lokasi_Rak) === upperV2_(location)) {
      if (found) throw new Error('Duplikasi saldo ditemukan untuk ' + variantId + ' / ' + location + '. Periksa STOK_VARIASI.');
      found = { row: r._row, qty: Number(r.Kuantitas) || 0 };
    }
  });
  return found;
}

function writeStockQuantityV2_(sheet, variantId, barcode, location, qty, now) {
  var existing = findStockRowV2_(sheet, variantId, location);
  if (existing) {
    var old = sheet.getRange(existing.row, 4).getValue();
    var oldUpdated = sheet.getRange(existing.row, 5).getValue();
    sheet.getRange(existing.row, 4).setValue(qty);
    sheet.getRange(existing.row, 5).setValue(now);
    return { type: 'update', row: existing.row, oldQty: old, oldUpdated: oldUpdated };
  }
  var row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 5).setValues([[variantId, barcode, location, qty, now]]);
  return { type: 'append', row: row };
}

function rollbackStockChangesV2_(sheet, changes) {
  for (var i = changes.length - 1; i >= 0; i--) {
    var c = changes[i];
    try {
      if (c.type === 'append') sheet.deleteRow(c.row);
      else if (c.type === 'update') { sheet.getRange(c.row, 4).setValue(c.oldQty); sheet.getRange(c.row, 5).setValue(c.oldUpdated); }
    } catch (e) { console.error('Rollback stock error: ' + e); }
  }
}

function readObjectsV2_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  var headers = values[0].map(function (h) { return textV2_(h); });
  return values.slice(1).map(function (row, i) {
    var obj = { _row: i + 2 };
    headers.forEach(function (h, j) { if (h) obj[h] = row[j]; });
    return obj;
  });
}

function textV2_(val) { return val == null ? '' : String(val).trim(); }
function upperV2_(val) { return textV2_(val).toUpperCase(); }
function booleanV2_(val, fallback) {
  if (val === true || val === false) return val;
  var v = upperV2_(val);
  if (v === 'TRUE' || v === '1' || v === 'YA' || v === 'AKTIF') return true;
  if (v === 'FALSE' || v === '0' || v === 'TIDAK' || v === 'NONAKTIF') return false;
  return fallback;
}

/**
 * PENGUJIAN OTOMATIS END-TO-END V6.6 (Opsional - hanya berjalan jika dipanggil manual)
 */
function testSimulasiV6() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAllOperationalSheets_(ss);

  var testSku = 'TEST-SKU-V6';
  var testProdBarcode = 'TEST-PROD-V6';
  var testTailorName = 'PENJAHIT UJI V6';
  var steps = [];

  try {
    cleanupSimulationData_(ss);

    var r1 = saveMasterTailor_(ss, {
      barcode_kategori: testSku,
      nama_kategori: 'Croptop Uji V6',
      variants: [
        { warna: 'Hitam', ukuran: 'XL', upah_per_pcs: 5000, status_aktif: true },
        { warna: 'Putih', ukuran: 'L', upah_per_pcs: 5000, status_aktif: true }
      ]
    });
    if (!r1.success) throw new Error('Step 1 Gagal: ' + r1.message);
    steps.push('1. Master Tailor multi-variasi tersimpan.');

    var r3 = submitStokTailor_(ss, {
      barcode_kategori: testSku,
      warna: 'Hitam',
      ukuran: 'XL',
      nama_penjahit: testTailorName,
      qty: 10
    });
    if (!r3.success || r3.data.stok_ready !== 10) {
      throw new Error('Step 2 Gagal: ' + r3.message);
    }

    var r4 = transferTailorToWarehouse_(ss, {
      id_tailor: r3.data.id_tailor,
      nama_penjahit: testTailorName,
      qty_lolos: 8,
      qty_gagal: 2,
      alasan_gagal: 'Simulasi Reject',
      nama_penerima: 'ADMIN UJI'
    });
    if (!r4.success || r4.data.stok_gudang_polos !== 8) {
      throw new Error('Step 3 Gagal: ' + r4.message);
    }
    steps.push('2. Terima Setoran ke STOK_TAILOR (GUDANG POLOS = 8 pcs) lulus.');

    var rProd = createProductActionV2_(ss, {
      transaction_id: 'TEST-NEW-' + Date.now(),
      barcode: testProdBarcode,
      nama: 'Croptop Desain Uji',
      status_aktif: true,
      default_hpp: 35000,
      default_harga_jual: 85000,
      attributes: [
        { name: 'Warna', type: 'text', values: ['Hitam'] },
        { name: 'Ukuran', type: 'text', values: ['XL'] }
      ],
      variants: [
        { attributes: { Warna: 'Hitam', Ukuran: 'XL' }, active: true }
      ],
      barcode_kategori_tailor: testSku,
      nama_pelapor: 'System Test'
    });
    if (!rProd.success) throw new Error('Setup Produk Induk Gagal: ' + rProd.message);
    var prodData = getProductActionV2_(ss, { barcode: testProdBarcode });
    var varHitamId = prodData.data.variants[0].id;
    steps.push('3. Auto-Match saat CREATE_PRODUCT ke kategori polos lulus.');

    var rBulk = saveBulkKeluarV2_(ss, {
      batch_id: 'TEST-BLK-' + Date.now(),
      nama_pelapor: 'Admin Uji',
      items: [
        { barcode: testProdBarcode, variant_id: varHitamId, qty: 5 }
      ]
    });
    if (!rBulk.success || rBulk.data.total_qty !== 5 || rBulk.data.polos_deducted[0].sisa_stok !== 3) {
      throw new Error('Step 4 Gagal Bulk Keluar: ' + rBulk.message);
    }

    var rOverOrder = saveBulkKeluarV2_(ss, {
      batch_id: 'TEST-BLK-OVER-' + Date.now(),
      nama_pelapor: 'Admin Uji',
      items: [
        { barcode: testProdBarcode, variant_id: varHitamId, qty: 5 }
      ]
    });
    if (rOverOrder.success) {
      throw new Error('Step 5 Gagal: Over-order melebihi sisa Gudang Polos (3 pcs) seharusnya ditolak.');
    }
    steps.push('4. Bulk Keluar 100% Gudang Polos (sisa 3 pcs) & proteksi Over-Order lulus.');

    var r7 = generatePaySlip_(ss, {
      nama_penjahit: testTailorName,
      nama_admin: 'ADMIN UJI'
    });
    if (!r7.success || r7.data.total_upah !== 40000) {
      throw new Error('Step 6 Gagal: ' + r7.message);
    }
    steps.push('5. Penerbitan Slip Gaji (' + r7.data.slip_id + ') & status PAID lulus.');

    var summary = 'SEMUA PENGUJIAN V6.6 LULUS:\n' + steps.join('\n');
    Logger.log(summary);
    return { success: true, message: summary };
  } finally {
    cleanupSimulationData_(ss);
    invalidateAppCache_();
  }
}