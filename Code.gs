/**
 * INA BASIC STORE - Warehouse & Production System V6.2 (Production Ready)
 * Fitur Lengkap:
 * - Gudang Utama: MASTER_PRODUK, ATRIBUT_PRODUK, VARIASI_PRODUK (HPP & Harga Jual), STOK_VARIASI, OPERATION_LOG, MASTER_LOKASI
 * - Stock Opname Sesi: OPNAME_SESI, OPNAME_DETAIL
 * - Gudang Penjahit & Upah Otomatis: MASTER_TAILOR, STOK_TAILOR, LOG_TAILOR
 * - Pembacaan Real-Time Semua Lokasi Aktif dari MASTER_LOKASI (tanpa tertahan cache 2 lokasi)
 * - Fungsi migrasiKeProduksi() untuk migrasi aman ke Spreadsheet Produksi tanpa menghapus data existing
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
  SHEET_LAPORAN_JAHIT: 'LAPORAN_JAHIT',
  SHEET_MASTER_TAILOR: 'MASTER_TAILOR',
  SHEET_STOK_TAILOR: 'STOK_TAILOR',
  SHEET_LOG_TAILOR: 'LOG_TAILOR',
  DEPRECATED_SHEETS: [
    'MASTER_BAHAN',
    'PENGIRIMAN_BAHAN',
    'STOK_BAHAN',
    'LOG_BAHAN',
    'STOK_LOKASI',
    'MIGRASI_ALOKASI',
    'WO_JAHIT',
    'WO_JAHIT_ITEM'
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
    LAPORAN_JAHIT: ['Report_ID', 'Transaction_ID', 'WO_ID', 'Nomor_WO', 'Nama_Penjahit', 'Qty_Bagus', 'Qty_Cacat', 'Status', 'Waktu_Lapor', 'Qty_Terima', 'Qty_Lolos', 'Qty_Gagal', 'Alasan_Gagal', 'Barcode_WIP', 'ID_Variasi_WIP', 'Lokasi_Tujuan', 'Nama_Penerima', 'Waktu_QC'],
    MASTER_TAILOR: ['ID_Tailor', 'Barcode_Kategori', 'Nama_Kategori', 'Warna', 'Ukuran', 'Upah_Per_Pcs', 'Barcode_Produk_Tujuan', 'ID_Variasi_Tujuan', 'Status_Aktif'],
    STOK_TAILOR: ['ID_Tailor', 'Nama_Penjahit', 'Barcode_Kategori', 'Nama_Kategori', 'Warna', 'Ukuran', 'Stok_Ready', 'Updated_At'],
    LOG_TAILOR: ['Log_ID', 'Waktu', 'Jenis_Aktivitas', 'Nama_Penjahit', 'ID_Tailor', 'Detail_Item', 'Qty_Bagus', 'Qty_Reject', 'Upah_Satuan', 'Total_Upah', 'Status_Upah', 'ID_Slip']
  }
};

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
      'IBS_BOOTSTRAP_V5',
      'IBS_LOCATIONS_V5',
      'IBS_MASTER_TAILOR_V6',
      'IBS_STOK_TAILOR_V6',
      'IBS_UNPAID_UPAH_V6',
      'IBS_OPNAME_SESSIONS_V5',
      'IBS_BOOTSTRAP_V6_1',
      'IBS_LOCATIONS_V6_1',
      'IBS_MASTER_TAILOR_V6_1',
      'IBS_STOK_TAILOR_V6_1',
      'IBS_UNPAID_UPAH_V6_1',
      'IBS_OPNAME_SESSIONS_V6_1',
      'IBS_BOOTSTRAP_V6_2',
      'IBS_MASTER_TAILOR_V6_2',
      'IBS_STOK_TAILOR_V6_2',
      'IBS_UNPAID_UPAH_V6_2',
      'IBS_OPNAME_SESSIONS_V6_2'
    ]);
  } catch (e) {}
}

/**
 * Otomatis merapikan baris baru pada MASTER_LOKASI dan membersihkan cache ketika sheet diedit langsung.
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    var sheetName = sheet.getName();
    var watchedSheets = [
      IBS.SHEET_LOCATIONS,
      IBS.SHEET_MASTER,
      IBS.SHEET_ATTRIBUTES,
      IBS.SHEET_VARIANTS,
      IBS.SHEET_STOCK,
      IBS.SHEET_MASTER_TAILOR,
      IBS.SHEET_STOK_TAILOR,
      IBS.SHEET_LOG_TAILOR,
      IBS.SHEET_OPNAME_SESI
    ];

    if (watchedSheets.indexOf(sheetName) !== -1) {
      invalidateAppCache_();
      if (sheetName === IBS.SHEET_LOCATIONS && e.range.getRow() >= 2) {
        var row = e.range.getRow();
        var namaVal = textV2_(sheet.getRange(row, 2).getValue());
        var kodeVal = textV2_(sheet.getRange(row, 1).getValue());
        if (namaVal && !kodeVal) {
          sheet.getRange(row, 1).setValue(upperV2_(namaVal).replace(/\s+/g, '-')).setNumberFormat('@');
        }
        var statusCell = sheet.getRange(row, 4);
        if (namaVal && statusCell.getValue() === '') {
          statusCell.setValue(true);
        }
        if (namaVal) {
          applyRowCheckboxV2_(sheet, row, 1, 4);
          var createdCell = sheet.getRange(row, 5);
          if (createdCell.getValue() === '') {
            createdCell.setValue(new Date());
          }
        }
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
 * Jalankan fungsi ini SEKALI setelah Anda menempelkan kode ke Spreadsheet Produksi yang sudah
 * berisi MASTER_PRODUK, ATRIBUT_PRODUK, VARIASI_PRODUK, STOK_VARIASI, OPERATION_LOG, dan MASTER_LOKASI.
 * - Tidak menghapus atau menimpa data MASTER_PRODUK, VARIASI_PRODUK, STOK_VARIASI, maupun MASTER_LOKASI yang ada.
 * - Menambahkan kolom baru (HPP, Harga_Jual, Margin_Rp, Margin_Persen, Total_Omzet, Total_Profit) secara aman.
 * - Membuat tab MASTER_TAILOR, STOK_TAILOR, LOG_TAILOR, OPNAME_SESI, OPNAME_DETAIL jika belum ada.
 * - Membersihkan sisa baris simulasi testing (TEST-*) jika pernah dijalankan.
 * - Membangun ulang sheet DASHBOARD dengan data produksi Anda.
 */
function migrasiKeProduksi() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Lengkapi seluruh sheet & header operasional tanpa menghapus data lama
  ensureAllOperationalSheets_(ss);

  // 2. Bersihkan sisa baris simulasi testing jika ada
  var cleanedTestCount = cleanupSimulationData_(ss);

  // 3. Rapikan format kolom, formula margin, checkbox, dan hapus tab lama yang tidak terpakai
  var res = setup();

  // 4. Pastikan seluruh baris di MASTER_LOKASI memiliki Kode_Lokasi & checkbox Status_Aktif yang rapi
  var totalLocations = normalizeMasterLokasiSheet_(ss);

  // 5. Bangun ulang Dashboard Eksekutif dengan data produksi asli
  if (typeof buildDashboardSummary === 'function') {
    buildDashboardSummary();
  }

  invalidateAppCache_();

  var summary = 'Migrasi Produksi V6.2 Berhasil! ' +
    'Data MASTER_PRODUK & STOK_VARIASI aman dipertahankan. ' +
    'Total lokasi aktif di MASTER_LOKASI: ' + totalLocations + ' lokasi.' +
    (cleanedTestCount > 0 ? ' Baris simulasi testing dibersihkan: ' + cleanedTestCount + '.' : '');

  Logger.log(summary);
  try {
    ss.toast(summary, 'Migrasi Produksi Selesai', 8);
  } catch (e) {}

  return {
    success: true,
    message: summary,
    active_locations_count: totalLocations,
    cleaned_test_rows: cleanedTestCount,
    deleted_deprecated_sheets: res.deleted_sheets || []
  };
}

/**
 * Dialog interaktif di menu Google Sheets untuk menambah lokasi baru ke MASTER_LOKASI.
 */
function promptTambahLokasiRak() {
  var ui = SpreadsheetApp.getUi();
  var resNama = ui.prompt(
    'Tambah Lokasi Rak Baru',
    'Masukkan Nama Lokasi (contoh: Rak B, Rak Display, Gudang Cabang):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resNama.getSelectedButton() !== ui.Button.OK) return;

  var namaLokasi = textV2_(resNama.getResponseText());
  if (!namaLokasi) {
    ui.alert('Nama lokasi tidak boleh kosong.');
    return;
  }

  var defaultKode = upperV2_(namaLokasi).replace(/[^A-Z0-9]+/g, '-').replace(/^-|-$/g, '');
  var resKode = ui.prompt(
    'Kode Lokasi',
    'Masukkan Kode Lokasi singkat (atau klik OK untuk menggunakan "' + defaultKode + '"):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resKode.getSelectedButton() !== ui.Button.OK) return;

  var kodeLokasi = textV2_(resKode.getResponseText()) || defaultKode;
  var result = tambahLokasiMaster(kodeLokasi, namaLokasi, 'warehouse');
  ui.alert(result.message);
}

/**
 * Menambahkan lokasi baru ke MASTER_LOKASI dengan validasi anti-duplikat dan format otomatis.
 */
function tambahLokasiMaster(kodeLokasi, namaLokasi, jenisLokasi) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var locSheet = ensureSheetHeadersV2_(ss, IBS.SHEET_LOCATIONS, IBS.HEADERS.MASTER_LOKASI);
  var nama = textV2_(namaLokasi);
  var kode = upperV2_(kodeLokasi || nama.replace(/[^a-zA-Z0-9]+/g, '-'));
  var jenis = textV2_(jenisLokasi).toLowerCase() || 'warehouse';

  if (!nama) {
    return { success: false, message: 'Nama_Lokasi wajib diisi.' };
  }

  var existing = locSheet.getLastRow() >= 2 ? readObjectsV2_(locSheet) : [];
  var dup = existing.some(function (r) {
    var existingNama = textV2_(r.Nama_Lokasi || r['Nama Lokasi'] || r.Lokasi);
    return upperV2_(existingNama) === upperV2_(nama);
  });
  if (dup) {
    return { success: false, message: 'Lokasi "' + nama + '" sudah terdaftar di MASTER_LOKASI.' };
  }

  var newRow = locSheet.getLastRow() + 1;
  locSheet.getRange(newRow, 1, 1, 5).setValues([[kode, nama, jenis, true, new Date()]]);
  locSheet.getRange(newRow, 1, 1, 2).setNumberFormat('@');
  applyRowCheckboxV2_(locSheet, newRow, 1, 4);
  invalidateAppCache_();
  SpreadsheetApp.flush();

  return {
    success: true,
    message: 'Lokasi "' + nama + '" berhasil ditambahkan ke MASTER_LOKASI.',
    data: { kode: kode, value: nama, label: nama, kind: jenis }
  };
}

function normalizeMasterLokasiSheet_(ss) {
  var locSheet = ensureSheetHeadersV2_(ss, IBS.SHEET_LOCATIONS, IBS.HEADERS.MASTER_LOKASI);
  var lastRow = locSheet.getLastRow();
  if (lastRow < 2) {
    locSheet.appendRow(['RAK-UTAMA', 'Rak Utama', 'warehouse', true, new Date()]);
    lastRow = 2;
  }

  var vals = locSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var updated = false;
  for (var i = 0; i < vals.length; i++) {
    var nama = textV2_(vals[i][1]);
    if (!nama) continue;
    if (!textV2_(vals[i][0])) {
      vals[i][0] = upperV2_(nama).replace(/[^A-Z0-9]+/g, '-');
      updated = true;
    }
    if (!textV2_(vals[i][2])) {
      vals[i][2] = 'warehouse';
      updated = true;
    }
    if (vals[i][3] === '' || vals[i][3] === null || vals[i][3] === undefined) {
      vals[i][3] = true;
      updated = true;
    }
    if (!vals[i][4]) {
      vals[i][4] = new Date();
      updated = true;
    }
  }
  if (updated) {
    locSheet.getRange(2, 1, vals.length, 5).setValues(vals);
  }
  locSheet.getRange(2, 1, lastRow - 1, 2).setNumberFormat('@');
  cleanAndApplyCheckboxV2_(locSheet, 4);
  return getActiveLocationsV2_(ss).length;
}

function cleanupSimulationData_(ss) {
  var testSku = 'TEST-SKU-V6';
  var testProdBarcode = 'TEST-PROD-V6';
  var testTailorName = 'PENJAHIT UJI V6';
  var deletedCount = 0;

  var targets = [
    { sheet: IBS.SHEET_MASTER_TAILOR, cols: [2] },
    { sheet: IBS.SHEET_STOK_TAILOR, cols: [2, 3] },
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
    IBS.SHEET_OPNAME_DETAIL, IBS.SHEET_LAPORAN_JAHIT,
    IBS.SHEET_MASTER_TAILOR, IBS.SHEET_STOK_TAILOR, IBS.SHEET_LOG_TAILOR
  ];

  sheets.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh) {
      sh.setFrozenRows(1);
      var width = Math.max(sh.getLastColumn(), 1);
      sh.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#d9eaf7');
    }
  });

  // Format kolom teks (@) agar kode/barcode tidak berubah menjadi angka
  ss.getSheetByName(IBS.SHEET_MASTER).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_ATTRIBUTES).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_VARIANTS).getRange('A:B').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_STOCK).getRange('A:B').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_LOG).getRange('B:B').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_LOG).getRange('D:F').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_LOG).getRange('I:I').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_LOCATIONS).getRange('A:B').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_OPNAME_DETAIL).getRange('B:D').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_LAPORAN_JAHIT).getRange('A:D').setNumberFormat('@');
  ss.getSheetByName(IBS.SHEET_LAPORAN_JAHIT).getRange('N:O').setNumberFormat('@');

  // Format kolom teks & angka pada 3 sheet khusus penjahit
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

  var shVar = ss.getSheetByName(IBS.SHEET_VARIANTS);
  repairVariasiProdukSheet_(ss);
  shVar.getRange('F2:G').setNumberFormat('dd/mm/yyyy HH:mm');
  shVar.getRange('H2:J').setNumberFormat('#,##0');
  shVar.getRange('K2:K').setNumberFormat('0.0%');
  syncVariantMarginFormulas_(shVar);

  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  shLog.getRange('O2:R').setNumberFormat('#,##0');

  var locSheet = ss.getSheetByName(IBS.SHEET_LOCATIONS);

  cleanAndApplyCheckboxV2_(ss.getSheetByName(IBS.SHEET_MASTER), 3);
  cleanAndApplyCheckboxV2_(ss.getSheetByName(IBS.SHEET_ATTRIBUTES), 8);
  cleanAndApplyCheckboxV2_(shVar, 5);
  cleanAndApplyCheckboxV2_(locSheet, 4);
  cleanAndApplyCheckboxV2_(shMasterTailor, 9);

  invalidateAppCache_();

  var msg = 'Setup V6.2 (MASTER_TAILOR, STOK_TAILOR, LOG_TAILOR) selesai.';
  if (deletedSheets.length > 0) {
    msg += ' Tab lama dihapus: ' + deletedSheets.join(', ');
  }
  Logger.log(msg);
  try {
    ss.toast(msg, 'Ina Basic Store V6.2', 6);
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
  ensureSheetHeadersV2_(ss, IBS.SHEET_LAPORAN_JAHIT, IBS.HEADERS.LAPORAN_JAHIT);
  ensureSheetHeadersV2_(ss, IBS.SHEET_MASTER_TAILOR, IBS.HEADERS.MASTER_TAILOR);
  ensureSheetHeadersV2_(ss, IBS.SHEET_STOK_TAILOR, IBS.HEADERS.STOK_TAILOR);
  ensureSheetHeadersV2_(ss, IBS.SHEET_LOG_TAILOR, IBS.HEADERS.LOG_TAILOR);
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
  var expected = IBS.HEADERS.VARIASI_PRODUK; // 11 kolom: A..K
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
    IBS.SHEET_OPNAME_DETAIL, IBS.SHEET_LAPORAN_JAHIT,
    IBS.SHEET_MASTER_TAILOR, IBS.SHEET_STOK_TAILOR, IBS.SHEET_LOG_TAILOR
  ];
  var missing = required.some(function (name) { return !ss.getSheetByName(name); });
  if (missing) {
    ensureAllOperationalSheets_(ss);
  }
}

/**
 * Membaca SELURUH lokasi aktif di MASTER_LOKASI secara real-time (tanpa cache)
 * agar setiap penambahan/perubahan baris di MASTER_LOKASI langsung tampil di Web App.
 */
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
  var cachedPart = getCachedV5_('IBS_BOOTSTRAP_V6_2', function () {
    return {
      master_tailor: getMasterTailorList_(ss, false).items,
      stok_tailor: getStokTailorList_(ss, false).items,
      opname_sessions: getOpnameSessions_(ss).data.sessions
    };
  });
  return {
    locations: getActiveLocationsV2_(ss),
    master_tailor: cachedPart.master_tailor,
    stok_tailor: cachedPart.stok_tailor,
    opname_sessions: cachedPart.opname_sessions,
    server_time: new Date().toISOString()
  };
}

function doGet(e) {
  var action = e && e.parameter && e.parameter.action ? upperV2_(e.parameter.action) : '';
  var ss = SpreadsheetApp.getActiveSpreadsheet();

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
  return jsonResponseV2_(true, 'Ina Basic Store Warehouse API V6.2 aktif.', {
    version: '6.2',
    supported_actions: [
      'GET_BOOTSTRAP', 'GET_PRODUCT', 'GET_LOCATIONS', 'ADD_LOCATION', 'GET_OPNAME_SESSIONS',
      'GET_MASTER_TAILOR', 'GET_STOK_TAILOR', 'GET_TAILOR_SKU', 'GET_UNPAID_UPAH',
      'CREATE_PRODUCT', 'ADD_VARIANTS', 'UPDATE_PRODUCT', 'SET_VARIANT_STATUS',
      'UPDATE_VARIANT_PRICING', 'SET_VARIANT_PRICING', 'SET_VARIANT_PRICING_BULK',
      'SAVE_TRANSACTION', 'DEACTIVATE_PRODUCT',
      'CREATE_OPNAME_SESSION', 'SUBMIT_OPNAME_COUNT', 'APPROVE_OPNAME_SESSION',
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

  var action = textV2_(payload.action).toUpperCase();
  if (!action && payload.activity) {
    var legacy = textV2_(payload.activity).toLowerCase();
    if (legacy === 'cekproduk') action = 'GET_PRODUCT';
    else if (legacy === 'editproduk') action = 'UPDATE_PRODUCT';
    else if (legacy === 'hapusproduk') action = 'DEACTIVATE_PRODUCT';
    else return jsonResponseV2_(false, 'Aplikasi lama tidak kompatibel. Perbarui aplikasi ke versi terbaru.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // READ-ONLY ACTIONS (Tanpa LockService agar cepat)
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

  // WRITE / MUTATION ACTIONS (Dengan LockService + Auto Cache Invalidation)
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(IBS.LOCK_TIMEOUT_MS)) return jsonResponseV2_(false, 'Server sedang memproses transaksi lain. Coba lagi.');
  try {
    assertV2SheetsReady_(ss);
    var result;
    switch (action) {
      case 'ADD_LOCATION':
        result = tambahLokasiMaster(payload.kode_lokasi, payload.nama_lokasi, payload.jenis_lokasi);
        break;
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
    locations: getActiveLocationsV2_(ss)
  };
}

function createProductActionV2_(ss, payload) {
  var barcode = textV2_(payload.barcode);
  var name = textV2_(payload.nama);
  var trxId = textV2_(payload.transaction_id);
  var attributes = normalizeAttributesV2_(payload.attributes || []);
  var variants = normalizeVariantsV2_(payload.variants || [], attributes);
  var productActive = booleanV2_(payload.status_aktif, true);
  var defaultHpp = Math.max(0, Number(payload.default_hpp) || 0);
  var defaultHargaJual = Math.max(0, Number(payload.default_harga_jual) || 0);

  if (!trxId) return { success: false, message: 'ID transaksi wajib diisi.' };
  if (!barcode) return { success: false, message: 'Barcode wajib diisi.' };
  if (!name || name.length < 2) return { success: false, message: 'Nama produk minimal 2 karakter.' };
  if (findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode)) return { success: false, message: 'Barcode sudah terdaftar.' };
  if (trxExistsV2_(ss.getSheetByName(IBS.SHEET_LOG), trxId)) return { success: false, message: 'ID transaksi sudah pernah diproses.' };
  if (!variants.length) return { success: false, message: 'Minimal satu kombinasi variasi harus dipilih.' };
  if (!variants.some(function (v) { return v.active; })) return { success: false, message: 'Minimal satu variasi harus berstatus aktif.' };

  var shMaster = ss.getSheetByName(IBS.SHEET_MASTER);
  var shAttrs = ss.getSheetByName(IBS.SHEET_ATTRIBUTES);
  var shVariants = ss.getSheetByName(IBS.SHEET_VARIANTS);
  var shStock = ss.getSheetByName(IBS.SHEET_STOCK);
  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  var activeLocs = getActiveLocationsV2_(ss);
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

    var variantRows = [];
    var stockRows = [];
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
      activeLocs.forEach(function (loc) { stockRows.push([id, barcode, loc.value, 0, now]); });
    });
    var vStart = appendRowsV2_(shVariants, variantRows, starts);
    applyRowCheckboxV2_(shVariants, vStart, variantRows.length, 5);
    shVariants.getRange(vStart, 8, variantRows.length, 3).setNumberFormat('#,##0');
    shVariants.getRange(vStart, 11, variantRows.length, 1).setNumberFormat('0.0%');

    appendRowsV2_(shStock, stockRows, starts);
    appendRowsV2_(shLog, [[now, trxId, 'ProdukBaru', barcode, '', '', 0, 'SUCCESS', '', name, '', '', name, textV2_(payload.nama_pelapor), defaultHpp, defaultHargaJual, 0, 0]], starts);
    SpreadsheetApp.flush();
    return { success: true, message: 'Produk dan ' + variants.length + ' variasi berhasil didaftarkan.', data: { barcode: barcode, nama: name, jumlah_variasi: variants.length } };
  } catch (err) {
    rollbackAppendsV2_(starts);
    throw err;
  }
}

function addVariantsActionV2_(ss, payload) {
  var barcode = textV2_(payload.barcode);
  var trxId = textV2_(payload.transaction_id);
  var defaultHpp = Math.max(0, Number(payload.default_hpp) || 0);
  var defaultHargaJual = Math.max(0, Number(payload.default_harga_jual) || 0);

  if (!barcode || !trxId) return { success: false, message: 'Barcode dan ID transaksi wajib diisi.' };
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
  var shStock = ss.getSheetByName(IBS.SHEET_STOCK);
  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  var activeLocs = getActiveLocationsV2_(ss);
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

    var vr = [], sr = [];
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
      activeLocs.forEach(function (loc) { sr.push([id, barcode, loc.value, 0, now]); });
    });
    var vStart = appendRowsV2_(shVariants, vr, starts);
    applyRowCheckboxV2_(shVariants, vStart, vr.length, 5);
    shVariants.getRange(vStart, 8, vr.length, 3).setNumberFormat('#,##0');
    shVariants.getRange(vStart, 11, vr.length, 1).setNumberFormat('0.0%');

    appendRowsV2_(shStock, sr, starts);
    appendRowsV2_(shLog, [[now, trxId, 'TambahVariasi', barcode, '', '', 0, 'SUCCESS', '', product.name, '', '', product.name, textV2_(payload.nama_pelapor), defaultHpp, defaultHargaJual, 0, 0]], starts);
    SpreadsheetApp.flush();
    return { success: true, message: variants.length + ' variasi baru berhasil ditambahkan.' };
  } catch (err) {
    rollbackAppendsV2_(starts);
    changedVariantRows.forEach(function (x) { try { shVariants.getRange(x.row, 5).setValue(x.oldActive); shVariants.getRange(x.row, 7).setValue(x.oldUpdated); } catch (restoreError) {} });
    throw err;
  }
}

function updateVariantPricingActionV5_(ss, payload) {
  var barcode = textV2_(payload.barcode);
  var items = payload.items || [];
  var pelapor = textV2_(payload.nama_pelapor) || 'Admin';
  var trxId = textV2_(payload.transaction_id) || ('PRICE-' + Utilities.getUuid().substring(0, 8).toUpperCase());

  if (!Array.isArray(items) || !items.length) return { success: false, message: 'Tidak ada data harga/HPP yang dikirim.' };

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
    shVar.getRange(rNum, 8, 1, 3).setNumberFormat('#,##0');
    shVar.getRange(rNum, 11).setNumberFormat('0.0%');
    lastMarginPersen = hargaJual > 0 ? Math.round(((hargaJual - hpp) / hargaJual) * 1000) / 10 : 0;
    updatedCount++;
  });

  if (updatedCount > 0 && product) {
    shLog.appendRow([now, trxId, 'UpdateHargaHPP', barcode, '', '', updatedCount, 'SUCCESS', '', product.name, '', updatedCount + ' variasi', product.name + ' (Update HPP & Harga)', pelapor, 0, 0, 0, 0]);
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
  var name = textV2_(payload.nama);
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
    log.appendRow([now, trxId, 'EditProduk', barcode, '', '', 0, 'SUCCESS', '', name || product.name, '', '', name || product.name, textV2_(payload.nama_pelapor), 0, 0, 0, 0]);
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
    log.appendRow([now, trxId, 'StatusVariasi', barcode, '', '', 0, 'SUCCESS', variantId, product.name, JSON.stringify(row.attributes || {}), row.label, product.name + ' · ' + row.label, textV2_(payload.nama_pelapor), row.hpp, row.harga_jual, 0, 0]);
    SpreadsheetApp.flush();
    return { success: true, message: active ? 'Variasi diaktifkan.' : 'Variasi dinonaktifkan.' };
  } catch (err) {
    try { sh.getRange(row.row, 5).setValue(oldActive); sh.getRange(row.row, 7).setValue(oldUpdated); } catch (restoreError) { console.error('Rollback status variasi: ' + restoreError); }
    throw err;
  }
}

function saveStockTransactionV2_(ss, payload) {
  var trxId = textV2_(payload.transaction_id);
  var action = textV2_(payload.activity);
  var barcode = textV2_(payload.barcode);
  var variantId = textV2_(payload.variant_id);
  var from = textV2_(payload.location_from);
  var to = textV2_(payload.location_to);
  var qty = Number(payload.qty);
  var pelapor = textV2_(payload.nama_pelapor);

  if (!trxId || !barcode || !variantId) return { success: false, message: 'ID transaksi, barcode, dan variasi wajib diisi.' };
  if (!pelapor) return { success: false, message: 'Nama pelapor wajib diisi untuk transaksi ini.' };
  if (action === 'Opname') return { success: false, message: 'Stock opname hanya dapat dilakukan melalui sesi opname.' };
  if (['Masuk', 'Keluar', 'Mutasi'].indexOf(action) === -1) return { success: false, message: 'Aktivitas tidak valid.' };
  if (textV2_(payload.qty) === '' || !isFinite(qty) || qty < 1 || Math.floor(qty) !== qty) return { success: false, message: 'Kuantitas tidak valid.' };
  if ((action === 'Keluar' || action === 'Mutasi') && !isValidLocationV2_(ss, from)) return { success: false, message: 'Lokasi asal tidak valid atau nonaktif.' };
  if ((action === 'Masuk' || action === 'Mutasi') && !isValidLocationV2_(ss, to)) return { success: false, message: 'Lokasi tujuan tidak valid atau nonaktif.' };
  if (action === 'Mutasi' && upperV2_(from) === upperV2_(to)) return { success: false, message: 'Lokasi asal dan tujuan tidak boleh sama.' };

  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  if (trxExistsV2_(shLog, trxId)) return { success: false, message: 'Transaksi ini sudah pernah diproses.' };
  var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode);
  if (!product) return { success: false, message: 'Barcode belum terdaftar.' };
  if (!product.active) return { success: false, message: 'Produk sedang nonaktif.' };
  var variant = findVariantRowV2_(ss.getSheetByName(IBS.SHEET_VARIANTS), variantId);
  if (!variant || upperV2_(variant.barcode) !== upperV2_(barcode)) return { success: false, message: 'Variasi tidak cocok dengan barcode produk.' };
  if (!variant.active) return { success: false, message: 'Variasi sedang nonaktif.' };

  var stockSheet = ss.getSheetByName(IBS.SHEET_STOCK);
  var stockFrom = (action === 'Keluar' || action === 'Mutasi') ? findStockRowV2_(stockSheet, variantId, from) : null;
  var stockTo = (action === 'Masuk' || action === 'Mutasi') ? findStockRowV2_(stockSheet, variantId, to) : null;
  var available = stockFrom ? stockFrom.qty : 0;
  if ((action === 'Keluar' || action === 'Mutasi') && available < qty) {
    return { success: false, message: 'Stok ' + variant.label + ' di ' + from + ' tidak cukup. Tersedia: ' + available };
  }

  var hppSatuan = variant.hpp || 0;
  var hargaJualSatuan = 0;
  var totalOmzet = 0;
  var totalProfit = 0;

  if (action === 'Keluar') {
    hargaJualSatuan = (payload.harga_jual !== undefined && payload.harga_jual !== '' && Number(payload.harga_jual) >= 0)
      ? Number(payload.harga_jual)
      : (variant.harga_jual || 0);
    totalOmzet = qty * hargaJualSatuan;
    totalProfit = hargaJualSatuan > 0 ? (qty * (hargaJualSatuan - hppSatuan)) : 0;
  }

  var changes = [];
  var now = new Date();
  try {
    if (action === 'Keluar' || action === 'Mutasi') {
      changes.push(writeStockQuantityV2_(stockSheet, variantId, barcode, from, available - qty, now));
    }
    if (action === 'Masuk' || action === 'Mutasi') {
      var toQty = stockTo ? stockTo.qty : 0;
      changes.push(writeStockQuantityV2_(stockSheet, variantId, barcode, to, toQty + qty, now));
    }
    var display = product.name + (variant.label ? ' · ' + variant.label : '');
    shLog.appendRow([
      now, trxId, action, barcode, from, to, qty, 'SUCCESS', variantId,
      product.name, JSON.stringify(variant.attributes || {}), variant.label, display, pelapor,
      hppSatuan, hargaJualSatuan, totalOmzet, totalProfit
    ]);
    var lastLogRow = shLog.getLastRow();
    shLog.getRange(lastLogRow, 15, 1, 4).setNumberFormat('#,##0');

    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Transaksi ' + action.toLowerCase() + ' berhasil disimpan.',
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
    rollbackStockChangesV2_(stockSheet, changes);
    throw err;
  }
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
  var allItems = getCachedV5_('IBS_MASTER_TAILOR_V6_2', function () {
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
        barcode_produk_tujuan: textV2_(r.Barcode_Produk_Tujuan),
        id_variasi_tujuan: textV2_(r.ID_Variasi_Tujuan),
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

function getStokTailorList_(ss, includeZero) {
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

  var allItems = getCachedV5_('IBS_STOK_TAILOR_V6_2', buildAll);
  var items = includeZero
    ? allItems
    : allItems.filter(function (it) { return it.stok_ready > 0; });
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
      stokByTailorId[idT] = { total: 0, by_tailor: [] };
    }
    stokByTailorId[idT].total += qty;
    if (qty > 0) {
      stokByTailorId[idT].by_tailor.push({
        nama_penjahit: textV2_(st.Nama_Penjahit),
        stok_ready: qty
      });
    }
  });

  var variants = activeVariants.map(function (r) {
    var idT = textV2_(r.ID_Tailor);
    var stInfo = stokByTailorId[idT] || { total: 0, by_tailor: [] };
    return {
      id_tailor: idT,
      barcode_kategori: textV2_(r.Barcode_Kategori),
      nama_kategori: textV2_(r.Nama_Kategori),
      warna: textV2_(r.Warna) || '-',
      ukuran: textV2_(r.Ukuran) || '-',
      upah_per_pcs: Number(r.Upah_Per_Pcs) || 0,
      stok_ready_total: stInfo.total,
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
  var barcodeKategori = textV2_(payload.barcode_kategori || payload.barcode);
  var namaKategori = textV2_(payload.nama_kategori || payload.nama);
  var defaultUpah = Number(payload.upah_per_pcs);
  var barcodeProdukTujuan = textV2_(payload.barcode_produk_tujuan);
  var idVariasiTujuan = textV2_(payload.id_variasi_tujuan);

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
        warna: textV2_(v.warna) || '-',
        ukuran: textV2_(v.ukuran) || '-',
        upah_per_pcs: v.upah_per_pcs !== undefined ? Number(v.upah_per_pcs) : defaultUpah,
        status_aktif: booleanV2_(v.status_aktif, booleanV2_(payload.status_aktif, true))
      };
    });
  } else {
    itemsToSave = [{
      id_tailor: textV2_(payload.id_tailor),
      warna: textV2_(payload.warna) || '-',
      ukuran: textV2_(payload.ukuran) || '-',
      upah_per_pcs: defaultUpah,
      status_aktif: booleanV2_(payload.status_aktif, true)
    }];
  }

  for (var i = 0; i < itemsToSave.length; i++) {
    if (isNaN(itemsToSave[i].upah_per_pcs) || itemsToSave[i].upah_per_pcs < 0) {
      return { success: false, message: 'Tarif upah per pcs harus berupa angka >= 0.' };
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

    var keepBarcodeTujuan = barcodeProdukTujuan || (targetRow ? textV2_(targetRow.Barcode_Produk_Tujuan) : '');
    var keepVariasiTujuan = idVariasiTujuan || (targetRow ? textV2_(targetRow.ID_Variasi_Tujuan) : '');

    if (targetRow) {
      shMaster.getRange(targetRow._row, 2, 1, 8).setValues([[
        barcodeKategori, namaKategori, item.warna, item.ukuran, item.upah_per_pcs, keepBarcodeTujuan, keepVariasiTujuan, item.status_aktif
      ]]);
      shMaster.getRange(targetRow._row, 6).setNumberFormat('#,##0');
      applyRowCheckboxV2_(shMaster, targetRow._row, 1, 9);
    } else {
      var newRow = shMaster.getLastRow() + 1;
      shMaster.getRange(newRow, 1, 1, 9).setValues([[
        finalId, barcodeKategori, namaKategori, item.warna, item.ukuran, item.upah_per_pcs, keepBarcodeTujuan, keepVariasiTujuan, item.status_aktif
      ]]);
      shMaster.getRange(newRow, 1, 1, 2).setNumberFormat('@');
      shMaster.getRange(newRow, 6).setNumberFormat('#,##0');
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
        Barcode_Produk_Tujuan: keepBarcodeTujuan,
        ID_Variasi_Tujuan: keepVariasiTujuan,
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
  var namaPenjahit = textV2_(payload.nama_penjahit);
  var qty = Number(payload.qty !== undefined ? payload.qty : payload.qty_bagus);

  if (!namaPenjahit) {
    return { success: false, message: 'Nama penjahit wajib diisi.' };
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
  var namaPenjahit = textV2_(payload.nama_penjahit);
  var qtyLolos = Number(payload.qty_lolos) || 0;
  var qtyGagal = Number(payload.qty_gagal) || 0;
  var totalKirim = qtyLolos + qtyGagal;
  var alasanGagal = textV2_(payload.alasan_gagal);
  var namaPenerima = textV2_(payload.nama_penerima || payload.nama_qc) || 'Admin';
  var lokasiTujuan = textV2_(payload.lokasi_tujuan) || 'Rak Utama';

  if (!idTailor || !namaPenjahit) {
    return { success: false, message: 'ID kategori penjahit dan nama penjahit wajib diisi.' };
  }
  if (qtyLolos < 0 || qtyGagal < 0 || totalKirim <= 0) {
    return { success: false, message: 'Jumlah barang dipindahkan minimal 1 pcs dan tidak boleh negatif.' };
  }
  if (!isValidLocationV2_(ss, lokasiTujuan)) {
    return { success: false, message: 'Lokasi tujuan gudang tidak valid atau nonaktif.' };
  }

  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  if (trxId && trxExistsV2_(shLog, trxId)) {
    return { success: false, message: 'Transaksi pemindahan ini sudah pernah diproses.' };
  }

  var shLogTailor = ensureSheetHeadersV2_(ss, IBS.SHEET_LOG_TAILOR, IBS.HEADERS.LOG_TAILOR);
  if (trxId && shLogTailor.getLastRow() >= 2) {
    var dupTailorLog = readObjectsV2_(shLogTailor).some(function (r) { return textV2_(r.Log_ID) === trxId; });
    if (dupTailorLog) {
      return { success: false, message: 'Transaksi pemindahan ini sudah tercatat di LOG_TAILOR.' };
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
        textV2_(r.ID_Variasi_Tujuan) === '' &&
        booleanV2_(r.Status_Aktif, true);
    })[0];
    if (skuFallback) {
      upahSatuan = Number(skuFallback.Upah_Per_Pcs) || 0;
    }
  }

  var barcodeTujuan = textV2_(payload.barcode_produk_tujuan || payload.barcode_wip || payload.barcode) ||
    (masterRec ? textV2_(masterRec.Barcode_Produk_Tujuan) : '') ||
    textV2_(stokRec.Barcode_Kategori);
  var variantId = textV2_(payload.id_variasi_tujuan || payload.id_variasi_wip || payload.variant_id) ||
    (masterRec ? textV2_(masterRec.ID_Variasi_Tujuan) : '');

  var product = null;
  var variant = null;
  if (qtyLolos > 0) {
    if (!barcodeTujuan || !variantId) {
      return { success: false, message: 'Produk dan variasi tujuan di gudang utama wajib dipilih untuk barang siap jual.' };
    }
    product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcodeTujuan);
    if (!product) return { success: false, message: 'Barcode produk tujuan tidak ditemukan di MASTER_PRODUK.' };
    if (!product.active) return { success: false, message: 'Produk tujuan sedang nonaktif.' };

    variant = findVariantRowV2_(ss.getSheetByName(IBS.SHEET_VARIANTS), variantId);
    if (!variant || upperV2_(variant.barcode) !== upperV2_(barcodeTujuan)) {
      return { success: false, message: 'Variasi tidak cocok dengan barcode produk tujuan.' };
    }
    if (!variant.active) return { success: false, message: 'Variasi produk tujuan sedang nonaktif.' };
  }

  var now = new Date();
  var stockSheet = ss.getSheetByName(IBS.SHEET_STOCK);
  var changes = [];
  var oldTailorQty = currentTailorQty;
  var oldTailorUpdated = shStokTailor.getRange(stokRec._row, 8).getValue();

  try {
    var sisaStokTailor = currentTailorQty - totalKirim;
    shStokTailor.getRange(stokRec._row, 7, 1, 2).setValues([[sisaStokTailor, now]]);
    shStokTailor.getRange(stokRec._row, 7).setNumberFormat('#,##0');

    if (qtyLolos > 0) {
      var existingStock = findStockRowV2_(stockSheet, variantId, lokasiTujuan);
      var currentMainQty = existingStock ? existingStock.qty : 0;
      changes.push(writeStockQuantityV2_(stockSheet, variantId, barcodeTujuan, lokasiTujuan, currentMainQty + qtyLolos, now));

      var displayProd = product.name + (variant.label ? ' · ' + variant.label : '');
      var logMainId = trxId || ('TRF-' + Utilities.formatDate(now, 'Asia/Jakarta', 'yyMMdd') + '-' + Math.floor(Math.random() * 9000 + 1000));
      shLog.appendRow([
        now, logMainId, 'Masuk', barcodeTujuan, 'Penjahit: ' + namaPenjahit, lokasiTujuan, qtyLolos, 'SUCCESS', variantId,
        product.name, JSON.stringify(variant.attributes || {}), variant.label, displayProd,
        namaPenerima + ' (Dari Penjahit: ' + namaPenjahit + ')',
        variant.hpp || 0, 0, 0, 0
      ]);
      shLog.getRange(shLog.getLastRow(), 15, 1, 4).setNumberFormat('#,##0');

      if (masterRec && (!textV2_(masterRec.Barcode_Produk_Tujuan) || !textV2_(masterRec.ID_Variasi_Tujuan))) {
        shMasterTailor.getRange(masterRec._row, 7, 1, 2).setValues([[barcodeTujuan, variantId]]);
        shMasterTailor.getRange(masterRec._row, 7, 1, 2).setNumberFormat('@');
      }
    }

    var totalUpah = qtyLolos * upahSatuan;
    var statusUpah = totalUpah > 0 ? 'UNPAID' : 'NO_PAY';
    var detailItem = textV2_(stokRec.Nama_Kategori) + ' · ' + (textV2_(stokRec.Warna) || '-') + ' · ' + (textV2_(stokRec.Ukuran) || '-');
    if (alasanGagal) detailItem += ' [Reject: ' + alasanGagal + ']';
    var logTailorId = trxId || ('TRF-' + Utilities.formatDate(now, 'Asia/Jakarta', 'yyMMdd') + '-' + Math.floor(Math.random() * 9000 + 1000));

    shLogTailor.appendRow([
      logTailorId, now, 'TRANSFER_GUDANG', namaPenjahit, idTailor, detailItem,
      qtyLolos, qtyGagal, upahSatuan, totalUpah, statusUpah, ''
    ]);
    shLogTailor.getRange(shLogTailor.getLastRow(), 7, 1, 4).setNumberFormat('#,##0');

    SpreadsheetApp.flush();
    return {
      success: true,
      message: 'Berhasil memindahkan +' + qtyLolos + ' pcs ke ' + lokasiTujuan + ' (Upah tercatat: Rp ' + totalUpah.toLocaleString('id-ID') + ').',
      data: {
        log_id: logTailorId,
        qty_lolos: qtyLolos,
        qty_gagal: qtyGagal,
        sisa_stok_tailor: sisaStokTailor,
        upah_satuan: upahSatuan,
        total_upah: totalUpah,
        status_upah: statusUpah
      }
    };
  } catch (err) {
    rollbackStockChangesV2_(stockSheet, changes);
    try {
      shStokTailor.getRange(stokRec._row, 7, 1, 2).setValues([[oldTailorQty, oldTailorUpdated]]);
    } catch (rbErr) {}
    throw err;
  }
}

function getUnpaidUpah_(ss) {
  return getCachedV5_('IBS_UNPAID_UPAH_V6_2', function () {
    var shLog = ensureSheetHeadersV2_(ss, IBS.SHEET_LOG_TAILOR, IBS.HEADERS.LOG_TAILOR);
    if (shLog.getLastRow() < 2) {
      return { tailors: [], grand_total_upah: 0, total_penjahit: 0 };
    }

    var grouped = {};
    readObjectsV2_(shLog).forEach(function (r) {
      if (upperV2_(r.Status_Upah) !== 'UNPAID') return;
      var totalUpah = Number(r.Total_Upah) || 0;
      if (totalUpah <= 0) return;

      var nama = textV2_(r.Nama_Penjahit) || 'Tanpa Nama';
      var key = upperV2_(nama);
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
  var namaPenjahit = textV2_(payload.nama_penjahit);
  var requestedIds = Array.isArray(payload.log_ids) ? payload.log_ids.map(function (id) { return textV2_(id); }) : [];
  var namaAdmin = textV2_(payload.nama_admin || payload.nama_pelapor) || 'Admin';

  if (!namaPenjahit) {
    return { success: false, message: 'Nama penjahit wajib dipilih untuk pencetakan slip upah.' };
  }

  var shLog = ensureSheetHeadersV2_(ss, IBS.SHEET_LOG_TAILOR, IBS.HEADERS.LOG_TAILOR);
  if (shLog.getLastRow() < 2) {
    return { success: false, message: 'Belum ada riwayat upah pada LOG_TAILOR.' };
  }

  var rows = readObjectsV2_(shLog);
  var matched = rows.filter(function (r) {
    if (upperV2_(r.Nama_Penjahit) !== upperV2_(namaPenjahit)) return false;
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

  matched.forEach(function (r) {
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

  shLog.appendRow([
    slipId, now, 'PEMBAYARAN_SLIP', matched[0].Nama_Penjahit, '',
    'Pelunasan ' + matched.length + ' transaksi oleh ' + namaAdmin,
    totalQtyBagus, totalQtyReject, 0, totalUpah, 'PAID', slipId
  ]);
  shLog.getRange(shLog.getLastRow(), 7, 1, 4).setNumberFormat('#,##0');

  SpreadsheetApp.flush();
  return {
    success: true,
    message: 'Slip Upah ' + slipId + ' berhasil diterbitkan & ditandai LUNAS (PAID).',
    data: {
      slip_id: slipId,
      nama_penjahit: textV2_(matched[0].Nama_Penjahit),
      tanggal_bayar: now.toISOString(),
      dibayar_oleh: namaAdmin,
      jumlah_item: matched.length,
      total_qty_bagus: totalQtyBagus,
      total_qty_reject: totalQtyReject,
      total_upah: totalUpah,
      rincian: rincian
    }
  };
}

// ==========================================
// MODUL STOCK OPNAME SESI
// ==========================================
function getOpnameSessions_(ss) {
  return getCachedV5_('IBS_OPNAME_SESSIONS_V6_2', function () {
    var sh = ss.getSheetByName(IBS.SHEET_OPNAME_SESI);
    if (!sh || sh.getLastRow() < 2) return { success: true, data: { sessions: [] } };

    var rows = readObjectsV2_(sh);
    var activeSessions = rows.filter(function (r) { return upperV2_(r.Status) === 'ACTIVE'; }).map(function (r) {
      return {
        session_id: textV2_(r.Session_ID),
        nama_sesi: textV2_(r.Nama_Sesi),
        lokasi: textV2_(r.Lokasi),
        dibuat_oleh: textV2_(r.Dibuat_Oleh),
        waktu_mulai: r.Waktu_Mulai
      };
    });
    return { success: true, data: { sessions: activeSessions } };
  });
}

function createOpnameSession_(ss, payload) {
  var name = textV2_(payload.nama_sesi);
  var loc = textV2_(payload.lokasi);
  var pelapor = textV2_(payload.nama_pelapor);

  if (!name || !loc || !pelapor) return { success: false, message: 'Nama sesi, lokasi, dan pelapor wajib diisi.' };
  if (!isValidLocationV2_(ss, loc)) return { success: false, message: 'Lokasi tidak valid.' };

  var sessionId = 'OPN-' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyMMdd') + '-' + Math.floor(Math.random() * 9000 + 1000);
  var sh = ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_SESI, IBS.HEADERS.OPNAME_SESI);
  sh.appendRow([sessionId, name, loc, 'ACTIVE', pelapor, new Date(), '']);
  SpreadsheetApp.flush();
  return { success: true, message: 'Sesi opname berhasil dibuat.', data: { session_id: sessionId } };
}

function submitOpnameCount_(ss, payload) {
  var sessionId = textV2_(payload.session_id);
  var barcode = textV2_(payload.barcode);
  var variantId = textV2_(payload.variant_id);
  var qty = Number(payload.qty);
  var pelapor = textV2_(payload.nama_pelapor);

  if (!sessionId || !barcode || !variantId || !pelapor || isNaN(qty) || qty < 0) return { success: false, message: 'Data tidak lengkap atau tidak valid.' };

  var shSesi = ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_SESI, IBS.HEADERS.OPNAME_SESI);
  var sesiRows = readObjectsV2_(shSesi);
  var sesi = sesiRows.filter(function (r) { return textV2_(r.Session_ID) === sessionId; })[0];
  if (!sesi) return { success: false, message: 'Sesi tidak ditemukan.' };
  assertEditableStatus_(sesi.Status, 'Sesi opname ' + sessionId);

  var variant = findVariantRowV2_(ss.getSheetByName(IBS.SHEET_VARIANTS), variantId);
  if (!variant) return { success: false, message: 'Variasi tidak ditemukan.' };

  var shStock = ss.getSheetByName(IBS.SHEET_STOCK);
  var stockRow = findStockRowV2_(shStock, variantId, textV2_(sesi.Lokasi));
  var snapshotQty = stockRow ? stockRow.qty : 0;

  var shDetail = ensureSheetHeadersV2_(ss, IBS.SHEET_OPNAME_DETAIL, IBS.HEADERS.OPNAME_DETAIL);
  var detailRows = readObjectsV2_(shDetail);
  var existing = detailRows.filter(function (r) { return textV2_(r.Session_ID) === sessionId && textV2_(r.ID_Variasi) === variantId; })[0];

  var now = new Date();
  if (existing) {
    shDetail.getRange(existing._row, 7).setValue(qty);
    shDetail.getRange(existing._row, 8).setValue(pelapor);
    shDetail.getRange(existing._row, 9).setValue(now);
    shDetail.getRange(existing._row, 10).setValue('RECOUNT');
  } else {
    var detailId = sessionId + '-' + variantId;
    shDetail.appendRow([detailId, sessionId, barcode, variantId, variant.label, snapshotQty, qty, pelapor, now, 'COUNTED']);
  }

  SpreadsheetApp.flush();
  return { success: true, message: 'Hasil hitung disimpan (Sistem: ' + snapshotQty + ', Fisik: ' + qty + ').' };
}

function approveOpnameSession_(ss, payload) {
  var sessionId = textV2_(payload.session_id);
  var pelapor = textV2_(payload.nama_pelapor);
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

  var shStock = ss.getSheetByName(IBS.SHEET_STOCK);
  var shLog = ss.getSheetByName(IBS.SHEET_LOG);
  var shVar = ss.getSheetByName(IBS.SHEET_VARIANTS);
  var now = new Date();
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
        var product = findMasterV2_(ss.getSheetByName(IBS.SHEET_MASTER), barcode);
        var variant = findVariantRowV2_(shVar, variantId);
        var display = (product ? product.name : barcode) + ' · ' + textV2_(c.Label_Variasi);
        var hppSatuan = variant ? variant.hpp : 0;
        logRows.push([now, trxId, 'Opname', barcode, loc, loc, diff, 'SUCCESS', variantId, product ? product.name : barcode, '', textV2_(c.Label_Variasi), display, pelapor, hppSatuan, 0, 0, 0]);
      }
    });

    if (logRows.length) {
      var start = shLog.getLastRow() + 1;
      shLog.getRange(start, 1, logRows.length, logRows[0].length).setValues(logRows);
      shLog.getRange(start, 15, logRows.length, 4).setNumberFormat('#,##0');
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
 * PENGUJIAN OTOMATIS END-TO-END (Opsional - hanya berjalan jika dipanggil manual)
 * Seluruh data sementara otomatis dihapus kembali pada blok finally.
 */
function testSimulasiV6() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAllOperationalSheets_(ss);

  var testSku = 'TEST-SKU-V6';
  var testProdBarcode = 'TEST-PROD-V6';
  var testTailorName = 'Penjahit Uji V6';
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

    var r2Valid = getTailorSkuAction_(ss, { sku: testSku });
    if (!r2Valid.success || r2Valid.data.variants.length !== 2) {
      throw new Error('Step 2a Gagal: Validasi SKU valid tidak mengembalikan 2 variasi.');
    }
    var r2Invalid = getTailorSkuAction_(ss, { sku: 'SKU-PALSU-999' });
    if (r2Invalid.success) {
      throw new Error('Step 2b Gagal: SKU tidak terdaftar seharusnya ditolak.');
    }
    steps.push('2. Validasi QR SKU ke Master Tailor (valid & tidak valid) lulus.');

    var r3 = submitStokTailor_(ss, {
      barcode_kategori: testSku,
      warna: 'Hitam',
      ukuran: 'XL',
      nama_penjahit: testTailorName,
      qty: 10
    });
    if (!r3.success || r3.data.stok_ready !== 10) {
      throw new Error('Step 3 Gagal: ' + r3.message);
    }
    steps.push('3. Input hasil jahitan (+10 pcs) ke STOK_TAILOR lulus.');

    var rProd = createProductActionV2_(ss, {
      transaction_id: 'TEST-NEW-' + Date.now(),
      barcode: testProdBarcode,
      nama: 'Croptop Jadi Uji',
      status_aktif: true,
      attributes: [{ name: 'Warna', type: 'text', values: ['Hitam'] }],
      variants: [{ attributes: { Warna: 'Hitam' }, active: true }],
      nama_pelapor: 'System Test'
    });
    if (!rProd.success) throw new Error('Setup Produk Tujuan Gagal: ' + rProd.message);
    var prodData = getProductActionV2_(ss, { barcode: testProdBarcode });
    var targetVarId = prodData.data.variants[0].id;

    var r4 = transferTailorToWarehouse_(ss, {
      id_tailor: r3.data.id_tailor,
      nama_penjahit: testTailorName,
      qty_lolos: 8,
      qty_gagal: 2,
      alasan_gagal: 'Simulasi Reject',
      barcode_produk_tujuan: testProdBarcode,
      id_variasi_tujuan: targetVarId,
      lokasi_tujuan: 'Rak Utama',
      nama_penerima: 'Admin Uji'
    });
    if (!r4.success || r4.data.sisa_stok_tailor !== 0 || r4.data.total_upah !== 40000) {
      throw new Error('Step 4 Gagal: ' + r4.message);
    }
    steps.push('4. Terima ke Gudang Utama & QC (8 lolos, 2 reject, upah Rp 40.000) lulus.');

    invalidateAppCache_();
    var r5 = getUnpaidUpah_(ss);
    var foundTailor = r5.tailors.filter(function (t) { return upperV2_(t.nama_penjahit) === upperV2_(testTailorName); })[0];
    if (!foundTailor || foundTailor.total_upah !== 40000) {
      throw new Error('Step 5 Gagal: Tagihan UNPAID tidak sesuai.');
    }
    steps.push('5. Rekap upah UNPAID (Rp 40.000) terhitung akurat.');

    var r6 = generatePaySlip_(ss, {
      nama_penjahit: testTailorName,
      nama_admin: 'Admin Uji'
    });
    if (!r6.success || r6.data.total_upah !== 40000) {
      throw new Error('Step 6 Gagal: ' + r6.message);
    }
    steps.push('6. Penerbitan Slip Gaji (' + r6.data.slip_id + ') & status PAID lulus.');

    var summary = 'SEMUA 6 PENGUJIAN LULUS:\n' + steps.join('\n');
    Logger.log(summary);
    return { success: true, message: summary };
  } finally {
    cleanupSimulationData_(ss);
    invalidateAppCache_();
  }
}