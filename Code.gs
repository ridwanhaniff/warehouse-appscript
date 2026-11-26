/**
 * ============================================================================
 * INA BASIC STORE · BACKEND & EXECUTIVE DASHBOARD V6.6 (GOOGLE APPS SCRIPT)
 * Arsitektur: 100% Made-to-Order (MTO) & Gudang Polos
 * ============================================================================
 * Fitur Utama:
 * 1. Manajemen Master Produk, Variasi, Harga Modal (HPP) & Harga Jual.
 * 2. Master Barang Jahit & Ongkos Jahit (MASTER_TAILOR) + Saldo WIP Penjahit
 *    & Gudang Polos Siap Pakai (STOK_TAILOR).
 * 3. Auto-Match 1-Klik Produk Desain ke Kategori Kaos Polos (MAPPING_POLOS_PRODUK).
 * 4. Transaksi Barang Keluar Multi-Produk (SAVE_BULK_KELUAR) dengan validasi
 *    atomik & pemotongan saldo GUDANG POLOS secara real-time.
 * 5. Terima Setoran Jahitan (QC) ke Gudang Polos + Pencatatan Upah UNPAID &
 *    Penerbitan Slip Gaji Lunas (GENERATE_PAY_SLIP).
 * 6. Sesi Opname Fisik Khusus Gudang Polos (CREATE / SUBMIT / APPROVE).
 * 7. Sheet DASHBOARD Eksekutif Interaktif + Filter Real-Time & Ringkasan Semua Tab.
 * ============================================================================
 */

const APP_VERSION = 'V6.6';
const DEFAULT_GUDANG_POLOS = 'GUDANG POLOS';
const READ_ONLY_ACTIONS = ['GET_BOOTSTRAP', 'GET_UNPAID_UPAH', 'GET_PRODUCT', 'GET_TAILOR_SKU'];

const SHEETS = {
  MASTER_PRODUK: 'MASTER_PRODUK',
  ATRIBUT_PRODUK: 'ATRIBUT_PRODUK',
  VARIASI_PRODUK: 'VARIASI_PRODUK',
  STOK_VARIASI: 'STOK_VARIASI',
  OPERATION_LOG: 'OPERATION_LOG',
  MASTER_LOKASI: 'MASTER_LOKASI',
  OPNAME_SESI: 'OPNAME_SESI',
  OPNAME_DETAIL: 'OPNAME_DETAIL',
  MASTER_TAILOR: 'MASTER_TAILOR',
  STOK_TAILOR: 'STOK_TAILOR',
  LOG_TAILOR: 'LOG_TAILOR',
  MAPPING_POLOS_PRODUK: 'MAPPING_POLOS_PRODUK',
  DASHBOARD: 'DASHBOARD'
};

const SHEET_HEADERS = {
  MASTER_PRODUK: ['Barcode', 'Nama_Produk', 'Status_Aktif', 'Updated_At'],
  ATRIBUT_PRODUK: ['Barcode', 'Nama_Atribut', 'Tipe_Atribut', 'Daftar_Nilai'],
  VARIASI_PRODUK: ['ID_Variasi', 'Barcode', 'Label_Variasi', 'Atribut_JSON', 'HPP', 'Harga_Jual', 'Status_Aktif', 'Updated_At'],
  STOK_VARIASI: ['ID_Variasi', 'Barcode', 'Lokasi', 'Stok', 'Updated_At'],
  OPERATION_LOG: ['Log_ID', 'Transaction_ID', 'Waktu', 'Tipe', 'Barcode', 'ID_Variasi', 'Detail_Item', 'Lokasi', 'Qty', 'HPP', 'Harga_Jual', 'Omzet', 'Profit', 'Pelapor', 'Keterangan'],
  MASTER_LOKASI: ['Kode_Lokasi', 'Nama_Lokasi', 'Status_Aktif'],
  OPNAME_SESI: ['Session_ID', 'Nama_Sesi', 'Lokasi', 'Status', 'Dibuat_Oleh', 'Waktu_Buat', 'Disetujui_Oleh', 'Waktu_Selesai'],
  OPNAME_DETAIL: ['Detail_ID', 'Session_ID', 'Barcode', 'ID_Tailor', 'Detail_Item', 'Stok_Sistem', 'Stok_Fisik', 'Selisih', 'Pelapor', 'Waktu'],
  MASTER_TAILOR: ['ID_Tailor', 'Barcode_Kategori', 'Nama_Kategori', 'Warna', 'Ukuran', 'Upah_Per_Pcs', 'Status_Aktif', 'Updated_At'],
  STOK_TAILOR: ['ID_Stok_Tailor', 'ID_Tailor', 'Nama_Penjahit', 'Barcode_Kategori', 'Nama_Kategori', 'Warna', 'Ukuran', 'Stok_Ready', 'Updated_At'],
  LOG_TAILOR: ['Log_ID', 'Transaction_ID', 'Waktu', 'Tipe', 'ID_Tailor', 'Nama_Penjahit', 'Detail_Item', 'Qty_Bagus', 'Qty_Rusak', 'Upah_Satuan', 'Total_Upah', 'Status_Upah', 'Slip_ID', 'Pelapor', 'Keterangan'],
  MAPPING_POLOS_PRODUK: ['Mapping_ID', 'Barcode_Produk', 'ID_Variasi', 'Label_Variasi', 'ID_Tailor', 'Detail_Polos', 'Qty_Per_Unit', 'Updated_At']
};

// ============================================================================
// 1. MENU SPREADSHEET & INISIALISASI / MIGRASI AMAN
// ============================================================================

/**
 * Menambahkan menu kustom "Ina Basic Store" saat Google Spreadsheet dibuka
 * sekaligus memperbarui Dashboard jika ada perubahan data.
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Ina Basic Store')
      .addItem('Refresh Dashboard Sekarang', 'refreshDashboard')
      .addItem('Reset Filter Dashboard', 'resetFilterDashboard')
      .addSeparator()
      .addItem('Aktifkan Auto-Refresh Dashboard (5 Menit)', 'aktifkanAutoRefreshDashboard')
      .addItem('Nonaktifkan Auto-Refresh Dashboard', 'nonaktifkanAutoRefreshDashboard')
      .addSeparator()
      .addItem('Migrasi / Sinkronisasi Struktur V6.6', 'migrasiKeProduksi')
      .addToUi();
  } catch (e) {
    // Abaikan jika dijalankan dari trigger non-UI
  }
  try {
    refreshDashboardIfDirty();
  } catch (e) {
    // Abaikan error minor saat buka file
  }
}

/**
 * Trigger otomatis saat pengguna mengubah dropdown filter, tanggal kalender,
 * atau mencentang kotak Reset Filter di sheet DASHBOARD (baris 6).
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEETS.DASHBOARD) return;

    const a1 = e.range.getA1Notation();
    const watched = [
      DASH_FILTER_CELLS.PERIODE,
      DASH_FILTER_CELLS.DATE_FROM,
      DASH_FILTER_CELLS.DATE_TO,
      DASH_FILTER_CELLS.KATEGORI,
      DASH_FILTER_CELLS.PENJAHIT,
      DASH_FILTER_CELLS.RESET_CB
    ];
    if (watched.indexOf(a1) === -1) return;

    // Umpan balik visual instan saat filter diubah
    try {
      sheet.getRange('N2:P3')
        .setValue('⏳  MEMPERBARUI DATA...')
        .setBackground('#78350F')
        .setFontColor('#FEF3C7');
      SpreadsheetApp.flush();
    } catch (errFlush) { }

    if (a1 === DASH_FILTER_CELLS.RESET_CB && parseBool_(e.range.getValue(), false)) {
      resetFilterDashboard();
      return;
    }

    // Jika pengguna memilih tanggal manual di F6 / G6, otomatis ubah mode periode ke CUSTOM TANGGAL
    if (a1 === DASH_FILTER_CELLS.DATE_FROM || a1 === DASH_FILTER_CELLS.DATE_TO) {
      sheet.getRange(DASH_FILTER_CELLS.PERIODE).setValue('CUSTOM TANGGAL');
    }

    refreshDashboard({ editedCell: a1 });
  } catch (err) {
    // Abaikan error minor saat edit manual
  }
}

/**
 * Mengembalikan seluruh filter Dashboard ke kondisi default (SEMUA WAKTU / SEMUA DATA)
 * lalu merender ulang sheet DASHBOARD.
 */
function resetFilterDashboard() {
  return refreshDashboard({ reset: true });
}

function resetDashboardFilters() {
  return resetFilterDashboard();
}

/**
 * Inisialisasi seluruh struktur sheet untuk spreadsheet baru.
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAllSheetsStructure_(ss);
  ensureDefaultLocations_(ss);
  syncAllMasterTailorToGudangPolos_(ss);
  refreshDashboard();
  return 'Setup Ina Basic Store ' + APP_VERSION + ' selesai.';
}

/**
 * Migrasi aman untuk spreadsheet yang sudah berisi data produksi:
 * - Menambahkan sheet/kolom baru yang belum ada tanpa menghapus data lama.
 * - Membersihkan baris sisa pengujian (TEST-SKU-V6, TEST-PROD-V6).
 * - Memastikan setiap varian aktif di MASTER_TAILOR memiliki baris di GUDANG POLOS.
 * - Membangun ulang sheet DASHBOARD eksekutif.
 */
function migrasiKeProduksi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAllSheetsStructure_(ss);
  ensureDefaultLocations_(ss);
  cleanupTestArtifacts_(ss);
  syncAllMasterTailorToGudangPolos_(ss);
  refreshDashboard();
  try {
    SpreadsheetApp.getUi().alert('Migrasi & Sinkronisasi Struktur V6.6 berhasil dijalankan!');
  } catch (e) { }
  return 'Migrasi ke produksi ' + APP_VERSION + ' berhasil.';
}

/**
 * Memastikan seluruh 12 sheet operasional & header kolom tersedia.
 */
function ensureAllSheetsStructure_(ss) {
  Object.keys(SHEET_HEADERS).forEach(function (key) {
    const sheetName = SHEETS[key];
    const headers = SHEET_HEADERS[key];
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      formatHeaderRow_(sheet, headers.length);
    } else if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      formatHeaderRow_(sheet, headers.length);
    } else {
      const existingLastCol = sheet.getLastColumn();
      const existingHeaders = sheet.getRange(1, 1, 1, existingLastCol).getValues()[0].map(function (h) {
        return String(h || '').trim();
      });
      headers.forEach(function (reqH) {
        if (existingHeaders.indexOf(reqH) === -1) {
          const newCol = sheet.getLastColumn() + 1;
          sheet.getRange(1, newCol).setValue(reqH);
          existingHeaders.push(reqH);
        }
      });
      formatHeaderRow_(sheet, sheet.getLastColumn());
    }
  });
}

function formatHeaderRow_(sheet, numCols) {
  const range = sheet.getRange(1, 1, 1, numCols);
  range
    .setBackground('#0f172a')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10);
  sheet.setFrozenRows(1);
}

function ensureDefaultLocations_(ss) {
  const sheet = ss.getSheetByName(SHEETS.MASTER_LOKASI);
  const rows = getSheetObjects_(sheet);
  const existing = {};
  rows.forEach(function (r) {
    existing[normText_(r.Kode_Lokasi)] = true;
  });

  const defaults = [
    ['GUDANG POLOS', 'Gudang Bahan Kaos Polos', true],
    ['GUDANG UTAMA', 'Gudang Utama Operasional', true]
  ];
  defaults.forEach(function (d) {
    if (!existing[normText_(d[0])]) {
      sheet.appendRow(d);
    }
  });
}

function cleanupTestArtifacts_(ss) {
  const testCodes = ['TEST-SKU-V6', 'TEST-PROD-V6'];
  const targets = [
    { sheetName: SHEETS.MASTER_TAILOR, colName: 'Barcode_Kategori' },
    { sheetName: SHEETS.STOK_TAILOR, colName: 'Barcode_Kategori' },
    { sheetName: SHEETS.MASTER_PRODUK, colName: 'Barcode' },
    { sheetName: SHEETS.VARIASI_PRODUK, colName: 'Barcode' },
    { sheetName: SHEETS.MAPPING_POLOS_PRODUK, colName: 'Barcode_Produk' }
  ];

  targets.forEach(function (t) {
    const sh = ss.getSheetByName(t.sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    const data = sh.getDataRange().getValues();
    const headers = data[0].map(function (h) { return String(h || '').trim(); });
    const colIdx = headers.indexOf(t.colName);
    if (colIdx === -1) return;

    for (let r = data.length - 1; r >= 1; r--) {
      const val = normText_(data[r][colIdx]);
      if (testCodes.indexOf(val) !== -1) {
        sh.deleteRow(r + 1);
      }
    }
  });
}

function syncAllMasterTailorToGudangPolos_(ss) {
  const mtSheet = ss.getSheetByName(SHEETS.MASTER_TAILOR);
  const stSheet = ss.getSheetByName(SHEETS.STOK_TAILOR);
  const masters = getSheetObjects_(mtSheet);
  const stoks = getSheetObjects_(stSheet);

  const polosIndex = {};
  stoks.forEach(function (s) {
    if (normText_(s.Nama_Penjahit) === DEFAULT_GUDANG_POLOS) {
      polosIndex[String(s.ID_Tailor || '').trim()] = true;
    }
  });

  const nowStr = formatNow_();
  const rowsToInsert = [];
  masters.forEach(function (m) {
    const idTailor = String(m.ID_Tailor || '').trim();
    if (!idTailor || polosIndex[idTailor]) return;
    rowsToInsert.push([
      generateId_('STK'),
      idTailor,
      DEFAULT_GUDANG_POLOS,
      String(m.Barcode_Kategori || '').trim(),
      String(m.Nama_Kategori || '').trim(),
      String(m.Warna || '-').trim(),
      String(m.Ukuran || '-').trim(),
      0,
      nowStr
    ]);
    polosIndex[idTailor] = true;
  });

  if (rowsToInsert.length > 0) {
    stSheet.getRange(stSheet.getLastRow() + 1, 1, rowsToInsert.length, rowsToInsert[0].length)
      .setValues(rowsToInsert);
    SpreadsheetApp.flush();
  }
}

// ============================================================================
// 2. HTTP ROUTER (doGet & doPost) + KEAMANAN TOKEN & LOCK
// ============================================================================

function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    if (!validateApiToken_(params.token)) {
      return jsonResponse_({ success: false, message: 'Akses ditolak: Token API tidak valid.' });
    }
    const action = normText_(params.action || 'GET_BOOTSTRAP');
    if (READ_ONLY_ACTIONS.indexOf(action) === -1) {
      return jsonResponse_({ success: false, message: 'Metode tidak diizinkan: Gunakan HTTP POST untuk transaksi.' });
    }
    return dispatchAction_(action, params);
  } catch (err) {
    return jsonResponse_({ success: false, message: 'Server Error: ' + err.message });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  let hasLock = false;
  try {
    let body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }
    const params = Object.assign({}, (e && e.parameter) ? e.parameter : {}, body);
    if (!validateApiToken_(params.token)) {
      return jsonResponse_({ success: false, message: 'Akses ditolak: Token API tidak valid.' });
    }

    const action = normText_(params.action || '');
    if (READ_ONLY_ACTIONS.indexOf(action) === -1) {
      hasLock = lock.tryLock(15000);
      if (!hasLock) {
        return jsonResponse_({ success: false, message: 'Server sedang memproses transaksi lain. Silakan coba beberapa detik lagi.' });
      }
    }

    return dispatchAction_(action, params);
  } catch (err) {
    return jsonResponse_({ success: false, message: 'Server Error: ' + err.message });
  } finally {
    if (hasLock) {
      try { lock.releaseLock(); } catch (e) { }
    }
  }
}

function validateApiToken_(providedToken) {
  const props = PropertiesService.getScriptProperties();
  const configuredToken = String(props.getProperty('IBS_API_TOKEN') || '').trim();
  const strictMode = String(props.getProperty('IBS_STRICT_AUTH') || '').trim() === '1';
  if (!configuredToken) {
    return !strictMode;
  }
  return String(providedToken || '').trim() === configuredToken;
}

function dispatchAction_(action, payload) {
  switch (action) {
    case 'GET_BOOTSTRAP':
      return jsonResponse_(handleGetBootstrap_());
    case 'GET_UNPAID_UPAH':
      return jsonResponse_({ success: true, data: buildUnpaidUpahSummary_() });
    case 'GET_PRODUCT':
      return jsonResponse_(handleGetProduct_(payload));
    case 'CREATE_PRODUCT':
      return jsonResponse_(handleCreateProduct_(payload));
    case 'UPDATE_PRODUCT':
      return jsonResponse_(handleUpdateProduct_(payload));
    case 'ADD_VARIANTS':
      return jsonResponse_(handleAddVariants_(payload));
    case 'SET_VARIANT_STATUS':
      return jsonResponse_(handleSetVariantStatus_(payload));
    case 'UPDATE_VARIANT_PRICING':
      return jsonResponse_(handleUpdateVariantPricing_(payload));
    case 'SET_VARIANT_PRICING_BULK':
      return jsonResponse_(handleSetVariantPricingBulk_(payload));
    case 'SAVE_POLOS_MAPPING':
      return jsonResponse_(handleSavePolosMapping_(payload));
    case 'AUTO_MAP_POLOS_CATEGORY':
      return jsonResponse_(handleAutoMapPolosCategory_(payload));
    case 'SAVE_BULK_KELUAR':
      return jsonResponse_(handleSaveBulkKeluar_(payload));
    case 'SAVE_MASTER_TAILOR':
      return jsonResponse_(handleSaveMasterTailor_(payload));
    case 'SUBMIT_STOK_TAILOR':
      return jsonResponse_(handleSubmitStokTailor_(payload));
    case 'TRANSFER_TAILOR_TO_WAREHOUSE':
      return jsonResponse_(handleTransferTailorToWarehouse_(payload));
    case 'GENERATE_PAY_SLIP':
      return jsonResponse_(handleGeneratePaySlip_(payload));
    case 'CREATE_OPNAME_SESSION':
      return jsonResponse_(handleCreateOpnameSession_(payload));
    case 'GET_TAILOR_SKU':
      return jsonResponse_(handleGetTailorSku_(payload));
    case 'SUBMIT_OPNAME_COUNT':
      return jsonResponse_(handleSubmitOpnameCount_(payload));
    case 'APPROVE_OPNAME_SESSION':
      return jsonResponse_(handleApproveOpnameSession_(payload));
    default:
      return jsonResponse_({ success: false, message: 'Action tidak dikenal: ' + action });
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================================
// 3. BOOTSTRAP & SINKRONISASI DATA UTAMA
// ============================================================================

function handleGetBootstrap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAllSheetsStructure_(ss);

  const masterTailorRows = getSheetObjects_(ss.getSheetByName(SHEETS.MASTER_TAILOR));
  const upahByTailorId = {};
  const masterTailor = masterTailorRows.map(function (m) {
    const idT = String(m.ID_Tailor || '').trim();
    const upah = Number(m.Upah_Per_Pcs) || 0;
    const isAct = parseBool_(m.Status_Aktif, true);
    upahByTailorId[idT] = upah;
    return {
      id_tailor: idT,
      barcode_kategori: String(m.Barcode_Kategori || '').trim(),
      nama_kategori: String(m.Nama_Kategori || '').trim(),
      warna: String(m.Warna || '-').trim(),
      ukuran: String(m.Ukuran || '-').trim(),
      upah_per_pcs: upah,
      status_aktif: isAct,
      aktif: isAct,
      updated_at: m.Updated_At || ''
    };
  });

  const stokTailorRows = getSheetObjects_(ss.getSheetByName(SHEETS.STOK_TAILOR));
  const stokTailor = [];
  const gudangPolos = [];

  stokTailorRows.forEach(function (s) {
    const idT = String(s.ID_Tailor || '').trim();
    const penjahit = normText_(s.Nama_Penjahit);
    const qty = Number(s.Stok_Ready) || 0;
    const upahSat = upahByTailorId[idT] || 0;
    const itemObj = {
      id_stok_tailor: String(s.ID_Stok_Tailor || '').trim(),
      id_tailor: idT,
      nama_penjahit: penjahit,
      barcode_kategori: String(s.Barcode_Kategori || '').trim(),
      nama_kategori: String(s.Nama_Kategori || '').trim(),
      warna: String(s.Warna || '-').trim(),
      ukuran: String(s.Ukuran || '-').trim(),
      stok_ready: qty,
      stok_tersedia: qty,
      stok_gudang: qty,
      stok: qty,
      upah_per_pcs: upahSat,
      estimasi_upah: qty * upahSat,
      updated_at: s.Updated_At || ''
    };
    stokTailor.push(itemObj);
    if (penjahit === DEFAULT_GUDANG_POLOS) {
      gudangPolos.push(itemObj);
    }
  });

  const mappingRows = getSheetObjects_(ss.getSheetByName(SHEETS.MAPPING_POLOS_PRODUK));
  const polosMappings = mappingRows.map(function (mp) {
    return {
      mapping_id: String(mp.Mapping_ID || '').trim(),
      barcode_produk: String(mp.Barcode_Produk || '').trim(),
      id_variasi: String(mp.ID_Variasi || '*').trim(),
      label_variasi: String(mp.Label_Variasi || '').trim(),
      id_tailor: String(mp.ID_Tailor || '').trim(),
      detail_polos: String(mp.Detail_Polos || '').trim(),
      qty_per_unit: Math.max(1, Number(mp.Qty_Per_Unit) || 1),
      updated_at: mp.Updated_At || ''
    };
  });

  const opnameSessions = buildActiveOpnameSessions_(ss);
  const unpaidUpah = buildUnpaidUpahSummary_(ss);
  const recentActivities = buildRecentActivities_(ss, 15);

  return {
    success: true,
    data: {
      master_tailor: masterTailor,
      stok_tailor: stokTailor,
      gudang_polos: gudangPolos,
      polos_mappings: polosMappings,
      opname_sessions: opnameSessions,
      unpaid_upah: unpaidUpah,
      unpaid_summary: {
        grand_total_upah: unpaidUpah.grand_total_upah,
        total_penjahit: unpaidUpah.total_penjahit
      },
      recent_activities: recentActivities
    }
  };
}

function buildActiveOpnameSessions_(ss) {
  const sesiRows = getSheetObjects_(ss.getSheetByName(SHEETS.OPNAME_SESI));
  const detailRows = getSheetObjects_(ss.getSheetByName(SHEETS.OPNAME_DETAIL));

  const countBySession = {};
  detailRows.forEach(function (d) {
    const sid = String(d.Session_ID || '').trim();
    if (!sid) return;
    countBySession[sid] = (countBySession[sid] || 0) + 1;
  });

  const activeList = [];
  sesiRows.forEach(function (s) {
    if (normText_(s.Status) !== 'ACTIVE') return;
    const sid = String(s.Session_ID || '').trim();
    activeList.push({
      session_id: sid,
      nama_sesi: String(s.Nama_Sesi || sid).trim(),
      lokasi: String(s.Lokasi || DEFAULT_GUDANG_POLOS).trim(),
      status: 'ACTIVE',
      dibuat_oleh: String(s.Dibuat_Oleh || '-').trim(),
      waktu_buat: s.Waktu_Buat || '',
      counted_variants: countBySession[sid] || 0
    });
  });
  return activeList.reverse();
}

function buildUnpaidUpahSummary_(ssOpt) {
  const ss = ssOpt || SpreadsheetApp.getActiveSpreadsheet();
  const logRows = getSheetObjects_(ss.getSheetByName(SHEETS.LOG_TAILOR));

  const byTailor = {};
  let grandTotal = 0;

  logRows.forEach(function (row) {
    if (normText_(row.Tipe) !== 'SETOR_GUDANG') return;
    if (normText_(row.Status_Upah) !== 'UNPAID') return;
    const totalUpah = Number(row.Total_Upah) || 0;
    const qtyBagus = Number(row.Qty_Bagus) || 0;
    if (totalUpah <= 0 && qtyBagus <= 0) return;

    const namaPenjahit = normText_(row.Nama_Penjahit) || 'TANPA NAMA';
    if (!byTailor[namaPenjahit]) {
      byTailor[namaPenjahit] = {
        nama_penjahit: namaPenjahit,
        jumlah_transaksi: 0,
        total_qty_bagus: 0,
        total_upah: 0,
        items: []
      };
    }

    byTailor[namaPenjahit].jumlah_transaksi += 1;
    byTailor[namaPenjahit].total_qty_bagus += qtyBagus;
    byTailor[namaPenjahit].total_upah += totalUpah;
    byTailor[namaPenjahit].items.push({
      log_id: String(row.Log_ID || '').trim(),
      waktu: row.Waktu || '',
      detail_item: String(row.Detail_Item || '').trim(),
      qty_bagus: qtyBagus,
      upah_satuan: Number(row.Upah_Satuan) || 0,
      total_upah: totalUpah
    });

    grandTotal += totalUpah;
  });

  const tailors = Object.values(byTailor).sort(function (a, b) {
    return b.total_upah - a.total_upah;
  });

  return {
    tailors: tailors,
    grand_total_upah: grandTotal,
    total_penjahit: tailors.length
  };
}

function buildRecentActivities_(ss, limit) {
  const maxItems = limit || 15;
  const opRows = getSheetObjects_(ss.getSheetByName(SHEETS.OPERATION_LOG));
  const tlRows = getSheetObjects_(ss.getSheetByName(SHEETS.LOG_TAILOR));
  const combined = [];

  opRows.slice(-30).forEach(function (r) {
    const tipe = normText_(r.Tipe);
    let category = 'PRODUK';
    let actionText = tipe;
    if (tipe === 'KELUAR') {
      category = 'KELUAR';
      actionText = 'Barang Keluar (' + Math.abs(Number(r.Qty) || 0) + ' pcs)';
    } else if (tipe === 'OPNAME') {
      category = 'OPNAME';
      actionText = 'Opname Gudang Polos';
    } else if (tipe === 'CREATE_PRODUCT') {
      actionText = 'Daftarkan Produk Baru';
    } else if (tipe === 'ADD_VARIANTS') {
      actionText = 'Tambah Variasi Produk';
    } else if (tipe === 'UPDATE_PRICING') {
      actionText = 'Ubah Modal / Harga Jual';
    } else if (tipe === 'UPDATE_PRODUCT' || tipe === 'SET_VARIANT_STATUS') {
      actionText = 'Perbarui Data Produk';
    }

    combined.push({
      category: category,
      actor: String(r.Pelapor || 'ADMIN').trim(),
      action_text: actionText,
      detail_text: String(r.Detail_Item || r.Barcode || '').trim(),
      timestamp: r.Waktu || '',
      _ts: parseDateSafe_(r.Waktu)
    });
  });

  tlRows.slice(-30).forEach(function (r) {
    const tipe = normText_(r.Tipe);
    let category = 'SETOR_GUDANG';
    let actionText = tipe;
    if (tipe === 'SETOR_GUDANG') {
      category = 'SETOR_GUDANG';
      actionText = 'Terima Setoran (' + (Number(r.Qty_Bagus) || 0) + ' pcs)';
    } else if (tipe === 'INPUT_JAHIT') {
      category = 'SETOR_GUDANG';
      actionText = 'Input Hasil Jahit (' + (Number(r.Qty_Bagus) || 0) + ' pcs)';
    } else if (tipe === 'BAYAR_UPAH') {
      category = 'UPAH';
      actionText = 'Bayar Upah (' + formatRp_(r.Total_Upah) + ')';
    }

    combined.push({
      category: category,
      actor: String(r.Pelapor || r.Nama_Penjahit || 'ADMIN').trim(),
      action_text: actionText,
      detail_text: String(r.Detail_Item || r.Nama_Penjahit || '').trim(),
      timestamp: r.Waktu || '',
      _ts: parseDateSafe_(r.Waktu)
    });
  });

  combined.sort(function (a, b) {
    return b._ts - a._ts;
  });

  return combined.slice(0, maxItems).map(function (x) {
    return {
      category: x.category,
      actor: x.actor,
      action_text: x.action_text,
      detail_text: x.detail_text,
      timestamp: x.timestamp
    };
  });
}

// ============================================================================
// 4. MASTER PRODUK, VARIASI, PRICING & AUTO-MATCH GUDANG POLOS
// ============================================================================

function handleGetProduct_(payload) {
  const barcode = normText_(payload.barcode);
  if (!barcode) {
    return { success: false, message: 'Barcode produk wajib diisi.' };
  }
  const includeInactive = parseBool_(payload.include_inactive, false);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prodRows = getSheetObjects_(ss.getSheetByName(SHEETS.MASTER_PRODUK));
  const prod = prodRows.find(function (p) {
    return normText_(p.Barcode) === barcode;
  });

  if (!prod) {
    return { success: false, message: 'Produk dengan barcode "' + barcode + '" tidak ditemukan.' };
  }

  const isProdActive = parseBool_(prod.Status_Aktif, true);
  if (!isProdActive && !includeInactive) {
    return { success: false, message: 'Produk "' + prod.Nama_Produk + '" berstatus NONAKTIF.' };
  }

  const attrRows = getSheetObjects_(ss.getSheetByName(SHEETS.ATRIBUT_PRODUK)).filter(function (a) {
    return normText_(a.Barcode) === barcode;
  });

  const attributeDefinitions = attrRows.map(function (a) {
    const vals = String(a.Daftar_Nilai || '').split(',').map(function (s) {
      return s.trim();
    }).filter(Boolean);
    return {
      name: String(a.Nama_Atribut || '').trim(),
      type: String(a.Tipe_Atribut || 'text').trim(),
      values: vals.map(function (v) { return { value: v }; })
    };
  });

  const varRows = getSheetObjects_(ss.getSheetByName(SHEETS.VARIASI_PRODUK)).filter(function (v) {
    return normText_(v.Barcode) === barcode;
  });

  const variants = [];
  varRows.forEach(function (v) {
    const isVarActive = parseBool_(v.Status_Aktif, true);
    if (!isVarActive && !includeInactive) return;

    let attrsObj = {};
    try {
      if (v.Atribut_JSON) attrsObj = JSON.parse(String(v.Atribut_JSON));
    } catch (e) { }

    const cs = extractColorSizeFromVariant_(attrsObj, v.Label_Variasi);
    const hpp = Number(v.HPP) || 0;
    const hargaJual = Number(v.Harga_Jual) || 0;

    variants.push({
      id: String(v.ID_Variasi || '').trim(),
      id_varian: String(v.ID_Variasi || '').trim(),
      label: String(v.Label_Variasi || 'Standar').trim(),
      warna: cs.warna,
      ukuran: cs.ukuran,
      attributes: attrsObj,
      hpp: hpp,
      harga_modal: hpp,
      harga_jual: hargaJual,
      active: isVarActive
    });
  });

  const mappingRows = getSheetObjects_(ss.getSheetByName(SHEETS.MAPPING_POLOS_PRODUK)).filter(function (mp) {
    return normText_(mp.Barcode_Produk) === barcode;
  }).map(function (mp) {
    return {
      mapping_id: String(mp.Mapping_ID || '').trim(),
      barcode_produk: String(mp.Barcode_Produk || '').trim(),
      id_variasi: String(mp.ID_Variasi || '*').trim(),
      label_variasi: String(mp.Label_Variasi || '').trim(),
      id_tailor: String(mp.ID_Tailor || '').trim(),
      detail_polos: String(mp.Detail_Polos || '').trim(),
      qty_per_unit: Math.max(1, Number(mp.Qty_Per_Unit) || 1),
      updated_at: mp.Updated_At || ''
    };
  });

  return {
    success: true,
    data: {
      barcode: String(prod.Barcode || '').trim(),
      nama: String(prod.Nama_Produk || '').trim(),
      nama_produk: String(prod.Nama_Produk || '').trim(),
      aktif: isProdActive,
      attributeDefinitions: attributeDefinitions,
      attributes: attributeDefinitions,
      variants: variants,
      polos_mappings: mappingRows
    }
  };
}

function handleCreateProduct_(payload) {
  const barcode = normText_(payload.barcode);
  const nama = String(payload.nama || '').trim();
  const pelapor = normText_(payload.nama_pelapor) || 'ADMIN';
  const trxId = String(payload.transaction_id || generateId_('NEW')).trim();

  if (!barcode || nama.length < 2) {
    return { success: false, message: 'Barcode dan Nama Produk wajib diisi.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (isTransactionLogged_(ss, SHEETS.OPERATION_LOG, trxId)) {
    return { success: true, message: 'Produk sudah tercatat sebelumnya.' };
  }

  const mpSheet = ss.getSheetByName(SHEETS.MASTER_PRODUK);
  const existing = getSheetObjects_(mpSheet).find(function (p) {
    return normText_(p.Barcode) === barcode;
  });
  if (existing) {
    return { success: false, message: 'Barcode "' + barcode + '" sudah terdaftar.' };
  }

  const nowStr = formatNow_();
  const statusAktif = parseBool_(payload.status_aktif, true);
  mpSheet.appendRow([barcode, nama, statusAktif, nowStr]);

  const attrDefs = Array.isArray(payload.attributes) ? payload.attributes : [];
  const attrSheet = ss.getSheetByName(SHEETS.ATRIBUT_PRODUK);
  const attrRowsToInsert = [];
  attrDefs.forEach(function (d) {
    const vals = Array.isArray(d.values) ? d.values.join(', ') : String(d.values || '');
    if (d.name && vals) {
      attrRowsToInsert.push([barcode, String(d.name).trim(), String(d.type || 'text').trim(), vals]);
    }
  });
  if (attrRowsToInsert.length > 0) {
    attrSheet.getRange(attrSheet.getLastRow() + 1, 1, attrRowsToInsert.length, 4).setValues(attrRowsToInsert);
  }

  const defaultHpp = Math.max(0, Number(payload.default_hpp) || 0);
  const defaultHargaJual = Math.max(0, Number(payload.default_harga_jual) || 0);
  const variantsInput = Array.isArray(payload.variants) && payload.variants.length
    ? payload.variants
    : [{ attributes: {}, active: true }];

  const varSheet = ss.getSheetByName(SHEETS.VARIASI_PRODUK);
  const varRowsToInsert = variantsInput.map(function (v, idx) {
    const attrs = (v && v.attributes && typeof v.attributes === 'object') ? v.attributes : {};
    const attrKeys = Object.keys(attrs);
    const label = attrKeys.length > 0
      ? attrKeys.map(function (k) { return attrs[k]; }).join(' · ')
      : 'Standar';
    __idSeqCounter = (__idSeqCounter + 1) % 46656;
    const seqVar = __idSeqCounter.toString(36).toUpperCase().padStart(3, '0');
    const idVar = barcode + '-V' + String(idx + 1).padStart(2, '0') + '-' + seqVar + Math.random().toString(36).substring(2, 5).toUpperCase();
    return [
      idVar,
      barcode,
      label,
      JSON.stringify(attrs),
      defaultHpp,
      defaultHargaJual,
      parseBool_(v.active, true),
      nowStr
    ];
  });
  if (varRowsToInsert.length > 0) {
    varSheet.getRange(varSheet.getLastRow() + 1, 1, varRowsToInsert.length, 8).setValues(varRowsToInsert);
    SpreadsheetApp.flush();
  }

  let mapMsg = '';
  const katTailor = normText_(payload.barcode_kategori_tailor);
  if (katTailor) {
    const mapRes = autoMapPolosCategoryInternal_(
      ss,
      barcode,
      katTailor,
      Math.max(1, Number(payload.qty_per_unit) || 1),
      parseBool_(payload.auto_create_missing_tailor, true)
    );
    if (mapRes.success) {
      mapMsg = ' & ' + mapRes.message;
    }
  }

  appendOperationLog_(ss, {
    transaction_id: trxId,
    tipe: 'CREATE_PRODUCT',
    barcode: barcode,
    id_variasi: '-',
    detail_item: nama + ' (' + variantsInput.length + ' varian)',
    lokasi: DEFAULT_GUDANG_POLOS,
    qty: 0,
    hpp: defaultHpp,
    harga_jual: defaultHargaJual,
    omzet: 0,
    profit: 0,
    pelapor: pelapor,
    keterangan: 'Produk baru terdaftar' + mapMsg
  });

  markDashboardDirty_();
  return {
    success: true,
    message: 'Produk "' + nama + '" (' + variantsInput.length + ' varian) berhasil disimpan' + mapMsg + '.'
  };
}

function handleUpdateProduct_(payload) {
  const barcode = normText_(payload.barcode);
  const nama = String(payload.nama || '').trim();
  const statusAktif = parseBool_(payload.status_aktif, true);
  const pelapor = normText_(payload.nama_pelapor) || 'ADMIN';
  const trxId = String(payload.transaction_id || generateId_('UPD')).trim();

  if (!barcode || nama.length < 2) {
    return { success: false, message: 'Barcode dan Nama Produk wajib diisi.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (isTransactionLogged_(ss, SHEETS.OPERATION_LOG, trxId)) {
    return { success: true, message: 'Perubahan data produk sudah tercatat sebelumnya.' };
  }

  const sheet = ss.getSheetByName(SHEETS.MASTER_PRODUK);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(function (h) { return String(h || '').trim(); });
  const cBar = headers.indexOf('Barcode');
  const cNama = headers.indexOf('Nama_Produk');
  const cAct = headers.indexOf('Status_Aktif');
  const cUpd = headers.indexOf('Updated_At');

  let updated = false;
  for (let r = 1; r < data.length; r++) {
    if (normText_(data[r][cBar]) === barcode) {
      sheet.getRange(r + 1, cNama + 1).setValue(nama);
      sheet.getRange(r + 1, cAct + 1).setValue(statusAktif);
      if (cUpd !== -1) sheet.getRange(r + 1, cUpd + 1).setValue(formatNow_());
      updated = true;
      break;
    }
  }

  if (!updated) {
    return { success: false, message: 'Produk tidak ditemukan.' };
  }

  appendOperationLog_(ss, {
    transaction_id: trxId,
    tipe: 'UPDATE_PRODUCT',
    barcode: barcode,
    id_variasi: '-',
    detail_item: nama + ' (' + (statusAktif ? 'AKTIF' : 'NONAKTIF') + ')',
    lokasi: DEFAULT_GUDANG_POLOS,
    qty: 0,
    hpp: 0,
    harga_jual: 0,
    omzet: 0,
    profit: 0,
    pelapor: pelapor,
    keterangan: 'Update master produk'
  });

  markDashboardDirty_();
  return { success: true, message: 'Data produk "' + nama + '" berhasil diperbarui.' };
}

function handleAddVariants_(payload) {
  const barcode = normText_(payload.barcode);
  const pelapor = normText_(payload.nama_pelapor) || 'ADMIN';
  const trxId = String(payload.transaction_id || generateId_('ADDV')).trim();
  const newVars = Array.isArray(payload.variants) ? payload.variants : [];

  if (!barcode || !newVars.length) {
    return { success: false, message: 'Pilih minimal satu kombinasi variasi baru.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (isTransactionLogged_(ss, SHEETS.OPERATION_LOG, trxId)) {
    return { success: true, message: 'Variasi baru sudah ditambahkan.' };
  }

  // Perbarui definisi atribut produk jika ada nilai baru
  const defs = Array.isArray(payload.attributes) ? payload.attributes : [];
  if (defs.length > 0) {
    const attrSheet = ss.getSheetByName(SHEETS.ATRIBUT_PRODUK);
    const attrData = attrSheet.getDataRange().getValues();
    const aHeaders = attrData[0].map(function (h) { return String(h || '').trim(); });
    const cBar = aHeaders.indexOf('Barcode');
    const cName = aHeaders.indexOf('Nama_Atribut');
    const cVals = aHeaders.indexOf('Daftar_Nilai');

    defs.forEach(function (d) {
      const dName = String(d.name || '').trim();
      const incomingVals = Array.isArray(d.values) ? d.values.map(function (x) { return String(x).trim(); }).filter(Boolean) : [];
      if (!dName || !incomingVals.length) return;

      let foundRow = -1;
      for (let r = 1; r < attrData.length; r++) {
        if (normText_(attrData[r][cBar]) === barcode && normText_(attrData[r][cName]) === normText_(dName)) {
          foundRow = r + 1;
          const existingVals = String(attrData[r][cVals] || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean);
          const mergedSet = {};
          existingVals.concat(incomingVals).forEach(function (v) { mergedSet[v] = true; });
          attrSheet.getRange(foundRow, cVals + 1).setValue(Object.keys(mergedSet).join(', '));
          break;
        }
      }
      if (foundRow === -1) {
        attrSheet.appendRow([barcode, dName, 'text', incomingVals.join(', ')]);
      }
    });
  }

  const varSheet = ss.getSheetByName(SHEETS.VARIASI_PRODUK);
  const existingVars = getSheetObjects_(varSheet).filter(function (v) {
    return normText_(v.Barcode) === barcode;
  });

  // Inherit harga dari varian yang sudah ada jika default_hpp / default_harga_jual tidak diisi
  let inheritHpp = Math.max(0, Number(payload.default_hpp) || 0);
  let inheritHj = Math.max(0, Number(payload.default_harga_jual) || 0);
  if ((inheritHpp <= 0 || inheritHj <= 0) && existingVars.length > 0) {
    if (inheritHpp <= 0) inheritHpp = Number(existingVars[0].HPP) || 0;
    if (inheritHj <= 0) inheritHj = Number(existingVars[0].Harga_Jual) || 0;
  }

  const nowStr = formatNow_();
  let addedCount = 0;
  const rowsToInsert = [];

  newVars.forEach(function (v) {
    const attrs = (v && v.attributes && typeof v.attributes === 'object') ? v.attributes : {};
    const keys = Object.keys(attrs);
    const label = keys.length > 0 ? keys.map(function (k) { return attrs[k]; }).join(' · ') : 'Standar';

    const duplicate = existingVars.some(function (ev) {
      return normText_(ev.Label_Variasi) === normText_(label);
    });
    if (duplicate) return;

    __idSeqCounter = (__idSeqCounter + 1) % 46656;
    const seqVar = __idSeqCounter.toString(36).toUpperCase().padStart(3, '0');
    const idVar = barcode + '-V' + String(existingVars.length + addedCount + 1).padStart(2, '0') + '-' + seqVar + Math.random().toString(36).substring(2, 5).toUpperCase();
    rowsToInsert.push([
      idVar,
      barcode,
      label,
      JSON.stringify(attrs),
      inheritHpp,
      inheritHj,
      true,
      nowStr
    ]);
    addedCount++;
  });

  if (rowsToInsert.length > 0) {
    varSheet.getRange(varSheet.getLastRow() + 1, 1, rowsToInsert.length, 8).setValues(rowsToInsert);
    SpreadsheetApp.flush();
  }

  let mapMsg = '';
  const katTailor = normText_(payload.barcode_kategori_tailor);
  if (katTailor) {
    const mapRes = autoMapPolosCategoryInternal_(
      ss,
      barcode,
      katTailor,
      Math.max(1, Number(payload.qty_per_unit) || 1),
      parseBool_(payload.auto_create_missing_tailor, true)
    );
    if (mapRes.success) mapMsg = ' & ' + mapRes.message;
  }

  appendOperationLog_(ss, {
    transaction_id: trxId,
    tipe: 'ADD_VARIANTS',
    barcode: barcode,
    id_variasi: '-',
    detail_item: '+' + addedCount + ' variasi baru',
    lokasi: DEFAULT_GUDANG_POLOS,
    qty: 0,
    hpp: inheritHpp,
    harga_jual: inheritHj,
    omzet: 0,
    profit: 0,
    pelapor: pelapor,
    keterangan: 'Tambah variasi produk' + mapMsg
  });

  markDashboardDirty_();
  return {
    success: true,
    message: addedCount + ' variasi baru berhasil ditambahkan' + mapMsg + '.'
  };
}

function handleSetVariantStatus_(payload) {
  const barcode = normText_(payload.barcode);
  const varId = String(payload.variant_id || '').trim();
  const active = parseBool_(payload.active, true);
  const pelapor = normText_(payload.nama_pelapor) || 'ADMIN';
  const trxId = String(payload.transaction_id || generateId_('VARST')).trim();

  if (!barcode || !varId) {
    return { success: false, message: 'Barcode dan ID Variasi wajib diisi.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (isTransactionLogged_(ss, SHEETS.OPERATION_LOG, trxId)) {
    return { success: true, message: 'Perubahan status variasi sudah tercatat sebelumnya.' };
  }

  const sheet = ss.getSheetByName(SHEETS.VARIASI_PRODUK);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(function (h) { return String(h || '').trim(); });
  const cId = headers.indexOf('ID_Variasi');
  const cBar = headers.indexOf('Barcode');
  const cLbl = headers.indexOf('Label_Variasi');
  const cAct = headers.indexOf('Status_Aktif');
  const cUpd = headers.indexOf('Updated_At');

  let label = varId;
  let found = false;
  for (let r = 1; r < data.length; r++) {
    const rowId = String(data[r][cId] || '').trim();
    const rowBar = cBar !== -1 ? normText_(data[r][cBar]) : barcode;
    if (rowId === varId && rowBar === barcode) {
      label = String(data[r][cLbl] || varId).trim();
      sheet.getRange(r + 1, cAct + 1).setValue(active);
      if (cUpd !== -1) sheet.getRange(r + 1, cUpd + 1).setValue(formatNow_());
      found = true;
      break;
    }
  }

  if (!found) {
    return { success: false, message: 'Varian tidak ditemukan.' };
  }

  appendOperationLog_(ss, {
    transaction_id: trxId,
    tipe: 'SET_VARIANT_STATUS',
    barcode: barcode,
    id_variasi: varId,
    detail_item: label + ' -> ' + (active ? 'AKTIF' : 'NONAKTIF'),
    lokasi: DEFAULT_GUDANG_POLOS,
    qty: 0,
    hpp: 0,
    harga_jual: 0,
    omzet: 0,
    profit: 0,
    pelapor: pelapor,
    keterangan: 'Ubah status variasi'
  });

  markDashboardDirty_();
  return { success: true, message: 'Status variasi "' + label + '" diubah menjadi ' + (active ? 'Aktif' : 'Nonaktif') + '.' };
}

function handleUpdateVariantPricing_(payload) {
  const barcode = normText_(payload.barcode);
  const pelapor = normText_(payload.nama_pelapor) || 'ADMIN';
  const trxId = String(payload.transaction_id || generateId_('PRC')).trim();
  const items = Array.isArray(payload.items) ? payload.items : [];

  if (!barcode || !items.length) {
    return { success: false, message: 'Data harga varian tidak lengkap.' };
  }

  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    if (it.hpp !== undefined && it.hpp !== null && it.hpp !== '' && (isNaN(Number(it.hpp)) || Number(it.hpp) < 0)) {
      return { success: false, message: 'Harga Modal (HPP) tidak boleh bernilai negatif.' };
    }
    if (it.harga_jual !== undefined && it.harga_jual !== null && it.harga_jual !== '' && (isNaN(Number(it.harga_jual)) || Number(it.harga_jual) < 0)) {
      return { success: false, message: 'Harga Jual tidak boleh bernilai negatif.' };
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (isTransactionLogged_(ss, SHEETS.OPERATION_LOG, trxId)) {
    return { success: true, message: 'Perubahan harga varian sudah tercatat sebelumnya.' };
  }

  const sheet = ss.getSheetByName(SHEETS.VARIASI_PRODUK);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(function (h) { return String(h || '').trim(); });
  const cId = headers.indexOf('ID_Variasi');
  const cBar = headers.indexOf('Barcode');
  const cHpp = headers.indexOf('HPP');
  const cHj = headers.indexOf('Harga_Jual');
  const cUpd = headers.indexOf('Updated_At');

  const itemMap = {};
  items.forEach(function (it) {
    itemMap[String(it.variant_id || '').trim()] = it;
  });

  const nowStr = formatNow_();
  let updatedCount = 0;

  for (let r = 1; r < data.length; r++) {
    const vid = String(data[r][cId] || '').trim();
    const rowBar = cBar !== -1 ? normText_(data[r][cBar]) : barcode;
    if (rowBar === barcode && itemMap[vid]) {
      const it = itemMap[vid];
      if (it.hpp !== undefined && it.hpp !== null && it.hpp !== '') {
        sheet.getRange(r + 1, cHpp + 1).setValue(Math.max(0, Number(it.hpp) || 0));
      }
      if (it.harga_jual !== undefined && it.harga_jual !== null && it.harga_jual !== '') {
        sheet.getRange(r + 1, cHj + 1).setValue(Math.max(0, Number(it.harga_jual) || 0));
      }
      if (cUpd !== -1) sheet.getRange(r + 1, cUpd + 1).setValue(nowStr);
      updatedCount++;
    }
  }

  if (updatedCount === 0) {
    return { success: false, message: 'Tidak ada varian yang cocok untuk diperbarui harganya.' };
  }

  appendOperationLog_(ss, {
    transaction_id: trxId,
    tipe: 'UPDATE_PRICING',
    barcode: barcode,
    id_variasi: items[0].variant_id || '-',
    detail_item: 'Update harga ' + updatedCount + ' varian',
    lokasi: DEFAULT_GUDANG_POLOS,
    qty: 0,
    hpp: Math.max(0, Number(items[0].hpp) || 0),
    harga_jual: Math.max(0, Number(items[0].harga_jual) || 0),
    omzet: 0,
    profit: 0,
    pelapor: pelapor,
    keterangan: 'Perbarui HPP & Harga Jual'
  });

  markDashboardDirty_();
  return { success: true, message: 'Harga ' + updatedCount + ' varian berhasil diperbarui.' };
}

function handleSetVariantPricingBulk_(payload) {
  const barcode = normText_(payload.barcode);
  const pelapor = normText_(payload.nama_pelapor) || 'ADMIN';
  const trxId = String(payload.transaction_id || generateId_('PRCB')).trim();
  const hasHpp = payload.hpp !== undefined && payload.hpp !== null && payload.hpp !== '';
  const hasHj = payload.harga_jual !== undefined && payload.harga_jual !== null && payload.harga_jual !== '';

  if (!barcode || (!hasHpp && !hasHj)) {
    return { success: false, message: 'Isi minimal Harga Modal (HPP) atau Harga Jual.' };
  }

  const newHpp = hasHpp ? Number(payload.hpp) : null;
  const newHj = hasHj ? Number(payload.harga_jual) : null;
  if ((newHpp !== null && (isNaN(newHpp) || newHpp < 0)) || (newHj !== null && (isNaN(newHj) || newHj < 0))) {
    return { success: false, message: 'Harga Modal (HPP) dan Harga Jual tidak boleh bernilai negatif.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (isTransactionLogged_(ss, SHEETS.OPERATION_LOG, trxId)) {
    return { success: true, message: 'Perubahan harga massal sudah tercatat sebelumnya.' };
  }

  const sheet = ss.getSheetByName(SHEETS.VARIASI_PRODUK);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(function (h) { return String(h || '').trim(); });
  const cBar = headers.indexOf('Barcode');
  const cHpp = headers.indexOf('HPP');
  const cHj = headers.indexOf('Harga_Jual');
  const cUpd = headers.indexOf('Updated_At');

  const nowStr = formatNow_();
  let count = 0;

  for (let r = 1; r < data.length; r++) {
    if (normText_(data[r][cBar]) === barcode) {
      if (newHpp !== null) sheet.getRange(r + 1, cHpp + 1).setValue(newHpp);
      if (newHj !== null) sheet.getRange(r + 1, cHj + 1).setValue(newHj);
      if (cUpd !== -1) sheet.getRange(r + 1, cUpd + 1).setValue(nowStr);
      count++;
    }
  }

  if (count === 0) {
    return { success: false, message: 'Tidak ada varian pada produk "' + barcode + '" untuk diperbarui.' };
  }

  appendOperationLog_(ss, {
    transaction_id: trxId,
    tipe: 'UPDATE_PRICING',
    barcode: barcode,
    id_variasi: '*',
    detail_item: 'Samakan harga semua (' + count + ' varian)',
    lokasi: DEFAULT_GUDANG_POLOS,
    qty: 0,
    hpp: newHpp || 0,
    harga_jual: newHj || 0,
    omzet: 0,
    profit: 0,
    pelapor: pelapor,
    keterangan: 'Update harga massal'
  });

  markDashboardDirty_();
  return { success: true, message: 'Harga seluruh (' + count + ') varian berhasil disamakan.' };
}

function handleSavePolosMapping_(payload) {
  const barcodeProduk = normText_(payload.barcode_produk);
  const idVariasi = String(payload.id_variasi || '*').trim();
  const idTailor = String(payload.id_tailor || '').trim();
  const qtyPerUnit = Math.max(1, Number(payload.qty_per_unit) || 1);

  if (!barcodeProduk) {
    return { success: false, message: 'Barcode produk wajib diisi.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapSheet = ss.getSheetByName(SHEETS.MAPPING_POLOS_PRODUK);
  const data = mapSheet.getDataRange().getValues();
  const headers = data[0].map(function (h) { return String(h || '').trim(); });
  const cBar = headers.indexOf('Barcode_Produk');
  const cVar = headers.indexOf('ID_Variasi');

  // Hapus baris mapping lama untuk (barcodeProduk, idVariasi)
  for (let r = data.length - 1; r >= 1; r--) {
    if (normText_(data[r][cBar]) === barcodeProduk && String(data[r][cVar] || '').trim() === idVariasi) {
      mapSheet.deleteRow(r + 1);
    }
  }

  if (idTailor) {
    const mtRows = getSheetObjects_(ss.getSheetByName(SHEETS.MASTER_TAILOR));
    const mt = mtRows.find(function (m) { return String(m.ID_Tailor || '').trim() === idTailor; });
    if (!mt) {
      return { success: false, message: 'ID Kaos Polos "' + idTailor + '" tidak ditemukan.' };
    }

    const varRows = getSheetObjects_(ss.getSheetByName(SHEETS.VARIASI_PRODUK));
    const vr = varRows.find(function (v) { return String(v.ID_Variasi || '').trim() === idVariasi; });
    const labelVar = vr ? String(vr.Label_Variasi || idVariasi) : idVariasi;
    const detailPolos = mt.Nama_Kategori + ' · ' + mt.Warna + ' · ' + mt.Ukuran;

    mapSheet.appendRow([
      generateId_('MAP'),
      barcodeProduk,
      idVariasi,
      labelVar,
      idTailor,
      detailPolos,
      qtyPerUnit,
      formatNow_()
    ]);
  }

  markDashboardDirty_();
  const updatedMappings = getSheetObjects_(mapSheet).map(function (mp) {
    return {
      mapping_id: String(mp.Mapping_ID || '').trim(),
      barcode_produk: String(mp.Barcode_Produk || '').trim(),
      id_variasi: String(mp.ID_Variasi || '*').trim(),
      label_variasi: String(mp.Label_Variasi || '').trim(),
      id_tailor: String(mp.ID_Tailor || '').trim(),
      detail_polos: String(mp.Detail_Polos || '').trim(),
      qty_per_unit: Math.max(1, Number(mp.Qty_Per_Unit) || 1),
      updated_at: mp.Updated_At || ''
    };
  });

  return {
    success: true,
    message: idTailor ? 'Tautan Gudang Polos disimpan.' : 'Tautan Gudang Polos dihapus.',
    data: { polos_mappings: updatedMappings }
  };
}

function handleAutoMapPolosCategory_(payload) {
  const barcodeProduk = normText_(payload.barcode_produk);
  const barcodeKatTailor = normText_(payload.barcode_kategori_tailor);
  const qtyPerUnit = Math.max(1, Number(payload.qty_per_unit) || 1);
  const autoCreateMissing = parseBool_(payload.auto_create_missing_tailor, false);

  if (!barcodeProduk || !barcodeKatTailor) {
    return { success: false, message: 'Barcode produk dan kategori kaos polos wajib dipilih.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const res = autoMapPolosCategoryInternal_(ss, barcodeProduk, barcodeKatTailor, qtyPerUnit, autoCreateMissing);
  if (!res.success) return res;

  markDashboardDirty_();
  const boot = handleGetBootstrap_();
  return {
    success: true,
    message: res.message,
    data: {
      polos_mappings: boot.data.polos_mappings,
      master_tailor: boot.data.master_tailor
    }
  };
}

function autoMapPolosCategoryInternal_(ss, barcodeProduk, barcodeKatTailor, qtyPerUnit, autoCreateMissing) {
  const varSheet = ss.getSheetByName(SHEETS.VARIASI_PRODUK);
  const prodVars = getSheetObjects_(varSheet).filter(function (v) {
    return normText_(v.Barcode) === barcodeProduk && parseBool_(v.Status_Aktif, true);
  });

  if (!prodVars.length) {
    return { success: false, message: 'Produk tidak memiliki variasi aktif untuk ditautkan.' };
  }

  const mtSheet = ss.getSheetByName(SHEETS.MASTER_TAILOR);
  const stSheet = ss.getSheetByName(SHEETS.STOK_TAILOR);
  let catTailors = getSheetObjects_(mtSheet).filter(function (m) {
    return normText_(m.Barcode_Kategori) === barcodeKatTailor && parseBool_(m.Status_Aktif, true);
  });

  if (!catTailors.length) {
    return { success: false, message: 'Kategori kaos polos "' + barcodeKatTailor + '" tidak ditemukan di Master Tailor.' };
  }

  // Simpan peta Qty_Per_Unit eksisting per ID_Variasi agar aturan rasio kustom (> 1) tidak hilang
  const mapSheet = ss.getSheetByName(SHEETS.MAPPING_POLOS_PRODUK);
  const mapData = mapSheet.getDataRange().getValues();
  const mHeaders = mapData[0].map(function (h) { return String(h || '').trim(); });
  const cBar = mHeaders.indexOf('Barcode_Produk');
  const cVarId = mHeaders.indexOf('ID_Variasi');
  const cQtyUnit = mHeaders.indexOf('Qty_Per_Unit');

  const existingQtyByVarId = {};
  for (let r = 1; r < mapData.length; r++) {
    if (normText_(mapData[r][cBar]) === barcodeProduk) {
      const exVar = String(mapData[r][cVarId] || '').trim();
      const exQty = Math.max(1, Number(mapData[r][cQtyUnit]) || 1);
      if (exVar) existingQtyByVarId[exVar] = exQty;
    }
  }

  const sampleCat = catTailors[0];
  const lookup = {};
  catTailors.forEach(function (m) {
    const key = normalizeMatchToken_(m.Warna) + '||' + normalizeMatchToken_(m.Ukuran);
    if (!lookup[key]) lookup[key] = m;
  });

  const nowStr = formatNow_();
  let matchedCount = 0;
  let createdTailorCount = 0;
  const newMappings = [];
  const newMtRows = [];
  const newStRows = [];

  prodVars.forEach(function (v) {
    let attrs = {};
    try {
      if (v.Atribut_JSON) attrs = JSON.parse(String(v.Atribut_JSON));
    } catch (e) { }
    const cs = extractColorSizeFromVariant_(attrs, v.Label_Variasi);
    const normW = normalizeMatchToken_(cs.warna);
    const normU = normalizeMatchToken_(cs.ukuran);

    let found = lookup[normW + '||' + normU] || null;
    if (!found && normU === '-') found = lookup[normW + '||ALLSIZE'] || null;
    if (!found && normU === 'ALLSIZE') found = lookup[normW + '||' + '-'] || null;

    if (!found && autoCreateMissing) {
      const newIdTailor = generateId_('TLR');
      const cleanWarna = cs.warna || '-';
      const cleanUkuran = normText_(cs.ukuran || 'ALLSIZE') || 'ALLSIZE';
      const upahDefault = Number(sampleCat.Upah_Per_Pcs) || 0;

      newMtRows.push([
        newIdTailor,
        barcodeKatTailor,
        sampleCat.Nama_Kategori,
        cleanWarna,
        cleanUkuran,
        upahDefault,
        true,
        nowStr
      ]);

      newStRows.push([
        generateId_('STK'),
        newIdTailor,
        DEFAULT_GUDANG_POLOS,
        barcodeKatTailor,
        sampleCat.Nama_Kategori,
        cleanWarna,
        cleanUkuran,
        0,
        nowStr
      ]);

      found = {
        ID_Tailor: newIdTailor,
        Barcode_Kategori: barcodeKatTailor,
        Nama_Kategori: sampleCat.Nama_Kategori,
        Warna: cleanWarna,
        Ukuran: cleanUkuran
      };
      lookup[normW + '||' + normalizeMatchToken_(cleanUkuran)] = found;
      createdTailorCount++;
    }

    if (found) {
      matchedCount++;
      const vidStr = String(v.ID_Variasi || '').trim();
      const effectiveQty = existingQtyByVarId[vidStr] || existingQtyByVarId['*'] || qtyPerUnit;
      const detailPolos = found.Nama_Kategori + ' · ' + found.Warna + ' · ' + found.Ukuran;
      newMappings.push([
        generateId_('MAP'),
        barcodeProduk,
        vidStr,
        String(v.Label_Variasi || '').trim(),
        String(found.ID_Tailor || '').trim(),
        detailPolos,
        effectiveQty,
        nowStr
      ]);
    }
  });

  if (!newMappings.length) {
    return {
      success: false,
      message: 'Tidak ada warna & ukuran yang cocok dengan "' + sampleCat.Nama_Kategori + '". Aktifkan opsi Buat Otomatis Varian Kurang.'
    };
  }

  if (newMtRows.length > 0) {
    mtSheet.getRange(mtSheet.getLastRow() + 1, 1, newMtRows.length, 8).setValues(newMtRows);
  }
  if (newStRows.length > 0) {
    stSheet.getRange(stSheet.getLastRow() + 1, 1, newStRows.length, 9).setValues(newStRows);
  }

  // Ganti mapping lama untuk produk ini dengan hasil auto-match baru
  for (let r = mapData.length - 1; r >= 1; r--) {
    if (normText_(mapData[r][cBar]) === barcodeProduk) {
      mapSheet.deleteRow(r + 1);
    }
  }

  mapSheet.getRange(mapSheet.getLastRow() + 1, 1, newMappings.length, 8).setValues(newMappings);
  SpreadsheetApp.flush();

  let msg = matchedCount + '/' + prodVars.length + ' varian berhasil ditautkan ke ' + sampleCat.Nama_Kategori;
  if (createdTailorCount > 0) {
    msg += ' (+' + createdTailorCount + ' varian polos baru dibuat)';
  }
  return { success: true, message: msg };
}

// ============================================================================
// 5. TRANSAKSI BARANG KELUAR MULTI-PRODUK (100% POTONG GUDANG POLOS)
// ============================================================================

function handleSaveBulkKeluar_(payload) {
  const batchId = String(payload.batch_id || generateId_('BLK')).trim();
  const pelapor = normText_(payload.nama_pelapor) || 'ADMIN';
  const items = Array.isArray(payload.items) ? payload.items : [];

  if (!items.length) {
    return { success: false, message: 'Keranjang barang keluar masih kosong.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (isTransactionLogged_(ss, SHEETS.OPERATION_LOG, batchId)) {
    return { success: true, message: 'Transaksi pesanan ini sudah diproses sebelumnya.' };
  }

  const prodRows = getSheetObjects_(ss.getSheetByName(SHEETS.MASTER_PRODUK));
  const varRows = getSheetObjects_(ss.getSheetByName(SHEETS.VARIASI_PRODUK));
  const mtRows = getSheetObjects_(ss.getSheetByName(SHEETS.MASTER_TAILOR));
  const mapRows = getSheetObjects_(ss.getSheetByName(SHEETS.MAPPING_POLOS_PRODUK));
  const stSheet = ss.getSheetByName(SHEETS.STOK_TAILOR);
  const stData = stSheet.getDataRange().getValues();
  const stHeaders = stData[0].map(function (h) { return String(h || '').trim(); });

  const cStIdTailor = stHeaders.indexOf('ID_Tailor');
  const cStPenjahit = stHeaders.indexOf('Nama_Penjahit');
  const cStReady = stHeaders.indexOf('Stok_Ready');
  const cStUpd = stHeaders.indexOf('Updated_At');

  const mtActiveMap = {};
  mtRows.forEach(function (m) {
    mtActiveMap[String(m.ID_Tailor || '').trim()] = parseBool_(m.Status_Aktif, true);
  });

  // Agregasi seluruh baris GUDANG POLOS per ID_Tailor (aman jika terdapat duplikasi baris)
  const polosRowsByTailor = {};
  for (let r = 1; r < stData.length; r++) {
    if (normText_(stData[r][cStPenjahit]) === DEFAULT_GUDANG_POLOS) {
      const idT = String(stData[r][cStIdTailor] || '').trim();
      if (!polosRowsByTailor[idT]) {
        polosRowsByTailor[idT] = { totalStok: 0, rows: [] };
      }
      const rowQty = Number(stData[r][cStReady]) || 0;
      polosRowsByTailor[idT].totalStok += rowQty;
      polosRowsByTailor[idT].rows.push({
        rowIndex: r + 1,
        stok: rowQty
      });
    }
  }

  const demandByTailor = {};
  const preparedLines = [];

  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    const barcode = normText_(it.barcode);
    const varId = String(it.variant_id || '').trim();
    const qty = Math.floor(Number(it.qty) || 0);

    if (!barcode || !varId || qty < 1) {
      return { success: false, message: 'Item ke-' + (i + 1) + ' tidak valid.' };
    }

    const prod = prodRows.find(function (p) { return normText_(p.Barcode) === barcode; });
    const variant = varRows.find(function (v) { return String(v.ID_Variasi || '').trim() === varId; });
    if (!prod || !variant) {
      return { success: false, message: 'Produk atau variasi "' + varId + '" tidak ditemukan.' };
    }
    if (!parseBool_(prod.Status_Aktif, true) || !parseBool_(variant.Status_Aktif, true)) {
      return {
        success: false,
        message: 'Produk atau variasi "' + prod.Nama_Produk + ' - ' + variant.Label_Variasi + '" berstatus NONAKTIF.'
      };
    }

    const hpp = Number(variant.HPP) || 0;
    const hargaJual = Number(variant.Harga_Jual) || 0;
    if (hpp <= 0 || hargaJual <= 0) {
      return {
        success: false,
        message: 'Modal (HPP) & Harga Jual untuk "' + prod.Nama_Produk + ' - ' + variant.Label_Variasi + '" belum diatur.'
      };
    }

    const mapping = mapRows.find(function (m) {
      return normText_(m.Barcode_Produk) === barcode && String(m.ID_Variasi || '').trim() === varId;
    }) || mapRows.find(function (m) {
      return normText_(m.Barcode_Produk) === barcode && String(m.ID_Variasi || '').trim() === '*';
    });

    if (!mapping) {
      return {
        success: false,
        message: 'Varian "' + prod.Nama_Produk + ' (' + variant.Label_Variasi + ')" belum ditautkan ke Gudang Polos.'
      };
    }

    const idTailor = String(mapping.ID_Tailor || '').trim();
    if (mtActiveMap[idTailor] === false) {
      return {
        success: false,
        message: 'Bahan kaos polos "' + (mapping.Detail_Polos || idTailor) + '" berstatus NONAKTIF.'
      };
    }

    const qtyPerUnit = Math.max(1, Number(mapping.Qty_Per_Unit) || 1);
    const neededPolos = qty * qtyPerUnit;

    if (!demandByTailor[idTailor]) {
      demandByTailor[idTailor] = {
        needed: 0,
        detail: String(mapping.Detail_Polos || idTailor)
      };
    }
    demandByTailor[idTailor].needed += neededPolos;

    preparedLines.push({
      barcode: barcode,
      nama_produk: String(prod.Nama_Produk || barcode),
      variant_id: varId,
      label_variasi: String(variant.Label_Variasi || 'Standar'),
      qty: qty,
      hpp: hpp,
      harga_jual: hargaJual,
      omzet: qty * hargaJual,
      profit: qty * (hargaJual - hpp),
      id_tailor: idTailor,
      detail_polos: String(mapping.Detail_Polos || idTailor),
      needed_polos: neededPolos
    });
  }

  // Validasi kecukupan stok Gudang Polos sebelum menulis apa pun (Atomik)
  const tailorIds = Object.keys(demandByTailor);
  for (let t = 0; t < tailorIds.length; t++) {
    const idT = tailorIds[t];
    const req = demandByTailor[idT];
    const avail = polosRowsByTailor[idT] ? polosRowsByTailor[idT].totalStok : 0;
    if (avail < req.needed) {
      return {
        success: false,
        message: 'Stok Gudang Polos tidak cukup untuk ' + req.detail + ' (butuh ' + req.needed + ' pcs, tersedia ' + avail + ' pcs).'
      };
    }
  }

  // 1. Potong saldo GUDANG POLOS di STOK_TAILOR (mendukung multi-baris per ID_Tailor)
  const nowStr = formatNow_();
  tailorIds.forEach(function (idT) {
    let remainingToDeduct = demandByTailor[idT].needed;
    const rows = polosRowsByTailor[idT].rows;
    for (let k = 0; k < rows.length && remainingToDeduct > 0; k++) {
      const deduct = Math.min(rows[k].stok, remainingToDeduct);
      const sisaBaru = rows[k].stok - deduct;
      remainingToDeduct -= deduct;
      stSheet.getRange(rows[k].rowIndex, cStReady + 1).setValue(sisaBaru);
      if (cStUpd !== -1) stSheet.getRange(rows[k].rowIndex, cStUpd + 1).setValue(nowStr);
    }
  });

  // 2. Batch insert seluruh baris log penjualan / barang keluar ke OPERATION_LOG
  let totalPcs = 0;
  let totalOmzet = 0;
  const logRowsBatch = preparedLines.map(function (ln) {
    totalPcs += ln.qty;
    totalOmzet += ln.omzet;
    return [
      generateId_('LOG'),
      batchId,
      nowStr,
      'KELUAR',
      ln.barcode,
      ln.variant_id,
      ln.nama_produk + ' · ' + ln.label_variasi,
      DEFAULT_GUDANG_POLOS,
      -ln.qty,
      ln.hpp,
      ln.harga_jual,
      ln.omzet,
      ln.profit,
      pelapor,
      'Potong polos: ' + ln.detail_polos + ' (-' + ln.needed_polos + ' pcs)'
    ];
  });

  const opSheet = ss.getSheetByName(SHEETS.OPERATION_LOG);
  opSheet.getRange(opSheet.getLastRow() + 1, 1, logRowsBatch.length, logRowsBatch[0].length)
    .setValues(logRowsBatch);
  SpreadsheetApp.flush();

  markDashboardDirty_();
  return {
    success: true,
    message: preparedLines.length + ' item (' + totalPcs + ' pcs · ' + formatRp_(totalOmzet) + ') keluar & saldo Gudang Polos terpotong.'
  };
}

// ============================================================================
// 6. PENJAHIT, GUDANG POLOS & UPAH JAHIT
// ============================================================================

function handleSaveMasterTailor_(payload) {
  const editId = String(payload.id_tailor || '').trim();
  const barcodeKat = normText_(payload.barcode_kategori);
  const namaKat = String(payload.nama_kategori || '').trim();
  const upahDefault = Math.max(0, Number(payload.upah_per_pcs) || 0);
  const statusAktif = parseBool_(payload.status_aktif, true);

  if (!barcodeKat || !namaKat) {
    return { success: false, message: 'Kode QR dan Nama Kaos Polos wajib diisi.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mtSheet = ss.getSheetByName(SHEETS.MASTER_TAILOR);
  const stSheet = ss.getSheetByName(SHEETS.STOK_TAILOR);
  const nowStr = formatNow_();

  if (editId) {
    const warna = String(payload.warna || '-').trim() || '-';
    const ukuran = normText_(payload.ukuran || '-') || '-';

    const data = mtSheet.getDataRange().getValues();
    const headers = data[0].map(function (h) { return String(h || '').trim(); });
    const cId = headers.indexOf('ID_Tailor');
    const cBar = headers.indexOf('Barcode_Kategori');
    const cNama = headers.indexOf('Nama_Kategori');
    const cWarna = headers.indexOf('Warna');
    const cUkuran = headers.indexOf('Ukuran');
    const cUpah = headers.indexOf('Upah_Per_Pcs');
    const cAct = headers.indexOf('Status_Aktif');
    const cUpd = headers.indexOf('Updated_At');

    let found = false;
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][cId] || '').trim() === editId) {
        mtSheet.getRange(r + 1, cBar + 1).setValue(barcodeKat);
        mtSheet.getRange(r + 1, cNama + 1).setValue(namaKat);
        mtSheet.getRange(r + 1, cWarna + 1).setValue(warna);
        mtSheet.getRange(r + 1, cUkuran + 1).setValue(ukuran);
        mtSheet.getRange(r + 1, cUpah + 1).setValue(upahDefault);
        mtSheet.getRange(r + 1, cAct + 1).setValue(statusAktif);
        if (cUpd !== -1) mtSheet.getRange(r + 1, cUpd + 1).setValue(nowStr);
        found = true;
        break;
      }
    }

    if (!found) {
      return { success: false, message: 'ID Master Tailor tidak ditemukan.' };
    }

    // Sinkronkan label di STOK_TAILOR
    const stData = stSheet.getDataRange().getValues();
    const sHeaders = stData[0].map(function (h) { return String(h || '').trim(); });
    const sId = sHeaders.indexOf('ID_Tailor');
    const sBar = sHeaders.indexOf('Barcode_Kategori');
    const sNama = sHeaders.indexOf('Nama_Kategori');
    const sWarna = sHeaders.indexOf('Warna');
    const sUkuran = sHeaders.indexOf('Ukuran');
    for (let sr = 1; sr < stData.length; sr++) {
      if (String(stData[sr][sId] || '').trim() === editId) {
        stSheet.getRange(sr + 1, sBar + 1).setValue(barcodeKat);
        stSheet.getRange(sr + 1, sNama + 1).setValue(namaKat);
        stSheet.getRange(sr + 1, sWarna + 1).setValue(warna);
        stSheet.getRange(sr + 1, sUkuran + 1).setValue(ukuran);
      }
    }

    // Sinkronkan label Detail_Polos di MAPPING_POLOS_PRODUK
    const mapSheet = ss.getSheetByName(SHEETS.MAPPING_POLOS_PRODUK);
    const mapData = mapSheet.getDataRange().getValues();
    const mHeaders = mapData[0].map(function (h) { return String(h || '').trim(); });
    const mTailorId = mHeaders.indexOf('ID_Tailor');
    const mDetail = mHeaders.indexOf('Detail_Polos');
    const mUpd = mHeaders.indexOf('Updated_At');
    const newDetailPolos = namaKat + ' · ' + warna + ' · ' + ukuran;

    if (mTailorId !== -1 && mDetail !== -1) {
      for (let mr = 1; mr < mapData.length; mr++) {
        if (String(mapData[mr][mTailorId] || '').trim() === editId) {
          mapSheet.getRange(mr + 1, mDetail + 1).setValue(newDetailPolos);
          if (mUpd !== -1) mapSheet.getRange(mr + 1, mUpd + 1).setValue(nowStr);
        }
      }
    }

    markDashboardDirty_();
    return { success: true, message: 'Varian kaos polos & ongkos jahit berhasil diperbarui.' };
  }

  // Mode Tambah Baru (single atau batch variants)
  const variantsInput = Array.isArray(payload.variants) && payload.variants.length
    ? payload.variants
    : [{
      warna: String(payload.warna || '-').trim() || '-',
      ukuran: normText_(payload.ukuran || '-') || '-',
      upah_per_pcs: upahDefault,
      status_aktif: statusAktif
    }];

  const existingMasters = getSheetObjects_(mtSheet);
  let savedCount = 0;
  const newMtRows = [];
  const newStRows = [];

  variantsInput.forEach(function (v) {
    const warna = String(v.warna || '-').trim() || '-';
    const ukuran = normText_(v.ukuran || '-') || '-';
    const upah = Math.max(0, Number(v.upah_per_pcs !== undefined ? v.upah_per_pcs : upahDefault) || 0);
    const isAct = parseBool_(v.status_aktif, statusAktif);

    const dup = existingMasters.some(function (m) {
      return normText_(m.Barcode_Kategori) === barcodeKat &&
        normalizeMatchToken_(m.Warna) === normalizeMatchToken_(warna) &&
        normalizeMatchToken_(m.Ukuran) === normalizeMatchToken_(ukuran);
    });
    if (dup) return;

    const newId = generateId_('TLR');
    newMtRows.push([newId, barcodeKat, namaKat, warna, ukuran, upah, isAct, nowStr]);
    newStRows.push([generateId_('STK'), newId, DEFAULT_GUDANG_POLOS, barcodeKat, namaKat, warna, ukuran, 0, nowStr]);
    savedCount++;
  });

  if (savedCount === 0) {
    return { success: false, message: 'Semua kombinasi warna & ukuran tersebut sudah terdaftar pada kode ' + barcodeKat + '.' };
  }

  mtSheet.getRange(mtSheet.getLastRow() + 1, 1, newMtRows.length, 8).setValues(newMtRows);
  stSheet.getRange(stSheet.getLastRow() + 1, 1, newStRows.length, 9).setValues(newStRows);
  SpreadsheetApp.flush();

  markDashboardDirty_();
  return {
    success: true,
    message: savedCount + ' pilihan kaos polos (' + namaKat + ') berhasil disimpan.'
  };
}

function handleSubmitStokTailor_(payload) {
  const trxId = String(payload.transaction_id || generateId_('JHT')).trim();
  const idTailor = String(payload.id_tailor || '').trim();
  const namaPenjahit = normText_(payload.nama_penjahit);
  const qty = Math.floor(Number(payload.qty) || 0);

  if (!idTailor || !namaPenjahit || qty < 1) {
    return { success: false, message: 'Penjahit, varian kaos polos, dan jumlah hasil jahit (>= 1) wajib diisi.' };
  }
  if (namaPenjahit === DEFAULT_GUDANG_POLOS) {
    return { success: false, message: 'Gunakan menu Terima Setoran Jahitan untuk memasukkan stok ke Gudang Polos.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (isTransactionLogged_(ss, SHEETS.LOG_TAILOR, trxId)) {
    return { success: true, message: 'Hasil jahit sudah tercatat.' };
  }

  const mt = getSheetObjects_(ss.getSheetByName(SHEETS.MASTER_TAILOR)).find(function (m) {
    return String(m.ID_Tailor || '').trim() === idTailor;
  });
  if (!mt) {
    return { success: false, message: 'Varian kaos polos tidak ditemukan di Master Tailor.' };
  }
  if (!parseBool_(mt.Status_Aktif, true)) {
    return { success: false, message: 'Varian kaos polos "' + mt.Nama_Kategori + ' · ' + mt.Warna + ' · ' + mt.Ukuran + '" berstatus NONAKTIF.' };
  }

  const stSheet = ss.getSheetByName(SHEETS.STOK_TAILOR);
  const stData = stSheet.getDataRange().getValues();
  const headers = stData[0].map(function (h) { return String(h || '').trim(); });
  const cIdT = headers.indexOf('ID_Tailor');
  const cPen = headers.indexOf('Nama_Penjahit');
  const cReady = headers.indexOf('Stok_Ready');
  const cUpd = headers.indexOf('Updated_At');

  const nowStr = formatNow_();
  let updated = false;
  for (let r = 1; r < stData.length; r++) {
    if (String(stData[r][cIdT] || '').trim() === idTailor && normText_(stData[r][cPen]) === namaPenjahit) {
      const cur = Number(stData[r][cReady]) || 0;
      stSheet.getRange(r + 1, cReady + 1).setValue(cur + qty);
      if (cUpd !== -1) stSheet.getRange(r + 1, cUpd + 1).setValue(nowStr);
      updated = true;
      break;
    }
  }

  if (!updated) {
    stSheet.appendRow([
      generateId_('STK'),
      idTailor,
      namaPenjahit,
      String(mt.Barcode_Kategori || '').trim(),
      String(mt.Nama_Kategori || '').trim(),
      String(mt.Warna || '-').trim(),
      String(mt.Ukuran || '-').trim(),
      qty,
      nowStr
    ]);
  }

  const detailItem = mt.Nama_Kategori + ' · ' + mt.Warna + ' · ' + mt.Ukuran;
  const logSheet = ss.getSheetByName(SHEETS.LOG_TAILOR);
  const logRow = [[
    generateId_('TLG'),
    trxId,
    nowStr,
    'INPUT_JAHIT',
    idTailor,
    namaPenjahit,
    detailItem,
    qty,
    0,
    Number(mt.Upah_Per_Pcs) || 0,
    0,
    'WIP',
    '',
    namaPenjahit,
    'Hasil jahit masuk stok penjahit'
  ]];
  logSheet.getRange(logSheet.getLastRow() + 1, 1, 1, 15).setValues(logRow);
  SpreadsheetApp.flush();

  markDashboardDirty_();
  return {
    success: true,
    message: '+' + qty + ' pcs ' + detailItem + ' dicatat pada penjahit ' + namaPenjahit + '.'
  };
}

function handleTransferTailorToWarehouse_(payload) {
  const trxId = String(payload.transaction_id || generateId_('TRF')).trim();
  const idTailor = String(payload.id_tailor || '').trim();
  const namaPenjahit = normText_(payload.nama_penjahit);
  const qtyLolos = Math.floor(Number(payload.qty_lolos) || 0);
  const qtyGagal = Math.floor(Number(payload.qty_gagal) || 0);
  const alasanGagal = String(payload.alasan_gagal || '').trim();
  const penerima = normText_(payload.nama_penerima) || 'ADMIN';

  if (!idTailor || !namaPenjahit || qtyLolos < 0 || qtyGagal < 0 || (qtyLolos + qtyGagal) <= 0) {
    return { success: false, message: 'Data setoran jahitan tidak lengkap atau jumlah tidak valid.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (isTransactionLogged_(ss, SHEETS.LOG_TAILOR, trxId)) {
    return { success: true, message: 'Setoran jahitan sudah diproses.' };
  }

  const mt = getSheetObjects_(ss.getSheetByName(SHEETS.MASTER_TAILOR)).find(function (m) {
    return String(m.ID_Tailor || '').trim() === idTailor;
  });
  if (!mt) {
    return { success: false, message: 'Data Master Tailor tidak ditemukan.' };
  }
  if (!parseBool_(mt.Status_Aktif, true)) {
    return { success: false, message: 'Varian kaos polos "' + mt.Nama_Kategori + ' · ' + mt.Warna + ' · ' + mt.Ukuran + '" berstatus NONAKTIF.' };
  }

  const stSheet = ss.getSheetByName(SHEETS.STOK_TAILOR);
  const stData = stSheet.getDataRange().getValues();
  const headers = stData[0].map(function (h) { return String(h || '').trim(); });
  const cIdT = headers.indexOf('ID_Tailor');
  const cPen = headers.indexOf('Nama_Penjahit');
  const cReady = headers.indexOf('Stok_Ready');
  const cUpd = headers.indexOf('Updated_At');

  const tailorRows = [];
  let totalTailorStok = 0;
  let polosRowIdx = -1;
  let polosStok = 0;

  for (let r = 1; r < stData.length; r++) {
    if (String(stData[r][cIdT] || '').trim() !== idTailor) continue;
    const pen = normText_(stData[r][cPen]);
    if (pen === namaPenjahit) {
      const q = Number(stData[r][cReady]) || 0;
      totalTailorStok += q;
      tailorRows.push({ rowIndex: r + 1, stok: q });
    } else if (pen === DEFAULT_GUDANG_POLOS && polosRowIdx === -1) {
      polosRowIdx = r + 1;
      polosStok = Number(stData[r][cReady]) || 0;
    }
  }

  const totalProses = qtyLolos + qtyGagal;
  if (tailorRows.length === 0 || totalTailorStok < totalProses) {
    return {
      success: false,
      message: 'Stok jahitan di ' + namaPenjahit + ' hanya ' + totalTailorStok + ' pcs (tidak cukup untuk ' + totalProses + ' pcs).'
    };
  }

  const nowStr = formatNow_();
  // 1. Kurangi stok dari penjahit
  let remToDeduct = totalProses;
  for (let k = 0; k < tailorRows.length && remToDeduct > 0; k++) {
    const deduct = Math.min(tailorRows[k].stok, remToDeduct);
    const sisa = tailorRows[k].stok - deduct;
    remToDeduct -= deduct;
    stSheet.getRange(tailorRows[k].rowIndex, cReady + 1).setValue(sisa);
    if (cUpd !== -1) stSheet.getRange(tailorRows[k].rowIndex, cUpd + 1).setValue(nowStr);
  }

  // 2. Tambahkan qtyLolos ke GUDANG POLOS
  if (polosRowIdx !== -1) {
    stSheet.getRange(polosRowIdx, cReady + 1).setValue(polosStok + qtyLolos);
    if (cUpd !== -1) stSheet.getRange(polosRowIdx, cUpd + 1).setValue(nowStr);
  } else {
    stSheet.appendRow([
      generateId_('STK'),
      idTailor,
      DEFAULT_GUDANG_POLOS,
      String(mt.Barcode_Kategori || '').trim(),
      String(mt.Nama_Kategori || '').trim(),
      String(mt.Warna || '-').trim(),
      String(mt.Ukuran || '-').trim(),
      qtyLolos,
      nowStr
    ]);
  }

  // 3. Catat log setoran & upah penjahit (UNPAID) secara atomik
  const upahSatuan = Number(mt.Upah_Per_Pcs) || 0;
  const totalUpah = qtyLolos * upahSatuan;
  const detailItem = mt.Nama_Kategori + ' · ' + mt.Warna + ' · ' + mt.Ukuran;
  const statusUpah = qtyLolos > 0 ? 'UNPAID' : 'N/A';

  const logSheet = ss.getSheetByName(SHEETS.LOG_TAILOR);
  const logRow = [[
    generateId_('TLG'),
    trxId,
    nowStr,
    'SETOR_GUDANG',
    idTailor,
    namaPenjahit,
    detailItem,
    qtyLolos,
    qtyGagal,
    upahSatuan,
    totalUpah,
    statusUpah,
    '',
    penerima,
    qtyGagal > 0 ? ('Rusak: ' + qtyGagal + ' pcs' + (alasanGagal ? ' (' + alasanGagal + ')' : '')) : 'Masuk Gudang Polos'
  ]];
  logSheet.getRange(logSheet.getLastRow() + 1, 1, 1, 15).setValues(logRow);
  SpreadsheetApp.flush();

  markDashboardDirty_();
  return {
    success: true,
    message: qtyLolos + ' pcs masuk Gudang Polos · Upah ' + namaPenjahit + ' (' + formatRp_(totalUpah) + ') tercatat.'
  };
}

function handleGeneratePaySlip_(payload) {
  const namaPenjahit = normText_(payload.nama_penjahit);
  const logIds = Array.isArray(payload.log_ids) ? payload.log_ids.map(function (x) { return String(x).trim(); }) : [];
  const adminName = normText_(payload.nama_admin) || 'ADMIN';
  const trxId = String(payload.transaction_id || '').trim();

  if (!namaPenjahit || !logIds.length) {
    return { success: false, message: 'Pilih penjahit dan minimal satu transaksi upah.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEETS.LOG_TAILOR);
  const data = logSheet.getDataRange().getValues();
  const headers = data[0].map(function (h) { return String(h || '').trim(); });

  const cLogId = headers.indexOf('Log_ID');
  const cTrxId = headers.indexOf('Transaction_ID');
  const cWaktu = headers.indexOf('Waktu');
  const cTipe = headers.indexOf('Tipe');
  const cPenjahit = headers.indexOf('Nama_Penjahit');
  const cDetail = headers.indexOf('Detail_Item');
  const cBagus = headers.indexOf('Qty_Bagus');
  const cUpahSat = headers.indexOf('Upah_Satuan');
  const cTotalUpah = headers.indexOf('Total_Upah');
  const cStatusUpah = headers.indexOf('Status_Upah');
  const cSlipId = headers.indexOf('Slip_ID');
  const cPelapor = headers.indexOf('Pelapor');

  // Idempotensi: jika transaction_id sudah tercatat pada BAYAR_UPAH, rekonstruksi & kembalikan slip yang sudah ada
  if (trxId) {
    for (let r = data.length - 1; r >= 1; r--) {
      if (normText_(data[r][cTipe]) === 'BAYAR_UPAH' && String(data[r][cTrxId] || '').trim() === trxId) {
        const existingSlipId = String(data[r][cSlipId] || '').trim();
        const existingPenjahit = normText_(data[r][cPenjahit]) || namaPenjahit;
        const existingTanggal = data[r][cWaktu] || '';
        const existingAdmin = String(data[r][cPelapor] || adminName).trim();
        const existingTotal = Number(data[r][cTotalUpah]) || 0;

        const existingRincian = [];
        for (let k = 1; k < data.length; k++) {
          if (normText_(data[k][cTipe]) === 'SETOR_GUDANG' && String(data[k][cSlipId] || '').trim() === existingSlipId) {
            existingRincian.push({
              log_id: String(data[k][cLogId] || '').trim(),
              detail_item: String(data[k][cDetail] || '').trim(),
              qty_bagus: Number(data[k][cBagus]) || 0,
              upah_satuan: Number(data[k][cUpahSat]) || 0,
              total_upah: Number(data[k][cTotalUpah]) || 0
            });
          }
        }

        return {
          success: true,
          message: 'Slip upah ' + existingSlipId + ' (' + formatRp_(existingTotal) + ') sudah diterbitkan sebelumnya.',
          data: {
            slip_id: existingSlipId,
            nama_penjahit: existingPenjahit,
            tanggal_bayar: existingTanggal,
            dibayar_oleh: existingAdmin,
            jumlah_item: existingRincian.length,
            total_upah: existingTotal,
            rincian: existingRincian
          }
        };
      }
    }
  }

  const slipId = 'SLIP-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd') + '-' + Math.random().toString(36).substring(2, 6).toUpperCase();
  const nowStr = formatNow_();
  const matchedRows = [];
  const rincian = [];
  let totalPaid = 0;

  for (let r = 1; r < data.length; r++) {
    const lid = String(data[r][cLogId] || '').trim();
    if (logIds.indexOf(lid) === -1) continue;
    if (normText_(data[r][cTipe]) !== 'SETOR_GUDANG') continue;
    if (normText_(data[r][cPenjahit]) !== namaPenjahit) continue;
    if (normText_(data[r][cStatusUpah]) !== 'UNPAID') continue;

    const upahRow = Number(data[r][cTotalUpah]) || 0;
    matchedRows.push(r + 1);
    totalPaid += upahRow;
    rincian.push({
      log_id: lid,
      detail_item: String(data[r][cDetail] || '').trim(),
      qty_bagus: Number(data[r][cBagus]) || 0,
      upah_satuan: Number(data[r][cUpahSat]) || 0,
      total_upah: upahRow
    });
  }

  if (!rincian.length) {
    return { success: false, message: 'Tidak ada tagihan upah UNPAID yang cocok untuk dilunasi.' };
  }

  // Perbarui status PAID & Slip_ID pada baris SETOR_GUDANG
  matchedRows.forEach(function (rowIdx) {
    logSheet.getRange(rowIdx, cStatusUpah + 1).setValue('PAID');
    logSheet.getRange(rowIdx, cSlipId + 1).setValue(slipId);
  });

  // Tulis baris BAYAR_UPAH secara atomik
  const bayarRow = [[
    generateId_('TLG'),
    trxId || slipId,
    nowStr,
    'BAYAR_UPAH',
    '-',
    namaPenjahit,
    'Pelunasan Slip ' + slipId + ' (' + rincian.length + ' setoran)',
    rincian.reduce(function (s, x) { return s + x.qty_bagus; }, 0),
    0,
    0,
    totalPaid,
    'PAID',
    slipId,
    adminName,
    'Slip Upah Lunas'
  ]];
  logSheet.getRange(logSheet.getLastRow() + 1, 1, 1, 15).setValues(bayarRow);
  SpreadsheetApp.flush();

  markDashboardDirty_();
  return {
    success: true,
    message: 'Slip upah ' + slipId + ' (' + formatRp_(totalPaid) + ') berhasil diterbitkan.',
    data: {
      slip_id: slipId,
      nama_penjahit: namaPenjahit,
      tanggal_bayar: nowStr,
      dibayar_oleh: adminName,
      jumlah_item: rincian.length,
      total_upah: totalPaid,
      rincian: rincian
    }
  };
}

// ============================================================================
// 7. OPNAME STOK FISIK GUDANG POLOS
// ============================================================================

function handleCreateOpnameSession_(payload) {
  const namaSesi = String(payload.nama_sesi || '').trim();
  const lokasi = normText_(payload.lokasi || DEFAULT_GUDANG_POLOS) || DEFAULT_GUDANG_POLOS;
  const pelapor = normText_(payload.nama_pelapor) || 'ADMIN';
  const trxId = String(payload.transaction_id || '').trim();

  if (!namaSesi) {
    return { success: false, message: 'Nama sesi opname wajib diisi.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sesiSheet = ss.getSheetByName(SHEETS.OPNAME_SESI);

  // Idempotensi berbasis transaction_id
  if (trxId) {
    try {
      const cachedSessionId = CacheService.getScriptCache().get('OPN_TRX_' + trxId);
      if (cachedSessionId) {
        const existingSesi = getSheetObjects_(sesiSheet).find(function (s) {
          return String(s.Session_ID || '').trim() === cachedSessionId;
        });
        if (existingSesi) {
          return {
            success: true,
            message: 'Sesi opname "' + existingSesi.Nama_Sesi + '" sudah dibuat.',
            data: {
              session_id: String(existingSesi.Session_ID || '').trim(),
              nama_sesi: String(existingSesi.Nama_Sesi || namaSesi).trim(),
              lokasi: String(existingSesi.Lokasi || lokasi).trim()
            }
          };
        }
      }
    } catch (e) { }
  }

  const sessionId = generateId_('OPN');
  const nowStr = formatNow_();

  const rowValues = [[
    sessionId,
    namaSesi,
    lokasi,
    'ACTIVE',
    pelapor,
    nowStr,
    '',
    ''
  ]];
  sesiSheet.getRange(sesiSheet.getLastRow() + 1, 1, 1, 8).setValues(rowValues);
  SpreadsheetApp.flush();

  if (trxId) {
    try {
      CacheService.getScriptCache().put('OPN_TRX_' + trxId, sessionId, 21600);
    } catch (e) { }
  }

  markDashboardDirty_();
  return {
    success: true,
    message: 'Sesi opname "' + namaSesi + '" berhasil dibuat.',
    data: {
      session_id: sessionId,
      nama_sesi: namaSesi,
      lokasi: lokasi
    }
  };
}

function handleGetTailorSku_(payload) {
  const sku = normText_(payload.sku);
  if (!sku) {
    return { success: false, message: 'Kode QR kaos polos wajib diisi.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mtRows = getSheetObjects_(ss.getSheetByName(SHEETS.MASTER_TAILOR)).filter(function (m) {
    return (normText_(m.Barcode_Kategori) === sku || normText_(m.ID_Tailor) === sku) && parseBool_(m.Status_Aktif, true);
  });

  if (!mtRows.length) {
    return { success: false, message: 'Kode kaos polos "' + sku + '" tidak ditemukan di Master Tailor.' };
  }

  const stRows = getSheetObjects_(ss.getSheetByName(SHEETS.STOK_TAILOR));
  const polosStokByTailor = {};
  stRows.forEach(function (s) {
    if (normText_(s.Nama_Penjahit) === DEFAULT_GUDANG_POLOS) {
      const idT = String(s.ID_Tailor || '').trim();
      polosStokByTailor[idT] = (polosStokByTailor[idT] || 0) + (Number(s.Stok_Ready) || 0);
    }
  });

  const variants = mtRows.map(function (m) {
    const idT = String(m.ID_Tailor || '').trim();
    return {
      id_tailor: idT,
      barcode_kategori: String(m.Barcode_Kategori || '').trim(),
      nama_kategori: String(m.Nama_Kategori || '').trim(),
      warna: String(m.Warna || '-').trim(),
      ukuran: String(m.Ukuran || '-').trim(),
      upah_per_pcs: Number(m.Upah_Per_Pcs) || 0,
      stok_gudang_polos: polosStokByTailor[idT] || 0
    };
  });

  return {
    success: true,
    data: {
      barcode_kategori: String(mtRows[0].Barcode_Kategori || sku).trim(),
      nama_kategori: String(mtRows[0].Nama_Kategori || sku).trim(),
      variants: variants
    }
  };
}

function handleSubmitOpnameCount_(payload) {
  const sessionId = String(payload.session_id || '').trim();
  const idTailor = String(payload.id_tailor || '').trim();
  const rawQty = payload.qty;
  const pelapor = normText_(payload.nama_pelapor) || 'ADMIN';

  if (
    !sessionId ||
    !idTailor ||
    rawQty === undefined ||
    rawQty === null ||
    String(rawQty).trim() === '' ||
    typeof rawQty === 'boolean'
  ) {
    return { success: false, message: 'Sesi opname, varian kaos polos, dan jumlah fisik wajib diisi.' };
  }

  const parsedNum = Number(rawQty);
  if (isNaN(parsedNum) || parsedNum < 0) {
    return { success: false, message: 'Jumlah fisik opname harus berupa angka >= 0.' };
  }
  const qtyFisik = Math.floor(parsedNum);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sesi = getSheetObjects_(ss.getSheetByName(SHEETS.OPNAME_SESI)).find(function (s) {
    return String(s.Session_ID || '').trim() === sessionId;
  });
  if (!sesi || normText_(sesi.Status) !== 'ACTIVE') {
    return { success: false, message: 'Sesi opname tidak ditemukan atau sudah dikunci.' };
  }

  const mt = getSheetObjects_(ss.getSheetByName(SHEETS.MASTER_TAILOR)).find(function (m) {
    return String(m.ID_Tailor || '').trim() === idTailor;
  });
  if (!mt) {
    return { success: false, message: 'Varian kaos polos tidak ditemukan.' };
  }

  const stRows = getSheetObjects_(ss.getSheetByName(SHEETS.STOK_TAILOR));
  let stokSistem = 0;
  stRows.forEach(function (s) {
    if (String(s.ID_Tailor || '').trim() === idTailor && normText_(s.Nama_Penjahit) === DEFAULT_GUDANG_POLOS) {
      stokSistem += (Number(s.Stok_Ready) || 0);
    }
  });
  const selisih = qtyFisik - stokSistem;
  const detailItem = mt.Nama_Kategori + ' · ' + mt.Warna + ' · ' + mt.Ukuran;

  const detSheet = ss.getSheetByName(SHEETS.OPNAME_DETAIL);
  const data = detSheet.getDataRange().getValues();
  const headers = data[0].map(function (h) { return String(h || '').trim(); });
  const cSesi = headers.indexOf('Session_ID');
  const cIdT = headers.indexOf('ID_Tailor');
  const cSys = headers.indexOf('Stok_Sistem');
  const cFis = headers.indexOf('Stok_Fisik');
  const cSel = headers.indexOf('Selisih');
  const cPel = headers.indexOf('Pelapor');
  const cWkt = headers.indexOf('Waktu');

  const nowStr = formatNow_();
  let updated = false;
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][cSesi] || '').trim() === sessionId && String(data[r][cIdT] || '').trim() === idTailor) {
      detSheet.getRange(r + 1, cSys + 1).setValue(stokSistem);
      detSheet.getRange(r + 1, cFis + 1).setValue(qtyFisik);
      detSheet.getRange(r + 1, cSel + 1).setValue(selisih);
      detSheet.getRange(r + 1, cPel + 1).setValue(pelapor);
      detSheet.getRange(r + 1, cWkt + 1).setValue(nowStr);
      updated = true;
      break;
    }
  }

  if (!updated) {
    detSheet.appendRow([
      generateId_('OPD'),
      sessionId,
      String(mt.Barcode_Kategori || '').trim(),
      idTailor,
      detailItem,
      stokSistem,
      qtyFisik,
      selisih,
      pelapor,
      nowStr
    ]);
  }
  SpreadsheetApp.flush();

  markDashboardDirty_();
  return {
    success: true,
    message: 'Tercatat: ' + detailItem + ' = ' + qtyFisik + ' pcs (selisih ' + (selisih >= 0 ? '+' + selisih : selisih) + ').'
  };
}

function handleApproveOpnameSession_(payload) {
  const sessionId = String(payload.session_id || '').trim();
  const pelapor = normText_(payload.nama_pelapor) || 'ADMIN';

  if (!sessionId) {
    return { success: false, message: 'Session ID opname wajib diisi.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sesiSheet = ss.getSheetByName(SHEETS.OPNAME_SESI);
  const sesiData = sesiSheet.getDataRange().getValues();
  const sHeaders = sesiData[0].map(function (h) { return String(h || '').trim(); });
  const cSid = sHeaders.indexOf('Session_ID');
  const cNamaSesi = sHeaders.indexOf('Nama_Sesi');
  const cStatus = sHeaders.indexOf('Status');
  const cBy = sHeaders.indexOf('Disetujui_Oleh');
  const cEnd = sHeaders.indexOf('Waktu_Selesai');

  let sesiRowIdx = -1;
  let namaSesi = sessionId;
  for (let r = 1; r < sesiData.length; r++) {
    if (String(sesiData[r][cSid] || '').trim() === sessionId) {
      if (normText_(sesiData[r][cStatus]) !== 'ACTIVE') {
        return { success: false, message: 'Sesi opname ini sudah dikunci sebelumnya.' };
      }
      sesiRowIdx = r + 1;
      namaSesi = String(sesiData[r][cNamaSesi] || sessionId).trim();
      break;
    }
  }

  if (sesiRowIdx === -1) {
    return { success: false, message: 'Sesi opname tidak ditemukan.' };
  }

  const details = getSheetObjects_(ss.getSheetByName(SHEETS.OPNAME_DETAIL)).filter(function (d) {
    return String(d.Session_ID || '').trim() === sessionId;
  });

  if (details.length === 0) {
    return { success: false, message: 'Tidak dapat mengunci sesi opname yang masih kosong (0 varian dihitung).' };
  }

  const stSheet = ss.getSheetByName(SHEETS.STOK_TAILOR);
  const stData = stSheet.getDataRange().getValues();
  const stHeaders = stData[0].map(function (h) { return String(h || '').trim(); });
  const cStIdT = stHeaders.indexOf('ID_Tailor');
  const cStPen = stHeaders.indexOf('Nama_Penjahit');
  const cStReady = stHeaders.indexOf('Stok_Ready');
  const cStUpd = stHeaders.indexOf('Updated_At');

  const mtRows = getSheetObjects_(ss.getSheetByName(SHEETS.MASTER_TAILOR));
  const nowStr = formatNow_();
  const opLogsBatch = [];

  details.forEach(function (d) {
    const idTailor = String(d.ID_Tailor || '').trim();
    const qtyFisik = Number(d.Stok_Fisik) || 0;

    let foundRow = -1;
    let currentSystemStok = 0;
    for (let r = 1; r < stData.length; r++) {
      if (String(stData[r][cStIdT] || '').trim() === idTailor && normText_(stData[r][cStPen]) === DEFAULT_GUDANG_POLOS) {
        if (foundRow === -1) foundRow = r + 1;
        currentSystemStok += (Number(stData[r][cStReady]) || 0);
      }
    }

    // Hitung selisih aktual terhadap saldo terkini saat sesi dikunci
    const actualSelisih = qtyFisik - currentSystemStok;

    if (foundRow !== -1) {
      stSheet.getRange(foundRow, cStReady + 1).setValue(qtyFisik);
      if (cStUpd !== -1) stSheet.getRange(foundRow, cStUpd + 1).setValue(nowStr);
      // Nolkan baris duplikat GUDANG POLOS lain untuk ID_Tailor yang sama jika ada
      for (let r = 1; r < stData.length; r++) {
        if (r + 1 !== foundRow && String(stData[r][cStIdT] || '').trim() === idTailor && normText_(stData[r][cStPen]) === DEFAULT_GUDANG_POLOS) {
          stSheet.getRange(r + 1, cStReady + 1).setValue(0);
          if (cStUpd !== -1) stSheet.getRange(r + 1, cStUpd + 1).setValue(nowStr);
        }
      }
    } else {
      const mt = mtRows.find(function (m) { return String(m.ID_Tailor || '').trim() === idTailor; });
      if (mt) {
        stSheet.appendRow([
          generateId_('STK'),
          idTailor,
          DEFAULT_GUDANG_POLOS,
          String(mt.Barcode_Kategori || '').trim(),
          String(mt.Nama_Kategori || '').trim(),
          String(mt.Warna || '-').trim(),
          String(mt.Ukuran || '-').trim(),
          qtyFisik,
          nowStr
        ]);
      }
    }

    opLogsBatch.push([
      generateId_('LOG'),
      sessionId,
      nowStr,
      'OPNAME',
      String(d.Barcode || '').trim(),
      idTailor,
      String(d.Detail_Item || idTailor).trim(),
      DEFAULT_GUDANG_POLOS,
      actualSelisih,
      0,
      0,
      0,
      0,
      pelapor,
      'Opname "' + namaSesi + '": fisik ' + qtyFisik + ' pcs (selisih aktual ' + (actualSelisih >= 0 ? '+' + actualSelisih : actualSelisih) + ')'
    ]);
  });

  if (opLogsBatch.length > 0) {
    const opSheet = ss.getSheetByName(SHEETS.OPERATION_LOG);
    opSheet.getRange(opSheet.getLastRow() + 1, 1, opLogsBatch.length, opLogsBatch[0].length)
      .setValues(opLogsBatch);
  }

  sesiSheet.getRange(sesiRowIdx, cStatus + 1).setValue('APPROVED');
  if (cBy !== -1) sesiSheet.getRange(sesiRowIdx, cBy + 1).setValue(pelapor);
  if (cEnd !== -1) sesiSheet.getRange(sesiRowIdx, cEnd + 1).setValue(nowStr);
  SpreadsheetApp.flush();

  markDashboardDirty_();
  return {
    success: true,
    message: 'Sesi opname "' + namaSesi + '" selesai & dikunci (' + details.length + ' varian disesuaikan).'
  };
}

// ============================================================================
// 8. EXECUTIVE DASHBOARD INTERAKTIF, FILTER REAL-TIME & RINGKASAN SEMUA TAB
// ============================================================================

const DASH_FILTER_CELLS = {
  PERIODE: 'B6',
  DATE_FROM: 'F6',
  DATE_TO: 'G6',
  KATEGORI: 'J6',
  PENJAHIT: 'N6',
  RESET_CB: 'P6'
};

function markDashboardDirty_() {
  try {
    PropertiesService.getScriptProperties().setProperty('IBS_DASHBOARD_DIRTY', '1');
  } catch (e) { }
}

function refreshDashboardIfDirty() {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('IBS_DASHBOARD_DIRTY') === '1') {
      refreshDashboard();
      props.deleteProperty('IBS_DASHBOARD_DIRTY');
    }
  } catch (e) { }
}

/**
 * Mengaktifkan pemicu otomatis (Time-Driven Trigger) setiap 5 menit
 * untuk memperbarui sheet DASHBOARD jika ada transaksi baru.
 */
function aktifkanAutoRefreshDashboard() {
  nonaktifkanAutoRefreshDashboard(true);
  ScriptApp.newTrigger('refreshDashboardIfDirty')
    .timeBased()
    .everyMinutes(5)
    .create();
  refreshDashboard();
  try {
    SpreadsheetApp.getUi().alert('Auto-Refresh Dashboard (setiap 5 menit) berhasil diaktifkan!');
  } catch (e) { }
}

function nonaktifkanAutoRefreshDashboard(silent) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === 'refreshDashboardIfDirty' || t.getHandlerFunction() === 'refreshDashboard') {
      ScriptApp.deleteTrigger(t);
    }
  });
  if (!silent) {
    try {
      SpreadsheetApp.getUi().alert('Auto-Refresh Dashboard dinonaktifkan.');
    } catch (e) { }
  }
}

/**
 * Alias kompatibilitas apabila tombol lama di sheet DASHBOARD masih tertaut ke buildDashboardSummary
 */
function buildDashboardSummary() {
  return refreshDashboard();
}

/**
 * Membangun tampilan sheet DASHBOARD eksekutif interaktif:
 * - Palet warna editorial khusus (Midnight Obsidian, Warm Alabaster, Deep Pine, Indigo Slate, Terracotta Copper)
 * - Baris Filter Interaktif Real-Time: Periode Hari/Tanggal, Kalender Dari-Sampai, Kategori Polos, Penjahit & Reset
 * - 4 Kartu KPI dengan Sparkline Dinamis
 * - 4 Grafik Analitik Interaktif yang mengikuti filter aktif
 * - Ringkasan menyeluruh dari SELURUH 12 TAB spreadsheet
 */
function refreshDashboard(options) {
  const opts = options || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureAllSheetsStructure_(ss);

  const locale = String(ss.getSpreadsheetLocale() || 'en_US').toLowerCase();
  const isIndoLocale = locale.indexOf('id') === 0;

  let dash = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!dash) {
    dash = ss.insertSheet(SHEETS.DASHBOARD, 0);
  }

  // Pastikan kapasitas kolom & baris mencukupi sebelum merender
  if (dash.getMaxColumns() < 28) {
    dash.insertColumnsAfter(dash.getMaxColumns(), 28 - dash.getMaxColumns());
  }
  if (dash.getMaxRows() < 90) {
    dash.insertRowsAfter(dash.getMaxRows(), 90 - dash.getMaxRows());
  }

  // --------------------------------------------------------------------------
  // 1. BACA STATE FILTER AKTIF DARI SHEET SEBELUM DIRESET
  // --------------------------------------------------------------------------
  let selPeriode = 'SEMUA WAKTU';
  let rawFrom = null;
  let rawTo = null;
  let selKategori = 'SEMUA KATEGORI';
  let selPenjahit = 'SEMUA PENJAHIT';

  if (!opts.reset) {
    try {
      const prevPeriode = normText_(dash.getRange(DASH_FILTER_CELLS.PERIODE).getValue());
      if (prevPeriode) selPeriode = prevPeriode;
      rawFrom = dash.getRange(DASH_FILTER_CELLS.DATE_FROM).getValue();
      rawTo = dash.getRange(DASH_FILTER_CELLS.DATE_TO).getValue();
      const prevKat = normText_(dash.getRange(DASH_FILTER_CELLS.KATEGORI).getValue());
      if (prevKat) selKategori = prevKat;
      const prevPen = normText_(dash.getRange(DASH_FILTER_CELLS.PENJAHIT).getValue());
      if (prevPen) selPenjahit = prevPen;
    } catch (e) { }
  }

  const dateWindow = resolveDashboardDateWindow_(selPeriode, rawFrom, rawTo);
  selPeriode = dateWindow.periode;

  // Hapus seluruh grafik & validasi lama agar bersih saat dirender ulang
  dash.getCharts().forEach(function (chart) {
    dash.removeChart(chart);
  });
  try {
    dash.showColumns(1, dash.getMaxColumns());
  } catch (e) { }
  dash.getRange(1, 1, dash.getMaxRows(), dash.getMaxColumns()).clearDataValidations();
  dash.clear();
  dash.setHiddenGridlines(true);

  // --------------------------------------------------------------------------
  // 2. MUAT & AGREGASI DATA DARI SELURUH 12 TAB OPERASIONAL
  // --------------------------------------------------------------------------
  const prodRows = getSheetObjects_(ss.getSheetByName(SHEETS.MASTER_PRODUK));
  const attrRows = getSheetObjects_(ss.getSheetByName(SHEETS.ATRIBUT_PRODUK));
  const varRows = getSheetObjects_(ss.getSheetByName(SHEETS.VARIASI_PRODUK));
  const stokVarRows = getSheetObjects_(ss.getSheetByName(SHEETS.STOK_VARIASI));
  const opRows = getSheetObjects_(ss.getSheetByName(SHEETS.OPERATION_LOG));
  const lokasiRows = getSheetObjects_(ss.getSheetByName(SHEETS.MASTER_LOKASI));
  const opSesiRows = getSheetObjects_(ss.getSheetByName(SHEETS.OPNAME_SESI));
  const opDetailRows = getSheetObjects_(ss.getSheetByName(SHEETS.OPNAME_DETAIL));
  const mtRows = getSheetObjects_(ss.getSheetByName(SHEETS.MASTER_TAILOR));
  const stRows = getSheetObjects_(ss.getSheetByName(SHEETS.STOK_TAILOR));
  const tlRows = getSheetObjects_(ss.getSheetByName(SHEETS.LOG_TAILOR));
  const mapRows = getSheetObjects_(ss.getSheetByName(SHEETS.MAPPING_POLOS_PRODUK));

  // Cari tanggal transaksi paling awal & paling akhir untuk mengisi kotak tanggal agar tidak kosong
  let earliestTs = null;
  let latestTs = null;
  opRows.concat(tlRows).forEach(function (r) {
    const rawW = getRowField_(r, ['Waktu', 'Timestamp', 'Tanggal', 'Waktu_Transaksi', 'Created_At']);
    const ts = parseDateSafe_(rawW);
    if (ts > 0) {
      if (earliestTs === null || ts < earliestTs) earliestTs = ts;
      if (latestTs === null || ts > latestTs) latestTs = ts;
    }
  });

  // Indeks Master Tailor & Daftar Pilihan Dropdown Kategori / Penjahit
  const mtById = {};
  const categorySet = {};
  const tailorSet = {};

  mtRows.forEach(function (m) {
    const idT = String(getRowField_(m, ['ID_Tailor', 'Kode_Tailor']) || '').trim();
    const katName = normText_(getRowField_(m, ['Nama_Kategori', 'Kategori', 'Barcode_Kategori']));
    if (katName) categorySet[katName] = true;
    mtById[idT] = {
      id_tailor: idT,
      barcode_kategori: normText_(getRowField_(m, ['Barcode_Kategori', 'SKU'])),
      nama_kategori: String(getRowField_(m, ['Nama_Kategori', 'Kategori']) || '').trim(),
      warna: String(getRowField_(m, ['Warna']) || '-').trim(),
      ukuran: String(getRowField_(m, ['Ukuran']) || '-').trim(),
      upah_per_pcs: parseNumSafe_(getRowField_(m, ['Upah_Per_Pcs', 'Upah'])),
      aktif: parseBool_(getRowField_(m, ['Status_Aktif', 'Aktif']), true)
    };
  });

  stRows.forEach(function (s) {
    const pen = normText_(getRowField_(s, ['Nama_Penjahit', 'Penjahit', 'Nama_Tailor']));
    if (isValidTailorName_(pen)) tailorSet[pen] = true;
    const kat = normText_(getRowField_(s, ['Nama_Kategori', 'Kategori']));
    if (kat) categorySet[kat] = true;
  });
  tlRows.forEach(function (t) {
    const pen = normText_(getRowField_(t, ['Nama_Penjahit', 'Penjahit', 'Nama_Tailor']));
    if (isValidTailorName_(pen)) tailorSet[pen] = true;
  });

  const kategoriOptions = ['SEMUA KATEGORI'].concat(Object.keys(categorySet).sort());
  const penjahitOptions = ['SEMUA PENJAHIT'].concat(Object.keys(tailorSet).sort());
  if (kategoriOptions.indexOf(selKategori) === -1) selKategori = 'SEMUA KATEGORI';
  if (penjahitOptions.indexOf(selPenjahit) === -1) selPenjahit = 'SEMUA PENJAHIT';

  // Indeks Mapping Produk -> Kaos Polos (mendukung spesifik ID_Variasi maupun wildcard '*')
  const mapByVarId = {};
  const mapByProdBarcode = {};
  const mappedVarSet = {};
  const mappedWildcardBarcodes = {};

  mapRows.forEach(function (mp) {
    const vid = String(mp.ID_Variasi || '').trim();
    const bcode = normText_(mp.Barcode_Produk);
    const idT = String(mp.ID_Tailor || '').trim();
    const mt = mtById[idT];
    const katNorm = mt ? normText_(mt.nama_kategori) : '';
    if (vid && vid !== '*') {
      mapByVarId[vid] = katNorm;
      mappedVarSet[vid] = true;
    } else if (vid === '*' && bcode) {
      mappedWildcardBarcodes[bcode] = true;
    }
    if (bcode && katNorm) mapByProdBarcode[bcode] = katNorm;
  });

  // A. Agregasi Gudang Polos & Penjahit WIP (STOK_TAILOR + MASTER_TAILOR)
  const polosQtyByTailorId = {};
  const tailorStats = {};
  let totalWipPcs = 0;
  let totalWipEstUpah = 0;

  stRows.forEach(function (s) {
    const idT = String(getRowField_(s, ['ID_Tailor', 'Kode_Tailor']) || '').trim();
    const pen = normText_(getRowField_(s, ['Nama_Penjahit', 'Penjahit', 'Nama_Tailor']));
    const qty = parseNumSafe_(getRowField_(s, ['Stok_Ready', 'Stok', 'Qty']));
    const mt = mtById[idT];
    const katNorm = normText_(getRowField_(s, ['Nama_Kategori', 'Kategori']) || (mt ? mt.nama_kategori : ''));
    const upahSat = mt ? mt.upah_per_pcs : 0;

    if (pen === DEFAULT_GUDANG_POLOS) {
      polosQtyByTailorId[idT] = (polosQtyByTailorId[idT] || 0) + qty;
    } else if (isValidTailorName_(pen)) {
      if (selKategori !== 'SEMUA KATEGORI' && katNorm !== selKategori) return;
      if (selPenjahit !== 'SEMUA PENJAHIT' && pen !== selPenjahit) return;

      if (!tailorStats[pen]) {
        tailorStats[pen] = { nama: pen, wip: 0, setorBagus: 0, setorRusak: 0, unpaid: 0, paid: 0 };
      }
      tailorStats[pen].wip += qty;
      if (qty > 0) {
        totalWipPcs += qty;
        totalWipEstUpah += qty * upahSat;
      }
    }
  });

  let totalPolosPcs = 0;
  let totalPolosNilai = 0;
  let activePolosVariants = 0;
  const catSummary = {};
  const lowPolosList = [];

  Object.keys(mtById).forEach(function (idT) {
    const m = mtById[idT];
    if (!m.aktif) return;
    const katNorm = normText_(m.nama_kategori);
    if (selKategori !== 'SEMUA KATEGORI' && katNorm !== selKategori) return;

    const qty = polosQtyByTailorId[idT] || 0;
    const nilai = qty * m.upah_per_pcs;
    totalPolosPcs += qty;
    totalPolosNilai += nilai;
    activePolosVariants += 1;

    const key = m.barcode_kategori || katNorm;
    if (!catSummary[key]) {
      catSummary[key] = {
        sku: m.barcode_kategori,
        nama: m.nama_kategori,
        varian: 0,
        qty: 0,
        nilai: 0
      };
    }
    catSummary[key].varian += 1;
    catSummary[key].qty += qty;
    catSummary[key].nilai += nilai;

    if (qty <= 5) {
      lowPolosList.push({
        sku: m.barcode_kategori,
        detail: m.nama_kategori + ' · ' + m.warna + ' · ' + m.ukuran,
        qty: qty
      });
    }
  });
  lowPolosList.sort(function (a, b) { return a.qty - b.qty; });

  // B. Agregasi Log Penjahit & Upah (LOG_TAILOR - terfilter tanggal, kategori, penjahit)
  let filteredSetorBagus = 0;
  let filteredSetorRusak = 0;
  let filteredPaidUpah = 0;
  let totalUnpaidUpah = 0;
  const unpaidTailorSet = {};
  const dailyBuckets = {};

  tlRows.forEach(function (r) {
    const tipe = normText_(getRowField_(r, ['Tipe', 'Tipe_Transaksi', 'Jenis', 'Type']));
    const rawWaktu = getRowField_(r, ['Waktu', 'Timestamp', 'Tanggal', 'Waktu_Transaksi']);
    const idT = String(getRowField_(r, ['ID_Tailor', 'Kode_Tailor']) || '').trim();
    const rawPen = normText_(getRowField_(r, ['Nama_Penjahit', 'Penjahit', 'Nama_Tailor']));
    const pen = isValidTailorName_(rawPen) ? rawPen : '';
    const mt = mtById[idT];
    const katNorm = mt ? normText_(mt.nama_kategori) : '';

    if (selKategori !== 'SEMUA KATEGORI' && katNorm !== selKategori) return;
    if (selPenjahit !== 'SEMUA PENJAHIT' && pen !== selPenjahit) return;

    if (pen && !tailorStats[pen]) {
      tailorStats[pen] = { nama: pen, wip: 0, setorBagus: 0, setorRusak: 0, unpaid: 0, paid: 0 };
    }

    const statusUpah = normText_(getRowField_(r, ['Status_Upah', 'Status_Bayar', 'Status']));
    const bagus = parseNumSafe_(getRowField_(r, ['Qty_Bagus', 'Qty_Lolos', 'Qty', 'Jumlah']));
    const rusak = parseNumSafe_(getRowField_(r, ['Qty_Rusak', 'Qty_Gagal', 'Rusak']));
    const upahSat = parseNumSafe_(getRowField_(r, ['Upah_Satuan'])) || (mt ? mt.upah_per_pcs : 0);
    const rawUpah = parseNumSafe_(getRowField_(r, ['Total_Upah', 'Upah', 'Nominal']));
    const totalUpahRow = rawUpah > 0 ? rawUpah : (bagus * upahSat);

    // Saldo tagihan UNPAID (akumulasi kewajiban aktif sesuai filter penjahit/kategori)
    if (tipe === 'SETOR_GUDANG' && statusUpah === 'UNPAID') {
      if (totalUpahRow > 0) {
        totalUnpaidUpah += totalUpahRow;
        if (pen && tailorStats[pen]) {
          tailorStats[pen].unpaid += totalUpahRow;
          unpaidTailorSet[pen] = true;
        }
      }
    }

    // Metrik berbasis periode waktu (Setoran QC & Pelunasan Slip)
    if (!isWithinDateWindow_(rawWaktu, dateWindow)) return;

    if (tipe === 'SETOR_GUDANG') {
      filteredSetorBagus += bagus;
      filteredSetorRusak += rusak;
      if (pen && tailorStats[pen]) {
        tailorStats[pen].setorBagus += bagus;
        tailorStats[pen].setorRusak += rusak;
      }

      const dKey = extractShortDateKey_(rawWaktu);
      if (!dailyBuckets[dKey]) dailyBuckets[dKey] = { label: dKey, keluar: 0, masuk: 0, omzet: 0, profit: 0 };
      dailyBuckets[dKey].masuk += bagus;
    } else if (tipe === 'BAYAR_UPAH') {
      filteredPaidUpah += totalUpahRow;
      if (pen && tailorStats[pen]) {
        tailorStats[pen].paid += totalUpahRow;
      }
    }
  });

  // C. Agregasi Penjualan / Barang Keluar & Performa Produk (OPERATION_LOG + MASTER_PRODUK + VARIASI_PRODUK)
  const prodPerfMap = {};
  let activeProdCount = 0;
  let inactiveProdCount = 0;

  prodRows.forEach(function (p) {
    const bcode = normText_(getRowField_(p, ['Barcode', 'Kode_Produk']));
    const isAct = parseBool_(getRowField_(p, ['Status_Aktif', 'Aktif']), true);
    if (isAct) activeProdCount++;
    else inactiveProdCount++;
    prodPerfMap[bcode] = {
      barcode: bcode,
      nama: String(getRowField_(p, ['Nama_Produk', 'Nama']) || bcode).trim(),
      aktif: isAct,
      varianAktif: 0,
      qtyKeluar: 0,
      omzet: 0,
      profit: 0
    };
  });

  let activeVarCount = 0;
  let inactiveVarCount = 0;
  let sumHpp = 0;
  let sumHargaJual = 0;
  let pricedVarCount = 0;
  const varPriceById = {};

  varRows.forEach(function (v) {
    const vid = String(getRowField_(v, ['ID_Variasi', 'ID_Varian']) || '').trim();
    const bcode = normText_(getRowField_(v, ['Barcode', 'Kode_Produk']));
    const isAct = parseBool_(getRowField_(v, ['Status_Aktif', 'Aktif']), true);
    if (isAct) {
      activeVarCount++;
      if (prodPerfMap[bcode]) prodPerfMap[bcode].varianAktif += 1;
      if (vid && mappedWildcardBarcodes[bcode]) {
        mappedVarSet[vid] = true;
      }
    } else {
      inactiveVarCount++;
    }
    const hpp = parseNumSafe_(getRowField_(v, ['HPP', 'Harga_Modal']));
    const hj = parseNumSafe_(getRowField_(v, ['Harga_Jual', 'Harga']));
    if (vid) varPriceById[vid] = { hpp: hpp, hj: hj };
    if (hj > 0) {
      sumHpp += hpp;
      sumHargaJual += hj;
      pricedVarCount++;
    }
  });

  let filteredKeluarPcs = 0;
  let filteredOmzet = 0;
  let filteredProfit = 0;

  opRows.forEach(function (r) {
    const tipe = normText_(getRowField_(r, ['Tipe', 'Tipe_Transaksi', 'Jenis', 'Type']));
    if (tipe !== 'KELUAR' && tipe !== 'PENJUALAN' && tipe !== 'OUT') return;

    const rawWaktu = getRowField_(r, ['Waktu', 'Timestamp', 'Tanggal', 'Waktu_Transaksi', 'Created_At']);
    if (!isWithinDateWindow_(rawWaktu, dateWindow)) return;

    const bcode = normText_(getRowField_(r, ['Barcode', 'Kode_Produk']));
    const vid = String(getRowField_(r, ['ID_Variasi', 'ID_Varian']) || '').trim();
    const mappedKat = mapByVarId[vid] || mapByProdBarcode[bcode] || '';
    if (selKategori !== 'SEMUA KATEGORI' && mappedKat !== selKategori) return;

    const qtyOut = Math.abs(parseNumSafe_(getRowField_(r, ['Qty', 'Jumlah', 'Qty_Keluar'])));
    const fallbackPrice = varPriceById[vid] || { hpp: 0, hj: 0 };
    const rowHj = parseNumSafe_(getRowField_(r, ['Harga_Jual', 'Harga'])) || fallbackPrice.hj;
    const rowHpp = parseNumSafe_(getRowField_(r, ['HPP', 'Harga_Modal'])) || fallbackPrice.hpp;
    const rawOmzet = parseNumSafe_(getRowField_(r, ['Omzet', 'Total_Omzet', 'Total_Harga', 'Subtotal']));
    const rawProfit = parseNumSafe_(getRowField_(r, ['Profit', 'Laba', 'Total_Profit']));

    const omzet = rawOmzet > 0 ? rawOmzet : (qtyOut * rowHj);
    const profit = rawProfit !== 0 ? rawProfit : (qtyOut * Math.max(0, rowHj - rowHpp));

    filteredKeluarPcs += qtyOut;
    filteredOmzet += omzet;
    filteredProfit += profit;

    if (!prodPerfMap[bcode]) {
      prodPerfMap[bcode] = {
        barcode: bcode,
        nama: String(getRowField_(r, ['Detail_Item', 'Nama_Produk', 'Keterangan']) || bcode).split('·')[0].trim(),
        aktif: true,
        varianAktif: 1,
        qtyKeluar: 0,
        omzet: 0,
        profit: 0
      };
    }
    prodPerfMap[bcode].qtyKeluar += qtyOut;
    prodPerfMap[bcode].omzet += omzet;
    prodPerfMap[bcode].profit += profit;

    const dKey = extractShortDateKey_(rawWaktu);
    if (!dailyBuckets[dKey]) dailyBuckets[dKey] = { label: dKey, keluar: 0, masuk: 0, omzet: 0, profit: 0 };
    dailyBuckets[dKey].keluar += qtyOut;
    dailyBuckets[dKey].omzet += omzet;
    dailyBuckets[dKey].profit += profit;
  });

  // D. Agregasi Tab Pendukung (OPNAME_SESI, OPNAME_DETAIL, MAPPING_POLOS_PRODUK, MASTER_LOKASI, STOK_VARIASI, ATRIBUT_PRODUK)
  let opnameActiveCount = 0;
  let opnameApprovedCount = 0;
  opSesiRows.forEach(function (s) {
    if (normText_(s.Status) === 'ACTIVE') opnameActiveCount++;
    else opnameApprovedCount++;
  });

  let opnameItemCount = 0;
  let opnameTotalSelisih = 0;
  opDetailRows.forEach(function (d) {
    if (!isWithinDateWindow_(d.Waktu, dateWindow)) return;
    opnameItemCount++;
    opnameTotalSelisih += (Number(d.Selisih) || 0);
  });

  const mappedVarCount = Object.keys(mappedVarSet).length;
  const activeLokasiList = lokasiRows.filter(function (l) {
    return parseBool_(l.Status_Aktif, true);
  }).map(function (l) {
    return String(l.Kode_Lokasi || '').trim();
  });

  // E. Aktivitas Gabungan Terfilter (OPERATION_LOG + LOG_TAILOR)
  const filteredActivities = buildFilteredActivities_(opRows, tlRows, dateWindow, selKategori, selPenjahit, mtById, mapByVarId, mapByProdBarcode, 10);

  // --------------------------------------------------------------------------
  // 3. KANVAS UTAMA, PALET WARNA EDITORIAL & TOP HERO HEADER (B2:P3)
  // --------------------------------------------------------------------------
  const PALETTE = {
    CANVAS: '#F7F6F2',        // Warm Alabaster
    CARD_BG: '#FFFFFF',       // Crisp Porcelain
    CARD_BORDER: '#E2DFD7',   // Warm Stone Border
    HEADER_BG: '#111C18',     // Midnight Obsidian Forest
    HEADER_ACCENT: '#D4AF37', // Champagne Gold
    TEXT_DARK: '#18221F',     // Deep Charcoal
    TEXT_MUTED: '#5C6B66',    // Muted Sage Slate
    PINE: '#1B4332',          // Deep Pine
    PINE_TINT: '#E8F3EE',
    INDIGO: '#1E3A5F',        // Indigo Slate
    INDIGO_TINT: '#EAF0F8',
    COPPER: '#9C4221',        // Terracotta Copper
    COPPER_TINT: '#FDF0EB',
    AMBER: '#92400E',         // Warm Ochre
    AMBER_TINT: '#FEF6E7'
  };

  const TOTAL_ROWS = 76;
  dash.getRange(1, 1, TOTAL_ROWS, 17)
    .setBackground(PALETTE.CANVAS)
    .setFontFamily('Roboto')
    .setVerticalAlignment('middle');

  // Top Hero Header Banner (B2:P3)
  dash.getRange('B2:P3')
    .setBackground(PALETTE.HEADER_BG)
    .setBorder(true, true, true, true, false, false, PALETTE.HEADER_ACCENT, SpreadsheetApp.BorderStyle.SOLID);

  dash.getRange('B2:B3').merge()
    .setValue('IB')
    .setBackground('#1F352D')
    .setFontColor(PALETTE.HEADER_ACCENT)
    .setFontWeight('bold')
    .setFontSize(15)
    .setHorizontalAlignment('center');

  dash.getRange('C2:L2').merge()
    .setValue('INA BASIC STORE  ·  EXECUTIVE COMMAND CENTER ' + APP_VERSION)
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(13)
    .setVerticalAlignment('bottom');

  dash.getRange('C3:L3').merge()
    .setValue('100% Made-to-Order & Gudang Polos  ·  Ringkasan Seluruh Tab & Analitik Interaktif')
    .setFontColor('#A3BEB3')
    .setFontSize(8.5)
    .setVerticalAlignment('top');

  // Live Status Badge di Kanan Header (N2:P3)
  const filterBadgeText = (selPeriode === 'SEMUA WAKTU' && selKategori === 'SEMUA KATEGORI' && selPenjahit === 'SEMUA PENJAHIT')
    ? '● LIVE · SEMUA DATA\nUpdate: ' + formatCompactTimestamp_(formatNow_())
    : '● LIVE · TERFILTER (' + selPeriode + ')\nUpdate: ' + formatCompactTimestamp_(formatNow_());

  dash.getRange('N2:P3').merge()
    .setValue(filterBadgeText)
    .setBackground('#1B2E27')
    .setFontColor('#D1FAE5')
    .setFontWeight('bold')
    .setFontSize(8)
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, false, false, '#2D5A47', SpreadsheetApp.BorderStyle.SOLID);

  // --------------------------------------------------------------------------
  // 4. INTERACTIVE FILTER & CONTROL DOCK (BARIS 5–6)
  // --------------------------------------------------------------------------
  dash.getRange('B5:P6')
    .setBackground(PALETTE.CARD_BG)
    .setBorder(true, true, true, true, false, false, PALETTE.CARD_BORDER, SpreadsheetApp.BorderStyle.SOLID);

  // Label Header Filter (Baris 5)
  dash.getRange('B5:D5').merge().setValue('📅  FILTER HARI / PERIODE').setFontColor(PALETTE.TEXT_MUTED).setFontWeight('bold').setFontSize(7.5);
  dash.getRange('F5').setValue('DARI TANGGAL').setFontColor(PALETTE.TEXT_MUTED).setFontWeight('bold').setFontSize(7.5);
  dash.getRange('G5:H5').merge().setValue('SAMPAI TANGGAL').setFontColor(PALETTE.TEXT_MUTED).setFontWeight('bold').setFontSize(7.5);
  dash.getRange('J5:L5').merge().setValue('👕  FILTER KATEGORI POLOS').setFontColor(PALETTE.TEXT_MUTED).setFontWeight('bold').setFontSize(7.5);
  dash.getRange('N5:O5').merge().setValue('🧵  FILTER PENJAHIT').setFontColor(PALETTE.TEXT_MUTED).setFontWeight('bold').setFontSize(7.5);
  dash.getRange('P5').setValue('RESET').setFontColor(PALETTE.COPPER).setFontWeight('bold').setFontSize(7.5).setHorizontalAlignment('center');

  // Input Controls (Baris 6) dengan Dropdown & Calendar Date Validation
  const defaultFromDate = dateWindow.fromDateObj || (earliestTs ? new Date(earliestTs) : new Date());
  const defaultToDate = dateWindow.toDateObj || (latestTs ? new Date(latestTs) : new Date());

  const periodeList = ['SEMUA WAKTU', 'HARI INI', 'KEMARIN', '7 HARI TERAKHIR', '30 HARI TERAKHIR', 'BULAN INI', 'CUSTOM TANGGAL'];
  const rulePeriode = SpreadsheetApp.newDataValidation().requireValueInList(periodeList, true).setAllowInvalid(false).build();
  dash.getRange(DASH_FILTER_CELLS.PERIODE)
    .setValue(selPeriode)
    .setDataValidation(rulePeriode)
    .setBackground('#EEF3F0')
    .setFontColor(PALETTE.PINE)
    .setFontWeight('bold')
    .setFontSize(9)
    .setBorder(true, true, true, true, false, false, '#B7CFC3', SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange('C6:D6').merge()
    .setValue(dateWindow.caption)
    .setFontColor(PALETTE.TEXT_MUTED)
    .setFontSize(8);

  const ruleDate = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();
  dash.getRange(DASH_FILTER_CELLS.DATE_FROM)
    .setValue(defaultFromDate)
    .setNumberFormat('yyyy-mm-dd')
    .setDataValidation(ruleDate)
    .setBackground('#F3F2ED')
    .setFontColor(PALETTE.TEXT_DARK)
    .setFontWeight('bold')
    .setFontSize(8.5)
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, false, false, '#D6D3C9', SpreadsheetApp.BorderStyle.SOLID);

  dash.getRange(DASH_FILTER_CELLS.DATE_TO)
    .setValue(defaultToDate)
    .setNumberFormat('yyyy-mm-dd')
    .setDataValidation(ruleDate)
    .setBackground('#F3F2ED')
    .setFontColor(PALETTE.TEXT_DARK)
    .setFontWeight('bold')
    .setFontSize(8.5)
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, false, false, '#D6D3C9', SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange('H6')
    .setValue('Klik 2x kalender')
    .setFontColor(PALETTE.TEXT_MUTED)
    .setFontSize(7.5)
    .setHorizontalAlignment('center');

  const ruleKategori = SpreadsheetApp.newDataValidation().requireValueInList(kategoriOptions, true).setAllowInvalid(false).build();
  dash.getRange(DASH_FILTER_CELLS.KATEGORI)
    .setValue(selKategori)
    .setDataValidation(ruleKategori)
    .setBackground('#EEF2F7')
    .setFontColor(PALETTE.INDIGO)
    .setFontWeight('bold')
    .setFontSize(8.5)
    .setBorder(true, true, true, true, false, false, '#C3D1E4', SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange('K6:L6').merge()
    .setValue(Object.keys(catSummary).length + ' Kategori Aktif')
    .setFontColor(PALETTE.TEXT_MUTED)
    .setFontSize(8);

  const rulePenjahit = SpreadsheetApp.newDataValidation().requireValueInList(penjahitOptions, true).setAllowInvalid(false).build();
  dash.getRange(DASH_FILTER_CELLS.PENJAHIT)
    .setValue(selPenjahit)
    .setDataValidation(rulePenjahit)
    .setBackground('#FDF4F0')
    .setFontColor(PALETTE.COPPER)
    .setFontWeight('bold')
    .setFontSize(8.5)
    .setBorder(true, true, true, true, false, false, '#E8C5B8', SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange('O6')
    .setValue(Object.keys(tailorStats).length + ' Penjahit')
    .setFontColor(PALETTE.TEXT_MUTED)
    .setFontSize(8);

  dash.getRange(DASH_FILTER_CELLS.RESET_CB)
    .insertCheckboxes()
    .setValue(false)
    .setBackground('#FDF0EB')
    .setHorizontalAlignment('center');

  // --------------------------------------------------------------------------
  // 5. BARIS 4 KARTU KPI EKSEKUTIF DENGAN SPARKLINE VISUAL (BARIS 8–11)
  // --------------------------------------------------------------------------
  const marginPct = filteredOmzet > 0 ? Math.round((filteredProfit / filteredOmzet) * 100) : 0;
  const totalQcProcessed = filteredSetorBagus + filteredSetorRusak;
  const qcPassPct = totalQcProcessed > 0 ? Math.round((filteredSetorBagus / totalQcProcessed) * 100) : 100;

  const kpiCards = [
    {
      cStart: 2, // B:D
      accent: PALETTE.PINE,
      tint: PALETTE.PINE_TINT,
      title: 'GUDANG POLOS SIAP PAKAI',
      mainVal: formatInt_(totalPolosPcs) + ' Pcs',
      subVal: Object.keys(catSummary).length + ' Kategori · ' + activePolosVariants + ' Varian · Nilai ' + formatRp_(totalPolosNilai),
      sparkFormula: buildSparklineBarFormula_(Math.max(1, totalPolosPcs), Math.max(10, totalPolosPcs + totalWipPcs), PALETTE.PINE, isIndoLocale)
    },
    {
      cStart: 6, // F:H
      accent: PALETTE.COPPER,
      tint: PALETTE.COPPER_TINT,
      title: 'OMZET & PROFIT (' + selPeriode + ')',
      mainVal: formatRp_(filteredOmzet),
      subVal: 'Profit ' + formatRp_(filteredProfit) + ' (' + marginPct + '%) · ' + formatInt_(filteredKeluarPcs) + ' Pcs Keluar',
      sparkFormula: buildSparklineBarFormula_(Math.max(1, filteredProfit), Math.max(10, filteredOmzet), PALETTE.COPPER, isIndoLocale)
    },
    {
      cStart: 10, // J:L
      accent: PALETTE.INDIGO,
      tint: PALETTE.INDIGO_TINT,
      title: 'PRODUKSI & STOK PENJAHIT (WIP)',
      mainVal: formatInt_(totalWipPcs) + ' Pcs WIP',
      subVal: 'Setor QC: ' + formatInt_(filteredSetorBagus) + ' Bagus · ' + filteredSetorRusak + ' Rusak (' + qcPassPct + '% Lolos)',
      sparkFormula: buildSparklineBarFormula_(qcPassPct, 100, PALETTE.INDIGO, isIndoLocale)
    },
    {
      cStart: 14, // N:P
      accent: PALETTE.AMBER,
      tint: PALETTE.AMBER_TINT,
      title: 'UPAH PENJAHIT & AUDIT STOK',
      mainVal: formatRp_(totalUnpaidUpah) + ' UNPAID',
      subVal: 'Lunas: ' + formatRp_(filteredPaidUpah) + ' · ' + lowPolosList.length + ' Varian Menipis (≤5)',
      sparkFormula: buildSparklineBarFormula_(Math.max(1, filteredPaidUpah), Math.max(10, filteredPaidUpah + totalUnpaidUpah), PALETTE.AMBER, isIndoLocale)
    }
  ];

  kpiCards.forEach(function (k) {
    dash.getRange(8, k.cStart, 4, 3)
      .setBackground(PALETTE.CARD_BG)
      .setBorder(true, true, true, true, false, false, PALETTE.CARD_BORDER, SpreadsheetApp.BorderStyle.SOLID);

    dash.getRange(8, k.cStart, 1, 3).merge()
      .setValue('  ' + k.title)
      .setBackground(k.tint)
      .setFontColor(k.accent)
      .setFontWeight('bold')
      .setFontSize(8);

    dash.getRange(9, k.cStart, 1, 3).merge()
      .setValue('  ' + k.mainVal)
      .setFontColor(PALETTE.TEXT_DARK)
      .setFontWeight('bold')
      .setFontSize(14);

    dash.getRange(10, k.cStart, 1, 3).merge()
      .setValue('  ' + k.subVal)
      .setFontColor(PALETTE.TEXT_MUTED)
      .setFontSize(7.8);

    dash.getRange(11, k.cStart, 1, 3).merge()
      .setFormula(k.sparkFormula);
  });

  // --------------------------------------------------------------------------
  // 6. SIAPKAN TABEL HELPER GRAFIK DI KOLOM S:X (WARNA MENYATU DENGAN KANVAS)
  // --------------------------------------------------------------------------
  const sortedDateKeys = Object.keys(dailyBuckets).sort();
  const trendSlice = sortedDateKeys.slice(-12).map(function (k) { return dailyBuckets[k]; });
  if (trendSlice.length === 0) {
    trendSlice.push({ label: 'Awal', keluar: 0, masuk: 0, omzet: 0, profit: 0 });
    trendSlice.push({ label: 'Periode Ini', keluar: 0, masuk: 0, omzet: 0, profit: 0 });
  } else if (trendSlice.length === 1) {
    trendSlice.unshift({ label: 'Sebelumnya', keluar: 0, masuk: 0, omzet: 0, profit: 0 });
  }

  // Tabel Helper 1 (S1:U14): Tren Harian Keluar & Masuk Polos
  const flowTable = [['Tanggal', 'Barang Keluar (Pcs)', 'Setor Polos Masuk (Pcs)']];
  // Tabel Helper 2 (V1:X14): Tren Omzet & Profit
  const financeTable = [['Tanggal', 'Omzet (Rp)', 'Profit (Rp)']];
  trendSlice.forEach(function (pt) {
    flowTable.push([String(pt.label), Number(pt.keluar) || 0, Number(pt.masuk) || 0]);
    financeTable.push([String(pt.label), Number(pt.omzet) || 0, Number(pt.profit) || 0]);
  });
  dash.getRange(1, 19, flowTable.length, 3).setValues(flowTable);
  dash.getRange(1, 22, financeTable.length, 3).setValues(financeTable);

  // Tabel Helper 3 (S17:T25): Komposisi Gudang Polos per Kategori (Donut)
  const catSorted = Object.values(catSummary).sort(function (a, b) { return b.qty - a.qty; });
  const hasPolosStock = catSorted.some(function (c) { return c.qty > 0; });
  const donutTable = [['Kategori', 'Stok Ready (Pcs)']];
  if (hasPolosStock) {
    catSorted.slice(0, 6).forEach(function (c) {
      if (c.qty > 0) donutTable.push([String(c.nama), Number(c.qty) || 0]);
    });
  } else {
    donutTable.push(['Belum Ada Stok', 1]);
  }
  dash.getRange(17, 19, donutTable.length, 2).setValues(donutTable);

  // Tabel Helper 4 (V17:W25): Top 6 Produk Terlaris (Horizontal Bar)
  const prodSorted = Object.values(prodPerfMap).sort(function (a, b) {
    if (b.qtyKeluar !== a.qtyKeluar) return b.qtyKeluar - a.qtyKeluar;
    return b.omzet - a.omzet;
  });
  const topProdTable = [['Produk', 'Terjual (Pcs)']];
  const topProds = prodSorted.slice(0, 6);
  if (topProds.length === 0) {
    topProdTable.push(['Belum Ada Penjualan', 0]);
  } else {
    topProds.forEach(function (p) {
      topProdTable.push([String(p.nama).slice(0, 24), Number(p.qtyKeluar) || 0]);
    });
  }
  dash.getRange(17, 22, topProdTable.length, 2).setValues(topProdTable);

  // Samarkan helper table ke warna kanvas (tanpa hideColumns agar seri grafik selalu dirender Google Sheets)
  dash.getRange(1, 19, 28, 7).setBackground(PALETTE.CANVAS).setFontColor(PALETTE.CANVAS);
  SpreadsheetApp.flush();

  // --------------------------------------------------------------------------
  // 7. AREA 4 GRAFIK ANALITIK UTAMA (BARIS 13–24 & BARIS 26–37)
  // --------------------------------------------------------------------------
  dash.getRange('B13:H24').setBackground(PALETTE.CARD_BG).setBorder(true, true, true, true, false, false, PALETTE.CARD_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange('B13:H13').merge().setValue('  📈  TREN OMZET & PROFIT PENJUALAN (' + selPeriode + ')').setBackground('#F3F2ED').setFontColor(PALETTE.TEXT_DARK).setFontWeight('bold').setFontSize(9);

  dash.getRange('J13:P24').setBackground(PALETTE.CARD_BG).setBorder(true, true, true, true, false, false, PALETTE.CARD_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange('J13:P13').merge().setValue('  📊  PERBANDINGAN ARUS BARANG: SETORAN MASUK POLOS VS BARANG KELUAR').setBackground('#F3F2ED').setFontColor(PALETTE.TEXT_DARK).setFontWeight('bold').setFontSize(9);

  dash.getRange('B26:H37').setBackground(PALETTE.CARD_BG).setBorder(true, true, true, true, false, false, PALETTE.CARD_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange('B26:H26').merge().setValue('  ◔  KOMPOSISI SALDO GUDANG POLOS PER KATEGORI').setBackground('#F3F2ED').setFontColor(PALETTE.TEXT_DARK).setFontWeight('bold').setFontSize(9);

  dash.getRange('J26:P37').setBackground(PALETTE.CARD_BG).setBorder(true, true, true, true, false, false, PALETTE.CARD_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange('J26:P26').merge().setValue('  🏆  TOP PRODUK TERLARIS (PCS KELUAR PADA PERIODE TERFILTER)').setBackground('#F3F2ED').setFontColor(PALETTE.TEXT_DARK).setFontWeight('bold').setFontSize(9);

  // Chart 1: Area Chart Omzet & Profit (B14)
  const chartFinance = dash.newChart()
    .asAreaChart()
    .addRange(dash.getRange(1, 22, financeTable.length, 3))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setNumHeaders(1)
    .setOption('useFirstColumnAsDomain', true)
    .setPosition(14, 2, 4, 4)
    .setOption('width', 575)
    .setOption('height', 232)
    .setOption('backgroundColor', '#FFFFFF')
    .setOption('colors', [PALETTE.PINE, PALETTE.COPPER])
    .setOption('areaOpacity', 0.16)
    .setOption('pointSize', 5)
    .setOption('lineWidth', 2.5)
    .setOption('legend', { position: 'top', textStyle: { fontSize: 8.5, color: PALETTE.TEXT_DARK } })
    .setOption('hAxis', { textStyle: { fontSize: 8, color: PALETTE.TEXT_MUTED } })
    .setOption('vAxis', { textStyle: { fontSize: 8, color: PALETTE.TEXT_MUTED }, minValue: 0, viewWindow: { min: 0 } })
    .setOption('chartArea', { left: 58, top: 34, width: '84%', height: '68%' })
    .build();
  dash.insertChart(chartFinance);

  // Chart 2: Column Chart Masuk Polos vs Keluar (J14)
  const chartFlow = dash.newChart()
    .asColumnChart()
    .addRange(dash.getRange(1, 19, flowTable.length, 3))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setNumHeaders(1)
    .setOption('useFirstColumnAsDomain', true)
    .setPosition(14, 10, 4, 4)
    .setOption('width', 575)
    .setOption('height', 232)
    .setOption('backgroundColor', '#FFFFFF')
    .setOption('colors', [PALETTE.INDIGO, PALETTE.HEADER_ACCENT])
    .setOption('legend', { position: 'top', textStyle: { fontSize: 8.5, color: PALETTE.TEXT_DARK } })
    .setOption('hAxis', { textStyle: { fontSize: 8, color: PALETTE.TEXT_MUTED } })
    .setOption('vAxis', { textStyle: { fontSize: 8, color: PALETTE.TEXT_MUTED }, minValue: 0, viewWindow: { min: 0 } })
    .setOption('chartArea', { left: 46, top: 34, width: '87%', height: '68%' })
    .build();
  dash.insertChart(chartFlow);

  // Chart 3: Donut Chart Komposisi Gudang Polos (B27)
  const chartDonut = dash.newChart()
    .asPieChart()
    .addRange(dash.getRange(17, 19, donutTable.length, 2))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setNumHeaders(1)
    .setPosition(27, 2, 4, 4)
    .setOption('width', 575)
    .setOption('height', 232)
    .setOption('backgroundColor', '#FFFFFF')
    .setOption('pieHole', 0.54)
    .setOption('colors', hasPolosStock ? [PALETTE.PINE, PALETTE.INDIGO, PALETTE.COPPER, PALETTE.HEADER_ACCENT, '#2D6A4F', '#4A6FA5'] : ['#CBD5E1'])
    .setOption('legend', { position: 'right', textStyle: { fontSize: 8.5, color: PALETTE.TEXT_DARK } })
    .setOption('pieSliceText', hasPolosStock ? 'percentage' : 'none')
    .setOption('chartArea', { left: 20, top: 16, width: '90%', height: '82%' })
    .build();
  dash.insertChart(chartDonut);

  // Chart 4: Horizontal Bar Chart Top Produk Terlaris (J27)
  const chartTopProd = dash.newChart()
    .asBarChart()
    .addRange(dash.getRange(17, 22, topProdTable.length, 2))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setNumHeaders(1)
    .setOption('useFirstColumnAsDomain', true)
    .setPosition(27, 10, 4, 4)
    .setOption('width', 575)
    .setOption('height', 232)
    .setOption('backgroundColor', '#FFFFFF')
    .setOption('legend', { position: 'none' })
    .setOption('colors', [PALETTE.PINE])
    .setOption('hAxis', { textStyle: { fontSize: 8, color: PALETTE.TEXT_MUTED }, minValue: 0, viewWindow: { min: 0 } })
    .setOption('vAxis', { textStyle: { fontSize: 8.5, color: PALETTE.TEXT_DARK } })
    .setOption('chartArea', { left: 135, top: 16, width: '70%', height: '76%' })
    .build();
  dash.insertChart(chartTopProd);

  // --------------------------------------------------------------------------
  // 8. RINGKASAN LENGKAP DARI SELURUH 12 TAB SPREADSHEET (BATCH RENDERING)
  // --------------------------------------------------------------------------
  // ==================== KOLOM KIRI (B:H) ====================
  let lRow = 39;

  // TABEL L1: Ringkasan Tab MASTER_TAILOR & STOK_TAILOR (Gudang Polos per Kategori + Sparkline)
  dash.getRange(lRow, 2, 1, 7).merge()
    .setValue('  1. TAB STOK_TAILOR & MASTER_TAILOR · SALDO GUDANG POLOS PER KATEGORI')
    .setBackground(PALETTE.PINE).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(9);
  lRow++;
  dash.getRange(lRow, 2, 1, 7)
    .setValues([['Kode QR', 'Nama Kaos Polos', 'Varian', 'Stok Ready', 'Visual Bar', 'Nilai Upah', 'Status']])
    .setBackground(PALETTE.PINE_TINT).setFontColor(PALETTE.PINE).setFontWeight('bold').setFontSize(8);
  lRow++;

  const maxCatQty = catSorted.reduce(function (mx, c) { return Math.max(mx, c.qty); }, 1);
  const catSlice = catSorted.slice(0, 8);
  if (catSlice.length === 0) {
    dash.getRange(lRow, 2, 1, 7).merge().setValue('Belum ada data kategori kaos polos.').setBackground('#FFFFFF').setFontColor(PALETTE.TEXT_MUTED).setFontSize(8.5);
    lRow++;
  } else {
    const nCat = catSlice.length;
    const vals = [];
    const bgs = [];
    const fgs = [];
    const fws = [];
    const aligns = [];
    const formulas = [];

    catSlice.forEach(function (c) {
      const stLabel = c.qty === 0 ? 'HABIS' : (c.qty <= 10 ? 'MENIPIS' : 'AMAN');
      const stBg = c.qty === 0 ? '#FEE2E2' : (c.qty <= 10 ? PALETTE.AMBER_TINT : PALETTE.PINE_TINT);
      const stFg = c.qty === 0 ? '#B91C1C' : (c.qty <= 10 ? PALETTE.AMBER : PALETTE.PINE);

      vals.push([c.sku, c.nama, c.varian + ' var', c.qty, '', c.nilai, stLabel]);
      bgs.push(['#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', stBg]);
      fgs.push([PALETTE.TEXT_DARK, PALETTE.TEXT_DARK, PALETTE.TEXT_DARK, PALETTE.TEXT_DARK, PALETTE.TEXT_DARK, PALETTE.TEXT_DARK, stFg]);
      fws.push(['normal', 'normal', 'normal', 'bold', 'normal', 'normal', 'bold']);
      aligns.push(['left', 'left', 'left', 'right', 'left', 'right', 'center']);
      formulas.push([buildSparklineBarFormula_(c.qty, maxCatQty, PALETTE.PINE, isIndoLocale)]);
    });

    dash.getRange(lRow, 2, nCat, 7)
      .setValues(vals)
      .setBackgrounds(bgs)
      .setFontColors(fgs)
      .setFontWeights(fws)
      .setHorizontalAlignments(aligns)
      .setFontSize(8.2)
      .setBorder(false, false, true, false, false, false, '#ECEAE2', SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(lRow, 6, nCat, 1).setFormulas(formulas);
    dash.getRange(lRow, 7, nCat, 1).setNumberFormat('"Rp "#,##0');
    dash.getRange(lRow, 8, nCat, 1).setFontSize(7.5);
    lRow += nCat;
  }

  // TABEL L2: Ringkasan Tab MASTER_PRODUK, VARIASI_PRODUK & OPERATION_LOG (Performa Produk Terfilter)
  lRow += 1;
  dash.getRange(lRow, 2, 1, 7).merge()
    .setValue('  2. TAB MASTER_PRODUK, VARIASI & OPERATION_LOG · PERFORMA PRODUK (' + selPeriode + ')')
    .setBackground(PALETTE.COPPER).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(9);
  lRow++;
  dash.getRange(lRow, 2, 1, 7)
    .setValues([['Barcode', 'Nama Produk Desain', 'Varian', 'Keluar (Pcs)', 'Omzet', 'Profit Kotor', 'Status']])
    .setBackground(PALETTE.COPPER_TINT).setFontColor(PALETTE.COPPER).setFontWeight('bold').setFontSize(8);
  lRow++;

  const prodSlice = prodSorted.slice(0, 8);
  if (prodSlice.length === 0) {
    dash.getRange(lRow, 2, 1, 7).merge().setValue('Belum ada produk terdaftar di MASTER_PRODUK.').setBackground('#FFFFFF').setFontColor(PALETTE.TEXT_MUTED).setFontSize(8.5);
    lRow++;
  } else {
    const nProd = prodSlice.length;
    const vals = [];
    const bgs = [];
    const fgs = [];
    const fws = [];
    const aligns = [];

    prodSlice.forEach(function (p) {
      vals.push([p.barcode, p.nama, p.varianAktif + ' var', p.qtyKeluar, p.omzet, p.profit, p.aktif ? 'AKTIF' : 'NONAKTIF']);
      bgs.push(['#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', p.aktif ? PALETTE.PINE_TINT : '#F1F5F9']);
      fgs.push([PALETTE.TEXT_DARK, PALETTE.TEXT_DARK, PALETTE.TEXT_DARK, PALETTE.TEXT_DARK, PALETTE.TEXT_DARK, PALETTE.PINE, p.aktif ? PALETTE.PINE : PALETTE.TEXT_MUTED]);
      fws.push(['normal', 'normal', 'normal', 'bold', 'normal', 'bold', 'bold']);
      aligns.push(['left', 'left', 'left', 'right', 'right', 'right', 'center']);
    });

    dash.getRange(lRow, 2, nProd, 7)
      .setValues(vals)
      .setBackgrounds(bgs)
      .setFontColors(fgs)
      .setFontWeights(fws)
      .setHorizontalAlignments(aligns)
      .setFontSize(8.2)
      .setBorder(false, false, true, false, false, false, '#ECEAE2', SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(lRow, 6, nProd, 2).setNumberFormat('"Rp "#,##0');
    dash.getRange(lRow, 8, nProd, 1).setFontSize(7.5);
    lRow += nProd;
  }

  // TABEL L3: Prioritas Jahit Ulang (Stok Menipis <= 5 Pcs)
  lRow += 1;
  dash.getRange(lRow, 2, 1, 7).merge()
    .setValue('  3. PRIORITAS JAHIT ULANG · VARIAN GUDANG POLOS MENIPIS (≤ 5 PCS)')
    .setBackground(PALETTE.AMBER).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(9);
  lRow++;
  dash.getRange(lRow, 2, 1, 7)
    .setValues([['Kode QR', 'Varian Kaos Polos (Warna · Ukuran)', '', '', 'Sisa Stok', 'Level Stok', 'Urgensi']])
    .setBackground(PALETTE.AMBER_TINT).setFontColor(PALETTE.AMBER).setFontWeight('bold').setFontSize(8);
  dash.getRange(lRow, 3, 1, 3).merge();
  lRow++;

  const lowSlice = lowPolosList.slice(0, 8);
  if (lowSlice.length === 0) {
    dash.getRange(lRow, 2, 1, 7).merge().setValue('✓ Seluruh varian kaos polos berada di atas batas aman (> 5 pcs).').setBackground('#FFFFFF').setFontColor(PALETTE.PINE).setFontWeight('bold').setFontSize(8.5);
    lRow++;
  } else {
    const nLow = lowSlice.length;
    const vals = [];
    const bgs = [];
    const fgs = [];
    const fws = [];
    const aligns = [];
    const formulas = [];

    lowSlice.forEach(function (lp) {
      const urgBg = lp.qty === 0 ? '#FEE2E2' : PALETTE.AMBER_TINT;
      const urgFg = lp.qty === 0 ? '#B91C1C' : PALETTE.AMBER;
      vals.push([lp.sku, lp.detail, '', '', lp.qty + ' pcs', '', lp.qty === 0 ? 'HABIS' : 'MENIPIS']);
      bgs.push(['#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', urgBg]);
      fgs.push([PALETTE.TEXT_DARK, PALETTE.TEXT_DARK, PALETTE.TEXT_DARK, PALETTE.TEXT_DARK, PALETTE.TEXT_DARK, PALETTE.TEXT_DARK, urgFg]);
      fws.push(['bold', 'normal', 'normal', 'normal', 'bold', 'normal', 'bold']);
      aligns.push(['left', 'left', 'left', 'left', 'right', 'left', 'center']);
      formulas.push([buildSparklineBarFormula_(Math.max(0.2, lp.qty), 5, lp.qty === 0 ? '#B91C1C' : PALETTE.AMBER, isIndoLocale)]);
    });

    dash.getRange(lRow, 2, nLow, 7)
      .setValues(vals)
      .setBackgrounds(bgs)
      .setFontColors(fgs)
      .setFontWeights(fws)
      .setHorizontalAlignments(aligns)
      .setFontSize(8.2)
      .setBorder(false, false, true, false, false, false, '#ECEAE2', SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(lRow, 7, nLow, 1).setFormulas(formulas);
    dash.getRange(lRow, 8, nLow, 1).setFontSize(7.5);
    for (let idx = 0; idx < nLow; idx++) {
      dash.getRange(lRow + idx, 3, 1, 3).merge();
    }
    lRow += nLow;
  }

  // ==================== KOLOM KANAN (J:P) ====================
  let rRow = 39;

  // TABEL R1: Ringkasan Tab LOG_TAILOR & STOK_TAILOR (Penjahit, QC & Upah)
  dash.getRange(rRow, 10, 1, 7).merge()
    .setValue('  4. TAB LOG_TAILOR & STOK_TAILOR · REKAP PENJAHIT, QC & UPAH')
    .setBackground(PALETTE.INDIGO).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(9);
  rRow++;
  dash.getRange(rRow, 10, 1, 7)
    .setValues([['Nama Penjahit', 'Stok WIP', 'Setor Bagus', 'Rusak', 'Tagihan UNPAID', 'Sudah Lunas', 'Status']])
    .setBackground(PALETTE.INDIGO_TINT).setFontColor(PALETTE.INDIGO).setFontWeight('bold').setFontSize(8);
  rRow++;

  const tailorArr = Object.values(tailorStats).sort(function (a, b) {
    if (b.unpaid !== a.unpaid) return b.unpaid - a.unpaid;
    return b.wip - a.wip;
  });

  const tailorSlice = tailorArr.slice(0, 8);
  if (tailorSlice.length === 0) {
    dash.getRange(rRow, 10, 1, 7).merge().setValue('Belum ada data aktivitas penjahit pada filter ini.').setBackground('#FFFFFF').setFontColor(PALETTE.TEXT_MUTED).setFontSize(8.5);
    rRow++;
  } else {
    const nTailor = tailorSlice.length;
    const vals = [];
    const bgs = [];
    const fgs = [];
    const fws = [];
    const aligns = [];

    tailorSlice.forEach(function (t) {
      const hasDebt = t.unpaid > 0;
      vals.push([t.nama, t.wip + ' pcs', t.setorBagus + ' pcs', t.setorRusak + ' pcs', t.unpaid, t.paid, hasDebt ? 'UNPAID' : 'LUNAS']);
      bgs.push(['#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', hasDebt ? PALETTE.COPPER_TINT : PALETTE.PINE_TINT]);
      fgs.push([PALETTE.TEXT_DARK, PALETTE.TEXT_DARK, PALETTE.TEXT_DARK, PALETTE.TEXT_DARK, hasDebt ? PALETTE.COPPER : PALETTE.TEXT_MUTED, PALETTE.PINE, hasDebt ? PALETTE.COPPER : PALETTE.PINE]);
      fws.push(['bold', 'normal', 'normal', 'normal', hasDebt ? 'bold' : 'normal', 'normal', 'bold']);
      aligns.push(['left', 'left', 'left', 'left', 'right', 'right', 'center']);
    });

    dash.getRange(rRow, 10, nTailor, 7)
      .setValues(vals)
      .setBackgrounds(bgs)
      .setFontColors(fgs)
      .setFontWeights(fws)
      .setHorizontalAlignments(aligns)
      .setFontSize(8.2)
      .setBorder(false, false, true, false, false, false, '#ECEAE2', SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(rRow, 14, nTailor, 2).setNumberFormat('"Rp "#,##0');
    dash.getRange(rRow, 16, nTailor, 1).setFontSize(7.5);
    rRow += nTailor;
  }

  // TABEL R2: Ringkasan Terpadu Seluruh Tab Master, Mapping, Opname & Lokasi
  rRow += 1;
  dash.getRange(rRow, 10, 1, 7).merge()
    .setValue('  5. RINGKASAN TAB MASTER, MAPPING POLOS, OPNAME FISIK & LOKASI')
    .setBackground(PALETTE.HEADER_BG).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(9);
  rRow++;
  dash.getRange(rRow, 10, 1, 7)
    .setValues([['Nama Tab / Modul', 'Ringkasan Data & Metrik Utama', '', '', '', 'Total / Nilai', 'Status']])
    .setBackground('#E6ECE9').setFontColor(PALETTE.HEADER_BG).setFontWeight('bold').setFontSize(8);
  dash.getRange(rRow, 11, 1, 4).merge();
  rRow++;

  const avgMargin = sumHargaJual > 0 ? Math.round(((sumHargaJual - sumHpp) / sumHargaJual) * 100) : 0;
  const mapCoveragePct = activeVarCount > 0 ? Math.min(100, Math.round((mappedVarCount / activeVarCount) * 100)) : 0;

  const allTabsAuditRows = [
    {
      tab: 'MASTER_PRODUK & ATRIBUT',
      desc: activeProdCount + ' Produk Aktif · ' + inactiveProdCount + ' Nonaktif · ' + attrRows.length + ' Definisi Atribut',
      val: (activeProdCount + inactiveProdCount) + ' Produk',
      badge: 'AKTIF',
      bg: PALETTE.PINE_TINT,
      fg: PALETTE.PINE
    },
    {
      tab: 'VARIASI_PRODUK & HARGA',
      desc: activeVarCount + ' Varian Aktif · ' + inactiveVarCount + ' Nonaktif · Rata-rata Margin ' + avgMargin + '%',
      val: activeVarCount + ' Varian',
      badge: pricedVarCount + '/' + (activeVarCount + inactiveVarCount) + ' HARGA',
      bg: PALETTE.INDIGO_TINT,
      fg: PALETTE.INDIGO
    },
    {
      tab: 'MAPPING_POLOS_PRODUK',
      desc: mappedVarCount + ' dari ' + activeVarCount + ' varian aktif sudah tertaut ke Gudang Polos (' + mapCoveragePct + '%)',
      val: mapRows.length + ' Tautan',
      badge: mapCoveragePct + '% TAUT',
      bg: mapCoveragePct >= 80 ? PALETTE.PINE_TINT : PALETTE.AMBER_TINT,
      fg: mapCoveragePct >= 80 ? PALETTE.PINE : PALETTE.AMBER
    },
    {
      tab: 'OPNAME_SESI & DETAIL',
      desc: opnameActiveCount + ' Sesi Aktif · ' + opnameApprovedCount + ' Selesai · Selisih Periode: ' + (opnameTotalSelisih >= 0 ? '+' + opnameTotalSelisih : opnameTotalSelisih) + ' pcs',
      val: opnameItemCount + ' Audit',
      badge: opnameActiveCount > 0 ? (opnameActiveCount + ' AKTIF') : 'TERKUNCI',
      bg: opnameActiveCount > 0 ? PALETTE.AMBER_TINT : PALETTE.PINE_TINT,
      fg: opnameActiveCount > 0 ? PALETTE.AMBER : PALETTE.PINE
    },
    {
      tab: 'MASTER_LOKASI & STOK_VAR',
      desc: 'Lokasi Aktif: ' + (activeLokasiList.join(', ') || '-') + ' · Stok Variasi (Arsip V5): ' + stokVarRows.length + ' baris',
      val: activeLokasiList.length + ' Lokasi',
      badge: 'SIAP',
      bg: PALETTE.PINE_TINT,
      fg: PALETTE.PINE
    }
  ];

  const nAudit = allTabsAuditRows.length;
  const auditVals = [];
  const auditBgs = [];
  const auditFgs = [];
  const auditFws = [];
  const auditAligns = [];

  allTabsAuditRows.forEach(function (item) {
    auditVals.push([item.tab, item.desc, '', '', '', item.val, item.badge]);
    auditBgs.push(['#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', item.bg]);
    auditFgs.push([PALETTE.TEXT_DARK, PALETTE.TEXT_MUTED, PALETTE.TEXT_MUTED, PALETTE.TEXT_MUTED, PALETTE.TEXT_MUTED, PALETTE.TEXT_DARK, item.fg]);
    auditFws.push(['bold', 'normal', 'normal', 'normal', 'normal', 'bold', 'bold']);
    auditAligns.push(['left', 'left', 'left', 'left', 'left', 'right', 'center']);
  });

  dash.getRange(rRow, 10, nAudit, 7)
    .setValues(auditVals)
    .setBackgrounds(auditBgs)
    .setFontColors(auditFgs)
    .setFontWeights(auditFws)
    .setHorizontalAlignments(auditAligns)
    .setFontSize(8)
    .setBorder(false, false, true, false, false, false, '#ECEAE2', SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange(rRow, 10, nAudit, 1).setFontSize(7.8);
  dash.getRange(rRow, 16, nAudit, 1).setFontSize(7.5);
  for (let idx = 0; idx < nAudit; idx++) {
    dash.getRange(rRow + idx, 11, 1, 4).merge();
  }
  rRow += nAudit;

  // TABEL R3: Log Aktivitas Operasional Terbaru (OPERATION_LOG + LOG_TAILOR Terfilter)
  rRow += 1;
  dash.getRange(rRow, 10, 1, 7).merge()
    .setValue('  6. LOG AKTIVITAS OPERASIONAL TERBARU (' + selPeriode + ')')
    .setBackground('#2D3A35').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(9);
  rRow++;
  dash.getRange(rRow, 10, 1, 7)
    .setValues([['Waktu', 'Petugas', 'Aktivitas', 'Rincian Transaksi', '', '', 'Modul']])
    .setBackground('#EAECEB').setFontColor(PALETTE.TEXT_DARK).setFontWeight('bold').setFontSize(8);
  dash.getRange(rRow, 13, 1, 3).merge();
  rRow++;

  const actSlice = filteredActivities.slice(0, 8);
  if (actSlice.length === 0) {
    dash.getRange(rRow, 10, 1, 7).merge().setValue('Tidak ada aktivitas transaksi pada rentang filter yang dipilih.').setBackground('#FFFFFF').setFontColor(PALETTE.TEXT_MUTED).setFontSize(8.5);
    rRow++;
  } else {
    const nAct = actSlice.length;
    const vals = [];
    const bgs = [];
    const fgs = [];
    const fws = [];
    const aligns = [];

    actSlice.forEach(function (act) {
      const badge = getActivityBadgeStyle_(act.category);
      vals.push([formatCompactTimestamp_(act.timestamp), act.actor, act.action_text, act.detail_text, '', '', badge.label]);
      bgs.push(['#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', '#FFFFFF', badge.bg]);
      fgs.push([PALETTE.TEXT_MUTED, PALETTE.TEXT_DARK, PALETTE.TEXT_DARK, PALETTE.TEXT_MUTED, PALETTE.TEXT_MUTED, PALETTE.TEXT_MUTED, badge.fg]);
      fws.push(['normal', 'bold', 'normal', 'normal', 'normal', 'normal', 'bold']);
      aligns.push(['left', 'left', 'left', 'left', 'left', 'left', 'center']);
    });

    dash.getRange(rRow, 10, nAct, 7)
      .setValues(vals)
      .setBackgrounds(bgs)
      .setFontColors(fgs)
      .setFontWeights(fws)
      .setHorizontalAlignments(aligns)
      .setFontSize(8)
      .setBorder(false, false, true, false, false, false, '#ECEAE2', SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(rRow, 10, nAct, 1).setFontSize(7.5);
    dash.getRange(rRow, 16, nAct, 1).setFontSize(7.5);
    for (let idx = 0; idx < nAct; idx++) {
      dash.getRange(rRow + idx, 13, 1, 3).merge();
    }
    rRow += nAct;
  }

  // --------------------------------------------------------------------------
  // 9. ATUR LEBAR KOLOM & TINGGI BARIS AGAR PRESISI DAN MUDAH DIBACA
  // --------------------------------------------------------------------------
  const colWidths = [
    16,  // A: Margin Kiri
    108, // B: KPI 1 / Kolom Tabel Kiri 1
    108, // C: KPI 1 / Kolom Tabel Kiri 2
    82,  // D: KPI 1 / Kolom Tabel Kiri 3
    14,  // E: Gutter 1
    104, // F: KPI 2 / Dari Tanggal / Kolom Tabel Kiri 4
    104, // G: KPI 2 / Sampai Tanggal / Kolom Tabel Kiri 5
    84,  // H: KPI 2 / Kolom Tabel Kiri 6
    18,  // I: Gutter Tengah (pemisah blok kiri & kanan)
    110, // J: KPI 3 / Filter Kategori / Kolom Tabel Kanan 1
    98,  // K: KPI 3 / Kolom Tabel Kanan 2
    92,  // L: KPI 3 / Kolom Tabel Kanan 3
    14,  // M: Gutter 3
    106, // N: KPI 4 / Filter Penjahit / Kolom Tabel Kanan 4
    98,  // O: KPI 4 / Kolom Tabel Kanan 5
    84,  // P: KPI 4 / Reset Checkbox / Kolom Tabel Kanan 6
    16   // Q: Margin Kanan
  ];

  colWidths.forEach(function (w, idx) {
    dash.setColumnWidth(idx + 1, w);
  });

  // Lebar kecil untuk kolom helper grafik (S:Z)
  for (let c = 18; c <= 26; c++) {
    dash.setColumnWidth(c, 12);
  }

  dash.setRowHeight(1, 12);
  dash.setRowHeight(2, 26);
  dash.setRowHeight(3, 22);
  dash.setRowHeight(4, 12);
  dash.setRowHeight(5, 18);
  dash.setRowHeight(6, 26);
  dash.setRowHeight(7, 12);
  dash.setRowHeight(8, 22);
  dash.setRowHeight(9, 28);
  dash.setRowHeight(10, 18);
  dash.setRowHeight(11, 14);
  dash.setRowHeight(12, 14);
  dash.setRowHeight(13, 26);
  for (let r = 14; r <= 24; r++) dash.setRowHeight(r, 22);
  dash.setRowHeight(25, 14);
  dash.setRowHeight(26, 26);
  for (let r = 27; r <= 37; r++) dash.setRowHeight(r, 22);
  dash.setRowHeight(38, 14);

  const maxDataRow = Math.max(lRow, rRow);
  for (let r = 39; r <= maxDataRow; r++) {
    dash.setRowHeight(r, 24);
  }
}

/**
 * Menghasilkan formula SPARKLINE tipe bar yang kompatibel dengan locale Indonesia (id_ID) maupun Inggris (en_US)
 */
function buildSparklineBarFormula_(val, maxVal, colorHex, isIndoLocale) {
  const v = Number(val) || 0;
  const m = Math.max(1, Number(maxVal) || 1);
  if (isIndoLocale) {
    return '=SPARKLINE(' + v + ';{"charttype"\\"bar";"max"\\' + m + ';"color1"\\"' + colorHex + '"})';
  }
  return '=SPARKLINE(' + v + ',{"charttype","bar";"max",' + m + ';"color1","' + colorHex + '"})';
}

/**
 * Menghitung rentang tanggal awal & akhir berdasarkan pilihan periode filter
 */
function resolveDashboardDateWindow_(periode, rawFrom, rawTo) {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0, 0);
  const todayEnd = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);

  const p = normText_(periode || 'SEMUA WAKTU');
  if (p === 'HARI INI') {
    return {
      periode: 'HARI INI',
      fromMs: todayStart.getTime(),
      toMs: todayEnd.getTime(),
      fromDateObj: todayStart,
      toDateObj: todayStart,
      caption: Utilities.formatDate(todayStart, tz, 'dd MMM yyyy')
    };
  }
  if (p === 'KEMARIN') {
    const yStart = new Date(todayStart.getTime() - 86400000);
    const yEnd = new Date(todayEnd.getTime() - 86400000);
    return {
      periode: 'KEMARIN',
      fromMs: yStart.getTime(),
      toMs: yEnd.getTime(),
      fromDateObj: yStart,
      toDateObj: yStart,
      caption: Utilities.formatDate(yStart, tz, 'dd MMM yyyy')
    };
  }
  if (p === '7 HARI TERAKHIR') {
    const d7 = new Date(todayStart.getTime() - 6 * 86400000);
    return {
      periode: '7 HARI TERAKHIR',
      fromMs: d7.getTime(),
      toMs: todayEnd.getTime(),
      fromDateObj: d7,
      toDateObj: todayStart,
      caption: Utilities.formatDate(d7, tz, 'dd MMM') + ' - ' + Utilities.formatDate(todayStart, tz, 'dd MMM yyyy')
    };
  }
  if (p === '30 HARI TERAKHIR') {
    const d30 = new Date(todayStart.getTime() - 29 * 86400000);
    return {
      periode: '30 HARI TERAKHIR',
      fromMs: d30.getTime(),
      toMs: todayEnd.getTime(),
      fromDateObj: d30,
      toDateObj: todayStart,
      caption: Utilities.formatDate(d30, tz, 'dd MMM') + ' - ' + Utilities.formatDate(todayStart, tz, 'dd MMM yyyy')
    };
  }
  if (p === 'BULAN INI') {
    const mStart = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);
    return {
      periode: 'BULAN INI',
      fromMs: mStart.getTime(),
      toMs: todayEnd.getTime(),
      fromDateObj: mStart,
      toDateObj: todayStart,
      caption: Utilities.formatDate(mStart, tz, 'MMMM yyyy')
    };
  }
  if (p === 'CUSTOM TANGGAL') {
    const fParsed = parseFilterDateBoundary_(rawFrom, false);
    const tParsed = parseFilterDateBoundary_(rawTo, true);
    if (fParsed || tParsed) {
      const fLabel = fParsed ? Utilities.formatDate(fParsed, tz, 'dd MMM yyyy') : 'Awal';
      const tLabel = tParsed ? Utilities.formatDate(tParsed, tz, 'dd MMM yyyy') : 'Sekarang';
      return {
        periode: 'CUSTOM TANGGAL',
        fromMs: fParsed ? fParsed.getTime() : null,
        toMs: tParsed ? tParsed.getTime() : null,
        fromDateObj: fParsed || '',
        toDateObj: tParsed || '',
        caption: fLabel + ' - ' + tLabel
      };
    }
  }

  return {
    periode: 'SEMUA WAKTU',
    fromMs: null,
    toMs: null,
    fromDateObj: '',
    toDateObj: '',
    caption: 'Seluruh Riwayat Transaksi'
  };
}

function parseFilterDateBoundary_(raw, isEndOfDay) {
  if (!raw) return null;
  let d = null;
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    d = new Date(raw.getFullYear(), raw.getMonth(), raw.getDate());
  } else {
    const s = String(raw).trim();
    const isoM = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    const idM = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (isoM) {
      d = new Date(Number(isoM[1]), Number(isoM[2]) - 1, Number(isoM[3]));
    } else if (idM) {
      d = new Date(Number(idM[3]), Number(idM[2]) - 1, Number(idM[1]));
    }
  }
  if (!d || isNaN(d.getTime())) return null;
  if (isEndOfDay) {
    d.setHours(23, 59, 59, 999);
  } else {
    d.setHours(0, 0, 0, 0);
  }
  return d;
}

function isWithinDateWindow_(rawTimestamp, dateWindow) {
  if (!dateWindow || (dateWindow.fromMs === null && dateWindow.toMs === null)) return true;
  const ts = parseDateSafe_(rawTimestamp);
  if (!ts) return false;
  if (dateWindow.fromMs !== null && ts < dateWindow.fromMs) return false;
  if (dateWindow.toMs !== null && ts > dateWindow.toMs) return false;
  return true;
}

function buildFilteredActivities_(opRows, tlRows, dateWindow, selKategori, selPenjahit, mtById, mapByVarId, mapByProdBarcode, limit) {
  const combined = [];

  if (selPenjahit === 'SEMUA PENJAHIT') {
    opRows.slice(-60).forEach(function (r) {
      const rawWaktu = getRowField_(r, ['Waktu', 'Timestamp', 'Tanggal', 'Waktu_Transaksi', 'Created_At']);
      if (!isWithinDateWindow_(rawWaktu, dateWindow)) return;
      const bcode = normText_(getRowField_(r, ['Barcode', 'Kode_Produk']));
      const vid = String(getRowField_(r, ['ID_Variasi', 'ID_Varian']) || '').trim();
      const mappedKat = mapByVarId[vid] || mapByProdBarcode[bcode] || '';
      if (selKategori !== 'SEMUA KATEGORI' && mappedKat !== selKategori) return;

      const tipe = normText_(getRowField_(r, ['Tipe', 'Tipe_Transaksi', 'Jenis', 'Type']));
      const qtyVal = Math.abs(parseNumSafe_(getRowField_(r, ['Qty', 'Jumlah', 'Qty_Keluar'])));
      let category = 'PRODUK';
      let actionText = tipe || 'TRANSAKSI';
      if (tipe === 'KELUAR' || tipe === 'PENJUALAN' || tipe === 'OUT') {
        category = 'KELUAR';
        actionText = 'Keluar ' + qtyVal + ' pcs';
      } else if (tipe === 'OPNAME') {
        category = 'OPNAME';
        actionText = 'Opname Polos';
      } else if (tipe === 'CREATE_PRODUCT') {
        actionText = 'Produk Baru';
      } else if (tipe === 'ADD_VARIANTS') {
        actionText = '+Variasi Baru';
      } else if (tipe === 'UPDATE_PRICING') {
        actionText = 'Update Harga';
      } else {
        actionText = 'Update Produk';
      }

      combined.push({
        category: category,
        actor: String(getRowField_(r, ['Pelapor', 'Admin', 'User']) || 'ADMIN').trim(),
        action_text: actionText,
        detail_text: String(getRowField_(r, ['Detail_Item', 'Nama_Produk', 'Keterangan']) || bcode).trim(),
        timestamp: rawWaktu || '',
        _ts: parseDateSafe_(rawWaktu)
      });
    });
  }

  tlRows.slice(-60).forEach(function (r) {
    const rawWaktu = getRowField_(r, ['Waktu', 'Timestamp', 'Tanggal', 'Waktu_Transaksi']);
    if (!isWithinDateWindow_(rawWaktu, dateWindow)) return;
    const rawPen = normText_(getRowField_(r, ['Nama_Penjahit', 'Penjahit', 'Nama_Tailor']));
    const pen = isValidTailorName_(rawPen) ? rawPen : '';
    const idT = String(getRowField_(r, ['ID_Tailor', 'Kode_Tailor']) || '').trim();
    const mt = mtById[idT];
    const katNorm = mt ? normText_(mt.nama_kategori) : '';

    if (selKategori !== 'SEMUA KATEGORI' && katNorm !== selKategori) return;
    if (selPenjahit !== 'SEMUA PENJAHIT' && pen !== selPenjahit) return;

    const tipe = normText_(getRowField_(r, ['Tipe', 'Tipe_Transaksi', 'Jenis', 'Type']));
    const qtyBagus = parseNumSafe_(getRowField_(r, ['Qty_Bagus', 'Qty_Lolos', 'Qty', 'Jumlah']));
    const totalUpah = parseNumSafe_(getRowField_(r, ['Total_Upah', 'Upah', 'Nominal']));
    let category = 'SETOR_GUDANG';
    let actionText = tipe || 'JAHIT';
    if (tipe === 'SETOR_GUDANG') {
      category = 'SETOR_GUDANG';
      actionText = 'Setor +' + qtyBagus + ' pcs';
    } else if (tipe === 'INPUT_JAHIT') {
      category = 'SETOR_GUDANG';
      actionText = 'Jahit +' + qtyBagus + ' pcs';
    } else if (tipe === 'BAYAR_UPAH') {
      category = 'UPAH';
      actionText = 'Lunas ' + formatRp_(totalUpah);
    }

    combined.push({
      category: category,
      actor: String(getRowField_(r, ['Pelapor', 'Admin']) || pen || 'ADMIN').trim(),
      action_text: actionText,
      detail_text: String(getRowField_(r, ['Detail_Item', 'Nama_Kategori', 'Keterangan']) || pen).trim(),
      timestamp: rawWaktu || '',
      _ts: parseDateSafe_(rawWaktu)
    });
  });

  combined.sort(function (a, b) {
    return b._ts - a._ts;
  });

  return combined.slice(0, limit || 10);
}

/**
 * Mengambil nilai kolom dari objek baris secara fleksibel terhadap variasi nama header
 */
function getRowField_(rowObj, candidates) {
  if (!rowObj) return '';
  for (let i = 0; i < candidates.length; i++) {
    const key = candidates[i];
    if (rowObj[key] !== undefined && rowObj[key] !== null && rowObj[key] !== '') {
      return rowObj[key];
    }
  }
  const rowKeys = Object.keys(rowObj);
  for (let i = 0; i < candidates.length; i++) {
    const targetNorm = normText_(candidates[i]);
    for (let j = 0; j < rowKeys.length; j++) {
      if (normText_(rowKeys[j]) === targetNorm && rowObj[rowKeys[j]] !== '' && rowObj[rowKeys[j]] !== null) {
        return rowObj[rowKeys[j]];
      }
    }
  }
  return '';
}

/**
 * Mengonversi angka maupun teks berformat Rupiah/ribuan ("Rp 150.000,50" / "Rp 150.000,-") menjadi Number murni
 */
function parseNumSafe_(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (val === undefined || val === null || val === '') return 0;
  let s = String(val).trim();
  if (!s) return 0;
  s = s.replace(/rp/gi, '').replace(/,-$/g, '').replace(/\s+/g, '');
  if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) {
    s = s.replace(/,/g, '');
  } else if (/^-?\d+,\d+$/.test(s)) {
    s = s.replace(',', '.');
  }
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

/**
 * Memastikan nama penjahit valid (bukan kode ID TLR-xxx atau tipe log INPUT_JAHIT)
 */
function isValidTailorName_(name) {
  const n = normText_(name);
  if (!n || n === '-' || n === DEFAULT_GUDANG_POLOS || n === 'GUDANG UTAMA' || n === 'ADMIN' || n === 'SISTEM') {
    return false;
  }
  if (/^(TLR|STK|TLG|LOG|OPN|OPD|MAP|BLK|TRF|JHT|SLIP)-/.test(n)) {
    return false;
  }
  const reserved = ['INPUT_JAHIT', 'SETOR_GUDANG', 'BAYAR_UPAH', 'WIP', 'UNPAID', 'PAID', 'N/A', 'KELUAR', 'OPNAME'];
  if (reserved.indexOf(n) !== -1) {
    return false;
  }
  return true;
}

function extractShortDateKey_(raw) {
  if (!raw) return 'Hari Ini';
  if (raw instanceof Date) {
    return Utilities.formatDate(raw, Session.getScriptTimeZone(), 'MM-dd');
  }
  const s = String(raw).trim();
  const m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[2] + '-' + m[3];
  const mId = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (mId) return mId[2].padStart(2, '0') + '-' + mId[1].padStart(2, '0');
  return s.slice(0, 10) || 'Hari Ini';
}

function formatCompactTimestamp_(raw) {
  if (!raw) return '-';
  if (raw instanceof Date) {
    return Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  }
  const s = String(raw).trim();
  return s.length >= 16 ? s.slice(0, 16) : s;
}

function getActivityBadgeStyle_(category) {
  const cat = normText_(category);
  if (cat === 'KELUAR') {
    return { label: 'Keluar', bg: '#E8F3EE', fg: '#1B4332' };
  }
  if (cat === 'SETOR_GUDANG') {
    return { label: 'Setoran', bg: '#EAF0F8', fg: '#1E3A5F' };
  }
  if (cat === 'UPAH') {
    return { label: 'Lunas', bg: '#FDF0EB', fg: '#9C4221' };
  }
  if (cat === 'OPNAME') {
    return { label: 'Opname', bg: '#FEF6E7', fg: '#92400E' };
  }
  return { label: 'Produk', bg: '#EAECEB', fg: '#18221F' };
}

function formatInt_(num) {
  const n = Math.round(Number(num) || 0);
  return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

// ============================================================================
// 9. UTILITAS UMUM & NORMALISASI
// ============================================================================

function getSheetObjects_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function (h) { return String(h || '').trim(); });
  const list = [];

  for (let r = 1; r < values.length; r++) {
    const obj = {};
    let hasAny = false;
    for (let c = 0; c < headers.length; c++) {
      const h = headers[c];
      if (!h) continue;
      obj[h] = values[r][c];
      if (values[r][c] !== '' && values[r][c] !== null && values[r][c] !== undefined) {
        hasAny = true;
      }
    }
    if (hasAny) list.push(obj);
  }
  return list;
}

function appendOperationLog_(ss, entry) {
  const sheet = ss.getSheetByName(SHEETS.OPERATION_LOG);
  const row = [[
    generateId_('LOG'),
    entry.transaction_id || generateId_('TRX'),
    formatNow_(),
    entry.tipe || 'INFO',
    entry.barcode || '-',
    entry.id_variasi || '-',
    entry.detail_item || '-',
    entry.lokasi || DEFAULT_GUDANG_POLOS,
    Number(entry.qty) || 0,
    Number(entry.hpp) || 0,
    Number(entry.harga_jual) || 0,
    Number(entry.omzet) || 0,
    Number(entry.profit) || 0,
    entry.pelapor || 'ADMIN',
    entry.keterangan || ''
  ]];
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, 15).setValues(row);
}

function isTransactionLogged_(ss, sheetName, transactionId) {
  if (!transactionId) return false;
  const rows = getSheetObjects_(ss.getSheetByName(sheetName));
  return rows.some(function (r) {
    return String(r.Transaction_ID || '').trim() === String(transactionId).trim();
  });
}

function normText_(val) {
  return String(val === undefined || val === null ? '' : val)
    .trim()
    .toUpperCase()
    .replace(/\s+/g, ' ');
}

function normalizeMatchToken_(val) {
  const str = normText_(val).replace(/[_-]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!str || str === '-' || str === 'STANDAR' || str === 'DEFAULT') return '-';
  if (str === 'ALL SIZE' || str === 'ALLSIZE' || str === 'FREE SIZE' || str === 'ONE SIZE') return 'ALLSIZE';
  if (str === '2XL') return 'XXL';
  if (str === '3XL') return 'XXXL';
  return str;
}

function extractColorSizeFromVariant_(attrsObj, labelVariasi) {
  let warna = '';
  let ukuran = '';
  if (attrsObj && typeof attrsObj === 'object') {
    Object.keys(attrsObj).forEach(function (k) {
      const kl = k.toLowerCase();
      if (!warna && (kl.indexOf('warna') !== -1 || kl.indexOf('color') !== -1)) {
        warna = String(attrsObj[k] || '').trim();
      } else if (!ukuran && (kl.indexOf('ukuran') !== -1 || kl.indexOf('size') !== -1)) {
        ukuran = String(attrsObj[k] || '').trim();
      }
    });
    const keys = Object.keys(attrsObj);
    if (!warna && !ukuran && keys.length >= 2) {
      warna = String(attrsObj[keys[0]] || '').trim();
      ukuran = String(attrsObj[keys[1]] || '').trim();
    } else if (!warna && !ukuran && keys.length === 1) {
      const singleVal = String(attrsObj[keys[0]] || '').trim();
      if (/^(XS|S|M|L|XL|XXL|2XL|3XL|XXXL|ALL\s*SIZE)$/i.test(singleVal)) {
        ukuran = singleVal;
      } else {
        warna = singleVal;
      }
    }
  }

  if (!warna && !ukuran && labelVariasi) {
    const parts = String(labelVariasi).split(/·|\||\/|-/).map(function (s) { return s.trim(); }).filter(Boolean);
    if (parts.length >= 2) {
      warna = parts[0];
      ukuran = parts[1];
    } else if (parts.length === 1 && parts[0].toLowerCase() !== 'standar') {
      if (/^(XS|S|M|L|XL|XXL|2XL|3XL|XXXL|ALL\s*SIZE)$/i.test(parts[0])) {
        ukuran = parts[0];
      } else {
        warna = parts[0];
      }
    }
  }

  return {
    warna: warna || '-',
    ukuran: ukuran || '-'
  };
}

function parseBool_(val, defaultVal) {
  if (val === undefined || val === null || val === '') return defaultVal;
  if (typeof val === 'boolean') return val;
  const s = String(val).trim().toLowerCase();
  if (s === 'true' || s === '1' || s === 'ya' || s === 'yes' || s === 'aktif') return true;
  if (s === 'false' || s === '0' || s === 'tidak' || s === 'no' || s === 'nonaktif') return false;
  return defaultVal;
}

let __idSeqCounter = 0;

function generateId_(prefix) {
  __idSeqCounter = (__idSeqCounter + 1) % 46656;
  const seqStr = __idSeqCounter.toString(36).toUpperCase().padStart(3, '0');
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmmss');
  const rnd = Math.random().toString(36).substring(2, 6).toUpperCase();
  return prefix + '-' + ts + '-' + seqStr + rnd;
}

function formatNow_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function parseDateSafe_(raw) {
  if (!raw) return 0;
  if (raw instanceof Date) return isNaN(raw.getTime()) ? 0 : raw.getTime();
  const s = String(raw).trim();
  if (!s) return 0;

  // 1. Format standar sistem: YYYY-MM-DD[ HH:mm[:ss]] dievaluasi dalam zona waktu lokal skrip
  const isoLocal = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
  if (isoLocal) {
    const d = new Date(
      Number(isoLocal[1]),
      Number(isoLocal[2]) - 1,
      Number(isoLocal[3]),
      Number(isoLocal[4] || 0),
      Number(isoLocal[5] || 0),
      Number(isoLocal[6] || 0)
    );
    return isNaN(d.getTime()) ? 0 : d.getTime();
  }

  // 2. Format lokal Indonesia: DD/MM/YYYY atau DD-MM-YYYY
  const idLocal = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[\sT](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
  if (idLocal) {
    const d = new Date(
      Number(idLocal[3]),
      Number(idLocal[2]) - 1,
      Number(idLocal[1]),
      Number(idLocal[4] || 0),
      Number(idLocal[5] || 0),
      Number(idLocal[6] || 0)
    );
    return isNaN(d.getTime()) ? 0 : d.getTime();
  }

  const fallback = new Date(s);
  return isNaN(fallback.getTime()) ? 0 : fallback.getTime();
}

function formatRp_(num) {
  const n = Math.round(Number(num) || 0);
  return 'Rp ' + n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}