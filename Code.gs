/**
 * INA BASIC STORE WAREHOUSE - Backend API (Google Apps Script)
 * Deploy: Deploy > New deployment > Web app
 *   - Execute as: Me
 *   - Who has access: Anyone
 */

// ===================== KONFIGURASI =====================
var SHEET_MASTER = 'MASTER_PRODUK';
var SHEET_STOK   = 'STOK_LOKASI';
var SHEET_LOG    = 'OPERATION_LOG';

var HEADERS = {
  MASTER_PRODUK: ['SKU', 'Barcode', 'Nama', 'Status_Aktif'],
  STOK_LOKASI:   ['SKU', 'Lokasi_Rak', 'Kuantitas'],
  OPERATION_LOG: ['Timestamp', 'ID_Transaksi', 'Aktivitas', 'SKU', 'Lokasi_Asal', 'Lokasi_Tujuan', 'Qty', 'Status']
};

var LOCK_TIMEOUT_MS = 10000; // antrean 10 detik
var ACTIVITIES = { 'masuk': 'Masuk', 'keluar': 'Keluar', 'mutasi': 'Mutasi', 'opname': 'Opname' };

// ===================== SETUP =====================
/**
 * Jalankan sekali secara manual dari editor Apps Script.
 * Membuat 3 sheet beserta header jika belum ada.
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(HEADERS).forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);

    var headers = HEADERS[name];
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
      sheet.setFrozenRows(1);
    }
  });

  // Paksa kolom identitas jadi teks agar barcode/SKU angka tidak berubah format
  // (mis. "00123" tidak menjadi 123, barcode panjang tidak jadi notasi ilmiah).
  ss.getSheetByName(SHEET_MASTER).getRange('A:B').setNumberFormat('@');
  ss.getSheetByName(SHEET_STOK).getRange('A:B').setNumberFormat('@');
  ss.getSheetByName(SHEET_LOG).getRange('B:B').setNumberFormat('@');
  ss.getSheetByName(SHEET_LOG).getRange('D:F').setNumberFormat('@');

  Logger.log('Setup selesai.');
}

// ===================== ENDPOINT =====================
function doPost(e) {
  var lock = LockService.getScriptLock();

  // 1. Antrean lock 10 detik
  if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
    return respond_(false, 'Server sibuk, silakan coba lagi.');
  }

  var undo = []; // jejak perubahan stok untuk rollback bila pencatatan log gagal

  try {
    // ---- Parse & validasi payload ----
    var p;
    try {
      p = JSON.parse(e.postData.contents);
    } catch (err) {
      return respond_(false, 'Payload bukan JSON yang valid.');
    }

    var trxId    = str_(p.transaction_id);
    var activity = ACTIVITIES[str_(p.activity).toLowerCase()];
    var barcode  = str_(p.barcode);
    var from     = str_(p.location_from);
    var to       = str_(p.location_to);
    var qty      = Number(p.qty);

    if (!trxId)    return respond_(false, 'transaction_id wajib diisi.');
    if (!activity) return respond_(false, "activity harus 'Masuk', 'Keluar', atau 'Mutasi'.");
    if (!barcode)  return respond_(false, 'barcode wajib diisi.');
    if (str_(p.qty) === '') return respond_(false, 'qty wajib diisi.');
    var minQty = (activity === 'Opname') ? 0 : 1;
    if (!isFinite(qty) || qty < minQty || Math.floor(qty) !== qty) {
      return respond_(false, activity === 'Opname'
        ? 'qty hasil hitung harus bilangan bulat >= 0.'
        : 'qty harus berupa bilangan bulat positif.');
    }
    if ((activity === 'Keluar' || activity === 'Mutasi') && !from) {
      return respond_(false, 'location_from wajib diisi untuk aktivitas ' + activity + '.');
    }
    if ((activity === 'Masuk' || activity === 'Mutasi' || activity === 'Opname') && !to) {
      return respond_(false, 'location_to wajib diisi untuk aktivitas ' + activity + '.');
    }
    if (activity === 'Mutasi' && from === to) {
      return respond_(false, 'Lokasi asal dan tujuan tidak boleh sama.');
    }

    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var shMaster = ss.getSheetByName(SHEET_MASTER);
    var shStok   = ss.getSheetByName(SHEET_STOK);
    var shLog    = ss.getSheetByName(SHEET_LOG);
    if (!shMaster || !shStok || !shLog) {
      return respond_(false, 'Sheet belum lengkap. Jalankan setup() terlebih dahulu.');
    }

    // ---- 2. Cegah double submission ----
    if (trxExists_(shLog, trxId)) {
      return respond_(false, 'transaction_id sudah pernah diproses: ' + trxId);
    }

    // ---- 3. Cari SKU berdasarkan barcode ----
    var produk = findProductByBarcode_(shMaster, barcode);
    if (!produk) {
      return respond_(false, 'Barcode tidak ditemukan di MASTER_PRODUK: ' + barcode);
    }
    if (!produk.aktif) {
      return respond_(false, 'Produk nonaktif (Status_Aktif = FALSE): ' + produk.sku);
    }
    var sku = produk.sku;

    // ---- 4. Cek stok untuk Keluar / Mutasi ----
    var stokData = readStok_(shStok);
    var stokAsal = null;

    if (activity === 'Keluar' || activity === 'Mutasi') {
      stokAsal = findStokRow_(stokData, sku, from);
      var tersedia = stokAsal ? stokAsal.qty : 0;
      if (tersedia < qty) {
        return respond_(false,
          'Stok tidak cukup di ' + from + '. Tersedia: ' + tersedia + ', diminta: ' + qty + '.');
      }
    }

    // ---- 5. Update STOK_LOKASI ----
    if (activity === 'Keluar' || activity === 'Mutasi') {
      undo.push({ type: 'set', row: stokAsal.row, old: stokAsal.qty });
      shStok.getRange(stokAsal.row, 3).setValue(stokAsal.qty - qty);
    }

    if (activity === 'Masuk' || activity === 'Mutasi') {
      var stokTujuan = findStokRow_(stokData, sku, to);
      if (stokTujuan) {
        undo.push({ type: 'set', row: stokTujuan.row, old: stokTujuan.qty });
        shStok.getRange(stokTujuan.row, 3).setValue(stokTujuan.qty + qty);
      } else {
        shStok.appendRow([sku, to, qty]);
        undo.push({ type: 'append', row: shStok.getLastRow() });
      }
    }

    // ---- 5b. Stock Opname: stok di location_to DITIMPA dengan hasil hitung fisik (qty) ----
    var logStatus = 'SUCCESS';
    var message = 'Transaksi ' + activity + ' berhasil dicatat.';
    var extra = null;

    if (activity === 'Opname') {
      var stokOp = findStokRow_(stokData, sku, to);
      var sebelum = stokOp ? stokOp.qty : 0;
      var selisih = qty - sebelum;

      if (stokOp) {
        undo.push({ type: 'set', row: stokOp.row, old: stokOp.qty });
        shStok.getRange(stokOp.row, 3).setValue(qty);
      } else if (qty > 0) {
        shStok.appendRow([sku, to, qty]);
        undo.push({ type: 'append', row: shStok.getLastRow() });
      }

      var selisihTxt = (selisih > 0 ? '+' : '') + selisih;
      logStatus = 'SUCCESS (selisih: ' + selisihTxt + ')';
      message = 'Stock Opname berhasil. Sistem: ' + sebelum + ', hitung fisik: ' + qty +
                ', selisih: ' + selisihTxt + '.';
      extra = { stok_sebelum: sebelum, stok_sesudah: qty, selisih: selisih };
    }

    // ---- 6. Catat ke OPERATION_LOG ----
    shLog.appendRow([new Date(), trxId, activity, sku, from, to, qty, logStatus]);
    SpreadsheetApp.flush();

    var result = {
      transaction_id: trxId,
      sku: sku,
      nama: produk.nama,
      activity: activity,
      qty: qty
    };
    if (extra) {
      result.stok_sebelum = extra.stok_sebelum;
      result.stok_sesudah = extra.stok_sesudah;
      result.selisih = extra.selisih;
    }
    return respond_(true, message, result);

  } catch (err) {
    // Rollback perubahan stok jika terjadi error di tengah proses
    try {
      for (var i = undo.length - 1; i >= 0; i--) {
        var u = undo[i];
        var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STOK);
        if (u.type === 'set') sh.getRange(u.row, 3).setValue(u.old);
        else if (u.type === 'append') sh.deleteRow(u.row);
      }
      SpreadsheetApp.flush();
    } catch (rollbackErr) {
      Logger.log('Rollback gagal: ' + rollbackErr);
    }
    Logger.log('doPost error: ' + err);
    return respond_(false, 'Terjadi kesalahan server: ' + err.message);

  } finally {
    lock.releaseLock();
  }
}

// Opsional: untuk tes cepat lewat browser
function doGet() {
  return respond_(true, 'Ina Basic Store Warehouse API aktif. Gunakan POST.');
}

// ===================== HELPER =====================
function str_(v) {
  return (v === null || v === undefined) ? '' : String(v).trim();
}

function respond_(success, message, data) {
  var out = { success: success, message: message };
  if (data) out.data = data;
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function trxExists_(shLog, trxId) {
  var last = shLog.getLastRow();
  if (last < 2) return false;
  var ids = shLog.getRange(2, 2, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (str_(ids[i][0]) === trxId) return true;
  }
  return false;
}

function findProductByBarcode_(shMaster, barcode) {
  var last = shMaster.getLastRow();
  if (last < 2) return null;
  var rows = shMaster.getRange(2, 1, last - 1, 4).getValues(); // SKU, Barcode, Nama, Status_Aktif
  for (var i = 0; i < rows.length; i++) {
    if (str_(rows[i][1]) === barcode) {
      var status = rows[i][3];
      var aktif = (status === true) || (str_(status).toUpperCase() === 'TRUE');
      return { sku: str_(rows[i][0]), nama: str_(rows[i][2]), aktif: aktif };
    }
  }
  return null;
}

function readStok_(shStok) {
  var last = shStok.getLastRow();
  if (last < 2) return [];
  var rows = shStok.getRange(2, 1, last - 1, 3).getValues();
  return rows.map(function (r, i) {
    return { row: i + 2, sku: str_(r[0]), lokasi: str_(r[1]), qty: Number(r[2]) || 0 };
  });
}

function findStokRow_(stokData, sku, lokasi) {
  for (var i = 0; i < stokData.length; i++) {
    if (stokData[i].sku === sku && stokData[i].lokasi === lokasi) return stokData[i];
  }
  return null;
}
