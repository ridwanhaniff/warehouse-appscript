/**
 * INA BASIC STORE WAREHOUSE - Backend (Dengan Kelola Produk)
 */

var SHEET_MASTER = 'MASTER_PRODUK';
var SHEET_STOK   = 'STOK_LOKASI';
var SHEET_LOG    = 'OPERATION_LOG';

var HEADERS = {
  MASTER_PRODUK: ['Barcode', 'Nama', 'Status_Aktif'],
  STOK_LOKASI:   ['Barcode', 'Lokasi_Rak', 'Kuantitas'],
  OPERATION_LOG: ['Timestamp', 'ID_Transaksi', 'Aktivitas', 'Barcode', 'Lokasi_Asal', 'Lokasi_Tujuan', 'Qty', 'Status']
};

var LOCK_TIMEOUT_MS = 10000;
// DITAMBAHKAN: Aktivitas Kelola (Cek, Edit, Hapus)
var ACTIVITIES = { 
  'masuk': 'Masuk', 'keluar': 'Keluar', 'mutasi': 'Mutasi', 
  'opname': 'Opname', 'produkbaru': 'ProdukBaru',
  'cekproduk': 'CekProduk', 'editproduk': 'EditProduk', 'hapusproduk': 'HapusProduk' 
};

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(HEADERS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold').setBackground('#d9eaf7');
      sh.setFrozenRows(1);
    }
  });

  var shMaster = ss.getSheetByName(SHEET_MASTER);
  var shStok = ss.getSheetByName(SHEET_STOK);
  var shLog = ss.getSheetByName(SHEET_LOG);
  
  shMaster.getRange('A:A').setNumberFormat('@');
  shStok.getRange('A:B').setNumberFormat('@');
  shLog.getRange('B:B').setNumberFormat('@');
  shLog.getRange('D:F').setNumberFormat('@');

  var masterLast = shMaster.getLastRow();
  if (shMaster.getMaxRows() < 2) shMaster.insertRowsAfter(1, 1);
  var statusColumnRange = shMaster.getRange(2, 3, shMaster.getMaxRows() - 1, 1);
  statusColumnRange.clearDataValidations();
  statusColumnRange.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build());

  [shMaster, shStok, shLog].forEach(function (sh) {
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#d9eaf7');
    sh.setFrozenRows(1);
  });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_TIMEOUT_MS)) return respond_(false, 'Server sibuk, coba lagi.');

  var undo = [];
  try {
    var p;
    try { p = JSON.parse(e.postData.contents); } catch (err) { return respond_(false, 'Format JSON salah.'); }

    var activityRaw = str_(p.activity).toLowerCase();
    var trxId = str_(p.transaction_id);
    var activity = ACTIVITIES[activityRaw];
    var barcode = str_(p.barcode);
    var from = str_(p.location_from);
    var to = str_(p.location_to);
    var qty = Number(p.qty);

    if (!activity || !barcode) return respond_(false, 'Data transaksi tidak lengkap.');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var shMaster = ss.getSheetByName(SHEET_MASTER);
    var shStok = ss.getSheetByName(SHEET_STOK);
    var shLog = ss.getSheetByName(SHEET_LOG);

    // ==========================================
    // FITUR KELOLA: CEK, EDIT, HAPUS PRODUK
    // ==========================================
    if (activity === 'CekProduk') {
      var produkCek = findProductByBarcode_(shMaster, barcode);
      if (!produkCek) return respond_(false, 'Produk tidak ditemukan di database.');
      return respond_(true, 'Data ditemukan', produkCek);
    }

    if (activity === 'EditProduk') {
      var produkEdit = findProductByBarcode_(shMaster, barcode);
      if (!produkEdit) return respond_(false, 'Produk tidak ditemukan.');
      var newNama = str_(p.nama);
      var newStatus = toBoolean_(p.status_aktif, true);
      
      shMaster.getRange(produkEdit.row, 2).setValue(newNama);
      shMaster.getRange(produkEdit.row, 3).setValue(newStatus);
      shLog.appendRow([new Date(), trxId, 'EditProduk', "'" + barcode, '', '', 0, 'SUCCESS']);
      SpreadsheetApp.flush();
      return respond_(true, 'Perubahan produk berhasil disimpan.');
    }

    if (activity === 'HapusProduk') {
      var produkHapus = findProductByBarcode_(shMaster, barcode);
      if (!produkHapus) return respond_(false, 'Produk tidak ditemukan.');
      
      // Hapus dari Master
      shMaster.deleteRow(produkHapus.row);
      
      // Hapus semua riwayat stok lokasi terkait produk ini
      var stokDataRaw = readStok_(shStok);
      for (var i = stokDataRaw.length - 1; i >= 0; i--) {
        if (stokDataRaw[i].barcode === barcode) {
          shStok.deleteRow(stokDataRaw[i].row);
        }
      }
      shLog.appendRow([new Date(), trxId, 'HapusProduk', "'" + barcode, '', '', 0, 'SUCCESS']);
      SpreadsheetApp.flush();
      return respond_(true, 'Produk dan stok terkait berhasil dihapus permanen.');
    }
    // ==========================================

    if (!trxId) return respond_(false, 'ID Transaksi wajib diisi.');
    if (activity === 'ProdukBaru') return registerProduct_(p, trxId, barcode, to, qty);

    if (str_(p.qty) === '') return respond_(false, 'Kuantitas wajib diisi.');
    var minQty = activity === 'Opname' ? 0 : 1;
    if (!isFinite(qty) || qty < minQty || Math.floor(qty) !== qty) return respond_(false, 'Kuantitas tidak valid.');
    
    if ((activity === 'Keluar' || activity === 'Mutasi') && !from) return respond_(false, 'Lokasi asal wajib diisi.');
    if ((activity === 'Masuk' || activity === 'Mutasi' || activity === 'Opname') && !to) return respond_(false, 'Lokasi tujuan wajib diisi.');
    if (activity === 'Mutasi' && from === to) return respond_(false, 'Lokasi asal dan tujuan tidak boleh sama.');

    if (trxExists_(shLog, trxId)) return respond_(false, 'Transaksi sudah pernah diproses.');

    var produk = findProductByBarcode_(shMaster, barcode);
    if (!produk) return respond_(false, 'Barcode belum terdaftar di Master Produk.');
    if (!produk.aktif) return respond_(false, 'Produk sedang nonaktif.');

    var stokData = readStok_(shStok);
    var stokAsal = null;

    if (activity === 'Keluar' || activity === 'Mutasi') {
      stokAsal = findStokRow_(stokData, barcode, from);
      var tersedia = stokAsal ? stokAsal.qty : 0;
      if (tersedia < qty) return respond_(false, 'Stok di ' + from + ' tidak cukup. Tersedia: ' + tersedia);
    }

    if (activity === 'Keluar' || activity === 'Mutasi') {
      undo.push({ type: 'set', row: stokAsal.row, old: stokAsal.qty });
      shStok.getRange(stokAsal.row, 3).setValue(stokAsal.qty - qty);
    }

    if (activity === 'Masuk' || activity === 'Mutasi') {
      var stokTujuan = findStokRow_(stokData, barcode, to);
      if (stokTujuan) {
        undo.push({ type: 'set', row: stokTujuan.row, old: stokTujuan.qty });
        shStok.getRange(stokTujuan.row, 3).setValue(stokTujuan.qty + qty);
      } else {
        shStok.appendRow(["'" + barcode, to, qty]); 
        undo.push({ type: 'append', row: shStok.getLastRow() });
      }
    }

    var logStatus = 'SUCCESS';
    if (activity === 'Opname') {
      var stokOp = findStokRow_(stokData, barcode, to);
      if (stokOp) {
        undo.push({ type: 'set', row: stokOp.row, old: stokOp.qty });
        shStok.getRange(stokOp.row, 3).setValue(qty);
      } else if (qty > 0) {
        shStok.appendRow(["'" + barcode, to, qty]); 
        undo.push({ type: 'append', row: shStok.getLastRow() });
      }
      logStatus = 'SUCCESS (Hitung fisik: ' + qty + ')';
    }

    shLog.appendRow([new Date(), trxId, activity, "'" + barcode, from, to, qty, logStatus]);
    SpreadsheetApp.flush();
    return respond_(true, 'Transaksi berhasil disimpan.', { barcode: barcode, nama: produk.nama });

  } catch (err) {
    for (var i = undo.length - 1; i >= 0; i--) {
      var u = undo[i];
      var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STOK);
      if (u.type === 'set') sh.getRange(u.row, 3).setValue(u.old);
      else if (u.type === 'append') sh.deleteRow(u.row);
    }
    return respond_(false, 'Error Server: ' + err.message);
  } finally {
    lock.releaseLock();
  }
}

function registerProduct_(p, trxId, barcode, locTo, qty) {
  var nama = str_(p.nama);
  var statusAktif = toBoolean_(p.status_aktif, true);

  if (!barcode) return respond_(false, 'Barcode wajib diisi.');
  if (!nama || nama.length < 2) return respond_(false, 'Nama produk wajib diisi dengan benar.');
  if (!locTo) return respond_(false, 'Lokasi awal wajib dipilih.');
  if (!isFinite(qty) || qty < 0 || Math.floor(qty) !== qty) return respond_(false, 'Stok awal harus bilangan bulat minimal 0.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shMaster = ss.getSheetByName(SHEET_MASTER);
  var shStok = ss.getSheetByName(SHEET_STOK);
  var shLog = ss.getSheetByName(SHEET_LOG);

  if (trxExists_(shLog, trxId)) return respond_(false, 'Transaksi terdeteksi duplikat.');

  var produk = findProductByBarcode_(shMaster, barcode);
  if (produk) return respond_(false, 'Barcode sudah terdaftar sebelumnya.');

  var last = shMaster.getLastRow();
  var masterRow = last + 1;
  if (masterRow > shMaster.getMaxRows()) shMaster.insertRowsAfter(shMaster.getMaxRows(), masterRow - shMaster.getMaxRows());
  
  var stockRow = null, logRow = null;
  try {
    var statusCell = shMaster.getRange(masterRow, 3);
    statusCell.clearDataValidations();
    statusCell.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build());
    
    shMaster.getRange(masterRow, 1, 1, 2).setNumberFormat('@').setValues([["'" + barcode, nama]]);
    statusCell.setValue(statusAktif);

    shStok.appendRow(["'" + barcode, locTo, qty]);
    stockRow = shStok.getLastRow();
    
    shLog.appendRow([new Date(), trxId, 'ProdukBaru', "'" + barcode, '', locTo, qty, 'SUCCESS']);
    logRow = shLog.getLastRow();
    
    SpreadsheetApp.flush();
  } catch (err) {
    try { if (logRow && shLog.getLastRow() >= logRow) shLog.deleteRow(logRow); } catch (e) {}
    try { if (stockRow && shStok.getLastRow() >= stockRow) shStok.deleteRow(stockRow); } catch (e) {}
    try { if (shMaster.getLastRow() >= masterRow) shMaster.deleteRow(masterRow); } catch (e) {}
    throw err;
  }

  return respond_(true, 'Produk berhasil didaftarkan.', { barcode: barcode, nama: nama });
}

function doGet() { return respond_(true, 'API Aktif.'); }
function str_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
function toBoolean_(value, defaultValue) {
  if (value === true || value === false) return value;
  var normalized = str_(value).toUpperCase();
  if (normalized === 'TRUE' || normalized === '1' || normalized === 'YA' || normalized === 'AKTIF') return true;
  if (normalized === 'FALSE' || normalized === '0' || normalized === 'TIDAK' || normalized === 'NONAKTIF') return false;
  return defaultValue;
}
function respond_(success, message, data) {
  var out = { success: success, message: message };
  if (data) out.data = data;
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function trxExists_(shLog, trxId) {
  var last = shLog.getLastRow();
  if (last < 2) return false;
  var ids = shLog.getRange(2, 2, last - 1, 1).getDisplayValues();
  for (var i = 0; i < ids.length; i++) { if (str_(ids[i][0]) === trxId) return true; }
  return false;
}
function findProductByBarcode_(shMaster, barcode) {
  var last = shMaster.getLastRow();
  if (last < 2) return null;
  var rows = shMaster.getRange(2, 1, last - 1, 3).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (str_(rows[i][0]).toUpperCase() === barcode.toUpperCase()) {
      // DITAMBAHKAN: Return 'row' agar data dapat dihapus/diedit pada index yang tepat
      return { row: i + 2, barcode: str_(rows[i][0]), nama: str_(rows[i][1]), aktif: toBoolean_(rows[i][2], false) };
    }
  }
  return null;
}
function readStok_(shStok) {
  var last = shStok.getLastRow();
  if (last < 2) return [];
  return shStok.getRange(2, 1, last - 1, 3).getValues().map(function (r, i) { return { row: i + 2, barcode: str_(r[0]), lokasi: str_(r[1]), qty: Number(r[2]) || 0 }; });
}
function findStokRow_(stokData, barcode, lokasi) {
  for (var i = 0; i < stokData.length; i++) { if (stokData[i].barcode === barcode && stokData[i].lokasi === lokasi) return stokData[i]; }
  return null;
}