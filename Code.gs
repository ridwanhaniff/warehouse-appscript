/**
 * INA BASIC STORE WAREHOUSE - Backend (Dengan Mitigasi Ketat)
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
var ACTIVITIES = { 'masuk': 'Masuk', 'keluar': 'Keluar', 'mutasi': 'Mutasi', 'opname': 'Opname', 'produkbaru': 'ProdukBaru' };

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
  var codeMap = {};

  if (shMaster.getLastRow() >= 1) {
    var oldHeader = shMaster.getRange(1, 1, 1, Math.max(3, shMaster.getLastColumn())).getDisplayValues()[0];
    var isLegacyMaster = str_(oldHeader[0]).toUpperCase() === 'SKU' && str_(oldHeader[1]).toUpperCase() === 'BARCODE';

    if (isLegacyMaster) {
      shMaster.getRange(1, 3, shMaster.getMaxRows(), 1).clearDataValidations();
      var oldLast = shMaster.getLastRow();
      var oldRows = oldLast >= 2 ? shMaster.getRange(2, 1, oldLast - 1, 4).getValues() : [];
      var newRows = [];

      oldRows.forEach(function (row) {
        var oldSku = str_(row[0]), oldBarcode = str_(row[1]);
        var barcode = oldBarcode || oldSku;
        if (!barcode) return;
        var nama = str_(row[2]), aktif = toBoolean_(row[3], false);
        if (oldSku) codeMap[oldSku] = barcode;
        if (oldBarcode) codeMap[oldBarcode] = barcode;
        newRows.push([barcode, nama, aktif]);
      });

      shMaster.clearContents();
      shMaster.getRange(1, 1, 1, 3).setValues([HEADERS.MASTER_PRODUK]);
      if (newRows.length) shMaster.getRange(2, 1, newRows.length, 3).setValues(newRows);
    }
  }

  if (Object.keys(codeMap).length) {
    remapIdentifierColumn_(shStok, 1, codeMap);
    remapIdentifierColumn_(shLog, 4, codeMap);
  }

  if (shStok.getLastRow() >= 1 && str_(shStok.getRange(1, 1).getDisplayValue()).toUpperCase() === 'SKU') shStok.getRange(1, 1).setValue('Barcode');
  if (shLog.getLastRow() >= 1 && str_(shLog.getRange(1, 4).getDisplayValue()).toUpperCase() === 'SKU') shLog.getRange(1, 4).setValue('Barcode');

  shMaster.getRange('A:A').setNumberFormat('@');
  shStok.getRange('A:B').setNumberFormat('@');
  shLog.getRange('B:B').setNumberFormat('@');
  shLog.getRange('D:F').setNumberFormat('@');

  var masterLast = shMaster.getLastRow();
  if (shMaster.getMaxRows() < 2) shMaster.insertRowsAfter(1, 1);
  var statusColumnRange = shMaster.getRange(2, 3, shMaster.getMaxRows() - 1, 1);
  statusColumnRange.clearDataValidations();

  if (masterLast >= 2) {
    var statusDataRange = shMaster.getRange(2, 3, masterLast - 1, 1);
    statusDataRange.setValues(statusDataRange.getValues().map(function (r) { return [toBoolean_(r[0], false)]; }));
  }
  statusColumnRange.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build());

  [shMaster, shStok, shLog].forEach(function (sh) {
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#d9eaf7');
    sh.setFrozenRows(1);
  });
}

function remapIdentifierColumn_(sheet, col, map) {
  var last = sheet.getLastRow();
  if (last < 2) return;
  var range = sheet.getRange(2, col, last - 1, 1);
  var values = range.getValues();
  var changed = false;
  values = values.map(function (row) {
    var oldCode = str_(row[0]);
    if (oldCode && map[oldCode]) { changed = true; return [map[oldCode]]; }
    return row;
  });
  if (changed) range.setValues(values);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_TIMEOUT_MS)) return respond_(false, 'Server sibuk, coba lagi.');

  var undo = [];
  try {
    var p;
    try { p = JSON.parse(e.postData.contents); } catch (err) { return respond_(false, 'Format JSON salah.'); }

    var trxId = str_(p.transaction_id);
    var activity = ACTIVITIES[str_(p.activity).toLowerCase()];
    var barcode = str_(p.barcode);
    var from = str_(p.location_from);
    var to = str_(p.location_to);
    var qty = Number(p.qty);

    if (!trxId || !activity || !barcode) return respond_(false, 'Data transaksi tidak lengkap.');

    if (activity === 'ProdukBaru') return registerProduct_(p, trxId, barcode, to, qty);

    if (str_(p.qty) === '') return respond_(false, 'Kuantitas wajib diisi.');
    var minQty = activity === 'Opname' ? 0 : 1;
    if (!isFinite(qty) || qty < minQty || Math.floor(qty) !== qty) return respond_(false, 'Kuantitas tidak valid.');
    
    if ((activity === 'Keluar' || activity === 'Mutasi') && !from) return respond_(false, 'Lokasi asal wajib diisi.');
    if ((activity === 'Masuk' || activity === 'Mutasi' || activity === 'Opname') && !to) return respond_(false, 'Lokasi tujuan wajib diisi.');
    if (activity === 'Mutasi' && from === to) return respond_(false, 'Lokasi asal dan tujuan tidak boleh sama.');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var shMaster = ss.getSheetByName(SHEET_MASTER);
    var shStok = ss.getSheetByName(SHEET_STOK);
    var shLog = ss.getSheetByName(SHEET_LOG);

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
        shStok.appendRow(["'" + barcode, to, qty]); // MITIGASI: Paksa format teks dengan kutip tunggal
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
        shStok.appendRow(["'" + barcode, to, qty]); // MITIGASI
        undo.push({ type: 'append', row: shStok.getLastRow() });
      }
      logStatus = 'SUCCESS (Hitung fisik: ' + qty + ')';
    }

    // MITIGASI: Paksa Barcode Teks Absolut saat mencatat log
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

  // MITIGASI 1: Cegah Bypass input kosong (Backend Level)
  if (!barcode) return respond_(false, 'Barcode wajib diisi.');
  if (!nama || nama.length < 2) return respond_(false, 'Nama produk wajib diisi dengan benar.');
  if (!locTo) return respond_(false, 'Lokasi awal wajib dipilih.');
  if (!isFinite(qty) || qty < 0 || Math.floor(qty) !== qty) return respond_(false, 'Stok awal harus bilangan bulat minimal 0.');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shMaster = ss.getSheetByName(SHEET_MASTER);
  var shStok = ss.getSheetByName(SHEET_STOK);
  var shLog = ss.getSheetByName(SHEET_LOG);

  // MITIGASI 2: Cek Struktur Kolom Master Produk
  var mHeader = shMaster.getRange(1, 1, 1, 3).getDisplayValues()[0];
  if (str_(mHeader[0]).toUpperCase() !== 'BARCODE' || str_(mHeader[1]).toUpperCase() !== 'NAMA') {
    return respond_(false, 'Sistem dicegah! Struktur kolom Master Produk tidak valid. Buka Google Sheets Anda dan klik menu "Ina Warehouse > Bangun / Perbarui Dashboard" terlebih dahulu.');
  }

  if (trxExists_(shLog, trxId)) return respond_(false, 'Transaksi terdeteksi duplikat.');

  var last = shMaster.getLastRow();
  if (last >= 2) {
    var existing = shMaster.getRange(2, 1, last - 1, 1).getDisplayValues();
    for (var i = 0; i < existing.length; i++) {
      if (str_(existing[i][0]).toUpperCase() === barcode.toUpperCase()) {
        return respond_(false, 'Barcode sudah terdaftar sebelumnya.');
      }
    }
  }

  var masterRow = last + 1;
  if (masterRow > shMaster.getMaxRows()) shMaster.insertRowsAfter(shMaster.getMaxRows(), masterRow - shMaster.getMaxRows());
  
  var stockRow = null, logRow = null;
  try {
    var statusCell = shMaster.getRange(masterRow, 3);
    statusCell.clearDataValidations();
    statusCell.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build());
    
    // MITIGASI 3: Force Format
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
      return { barcode: str_(rows[i][0]), nama: str_(rows[i][1]), aktif: toBoolean_(rows[i][2], false) };
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