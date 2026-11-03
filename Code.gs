/**
 * INA BASIC STORE WAREHOUSE - Backend API Final
 */

var SHEET_MASTER = 'MASTER_PRODUK';
var SHEET_STOK   = 'STOK_LOKASI';
var SHEET_LOG    = 'OPERATION_LOG';

var HEADERS = {
  MASTER_PRODUK: ['SKU', 'Barcode', 'Nama', 'Status_Aktif'],
  STOK_LOKASI:   ['SKU', 'Lokasi_Rak', 'Kuantitas'],
  OPERATION_LOG: ['Timestamp', 'ID_Transaksi', 'Aktivitas', 'SKU', 'Lokasi_Asal', 'Lokasi_Tujuan', 'Qty', 'Status']
};

var LOCK_TIMEOUT_MS = 10000;
var ACTIVITIES = { 'masuk': 'Masuk', 'keluar': 'Keluar', 'mutasi': 'Mutasi', 'opname': 'Opname', 'produkbaru': 'ProdukBaru' };

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]).setFontWeight('bold').setBackground('#f3f3f3');
      sheet.setFrozenRows(1);
    }
  });
  ss.getSheetByName(SHEET_MASTER).getRange('A:B').setNumberFormat('@');
  ss.getSheetByName(SHEET_STOK).getRange('A:B').setNumberFormat('@');
  ss.getSheetByName(SHEET_LOG).getRange('B:B').setNumberFormat('@');
  ss.getSheetByName(SHEET_LOG).getRange('D:F').setNumberFormat('@');
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_TIMEOUT_MS)) return respond_(false, 'Server sibuk, coba lagi.');

  var undo = []; 
  try {
    var p;
    try { p = JSON.parse(e.postData.contents); } catch (err) { return respond_(false, 'Format JSON salah.'); }

    var trxId    = str_(p.transaction_id);
    var activity = ACTIVITIES[str_(p.activity).toLowerCase()];
    var barcode  = str_(p.barcode);
    var from     = str_(p.location_from);
    var to       = str_(p.location_to);
    var qty      = Number(p.qty);

    if (!trxId) return respond_(false, 'transaction_id wajib diisi.');
    if (!activity) return respond_(false, 'activity tidak valid.');
    if (!barcode) return respond_(false, 'barcode wajib diisi.');

    if (activity === 'ProdukBaru') return registerProduct_(p, trxId, barcode, to, qty);

    if (str_(p.qty) === '') return respond_(false, 'Kuantitas wajib diisi.');
    var minQty = (activity === 'Opname') ? 0 : 1;
    if (!isFinite(qty) || qty < minQty || Math.floor(qty) !== qty) return respond_(false, 'Kuantitas harus bilangan bulat valid.');
    
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
    if (!produk.aktif) return respond_(false, 'Produk nonaktif.');
    var sku = produk.sku;

    var stokData = readStok_(shStok);
    var stokAsal = null;

    if (activity === 'Keluar' || activity === 'Mutasi') {
      stokAsal = findStokRow_(stokData, sku, from);
      var tersedia = stokAsal ? stokAsal.qty : 0;
      if (tersedia < qty) return respond_(false, 'Stok di ' + from + ' tidak cukup. Tersedia: ' + tersedia);
    }

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

    var logStatus = 'SUCCESS';
    var message = 'Transaksi berhasil disimpan.';

    if (activity === 'Opname') {
      var stokOp = findStokRow_(stokData, sku, to);
      var sebelum = stokOp ? stokOp.qty : 0;
      
      if (stokOp) {
        undo.push({ type: 'set', row: stokOp.row, old: stokOp.qty });
        shStok.getRange(stokOp.row, 3).setValue(qty);
      } else if (qty > 0) {
        shStok.appendRow([sku, to, qty]);
        undo.push({ type: 'append', row: shStok.getLastRow() });
      }
      logStatus = 'SUCCESS (Hitung fisik: ' + qty + ')';
    }

    shLog.appendRow([new Date(), trxId, activity, sku, from, to, qty, logStatus]);
    SpreadsheetApp.flush();

    return respond_(true, message, { sku: sku, nama: produk.nama });

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
  var sku = str_(p.sku);
  var nama = str_(p.nama);

  if (!sku || !nama || !locTo) return respond_(false, 'SKU, Nama, dan Lokasi wajib diisi.');
  if (typeof qty !== 'number' || qty < 0) return respond_(false, 'Stok awal harus bilangan positif.');
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shMaster = ss.getSheetByName(SHEET_MASTER);
  var shStok = ss.getSheetByName(SHEET_STOK);
  var shLog = ss.getSheetByName(SHEET_LOG);

  if (trxExists_(shLog, trxId)) return respond_(false, 'Transaksi terdeteksi duplikat.');

  var last = shMaster.getLastRow();
  if (last >= 2) {
    var rows = shMaster.getRange(2, 1, last - 1, 2).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (str_(rows[i][1]) === barcode) return respond_(false, 'Barcode sudah dipakai oleh SKU: ' + str_(rows[i][0]));
      if (str_(rows[i][0]).toLowerCase() === sku.toLowerCase()) return respond_(false, 'SKU sudah terdaftar.');
    }
  }

  var rowIdx = last + 1;
  try {
    shMaster.getRange(rowIdx, 1, 1, 4).setNumberFormats([['@', '@', '@', 'General']]).setValues([[sku, barcode, nama, true]]);
    if(qty >= 0) shStok.appendRow([sku, locTo, qty]);
    shLog.appendRow([new Date(), trxId, 'ProdukBaru', sku, '', locTo, qty, 'SUCCESS']);
    SpreadsheetApp.flush();
  } catch (err) {
    try { if (shMaster.getLastRow() >= rowIdx) shMaster.deleteRow(rowIdx); } catch (e) {}
    throw err;
  }
  return respond_(true, 'Produk berhasil didaftarkan.', { sku: sku, nama: nama });
}

function doGet() { return respond_(true, 'API Aktif.'); }
function str_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
function respond_(success, message, data) {
  var out = { success: success, message: message };
  if (data) out.data = data;
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function trxExists_(shLog, trxId) {
  var last = shLog.getLastRow();
  if (last < 2) return false;
  var ids = shLog.getRange(2, 2, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) { if (str_(ids[i][0]) === trxId) return true; }
  return false;
}
function findProductByBarcode_(shMaster, barcode) {
  var last = shMaster.getLastRow();
  if (last < 2) return null;
  var rows = shMaster.getRange(2, 1, last - 1, 4).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (str_(rows[i][1]) === barcode) {
      return { sku: str_(rows[i][0]), nama: str_(rows[i][2]), aktif: (str_(rows[i][3]).toUpperCase() === 'TRUE' || rows[i][3] === true) };
    }
  }
  return null;
}
function readStok_(shStok) {
  var last = shStok.getLastRow();
  if (last < 2) return [];
  return shStok.getRange(2, 1, last - 1, 3).getValues().map(function (r, i) {
    return { row: i + 2, sku: str_(r[0]), lokasi: str_(r[1]), qty: Number(r[2]) || 0 };
  });
}
function findStokRow_(stokData, sku, lokasi) {
  for (var i = 0; i < stokData.length; i++) { if (stokData[i].sku === sku && stokData[i].lokasi === lokasi) return stokData[i]; }
  return null;
}