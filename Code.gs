/**
 * INA BASIC STORE WAREHOUSE - Backend
 * Identifier produk tunggal: Barcode (dipakai juga sebagai kunci stok/transaksi).
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
var ACTIVITIES = {
  'masuk': 'Masuk',
  'keluar': 'Keluar',
  'mutasi': 'Mutasi',
  'opname': 'Opname',
  'produkbaru': 'ProdukBaru'
};

/**
 * Jalankan sekali setelah mengganti kode.
 * Fungsi ini juga memigrasikan master lama [SKU, Barcode, Nama, Status_Aktif]
 * menjadi [Barcode, Nama, Status_Aktif], serta menyelaraskan kunci stok/log.
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(HEADERS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEADERS[name].length)
        .setValues([HEADERS[name]])
        .setFontWeight('bold')
        .setBackground('#d9eaf7');
      sh.setFrozenRows(1);
    }
  });

  var shMaster = ss.getSheetByName(SHEET_MASTER);
  var shStok = ss.getSheetByName(SHEET_STOK);
  var shLog = ss.getSheetByName(SHEET_LOG);
  var codeMap = {};

  // Migrasi master lama: SKU dan Barcode digabung menjadi satu kolom Barcode.
  if (shMaster.getLastRow() >= 1) {
    var oldHeader = shMaster.getRange(1, 1, 1, Math.max(3, shMaster.getLastColumn())).getDisplayValues()[0];
    var isLegacyMaster = str_(oldHeader[0]).toUpperCase() === 'SKU' &&
                         str_(oldHeader[1]).toUpperCase() === 'BARCODE';

    if (isLegacyMaster) {
      var oldLast = shMaster.getLastRow();
      var oldRows = oldLast >= 2 ? shMaster.getRange(2, 1, oldLast - 1, 4).getValues() : [];
      var newRows = [];

      oldRows.forEach(function (row) {
        var oldSku = str_(row[0]);
        var oldBarcode = str_(row[1]);
        var barcode = oldBarcode || oldSku;
        if (!barcode) return;
        var nama = str_(row[2]);
        var aktif = toBoolean_(row[3], false);
        if (oldSku) codeMap[oldSku] = barcode;
        if (oldBarcode) codeMap[oldBarcode] = barcode;
        newRows.push([barcode, nama, aktif]);
      });

      shMaster.clearContents();
      shMaster.getRange(1, 1, 1, 3).setValues([HEADERS.MASTER_PRODUK]);
      if (newRows.length) shMaster.getRange(2, 1, newRows.length, 3).setValues(newRows);
    } else {
      // Pastikan nama kolom baru konsisten jika sheet sudah memakai skema 3 kolom.
      var current = shMaster.getRange(1, 1, 1, 3).getDisplayValues()[0];
      if (str_(current[0]).toUpperCase() === 'BARCODE' && str_(current[1]).toUpperCase() === 'NAMA') {
        shMaster.getRange(1, 1, 1, 3).setValues([HEADERS.MASTER_PRODUK]);
      }
    }
  }

  // Ubah nilai kunci lama di STOK_LOKASI dan OPERATION_LOG jika SKU berbeda dari Barcode.
  if (Object.keys(codeMap).length) {
    remapIdentifierColumn_(shStok, 1, codeMap);
    remapIdentifierColumn_(shLog, 4, codeMap);
  }

  // Selaraskan header kunci lama menjadi Barcode.
  if (shStok.getLastRow() >= 1 && str_(shStok.getRange(1, 1).getDisplayValue()).toUpperCase() === 'SKU') {
    shStok.getRange(1, 1).setValue('Barcode');
  }
  if (shLog.getLastRow() >= 1 && str_(shLog.getRange(1, 4).getDisplayValue()).toUpperCase() === 'SKU') {
    shLog.getRange(1, 4).setValue('Barcode');
  }

  // Format identitas sebagai teks agar kode seperti A-R01-T01 tidak berubah.
  shMaster.getRange('A:A').setNumberFormat('@');
  shStok.getRange('A:B').setNumberFormat('@');
  shLog.getRange('B:B').setNumberFormat('@');
  shLog.getRange('D:F').setNumberFormat('@');

  // Status aktif memakai checkbox dengan nilai boolean TRUE/FALSE.
  var masterLast = shMaster.getLastRow();
  if (masterLast >= 2) {
    var statusRange = shMaster.getRange(2, 3, masterLast - 1, 1);
    var statusValues = statusRange.getValues().map(function (r) {
      return [toBoolean_(r[0], false)];
    });
    statusRange.setValues(statusValues);
  }
  var validationRows = Math.max(200, Math.min(2000, shMaster.getMaxRows() - 1));
  if (masterLast > validationRows + 1) validationRows = masterLast - 1;
  shMaster.getRange(2, 3, validationRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(false).build()
  );

  // Perapian header dan pembekuan baris pertama.
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
    if (oldCode && map[oldCode]) {
      changed = true;
      return [map[oldCode]];
    }
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
    try {
      if (!e || !e.postData || !e.postData.contents) throw new Error('Payload kosong');
      p = JSON.parse(e.postData.contents);
    } catch (err) {
      return respond_(false, 'Format JSON salah atau payload kosong.');
    }

    var trxId = str_(p.transaction_id);
    var activity = ACTIVITIES[str_(p.activity).toLowerCase()];
    var barcode = str_(p.barcode);
    var from = str_(p.location_from);
    var to = str_(p.location_to);
    var qty = Number(p.qty);

    if (!trxId) return respond_(false, 'transaction_id wajib diisi.');
    if (!activity) return respond_(false, 'activity tidak valid.');
    if (!barcode) return respond_(false, 'barcode wajib diisi.');

    if (activity === 'ProdukBaru') return registerProduct_(p, trxId, barcode, to, qty);

    if (str_(p.qty) === '') return respond_(false, 'Kuantitas wajib diisi.');
    var minQty = activity === 'Opname' ? 0 : 1;
    if (!isFinite(qty) || qty < minQty || Math.floor(qty) !== qty) {
      return respond_(false, 'Kuantitas harus bilangan bulat valid.');
    }

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
        shStok.appendRow([barcode, to, qty]);
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
        shStok.appendRow([barcode, to, qty]);
        undo.push({ type: 'append', row: shStok.getLastRow() });
      }
      logStatus = 'SUCCESS (Hitung fisik: ' + qty + ')';
    }

    shLog.appendRow([new Date(), trxId, activity, barcode, from, to, qty, logStatus]);
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

/** Pendaftaran produk: Barcode menjadi satu-satunya identitas produk. */
function registerProduct_(p, trxId, barcode, locTo, qty) {
  var nama = str_(p.nama);
  var statusAktif = toBoolean_(p.status_aktif, true);

  if (!barcode || !nama || !locTo) return respond_(false, 'Barcode, Nama, dan Lokasi wajib diisi.');
  if (!isFinite(qty) || qty < 0 || Math.floor(qty) !== qty) {
    return respond_(false, 'Stok awal harus bilangan bulat minimal 0.');
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shMaster = ss.getSheetByName(SHEET_MASTER);
  var shStok = ss.getSheetByName(SHEET_STOK);
  var shLog = ss.getSheetByName(SHEET_LOG);

  if (trxExists_(shLog, trxId)) return respond_(false, 'Transaksi terdeteksi duplikat.');

  var last = shMaster.getLastRow();
  if (last >= 2) {
    var existing = shMaster.getRange(2, 1, last - 1, 1).getDisplayValues();
    for (var i = 0; i < existing.length; i++) {
      if (str_(existing[i][0]).toUpperCase() === barcode.toUpperCase()) {
        return respond_(false, 'Barcode sudah terdaftar. Gunakan aktivitas transaksi untuk mengubah stok produk ini.');
      }
    }
  }

  var masterRow = last + 1;
  var stockRow = null;
  var logRow = null;
  try {
    shMaster.getRange(masterRow, 1, 1, 3).setNumberFormat('@').setValues([[barcode, nama, statusAktif]]);
    shStok.appendRow([barcode, locTo, qty]);
    stockRow = shStok.getLastRow();
    shLog.appendRow([new Date(), trxId, 'ProdukBaru', barcode, '', locTo, qty, 'SUCCESS']);
    logRow = shLog.getLastRow();
    SpreadsheetApp.flush();
  } catch (err) {
    try { if (logRow && shLog.getLastRow() >= logRow) shLog.deleteRow(logRow); } catch (ignoreLog) {}
    try { if (stockRow && shStok.getLastRow() >= stockRow) shStok.deleteRow(stockRow); } catch (ignoreStock) {}
    try { if (shMaster.getLastRow() >= masterRow) shMaster.deleteRow(masterRow); } catch (ignoreMaster) {}
    throw err;
  }

  return respond_(true, 'Produk berhasil didaftarkan.', { barcode: barcode, nama: nama });
}

function doGet() {
  return respond_(true, 'API Aktif. Skema produk menggunakan Barcode sebagai identitas tunggal.');
}

function str_(v) {
  return (v === null || v === undefined) ? '' : String(v).trim();
}

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
  for (var i = 0; i < ids.length; i++) {
    if (str_(ids[i][0]) === trxId) return true;
  }
  return false;
}

function findProductByBarcode_(shMaster, barcode) {
  var last = shMaster.getLastRow();
  if (last < 2) return null;
  var rows = shMaster.getRange(2, 1, last - 1, 3).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (str_(rows[i][0]).toUpperCase() === barcode.toUpperCase()) {
      return {
        barcode: str_(rows[i][0]),
        nama: str_(rows[i][1]),
        aktif: toBoolean_(rows[i][2], false)
      };
    }
  }
  return null;
}

function readStok_(shStok) {
  var last = shStok.getLastRow();
  if (last < 2) return [];
  return shStok.getRange(2, 1, last - 1, 3).getValues().map(function (r, i) {
    return { row: i + 2, barcode: str_(r[0]), lokasi: str_(r[1]), qty: Number(r[2]) || 0 };
  });
}

function findStokRow_(stokData, barcode, lokasi) {
  for (var i = 0; i < stokData.length; i++) {
    if (stokData[i].barcode === barcode && stokData[i].lokasi === lokasi) return stokData[i];
  }
  return null;
}
