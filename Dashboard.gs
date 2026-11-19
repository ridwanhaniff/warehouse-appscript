/**
 * INA BASIC STORE - Executive Warehouse & Production Dashboard V6.6 (100% Made-to-Order / Gudang Polos)
 * Tata Letak Grid Eksekutif 2-Kolom Berdampingan (Side-by-Side Panels):
 * - Kanvas utama 15 kolom (A:O) dengan margin luar (A & O) serta pemisah tengah (H).
 * - Blok KPI Eksekutif 2x4 berfokus pada Saldo Gudang Polos, Stok Penjahit, Omzet/Profit, & Audit Polos Menipis.
 * - Band 1 (Kiri vs Kanan): Monitoring Stok Penjahit & Upah (UPPERCASE) || Rekap Saldo Gudang Polos per Kategori.
 * - Band 2 (Kiri vs Kanan): Top 10 Produk Terlaris Bulan Ini || Peringatan Stok Menipis Gudang Polos (<= 5 Pcs).
 * - Band 3 (Lebar Penuh B:N): 8 Log Aktivitas Terbaru (Pesanan Keluar & Gudang Penjahit).
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ina Basic Store')
    .addItem('Refresh Dashboard Ringkasan', 'buildDashboardSummary')
    .addItem('Aktifkan Auto-Refresh Dashboard (5 Menit)', 'aktifkanAutoRefreshDashboard')
    .addSeparator()
    .addItem('Migrasi Aman ke Sheet Produksi (V6.6)', 'migrasiKeProduksi')
    .addItem('Jalankan Setup / Perbarui Struktur', 'setup')
    .addItem('Hapus Tab Lama Tidak Terpakai', 'hapusTabTidakTerpakai')
    .addToUi();

  try {
    refreshDashboardIfDirty_();
  } catch (e) {
    Logger.log('onOpen refreshDashboardIfDirty_: ' + e);
  }
}

function parseSheetDateV2_(raw) {
  if (raw instanceof Date) return raw;
  if (typeof raw === 'number' && isFinite(raw)) {
    return new Date(Math.round((raw - 25569) * 86400 * 1000));
  }
  var str = textV2_(raw);
  if (!str) return new Date(NaN);
  var m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[,\sT]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    var day = parseInt(m[1], 10);
    var month = parseInt(m[2], 10) - 1;
    var year = parseInt(m[3], 10);
    var hour = m[4] ? parseInt(m[4], 10) : 0;
    var min = m[5] ? parseInt(m[5], 10) : 0;
    var sec = m[6] ? parseInt(m[6], 10) : 0;
    return new Date(year, month, day, hour, min, sec);
  }
  return new Date(str);
}

function markDashboardDirty_() {
  try {
    PropertiesService.getScriptProperties().setProperty('IBS_DASHBOARD_DIRTY', '1');
  } catch (e) {}
}

function refreshDashboardIfDirty_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName('DASHBOARD');
  var props = PropertiesService.getScriptProperties();
  var isDirty = props.getProperty('IBS_DASHBOARD_DIRTY') === '1';
  var isMissingOrEmpty = !dash || dash.getLastRow() < 5;

  if (!isDirty && !isMissingOrEmpty) {
    return false;
  }

  buildDashboardSummary();
  try {
    props.deleteProperty('IBS_DASHBOARD_DIRTY');
  } catch (e) {}
  return true;
}

function aktifkanAutoRefreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var res = installDashboardAutoTrigger_(true);
  try {
    ss.toast(res.message, 'Ina Basic Store', 6);
  } catch (e) {}
  return res;
}

function installDashboardAutoTrigger_(interactive) {
  try {
    var handlerName = 'refreshDashboardIfDirty_';
    var triggers = ScriptApp.getProjectTriggers();
    var exists = triggers.some(function (t) {
      return t.getHandlerFunction() === handlerName;
    });
    if (!exists) {
      ScriptApp.newTrigger(handlerName)
        .timeBased()
        .everyMinutes(5)
        .create();
    }
    return {
      success: true,
      message: 'Auto-refresh dashboard setiap 5 menit aktif.'
    };
  } catch (e) {
    var msg = 'Trigger 5 menit dilewati (izin script.scriptapp belum aktif). Dashboard tetap otomatis diperbarui saat spreadsheet dibuka atau lewat menu Refresh Dashboard Ringkasan.';
    Logger.log('installDashboardAutoTrigger_: ' + e);
    return {
      success: false,
      message: interactive ? msg : ''
    };
  }
}

function buildDashboardSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  assertV2SheetsReady_(ss);

  var dashName = 'DASHBOARD';
  var dash = ss.getSheetByName(dashName);
  if (!dash) {
    dash = ss.insertSheet(dashName, 0);
  }
  dash.clear();
  dash.setHiddenGridlines(true);
  dash.setTabColor('#0f172a');

  var fmtNum_ = function (n) {
    return Math.round(Number(n) || 0).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  };
  var fmtIdr_ = function (n) {
    return 'Rp ' + fmtNum_(n);
  };

  // ==========================================
  // 1. PENGAMBILAN & AGREGASI DATA V6.5 (MADE-TO-ORDER)
  // ==========================================
  var masterRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_MASTER));
  var varRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_VARIANTS));
  var logRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_LOG));
  var opnameRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_OPNAME_SESI));
  var masterTailorRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_MASTER_TAILOR));
  var stokTailorRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_STOK_TAILOR));
  var logTailorRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_LOG_TAILOR));
  var mappingRows = readObjectsV2_(ss.getSheetByName(IBS.SHEET_MAPPING_POLOS));

  var now = new Date();
  var curMonth = now.getMonth();
  var curYear = now.getFullYear();
  var periodeLabel = Utilities.formatDate(now, 'Asia/Jakarta', 'MMMM yyyy');

  // Master Produk Desain
  var masterMap = {};
  var activeSkuCount = 0;
  masterRows.forEach(function (m) {
    var bCode = upperV2_(m.Barcode);
    var isAct = booleanV2_(m.Status_Aktif, true);
    masterMap[bCode] = {
      barcode: textV2_(m.Barcode),
      nama: textV2_(m.Nama),
      aktif: isAct
    };
    if (isAct) activeSkuCount++;
  });

  // Variasi Produk Desain & Kelengkapan Harga
  var activeVariantCount = 0;
  var missingPriceCount = 0;
  var totalMarginPctSum = 0;
  var pricedVariantCount = 0;

  varRows.forEach(function (v) {
    var bCode = textV2_(v.Barcode_Induk);
    var mInfo = masterMap[upperV2_(bCode)];
    var active = booleanV2_(v.Status_Aktif, true) && (!mInfo || mInfo.aktif);
    var hpp = Number(v.HPP) || 0;
    var hj = Number(v.Harga_Jual) || 0;

    if (active) {
      activeVariantCount++;
      if (hpp <= 0 || hj <= 0) {
        missingPriceCount++;
      } else {
        totalMarginPctSum += ((hj - hpp) / hj);
        pricedVariantCount++;
      }
    }
  });

  // Agregasi Penjualan & Top Produk Bulan Ini
  var qtyKeluarBulanIni = 0;
  var trxKeluarBulanIni = 0;
  var omzetBulanIni = 0;
  var profitBulanIni = 0;
  var topSalesMap = {};

  logRows.forEach(function (lg) {
    if (upperV2_(lg.Aktivitas) !== 'KELUAR' || upperV2_(lg.Status) !== 'SUCCESS') return;
    var ts = parseSheetDateV2_(lg.Timestamp);
    if (isNaN(ts.getTime()) || ts.getMonth() !== curMonth || ts.getFullYear() !== curYear) return;

    var qty = Number(lg.Qty) || 0;
    var omzet = Number(lg.Total_Omzet) || 0;
    var profit = Number(lg.Total_Profit) || 0;
    var vid = textV2_(lg.ID_Variasi) || (textV2_(lg.Barcode) + '|' + textV2_(lg.Label_Variasi));

    qtyKeluarBulanIni += qty;
    trxKeluarBulanIni += 1;
    omzetBulanIni += omzet;
    profitBulanIni += profit;

    if (!topSalesMap[vid]) {
      topSalesMap[vid] = {
        barcode: textV2_(lg.Barcode),
        display: textV2_(lg.Produk_Tampilan) || textV2_(lg.Nama_Produk) || textV2_(lg.Barcode),
        qty: 0,
        omzet: 0,
        profit: 0
      };
    }
    topSalesMap[vid].qty += qty;
    topSalesMap[vid].omzet += omzet;
    topSalesMap[vid].profit += profit;
  });

  // Master Tailor & Saldo Gudang Polos per ID_Tailor
  var masterTailorMap = {};
  var activeTailorVariants = [];
  var activePolosSkuSet = {};

  masterTailorRows.forEach(function (mt) {
    var idT = textV2_(mt.ID_Tailor);
    if (!idT) return;
    var isAct = booleanV2_(mt.Status_Aktif, true);
    var rec = {
      id_tailor: idT,
      sku: textV2_(mt.Barcode_Kategori),
      nama: textV2_(mt.Nama_Kategori),
      warna: textV2_(mt.Warna) || '-',
      ukuran: textV2_(mt.Ukuran) || '-',
      upah: Number(mt.Upah_Per_Pcs) || 0,
      aktif: isAct
    };
    masterTailorMap[idT] = rec;
    if (isAct) {
      activeTailorVariants.push(rec);
      if (rec.sku) activePolosSkuSet[upperV2_(rec.sku)] = true;
    }
  });

  var tailorSummaryMap = {};
  var totalStokTailor = 0;
  var totalStokGudangPolos = 0;
  var totalNilaiUpahPolos = 0;
  var totalEstUpahBerjalan = 0;
  var gudangPolosPerTailorId = {};
  var gudangPolosByCategory = {};
  var gudangPolosNameUpper = upperV2_((typeof IBS !== 'undefined' && IBS.TAILOR_GUDANG_POLOS_NAME) ? IBS.TAILOR_GUDANG_POLOS_NAME : 'GUDANG POLOS');

  var ensureTailorSummary = function (namaRaw) {
    var nama = upperV2_(namaRaw) || 'TANPA NAMA';
    var key = nama;
    if (!tailorSummaryMap[key]) {
      tailorSummaryMap[key] = {
        nama_penjahit: nama,
        varianSet: {},
        stokReady: 0,
        estUpahBerjalan: 0,
        lolosBulanIni: 0,
        rejectBulanIni: 0,
        unpaidUpah: 0,
        paidBulanIni: 0
      };
    }
    return tailorSummaryMap[key];
  };

  stokTailorRows.forEach(function (st) {
    var idT = textV2_(st.ID_Tailor);
    var qty = Number(st.Stok_Ready) || 0;
    var mtInfo = masterTailorMap[idT];
    var upahSat = mtInfo ? mtInfo.upah : 0;

    if (upperV2_(st.Nama_Penjahit) === gudangPolosNameUpper) {
      if (idT) {
        gudangPolosPerTailorId[idT] = (gudangPolosPerTailorId[idT] || 0) + qty;
      }
      if (qty > 0) {
        totalStokGudangPolos += qty;
        totalNilaiUpahPolos += (qty * upahSat);

        var katKey = upperV2_(st.Barcode_Kategori || (mtInfo ? mtInfo.sku : idT));
        if (!gudangPolosByCategory[katKey]) {
          gudangPolosByCategory[katKey] = {
            sku: textV2_(st.Barcode_Kategori) || (mtInfo ? mtInfo.sku : '-'),
            nama: textV2_(st.Nama_Kategori) || (mtInfo ? mtInfo.nama : '-'),
            varianSet: {},
            totalQty: 0,
            nilaiUpah: 0
          };
        }
        gudangPolosByCategory[katKey].varianSet[idT] = true;
        gudangPolosByCategory[katKey].totalQty += qty;
        gudangPolosByCategory[katKey].nilaiUpah += (qty * upahSat);
      }
      return;
    }

    if (qty <= 0) return;

    var tObj = ensureTailorSummary(st.Nama_Penjahit);
    tObj.varianSet[idT || textV2_(st.Barcode_Kategori)] = true;
    tObj.stokReady += qty;
    tObj.estUpahBerjalan += (qty * upahSat);

    totalStokTailor += qty;
    totalEstUpahBerjalan += (qty * upahSat);
  });

  var qcLolosBulanIni = 0;
  var qcRejectBulanIni = 0;
  var totalUnpaidUpah = 0;
  var paidUpahBulanIni = 0;

  logTailorRows.forEach(function (lt) {
    var jenis = upperV2_(lt.Jenis_Aktivitas);
    var statusUpah = upperV2_(lt.Status_Upah);
    var qtyBagus = Number(lt.Qty_Bagus) || 0;
    var qtyReject = Number(lt.Qty_Reject) || 0;
    var totUpah = Number(lt.Total_Upah) || 0;
    var ts = parseSheetDateV2_(lt.Waktu);
    var isThisMonth = !isNaN(ts.getTime()) && ts.getMonth() === curMonth && ts.getFullYear() === curYear;

    if (jenis === 'TRANSFER_GUDANG') {
      var tObj = ensureTailorSummary(lt.Nama_Penjahit);
      if (isThisMonth) {
        tObj.lolosBulanIni += qtyBagus;
        tObj.rejectBulanIni += qtyReject;
        qcLolosBulanIni += qtyBagus;
        qcRejectBulanIni += qtyReject;
      }
      if (statusUpah === 'UNPAID' && totUpah > 0) {
        tObj.unpaidUpah += totUpah;
        totalUnpaidUpah += totUpah;
      }
      if (statusUpah === 'PAID' && isThisMonth && totUpah > 0) {
        tObj.paidBulanIni += totUpah;
        paidUpahBulanIni += totUpah;
      }
    }
  });

  var tailorList = Object.keys(tailorSummaryMap).map(function (k) { return tailorSummaryMap[k]; });
  tailorList.sort(function (a, b) {
    return (b.unpaidUpah - a.unpaidUpah) || (b.stokReady - a.stokReady);
  });

  var activeTailorCount = tailorList.filter(function (t) { return t.stokReady > 0 || t.unpaidUpah > 0; }).length;
  var unpaidTailorCount = tailorList.filter(function (t) { return t.unpaidUpah > 0; }).length;

  // Audit Stok Menipis Gudang Polos (<= 5 Pcs) dari MASTER_TAILOR & STOK_TAILOR ("GUDANG POLOS")
  var lowPolosRows = [];
  var lowPolosCount = 0;
  var outOfStockPolosCount = 0;

  activeTailorVariants.forEach(function (mt) {
    var sisaPolos = Number(gudangPolosPerTailorId[mt.id_tailor]) || 0;
    if (sisaPolos <= 5) {
      lowPolosCount++;
      if (sisaPolos === 0) outOfStockPolosCount++;
      lowPolosRows.push([
        mt.sku || mt.id_tailor,
        mt.nama,
        mt.warna,
        mt.ukuran,
        sisaPolos,
        sisaPolos === 0 ? 'HABIS (0 PCS)' : 'MENIPIS · JAHIT ULANG'
      ]);
    }
  });
  lowPolosRows.sort(function (a, b) {
    return (a[4] - b[4]) || a[1].localeCompare(b[1]) || a[2].localeCompare(b[2]);
  });

  var activeOpnameCount = opnameRows.filter(function (o) { return upperV2_(o.Status) === 'ACTIVE'; }).length;
  var activePolosOpnameCount = opnameRows.filter(function (o) {
    return upperV2_(o.Status) === 'ACTIVE' && upperV2_(o.Lokasi) === gudangPolosNameUpper;
  }).length;
  var avgMargin = pricedVariantCount > 0 ? (totalMarginPctSum / pricedVariantCount) : 0;
  var mappedVariantCount = mappingRows.filter(function (m) { return textV2_(m.Barcode_Produk) !== '' && textV2_(m.ID_Tailor) !== ''; }).length;

  // ==========================================
  // 2. KANVAS DASAR & LEBAR KOLOM (15 KOLOM A:O)
  // ==========================================
  var colWidths = [
    18,  // A: Margin Kiri
    125, // B: Panel Kiri Col 1
    185, // C: Panel Kiri Col 2
    90,  // D: Panel Kiri Col 3
    115, // E: Panel Kiri Col 4
    115, // F: Panel Kiri Col 5
    125, // G: Panel Kiri Col 6
    24,  // H: Pemisah Tengah (Gutter)
    115, // I: Panel Kanan Col 1
    185, // J: Panel Kanan Col 2
    90,  // K: Panel Kanan Col 3
    115, // L: Panel Kanan Col 4
    115, // M: Panel Kanan Col 5
    145, // N: Panel Kanan Col 6
    18   // O: Margin Kanan
  ];
  for (var c = 0; c < colWidths.length; c++) {
    dash.setColumnWidth(c + 1, colWidths[c]);
  }

  dash.getRange(1, 1, 80, 15).setBackground('#f1f5f9').setFontFamily('Arial');
  dash.setRowHeight(1, 10);

  // ==========================================
  // 3. HEADER UTAMA EKSEKUTIF (B2:N3)
  // ==========================================
  dash.getRange('B2:N2').merge()
    .setValue('  INA BASIC STORE · EXECUTIVE MADE-TO-ORDER & GUDANG POLOS DASHBOARD')
    .setFontSize(13).setFontWeight('bold')
    .setBackground('#0f172a').setFontColor('#ffffff')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  dash.setRowHeight(2, 38);

  dash.getRange('B3:N3').merge()
    .setValue('  Periode Laporan: ' + periodeLabel + '   •   Diperbarui: ' + Utilities.formatDate(now, 'Asia/Jakarta', 'dd MMM yyyy HH:mm:ss') + ' WIB   •   Sistem 100% Made-to-Order V6.6' + (activeOpnameCount > 0 ? '   •   Opname Aktif: ' + activeOpnameCount + ' Sesi' : ''))
    .setFontSize(9).setFontColor('#cbd5e1')
    .setBackground('#1e293b')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  dash.setRowHeight(3, 24);
  dash.setRowHeight(4, 12);

  // ==========================================
  // 4. KARTU KPI EKSEKUTIF 2x4 (BERDAMPINGAN)
  // ==========================================
  var renderModernKpi_ = function (topRow, startCol, spanCols, title, mainVal, numFormat, subText, stripeColor) {
    var stripeRange = dash.getRange(topRow, startCol, 1, spanCols);
    var titleRange = dash.getRange(topRow + 1, startCol, 1, spanCols);
    var valRange = dash.getRange(topRow + 2, startCol, 1, spanCols);
    var subRange = dash.getRange(topRow + 3, startCol, 1, spanCols);
    var cardBox = dash.getRange(topRow, startCol, 4, spanCols);

    stripeRange.merge().setBackground(stripeColor);

    titleRange.merge().setValue('  ' + title)
      .setFontSize(8).setFontWeight('bold')
      .setBackground('#ffffff').setFontColor('#64748b')
      .setHorizontalAlignment('left').setVerticalAlignment('bottom');

    valRange.merge().setValue(mainVal)
      .setFontSize(14).setFontWeight('bold')
      .setBackground('#ffffff').setFontColor('#0f172a')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    if (numFormat) {
      valRange.setNumberFormat('"  "' + numFormat);
    } else {
      valRange.setValue('  ' + mainVal);
    }

    subRange.merge().setValue('  ' + subText)
      .setFontSize(8.5).setFontColor('#475569')
      .setBackground('#ffffff')
      .setHorizontalAlignment('left').setVerticalAlignment('top');

    cardBox.setBorder(true, true, true, true, false, false, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  };

  renderModernKpi_(5, 2, 3, 'SALDO GUDANG POLOS SIAP PAKAI', totalStokGudangPolos, '#,##0" Pcs"', Object.keys(activePolosSkuSet).length + ' Kategori Polos · ' + activeTailorVariants.length + ' Varian Aktif', '#0d9488');
  renderModernKpi_(5, 5, 3, 'STOK DI TANGAN PENJAHIT', totalStokTailor, '#,##0" Pcs"', activeTailorCount + ' Penjahit · Est. Upah: ' + fmtIdr_(totalEstUpahBerjalan), '#2563eb');
  renderModernKpi_(5, 9, 3, 'OMZET PESANAN (' + periodeLabel.toUpperCase() + ')', omzetBulanIni, '"Rp "#,##0', fmtNum_(qtyKeluarBulanIni) + ' Pcs Terjual (' + trxKeluarBulanIni + ' Trx)', '#059669');
  renderModernKpi_(5, 12, 3, 'ESTIMASI PROFIT (' + periodeLabel.toUpperCase() + ')', profitBulanIni, '"Rp "#,##0', 'Margin Rata-Rata Aktif: ' + (avgMargin * 100).toFixed(1) + '%', '#16a34a');

  dash.setRowHeight(5, 5);
  dash.setRowHeight(6, 20);
  dash.setRowHeight(7, 30);
  dash.setRowHeight(8, 20);
  dash.setRowHeight(9, 10);

  renderModernKpi_(10, 2, 3, 'TAGIHAN UPAH BELUM BAYAR (UNPAID)', totalUnpaidUpah, '"Rp "#,##0', unpaidTailorCount + ' Penjahit Menunggu Slip Gaji', '#7c3aed');
  renderModernKpi_(10, 5, 3, 'SETORAN LOLOS QC (' + periodeLabel.toUpperCase() + ')', qcLolosBulanIni, '#,##0" Pcs"', 'Reject: ' + qcRejectBulanIni + ' Pcs · Lunas: ' + fmtIdr_(paidUpahBulanIni), '#334155');
  renderModernKpi_(10, 9, 3, 'KATALOG DESAIN & TAUTAN POLOS', activeVariantCount + ' Variasi', null, activeSkuCount + ' Desain · ' + mappedVariantCount + ' Tautan Polos Aktif', '#4f46e5');
  renderModernKpi_(10, 12, 3, 'STOK MENIPIS GUDANG POLOS (<= 5 PCS)', lowPolosCount + ' Varian Polos', null, outOfStockPolosCount + ' Habis (0 Pcs) · ' + (activePolosOpnameCount > 0 ? activePolosOpnameCount + ' Opname Polos Aktif' : missingPriceCount + ' Harga Kosong'), '#d97706');

  dash.setRowHeight(10, 5);
  dash.setRowHeight(11, 20);
  dash.setRowHeight(12, 30);
  dash.setRowHeight(13, 20);
  dash.setRowHeight(14, 14);

  // ==========================================
  // HELPER RENDER PANEL TABEL BERDAMPINGAN
  // ==========================================
  var renderSidePanel_ = function (startRow, startCol, titleText, headerBg, headers, dataRows, totalRow, numFormats, rightCols, emptyMsg) {
    var numCols = 6;
    var r = startRow;

    dash.getRange(r, startCol, 1, numCols).merge()
      .setValue('  ' + titleText)
      .setFontSize(9.5).setFontWeight('bold')
      .setBackground(headerBg).setFontColor('#ffffff')
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
    r++;

    dash.getRange(r, startCol, 1, numCols).setValues([headers])
      .setFontSize(8.5).setFontWeight('bold')
      .setBackground('#e2e8f0').setFontColor('#1e293b')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(true);
    r++;

    var dataStart = r;
    if (dataRows.length > 0) {
      dash.getRange(dataStart, startCol, dataRows.length, numCols).setValues(dataRows);

      var bgs = [];
      for (var i = 0; i < dataRows.length; i++) {
        var shade = (i % 2 === 0) ? '#ffffff' : '#f8fafc';
        bgs.push([shade, shade, shade, shade, shade, shade]);
      }
      dash.getRange(dataStart, startCol, dataRows.length, numCols)
        .setBackgrounds(bgs)
        .setFontSize(9).setFontColor('#1e293b')
        .setVerticalAlignment('middle').setWrap(true);

      Object.keys(numFormats || {}).forEach(function (colOffsetStr) {
        var offset = Number(colOffsetStr);
        dash.getRange(dataStart, startCol + offset, dataRows.length, 1).setNumberFormat(numFormats[colOffsetStr]);
      });
      (rightCols || []).forEach(function (offset) {
        dash.getRange(dataStart, startCol + offset, dataRows.length, 1).setHorizontalAlignment('right');
      });

      r += dataRows.length;

      if (totalRow) {
        dash.getRange(r, startCol, 1, numCols).setValues([totalRow])
          .setBackground('#f1f5f9').setFontWeight('bold').setFontSize(9).setFontColor('#0f172a')
          .setVerticalAlignment('middle');
        Object.keys(numFormats || {}).forEach(function (colOffsetStr) {
          var offset = Number(colOffsetStr);
          if (typeof totalRow[offset] === 'number') {
            dash.getRange(r, startCol + offset, 1, 1).setNumberFormat(numFormats[colOffsetStr]);
          }
        });
        (rightCols || []).forEach(function (offset) {
          dash.getRange(r, startCol + offset, 1, 1).setHorizontalAlignment('right');
        });
        r++;
      }
    } else {
      dash.getRange(dataStart, startCol, 1, numCols).merge()
        .setValue(emptyMsg)
        .setFontStyle('italic').setFontSize(9).setFontColor('#64748b')
        .setBackground('#ffffff').setHorizontalAlignment('center').setVerticalAlignment('middle');
      r++;
    }

    return {
      startRow: startRow,
      dataStartRow: dataStart,
      endRow: r - 1,
      hasData: dataRows.length > 0
    };
  };

  var padAndFrameSidePanel_ = function (panelInfo, startCol, targetEndRow) {
    var numCols = 6;
    if (panelInfo.endRow < targetEndRow) {
      var padCount = targetEndRow - panelInfo.endRow;
      dash.getRange(panelInfo.endRow + 1, startCol, padCount, numCols)
        .setBackground('#ffffff');
    }
    dash.getRange(panelInfo.startRow, startCol, targetEndRow - panelInfo.startRow + 1, numCols)
      .setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  };

  // ==========================================
  // 5. BAND 1 (KIRI vs KANAN): PENJAHIT & SALDO GUDANG POLOS
  // ==========================================
  var band1Start = 15;

  var p1Rows = tailorList.map(function (t) {
    var vCnt = Object.keys(t.varianSet).length;
    return [
      t.nama_penjahit,
      vCnt > 0 ? (vCnt + ' Variasi Aktif') : 'Riwayat Setoran',
      t.stokReady,
      t.estUpahBerjalan,
      t.lolosBulanIni,
      t.unpaidUpah
    ];
  });
  var p1Total = [
    'TOTAL PENJAHIT',
    tailorList.length + ' Penjahit',
    totalStokTailor,
    totalEstUpahBerjalan,
    qcLolosBulanIni,
    totalUnpaidUpah
  ];

  var panel1 = renderSidePanel_(
    band1Start, 2,
    '1. MONITORING STOK PENJAHIT & REKAP UPAH (MASTER TAILOR)',
    '#0f766e',
    ['Nama Penjahit', 'Varian Dikerjakan', 'Stok Ready', 'Est. Upah Jalan', 'Lolos QC (Bln)', 'Tagihan UNPAID'],
    p1Rows,
    p1Total,
    { 2: '#,##0', 3: '"Rp "#,##0', 4: '#,##0', 5: '"Rp "#,##0' },
    [2, 3, 4, 5],
    'Belum ada stok atau tagihan penjahit aktif.'
  );

  if (panel1.hasData) {
    for (var i1 = 0; i1 < p1Rows.length; i1++) {
      if (p1Rows[i1][5] > 0) {
        dash.getRange(panel1.dataStartRow + i1, 7).setBackground('#f5f3ff').setFontColor('#5b21b6').setFontWeight('bold');
      }
    }
  }

  var polosCatList = Object.keys(gudangPolosByCategory).map(function (k) { return gudangPolosByCategory[k]; });
  polosCatList.sort(function (a, b) { return b.totalQty - a.totalQty; });

  var p2Rows = polosCatList.map(function (cat) {
    var share = totalStokGudangPolos > 0 ? (cat.totalQty / totalStokGudangPolos) : 0;
    var vCount = Object.keys(cat.varianSet).length;
    return [
      cat.sku,
      cat.nama,
      vCount + ' Varian',
      cat.totalQty,
      cat.nilaiUpah,
      share
    ];
  });
  var p2Total = [
    'TOTAL POLOS',
    polosCatList.length + ' Kategori Polos',
    activeTailorVariants.length + ' Varian',
    totalStokGudangPolos,
    totalNilaiUpahPolos,
    totalStokGudangPolos > 0 ? 1 : 0
  ];

  var panel2 = renderSidePanel_(
    band1Start, 9,
    '2. SALDO GUDANG POLOS SIAP PAKAI (PER KATEGORI)',
    '#1e3a8a',
    ['Kode SKU', 'Kategori Kaos Polos', 'Varian Ada', 'Stok Siap (Pcs)', 'Nilai Upah Jahit', 'Porsi Stok'],
    p2Rows,
    p2Total,
    { 0: '@', 3: '#,##0', 4: '"Rp "#,##0', 5: '0.0%' },
    [3, 4, 5],
    'Belum ada saldo kaos polos di Gudang Polos.'
  );

  var band1End = Math.max(panel1.endRow, panel2.endRow);
  dash.setRowHeight(band1Start, 28);
  dash.setRowHeight(band1Start + 1, 30);
  for (var r1 = band1Start + 2; r1 <= band1End; r1++) {
    dash.setRowHeight(r1, 25);
  }
  padAndFrameSidePanel_(panel1, 2, band1End);
  padAndFrameSidePanel_(panel2, 9, band1End);
  dash.setRowHeight(band1End + 1, 14);

  // ==========================================
  // 6. BAND 2 (KIRI vs KANAN): TOP 10 TERLARIS & STOK MENIPIS GUDANG POLOS
  // ==========================================
  var band2Start = band1End + 2;

  var topSalesList = Object.keys(topSalesMap).map(function (k) { return topSalesMap[k]; });
  topSalesList.sort(function (a, b) { return (b.qty - a.qty) || (b.omzet - a.omzet); });
  var top10 = topSalesList.slice(0, 10);

  var sumTopQty = 0;
  var sumTopOmzet = 0;
  var sumTopProfit = 0;
  var p3Rows = top10.map(function (item, idx) {
    sumTopQty += item.qty;
    sumTopOmzet += item.omzet;
    sumTopProfit += item.profit;
    var mPct = item.omzet > 0 ? (item.profit / item.omzet) : 0;
    return [
      '#' + (idx + 1) + ' · ' + item.barcode,
      item.display,
      item.qty,
      item.omzet,
      item.profit,
      mPct
    ];
  });
  var p3Total = [
    'TOTAL TOP ' + top10.length,
    'Akumulasi Produk Terlaris',
    sumTopQty,
    sumTopOmzet,
    sumTopProfit,
    sumTopOmzet > 0 ? (sumTopProfit / sumTopOmzet) : 0
  ];

  var panel3 = renderSidePanel_(
    band2Start, 2,
    '3. TOP 10 PRODUK & VARIASI TERLARIS (' + periodeLabel.toUpperCase() + ')',
    '#065f46',
    ['Rank & Barcode', 'Nama Produk & Variasi', 'Terjual', 'Total Omzet', 'Estimasi Profit', 'Margin'],
    p3Rows,
    p3Total,
    { 0: '@', 2: '#,##0', 3: '"Rp "#,##0', 4: '"Rp "#,##0', 5: '0.0%' },
    [2, 3, 4, 5],
    'Belum ada transaksi penjualan pada periode ini.'
  );
  if (panel3.hasData) {
    dash.getRange(panel3.dataStartRow, 2, p3Rows.length, 1).setFontWeight('bold');
  }

  var displayLowPolos = lowPolosRows.slice(0, 14);
  var panel4 = renderSidePanel_(
    band2Start, 9,
    '4. PERINGATAN STOK MENIPIS GUDANG POLOS (<= 5 PCS)',
    '#92400e',
    ['Kode SKU', 'Kategori Kaos Polos', 'Warna', 'Ukuran', 'Sisa Polos', 'Status Gudang'],
    displayLowPolos,
    null,
    { 0: '@', 4: '#,##0' },
    [4],
    'Semua varian kaos polos di Gudang Polos memiliki stok > 5 pcs.'
  );

  if (panel4.hasData) {
    dash.getRange(panel4.dataStartRow, 11, displayLowPolos.length, 2).setHorizontalAlignment('center');
    dash.getRange(panel4.dataStartRow, 14, displayLowPolos.length, 1).setHorizontalAlignment('center');
    for (var i4 = 0; i4 < displayLowPolos.length; i4++) {
      var sisaQty = displayLowPolos[i4][4];
      var stCell = dash.getRange(panel4.dataStartRow + i4, 14);
      if (sisaQty === 0) {
        stCell.setBackground('#fee2e2').setFontColor('#991b1b').setFontWeight('bold');
      } else {
        stCell.setBackground('#fef3c7').setFontColor('#92400e').setFontWeight('bold');
      }
    }
  }

  var band2End = Math.max(panel3.endRow, panel4.endRow);
  dash.setRowHeight(band2Start, 28);
  dash.setRowHeight(band2Start + 1, 30);
  for (var r2 = band2Start + 2; r2 <= band2End; r2++) {
    dash.setRowHeight(r2, 25);
  }
  padAndFrameSidePanel_(panel3, 2, band2End);
  padAndFrameSidePanel_(panel4, 9, band2End);
  dash.setRowHeight(band2End + 1, 14);

  // ==========================================
  // 7. BAND 3 (LEBAR PENUH B:N): 8 LOG AKTIVITAS TERBARU
  // ==========================================
  var band3Start = band2End + 2;

  dash.getRange(band3Start, 2, 1, 13).merge()
    .setValue('  5. LOG AKTIVITAS TERBARU (PESANAN KELUAR & GUDANG PENJAHIT)')
    .setFontSize(9.5).setFontWeight('bold')
    .setBackground('#1e293b').setFontColor('#ffffff')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  dash.setRowHeight(band3Start, 28);

  var logHeaderRow = band3Start + 1;
  dash.getRange(logHeaderRow, 2).setValue('Waktu (WIB)');
  dash.getRange(logHeaderRow, 3).setValue('Modul');
  dash.getRange(logHeaderRow, 4, 1, 2).merge().setValue('Aktivitas');
  dash.getRange(logHeaderRow, 6, 1, 3).merge().setValue('Detail Produk / Item Jahit');
  dash.getRange(logHeaderRow, 9, 1, 2).merge().setValue('Asal / Penjahit');
  dash.getRange(logHeaderRow, 11, 1, 2).merge().setValue('Tujuan / Status Upah');
  dash.getRange(logHeaderRow, 13).setValue('Qty (Pcs)');
  dash.getRange(logHeaderRow, 14).setValue('Admin / Pelapor');

  dash.getRange(logHeaderRow, 2, 1, 13)
    .setFontSize(8.5).setFontWeight('bold')
    .setBackground('#e2e8f0').setFontColor('#1e293b')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  dash.setRowHeight(logHeaderRow, 28);

  var recentEvents = [];
  logRows.slice(-15).forEach(function (lg) {
    var ts = parseSheetDateV2_(lg.Timestamp);
    if (isNaN(ts.getTime())) return;
    recentEvents.push({
      ts: ts,
      modul: 'PESANAN',
      aktivitas: textV2_(lg.Aktivitas),
      detail: textV2_(lg.Produk_Tampilan) || textV2_(lg.Barcode),
      asal: textV2_(lg.Lokasi_Asal) || 'Gudang Polos',
      tujuan: textV2_(lg.Lokasi_Tujuan) || textV2_(lg.Status),
      qty: Number(lg.Qty) || 0,
      pelapor: textV2_(lg.Nama_Pelapor) || '-'
    });
  });

  logTailorRows.slice(-15).forEach(function (lt) {
    var ts = parseSheetDateV2_(lt.Waktu);
    if (isNaN(ts.getTime())) return;
    recentEvents.push({
      ts: ts,
      modul: 'POLOS/JAHIT',
      aktivitas: textV2_(lt.Jenis_Aktivitas),
      detail: textV2_(lt.Detail_Item),
      asal: upperV2_(lt.Nama_Penjahit) || '-',
      tujuan: textV2_(lt.Status_Upah) || '-',
      qty: Number(lt.Qty_Bagus) || 0,
      pelapor: textV2_(lt.ID_Slip) || upperV2_(lt.Nama_Penjahit) || '-'
    });
  });

  recentEvents.sort(function (a, b) { return b.ts.getTime() - a.ts.getTime(); });
  var topRecent = recentEvents.slice(0, 8);
  var band3End = logHeaderRow;

  if (topRecent.length > 0) {
    topRecent.forEach(function (ev, idx) {
      var r = logHeaderRow + 1 + idx;
      var rowBg = (idx % 2 === 0) ? '#ffffff' : '#f8fafc';
      dash.getRange(r, 2, 1, 13).setBackground(rowBg).setFontSize(9).setFontColor('#1e293b').setVerticalAlignment('middle');

      dash.getRange(r, 2).setValue(Utilities.formatDate(ev.ts, 'Asia/Jakarta', 'dd/MM HH:mm')).setHorizontalAlignment('center');
      dash.getRange(r, 3).setValue(ev.modul).setHorizontalAlignment('center').setFontWeight('bold')
        .setFontColor(ev.modul === 'POLOS/JAHIT' ? '#0f766e' : '#1d4ed8');
      dash.getRange(r, 4, 1, 2).merge().setValue(ev.aktivitas).setHorizontalAlignment('center');
      dash.getRange(r, 6, 1, 3).merge().setValue(ev.detail).setWrap(true);
      dash.getRange(r, 9, 1, 2).merge().setValue(ev.asal);
      dash.getRange(r, 11, 1, 2).merge().setValue(ev.tujuan);
      dash.getRange(r, 13).setValue(ev.qty).setNumberFormat('#,##0').setHorizontalAlignment('right').setFontWeight('bold');
      dash.getRange(r, 14).setValue(ev.pelapor);

      dash.setRowHeight(r, 24);
      band3End = r;
    });
  } else {
    var emptyLogRow = logHeaderRow + 1;
    dash.getRange(emptyLogRow, 2, 1, 13).merge()
      .setValue('Belum ada riwayat aktivitas transaksi yang tercatat.')
      .setFontStyle('italic').setFontSize(9).setFontColor('#64748b')
      .setBackground('#ffffff').setHorizontalAlignment('center').setVerticalAlignment('middle');
    dash.setRowHeight(emptyLogRow, 26);
    band3End = emptyLogRow;
  }

  dash.getRange(band3Start, 2, band3End - band3Start + 1, 13)
    .setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);

  dash.setRowHeight(band3End + 1, 18);
  dash.getRange(band3End + 1, 1, 1, 15).setBackground('#f1f5f9');

  var maxRows = dash.getMaxRows();
  if (maxRows > band3End + 1) {
    dash.getRange(band3End + 2, 1, maxRows - (band3End + 1), dash.getMaxColumns()).setBackground('#ffffff');
  }
  var maxCols = dash.getMaxColumns();
  if (maxCols > 15) {
    dash.getRange(1, 16, maxRows, maxCols - 15).setBackground('#ffffff');
  }

  dash.setFrozenRows(3);
  dash.setFrozenColumns(0);

  try {
    PropertiesService.getScriptProperties().deleteProperty('IBS_DASHBOARD_DIRTY');
  } catch (e) {}

  SpreadsheetApp.flush();
  try {
    ss.toast('Dashboard Made-to-Order & Audit Gudang Polos V6.6 berhasil diperbarui.', 'Ina Basic Store', 4);
  } catch (e) {}
}