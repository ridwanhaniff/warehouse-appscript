/**
 * INA BASIC STORE WAREHOUSE - Dashboard & perapian Google Sheet
 *
 * File ini DITAMBAHKAN sebagai file baru di project Apps Script yang sama dengan Code.gs
 * (Files + > Script > beri nama "Dashboard"). Tidak mengubah API doPost.
 *
 * Cara pakai:
 *   1. Jalankan setupDashboard() sekali dari editor (atau lewat menu "Ina Warehouse" di Google Sheet).
 *   2. Selesai. Dashboard memakai RUMUS, jadi angkanya berubah otomatis setiap ada transaksi baru.
 *
 * Aman dijalankan berulang kali: dashboard disusun ulang, data di MASTER_PRODUK,
 * STOK_LOKASI, dan OPERATION_LOG tidak disentuh (hanya tampilannya dirapikan).
 */

var SHEET_DASH      = 'DASHBOARD';
var SHEET_DASH_DATA = '_DASH_DATA';   // sheet bantu (disembunyikan) untuk grafik tren
var TIMEZONE        = 'Asia/Jakarta';
var DEFAULT_PERIOD   = 30;            // periode analisis (hari)
var DEFAULT_LOWSTOCK = 10;            // batas stok menipis (unit)

var COLOR = {
  page:   '#f1f5f9',
  dark:   '#0f172a',
  text:   '#1e293b',
  muted:  '#64748b',
  line:   '#e2e8f0',
  indigo: '#4f46e5',
  green:  '#059669',
  red:    '#e11d48',
  amber:  '#d97706',
  sky:    '#0284c7'
};

// ===================== MENU =====================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ina Warehouse')
    .addItem('Susun ulang Dashboard', 'setupDashboard')
    .addItem('Rapikan sheet data', 'rapikanSheetData')
    .addToUi();
}

// ===================== ENTRY POINT =====================
function setupDashboard() {
  setup(); // pastikan 3 sheet data ada (dari Code.gs)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TIMEZONE);

  rapikanSheetData();
  buildDashData_(ss);
  buildDashboard_(ss);

  var sh = ss.getSheetByName(SHEET_DASH);
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1);
  SpreadsheetApp.flush();
}

// ===================== MERAPIKAN SHEET DATA =====================
function rapikanSheetData() {
  setup();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // MASTER_PRODUK
  var master = ss.getSheetByName(SHEET_MASTER);
  styleDataSheet_(master, [120, 160, 300, 110], COLOR.indigo);
  master.getRange('D2:D').setHorizontalAlignment('center');
  master.setConditionalFormatRules([
    formulaRule_('=AND($A2<>"",UPPER($D2&"")="FALSE")', '#fee2e2', '#991b1b', master.getRange('A2:D'), false)
  ]);

  // STOK_LOKASI
  var stok = ss.getSheetByName(SHEET_STOK);
  styleDataSheet_(stok, [120, 160, 110], COLOR.green);
  stok.getRange('C2:C').setNumberFormat('#,##0').setHorizontalAlignment('center');
  stok.setConditionalFormatRules([
    formulaRule_('=AND($C2<>"",$C2<=0)', '#fecaca', '#991b1b', stok.getRange('A2:C'), true),
    formulaRule_('=AND($C2<>"",$C2>0,$C2<=INDIRECT("DASHBOARD!K4"))', '#fef3c7', '#92400e', stok.getRange('A2:C'), false)
  ]);

  // OPERATION_LOG
  var log = ss.getSheetByName(SHEET_LOG);
  styleDataSheet_(log, [150, 200, 100, 120, 120, 120, 70, 210], COLOR.amber);
  log.getRange('A2:A').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  log.getRange('G2:G').setNumberFormat('#,##0');
  log.getRange('C2:C').setHorizontalAlignment('center');
  log.getRange('G2:G').setHorizontalAlignment('center');
  var c = log.getRange('C2:C');
  log.setConditionalFormatRules([
    textRule_('Masuk',      '#d1fae5', '#065f46', c),
    textRule_('Keluar',     '#ffe4e6', '#9f1239', c),
    textRule_('Mutasi',     '#e0f2fe', '#075985', c),
    textRule_('Opname',     '#fef3c7', '#92400e', c),
    textRule_('ProdukBaru', '#ede9fe', '#5b21b6', c)
  ]);
}

function styleDataSheet_(sheet, widths, tabColor) {
  var n = widths.length;
  var rows = Math.max(sheet.getMaxRows(), 2);

  for (var i = 0; i < n; i++) sheet.setColumnWidth(i + 1, widths[i]);
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 32);

  sheet.getBandings().forEach(function (b) { b.remove(); });
  var area = sheet.getRange(1, 1, rows, n);
  var band = area.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  band.setHeaderRowColor(COLOR.dark).setFirstRowColor('#ffffff').setSecondRowColor('#f8fafc');

  area.setFontFamily('Roboto').setFontSize(10).setVerticalAlignment('middle');
  sheet.getRange(1, 1, 1, n)
    .setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setTabColor(tabColor);
}

function formulaRule_(formula, bg, fg, range, bold) {
  var b = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula)
    .setBackground(bg).setFontColor(fg).setRanges([range]);
  if (bold) b.setBold(true);
  return b.build();
}

function textRule_(text, bg, fg, range) {
  return SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(text)
    .setBackground(bg).setFontColor(fg).setBold(true).setRanges([range]).build();
}

// ===================== DATA BANTU GRAFIK =====================
function buildDashData_(ss) {
  var d = ss.getSheetByName(SHEET_DASH_DATA) || ss.insertSheet(SHEET_DASH_DATA);
  d.clear();
  d.getRange('A1:C1').setValues([['Tanggal', 'Masuk', 'Keluar']]);

  var rows = [];
  for (var r = 2; r <= 15; r++) {
    var day = '(TODAY()-15+ROW())'; // baris 2 = 13 hari lalu ... baris 15 = hari ini
    var base = 'OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"';
    var range = '",OPERATION_LOG!$A$2:$A,">="&' + day + ',OPERATION_LOG!$A$2:$A,"<"&(' + day + '+1))';
    rows.push([
      '=TEXT(' + day + ',"dd MMM")',
      '=SUMIFS(' + base + 'Masuk' + range,
      '=SUMIFS(' + base + 'Keluar' + range
    ]);
  }
  d.getRange(2, 1, 14, 3).setFormulas(rows);
  d.hideSheet();
}

// ===================== DASHBOARD =====================
function topQuery_(activity) {
  return '=IFERROR(QUERY(OPERATION_LOG!$A$1:$H,"select D, count(B), sum(G) where C=\'' + activity +
    '\' and A >= datetime \'"&TEXT(TODAY()-$D$4,"yyyy-mm-dd")&" 00:00:00\' group by D order by sum(G) desc limit 10 ' +
    'label D \'SKU\', count(B) \'Transaksi\', sum(G) \'Total Qty\'",1),{"SKU","Transaksi","Total Qty"})';
}

function spark_(valueCol, r, first, last, color) {
  return '=IF(' + valueCol + r + '="","",IFERROR(SPARKLINE(' + valueCol + r +
    ',{"charttype","bar";"max",MAX($' + valueCol + '$' + first + ':$' + valueCol + '$' + last +
    ');"color1","' + color + '"}),""))';
}

function nameLookup_(skuCell, notFound) {
  return '=IF(' + skuCell + '="","",IFERROR(VLOOKUP(' + skuCell + ',MASTER_PRODUK!$A$2:$C,3,FALSE),"' + notFound + '"))';
}

function buildDashboard_(ss) {
  var NEED_COLS = 15, NEED_ROWS = 54;
  var period = DEFAULT_PERIOD, low = DEFAULT_LOWSTOCK;

  var sh = ss.getSheetByName(SHEET_DASH);
  if (sh) {
    // pertahankan nilai yang sudah diatur pengguna
    var pv = sh.getRange('D4').getValue();
    var lv = sh.getRange('K4').getValue();
    if (pv !== '' && isFinite(Number(pv)) && Number(pv) >= 1) period = Number(pv);
    if (lv !== '' && isFinite(Number(lv)) && Number(lv) >= 0) low = Number(lv);
  } else {
    sh = ss.insertSheet(SHEET_DASH, 0);
  }

  // ---- reset ----
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
  sh.getCharts().forEach(function (ch) { sh.removeChart(ch); });
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getDataRange().clearDataValidations();

  if (sh.getMaxColumns() < NEED_COLS) sh.insertColumnsAfter(sh.getMaxColumns(), NEED_COLS - sh.getMaxColumns());
  else if (sh.getMaxColumns() > NEED_COLS) sh.deleteColumns(NEED_COLS + 1, sh.getMaxColumns() - NEED_COLS);
  if (sh.getMaxRows() < NEED_ROWS) sh.insertRowsAfter(sh.getMaxRows(), NEED_ROWS - sh.getMaxRows());
  else if (sh.getMaxRows() > NEED_ROWS) sh.deleteRows(NEED_ROWS + 1, sh.getMaxRows() - NEED_ROWS);

  var R = function (a1) { return sh.getRange(a1); };

  // ---- dasar tampilan ----
  sh.setHiddenGridlines(true);
  sh.setTabColor(COLOR.indigo);
  sh.setColumnWidth(1, 14);
  for (var c = 2; c <= 7; c++) sh.setColumnWidth(c, 100);
  sh.setColumnWidth(8, 14);
  for (c = 9; c <= 14; c++) sh.setColumnWidth(c, 100);
  sh.setColumnWidth(15, 14);
  sh.setRowHeights(1, NEED_ROWS, 22);
  R('A1:O54').setBackground(COLOR.page).setFontFamily('Roboto').setFontSize(10)
    .setFontColor(COLOR.text).setVerticalAlignment('middle');

  [1, 5, 8, 21, 30, 38, 51].forEach(function (r) { sh.setRowHeight(r, 12); });
  sh.setRowHeight(2, 46);
  sh.setRowHeight(3, 24);
  sh.setRowHeight(7, 46);

  // ---- helper gaya ----
  function section(a1, textOrFormula) {
    var r = R(a1).merge();
    if (textOrFormula.charAt(0) === '=') r.setFormula(textOrFormula); else r.setValue(textOrFormula);
    r.setBackground('#e0e7ff').setFontColor('#312e81').setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('left');
    r.setBorder(null, null, true, null, null, null, COLOR.indigo, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
  function thead(a1) {
    R(a1).setBackground(COLOR.dark).setFontColor('#ffffff').setFontWeight('bold')
      .setFontSize(9).setHorizontalAlignment('center');
  }
  function tbody(a1) {
    R(a1).setBackground('#ffffff')
      .setBorder(null, null, true, null, null, true, COLOR.line, SpreadsheetApp.BorderStyle.SOLID);
  }
  function kpi(labelA1, valueA1, label, formula, color, fmt) {
    R(labelA1).merge().setValue(label).setFontSize(8).setFontWeight('bold').setFontColor(COLOR.muted)
      .setHorizontalAlignment('center').setBackground('#ffffff');
    R(valueA1).merge().setFormula(formula).setFontSize(22).setFontWeight('bold').setFontColor(color)
      .setHorizontalAlignment('center').setBackground('#ffffff').setNumberFormat(fmt || '#,##0');
    var box = R(labelA1.split(':')[0] + ':' + valueA1.split(':')[1]);
    box.setBorder(true, true, true, true, false, false, COLOR.line, SpreadsheetApp.BorderStyle.SOLID);
    R(labelA1).setBorder(true, null, null, null, null, null, color, SpreadsheetApp.BorderStyle.SOLID_THICK);
  }

  // ---- banner & pengaturan ----
  R('B2:N2').merge().setValue('  INA BASIC STORE  ·  WAREHOUSE DASHBOARD')
    .setBackground(COLOR.dark).setFontColor('#ffffff').setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('left');
  R('B3:N3').merge()
    .setFormula('="  Real-time dari OPERATION_LOG, STOK_LOKASI, MASTER_PRODUK  ·  Dihitung ulang: "&TEXT(NOW(),"dd MMM yyyy HH:mm")')
    .setBackground('#1e293b').setFontColor('#cbd5e1').setFontSize(9).setHorizontalAlignment('left');

  R('B4:C4').merge().setValue('Periode analisis (hari)').setFontWeight('bold').setFontColor(COLOR.muted);
  R('D4').setValue(period);
  R('I4:J4').merge().setValue('Batas stok menipis (unit)').setFontWeight('bold').setFontColor(COLOR.muted);
  R('K4').setValue(low);
  R('L4:N4').merge().setValue('◄ sel kuning bisa Anda ubah').setFontColor(COLOR.muted).setFontSize(9);
  ['D4', 'K4'].forEach(function (a1) {
    R(a1).setBackground('#fef9c3').setFontWeight('bold').setHorizontalAlignment('center')
      .setBorder(true, true, true, true, false, false, COLOR.amber, SpreadsheetApp.BorderStyle.SOLID);
  });
  R('D4').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 365).setAllowInvalid(false).setHelpText('Isi 1 - 365 hari').build());
  R('K4').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).setHelpText('Isi angka 0 atau lebih').build());

  // ---- KPI (baris 6-7) ----
  kpi('B6:C6', 'B7:C7', 'PRODUK AKTIF',
      '=SUMPRODUCT(--(MASTER_PRODUK!$A$2:$A<>""),--(UPPER(MASTER_PRODUK!$D$2:$D&"")="TRUE"))', COLOR.indigo);
  kpi('D6:E6', 'D7:E7', 'TOTAL STOK (UNIT)', '=SUM(STOK_LOKASI!$C$2:$C)', COLOR.text);
  kpi('F6:G6', 'F7:G7', 'MASUK HARI INI',
      '=SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Masuk",OPERATION_LOG!$A$2:$A,">="&TODAY())', COLOR.green);
  kpi('I6:J6', 'I7:J7', 'KELUAR HARI INI',
      '=SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Keluar",OPERATION_LOG!$A$2:$A,">="&TODAY())', COLOR.red);
  kpi('K6:L6', 'K7:L7', 'TRANSAKSI HARI INI',
      '=COUNTIFS(OPERATION_LOG!$A$2:$A,">="&TODAY())', COLOR.sky);
  kpi('M6:N6', 'M7:N7', 'SKU STOK MENIPIS',
      '=IFERROR(ROWS(QUERY(QUERY(STOK_LOKASI!$A$1:$C,"select A, sum(C) where A is not null group by A",1),' +
      '"select Col1 where Col2 <= "&$K$4,1))-1,0)', COLOR.amber);

  // ---- BARIS 9-20: Top Keluar | Stok Menipis ----
  section('B9:G9', '="TOP 10 BARANG PALING SERING KELUAR  ·  "&$D$4&" HARI TERAKHIR"');
  section('I9:N9', '="STOK MENIPIS  ·  BATAS ≤ "&$K$4&" UNIT"');

  R('B10:C10').merge().setValue('Nama Produk');
  R('G10').setValue('Visual');
  R('I10:J10').merge().setValue('Nama Produk');
  R('M10:N10').merge().setValue('Status');
  R('D10').setFormula(topQuery_('Keluar'));
  R('K10').setFormula('=IFERROR(QUERY(QUERY(STOK_LOKASI!$A$1:$C,"select A, sum(C) where A is not null group by A ' +
    'label sum(C) \'Stok\'",1),"select Col1, Col2 where Col2 <= "&$K$4&" order by Col2 asc limit 10",1),{"SKU","Stok"})');
  thead('B10:G10'); thead('I10:N10');

  for (var r = 11; r <= 20; r++) {
    R('B' + r + ':C' + r).merge().setFormula(nameLookup_('D' + r, '(tidak ada di master)')).setHorizontalAlignment('left');
    R('G' + r).setFormula(spark_('F', r, 11, 20, COLOR.red));
    R('I' + r + ':J' + r).merge().setFormula(nameLookup_('K' + r, '-')).setHorizontalAlignment('left');
    R('M' + r + ':N' + r).merge()
      .setFormula('=IF(L' + r + '="","",IF(L' + r + '<=0,"HABIS","MENIPIS"))').setHorizontalAlignment('center');
  }
  tbody('B11:G20'); tbody('I11:N20');
  R('D11:F20').setNumberFormat('#,##0').setHorizontalAlignment('center');
  R('K11:K20').setHorizontalAlignment('center');
  R('L11:L20').setNumberFormat('#,##0').setHorizontalAlignment('center');

  // ---- BARIS 22-37: Tren (grafik) | Stok per toko + Aktivitas ----
  section('B22:G22', 'TREN MASUK vs KELUAR  ·  14 HARI TERAKHIR');
  section('I22:N22', 'STOK PER TOKO');
  R('B23:G37').setBackground('#ffffff');

  R('I23').setFormula('=IFERROR(QUERY(STOK_LOKASI!$A$1:$C,"select B, sum(C), count(A) where B is not null and C > 0 ' +
    'group by B order by sum(C) desc limit 6 label B \'Toko\', sum(C) \'Total Unit\', count(A) \'SKU Ber-stok\'",1),' +
    '{"Toko","Total Unit","SKU Ber-stok"})');
  R('L23').setValue('Porsi');
  R('M23:N23').merge().setValue('Visual');
  thead('I23:N23');
  for (r = 24; r <= 29; r++) {
    R('L' + r).setFormula('=IF(J' + r + '="","",J' + r + '/SUM($J$24:$J$29))').setNumberFormat('0%');
    R('M' + r + ':N' + r).merge().setFormula(spark_('J', r, 24, 29, COLOR.indigo));
  }
  tbody('I24:N29');
  R('I24:I29').setFontWeight('bold');
  R('J24:L29').setHorizontalAlignment('center');
  R('J24:J29').setNumberFormat('#,##0');

  section('I31:N31', '="AKTIVITAS  ·  "&$D$4&" HARI TERAKHIR"');
  R('I32').setFormula('=IFERROR(QUERY(OPERATION_LOG!$A$1:$H,"select C, count(B) where C is not null and A >= datetime \'"' +
    '&TEXT(TODAY()-$D$4,"yyyy-mm-dd")&" 00:00:00\' group by C order by count(B) desc limit 5 ' +
    'label C \'Aktivitas\', count(B) \'Transaksi\'",1),{"Aktivitas","Transaksi"})');
  R('K32:N32').merge().setValue('Visual');
  thead('I32:N32');
  for (r = 33; r <= 37; r++) {
    R('K' + r + ':N' + r).merge().setFormula(spark_('J', r, 33, 37, COLOR.sky));
  }
  tbody('I33:N37');
  R('I33:I37').setFontWeight('bold');
  R('J33:J37').setHorizontalAlignment('center').setNumberFormat('#,##0');

  // ---- BARIS 39-50: Transaksi terbaru | Top Masuk ----
  section('B39:G39', 'TRANSAKSI TERBARU');
  section('I39:N39', '="TOP 10 BARANG PALING BANYAK MASUK  ·  "&$D$4&" HARI TERAKHIR"');

  R('B40').setFormula('=IFERROR(QUERY(OPERATION_LOG!$A$1:$H,"select A, C, D, G, E, F where A is not null ' +
    'order by A desc limit 10 label A \'Waktu\', C \'Aktivitas\', D \'SKU\', G \'Qty\', E \'Dari\', F \'Ke\'",1),' +
    '{"Waktu","Aktivitas","SKU","Qty","Dari","Ke"})');
  thead('B40:G40');
  tbody('B41:G50');
  R('B41:B50').setNumberFormat('dd MMM HH:mm').setHorizontalAlignment('center');
  R('C41:G50').setHorizontalAlignment('center');
  R('E41:E50').setNumberFormat('#,##0');

  R('I40:J40').merge().setValue('Nama Produk');
  R('N40').setValue('Visual');
  R('K40').setFormula(topQuery_('Masuk'));
  thead('I40:N40');
  for (r = 41; r <= 50; r++) {
    R('I' + r + ':J' + r).merge().setFormula(nameLookup_('K' + r, '(tidak ada di master)')).setHorizontalAlignment('left');
    R('N' + r).setFormula(spark_('M', r, 41, 50, COLOR.green));
  }
  tbody('I41:N50');
  R('K41:K50').setHorizontalAlignment('center');
  R('L41:M50').setNumberFormat('#,##0').setHorizontalAlignment('center');

  // ---- catatan kaki ----
  R('B52:N52').merge()
    .setValue('Catatan: "Keluar" hanya menghitung aktivitas Keluar (Pindah Lokasi dan Stock Opname tidak dihitung sebagai barang keluar). ' +
              'Semua angka dihitung otomatis dari data; ubah periode dan batas stok di sel kuning.')
    .setFontSize(8).setFontColor(COLOR.muted).setWrap(true).setVerticalAlignment('top');
  sh.setRowHeight(52, 30);

  // ---- format bersyarat ----
  sh.setConditionalFormatRules([
    textRule_('HABIS',   '#fee2e2', '#b91c1c', R('M11:M20')),
    textRule_('MENIPIS', '#fef3c7', '#b45309', R('M11:M20')),
    textRule_('Masuk',      '#d1fae5', '#065f46', R('C41:C50')),
    textRule_('Keluar',     '#ffe4e6', '#9f1239', R('C41:C50')),
    textRule_('Mutasi',     '#e0f2fe', '#075985', R('C41:C50')),
    textRule_('Opname',     '#fef3c7', '#92400e', R('C41:C50')),
    textRule_('ProdukBaru', '#ede9fe', '#5b21b6', R('C41:C50'))
  ]);

  // ---- grafik tren (14 hari) ----
  var data = ss.getSheetByName(SHEET_DASH_DATA);
  var chart = sh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(data.getRange('A1:C15'))
    .setNumHeaders(1)
    .setPosition(23, 2, 2, 2)
    .setOption('legend', { position: 'bottom' })
    .setOption('colors', ['#10b981', '#f43f5e'])
    .setOption('backgroundColor', '#ffffff')
    .setOption('width', 596)
    .setOption('height', 316)
    .setOption('vAxis', { minValue: 0, format: '#,##0', gridlines: { color: '#e5e7eb' } })
    .build();
  sh.insertChart(chart);

  // ---- proteksi (hanya peringatan; sel kuning tetap bebas diubah) ----
  try {
    var prot = sh.protect().setDescription('Dashboard otomatis').setWarningOnly(true);
    prot.setUnprotectedRanges([R('D4'), R('K4')]);
  } catch (e) {
    Logger.log('Proteksi dashboard dilewati: ' + e);
  }

  sh.setActiveRange(R('D4'));
}
