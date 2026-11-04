/**
 * INA BASIC STORE WAREHOUSE - Dashboard
 * Dashboard menampilkan ringkasan walaupun belum ada transaksi.
 */

var SHEET_DASH = 'DASHBOARD';
var SHEET_DASH_DATA = '_DASH_DATA';
var TIMEZONE = 'Asia/Jakarta';

var COLOR = {
  page: '#f1f5f9', dark: '#0f172a', text: '#1e293b', muted: '#64748b',
  line: '#cbd5e1', indigo: '#4f46e5', green: '#059669', red: '#e11d48',
  amber: '#d97706', sky: '#0284c7'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ina Warehouse')
    .addItem('Bangun / Perbarui Dashboard', 'setupDashboard')
    .addToUi();
}

function setupDashboard() {
  setup();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TIMEZONE);

  // Pastikan kedua sheet sudah ada sebelum formula antar-sheet dibuat.
  // _DASH_DATA memakai referensi DASHBOARD!D4, jadi tab DASHBOARD harus dibuat dahulu.
  if (!ss.getSheetByName(SHEET_DASH)) ss.insertSheet(SHEET_DASH);
  if (!ss.getSheetByName(SHEET_DASH_DATA)) ss.insertSheet(SHEET_DASH_DATA);

  buildDashData_(ss);
  buildDashboard_(ss);
  var sh = ss.getSheetByName(SHEET_DASH);
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1);
  SpreadsheetApp.flush();
}

function buildDashData_(ss) {
  var d = ss.getSheetByName(SHEET_DASH_DATA) || ss.insertSheet(SHEET_DASH_DATA);
  d.clear();
  d.getRange('A1:H1').setValues([['Tanggal', 'Masuk', 'Keluar', '', 'TGL_AWAL', 'TGL_AKHIR', '', 'LOKASI']]);

  var rows = [];
  for (var r = 2; r <= 15; r++) {
    var dateFormula = '=TODAY()-' + (15 - r);
    var dateRef = '$A' + r;
    rows.push([
      dateFormula,
      '=SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Masuk",OPERATION_LOG!$A$2:$A,">="&' + dateRef + ',OPERATION_LOG!$A$2:$A,"<"&(' + dateRef + '+1))',
      '=SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Keluar",OPERATION_LOG!$A$2:$A,">="&' + dateRef + ',OPERATION_LOG!$A$2:$A,"<"&(' + dateRef + '+1))'
    ]);
  }
  d.getRange(2, 1, 14, 3).setFormulas(rows);
  d.getRange('A2:A15').setNumberFormat('dd MMM');
  d.getRange('B2:C15').setNumberFormat('0');

  d.getRange('E2').setFormula(
    '=IFS(DASHBOARD!$D$4="Hari Ini",TODAY(),DASHBOARD!$D$4="Kemarin",TODAY()-1,' +
    'DASHBOARD!$D$4="7 Hari Terakhir",TODAY()-7,DASHBOARD!$D$4="30 Hari Terakhir",TODAY()-30,' +
    'DASHBOARD!$D$4="Bulan Ini",EOMONTH(TODAY(),-1)+1,DASHBOARD!$D$4="Semua Waktu",DATE(2000,1,1),TRUE,TODAY()-30)'
  );
  d.getRange('F2').setFormula('=IF(DASHBOARD!$D$4="Kemarin",TODAY(),TODAY()+1)');
  d.getRange('H2').setValue('Semua Lokasi');
  d.getRange('H3').setFormula('=IFERROR(QUERY(STOK_LOKASI!B2:B,"select Col1 where Col1 is not null group by Col1 label Col1 \'\'",0),"")');
  d.hideSheet();
}

function topQuery_(activity) {
  var locFilter = activity === 'Keluar' ? 'E' : 'F';
  return '=IFERROR(QUERY(OPERATION_LOG!$A$1:$H,"select D, count(B), sum(G) where C=\'' + activity +
    '\' and A >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and A < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' ' +
    '"&IF($G$4="Semua Lokasi",""," and ' + locFilter + '=\'"&$G$4&"\' ")&" ' +
    'group by D order by sum(G) desc limit 10 label D \'Barcode\', count(B) \'Transaksi\', sum(G) \'Total Qty\'",1),{"Barcode","Transaksi","Total Qty"})';
}

function spark_(valueCol, r, first, last, color) {
  return '=IF(' + valueCol + r + '="","",IFERROR(SPARKLINE(' + valueCol + r +
    ',{"charttype","bar";"max",MAX($' + valueCol + '$' + first + ':$' + valueCol + '$' + last + ');"color1","' + color + '"}),""))';
}

function nameLookup_(barcodeCell, notFound) {
  // MITIGASI KETAT: Membungkus barcode dengan fungsi TO_TEXT()
  // Ini memaksa Google Sheets mencocokkan Angka dan Teks dengan sempurna
  return '=IF(' + barcodeCell + '="","",IFERROR(VLOOKUP(TO_TEXT(' + barcodeCell + '),MASTER_PRODUK!$A$2:$B,2,FALSE),"' + notFound + '"))';
}

function buildDashboard_(ss) {
  var sh = ss.getSheetByName(SHEET_DASH) || ss.insertSheet(SHEET_DASH, 0);
  // Clear chart objects too; otherwise repeated rebuilds can leave overlapping charts.
  sh.getCharts().forEach(function (chart) { sh.removeChart(chart); });
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getDataRange().clearDataValidations();

  var R = function (a1) { return sh.getRange(a1); };

  sh.setColumnWidth(1, 14);
  for (var c = 2; c <= 7; c++) sh.setColumnWidth(c, 100);
  sh.setColumnWidth(8, 14);
  for (c = 9; c <= 14; c++) sh.setColumnWidth(c, 100);
  sh.setColumnWidth(15, 14);
  sh.setRowHeights(1, 54, 22);
  [1, 5, 8, 21, 30, 38, 51].forEach(function (r) { sh.setRowHeight(r, 12); });
  sh.setRowHeight(2, 46);
  sh.setRowHeight(3, 24);
  sh.setRowHeight(7, 54);

  R('A1:O54').setBackground(COLOR.page).setFontFamily('Roboto').setFontSize(10).setVerticalAlignment('middle');
  sh.setHiddenGridlines(true);

  R('B2:N2').merge().setValue('  INA BASIC STORE  ·  WAREHOUSE DASHBOARD').setBackground(COLOR.dark).setFontColor('#ffffff').setFontSize(18).setFontWeight('bold');
  R('B3:N3').merge().setFormula('="  Dihitung ulang: "&TEXT(NOW(),"dd MMM yyyy HH:mm")').setBackground('#1e293b').setFontColor('#cbd5e1');

  R('B4:C4').merge().setValue('Filter Waktu');
  R('D4').setValue('30 Hari Terakhir').setBackground('#fef9c3');
  R('E4:F4').merge().setValue('Filter Lokasi').setHorizontalAlignment('right');
  R('G4:H4').merge().setValue('Semua Lokasi').setBackground('#fef9c3');
  R('I4:J4').merge().setValue('Batas Kritis').setHorizontalAlignment('right');
  R('K4').setValue(10).setBackground('#fef9c3');

  R('D4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Hari Ini', 'Kemarin', '7 Hari Terakhir', '30 Hari Terakhir', 'Bulan Ini', 'Semua Waktu'], true).build());
  R('G4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName(SHEET_DASH_DATA).getRange('H2:H100'), true).build());

  function kpi(lblRange, valRange, label, formula, col) {
    var labelCell = R(lblRange);
    if (String(label).charAt(0) === '=') labelCell.merge().setFormula(label);
    else labelCell.merge().setValue(label);
    labelCell.setFontSize(8).setFontColor(COLOR.muted).setHorizontalAlignment('center');
    R(valRange).merge().setFormula(formula).setFontSize(22).setFontColor(col).setHorizontalAlignment('center').setFontWeight('bold');
    R(lblRange.split(':')[0] + ':' + valRange.split(':')[1]).setBorder(true, true, true, true, false, false, COLOR.line, SpreadsheetApp.BorderStyle.SOLID);
  }

  function section(a1, formulaOrText) {
    var rg = R(a1);
    rg.merge();
    if (String(formulaOrText).charAt(0) === '=') rg.setFormula(formulaOrText);
    else rg.setValue(formulaOrText);
    rg.setBackground(COLOR.dark).setFontColor('#ffffff').setFontWeight('bold')
      .setBorder(true, true, false, true, null, null, COLOR.dark, SpreadsheetApp.BorderStyle.SOLID);
  }

  // KPI: master sekarang hanya memiliki Barcode, Nama, Status_Aktif.
  kpi('B6:C6', 'B7:C7', 'PRODUK AKTIF', '=COUNTIF(MASTER_PRODUK!$C$2:$C,TRUE)', COLOR.indigo);
  kpi('D6:E6', 'D7:E7', 'TOTAL STOK', '=IF($G$4="Semua Lokasi",SUM(STOK_LOKASI!$C$2:$C),SUMIFS(STOK_LOKASI!$C$2:$C,STOK_LOKASI!$B$2:$B,$G$4))', COLOR.text);
  kpi('F6:G6', 'F7:G7', '="MASUK ("&UPPER($D$4)&")"', '=IF($G$4="Semua Lokasi",SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Masuk",OPERATION_LOG!$A$2:$A,">="&_DASH_DATA!$E$2,OPERATION_LOG!$A$2:$A,"<"&_DASH_DATA!$F$2),SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Masuk",OPERATION_LOG!$A$2:$A,">="&_DASH_DATA!$E$2,OPERATION_LOG!$A$2:$A,"<"&_DASH_DATA!$F$2,OPERATION_LOG!$F$2:$F,$G$4))', COLOR.green);
  kpi('I6:J6', 'I7:J7', '="KELUAR ("&UPPER($D$4)&")"', '=IF($G$4="Semua Lokasi",SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Keluar",OPERATION_LOG!$A$2:$A,">="&_DASH_DATA!$E$2,OPERATION_LOG!$A$2:$A,"<"&_DASH_DATA!$F$2),SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Keluar",OPERATION_LOG!$A$2:$A,">="&_DASH_DATA!$E$2,OPERATION_LOG!$A$2:$A,"<"&_DASH_DATA!$F$2,OPERATION_LOG!$E$2:$E,$G$4))', COLOR.red);
  kpi('K6:L6', 'K7:L7', '="TRX ("&UPPER($D$4)&")"', '=IFERROR(ROWS(QUERY(OPERATION_LOG!$A$2:$H,"select A where A >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and A < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' "&IF($G$4="Semua Lokasi",""," and (E=\'"&$G$4&"\' or F=\'"&$G$4&"\') ")&" ",0)),0)', COLOR.sky);
  kpi('M6:N6', 'M7:N7', 'STOK KRITIS', '=IFERROR(ROWS(QUERY(QUERY(STOK_LOKASI!$A$1:$C,"select A, sum(C) where A is not null "&IF($G$4="Semua Lokasi",""," and B=\'"&$G$4&"\' ")&" group by A",1),"select Col1 where Col2 <= "&$K$4,1))-1,0)', COLOR.amber);

  // TOP BARANG KELUAR
  section('B9:G9', '="TOP 10 BARANG KELUAR  ·  "&UPPER($D$4)');
  R('B10:C10').merge().setValue('Nama Produk').setBackground('#e2e8f0').setFontWeight('bold');
  R('D10').setFormula(topQuery_('Keluar')).setBackground('#e2e8f0').setFontWeight('bold');
  R('G10').setValue('Visual Bar').setBackground('#e2e8f0').setFontWeight('bold');
  for (var r = 11; r <= 20; r++) {
    R('B' + r + ':C' + r).merge().setFormula(nameLookup_('D' + r, '(tidak ada)')).setBackground('#ffffff');
    R('G' + r).setFormula(spark_('F', r, 11, 20, COLOR.red)).setBackground('#ffffff');
  }

  // STOK MENIPIS / HABIS
  section('I9:N9', '="STOK KRITIS  ·  BATAS ≤ "&$K$4&" UNIT"');
  R('I10:J10').merge().setValue('Nama Produk').setBackground('#e2e8f0').setFontWeight('bold');
  R('K10').setFormula('=IFERROR(QUERY(QUERY(STOK_LOKASI!$A$1:$C,"select A, sum(C) where A is not null "&IF($G$4="Semua Lokasi",""," and B=\'"&$G$4&"\' ")&" group by A label sum(C) \'Stok\'",1),"select Col1, Col2 where Col2 <= "&$K$4&" order by Col2 asc limit 10",1),{"Barcode","Stok"})').setBackground('#e2e8f0').setFontWeight('bold');
  R('M10:N10').merge().setValue('Status').setBackground('#e2e8f0').setFontWeight('bold');
  for (r = 11; r <= 20; r++) {
    R('I' + r + ':J' + r).merge().setFormula(nameLookup_('K' + r, '-')).setBackground('#ffffff');
    R('M' + r + ':N' + r).merge().setFormula('=IF(L' + r + '="","",IF(L' + r + '<=0,"HABIS","MENIPIS"))').setBackground('#ffffff').setHorizontalAlignment('center');
  }

  // TREND GRAFIK
  section('B22:G22', '="TREN MASUK vs KELUAR  ·  14 HARI TERAKHIR"');
  R('B23:G37').setBackground('#ffffff');
  R('B38:G38').merge().setFormula('=IF(COUNTA(OPERATION_LOG!A2:A)=0,"Belum ada transaksi. Grafik akan mulai bergerak setelah transaksi pertama.","Grafik menampilkan jumlah unit Masuk dan Keluar per hari.")').setFontColor(COLOR.muted).setFontSize(8).setHorizontalAlignment('center');

  // PERSEBARAN STOK
  section('I22:N22', '="PERSEBARAN STOK PER LOKASI"');
  R('I23').setFormula('=IFERROR(QUERY(STOK_LOKASI!$A$1:$C,"select B, sum(C), count(A) where B is not null and C > 0 "&IF($G$4="Semua Lokasi",""," and B=\'"&$G$4&"\' ")&" group by B order by sum(C) desc limit 6 label B \'Toko/Rak\', sum(C) \'Total Unit\', count(A) \'Barcode\'",1),{"Toko/Rak","Total Unit","Barcode"})').setBackground('#e2e8f0').setFontWeight('bold');
  R('L23').setValue('Porsi').setBackground('#e2e8f0').setFontWeight('bold');
  R('M23:N23').merge().setValue('Visual Bar').setBackground('#e2e8f0').setFontWeight('bold');
  for (r = 24; r <= 29; r++) {
    R('L' + r).setFormula('=IF(J' + r + '="","",IFERROR(J' + r + '/SUM($J$24:$J$29),0))').setNumberFormat('0%').setBackground('#ffffff');
    R('M' + r + ':N' + r).merge().setFormula(spark_('J', r, 24, 29, COLOR.indigo)).setBackground('#ffffff');
  }

  // AKTIVITAS
  section('I31:N31', '="AKTIVITAS  ·  "&UPPER($D$4)');
  R('I32').setFormula('=IFERROR(QUERY(OPERATION_LOG!$A$1:$H,"select C, count(B) where C is not null and A >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and A < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' "&IF($G$4="Semua Lokasi",""," and (E=\'"&$G$4&"\' or F=\'"&$G$4&"\') ")&" group by C order by count(B) desc limit 5 label C \'Aktivitas\', count(B) \'Transaksi\'",1),{"Aktivitas","Transaksi"})').setBackground('#e2e8f0').setFontWeight('bold');
  R('K32:N32').merge().setValue('Visual Bar').setBackground('#e2e8f0').setFontWeight('bold');
  for (r = 33; r <= 37; r++) {
    R('K' + r + ':N' + r).merge().setFormula(spark_('J', r, 33, 37, COLOR.sky)).setBackground('#ffffff');
  }

  // RIWAYAT TRANSAKSI
  section('B39:G39', '="10 TRANSAKSI TERBARU"');
  R('B40').setFormula('=IFERROR(QUERY(OPERATION_LOG!$A$1:$H,"select A, C, D, G, E, F where A is not null "&IF($G$4="Semua Lokasi",""," and (E=\'"&$G$4&"\' or F=\'"&$G$4&"\') ")&" order by A desc limit 10 label A \'Waktu\', C \'Aktivitas\', D \'Barcode\', G \'Qty\', E \'Dari\', F \'Ke\'",1),{"Waktu","Aktivitas","Barcode","Qty","Dari","Ke"})').setBackground('#e2e8f0').setFontWeight('bold');
  R('B41:B50').setNumberFormat('dd MMM HH:mm').setBackground('#ffffff');

  // TOP BARANG MASUK
  section('I39:N39', '="TOP 10 BARANG MASUK  ·  "&UPPER($D$4)');
  R('I40:J40').merge().setValue('Nama Produk').setBackground('#e2e8f0').setFontWeight('bold');
  R('K40').setFormula(topQuery_('Masuk')).setBackground('#e2e8f0').setFontWeight('bold');
  R('N40').setValue('Visual Bar').setBackground('#e2e8f0').setFontWeight('bold');
  for (r = 41; r <= 50; r++) {
    R('I' + r + ':J' + r).merge().setFormula(nameLookup_('K' + r, '(tidak ada)')).setBackground('#ffffff');
    R('N' + r).setFormula(spark_('M', r, 41, 50, COLOR.green)).setBackground('#ffffff');
  }

  sh.getRangeList(['D11:F20', 'K11:L20', 'I24:K29', 'I33:J37', 'C41:G50', 'K41:M50']).getRanges().forEach(function (range) {
    range.setBackground('#ffffff').setHorizontalAlignment('center');
  });

  // Grafik tetap dibuat saat semua nilai 0. Dalam kondisi itu batangnya memang belum terlihat.
  var data = ss.getSheetByName(SHEET_DASH_DATA);
  var chart = sh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(data.getRange('A1:C15'))
    .setNumHeaders(1)
    .setPosition(23, 2, 2, 2)
    .setOption('title', 'Jumlah unit per hari')
    .setOption('legend', { position: 'bottom' })
    .setOption('colors', ['#10b981', '#f43f5e'])
    .setOption('backgroundColor', '#ffffff')
    .setOption('vAxis', { minValue: 0, viewWindow: { min: 0 }, format: '0' })
    .setOption('hAxis', { slantedText: false })
    .setOption('width', 596)
    .setOption('height', 316)
    .build();
  sh.insertChart(chart);

  var ruleHabis = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('HABIS').setBackground('#fee2e2').setFontColor('#b91c1c').setRanges([R('M11:M20')]).build();
  var ruleTipis = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('MENIPIS').setBackground('#fef3c7').setFontColor('#b45309').setRanges([R('M11:M20')]).build();
  sh.setConditionalFormatRules([ruleHabis, ruleTipis]);
}
