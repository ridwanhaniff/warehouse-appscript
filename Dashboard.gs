/**
 * INA BASIC STORE WAREHOUSE - Dashboard V3 Final (Bebas Array Error)
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
    .addItem('Susun ulang Dashboard', 'setupDashboard')
    .addToUi();
}

function setupDashboard() {
  setup(); // Panggil dari Code.gs
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TIMEZONE);
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
    var day = '(TODAY()-15+ROW())';
    var base = 'OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"';
    var range = '",OPERATION_LOG!$A$2:$A,">="&' + day + ',OPERATION_LOG!$A$2:$A,"<"&(' + day + '+1))';
    rows.push(['=TEXT(' + day + ',"dd MMM")', '=SUMIFS(' + base + 'Masuk' + range, '=SUMIFS(' + base + 'Keluar' + range]);
  }
  d.getRange(2, 1, 14, 3).setFormulas(rows);

  d.getRange('E2').setFormula(
    '=IFS(DASHBOARD!$D$4="Hari Ini", TODAY(), DASHBOARD!$D$4="Kemarin", TODAY()-1, ' +
    'DASHBOARD!$D$4="7 Hari Terakhir", TODAY()-7, DASHBOARD!$D$4="30 Hari Terakhir", TODAY()-30, ' +
    'DASHBOARD!$D$4="Bulan Ini", EOMONTH(TODAY(),-1)+1, DASHBOARD!$D$4="Semua Waktu", DATE(2000,1,1), TRUE, TODAY()-30)'
  );
  d.getRange('F2').setFormula('=IF(DASHBOARD!$D$4="Kemarin", TODAY(), TODAY()+1)');

  d.getRange('H2').setValue('Semua Lokasi');
  d.getRange('H3').setFormula('=IFERROR(QUERY(STOK_LOKASI!B2:B, "select B where B is not null group by B label B \'\'", 0), "")');
  d.hideSheet();
}

// Rumus ini secara mandiri mencetak 3 kolom (SKU, Transaksi, Total Qty) ke arah kanan.
function topQuery_(activity) {
  var locFilter = (activity === 'Keluar') ? 'E' : 'F'; 
  return '=IFERROR(QUERY(OPERATION_LOG!$A$1:$H,"select D, count(B), sum(G) where C=\'' + activity +
    '\' and A >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and A < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' ' +
    '"&IF($G$4="Semua Lokasi",""," and ' + locFilter + '=\'"&$G$4&"\' ")&" ' +
    'group by D order by sum(G) desc limit 10 label D \'SKU\', count(B) \'Transaksi\', sum(G) \'Total Qty\'",1),{"SKU","Transaksi","Total Qty"})';
}

function spark_(valueCol, r, first, last, color) {
  return '=IF(' + valueCol + r + '="","",IFERROR(SPARKLINE(' + valueCol + r +
    ',{"charttype","bar";"max",MAX($' + valueCol + '$' + first + ':$' + valueCol + '$' + last + ');"color1","' + color + '"}),""))';
}

function nameLookup_(skuCell, notFound) {
  return '=IF(' + skuCell + '="","",IFERROR(VLOOKUP(' + skuCell + ',MASTER_PRODUK!$A$2:$C,3,FALSE),"' + notFound + '"))';
}

function buildDashboard_(ss) {
  var sh = ss.getSheetByName(SHEET_DASH) || ss.insertSheet(SHEET_DASH, 0);
  sh.clear();
  sh.clearConditionalFormatRules();
  sh.getDataRange().clearDataValidations();

  var R = function (a1) { return sh.getRange(a1); };

  // Set Dimensi Kolom & Baris
  sh.setColumnWidth(1, 14); for (var c = 2; c <= 7; c++) sh.setColumnWidth(c, 100);
  sh.setColumnWidth(8, 14); for (c = 9; c <= 14; c++) sh.setColumnWidth(c, 100); sh.setColumnWidth(15, 14);
  sh.setRowHeights(1, 54, 22);
  [1, 5, 8, 21, 30, 38, 51].forEach(function (r) { sh.setRowHeight(r, 12); });
  sh.setRowHeight(2, 46); sh.setRowHeight(3, 24); sh.setRowHeight(7, 54);

  R('A1:O54').setBackground(COLOR.page).setFontFamily('Roboto').setFontSize(10).setVerticalAlignment('middle');
  sh.setHiddenGridlines(true);

  // HEADER & BANNER
  R('B2:N2').merge().setValue('  INA BASIC STORE  ·  WAREHOUSE DASHBOARD').setBackground(COLOR.dark).setFontColor('#ffffff').setFontSize(18).setFontWeight('bold');
  R('B3:N3').merge().setFormula('="  Dihitung ulang: "&TEXT(NOW(),"dd MMM yyyy HH:mm")').setBackground('#1e293b').setFontColor('#cbd5e1');

  // FILTER AREA (Kuning)
  R('B4:C4').merge().setValue('Filter Waktu'); R('D4').setValue('30 Hari Terakhir').setBackground('#fef9c3');
  R('E4:F4').merge().setValue('Filter Lokasi').setHorizontalAlignment('right'); R('G4:H4').merge().setValue('Semua Lokasi').setBackground('#fef9c3');
  R('I4:J4').merge().setValue('Batas Kritis').setHorizontalAlignment('right'); R('K4').setValue(10).setBackground('#fef9c3');
  
  R('D4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Hari Ini', 'Kemarin', '7 Hari Terakhir', '30 Hari Terakhir', 'Bulan Ini', 'Semua Waktu'], true).build());
  R('G4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName(SHEET_DASH_DATA).getRange('H2:H100'), true).build());
  
  // FUNGSI HELPER UI
  function kpi(lblRange, valRange, label, formula, col) {
    R(lblRange).merge().setValue(label).setFontSize(8).setFontColor(COLOR.muted).setHorizontalAlignment('center');
    R(valRange).merge().setFormula(formula).setFontSize(22).setFontColor(col).setHorizontalAlignment('center').setFontWeight('bold');
    R(lblRange.split(':')[0] + ':' + valRange.split(':')[1]).setBorder(true, true, true, true, false, false, COLOR.line, SpreadsheetApp.BorderStyle.SOLID);
  }
  function section(a1, label) {
    R(a1).merge().setFormula(label).setBackground(COLOR.dark).setFontColor('#ffffff').setFontWeight('bold').setBorder(true, true, false, true, null, null, COLOR.dark, SpreadsheetApp.BorderStyle.SOLID);
  }

  // KPI CARDS
  kpi('B6:C6', 'B7:C7', 'PRODUK AKTIF', '=SUMPRODUCT(--(MASTER_PRODUK!$A$2:$A<>""),--(UPPER(MASTER_PRODUK!$D$2:$D&"")="TRUE"))', COLOR.indigo);
  kpi('D6:E6', 'D7:E7', 'TOTAL STOK', '=IF($G$4="Semua Lokasi", SUM(STOK_LOKASI!$C$2:$C), SUMIFS(STOK_LOKASI!$C$2:$C, STOK_LOKASI!$B$2:$B, $G$4))', COLOR.text);
  kpi('F6:G6', 'F7:G7', '="MASUK (" & UPPER($D$4) & ")"', '=IF($G$4="Semua Lokasi", SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Masuk",OPERATION_LOG!$A$2:$A,">="&_DASH_DATA!$E$2,OPERATION_LOG!$A$2:$A,"<"&_DASH_DATA!$F$2), SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Masuk",OPERATION_LOG!$A$2:$A,">="&_DASH_DATA!$E$2,OPERATION_LOG!$A$2:$A,"<"&_DASH_DATA!$F$2, OPERATION_LOG!$F$2:$F, $G$4))', COLOR.green);
  kpi('I6:J6', 'I7:J7', '="KELUAR (" & UPPER($D$4) & ")"', '=IF($G$4="Semua Lokasi", SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Keluar",OPERATION_LOG!$A$2:$A,">="&_DASH_DATA!$E$2,OPERATION_LOG!$A$2:$A,"<"&_DASH_DATA!$F$2), SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Keluar",OPERATION_LOG!$A$2:$A,">="&_DASH_DATA!$E$2,OPERATION_LOG!$A$2:$A,"<"&_DASH_DATA!$F$2, OPERATION_LOG!$E$2:$E, $G$4))', COLOR.red);
  kpi('K6:L6', 'K7:L7', '="TRX (" & UPPER($D$4) & ")"', '=IFERROR(ROWS(QUERY(OPERATION_LOG!$A$2:$H, "select A where A >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and A < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' "&IF($G$4="Semua Lokasi",""," and (E=\'"&$G$4&"\' or F=\'"&$G$4&"\') " )&" ", 0)), 0)', COLOR.sky);
  kpi('M6:N6', 'M7:N7', 'STOK KRITIS', '=IFERROR(ROWS(QUERY(QUERY(STOK_LOKASI!$A$1:$C,"select A, sum(C) where A is not null "&IF($G$4="Semua Lokasi",""," and B=\'"&$G$4&"\' ")&" group by A",1),"select Col1 where Col2 <= "&$K$4,1))-1,0)', COLOR.amber);

  // TABEL 1: BARANG KELUAR
  section('B9:G9', '="TOP 10 BARANG KELUAR  ·  "&UPPER($D$4)');
  R('B10:C10').merge().setValue('Nama Produk').setBackground('#e2e8f0').setFontWeight('bold');
  
  // D10 otomatis mencetak (SKU, Transaksi, Total Qty) secara horizontal! Sel E10 & F10 HARUS DIBIARKAN KOSONG agar tidak #REF!
  R('D10').setFormula(topQuery_('Keluar')).setBackground('#e2e8f0').setFontWeight('bold');
  R('G10').setValue('Visual Bar').setBackground('#e2e8f0').setFontWeight('bold');
  
  for (var r = 11; r <= 20; r++) {
    R('B' + r + ':C' + r).merge().setFormula(nameLookup_('D' + r, '(tidak ada)')).setBackground('#ffffff');
    R('G' + r).setFormula(spark_('F', r, 11, 20, COLOR.red)).setBackground('#ffffff');
  }

  // TABEL 2: STOK MENIPIS
  section('I9:N9', '="STOK KRITIS  ·  BATAS ≤ "&$K$4&" UNIT"');
  R('I10:J10').merge().setValue('Nama Produk').setBackground('#e2e8f0').setFontWeight('bold');
  
  R('K10').setFormula('=IFERROR(QUERY(QUERY(STOK_LOKASI!$A$1:$C,"select A, sum(C) where A is not null "&IF($G$4="Semua Lokasi",""," and B=\'"&$G$4&"\' ")&" group by A label sum(C) \'Stok\'",1),"select Col1, Col2 where Col2 <= "&$K$4&" order by Col2 asc limit 10",1),{"SKU","Stok"})').setBackground('#e2e8f0').setFontWeight('bold');
  R('M10:N10').merge().setValue('Status').setBackground('#e2e8f0').setFontWeight('bold');
  
  for (r = 11; r <= 20; r++) {
    R('I' + r + ':J' + r).merge().setFormula(nameLookup_('K' + r, '-')).setBackground('#ffffff');
    R('M' + r + ':N' + r).merge().setFormula('=IF(L' + r + '="","",IF(L' + r + '<=0,"HABIS","MENIPIS"))').setBackground('#ffffff').setHorizontalAlignment('center');
  }

  // AREA KOSONG UNTUK GRAFIK
  section('B22:G22', '="TREN MASUK vs KELUAR  ·  14 HARI TERAKHIR"');
  R('B23:G37').setBackground('#ffffff');

  // TABEL 3: PORSI TOKO
  section('I22:N22', '="PERSEBARAN STOK PER LOKASI"');
  R('I23').setFormula('=IFERROR(QUERY(STOK_LOKASI!$A$1:$C,"select B, sum(C), count(A) where B is not null and C > 0 "&IF($G$4="Semua Lokasi",""," and B=\'"&$G$4&"\' ")&" group by B order by sum(C) desc limit 6 label B \'Toko/Rak\', sum(C) \'Total Unit\', count(A) \'SKU\'",1),{"Toko/Rak","Total Unit","SKU"})').setBackground('#e2e8f0').setFontWeight('bold');
  R('L23').setValue('Porsi').setBackground('#e2e8f0').setFontWeight('bold');
  R('M23:N23').merge().setValue('Visual Bar').setBackground('#e2e8f0').setFontWeight('bold');
  for (r = 24; r <= 29; r++) {
    R('L' + r).setFormula('=IF(J' + r + '="","",J' + r + '/SUM($J$24:$J$29))').setNumberFormat('0%').setBackground('#ffffff');
    R('M' + r + ':N' + r).merge().setFormula(spark_('J', r, 24, 29, COLOR.indigo)).setBackground('#ffffff');
  }

  // TABEL 4: AKTIVITAS
  section('I31:N31', '="AKTIVITAS  ·  "&UPPER($D$4)');
  R('I32').setFormula('=IFERROR(QUERY(OPERATION_LOG!$A$1:$H,"select C, count(B) where C is not null and A >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and A < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' "&IF($G$4="Semua Lokasi",""," and (E=\'"&$G$4&"\' or F=\'"&$G$4&"\') " )&" group by C order by count(B) desc limit 5 label C \'Aktivitas\', count(B) \'Transaksi\'",1),{"Aktivitas","Transaksi"})').setBackground('#e2e8f0').setFontWeight('bold');
  R('K32:N32').merge().setValue('Visual Bar').setBackground('#e2e8f0').setFontWeight('bold');
  for (r = 33; r <= 37; r++) {
    R('K' + r + ':N' + r).merge().setFormula(spark_('J', r, 33, 37, COLOR.sky)).setBackground('#ffffff');
  }

  // TABEL 5: RIWAYAT TERBARU
  section('B39:G39', '="10 TRANSAKSI TERBARU"');
  R('B40').setFormula('=IFERROR(QUERY(OPERATION_LOG!$A$1:$H,"select A, C, D, G, E, F where A is not null "&IF($G$4="Semua Lokasi",""," and (E=\'"&$G$4&"\' or F=\'"&$G$4&"\') " )&" order by A desc limit 10 label A \'Waktu\', C \'Aktivitas\', D \'SKU\', G \'Qty\', E \'Dari\', F \'Ke\'",1),{"Waktu","Aktivitas","SKU","Qty","Dari","Ke"})').setBackground('#e2e8f0').setFontWeight('bold');
  R('B41:B50').setNumberFormat('dd MMM HH:mm').setBackground('#ffffff');
  
  // TABEL 6: BARANG MASUK
  section('I39:N39', '="TOP 10 BARANG MASUK  ·  "&UPPER($D$4)');
  R('I40:J40').merge().setValue('Nama Produk').setBackground('#e2e8f0').setFontWeight('bold');
  
  // I40 otomatis mencetak (SKU, Transaksi, Total Qty) secara horizontal! Sel L40 & M40 HARUS DIBIARKAN KOSONG agar tidak #REF!
  R('K40').setFormula(topQuery_('Masuk')).setBackground('#e2e8f0').setFontWeight('bold');
  R('N40').setValue('Visual Bar').setBackground('#e2e8f0').setFontWeight('bold');
  
  for (r = 41; r <= 50; r++) {
    R('I' + r + ':J' + r).merge().setFormula(nameLookup_('K' + r, '(tidak ada)')).setBackground('#ffffff');
    R('N' + r).setFormula(spark_('M', r, 41, 50, COLOR.green)).setBackground('#ffffff');
  }
  
  // Rapi-rapi garis border untuk semua background putih
  sh.getRangeList(['D11:F20','K11:L20','I24:K29','I33:J37','C41:G50','K41:M50']).getRanges().forEach(range => {
    range.setBackground('#ffffff').setHorizontalAlignment('center');
  });

  // MASUKAN GRAFIK
  var data = ss.getSheetByName(SHEET_DASH_DATA);
  var chart = sh.newChart().setChartType(Charts.ChartType.COLUMN).addRange(data.getRange('A1:C15'))
    .setNumHeaders(1).setPosition(23, 2, 2, 2).setOption('legend', { position: 'bottom' })
    .setOption('colors', ['#10b981', '#f43f5e']).setOption('backgroundColor', '#ffffff')
    .setOption('width', 596).setOption('height', 316).build();
  sh.insertChart(chart);

  // ATUR CONDITIONAL FORMATTING
  var ruleHabis = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('HABIS').setBackground('#fee2e2').setFontColor('#b91c1c').setRanges([R('M11:M20')]).build();
  var ruleTipis = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('MENIPIS').setBackground('#fef3c7').setFontColor('#b45309').setRanges([R('M11:M20')]).build();
  sh.setConditionalFormatRules([ruleHabis, ruleTipis]);
}