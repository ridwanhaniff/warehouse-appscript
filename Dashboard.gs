/**
 * INA BASIC STORE - Dashboard V2
 * Reads stock by ID_Variasi and location; legacy operation rows remain visible
 * with product-name fallback where the old log has no variation columns.
 */
var IBS_DASH = {
  DASH: 'DASHBOARD',
  DATA: '_DASH_DATA',
  TZ: 'Asia/Jakarta',
  COLOR: {
    page: '#f1f5f9', dark: '#0f172a', text: '#1e293b', muted: '#64748b',
    line: '#cbd5e1', indigo: '#4f46e5', green: '#059669', red: '#e11d48', amber: '#d97706', sky: '#0284c7'
  }
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Ina Warehouse')
    .addItem('Siapkan / Perbarui Dashboard V2', 'setupDashboard')
    .addItem('Siapkan Sheet Sistem V2', 'setup')
    .addItem('Jalankan Migrasi Stok Lama', 'showMigrationReminderV2_')
    .addToUi();
}

function showMigrationReminderV2_() {
  SpreadsheetApp.getUi().alert(
    'Migrasi stok lama',
    'Isi sheet MIGRASI_ALOKASI dengan pembagian stok lama ke ID_Variasi. Pastikan total per Barcode_Induk + Lokasi_Rak sama persis dengan STOK_LOKASI. Setelah valid, jalankan fungsi migrateLegacyStockV2 dari editor Apps Script. Sheet lama tidak dihapus.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function setupDashboard() {
  setup();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(IBS_DASH.TZ);
  if (!ss.getSheetByName(IBS_DASH.DASH)) ss.insertSheet(IBS_DASH.DASH);
  if (!ss.getSheetByName(IBS_DASH.DATA)) ss.insertSheet(IBS_DASH.DATA);
  buildDashDataV2_(ss);
  buildDashboardV2_(ss);
  var dash = ss.getSheetByName(IBS_DASH.DASH);
  ss.setActiveSheet(dash);
  ss.moveActiveSheet(1);
  SpreadsheetApp.flush();
}

function buildDashDataV2_(ss) {
  var sh = ss.getSheetByName(IBS_DASH.DATA) || ss.insertSheet(IBS_DASH.DATA);
  sh.clear();
  sh.getRange('A1:H1').setValues([['Tanggal', 'Masuk', 'Keluar', '', 'TGL_AWAL', 'TGL_AKHIR', '', 'LOKASI']]);
  var formulas = [];
  for (var r = 2; r <= 15; r++) {
    var dateRef = '$A' + r;
    formulas.push([
      '=TODAY()-' + (15 - r),
      '=SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Masuk",OPERATION_LOG!$A$2:$A,">="&' + dateRef + ',OPERATION_LOG!$A$2:$A,"<"&(' + dateRef + '+1))',
      '=SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Keluar",OPERATION_LOG!$A$2:$A,">="&' + dateRef + ',OPERATION_LOG!$A$2:$A,"<"&(' + dateRef + '+1))'
    ]);
  }
  sh.getRange(2, 1, formulas.length, 3).setFormulas(formulas);
  sh.getRange('A2:A15').setNumberFormat('dd MMM');
  sh.getRange('B2:C15').setNumberFormat('0');
  sh.getRange('E2').setFormula('=IFS(DASHBOARD!$D$4="Hari Ini",TODAY(),DASHBOARD!$D$4="Kemarin",TODAY()-1,DASHBOARD!$D$4="7 Hari Terakhir",TODAY()-7,DASHBOARD!$D$4="30 Hari Terakhir",TODAY()-30,DASHBOARD!$D$4="Bulan Ini",EOMONTH(TODAY(),-1)+1,DASHBOARD!$D$4="Semua Waktu",DATE(2000,1,1),TRUE,TODAY()-30)');
  sh.getRange('F2').setFormula('=IF(DASHBOARD!$D$4="Kemarin",TODAY(),TODAY()+1)');
  sh.getRange('H2').setValue('Semua Lokasi');
  sh.getRange('H3').setFormula('=IFERROR(SORT(UNIQUE(FILTER(STOK_VARIASI!C2:C,STOK_VARIASI!C2:C<>""))),"")');

  sh.getRange('J1:P1').setValues([['Timestamp', 'Aktivitas', 'Item_Key', 'Produk_Tampilan', 'Qty', 'Lokasi_Asal', 'Lokasi_Tujuan']]);
  sh.getRange('J2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",OPERATION_LOG!A2:A))');
  sh.getRange('K2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",OPERATION_LOG!C2:C))');
  sh.getRange('L2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",IF(OPERATION_LOG!I2:I<>"",OPERATION_LOG!I2:I,OPERATION_LOG!D2:D)))');
  sh.getRange('M2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",IF(OPERATION_LOG!M2:M<>"",OPERATION_LOG!M2:M,IFERROR(VLOOKUP(OPERATION_LOG!D2:D,MASTER_PRODUK!A:B,2,FALSE),OPERATION_LOG!D2:D))))');
  sh.getRange('N2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",OPERATION_LOG!G2:G))');
  sh.getRange('O2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",OPERATION_LOG!E2:E))');
  sh.getRange('P2').setFormula('=ARRAYFORMULA(IF(OPERATION_LOG!A2:A="","",OPERATION_LOG!F2:F))');
  sh.hideSheet();
}

function topItemQueryV2_(activity) {
  var locClause = activity === 'Keluar' ? 'O' : 'P';
  return `=IFERROR(QUERY(_DASH_DATA!$J$1:$P,"select M, count(J), sum(N) where K='${activity}' and J >= datetime '"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00' and J < datetime '"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00' "&IF($G$4="Semua Lokasi",""," and ${locClause}='"&$G$4&"' ")&" group by M order by sum(N) desc limit 10 label M 'Produk / Variasi', count(J) 'Transaksi', sum(N) 'Total Qty'",1),{"Produk / Variasi","Transaksi","Total Qty"})`;
}

function sparkFormulaV2_(col, row, first, last, color) {
  return '=IF(' + col + row + '="","",IFERROR(SPARKLINE(' + col + row + ',{"charttype","bar";"max",MAX($' + col + '$' + first + ':$' + col + '$' + last + ');"color1","' + color + '"}),""))';
}

function buildDashboardV2_(ss) {
  var sh = ss.getSheetByName(IBS_DASH.DASH) || ss.insertSheet(IBS_DASH.DASH);
  sh.getCharts().forEach(function (chart) { sh.removeChart(chart); });
  sh.clear(); sh.clearConditionalFormatRules(); sh.getDataRange().clearDataValidations();
  var C = IBS_DASH.COLOR;
  var R = function (a1) { return sh.getRange(a1); };
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 14); for (var c = 2; c <= 14; c++) sh.setColumnWidth(c, 100); sh.setColumnWidth(15, 14);
  sh.setRowHeights(1, 54, 22); [1,5,8,21,30,38,51].forEach(function (r) { sh.setRowHeight(r, 12); });
  sh.setRowHeight(2, 46); sh.setRowHeight(3, 24); sh.setRowHeight(7, 54);
  R('A1:O54').setBackground(C.page).setFontFamily('Roboto').setFontSize(10).setVerticalAlignment('middle');
  R('B2:N2').merge().setValue('  INA BASIC STORE  ·  WAREHOUSE DASHBOARD').setBackground(C.dark).setFontColor('#fff').setFontSize(18).setFontWeight('bold');
  R('B3:N3').merge().setFormula('="  Diperbarui: "&TEXT(NOW(),"dd MMM yyyy HH:mm")').setBackground('#1e293b').setFontColor('#cbd5e1');

  R('B4:C4').merge().setValue('Filter Waktu'); R('D4').setValue('30 Hari Terakhir').setBackground('#fef9c3');
  R('E4:F4').merge().setValue('Filter Lokasi').setHorizontalAlignment('right'); R('G4:H4').merge().setValue('Semua Lokasi').setBackground('#fef9c3');
  R('I4:J4').merge().setValue('Batas Kritis').setHorizontalAlignment('right'); R('K4').setValue(10).setBackground('#fef9c3');
  R('D4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Hari Ini','Kemarin','7 Hari Terakhir','30 Hari Terakhir','Bulan Ini','Semua Waktu'], true).build());
  R('G4').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName(IBS_DASH.DATA).getRange('H2:H100'), true).build());

  function kpi(labelRange, valueRange, label, formula, color) {
    R(labelRange).merge().setValue(label).setFontSize(8).setFontColor(C.muted).setHorizontalAlignment('center');
    R(valueRange).merge().setFormula(formula).setFontSize(22).setFontColor(color).setHorizontalAlignment('center').setFontWeight('bold');
    R(labelRange.split(':')[0] + ':' + valueRange.split(':')[1]).setBorder(true,true,true,true,false,false,C.line,SpreadsheetApp.BorderStyle.SOLID);
  }
  function section(range, title) { R(range).merge().setValue(title).setBackground(C.dark).setFontColor('#fff').setFontWeight('bold'); }

  kpi('B6:C6','B7:C7','PRODUK AKTIF','=COUNTIF(MASTER_PRODUK!$C$2:$C,TRUE)',C.indigo);
  kpi('D6:E6','D7:E7','VARIASI AKTIF','=COUNTIF(VARIASI_PRODUK!$E$2:$E,TRUE)',C.sky);
  kpi('F6:G6','F7:G7','TOTAL STOK','=IF($G$4="Semua Lokasi",SUM(STOK_VARIASI!$D$2:$D),SUMIFS(STOK_VARIASI!$D$2:$D,STOK_VARIASI!$C$2:$C,$G$4))',C.text);
  kpi('I6:J6','I7:J7','MASUK','=IF($G$4="Semua Lokasi",SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Masuk",OPERATION_LOG!$A$2:$A,">="&_DASH_DATA!$E$2,OPERATION_LOG!$A$2:$A,"<"&_DASH_DATA!$F$2),SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Masuk",OPERATION_LOG!$A$2:$A,">="&_DASH_DATA!$E$2,OPERATION_LOG!$A$2:$A,"<"&_DASH_DATA!$F$2,OPERATION_LOG!$F$2:$F,$G$4))',C.green);
  kpi('K6:L6','K7:L7','KELUAR','=IF($G$4="Semua Lokasi",SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Keluar",OPERATION_LOG!$A$2:$A,">="&_DASH_DATA!$E$2,OPERATION_LOG!$A$2:$A,"<"&_DASH_DATA!$F$2),SUMIFS(OPERATION_LOG!$G$2:$G,OPERATION_LOG!$C$2:$C,"Keluar",OPERATION_LOG!$A$2:$A,">="&_DASH_DATA!$E$2,OPERATION_LOG!$A$2:$A,"<"&_DASH_DATA!$F$2,OPERATION_LOG!$E$2:$E,$G$4))',C.red);
  kpi('M6:N6','M7:N7','STOK KRITIS','=IFERROR(ROWS(QUERY(QUERY(STOK_VARIASI!$A$1:$D,"select A, sum(D) where A is not null "&IF($G$4="Semua Lokasi","","and C=\'"&$G$4&"\' ")&"group by A",1),"select Col1 where Col2 <= "&$K$4,1))-1,0)',C.amber);

  // Top 10 keluar / masuk, using snapshot display name and variant label for V2 log rows.
  section('B9:G9','TOP 10 VARIASI KELUAR');
  R('B10:C10').merge().setValue('Produk / Variasi').setBackground('#e2e8f0').setFontWeight('bold');
  R('D10').setFormula(topItemQueryV2_('Keluar')).setBackground('#e2e8f0').setFontWeight('bold');
  R('G10').setValue('Visual').setBackground('#e2e8f0').setFontWeight('bold');
  for (var r = 11; r <= 20; r++) { R('B'+r+':C'+r).merge().setFormula('=IF(D'+r+'="","",D'+r+')').setBackground('#fff'); R('G'+r).setFormula(sparkFormulaV2_('F',r,11,20,C.red)).setBackground('#fff'); }

  // Critical stock grouped by variation ID, then resolve readable display label.
  section('I9:N9','STOK KRITIS PER VARIASI');
  R('I10:J10').merge().setValue('Produk / Variasi').setBackground('#e2e8f0').setFontWeight('bold');
  R('K10').setValue('ID Variasi').setBackground('#e2e8f0').setFontWeight('bold');
  R('L10').setValue('Stok').setBackground('#e2e8f0').setFontWeight('bold');
  R('M10:N10').merge().setValue('Status').setBackground('#e2e8f0').setFontWeight('bold');
  R('K11').setFormula('=IFERROR(QUERY(QUERY(STOK_VARIASI!$A$1:$D,"select A, sum(D) where A is not null "&IF($G$4="Semua Lokasi","","and C=\'"&$G$4&"\' ")&"group by A label A \'\', sum(D) \'\'",1),"select Col1, Col2 where Col2 <= "&$K$4&" order by Col2 asc limit 10",0),{"",""})');
  for (r = 11; r <= 20; r++) {
    R('I'+r+':J'+r).merge().setFormula('=IF(K'+r+'="","",IFERROR(VLOOKUP(VLOOKUP(K'+r+',VARIASI_PRODUK!$A:$C,2,FALSE),MASTER_PRODUK!$A:$B,2,FALSE),"")&" · "&IFERROR(VLOOKUP(K'+r+',VARIASI_PRODUK!$A:$C,3,FALSE),K'+r+'))').setBackground('#fff');
    R('M'+r+':N'+r).merge().setFormula('=IF(L'+r+'="","",IF(L'+r+'<=0,"HABIS","MENIPIS"))').setBackground('#fff').setHorizontalAlignment('center');
  }

  section('B22:G22','TREN BARANG MASUK VS KELUAR · 14 HARI');
  R('B23:G37').setBackground('#fff');
  R('B38:G38').merge().setFormula('=IF(COUNTA(OPERATION_LOG!A2:A)=0,"Belum ada transaksi. Grafik akan mulai bergerak setelah transaksi pertama.","Grafik menampilkan jumlah unit Masuk dan Keluar per hari.")').setFontColor(C.muted).setFontSize(8).setHorizontalAlignment('center');

  section('I22:N22','PERSEBARAN STOK PER LOKASI');
  R('I23').setFormula('=IFERROR(QUERY(STOK_VARIASI!$A$1:$D,"select C, sum(D), count(A) where C is not null "&IF($G$4="Semua Lokasi","","and C=\'"&$G$4&"\' ")&"group by C order by sum(D) desc limit 6 label C \'Lokasi\', sum(D) \'Total Unit\', count(A) \'Variasi\'",1),{"Lokasi","Total Unit","Variasi"})').setBackground('#e2e8f0').setFontWeight('bold');
  R('L23').setValue('Porsi').setBackground('#e2e8f0').setFontWeight('bold'); R('M23:N23').merge().setValue('Visual').setBackground('#e2e8f0').setFontWeight('bold');
  for (r = 24; r <= 29; r++) { R('L'+r).setFormula('=IF(J'+r+'="","",IFERROR(J'+r+'/SUM($J$24:$J$29),0))').setNumberFormat('0%').setBackground('#fff'); R('M'+r+':N'+r).merge().setFormula(sparkFormulaV2_('J',r,24,29,C.indigo)).setBackground('#fff'); }

  section('I31:N31','AKTIVITAS TRANSAKSI');
  R('I32').setFormula('=IFERROR(QUERY(OPERATION_LOG!$A$1:$M,"select C, count(B) where C is not null and A >= datetime \'"&TEXT(_DASH_DATA!$E$2,"yyyy-mm-dd")&" 00:00:00\' and A < datetime \'"&TEXT(_DASH_DATA!$F$2,"yyyy-mm-dd")&" 00:00:00\' "&IF($G$4="Semua Lokasi","","and (E=\'"&$G$4&"\' or F=\'"&$G$4&"\') ")&"group by C order by count(B) desc limit 5 label C \'Aktivitas\', count(B) \'Transaksi\'",1),{"Aktivitas","Transaksi"})').setBackground('#e2e8f0').setFontWeight('bold');
  R('K32:N32').merge().setValue('Visual').setBackground('#e2e8f0').setFontWeight('bold');
  for (r = 33; r <= 37; r++) R('K'+r+':N'+r).merge().setFormula(sparkFormulaV2_('J',r,33,37,C.sky)).setBackground('#fff');

  section('B39:G39','10 TRANSAKSI TERBARU');
  R('B40').setFormula('=IFERROR(QUERY(_DASH_DATA!$J$1:$P,"select J, K, M, N, O, P where J is not null "&IF($G$4="Semua Lokasi","","and (O=\'"&$G$4&"\' or P=\'"&$G$4&"\') ")&"order by J desc limit 10 label J \'Waktu\', K \'Aktivitas\', M \'Produk / Variasi\', N \'Qty\', O \'Dari\', P \'Ke\'",1),{"Waktu","Aktivitas","Produk / Variasi","Qty","Dari","Ke"})').setBackground('#e2e8f0').setFontWeight('bold');
  R('B41:B50').setNumberFormat('dd MMM HH:mm').setBackground('#fff');

  section('I39:N39','TOP 10 VARIASI MASUK');
  R('I40:J40').merge().setValue('Produk / Variasi').setBackground('#e2e8f0').setFontWeight('bold');
  R('K40').setFormula(topItemQueryV2_('Masuk')).setBackground('#e2e8f0').setFontWeight('bold');
  R('N40').setValue('Visual').setBackground('#e2e8f0').setFontWeight('bold');
  for (r = 41; r <= 50; r++) { R('I'+r+':J'+r).merge().setFormula('=IF(K'+r+'="","",K'+r+')').setBackground('#fff'); R('N'+r).setFormula(sparkFormulaV2_('M',r,41,50,C.green)).setBackground('#fff'); }

  // Center/format data blocks.
  sh.getRangeList(['D11:F20','K11:L20','I24:K29','I33:J37','B41:G50','K41:M50']).getRanges().forEach(function (rg) { rg.setBackground('#fff').setHorizontalAlignment('center'); });
  var data = ss.getSheetByName(IBS_DASH.DATA);
  var chart = sh.newChart().setChartType(Charts.ChartType.COLUMN).addRange(data.getRange('A1:C15')).setNumHeaders(1)
    .setPosition(23,2,2,2).setOption('title','Jumlah unit per hari').setOption('legend',{position:'bottom'})
    .setOption('colors',['#10b981','#f43f5e']).setOption('backgroundColor','#ffffff')
    .setOption('vAxis',{minValue:0,viewWindow:{min:0},format:'0'}).setOption('hAxis',{slantedText:false}).setOption('width',596).setOption('height',316).build();
  sh.insertChart(chart);
  var habis = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('HABIS').setBackground('#fee2e2').setFontColor('#b91c1c').setRanges([R('M11:M20')]).build();
  var menipis = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('MENIPIS').setBackground('#fef3c7').setFontColor('#b45309').setRanges([R('M11:M20')]).build();
  sh.setConditionalFormatRules([habis,menipis]);
}
