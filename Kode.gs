function doGet(e) {
  var page = e.parameter.page || 'login';
  var token = e.parameter.token || '';
  var validPages = ['dashboard', 'input_purbalingga', 'input_banjarnegara', 'laporan'];

  if (page === 'login') {
    var tmpl = HtmlService.createTemplateFromFile('login');
    tmpl.token = '';
    tmpl.session = '{}';
    return tmpl.evaluate()
      .setTitle('Login - Sistem Terminal')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (validPages.indexOf(page) === -1) page = 'dashboard';

  var session = checkSession(token);
  if (!session) {
    var tmplLogin = HtmlService.createTemplateFromFile('login');
    tmplLogin.token = '';
    tmplLogin.session = '{}';
    return tmplLogin.evaluate()
      .setTitle('Login - Sistem Terminal')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Cek akses halaman sesuai role
var role = session.role;
var terminal = session.terminal;

if (role === 'tamu') {
  // Tamu: hanya dashboard
  if (page !== 'dashboard') page = 'dashboard';

} else if (role === 'petugas') {
  // Petugas purbalingga: dashboard, input_purbalingga, laporan
  // Petugas banjarnegara: dashboard, input_banjarnegara, laporan
  var allowed = ['dashboard', 'laporan'];
  if (terminal === 'PURBALINGGA') allowed.push('input_purbalingga');
  else if (terminal === 'BANJARNEGARA') allowed.push('input_banjarnegara');

  if (allowed.indexOf(page) === -1) {
    // Akses halaman yang diblokir → redirect ke dashboard
    page = 'dashboard';
  }
}
// admin: semua halaman boleh diakses
//INI BATASNYA
  try {
    var tmpl = HtmlService.createTemplateFromFile(page);
    tmpl.token = token;
    tmpl.session = JSON.stringify(session);
    return tmpl.evaluate()
      .setTitle('Sistem Terminal Perhubungan')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch(err) {
    var tmplErr = HtmlService.createTemplateFromFile('login');
    tmplErr.token = '';
    tmplErr.session = '{}';
    return tmplErr.evaluate()
      .setTitle('Login - Sistem Terminal')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}
function resetSemuaPoId() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var harian = ss.getSheetByName('data_harian');
  var data = harian.getDataRange().getValues();
  var header = data[0];
  var colPoId = header.indexOf('po_id');
  
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    harian.getRange(i + 1, colPoId + 1).setValue('');
    count++;
  }
  Logger.log('Reset ' + count + ' baris po_id selesai!');
}
function autoFillPoId() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var masterSheet = ss.getSheetByName('master_po');
  var masterData = masterSheet.getDataRange().getValues();
  var headerMaster = masterData[0];
  
  var colPoId      = headerMaster.indexOf('po_id');
  var colNamaPo    = headerMaster.indexOf('nama_po');
  var colLayanan   = headerMaster.indexOf('layanan');
  var colTerminal  = headerMaster.indexOf('terminal');
  
  // Buat map dengan key: "NAMA_PO|LAYANAN|TERMINAL" → po_id
  // Ini biar Sinar Jaya PBG dan Sinar Jaya BNA gak ketuker
  var poMap = {};
  for (var i = 1; i < masterData.length; i++) {
    var id       = masterData[i][colPoId];
    var nama     = masterData[i][colNamaPo];
    var layanan  = masterData[i][colLayanan];
    var terminal = masterData[i][colTerminal];
    if (!nama) continue;
    var key = nama.toString().trim().toUpperCase() + '|' + 
              layanan.toString().trim().toUpperCase() + '|' + 
              terminal.toString().trim().toUpperCase();
    poMap[key] = id;
  }
  
  var harian = ss.getSheetByName('data_harian');
  var harianData = harian.getDataRange().getValues();
  var headerHarian = harianData[0];
  
  var colPoIdH    = headerHarian.indexOf('po_id');
  var colNamaPoH  = headerHarian.indexOf('nama_po');
  var colLayananH = headerHarian.indexOf('layanan');
  var colTerminalH = headerHarian.indexOf('terminal');
  
  var filled = 0;
  var notFound = [];
  
  for (var j = 1; j < harianData.length; j++) {
    var poIdVal    = harianData[j][colPoIdH];
    var namaPoVal  = harianData[j][colNamaPoH];
    var layananVal = harianData[j][colLayananH];
    var terminalVal = harianData[j][colTerminalH];
    
    // Skip kalau po_id udah ada
    if (poIdVal !== '' && poIdVal !== null && poIdVal !== undefined) continue;
    if (!namaPoVal) continue;
    
    var key = namaPoVal.toString().trim().toUpperCase() + '|' + 
              layananVal.toString().trim().toUpperCase() + '|' + 
              terminalVal.toString().trim().toUpperCase();
    
    if (poMap[key]) {
      harian.getRange(j + 1, colPoIdH + 1).setValue(poMap[key]);
      filled++;
    } else {
      notFound.push(namaPoVal + ' (' + layananVal + ' - ' + terminalVal + ')');
    }
  }
  
  var uniqueNotFound = notFound.filter(function(v, i, a) { return a.indexOf(v) === i; });
  
  Logger.log('Berhasil diisi: ' + filled + ' baris');
  if (uniqueNotFound.length > 0) {
    Logger.log('Tidak ditemukan di master_po:');
    uniqueNotFound.forEach(function(n) { Logger.log('  - ' + n); });
  }
  Logger.log('Selesai!');
}

function debugPO() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('master_po');
  var data = sheet.getDataRange().getValues();
  
  Logger.log('Total PO: ' + (data.length - 1));
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    Logger.log('po_id:' + row[0] + ' | nama:' + row[1] + ' | layanan:' + row[3] + ' | terminal:' + row[4]);
  }
}
function bersihkanAnomalyRekap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rekap = ss.getSheetByName('rekap_bulanan');
  const data = rekap.getDataRange().getValues();
  const header = data[0];
  const colBulan = header.indexOf('bulan');

  // Kumpulkan baris anomali (dari bawah supaya index tidak geser saat hapus)
  let hapus = [];
  for (let i = 1; i < data.length; i++) {
    const bulan = data[i][colBulan];
    if (!bulan || bulan.toString().includes('NaN') || bulan.toString().trim() === '') {
      hapus.push(i + 1); // +1 karena index sheet mulai dari 1
    }
  }

  if (hapus.length === 0) {
    Logger.log('Tidak ada anomali ditemukan.');
    return { success: true, deleted: 0 };
  }

  // Hapus dari bawah ke atas
  for (let i = hapus.length - 1; i >= 0; i--) {
    rekap.deleteRow(hapus[i]);
  }

  Logger.log(`Berhasil menghapus ${hapus.length} baris anomali.`);
  return { success: true, deleted: hapus.length };
}
// Baca satu sheet harian, return array of rows
function bacaSheetHarian(raw, poMap) {
  let hasil = [];
  let currentTanggal = '';
  let currentLayanan = '';
  let readingData = false;

  for (let i = 0; i < raw.length; i++) {
    const row = raw[i];
    const col0 = row[0] ? row[0].toString().trim() : '';
    const col1 = row[1] ? row[1].toString().trim() : '';
    const col2 = row[2] ? row[2].toString().trim() : '';
    const rowJoined = row.join('').toUpperCase();

    // Deteksi tanggal — kolom A atau B bertulisan "Tanggal"
    if (col0.toUpperCase().includes('TANGGAL') || col1.toUpperCase().includes('TANGGAL')) {
      // Nilai tanggal bisa di kolom B, C, atau D
      const tglRaw = row[2] || row[3] || row[1] || '';
      // Coba parse, kalau kosong coba kolom lain
      const parsed = parseTanggalIndo(tglRaw);
      if (parsed) currentTanggal = parsed;
      readingData = false;
      continue;
    }

    // Deteksi AKAP
    if (rowJoined.includes('AKAP') && rowJoined.includes('TERMINAL')) {
      currentLayanan = 'AKAP';
      readingData = false;
      continue;
    }

    // Deteksi AKDP
    if (rowJoined.includes('AKDP') && rowJoined.includes('TERMINAL')) {
      currentLayanan = 'AKDP';
      readingData = false;
      continue;
    }

    // Deteksi baris nomor kolom (1,2,3,4,5)
    // Cek row[0]==1 dan row[1]==2 dan row[2]==3
    const r0 = parseInt(row[0]);
    const r1 = parseInt(row[1]);
    const r2 = parseInt(row[2]);
    const r3 = parseInt(row[3]);
    if (r0 === 1 && r1 === 2 && r2 === 3 && r3 === 4) {
      readingData = true;
      continue;
    }

    // Stop kalau ketemu JUMLAH
    if (col0.toUpperCase().includes('JUMLAH') || col1.toUpperCase().includes('JUMLAH')) {
      readingData = false;
      continue;
    }

    if (!readingData) continue;
    if (!currentTanggal || !currentLayanan) continue;

    const namaPo  = row[4] ? row[4].toString().trim() : '';
    const nomorBus= row[2] ? row[2].toString().trim() : '';

    if (!namaPo && !nomorBus) continue;

    // Skip kalau no urut bukan angka (baris header nyasar)
    const noUrut = parseInt(row[0]);
    if (isNaN(noUrut)) continue;

    const key  = namaPo.toUpperCase();
    const poId = poMap[key] || '';

    const id = 'DH' + currentTanggal.replace(/-/g,'') +
               currentLayanan + '_' + hasil.length +
               '_' + Math.floor(Math.random() * 9999);

    hasil.push([
      id,
      currentTanggal,
      formatJam(row[1]),           // jam_datang
      formatJam(row[15]),          // jam_berangkat
      'PURBALINGGA',               // terminal
      currentLayanan,              // layanan
      poId,                        // po_id
      namaPo,                      // nama_po
      row[3] ? row[3].toString().trim() : '',   // koperasi
      row[16] ? row[16].toString().trim() : '', // trayek
      nomorBus,                    // nomor_bus
      row[5]  || 0,                // jml_po
      row[6]  || 0,                // BK
      row[7]  || 0,                // BS
      row[8]  || 0,                // BB
      row[9]  || 0,                // MPU
      row[10] || 0,                // jml_seat
      row[11] || 0,                // jml_pnp
      row[13] || 0,                // pnp_naik
      row[12] || 0,                // pnp_turun
      row[17] || 0,                // laik
      row[18] || 0                 // tidak_laik
    ]);
  }

  return hasil;
}

// Helper: parse tanggal Indonesia ke YYYY-MM-DD
function parseTanggalIndo(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  const bulanMap = {
    'JANUARI':'01','FEBRUARI':'02','MARET':'03','APRIL':'04',
    'MEI':'05','JUNI':'06','JULI':'07','AGUSTUS':'08',
    'SEPTEMBER':'09','OKTOBER':'10','NOVEMBER':'11','DESEMBER':'12'
  };
  const parts = val.toString().trim().toUpperCase().split(' ');
  if (parts.length >= 3) {
    const tgl = parts[0].padStart(2, '0');
    const bln = bulanMap[parts[1]] || '01';
    const thn = parts[2];
    return `${thn}-${bln}-${tgl}`;
  }
  return val.toString();
}

// Helper: format jam ke HH:MM
function formatJam(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return String(val.getHours()).padStart(2, '0') + ':' +
           String(val.getMinutes()).padStart(2, '0');
  }
  return val.toString().trim();
}

// Hapus semua sheet import (angka 1-31) setelah import selesai
function hapusSheetImport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const skipSheets = ['master_po', 'master_trayek', 'master_user',
                      'data_harian', 'rekap_bulanan', 'data_pengunjung'];
  const skipKeywords = ['REKAP', 'MASTER', 'PENGUNJUNG'];

  for (let s = allSheets.length - 1; s >= 0; s--) {
    const sheet = allSheets[s];
    const name  = sheet.getName().trim();
    if (skipSheets.includes(name)) continue;
    const isRekap = skipKeywords.some(k => name.toUpperCase().includes(k));
    if (isRekap) continue;
    if (/^\d+$/.test(name)) {
      ss.deleteSheet(sheet);
    }
  }
  Logger.log('Semua sheet import berhasil dihapus.');
}

function fixHurufBesar() {
  var ss = getSpreadsheet();
  var sheets = ['data_harian', 'rekap_bulanan', 'master_po', 'master_trayek'];
  
  // Kolom yang perlu di-uppercase per sheet
  var kolonUpper = {
    'data_harian': [5, 6, 8, 9, 10, 11],  // terminal, layanan, nama_po, koperasi, trayek, nomor_bus (kolom E,F,H,I,J,K)
    'rekap_bulanan': [4, 5],               // terminal, layanan
    'master_po': [2, 3, 4, 5],            // nama_po, nama_koperasi, layanan, terminal
    'master_trayek': [3, 4]               // nama_po, rute
  };

  for (var sheetName in kolonUpper) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    var data = sheet.getDataRange().getValues();
    var koloms = kolonUpper[sheetName];

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      for (var k = 0; k < koloms.length; k++) {
        var colIdx = koloms[k] - 1; // convert ke 0-based
        var val = data[i][colIdx];
        if (val && typeof val === 'string') {
          var upper = val.trim().toUpperCase();
          if (upper !== val) {
            sheet.getRange(i + 1, koloms[k]).setValue(upper);
          }
        }
      }
    }
    Logger.log('Done: ' + sheetName);
  }
  Logger.log('Semua selesai!');
}

function debugHarian2() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('data_harian');
  var data = sheet.getDataRange().getValues();
  
  Logger.log('Total rows: ' + (data.length - 1));
  
  for (var i = 1; i <= Math.min(3, data.length-1); i++) {
    Logger.log('Row ' + i + ':');
    Logger.log('  terminal: [' + data[i][4] + ']');
    Logger.log('  layanan: [' + data[i][5] + ']');
    Logger.log('  tanggal: ' + data[i][1]);
    Logger.log('  tanggal type: ' + (data[i][1] instanceof Date ? 'Date' : typeof data[i][1]));
  }
  
  // Test filter langsung
  var result = getDataHarian('PURBALINGGA', 'Semua', '2026-04-27', '2026-04-30');
  Logger.log('Test PURBALINGGA kapital: ' + result.length + ' hasil');
  
  var result2 = getDataHarian('Purbalingga', 'Semua', '2026-04-27', '2026-04-30');
  Logger.log('Test Purbalingga normal: ' + result2.length + ' hasil');
}

function testSimple2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('data_harian');
  var data = sheet.getDataRange().getValues();
  
  Logger.log('Total rows: ' + (data.length - 1));
  Logger.log('Row 1 terminal: [' + data[1][4] + ']');
  Logger.log('Row 1 layanan: [' + data[1][5] + ']');
  Logger.log('Row 1 tanggal: ' + data[1][1]);
  Logger.log('Row 1 tanggal type: ' + (data[1][1] instanceof Date ? 'Date' : typeof data[1][1]));
}
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== 'data_harian') return;
  
  var range = e.range;
  var row = range.getRow();
  if (row <= 1) return;
  
  // Cek kalau baris dihapus (kolom A jadi kosong)
  var idVal = sheet.getRange(row, 1).getValue();
  if (idVal === '' || idVal === null) {
    // Regenerate rekap untuk bulan yang terpengaruh
    var tglVal = sheet.getRange(row, 2).getValue();
    var terminal = sheet.getRange(row, 5).getValue();
    var layanan = sheet.getRange(row, 6).getValue();
    if (tglVal && terminal && layanan) {
      updateRekapBulanan(tglVal, terminal, layanan);
    }
  }
}
