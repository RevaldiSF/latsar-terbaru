function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function formatJam(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Jakarta', 'HH:mm');
  }
  return String(val);
}

function getPOByLayanan(layanan, terminal) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('master_po');
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[3] === layanan && row[4] === terminal) {
      result.push({ po_id: row[0], nama_po: row[1], nama_koperasi: row[2] });
    }
  }
  return result;
}

function getTrayekByPO(po_id) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('master_trayek');
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === po_id) {
      result.push({ trayek_id: data[i][0], rute: data[i][3] });
    }
  }
  return result;
}

// Convert tanggal apapun formatnya jadi "yyyy-MM-dd" string
function normalizeTanggal(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Jakarta', 'yyyy-MM-dd');
  }
  var s = String(val).trim();
  // Sudah format yyyy-MM-dd
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // Format dd/MM/yyyy
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
    var p = s.split('/');
    return p[2] + '-' + p[1] + '-' + p[0];
  }
  // Coba parse
  try {
    var d = new Date(s);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, 'Asia/Jakarta', 'yyyy-MM-dd');
    }
  } catch(e) {}
  return s;
}

// Convert "yyyy-MM-dd" jadi angka untuk compare
function dateToNum(str) {
  var s = normalizeTanggal(str);
  if (!s) return 0;
  var p = s.split('-');
  if (p.length !== 3) return 0;
  return parseInt(p[0]) * 10000 + parseInt(p[1]) * 100 + parseInt(p[2]);
}

// Format display "dd/MM/yyyy"
function formatDisplay(val) {
  var s = normalizeTanggal(val);
  if (!s) return '';
  var p = s.split('-');
  if (p.length !== 3) return s;
  return p[2] + '/' + p[1] + '/' + p[0];
}

function saveDataHarian(formData, token) {
  var session = checkSession(token);
  if (!session) return { success: false, message: 'Session expired!' };

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('data_harian');
  var tglStr = normalizeTanggal(formData.tanggal);

  // Hitung jumlah PO unik
  var allData = sheet.getDataRange().getValues();
  var poUnik = {};
  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    if (!row[0]) continue;
    var rowTgl = normalizeTanggal(row[1]);
    if (rowTgl === tglStr && row[4] === formData.terminal && row[5] === formData.layanan) {
      poUnik[row[7]] = true;
    }
  }
  poUnik[formData.nama_po] = true;
  var totalPOUnik = Object.keys(poUnik).length;

  // Update jml_po di semua baris yang sama tanggal+terminal+layanan
  for (var j = 1; j < allData.length; j++) {
    var r = allData[j];
    if (!r[0]) continue;
    var rTgl = normalizeTanggal(r[1]);
    if (rTgl === tglStr && r[4] === formData.terminal && r[5] === formData.layanan) {
      sheet.getRange(j + 1, 12).setValue(totalPOUnik);
    }
  }

  var lastRow = sheet.getLastRow();
  var newId = 'DH' + String(lastRow).padStart(5, '0');
  var keterangan = formData.keterangan || '';

  sheet.appendRow([
    newId, tglStr,
    formData.jam_datang, formData.jam_berangkat,
    formData.terminal, formData.layanan,
    formData.po_id, formData.nama_po, formData.koperasi, formData.trayek,
    formData.nomor_bus, totalPOUnik,
    formData.BK, formData.BS, formData.BB, formData.MPU,
    formData.jml_seat, formData.jml_pnp,
    formData.pnp_naik, formData.pnp_turun,
    formData.laik, formData.tidak_laik,
    keterangan
  ]);

  updateRekapBulanan(tglStr, formData.terminal, formData.layanan);
  return { success: true, message: 'Data berhasil disimpan!' };
}

function getDataHarian(terminal, layanan, tanggalMulai, tanggalAkhir) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('data_harian');
  var data = sheet.getDataRange().getValues();
  var result = [];

 function formatJam(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Jakarta', 'HH:mm');
  }
  return String(val);
}

  // Parse filter ke angka yyyyMMdd
  var mulaiParts = tanggalMulai.split('-');
  var akhirParts = tanggalAkhir.split('-');
  var mulaiNum = parseInt(mulaiParts[0]) * 10000 + parseInt(mulaiParts[1]) * 100 + parseInt(mulaiParts[2]);
  var akhirNum = parseInt(akhirParts[0]) * 10000 + parseInt(akhirParts[1]) * 100 + parseInt(akhirParts[2]);

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    // Tanggal dari sheet — SELALU pakai Utilities.formatDate untuk Date object
    var tglNum = 0;
    var tglDisplay = '';
    var tglVal = row[1];

    if (tglVal instanceof Date) {
      var formatted = Utilities.formatDate(tglVal, 'Asia/Jakarta', 'yyyyMMdd');
      tglNum = parseInt(formatted);
      tglDisplay = Utilities.formatDate(tglVal, 'Asia/Jakarta', 'dd/MM/yyyy');
    } else {
      var s = String(tglVal).trim();
      if (s.indexOf('-') > -1) {
        var p = s.split('-');
        tglNum = parseInt(p[0]) * 10000 + parseInt(p[1]) * 100 + parseInt(p[2]);
        tglDisplay = p[2] + '/' + p[1] + '/' + p[0];
      }
    }

    if (tglNum === 0) continue;

    var matchTerminal = terminal === 'Semua' || String(row[4]).trim() === String(terminal).trim();
    var matchLayanan = layanan === 'Semua' || String(row[5]).trim() === String(layanan).trim();
    var matchTanggal = tglNum >= mulaiNum && tglNum <= akhirNum;

    if (matchTerminal && matchLayanan && matchTanggal) {
      result.push({
  id: row[0],
  tanggal: tglDisplay,
  jam_datang: formatJam(row[2]),
  jam_berangkat: formatJam(row[3]),
  terminal: row[4],
  layanan: row[5],
  nama_po: row[7],
  koperasi: row[8],
  trayek: row[9],
  nomor_bus: row[10],
  jml_po: row[11],
  BK: row[12], BS: row[13], BB: row[14], MPU: row[15],
  jml_seat: row[16], jml_pnp: row[17],
  pnp_naik: row[18], pnp_turun: row[19],
  laik: row[20], tidak_laik: row[21],
  keterangan: row[22] || ''  // tambah ini
});
    }
  }
  return result;
}

function getRekapBulanan(terminal, layanan, bulan, tahun) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('rekap_bulanan');
  var data = sheet.getDataRange().getValues();
  var result = [];

  var bulanInt = parseInt(bulan);
  var tahunInt = parseInt(tahun);

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    // Parse bulan dari kolom bulan (index 1) — format "4-2026"
    var bulanKey = String(row[1]);
    var rowBulan = 0, rowTahun = 0;

    if (bulanKey.indexOf('-') > -1) {
      var parts = bulanKey.split('-');
      rowBulan = parseInt(parts[0]);
      rowTahun = parseInt(parts[1]);
    } else if (row[2]) {
      // Fallback: parse dari kolom tanggal
      var tglNorm = normalizeTanggal(row[2]);
      if (tglNorm) {
        var tp = tglNorm.split('-');
        rowBulan = parseInt(tp[1]);
        rowTahun = parseInt(tp[0]);
      }
    }

    var matchTerminal = terminal === 'Semua' || row[3] === terminal;
    var matchLayanan = layanan === 'Semua' || row[4] === layanan;
    var matchBulan = rowBulan === bulanInt && rowTahun === tahunInt;

    if (matchTerminal && matchLayanan && matchBulan) {
      result.push({
        tanggal: formatDisplay(row[2]),
        terminal: row[3], layanan: row[4],
        jml_po: row[5], bus_msk: row[6], bus_klr: row[7],
        BK: row[8], BS: row[9], BB: row[10], MPU: row[11],
        seat: row[12], penumpang: row[13], load_factor: row[14],
        pnp_naik: row[15], pnp_turun: row[16],
        laik_jalan: row[17], tidak_laik: row[18]
      });
    }
  }
  return result;
}

function savePengunjung(formData, token) {
  var session = checkSession(token);
  if (!session) return { success: false, message: 'Session expired!' };

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('data_pengunjung');
  var lastRow = sheet.getLastRow();
  var newId = 'PG' + String(lastRow).padStart(5, '0');
  var tglStr = normalizeTanggal(formData.tanggal);

  sheet.appendRow([newId, tglStr, formData.terminal, formData.jumlah_pengunjung, formData.keterangan || '']);
  return { success: true, message: 'Data pengunjung berhasil disimpan!' };
}

function getDataPengunjung(terminal, bulan, tahun) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('data_pengunjung');
  var data = sheet.getDataRange().getValues();
  var result = [];

  var bulanInt = parseInt(bulan);
  var tahunInt = parseInt(tahun);

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;

    var tglNorm = normalizeTanggal(row[1]);
    if (!tglNorm) continue;
    var tp = tglNorm.split('-');
    var rowBulan = parseInt(tp[1]);
    var rowTahun = parseInt(tp[0]);

    var matchTerminal = String(terminal).trim().toUpperCase() === 'SEMUA' || String(row[2]).trim().toUpperCase() === String(terminal).trim().toUpperCase();
    var matchBulan = rowBulan === bulanInt && rowTahun === tahunInt;

    if (matchTerminal && matchBulan) {
      result.push({
        id: row[0],
        tanggal: formatDisplay(row[1]),
        tanggal_raw: formatDisplay(row[1]).replace(/\//g, ' ').split(' ').map(function(v,idx) {
          if (idx === 1) {
            var bln = ['','Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
            return bln[parseInt(v)] || v;
          }
          return v;
        }).join(' '),
        terminal: row[2],
        jumlah_pengunjung: row[3],
        keterangan: row[4] || ''
      });
    }
  }

  result.sort(function(a, b) {
    return dateToNum(a.tanggal.split('/').reverse().join('-')) - dateToNum(b.tanggal.split('/').reverse().join('-'));
  });

  return result;
}

function getRekapPengunjungDashboard(terminal, bulan, tahun) {
  var data = getDataPengunjung(terminal, bulan, tahun);
  return data;
}
function testFilterHarian() {
  var result = getDataHarian('PURBALINGGA', 'Semua', '2026-04-27', '2026-04-30');
  Logger.log('Jumlah hasil: ' + result.length);
  if (result.length > 0) {
    Logger.log('Contoh row pertama: ' + JSON.stringify(result[0]));
  }
}

function debugDataHarian() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('data_harian');
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i <= Math.min(5, data.length-1); i++) {
    var row = data[i];
    Logger.log('--- Row ' + i + ' ---');
    Logger.log('id: ' + row[0]);
    Logger.log('tanggal raw: ' + row[1]);
    Logger.log('tanggal type: ' + typeof row[1]);
    Logger.log('tanggal isDate: ' + (row[1] instanceof Date));
    Logger.log('normalized: ' + normalizeTanggal(row[1]));
    Logger.log('dateToNum: ' + dateToNum(row[1]));
    Logger.log('terminal: ' + row[4]);
    Logger.log('layanan: ' + row[5]);
  }
  
  // Test filter dengan tanggal hari ini
  var today = '2026-04-30';
  Logger.log('=== TEST FILTER ===');
  Logger.log('mulaiNum: ' + dateToNum(today));
  Logger.log('akhirNum: ' + dateToNum(today));
}


function generateExcelLengkap(terminal, bulan, tahun) {
  var bulanInt = parseInt(bulan);
  var tahunInt = parseInt(tahun);
  var BULAN_NAMA = ['','Januari','Februari','Maret','April','Mei','Juni',
    'Juli','Agustus','September','Oktober','November','Desember'];
  var bulanNama = BULAN_NAMA[bulanInt];
  var jumlahHari = new Date(tahunInt, bulanInt, 0).getDate();

  var rekapAKAP = getRekapBulanan(terminal, 'AKAP', bulan, tahun);
  var rekapAKDP = getRekapBulanan(terminal, 'AKDP', bulan, tahun);
  var dataPengunjung = getDataPengunjung(terminal, bulan, tahun);

  var ss = getSpreadsheet();
  var harian = ss.getSheetByName('data_harian');
  var harianData = harian.getDataRange().getValues();

  var harianPerTanggal = {};
  for (var i = 1; i < harianData.length; i++) {
    var row = harianData[i];
    if (!row[0]) continue;

    var tglVal = row[1];
    var tglStr = '';
    var tglBulan = 0, tglTahun = 0, tglHari = 0;

    if (tglVal instanceof Date) {
      tglBulan = tglVal.getMonth() + 1;
      tglTahun = tglVal.getFullYear();
      tglHari = parseInt(Utilities.formatDate(tglVal, 'Asia/Jakarta', 'd'));
    } else {
      var s = String(tglVal).trim();
      if (s.indexOf('-') > -1) {
        var p = s.split('-');
        tglTahun = parseInt(p[0]);
        tglBulan = parseInt(p[1]);
        tglHari = parseInt(p[2]);
      }
    }

    if (row[4] !== terminal || tglBulan !== bulanInt || tglTahun !== tahunInt) continue;

    if (!harianPerTanggal[tglHari]) harianPerTanggal[tglHari] = { AKAP: [], AKDP: [] };
    var layanan = String(row[5]);
    if (layanan === 'AKAP' || layanan === 'AKDP') {
      harianPerTanggal[tglHari][layanan].push({
        jam_datang: formatJam(row[2]),
        jam_berangkat: formatJam(row[3]),
        nomor_bus: row[10],
        koperasi: row[8],
        nama_po: row[7],
        jml_po: row[11],
        BK: row[12], BS: row[13], BB: row[14], MPU: row[15],
        jml_seat: row[16], jml_pnp: row[17],
        pnp_turun: row[19], pnp_naik: row[18],
        trayek: row[9],
        laik: row[20], tidak_laik: row[21]
      });
    }
  }

  var namaFile = 'Laporan_' + terminal + '_' + bulanNama + '_' + tahunInt;
  var newSS = SpreadsheetApp.create(namaFile);

  var sheetAKAP = newSS.getActiveSheet();
  sheetAKAP.setName('Rekap Bulanan AKAP');
  buatSheetRekap(sheetAKAP, rekapAKAP, 'AKAP', terminal, bulanNama, tahunInt, false);

  var sheetAKDP = newSS.insertSheet('Rekap Bulanan AKDP');
  buatSheetRekap(sheetAKDP, rekapAKDP, 'AKDP', terminal, bulanNama, tahunInt, true);

  var sheetPengunjung = newSS.insertSheet('Rekap Pengunjung');
  buatSheetPengunjung(sheetPengunjung, dataPengunjung, terminal, bulanNama, tahunInt);

  for (var h = 1; h <= jumlahHari; h++) {
    var sheetHari = newSS.insertSheet(String(h));
    var dataHari = harianPerTanggal[h] || { AKAP: [], AKDP: [] };
    var tglHariStr = h + ' ' + bulanNama + ' ' + tahunInt;
    buatSheetHarian(sheetHari, dataHari, terminal, tglHariStr, h, bulanInt, tahunInt);
  }

  var ssId = newSS.getId();
  SpreadsheetApp.flush();
  var url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?format=xlsx';

  PropertiesService.getScriptProperties().setProperty('temp_ss_id', ssId);
  ScriptApp.newTrigger('hapusSpreadsheetSementara').timeBased().after(120000).create();

  return { success: true, url: url, nama: namaFile };
}

function hapusSpreadsheetSementara() {
  var ssId = PropertiesService.getScriptProperties().getProperty('temp_ss_id');
  if (ssId) {
    try { DriveApp.getFileById(ssId).setTrashed(true); } catch(e) {}
    PropertiesService.getScriptProperties().deleteProperty('temp_ss_id');
  }
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'hapusSpreadsheetSementara') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function buatSheetRekap(sheet, data, layanan, terminal, bulanNama, tahun, isAKDP) {
  var r = 1;
  sheet.getRange(1, 1, 100, 30).setFontFamily('Calibri');
  sheet.getRange(r, 1).setValue('REKAPITULASI PENCATATAN ARUS ' + layanan + ' TERMINAL ' + terminal.toUpperCase());
  sheet.getRange(r, 1, 1, 17).merge().setFontWeight('bold').setHorizontalAlignment('center');
  r++;
  sheet.getRange(r, 1).setValue('BULAN'); sheet.getRange(r, 2).setValue(bulanNama.toUpperCase()); r++;
  sheet.getRange(r, 1).setValue('TERMINAL'); sheet.getRange(r, 2).setValue(terminal.toUpperCase()); r++;
  sheet.getRange(r, 1).setValue('KAB / KOTA'); sheet.getRange(r, 2).setValue(terminal.toUpperCase()); r += 2;

  var totalCols = isAKDP ? 18 : 17;
  var colAfterBus = isAKDP ? 10 : 9;

  sheet.getRange(r, 1).setValue('NO'); sheet.getRange(r, 1, 2, 1).merge();
  sheet.getRange(r, 2).setValue('TANGGAL'); sheet.getRange(r, 2, 2, 1).merge();
  sheet.getRange(r, 3).setValue('JUMLAH PO'); sheet.getRange(r, 3, 2, 1).merge();
  sheet.getRange(r, 4).setValue('KENDARAAN'); sheet.getRange(r, 4, 1, 2).merge();
  sheet.getRange(r, 6).setValue('JENIS BUS'); sheet.getRange(r, 6, 1, isAKDP ? 4 : 3).merge();
  sheet.getRange(r, colAfterBus).setValue('JUMLAH SEAT TERSEDIA'); sheet.getRange(r, colAfterBus, 2, 1).merge();
  sheet.getRange(r, colAfterBus+1).setValue('JUMLAH PNP'); sheet.getRange(r, colAfterBus+1, 2, 1).merge();
  sheet.getRange(r, colAfterBus+2).setValue('LOAD FACTOR (%)'); sheet.getRange(r, colAfterBus+2, 2, 1).merge();
  sheet.getRange(r, colAfterBus+3).setValue('PENUMPANG'); sheet.getRange(r, colAfterBus+3, 1, 2).merge();
  sheet.getRange(r, colAfterBus+5).setValue('KONDISI KENDARAAN'); sheet.getRange(r, colAfterBus+5, 1, 2).merge();
  sheet.getRange(r, colAfterBus+7).setValue('KETERANGAN'); sheet.getRange(r, colAfterBus+7, 2, 1).merge();
  sheet.getRange(r, 1, 1, totalCols).setFontWeight('bold').setBackground('#d9d9d9').setHorizontalAlignment('center').setBorder(true,true,true,true,true,true);
  r++;

  sheet.getRange(r, 4).setValue('MASUK'); sheet.getRange(r, 5).setValue('KELUAR');
  sheet.getRange(r, 6).setValue('BK'); sheet.getRange(r, 7).setValue('BS'); sheet.getRange(r, 8).setValue('BB');
  if (isAKDP) sheet.getRange(r, 9).setValue('MPU');
  sheet.getRange(r, colAfterBus+3).setValue('TURUN'); sheet.getRange(r, colAfterBus+4).setValue('NAIK');
  sheet.getRange(r, colAfterBus+5).setValue('LAIK JALAN'); sheet.getRange(r, colAfterBus+6).setValue('TDK LAIK JALAN');
  sheet.getRange(r, 1, 1, totalCols).setFontWeight('bold').setBackground('#d9d9d9').setHorizontalAlignment('center').setWrap(true).setBorder(true,true,true,true,true,true);
  r++;

  var tot = { jml_po:0, bus_msk:0, bus_klr:0, BK:0, BS:0, BB:0, MPU:0, seat:0, pnp:0, turun:0, naik:0, laik:0, tidak:0 };
  data.forEach(function(d, idx) {
    var lf = Number(d.seat) > 0 ? ((Number(d.penumpang)/Number(d.seat))*100).toFixed(2)+'%' : '0%';
    var rowData = [idx+1, d.tanggal, d.jml_po, d.bus_msk, d.bus_klr, d.BK, d.BS, d.BB];
    if (isAKDP) rowData.push(d.MPU);
    rowData = rowData.concat([d.seat, d.penumpang, lf, d.pnp_turun, d.pnp_naik, d.laik_jalan, d.tidak_laik, '']);
    sheet.getRange(r, 1, 1, rowData.length).setValues([rowData]).setBorder(true,true,true,true,true,true).setHorizontalAlignment('center');
    tot.jml_po = Math.max(tot.jml_po, Number(d.jml_po)||0);
    tot.bus_msk += Number(d.bus_msk)||0; tot.bus_klr += Number(d.bus_klr)||0;
    tot.BK += Number(d.BK)||0; tot.BS += Number(d.BS)||0; tot.BB += Number(d.BB)||0; tot.MPU += Number(d.MPU)||0;
    tot.seat += Number(d.seat)||0; tot.pnp += Number(d.penumpang)||0;
    tot.turun += Number(d.pnp_turun)||0; tot.naik += Number(d.pnp_naik)||0;
    tot.laik += Number(d.laik_jalan)||0; tot.tidak += Number(d.tidak_laik)||0;
    r++;
  });

  var totLF = tot.seat > 0 ? ((tot.pnp/tot.seat)*100).toFixed(2)+'%' : '0%';
  var totRow = ['JUMLAH :', '', tot.jml_po, tot.bus_msk, tot.bus_klr, tot.BK, tot.BS, tot.BB];
  if (isAKDP) totRow.push(tot.MPU);
  totRow = totRow.concat([tot.seat, tot.pnp, totLF, tot.turun, tot.naik, tot.laik, tot.tidak, '']);
  sheet.getRange(r, 1, 1, totRow.length).setValues([totRow]).setFontWeight('bold').setBackground('#f0f0f0').setBorder(true,true,true,true,true,true).setHorizontalAlignment('center');
  r += 3;

  var ttdCol = totalCols - 3;
  sheet.getRange(r, ttdCol).setValue('Mengetahui,'); sheet.getRange(r, ttdCol, 1, 4).merge().setHorizontalAlignment('center'); r++;
  sheet.getRange(r, ttdCol).setValue('Plt. Kepala Seksi Angkutan BPSPP Wilayah V /'); sheet.getRange(r, ttdCol, 1, 4).merge().setHorizontalAlignment('center'); r++;
  sheet.getRange(r, ttdCol).setValue('Kepala Seksi Lalu Lintas BPSPP Wilayah V'); sheet.getRange(r, ttdCol, 1, 4).merge().setHorizontalAlignment('center'); r += 4;
  sheet.getRange(r, ttdCol).setValue('BAMBANG SETIARTO, S.Hut'); sheet.getRange(r, ttdCol, 1, 4).merge().setFontWeight('bold').setHorizontalAlignment('center'); r++;
  sheet.getRange(r, ttdCol).setValue('Penata Tk. I'); sheet.getRange(r, ttdCol, 1, 4).merge().setHorizontalAlignment('center'); r++;
  sheet.getRange(r, ttdCol).setValue('NIP. 19720717 199803 1 006'); sheet.getRange(r, ttdCol, 1, 4).merge().setHorizontalAlignment('center');

  // Auto-resize kolom supaya header tidak terpotong
  sheet.autoResizeColumns(1, totalCols);
  sheet.setColumnWidth(1, 35);  // NO
  sheet.setColumnWidth(2, 80);  // TANGGAL
  sheet.setColumnWidth(3, 70);  // JUMLAH PO
}

function buatSheetPengunjung(sheet, data, terminal, bulanNama, tahun) {
  var r = 1;
  sheet.getRange(1, 1, 100, 10).setFontFamily('Calibri');
  sheet.getRange(r, 1).setValue('REKAPITULASI PENCATATAN PENGUNJUNG TERMINAL ' + terminal.toUpperCase());
  sheet.getRange(r, 1, 1, 4).merge().setFontWeight('bold').setHorizontalAlignment('center'); r++;
  sheet.getRange(r, 1).setValue('BULAN'); sheet.getRange(r, 2).setValue(bulanNama.toUpperCase()); r++;
  sheet.getRange(r, 1).setValue('TERMINAL'); sheet.getRange(r, 2).setValue(terminal.toUpperCase()); r++;
  sheet.getRange(r, 1).setValue('KAB/KOTA'); sheet.getRange(r, 2).setValue(terminal.toUpperCase()); r += 2;

  sheet.getRange(r, 1, 1, 4).setValues([['NO','TANGGAL','JUMLAH PENGUNJUNG','KETERANGAN']]).setFontWeight('bold').setBackground('#d9d9d9').setHorizontalAlignment('center').setBorder(true,true,true,true,true,true); r++;

  var tot = 0;
  data.forEach(function(d, idx) {
    sheet.getRange(r, 1, 1, 4).setValues([[idx+1, d.tanggal_raw||d.tanggal, d.jumlah_pengunjung, d.keterangan||'']]).setBorder(true,true,true,true,true,true).setHorizontalAlignment('center');
    tot += Number(d.jumlah_pengunjung)||0;
    r++;
  });
  sheet.getRange(r, 1, 1, 4).setValues([['JUMLAH','',tot,'']]).setFontWeight('bold').setBackground('#f0f0f0').setBorder(true,true,true,true,true,true).setHorizontalAlignment('center'); r += 3;
  sheet.getRange(r, 1).setValue('Mengetahui,'); sheet.getRange(r, 1, 1, 4).merge().setHorizontalAlignment('center'); r++;
  sheet.getRange(r, 1).setValue('Plt. Kepala Seksi Angkutan BPSPP Wilayah V /'); sheet.getRange(r, 1, 1, 4).merge().setHorizontalAlignment('center'); r++;
  sheet.getRange(r, 1).setValue('Kepala Seksi Lalu Lintas BPSPP Wilayah V'); sheet.getRange(r, 1, 1, 4).merge().setHorizontalAlignment('center'); r += 4;
  sheet.getRange(r, 1).setValue('BAMBANG SETIARTO, S.Hut'); sheet.getRange(r, 1, 1, 4).merge().setFontWeight('bold').setHorizontalAlignment('center'); r++;
  sheet.getRange(r, 1).setValue('Penata Tk. I'); sheet.getRange(r, 1, 1, 4).merge().setHorizontalAlignment('center'); r++;
  sheet.getRange(r, 1).setValue('NIP. 19720717 199803 1 006'); sheet.getRange(r, 1, 1, 4).merge().setHorizontalAlignment('center');
}

function buatSheetHarian(sheet, dataHari, terminal, tglHariStr, hari, bulan, tahun) {
  var r = 1;
  sheet.getRange(1, 1, 200, 20).setFontFamily('Calibri').setFontSize(10);

  // ── AKAP ──────────────────────────────────────────────────────────────
  sheet.getRange(r, 1).setValue('PENCATATAN ARUS AKAP - TERMINAL ' + terminal.toUpperCase());
  sheet.getRange(r, 1, 1, 19).merge().setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center'); r++;
  sheet.getRange(r, 1).setValue('Hari');    sheet.getRange(r, 3).setValue(':'); sheet.getRange(r, 4).setValue(getDayName(hari, bulan, tahun)); r++;
  sheet.getRange(r, 1).setValue('Tanggal'); sheet.getRange(r, 3).setValue(':'); sheet.getRange(r, 4).setValue(tglHariStr); r++;
  sheet.getRange(r, 1).setValue('Shift');   sheet.getRange(r, 3).setValue(':'); sheet.getRange(r, 4).setValue(1); r++;

  var hAKAP = [['NO','JAM DATANG','NOMOR KENDARAAN','PT/KOPERASI','NAMA PO','JML PO','BK','BS','BB','JUML SEAT','JUML PNP','TURUN','NAIK','JAM BERANGKAT','TRAYEK/JURUSAN','LAIK JALAN','TDK LAIK JALAN','KET']];
  sheet.getRange(r, 1, 1, 18).setValues(hAKAP)
    .setFontWeight('bold').setBackground('#d9d9d9')
    .setHorizontalAlignment('center').setWrap(true)
    .setBorder(true,true,true,true,true,true); r++;

  var totAKAP = {po:0,BK:0,BS:0,BB:0,seat:0,pnp:0,turun:0,naik:0,laik:0,tidak:0};
  if (dataHari.AKAP.length === 0) {
    sheet.getRange(r, 1, 1, 18).merge().setValue('Tidak ada data').setHorizontalAlignment('center'); r++;
  } else {
    // Hitung PO muncul pertama kali untuk masing-masing nama_po
    var poSudahAKAP = {};
    dataHari.AKAP.forEach(function(d, idx) {
      var namaPo = String(d.nama_po).trim();
      var tampilPO = '';
      if (!poSudahAKAP[namaPo]) {
        poSudahAKAP[namaPo] = true;
        tampilPO = 1; // Tampilkan 1 hanya di kemunculan pertama PO ini
      }
      var rowData = [idx+1, d.jam_datang, d.nomor_bus, d.koperasi, d.nama_po, tampilPO, d.BK, d.BS, d.BB, d.jml_seat, d.jml_pnp, d.pnp_turun, d.pnp_naik, d.jam_berangkat, d.trayek, d.laik, d.tidak_laik, ''];
      sheet.getRange(r, 1, 1, 18).setValues([rowData])
        .setBorder(true,true,true,true,true,true)
        .setHorizontalAlignment('center');
      totAKAP.BK+=Number(d.BK)||0; totAKAP.BS+=Number(d.BS)||0; totAKAP.BB+=Number(d.BB)||0;
      totAKAP.seat+=Number(d.jml_seat)||0; totAKAP.pnp+=Number(d.jml_pnp)||0;
      totAKAP.turun+=Number(d.pnp_turun)||0; totAKAP.naik+=Number(d.pnp_naik)||0;
      totAKAP.laik+=Number(d.laik)||0; totAKAP.tidak+=Number(d.tidak_laik)||0;
      r++;
    });
    totAKAP.po = Object.keys(poSudahAKAP).length; // jumlah PO unik
  }
  sheet.getRange(r, 1, 1, 18).setValues([['JUMLAH','','','','',totAKAP.po,totAKAP.BK,totAKAP.BS,totAKAP.BB,totAKAP.seat,totAKAP.pnp,totAKAP.turun,totAKAP.naik,'','',totAKAP.laik,totAKAP.tidak,'']])
    .setFontWeight('bold').setBackground('#f0f0f0')
    .setBorder(true,true,true,true,true,true).setHorizontalAlignment('center'); r += 3;

  // ── AKDP ──────────────────────────────────────────────────────────────
  sheet.getRange(r, 1).setValue('PENCATATAN ARUS AKDP - TERMINAL ' + terminal.toUpperCase());
  sheet.getRange(r, 1, 1, 19).merge().setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center'); r++;
  sheet.getRange(r, 1).setValue('Hari');    sheet.getRange(r, 3).setValue(':'); sheet.getRange(r, 4).setValue(getDayName(hari, bulan, tahun)); r++;
  sheet.getRange(r, 1).setValue('Tanggal'); sheet.getRange(r, 3).setValue(':'); sheet.getRange(r, 4).setValue(tglHariStr); r++;
  sheet.getRange(r, 1).setValue('Shift');   sheet.getRange(r, 3).setValue(':'); sheet.getRange(r, 4).setValue(1); r++;

  var hAKDP = [['NO','JAM DATANG','NOMOR KENDARAAN','PT/KOPERASI','NAMA PO','JML PO','BK','BS','BB','MPU','JUML SEAT','JUML PNP','TURUN','NAIK','JAM BERANGKAT','TRAYEK/JURUSAN','LAIK JALAN','TDK LAIK JALAN','KET']];
  sheet.getRange(r, 1, 1, 19).setValues(hAKDP)
    .setFontWeight('bold').setBackground('#d9d9d9')
    .setHorizontalAlignment('center').setWrap(true)
    .setBorder(true,true,true,true,true,true); r++;

  var totAKDP = {po:0,BK:0,BS:0,BB:0,MPU:0,seat:0,pnp:0,turun:0,naik:0,laik:0,tidak:0};
  if (dataHari.AKDP.length === 0) {
    sheet.getRange(r, 1, 1, 19).merge().setValue('Tidak ada data').setHorizontalAlignment('center'); r++;
  } else {
    var poSudahAKDP = {};
    dataHari.AKDP.forEach(function(d, idx) {
      var namaPo = String(d.nama_po).trim();
      var tampilPO = '';
      if (!poSudahAKDP[namaPo]) {
        poSudahAKDP[namaPo] = true;
        tampilPO = 1;
      }
      var rowData = [idx+1, d.jam_datang, d.nomor_bus, d.koperasi, d.nama_po, tampilPO, d.BK, d.BS, d.BB, d.MPU, d.jml_seat, d.jml_pnp, d.pnp_turun, d.pnp_naik, d.jam_berangkat, d.trayek, d.laik, d.tidak_laik, ''];
      sheet.getRange(r, 1, 1, 19).setValues([rowData])
        .setBorder(true,true,true,true,true,true)
        .setHorizontalAlignment('center');
      totAKDP.BK+=Number(d.BK)||0; totAKDP.BS+=Number(d.BS)||0; totAKDP.BB+=Number(d.BB)||0; totAKDP.MPU+=Number(d.MPU)||0;
      totAKDP.seat+=Number(d.jml_seat)||0; totAKDP.pnp+=Number(d.jml_pnp)||0;
      totAKDP.turun+=Number(d.pnp_turun)||0; totAKDP.naik+=Number(d.pnp_naik)||0;
      totAKDP.laik+=Number(d.laik)||0; totAKDP.tidak+=Number(d.tidak_laik)||0;
      r++;
    });
    totAKDP.po = Object.keys(poSudahAKDP).length;
  }
  sheet.getRange(r, 1, 1, 19).setValues([['JUMLAH','','','','',totAKDP.po,totAKDP.BK,totAKDP.BS,totAKDP.BB,totAKDP.MPU,totAKDP.seat,totAKDP.pnp,totAKDP.turun,totAKDP.naik,'','',totAKDP.laik,totAKDP.tidak,'']])
    .setFontWeight('bold').setBackground('#f0f0f0')
    .setBorder(true,true,true,true,true,true).setHorizontalAlignment('center');

  // ── Auto-resize semua kolom biar header tidak terpotong ───────────────
  sheet.autoResizeColumns(1, 19);
  // Set lebar minimum untuk kolom penting
  sheet.setColumnWidth(1, 35);  // NO
  sheet.setColumnWidth(2, 70);  // JAM DATANG
  sheet.setColumnWidth(3, 100); // NOMOR KENDARAAN
  sheet.setColumnWidth(4, 120); // PT/KOPERASI
  sheet.setColumnWidth(5, 120); // NAMA PO
  sheet.setColumnWidth(6, 50);  // JML PO
  sheet.setColumnWidth(15, 70); // JAM BERANGKAT
  sheet.setColumnWidth(16, 160); // TRAYEK/JURUSAN
}

function getDayName(hari, bulan, tahun) {
  var days = ['Minggu','Senin','Selasa','Rabu','Kamis','Jumat','Sabtu'];
  return days[new Date(tahun, bulan-1, hari).getDay()];
}

function getDataByNomorBus(nomorBus, terminal, layanan) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('data_harian');
  var data = sheet.getDataRange().getValues();
  var header = data[0];

  var colNomorBus  = header.indexOf('nomor_bus');
  var colTerminal  = header.indexOf('terminal');
  var colLayanan   = header.indexOf('layanan');
  var colPoId      = header.indexOf('po_id');
  var colNamaPO    = header.indexOf('nama_po');
  var colKoperasi  = header.indexOf('koperasi');
  var colTrayek    = header.indexOf('trayek');
  var colSeat      = header.indexOf('jml_seat');
  var colBK        = header.indexOf('BK');
  var colBS        = header.indexOf('BS');
  var colBB        = header.indexOf('BB');
  var colMPU       = header.indexOf('MPU');

  // Normalisasi input: hapus spasi, uppercase
  var nomorClean   = String(nomorBus).trim().toUpperCase().replace(/\s+/g, '');
  var terminalUp   = terminal ? terminal.toString().trim().toUpperCase() : '';
  var layananUp    = layanan  ? layanan.toString().trim().toUpperCase()  : '';

  // Cari dari bawah supaya dapat data terbaru
  for (var i = data.length - 1; i >= 1; i--) {
    var row = data[i];
    if (!row[0]) continue;

    var rowBus      = String(row[colNomorBus]  || '').trim().toUpperCase().replace(/\s+/g, '');
    var rowTerminal = String(row[colTerminal]  || '').trim().toUpperCase();
    var rowLayanan  = String(row[colLayanan]   || '').trim().toUpperCase();

    if (rowBus !== nomorClean) continue;
    if (terminalUp && rowTerminal !== terminalUp) continue;
    if (layananUp  && rowLayanan  !== layananUp)  continue;

    // Tentukan jenis bus dari kolom BK/BS/BB/MPU
    var jenisBus = '';
    if      (Number(row[colBK])  > 0) jenisBus = 'BK';
    else if (Number(row[colBS])  > 0) jenisBus = 'BS';
    else if (Number(row[colBB])  > 0) jenisBus = 'BB';
    else if (Number(row[colMPU]) > 0) jenisBus = 'MPU';

    return {
      found:     true,
      po_id:     row[colPoId]   || '',
      nama_po:   row[colNamaPO] || '',
      koperasi:  row[colKoperasi] || '',
      trayek:    row[colTrayek] || '',
      jml_seat:  Number(row[colSeat]) || 0,
      jenis_bus: jenisBus
    };
  }

  return { found: false };
}
