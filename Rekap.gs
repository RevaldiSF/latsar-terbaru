function updateRekapBulanan(tanggal, terminal, layanan) {
  var ss = getSpreadsheet();
  var harian = ss.getSheetByName('data_harian');
  var rekap = ss.getSheetByName('rekap_bulanan');

  var tglNorm = normalizeTanggal(tanggal);
  var tp = tglNorm.split('-');
  var bulanKey = parseInt(tp[1]) + '-' + parseInt(tp[0]);
  var tglNum = dateToNum(tglNorm);

  var harianData = harian.getDataRange().getValues();

  var jml_po = 0, bus_msk = 0, bus_klr = 0;
  var BK = 0, BS = 0, BB = 0, MPU = 0;
  var seat = 0, penumpang = 0;
  var pnp_naik = 0, pnp_turun = 0;
  var laik_jalan = 0, tidak_laik = 0;
  var poSet = {};

  for (var i = 1; i < harianData.length; i++) {
    var row = harianData[i];
    if (!row[0]) continue;

    var rowTglNum = dateToNum(row[1]);
    if (rowTglNum === tglNum && row[4] === terminal && row[5] === layanan) {
      poSet[row[7]] = true;
      bus_msk++;

      // Bus keluar hanya dihitung kalau BUKAN bermalam
      var keterangan = String(row[22] || '').toUpperCase().trim();
      if (keterangan !== 'BERMALAM') {
        bus_klr++;
      }

      BK += Number(row[12]) || 0;
      BS += Number(row[13]) || 0;
      BB += Number(row[14]) || 0;
      MPU += Number(row[15]) || 0;
      seat += Number(row[16]) || 0;
      penumpang += Number(row[17]) || 0;
      pnp_naik += Number(row[18]) || 0;
      pnp_turun += Number(row[19]) || 0;
      laik_jalan += Number(row[20]) || 0;
      tidak_laik += Number(row[21]) || 0;
    }
  }

  jml_po = Object.keys(poSet).length;
  var load_factor = seat > 0 ? parseFloat(((penumpang / seat) * 100).toFixed(2)) : 0;

  var rekapData = rekap.getDataRange().getValues();
  var existingRow = -1;

  for (var j = 1; j < rekapData.length; j++) {
    if (!rekapData[j][0]) continue;
    var rekapTglNum = dateToNum(rekapData[j][2]);
    if (rekapTglNum === tglNum && rekapData[j][3] === terminal && rekapData[j][4] === layanan) {
      existingRow = j + 1;
      break;
    }
  }

  var rowId = existingRow > 0
    ? rekapData[existingRow - 1][0]
    : 'RB' + String(rekap.getLastRow()).padStart(5, '0');

  var rowData = [
    rowId, bulanKey, tglNorm,
    terminal, layanan, jml_po,
    bus_msk, bus_klr,
    BK, BS, BB, MPU,
    seat, penumpang, load_factor,
    pnp_naik, pnp_turun,
    laik_jalan, tidak_laik
  ];

  if (existingRow > 0) {
    rekap.getRange(existingRow, 1, 1, rowData.length).setValues([rowData]);
  } else {
    rekap.appendRow(rowData);
  }
}


function fixJmlPO() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('data_harian');
  var data = sheet.getDataRange().getValues();
  var header = data[0];

  var colPoId = header.indexOf('po_id');
  var colTgl = header.indexOf('tanggal');
  var colTerminal = header.indexOf('terminal');
  var colLayanan = header.indexOf('layanan');
  var colNamaPo = header.indexOf('nama_po');
  var colJmlPo = header.indexOf('jml_po');

  // Step 1: Hitung PO unik per kombinasi tanggal+terminal+layanan
  var kombiPO = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var tgl = normalizeTanggal(row[colTgl]);
    var key = tgl + '|' + String(row[colTerminal]).trim() + '|' + String(row[colLayanan]).trim();
    if (!kombiPO[key]) kombiPO[key] = {};
    kombiPO[key][String(row[colNamaPo]).trim()] = true;
  }

  // Step 2: Siapkan semua nilai jml_po sekaligus (batch update)
  var jmlPoValues = [];
  for (var j = 1; j < data.length; j++) {
    var r = data[j];
    if (!r[0]) { jmlPoValues.push([r[colJmlPo]]); continue; }
    var tgl2 = normalizeTanggal(r[colTgl]);
    var key2 = tgl2 + '|' + String(r[colTerminal]).trim() + '|' + String(r[colLayanan]).trim();
    var jumlahPO = kombiPO[key2] ? Object.keys(kombiPO[key2]).length : 0;
    jmlPoValues.push([jumlahPO]);
  }

  // Step 3: Update kolom jml_po sekaligus dalam satu operasi
  if (jmlPoValues.length > 0) {
    sheet.getRange(2, colJmlPo + 1, jmlPoValues.length, 1).setValues(jmlPoValues);
  }

  Logger.log('fixJmlPO selesai! Total rows: ' + jmlPoValues.length);
}

function rekapSkipYangSudahAda(bulan, tahun) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('data_harian');
  var data = sheet.getDataRange().getValues();
  var header = data[0];

  var colTgl      = header.indexOf('tanggal');
  var colTerminal = header.indexOf('terminal');
  var colLayanan  = header.indexOf('layanan');

  var bulanInt = parseInt(bulan);
  var tahunInt = parseInt(tahun);

  // Kumpulkan kombinasi tgl+terminal+layanan yang ADA DI data_harian untuk bulan ini
  var kombinasi = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var tgl = normalizeTanggal(row[colTgl]);
    if (!tgl) continue;
    var tp = tgl.split('-');
    if (parseInt(tp[1]) !== bulanInt || parseInt(tp[0]) !== tahunInt) continue;
    var key = tgl + '|' + row[colTerminal] + '|' + row[colLayanan];
    if (!kombinasi[key]) {
      kombinasi[key] = { tanggal: row[colTgl], terminal: row[colTerminal], layanan: row[colLayanan] };
    }
  }

  // Kumpulkan kombinasi yang SUDAH ADA di rekap_bulanan
  var rekapSheet = ss.getSheetByName('rekap_bulanan');
  var rekapData  = rekapSheet.getDataRange().getValues();
  var sudahAda   = {};
  for (var r = 1; r < rekapData.length; r++) {
    if (!rekapData[r][0]) continue;
    var tglRekap = normalizeTanggal(rekapData[r][2]);
    var keyRekap = tglRekap + '|' + rekapData[r][3] + '|' + rekapData[r][4];
    sudahAda[keyRekap] = true;
  }

  // Proses HANYA yang belum ada
  var count = 0;
  var skip  = 0;
  for (var k in kombinasi) {
    if (sudahAda[k]) {
      skip++;
      continue; // Udah ada, skip!
    }
    var item = kombinasi[k];
    updateRekapBulanan(item.tanggal, item.terminal, item.layanan);
    count++;
  }

  Logger.log('Bulan ' + bulan + '/' + tahun + ' selesai! Diproses: ' + count + ', Diskip: ' + skip);
}

// Wrapper per bulan — tinggal run ini
function rekapSkip_Jan2026() { rekapSkipYangSudahAda(1, 2026); }
function rekapSkip_Feb2026() { rekapSkipYangSudahAda(2, 2026); }
function rekapSkip_Mar2026() { rekapSkipYangSudahAda(3, 2026); }
function rekapSkip_Apr2026() { rekapSkipYangSudahAda(4, 2026); }
function rekapSkip_Mei2026() { rekapSkipYangSudahAda(5, 2026); }

// Jalankan untuk semua bulan yang lo punya
function regenerateSemuaBulan() {
  var bulanTahun = [
    {b:1, t:2026}, {b:2, t:2026}, {b:3, t:2026},
    {b:4, t:2026}, {b:5, t:2026}
    // Tambah sesuai bulan yang ada datanya
  ];

  for (var i = 0; i < bulanTahun.length; i++) {
    regenerateRekapPerBulan(bulanTahun[i].b, bulanTahun[i].t);
    Logger.log('Done: ' + bulanTahun[i].b + '/' + bulanTahun[i].t);
  }

  Logger.log('SEMUA SELESAI!');
}
// Tambah function wrapper per bulan
function rekap_Jan2026() { regenerateRekapPerBulan(1, 2026); }
function rekap_Feb2026() { regenerateRekapPerBulan(2, 2026); }
function rekap_Mar2026() { regenerateRekapPerBulan(3, 2026); }
function rekap_Apr2026() { regenerateRekapPerBulan(4, 2026); }
function rekap_Mei2026() { regenerateRekapPerBulan(5, 2026); }

function fixFormatJam() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('data_harian');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;

    // Fix jam_datang (kolom C = index 2)
    var jamDatang = data[i][2];
    if (jamDatang instanceof Date) {
      var jamStr = Utilities.formatDate(jamDatang, 'Asia/Jakarta', 'HH:mm');
      sheet.getRange(i + 1, 3).setValue(jamStr);
    }

    // Fix jam_berangkat (kolom D = index 3)
    var jamBerangkat = data[i][3];
    if (jamBerangkat instanceof Date) {
      var jamStr2 = Utilities.formatDate(jamBerangkat, 'Asia/Jakarta', 'HH:mm');
      sheet.getRange(i + 1, 4).setValue(jamStr2);
    }
  }
  Logger.log('fixFormatJam selesai!');
}

function sortDataHarian() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('data_harian');
  var data = sheet.getDataRange().getValues();
  var header = data[0];
  var rows = data.slice(1).filter(function(r) { return r[0]; });

  // Sort berdasarkan tanggal lalu jam_datang
  rows.sort(function(a, b) {
    var tglA = normalizeTanggal(a[1]);
    var tglB = normalizeTanggal(b[1]);
    if (tglA !== tglB) return tglA > tglB ? 1 : -1;
    var jamA = String(a[2] || '');
    var jamB = String(b[2] || '');
    return jamA > jamB ? 1 : -1;
  });

  // Re-numbering id
  rows = rows.map(function(r, idx) {
    r[0] = 'DH' + String(idx + 1).padStart(5, '0');
    return r;
  });

  // Tulis ulang
  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
  Logger.log('Sort selesai! Total: ' + rows.length + ' baris');
}
