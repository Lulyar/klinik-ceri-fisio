// ============================================================
// KODE.GS — Backend Google Apps Script
// Klinik Ceri Fisio — Rekam Medis Fisioterapi
// Database: Google Spreadsheets
// ============================================================

// GANTI dengan ID Spreadsheet Anda
var SPREADSHEET_ID = '10ERLKM0QxEs2Zs9IVEqjLd6fTJFdlIkqz3sxycQJ2pA';

// ============================================================
// SERVE WEB APP
// ============================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Klinik Ceri Fisio — Rekam Medis')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// HELPER: Get Sheet
// ============================================================
function getSheet(sheetName) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(sheetName);
}

// ============================================================
// AUTHENTICATION (3 roles: admin, fisioterapis, pasien)
// ============================================================
function login(username, password) {
  try {
    var sheet = getSheet('Users');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === username && data[i][1] === password) {
        return {
          success: true,
          user: {
            username: data[i][0],
            role: data[i][2],
            nama: data[i][3],
            idPasien: data[i][4] || ''
          }
        };
      }
    }
    return { success: false, message: 'Username atau password salah!' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ============================================================
// GENERATE NEXT ID
// ============================================================
function getNextId(sheetName, prefix) {
  var sheet = getSheet(sheetName);
  var data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return prefix + '001';
  }

  var maxNum = 0;
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][0]);
    if (id.indexOf(prefix) === 0) {
      var num = parseInt(id.replace(prefix, ''), 10);
      if (num > maxNum) maxNum = num;
    }
  }

  var nextNum = maxNum + 1;
  var padded = ('000' + nextNum).slice(-3);
  return prefix + padded;
}

// ============================================================
// GET ALL DATA FROM A SHEET
// ============================================================
function getData(sheetName) {
  try {
    var sheet = getSheet(sheetName);
    var data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return { success: true, headers: data[0] || [], rows: [] };
    }

    var headers = data[0];
    var rows = [];

    for (var i = 1; i < data.length; i++) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        obj[headers[j]] = data[i][j];
      }
      rows.push(obj);
    }

    return { success: true, headers: headers, rows: rows };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ============================================================
// ADD DATA TO SHEET
// ============================================================
function addData(sheetName, rowData) {
  try {
    var sheet = getSheet(sheetName);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    var newRow = [];
    for (var i = 0; i < headers.length; i++) {
      newRow.push(rowData[headers[i]] || '');
    }

    sheet.appendRow(newRow);
    return { success: true, message: 'Data berhasil ditambahkan!' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ============================================================
// UPDATE DATA IN SHEET (by ID in column A)
// ============================================================
function updateData(sheetName, id, rowData) {
  try {
    var sheet = getSheet(sheetName);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        for (var j = 0; j < headers.length; j++) {
          if (rowData.hasOwnProperty(headers[j])) {
            sheet.getRange(i + 1, j + 1).setValue(rowData[headers[j]]);
          }
        }
        return { success: true, message: 'Data berhasil diperbarui!' };
      }
    }

    return { success: false, message: 'ID tidak ditemukan!' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ============================================================
// UPDATE STATUS ONLY
// ============================================================
function updateStatus(sheetName, id, newStatus) {
  try {
    var sheet = getSheet(sheetName);
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var statusCol = headers.indexOf('status');

    if (statusCol === -1) {
      return { success: false, message: 'Kolom status tidak ditemukan!' };
    }

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
        return { success: true, message: 'Status berhasil diperbarui!' };
      }
    }

    return { success: false, message: 'ID tidak ditemukan!' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ============================================================
// DELETE DATA FROM SHEET (by ID in column A)
// ============================================================
function deleteData(sheetName, id) {
  try {
    var sheet = getSheet(sheetName);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Data berhasil dihapus!' };
      }
    }

    return { success: false, message: 'ID tidak ditemukan!' };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ============================================================
// GET PATIENT BY ID (for patient portal & cek pasien)
// ============================================================
function getPatientById(patientId) {
  try {
    var patientResult = getData('Pasien');
    var jadwalResult = getData('Jadwal');
    var rekamResult = getData('RekamMedis');

    if (!patientResult.success) return patientResult;

    var patient = null;
    for (var i = 0; i < patientResult.rows.length; i++) {
      if (String(patientResult.rows[i].id) === String(patientId)) {
        patient = patientResult.rows[i];
        break;
      }
    }

    if (!patient) {
      return { success: false, message: 'Pasien tidak ditemukan!' };
    }

    var jadwalList = [];
    if (jadwalResult.success) {
      for (var j = 0; j < jadwalResult.rows.length; j++) {
        if (String(jadwalResult.rows[j].idPasien) === String(patientId)) {
          jadwalList.push(jadwalResult.rows[j]);
        }
      }
    }

    var rekamList = [];
    if (rekamResult.success) {
      for (var k = 0; k < rekamResult.rows.length; k++) {
        if (String(rekamResult.rows[k].idPasien) === String(patientId)) {
          rekamList.push(rekamResult.rows[k]);
        }
      }
    }

    var totalSesi = jadwalList.length;
    var selesai = 0, menunggu = 0, dalamProses = 0;
    for (var s = 0; s < jadwalList.length; s++) {
      var st = String(jadwalList[s].status).toLowerCase();
      if (st === 'selesai') selesai++;
      else if (st === 'menunggu') menunggu++;
      else if (st === 'dalam proses') dalamProses++;
    }

    return {
      success: true,
      patient: patient,
      jadwal: jadwalList,
      rekamMedis: rekamList,
      stats: {
        totalSesi: totalSesi,
        selesai: selesai,
        menunggu: menunggu,
        dalamProses: dalamProses
      }
    };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ============================================================
// GET DASHBOARD DATA
// ============================================================
function getDashboardData(role, terapisName) {
  try {
    var jadwalResult = getData('Jadwal');
    var pasienResult = getData('Pasien');

    if (!jadwalResult.success) return jadwalResult;

    var allJadwal = jadwalResult.rows;
    var totalPasien = pasienResult.success ? pasienResult.rows.length : 0;

    if (role === 'fisioterapis' && terapisName) {
      allJadwal = allJadwal.filter(function(j) {
        return String(j.terapis) === String(terapisName);
      });
    }

    var today = new Date();
    var todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    var sesiHariIni = 0;
    var menunggu = 0;
    var selesai = 0;

    for (var i = 0; i < allJadwal.length; i++) {
      var tgl = allJadwal[i].tanggal;
      var jadwalDate = '';
      if (tgl instanceof Date) {
        jadwalDate = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        jadwalDate = String(tgl).substring(0, 10);
      }

      if (jadwalDate === todayStr) {
        sesiHariIni++;
      }

      var status = String(allJadwal[i].status).toLowerCase();
      if (status === 'menunggu') menunggu++;
      if (status === 'selesai') selesai++;
    }

    return {
      success: true,
      stats: {
        totalPasien: totalPasien,
        sesiHariIni: sesiHariIni,
        menunggu: menunggu,
        selesai: selesai
      },
      jadwal: allJadwal
    };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ============================================================
// GET REPORT DATA
// ============================================================
function getReportData() {
  try {
    var pasienResult = getData('Pasien');
    var jadwalResult = getData('Jadwal');

    var totalPasien = pasienResult.success ? pasienResult.rows.length : 0;
    var jadwalRows = jadwalResult.success ? jadwalResult.rows : [];

    var sesiSelesai = 0;
    var sesiMenunggu = 0;
    var terapiCount = {};

    for (var i = 0; i < jadwalRows.length; i++) {
      var status = String(jadwalRows[i].status).toLowerCase();
      if (status === 'selesai') sesiSelesai++;
      if (status === 'menunggu') sesiMenunggu++;

      var terapi = String(jadwalRows[i].jenisTerapi);
      if (terapi) {
        terapiCount[terapi] = (terapiCount[terapi] || 0) + 1;
      }
    }

    var totalSesi = jadwalRows.length;
    var kehadiran = totalSesi > 0 ? Math.round((sesiSelesai / totalSesi) * 100) : 0;

    var terapiFavorit = '-';
    var maxCount = 0;
    for (var key in terapiCount) {
      if (terapiCount[key] > maxCount) {
        maxCount = terapiCount[key];
        terapiFavorit = key;
      }
    }

    var distribusi = [];
    for (var t in terapiCount) {
      distribusi.push({ nama: t, jumlah: terapiCount[t] });
    }
    distribusi.sort(function(a, b) { return b.jumlah - a.jumlah; });

    return {
      success: true,
      stats: {
        totalPasien: totalPasien,
        sesiSelesai: sesiSelesai,
        kehadiran: kehadiran,
        sesiMenunggu: sesiMenunggu,
        terapiFavorit: terapiFavorit
      },
      distribusi: distribusi
    };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.message };
  }
}

// ============================================================
// SETUP: Create initial sheets + 20 dummy patients
// Jalankan fungsi ini SEKALI untuk membuat struktur sheet
// ============================================================
function setupSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  function initSheet(name, headers) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      sheet.clear();
    } else {
      sheet = ss.insertSheet(name);
    }
    sheet.appendRow(headers);
    return sheet;
  }

  // ==========================================
  // 1. Users Sheet (3 roles: admin, fisioterapis, pasien)
  // Headers: username, password, role, nama, idPasien
  // ==========================================
  var usersSheet = initSheet('Users', ['username', 'password', 'role', 'nama', 'idPasien']);
  usersSheet.appendRow(['admin', 'admin123', 'admin', 'Administrator', '']);
  usersSheet.appendRow(['fisioterapis', 'fisio123', 'fisioterapis', 'Anisah Ainiyah', '']);

  // ==========================================
  // 2. Terapis Sheet
  // ==========================================
  var terapisSheet = initSheet('Terapis', ['id', 'nama', 'spesialisasi', 'deskripsi', 'shiftMulai', 'shiftSelesai']);
  terapisSheet.appendRow(['PHY001', 'Anisah Ainiyah', 'Fisioterapis Utama', 'Fisioterapis berlisensi spesialis muskuloskeletal, neurologi, dan rehabilitasi pasca cedera.', '08:00', '17:00']);

  // ==========================================
  // 3. JenisTerapi Sheet
  // ==========================================
  var jenisSheet = initSheet('JenisTerapi', ['id', 'nama', 'deskripsi', 'durasi']);
  var terapiList = [
    ['TRP001', 'Fisioterapi Manual', 'Terapi manipulasi sendi dan jaringan lunak menggunakan tangan.', '45'],
    ['TRP002', 'Elektroterapi', 'Terapi menggunakan stimulasi listrik (TENS, Ultrasound, Diatermi).', '30'],
    ['TRP003', 'Terapi Latihan', 'Latihan terfokus untuk kekuatan otot, kelenturan, dan fungsi motorik.', '45'],
    ['TRP004', 'Hidroterapi', 'Latihan rehabilitasi fisik di dalam kolam air hangat.', '60'],
    ['TRP005', 'Terapi Okupasi', 'Terapi untuk membantu pasien melakukan aktivitas sehari-hari secara mandiri.', '45'],
    ['TRP006', 'Dry Needling', 'Metode penusukan jarum halus pada titik nyeri ketegangan otot.', '30']
  ];
  terapiList.forEach(function(row) { jenisSheet.appendRow(row); });

  // ==========================================
  // 4. Pasien Sheet (20 Dummy)
  // Headers: id, nama, tanggalLahir, alamat, jenisKelamin, riwayatPenyakit, tglDaftar, status
  // ==========================================
  var pasienSheet = initSheet('Pasien', ['id', 'nama', 'tanggalLahir', 'alamat', 'jenisKelamin', 'riwayatPenyakit', 'tglDaftar', 'status']);

  var firstNames = ['Andi', 'Siti', 'Budi', 'Dewi', 'Rudi', 'Rina', 'Joko', 'Maya', 'Hendra', 'Citra', 'Wawan', 'Lina', 'Eko', 'Santi', 'Dedi', 'Yudi', 'Tari', 'Yoga', 'Fajar', 'Guntur'];
  var lastNames = ['Pratama', 'Nurhaliza', 'Santoso', 'Lestari', 'Hidayat', 'Marlina', 'Susanto', 'Sari', 'Wijaya', 'Maharani', 'Saputra', 'Wulandari', 'Kusuma', 'Fitriani', 'Purnomo', 'Setiawan', 'Indah', 'Prasetyo', 'Nugroho', 'Pamungkas'];
  var genders = ['Laki-laki', 'Perempuan', 'Laki-laki', 'Perempuan', 'Laki-laki', 'Perempuan', 'Laki-laki', 'Perempuan', 'Laki-laki', 'Perempuan', 'Laki-laki', 'Perempuan', 'Laki-laki', 'Perempuan', 'Laki-laki', 'Laki-laki', 'Perempuan', 'Laki-laki', 'Laki-laki', 'Laki-laki'];

  var alamatList = [
    'Jl. Merpati No.12, Comal, Pemalang',
    'Jl. Kenanga No.5, Comal, Pemalang',
    'Jl. Mawar No.8, Comal, Pemalang',
    'Jl. Anggrek No.15, Comal, Pemalang',
    'Jl. Melati No.3, Comal, Pemalang',
    'Jl. Dahlia No.22, Comal, Pemalang',
    'Jl. Cempaka No.7, Comal, Pemalang',
    'Jl. Teratai No.11, Comal, Pemalang',
    'Jl. Bougenville No.9, Comal, Pemalang',
    'Jl. Flamboyan No.14, Comal, Pemalang',
    'Jl. Sakura No.6, Comal, Pemalang',
    'Jl. Kamboja No.18, Comal, Pemalang',
    'Jl. Tulip No.2, Comal, Pemalang',
    'Jl. Lavender No.20, Comal, Pemalang',
    'Jl. Seruni No.16, Comal, Pemalang',
    'Jl. Aster No.4, Comal, Pemalang',
    'Jl. Lili No.13, Comal, Pemalang',
    'Jl. Krisan No.10, Comal, Pemalang',
    'Jl. Violet No.19, Comal, Pemalang',
    'Jl. Gardenia No.1, Comal, Pemalang'
  ];

  var riwayatList = [
    'Riwayat hernia nucleus pulposus (HNP) L4-L5.',
    'Riwayat stroke iskemik 6 bulan lalu.',
    'Osteoarthritis genu bilateral sejak 3 tahun.',
    'Frozen shoulder kanan berulang sejak 2024.',
    'Ruptur ACL lutut kiri, operasi rekonstruksi 2 bulan lalu.',
    'Spondylosis cervical kronis, sering kambuh.',
    'Carpal Tunnel Syndrome bilateral, dominan kiri.',
    'Plantar fasciitis bilateral sejak 1 tahun.',
    'Cedera otot betis kanan saat olahraga lari.',
    'Sprain ankle grade II akibat olahraga futsal.',
    'Tennis elbow kiri akibat aktivitas repetitif.',
    'Hemiparesis kanan pasca stroke.',
    'Delayed motor development (usia 2 tahun).',
    'Ischialgia lumbosakral berulang.',
    'Myofascial pain syndrome upper trapezius.',
    'Rheumatoid arthritis jari tangan.',
    'Gangguan keseimbangan pada geriatri.',
    'Fraktur radius distal pasca pelepasan gips.',
    'Pneumonia ringan, aspek kardiorespirasi.',
    'Strain hamstring bilateral akibat sprint.'
  ];

  var pasienData = [];
  for (var i = 0; i < 20; i++) {
    var idNum = i + 1;
    var id = 'PX' + ('000' + idNum).slice(-3);
    var nama = firstNames[i] + ' ' + lastNames[i];

    var birthYear = 1960 + Math.floor(Math.random() * 45);
    var birthMonth = 1 + Math.floor(Math.random() * 12);
    var birthDay = 1 + Math.floor(Math.random() * 28);
    var tglLahir = birthYear + '-' + ('00' + birthMonth).slice(-2) + '-' + ('00' + birthDay).slice(-2);

    var gender = genders[i];
    var alamat = alamatList[i];
    var riwayat = riwayatList[i];
    var tglDaftar = '2026-06-25';
    var status = 'Aktif';

    pasienSheet.appendRow([id, nama, tglLahir, alamat, gender, riwayat, tglDaftar, status]);
    pasienData.push({ id: id, nama: nama });
  }

  // Add pasien user accounts (first 5 patients get login accounts)
  for (var p = 0; p < 5; p++) {
    var pUsername = 'pasien' + (p + 1);
    var pPassword = 'pasien' + (p + 1) + '123';
    usersSheet.appendRow([pUsername, pPassword, 'pasien', pasienData[p].nama, pasienData[p].id]);
  }

  // ==========================================
  // 5. Jadwal Sheet (10 Dummy)
  // ==========================================
  var jadwalSheet = initSheet('Jadwal', ['id', 'idPasien', 'namaPasien', 'jenisTerapi', 'terapis', 'tanggal', 'jam', 'ruang', 'status']);

  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var yesterdayStr = Utilities.formatDate(new Date(new Date().getTime() - 24 * 60 * 60 * 1000), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  var jadwalList = [
    ['APT001', pasienData[0].id, pasienData[0].nama, 'Fisioterapi Manual', 'Anisah Ainiyah', todayStr, '08:00', 'Ruang 1', 'Menunggu'],
    ['APT002', pasienData[1].id, pasienData[1].nama, 'Terapi Latihan', 'Anisah Ainiyah', todayStr, '08:30', 'Ruang 2', 'Dalam Proses'],
    ['APT003', pasienData[2].id, pasienData[2].nama, 'Elektroterapi', 'Anisah Ainiyah', todayStr, '09:00', 'Ruang 1', 'Selesai'],
    ['APT004', pasienData[3].id, pasienData[3].nama, 'Hidroterapi', 'Anisah Ainiyah', todayStr, '09:30', 'Ruang 3', 'Menunggu'],
    ['APT005', pasienData[4].id, pasienData[4].nama, 'Terapi Okupasi', 'Anisah Ainiyah', todayStr, '10:00', 'Ruang 2', 'Selesai'],
    ['APT006', pasienData[5].id, pasienData[5].nama, 'Dry Needling', 'Anisah Ainiyah', todayStr, '10:30', 'Ruang 1', 'Dalam Proses'],
    ['APT007', pasienData[6].id, pasienData[6].nama, 'Elektroterapi', 'Anisah Ainiyah', todayStr, '11:00', 'Ruang 1', 'Selesai'],
    ['APT008', pasienData[7].id, pasienData[7].nama, 'Fisioterapi Manual', 'Anisah Ainiyah', todayStr, '11:30', 'Ruang 2', 'Menunggu'],
    ['APT009', pasienData[8].id, pasienData[8].nama, 'Terapi Latihan', 'Anisah Ainiyah', todayStr, '13:00', 'Ruang 1', 'Selesai'],
    ['APT010', pasienData[9].id, pasienData[9].nama, 'Hidroterapi', 'Anisah Ainiyah', todayStr, '13:30', 'Ruang 3', 'Selesai'],
    ['APT011', pasienData[10].id, pasienData[10].nama, 'Dry Needling', 'Anisah Ainiyah', yesterdayStr, '08:00', 'Ruang 1', 'Selesai'],
    ['APT012', pasienData[11].id, pasienData[11].nama, 'Terapi Okupasi', 'Anisah Ainiyah', yesterdayStr, '09:00', 'Ruang 2', 'Selesai'],
    ['APT013', pasienData[12].id, pasienData[12].nama, 'Terapi Latihan', 'Anisah Ainiyah', yesterdayStr, '10:00', 'Ruang 1', 'Selesai'],
    ['APT014', pasienData[13].id, pasienData[13].nama, 'Fisioterapi Manual', 'Anisah Ainiyah', yesterdayStr, '11:00', 'Ruang 1', 'Selesai'],
    ['APT015', pasienData[14].id, pasienData[14].nama, 'Dry Needling', 'Anisah Ainiyah', yesterdayStr, '13:00', 'Ruang 2', 'Selesai'],
    ['APT016', pasienData[15].id, pasienData[15].nama, 'Elektroterapi', 'Anisah Ainiyah', yesterdayStr, '14:00', 'Ruang 1', 'Selesai'],
    ['APT017', pasienData[16].id, pasienData[16].nama, 'Terapi Latihan', 'Anisah Ainiyah', yesterdayStr, '15:00', 'Ruang 2', 'Selesai'],
    ['APT018', pasienData[17].id, pasienData[17].nama, 'Terapi Okupasi', 'Anisah Ainiyah', yesterdayStr, '16:00', 'Ruang 1', 'Selesai'],
    ['APT019', pasienData[18].id, pasienData[18].nama, 'Fisioterapi Manual', 'Anisah Ainiyah', yesterdayStr, '16:30', 'Ruang 2', 'Selesai'],
    ['APT020', pasienData[19].id, pasienData[19].nama, 'Terapi Latihan', 'Anisah Ainiyah', yesterdayStr, '17:00', 'Ruang 1', 'Selesai']
  ];
  jadwalList.forEach(function(row) { jadwalSheet.appendRow(row); });

  // ==========================================
  // 6. RekamMedis Sheet (20 Dummy)
  // Headers: id, idPasien, namaPasien, diagnosa, tindakan, catatan, pemeriksa, tglPeriksa, status
  // ==========================================
  var rekamSheet = initSheet('RekamMedis', ['id', 'idPasien', 'namaPasien', 'diagnosa', 'tindakan', 'catatan', 'pemeriksa', 'tglPeriksa', 'status']);
  var rekamList = [
    ['RM001', pasienData[0].id, pasienData[0].nama, 'Low Back Pain ec HNP L4-L5', 'Fisioterapi Manual + Traksi Lumbal', 'Nyeri pinggang berkurang pasca traksi. Edukasi core stability.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM002', pasienData[1].id, pasienData[1].nama, 'Hemiparesis Sinistra pasca Stroke', 'Terapi Latihan fungsional', 'Latihan berjalan dengan quad-cane, keseimbangan mulai stabil.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM003', pasienData[2].id, pasienData[2].nama, 'Osteoarthritis Genu Bilateral', 'Elektroterapi + Ultrasound', 'Ultrasound di area lutut untuk mengurangi kekakuan sendi.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM004', pasienData[3].id, pasienData[3].nama, 'Frozen Shoulder Dextra', 'Hidroterapi di kolam hangat', 'Relaksasi otot bahu, ROM abduksi meningkat 15 derajat.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM005', pasienData[4].id, pasienData[4].nama, 'Pasca Rekonstruksi ACL Sinistra', 'Terapi Okupasi + Penguatan otot', 'Mobilisasi sendi lutut pasca operasi, fokus ekstensi penuh.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM006', pasienData[5].id, pasienData[5].nama, 'Spondylosis Cervical', 'Dry Needling upper trapezius', 'Pelepasan spasme otot leher belakang, pusing berkurang.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM007', pasienData[6].id, pasienData[6].nama, 'Carpal Tunnel Syndrome Bilateral', 'Elektroterapi TENS pergelangan tangan', 'Kebas di jari-jari berkurang setelah TENS selama 20 menit.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM008', pasienData[7].id, pasienData[7].nama, 'Plantar Fasciitis Bilateral', 'Fisioterapi Manual + Deep tissue massage', 'Peregangan plantar fascia telapak kaki, nyeri tumit berkurang.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM009', pasienData[8].id, pasienData[8].nama, 'Strain Gastrocnemius Dextra', 'Terapi Latihan stretching betis', 'Restorasi panjang serat otot betis, jalan tidak pincang.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM010', pasienData[9].id, pasienData[9].nama, 'Sprain Ankle Dextra Grade II', 'Hidroterapi + latihan proprioseptif', 'Latihan keseimbangan di kolam hangat untuk mengurangi bengkak.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM011', pasienData[10].id, pasienData[10].nama, 'Tennis Elbow Sinistra', 'Dry Needling extensor carpi radialis', 'Nyeri tekan siku lateral berkurang drastis.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM012', pasienData[11].id, pasienData[11].nama, 'Hemiparesis Dextra pasca Stroke', 'Terapi Okupasi aktivitas makan mandiri', 'Melatih motorik halus tangan kanan memegang sendok.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM013', pasienData[12].id, pasienData[12].nama, 'Delayed Motor Development', 'Terapi Latihan merangkak & berdiri', 'Stimulasi motorik kasar untuk anak usia 2 tahun.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM014', pasienData[13].id, pasienData[13].nama, 'Ischialgia Lumbosakral Bilateral', 'Fisioterapi Manual + Stretching piriformis', 'Pelepasan jepitan saraf sciatic, kesemutan paha belakang berkurang.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM015', pasienData[14].id, pasienData[14].nama, 'Myofascial Pain Syndrome Upper Trapezius', 'Dry Needling trigger point bahu', 'Pelepasan taut band otot bahu, ROM leher kembali penuh.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM016', pasienData[15].id, pasienData[15].nama, 'Rheumatoid Arthritis Digiti Bilateral', 'Elektroterapi parafin bath', 'Mengurangi kekakuan pagi hari pada persendian jari tangan.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM017', pasienData[16].id, pasienData[16].nama, 'Gangguan Keseimbangan Geriatri', 'Terapi Latihan keseimbangan statis & dinamis', 'Mengurangi risiko jatuh pada pasien lansia.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM018', pasienData[17].id, pasienData[17].nama, 'Post Fraktur Radius Distal Dextra', 'Terapi Okupasi genggaman tangan', 'Meningkatkan kekuatan genggaman dan fleksi pergelangan tangan.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM019', pasienData[18].id, pasienData[18].nama, 'Pneumonia Ringan (Fisioterapi Dada)', 'Fisioterapi Manual chest physiotherapy', 'Postural drainage + tapotement untuk meluruhkan sputum.', 'Anisah Ainiyah', yesterdayStr, 'Aktif'],
    ['RM020', pasienData[19].id, pasienData[19].nama, 'Strain Hamstring Bilateral', 'Terapi Latihan penguatan eksentrik', 'Stretching hamstring dinamis, tidak ada nyeri saat ditarik.', 'Anisah Ainiyah', yesterdayStr, 'Aktif']
  ];
  rekamList.forEach(function(row) { rekamSheet.appendRow(row); });

  Logger.log('=== SETUP SELESAI ===');
  Logger.log('Semua sheet telah dibuat dengan 20 data pasien dummy.');
  Logger.log('Akun Login:');
  Logger.log('  Admin    -> admin / admin123');
  Logger.log('  Fisioterapis -> fisioterapis / fisio123');
  Logger.log('  Pasien 1 -> pasien1 / pasien1123');
  Logger.log('  Pasien 2 -> pasien2 / pasien2123');
  Logger.log('  Pasien 3 -> pasien3 / pasien3123');
  Logger.log('  Pasien 4 -> pasien4 / pasien4123');
  Logger.log('  Pasien 5 -> pasien5 / pasien5123');
}

// ============================================================
// SERVE WEB APP STANDALONE (CORS POST API)
// ============================================================
function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var action = params.action;
    var args = params.args || [];
    var result;

    if (action === 'login') {
      result = login(args[0], args[1]);
    } else if (action === 'getData') {
      result = getData(args[0]);
    } else if (action === 'addData') {
      result = addData(args[0], args[1]);
    } else if (action === 'updateData') {
      result = updateData(args[0], args[1], args[2]);
    } else if (action === 'updateStatus') {
      result = updateStatus(args[0], args[1], args[2]);
    } else if (action === 'deleteData') {
      result = deleteData(args[0], args[1]);
    } else if (action === 'getNextId') {
      result = getNextId(args[0], args[1]);
    } else if (action === 'getPatientById') {
      result = getPatientById(args[0]);
    } else {
      result = { success: false, message: 'Action not found' };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
