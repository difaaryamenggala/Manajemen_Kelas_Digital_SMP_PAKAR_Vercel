/**
 * HANDLER API UNTUK PROXIED VERCEL REQUESTS
 */
function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    const payload = requestData.payload || [];

    // Allowlist publik minimum untuk endpoint proxy. Jangan izinkan pemanggilan
    // fungsi internal seperti setupSheet() atau generateSequentialID() dari HTTP.
    const allowedActions = [
      'authenticateUser', 'getAllUsers', 'saveUser', 'bulkImportUsers', 'deleteUser',
      'getSiswaByKelas', 'getAllKelasList', 'getAllSiswaList', 'getDashboardSummary',
      'generateQRPresensiToken', 'getQRHistory', 'processDirectQRScanWithSelfie',
      'getPresensiData', 'submitIzinSakit', 'getIzinSakitData', 'updateStatusIzin',
      'getJurnalKBM', 'submitJurnalKBM', 'getNilaiSiswa', 'submitNilaiSiswa'
    ];

    if (allowedActions.indexOf(action) === -1 || typeof this[action] !== 'function') {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'Aksi tidak diizinkan atau tidak ditemukan.'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const result = this[action].apply(this, payload);

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: result
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Server Error: ' + err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
/**
 * ============================================================================
 * MANAJEMEN KELAS DIGITAL - BACKEND LOGIC (FULL UPDATED & FIXED)
 * ============================================================================
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Manajemen Kelas Digital')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * INSIALISASI STRUKTUR TABEL GOOGLE SHEETS
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const schemas = {
    'Users': ['UserID', 'Password', 'Nama', 'Role', 'Kelas', 'Email', 'NoHP'],
    'Presensi': ['PresensiID', 'Tanggal', 'UserID', 'Nama', 'Kelas', 'Status', 'Keterangan', 'FotoSelfie', 'Timestamp'],
    'IzinSakit': ['IzinID', 'TanggalMulai', 'TanggalSelesai', 'UserID', 'Nama', 'TipeIzin', 'Alasan', 'StatusApproval', 'CatatanWali', 'Timestamp'],
    'JurnalKBM': ['JurnalID', 'Tanggal', 'MataPelajaran', 'Topik', 'Deskripsi', 'PRTugas', 'GuruID', 'Timestamp', 'Kelas'],
    'NilaiSiswa': ['NilaiID', 'UserID', 'Nama', 'MataPelajaran', 'JenisEvaluasi', 'NilaiAngka', 'CatatanGuru', 'Timestamp'],
    'QRCodeHistory': ['QR_ID', 'GuruID', 'Kelas', 'Mapel', 'CreatedTimestamp', 'ExpireTimestamp', 'Status']
  };

  Object.keys(schemas).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    const expectedHeaders = schemas[sheetName];

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(expectedHeaders);
      sheet.getRange(1, 1, 1, expectedHeaders.length).setFontWeight('bold').setBackground('#f3f4f6');
    } else {
      const lastColumn = sheet.getLastColumn();
      if (lastColumn === 0) {
        sheet.appendRow(expectedHeaders);
        sheet.getRange(1, 1, 1, expectedHeaders.length).setFontWeight('bold').setBackground('#f3f4f6');
      } else {
        const existingHeaders = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
        expectedHeaders.forEach(h => {
          if (!existingHeaders.includes(h)) {
            const nextCol = sheet.getLastColumn() + 1;
            sheet.getRange(1, nextCol).setValue(h).setFontWeight('bold').setBackground('#f3f4f6');
          }
        });
      }
    }
  });

  SpreadsheetApp.flush();
}

/**
 * 1. AUTENTIKASI LOGIN
 */
function authenticateUser(userID, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Database Users kosong.' };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getDisplayValues();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (row[0] === userID && row[1] === password) {
      return {
        success: true,
        user: {
          userID: row[0],
          nama: row[2],
          role: row[3],
          kelas: row[4],
          email: row[5],
          noHp: row[6] || ''
        }
      };
    }
  }
  return { success: false, message: 'UserID atau Password salah.' };
}

/**
 * 2. MANAJEMEN USERS, BULK UPLOAD, & HAPUS DATA LAMA
 */
function getAllUsers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getDisplayValues();
}

function saveUser(userData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  
  if (sheet.getLastRow() > 1) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getDisplayValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === userData.userID) {
        sheet.getRange(i + 2, 2, 1, 6).setValues([[userData.password, userData.nama, userData.role, userData.kelas, userData.email, userData.noHp]]);
        SpreadsheetApp.flush();
        return { success: true, message: 'Data User Berhasil Diperbarui!' };
      }
    }
  }

  sheet.appendRow([userData.userID, userData.password, userData.nama, userData.role, userData.kelas, userData.email, userData.noHp || '']);
  SpreadsheetApp.flush();
  return { success: true, message: 'User Baru Berhasil Ditambahkan!' };
}

function bulkImportUsers(usersList, clearExistingData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Users');
    
    if (!sheet) {
      sheet = ss.insertSheet('Users');
      sheet.appendRow(['UserID', 'Password', 'Nama', 'Role', 'Kelas', 'Email', 'NoHP']);
    }

    if (clearExistingData) {
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
      }
    }

    if (!usersList || usersList.length === 0) {
      return { success: false, message: 'Data yang diimport kosong.' };
    }

    const rowsToAppend = usersList.map(u => [
      String(u.userID || u[0]),
      String(u.password || u[1]),
      String(u.nama || u[2]),
      String(u.role || u[3]),
      String(u.kelas || u[4]),
      String(u.email || u[5] || ''),
      String(u.noHp || u[6] || '')
    ]);

    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 7).setValues(rowsToAppend);
    SpreadsheetApp.flush();

    return { success: true, message: 'Berhasil mengimpor ' + rowsToAppend.length + ' data user!' };
  } catch (err) {
    return { success: false, message: 'Gagal Import: ' + err.message };
  }
}

function deleteUser(userID) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Data tidak ditemukan.' };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === userID) {
      sheet.deleteRow(i + 2);
      SpreadsheetApp.flush();
      return { success: true, message: 'User Berhasil Dihapus!' };
    }
  }
  return { success: false, message: 'UserID tidak ditemukan.' };
}

/**
 * 3. FILTER & PENGAMBILAN DATA KELAS & SISWA FLEKSIBEL
 */
function getSiswaByKelas(kelas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getDisplayValues();
  
  return data
    .filter(row => row[3] === 'Siswa' && (kelas === 'Semua' || row[4] === kelas))
    .map(row => ({
      userID: row[0],
      nama: row[2],
      kelas: row[4],
      email: row[5],
      noHp: row[6] || ''
    }));
}

function getAllKelasList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getDisplayValues();
  const kelasSet = new Set();

  data.forEach(row => {
    if (row[3] === 'Siswa' && row[4]) {
      kelasSet.add(row[4]);
    }
  });

  return Array.from(kelasSet);
}

function getAllSiswaList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getDisplayValues();
  return data.filter(row => row[3] === 'Siswa').map(row => {
    return { userID: row[0], nama: row[2], kelas: row[4], email: row[5] };
  });
}

/**
 * 4. RINGKASAN DASHBOARD
 */
function getDashboardSummary(user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  let totalHadir = 0, totalIzin = 0, totalSakit = 0, totalAlpa = 0;
  let pendingIzinCount = 0;
  let totalSiswa = 0, totalGuru = 0, totalWali = 0;

  const uSheet = ss.getSheetByName('Users');
  if (uSheet && uSheet.getLastRow() > 1) {
    const uData = uSheet.getRange(2, 1, uSheet.getLastRow() - 1, 7).getDisplayValues();
    uData.forEach(r => {
      if (r[3] === 'Siswa') totalSiswa++;
      if (r[3] === 'Guru') totalGuru++;
      if (r[3] === 'Wali Kelas') totalWali++;
    });
  }

  const pSheet = ss.getSheetByName('Presensi');
  if (pSheet && pSheet.getLastRow() > 1) {
    const pData = pSheet.getRange(2, 1, pSheet.getLastRow() - 1, 9).getDisplayValues();
    pData.forEach(row => {
      const matchKelas = (user.role === 'Admin' || user.role === 'Guru') ? true : (row[4] === user.kelas);
      const matchSiswa = (user.role === 'Siswa') ? (row[2] === user.userID) : true;

      if (row[1] === today && matchKelas && matchSiswa) {
        if (row[5] === 'Hadir') totalHadir++;
        if (row[5] === 'Izin') totalIzin++;
        if (row[5] === 'Sakit') totalSakit++;
        if (row[5] === 'Alpa') totalAlpa++;
      }
    });
  }

  const iSheet = ss.getSheetByName('IzinSakit');
  if (iSheet && iSheet.getLastRow() > 1) {
    const iData = iSheet.getRange(2, 1, iSheet.getLastRow() - 1, 10).getDisplayValues();
    iData.forEach(row => {
      if (row[7] === 'Pending') {
        if (user.role === 'Admin' || user.role === 'Wali Kelas' || user.role === 'Guru') pendingIzinCount++;
        else if (user.role === 'Siswa' && row[3] === user.userID) pendingIzinCount++;
      }
    });
  }

  return {
    today: today,
    presensi: { Hadir: totalHadir, Izin: totalIzin, Sakit: totalSakit, Alpa: totalAlpa },
    pendingIzin: pendingIzinCount,
    statsAdmin: { totalSiswa: totalSiswa, totalGuru: totalGuru, totalWali: totalWali }
  };
}

/**
 * 5. QR PRESENSI (GENERATE, KADALUARSA, RIWAYAT, & VALIDASI SCAN)
 */
function generateQRPresensiToken(guruUser, kelas, mapel, expireMinutes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let qSheet = ss.getSheetByName('QRCodeHistory');
  if (!qSheet) {
    qSheet = ss.insertSheet('QRCodeHistory');
    qSheet.appendRow(['QR_ID', 'GuruID', 'Kelas', 'Mapel', 'CreatedTimestamp', 'ExpireTimestamp', 'Status']);
  }

  const duration = parseInt(expireMinutes) || 15;
  const now = new Date();
  const expireTime = new Date(now.getTime() + duration * 60000);

  const qrID = 'QR-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const nowStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const expireStr = Utilities.formatDate(expireTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  qSheet.appendRow([qrID, guruUser.userID, kelas, mapel, nowStr, expireStr, 'Aktif']);
  SpreadsheetApp.flush();

  const qrPayload = JSON.stringify({
    qrID: qrID,
    kelas: kelas,
    mapel: mapel || 'Mata Pelajaran',
    created: now.getTime(),
    expire: expireTime.getTime()
  });

  const qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=' + encodeURIComponent(qrPayload);

  return { success: true, qrUrl: qrUrl, qrID: qrID, expireStr: expireStr, mapel: mapel };
}

function getQRHistory(guruUser) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSheet = ss.getSheetByName('QRCodeHistory');
  if (!qSheet || qSheet.getLastRow() <= 1) return [];

  const data = qSheet.getRange(2, 1, qSheet.getLastRow() - 1, 7).getDisplayValues();
  let filtered = data;
  if (guruUser.role === 'Guru') {
    filtered = data.filter(row => row[1] === guruUser.userID);
  }
  return filtered.reverse();
}

function processDirectQRScanWithSelfie(user, qrDataRaw, base64Image) {
  try {
    const data = JSON.parse(qrDataRaw);
    const nowTs = new Date().getTime();

    if (data.expire && nowTs > data.expire) {
      return { success: false, message: 'Gagal Presensi: QR Code telah KADALUARSA!' };
    }

    if (user.role === 'Siswa' && user.kelas !== data.kelas && data.kelas !== 'Semua') {
      return { success: false, message: 'Gagal Presensi: QR Code khusus untuk kelas ' + data.kelas };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Presensi');
    if (!sheet) {
      setupSheet();
      sheet = ss.getSheetByName('Presensi');
    }

    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    if (sheet.getLastRow() > 1) {
      const existingData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getDisplayValues();
      for (let i = 0; i < existingData.length; i++) {
        if (existingData[i][1] === today && existingData[i][2] === user.userID) {
          return { success: false, message: 'Anda sudah presensi hari ini!' };
        }
      }
    }

    let photoUrl = '-';
    if (base64Image) {
      try {
        const folderName = 'Presensi_Selfie_ClassDigital';
        let folder;
        const folders = DriveApp.getFoldersByName(folderName);
        folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

        const contentType = base64Image.substring(5, base64Image.indexOf(';'));
        const bytes = Utilities.base64Decode(base64Image.substr(base64Image.indexOf('base64,') + 7));
        const blob = Utilities.newBlob(bytes, contentType, user.userID + '_' + today + '.jpg');
        
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoUrl = file.getUrl();
      } catch (e) {
        photoUrl = 'Gagal Unggah Foto: ' + e.message;
      }
    }

    const presensiID = generateSequentialID('Presensi', 'PRS');
    sheet.appendRow([
      presensiID,
      today,
      user.userID,
      user.nama,
      user.kelas,
      'Hadir',
      'Scan QR (' + (data.mapel || '-') + ') [' + data.qrID + ']',
      photoUrl,
      timestamp
    ]);

    SpreadsheetApp.flush();
    return { success: true, message: 'Presensi QR Berhasil Divalidasi & Disimpan!' };

  } catch (err) {
    return { success: false, message: 'QR Code tidak valid: ' + err.message };
  }
}

/**
 * 6. REKAP & DATA PRESENSI
 */
function getPresensiData(user, dateFilter, page, limit, classFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Presensi');
  if (!sheet || sheet.getLastRow() <= 1) return { list: [], total: 0 };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
  
  let filtered = data.filter(row => {
    const matchUser = (user.role === 'Siswa') ? (row[2] === user.userID) : true;
    
    let matchKelas = true;
    if (user.role === 'Wali Kelas') {
      matchKelas = (row[4] === user.kelas);
    } else if (classFilter && classFilter.trim() !== '' && classFilter !== 'Semua') {
      matchKelas = (row[4] === classFilter);
    }

    const matchDate = (dateFilter && dateFilter.trim() !== '') ? (row[1] === dateFilter) : true;
    
    return matchUser && matchKelas && matchDate;
  });

  filtered.reverse();

  const p = parseInt(page) || 1;
  const l = parseInt(limit) || 10;
  const startIndex = (p - 1) * l;
  
  return { 
    list: filtered.slice(startIndex, startIndex + l), 
    total: filtered.length 
  };
}

/**
 * 7. PERMOHONAN IZIN & NOTIFIKASI WHATSAPP
 */
function submitIzinSakit(user, formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('IzinSakit');
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const izinID = generateSequentialID('IzinSakit', 'IZN');
  
  sheet.appendRow([
    izinID,
    String(formData.tanggalMulai),
    String(formData.tanggalSelesai),
    user.userID,
    user.nama,
    String(formData.tipeIzin),
    String(formData.alasan),
    'Pending',
    '-',
    timestamp
  ]);
  SpreadsheetApp.flush();

  let waWaliURL = '';
  const uSheet = ss.getSheetByName('Users');
  if (uSheet && uSheet.getLastRow() > 1) {
    const uData = uSheet.getRange(2, 1, uSheet.getLastRow() - 1, 7).getDisplayValues();
    const waliObj = uData.find(r => r[3] === 'Wali Kelas' && r[4] === user.kelas);
    
    if (waliObj && waliObj[6]) {
      let noHp = waliObj[6].trim().replace(/[^0-9]/g, '');
      if (noHp.startsWith('0')) noHp = '62' + noHp.slice(1);

      const pesanWA = `*PEMBERITAHUAN IZIN/SAKIT SISWA*\n\n` +
                      `Yth. Wali Kelas ${user.kelas},\n` +
                      `Siswa berikut mengajukan permohonan izin:\n\n` +
                      `• *Nama*: ${user.nama} (${user.userID})\n` +
                      `• *Jenis*: ${formData.tipeIzin}\n` +
                      `• *Tanggal*: ${formData.tanggalMulai} s/d ${formData.tanggalSelesai}\n` +
                      `• *Alasan*: ${formData.alasan}\n\n` +
                      `Mohon melakukan verifikasi melalui portal Kelas Digital. Terima Kasih.`;

      waWaliURL = `https://api.whatsapp.com/send?phone=${noHp}&text=${encodeURIComponent(pesanWA)}`;
    }
  }

  return { 
    success: true, 
    message: 'Pengajuan izin berhasil dikirim.', 
    waURL: waWaliURL 
  };
}

function getIzinSakitData(user, page, limit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('IzinSakit');
  if (!sheet || sheet.getLastRow() <= 1) return { list: [], total: 0 };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getDisplayValues();
  let filtered = data.filter(row => (user.role === 'Siswa') ? (row[3] === user.userID) : true);
  filtered.reverse();

  return { list: filtered.slice((page - 1) * limit, (page - 1) * limit + limit), total: filtered.length };
}

function updateStatusIzin(izinID, statusApproval, catatanWali) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('IzinSakit');
  if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Data tidak ditemukan.' };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getDisplayValues();
  let rowTarget = -1;
  let rowData = null;

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === izinID) {
      rowTarget = i + 2;
      rowData = data[i];
      break;
    }
  }

  if (rowTarget === -1) return { success: false, message: 'ID Izin tidak ditemukan.' };

  sheet.getRange(rowTarget, 8).setValue(statusApproval);
  sheet.getRange(rowTarget, 9).setValue(catatanWali || '-');
  SpreadsheetApp.flush();

  if (statusApproval === 'Approved') {
    const pSheet = ss.getSheetByName('Presensi');
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const presensiID = generateSequentialID('Presensi', 'PRS');

    pSheet.appendRow([
      presensiID,
      rowData[1],
      rowData[3],
      rowData[4],
      'Umum',
      rowData[5],
      'Disetujui Wali Kelas (' + izinID + ')',
      '-',
      timestamp
    ]);
    SpreadsheetApp.flush();
  }

  return { success: true, message: 'Status izin berhasil diperbarui!' };
}

/**
 * 8. JURNAL KBM & NILAI SISWA
 */
function getJurnalKBM(user, page, limit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('JurnalKBM');
  if (!sheet || sheet.getLastRow() <= 1) return { list: [], total: 0 };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getDisplayValues();
  let filtered = (user.role === 'Guru') ? data.filter(row => row[6] === user.userID) : data;
  filtered.reverse();

  return { list: filtered.slice((page - 1) * limit, (page - 1) * limit + limit), total: filtered.length };
}

function submitJurnalKBM(user, formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('JurnalKBM');
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const jurnalID = generateSequentialID('JurnalKBM', 'JRN');

  sheet.appendRow([
    jurnalID,
    String(formData.tanggal),
    String(formData.mapel),
    String(formData.topik),
    String(formData.deskripsi),
    String(formData.prTugas),
    user.userID,
    timestamp,
    String(formData.kelas)
  ]);
  SpreadsheetApp.flush();
  return { success: true, message: 'Jurnal KBM berhasil disimpan.' };
}

function getNilaiSiswa(user, page, limit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('NilaiSiswa');
  if (!sheet || sheet.getLastRow() <= 1) return { list: [], total: 0 };

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getDisplayValues();
  let filtered = (user.role === 'Siswa') ? data.filter(row => row[1] === user.userID) : data;
  filtered.reverse();

  return { list: filtered.slice((page - 1) * limit, (page - 1) * limit + limit), total: filtered.length };
}

function submitNilaiSiswa(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('NilaiSiswa');
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const nilaiID = generateSequentialID('NilaiSiswa', 'NLI');

  sheet.appendRow([
    nilaiID,
    String(formData.userID),
    String(formData.namaSiswa),
    String(formData.mapel),
    String(formData.jenisEvaluasi),
    Number(formData.nilaiAngka),
    String(formData.catatanGuru),
    timestamp
  ]);
  SpreadsheetApp.flush();
  return { success: true, message: 'Nilai siswa berhasil disimpan.' };
}

/**
 * UTILITY: GENERATE SEQUENTIAL ID
 */
function generateSequentialID(sheetName, prefix) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return prefix + '0001';
  const lastRowIndex = sheet.getLastRow() - 1;
  const nextNum = lastRowIndex + 1;
  return prefix + String(nextNum).padStart(4, '0');
}