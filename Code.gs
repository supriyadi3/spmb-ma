const SPREADSHEET_ID = '1gHYdY8SzuUWk2lz0vYK1PnPBgMzxWe6WcXkSfhm5Odc';
const FOLDER_KK_ID = '1RoIsMqYobFgQYEKu7Ay4cDK7v7MqoOWX';
const FOLDER_IJAZAH_ID = '1i2W66OG5WPZTrJe2KttGN2iSV2ZmIUW1';

/**
 * Pintu masuk utama API dari Netlify (Metode POST)
 */
function doPost(e) {
  let result;
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return buildJsonResponse({ status: 'error', message: 'Tidak ada payload data yang diterima.' });
    }

    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;

    if (action === 'login') {
      result = loginUser(requestData.username, requestData.password);
    } else if (action === 'submit') {
      result = submitPendaftaran(
        requestData.formData, 
        requestData.fileKkBase64, 
        requestData.fileKkName, 
        requestData.fileIjazahBase64, 
        requestData.fileIjazahName
      );
    } else if (action === 'getAll') {
      result = getAllPendaftar();
    } else if (action === 'update') {
      result = updatePendaftaran(requestData.formData);
    } else if (action === 'printF5') {
      result = { status: 'success', html: generatePrintF5(requestData.nisn) };
    } else if (action === 'getDownloadUrl') {
      result = { status: 'success', url: getDownloadUrl() };
    } else {
      result = { status: 'error', message: 'Aksi tidak dikenal.' };
    }
  } catch (error) {
    result = { status: 'error', message: 'Terjadi kesalahan sistem: ' + error.toString() };
  }

  return buildJsonResponse(result);
}

/**
 * Fungsi Pembantu mengambil data spreadsheet
 */
function getSpreadsheetData(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    return sheet.getDataRange().getValues();
  } catch (err) {
    return [];
  }
}

/**
 * Autentikasi Login (Admin via Sheet 'USER', Siswa via Sheet 'PENDAFTAR')
 */
function loginUser(username, password) {
  if (!username || !password) {
    return { status: 'error', message: 'Username dan Password wajib diisi!' };
  }

  const uPlain = username.toString().trim();
  const pPlain = password.toString().trim();

  // 1. Cek Login Admin
  const userData = getSpreadsheetData('USER');
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] && userData[i][0].toString().trim() === uPlain && 
        userData[i][1] && userData[i][1].toString().trim() === pPlain) {
      return { status: 'success', role: 'admin', name: 'Administrator' };
    }
  }
  
  // 2. Cek Login Siswa (Username & Password adalah NISN)
  const pendaftarData = getSpreadsheetData('PENDAFTAR');
  if (pendaftarData.length > 0) {
    const headers = pendaftarData[0].map(h => h.toString().trim());
    const nisnIdx = headers.indexOf('NISN');
    const namaIdx = headers.indexOf('Nama Lengkap');
    
    if (nisnIdx !== -1) {
      for (let i = 1; i < pendaftarData.length; i++) {
        if (pendaftarData[i][nisnIdx] && pendaftarData[i][nisnIdx].toString().trim() === uPlain && uPlain === pPlain) {
          let studentData = {};
          headers.forEach((header, index) => {
            let val = pendaftarData[i][index];
            if (val instanceof Date) {
              val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
            }
            studentData[header] = val === undefined || val === null ? '' : val.toString();
          });
          return { status: 'success', role: 'siswa', name: pendaftarData[i][namaIdx] || 'Siswa', data: studentData };
        }
      }
    }
  }
  return { status: 'error', message: 'Username atau Password (NISN) salah!' };
}

/**
 * Menyimpan data pendaftaran baru dan upload berkas dokumen ke Google Drive
 */
function submitPendaftaran(formData, fileKkBase64, fileKkName, fileIjazahBase64, fileIjazahName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('PENDAFTAR');
    const dataRange = sheet.getDataRange().getValues();
    const headers = dataRange[0].map(h => h.toString().trim());
    
    const nisnInput = formData['NISN'] ? formData['NISN'].toString().trim() : '';
    if (!nisnInput) {
      return { status: 'error', message: 'Gagal, NISN wajib diisi.' };
    }

    // Proteksi Duplikasi: Cek apakah NISN sudah terdaftar
    const nisnIdx = headers.indexOf('NISN');
    if (nisnIdx !== -1) {
      for (let i = 1; i < dataRange.length; i++) {
        if (dataRange[i][nisnIdx] && dataRange[i][nisnIdx].toString().trim() === nisnInput) {
          return { status: 'error', message: 'Gagal! Siswa dengan NISN ' + nisnInput + ' sudah terdaftar.' };
        }
      }
    }
    
    // Proses upload KK jika ada
    let linkKk = '';
    if (fileKkBase64 && fileKkName) {
      const folderKk = DriveApp.getFolderById(FOLDER_KK_ID);
      const decodedKk = Utilities.base64Decode(fileKkBase64.split(",")[1]);
      const mimeKk = fileKkBase64.split(",")[0].split(";")[0].split(":")[1];
      const blobKk = Utilities.newBlob(decodedKk, mimeKk, "KK_" + nisnInput + "_" + fileKkName);
      const fileKk = folderKk.createFile(blobKk);
      fileKk.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      linkKk = fileKk.getUrl();
    }
    
    // Proses upload Ijazah jika ada
    let linkIjazah = '';
    if (fileIjazahBase64 && fileIjazahName) {
      const folderIjazah = DriveApp.getFolderById(FOLDER_IJAZAH_ID);
      const decodedIjz = Utilities.base64Decode(fileIjazahBase64.split(",")[1]);
      const mimeIjz = fileIjazahBase64.split(",")[0].split(";")[0].split(":")[1];
      const blobIjz = Utilities.newBlob(decodedIjz, mimeIjz, "IJZ_" + nisnInput + "_" + fileIjazahName);
      const fileIjz = folderIjazah.createFile(blobIjz);
      fileIjz.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      linkIjazah = fileIjz.getUrl();
    }
    
    formData['Link File KK'] = linkKk;
    formData['Link File Ijazah'] = linkIjazah;
    formData['Tanggal Daftar'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    let rowValues = [];
    headers.forEach(header => {
      rowValues.push(formData[header] !== undefined && formData[header] !== null ? formData[header] : '');
    });
    
    sheet.appendRow(rowValues);
    return { status: 'success', message: 'Pendaftaran berhasil disimpan! Silakan login menggunakan NISN Anda.' };
  } catch (error) {
    return { status: 'error', message: 'Gagal simpan pendaftaran: ' + error.toString() };
  }
}

/**
 * Mengambil seluruh data pendaftar untuk Dashboard Admin
 */
function getAllPendaftar() {
  const data = getSpreadsheetData('PENDAFTAR');
  if (data.length <= 1) return [];
  const headers = data[0].map(h => h.toString().trim());
  let result = [];
  for(let i = 1; i < data.length; i++) {
    let row = {};
    headers.forEach((header, index) => {
      let val = data[i][index];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      row[header] = val === undefined || val === null ? '' : val.toString();
    });
    result.push(row);
  }
  return result;
}

/**
 * Memperbarui data pendaftar dari formulir edit siswa mandiri / admin
 */
function updatePendaftaran(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('PENDAFTAR');
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    const nisnIdx = headers.indexOf('NISN');
    
    if (nisnIdx === -1 || !formData['NISN']) {
      return { status: 'error', message: 'Parameter NISN tidak valid.' };
    }

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][nisnIdx] && data[i][nisnIdx].toString().trim() === formData['NISN'].toString().trim()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex !== -1) {
      headers.forEach((header, index) => {
        // Jangan timpa link berkas dengan string kosong saat siswa hanya meng-update teks biodata
        if (formData[header] !== undefined && header !== 'Link File KK' && header !== 'Link File Ijazah' && header !== 'Tanggal Daftar') {
          sheet.getRange(rowIndex, index + 1).setValue(formData[header]);
        }
      });
      return { status: 'success', message: 'Biodata berhasil diperbarui!' };
    }
    return { status: 'error', message: 'Data tidak ditemukan di database.' };
  } catch (error) {
    return { status: 'error', message: 'Gagal update data: ' + error.toString() };
  }
}

/**
 * Membuat dokumen cetak Formulir F5 berbasis HTML
 */
function generatePrintF5(nisn) {
  try {
    const pendaftarData = getSpreadsheetData('PENDAFTAR');
    if (pendaftarData.length <= 1) return "<h3>Error: Data pendaftar kosong.</h3>";
    
    const headers = pendaftarData[0].map(h => h.toString().trim());
    const nisnIdx = headers.indexOf('NISN');
    
    let studentRow = null;
    for (let i = 1; i < pendaftarData.length; i++) {
      if (pendaftarData[i][nisnIdx] && pendaftarData[i][nisnIdx].toString().trim() === nisn.toString().trim()) {
        studentRow = pendaftarData[i];
        break;
      }
    }
    
    if (!studentRow) return "<h3>Error: Data siswa dengan NISN " + nisn + " tidak ditemukan!</h3>";
    
    let d = {};
    headers.forEach((header, idx) => {
      let val = studentRow[idx];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd-MM-yyyy");
      }
      d[header] = val === undefined || val === null ? '' : val.toString().trim();
    });
    
    // Sinkronisasi pemetaan key alternatif untuk mencegah error pembacaan properti di HTML template
    let namaSekolah = d['Nama Sekolah Asal'] || d['Nama Sekolah'] || '-';
    let jenjangSekolah = d['Jenjang Asal'] || d['Jenjang Sekolah Asal'] || '-';
    let statusSekolah = d['Status Asal'] || d['Status Sekolah Asal'] || '-';
    let npsnSekolah = d['NPSN Asal'] || d['NPSN Sekolah Asal'] || '-';
    let alamatSekolah = d['Alamat Sekolah'] || d['Alamat Sekolah Asal'] || '-';

    let hasWali = d['Nama Wali'] && d['Nama Wali'] !== '' && d['Nama Wali'] !== d['Nama Ayah'] && d['Nama Wali'] !== d['Nama Ibu'];
    let waliHtmlSection = '';
    
    if (hasWali) {
      waliHtmlSection = `
        <div class="section-title">D. DATA WALI MURID</div>
        <table>
          <tr><td class="w-35">Nama Wali</td><td class="w-2">:</td><td>${d['Nama Wali']}</td></tr>
          <tr><td>NIK Wali</td><td>:</td><td>${d['NIK Wali'] || '-'}</td></tr>
          <tr><td>Hubungan Keluarga</td><td>:</td><td>${d['Hubungan Keluarga'] || '-'}</td></tr>
          <tr><td>Pendidikan Terakhir</td><td>:</td><td>${d['Pendidikan Wali'] || '-'}</td></tr>
          <tr><td>Pekerjaan Utama</td><td>:</td><td>${d['Pekerjaan Wali'] || '-'}</td></tr>
          <tr><td>Penghasilan Bulanan</td><td>:</td><td>${d['Penghasilan Wali'] || '-'}</td></tr>
        </table>
      `;
    }

    return `
      <!DOCTYPE html>
      <html>
      <head>
        <title>Formulir F5_${d['NISN']}</title>
        <style>
          @page { size: legal; margin: 20mm 15mm 20mm 15mm; }
          body { font-family: 'Arial', sans-serif; font-size: 11pt; color: #111; line-height: 1.4; padding: 0; margin: 0; }
          .kop-surat { text-align: center; border-bottom: 3px double #000; padding-bottom: 8px; margin-bottom: 15px; }
          .kop-instansi { font-size: 14pt; font-weight: bold; text-transform: uppercase; letter-spacing: 0.5px; }
          .kop-madrasah { font-size: 18pt; font-weight: bold; color: #065f46; text-transform: uppercase; margin: 2px 0; }
          .kop-sub { font-size: 9pt; font-style: italic; color: #444; }
          .doc-title { text-align: center; font-size: 13pt; font-weight: bold; text-decoration: underline; margin-bottom: 20px; text-transform: uppercase; }
          .section-title { font-size: 11pt; font-weight: bold; background-color: #f3f4f6; padding: 4px 8px; margin-top: 15px; margin-bottom: 8px; border-left: 4px solid #047857; }
          table { width: 100%; border-collapse: collapse; margin-bottom: 10px; }
          td { padding: 4px 5px; vertical-align: top; font-size: 10.5pt; }
          .w-35 { width: 35%; }
          .w-2 { width: 2%; }
          .footer-sign { width: 100%; margin-top: 35px; }
          .footer-sign td { text-align: center; width: 50%; }
          .space-sign { height: 75px; }
          .bold { font-weight: bold; }
          .no-print-btn { background: #047857; color: white; border: none; padding: 8px 16px; font-weight: bold; border-radius: 4px; cursor: pointer; margin-bottom: 20px; }
          @media print { .no-print-btn { display: none; } }
        </style>
      </head>
      <body>
        <button class="no-print-btn" onclick="window.print()">Pemicu Cetak Dokumen (Print)</button>
        
        <div class="kop-surat">
          <div class="kop-instansi">Yayasan Pendidikan Islam Al Istiqomah</div>
          <div class="kop-madrasah">Madrasah Aliyah Al Istiqomah</div>
          <div class="kop-sub">NSM: 131232010001 | Akreditasi A | Jl. Kalisari No. 09, Kotabaru, Karawang, Jawa Barat</div>
        </div>
        
        <div class="doc-title">Formulir Pendaftaran Peserta Didik Baru (F5)</div>
        
        <div class="section-title">A. BIODATA SISWA</div>
        <table>
          <tr><td class="w-35">Nama Lengkap Siswa</td><td class="w-2">:</td><td class="bold">${d['Nama Lengkap'] || '-'}</td></tr>
          <tr><td>Nomor Induk Siswa Nasional (NISN)</td><td>:</td><td class="bold">${d['NISN'] || '-'}</td></tr>
          <tr><td>NIK / No. KTP Calon Siswa</td><td>:</td><td>${d['NIK'] || '-'}</td></tr>
          <tr><td>Tempat, Tanggal Lahir</td><td>:</td><td>${d['Tempat Lahir'] || '-'}, ${d['Tanggal Lahir'] || '-'}</td></tr>
          <tr><td>Hobi Pembawaan</td><td>:</td><td>${d['Hobi'] || '-'}</td></tr>
          <tr><td>Cita-cita Siswa</td><td>:</td><td>${d['Cita-cita'] || '-'}</td></tr>
          <tr><td>Alamat Lengkap Rumah</td><td>:</td><td>${d['Alamat Lengkap'] || '-'}</td></tr>
          <tr><td>Riwayat Pendidikan Usia Dini</td><td>:</td><td>TK: ${d['Pernah TK'] || '-'} | PAUD: ${d['Pernah PAUD'] || '-'}</td></tr>
        </table>
        
        <div class="section-title">B. DATA SEKOLAH ASAL</div>
        <table>
          <tr><td class="w-35">Nama Satuan Pendidikan Asal</td><td class="w-2">:</td><td>${namaSekolah} (${jenjangSekolah})</td></tr>
          <tr><td>Status Kelembagaan Sekolah</td><td>:</td><td>${statusSekolah}</td></tr>
          <tr><td>Nomor Pokok Sekolah Nasional (NPSN)</td><td>:</td><td>${npsnSekolah}</td></tr>
          <tr><td>Alamat Lokasi Instansi Sekolah</td><td>:</td><td>${alamatSekolah}</td></tr>
        </table>
        
        <div class="section-title">C. DATA ORANG TUA</div>
        <table>
          <tr>
            <td class="w-35 bold" style="text-decoration: underline;">Data Ayah Kandung:</td>
            <td class="w-2"></td>
            <td class="bold" style="text-decoration: underline;">Data Ibu Kandung:</td>
          </tr>
          <tr>
            <td>
              Nama: ${d['Nama Ayah'] || '-'}<br>
              NIK: ${d['NIK Ayah'] || '-'}<br>
              Lahir: ${d['Tempat Lahir Ayah'] || '-'}, ${d['Tanggal Lahir Ayah'] || '-'}<br>
              Status: ${d['Status Ayah'] || '-'}<br>
              Pendidikan: ${d['Pendidikan Ayah'] || '-'}<br>
              Pekerjaan: ${d['Pekerjaan Ayah'] || '-'}<br>
              Penghasilan: ${d['Penghasilan Ayah'] || '-'}
            </td>
            <td></td>
            <td>
              Nama: ${d['Nama Ibu'] || '-'}<br>
              NIK: ${d['NIK Ibu'] || '-'}<br>
              Lahir: ${d['Tempat Lahir Ibu'] || '-'}, ${d['Tanggal Lahir Ibu'] || '-'}<br>
              Status: ${d['Status Ibu'] || '-'}<br>
              Pendidikan: ${d['Pendidikan Ibu'] || '-'}<br>
              Pekerjaan: ${d['Pekerjaan Ibu'] || '-'}<br>
              Penghasilan: ${d['Penghasilan Ibu'] || '-'}
            </td>
          </tr>
          <tr>
            <td colspan="3" style="padding-top:8px;">No. Kartu Keluarga (KK): <b>${d['No KK'] || '-'}</b> &nbsp;|&nbsp; Kepala Keluarga: <b>${d['Nama Kepala Keluarga'] || '-'}</b></td>
          </tr>
        </table>
        
        ${waliHtmlSection}
        
        <table class="footer-sign">
          <tr>
            <td>
              <br>Orang Tua / Wali Murid,
              <div class="space-sign"></div>
              ( ...................................................... )
            </td>
            <td>
              Karawang, ............................. 2026<br>
              Panitia Pendaftaran SPMB,
              <div class="space-sign"></div>
              ( ...................................................... )
            </td>
          </tr>
        </table>
      </body>
      </html>
    `;
  } catch(err) {
    return "<h3>Terjadi masalah internal pembuatan F5: " + err.toString() + "</h3>";
  }
}

function getDownloadUrl() {
  return "https://docs.google.com/spreadsheets/d/" + SPREADSHEET_ID + "/export?format=xlsx";
}

/**
 * Fungsi Pembantu Standarisasi Response JSON demi Keamanan CORS Netlify
 */
function buildJsonResponse(objekData) {
  return ContentService.createTextOutput(JSON.stringify(objekData))
    .setMimeType(ContentService.MimeType.JSON);
}
