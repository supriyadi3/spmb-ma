const SPREADSHEET_ID = '1gHYdY8SzuUWk2lz0vYK1PnPBgMzxWe6WcXkSfhm5Odc';
const FOLDER_KK_ID = '1RoIsMqYobFgQYEKu7Ay4cDK7v7MqoOWX';
const FOLDER_IJAZAH_ID = '1i2W66OG5WPZTrJe2KttGN2iSV2ZmIUW1';

// Tambahan fungsi doPost sebagai pintu masuk API dari Netlify
function doPost(e) {
  let result;
  try {
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
    result = { status: 'error', message: error.toString() };
  }

  // Mengembalikan respons dalam bentuk JSON (Wajib untuk integrasi lintas platform)
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// Menghapus fungsi doGet lama karena HTML sekarang di-host di Netlify

function getSpreadsheetData(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  return sheet.getDataRange().getValues();
}

function loginUser(username, password) {
  const userData = getSpreadsheetData('USER');
  for (let i = 1; i < userData.length; i++) {
    if (userData[i][0] && userData[i][0].toString().trim() === username.toString().trim() && 
        userData[i][1] && userData[i][1].toString().trim() === password.toString().trim()) {
      return { status: 'success', role: 'admin', name: 'Administrator' };
    }
  }
  
  const pendaftarData = getSpreadsheetData('PENDAFTAR');
  if (pendaftarData.length > 0) {
    const headers = pendaftarData[0].map(h => h.toString().trim());
    const nisnIdx = headers.indexOf('NISN');
    const namaIdx = headers.indexOf('Nama Lengkap');
    
    for (let i = 1; i < pendaftarData.length; i++) {
      if (pendaftarData[i][nisnIdx] && pendaftarData[i][nisnIdx].toString().trim() === username.toString().trim() && 
          username.toString().trim() === password.toString().trim()) {
        
        let studentData = {};
        headers.forEach((header, index) => {
          let val = pendaftarData[i][index];
          if (val instanceof Date) {
            val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
          studentData[header] = val === undefined || val === null ? '' : val.toString();
        });
        return { status: 'success', role: 'siswa', name: pendaftarData[i][namaIdx], data: studentData };
      }
    }
  }
  return { status: 'error', message: 'Username atau Password salah!' };
}

function submitPendaftaran(formData, fileKkBase64, fileKkName, fileIjazahBase64, fileIjazahName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('PENDAFTAR');
    const headers = sheet.getDataRange().getValues()[0].map(h => h.toString().trim());
    
    let linkKk = '';
    let linkIjazah = '';
    
    if (fileKkBase64) {
      const folderKk = DriveApp.getFolderById(FOLDER_KK_ID);
      const decodedKk = Utilities.base64Decode(fileKkBase64.split(",")[1]);
      const blobKk = Utilities.newBlob(decodedKk, fileKkBase64.split(",")[0].split(";")[0].split(":")[1], fileKkName);
      const fileKk = folderKk.createFile(blobKk);
      fileKk.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      linkKk = fileKk.getUrl();
    }
    
    if (fileIjazahBase64) {
      const folderIjazah = DriveApp.getFolderById(FOLDER_IJAZAH_ID);
      const decodedIjz = Utilities.base64Decode(fileIjazahBase64.split(",")[1]);
      const blobIjz = Utilities.newBlob(decodedIjz, fileIjazahBase64.split(",")[0].split(";")[0].split(":")[1], fileIjazahName);
      const fileIjz = folderIjazah.createFile(blobIjz);
      fileIjz.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      linkIjazah = fileIjz.getUrl();
    }
    
    formData['Link File KK'] = linkKk;
    formData['Link File Ijazah'] = linkIjazah;
    
    let rowValues = [];
    headers.forEach(header => {
      rowValues.push(formData[header] || '');
    });
    
    sheet.appendRow(rowValues);
    return { status: 'success', message: 'Pendaftaran berhasil disimpan!' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

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

function updatePendaftaran(formData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('PENDAFTAR');
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    const nisnIdx = headers.indexOf('NISN');
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][nisnIdx] && data[i][nisnIdx].toString().trim() === formData['NISN'].toString().trim()) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex !== -1) {
      headers.forEach((header, index) => {
        if (formData[header] !== undefined && header !== 'Link File KK' && header !== 'Link File Ijazah') {
          sheet.getRange(rowIndex, index + 1).setValue(formData[header]);
        }
      });
      return { status: 'success', message: 'Biodata berhasil diperbarui!' };
    }
    return { status: 'error', message: 'Data tidak ditemukan.' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function generatePrintF5(nisn) {
  // Perbaikan typo pembungkusan variabel let pada kode asli
  try {
    const pendaftarData = getSpreadsheetData('PENDAFTAR');
    if (pendaftarData.length <= 1) return "Data pendaftar kosong.";
    
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
    
    let hasWali = d['Nama Wali'] && d['Nama Wali'] !== '' && d['Nama Wali'] !== d['Nama Ayah'] && d['Nama Wali'] !== d['Nama Ibu'];
    let waliHtmlSection = '';
    
    if (hasWali) {
      waliHtmlSection = `
        <div class="section-title">D. DATA WALI MURID</div>
        <table>
          <tr><td class="w-35">Nama Wali</td><td class="w-2">:</td><td>${d['Nama Wali'] || '-'}</td></tr>
          <tr><td>NIK Wali</td><td>:</td><td>${d['NIK Wali'] || '-'}</td></tr>
          <tr><td>Tempat, Tanggal Lahir</td><td>:</td><td>${d['Tempat Lahir Wali'] || '-'}, ${d['Tanggal Lahir Wali'] || '-'}</td></tr>
          <tr><td>Pendidikan Terakhir</td><td>:</td><td>${d['Pendidikan Wali'] || '-'}</td></tr>
          <tr><td>Pekerjaan Utama</td><td>:</td><td>${d['Pekerjaan Wali'] || '-'}</td></tr>
          <tr><td>Penghasilan Bulanan</td><td>:</td><td>${d['Penghasilan Wali'] || '-'}</td></tr>
        </table>
      `;
    }

    let htmlOutput = `
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
          <tr><td class="w-35">Nama Satuan Pendidikan Asal</td><td class="w-2">:</td><td>${d['Nama Sekolah Asal'] || '-'} (${d['Jenjang Sekolah Asal'] || '-'})</td></tr>
          <tr><td>Status Kelembagaan Sekolah</td><td>:</td><td>${d['Status Sekolah Asal'] || '-'}</td></tr>
          <tr><td>Nomor Pokok Sekolah Nasional (NPSN)</td><td>:</td><td>${d['NPSN Sekolah Asal'] || '-'}</td></tr>
          <tr><td>Alamat Lokasi Instansi Sekolah</td><td>:</td><td>${d['Alamat Sekolah Asal'] || '-'}</td></tr>
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
    return htmlOutput;
  } catch(err) {
    return " terjadi masalah internal: " + err.toString();
  }
}

function getDownloadUrl() {
  return "https://docs.google.com/spreadsheets/d/" + SPREADSHEET_ID + "/export?format=xlsx";
}
