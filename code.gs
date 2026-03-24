// ==========================================
// KONFIGURASI UTAMA
// ==========================================
const SPREADSHEET_ID = '1HHUpqW_dJ9sFN9rNYEyv2NUqWBqHLM6H4CPqtbuLidY'; 
const CONFIG_FOLDER_ID = '1WkxAmoxn7kN--3a99UP7O_zECJEUi9bg'; // Folder ID untuk menyimpan Logo & Foto Barang

const SHEET_USERS = 'Users';
const SHEET_INVENTARIS = 'Inventaris';
const SHEET_LOG = 'Log';
const SHEET_CONFIG = 'Config';

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ==========================================
// 1. FUNGSI UTILITY & FORMATTING
// ==========================================

function convertSheetDate(dateVal) {
  if (!dateVal) return null;
  if (dateVal instanceof Date) return dateVal;
  if (typeof dateVal === 'string') {
    try {
      var parts = dateVal.split(' ');
      if (parts.length < 1) return new Date(); 
      
      var dateParts = parts[0].split('/'); 
      if (dateParts.length !== 3) return new Date(dateVal);
      var timeParts = parts.length > 1 ? parts[1].split(':') : [0, 0, 0];
      return new Date(dateParts[2], dateParts[1] - 1, dateParts[0], timeParts[0] || 0, timeParts[1] || 0, timeParts[2] || 0);
    } catch (e) {
      return new Date();
    }
  }
  return dateVal;
}

function formatDateToString(dateVal) {
  if (!dateVal) return "";
  try {
    if (typeof dateVal === 'string') return dateVal; 
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  } catch (e) {
    return String(dateVal);
  }
}

// FUNGSI INIT (DIPERBAIKI: Menambahkan Header Foto Barang)
function initializeSheets() {
  const ss = getSpreadsheet();
  
  // 1. Sheet Users
  let usersSheet = ss.getSheetByName(SHEET_USERS);
  if (!usersSheet) {
    usersSheet = ss.insertSheet(SHEET_USERS);
    usersSheet.appendRow(['ID', 'Username', 'Password', 'Nama', 'Role', 'Avatar', 'Last Login', 'Created Date']);
    usersSheet.appendRow(['1', 'admin', 'admin123', 'Administrator', 'Admin', '', new Date(), new Date()]);
    usersSheet.appendRow(['2', 'petugas', 'petugas123', 'Petugas Gudang', 'Petugas', '', new Date(), new Date()]);
  }
  
  // 2. Sheet Inventaris (UPDATED: Tambah Header Foto Barang)
  let inventarisSheet = ss.getSheetByName(SHEET_INVENTARIS);
  if (!inventarisSheet) {
    inventarisSheet = ss.insertSheet(SHEET_INVENTARIS);
    // Kolom ke-13 adalah Foto Barang
    inventarisSheet.appendRow(['ID', 'Kode Barang', 'Nama Barang', 'Tahun', 'Bulan', 'Kategori', 'Lokasi', 'Kondisi', 'Jumlah', 'QR Code', 'Created Date', 'Updated Date', 'Foto Barang']);
  }
  
  // 3. Sheet Log
  let logSheet = ss.getSheetByName(SHEET_LOG);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEET_LOG);
    logSheet.appendRow(['Timestamp', 'User', 'Action', 'Details']);
  }

  // 4. Sheet Config
  let configSheet = ss.getSheetByName(SHEET_CONFIG);
  if (!configSheet) {
    configSheet = ss.insertSheet(SHEET_CONFIG);
    configSheet.appendRow(['Key', 'Value']);
    
    // Default Values
    configSheet.appendRow(['instansi_baris1', 'PEMERINTAH KABUPATEN BENGKULU']);
    configSheet.appendRow(['instansi_baris2', 'DINAS PENDIDIKAN DAN KEBUDAYAAN']);
    configSheet.appendRow(['nama_sekolah', 'SMP NEGERI CONTOH SISTEM']);
    configSheet.appendRow(['alamat', 'Jl. Merpati No. 123, Ratu Samban, Kota Bengkulu, Kode Pos 38222']);
    configSheet.appendRow(['kontak', 'Telp: (0736) 123456 | Email: admin@sekolah.sch.id']);
    configSheet.appendRow(['logo_kiri', 'https://upload.wikimedia.org/wikipedia/commons/9/98/Kota_Bengkulu.png']);
    configSheet.appendRow(['logo_kanan', 'https://upload.wikimedia.org/wikipedia/commons/9/9c/Logo_of_Ministry_of_Education_and_Culture_of_Republic_of_Indonesia.svg']);
    configSheet.appendRow(['kota_surat', 'Bengkulu']);
    configSheet.appendRow(['nama_kepsek', '(Nama Kepala Sekolah)']);
    configSheet.appendRow(['nip_kepsek', '19800101 200001 1 001']);
    configSheet.appendRow(['nama_petugas', '(Nama Petugas Barang)']);
    configSheet.appendRow(['nip_petugas', '-']);
  }
}

function doGet() {
  initializeSheets();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Sistem Inventaris QR Code')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setFaviconUrl('https://cdn-icons-png.flaticon.com/512/201/201614.png');
}

// ==========================================
// 2. AUTENTIKASI & SESSION
// ==========================================
function login(username, password) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]) === username && String(data[i][2]) === password) {
        
        const now = new Date();
        sheet.getRange(i + 1, 7).setValue(now);

        const rawId = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, (now.getTime() + Math.random().toString()));
        const sessionId = Utilities.base64Encode(rawId);
        
        const cache = CacheService.getScriptCache();
        const userData = {
            id: data[i][0].toString(),
            username: data[i][1],
            nama: data[i][3],
            role: data[i][4],
            avatar: data[i][5]
        };
        cache.put(sessionId, JSON.stringify(userData), 21600); // 6 jam
        addLog(data[i][3], 'Login', 'User berhasil login');
        
        return {
          success: true,
          user: {
            id: data[i][0],
            username: data[i][1],
            nama: data[i][3],
            role: data[i][4],
            avatar: data[i][5],
            lastLogin: formatDateToString(now),
            sessionId: sessionId
          }
        };
      }
    }
    return { success: false, message: 'Username atau password salah' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function checkSession(sessionId) {
  try {
    if (!sessionId) return { success: false };
    const cache = CacheService.getScriptCache();
    const cachedData = cache.get(sessionId);

    if (cachedData) {
      const userData = JSON.parse(cachedData);
      cache.put(sessionId, cachedData, 21600); // Perpanjang
      return { success: true, user: userData };
    }
    return { success: false };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function logout(sessionId) {
  try {
    const cache = CacheService.getScriptCache();
    const cachedData = cache.get(sessionId);
    let nama = "Unknown";
    if (cachedData) {
       nama = JSON.parse(cachedData).nama;
       cache.remove(sessionId);
    }
    addLog(nama, 'Logout', 'User logout dari sistem');
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// 3. FUNGSI KONFIGURASI & UPLOAD LOGO
// ==========================================

function getConfig() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_CONFIG);
    const data = sheet.getDataRange().getValues();
    
    let config = {};
    for (let i = 1; i < data.length; i++) {
      config[data[i][0]] = data[i][1];
    }
    return { success: true, data: config };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function saveConfig(formObject) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_CONFIG);
    const data = sheet.getDataRange().getValues();
    
    const updateOrInsert = (key, value) => {
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === key) {
          sheet.getRange(i + 1, 2).setValue(value);
          found = true;
          break;
        }
      }
      if (!found) sheet.appendRow([key, value]);
    };

    for (const key in formObject) {
      updateOrInsert(key, formObject[key]);
    }

    addLog('Admin', 'Update Config', 'Memperbarui konfigurasi aplikasi');
    return { success: true, message: 'Konfigurasi berhasil disimpan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function uploadLogo(data, filename, type) {
  try {
    const decoded = Utilities.base64Decode(data);
    const blob = Utilities.newBlob(decoded, type, filename);
    const folder = DriveApp.getFolderById(CONFIG_FOLDER_ID);
    const file = folder.createFile(blob);
    
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const fileId = file.getId();
    const directUrl = "https://drive.google.com/uc?export=view&id=" + fileId;
    
    return { success: true, url: directUrl };
  } catch (error) {
    return { success: false, message: "Upload gagal: " + error.toString() };
  }
}

// ==========================================
// 4. FUNGSI UPLOAD FOTO BARANG (BARU)
// ==========================================

function uploadInventoryFile(data, filename, type) {
  try {
    const decoded = Utilities.base64Decode(data);
    const blob = Utilities.newBlob(decoded, type, filename);
    
    // Menggunakan Folder ID yang sama dengan Config
    const folder = DriveApp.getFolderById(CONFIG_FOLDER_ID); 
    const file = folder.createFile(blob);
    
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Format URL sesuai permintaan: https://drive.google.com/file/d/[ID]/view
    const fileId = file.getId();
    const viewUrl = "https://drive.google.com/file/d/" + fileId + "/view";
    
    return { success: true, url: viewUrl };
  } catch (error) {
    return { success: false, message: "Upload gagal: " + error.toString() };
  }
}

// ==========================================
// 5. FUNGSI DASHBOARD & INVENTARIS
// ==========================================

function getDashboardStats() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_INVENTARIS);
    
    if (!sheet || sheet.getLastRow() <= 1) {
       return { success: true, stats: { totalBarang: 0, totalItem: 0, barangBaik: 0, barangRusak: 0, barangHilang: 0 } };
    }

    const data = sheet.getDataRange().getValues();
    let totalBarang = 0, barangBaik = 0, barangRusak = 0, barangHilang = 0;
    
    for (let i = 1; i < data.length; i++) {
      let jumlah = parseInt(data[i][8]);
      if (isNaN(jumlah)) jumlah = 0;
      
      totalBarang += jumlah;
      const kondisi = String(data[i][7]).toLowerCase().trim();
      
      if (kondisi === 'baik') barangBaik += jumlah;
      else if (kondisi === 'rusak') barangRusak += jumlah;
      else if (kondisi === 'hilang') barangHilang += jumlah;
    }
    
    return {
      success: true,
      stats: {
        totalBarang: totalBarang,
        totalItem: data.length - 1,
        barangBaik: barangBaik,
        barangRusak: barangRusak,
        barangHilang: barangHilang
      }
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getAllInventaris() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_INVENTARIS);
    
    if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [] };

    const data = sheet.getDataRange().getValues();
    const items = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      items.push({
        id: data[i][0],
        kodeBarang: String(data[i][1]),
        namaBarang: String(data[i][2]),
        tahun: String(data[i][3] || ''),
        bulan: String(data[i][4] || ''),
        kategori: String(data[i][5]),
        lokasi: String(data[i][6]),
        kondisi: String(data[i][7]),
        jumlah: data[i][8],
        qrCode: String(data[i][9]),
        createdDate: formatDateToString(data[i][10]),
        updatedDate: formatDateToString(data[i][11]),
        foto: String(data[i][12] || '') // Kolom 13 (Index 12)
      });
    }
    return { success: true, data: items };
  } catch (error) {
    return { success: false, message: "Server Error: " + error.toString() };
  }
}

function getInventarisById(id) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_INVENTARIS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === id.toString()) {
        return {
          success: true,
          data: {
            id: data[i][0],
            kodeBarang: data[i][1],
            namaBarang: data[i][2],
            tahun: data[i][3],
            bulan: data[i][4],
            kategori: data[i][5],
            lokasi: data[i][6],
            kondisi: data[i][7],
            jumlah: data[i][8],
            qrCode: data[i][9],
            createdDate: formatDateToString(data[i][10]),
            updatedDate: formatDateToString(data[i][11]),
            foto: data[i][12] || '' // Kolom 13
          }
        };
      }
    }
    return { success: false, message: 'ID Barang tidak ditemukan' };
  } catch (error) {
    return { success: false, message: 'Server Error: ' + error.toString() };
  }
}

function addInventaris(item) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_INVENTARIS);
    const data = sheet.getDataRange().getValues();
    
    // Cek Duplikat Kode Barang
    const inputKode = String(item.kodeBarang).trim().toUpperCase();
    for (let i = 1; i < data.length; i++) {
      const existingKode = String(data[i][1]).trim().toUpperCase();
      if (existingKode === inputKode) {
        return { success: false, message: 'Kode Barang "' + item.kodeBarang + '" sudah ada!' };
      }
    }

    const lastRow = sheet.getLastRow();
    const id = lastRow > 0 ? lastRow : 1; 
    
    const qrData = `INV-${id}-${item.kodeBarang}`;
    const qrCode = `https://quickchart.io/chart?cht=qr&chs=200x200&chl=${encodeURIComponent(qrData)}`;
    const now = new Date();
    
    // Append Row (Perhatikan penambahan item.foto di akhir)
    sheet.appendRow([
      id,
      item.kodeBarang,
      item.namaBarang,
      item.tahun,
      item.bulan,
      item.kategori,
      item.lokasi,
      item.kondisi,
      item.jumlah,
      qrCode,
      now,
      now,
      item.foto || '' // Simpan URL Foto
    ]);

    const namaUser = item.editorName || 'Unknown User';
    addLog(namaUser, 'Tambah Barang', `Menambahkan barang: ${item.namaBarang} (${item.kodeBarang})`);
    
    return { success: true, message: 'Barang berhasil ditambahkan', id: id, qrCode: qrCode };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function updateInventaris(item) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_INVENTARIS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === item.id.toString()) {
        const row = i + 1;
        sheet.getRange(row, 2).setValue(item.kodeBarang);
        sheet.getRange(row, 3).setValue(item.namaBarang);
        sheet.getRange(row, 4).setValue(item.tahun);
        sheet.getRange(row, 5).setValue(item.bulan);
        sheet.getRange(row, 6).setValue(item.kategori);
        sheet.getRange(row, 7).setValue(item.lokasi);
        sheet.getRange(row, 8).setValue(item.kondisi);
        sheet.getRange(row, 9).setValue(item.jumlah);
        sheet.getRange(row, 12).setValue(new Date());
        
        // Update Foto (Kolom 13)
        sheet.getRange(row, 13).setValue(item.foto); 

        const namaUser = item.editorName || 'Unknown User';
        addLog(namaUser, 'Update Barang', `Mengupdate barang: ${item.namaBarang}`);
        
        return { success: true, message: 'Barang berhasil diupdate' };
      }
    }
    return { success: false, message: 'Barang tidak ditemukan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function deleteInventaris(id, editorName) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_INVENTARIS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === id.toString()) {
        const namaBarang = data[i][2];
        sheet.deleteRow(i + 1);
        
        const namaUser = editorName || 'Unknown User';
        addLog(namaUser, 'Hapus Barang', `Menghapus barang: ${namaBarang}`);
        
        return { success: true, message: 'Barang berhasil dihapus' };
      }
    }
    return { success: false, message: 'Barang tidak ditemukan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// 6. FUNGSI USER MANAGEMENT & LOGS
// ==========================================

function addLog(user, action, details) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_LOG);
    sheet.appendRow([new Date(), user, action, details]);
  } catch (error) {
    Logger.log('Error adding log: ' + error.toString());
  }
}

function getAllLogs() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_LOG);
    if (sheet.getLastRow() <= 1) return { success: true, data: [] };

    const data = sheet.getDataRange().getValues();
    const logs = [];
    for (let i = 1; i < data.length; i++) {
      logs.push({
        timestamp: formatDateToString(convertSheetDate(data[i][0])),
        user: data[i][1],
        action: data[i][2],
        details: data[i][3]
      });
    }
    return { success: true, data: logs.reverse() };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function getAllUsers() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_USERS);
    if (sheet.getLastRow() <= 1) return { success: true, data: [] };

    const data = sheet.getDataRange().getValues();
    const users = [];
    for (let i = 1; i < data.length; i++) {
      users.push({
        id: data[i][0],
        username: data[i][1],
        nama: data[i][3],
        role: data[i][4],
        avatar: data[i][5],
        lastLogin: formatDateToString(convertSheetDate(data[i][6])),
        createdDate: formatDateToString(convertSheetDate(data[i][7]))
      });
    }
    return { success: true, data: users };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function addUser(user) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase() === String(user.username).toLowerCase()) {
        return { success: false, message: 'Username sudah digunakan!' };
      }
    }

    let maxId = 0;
    for (let i = 1; i < data.length; i++) {
      const currentId = parseInt(data[i][0]);
      if (currentId > maxId) maxId = currentId;
    }
    const newId = maxId + 1;

    sheet.appendRow([
      newId,
      user.username,
      user.password, 
      user.nama,
      user.role,
      '', '', new Date()
    ]);
    addLog('Admin', 'Tambah User', `Menambahkan pengguna baru: ${user.username}`);
    return { success: true, message: 'Pengguna baru berhasil ditambahkan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function updateUser(user) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === user.id.toString()) {
        sheet.getRange(i + 1, 2).setValue(user.username);
        if(user.password) sheet.getRange(i + 1, 3).setValue(user.password);
        sheet.getRange(i + 1, 4).setValue(user.nama);
        sheet.getRange(i + 1, 5).setValue(user.role);
        return { success: true, message: 'Data pengguna berhasil diperbarui' };
      }
    }
    return { success: false, message: 'User tidak ditemukan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

function deleteUser(id) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === id.toString()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Pengguna berhasil dihapus' };
      }
    }
    return { success: false, message: 'User tidak ditemukan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// ==========================================
// 7. FUNGSI PDF GENERATOR
// ==========================================

function imageToDataUri(url) {
  if (!url) return "";
  try {
    let blob;
    const driveIdMatch = url.match(/[-\w]{25,}/);
    if (driveIdMatch) {
      const id = driveIdMatch[0];
      blob = DriveApp.getFileById(id).getBlob();
    } else {
      blob = UrlFetchApp.fetch(url).getBlob();
    }
    const base64 = Utilities.base64Encode(blob.getBytes());
    return "data:" + blob.getContentType() + ";base64," + base64;
  } catch (e) {
    Logger.log("Gagal konversi gambar: " + url);
    return ""; 
  }
}

function generatePDF() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_INVENTARIS);
    
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: "Data inventaris kosong." };

    const configResult = getConfig();
    const DB_CONF = configResult.success ? configResult.data : {};

    const logoKiriBase64 = imageToDataUri(DB_CONF.logo_kiri);
    const logoKananBase64 = imageToDataUri(DB_CONF.logo_kanan);

    const CONF = {
      instansi_baris1: DB_CONF.instansi_baris1 || "PEMERINTAH",
      instansi_baris2: DB_CONF.instansi_baris2 || "DINAS TERKAIT",
      nama_sekolah:    DB_CONF.nama_sekolah || "NAMA INSTANSI",
      alamat:          DB_CONF.alamat || "Alamat Belum Diatur",
      kontak:          DB_CONF.kontak || "Kontak Belum Diatur",
      logo_kiri:       logoKiriBase64, 
      logo_kanan:      logoKananBase64,
      kota_surat:      DB_CONF.kota_surat || "Kota",
      nama_kepsek:     DB_CONF.nama_kepsek || "(Nama Kepala Sekolah)",
      nip_kepsek:      DB_CONF.nip_kepsek || "-",
      nama_petugas:    DB_CONF.nama_petugas || "(Nama Petugas)",
      nip_petugas:     DB_CONF.nip_petugas || "-"
    };

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    const now = new Date();
    const months = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    const dateStr = `${now.getDate()} ${months[now.getMonth()]} ${now.getFullYear()}`;
    
    let html = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>Laporan Inventaris</title>
        <style>
          @page { size: A4 landscape; margin: 0; }
          body { font-family: 'Times New Roman', Times, serif; margin: 0; padding: 40px 50px; color: #000; background: #fff; }
          .kop-container { width: 100%; border-bottom: 5px double #000; padding-bottom: 10px; margin-bottom: 25px; display: table; }
          .kop-column { display: table-cell; vertical-align: middle; text-align: center; }
          .col-logo { width: 15%; }
          .col-text { width: 70%; }
          .kop-logo-img { max-width: 85px; max-height: 85px; object-fit: contain; } 
          .kop-text h3 { margin: 0; font-size: 16px; font-weight: normal; text-transform: uppercase; letter-spacing: 1px; }
          .kop-text h2 { margin: 5px 0; font-size: 18px; font-weight: bold; text-transform: uppercase; }
          .kop-text h1 { margin: 0; font-size: 22px; font-weight: 900; text-transform: uppercase; letter-spacing: 1.5px; }
          .kop-text p { margin: 5px 0 0; font-size: 11px; font-style: italic; }
          .judul-laporan { text-align: center; margin-bottom: 20px; font-family: Arial, sans-serif; }
          .judul-laporan h2 { margin: 0; font-size: 16px; text-decoration: underline; text-transform: uppercase; }
          .judul-laporan p { margin: 5px 0; font-size: 12px; }
          table { width: 100%; border-collapse: collapse; font-size: 11px; font-family: Arial, sans-serif; }
          thead th { background-color: #e0e0e0; border: 1px solid #000; padding: 8px; text-align: center; font-weight: bold; text-transform: uppercase; }
          tbody td { border: 1px solid #000; padding: 6px; vertical-align: middle; }
          .text-center { text-align: center; }
          .ttd-container { margin-top: 50px; width: 100%; display: table; font-family: Arial, sans-serif; font-size: 12px; }
          .ttd-box { display: table-cell; width: 30%; text-align: center; }
          .spacer { height: 70px; }
        </style>
      </head>
      <body>
        <div class="kop-container">
          <div class="kop-column col-logo">${CONF.logo_kiri ? `<img src="${CONF.logo_kiri}" class="kop-logo-img">` : ''}</div>
          <div class="kop-column col-text kop-text">
            <h3>${CONF.instansi_baris1}</h3>
            <h2>${CONF.instansi_baris2}</h2>
            <h1>${CONF.nama_sekolah}</h1>
            <p>${CONF.alamat}<br>${CONF.kontak}</p>
          </div>
          <div class="kop-column col-logo">${CONF.logo_kanan ? `<img src="${CONF.logo_kanan}" class="kop-logo-img">` : ''}</div>
        </div>
        <div class="judul-laporan"><h2>Laporan Data Aset & Inventaris</h2><p>Per Tanggal: ${dateStr}</p></div>
        <table>
          <thead>
            <tr>
              <th style="width: 5%">No</th><th style="width: 15%">Kode Barang</th><th style="width: 25%">Nama Barang</th>
              <th style="width: 15%">Kategori</th><th style="width: 15%">Lokasi</th><th style="width: 10%">Kondisi</th>
              <th style="width: 5%">Jml</th><th style="width: 10%">Thn/Bln</th>
            </tr>
          </thead>
          <tbody>
    `;
    data.forEach((row, index) => {
      const kondisi = String(row[7]).trim();
      html += `
        <tr>
          <td class="text-center">${index + 1}</td>
          <td class="text-center"><b>${row[1]}</b></td>
          <td>${row[2]}</td><td>${row[5]}</td><td>${row[6]}</td>
          <td class="text-center">${kondisi}</td>
          <td class="text-center">${row[8]}</td>
          <td class="text-center">${row[3]}/${row[4]}</td>
        </tr>`;
    });
    html += `
          </tbody>
        </table>
        <div class="ttd-container">
          <div class="ttd-box"><br>Mengetahui,<br>Kepala Sekolah<br><div class="spacer"></div><b><u>${CONF.nama_kepsek}</u></b><br>NIP. ${CONF.nip_kepsek}</div>
          <div style="display: table-cell; width: 40%;"></div>
          <div class="ttd-box">${CONF.kota_surat}, ${dateStr}<br>Pengurus Barang<br><div class="spacer"></div><b><u>${CONF.nama_petugas}</u></b><br>NIP. ${CONF.nip_petugas}</div>
        </div>
      </body>
      </html>`;
      
    const blob = HtmlService.createHtmlOutput(html).getAs('application/pdf');
    blob.setName(`Laporan_Inventaris_${dateStr.replace(/ /g, '_')}.pdf`);
    const base64 = Utilities.base64Encode(blob.getBytes());
    return { success: true, base64: base64, filename: `Laporan_Inventaris_${dateStr.replace(/ /g, '_')}.pdf` };
  } catch (error) {
    return { success: false, message: "Error PDF: " + error.toString() };
  }
}

function getAllSystemData() {
  const stats = getDashboardStats(); 
  const inventaris = getAllInventaris();
  const logs = getAllLogs();
  const users = getAllUsers();
  const config = getConfig();

  return {
    success: true,
    data: {
      stats: stats.success ? stats.stats : {},
      inventaris: inventaris.success ? inventaris.data : [],
      logs: logs.success ? logs.data : [],
      users: users.success ? users.data : [],
      config: config.success ? config.data : {}
    }
  };
}
