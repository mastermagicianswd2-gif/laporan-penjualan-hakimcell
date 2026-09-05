/**
 * SISTEM LAPORAN PENJUALAN - "AGEN HAKIM CELL"
 * Backend Script (Google Apps Script) - Versi Stabil Premium (Anti-Error & Aman V8 Serialization)
 * Penambahan Fitur: Edit Transaksi, Hapus Transaksi, dan Hapus Agen (Real-time Balance Adjustment)
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('AGEN HAKIM CELL - Sistem Laporan Penjualan')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Penolong konversi objek Tanggal (Date) ke format string ISO sebelum dikirim ke Client.
 * Ini untuk mencegah kegagalan serialisasi Google V8 yang menyebabkan "Error tak dikenal".
 */
function safeDate(val) {
  if (!val) return "";
  if (Object.prototype.toString.call(val) === '[object Date]') {
    try {
      return val.toISOString();
    } catch (e) {
      return val.toString();
    }
  }
  // Coba konversi jika berupa string tanggal biasa
  if (typeof val === 'string' && val.trim() !== '') {
    const d = new Date(val);
    if (!isNaN(d.getTime())) {
      return d.toISOString();
    }
  }
  return val.toString();
}

/**
 * 1. SISTEM SELF-HEALING DATABASE
 */
function initDatabase() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error("Spreadsheet tidak terdeteksi! Pastikan script ini dibuat melalui menu 'Ekstensi' > 'Apps Script' di dalam Google Sheets Anda (Container-bound).");
    }
    
    const sheetsConfig = {
      'Admins': ['Nama Pemilik', 'Nama Toko', 'Username', 'Password', 'Tanggal Daftar'],
      'Agents': ['Nama Agen', 'Nama Toko', 'Username', 'Password', 'Saldo Modal', 'Dibuat Oleh', 'Tanggal Daftar'],
      'Sales': ['ID Transaksi', 'Username Agen', 'Nama Produk', 'Modal', 'Harga Jual', 'Keuntungan', 'Tanggal Transaksi']
    };
    
    for (let sheetName in sheetsConfig) {
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.appendRow(sheetsConfig[sheetName]);
        sheet.getRange(1, 1, 1, sheetsConfig[sheetName].length).setFontWeight("bold");
      } else {
        let lastCol = sheet.getLastColumn();
        let headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
        let existingHeaders = headers.map(h => h.toString().trim().toLowerCase());
        let expectedHeaders = sheetsConfig[sheetName];
        
        expectedHeaders.forEach(header => {
          if (existingHeaders.indexOf(header.toLowerCase()) === -1) {
            let nextCol = sheet.getLastColumn() + 1;
            sheet.getRange(1, nextCol).setValue(header).setFontWeight("bold");
          }
        });
      }
    }
    return { success: true, message: "Database AGEN HAKIM CELL berhasil diinisialisasi & diperbaiki!" };
  } catch (e) {
    throw new Error("Gagal menginisialisasi database: " + e.message);
  }
}

/**
 * Helper untuk Dynamic Header Mapping
 * Mengembalikan objek mapping indeks kolom berbasis 1: { "nama header": index_kolom }
 */
function getHeaderMap(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol <= 0) return {};
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((header, index) => {
    if (header) {
      map[header.toString().trim().toLowerCase()] = index + 1;
    }
  });
  return map;
}

/**
 * Memeriksa apakah username sudah terdaftar (Case-Insensitive)
 */
function isUsernameTaken(username) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return false;
  username = username.toString().trim().toLowerCase();
  
  // Periksa sheet Admins
  const adminSheet = ss.getSheetByName('Admins');
  if (adminSheet && adminSheet.getLastRow() > 1) {
    const adminMap = getHeaderMap(adminSheet);
    const userCol = adminMap['username'];
    if (userCol) {
      const adminData = adminSheet.getRange(2, 1, adminSheet.getLastRow() - 1, adminSheet.getLastColumn()).getValues();
      const userColIdx = userCol - 1;
      for (let i = 0; i < adminData.length; i++) {
        if (adminData[i][userColIdx] && adminData[i][userColIdx].toString().trim().toLowerCase() === username) return true;
      }
    }
  }
  
  // Periksa sheet Agents
  const agentSheet = ss.getSheetByName('Agents');
  if (agentSheet && agentSheet.getLastRow() > 1) {
    const agentMap = getHeaderMap(agentSheet);
    const userCol = agentMap['username'];
    if (userCol) {
      const agentData = agentSheet.getRange(2, 1, agentSheet.getLastRow() - 1, agentSheet.getLastColumn()).getValues();
      const userColIdx = userCol - 1;
      for (let i = 0; i < agentData.length; i++) {
        if (agentData[i][userColIdx] && agentData[i][userColIdx].toString().trim().toLowerCase() === username) return true;
      }
    }
  }
  
  return false;
}

/**
 * 2. REGISTRASI ADMIN MANDIRI
 */
function registerAdmin(namaPemilik, namaToko, username, password) {
  initDatabase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Admins');
  if (!sheet) throw new Error("Database Admin tidak ditemukan!");
  
  username = username.toString().trim();
  if (isUsernameTaken(username)) {
    return { success: false, message: "Username '" + username + "' sudah terdaftar di sistem!" };
  }
  
  const map = getHeaderMap(sheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = [];
  
  headers.forEach(h => {
    const key = h.toString().trim().toLowerCase();
    if (key === 'nama pemilik') row.push(namaPemilik);
    else if (key === 'nama toko') row.push(namaToko);
    else if (key === 'username') row.push(username);
    else if (key === 'password') row.push(password);
    else if (key === 'tanggal daftar') row.push(new Date());
    else row.push('');
  });
  
  sheet.appendRow(row);
  return { success: true, message: "Pendaftaran Admin berhasil! Silakan login." };
}

/**
 * 2. SISTEM LOGIN MULTI-USER
 */
function loginUser(username, password) {
  initDatabase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  username = username.toString().trim().toLowerCase();
  password = password.toString();
  
  // 1. Cek Admin
  const adminSheet = ss.getSheetByName('Admins');
  if (adminSheet && adminSheet.getLastRow() > 1) {
    const adminMap = getHeaderMap(adminSheet);
    const userCol = adminMap['username'];
    const passCol = adminMap['password'];
    const namaCol = adminMap['nama pemilik'];
    const tokoCol = adminMap['nama toko'];
    
    if (userCol && passCol) {
      const data = adminSheet.getRange(2, 1, adminSheet.getLastRow() - 1, adminSheet.getLastColumn()).getValues();
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const u = row[userCol - 1] ? row[userCol - 1].toString().trim().toLowerCase() : '';
        const p = row[passCol - 1] ? row[passCol - 1].toString() : '';
        if (u === username && p === password) {
          return {
            success: true,
            role: 'admin',
            userData: {
              namaPemilik: namaCol ? row[namaCol - 1].toString() : username,
              namaToko: tokoCol ? row[tokoCol - 1].toString() : 'AGEN HAKIM CELL',
              username: u
            }
          };
        }
      }
    }
  }
  
  // 2. Cek Agen
  const agentSheet = ss.getSheetByName('Agents');
  if (agentSheet && agentSheet.getLastRow() > 1) {
    const agentMap = getHeaderMap(agentSheet);
    const userCol = agentMap['username'];
    const passCol = agentMap['password'];
    const namaCol = agentMap['nama agen'];
    const tokoCol = agentMap['nama toko'];
    const saldoCol = agentMap['saldo modal'];
    
    if (userCol && passCol) {
      const data = agentSheet.getRange(2, 1, agentSheet.getLastRow() - 1, agentSheet.getLastColumn()).getValues();
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const u = row[userCol - 1] ? row[userCol - 1].toString().trim().toLowerCase() : '';
        const p = row[passCol - 1] ? row[passCol - 1].toString() : '';
        if (u === username && p === password) {
          return {
            success: true,
            role: 'agent',
            userData: {
              namaAgen: namaCol ? row[namaCol - 1].toString() : username,
              namaToko: tokoCol ? row[tokoCol - 1].toString() : 'Kemitraan',
              username: u,
              saldoModal: saldoCol ? parseFloat(row[saldoCol - 1] || 0) : 0
            }
          };
        }
      }
    }
  }
  
  return { success: false, message: "Username atau password salah!" };
}

/**
 * 3. PENDAFTARAN AGEN BARU OLEH ADMIN
 */
function registerAgent(namaAgen, namaToko, username, password, adminUsername) {
  initDatabase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Agents');
  if (!sheet) throw new Error("Database Agen tidak ditemukan!");
  
  username = username.toString().trim();
  if (isUsernameTaken(username)) {
    return { success: false, message: "Username '" + username + "' sudah digunakan!" };
  }
  
  const map = getHeaderMap(sheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = [];
  
  headers.forEach(h => {
    const key = h.toString().trim().toLowerCase();
    if (key === 'nama agen') row.push(namaAgen);
    else if (key === 'nama toko') row.push(namaToko);
    else if (key === 'username') row.push(username);
    else if (key === 'password') row.push(password);
    else if (key === 'saldo modal') row.push(0);
    else if (key === 'dibuat oleh') row.push(adminUsername);
    else if (key === 'tanggal daftar') row.push(new Date());
    else row.push('');
  });
  
  sheet.appendRow(row);
  return { success: true, message: "Agen '" + namaAgen + "' berhasil didaftarkan!" };
}

/**
 * 3. HAPUS AGEN OLEH ADMIN (Permanen)
 */
function deleteAgent(agentUsername) {
  initDatabase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Agents');
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: false, message: "Data agen kosong!" };
  }
  
  const map = getHeaderMap(sheet);
  const userCol = map['username'];
  if (!userCol) return { success: false, message: "Kolom Username tidak ditemukan!" };
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const userColIdx = userCol - 1;
  
  for (let i = 0; i < data.length; i++) {
    const u = data[i][userColIdx] ? data[i][userColIdx].toString().trim().toLowerCase() : '';
    if (u === agentUsername.toString().trim().toLowerCase()) {
      sheet.deleteRow(i + 2);
      return { success: true, message: "Akun Mitra Agen '" + agentUsername + "' berhasil dihapus secara permanen!" };
    }
  }
  return { success: false, message: "Agen tidak ditemukan!" };
}

/**
 * 3. ATUR SALDO MODAL AGEN OLEH ADMIN
 */
function updateAgentBalance(agentUsername, newBalance) {
  initDatabase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Agents');
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: false, message: "Database Agen kosong atau belum ada!" };
  }
  
  const map = getHeaderMap(sheet);
  const userCol = map['username'];
  const saldoCol = map['saldo modal'];
  
  if (!userCol || !saldoCol) {
    return { success: false, message: "Struktur kolom database Agen tidak valid!" };
  }
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const userColIdx = userCol - 1;
  
  for (let i = 0; i < data.length; i++) {
    const u = data[i][userColIdx] ? data[i][userColIdx].toString().trim().toLowerCase() : '';
    if (u === agentUsername.toString().trim().toLowerCase()) {
      const rowNum = i + 2; 
      sheet.getRange(rowNum, saldoCol).setValue(parseFloat(newBalance));
      return { success: true, message: "Saldo Agen '" + agentUsername + "' sukses diubah menjadi Rp " + parseFloat(newBalance).toLocaleString('id-ID') };
    }
  }
  return { success: false, message: "Agen tidak ditemukan!" };
}

/**
 * 4. PENGIRIMAN LAPORAN PENJUALAN AGEN & VALIDASI SALDO
 */
function submitSale(agentUsername, namaProduk, modal, jual) {
  initDatabase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const agentSheet = ss.getSheetByName('Agents');
  if (!agentSheet || agentSheet.getLastRow() < 2) {
    return { success: false, message: "Database Agen tidak dapat diakses!" };
  }
  
  const agentMap = getHeaderMap(agentSheet);
  const uCol = agentMap['username'];
  const saldoCol = agentMap['saldo modal'];
  
  if (!uCol || !saldoCol) {
    return { success: false, message: "Struktur kolom database Agen tidak lengkap!" };
  }
  
  const agentData = agentSheet.getRange(2, 1, agentSheet.getLastRow() - 1, agentSheet.getLastColumn()).getValues();
  const userColIdx = uCol - 1;
  const saldoColIdx = saldoCol - 1;
  
  let agentRowIndex = -1;
  let currentSaldo = 0;
  
  for (let i = 0; i < agentData.length; i++) {
    const u = agentData[i][userColIdx] ? agentData[i][userColIdx].toString().trim().toLowerCase() : '';
    if (u === agentUsername.toString().trim().toLowerCase()) {
      agentRowIndex = i + 2;
      currentSaldo = parseFloat(agentData[i][saldoColIdx] || 0);
      break;
    }
  }
  
  if (agentRowIndex === -1) {
    return { success: false, message: "Autentikasi Agen tidak valid!" };
  }
  
  modal = parseFloat(modal);
  jual = parseFloat(jual);
  
  if (modal > currentSaldo) {
    return { 
      success: false, 
      message: "DITOLAK: Saldo modal tidak mencukupi!\nSisa Saldo Anda: Rp " + currentSaldo.toLocaleString('id-ID') + "\nHarga Modal Transaksi: Rp " + modal.toLocaleString('id-ID')
    };
  }
  
  const newSaldo = currentSaldo - modal;
  agentSheet.getRange(agentRowIndex, saldoCol).setValue(newSaldo);
  
  const salesSheet = ss.getSheetByName('Sales');
  if (!salesSheet) throw new Error("Database Sales tidak ditemukan!");
  
  const headers = salesSheet.getRange(1, 1, 1, salesSheet.getLastColumn()).getValues()[0];
  const row = [];
  const txId = "TX-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000);
  const keuntungan = jual - modal;
  
  headers.forEach(h => {
    const key = h.toString().trim().toLowerCase();
    if (key === 'id transaksi') row.push(txId);
    else if (key === 'username agen') row.push(agentUsername);
    else if (key === 'nama produk') row.push(namaProduk);
    else if (key === 'modal') row.push(modal);
    else if (key === 'harga jual') row.push(jual);
    else if (key === 'keuntungan') row.push(keuntungan);
    else if (key === 'tanggal transaksi') row.push(new Date());
    else row.push('');
  });
  
  salesSheet.appendRow(row);
  
  return {
    success: true,
    message: "Laporan penjualan berhasil dikirim dan tercatat!",
    updatedSaldo: newSaldo
  };
}

/**
 * 4. EDIT LAPORAN TRANSAKSI (Sisi Agen & Admin)
 * Secara otomatis mengecek kecukupan saldo jika nominal modal diubah.
 */
function editSale(idTransaksi, namaProduk, newModal, newJual) {
  initDatabase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName('Sales');
  if (!salesSheet || salesSheet.getLastRow() < 2) return { success: false, message: "Transaksi tidak ditemukan!" };

  const salesMap = getHeaderMap(salesSheet);
  const idColIdx = salesMap['id transaksi'] - 1;
  const userColIdx = salesMap['username agen'] - 1;
  const modalColIdx = salesMap['modal'] - 1;
  const jualColIdx = salesMap['harga jual'] - 1;
  const untungColIdx = salesMap['keuntungan'] - 1;
  const prodColIdx = salesMap['nama produk'] - 1;

  const salesData = salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, salesSheet.getLastColumn()).getValues();
  let saleRowIndex = -1;
  let agentUsername = '';
  let oldModal = 0;

  for (let i = 0; i < salesData.length; i++) {
    if (salesData[i][idColIdx].toString().trim() === idTransaksi.toString().trim()) {
      saleRowIndex = i + 2;
      agentUsername = salesData[i][userColIdx].toString().trim();
      oldModal = parseFloat(salesData[i][modalColIdx] || 0);
      break;
    }
  }

  if (saleRowIndex === -1) return { success: false, message: "Transaksi tidak ditemukan!" };

  newModal = parseFloat(newModal);
  newJual = parseFloat(newJual);
  const diff = newModal - oldModal; // Selisih pemotongan saldo baru

  // Ambil saldo agen saat ini
  const agentSheet = ss.getSheetByName('Agents');
  if (!agentSheet || agentSheet.getLastRow() < 2) return { success: false, message: "Database agen tidak ditemukan!" };
  const agentMap = getHeaderMap(agentSheet);
  const agentUserColIdx = agentMap['username'] - 1;
  const agentSaldoColIdx = agentMap['saldo modal'] - 1;

  const agentData = agentSheet.getRange(2, 1, agentSheet.getLastRow() - 1, agentSheet.getLastColumn()).getValues();
  let agentRowIndex = -1;
  let currentSaldo = 0;

  for (let j = 0; j < agentData.length; j++) {
    if (agentData[j][agentUserColIdx].toString().trim().toLowerCase() === agentUsername.toLowerCase()) {
      agentRowIndex = j + 2;
      currentSaldo = parseFloat(agentData[j][agentSaldoColIdx] || 0);
      break;
    }
  }

  if (agentRowIndex === -1) return { success: false, message: "Mitra agen pemilik transaksi tidak ditemukan!" };

  // Validasi kecukupan saldo jika ada penambahan modal
  if (diff > currentSaldo) {
    return {
      success: false,
      message: "Edit Gagal: Sisa saldo modal agen tidak mencukupi!\nSisa Saldo: Rp " + currentSaldo.toLocaleString('id-ID') + "\nButuh penambahan saldo sebesar: Rp " + (diff - currentSaldo).toLocaleString('id-ID')
    };
  }

  // Update Saldo Kerja Agen
  const newSaldo = currentSaldo - diff;
  agentSheet.getRange(agentRowIndex, agentSaldoColIdx + 1).setValue(newSaldo);

  // Update Data Penjualan
  salesSheet.getRange(saleRowIndex, prodColIdx + 1).setValue(namaProduk);
  salesSheet.getRange(saleRowIndex, modalColIdx + 1).setValue(newModal);
  salesSheet.getRange(saleRowIndex, jualColIdx + 1).setValue(newJual);
  salesSheet.getRange(saleRowIndex, untungColIdx + 1).setValue(newJual - newModal);

  return { success: true, message: "Transaksi '" + idTransaksi + "' berhasil diperbarui!", updatedSaldo: newSaldo };
}

/**
 * 4. HAPUS LAPORAN TRANSAKSI (Sisi Agen & Admin)
 * Melakukan pengembalian saldo modal (Refund) secara penuh ke akun agen.
 */
function deleteSale(idTransaksi) {
  initDatabase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName('Sales');
  if (!salesSheet || salesSheet.getLastRow() < 2) return { success: false, message: "Transaksi tidak ditemukan!" };

  const salesMap = getHeaderMap(salesSheet);
  const idColIdx = salesMap['id transaksi'] - 1;
  const userColIdx = salesMap['username agen'] - 1;
  const modalColIdx = salesMap['modal'] - 1;

  const salesData = salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, salesSheet.getLastColumn()).getValues();
  let saleRowIndex = -1;
  let agentUsername = '';
  let oldModal = 0;

  for (let i = 0; i < salesData.length; i++) {
    if (salesData[i][idColIdx].toString().trim() === idTransaksi.toString().trim()) {
      saleRowIndex = i + 2;
      agentUsername = salesData[i][userColIdx].toString().trim();
      oldModal = parseFloat(salesData[i][modalColIdx] || 0);
      break;
    }
  }

  if (saleRowIndex === -1) return { success: false, message: "Transaksi tidak ditemukan!" };

  // Refund saldo agen
  const agentSheet = ss.getSheetByName('Agents');
  if (agentSheet && agentSheet.getLastRow() > 1) {
    const agentMap = getHeaderMap(agentSheet);
    const agentUserColIdx = agentMap['username'] - 1;
    const agentSaldoColIdx = agentMap['saldo modal'] - 1;

    const agentData = agentSheet.getRange(2, 1, agentSheet.getLastRow() - 1, agentSheet.getLastColumn()).getValues();
    for (let j = 0; j < agentData.length; j++) {
      if (agentData[j][agentUserColIdx].toString().trim().toLowerCase() === agentUsername.toLowerCase()) {
        const agentRowIndex = j + 2;
        const currentSaldo = parseFloat(agentData[j][agentSaldoColIdx] || 0);
        agentSheet.getRange(agentRowIndex, agentSaldoColIdx + 1).setValue(currentSaldo + oldModal);
        break;
      }
    }
  }

  // Hapus dari sheet Sales
  salesSheet.deleteRow(saleRowIndex);

  return { success: true, message: "Laporan transaksi berhasil dihapus dan saldo modal kerja agen telah di-refund!" };
}

/**
 * 5. GET DATA UNTUK DASHBOARD ADMIN
 * Seluruh tanggal dikonversi secara protektif via `safeDate()` untuk mencegah error Google Apps Script boundary parsing.
 */
function getAdminDashboardData(adminUsername) {
  initDatabase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const agents = [];
  const agentSheet = ss.getSheetByName('Agents');
  if (agentSheet && agentSheet.getLastRow() > 1) {
    const map = getHeaderMap(agentSheet);
    const uCol = map['username'];
    const creatorCol = map['dibuat oleh'];
    const namaCol = map['nama agen'];
    const tokoCol = map['nama toko'];
    const saldoCol = map['saldo modal'];
    const tglCol = map['tanggal daftar'];
    
    if (uCol && creatorCol) {
      const data = agentSheet.getRange(2, 1, agentSheet.getLastRow() - 1, agentSheet.getLastColumn()).getValues();
      data.forEach(row => {
        const creatorVal = row[creatorCol - 1] ? row[creatorCol - 1].toString().trim().toLowerCase() : '';
        if (creatorVal === adminUsername.toString().trim().toLowerCase()) {
          agents.push({
            namaAgen: namaCol && row[namaCol - 1] ? row[namaCol - 1].toString() : '',
            namaToko: tokoCol && row[tokoCol - 1] ? row[tokoCol - 1].toString() : '',
            username: row[uCol - 1] ? row[uCol - 1].toString() : '',
            saldoModal: saldoCol ? parseFloat(row[saldoCol - 1] || 0) : 0,
            tanggalDaftar: tglCol ? safeDate(row[tglCol - 1]) : ''
          });
        }
      });
    }
  }
  
  const agentUsernames = agents.map(a => a.username.toLowerCase());
  const sales = [];
  const salesSheet = ss.getSheetByName('Sales');
  if (salesSheet && salesSheet.getLastRow() > 1) {
    const map = getHeaderMap(salesSheet);
    const uCol = map['username agen'];
    const idCol = map['id transaksi'];
    const prodCol = map['nama produk'];
    const modalCol = map['modal'];
    const jualCol = map['harga jual'];
    const untungCol = map['keuntungan'];
    const tglCol = map['tanggal transaksi'];
    
    if (uCol) {
      const data = salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, salesSheet.getLastColumn()).getValues();
      data.forEach(row => {
        const agentUser = row[uCol - 1] ? row[uCol - 1].toString().trim().toLowerCase() : '';
        if (agentUsernames.indexOf(agentUser) !== -1) {
          sales.push({
            idTransaksi: idCol && row[idCol - 1] ? row[idCol - 1].toString() : '',
            usernameAgen: agentUser,
            namaProduk: prodCol && row[prodCol - 1] ? row[prodCol - 1].toString() : '',
            modal: modalCol ? parseFloat(row[modalCol - 1] || 0) : 0,
            hargaJual: jualCol ? parseFloat(row[jualCol - 1] || 0) : 0,
            keuntungan: untungCol ? parseFloat(row[untungCol - 1] || 0) : 0,
            tanggalTransaksi: tglCol ? safeDate(row[tglCol - 1]) : ''
          });
        }
      });
    }
  }
  
  return {
    agents: agents,
    sales: sales
  };
}

/**
 * 4. GET DATA UNTUK DASHBOARD AGEN
 */
function getAgentDashboardData(agentUsername) {
  initDatabase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const agentSheet = ss.getSheetByName('Agents');
  let saldoModal = 0;
  let namaToko = '';
  let namaAgen = '';
  
  if (agentSheet && agentSheet.getLastRow() > 1) {
    const map = getHeaderMap(agentSheet);
    const uCol = map['username'];
    const saldoCol = map['saldo modal'];
    const tokoCol = map['nama toko'];
    const namaCol = map['nama agen'];
    
    if (uCol) {
      const data = agentSheet.getRange(2, 1, agentSheet.getLastRow() - 1, agentSheet.getLastColumn()).getValues();
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const u = row[uCol - 1] ? row[uCol - 1].toString().trim().toLowerCase() : '';
        if (u === agentUsername.toString().trim().toLowerCase()) {
          saldoModal = saldoCol ? parseFloat(row[saldoCol - 1] || 0) : 0;
          namaToko = tokoCol && row[tokoCol - 1] ? row[tokoCol - 1].toString() : '';
          namaAgen = namaCol && row[namaCol - 1] ? row[namaCol - 1].toString() : '';
          break;
        }
      }
    }
  }
  
  const sales = [];
  const salesSheet = ss.getSheetByName('Sales');
  if (salesSheet && salesSheet.getLastRow() > 1) {
    const map = getHeaderMap(salesSheet);
    const uCol = map['username agen'];
    const idCol = map['id transaksi'];
    const prodCol = map['nama produk'];
    const modalCol = map['modal'];
    const jualCol = map['harga jual'];
    const untungCol = map['keuntungan'];
    const tglCol = map['tanggal transaksi'];
    
    if (uCol) {
      const data = salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, salesSheet.getLastColumn()).getValues();
      data.forEach(row => {
        const agentUser = row[uCol - 1] ? row[uCol - 1].toString().trim().toLowerCase() : '';
        if (agentUser === agentUsername.toString().trim().toLowerCase()) {
          sales.push({
            idTransaksi: idCol && row[idCol - 1] ? row[idCol - 1].toString() : '',
            namaProduk: prodCol && row[prodCol - 1] ? row[prodCol - 1].toString() : '',
            modal: modalCol ? parseFloat(row[modalCol - 1] || 0) : 0,
            hargaJual: jualCol ? parseFloat(row[jualCol - 1] || 0) : 0,
            keuntungan: untungCol ? parseFloat(row[untungCol - 1] || 0) : 0,
            tanggalTransaksi: tglCol ? safeDate(row[tglCol - 1]) : ''
          });
        }
      });
    }
  }
  
  return {
    saldoModal: saldoModal,
    namaToko: namaToko,
    namaAgen: namaAgen,
    sales: sales
  };
}
