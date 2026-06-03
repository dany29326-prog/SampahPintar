// ========== KONFIGURASI ==========
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// ========== FUNGSI WEB APP (API) ==========
function doGet(e) {
  return handleCors(e);
}

function doPost(e) {
  return handleCors(e);
}

function handleCors(e) {
  const response = ContentService.createTextOutput();
  response.setMimeType(ContentService.MimeType.JSON);
  
  try {
    let params = {};
    if (e && e.parameter) {
      params = e.parameter;
    }
    // Untuk POST form data
    if (e && e.postData && e.postData.contents) {
      const postData = JSON.parse(e.postData.contents);
      params = postData;
    }
    
    const action = params.action;
    if (action) {
      const result = handleAction(params);
      response.setContent(JSON.stringify(result));
    } else {
      response.setContent(JSON.stringify({ status: 'success', message: 'SampahPintar API is running' }));
    }
  } catch(err) {
    response.setContent(JSON.stringify({ status: 'error', message: err.toString() }));
  }
  return response;
}

function handleAction(params) {
  const action = params.action;
  try {
    // Pastikan semua sheet ada
    ensureSheetsExist();
    
    switch(action) {
      case 'getPelanggan': return getPelanggan();
      case 'savePelanggan': return savePelanggan(params);
      case 'deletePelanggan': return deletePelanggan(params);
      case 'getPembayaran': return getPembayaran();
      case 'savePembayaran': return savePembayaran(params);
      case 'sync': return syncData(params);
      default: return { status: 'error', message: 'Unknown action' };
    }
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

// ========== PASTIKAN SHEET EXIST ==========
function ensureSheetsExist() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('pelanggan');
  if (!sheet) createPelangganSheet(ss);
  sheet = ss.getSheetByName('pembayaran');
  if (!sheet) createPembayaranSheet(ss);
  sheet = ss.getSheetByName('log_aktivitas');
  if (!sheet) createLogAktivitasSheet(ss);
  sheet = ss.getSheetByName('settings');
  if (!sheet) createSettingsSheet(ss);
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ========== PELANGGAN CRUD ==========
function getPelanggan() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('pelanggan');
  if (!sheet) return { status: 'error', message: 'Sheet pelanggan tidak ditemukan' };
  
  const data = sheet.getDataRange().getValues();
  const pelanggan = [];
  for (let i = 1; i < data.length; i++) {
    pelanggan.push({
      id: data[i][0],
      nama: data[i][1],
      alamat: data[i][2],
      hp: data[i][3],
      kategori: data[i][4],
      status: data[i][5]
    });
  }
  return { status: 'success', data: pelanggan };
}

function savePelanggan(params) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('pelanggan');
  if (!sheet) return { status: 'error', message: 'Sheet pelanggan tidak ditemukan' };
  
  let data = params.data;
  if (typeof data === 'string') data = JSON.parse(data);
  
  if (data.id && data.id !== 'null' && data.id !== '') {
    // Update atau Insert (Upsert)
    const ids = sheet.getRange('A:A').getValues();
    let found = false;
    for (let i = 1; i < ids.length; i++) {
      if (ids[i][0] == data.id) {
        sheet.getRange(i+1, 1, 1, 6).setValues([[
          data.id, data.nama, data.alamat, data.hp, data.kategori, data.status
        ]]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([data.id, data.nama, data.alamat, data.hp, data.kategori, data.status]);
    }
  } else {
    // Insert new
    data.id = generateId('PLG-');
    sheet.appendRow([data.id, data.nama, data.alamat, data.hp, data.kategori, data.status]);
  }
  
  logActivity('save_pelanggan', `Menyimpan pelanggan: ${data.nama}`);
  return { status: 'success', data: data };
}

function deletePelanggan(params) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('pelanggan');
  if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan' };
  
  const id = params.id;
  const ids = sheet.getRange('A:A').getValues();
  for (let i = 1; i < ids.length; i++) {
    if (ids[i][0] == id) {
      sheet.deleteRow(i+1);
      logActivity('delete_pelanggan', `Menghapus pelanggan ID: ${id}`);
      break;
    }
  }
  return { status: 'success' };
}


// ========== PEMBAYARAN CRUD ==========
function getPembayaran() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('pembayaran');
  if (!sheet) return { status: 'error', message: 'Sheet pembayaran tidak ditemukan' };
  
  const data = sheet.getDataRange().getValues();
  const pembayaran = [];
  for (let i = 1; i < data.length; i++) {
    pembayaran.push({
      id: data[i][0],
      nomor_kwitansi: data[i][1],
      tanggal_bayar: data[i][2],
      pelanggan_id: data[i][3],
      nama_pelanggan: data[i][4],
      kategori_pelanggan: data[i][5],
      periode_bayar: data[i][6],
      nominal: data[i][7],
      denda: data[i][8],
      total: data[i][9],
      status: data[i][10],
      metode_pembayaran: data[i][11],
      petugas: data[i][12],
      catatan: data[i][13]
    });
  }
  return { status: 'success', data: pembayaran };
}

function savePembayaran(params) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('pembayaran');
  if (!sheet) return { status: 'error', message: 'Sheet pembayaran tidak ditemukan' };
  
  let data = params.data;
  if (typeof data === 'string') data = JSON.parse(data);
  
  if (!data.id || data.id === 'null' || data.id === '') {
    data.id = generateId('TRX-');
  }
  if (!data.nomor_kwitansi) {
    data.nomor_kwitansi = generateId('KW-');
  }
  const ids = sheet.getRange('A:A').getValues();
  let found = false;
  for (let i = 1; i < ids.length; i++) {
    if (ids[i][0] == data.id) {
      sheet.getRange(i+1, 1, 1, 14).setValues([[
        data.id, data.nomor_kwitansi, data.tanggal_bayar, data.pelanggan_id, data.nama_pelanggan, data.kategori_pelanggan, data.periode_bayar, data.nominal, data.denda, data.total, data.status || 'LUNAS', data.metode_pembayaran, data.petugas, data.catatan
      ]]);
      found = true;
      break;
    }
  }
  
  if (!found) {
    if (!data.id || data.id === 'null' || data.id === '') data.id = generateId('TRX-');
    if (!data.nomor_kwitansi) data.nomor_kwitansi = generateId('KW-');
    sheet.appendRow([
        data.id, data.nomor_kwitansi, data.tanggal_bayar, data.pelanggan_id, data.nama_pelanggan, data.kategori_pelanggan, data.periode_bayar, data.nominal, data.denda, data.total, data.status || 'LUNAS', data.metode_pembayaran, data.petugas, data.catatan
    ]);
  }
  
  logActivity('save_pembayaran', `Pembayaran: ${data.nomor_kwitansi} - Rp ${data.total}`);
  return { status: 'success', data: data };
}

// ========== SINKRONISASI ==========
function syncData(params) {
  let data = params.data;
  if (typeof data === 'string') data = JSON.parse(data);
  let synced = 0;
  
  if (data.pelanggan && data.pelanggan.length > 0) {
    for (const pel of data.pelanggan) {
      try {
        savePelanggan({ data: JSON.stringify(pel) });
        synced++;
      } catch(e) {}
    }
  }
  if (data.pembayaran && data.pembayaran.length > 0) {
    for (const pay of data.pembayaran) {
      try {
        savePembayaran({ data: JSON.stringify(pay) });
        synced++;
      } catch(e) {}
    }
  }
  logActivity('sync', `Sinkronisasi ${synced} data dari client`);
  return { status: 'success', synced: synced };
}

// ========== LOG ACTIVITY ==========
function logActivity(aksi, detail) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('log_aktivitas');
    if (!sheet) {
      sheet = ss.insertSheet('log_aktivitas');
      sheet.getRange(1,1,1,5).setValues([['timestamp','aksi','detail','user','jenis']]);
    }
    sheet.appendRow([new Date().toISOString(), aksi, detail, Session.getActiveUser().getEmail(), 'web']);
  } catch(e) {}
}

function generateId(prefix) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let result = prefix;
  for (let i = 0; i < 5; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

// ========== CREATE SHEETS (JIKA BELUM ADA) ==========
function createPelangganSheet(ss) {
  let sheet = ss.insertSheet('pelanggan');
  const headers = ['id','nama','alamat','hp','kategori','status','tanggal_daftar'];
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#667eea').setFontColor('#ffffff');
  sheet.setColumnWidths(1, 150); sheet.setColumnWidths(2, 200); sheet.setColumnWidths(3, 250);
  sheet.setColumnWidths(4, 150); sheet.setColumnWidths(5, 100); sheet.setColumnWidths(6, 100); sheet.setColumnWidths(7, 150);
  
  // Data validation
  const kategoriRule = SpreadsheetApp.newDataValidation().requireValueInList(['RT','Toko','Kantor'], true).build();
  sheet.getRange('E2:E').setDataValidation(kategoriRule);
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['AKTIF','NONAKTIF'], true).build();
  sheet.getRange('F2:F').setDataValidation(statusRule);
}


function createPembayaranSheet(ss) {
  let sheet = ss.insertSheet('pembayaran');
  const headers = ['id','nomor_kwitansi','tanggal_bayar','pelanggan_id','nama_pelanggan','kategori_pelanggan','periode_bayar','nominal','denda','total','status','metode_pembayaran','petugas','catatan'];
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#667eea').setFontColor('#ffffff');
  sheet.getRange('H:J').setNumberFormat('Rp #,##0');
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['LUNAS','BELUM LUNAS'], true).build();
  sheet.getRange('K2:K').setDataValidation(statusRule);
}

function createLogAktivitasSheet(ss) {
  let sheet = ss.insertSheet('log_aktivitas');
  sheet.getRange(1,1,1,5).setValues([['timestamp','aksi','detail','user','jenis']]);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,5).setFontWeight('bold').setBackground('#667eea').setFontColor('#ffffff');
}

function createSettingsSheet(ss) {
  let sheet = ss.insertSheet('settings');
  sheet.getRange(1,1,1,3).setValues([['key','value','keterangan']]);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,3).setFontWeight('bold').setBackground('#667eea').setFontColor('#ffffff');
  sheet.getRange(2,1,4,3).setValues([
    ['denda_persen','10','Denda dalam persen'],
    ['admin_name','SampahPintar','Nama admin'],
    ['company_name','SampahPintar System','Nama perusahaan'],
    ['wa_template','Terima kasih {nama}','Template WA']
  ]);
}

// ========== FUNGSI MANUAL UNTUK MENJALANKAN SETUP ==========
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('📊 SampahPintar');
  menu.addItem('🔄 Setup Ulang Sheet', 'setupCompleteSystem');
  menu.addItem('🗑️ Reset Semua Data', 'resetAllData');
  menu.addToUi();
}

function setupCompleteSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ['pelanggan','pembayaran','log_aktivitas','settings'];
  sheets.forEach(name => { const s = ss.getSheetByName(name); if(s) ss.deleteSheet(s); });
  createPelangganSheet(ss);
  createPembayaranSheet(ss);
  createLogAktivitasSheet(ss);
  createSettingsSheet(ss);
  SpreadsheetApp.getUi().alert('✅ Setup Selesai! Semua sheet telah dibuat.');
}

function resetAllData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('⚠️ Peringatan', 'Hapus semua data?', ui.ButtonSet.YES_NO);
  if (response === ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ['pelanggan','pembayaran','log_aktivitas'].forEach(name => {
      const sheet = ss.getSheetByName(name);
      if(sheet && sheet.getLastRow()>1) sheet.deleteRows(2, sheet.getLastRow()-1);
    });
    ui.alert('✅ Data direset');
  }
}