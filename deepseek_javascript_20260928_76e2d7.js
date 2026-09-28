/**
 * ============================================================
 * DISPENDIK RUN 2026 - Backend (v9 Final)
 * Auto Setup + LockService + Cache + Multi-CS
 * ============================================================
 */

// ==================== KONFIGURASI ====================
const APP_NAME = 'DISPENDIK RUN 2026';
const FIXED_SPREADSHEET_ID = '1mVhP_okPc97bwMxwvL6Gs2u6jZrxdROm31OUvXaqBgI';
const DRIVE_FOLDER_NAME = 'DISPENDIK_RUN_2026_PROOFS';

const SHEET_PARTICIPANTS = 'Participants';
const SHEET_SETTINGS = 'Settings';
const SHEET_CATEGORIES = 'Categories';
const SHEET_JERSEYS = 'Jerseys';
const SHEET_CS = 'CustomerService';
const SHEET_LOG = 'Log';

const DEFAULT_ADMIN_USER = 'admin';
const DEFAULT_ADMIN_PASS = 'admin2026';
const DEFAULT_RESET_PASS = 'dispendik';

const DEFAULT_CATEGORIES = [
  ['5K', '5K Fun Run', 50000, 250],
  ['10K', '10K Race', 75000, 250]
];

const DEFAULT_JERSEYS = [
  ['S', 'Small', 1],
  ['M', 'Medium', 2],
  ['L', 'Large', 3],
  ['XL', 'Extra Large', 4],
  ['XXL', 'Double Extra Large', 5]
];

const DEFAULT_CS = [
  ['Admin Registration', '6281234567890', 1],
  ['Admin Payment', '6281234567891', 2]
];

const CACHE_SEC = 20;
const CACHE_KEY_SETTINGS = 'dsp_settings';
const CACHE_KEY_CATEGORIES = 'dsp_categories';
const CACHE_KEY_JERSEYS = 'dsp_jerseys';
const CACHE_KEY_CS = 'dsp_cs';
const CACHE_KEY_STATS = 'dsp_stats';

// ==================== LOCK HELPER ====================
function withLock(fn) {
  const lock = LockService.getScriptLock();
  const start = Date.now();
  while (true) {
    try {
      if (lock.tryLock(5000)) {
        try { return fn(); } finally { lock.releaseLock(); }
      }
    } catch (e) { Logger.log('Lock err: ' + e); }
    if (Date.now() - start > 15000) {
      throw new Error('Sistem sedang sibuk. Silakan coba lagi.');
    }
    Utilities.sleep(200);
  }
}

function clearCache(keys) {
  try {
    const cache = CacheService.getScriptCache();
    if (keys === 'all') {
      cache.removeAll([CACHE_KEY_SETTINGS, CACHE_KEY_CATEGORIES, CACHE_KEY_JERSEYS, CACHE_KEY_CS, CACHE_KEY_STATS]);
    } else if (Array.isArray(keys)) {
      cache.removeAll(keys);
    } else if (keys) {
      cache.remove(keys);
    }
  } catch (e) { Logger.log('Cache err: ' + e); }
}

// ==================== SPREADSHEET & FOLDER ====================
function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID') || FIXED_SPREADSHEET_ID;

  try {
    const ss = SpreadsheetApp.openById(ssId);
    props.setProperty('SPREADSHEET_ID', ssId);
    return ss;
  } catch (e) {
    // Jika ID fixed error, coba buat baru
    props.deleteProperty('SPREADSHEET_ID');
    const ss = SpreadsheetApp.create(APP_NAME + '_DATABASE');
    props.setProperty('SPREADSHEET_ID', ss.getId());
    initSheets(ss);
    return ss;
  }
}

function getFolder() {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty('DRIVE_FOLDER_ID');

  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (e) { props.deleteProperty('DRIVE_FOLDER_ID'); }
  }

  const existing = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  let folder = existing.hasNext() ? existing.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
  props.setProperty('DRIVE_FOLDER_ID', folder.getId());
  return folder;
}

function getSheet(name) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    initSheet(name, sheet);
  }
  return sheet;
}

function initSheet(name, sheet) {
  sheet.clear();
  let headers = [];
  if (name === SHEET_PARTICIPANTS) {
    headers = ['ID', 'Nama', 'NoHP', 'Email', 'Password', 'JenisKelamin', 'Kategori', 'Jersey', 'BIB', 'Status', 'ProofURL', 'TanggalDaftar', 'TanggalVerifikasi', 'Harga'];
  } else if (name === SHEET_SETTINGS) {
    headers = ['Key', 'Value'];
  } else if (name === SHEET_CATEGORIES) {
    headers = ['Kode', 'Nama', 'Harga', 'Kuota'];
  } else if (name === SHEET_JERSEYS) {
    headers = ['Kode', 'Label', 'Urutan'];
  } else if (name === SHEET_CS) {
    headers = ['Nama', 'NomorWA', 'Urutan'];
  } else if (name === SHEET_LOG) {
    headers = ['Timestamp', 'Action', 'User', 'Detail'];
  }

  if (headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff').setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }

  // Default data
  if (name === SHEET_SETTINGS) {
    const defaults = [
      ['title', APP_NAME],
      ['date', '24 MEI 2026 - ALUN-ALUN DISPENDIK'],
      ['bank', 'Bank Jatim: 1234567890 a.n. DISPENDIK RUN EVENT'],
      ['qris', 'https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=DISPENDIK-RUN-2026-QRIS-PAYMENT'],
      ['adminUser', DEFAULT_ADMIN_USER],
      ['adminPass', DEFAULT_ADMIN_PASS]
    ];
    sheet.getRange(2, 1, defaults.length, 2).setValues(defaults);
  }
  if (name === SHEET_CATEGORIES) {
    sheet.getRange(2, 1, DEFAULT_CATEGORIES.length, 4).setValues(DEFAULT_CATEGORIES);
  }
  if (name === SHEET_JERSEYS) {
    sheet.getRange(2, 1, DEFAULT_JERSEYS.length, 3).setValues(DEFAULT_JERSEYS);
  }
  if (name === SHEET_CS) {
    sheet.getRange(2, 1, DEFAULT_CS.length, 3).setValues(DEFAULT_CS);
  }
}

function initSheets(ss) {
  const needed = [SHEET_PARTICIPANTS, SHEET_SETTINGS, SHEET_CATEGORIES, SHEET_JERSEYS, SHEET_CS, SHEET_LOG];
  needed.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      initSheet(name, sheet);
    }
  });
  // Hapus Sheet default "Sheet1" jika ada
  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1) ss.deleteSheet(sheet1);
}

// ==================== SETUP APLIKASI ====================
function setupApplication() {
  Logger.log('🚀 SETUP ' + APP_NAME);
  const ss = getSpreadsheet();
  initSheets(ss);
  const folder = getFolder();

  Logger.log('✅ SETUP SELESAI');
  Logger.log('📊 Spreadsheet: ' + ss.getUrl());
  Logger.log('📁 Folder: ' + folder.getUrl());
  Logger.log('🔑 Admin: ' + DEFAULT_ADMIN_USER + ' / ' + DEFAULT_ADMIN_PASS);

  return {
    status: 'success',
    spreadsheetUrl: ss.getUrl(),
    folderUrl: folder.getUrl()
  };
}

function checkSetupStatus() {
  try {
    const ss = getSpreadsheet();
    const folder = getFolder();
    return {
      setup: true,
      spreadsheetUrl: ss.getUrl(),
      folderUrl: folder.getUrl()
    };
  } catch (e) {
    return { setup: false, message: e.toString() };
  }
}

function clearAllCache() {
  clearCache('all');
  Logger.log('✅ Cache dibersihkan');
}

// ==================== WEB APP ====================
function doGet(e) {
  try {
    const ss = getSpreadsheet();
    initSheets(ss);
  } catch (err) {
    Logger.log('doGet setup err: ' + err);
  }

  const template = HtmlService.createTemplateFromFile('Index');
  template.gasUrl = ScriptApp.getService().getUrl();
  return template.evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    const data = params.data || {};
    let result;

    switch (action) {
      case 'register':          result = registerParticipant(data); break;
      case 'login':             result = loginUser(data); break;
      case 'uploadProof':       result = uploadProof(data); break;
      case 'getSettings':       result = getSettings(); break;
      case 'saveSettings':      result = saveSettings(data); break;
      case 'getCategories':     result = getCategories(); break;
      case 'saveCategories':    result = saveCategories(data); break;
      case 'getJerseys':        result = getJerseys(); break;
      case 'saveJerseys':       result = saveJerseys(data); break;
      case 'getCs':             result = getCs(); break;
      case 'saveCs':            result = saveCs(data); break;
      case 'getParticipants':   result = getParticipants(); break;
      case 'verifyParticipant': result = verifyParticipant(data); break;
      case 'resetPassword':     result = resetPassword(data); break;
      case 'checkStatus':       result = checkStatus(data); break;
      case 'getPublicStats':    result = getPublicStats(); break;
      case 'getSetupStatus':    result = checkSetupStatus(); break;
      default:
        result = { status: 'error', message: 'Action tidak dikenal: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('doPost err: ' + err);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==================== HELPER ====================
function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(row => row[0] !== '' && row[0] !== null)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function logAction(action, user, detail) {
  try {
    getSheet(SHEET_LOG).appendRow([new Date().toLocaleString('id-ID'), action, user || '-', detail || '-']);
  } catch (e) { Logger.log('log err: ' + e); }
}

function uploadToDrive(base64Data, fileName, mimeType) {
  const folder = getFolder();
  const base64Clean = base64Data.indexOf(',') > -1 ? base64Data.split(',')[1] : base64Data;
  const bytes = Utilities.base64Decode(base64Clean);
  const blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg', fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {
    id: file.getId(),
    url: 'https://drive.google.com/uc?export=view&id=' + file.getId(),
    preview: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000',
    view: 'https://drive.google.com/file/d/' + file.getId() + '/view',
    download: 'https://drive.google.com/uc?export=download&id=' + file.getId()
  };
}

function mapParticipant(user) {
  return {
    id: user.ID,
    name: user.Nama,
    email: user.Email,
    phone: user.NoHP,
    gender: user.JenisKelamin || '-',
    category: user.Kategori,
    jersey: user.Jersey,
    bib: user.BIB,
    status: user.Status,
    proof: user.ProofURL || '',
    price: Number(user.Harga) || 0,
    date: user.TanggalDaftar
  };
}

// ==================== KATEGORI ====================
function getCategories(noCache) {
  try {
    if (!noCache) {
      const c = CacheService.getScriptCache().get(CACHE_KEY_CATEGORIES);
      if (c) return { status: 'success', categories: JSON.parse(c) };
    }

    const sheet = getSheet(SHEET_CATEGORIES);
    let rows = sheet.getDataRange().getValues();
    if (rows.length < 2) { initSheet(SHEET_CATEGORIES, sheet); rows = sheet.getDataRange().getValues(); }

    const categories = rows.slice(1).filter(r => r[0]).map(r => ({
      code: String(r[0]), name: String(r[1]),
      price: Number(r[2]) || 0, quota: Number(r[3]) || 0
    }));

    CacheService.getScriptCache().put(CACHE_KEY_CATEGORIES, JSON.stringify(categories), CACHE_SEC);
    return { status: 'success', categories };
  } catch (err) {
    return { status: 'error', message: err.message, categories: [] };
  }
}

function saveCategories(data) {
  return withLock(() => {
    try {
      const cats = data.categories || [];
      if (cats.length === 0) return { status: 'error', message: 'Minimal 1 kategori.' };

      for (const c of cats) {
        if (!c.code || !c.name) return { status: 'error', message: 'Kode & nama wajib diisi.' };
        if (Number(c.price) < 0 || Number(c.quota) < 0) return { status: 'error', message: 'Harga & kuota tidak boleh negatif.' };
      }

      const codes = cats.map(c => String(c.code).toUpperCase());
      if (new Set(codes).size !== codes.length) return { status: 'error', message: 'Kode tidak boleh duplikat.' };

      const sheet = getSheet(SHEET_CATEGORIES);
      sheet.clear();
      sheet.getRange(1, 1, 1, 4).setValues([['Kode', 'Nama', 'Harga', 'Kuota']])
        .setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff').setHorizontalAlignment('center');
      sheet.getRange(2, 1, cats.length, 4).setValues(cats.map(c => [String(c.code).toUpperCase(), String(c.name), Number(c.price), Number(c.quota)]));
      sheet.setFrozenRows(1);

      clearCache([CACHE_KEY_CATEGORIES, CACHE_KEY_STATS]);
      logAction('SAVE_CATEGORIES', 'admin', cats.length + ' kategori');
      return { status: 'success', message: 'Kategori disimpan!' };
    } catch (err) { return { status: 'error', message: err.message }; }
  });
}

// ==================== JERSEYS ====================
function getJerseys(noCache) {
  try {
    if (!noCache) {
      const c = CacheService.getScriptCache().get(CACHE_KEY_JERSEYS);
      if (c) return { status: 'success', jerseys: JSON.parse(c) };
    }

    const sheet = getSheet(SHEET_JERSEYS);
    let rows = sheet.getDataRange().getValues();
    if (rows.length < 2) { initSheet(SHEET_JERSEYS, sheet); rows = sheet.getDataRange().getValues(); }

    const jerseys = rows.slice(1).filter(r => r[0]).map((r, i) => ({
      code: String(r[0]), label: String(r[1]), order: Number(r[2]) || i + 1
    })).sort((a, b) => a.order - b.order);

    CacheService.getScriptCache().put(CACHE_KEY_JERSEYS, JSON.stringify(jerseys), CACHE_SEC);
    return { status: 'success', jerseys };
  } catch (err) {
    return { status: 'error', message: err.message, jerseys: [] };
  }
}

function saveJerseys(data) {
  return withLock(() => {
    try {
      const js = data.jerseys || [];
      if (js.length === 0) return { status: 'error', message: 'Minimal 1 ukuran.' };
      for (const j of js) {
        if (!j.code || !j.label) return { status: 'error', message: 'Kode & label wajib diisi.' };
      }
      const codes = js.map(j => String(j.code).toUpperCase());
      if (new Set(codes).size !== codes.length) return { status: 'error', message: 'Kode tidak boleh duplikat.' };

      const sheet = getSheet(SHEET_JERSEYS);
      sheet.clear();
      sheet.getRange(1, 1, 1, 3).setValues([['Kode', 'Label', 'Urutan']])
        .setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff').setHorizontalAlignment('center');
      sheet.getRange(2, 1, js.length, 3).setValues(js.map((j, i) => [String(j.code).toUpperCase(), String(j.label), Number(j.order) || i + 1]));
      sheet.setFrozenRows(1);

      clearCache([CACHE_KEY_JERSEYS]);
      logAction('SAVE_JERSEYS', 'admin', js.length + ' ukuran');
      return { status: 'success', message: 'Ukuran kaos disimpan!' };
    } catch (err) { return { status: 'error', message: err.message }; }
  });
}

// ==================== CS ====================
function getCs(noCache) {
  try {
    if (!noCache) {
      const c = CacheService.getScriptCache().get(CACHE_KEY_CS);
      if (c) return { status: 'success', cs: JSON.parse(c) };
    }

    const sheet = getSheet(SHEET_CS);
    let rows = sheet.getDataRange().getValues();
    if (rows.length < 2) { initSheet(SHEET_CS, sheet); rows = sheet.getDataRange().getValues(); }

    const cs = rows.slice(1).filter(r => r[0]).map((r, i) => ({
      name: String(r[0]),
      phone: String(r[1]).replace(/[^0-9]/g, ''),
      order: Number(r[2]) || i + 1
    })).sort((a, b) => a.order - b.order);

    CacheService.getScriptCache().put(CACHE_KEY_CS, JSON.stringify(cs), CACHE_SEC);
    return { status: 'success', cs };
  } catch (err) {
    return { status: 'error', message: err.message, cs: [] };
  }
}

function saveCs(data) {
  return withLock(() => {
    try {
      const list = data.cs || [];
      if (list.length === 0) return { status: 'error', message: 'Minimal 1 kontak CS.' };
      for (const c of list) {
        if (!c.name || !c.phone) return { status: 'error', message: 'Nama & nomor WA wajib diisi.' };
      }

      const sheet = getSheet(SHEET_CS);
      sheet.clear();
      sheet.getRange(1, 1, 1, 3).setValues([['Nama', 'NomorWA', 'Urutan']])
        .setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff').setHorizontalAlignment('center');
      sheet.getRange(2, 1, list.length, 3).setValues(list.map((c, i) => [
        String(c.name),
        String(c.phone).replace(/[^0-9]/g, ''),
        Number(c.order) || i + 1
      ]));
      sheet.setFrozenRows(1);

      clearCache([CACHE_KEY_CS]);
      logAction('SAVE_CS', 'admin', list.length + ' kontak CS');
      return { status: 'success', message: 'Kontak CS disimpan!' };
    } catch (err) { return { status: 'error', message: err.message }; }
  });
}

// ==================== SETTINGS ====================
function getSettings(noCache) {
  try {
    if (!noCache) {
      const c = CacheService.getScriptCache().get(CACHE_KEY_SETTINGS);
      if (c) return { status: 'success', settings: JSON.parse(c) };
    }

    const sheet = getSheet(SHEET_SETTINGS);
    let rows = sheet.getDataRange().getValues();
    if (rows.length < 2) { initSheet(SHEET_SETTINGS, sheet); rows = sheet.getDataRange().getValues(); }

    const settings = {};
    rows.slice(1).forEach(r => { if (r[0]) settings[r[0]] = r[1]; });
    if (!settings.adminUser) settings.adminUser = DEFAULT_ADMIN_USER;
    if (!settings.adminPass) settings.adminPass = DEFAULT_ADMIN_PASS;

    CacheService.getScriptCache().put(CACHE_KEY_SETTINGS, JSON.stringify(settings), CACHE_SEC);
    return { status: 'success', settings };
  } catch (err) {
    return { status: 'error', message: err.message };
  }
}

function saveSettings(data) {
  return withLock(() => {
    try {
      const sheet = getSheet(SHEET_SETTINGS);
      const rows = sheet.getDataRange().getValues();
      const keys = ['title', 'date', 'bank', 'qris', 'adminUser', 'adminPass'];

      keys.forEach(key => {
        if (data[key] === undefined || data[key] === '') return;
        let found = false;
        for (let i = 1; i < rows.length; i++) {
          if (rows[i][0] === key) {
            sheet.getRange(i + 1, 2).setValue(data[key]);
            found = true;
            break;
          }
        }
        if (!found) sheet.appendRow([key, data[key]]);
      });

      clearCache([CACHE_KEY_SETTINGS]);
      logAction('SAVE_SETTINGS', 'admin', 'Settings updated');
      return { status: 'success', message: 'Pengaturan disimpan!' };
    } catch (err) { return { status: 'error', message: err.message }; }
  });
}

// ==================== PUBLIC STATS ====================
function getPublicStats(noCache) {
  try {
    if (!noCache) {
      const c = CacheService.getScriptCache().get(CACHE_KEY_STATS);
      if (c) return JSON.parse(c);
    }

    const catRes = getCategories(true);
    const categories = catRes.categories || [];
    const participants = sheetToObjects(getSheet(SHEET_PARTICIPANTS));

    const stats = categories.map(cat => {
      const used = participants.filter(p => String(p.Kategori) === cat.code && (p.Status === 'Verified' || p.Status === 'Pending')).length;
      const verified = participants.filter(p => String(p.Kategori) === cat.code && p.Status === 'Verified').length;
      const pending = participants.filter(p => String(p.Kategori) === cat.code && p.Status === 'Pending').length;

      return {
        code: cat.code, name: cat.name, price: cat.price, quota: cat.quota,
        used, verified, pending, remaining: Math.max(0, cat.quota - used)
      };
    });

    const totalQuota = categories.reduce((s, c) => s + Number(c.quota), 0);
    const totalUsed = stats.reduce((s, x) => s + x.used, 0);
    const result = {
      status: 'success', stats,
      totalQuota, totalUsed,
      totalRemaining: Math.max(0, totalQuota - totalUsed)
    };

    CacheService.getScriptCache().put(CACHE_KEY_STATS, JSON.stringify(result), 10);
    return result;
  } catch (err) {
    return { status: 'error', message: err.message, stats: [], totalQuota: 0, totalUsed: 0, totalRemaining: 0 };
  }
}

// ==================== REGISTER ====================
function registerParticipant(data) {
  return withLock(() => {
    try {
      if (!data.name || !data.email || !data.phone || !data.password) return { status: 'error', message: 'Data tidak lengkap.' };
      if (data.password.length < 6) return { status: 'error', message: 'Password minimal 6 karakter.' };
      if (!data.gender || (data.gender !== 'L' && data.gender !== 'P')) return { status: 'error', message: 'Jenis kelamin wajib dipilih.' };
      if (!data.category) return { status: 'error', message: 'Kategori wajib dipilih.' };
      if (!data.jersey) return { status: 'error', message: 'Ukuran kaos wajib dipilih.' };

      const sheet = getSheet(SHEET_PARTICIPANTS);
      const existing = sheetToObjects(sheet);

      if (existing.some(p => String(p.Email || '').toLowerCase() === String(data.email).toLowerCase())) {
        return { status: 'error', message: 'Email sudah terdaftar.' };
      }

      const catRes = getCategories(true);
      const category = (catRes.categories || []).find(c => c.code === data.category);
      if (!category) return { status: 'error', message: 'Kategori tidak valid.' };

      const jerseyRes = getJerseys(true);
      if (!(jerseyRes.jerseys || []).find(j => j.code === data.jersey)) return { status: 'error', message: 'Ukuran kaos tidak valid.' };

      const usedInCat = existing.filter(p => String(p.Kategori) === data.category && (p.Status === 'Verified' || p.Status === 'Pending')).length;
      if (usedInCat >= category.quota) return { status: 'error', message: 'Kuota ' + category.name + ' penuh.' };

      // Generate unique ID
      let counter = existing.length + 1;
      let newId = 'REG-' + (1000 + counter);
      const ids = existing.map(p => p.ID);
      while (ids.indexOf(newId) > -1) { counter++; newId = 'REG-' + (1000 + counter); }

      sheet.appendRow([
        newId, data.name, data.phone, data.email, data.password,
        data.gender, data.category, data.jersey, 'Pending', 'Pending',
        '', new Date().toLocaleString('id-ID'), '', category.price
      ]);

      clearCache([CACHE_KEY_STATS]);
      logAction('REGISTER', data.email, data.name + ' (' + data.gender + ') - ' + data.category);

      return {
        status: 'success',
        message: 'Pendaftaran berhasil!',
        user: {
          id: newId, name: data.name, email: data.email, phone: data.phone,
          gender: data.gender, category: data.category, jersey: data.jersey,
          bib: 'Pending', status: 'Pending', proof: '', price: category.price
        }
      };
    } catch (err) {
      return { status: 'error', message: 'Gagal: ' + err.message };
    }
  });
}

// ==================== LOGIN ====================
function loginUser(data) {
  try {
    const { role, identifier, password } = data;
    if (!identifier || !password) return { status: 'error', message: 'Kredensial wajib diisi.' };

    if (role === 'admin') {
      const s = (getSettings().settings) || {};
      const u = s.adminUser || DEFAULT_ADMIN_USER;
      const p = s.adminPass || DEFAULT_ADMIN_PASS;
      if (identifier === u && password === p) {
        logAction('LOGIN_ADMIN', identifier, 'Admin login');
        return { status: 'success', user: { role: 'admin', name: 'Administrator Event' } };
      }
      return { status: 'error', message: 'Username/Password Admin salah.' };
    }

    const participants = sheetToObjects(getSheet(SHEET_PARTICIPANTS));
    const user = participants.find(p =>
      (String(p.Email || '').toLowerCase() === String(identifier).toLowerCase() || String(p.NoHP || '') === String(identifier)) &&
      String(p.Password || '') === String(password)
    );

    if (!user) return { status: 'error', message: 'Email/No HP atau Password salah.' };
    logAction('LOGIN_PESERTA', user.Email, user.Nama);
    return { status: 'success', user: mapParticipant(user) };
  } catch (err) {
    return { status: 'error', message: 'Login gagal: ' + err.message };
  }
}

// ==================== UPLOAD PROOF ====================
function uploadProof(data) {
  try {
    const { userId, base64Image, fileName } = data;
    if (!userId || !base64Image) return { status: 'error', message: 'Data tidak lengkap.' };

    const up = uploadToDrive(base64Image, (fileName || 'PROOF_' + userId) + '_' + Date.now() + '.jpg', 'image/jpeg');

    return withLock(() => {
      try {
        const sheet = getSheet(SHEET_PARTICIPANTS);
        const rows = sheet.getDataRange().getValues();
        let found = false;
        for (let i = 1; i < rows.length; i++) {
          if (rows[i][0] === userId) {
            sheet.getRange(i + 1, 11).setValue(up.url);
            sheet.getRange(i + 1, 10).setValue('Pending');
            found = true;
            break;
          }
        }
        if (!found) return { status: 'error', message: 'Peserta tidak ditemukan.' };

        clearCache([CACHE_KEY_STATS]);
        logAction('UPLOAD_PROOF', userId, 'Bukti diunggah');

        return {
          status: 'success', message: 'Bukti berhasil diunggah!',
          proofUrl: up.url, previewUrl: up.preview, viewUrl: up.view, downloadUrl: up.download
        };
      } catch (err) {
        return { status: 'error', message: 'Gagal simpan: ' + err.message };
      }
    });
  } catch (err) {
    return { status: 'error', message: 'Upload gagal: ' + err.message };
  }
}

// ==================== PARTICIPANTS ====================
function getParticipants() {
  try {
    const participants = sheetToObjects(getSheet(SHEET_PARTICIPANTS)).map(mapParticipant);
    return { status: 'success', participants };
  } catch (err) {
    return { status: 'error', message: err.message, participants: [] };
  }
}

// ==================== VERIFY ====================
function verifyParticipant(data) {
  return withLock(() => {
    try {
      const { userId, newStatus } = data;
      if (!userId || !newStatus) return { status: 'error', message: 'Parameter tidak lengkap.' };

      const sheet = getSheet(SHEET_PARTICIPANTS);
      const rows = sheet.getDataRange().getValues();
      let rowIdx = -1, name = '', gender = 'L';

      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === userId) { rowIdx = i; name = rows[i][1]; gender = rows[i][5] || 'L'; break; }
      }
      if (rowIdx === -1) return { status: 'error', message: 'Peserta tidak ditemukan.' };

      let newBib = '';
      if (newStatus === 'Verified') {
        const prefix = gender === 'P' ? 'P' : 'L';
        const bibs = rows.slice(1)
          .map(r => String(r[8] || ''))
          .filter(b => b.indexOf(prefix + '-') === 0)
          .map(b => parseInt(b.replace(prefix + '-', ''), 10))
          .filter(n => !isNaN(n));
        const nextNum = (bibs.length > 0 ? Math.max.apply(null, bibs) : 0) + 1;
        newBib = prefix + '-' + String(nextNum).padStart(3, '0');
      } else if (newStatus === 'Rejected') {
        newBib = 'Rejected';
      }

      sheet.getRange(rowIdx + 1, 10).setValue(newStatus);
      if (newBib) sheet.getRange(rowIdx + 1, 9).setValue(newBib);
      sheet.getRange(rowIdx + 1, 13).setValue(new Date().toLocaleString('id-ID'));

      clearCache([CACHE_KEY_STATS]);
      logAction('VERIFY', userId, name + ' → ' + newStatus + ' (' + newBib + ')');

      return { status: 'success', message: 'Status diubah ke ' + newStatus, bib: newBib };
    } catch (err) { return { status: 'error', message: err.message }; }
  });
}

// ==================== RESET PASSWORD ====================
function resetPassword(data) {
  return withLock(() => {
    try {
      const { userId } = data;
      if (!userId) return { status: 'error', message: 'User ID wajib diisi.' };

      const sheet = getSheet(SHEET_PARTICIPANTS);
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === userId) {
          sheet.getRange(i + 1, 5).setValue(DEFAULT_RESET_PASS);
          logAction('RESET_PASS', userId, 'Password direset');
          return { status: 'success', message: 'Password direset menjadi: ' + DEFAULT_RESET_PASS };
        }
      }
      return { status: 'error', message: 'Peserta tidak ditemukan.' };
    } catch (err) { return { status: 'error', message: err.message }; }
  });
}

// ==================== CHECK STATUS ====================
function checkStatus(data) {
  try {
    const { email } = data;
    if (!email) return { status: 'error', message: 'Email wajib diisi.' };

    const participants = sheetToObjects(getSheet(SHEET_PARTICIPANTS));
    const user = participants.find(p => String(p.Email || '').toLowerCase() === String(email).toLowerCase());
    if (!user) return { status: 'error', message: 'Email tidak terdaftar.' };

    return { status: 'success', user: mapParticipant(user) };
  } catch (err) {
    return { status: 'error', message: err.message };
  }
}

// ==================== UTILITY ====================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function testDatabase() {
  Logger.log('Setup: ' + JSON.stringify(checkSetupStatus()));
  Logger.log('Categories: ' + JSON.stringify(getCategories()));
  Logger.log('Jerseys: ' + JSON.stringify(getJerseys()));
  Logger.log('CS: ' + JSON.stringify(getCs()));
  Logger.log('Stats: ' + JSON.stringify(getPublicStats()));
  Logger.log('Participants: ' + (getParticipants().participants || []).length);
}