const TZ = 'Asia/Bangkok';
const SHEETS = {
  CHECKPOINTS: 'Checkpoints',
  GUARDS: 'Guards',
  REPORTS: 'Reports',
  SETTINGS: 'Settings'
};

const HEADERS = {
  Checkpoints: ['CheckpointID','PostID','CheckpointName','Zone','QRToken','Active','Latitude','Longitude','Radius','CreatedAt','UpdatedAt'],
  Guards: ['GuardID','Name','Active','CreatedAt','UpdatedAt'],
  Reports: ['ReportID','Timestamp','CheckpointID','CheckpointName','PostID','GuardID','GuardName','Situation','EscapeTrace','Lighting','SuspiciousPerson','AbnormalItem','Details','ActionTaken','QRToken','Latitude','Longitude','DistanceMeters'],
  Settings: ['Key','Value','UpdatedAt']
};

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TZ);

  Object.keys(HEADERS).forEach(name => ensureSheet_(ss, name, HEADERS[name]));
  seedCheckpoints_(ss);
  seedSettings_(ss);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_PASSWORD')) {
    const password = 'ADM-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    props.setProperty('ADMIN_PASSWORD', password);
    Logger.log('Generated admin password: ' + password);
  } else {
    Logger.log('Admin password already exists. Run showAdminPassword() if needed.');
  }

  Logger.log('Database setup complete. Deploy this script as Web app: Execute as Me / Who has access: Anyone.');
}

function showAdminPassword() {
  Logger.log('ADMIN_PASSWORD = ' + (PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || 'NOT SET'));
}

function setAdminPassword() {
  const newPassword = 'CHANGE_THIS_PASSWORD';
  if (newPassword === 'CHANGE_THIS_PASSWORD' || newPassword.length < 8) {
    throw new Error('Edit setAdminPassword() and set a password with at least 8 characters first.');
  }
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', newPassword);
  Logger.log('Admin password updated.');
}

function doGet() {
  return json_({ ok: true, data: { service: 'Security Check-in API', status: 'online', time: nowText_() } });
}

function doPost(e) {
  try {
    const body = parseBody_(e);
    const action = String(body.action || '').trim();
    let data;

    switch (action) {
      case 'bootstrap': data = bootstrap_(); break;
      case 'getCheckpoint': data = getCheckpoint_(body.token); break;
      case 'submitReport': data = submitReport_(body); break;
      case 'adminLogin': data = adminLogin_(body.password); break;
      case 'adminLogout': data = adminLogout_(body.sessionToken); break;
      case 'adminData': requireAdmin_(body.sessionToken); data = adminData_(); break;
      case 'saveGuard': requireAdmin_(body.sessionToken); data = saveGuard_(body.guard); break;
      case 'saveCheckpoint': requireAdmin_(body.sessionToken); data = saveCheckpoint_(body.checkpoint); break;
      case 'regenerateToken': requireAdmin_(body.sessionToken); data = regenerateToken_(body.checkpointId); break;
      default: throw new Error('Unknown action: ' + action);
    }

    return json_({ ok: true, data: data });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: err && err.message ? err.message : String(err) });
  }
}

function bootstrap_() {
  return {
    guards: getGuards_(true),
    appName: getSetting_('APP_NAME') || 'ระบบตรวจรักษาความปลอดภัย',
    serverTime: nowText_()
  };
}

function getCheckpoint_(token) {
  token = String(token || '').trim();
  if (!token) throw new Error('QR token is required.');
  const rows = sheetObjects_(SHEETS.CHECKPOINTS);
  const cp = rows.find(r => String(r.QRToken) === token && truthy_(r.Active));
  if (!cp) throw new Error('ไม่พบจุดตรวจ หรือ QR นี้ถูกปิดใช้งาน');
  return publicCheckpoint_(cp);
}

function submitReport_(body) {
  const token = String(body.checkpointToken || '').trim();
  const guardId = String(body.guardId || '').trim();
  if (!token) throw new Error('กรุณาสแกน QR จุดตรวจ');
  if (!guardId) throw new Error('กรุณาเลือกชื่อ รปภ.');

  const cp = sheetObjects_(SHEETS.CHECKPOINTS).find(r => String(r.QRToken) === token && truthy_(r.Active));
  if (!cp) throw new Error('QR จุดตรวจไม่ถูกต้องหรือถูกปิดใช้งาน');

  const guard = getGuards_(true).find(g => String(g.GuardID) === guardId);
  if (!guard) throw new Error('ไม่พบชื่อ รปภ. หรือบัญชีถูกปิดใช้งาน');

  const allowedSituation = ['ปกติ','เฝ้าระวัง','เร่งด่วน'];
  const allowedYesNo = ['ไม่พบ','พบ'];
  const allowedLighting = ['ปกติ','ผิดปกติ'];
  const situation = allowedSituation.includes(body.situation) ? body.situation : 'ปกติ';
  const escapeTrace = allowedYesNo.includes(body.escapeTrace) ? body.escapeTrace : 'ไม่พบ';
  const lighting = allowedLighting.includes(body.lighting) ? body.lighting : 'ปกติ';
  const suspiciousPerson = allowedYesNo.includes(body.suspiciousPerson) ? body.suspiciousPerson : 'ไม่พบ';
  const abnormalItem = allowedYesNo.includes(body.abnormalItem) ? body.abnormalItem : 'ไม่พบ';

  const guardLat = (body.latitude != null && body.latitude !== '' && !isNaN(Number(body.latitude))) ? Number(body.latitude) : null;
  const guardLng = (body.longitude != null && body.longitude !== '' && !isNaN(Number(body.longitude))) ? Number(body.longitude) : null;
  let distanceMeters = (body.distanceMeters != null && body.distanceMeters !== '' && !isNaN(Number(body.distanceMeters))) ? Number(body.distanceMeters) : null;

  // Anti-fraud GPS check if checkpoint has reference coordinates (Max 50 meters)
  const cpLat = (cp.Latitude != null && cp.Latitude !== '' && !isNaN(Number(cp.Latitude))) ? Number(cp.Latitude) : null;
  const cpLng = (cp.Longitude != null && cp.Longitude !== '' && !isNaN(Number(cp.Longitude))) ? Number(cp.Longitude) : null;
  const cpRadius = (cp.Radius != null && cp.Radius !== '' && !isNaN(Number(cp.Radius))) ? Number(cp.Radius) : 50;

  if (cpLat != null && cpLng != null) {
    if (guardLat == null || guardLng == null) {
      throw new Error('ไม่พบข้อมูลพิกัด GPS เพื่อยืนยันตัวตน กรุณาเปิด Location Service ก่อนส่งรายงาน');
    }
    const realDist = haversineDistance_(guardLat, guardLng, cpLat, cpLng);
    distanceMeters = Math.round(realDist);
    if (realDist > cpRadius) {
      throw new Error('อยู่นอกพื้นที่จุดตรวจ! ท่านอยู่ห่าง ' + Math.round(realDist) + ' เมตร (กำหนดไม่เกิน ' + cpRadius + ' เมตร) ไม่อนุญาตให้ลงเวลาตรวจ');
    }
  }

  const cache = CacheService.getScriptCache();
  const dupeKey = 'dupe_' + guardId + '_' + cp.CheckpointID;
  if (cache.get(dupeKey)) throw new Error('มีการส่งรายงานจุดนี้เมื่อสักครู่ กรุณารอสักครู่ก่อนส่งซ้ำ');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const now = new Date();
    const reportId = 'RPT-' + Utilities.formatDate(now, TZ, 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 4).toUpperCase();
    const reportSheet = getSheet_(SHEETS.REPORTS);

    ensureColumnHeader_(reportSheet, 'Latitude');
    ensureColumnHeader_(reportSheet, 'Longitude');
    ensureColumnHeader_(reportSheet, 'DistanceMeters');

    const row = [
      reportId,
      now,
      cp.CheckpointID,
      safeCell_(cp.CheckpointName),
      cp.PostID,
      guard.GuardID,
      safeCell_(guard.Name),
      situation,
      escapeTrace,
      lighting,
      suspiciousPerson,
      abnormalItem,
      safeCell_(String(body.details || '').trim()),
      safeCell_(String(body.actionTaken || '').trim()),
      token,
      guardLat != null ? guardLat : '',
      guardLng != null ? guardLng : '',
      distanceMeters != null ? distanceMeters : ''
    ];
    reportSheet.appendRow(row);
    cache.put(dupeKey, '1', 30);
    return { reportId, timestamp: formatDate_(now), checkpoint: cp.CheckpointName, guard: guard.Name, distanceMeters: distanceMeters };
  } finally {
    lock.releaseLock();
  }
}

function adminLogin_(password) {
  const actual = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!actual) throw new Error('Admin password is not configured. Run setupDatabase().');
  if (String(password || '') !== actual) throw new Error('รหัสผ่านไม่ถูกต้อง');
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('admin_' + token, '1', 21600);
  return { sessionToken: token, expiresInSeconds: 21600 };
}

function adminLogout_(token) {
  if (token) CacheService.getScriptCache().remove('admin_' + String(token));
  return { loggedOut: true };
}

function requireAdmin_(token) {
  token = String(token || '');
  if (!token || !CacheService.getScriptCache().get('admin_' + token)) {
    throw new Error('ADMIN_SESSION_EXPIRED');
  }
}

function adminData_() {
  const guards = getGuards_(false);
  const checkpoints = sheetObjects_(SHEETS.CHECKPOINTS).map(publicCheckpointAdmin_);
  const reports = sheetObjects_(SHEETS.REPORTS).reverse().map(serializeReport_);
  const todayKey = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const todayReports = reports.filter(r => r.DateKey === todayKey);

  const dailyLabels = [];
  const dailyCounts = [];
  for (let i = 6; i >= 0; i--) {
    const d = new Date(Date.now() - i * 86400000);
    const key = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    dailyLabels.push(Utilities.formatDate(d, TZ, 'dd/MM'));
    dailyCounts.push(reports.filter(r => r.DateKey === key).length);
  }

  const situation = { 'ปกติ': 0, 'เฝ้าระวัง': 0, 'เร่งด่วน': 0 };
  todayReports.forEach(r => { if (situation[r.Situation] !== undefined) situation[r.Situation]++; });

  return {
    guards,
    checkpoints,
    reports: reports.slice(0, 500),
    stats: {
      today: todayReports.length,
      activeCheckpoints: checkpoints.filter(x => x.Active).length,
      watch: situation['เฝ้าระวัง'],
      urgent: situation['เร่งด่วน']
    },
    charts: {
      daily: { labels: dailyLabels, values: dailyCounts },
      situation
    }
  };
}

function saveGuard_(guard) {
  guard = guard || {};
  const name = String(guard.Name || '').trim();
  if (!name) throw new Error('กรุณาระบุชื่อ รปภ.');
  const sheet = getSheet_(SHEETS.GUARDS);
  const rows = sheetObjects_(SHEETS.GUARDS);
  const now = new Date();

  if (guard.GuardID) {
    const idx = rows.findIndex(r => String(r.GuardID) === String(guard.GuardID));
    if (idx < 0) throw new Error('ไม่พบ รปภ.');
    const rowNo = idx + 2;
    sheet.getRange(rowNo, 2).setValue(safeCell_(name));
    sheet.getRange(rowNo, 3).setValue(Boolean(guard.Active));
    sheet.getRange(rowNo, 5).setValue(now);
    return { GuardID: guard.GuardID, Name: name, Active: Boolean(guard.Active) };
  }

  const id = 'G-' + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
  sheet.appendRow([id, safeCell_(name), guard.Active !== false, now, now]);
  return { GuardID: id, Name: name, Active: guard.Active !== false };
}

function saveCheckpoint_(checkpoint) {
  checkpoint = checkpoint || {};
  const name = String(checkpoint.CheckpointName || '').trim();
  const post = String(checkpoint.PostID || '').trim();
  const zone = String(checkpoint.Zone || '').trim();
  const lat = (checkpoint.Latitude != null && checkpoint.Latitude !== '' && !isNaN(Number(checkpoint.Latitude))) ? Number(checkpoint.Latitude) : '';
  const lng = (checkpoint.Longitude != null && checkpoint.Longitude !== '' && !isNaN(Number(checkpoint.Longitude))) ? Number(checkpoint.Longitude) : '';
  const radius = (checkpoint.Radius != null && checkpoint.Radius !== '' && !isNaN(Number(checkpoint.Radius))) ? Number(checkpoint.Radius) : 50;
  if (!name) throw new Error('กรุณาระบุชื่อจุดตรวจ');
  if (!post) throw new Error('กรุณาระบุป้อม/กลุ่มจุดตรวจ');

  const sheet = getSheet_(SHEETS.CHECKPOINTS);
  const rows = sheetObjects_(SHEETS.CHECKPOINTS);
  const now = new Date();

  ensureColumnHeader_(sheet, 'Latitude');
  ensureColumnHeader_(sheet, 'Longitude');
  ensureColumnHeader_(sheet, 'Radius');

  if (checkpoint.CheckpointID) {
    const idx = rows.findIndex(r => String(r.CheckpointID) === String(checkpoint.CheckpointID));
    if (idx < 0) throw new Error('ไม่พบจุดตรวจ');
    const rowNo = idx + 2;
    setRowValueByHeader_(sheet, rowNo, 'PostID', safeCell_(post));
    setRowValueByHeader_(sheet, rowNo, 'CheckpointName', safeCell_(name));
    setRowValueByHeader_(sheet, rowNo, 'Zone', safeCell_(zone));
    setRowValueByHeader_(sheet, rowNo, 'Active', Boolean(checkpoint.Active));
    setRowValueByHeader_(sheet, rowNo, 'Latitude', lat);
    setRowValueByHeader_(sheet, rowNo, 'Longitude', lng);
    setRowValueByHeader_(sheet, rowNo, 'Radius', radius);
    setRowValueByHeader_(sheet, rowNo, 'UpdatedAt', now);
    return { CheckpointID: checkpoint.CheckpointID, PostID: post, CheckpointName: name, Zone: zone, Active: Boolean(checkpoint.Active), Latitude: lat, Longitude: lng, Radius: radius };
  }

  const id = 'CP-' + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
  const qrToken = Utilities.getUuid();
  const newRowNo = sheet.getLastRow() + 1;
  sheet.appendRow([id]);
  setRowValueByHeader_(sheet, newRowNo, 'CheckpointID', id);
  setRowValueByHeader_(sheet, newRowNo, 'PostID', safeCell_(post));
  setRowValueByHeader_(sheet, newRowNo, 'CheckpointName', safeCell_(name));
  setRowValueByHeader_(sheet, newRowNo, 'Zone', safeCell_(zone));
  setRowValueByHeader_(sheet, newRowNo, 'QRToken', qrToken);
  setRowValueByHeader_(sheet, newRowNo, 'Active', checkpoint.Active !== false);
  setRowValueByHeader_(sheet, newRowNo, 'Latitude', lat);
  setRowValueByHeader_(sheet, newRowNo, 'Longitude', lng);
  setRowValueByHeader_(sheet, newRowNo, 'Radius', radius);
  setRowValueByHeader_(sheet, newRowNo, 'CreatedAt', now);
  setRowValueByHeader_(sheet, newRowNo, 'UpdatedAt', now);

  return { CheckpointID: id, PostID: post, CheckpointName: name, Zone: zone, QRToken: qrToken, Active: checkpoint.Active !== false, Latitude: lat, Longitude: lng, Radius: radius };
}

function regenerateToken_(checkpointId) {
  const rows = sheetObjects_(SHEETS.CHECKPOINTS);
  const idx = rows.findIndex(r => String(r.CheckpointID) === String(checkpointId || ''));
  if (idx < 0) throw new Error('ไม่พบจุดตรวจ');
  const token = Utilities.getUuid();
  const sheet = getSheet_(SHEETS.CHECKPOINTS);
  sheet.getRange(idx + 2, 5).setValue(token);
  sheet.getRange(idx + 2, 8).setValue(new Date());
  return { CheckpointID: checkpointId, QRToken: token };
}

function getGuards_(activeOnly) {
  return sheetObjects_(SHEETS.GUARDS)
    .filter(r => !activeOnly || truthy_(r.Active))
    .map(r => ({ GuardID: String(r.GuardID), Name: String(r.Name), Active: truthy_(r.Active) }));
}

function seedCheckpoints_(ss) {
  const sheet = ss.getSheetByName(SHEETS.CHECKPOINTS);
  if (sheet.getLastRow() > 1) return;
  const now = new Date();
  const seed = [
    ['CP001','POST01','ตึกอำนวยการ ชั้น 1','ตึกอำนวยการ'],
    ['CP002','POST01','ตึกอำนวยการ ชั้น 2','ตึกอำนวยการ'],
    ['CP003','POST01','ตึกอำนวยการ ชั้น 3','ตึกอำนวยการ'],
    ['CP004','POST02','แปลงเกษตร','พื้นที่ชั้นใน'],
    ['CP005','POST02','โดมกีฬา','พื้นที่ชั้นใน'],
    ['CP006','POST02','โรงอาหาร','พื้นที่ชั้นใน'],
    ['CP007','POST03','หลังอาคารเรือนนอนชาย','อาคารเรือนนอนชาย'],
    ['CP008','POST03','อาคารเรือนนอนชาย ชั้น 1','อาคารเรือนนอนชาย'],
    ['CP009','POST03','อาคารเรือนนอนชาย ชั้น 2','อาคารเรือนนอนชาย'],
    ['CP010','POST03','อาคารเรือนนอนชาย ชั้น 3','อาคารเรือนนอนชาย'],
    ['CP011','POST03','ห้องน้ำโยธา','อาคารเรือนนอนชาย']
  ];
  const rows = seed.map(x => [x[0],x[1],x[2],x[3],Utilities.getUuid(),true,now,now]);
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedSettings_(ss) {
  const sheet = ss.getSheetByName(SHEETS.SETTINGS);
  if (sheet.getLastRow() > 1) return;
  sheet.getRange(2, 1, 2, 3).setValues([
    ['APP_NAME','ระบบตรวจรักษาความปลอดภัย',new Date()],
    ['VERSION','1.0.0',new Date()]
  ]);
}

function getSetting_(key) {
  const row = sheetObjects_(SHEETS.SETTINGS).find(r => String(r.Key) === String(key));
  return row ? row.Value : '';
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  return sheet;
}

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: ' + name + '. Run setupDatabase() first.');
  return sheet;
}

function sheetObjects_(name) {
  const sheet = getSheet_(name);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).filter(row => row.some(v => v !== '')).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function publicCheckpoint_(r) {
  return {
    CheckpointID: String(r.CheckpointID), PostID: String(r.PostID), CheckpointName: String(r.CheckpointName),
    Zone: String(r.Zone), QRToken: String(r.QRToken), Active: truthy_(r.Active),
    Latitude: (r.Latitude != null && r.Latitude !== '' && !isNaN(Number(r.Latitude))) ? Number(r.Latitude) : null,
    Longitude: (r.Longitude != null && r.Longitude !== '' && !isNaN(Number(r.Longitude))) ? Number(r.Longitude) : null,
    Radius: (r.Radius != null && r.Radius !== '' && !isNaN(Number(r.Radius))) ? Number(r.Radius) : 50
  };
}

function publicCheckpointAdmin_(r) { return publicCheckpoint_(r); }

function serializeReport_(r) {
  const d = r.Timestamp instanceof Date ? r.Timestamp : new Date(r.Timestamp);
  return {
    ReportID: String(r.ReportID),
    Timestamp: formatDate_(d),
    DateKey: Utilities.formatDate(d, TZ, 'yyyy-MM-dd'),
    CheckpointID: String(r.CheckpointID), CheckpointName: String(r.CheckpointName), PostID: String(r.PostID),
    GuardID: String(r.GuardID), GuardName: String(r.GuardName), Situation: String(r.Situation),
    EscapeTrace: String(r.EscapeTrace), Lighting: String(r.Lighting), SuspiciousPerson: String(r.SuspiciousPerson),
    AbnormalItem: String(r.AbnormalItem), Details: String(r.Details || ''), ActionTaken: String(r.ActionTaken || ''),
    Latitude: (r.Latitude != null && r.Latitude !== '' && !isNaN(Number(r.Latitude))) ? Number(r.Latitude) : null,
    Longitude: (r.Longitude != null && r.Longitude !== '' && !isNaN(Number(r.Longitude))) ? Number(r.Longitude) : null,
    DistanceMeters: (r.DistanceMeters != null && r.DistanceMeters !== '' && !isNaN(Number(r.DistanceMeters))) ? Number(r.DistanceMeters) : null
  };
}

function haversineDistance_(lat1, lon1, lat2, lon2) {
  const R = 6371e3; // meters
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
            Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
            Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

function ensureColumnHeader_(sheet, headerName) {
  const lastCol = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  let colIdx = headers.indexOf(headerName);
  if (colIdx === -1) {
    colIdx = headers.length;
    sheet.getRange(1, colIdx + 1).setValue(headerName).setFontWeight('bold');
  }
  return colIdx + 1;
}

function setRowValueByHeader_(sheet, rowNo, headerName, value) {
  const col = ensureColumnHeader_(sheet, headerName);
  sheet.getRange(rowNo, col).setValue(value);
}

function parseBody_(e) {
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
  try { return JSON.parse(raw); } catch (_) { return e.parameter || {}; }
}

function safeCell_(value) {
  const s = String(value == null ? '' : value);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function truthy_(v) { return v === true || String(v).toLowerCase() === 'true' || String(v) === '1'; }
function nowText_() { return formatDate_(new Date()); }
function formatDate_(d) { return Utilities.formatDate(d, TZ, 'dd/MM/yyyy HH:mm:ss'); }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
