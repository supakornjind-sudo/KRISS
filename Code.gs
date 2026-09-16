/**
 * ระบบ "รายการของ & รายชื่อคน"
 * ฝั่งเซิร์ฟเวอร์ (Google Apps Script) — ทำหน้าที่เป็น API ให้หน้าเว็บบน GitHub Pages
 *  - เก็บข้อมูลใน Google Sheet
 *  - เก็บรูปสินค้าใน Google Drive
 *
 * ⚠️ ก่อนใช้งาน: ตั้งรหัสผ่าน APP_PASSWORD ใน พร็อพเพอร์ตี้ของสคริปต์ (ดู README)
 * วิธีติดตั้งดูใน README.md
 */

const CONFIG = {
  // รหัสผ่าน: แนะนำให้ตั้งใน "การตั้งค่าโปรเจ็กต์ > พร็อพเพอร์ตี้ของสคริปต์" ชื่อ APP_PASSWORD
  // (ถ้าไม่ได้ตั้งไว้ ระบบจะใช้ค่าด้านล่างนี้แทน — อย่าใส่รหัสจริงในไฟล์ที่อัปขึ้น GitHub)
  APP_PASSWORD: 'CHANGE-ME',
  APP_TITLE: 'รายการของ & รายชื่อ',
  SHEET_ITEMS: 'รายการของ',
  SHEET_PEOPLE: 'รายชื่อคน',
  SHEET_COLUMNS: 'ตั้งค่าช่อง',
  SHEET_SUMMARY: 'สรุป',
  SHEET_LOG: 'ประวัติการแก้ไข',
  ROOT_FOLDER: 'รายการของ & รายชื่อ (ข้อมูลระบบ)',
  IMAGE_FOLDER: 'รูปสินค้า',
  BACKUP_FOLDER: 'สำรองข้อมูล',
  DEFAULT_COLUMNS: [{ label: 'ส่งเงินแล้ว', type: 'check' }],
  MAX_LOG_ROWS: 1000,
};

const ITEM_HEADERS = ['ID', 'รูป', 'ชื่อสินค้า', 'ราคา/หน่วย', 'จำนวน', 'รวม (บาท)', 'นับในยอด',
  'หมวดหมู่', 'หมายเหตุ', 'Image File ID', 'Image URL', 'อัปเดตล่าสุด'];
const COLUMN_HEADERS = ['ID', 'ชื่อช่อง', 'ชนิด (check/text/number)', 'ลำดับ'];
const LOG_HEADERS = ['เวลา', 'ผู้ใช้', 'ส่วน', 'รายละเอียด'];

/* ------------------------------------------------------------------ */
/*  API สำหรับหน้าเว็บ (GitHub Pages)                                   */
/* ------------------------------------------------------------------ */
const API_ACTIONS = {
  ping: ping,
  getAllData: getAllData,
  getVersions: getVersions,
  saveItems: saveItems,
  savePeople: savePeople,
  uploadImage: uploadImage,
  getImageDataUrl: getImageDataUrl,
  deleteImage: deleteImage,
  createBackup: createBackup,
};
let CURRENT_USER = '';

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const pass = appPassword_();
    if (!pass || pass === 'CHANGE-ME' || pass.length < 6) {
      return json_({ ok: false, code: 'SETUP', error: 'ผู้ดูแลยังไม่ได้ตั้งรหัสผ่าน APP_PASSWORD ใน Apps Script (หรือรหัสสั้นกว่า 6 ตัว)' });
    }
    if (String(body.key || '').trim() !== pass) {
      Utilities.sleep(800); // ชะลอการเดารหัส
      return json_({ ok: false, code: 'AUTH', error: 'รหัสผ่านไม่ถูกต้อง' });
    }
    const fn = API_ACTIONS[body.action];
    if (!fn) return json_({ ok: false, error: 'ไม่รู้จักคำสั่ง: ' + body.action });
    CURRENT_USER = String(body.user || '').slice(0, 40);
    const result = fn.apply(null, Array.isArray(body.args) ? body.args : []);
    return json_({ ok: true, result: result });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

/** เปิดลิงก์ Web App ในเบราว์เซอร์เพื่อตรวจว่าระบบทำงาน */
function doGet() {
  return json_({ ok: true, app: CONFIG.APP_TITLE, message: 'API พร้อมใช้งาน — เปิดหน้าเว็บจาก GitHub Pages แล้วใส่ลิงก์นี้' });
}

function ping() {
  getSs_();
  return { app: CONFIG.APP_TITLE, time: new Date().toISOString() };
}

function appPassword_() {
  const p = PropertiesService.getScriptProperties().getProperty('APP_PASSWORD');
  return String(p || CONFIG.APP_PASSWORD || '').trim();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** รันครั้งแรกครั้งเดียวจากหน้า Apps Script เพื่อสร้างชีต/โฟลเดอร์และขอสิทธิ์ */
function setup() {
  const ss = getSs_();
  itemsSheet_();
  const cols = readColumns_();
  const people = readPeople_(cols);
  writePeople_(cols, people);
  logSheet_();
  summarySheet_();
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1') || ss.getSheetByName('แผ่นงาน1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  imageFolder_();
  backupFolder_();
  const pass = appPassword_();
  if (pass === 'CHANGE-ME' || pass.length < 6) {
    Logger.log('⚠️ ยังไม่ได้ตั้งรหัสผ่าน: ไปที่ การตั้งค่าโปรเจ็กต์ (รูปเฟือง) > พร็อพเพอร์ตี้ของสคริปต์ > เพิ่ม APP_PASSWORD (อย่างน้อย 6 ตัว)');
  } else {
    Logger.log('🔒 ตั้งรหัสผ่านแล้ว');
  }
  Logger.log('✅ ติดตั้งเสร็จแล้ว');
  Logger.log('Google Sheet: ' + ss.getUrl());
  Logger.log('โฟลเดอร์ Drive: ' + rootFolder_().getUrl());
}

/* ------------------------------------------------------------------ */
/*  API ที่หน้าเว็บเรียกใช้                                              */
/* ------------------------------------------------------------------ */
function getAllData() {
  const cols = readColumns_();
  return {
    items: readItems_(),
    columns: cols,
    people: readPeople_(cols),
    versions: getVersions(),
    meta: {
      sheetUrl: getSs_().getUrl(),
      folderUrl: rootFolder_().getUrl(),
      user: userEmail_(),
    },
  };
}

function getVersions() {
  return { items: getVersion_('ITEMS'), people: getVersion_('PEOPLE') };
}

function saveItems(items, baseVersion, notes) {
  return withLock_(function () {
    const cur = getVersion_('ITEMS');
    if (baseVersion !== null && baseVersion !== undefined && Number(baseVersion) !== cur) {
      return { conflict: true, items: readItems_(), version: cur };
    }
    writeItems_(items || []);
    const v = bumpVersion_('ITEMS');
    log_('รายการของ', notes);
    return { ok: true, version: v };
  });
}

function savePeople(columns, people, baseVersion, notes) {
  return withLock_(function () {
    const cur = getVersion_('PEOPLE');
    if (baseVersion !== null && baseVersion !== undefined && Number(baseVersion) !== cur) {
      const cols = readColumns_();
      return { conflict: true, columns: cols, people: readPeople_(cols), version: cur };
    }
    writeColumns_(columns || []);
    writePeople_(columns || [], people || []);
    const v = bumpVersion_('PEOPLE');
    log_('รายชื่อคน', notes);
    return { ok: true, version: v };
  });
}

/** รับรูปแบบ dataURL (base64) แล้วบันทึกลง Drive */
function uploadImage(dataUrl, name) {
  const m = String(dataUrl).match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!m) throw new Error('ไฟล์รูปไม่ถูกต้อง');
  const mime = m[1];
  const ext = mime.split('/')[1].replace('jpeg', 'jpg');
  const safeName = String(name || 'สินค้า').replace(/[\\/:*?"<>|]/g, '').slice(0, 60);
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), mime, safeName + '_' + stamp + '.' + ext);
  const file = imageFolder_().createFile(blob);
  let shared = true;
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    shared = false; // บางองค์กรไม่อนุญาตแชร์สาธารณะ → หน้าเว็บจะโหลดรูปผ่านสคริปต์แทน
  }
  return { id: file.getId(), url: imageUrl_(file.getId()), shared: shared };
}

/** ใช้เมื่อเบราว์เซอร์โหลดรูปจากลิงก์ไม่ได้ */
function getImageDataUrl(fileId) {
  const file = DriveApp.getFileById(fileId);
  if (!isInFolder_(file, imageFolder_())) throw new Error('ไม่อนุญาต');
  const blob = file.getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

/** ย้ายรูปไปถังขยะ (กู้คืนได้ใน Drive ภายใน 30 วัน) */
function deleteImage(fileId) {
  if (!fileId) return false;
  try {
    const file = DriveApp.getFileById(fileId);
    if (!isInFolder_(file, imageFolder_())) return false;
    file.setTrashed(true);
    return true;
  } catch (e) {
    return false;
  }
}

/** คัดลอกทั้งไฟล์ Google Sheet ไปไว้ในโฟลเดอร์สำรองข้อมูล */
function createBackup() {
  const ss = getSs_();
  const name = ss.getName() + ' - สำรอง ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const copy = DriveApp.getFileById(ss.getId()).makeCopy(name, backupFolder_());
  log_('สำรองข้อมูล', [name]);
  return { name: name, url: copy.getUrl() };
}

/** (ทางเลือก) ตั้งให้สำรองข้อมูลอัตโนมัติทุกวันตอนตี 2 — รันครั้งเดียว */
function setupDailyBackup() {
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return t.getHandlerFunction() === 'createBackup'; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('createBackup').timeBased().everyDays(1).atHour(2).create();
  Logger.log('ตั้งสำรองข้อมูลอัตโนมัติทุกวันแล้ว');
}

/** ถ้ามีคนแก้ใน Google Sheet โดยตรง ให้หน้าเว็บรู้ว่าต้องโหลดใหม่ */
function onEdit(e) {
  try {
    const name = e.range.getSheet().getName();
    if (name === CONFIG.SHEET_ITEMS) bumpVersion_('ITEMS');
    if (name === CONFIG.SHEET_PEOPLE || name === CONFIG.SHEET_COLUMNS) bumpVersion_('PEOPLE');
  } catch (err) { /* ไม่เป็นไร */ }
}

/* ------------------------------------------------------------------ */
/*  รายการของ                                                          */
/* ------------------------------------------------------------------ */
function itemsSheet_() {
  return sheet_(CONFIG.SHEET_ITEMS, ITEM_HEADERS, function (sh) {
    sh.setColumnWidth(1, 90);
    sh.setColumnWidth(2, 70);
    sh.setColumnWidth(3, 220);
    sh.hideColumns(10, 2);
  });
}

function readItems_() {
  const sh = itemsSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, ITEM_HEADERS.length).getValues();
  return vals
    .filter(function (r) { return r[0] !== '' || r[2] !== ''; })
    .map(function (r) {
      return {
        id: r[0] ? String(r[0]) : newId_('i'),
        name: str_(r[2]),
        price: num_(r[3]),
        qty: r[4] === '' ? 1 : num_(r[4]),
        active: r[6] === '' ? true : bool_(r[6]),
        category: str_(r[7]),
        note: str_(r[8]),
        imageId: str_(r[9]),
        imageUrl: str_(r[10]),
      };
    });
}

function writeItems_(items) {
  const sh = itemsSheet_();
  const n = ITEM_HEADERS.length;
  const last = sh.getLastRow();
  if (last > 1) {
    const old = sh.getRange(2, 1, last - 1, n);
    old.clearContent();
    old.clearDataValidations();
  }
  if (!items.length) return;
  const now = new Date();
  const rows = items.map(function (it, i) {
    const r = i + 2;
    const url = /^https:\/\/[^"\s]+$/.test(it.imageUrl || '') ? it.imageUrl : '';
    const imgId = /^[\w-]+$/.test(it.imageId || '') ? it.imageId : '';
    return [
      it.id || newId_('i'),
      url ? '=IMAGE("' + url + '")' : '',
      safe_(it.name),
      num_(it.price),
      num_(it.qty),
      '=D' + r + '*E' + r,
      !!it.active,
      safe_(it.category),
      safe_(it.note),
      imgId,
      url,
      now,
    ];
  });
  sh.getRange(2, 1, rows.length, n).setValues(rows);
  sh.getRange(2, 7, rows.length, 1)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  sh.getRange(2, 4, rows.length, 1).setNumberFormat('#,##0.00');
  sh.getRange(2, 6, rows.length, 1).setNumberFormat('#,##0.00');
  sh.getRange(2, 12, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  sh.setRowHeights(2, rows.length, 56);
}

/* ------------------------------------------------------------------ */
/*  รายชื่อคน + ช่องหลังชื่อ                                            */
/* ------------------------------------------------------------------ */
function columnsSheet_() {
  return sheet_(CONFIG.SHEET_COLUMNS, COLUMN_HEADERS, function (sh) {
    const rows = CONFIG.DEFAULT_COLUMNS.map(function (c, i) { return [newId_('c'), c.label, c.type, i + 1]; });
    if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
    sh.getRange(1, 6).setValue('หมายเหตุ: แนะนำให้เพิ่ม/ลบช่องจากหน้าเว็บ');
  });
}

function peopleSheet_() {
  return sheet_(CONFIG.SHEET_PEOPLE, ['ID', 'ชื่อ']);
}

function readColumns_() {
  const sh = columnsSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 4).getValues()
    .filter(function (r) { return r[1] !== ''; })
    .map(function (r, i) {
      const t = String(r[2]).toLowerCase();
      return {
        id: r[0] ? String(r[0]) : newId_('c'),
        label: str_(r[1]),
        type: (t === 'text' || t === 'number') ? t : 'check',
        order: r[3] === '' ? i : Number(r[3]),
      };
    })
    .sort(function (a, b) { return a.order - b.order; })
    .map(function (c) { return { id: c.id, label: c.label, type: c.type }; });
}

function writeColumns_(cols) {
  const sh = columnsSheet_();
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 4).clearContent();
  if (!cols.length) return;
  sh.getRange(2, 1, cols.length, 4).setValues(cols.map(function (c, i) {
    return [c.id, safe_(c.label), c.type, i + 1];
  }));
}

function readPeople_(cols) {
  const sh = peopleSheet_();
  const lastR = sh.getLastRow();
  const lastC = sh.getLastColumn();
  if (lastR < 2 || lastC < 2) return [];
  const head = sh.getRange(1, 1, 1, lastC);
  const labels = head.getValues()[0].map(String);
  const notes = head.getNotes()[0];
  const idx = {};
  cols.forEach(function (c) {
    let i = notes.indexOf(c.id);          // จับคู่ด้วย ID ที่ซ่อนในโน้ตหัวตาราง
    if (i < 2) i = labels.indexOf(c.label); // ถ้าไม่มี ใช้ชื่อหัวตาราง
    if (i >= 2) idx[c.id] = i;
  });
  return sh.getRange(2, 1, lastR - 1, lastC).getValues()
    .filter(function (r) { return r[0] !== '' || r[1] !== ''; })
    .map(function (r) {
      const values = {};
      cols.forEach(function (c) {
        const v = idx[c.id] === undefined ? '' : r[idx[c.id]];
        if (c.type === 'check') values[c.id] = bool_(v);
        else if (c.type === 'number') values[c.id] = v === '' ? '' : num_(v);
        else values[c.id] = str_(v);
      });
      return { id: r[0] ? String(r[0]) : newId_('p'), name: str_(r[1]), values: values };
    });
}

function writePeople_(cols, people) {
  const sh = peopleSheet_();
  const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  all.clearContent();
  all.clearDataValidations();
  all.clearNote();
  all.clearFormat();

  const headers = ['ID', 'ชื่อ'].concat(cols.map(function (c) { return c.label; }), ['อัปเดตล่าสุด']);
  const notes = ['', ''].concat(cols.map(function (c) { return c.id; }), ['']);
  if (sh.getMaxColumns() < headers.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
  }
  const head = sh.getRange(1, 1, 1, headers.length);
  head.setValues([headers.map(safe_)]);
  head.setNotes([notes]);
  styleHeader_(sh, headers.length);
  sh.setFrozenColumns(2);

  if (!people.length) return;
  const now = new Date();
  const rows = people.map(function (p) {
    const v = p.values || {};
    return [p.id || newId_('p'), safe_(p.name)].concat(cols.map(function (c) {
      const x = v[c.id];
      if (c.type === 'check') return !!x;
      if (c.type === 'number') return (x === '' || x === null || x === undefined) ? '' : num_(x);
      return safe_(x);
    }), [now]);
  });
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  const cb = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  cols.forEach(function (c, i) {
    const rg = sh.getRange(2, i + 3, rows.length, 1);
    if (c.type === 'check') rg.setDataValidation(cb);
    if (c.type === 'number') rg.setNumberFormat('#,##0.##');
  });
  sh.getRange(2, headers.length, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
}

/* ------------------------------------------------------------------ */
/*  ชีตสรุป + ประวัติ                                                   */
/* ------------------------------------------------------------------ */
function summarySheet_() {
  return sheet_(CONFIG.SHEET_SUMMARY, null, function (sh) {
    const it = "'" + CONFIG.SHEET_ITEMS + "'";
    const pp = "'" + CONFIG.SHEET_PEOPLE + "'";
    sh.getRange(1, 1, 5, 2).setValues([
      ['ยอดรวม (เฉพาะที่นับในยอด)', '=IFERROR(SUMPRODUCT(' + it + '!D2:D, ' + it + '!E2:E, --(' + it + '!G2:G=TRUE)), 0)'],
      ['จำนวนรายการที่นับ', '=COUNTIF(' + it + '!G2:G, TRUE)'],
      ['ยอดที่เอาออก (ไม่นับ)', '=IFERROR(SUMPRODUCT(' + it + '!D2:D, ' + it + '!E2:E, --(' + it + '!G2:G=FALSE)), 0)'],
      ['จำนวนคน', '=COUNTA(' + pp + '!B2:B)'],
      ['หารต่อคน', '=IFERROR(B1/B4, 0)'],
    ]);
    sh.getRange(1, 1, 5, 1).setFontWeight('bold');
    sh.getRange('B1').setNumberFormat('#,##0.00');
    sh.getRange('B3').setNumberFormat('#,##0.00');
    sh.getRange('B5').setNumberFormat('#,##0.00');
    sh.setColumnWidth(1, 240);
    sh.setColumnWidth(2, 140);
  });
}

function logSheet_() {
  return sheet_(CONFIG.SHEET_LOG, LOG_HEADERS);
}

function log_(section, notes) {
  try {
    const list = (notes || []).filter(Boolean);
    const detail = list.length ? list.slice(0, 20).join(' | ') + (list.length > 20 ? ' …' : '') : 'บันทึกข้อมูล';
    const sh = logSheet_();
    sh.appendRow([new Date(), userEmail_(), section, safe_(detail)]);
    const extra = sh.getLastRow() - 1 - CONFIG.MAX_LOG_ROWS;
    if (extra > 0) sh.deleteRows(2, extra);
  } catch (e) { /* ไม่ให้การบันทึกประวัติทำให้งานหลักล้ม */ }
}

/* ------------------------------------------------------------------ */
/*  Helpers                                                            */
/* ------------------------------------------------------------------ */
function getSs_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* สร้างใหม่ด้านล่าง */ }
  }
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (!ss) ss = SpreadsheetApp.create('ฐานข้อมูล - ' + CONFIG.APP_TITLE);
  props.setProperty('SPREADSHEET_ID', ss.getId());
  return ss;
}

function sheet_(name, headers, onCreate) {
  const ss = getSs_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      styleHeader_(sh, headers.length);
    }
    if (onCreate) onCreate(sh);
  }
  return sh;
}

function styleHeader_(sh, n) {
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#1d6b57').setFontColor('#ffffff');
}

function folderByProp_(key, name, parent) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(key);
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (e) { /* สร้างใหม่ */ }
  }
  const f = parent ? parent.createFolder(name) : DriveApp.createFolder(name);
  props.setProperty(key, f.getId());
  return f;
}
function rootFolder_() { return folderByProp_('ROOT_FOLDER_ID', CONFIG.ROOT_FOLDER, null); }
function imageFolder_() { return folderByProp_('IMAGE_FOLDER_ID', CONFIG.IMAGE_FOLDER, rootFolder_()); }
function backupFolder_() { return folderByProp_('BACKUP_FOLDER_ID', CONFIG.BACKUP_FOLDER, rootFolder_()); }

function isInFolder_(file, folder) {
  const it = file.getParents();
  while (it.hasNext()) if (it.next().getId() === folder.getId()) return true;
  return false;
}

function imageUrl_(id) { return 'https://drive.google.com/thumbnail?id=' + id + '&sz=w800'; }

function getVersion_(key) {
  return Number(PropertiesService.getScriptProperties().getProperty('VER_' + key) || 0);
}
function bumpVersion_(key) {
  const v = Date.now();
  PropertiesService.getScriptProperties().setProperty('VER_' + key, String(v));
  return v;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) throw new Error('ระบบกำลังบันทึกข้อมูลของคนอื่นอยู่ กรุณาลองใหม่');
  try { return fn(); } finally { lock.releaseLock(); }
}

function userEmail_() {
  if (CURRENT_USER) return CURRENT_USER;
  try { return Session.getActiveUser().getEmail() || 'ไม่ระบุชื่อ'; } catch (e) { return 'ไม่ระบุชื่อ'; }
}

function newId_(prefix) {
  return prefix + '_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

function num_(v) {
  const n = Number(String(v === null || v === undefined ? '' : v).replace(/,/g, ''));
  return isFinite(n) ? n : 0;
}
function bool_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function str_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  return v === null || v === undefined ? '' : String(v);
}
/** กันข้อความที่ขึ้นต้นด้วย = + - @ ไม่ให้กลายเป็นสูตรในชีต */
function safe_(v) {
  const s = str_(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
