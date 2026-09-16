/**
 * ระบบ "สต็อกสินค้า & รายชื่อคน (ส่งเงินรายสัปดาห์)"
 * ฝั่งเซิร์ฟเวอร์ (Google Apps Script) — เป็น API ให้หน้าเว็บบน GitHub Pages
 *  - เก็บข้อมูลใน Google Sheet
 *  - เก็บรูปสินค้าใน Google Drive
 *
 * วิธีติดตั้งดูใน README.md
 */

const CONFIG = {
  APP_TITLE: 'รายการของ & รายชื่อ',
  SHEET_ITEMS: 'รายการของ',
  SHEET_HISTORY: 'บันทึกรายวัน',
  SHEET_PEOPLE: 'รายชื่อคน',
  SHEET_PAY_HISTORY: 'บันทึกเงินรายสัปดาห์',
  SHEET_SUMMARY: 'สรุป',
  SHEET_LOG: 'ประวัติการแก้ไข',
  ROOT_FOLDER: 'รายการของ & รายชื่อ (ข้อมูลระบบ)',
  IMAGE_FOLDER: 'รูปสินค้า',
  BACKUP_FOLDER: 'สำรองข้อมูล',
  MAX_LOG_ROWS: 1000,
  HISTORY_SCAN_ROWS: 3000, // ค้นหาแถวล่าสุดในชีตบันทึกย้อนหลังกี่แถว
  WEEK_START: 1,           // วันเริ่มสัปดาห์ 1 = จันทร์ ... 7 = อาทิตย์
};

// คอลัมน์ในชีต "รายการของ"
const ITEM_HEADERS = ['ID', 'รูป', 'ชื่อสินค้า', 'หมวดหมู่', 'ยอดยกมา', 'วันนี้เพิ่ม', 'วันนี้เอาออก',
  'รวมยอด', 'วันที่ (ของช่องวันนี้)', 'หมายเหตุ', 'Image File ID', 'Image URL', 'อัปเดตล่าสุด'];
const IC = { id: 0, img: 1, name: 2, cat: 3, base: 4, inn: 5, out: 6, total: 7, day: 8, note: 9, imgId: 10, imgUrl: 11, updated: 12 };
const HIST_HEADERS = ['วันที่', 'Item ID', 'ชื่อสินค้า', 'ยอดยกมา', 'เพิ่ม', 'เอาออก', 'คงเหลือ', 'อัปเดตล่าสุด'];
// คอลัมน์ในชีต "รายชื่อคน"
const PEOPLE_HEADERS = ['ID', 'ชื่อ', 'เบอร์โทร', 'ยอดยกมา (บาท)', 'สัปดาห์นี้ส่งแล้ว (บาท)', 'ยอดรวมที่ส่งมา (บาท)',
  'สัปดาห์ที่เริ่ม', 'อัปเดตล่าสุด'];
const PC = { id: 0, name: 1, phone: 2, base: 3, week: 4, total: 5, weekKey: 6, updated: 7 };
const PAY_HEADERS = ['สัปดาห์ที่เริ่ม', 'Person ID', 'ชื่อ', 'ยอดยกมา', 'ส่งสัปดาห์นี้', 'ยอดรวม', 'อัปเดตล่าสุด'];
const LOG_HEADERS = ['เวลา', 'ผู้ใช้', 'ส่วน', 'รายละเอียด'];

/* ------------------------------------------------------------------ */
/*  API สำหรับหน้าเว็บ                                                  */
/* ------------------------------------------------------------------ */
const API_ACTIONS = {
  ping: ping,
  getAllData: getAllData,
  getVersions: getVersions,
  saveItems: saveItems,
  savePeople: savePeople,
  getHistory: getHistory,
  getPayHistory: getPayHistory,
  uploadImage: uploadImage,
  getImageDataUrl: getImageDataUrl,
  deleteImage: deleteImage,
  createBackup: createBackup,
};
let CURRENT_USER = '';

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
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
  return json_({ ok: true, app: CONFIG.APP_TITLE, message: 'API พร้อมใช้งาน — นำลิงก์นี้ไปใส่ในหน้าเว็บ' });
}

function ping() {
  getSs_();
  return { app: CONFIG.APP_TITLE, today: today_() };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** รันครั้งแรกจากหน้า Apps Script เพื่อสร้างชีต/โฟลเดอร์และขอสิทธิ์ (รันซ้ำได้ ไม่ลบข้อมูล) */
function setup() {
  const ss = getSs_();
  itemsSheet_();
  historySheet_();
  peopleSheet_();
  payHistorySheet_();
  logSheet_();
  refreshSummary_();
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1') || ss.getSheetByName('แผ่นงาน1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
  imageFolder_();
  backupFolder_();
  Logger.log('✅ ติดตั้งเสร็จแล้ว (วันที่ของระบบ: ' + today_() + ', สัปดาห์เริ่ม: ' + weekKey_() + ')');
  Logger.log('Google Sheet: ' + ss.getUrl());
  Logger.log('โฟลเดอร์ Drive: ' + rootFolder_().getUrl());
}

function getAllData() {
  let items = readItems_();
  // ขึ้นวันใหม่ → เก็บประวัติของวันก่อน แล้วยกยอดมา
  if (items.some(function (it) { return it.day !== today_(); })) {
    withLock_(function () {
      const fresh = readItems_();
      if (fresh.some(function (it) { return it.day !== today_(); })) {
        upsertHistory_(fresh);
        fresh.forEach(rollover_);
        writeItems_(fresh);
        bumpVersion_('ITEMS');
      }
      items = fresh;
    });
  }
  let people = readPeople_();
  if (people.some(function (p) { return p.weekKey !== weekKey_(); })) {
    withLock_(function () {
      const fresh = readPeople_();
      if (fresh.some(function (p) { return p.weekKey !== weekKey_(); })) {
        upsertPayHistory_(fresh);
        fresh.forEach(rolloverPerson_);
        writePeople_(fresh);
        bumpVersion_('PEOPLE');
      }
      people = fresh;
    });
  }
  return {
    items: items,
    people: people,
    versions: getVersions(),
    meta: {
      today: today_(),
      week: weekKey_(),
      sheetUrl: getSs_().getUrl(),
      folderUrl: rootFolder_().getUrl(),
    },
  };
}

function getVersions() {
  return { items: getVersion_('ITEMS'), people: getVersion_('PEOPLE'), today: today_(), week: weekKey_() };
}

function saveItems(items, baseVersion, notes) {
  return withLock_(function () {
    const cur = getVersion_('ITEMS');
    if (baseVersion !== null && baseVersion !== undefined && Number(baseVersion) !== cur) {
      return { conflict: true, items: readItems_(), version: cur };
    }
    const list = (items || []).map(normalizeItem_);
    upsertHistory_(list);
    list.forEach(rollover_);
    writeItems_(list);
    const v = bumpVersion_('ITEMS');
    log_('รายการของ', notes);
    return { ok: true, version: v };
  });
}

function savePeople(people, baseVersion, notes) {
  return withLock_(function () {
    const cur = getVersion_('PEOPLE');
    if (baseVersion !== null && baseVersion !== undefined && Number(baseVersion) !== cur) {
      return { conflict: true, people: readPeople_(), version: cur };
    }
    const list = (people || []).map(normalizePerson_);
    upsertPayHistory_(list);
    list.forEach(rolloverPerson_);
    writePeople_(list);
    const v = bumpVersion_('PEOPLE');
    log_('รายชื่อคน', notes);
    return { ok: true, version: v };
  });
}

/** ประวัติรายวันของสินค้า 1 รายการ (ใหม่สุดก่อน) */
function getHistory(itemId, limit) {
  const sh = historySheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, HIST_HEADERS.length).getValues();
  const out = [];
  for (let i = vals.length - 1; i >= 0 && out.length < (limit || 60); i--) {
    const r = vals[i];
    if (String(r[1]) !== String(itemId)) continue;
    out.push({ date: dayStr_(r[0]), base: num_(r[3]), inn: num_(r[4]), out: num_(r[5]), total: num_(r[6]) });
  }
  out.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; });
  return out;
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

function createBackup() {
  const ss = getSs_();
  const name = ss.getName() + ' - สำรอง ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const copy = DriveApp.getFileById(ss.getId()).makeCopy(name, backupFolder_());
  log_('สำรองข้อมูล', [name]);
  return { name: name, url: copy.getUrl() };
}

/** (ทางเลือก) ตั้งทริกเกอร์รายวัน: ตีหนึ่ง ยกยอด + เก็บประวัติ, ตีสอง สำรองข้อมูล — รันครั้งเดียว */
function setupDailyJobs() {
  ScriptApp.getProjectTriggers()
    .filter(function (t) { return ['createBackup', 'dailyRollover'].indexOf(t.getHandlerFunction()) >= 0; })
    .forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyRollover').timeBased().everyDays(1).atHour(0).nearMinute(5).create();
  ScriptApp.newTrigger('createBackup').timeBased().everyDays(1).atHour(2).create();
  Logger.log('ตั้งงานรายวันแล้ว');
}
function setupDailyBackup() { setupDailyJobs(); }

function dailyRollover() { getAllData(); } // ยกยอดสินค้า (รายวัน) และยอดเงิน (รายสัปดาห์)

/** ถ้ามีคนแก้ใน Google Sheet โดยตรง ให้หน้าเว็บรู้ว่าต้องโหลดใหม่ */
function onEdit(e) {
  try {
    const name = e.range.getSheet().getName();
    if (name === CONFIG.SHEET_ITEMS) bumpVersion_('ITEMS');
    if (name === CONFIG.SHEET_PEOPLE) bumpVersion_('PEOPLE');
  } catch (err) { /* ไม่เป็นไร */ }
}

/* ------------------------------------------------------------------ */
/*  สินค้า                                                             */
/* ------------------------------------------------------------------ */
let ITEMS_CHECKED_ = false;
function itemsSheet_() {
  const sh = sheet_(CONFIG.SHEET_ITEMS, ITEM_HEADERS, function (s) {
    s.setColumnWidth(1, 90);
    s.setColumnWidth(2, 70);
    s.setColumnWidth(3, 220);
    s.hideColumns(IC.imgId + 1, 2);
  });
  if (!ITEMS_CHECKED_) {
    ITEMS_CHECKED_ = true;
    migrateItems_(sh);
  }
  return sh;
}

/** แปลงชีตรูปแบบเดิม (ราคา/จำนวน) เป็นรูปแบบใหม่ — จำนวนเดิมจะกลายเป็น "ยอดยกมา" */
function migrateItems_(sh) {
  const lastC = Math.max(sh.getLastColumn(), 1);
  const head = sh.getRange(1, 1, 1, lastC).getValues()[0].map(String);
  if (head[3] === 'ราคา/หน่วย' && head[4] === 'จำนวน') {
    const last = sh.getLastRow();
    const rows = last > 1 ? sh.getRange(2, 1, last - 1, Math.max(lastC, 12)).getValues() : [];
    const items = rows.filter(function (r) { return r[0] !== '' || r[2] !== ''; }).map(function (r) {
      const note = [str_(r[8]), num_(r[3]) ? 'ราคาเดิม ' + num_(r[3]) : ''].filter(Boolean).join(' · ');
      return {
        id: r[0] ? String(r[0]) : newId_('i'), name: str_(r[2]), category: str_(r[7]),
        base: r[4] === '' ? 0 : num_(r[4]), inToday: 0, outToday: 0, day: today_(),
        note: note, imageId: str_(r[9]), imageUrl: str_(r[10]),
      };
    });
    // เก็บสำเนาชีตเดิมไว้ก่อน
    const ss = getSs_();
    const copy = sh.copyTo(ss);
    const oldName = CONFIG.SHEET_ITEMS + ' (แบบเดิม)';
    copy.setName(ss.getSheetByName(oldName) ? oldName + ' ' + Date.now() : oldName);
    const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
    all.clearContent(); all.clearDataValidations(); all.clearFormat();
    try { sh.showColumns(1, sh.getMaxColumns()); } catch (e) {}
    sh.getRange(1, 1, 1, ITEM_HEADERS.length).setValues([ITEM_HEADERS]);
    styleHeader_(sh, ITEM_HEADERS.length);
    sh.hideColumns(IC.imgId + 1, 2);
    writeItems_(items);
    refreshSummary_();
    bumpVersion_('ITEMS');
    log_('ระบบ', ['แปลงชีตรายการของเป็นแบบยอดยกมา/เพิ่ม/เอาออก (' + items.length + ' รายการ)']);
  } else if (head[0] === 'ID' && head.join('|') !== ITEM_HEADERS.join('|') && sh.getLastRow() <= 1) {
    sh.getRange(1, 1, 1, ITEM_HEADERS.length).setValues([ITEM_HEADERS]);
    styleHeader_(sh, ITEM_HEADERS.length);
  }
}

function normalizeItem_(it) {
  it = it || {};
  return {
    id: String(it.id || newId_('i')),
    name: str_(it.name),
    category: str_(it.category),
    base: num_(it.base),
    inToday: Math.max(0, num_(it.inToday)),
    outToday: Math.max(0, num_(it.outToday)),
    day: /^\d{4}-\d{2}-\d{2}$/.test(it.day || '') ? it.day : today_(),
    note: str_(it.note),
    imageId: /^[\w-]+$/.test(it.imageId || '') ? it.imageId : '',
    imageUrl: /^https:\/\/[^"\s]+$/.test(it.imageUrl || '') ? it.imageUrl : '',
  };
}

function rollover_(it) {
  if (it.day === today_()) return;
  it.base = round2_(it.base + it.inToday - it.outToday);
  it.inToday = 0;
  it.outToday = 0;
  it.day = today_();
}

function readItems_() {
  const sh = itemsSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, ITEM_HEADERS.length).getValues()
    .filter(function (r) { return r[IC.id] !== '' || r[IC.name] !== ''; })
    .map(function (r) {
      return normalizeItem_({
        id: r[IC.id] ? String(r[IC.id]) : newId_('i'),
        name: str_(r[IC.name]),
        category: str_(r[IC.cat]),
        base: r[IC.base],
        inToday: r[IC.inn],
        outToday: r[IC.out],
        day: dayStr_(r[IC.day]) || today_(),
        note: str_(r[IC.note]),
        imageId: str_(r[IC.imgId]),
        imageUrl: str_(r[IC.imgUrl]),
      });
    });
}

function writeItems_(items) {
  const sh = itemsSheet_();
  const n = ITEM_HEADERS.length;
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, n).clearContent();
  if (!items.length) return;
  const now = new Date();
  const rows = items.map(function (it, i) {
    const r = i + 2;
    const row = [];
    row[IC.id] = it.id;
    row[IC.img] = it.imageUrl ? '=IMAGE("' + it.imageUrl + '")' : '';
    row[IC.name] = safe_(it.name);
    row[IC.cat] = safe_(it.category);
    row[IC.base] = num_(it.base);
    row[IC.inn] = num_(it.inToday);
    row[IC.out] = num_(it.outToday);
    row[IC.total] = '=E' + r + '+F' + r + '-G' + r;
    row[IC.day] = it.day;
    row[IC.note] = safe_(it.note);
    row[IC.imgId] = it.imageId || '';
    row[IC.imgUrl] = it.imageUrl || '';
    row[IC.updated] = now;
    return row;
  });
  sh.getRange(2, IC.day + 1, rows.length, 1).setNumberFormat('@');
  sh.getRange(2, 1, rows.length, n).setValues(rows);
  sh.getRange(2, IC.base + 1, rows.length, 4).setNumberFormat('#,##0.##');
  sh.getRange(2, IC.total + 1, rows.length, 1).setFontWeight('bold');
  sh.getRange(2, IC.updated + 1, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  sh.setRowHeights(2, rows.length, 56);
}

/* ------------------------------------------------------------------ */
/*  บันทึกรายวัน                                                        */
/* ------------------------------------------------------------------ */
function historySheet_() {
  return sheet_(CONFIG.SHEET_HISTORY, HIST_HEADERS, function (s) {
    s.setColumnWidth(3, 200);
    s.getRange('A:A').setNumberFormat('@');
  });
}

/** เขียน/อัปเดตแถว (วันที่ + สินค้า) ด้วยค่าเพิ่ม/เอาออกล่าสุดของวันนั้น */
function upsertHistory_(items) {
  upsertRows_(historySheet_(), HIST_HEADERS.length, items.map(function (it) {
    return {
      key: it.day + '|' + it.id,
      moved: !!(it.inToday || it.outToday),
      row: [it.day, it.id, safe_(it.name), it.base, it.inToday, it.outToday,
        round2_(it.base + it.inToday - it.outToday), new Date()],
    };
  }));
}

/**
 * ตัวช่วยเขียนบันทึก: key = คอลัมน์ A|B, เทียบคอลัมน์ C..(n-1) ถ้าเหมือนเดิมไม่เขียนซ้ำ
 * แถวใหม่จะถูกเพิ่มเฉพาะเมื่อ moved = true
 */
function upsertRows_(sh, width, entries) {
  const last = sh.getLastRow();
  const scan = Math.min(Math.max(last - 1, 0), CONFIG.HISTORY_SCAN_ROWS);
  const start = last - scan + 1;
  const existing = {};
  let vals = [];
  if (scan > 0) {
    vals = sh.getRange(start, 1, scan, width).getValues();
    vals.forEach(function (r, i) { existing[dayStr_(r[0]) + '|' + r[1]] = i; });
  }
  const appends = [];
  entries.forEach(function (e) {
    const idx = existing[e.key];
    if (idx === undefined) {
      if (e.moved) appends.push(e.row);
      return;
    }
    const old = vals[idx];
    let same = true;
    for (let c = 2; c < width - 1; c++) {
      const a = e.row[c], b = old[c];
      if (typeof a === 'number' ? num_(b) !== a : String(b) !== String(a)) { same = false; break; }
    }
    if (!same) sh.getRange(start + idx, 1, 1, width).setValues([e.row]);
  });
  if (appends.length) {
    const r0 = sh.getLastRow() + 1;
    sh.getRange(r0, 1, appends.length, 1).setNumberFormat('@');
    sh.getRange(r0, 1, appends.length, width).setValues(appends);
    sh.getRange(r0, width, appends.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  }
}

/* ------------------------------------------------------------------ */
/*  รายชื่อคน + เงินรายสัปดาห์                                           */
/* ------------------------------------------------------------------ */
let PEOPLE_CHECKED_ = false;
function peopleSheet_() {
  const sh = sheet_(CONFIG.SHEET_PEOPLE, PEOPLE_HEADERS, function (s) {
    s.setColumnWidth(2, 180);
    s.setFrozenColumns(2);
  });
  if (!PEOPLE_CHECKED_) {
    PEOPLE_CHECKED_ = true;
    migratePeople_(sh);
  }
  return sh;
}

function payHistorySheet_() {
  return sheet_(CONFIG.SHEET_PAY_HISTORY, PAY_HEADERS, function (s) {
    s.setColumnWidth(3, 180);
    s.getRange('A:A').setNumberFormat('@');
  });
}

/** แปลงชีตรายชื่อแบบเดิม (ช่องหลังชื่อแบบกำหนดเอง) เป็นแบบใหม่ */
function migratePeople_(sh) {
  const lastC = Math.max(sh.getLastColumn(), 1);
  const head = sh.getRange(1, 1, 1, lastC).getValues()[0].map(String);
  if (head[0] !== 'ID' || head[1] !== 'ชื่อ') return;
  if (head.slice(0, PEOPLE_HEADERS.length).join('|') === PEOPLE_HEADERS.join('|')) return;
  const notes = sh.getRange(1, 1, 1, lastC).getNotes()[0];
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, lastC).getValues() : [];
  const phoneIdx = head.findIndex(function (h, i) { return i >= 2 && /เบอร์|โทร|phone/i.test(h); });
  const moneyIdx = head.findIndex(function (h, i) {
    return i >= 2 && /เงิน|โอน|จ่าย|บาท/.test(h) && rows.some(function (r) { return typeof r[i] === 'number'; });
  });
  const people = rows.filter(function (r) { return r[0] !== '' || r[1] !== ''; }).map(function (r) {
    return normalizePerson_({
      id: r[0] ? String(r[0]) : newId_('p'),
      name: str_(r[1]),
      phone: phoneIdx >= 0 ? str_(r[phoneIdx]) : '',
      base: moneyIdx >= 0 ? num_(r[moneyIdx]) : 0,
      week: 0,
      weekKey: weekKey_(),
    });
  });
  const ss = getSs_();
  const copy = sh.copyTo(ss);
  const oldName = CONFIG.SHEET_PEOPLE + ' (แบบเดิม)';
  copy.setName(ss.getSheetByName(oldName) ? oldName + ' ' + Date.now() : oldName);
  if (notes.some(Boolean)) copy.getRange(1, 1, 1, lastC).setNotes([notes]);
  writePeople_(people, sh);
  refreshSummary_();
  bumpVersion_('PEOPLE');
  log_('ระบบ', ['แปลงชีตรายชื่อเป็นแบบส่งเงินรายสัปดาห์ (' + people.length + ' คน)']);
}

function normalizePerson_(p) {
  p = p || {};
  return {
    id: String(p.id || newId_('p')),
    name: str_(p.name),
    phone: str_(p.phone),
    base: num_(p.base),
    week: Math.max(0, num_(p.week)),
    weekKey: /^\d{4}-\d{2}-\d{2}$/.test(p.weekKey || '') ? p.weekKey : weekKey_(),
  };
}

function rolloverPerson_(p) {
  if (p.weekKey === weekKey_()) return;
  p.base = round2_(p.base + p.week);
  p.week = 0;
  p.weekKey = weekKey_();
}

function readPeople_() {
  const sh = peopleSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, PEOPLE_HEADERS.length).getValues()
    .filter(function (r) { return r[PC.id] !== '' || r[PC.name] !== ''; })
    .map(function (r) {
      return normalizePerson_({
        id: r[PC.id] ? String(r[PC.id]) : newId_('p'),
        name: str_(r[PC.name]),
        phone: str_(r[PC.phone]),
        base: r[PC.base],
        week: r[PC.week],
        weekKey: dayStr_(r[PC.weekKey]) || weekKey_(),
      });
    });
}

function writePeople_(people, shOpt) {
  const sh = shOpt || peopleSheet_();
  const all = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());
  all.clearContent();
  all.clearDataValidations();
  all.clearNote();
  all.clearFormat();
  const n = PEOPLE_HEADERS.length;
  sh.getRange(1, 1, 1, n).setValues([PEOPLE_HEADERS]);
  styleHeader_(sh, n);
  sh.setFrozenColumns(2);
  if (!people.length) return;
  const now = new Date();
  const rows = people.map(function (p, i) {
    const r = i + 2;
    const row = [];
    row[PC.id] = p.id;
    row[PC.name] = safe_(p.name);
    row[PC.phone] = safe_(p.phone);
    row[PC.base] = num_(p.base);
    row[PC.week] = num_(p.week);
    row[PC.total] = '=D' + r + '+E' + r;
    row[PC.weekKey] = p.weekKey;
    row[PC.updated] = now;
    return row;
  });
  sh.getRange(2, PC.phone + 1, rows.length, 1).setNumberFormat('@');
  sh.getRange(2, PC.weekKey + 1, rows.length, 1).setNumberFormat('@');
  sh.getRange(2, 1, rows.length, n).setValues(rows);
  sh.getRange(2, PC.base + 1, rows.length, 3).setNumberFormat('#,##0.##');
  sh.getRange(2, PC.total + 1, rows.length, 1).setFontWeight('bold');
  sh.getRange(2, PC.updated + 1, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
}

function upsertPayHistory_(people) {
  upsertRows_(payHistorySheet_(), PAY_HEADERS.length, people.map(function (p) {
    return {
      key: p.weekKey + '|' + p.id,
      moved: !!p.week,
      row: [p.weekKey, p.id, safe_(p.name), p.base, p.week, round2_(p.base + p.week), new Date()],
    };
  }));
}

/** ประวัติการส่งเงินรายสัปดาห์ของ 1 คน (ใหม่สุดก่อน) */
function getPayHistory(personId, limit) {
  const sh = payHistorySheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, PAY_HEADERS.length).getValues();
  const out = [];
  for (let i = vals.length - 1; i >= 0 && out.length < (limit || 60); i--) {
    const r = vals[i];
    if (String(r[1]) !== String(personId)) continue;
    out.push({ week: dayStr_(r[0]), base: num_(r[3]), amount: num_(r[4]), total: num_(r[5]) });
  }
  out.sort(function (a, b) { return a.week < b.week ? 1 : a.week > b.week ? -1 : 0; });
  return out;
}

/* ------------------------------------------------------------------ */
/*  ชีตสรุป + ประวัติการแก้ไข                                           */
/* ------------------------------------------------------------------ */
function refreshSummary_() {
  const sh = sheet_(CONFIG.SHEET_SUMMARY, null);
  const it = "'" + CONFIG.SHEET_ITEMS + "'";
  const pp = "'" + CONFIG.SHEET_PEOPLE + "'";
  sh.getRange(1, 1, 12, 2).clearContent();
  sh.getRange(1, 1, 9, 2).setValues([
    ['จำนวนสินค้า', '=COUNTA(' + it + '!C2:C)'],
    ['รวมยอดคงเหลือทุกรายการ', '=SUM(' + it + '!H2:H)'],
    ['วันนี้เพิ่มรวม', '=SUM(' + it + '!F2:F)'],
    ['วันนี้เอาออกรวม', '=SUM(' + it + '!G2:G)'],
    ['สินค้าที่ยอดหมด (≤ 0)', '=COUNTIFS(' + it + '!C2:C, "<>", ' + it + '!H2:H, "<=0")'],
    ['จำนวนคน', '=COUNTA(' + pp + '!B2:B)'],
    ['สัปดาห์นี้ส่งแล้ว (คน)', '=COUNTIF(' + pp + '!E2:E, ">0")'],
    ['เงินที่ส่งสัปดาห์นี้ (บาท)', '=SUM(' + pp + '!E2:E)'],
    ['ยอดรวมเงินที่ส่งมาทั้งหมด (บาท)', '=SUM(' + pp + '!F2:F)'],
  ]);
  sh.getRange(1, 1, 9, 1).setFontWeight('bold');
  sh.getRange(1, 2, 9, 1).setNumberFormat('#,##0.##');
  sh.getRange(11, 1).setValue('หมายเหตุ: "วันนี้/สัปดาห์นี้" หมายถึงวันที่ในคอลัมน์ I ของชีตรายการของ และคอลัมน์ G ของชีตรายชื่อคน');
  sh.setColumnWidth(1, 240);
  sh.setColumnWidth(2, 140);
  return sh;
}

function logSheet_() {
  return sheet_(CONFIG.SHEET_LOG, LOG_HEADERS);
}

function log_(section, notes) {
  try {
    const list = (notes || []).filter(Boolean);
    const detail = list.length ? list.slice(0, 20).join(' | ') + (list.length > 20 ? ' …' : '') : 'บันทึกข้อมูล';
    const sh = logSheet_();
    sh.appendRow([new Date(), userName_(), section, safe_(detail)]);
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

function userName_() {
  if (CURRENT_USER) return CURRENT_USER;
  try { return Session.getActiveUser().getEmail() || 'ไม่ระบุชื่อ'; } catch (e) { return 'ไม่ระบุชื่อ'; }
}

function today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
/** วันที่ของวันแรกในสัปดาห์นี้ (ค่าเริ่มต้น: วันจันทร์) */
function weekKey_() {
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const dow = Number(Utilities.formatDate(now, tz, 'u')); // 1=จันทร์ … 7=อาทิตย์
  const back = (dow - CONFIG.WEEK_START + 7) % 7;
  return Utilities.formatDate(new Date(now.getTime() - back * 86400000), tz, 'yyyy-MM-dd');
}

function dayStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(v || '').trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}

function newId_(prefix) {
  return prefix + '_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}
function round2_(n) { return Math.round(n * 100) / 100; }
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
