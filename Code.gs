/**
 * Web App เก็บข้อมูลวัสดุ — ชีต "DT1.1.1 ข้อมูลดิบ" (โครงสร้างใหม่ 2026-06-19)
 *
 * คอลัมน์ (1-indexed):
 *   A(1) No.            [อ่าน]
 *   B(2) Material Code  [อ่าน]
 *   C(3) MainDept       [แก้ได้ • dropdown เดียว จากแท็บ Dept]
 *   D(4) SubDept        [แก้ได้ • หลายค่า จากแท็บ Dept]
 *   E(5) MainCat        [แก้ได้ • dropdown เดียว จากแท็บ MainCat]
 *   F(6) SubCat         [อ่าน]
 *   G(7) Material Name (1) [อ่าน]
 *   H(8) Material Name (2) [แก้ได้]
 *   I(9) Image          [แก้ได้ • อัปโหลดรูป]
 *   J(10) Spec          [แก้ได้]
 *   K(11) Brand         [แก้ได้]
 *   L(12) น้ำหนัก/ชิ้น (Kg) [แก้ได้ • ตัวเลข]
 *   M(13) กว้าง (mm)    [แก้ได้ • ตัวเลข]
 *   N(14) ยาว (mm)      [แก้ได้ • ตัวเลข]
 *   O(15) สูง (mm)      [แก้ได้ • ตัวเลข]
 *   P(16) หน่วยเก็บ (IC) [แก้ได้]
 *   Q(17) หน่วยซื้อ (PO) [แก้ได้]
 */

var SPREADSHEET_ID = '1FtoQo1BptJ3QnJxULZ-VZ130n0IKqNaV2XznCeo4kJE';
var SHEET_NAME = 'RawData1';
var DEPT_SHEET = 'Dept';
var MAINCAT_SHEET = 'MainCat';
var IMAGE_FOLDER_ID = '1CVDBgTgMbs9cHcC4yrJf2DKmrXD0bx4W';
var LOG_SHEET_NAME = '_EditLog';
var FIRST_DATA_ROW = 2;
var TOTAL_COLUMNS = 17;     // A..Q
var IMAGE_COL = 9;          // I
var NOTE_COL = 18;          // R — Note (กรอกได้)
var MARKER_COL = 19;        // S — marker "_web" (ทำเครื่องหมายแถวที่เพิ่มเองผ่านเว็บ เพื่อให้ลบได้เสมอ)

// ช่องที่แก้ไขได้ (key ฝั่งเว็บ → คอลัมน์ในชีต + ป้ายใน log + ชนิด)
var EDITABLE = [
  { key: 'mainDept',  col: 3,  label: 'MainDept (C)',           type: 'text' },
  { key: 'subDept',   col: 4,  label: 'SubDept (D)',            type: 'text' },
  { key: 'mainCat',   col: 5,  label: 'MainCat (E)',            type: 'text' },
  { key: 'name2',     col: 8,  label: 'Material Name (2) (H)',  type: 'text' },
  { key: 'image',     col: 9,  label: 'Image (I)',              type: 'text' },
  { key: 'spec',      col: 10, label: 'Spec (J)',               type: 'text' },
  { key: 'brand',     col: 11, label: 'Brand (K)',              type: 'text' },
  { key: 'weight',    col: 12, label: 'น้ำหนัก/ชิ้น (L)',       type: 'num'  },
  { key: 'width',     col: 13, label: 'กว้าง (M)',              type: 'num'  },
  { key: 'length',    col: 14, label: 'ยาว (N)',                type: 'num'  },
  { key: 'height',    col: 15, label: 'สูง (O)',                type: 'num'  },
  { key: 'unitStore', col: 16, label: 'หน่วยเก็บ (P)',          type: 'text' },
  { key: 'unitBuy',   col: 17, label: 'หน่วยซื้อ (Q)',          type: 'text' },
  { key: 'note',      col: 18, label: 'Note (R)',               type: 'text' }
];

/** ---------------- Web App ---------------- */
function doGet() {
  try { getLogSheet_(); } catch (e) {}
  try { var sh0 = getSheet_(); ensureMarkerCol_(sh0); migrateMarkerToS_(sh0); } catch (e) {}
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('BPI Product Data')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSheet_() {
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('ไม่พบชีต: ' + SHEET_NAME);
  return sh;
}

/** ให้แน่ใจว่ามีคอลัมน์ Note (R) + marker (S) พร้อมหัวตาราง */
function ensureMarkerCol_(sh) {
  if (sh.getMaxColumns() < MARKER_COL) sh.insertColumnsAfter(sh.getMaxColumns(), MARKER_COL - sh.getMaxColumns());
  var r1 = String(sh.getRange(1, NOTE_COL).getValue()).trim();
  if (r1 === '' || r1 === '_web') sh.getRange(1, NOTE_COL).setValue('Note');   // เดิม R เคยเป็น _web → เปลี่ยนเป็น Note
  if (!sh.getRange(1, MARKER_COL).getValue()) sh.getRange(1, MARKER_COL).setValue('_web');
}

/** ย้ายค่า marker เดิมจาก R(18) ไป S(19) ครั้งเดียว (ตอน R ยังเป็น marker) */
function migrateMarkerToS_(sh) {
  var lastRow = sh.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return;
  var rng = sh.getRange(FIRST_DATA_ROW, NOTE_COL, lastRow - FIRST_DATA_ROW + 1, 2); // R:S
  var vals = rng.getValues(), changed = false;
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === 'WEB' && String(vals[i][1]).trim() === '') {
      vals[i][1] = 'WEB'; vals[i][0] = ''; changed = true;   // S = marker, R = เคลียร์ (ไว้ใส่ Note)
    }
  }
  if (changed) rng.setValues(vals);
}

/** ---------------- อ่านข้อมูลวัสดุทั้งหมด ---------------- */
function getMaterials() {
  var sh = getSheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return [];

  var width = Math.min(MARKER_COL, sh.getMaxColumns());
  var v = sh.getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, width).getValues();
  var out = [];
  for (var i = 0; i < v.length; i++) {
    var r = v[i];
    if (r[0] === '' && r[1] === '') continue;   // ข้ามแถวว่าง (ไม่มีทั้ง No. และ Material Code)
    out.push({
      row: FIRST_DATA_ROW + i,
      manual: (width >= MARKER_COL && String(r[MARKER_COL - 1]).trim() !== ''),
      no: r[0],          // A
      code: r[1],        // B
      mainDept: r[2],    // C
      subDept: r[3],     // D
      mainCat: r[4],     // E
      subCat: r[5],      // F
      name1: r[6],       // G
      name2: r[7],       // H
      image: r[8],       // I
      spec: r[9],        // J
      brand: r[10],      // K
      weight: r[11],     // L
      width: r[12],      // M
      length: r[13],     // N
      height: r[14],     // O
      unitStore: r[15],  // P
      unitBuy: r[16],    // Q
      note: r[17]        // R
    });
  }
  return out;
}

/** ---------------- รายการตัวเลือก (dropdown) ---------------- */
function getLookups() {
  return {
    depts: readLookupColumn_(DEPT_SHEET, 2),     // Dept!B (แผนก)
    mainCats: readLookupColumn_(MAINCAT_SHEET, 2) // MainCat!B (หมวดหลัก)
  };
}

function readLookupColumn_(sheetName, col) {
  var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var vals = sh.getRange(2, col, lastRow - 1, 1).getValues();
  var seen = {}, out = [];
  for (var i = 0; i < vals.length; i++) {
    var s = (vals[i][0] === null || vals[i][0] === undefined) ? '' : String(vals[i][0]).trim();
    if (s !== '' && !seen[s]) { seen[s] = true; out.push(s); }
  }
  return out;
}

/** ---------------- บันทึก 1 แถว (auto-save) + log ---------------- */
function saveMaterialRow(payload) {
  if (!payload || !payload.row) throw new Error('ไม่มีข้อมูลแถวที่จะบันทึก');
  var row = Number(payload.row);
  if (row < FIRST_DATA_ROW) throw new Error('แถวไม่ถูกต้อง');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = getSheet_();
    if (row > sh.getLastRow()) throw new Error('แถวเกินขอบเขตข้อมูล');

    // อ่านช่วง C..R (3..18) มาก่อน — แก้เฉพาะช่อง editable, คง F/G (อ่านอย่างเดียว) ไว้ (ไม่แตะ marker S)
    var rng = sh.getRange(row, 3, 1, 16);
    var old = rng.getValues()[0];           // index 0 = C(3) ... 15 = R(18)
    var code = sh.getRange(row, 2).getValue();

    var when = new Date(), who = currentUser_(), logs = [];
    var arr = old.slice();
    for (var i = 0; i < EDITABLE.length; i++) {
      var f = EDITABLE[i];
      var off = f.col - 3;
      var nv = (f.type === 'num') ? toNumberOrText_(payload[f.key]) : toText_(payload[f.key]);
      if (String(old[off]) !== String(nv)) {
        logs.push([when, who, row, code, f.label, old[off], nv]);
        arr[off] = nv;
      }
    }

    if (!logs.length) return { ok: true, row: row, changes: 0, savedAt: nowString_() };

    rng.setValues([arr]);
    appendLogs_(logs);
    SpreadsheetApp.flush();
    return { ok: true, row: row, changes: logs.length, savedAt: nowString_() };
  } finally {
    lock.releaseLock();
  }
}

/** ---------------- เพิ่มแถวใหม่ (ต่อท้าย) ---------------- */
function addRow() {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getSheet_();
    ensureMarkerCol_(sh);
    var newRow = sh.getLastRow() + 1;
    var no = newRow - 1;
    sh.getRange(newRow, 1).setValue(no);            // คอลัมน์ A = No.
    sh.getRange(newRow, MARKER_COL).setValue('WEB'); // marker = เพิ่มเองผ่านเว็บ
    SpreadsheetApp.flush();
    return {
      row: newRow, no: no, manual: true, code: '', mainDept: '', subDept: '', mainCat: '', subCat: '',
      name1: '', name2: '', image: '', spec: '', brand: '', weight: '', width: '',
      length: '', height: '', unitStore: '', unitBuy: '', note: ''
    };
  } finally {
    lock.releaseLock();
  }
}

/** ---------------- ลบแถว (สำหรับรายการที่เพิ่มใหม่) ---------------- */
function deleteRow(row) {
  row = Number(row);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getSheet_();
    if (row < FIRST_DATA_ROW || row > sh.getLastRow()) throw new Error('แถวไม่ถูกต้อง');
    var code = sh.getRange(row, 2).getValue();
    sh.deleteRow(row);
    appendLogs_([[new Date(), currentUser_(), row, code, 'ลบแถว (DELETE ROW)', '', '']]);
    SpreadsheetApp.flush();
    return { ok: true, row: row };
  } finally {
    lock.releaseLock();
  }
}

/** ---------------- บันทึกแถวใหม่ (เขียน B..Q รวม Material Code/SubCat/Name1) ---------------- */
function saveNewRow(payload) {
  if (!payload || !payload.row) throw new Error('ไม่มีข้อมูลแถวที่จะบันทึก');
  var row = Number(payload.row);
  var map = [
    { col: 2,  key: 'code',      type: 'text' }, { col: 3,  key: 'mainDept', type: 'text' },
    { col: 4,  key: 'subDept',   type: 'text' }, { col: 5,  key: 'mainCat',  type: 'text' },
    { col: 6,  key: 'subCat',    type: 'text' }, { col: 7,  key: 'name1',    type: 'text' },
    { col: 8,  key: 'name2',     type: 'text' }, { col: 9,  key: 'image',    type: 'text' },
    { col: 10, key: 'spec',      type: 'text' }, { col: 11, key: 'brand',    type: 'text' },
    { col: 12, key: 'weight',    type: 'num'  }, { col: 13, key: 'width',    type: 'num'  },
    { col: 14, key: 'length',    type: 'num'  }, { col: 15, key: 'height',   type: 'num'  },
    { col: 16, key: 'unitStore', type: 'text' }, { col: 17, key: 'unitBuy',  type: 'text' },
    { col: 18, key: 'note',      type: 'text' }
  ];
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = getSheet_();
    if (row > sh.getLastRow()) throw new Error('แถวเกินขอบเขตข้อมูล');
    var rng = sh.getRange(row, 2, 1, 17);     // B..R (ไม่แตะ marker S)
    var old = rng.getValues()[0];
    var code = toText_(payload.code) || sh.getRange(row, 2).getValue();
    var when = new Date(), who = currentUser_(), logs = [], arr = old.slice();
    for (var i = 0; i < map.length; i++) {
      var off = map[i].col - 2;
      var nv = (map[i].type === 'num') ? toNumberOrText_(payload[map[i].key]) : toText_(payload[map[i].key]);
      if (String(old[off]) !== String(nv)) { logs.push([when, who, row, code, map[i].key, old[off], nv]); arr[off] = nv; }
    }
    if (logs.length) { rng.setValues([arr]); appendLogs_(logs); SpreadsheetApp.flush(); }
    return { ok: true, row: row, changes: logs.length, savedAt: nowString_() };
  } finally {
    lock.releaseLock();
  }
}

/** ---------------- อัปโหลดรูป → Drive แล้วเก็บลิงก์ที่คอลัมน์ I ---------------- */
function uploadImage(payload) {
  if (!payload || !payload.row || !payload.base64) throw new Error('ข้อมูลรูปไม่ครบ');
  var row = Number(payload.row);
  if (row < FIRST_DATA_ROW) throw new Error('แถวไม่ถูกต้อง');

  var bytes = Utilities.base64Decode(payload.base64);
  var blob = Utilities.newBlob(bytes, payload.mimeType || 'image/jpeg',
                               payload.filename || ('material_row' + row + '.jpg'));
  var folder = getImageFolder_();
  var file = folder.createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  var url = 'https://drive.google.com/file/d/' + file.getId() + '/view';

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getSheet_();
    if (row <= sh.getLastRow()) {
      var oldUrl = sh.getRange(row, IMAGE_COL).getValue();
      var code = sh.getRange(row, 2).getValue();
      sh.getRange(row, IMAGE_COL).setValue(url);
      if (String(oldUrl) !== String(url)) {
        appendLogs_([[new Date(), currentUser_(), row, code, 'Image (I)', oldUrl, url]]);
      }
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { ok: true, row: row, url: url, fileId: file.getId() };
}

function getImageFolder_() {
  try { return DriveApp.getFolderById(IMAGE_FOLDER_ID); }
  catch (e) { throw new Error('เข้าถึงโฟลเดอร์เก็บรูปไม่ได้ (ID: ' + IMAGE_FOLDER_ID + ')'); }
}

/** รันครั้งเดียวใน editor เพื่ออนุญาตสิทธิ์ Drive */
function authorizeDrive() {
  var folder = getImageFolder_();
  Logger.log('OK — โฟลเดอร์เก็บรูป: ' + folder.getName() + ' (' + folder.getId() + ')');
  return folder.getUrl();
}

/** ---------------- Log ---------------- */
function currentUser_() {
  var u = '';
  try { u = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!u) { try { u = Session.getEffectiveUser().getEmail(); } catch (e) {} }
  return u || '(ไม่ทราบ)';
}

function getLogSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET_NAME);
    sh.appendRow(['เวลา', 'ผู้แก้ไข', 'แถว', 'Material Code', 'ช่อง', 'ค่าเดิม', 'ค่าใหม่']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 7).setFontWeight('bold');
  }
  return sh;
}

function appendLogs_(rows) {
  if (!rows || !rows.length) return;
  var sh = getLogSheet_();
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

/** ---------------- helpers ---------------- */
function toText_(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
function toNumberOrText_(v) {
  var s = toText_(v);
  if (s === '') return '';
  var n = Number(s.replace(/,/g, ''));
  return isNaN(n) ? s : n;
}
function nowString_() { return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy HH:mm:ss'); }

/** ---------------- เครื่องมือสำรวจโครงสร้าง (ใช้ตอน setup) ---------------- */
function inspectDatabase() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var report = { spreadsheetName: ss.getName(), sheets: [] };
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    var info = { name: sh.getName(), lastRow: lastRow, lastColumn: lastCol, headers: [], sampleRows: [] };
    if (lastRow >= 1 && lastCol >= 1) {
      info.headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
      var n = Math.min(3, lastRow - 1);
      if (n > 0) info.sampleRows = sh.getRange(2, 1, n, lastCol).getValues();
    }
    report.sheets.push(info);
  }
  Logger.log(JSON.stringify(report, null, 2));
  return JSON.stringify(report);
}
