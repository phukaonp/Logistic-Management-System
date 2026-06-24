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
 *   R(18) จำนวน          [แก้ได้ • ตัวเลข]   (เพิ่ม 2026-06-23)
 *   S(19) Note           [แก้ได้]            (เพิ่ม 2026-06-23)
 *   T(20) _web marker    [ภายใน • ค่า "WEB"] (ย้ายจาก R เดิม)
 */

var SPREADSHEET_ID = '1FtoQo1BptJ3QnJxULZ-VZ130n0IKqNaV2XznCeo4kJE';
var SHEET_NAME = 'RawData1';
var DEPT_SHEET = 'Dept';
var MAINCAT_SHEET = 'MainCat';
var UNITS_SHEET = 'Units';        // หน่วย (คอลัมน์ B = unit_name)
var IMAGE_FOLDER_ID = '1CVDBgTgMbs9cHcC4yrJf2DKmrXD0bx4W';
var LOG_SHEET_NAME = '_EditLog';
var FIRST_DATA_ROW = 2;
var TOTAL_COLUMNS = 19;     // A..S (คอลัมน์ข้อมูลที่กรอก)
var IMAGE_COL = 9;          // I
var QTY_COL = 18;           // R — จำนวน (แก้ได้ • ตัวเลข)
var NOTE_COL = 19;          // S — Note (แก้ได้ • ข้อความ)
var MARKER_COL = 20;        // T — ทำเครื่องหมายแถวที่เพิ่มเองผ่านเว็บ (ค่า "WEB") • ย้ายจาก R เดิม

// ----- เมนู "เพิ่มจาก ERP" -----
var STAGING_SHEET = 'DT1.1.1_Staging';  // ชีตพักรายการที่เพิ่มจาก ERP (โครงสร้าง A..R เหมือน DT1.1.1)
var ALLMAT_SHEET  = 'All_Mat';          // ข้อมูลวัสดุทั้งหมดจาก ERP
var ERP_NOTES_SHEET = 'ERP_Notes';      // ชีตจดรายการที่ "ยังไม่มีใน ERP/All_Mat" (รอเพิ่มเข้า ERP)
var ALLMAT_CACHE_KEY = 'allmat_v1';     // cache รายการ All_Mat (แบบย่อ) เพื่อค้นหาเร็ว
var ALLMAT_CACHE_TTL = 21600;           // 6 ชม.
var WAREHOUSE_NAME = 'BPI';             // ชื่อคลัง (ใช้แสดงป้าย "อยู่ใน BPI แล้ว")
// หัวคอลัมน์ใน All_Mat ที่ใช้เติมให้อัตโนมัติ (จับคู่ด้วย "ชื่อหัว" ไม่ผูกเลขคอลัมน์)
var ALLMAT_COLS = {
  code:      'Product Code',   // → B Material Code (รหัสเต็ม PBPI… ไม่ใช่ 'PRODUCT CODE')
  subCat:    'SUB GROUP NAME', // → F SubCat
  name1:     'ชื่อในPO',       // → G Material Name (1)
  unitStore: 'Unit Name (IC)', // → P หน่วยเก็บ
  unitBuy:   'Unit Name (PO)'  // → Q หน่วยซื้อ
};

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
  { key: 'qty',       col: 18, label: 'จำนวน (R)',              type: 'num'  },
  { key: 'note',      col: 19, label: 'Note (S)',               type: 'text' }
];

/** ---------------- Web App ---------------- */
function doGet() {
  try { getLogSheet_(); } catch (e) {}
  try { ensureMarkerCol_(getSheet_()); } catch (e) {}
  try { maybeMigrateRST_(); } catch (e) {}   // ย้ายโครงสร้าง R→T อัตโนมัติครั้งเดียว
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

/** เลือกชีตปลายทางตาม key ('staging' = ชีตพัก, อื่น ๆ = DT1.1.1) */
function sheetFor_(sheetKey) {
  return (sheetKey === 'staging') ? getStagingSheet_() : getSheet_();
}

/** ชีตพัก (staging) — สร้างอัตโนมัติถ้ายังไม่มี โดยก็อปหัวตาราง A..R จาก DT1.1.1 */
function getStagingSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(STAGING_SHEET);
  if (!sh) {
    sh = ss.insertSheet(STAGING_SHEET);
    var src = getSheet_();
    var width = Math.max(MARKER_COL, src.getLastColumn());
    sh.getRange(1, 1, 1, width).setValues(src.getRange(1, 1, 1, width).getValues());
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, width).setFontWeight('bold');
  }
  ensureMarkerCol_(sh);
  return sh;
}

/** ให้แน่ใจว่ามีคอลัมน์ R(จำนวน) S(Note) T(_web marker) ครบ + หัวตาราง */
function ensureMarkerCol_(sh) {
  if (sh.getMaxColumns() < MARKER_COL) sh.insertColumnsAfter(sh.getMaxColumns(), MARKER_COL - sh.getMaxColumns());
  ensureHeader_(sh, QTY_COL, 'จำนวน');
  ensureHeader_(sh, NOTE_COL, 'Note');
  ensureHeader_(sh, MARKER_COL, '_web');
}
/** ตั้งหัวคอลัมน์เฉพาะเมื่อยังว่าง (ไม่ทับของเดิม) */
function ensureHeader_(sh, col, label) {
  if (!String(sh.getRange(1, col).getValue()).trim()) sh.getRange(1, col).setValue(label);
}

/** ---------------- ย้ายโครงสร้าง: marker _web จาก R(18) → T(20) + เพิ่ม R(จำนวน) S(Note) ----------------
 *  ย้ายเฉพาะเซลล์คอลัมน์ R ที่มีค่า "WEB" (marker เดิม) ไป T แล้วล้าง R — ปลอดภัยกับข้อมูลจำนวน (ตัวเลข)
 *  รันซ้ำได้ (idempotent): รอบถัดไปจะไม่เจอ "WEB" ใน R อีก จึงไม่ย้ายซ้ำ
 *  เรียกเองได้ใน editor ตอน setup หรือปล่อยให้ doGet เรียกอัตโนมัติครั้งแรก */
function migrateColumnsRST() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID), out = [];
  [SHEET_NAME, STAGING_SHEET].forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { out.push(name + ': ไม่พบชีต (ข้าม)'); return; }
    if (sh.getMaxColumns() < MARKER_COL) sh.insertColumnsAfter(sh.getMaxColumns(), MARKER_COL - sh.getMaxColumns());
    var last = sh.getLastRow(), moved = 0;
    if (last >= FIRST_DATA_ROW) {
      var n = last - FIRST_DATA_ROW + 1;
      var rCol = sh.getRange(FIRST_DATA_ROW, QTY_COL, n, 1).getValues();    // R เดิม (อาจมี marker)
      var tCol = sh.getRange(FIRST_DATA_ROW, MARKER_COL, n, 1).getValues(); // T ปลายทาง
      var chR = false, chT = false;
      for (var i = 0; i < n; i++) {
        if (String(rCol[i][0]).trim().toUpperCase() === 'WEB') {           // ย้ายเฉพาะค่า marker
          if (!String(tCol[i][0]).trim()) { tCol[i][0] = 'WEB'; chT = true; }
          rCol[i][0] = ''; chR = true; moved++;
        }
      }
      if (chT) sh.getRange(FIRST_DATA_ROW, MARKER_COL, n, 1).setValues(tCol);
      if (chR) sh.getRange(FIRST_DATA_ROW, QTY_COL, n, 1).setValues(rCol);
    }
    sh.getRange(1, QTY_COL).setValue('จำนวน');    // ตั้งหัวให้ตรงโครงสร้างใหม่
    sh.getRange(1, NOTE_COL).setValue('Note');
    sh.getRange(1, MARKER_COL).setValue('_web');
    out.push(name + ': ย้าย marker ' + moved + ' แถว → T, ตั้งหัว R/S/T เรียบร้อย');
  });
  SpreadsheetApp.flush();
  var msg = out.join('\n');
  Logger.log(msg);
  return msg;
}

/** เรียกจาก doGet — ย้ายโครงสร้างอัตโนมัติ "ครั้งเดียว" (กันด้วย Script Property + lock) */
function maybeMigrateRST_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('migratedRST') === '1') return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;        // ได้ล็อกไม่ทัน — รอบหน้าค่อยทำ
  try {
    if (props.getProperty('migratedRST') === '1') return;
    migrateColumnsRST();
    props.setProperty('migratedRST', '1');
  } catch (e) {
    // ปล่อยให้รอบถัดไปลองใหม่
  } finally {
    lock.releaseLock();
  }
}

/** ---------------- อ่านข้อมูลวัสดุทั้งหมด ---------------- */
function getMaterials(sheetKey) {
  var sh = sheetFor_(sheetKey);
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
      qty: r[17],        // R จำนวน
      note: r[18]        // S Note
    });
  }
  return out;
}

/** ---------------- รายการตัวเลือก (dropdown) ---------------- */
function getLookups() {
  return {
    depts: readLookupColumn_(DEPT_SHEET, 2),       // Dept!B (แผนก)
    mainCats: readLookupColumn_(MAINCAT_SHEET, 2), // MainCat!B (หมวดหลัก)
    units: readLookupColumn_(UNITS_SHEET, 2),      // Units!B (หน่วย)
    deptColors: getDeptColors()                    // สีแผนกที่ผู้ใช้ตั้งไว้ (override)
  };
}

/** ---------------- สีแผนก (เก็บใน Script Properties — แชร์ทุกผู้ใช้) ---------------- */
function getDeptColors() {
  var raw = PropertiesService.getScriptProperties().getProperty('deptColors');
  if (!raw) return {};
  try { return JSON.parse(raw); } catch (e) { return {}; }
}
function setDeptColors(json) {
  var obj = (typeof json === 'string') ? JSON.parse(json) : (json || {});
  PropertiesService.getScriptProperties().setProperty('deptColors', JSON.stringify(obj));
  return { ok: true };
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

/** ---------------- All_Mat (ERP) → staging ---------------- */
/** map ชื่อหัวคอลัมน์ (แถว 1) → เลขคอลัมน์ (1-based) */
function headerIndex_(sh) {
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var idx = {};
  for (var i = 0; i < hdr.length; i++) {
    var k = String(hdr[i]).trim();
    if (k && !(k in idx)) idx[k] = i + 1;
  }
  return idx;
}

/** เซ็ตของ Material Code ในชีตหนึ่ง (อ่านคอลัมน์ B) */
function codeSet_(sh) {
  var set = {}, last = sh.getLastRow();
  if (last < FIRST_DATA_ROW) return set;
  var col = sh.getRange(FIRST_DATA_ROW, 2, last - FIRST_DATA_ROW + 1, 1).getValues();
  for (var i = 0; i < col.length; i++) { var c = String(col[i][0]).trim(); if (c) set[c] = true; }
  return set;
}

/** ---------------- cache รายการ All_Mat (แบบย่อ) ---------------- */
/** อ่าน All_Mat เฉพาะคอลัมน์ที่ใช้ → array ย่อ [{code,subCat,name1,unitStore,unitBuy}] */
function readAllMatList_() {
  var am = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(ALLMAT_SHEET);
  if (!am) throw new Error('ไม่พบชีต: ' + ALLMAT_SHEET);
  var idx = headerIndex_(am);
  var cCode = idx[ALLMAT_COLS.code];
  if (!cCode) throw new Error('ไม่พบคอลัมน์ "' + ALLMAT_COLS.code + '" ใน ' + ALLMAT_SHEET);
  var lastRow = am.getLastRow();
  if (lastRow < 2) return [];
  var vals = am.getRange(2, 1, lastRow - 1, am.getLastColumn()).getValues();
  function pick(r, name) { var c = idx[name]; return c ? String(r[c - 1]).trim() : ''; }
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i], code = String(r[cCode - 1]).trim();
    if (!code) continue;
    out.push({ code: code, subCat: pick(r, ALLMAT_COLS.subCat), name1: pick(r, ALLMAT_COLS.name1),
               unitStore: pick(r, ALLMAT_COLS.unitStore), unitBuy: pick(r, ALLMAT_COLS.unitBuy) });
  }
  return out;
}

/** เก็บ list ลง CacheService (แบ่ง chunk กันเกิน 100KB/คีย์) */
function putAllMatCache_(cache, list) {
  var s = JSON.stringify(list), CH = 25000, n = Math.ceil(s.length / CH) || 1, obj = {};
  for (var i = 0; i < n; i++) obj[ALLMAT_CACHE_KEY + '_' + i] = s.substring(i * CH, (i + 1) * CH);
  cache.putAll(obj, ALLMAT_CACHE_TTL);
  cache.put(ALLMAT_CACHE_KEY, String(n), ALLMAT_CACHE_TTL);
}

/** ดึง list จาก cache (ถ้าหมดอายุ/ไม่มี → อ่านชีตแล้ว cache ใหม่) */
function getAllMatList_() {
  var cache = CacheService.getScriptCache();
  var meta = cache.get(ALLMAT_CACHE_KEY);
  if (meta) {
    var n = Number(meta), keys = [];
    for (var i = 0; i < n; i++) keys.push(ALLMAT_CACHE_KEY + '_' + i);
    var parts = cache.getAll(keys), joined = '', ok = true;
    for (var j = 0; j < n; j++) { var p = parts[ALLMAT_CACHE_KEY + '_' + j]; if (p == null) { ok = false; break; } joined += p; }
    if (ok) { try { return JSON.parse(joined); } catch (e) {} }
  }
  var list = readAllMatList_();
  putAllMatCache_(cache, list);
  return list;
}

/** รีเฟรช cache All_Mat ด้วยตนเอง (รันใน editor หรือเรียกจากปุ่มก็ได้) */
function refreshAllMatCache() {
  var list = readAllMatList_();
  putAllMatCache_(CacheService.getScriptCache(), list);
  return list.length;
}

/** ---------------- ค้นหา / เพิ่มจาก ERP ---------------- */
/** โหลดรายการ All_Mat ทั้งหมด (แบบย่อ) ไปไว้ฝั่ง client → ค้นหา/กรองได้ทันทีโดยไม่เรียก server ซ้ำ
 *  ธง inWarehouse/inStaging คำนวณฝั่ง client จากข้อมูล RawData1/Staging ที่โหลดอยู่แล้ว */
function getAllMatList() {
  return getAllMatList_();
}

/** เพิ่มหลายรายการจาก ERP เข้า staging พร้อมกัน (กดเลือกก่อน แล้วบันทึกทีเดียว) */
function addFromErpBatch(codes) {
  if (typeof codes === 'string') { try { codes = JSON.parse(codes); } catch (e) { codes = [codes]; } }
  if (!codes || !codes.length) throw new Error('ยังไม่ได้เลือกรายการ');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var list = getAllMatList_(), byCode = {};
    for (var i = 0; i < list.length; i++) byCode[list[i].code] = list[i];
    var existing = {}, whst = [getSheet_(), getStagingSheet_()];
    for (var w = 0; w < whst.length; w++) { var s = codeSet_(whst[w]); for (var k in s) existing[k] = true; }

    var sh = getStagingSheet_();
    var startRow = sh.getLastRow() + 1;
    var rows = [], created = [], logs = [], when = new Date(), who = currentUser_(), skipped = [];
    for (var c = 0; c < codes.length; c++) {
      var code = String(codes[c] == null ? '' : codes[c]).trim();
      if (!code || existing[code] || !byCode[code]) { if (code && existing[code]) skipped.push(code); continue; }
      existing[code] = true;                       // กันซ้ำภายใน batch เดียวกัน
      var it = byCode[code], rowNum = startRow + rows.length, no = rowNum - 1;
      var arr = new Array(MARKER_COL);
      for (var x = 0; x < MARKER_COL; x++) arr[x] = '';
      arr[0] = no; arr[1] = code; arr[5] = it.subCat; arr[6] = it.name1;
      arr[15] = it.unitStore; arr[16] = it.unitBuy; arr[MARKER_COL - 1] = 'WEB';
      rows.push(arr);
      logs.push([when, who, rowNum, code, 'เพิ่มจาก ERP → staging', '', code]);
      created.push({ row: rowNum, no: no, manual: true, code: code, mainDept: '', subDept: '', mainCat: '',
                     subCat: it.subCat, name1: it.name1, name2: '', image: '', spec: '', brand: '',
                     weight: '', width: '', length: '', height: '', unitStore: it.unitStore, unitBuy: it.unitBuy,
                     qty: '', note: '' });
    }
    if (rows.length) {
      sh.getRange(startRow, 1, rows.length, MARKER_COL).setValues(rows);
      appendLogs_(logs);
      SpreadsheetApp.flush();
    }
    return { added: created, skipped: skipped };
  } finally {
    lock.releaseLock();
  }
}

/** ---------------- จดรายการที่ยังไม่มีใน ERP (All_Mat) ---------------- */
/** ชีต ERP_Notes — สร้างอัตโนมัติพร้อมหัวตาราง */
function getErpNotesSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(ERP_NOTES_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ERP_NOTES_SHEET);
    sh.getRange(1, 1, 1, 9).setValues([['วันที่', 'ผู้บันทึก', 'คำค้น', 'รายละเอียด/ชื่อวัสดุ',
      'Subgroup (คาดว่า)', 'หน่วย', 'ยี่ห้อ/Spec', 'หมายเหตุ', 'สถานะ']]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 9).setFontWeight('bold');
  }
  return sh;
}

/** บันทึกโน้ต — ถ้ามีหลายหน่วย จะแยกเป็นหลายแถว (แถวละหน่วย) */
function addErpNote(payload) {
  payload = payload || {};
  var desc = toText_(payload.description);
  if (!desc) throw new Error('กรุณากรอกรายละเอียด/ชื่อวัสดุ');
  var units = payload.units;
  if (typeof units === 'string') { try { units = JSON.parse(units); } catch (e) { units = [units]; } }
  if (!units || !units.length) units = [toText_(payload.unit)];   // เผื่อส่ง unit เดี่ยว/ว่าง
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getErpNotesSheet_();
    var when = new Date(), who = currentUser_(), rows = [];
    for (var i = 0; i < units.length; i++) {
      rows.push([when, who, toText_(payload.query), desc, toText_(payload.subgroup),
        toText_(units[i]), toText_(payload.brandSpec), toText_(payload.note), 'รอเพิ่ม']);
    }
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 9).setValues(rows);
    SpreadsheetApp.flush();
    return { ok: true, count: rows.length, savedAt: nowString_() };
  } finally {
    lock.releaseLock();
  }
}

/** อ่านรายการที่บันทึกไว้ใน ERP_Notes (ใหม่สุดก่อน) */
function getErpNotes() {
  var sh = getErpNotesSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var v = sh.getRange(2, 1, last - 1, 9).getValues();
  var out = [];
  for (var i = 0; i < v.length; i++) {
    var r = v[i];
    if (!toText_(r[3]) && !toText_(r[2])) continue;   // ข้ามแถวว่าง
    var dt = r[0], dstr = (dt instanceof Date) ? Utilities.formatDate(dt, 'Asia/Bangkok', 'dd/MM/yyyy HH:mm') : toText_(dt);
    out.push({ row: i + 2, date: dstr, user: toText_(r[1]), query: toText_(r[2]), description: toText_(r[3]),
      subgroup: toText_(r[4]), unit: toText_(r[5]), brandSpec: toText_(r[6]), note: toText_(r[7]), status: toText_(r[8]) });
  }
  out.reverse();
  return out;
}

/** ลบ 1 แถวใน ERP_Notes */
function deleteErpNote(row) {
  row = Number(row);
  if (row < 2) throw new Error('แถวไม่ถูกต้อง');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = getErpNotesSheet_();
    if (row > sh.getLastRow()) throw new Error('แถวเกินขอบเขตข้อมูล');
    sh.deleteRow(row);
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** ---------------- บันทึก 1 แถว (auto-save) + log ---------------- */
function saveMaterialRow(payload) {
  if (!payload || !payload.row) throw new Error('ไม่มีข้อมูลแถวที่จะบันทึก');
  var row = Number(payload.row);
  if (row < FIRST_DATA_ROW) throw new Error('แถวไม่ถูกต้อง');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = sheetFor_(payload.sheetKey);
    if (row > sh.getLastRow()) throw new Error('แถวเกินขอบเขตข้อมูล');

    // อ่านช่วง C..S (3..19) มาก่อน — แก้เฉพาะช่อง editable, คง F/G (อ่านอย่างเดียว) ไว้ • ไม่แตะ T (marker)
    var rng = sh.getRange(row, 3, 1, 17);
    var old = rng.getValues()[0];           // index 0 = C(3) ... 16 = S(19)
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
function addRow(sheetKey) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheetFor_(sheetKey);
    ensureMarkerCol_(sh);
    var newRow = sh.getLastRow() + 1;
    var no = newRow - 1;
    sh.getRange(newRow, 1).setValue(no);            // คอลัมน์ A = No.
    sh.getRange(newRow, MARKER_COL).setValue('WEB'); // marker = เพิ่มเองผ่านเว็บ
    SpreadsheetApp.flush();
    return {
      row: newRow, no: no, manual: true, code: '', mainDept: '', subDept: '', mainCat: '', subCat: '',
      name1: '', name2: '', image: '', spec: '', brand: '', weight: '', width: '',
      length: '', height: '', unitStore: '', unitBuy: '', qty: '', note: ''
    };
  } finally {
    lock.releaseLock();
  }
}

/** ---------------- ลบแถว (สำหรับรายการที่เพิ่มใหม่) ---------------- */
function deleteRow(row, sheetKey) {
  row = Number(row);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheetFor_(sheetKey);
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
    { col: 18, key: 'qty',       type: 'num'  }, { col: 19, key: 'note',     type: 'text' }
  ];
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = sheetFor_(payload.sheetKey);
    if (row > sh.getLastRow()) throw new Error('แถวเกินขอบเขตข้อมูล');
    var rng = sh.getRange(row, 2, 1, 18);     // B..S (ไม่แตะ T = marker)
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
    var sh = sheetFor_(payload.sheetKey);
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
