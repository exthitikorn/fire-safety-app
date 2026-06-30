// ================================================================
//  SheetHelper.gs — Generic read / write / find helpers
//  ไม่ควรแก้ไขโดยตรง — ใช้ผ่าน function เหล่านี้เท่านั้น
// ================================================================

// ── Cache layer (CacheService) ──────────────────────────────────
// แต่ละ sheet ถูก cache เป็น JSON 1 key (TTL 5 นาที) ลดการอ่าน sheet ซ้ำ
// invalidate ทุกครั้งที่มีการเขียน (append/update/delete) + onEdit + seed
const _CACHE_PREFIX = 'sheetObjs_v1_';
const _CACHE_TTL    = 300;     // วินาที (5 นาที)
const _CACHE_MAX    = 95000;   // ปลอดภัยใต้ลิมิต 100KB ต่อ key

function _cacheKey(sheetName) { return _CACHE_PREFIX + sheetName; }

/** ล้าง cache ของชีตหนึ่ง — เรียกทุกครั้งที่มีการเขียน */
function invalidateSheetCache(sheetName) {
  try { CacheService.getScriptCache().remove(_cacheKey(sheetName)); } catch (_) {}
}

/** ล้าง cache ของทุกชีต entity — ใช้หลัง seed / bulk import */
function clearAllSheetCache() {
  try {
    const keys = [SHEET_BUILDING, SHEET_FLOOR, SHEET_LOCATION,
                  SHEET_DEPARTMENT, SHEET_EQUIPMENT, SHEET_EQUIPMENT_TYPE,
                  SHEET_INSPECTION].map(_cacheKey);
    CacheService.getScriptCache().removeAll(keys);
  } catch (_) {}
}

// Date.toJSON แปลง Date → ISO string ก่อน reviver จะกู้กลับได้
// จึงต้องอ่านค่าดิบจาก holder (this[key]) แล้วแทนด้วย marker { __d: epoch }
function _cacheReplacer(key, value) {
  const raw = this[key];
  if (raw instanceof Date) return { __d: raw.getTime() };
  return value;
}
function _cacheReviver(key, value) {
  if (value && typeof value === 'object' && typeof value.__d === 'number') {
    return new Date(value.__d);
  }
  return value;
}

/**
 * ดึงข้อมูลทุก row ของ sheet เป็น array ของ objects
 * อ่านจาก CacheService ก่อน (ถ้ามี) — ไม่งั้น batch getValues() แล้ว cache ไว้
 * @param {string} sheetName
 * @returns {Object[]}
 */
function getAllAsObjects(sheetName) {
  const cache = CacheService.getScriptCache();
  const key   = _cacheKey(sheetName);
  const hit   = cache.get(key);
  if (hit) {
    try { return JSON.parse(hit, _cacheReviver); } catch (_) { /* cache เสีย → อ่านใหม่ */ }
  }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const [headers, ...rows] = sheet.getDataRange().getValues();
  const objs = rows
    .filter(row => row[0] !== '' && row[0] !== null)   // skip empty rows
    .map(row => Object.fromEntries(headers.map((h, i) => [h, row[i]])));

  try {
    const ser = JSON.stringify(objs, _cacheReplacer);
    if (ser.length <= _CACHE_MAX) cache.put(key, ser, _CACHE_TTL);
  } catch (_) { /* ใหญ่เกินหรือ serialize ไม่ได้ → ข้ามการ cache (อ่านตรงต่อไป) */ }

  return objs;
}

/**
 * Append แถวใหม่โดย prepend UUID อัตโนมัติ
 * @param {string} sheetName
 * @param {Array}  rowData — ข้อมูลคอลัมน์ B ต่อไป (id จะถูก auto-generate)
 * @returns {string} UUID ที่สร้าง
 */
function appendRowGetId(sheetName, rowData) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`ไม่พบ Sheet: ${sheetName}`);
  const id = generateUUID();
  sheet.appendRow([id, ...rowData]);
  invalidateSheetCache(sheetName);
  return id;
}

/**
 * หา 1-based row index จาก id ใน column A
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} id
 * @returns {number} row index (≥ 2) หรือ -1 ถ้าไม่พบ
 */
function findRowById(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  // Batch read column A (ยกเว้น header row 1)
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const idx = ids.indexOf(id);
  return idx === -1 ? -1 : idx + 2;  // +2: offset header + 0-based
}

/**
 * อัปเดต row ทั้งหมดโดยใช้ id (batch setValues)
 * @param {string} sheetName
 * @param {string} id
 * @param {Array}  rowValues — ค่าทุกคอลัมน์ **รวม** id (col A)
 * @returns {boolean} true ถ้าอัปเดตสำเร็จ
 */
function updateRowById(sheetName, id, rowValues) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`ไม่พบ Sheet: ${sheetName}`);
  const row = findRowById(sheet, id);
  if (row === -1) return false;
  sheet.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);
  invalidateSheetCache(sheetName);
  return true;
}

/**
 * ลบ row โดยใช้ id (hard delete)
 * @param {string} sheetName
 * @param {string} id
 * @returns {boolean}
 */
function deleteRowById(sheetName, id) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`ไม่พบ Sheet: ${sheetName}`);
  const row = findRowById(sheet, id);
  if (row === -1) return false;
  sheet.deleteRow(row);
  invalidateSheetCache(sheetName);
  return true;
}

/**
 * ค้นหา object จาก id ใน array
 * @param {Object[]} objects
 * @param {string}   id
 * @returns {Object|null}
 */
function findById(objects, id) {
  return objects.find(o => o['id'] === id) || null;
}

// ── Date Utilities ───────────────────────────────────────────────

/**
 * แปลง Date object หรือ string เป็น Date
 * รองรับทั้ง 'dd/MM/yyyy' และ 'yyyy-MM-dd' (จาก HTML date input)
 * @param {Date|string|number} val
 * @returns {Date}
 */
function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  const s = String(val).trim();
  if (s.includes('/')) {
    // 'dd/MM/yyyy'
    const [d, m, y] = s.split('/').map(Number);
    return new Date(y, m - 1, d);
  }
  if (s.includes('-')) {
    // 'yyyy-MM-dd' (HTML input)
    const [y, m, d] = s.split('-').map(Number);
    return new Date(y, m - 1, d);
  }
  return new Date(val);
}

/**
 * Format Date → 'dd/MM/yyyy'
 * @param {Date|string} val
 * @returns {string}
 */
function formatDate(val) {
  if (!val) return '';
  const d = val instanceof Date ? val : parseDate(val);
  if (!d || isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

/**
 * สร้าง UUID v4 ผ่าน Apps Script built-in
 * @returns {string}
 */
function generateUUID() {
  return Utilities.getUuid();
}
