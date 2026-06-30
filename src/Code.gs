// ================================================================
//  Code.gs — Global Constants, Triggers, Setup, Seed Data
// ================================================================

// ── Sheet Names (ตรงกับชื่อ Tab จริงใน Spreadsheet) ─────────────
const SHEET_BUILDING    = 'อาคาร';
const SHEET_FLOOR       = 'ชั้น';
const SHEET_LOCATION    = 'สถานที่';
const SHEET_DEPARTMENT  = 'หน่วยงาน';
const SHEET_EQUIPMENT   = 'อุปกรณ์';
const SHEET_INSPECTION  = 'การตรวจสอบ';
const SHEET_DASHBOARD        = 'Dashboard';
const SHEET_EQUIPMENT_TYPE   = 'ประเภทอุปกรณ์';

// ── Inspection Result Enums ──────────────────────────────────────
const RESULT_PASS   = 'ปกติ';
const RESULT_FAIL   = 'ชำรุด';
const RESULT_REPAIR = 'อื่นๆ';

// ── Department Type Enums ────────────────────────────────────────
const DEPT_TYPE_OWNER     = 'รับผิดชอบ';
const DEPT_TYPE_INSPECTOR = 'ตรวจสอบ';

// ── Script Config Keys (เก็บใน PropertiesService ไม่ hardcode) ───
const CONFIG_SPREADSHEET_ID = 'SPREADSHEET_ID';
const CONFIG_UPCOMING_DAYS  = 'UPCOMING_DAYS';   // default: 7
const CONFIG_ALERT_EMAIL    = 'ALERT_EMAIL';
const CONFIG_ADMIN_PIN      = 'ADMIN_PIN';

// ── Header rows ต่อ Sheet ────────────────────────────────────────
const SHEET_HEADERS = {
  [SHEET_BUILDING]:    ['id', 'ชื่ออาคาร'],
  [SHEET_FLOOR]:       ['id', 'อาคาร_id', 'ชื่อชั้น'],
  [SHEET_LOCATION]:    ['id', 'ชั้น_id', 'ชื่อสถานที่'],
  [SHEET_DEPARTMENT]:  ['id', 'ชื่อหน่วยงาน'],
  [SHEET_EQUIPMENT]:      ['id', 'ชื่ออุปกรณ์', 'รหัสครุภัณฑ์', 'รายละเอียด', 'ประเภท', 'หน่วยงาน_id', 'สถานที่_id'],
  [SHEET_EQUIPMENT_TYPE]: ['id', 'ชื่อประเภท'],
  [SHEET_INSPECTION]:  [
    'id', 'วันที่ตรวจสอบ', 'อุปกรณ์_id', 'สถานที่_id',
    'หน่วยงานรับผิดชอบ_id', 'ผลการตรวจสอบ', 'หมายเหตุ',
    'หน่วยงานผู้ตรวจสอบ_id', 'ผู้ตรวจสอบ', 'วันที่ตรวจสอบครั้งถัดไป',
  ],
};

// ── Tab Colors (ตาม ui-spec.md §1) ──────────────────────────────
const TAB_COLORS = {
  [SHEET_DASHBOARD]:  '#1A237E',
  [SHEET_INSPECTION]: '#137333',
  [SHEET_BUILDING]:   '#0D47A1',
  [SHEET_FLOOR]:      '#1565C0',
  [SHEET_LOCATION]:   '#1976D2',
  [SHEET_DEPARTMENT]: '#42A5F5',
  [SHEET_EQUIPMENT]:      '#1E88E5',
  [SHEET_EQUIPMENT_TYPE]: '#0288D1',
};

// ================================================================
//  Triggers
// ================================================================

/** สร้าง Custom Menu เมื่อเปิด Spreadsheet */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔥 ระบบตรวจสอบ')
    .addItem('⟳ รีเฟรช Dashboard', 'refreshDashboard')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('⚙️ ตั้งค่า')
        .addItem('🏗 สร้าง / รีเซ็ต Sheets',    'setupSheets')
        .addItem('🏥 นำเข้าข้อมูลจริง',          'seedRealData')
        .addItem('🌱 ใส่ข้อมูลตัวอย่าง',         'seedData')
        .addItem('📧 ตั้งค่า Alert Email',        'promptAlertEmail')
        .addItem('⏰ ตั้ง Time-Driven Trigger',   'setupTimeTrigger')
    )
    .addToUi();
}

/**
 * Validate เมื่อแก้ไข Sheet การตรวจสอบ โดยตรง
 * ตรวจสอบ col F (ผลการตรวจสอบ) ให้เป็น enum ที่ถูกต้อง
 */
function onEdit(e) {
  const sheet = e.range.getSheet();
  // ล้าง cache ของชีตที่ถูกแก้ไขใน Sheet UI โดยตรง (กันข้อมูล cache ค้าง)
  invalidateSheetCache(sheet.getName());

  if (sheet.getName() !== SHEET_INSPECTION) return;

  const COL_RESULT = 6;  // F — ผลการตรวจสอบ (หมายเลข 1-based)
  if (e.range.getColumn() !== COL_RESULT) return;

  const val   = e.range.getValue();
  const valid = [RESULT_PASS, RESULT_FAIL, ''];
  if (val !== '' && !valid.includes(val)) {
    SpreadsheetApp.getUi().alert(
      `⚠️ ค่าไม่ถูกต้อง\n\nผลการตรวจสอบต้องเป็น:\n• ${RESULT_PASS}\n• ${RESULT_FAIL}`
    );
    e.range.setValue(e.oldValue || '');
  }
}

// ================================================================
//  Sheet Setup
// ================================================================

/**
 * สร้าง Sheets ทั้งหมดพร้อม header + formatting
 * ถ้า Sheet มีอยู่แล้ว จะ skip (ไม่ลบข้อมูลเดิม)
 */
function setupSheets() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();

  try {
    // 1. Master sheets
    [SHEET_BUILDING, SHEET_FLOOR, SHEET_LOCATION, SHEET_DEPARTMENT, SHEET_EQUIPMENT, SHEET_EQUIPMENT_TYPE]
      .forEach(name => _createSheetIfNeeded(ss, name));

    // 2. Transaction sheet + Conditional Formatting
    const insSheet = _createSheetIfNeeded(ss, SHEET_INSPECTION);
    _applyInspectionCF(insSheet);

    // 3. Dashboard sheet
    _createSheetIfNeeded(ss, SHEET_DASHBOARD);

    // 4. ลบ default Sheet1 / แผ่น1 ถ้ายังเหลืออยู่
    ['Sheet1', 'แผ่น1'].forEach(n => {
      const s = ss.getSheetByName(n);
      if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
    });

    // 5. เรียง Tab ตามลำดับ
    _reorderSheets(ss);

    ui.alert('✅ Setup เสร็จสิ้น', 'สร้าง Sheets และ Formatting ทั้งหมดแล้ว', ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('❌ เกิดข้อผิดพลาด', err.message, ui.ButtonSet.OK);
  }
}

/**
 * สร้าง Sheet ถ้ายังไม่มี — ตั้งสี Tab + เขียน Header row + Freeze
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function _createSheetIfNeeded(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  // Tab color
  if (TAB_COLORS[name]) sheet.setTabColor(TAB_COLORS[name]);

  // Header row — สร้างถ้าว่าง หรือเพิ่มคอลัมน์ที่หายไปใน sheet ที่มีอยู่แล้ว
  const headers = SHEET_HEADERS[name];
  if (headers) {
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
      headers.forEach(h => {
        if (!existing.includes(h)) {
          sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h);
          existing.push(h);
        }
      });
    }
  }

  // Format header row
  if (headers) {
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1A237E')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setFontSize(11)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 28);      // header height ตาม ui-spec §4
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }

  return sheet;
}

/**
 * ใช้ Conditional Formatting กับ Sheet การตรวจสอบ
 * ลำดับ rule สำคัญ: เกินกำหนด / ใกล้ครบ override สีผล
 */
function _applyInspectionCF(sheet) {
  sheet.clearConditionalFormatRules();

  const dataRange = sheet.getRange(`A2:J${sheet.getMaxRows()}`);
  const rules = [
    // เกินกำหนด — สูงสุด (override ทุกอย่าง)
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($J2<>"",$J2<TODAY())`)
      .setBackground('#FCE8E6').setFontColor('#7B0000').setBold(true)
      .setRanges([dataRange]).build(),

    // ใกล้ครบ ≤ 7 วัน
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($J2<>"",$J2>=TODAY(),$J2<=TODAY()+7)`)
      .setBackground('#FEF7E0').setFontColor('#FF6D00')
      .setRanges([dataRange]).build(),

    // ผ่าน
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$F2="${RESULT_PASS}"`)
      .setBackground('#E6F4EA').setFontColor('#137333')
      .setRanges([dataRange]).build(),

    // ไม่ผ่าน
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$F2="${RESULT_FAIL}"`)
      .setBackground('#FCE8E6').setFontColor('#A50000')
      .setRanges([dataRange]).build(),

    // ซ่อมแซม
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$F2="${RESULT_REPAIR}"`)
      .setBackground('#FEF7E0').setFontColor('#E37400')
      .setRanges([dataRange]).build(),
  ];

  sheet.setConditionalFormatRules(rules);
}

/** จัดลำดับ Tab ตาม ui-spec §1 */
function _reorderSheets(ss) {
  const order = [
    SHEET_DASHBOARD, SHEET_INSPECTION,
    SHEET_BUILDING, SHEET_FLOOR, SHEET_LOCATION,
    SHEET_DEPARTMENT, SHEET_EQUIPMENT, SHEET_EQUIPMENT_TYPE,
  ];
  order.forEach((name, idx) => {
    const sh = ss.getSheetByName(name);
    if (sh) ss.setActiveSheet(sh).moveActiveSheet(idx + 1);
  });
}

// ================================================================
//  Web App Entry Point
// ================================================================

/** Serve standalone web app */
function doGet(e) {
  var VALID = {dashboard:1, inspection:1, history:1, master:1, settings:1};
  var view = (e && e.parameter && VALID[e.parameter.view]) ? e.parameter.view : '';
  var tpl = HtmlService.createTemplateFromFile('Index');
  tpl.initialView = view;
  return tpl.evaluate()
    .setTitle('ระบบตรวจสอบอุปกรณ์ดับเพลิง')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ================================================================
//  Web App — Data Functions
// ================================================================

/** คืน Dashboard data ทั้งหมดสำหรับ Web App */
function getDashboardData() {
  const inspections = getAllInspections();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const upDays = parseInt(getConfig(CONFIG_UPCOMING_DAYS, '7'));
  const upcoming30 = new Date(today); upcoming30.setDate(upcoming30.getDate() + upDays);

  const total = inspections.length;
  const pass  = inspections.filter(r => r['ผลการตรวจสอบ'] === RESULT_PASS).length;
  const fail  = inspections.filter(r => r['ผลการตรวจสอบ'] === RESULT_FAIL).length;

  const eqMap  = {}; getAllEquipment().forEach(e  => { eqMap[e.id]  = e; });
  const locMap = {}; getAllLocations().forEach(l  => { locMap[l.id] = l; });
  const flMap  = {}; getAllFloors().forEach(f     => { flMap[f.id]  = f; });
  const bldMap = {}; getAllBuildings().forEach(b  => { bldMap[b.id] = b; });

  function locDisplay(locId) {
    const loc = locMap[locId];
    if (!loc) return locId;
    const fl  = flMap[loc['ชั้น_id']];
    const bld = fl ? bldMap[fl['อาคาร_id']] : null;
    return [bld ? bld['ชื่ออาคาร'] : '', fl ? fl['ชื่อชั้น'] : '', '—', loc['ชื่อสถานที่']].filter(Boolean).join(' ');
  }

  const buildings = Object.values(bldMap);  // ใช้ map ที่สร้างไว้แล้ว ไม่อ่านซ้ำ
  const buildingSummary = buildings.map(bld => {
    const flIds  = Object.values(flMap).filter(f => f['อาคาร_id'] === bld.id).map(f => f.id);
    const locIds = Object.values(locMap).filter(l => flIds.includes(l['ชั้น_id'])).map(l => l.id);
    const bIns   = inspections.filter(r => locIds.includes(r['สถานที่_id']));
    const bTotal = bIns.length;
    const bPass  = bIns.filter(r => r['ผลการตรวจสอบ'] === RESULT_PASS).length;
    const bFail  = bIns.filter(r => r['ผลการตรวจสอบ'] === RESULT_FAIL).length;
    return { name: bld['ชื่ออาคาร'], total: bTotal, pass: bPass, fail: bFail,
             pct: bTotal > 0 ? Math.round(bPass / bTotal * 100) : 0 };
  });

  const latestMap = {};
  inspections.forEach(r => {
    const key = `${r['สถานที่_id']}__${r['อุปกรณ์_id']}`;
    const d   = parseDate(r['วันที่ตรวจสอบ']);
    if (!d) return;
    if (!latestMap[key] || d > parseDate(latestMap[key]['วันที่ตรวจสอบ'])) latestMap[key] = r;
  });
  const latest = Object.values(latestMap);

  const overdue = latest
    .filter(r => { const nd = parseDate(r['วันที่ตรวจสอบครั้งถัดไป']); return nd && nd < today; })
    .map(r => {
      const nd   = parseDate(r['วันที่ตรวจสอบครั้งถัดไป']);
      const days = Math.floor((today - nd) / 86400000);
      return { equipment: (eqMap[r['อุปกรณ์_id']] || {})['ชื่ออุปกรณ์'] || r['อุปกรณ์_id'],
               location: locDisplay(r['สถานที่_id']), result: r['ผลการตรวจสอบ'],
               nextDate: formatDate(nd), daysOverdue: days };
    })
    .sort((a, b) => b.daysOverdue - a.daysOverdue);

  const upcoming = latest
    .filter(r => { const nd = parseDate(r['วันที่ตรวจสอบครั้งถัดไป']); return nd && nd >= today && nd <= upcoming30; })
    .map(r => {
      const nd   = parseDate(r['วันที่ตรวจสอบครั้งถัดไป']);
      const days = Math.floor((nd - today) / 86400000);
      return { equipment: (eqMap[r['อุปกรณ์_id']] || {})['ชื่ออุปกรณ์'] || r['อุปกรณ์_id'],
               location: locDisplay(r['สถานที่_id']), result: r['ผลการตรวจสอบ'],
               nextDate: formatDate(nd), daysLeft: days };
    })
    .sort((a, b) => a.daysLeft - b.daysLeft);

  const eqTypes = getAllEquipmentTypes();
  const typeIconMap = {};
  eqTypes.forEach(t => { typeIconMap[(t['ชื่อประเภท'] || '').trim()] = t['อิโมจิ'] || 'label|#1565C0'; });
  const typeMap = {};
  Object.values(eqMap).forEach(e => {
    const type = (e['ประเภท'] || '').trim() || 'ไม่ระบุ';
    if (!typeMap[type]) typeMap[type] = { type, count: 0, icon: typeIconMap[type] || 'label|#1565C0' };
    typeMap[type].count++;
  });
  const typeSummary = Object.values(typeMap).sort((a, b) => b.count - a.count);

  return { kpi: { total, pass, fail }, buildingSummary, typeSummary, overdue, upcoming };
}

/** คืน Inspection history พร้อม display-ready fields สำหรับ Web App */
function webGetInspectionHistory() {
  const inspections = getAllInspections();
  const eqMap   = {};  getAllEquipment().forEach(e       => { eqMap[e.id]    = e; });
  const locMap  = {};  getAllLocations().forEach(l       => { locMap[l.id]   = l; });
  const flMap   = {};  getAllFloors().forEach(f          => { flMap[f.id]    = f; });
  const bldMap  = {};  getAllBuildings().forEach(b       => { bldMap[b.id]   = b; });
  const deptMap = {};  getAllDepartments().forEach(d     => { deptMap[d.id]  = d; });
  const eqTypes = {};  getAllEquipmentTypes().forEach(t  => { eqTypes[(t['ชื่อประเภท']||'').trim()] = t['อิโมจิ'] || 'label|#1565C0'; });

  return inspections.map(r => {
    const eq   = eqMap[r['อุปกรณ์_id']]  || {};
    const loc  = locMap[r['สถานที่_id']] || {};
    const fl   = flMap[loc['ชั้น_id']]   || {};
    const bld  = bldMap[fl['อาคาร_id']]  || {};
    const type = (eq['ประเภท'] || '').trim();
    return {
      id:            r.id,
      equipmentId:   r['อุปกรณ์_id'] || '',
      date:          r['วันที่ตรวจสอบ']           ? formatDate(parseDate(r['วันที่ตรวจสอบ']))           : '',
      nextDate:      r['วันที่ตรวจสอบครั้งถัดไป'] ? formatDate(parseDate(r['วันที่ตรวจสอบครั้งถัดไป'])) : '',
      result:        r['ผลการตรวจสอบ'] || '',
      remarks:       r['หมายเหตุ']     || '',
      equipment:     eq['ชื่ออุปกรณ์']  || '',
      equipmentCode: eq['รหัสครุภัณฑ์'] || '',
      equipmentType: type,
      equipmentIcon: eqTypes[type] || 'label|#1565C0',
      building:      bld['ชื่ออาคาร']   || '',
      floor:         fl['ชื่อชั้น']      || '',
      location:      loc['ชื่อสถานที่']  || '',
      respDept:      (deptMap[r['หน่วยงานรับผิดชอบ_id']]  || {})['ชื่อหน่วยงาน'] || '',
      insDept:       (deptMap[r['หน่วยงานผู้ตรวจสอบ_id']] || {})['ชื่อหน่วยงาน'] || '',
      inspector:     r['ผู้ตรวจสอบ'] || '',
    };
  }).sort((a, b) => b.date.localeCompare(a.date));
}

/** คืน Master data ทั้งหมดสำหรับ Web App */
function getMasterData() {
  return {
    buildings:      getAllBuildings(),
    floors:         getAllFloors(),
    locations:      getAllLocations(),
    departments:    getAllDepartments(),
    equipment:      getAllEquipment(),
    equipmentTypes: getAllEquipmentTypes(),
  };
}

// ── Web CRUD Wrappers ────────────────────────────────────────────

function webAddBuilding(name) {
  try { const id = addBuilding({ 'ชื่ออาคาร': name }); return { success: true, id, masterData: getMasterData() }; }
  catch(e) { return { success: false, error: e.message }; }
}
function webDeleteBuilding(id) {
  try { deleteBuilding(id); return { success: true, masterData: getMasterData() }; }
  catch(e) { return { success: false, error: e.message }; }
}
function webAddFloor(buildingId, name) {
  try { const id = addFloor({ 'อาคาร_id': buildingId, 'ชื่อชั้น': name }); return { success: true, id, masterData: getMasterData() }; }
  catch(e) { return { success: false, error: e.message }; }
}
function webDeleteFloor(id) {
  try { deleteFloor(id); return { success: true, masterData: getMasterData() }; }
  catch(e) { return { success: false, error: e.message }; }
}
function webAddLocation(floorId, name) {
  try { const id = addLocation({ 'ชั้น_id': floorId, 'ชื่อสถานที่': name }); return { success: true, id, masterData: getMasterData() }; }
  catch(e) { return { success: false, error: e.message }; }
}
function webDeleteLocation(id) {
  try { deleteLocation(id); return { success: true, masterData: getMasterData() }; }
  catch(e) { return { success: false, error: e.message }; }
}
function webAddDepartment(name) {
  try { const id = addDepartment({ 'ชื่อหน่วยงาน': name }); return { success: true, id, masterData: getMasterData() }; }
  catch(e) { return { success: false, error: e.message }; }
}
function webDeleteDepartment(id) {
  try { deleteDepartment(id); return { success: true, masterData: getMasterData() }; }
  catch(e) { return { success: false, error: e.message }; }
}
function webAddEquipment(name, code, detail, type, deptId, locId) {
  try { const id = addEquipment({ 'ชื่ออุปกรณ์': name, 'รหัสครุภัณฑ์': code, 'รายละเอียด': detail, 'ประเภท': type, 'หน่วยงาน_id': deptId, 'สถานที่_id': locId }); return { success: true, id, masterData: getMasterData() }; }
  catch(e) { return { success: false, error: e.message }; }
}
function webDeleteEquipment(id) {
  try { deleteEquipment(id); return { success: true, masterData: getMasterData() }; }
  catch(e) { return { success: false, error: e.message }; }
}
function webAddEquipmentType(name, emoji) {
  try { const id = addEquipmentType({ 'ชื่อประเภท': name, 'อิโมจิ': emoji || '🏷' }); return { success: true, id, masterData: getMasterData() }; }
  catch(e) { return { success: false, error: e.message }; }
}
function webUpdateEquipmentType(id, name, emoji) {
  try { updateEquipmentType(id, { 'ชื่อประเภท': name, 'อิโมจิ': emoji || '🏷' }); return { success: true, masterData: getMasterData() }; }
  catch(e) { return { success: false, error: e.message }; }
}
function webDeleteEquipmentType(id) {
  try { deleteEquipmentType(id); return { success: true, masterData: getMasterData() }; }
  catch(e) { return { success: false, error: e.message }; }
}

function webSetAlertEmail(email) {
  PropertiesService.getScriptProperties().setProperty(CONFIG_ALERT_EMAIL, email);
}

function webAdminPinIsSet() {
  return !!getConfig(CONFIG_ADMIN_PIN);
}

function webCheckAdminPin(pin) {
  const stored = getConfig(CONFIG_ADMIN_PIN);
  return stored === String(pin);
}

function webSetAdminPin(pin) {
  PropertiesService.getScriptProperties().setProperty(CONFIG_ADMIN_PIN, String(pin));
}

function webUpdateBuilding(id, name) {
  try { updateBuilding(id, { 'ชื่ออาคาร': name }); return { success: true, masterData: getMasterData() }; }
  catch(e) { return { success: false, error: e.message }; }
}
function webUpdateFloor(id, buildingId, name) {
  try { updateFloor(id, { 'อาคาร_id': buildingId, 'ชื่อชั้น': name }); return { success: true, masterData: getMasterData() }; }
  catch(e) { return { success: false, error: e.message }; }
}
function webUpdateLocation(id, floorId, name) {
  try { updateLocation(id, { 'ชั้น_id': floorId, 'ชื่อสถานที่': name }); return { success: true, masterData: getMasterData() }; }
  catch(e) { return { success: false, error: e.message }; }
}
function webUpdateDepartment(id, name) {
  try { updateDepartment(id, { 'ชื่อหน่วยงาน': name }); return { success: true, masterData: getMasterData() }; }
  catch(e) { return { success: false, error: e.message }; }
}
function webUpdateEquipment(id, name, code, detail, type, deptId, locId) {
  try { updateEquipment(id, { 'ชื่ออุปกรณ์': name, 'รหัสครุภัณฑ์': code, 'รายละเอียด': detail, 'ประเภท': type, 'หน่วยงาน_id': deptId, 'สถานที่_id': locId }); return { success: true, masterData: getMasterData() }; }
  catch(e) { return { success: false, error: e.message }; }
}

function webSetupSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    [SHEET_BUILDING, SHEET_FLOOR, SHEET_LOCATION, SHEET_DEPARTMENT, SHEET_EQUIPMENT, SHEET_EQUIPMENT_TYPE]
      .forEach(name => _createSheetIfNeeded(ss, name));
    const insSheet = _createSheetIfNeeded(ss, SHEET_INSPECTION);
    _applyInspectionCF(insSheet);
    _createSheetIfNeeded(ss, SHEET_DASHBOARD);
    ['Sheet1', 'แผ่น1'].forEach(n => {
      const s = ss.getSheetByName(n);
      if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
    });
  } catch (err) {
    throw new Error('Setup ล้มเหลว: ' + err.message);
  }
}

function webSeedData() {
  try {
    addBuilding({ 'ชื่ออาคาร': 'อาคาร A — สำนักงานใหญ่' });
    addBuilding({ 'ชื่ออาคาร': 'อาคาร B — โกดังสินค้า' });
    const bC = addBuilding({ 'ชื่ออาคาร': 'อาคาร C — โรงงานผลิต' });
    const fC1 = addFloor({ 'อาคาร_id': bC, 'ชื่อชั้น': 'ชั้น 1' });
    addLocation({ 'ชั้น_id': fC1, 'ชื่อสถานที่': 'ทางเดินหลัก' });
    addDepartment({ 'ชื่อหน่วยงาน': 'ฝ่ายอาคารสถานที่',     'ประเภท': DEPT_TYPE_OWNER });
    addDepartment({ 'ชื่อหน่วยงาน': 'บริษัท SafeCheck จำกัด', 'ประเภท': DEPT_TYPE_INSPECTOR });
    addEquipment({ 'ชื่ออุปกรณ์': 'ถังดับเพลิงผงเคมีแห้ง', 'รหัสครุภัณฑ์': 'FE-DRY-001', 'รายละเอียด': 'ขนาด 4 kg, ABC type' });
    addEquipment({ 'ชื่ออุปกรณ์': 'ถังดับเพลิง CO₂',        'รหัสครุภัณฑ์': 'FE-CO2-001', 'รายละเอียด': 'ขนาด 5 kg' });
  } catch (err) {
    throw new Error('Seed ล้มเหลว: ' + err.message);
  }
}

// ================================================================
//  Config Helpers
// ================================================================

/** อ่านค่า config จาก ScriptProperties */
function getConfig(key, defaultValue = '') {
  return PropertiesService.getScriptProperties().getProperty(key) || defaultValue;
}

/** ตั้งค่า Alert Email ผ่าน dialog */
function promptAlertEmail() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt(
    '📧 Alert Email',
    'กรอก email สำหรับรับแจ้งเตือนอุปกรณ์เกินกำหนด:',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const email = res.getResponseText().trim();
  if (!email) return;
  PropertiesService.getScriptProperties().setProperty(CONFIG_ALERT_EMAIL, email);
  ui.alert('✅ บันทึก email แล้ว', email, ui.ButtonSet.OK);
}

// ================================================================
//  Time-Driven Trigger
// ================================================================

/** ตั้ง trigger รัน dailyJob() ทุกวัน 07:00 น. */
function setupTimeTrigger() {
  // ลบ trigger เดิมที่ซ้ำ
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dailyJob')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('dailyJob')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  SpreadsheetApp.getUi().alert(
    '✅ Time Trigger ตั้งแล้ว',
    'dailyJob() จะรันทุกวันเวลา 07:00 น.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/** Entry point สำหรับ time-driven trigger */
function dailyJob() {
  refreshDashboard();
  sendOverdueAlert();
}

// ================================================================
//  Seed Data — Phase 5 / Testing
// ================================================================

/**
 * ใส่ข้อมูล Master ตัวอย่างตาม ui-spec.md §2
 * (append เท่านั้น — ข้อมูลเดิมไม่หาย)
 */
function seedData() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.alert(
    '🌱 ใส่ข้อมูลตัวอย่าง',
    'จะ append ข้อมูลตัวอย่างลงใน Master sheets\nข้อมูลที่มีอยู่จะยังคงอยู่\n\nดำเนินการต่อ?',
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  try {
    // ── อาคาร ──────────────────────────────────────────────────
    const bA = addBuilding({ 'ชื่ออาคาร': 'อาคาร A — สำนักงานใหญ่' });
    const bB = addBuilding({ 'ชื่ออาคาร': 'อาคาร B — โกดังสินค้า' });
    const bC = addBuilding({ 'ชื่ออาคาร': 'อาคาร C — โรงงานผลิต' });
    const bD = addBuilding({ 'ชื่ออาคาร': 'อาคาร D — ที่จอดรถ' });

    // ── ชั้น ────────────────────────────────────────────────────
    const fA1  = addFloor({ 'อาคาร_id': bA, 'ชื่อชั้น': 'ชั้น 1' });
    const fA2  = addFloor({ 'อาคาร_id': bA, 'ชื่อชั้น': 'ชั้น 2' });
    const fA5  = addFloor({ 'อาคาร_id': bA, 'ชื่อชั้น': 'ชั้น 5' });
    const fB1  = addFloor({ 'อาคาร_id': bB, 'ชื่อชั้น': 'ชั้น 1' });
    const fB2  = addFloor({ 'อาคาร_id': bB, 'ชื่อชั้น': 'ชั้น 2' });
    const fC1  = addFloor({ 'อาคาร_id': bC, 'ชื่อชั้น': 'ชั้น 1' });
    const fC3  = addFloor({ 'อาคาร_id': bC, 'ชื่อชั้น': 'ชั้น 3' });
    const fDB1 = addFloor({ 'อาคาร_id': bD, 'ชื่อชั้น': 'ชั้น B1' });

    // ── สถานที่ ─────────────────────────────────────────────────
    const lA1lobby = addLocation({ 'ชั้น_id': fA1,  'ชื่อสถานที่': 'ล็อบบี้ชั้น 1' });
    const lA1elec  = addLocation({ 'ชั้น_id': fA1,  'ชื่อสถานที่': 'ห้องไฟฟ้า' });
    const lA2lift  = addLocation({ 'ชั้น_id': fA2,  'ชื่อสถานที่': 'โถงลิฟต์ ชั้น 2' });
    const lA5srv   = addLocation({ 'ชั้น_id': fA5,  'ชื่อสถานที่': 'ห้องเซิร์ฟเวอร์' });
    const lB1wh    = addLocation({ 'ชั้น_id': fB1,  'ชื่อสถานที่': 'คลังสินค้า A' });
    const lB1elec  = addLocation({ 'ชั้น_id': fB1,  'ชื่อสถานที่': 'ห้องไฟฟ้า' });
    const lB2wh    = addLocation({ 'ชั้น_id': fB2,  'ชื่อสถานที่': 'คลังสินค้า A' });
    const lC3lift  = addLocation({ 'ชั้น_id': fC3,  'ชื่อสถานที่': 'โถงลิฟต์ ชั้น 3' });
    const lDB1mach = addLocation({ 'ชั้น_id': fDB1, 'ชื่อสถานที่': 'ห้องเครื่อง' });

    // ── หน่วยงาน ────────────────────────────────────────────────
    const dFacility = addDepartment({ 'ชื่อหน่วยงาน': 'ฝ่ายอาคารสถานที่',     'ประเภท': DEPT_TYPE_OWNER });
    const dEng      = addDepartment({ 'ชื่อหน่วยงาน': 'ฝ่ายวิศวกรรม',          'ประเภท': DEPT_TYPE_OWNER });
    const dMaint    = addDepartment({ 'ชื่อหน่วยงาน': 'ฝ่ายซ่อมบำรุง',         'ประเภท': DEPT_TYPE_OWNER });
    const dFactory  = addDepartment({ 'ชื่อหน่วยงาน': 'ฝ่ายโรงงาน',            'ประเภท': DEPT_TYPE_OWNER });
    const dSafe     = addDepartment({ 'ชื่อหน่วยงาน': 'บริษัท SafeCheck จำกัด', 'ประเภท': DEPT_TYPE_INSPECTOR });

    // ── อุปกรณ์ ─────────────────────────────────────────────────
    const eDry = addEquipment({ 'ชื่ออุปกรณ์': 'ถังดับเพลิงผงเคมีแห้ง', 'รหัสครุภัณฑ์': 'FE-DRY-001', 'รายละเอียด': 'ขนาด 4 kg, ABC type' });
    const eCO2 = addEquipment({ 'ชื่ออุปกรณ์': 'ถังดับเพลิง CO₂',        'รหัสครุภัณฑ์': 'FE-CO2-001', 'รายละเอียด': 'ขนาด 5 kg' });
    const eEmg = addEquipment({ 'ชื่ออุปกรณ์': 'ไฟสำรองฉุกเฉิน',         'รหัสครุภัณฑ์': 'EL-EMG-001', 'รายละเอียด': 'LED 2×8W, backup 3 ชม.' });
    const eSpr = addEquipment({ 'ชื่ออุปกรณ์': 'สปริงเกอร์',               'รหัสครุภัณฑ์': 'SP-WET-001', 'รายละเอียด': 'Wet pipe system' });
    const eAlm = addEquipment({ 'ชื่ออุปกรณ์': 'สัญญาณเตือนไฟ',           'รหัสครุภัณฑ์': 'FA-ALM-001', 'รายละเอียด': 'Addressable type' });

    ui.alert(
      '✅ Seed Data เสร็จสิ้น',
      `เพิ่มข้อมูลตัวอย่าง:\n• 4 อาคาร\n• 8 ชั้น\n• 9 สถานที่\n• 5 หน่วยงาน\n• 5 อุปกรณ์`,
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert('❌ Seed Data ล้มเหลว', err.message, ui.ButtonSet.OK);
  }
}
