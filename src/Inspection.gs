// ================================================================
//  Inspection.gs — การตรวจสอบ CRUD + Sidebar bridge functions
//  Sheet: การตรวจสอบ
//  Columns: id(A) วันที่ตรวจสอบ(B) อุปกรณ์_id(C) สถานที่_id(D)
//           หน่วยงานรับผิดชอบ_id(E) ผลการตรวจสอบ(F) หมายเหตุ(G)
//           หน่วยงานผู้ตรวจสอบ_id(H) ผู้ตรวจสอบ(I) วันที่ตรวจสอบครั้งถัดไป(J)
// ================================================================

// Column index map (1-based) — ใช้อ้างอิงแทน hardcode number
const INS_COL = {
  id:                    1,  // A
  วันที่ตรวจสอบ:          2,  // B
  อุปกรณ์_id:            3,  // C
  สถานที่_id:             4,  // D
  หน่วยงานรับผิดชอบ_id:  5,  // E
  ผลการตรวจสอบ:           6,  // F
  หมายเหตุ:               7,  // G
  หน่วยงานผู้ตรวจสอบ_id: 8,  // H
  ผู้ตรวจสอบ:             9,  // I
  วันที่ตรวจสอบครั้งถัดไป: 10, // J
};

// ================================================================
//  Validation
// ================================================================

/**
 * ตรวจสอบความถูกต้องของข้อมูลก่อนบันทึก
 * @param {Object} data
 * @throws {Error} ถ้า validation ไม่ผ่าน
 */
function validateInspection(data) {
  // Required fields
  const required = [
    ['inspectionDate',     'วันที่ตรวจสอบ'],
    ['equipmentId',        'อุปกรณ์'],
    ['locationId',         'สถานที่'],
    ['responsibleDeptId',  'หน่วยงานรับผิดชอบ'],
    ['result',             'ผลการตรวจสอบ'],
    ['inspectorDeptId',    'หน่วยงานผู้ตรวจสอบ'],
    ['inspectorName',      'ชื่อผู้ตรวจสอบ'],
    ['nextInspectionDate', 'วันที่ตรวจสอบครั้งถัดไป'],
  ];
  required.forEach(([key, label]) => {
    if (!data[key] || String(data[key]).trim() === '') {
      throw new Error(`กรุณากรอก: ${label}`);
    }
  });

  // Result enum
  if (![RESULT_PASS, RESULT_FAIL].includes(data.result)) {
    throw new Error(`ผลการตรวจสอบต้องเป็น: ${RESULT_PASS} / ${RESULT_FAIL}`);
  }

  // หมายเหตุ required เมื่อ ไม่ผ่าน
  if (data.result === RESULT_FAIL && !String(data.remarks || '').trim()) {
    throw new Error('กรุณากรอกหมายเหตุ/รายละเอียด เมื่อผลการตรวจสอบ = ไม่ผ่าน');
  }

  // nextDate > inspectionDate
  const d1 = parseDate(data.inspectionDate);
  const d2 = parseDate(data.nextInspectionDate);
  if (!d1 || !d2) throw new Error('รูปแบบวันที่ไม่ถูกต้อง');
  if (d2 <= d1)   throw new Error('วันที่ตรวจครั้งถัดไปต้องมากกว่าวันที่ตรวจสอบ');
}

// ================================================================
//  CRUD
// ================================================================

/**
 * บันทึกผลการตรวจสอบใหม่
 * @param {Object} data
 * @returns {string} UUID ของรายการที่สร้าง
 */
function addInspection(data) {
  validateInspection(data);
  return appendRowGetId(SHEET_INSPECTION, [
    parseDate(data.inspectionDate),         // B — เก็บเป็น Date object
    data.equipmentId.trim(),                // C
    data.locationId.trim(),                 // D
    data.responsibleDeptId.trim(),          // E
    data.result,                            // F
    String(data.remarks || '').trim(),      // G
    data.inspectorDeptId.trim(),            // H
    data.inspectorName.trim(),              // I
    parseDate(data.nextInspectionDate),     // J — เก็บเป็น Date object
  ]);
}

/**
 * ดึงรายการตรวจสอบเดียว
 * @param {string} id
 * @returns {Object|null}
 */
function getInspection(id) {
  return findById(getAllInspections(), id);
}

/**
 * ดึงรายการตรวจสอบทั้งหมด
 * @returns {Object[]}
 */
function getAllInspections() {
  return getAllAsObjects(SHEET_INSPECTION);
}

/**
 * ดึงรายการตรวจสอบตามช่วงวันที่
 * @param {Date|string} startDate
 * @param {Date|string} endDate
 * @returns {Object[]}
 */
function getInspectionsByDateRange(startDate, endDate) {
  const s = parseDate(startDate).getTime();
  const e = parseDate(endDate).getTime();
  return getAllInspections().filter(r => {
    const t = parseDate(r['วันที่ตรวจสอบ']);
    return t && t.getTime() >= s && t.getTime() <= e;
  });
}

/**
 * ดึงรายการตรวจสอบตาม location
 * @param {string} locationId
 * @returns {Object[]}
 */
function getInspectionsByLocation(locationId) {
  return getAllInspections().filter(r => r['สถานที่_id'] === locationId);
}

/**
 * ดึงรายการตรวจสอบตามอุปกรณ์
 * @param {string} equipmentId
 * @returns {Object[]}
 */
function getInspectionsByEquipment(equipmentId) {
  return getAllInspections().filter(r => r['อุปกรณ์_id'] === equipmentId);
}

/**
 * ดึงรายการตรวจสอบตามผล
 * @param {'ผ่าน'|'ไม่ผ่าน'|'ซ่อมแซม'} result
 * @returns {Object[]}
 */
function getInspectionsByResult(result) {
  return getAllInspections().filter(r => r['ผลการตรวจสอบ'] === result);
}

/**
 * แก้ไขรายการตรวจสอบ
 * @param {string} id
 * @param {Object} data
 * @returns {boolean}
 */
function updateInspection(id, data) {
  validateInspection(data);
  return updateRowById(SHEET_INSPECTION, id, [
    id,
    parseDate(data.inspectionDate),
    data.equipmentId.trim(),
    data.locationId.trim(),
    data.responsibleDeptId.trim(),
    data.result,
    String(data.remarks || '').trim(),
    data.inspectorDeptId.trim(),
    data.inspectorName.trim(),
    parseDate(data.nextInspectionDate),
  ]);
}

/**
 * ลบรายการตรวจสอบ
 * @param {string} id
 * @returns {boolean}
 */
function deleteInspection(id) {
  return deleteRowById(SHEET_INSPECTION, id);
}

// ================================================================
//  Sidebar Bridge Functions
//  (เรียกจาก client-side ผ่าน google.script.run)
// ================================================================

/**
 * ดึงข้อมูล dropdown ทั้งหมดสำหรับ Sidebar form ในครั้งเดียว
 * รวม floors + locations เพื่อทำ cascade ฝั่ง client โดยไม่ต้อง round-trip
 * @returns {{ buildings, floors, locations, equipment, responsibleDepts, inspectorDepts }}
 */
function getDropdownData() {
  const departments = getAllDepartments();  // อ่านครั้งเดียว ใช้ทั้ง 2 ฝั่ง
  return {
    buildings:        getAllBuildings(),
    floors:           getAllFloors(),
    locations:        getAllLocations(),
    equipment:        getAllEquipment(),
    equipmentTypes:   getAllEquipmentTypes(),
    responsibleDepts: departments,
    inspectorDepts:   departments,
  };
}

/**
 * บันทึกผลจาก Sidebar form
 * @param {Object} data — ข้อมูลจาก form (inspectionDate ใน format yyyy-MM-dd)
 * @returns {{ success: boolean, id?: string, error?: string }}
 */
function saveInspection(data) {
  try {
    const id = addInspection(data);
    return { success: true, id: id };
  } catch (err) {
    console.error('saveInspection error:', err.message);
    return { success: false, error: err.message };
  }
}
