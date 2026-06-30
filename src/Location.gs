// ================================================================
//  Location.gs — สถานที่ CRUD
//  Sheet: สถานที่ | Columns: id (A), ชั้น_id (B), ชื่อสถานที่ (C)
// ================================================================

/**
 * เพิ่มสถานที่ใหม่
 * @param {{ ชั้น_id: string, ชื่อสถานที่: string }} data
 * @returns {string} UUID
 */
function addLocation(data) {
  if (!data || !data['ชั้น_id']) throw new Error('ชั้น_id ต้องไม่ว่างเปล่า');
  if (!String(data['ชื่อสถานที่'] || '').trim()) throw new Error('ชื่อสถานที่ ต้องไม่ว่างเปล่า');
  if (!getFloor(data['ชั้น_id'])) {
    throw new Error(`ไม่พบชั้น id: ${data['ชั้น_id']}`);
  }
  return appendRowGetId(SHEET_LOCATION, [data['ชั้น_id'], data['ชื่อสถานที่'].trim()]);
}

/**
 * ดึงสถานที่เดียวจาก id
 * @param {string} id
 * @returns {Object|null}
 */
function getLocation(id) {
  return findById(getAllLocations(), id);
}

/**
 * ดึงสถานที่ทั้งหมด
 * @returns {Object[]}
 */
function getAllLocations() {
  return getAllAsObjects(SHEET_LOCATION);
}

/**
 * ดึงสถานที่ของชั้น (ใช้สำหรับ cascade dropdown)
 * @param {string} floorId
 * @returns {Object[]}
 */
function getLocationsByFloor(floorId) {
  return getAllLocations().filter(l => l['ชั้น_id'] === floorId);
}

/**
 * แก้ไขสถานที่
 * @param {string} id
 * @param {{ ชั้น_id: string, ชื่อสถานที่: string }} data
 * @returns {boolean}
 */
function updateLocation(id, data) {
  if (!data || !data['ชั้น_id']) throw new Error('ชั้น_id ต้องไม่ว่างเปล่า');
  if (!String(data['ชื่อสถานที่'] || '').trim()) throw new Error('ชื่อสถานที่ ต้องไม่ว่างเปล่า');
  return updateRowById(SHEET_LOCATION, id, [id, data['ชั้น_id'], data['ชื่อสถานที่'].trim()]);
}

/**
 * ลบสถานที่ (hard delete)
 * ⚠️ ตรวจสอบ FK การตรวจสอบก่อนลบ
 * @param {string} id
 * @returns {boolean}
 */
function deleteLocation(id) {
  const used = getAllInspections().some(r => r['สถานที่_id'] === id);
  if (used) {
    throw new Error('ไม่สามารถลบได้ — มีรายการตรวจสอบที่อ้างอิงสถานที่นี้');
  }
  return deleteRowById(SHEET_LOCATION, id);
}
