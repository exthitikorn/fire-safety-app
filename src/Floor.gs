// ================================================================
//  Floor.gs — ชั้น CRUD
//  Sheet: ชั้น | Columns: id (A), อาคาร_id (B), ชื่อชั้น (C)
// ================================================================

/**
 * เพิ่มชั้นใหม่
 * @param {{ อาคาร_id: string, ชื่อชั้น: string }} data
 * @returns {string} UUID ของชั้นที่สร้าง
 */
function addFloor(data) {
  if (!data || !data['อาคาร_id']) throw new Error('อาคาร_id ต้องไม่ว่างเปล่า');
  if (!String(data['ชื่อชั้น'] || '').trim())  throw new Error('ชื่อชั้น ต้องไม่ว่างเปล่า');
  if (!getBuilding(data['อาคาร_id'])) {
    throw new Error(`ไม่พบอาคาร id: ${data['อาคาร_id']}`);
  }
  return appendRowGetId(SHEET_FLOOR, [data['อาคาร_id'], data['ชื่อชั้น'].trim()]);
}

/**
 * ดึงชั้นเดียวจาก id
 * @param {string} id
 * @returns {Object|null}
 */
function getFloor(id) {
  return findById(getAllFloors(), id);
}

/**
 * ดึงชั้นทั้งหมด
 * @returns {Object[]}
 */
function getAllFloors() {
  return getAllAsObjects(SHEET_FLOOR);
}

/**
 * ดึงชั้นของอาคาร (ใช้สำหรับ cascade dropdown)
 * @param {string} buildingId
 * @returns {Object[]}
 */
function getFloorsByBuilding(buildingId) {
  return getAllFloors().filter(f => f['อาคาร_id'] === buildingId);
}

/**
 * แก้ไขชั้น
 * @param {string} id
 * @param {{ อาคาร_id: string, ชื่อชั้น: string }} data
 * @returns {boolean}
 */
function updateFloor(id, data) {
  if (!data || !data['อาคาร_id']) throw new Error('อาคาร_id ต้องไม่ว่างเปล่า');
  if (!String(data['ชื่อชั้น'] || '').trim())  throw new Error('ชื่อชั้น ต้องไม่ว่างเปล่า');
  return updateRowById(SHEET_FLOOR, id, [id, data['อาคาร_id'], data['ชื่อชั้น'].trim()]);
}

/**
 * ลบชั้น (hard delete)
 * ⚠️ ตรวจสอบ FK สถานที่ก่อนลบ
 * @param {string} id
 * @returns {boolean}
 */
function deleteFloor(id) {
  const children = getAllLocations().filter(l => l['ชั้น_id'] === id);
  if (children.length > 0) {
    throw new Error(`ไม่สามารถลบได้ — มีสถานที่ที่อ้างอิงชั้นนี้ ${children.length} รายการ`);
  }
  return deleteRowById(SHEET_FLOOR, id);
}
