// ================================================================
//  Building.gs — อาคาร CRUD
//  Sheet: อาคาร | Columns: id (A), ชื่ออาคาร (B)
// ================================================================

/**
 * เพิ่มอาคารใหม่
 * @param {{ ชื่ออาคาร: string }} data
 * @returns {string} UUID ของอาคารที่สร้าง
 */
function addBuilding(data) {
  if (!data || !String(data.ชื่ออาคาร || '').trim()) {
    throw new Error('ชื่ออาคาร ต้องไม่ว่างเปล่า');
  }
  return appendRowGetId(SHEET_BUILDING, [data.ชื่ออาคาร.trim()]);
}

/**
 * ดึงอาคารเดียวจาก id
 * @param {string} id
 * @returns {{ id: string, ชื่ออาคาร: string }|null}
 */
function getBuilding(id) {
  return findById(getAllBuildings(), id);
}

/**
 * ดึงอาคารทั้งหมด
 * @returns {{ id: string, ชื่ออาคาร: string }[]}
 */
function getAllBuildings() {
  return getAllAsObjects(SHEET_BUILDING);
}

/**
 * แก้ไขชื่ออาคาร
 * @param {string} id
 * @param {{ ชื่ออาคาร: string }} data
 * @returns {boolean}
 */
function updateBuilding(id, data) {
  if (!data || !String(data.ชื่ออาคาร || '').trim()) {
    throw new Error('ชื่ออาคาร ต้องไม่ว่างเปล่า');
  }
  return updateRowById(SHEET_BUILDING, id, [id, data.ชื่ออาคาร.trim()]);
}

/**
 * ลบอาคาร (hard delete)
 * ⚠️ ควรตรวจสอบว่าไม่มี ชั้น FK ก่อน
 * @param {string} id
 * @returns {boolean}
 */
function deleteBuilding(id) {
  const children = getAllFloors().filter(f => f['อาคาร_id'] === id);
  if (children.length > 0) {
    throw new Error(`ไม่สามารถลบได้ — มีชั้นที่อ้างอิงอาคารนี้ ${children.length} รายการ`);
  }
  return deleteRowById(SHEET_BUILDING, id);
}
