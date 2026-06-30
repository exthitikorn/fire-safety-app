// ================================================================
//  Equipment.gs — อุปกรณ์ CRUD
//  Sheet: อุปกรณ์ | Columns: id (A), ชื่ออุปกรณ์ (B), รหัสครุภัณฑ์ (C), รายละเอียด (D), ประเภท (E)
// ================================================================

/**
 * เพิ่มอุปกรณ์ใหม่
 * @param {{ ชื่ออุปกรณ์: string, รหัสครุภัณฑ์?: string, รายละเอียด?: string }} data
 * @returns {string} UUID
 */
function addEquipment(data) {
  if (!data || !String(data['ชื่ออุปกรณ์'] || '').trim()) {
    throw new Error('ชื่ออุปกรณ์ ต้องไม่ว่างเปล่า');
  }
  return appendRowGetId(SHEET_EQUIPMENT, [
    data['ชื่ออุปกรณ์'].trim(),
    String(data['รหัสครุภัณฑ์'] || '').trim(),
    String(data['รายละเอียด']   || '').trim(),
    String(data['ประเภท']       || '').trim(),
    String(data['หน่วยงาน_id']  || '').trim(),
    String(data['สถานที่_id']   || '').trim(),
  ]);
}

/**
 * ดึงอุปกรณ์เดียวจาก id
 * @param {string} id
 * @returns {Object|null}
 */
function getEquipment(id) {
  return findById(getAllEquipment(), id);
}

/**
 * ดึงอุปกรณ์ทั้งหมด
 * @returns {Object[]}
 */
function getAllEquipment() {
  return getAllAsObjects(SHEET_EQUIPMENT);
}

/**
 * แก้ไขอุปกรณ์
 * @param {string} id
 * @param {{ ชื่ออุปกรณ์: string, รหัสครุภัณฑ์?: string, รายละเอียด?: string }} data
 * @returns {boolean}
 */
function updateEquipment(id, data) {
  if (!data || !String(data['ชื่ออุปกรณ์'] || '').trim()) {
    throw new Error('ชื่ออุปกรณ์ ต้องไม่ว่างเปล่า');
  }
  return updateRowById(SHEET_EQUIPMENT, id, [
    id,
    data['ชื่ออุปกรณ์'].trim(),
    String(data['รหัสครุภัณฑ์'] || '').trim(),
    String(data['รายละเอียด']   || '').trim(),
    String(data['ประเภท']       || '').trim(),
    String(data['หน่วยงาน_id']  || '').trim(),
    String(data['สถานที่_id']   || '').trim(),
  ]);
}

/**
 * ลบอุปกรณ์ (hard delete)
 * @param {string} id
 * @returns {boolean}
 */
function deleteEquipment(id) {
  const used = getAllInspections().some(r => r['อุปกรณ์_id'] === id);
  if (used) {
    throw new Error('ไม่สามารถลบได้ — มีรายการตรวจสอบที่อ้างอิงอุปกรณ์นี้');
  }
  return deleteRowById(SHEET_EQUIPMENT, id);
}
