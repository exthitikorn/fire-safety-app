// ================================================================
//  Department.gs — หน่วยงาน CRUD
//  Sheet: หน่วยงาน | Columns: id (A), ชื่อหน่วยงาน (B)
// ================================================================

/**
 * เพิ่มหน่วยงานใหม่
 * @param {{ ชื่อหน่วยงาน: string }} data
 * @returns {string} UUID
 */
function addDepartment(data) {
  if (!data || !String(data['ชื่อหน่วยงาน'] || '').trim()) {
    throw new Error('ชื่อหน่วยงาน ต้องไม่ว่างเปล่า');
  }
  return appendRowGetId(SHEET_DEPARTMENT, [data['ชื่อหน่วยงาน'].trim()]);
}

/**
 * ดึงหน่วยงานเดียวจาก id
 * @param {string} id
 * @returns {Object|null}
 */
function getDepartment(id) {
  return findById(getAllDepartments(), id);
}

/**
 * ดึงหน่วยงานทั้งหมด
 * @returns {Object[]}
 */
function getAllDepartments() {
  return getAllAsObjects(SHEET_DEPARTMENT);
}

/**
 * แก้ไขหน่วยงาน
 * @param {string} id
 * @param {{ ชื่อหน่วยงาน: string }} data
 * @returns {boolean}
 */
function updateDepartment(id, data) {
  if (!data || !String(data['ชื่อหน่วยงาน'] || '').trim()) {
    throw new Error('ชื่อหน่วยงาน ต้องไม่ว่างเปล่า');
  }
  return updateRowById(SHEET_DEPARTMENT, id, [id, data['ชื่อหน่วยงาน'].trim()]);
}

/**
 * ลบหน่วยงาน (hard delete)
 * @param {string} id
 * @returns {boolean}
 */
function deleteDepartment(id) {
  const inspections   = getAllInspections();  // อ่านครั้งเดียว
  const usedAsOwner    = inspections.some(r => r['หน่วยงานรับผิดชอบ_id']  === id);
  const usedAsInspect  = inspections.some(r => r['หน่วยงานผู้ตรวจสอบ_id'] === id);
  if (usedAsOwner || usedAsInspect) {
    throw new Error('ไม่สามารถลบได้ — มีรายการตรวจสอบที่อ้างอิงหน่วยงานนี้');
  }
  return deleteRowById(SHEET_DEPARTMENT, id);
}
