// ================================================================
//  EquipmentType.gs — ประเภทอุปกรณ์ CRUD
//  Sheet: ประเภทอุปกรณ์ | Columns: id (A), ชื่อประเภท (B), อิโมจิ (C)
// ================================================================

function addEquipmentType(data) {
  if (!data || !String(data['ชื่อประเภท'] || '').trim()) {
    throw new Error('ชื่อประเภท ต้องไม่ว่างเปล่า');
  }
  return appendRowGetId(SHEET_EQUIPMENT_TYPE, [data['ชื่อประเภท'].trim(), data['อิโมจิ'] || '🏷']);
}

function getEquipmentType(id) {
  return findById(getAllEquipmentTypes(), id);
}

function getAllEquipmentTypes() {
  return getAllAsObjects(SHEET_EQUIPMENT_TYPE);
}

function updateEquipmentType(id, data) {
  if (!data || !String(data['ชื่อประเภท'] || '').trim()) {
    throw new Error('ชื่อประเภท ต้องไม่ว่างเปล่า');
  }
  return updateRowById(SHEET_EQUIPMENT_TYPE, id, [id, data['ชื่อประเภท'].trim(), data['อิโมจิ'] || '🏷']);
}

function deleteEquipmentType(id) {
  return deleteRowById(SHEET_EQUIPMENT_TYPE, id);
}
