// ================================================================
//  Dashboard.gs — Dashboard refresh + Email alert
// ================================================================

// ── Dashboard Layout Row Constants ──────────────────────────────
// Row 1: Banner, Row 2: Updated, Row 3: empty
// Row 4: KPI labels, Row 5: KPI values, Row 6–7: empty
// Row 8: Building section header, Row 9: Building col headers
// Row 10+: Building data (dynamic)
// After buildings: Overdue table → Upcoming table

const DASH_ROW = {
  BANNER:     1,
  UPDATED:    2,
  KPI_LABEL:  4,
  KPI_VALUE:  5,
  BUILD_HDR:  8,
  BUILD_COLS: 9,
  BUILD_DATA: 10,
};

// KPI card positions (col, 1-based)
const DASH_KPI_COLS = {
  total:  1,   // A
  pass:   3,   // C
  fail:   5,   // E
};

// ================================================================
//  Main Entry Point
// ================================================================

/**
 * รีเฟรช Dashboard ทั้งหมด — เรียกจาก menu หรือ time trigger
 */
function refreshDashboard() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const dash = ss.getSheetByName(SHEET_DASHBOARD);
  if (!dash) {
    SpreadsheetApp.getUi().alert(
      '❌ ไม่พบ Sheet Dashboard',
      'กรุณารัน setupSheets() ก่อน',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  try {
    dash.clear();
    dash.clearConditionalFormatRules();

    _writeBanner(dash);
    updateKPICards(dash);
    const rowAfterBuilding = updateSummaryByBuilding(dash);
    const rowAfterOverdue  = updateOverdueList(dash,   rowAfterBuilding + 2);
    const rowAfterUpcoming = updateUpcomingList(dash,   rowAfterOverdue  + 2);
    _writeUpdatedTimestamp(dash);

    // Toast notification (ไม่ block UI)
    ss.toast('Dashboard อัปเดตแล้ว ✅', '🔥 ระบบตรวจสอบ', 5);
  } catch (err) {
    console.error('refreshDashboard error:', err.message);
    SpreadsheetApp.getUi().alert('❌ Dashboard Error', err.message,
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ================================================================
//  Banner & Timestamp
// ================================================================

function _writeBanner(dash) {
  // Row 1: Title
  dash.setRowHeight(DASH_ROW.BANNER, 52);
  dash.getRange(DASH_ROW.BANNER, 1, 1, 10).merge()
    .setValue('🔥 ระบบตรวจสอบอุปกรณ์ดับเพลิง')
    .setBackground('#1A237E').setFontColor('#FFFFFF')
    .setFontSize(20).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Row 2: placeholder for timestamp (written last)
  dash.setRowHeight(DASH_ROW.UPDATED, 20);
  dash.getRange(DASH_ROW.UPDATED, 1, 1, 10).merge()
    .setBackground('#FFFFFF');

  // Row 3: spacer
  dash.setRowHeight(3, 8);

  // Freeze banner + updated rows
  dash.setFrozenRows(2);

  // Column widths
  const colWidths = [130, 10, 130, 10, 130, 10, 130, 10, 220, 80, 80, 80, 80];
  colWidths.forEach((w, i) => {
    try { dash.setColumnWidth(i + 1, w); } catch (_) {}
  });
}

function _writeUpdatedTimestamp(dash) {
  const tz  = Session.getScriptTimeZone();
  const now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  dash.getRange(DASH_ROW.UPDATED, 1).setValue(`อัปเดตล่าสุด: ${now} น.`)
    .setFontColor('#5F6368').setFontSize(10)
    .setHorizontalAlignment('right').setFontStyle('italic');
}

// ================================================================
//  KPI Cards (Row 4–5)
// ================================================================

/**
 * คำนวณและเขียน KPI 4 ตัว (ทั้งหมด / ผ่าน / ไม่ผ่าน / ซ่อมแซม)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dash
 */
function updateKPICards(dash) {
  const inspections = getAllInspections();
  const total  = inspections.length;
  const pass   = inspections.filter(r => r['ผลการตรวจสอบ'] === RESULT_PASS).length;
  const fail   = inspections.filter(r => r['ผลการตรวจสอบ'] === RESULT_FAIL).length;

  const kpis = [
    { col: DASH_KPI_COLS.total, value: total, label: 'ทั้งหมด', color: '#1565C0', bg: '#E8EAF6' },
    { col: DASH_KPI_COLS.pass,  value: pass,  label: 'ผ่าน',    color: '#137333', bg: '#E6F4EA' },
    { col: DASH_KPI_COLS.fail,  value: fail,  label: 'ไม่ผ่าน', color: '#A50000', bg: '#FCE8E6' },
  ];

  dash.setRowHeight(DASH_ROW.KPI_LABEL, 18);
  dash.setRowHeight(DASH_ROW.KPI_VALUE, 72);

  kpis.forEach(({ col, value, label, color, bg }) => {
    // Label row
    dash.getRange(DASH_ROW.KPI_LABEL, col, 1, 2).merge()
      .setValue(label)
      .setBackground(bg).setFontColor('#5F6368').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');

    // Value row
    dash.getRange(DASH_ROW.KPI_VALUE, col, 1, 2).merge()
      .setValue(value)
      .setBackground(bg)
      .setFontColor(color).setFontSize(36).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false,
        color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });

  // Rows 6–7: spacers
  dash.setRowHeight(6, 12);
  dash.setRowHeight(7, 6);
}

// ================================================================
//  Building Summary Table
// ================================================================

/**
 * สร้างตารางสรุปรายอาคาร
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dash
 * @returns {number} row แรกที่ว่างหลังตาราง
 */
function updateSummaryByBuilding(dash) {
  const START = DASH_ROW.BUILD_HDR;  // Row 8

  // Section header
  _writeSectionHeader(dash, START, 1, 6, '🏢 สรุปรายอาคาร', '#1565C0');
  dash.setRowHeight(START, 30);

  // Column headers
  const colHdr = ['อาคาร', 'ทั้งหมด', 'ผ่าน', 'ไม่ผ่าน', '% ผ่าน', ''];
  dash.getRange(START + 1, 1, 1, colHdr.length).setValues([colHdr])
    .setBackground('#1A237E').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(10).setHorizontalAlignment('center');
  dash.setRowHeight(START + 1, 24);

  // Build location → building lookup
  const floors  = getAllFloors();
  const locs    = getAllLocations();
  const flBldMap = {};
  floors.forEach(f => { flBldMap[f['id']] = f['อาคาร_id']; });
  const locBldMap = {};
  locs.forEach(l => { locBldMap[l['id']] = flBldMap[l['ชั้น_id']]; });

  const buildings   = getAllBuildings();
  const inspections = getAllInspections();

  let row = DASH_ROW.BUILD_DATA;  // Row 10
  buildings.forEach((bld, i) => {
    const bIns  = inspections.filter(r => locBldMap[r['สถานที่_id']] === bld['id']);
    const total  = bIns.length;
    const pass   = bIns.filter(r => r['ผลการตรวจสอบ'] === RESULT_PASS).length;
    const fail   = bIns.filter(r => r['ผลการตรวจสอบ'] === RESULT_FAIL).length;
    const pct    = total > 0 ? Math.round(pass / total * 100) : 0;
    const bg     = i % 2 === 0 ? '#FFFFFF' : '#F8F9FA';

    dash.getRange(row, 1, 1, 6).setValues([[
      bld['ชื่ออาคาร'], total, pass, fail, `${pct}%`, '',
    ]]).setBackground(bg).setFontSize(10);

    dash.getRange(row, 1).setHorizontalAlignment('left');
    [2, 3, 4, 5].forEach(c => dash.getRange(row, c).setHorizontalAlignment('center'));

    if (pass > 0) dash.getRange(row, 3).setFontColor('#137333').setFontWeight('bold');
    if (fail > 0) dash.getRange(row, 4).setFontColor('#A50000').setFontWeight('bold');

    const pctColor = pct >= 80 ? '#137333' : pct >= 60 ? '#E37400' : '#A50000';
    dash.getRange(row, 5).setFontColor(pctColor).setFontWeight('bold');
    dash.setRowHeight(row, 22);
    row++;
  });

  if (buildings.length === 0) {
    dash.getRange(row, 1, 1, 5).merge()
      .setValue('— ยังไม่มีข้อมูลอาคาร —')
      .setFontColor('#5F6368').setHorizontalAlignment('center').setFontSize(10);
    row++;
  }

  return row;  // แถวแรกที่ว่าง
}

// ================================================================
//  Alert Tables
// ================================================================

/**
 * ตาราง อุปกรณ์ที่เกินกำหนดตรวจสอบ
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dash
 * @param {number} startRow
 * @returns {number} แถวแรกที่ว่างหลังตาราง
 */
function updateOverdueList(dash, startRow) {
  const today    = _today();
  const overdue  = _getLatestPerEquipLoc()
    .filter(r => {
      const nd = parseDate(r['วันที่ตรวจสอบครั้งถัดไป']);
      return nd && nd < today;
    })
    .map(r => {
      const nd   = parseDate(r['วันที่ตรวจสอบครั้งถัดไป']);
      const days = Math.floor((today - nd) / 86400000);
      return {
        อุปกรณ์:   _eqName(r['อุปกรณ์_id']),
        สถานที่:   _locDisplay(r['สถานที่_id']),
        ผลล่าสุด:  r['ผลการตรวจสอบ'],
        วันถัดไป:  formatDate(nd),
        เกิน:      `⚠️ ${days} วัน`,
      };
    })
    .sort((a, b) => parseInt(b.เกิน) - parseInt(a.เกิน));

  return _writeAlertTable(dash, startRow, {
    header:     '⚠️ อุปกรณ์ที่เกินกำหนดตรวจสอบ',
    headerBg:   '#7B0000',
    colHeaders: ['อุปกรณ์', 'สถานที่', 'ผลล่าสุด', 'วันตรวจถัดไป', 'เกินกำหนด'],
    data:       overdue,
    fields:     ['อุปกรณ์', 'สถานที่', 'ผลล่าสุด', 'วันถัดไป', 'เกิน'],
    rowBg:      '#FCE8E6',
    rowColor:   '#7B0000',
    emptyMsg:   '✅ ไม่มีอุปกรณ์ที่เกินกำหนด',
    emptyColor: '#137333',
  });
}

/**
 * ตาราง อุปกรณ์ที่ใกล้ครบกำหนด (≤ 7 วัน)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dash
 * @param {number} startRow
 * @returns {number} แถวแรกที่ว่างหลังตาราง
 */
function updateUpcomingList(dash, startRow) {
  const today      = _today();
  const upcoming30 = new Date(today);
  upcoming30.setDate(upcoming30.getDate() + parseInt(getConfig(CONFIG_UPCOMING_DAYS, '7')));

  const upcoming = _getLatestPerEquipLoc()
    .filter(r => {
      const nd = parseDate(r['วันที่ตรวจสอบครั้งถัดไป']);
      return nd && nd >= today && nd <= upcoming30;
    })
    .map(r => {
      const nd   = parseDate(r['วันที่ตรวจสอบครั้งถัดไป']);
      const days = Math.floor((nd - today) / 86400000);
      return {
        อุปกรณ์:  _eqName(r['อุปกรณ์_id']),
        สถานที่:  _locDisplay(r['สถานที่_id']),
        ผลล่าสุด: r['ผลการตรวจสอบ'],
        วันถัดไป: formatDate(nd),
        เหลือ:    `🔔 ${days} วัน`,
      };
    })
    .sort((a, b) => parseInt(a.เหลือ) - parseInt(b.เหลือ));

  return _writeAlertTable(dash, startRow, {
    header:     '🔔 อุปกรณ์ที่ใกล้ครบกำหนด (≤ 7 วัน)',
    headerBg:   '#FF6D00',
    colHeaders: ['อุปกรณ์', 'สถานที่', 'ผลล่าสุด', 'วันตรวจถัดไป', 'เหลืออีก'],
    data:       upcoming,
    fields:     ['อุปกรณ์', 'สถานที่', 'ผลล่าสุด', 'วันถัดไป', 'เหลือ'],
    rowBg:      '#FEF7E0',
    rowColor:   '#E37400',
    emptyMsg:   '✅ ไม่มีอุปกรณ์ที่ใกล้ครบกำหนด',
    emptyColor: '#137333',
  });
}

// ================================================================
//  Email Alert
// ================================================================

/**
 * ส่ง email แจ้งเตือนรายการที่เกินกำหนด
 * ต้องตั้งค่า CONFIG_ALERT_EMAIL ก่อน
 */
function sendOverdueAlert() {
  const email = getConfig(CONFIG_ALERT_EMAIL);
  if (!email) {
    console.log('sendOverdueAlert: CONFIG_ALERT_EMAIL ยังไม่ได้ตั้งค่า — ข้าม');
    return;
  }

  const today   = _today();
  const overdue = _getLatestPerEquipLoc().filter(r => {
    const nd = parseDate(r['วันที่ตรวจสอบครั้งถัดไป']);
    return nd && nd < today;
  });

  if (overdue.length === 0) {
    console.log('sendOverdueAlert: ไม่มีรายการที่เกินกำหนด');
    return;
  }

  const tz  = Session.getScriptTimeZone();
  const now = Utilities.formatDate(today, tz, 'dd/MM/yyyy');

  let html = `<h2 style="color:#1A237E">🔥 รายงานอุปกรณ์ดับเพลิงเกินกำหนด</h2>`;
  html += `<p>วันที่ <strong>${now}</strong> | จำนวน: <strong style="color:#A50000">${overdue.length} รายการ</strong></p>`;
  html += `<table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;font-size:13px;font-family:sans-serif">`;
  html += `<tr style="background:#1A237E;color:#fff">
    <th>อุปกรณ์</th><th>สถานที่</th><th>ผลล่าสุด</th>
    <th>วันตรวจถัดไป</th><th>เกินกำหนด (วัน)</th>
  </tr>`;

  overdue.forEach(r => {
    const nd   = parseDate(r['วันที่ตรวจสอบครั้งถัดไป']);
    const days = Math.floor((today - nd) / 86400000);
    html += `<tr style="background:#FCE8E6;color:#7B0000">
      <td>${_eqName(r['อุปกรณ์_id'])}</td>
      <td>${_locDisplay(r['สถานที่_id'])}</td>
      <td>${r['ผลการตรวจสอบ']}</td>
      <td>${formatDate(nd)}</td>
      <td style="text-align:center;font-weight:bold">⚠️ ${days}</td>
    </tr>`;
  });

  html += `</table>`;
  html += `<p style="color:#5F6368;font-size:11px;margin-top:16px">
    — ระบบตรวจสอบอุปกรณ์ดับเพลิง (Auto-alert via Google Apps Script)
  </p>`;

  MailApp.sendEmail({
    to:       email,
    subject:  `⚠️ อุปกรณ์ดับเพลิงเกินกำหนด ${overdue.length} รายการ (${now})`,
    htmlBody: html,
  });

  console.log(`sendOverdueAlert: ส่ง email ถึง ${email} จำนวน ${overdue.length} รายการ`);
}

// ================================================================
//  Private Helpers
// ================================================================

/** วันนี้ (00:00:00 local time) */
function _today() {
  const d = new Date();
  d.setHours(0, 0, 0, 0);
  return d;
}

/**
 * ดึงรายการตรวจสอบล่าสุดต่อ (สถานที่ + อุปกรณ์)
 * ใช้ group by key แล้วเก็บ record ที่ วันที่ตรวจสอบ ใหม่ที่สุด
 * @returns {Object[]}
 */
function _getLatestPerEquipLoc() {
  const map = {};
  getAllInspections().forEach(r => {
    const key = `${r['สถานที่_id']}__${r['อุปกรณ์_id']}`;
    const d   = parseDate(r['วันที่ตรวจสอบ']);
    if (!d) return;
    if (!map[key] || d > parseDate(map[key]['วันที่ตรวจสอบ'])) {
      map[key] = r;
    }
  });
  return Object.values(map);
}

// Cached lookup maps (rebuilt on each refreshDashboard call)
let _cacheEq  = null;
let _cacheLoc = null;
let _cacheFl  = null;
let _cacheBld = null;

function _ensureCache() {
  if (!_cacheEq) {
    _cacheEq  = {};  getAllEquipment().forEach(e  => { _cacheEq[e['id']]  = e; });
    _cacheLoc = {};  getAllLocations().forEach(l  => { _cacheLoc[l['id']] = l; });
    _cacheFl  = {};  getAllFloors().forEach(f     => { _cacheFl[f['id']]  = f; });
    _cacheBld = {};  getAllBuildings().forEach(b  => { _cacheBld[b['id']] = b; });
  }
}

/** คืนชื่ออุปกรณ์จาก id */
function _eqName(id) {
  _ensureCache();
  const e = _cacheEq[id];
  return e ? e['ชื่ออุปกรณ์'] : id;
}

/** คืน "อาคาร X ชั้น Y — สถานที่" จาก locationId */
function _locDisplay(locId) {
  _ensureCache();
  const loc = _cacheLoc[locId];
  if (!loc) return locId;
  const fl  = _cacheFl[loc['ชั้น_id']];
  const bld = fl ? _cacheBld[fl['อาคาร_id']] : null;
  return [
    bld ? bld['ชื่ออาคาร'] : '',
    fl  ? fl['ชื่อชั้น']   : '',
    '—',
    loc['ชื่อสถานที่'],
  ].filter(Boolean).join(' ');
}

/**
 * เขียน Section header (merged, colored)
 * @param {Sheet} dash
 * @param {number} row
 * @param {number} colStart 1-based
 * @param {number} colEnd   1-based
 * @param {string} text
 * @param {string} bgColor
 */
function _writeSectionHeader(dash, row, colStart, colEnd, text, bgColor) {
  dash.getRange(row, colStart, 1, colEnd - colStart + 1).merge()
    .setValue(text)
    .setBackground(bgColor).setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(12)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  dash.setRowHeight(row, 30);
}

/**
 * เขียนตาราง alert (overdue หรือ upcoming)
 * @returns {number} แถวแรกที่ว่างหลังตาราง
 */
function _writeAlertTable(dash, startRow, opts) {
  // Cache ล้างทุกครั้งที่เขียนตาราง
  _cacheEq = _cacheLoc = _cacheFl = _cacheBld = null;

  _writeSectionHeader(dash, startRow, 1, opts.colHeaders.length,
    opts.header, opts.headerBg);
  let row = startRow + 1;

  // Column headers
  dash.getRange(row, 1, 1, opts.colHeaders.length)
    .setValues([opts.colHeaders])
    .setBackground('#1A237E').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  dash.setRowHeight(row, 24);
  row++;

  if (!opts.data || opts.data.length === 0) {
    dash.getRange(row, 1, 1, opts.colHeaders.length).merge()
      .setValue(opts.emptyMsg)
      .setFontColor(opts.emptyColor || '#137333')
      .setHorizontalAlignment('center').setFontSize(10);
    dash.setRowHeight(row, 22);
    return row + 1;
  }

  opts.data.forEach((item, i) => {
    const values = opts.fields.map(f => item[f] || '');
    dash.getRange(row, 1, 1, values.length).setValues([values])
      .setBackground(opts.rowBg).setFontColor(opts.rowColor).setFontSize(10);
    // Last column (เกิน/เหลือ): center + bold
    dash.getRange(row, values.length)
      .setHorizontalAlignment('center').setFontWeight('bold');
    dash.setRowHeight(row, 22);
    row++;
  });

  return row;
}
