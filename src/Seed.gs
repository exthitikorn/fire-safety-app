// ================================================================
//  Seed.gs — นำเข้าข้อมูลจริง (อาคาร / ชั้น / สถานที่)
//  ข้อมูลจาก building.xlsx + floordepartment.xlsx
// ================================================================

// ── Menu wrapper (เรียกจาก Sheets menu) ─────────────────────────
function seedRealData() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.alert(
    '🏥 นำเข้าข้อมูลจริง',
    'จะลบข้อมูล อาคาร / ชั้น / สถานที่ ทั้งหมด แล้วนำเข้าใหม่\n\nดำเนินการต่อ?',
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;
  try {
    const counts = _doSeedRealData();
    ui.alert('✅ นำเข้าข้อมูลสำเร็จ',
      `อาคาร ${counts.buildings} | ชั้น ${counts.floors} | สถานที่ ${counts.locations}`,
      ui.ButtonSet.OK);
  } catch (err) {
    ui.alert('❌ เกิดข้อผิดพลาด', err.message, ui.ButtonSet.OK);
  }
}

// ── Web App wrapper (เรียกจาก google.script.run) ────────────────
function webSeedRealData() {
  try {
    const counts = _doSeedRealData();
    return {
      success: true,
      message: `อาคาร ${counts.buildings} | ชั้น ${counts.floors} | สถานที่ ${counts.locations}`,
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── Core logic ───────────────────────────────────────────────────
function _doSeedRealData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ล้างข้อมูล (คง header row 1 ไว้)
  [SHEET_LOCATION, SHEET_FLOOR, SHEET_BUILDING].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  });

  // ── อาคาร ────────────────────────────────────────────────────
  const BUILDINGS = [
    { id: 'cmi3vc92l00004e9mddtuket7', name: 'อาคารสมเด็จพระสังฆราช',            floors: 4 },
    { id: 'cmi3vcj1l00014e9mfxs7f3rm', name: 'อาคารอายุบวร',                      floors: 1 },
    { id: 'cmi3vcooy00024e9mvhca06ao', name: 'อาคารเฉลิมพระเกียรติ',              floors: 4 },
    { id: 'cmi3vd0mf00034e9m5do1s9oq', name: 'อาคารภูมิพิพัฒน์',                 floors: 7 },
    { id: 'cmi3vd8o600044e9ms0lj8p4w', name: 'อาคารเมตตาธรรม',                   floors: 4 },
    { id: 'cmi3vi9kc00054e9mzkrwtdht', name: 'อาคารพระเทพประสิทธิมล 71 (5ไร่)',  floors: 3 },
    { id: 'cmi3vipgg00064e9mlkh4qk65', name: 'อาคารคิลานุปัฏฐาก (5ไร่)',         floors: 1 },
    { id: 'cmi3viv4g00074e9m5a88klfu', name: 'อาคาร Modular 1 (5ไร่)',           floors: 1 },
    { id: 'cmicv3iai00014e73h3jkx13s', name: 'โรงพยาบาลอื่น',                    floors: 1 },
    { id: 'cmqrr5mn00b00s4diad211jne', name: 'อาคาร Modular 2 (5ไร่)',           floors: 1 },
    { id: 'cmqrr6a3l0b03s4di2ntmit0z', name: 'อาคาร Modular 3 (5ไร่)',           floors: 1 },
    { id: 'cmqrr65qi0b02s4digmoew9ma', name: 'อาคาร Modular 4 (5ไร่)',           floors: 1 },
  ];

  const bSheet = ss.getSheetByName(SHEET_BUILDING);
  BUILDINGS.forEach(b => bSheet.appendRow([b.id, b.name]));

  // ── ชั้น — สร้าง 1..N ต่ออาคาร ──────────────────────────────
  const floorMap = {}; // `${buildingId}_${n}` → uuid
  const fSheet   = ss.getSheetByName(SHEET_FLOOR);
  BUILDINGS.forEach(b => {
    for (let n = 1; n <= b.floors; n++) {
      const fid = Utilities.getUuid();
      floorMap[b.id + '_' + n] = fid;
      fSheet.appendRow([fid, b.id, 'ชั้น ' + n]);
    }
  });

  // ── สถานที่ — floor ว่าง → ชั้น 1 ───────────────────────────
  const LOCATIONS = [
    // อาคารสมเด็จพระสังฆราช
    { id: 'cmi3vd0mf00034e9m5do1s9oq', name: 'หอผู้ป่วยพิเศษพรีเมียม',                                   bId: 'cmi3vc92l00004e9mddtuket7', floor: 4 },
    { id: 'cmi3xj038000b4e9m6c3sb14w', name: 'หอผู้ป่วยอายุรกรรมชาย',                                    bId: 'cmi3vc92l00004e9mddtuket7', floor: 3 },
    { id: 'cmi3xj9qa000d4e9mgxhe00ud', name: 'หอผู้ป่วยอายุรกรรมหญิง',                                   bId: 'cmi3vc92l00004e9mddtuket7', floor: 3 },
    { id: 'cmi3xji6g000f4e9mryar326x', name: 'หอผู้ป่วยศัลยกรรมชาย',                                    bId: 'cmi3vc92l00004e9mddtuket7', floor: 2 },
    { id: 'cmi3xjtgm000h4e9m85brrmce', name: 'หอผู้ป่วยศัลยกรรมหญิง',                                   bId: 'cmi3vc92l00004e9mddtuket7', floor: 2 },
    { id: 'cmi3xk8ck000j4e9mwdkt0pin', name: 'หอผู้ปวยวิกฤต (ICU)',                                      bId: 'cmi3vc92l00004e9mddtuket7', floor: 1 },
    { id: 'cmi3xkfi9000l4e9me23gt180', name: 'ห้องนิรมัย',                                               bId: 'cmi3vc92l00004e9mddtuket7', floor: 1 },
    { id: 'cmipq0c6x002c4esdyf7mwr42', name: 'จุดรับ-ส่งผู้ป่วย',                                        bId: 'cmi3vc92l00004e9mddtuket7', floor: 1 },
    // อาคารอายุบวร
    { id: 'cmi3xkrgu000n4e9mxo9dayxs', name: 'คลินิกผู้สูงอายุคุณภาพ (คลินิกสุขใจ สูงวัย ประคับประคอง)', bId: 'cmi3vcj1l00014e9mfxs7f3rm', floor: 1 },
    { id: 'cmi3xkwq7000p4e9m2ienw7my', name: 'คลินิกกายภาพผู้สูงอายุ',                                  bId: 'cmi3vcj1l00014e9mfxs7f3rm', floor: 1 },
    { id: 'cmi3xl39w000r4e9m5o0uu980', name: 'คลินิกแพทย์แผนไทยและแพทย์ทางเลือก',                      bId: 'cmi3vcj1l00014e9mfxs7f3rm', floor: 1 },
    { id: 'cmi3xlmgf000v4e9mj7e2eapg', name: 'คลินิกให้คำปรึกษา/ARV clinic',                            bId: 'cmi3vcj1l00014e9mfxs7f3rm', floor: 1 },
    { id: 'cmi3xlrl5000x4e9mxa61yxjg', name: 'เรือนบุญ',                                                bId: 'cmi3vcj1l00014e9mfxs7f3rm', floor: 1 },
    { id: 'cmme9yait01b80hdi72o28iot', name: 'หอผู้ป่วยพิเศษศัลยกรรม',                                  bId: 'cmi3vcj1l00014e9mfxs7f3rm', floor: 1 },
    // อาคารเฉลิมพระเกียรติ
    { id: 'cmi3xlzpe000z4e9mqp5qs4ll', name: 'คลินิกพรีเมียม',                                          bId: 'cmi3vcooy00024e9mvhca06ao', floor: 4 },
    { id: 'cmi3xm7o700114e9m6nupza43', name: 'ห้องเจาะเลือด',                                           bId: 'cmi3vcooy00024e9mvhca06ao', floor: 2 },
    { id: 'cmi3xmejb00134e9mg68dsp6s', name: 'ห้อง X-RAY',                                              bId: 'cmi3vcooy00024e9mvhca06ao', floor: 2 },
    { id: 'cmi3xml0s00154e9mnrow3qd2', name: 'คลินิกทันตกรรม (ห้องฟัน)',                                bId: 'cmi3vcooy00024e9mvhca06ao', floor: 2 },
    { id: 'cmi3xmqip00174e9me7yqnrvt', name: 'คลินิกกุมารเวชกรรม',                                      bId: 'cmi3vcooy00024e9mvhca06ao', floor: 2 },
    { id: 'cmi3xmwp100194e9m1syibgl2', name: 'คลินิกต่างชาติ',                                          bId: 'cmi3vcooy00024e9mvhca06ao', floor: 2 },
    { id: 'cmi3xn84b001d4e9m275q43rj', name: 'คลินิกเบิกได้จ่ายตรง',                                    bId: 'cmi3vcooy00024e9mvhca06ao', floor: 1 },
    { id: 'cmi3xndk5001f4e9mdv2p4u5k', name: 'อนุมัติสิทธิ – ส่งต่อ',                                  bId: 'cmi3vcooy00024e9mvhca06ao', floor: 1 },
    { id: 'cmi3xnj7j001h4e9m2yqr2rlb', name: 'คลินิกตรวจโรคทั่วไป',                                    bId: 'cmi3vcooy00024e9mvhca06ao', floor: 1 },
    { id: 'cmi3xnq31001j4e9mcfjuv1ze', name: 'คลินิกจักษุ',                                             bId: 'cmi3vcooy00024e9mvhca06ao', floor: 1 },
    { id: 'cmi3xnxwn001l4e9mhkmknxo4', name: 'โสต ศอ นาสิก',                                           bId: 'cmi3vcooy00024e9mvhca06ao', floor: 1 },
    { id: 'cmi3xo50p001n4e9mcc1tkdxf', name: 'ห้องจ่ายยาผู้ป่วยนอก',                                   bId: 'cmi3vcooy00024e9mvhca06ao', floor: 1 },
    { id: 'cmi3xob1j001p4e9md4c0e2sp', name: 'ศูนย์บริการคนพิการแบบเบ็ดเสร็จ',                         bId: 'cmi3vcooy00024e9mvhca06ao', floor: 1 },
    { id: 'cmippqbiv00294esd6um8jcss', name: 'จุดรับ-ส่งผู้ป่วย',                                        bId: 'cmi3vcooy00024e9mvhca06ao', floor: 1 },
    { id: 'cmlytiu6s01bh1hdi3c5nu4kq', name: 'จุดคัดกรอง OPD',                                          bId: 'cmi3vcooy00024e9mvhca06ao', floor: 1 },
    // อาคารภูมิพิพัฒน์
    { id: 'cmi3xomy9001r4e9mr73k4jfp', name: 'หอผู้ป่วยพิเศษพรีเมียม',                                  bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 7 },
    { id: 'cmi3xps4u001t4e9my205mew7', name: 'หอผู้ป่วย ปกส.ชาย',                                     bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 6 },
    { id: 'cmi3xq0sh001v4e9mmx3bu30f', name: 'หอผู้ป่วย ปกส.หญิง',                                    bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 6 },
    { id: 'cmi4b24h70001ir09gho8cvc0', name: 'Stroke Unit',                                              bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 6 },
    { id: 'cmi3xqaar001x4e9m4v7fbm1u', name: 'หอผู้ป่วยสูติ – นรีเวชกรรม',                            bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 5 },
    { id: 'cmi3xqklk001z4e9mijsxn47a', name: 'หอผู้ป่วยกุมารเวชกรรม',                                  bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 5 },
    { id: 'cmi3xquh900214e9mb4s6lpey', name: 'PICU',                                                     bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 5 },
    { id: 'cmi3xr69j00234e9mjjo16b9q', name: 'ห้องผ่าตัด',                                              bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 4 },
    { id: 'cmi3xrzbk00254e9mxxyd1mu9', name: 'หอผู้ป่วยวิกฤต 2 (ICU2)',                                bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 4 },
    { id: 'cmi3xs7un00274e9mwwtv7dq6', name: 'หอผู้ป่วยวิกฤตเด็ก (NICU)',                              bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 4 },
    { id: 'cmi3xsfvq00294e9m5vup36ey', name: 'หอผู้ป่วยทารกแรกเกิด',                                   bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 4 },
    { id: 'cmi3xta8b002d4e9m9gqlpdi6', name: 'คลินิกสูติ – นรีเวชกรรม',                               bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 3 },
    { id: 'cmi3xth27002f4e9mfxq0leym', name: 'คลินิกศัลยกรรมกระดูก',                                   bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 3 },
    { id: 'cmi3xtpz7002h4e9mz99j7q7m', name: 'ห้องคลอด',                                               bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 3 },
    { id: 'cmi3xtwnt002j4e9m7rl01qlv', name: 'ห้องผ่าตัดเล็ก',                                         bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 3 },
    { id: 'cmi3xu9c2002l4e9mfy2sok8j', name: 'ศูนย์ประสานงานการผ่าตัดแบบวันเดียวกลับ (ODS)',          bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 3 },
    { id: 'cmi3xugfi002n4e9masiaumk4', name: 'คลินิกอายุรกรรม',                                         bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 2 },
    { id: 'cmi3xumc0002p4e9mgh6p5z9f', name: 'คลินิกศัลยกรรม',                                         bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 2 },
    { id: 'cmi3xurgi002r4e9manetm4fd', name: 'ห้องฉีดยา - ทำแผล',                                      bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 2 },
    { id: 'cmi3xuwk5002t4e9m0xtmailu', name: 'ห้องจ่ายยา',                                             bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 2 },
    { id: 'cmiqu6cp9001p4e6nd0mqtdwu', name: 'ห้องฉีดยาทำแผล',                                         bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 2 },
    { id: 'cmi3xv6bc002v4e9mqrlnehb2', name: 'ห้องตรวจอุบัติเหตุ-ฉุกเฉิน (ER)',                       bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 1 },
    { id: 'cmi3xvfgf002x4e9mhxhbzhjc', name: 'ห้องจ่ายยาผู้ป่วยใน – อุบัติเหตุ',                     bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 1 },
    { id: 'cmi3xvq35002z4e9mpmqs60z9', name: 'การเงิน',                                                bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 1 },
    { id: 'cmi3xw1xc00314e9mebqrbsqg', name: 'Health Tech',                                             bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 1 },
    { id: 'cmi3xw7hf00334e9mz09mjqcj', name: 'คลินิก ARI',                                             bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 1 },
    { id: 'cmi3xwfkm00354e9mchfvqxzm', name: 'Admission Center',                                        bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 1 },
    { id: 'cmi3xwlsu00374e9mia93ihzo', name: 'UMSC',                                                    bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 1 },
    { id: 'cmipq0klj002e4esde7a33jkg', name: 'จุดรับ-ส่งผู้ป่วย',                                       bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 1 },
    { id: 'cmkq7qdzr03s14ewt8pty90fe', name: 'ห้องสังเกตุอาการ (Observation Room)',                     bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 1 },
    { id: 'cml4uj0dm05si4ewt1djiyyvk', name: 'คลินิกประกันสังคม',                                      bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 1 },
    { id: 'cmlytjqn501bj1hdiy3s4z6h2', name: 'จุดคัดกรอง ER',                                          bId: 'cmi3vd0mf00034e9m5do1s9oq', floor: 1 },
    // อาคารเมตตาธรรม
    { id: 'cmi3xwvok00394e9m2ph91zmn', name: 'กลุ่มงานเวชกรรมฟื้นฟู (กายภาพบำบัด)',                   bId: 'cmi3vd8o600044e9ms0lj8p4w', floor: 4 },
    { id: 'cmi3xx2vb003b4e9mstn5zpe7', name: 'พยาธิวิทยากายวิภาค',                                     bId: 'cmi3vd8o600044e9ms0lj8p4w', floor: 4 },
    { id: 'cmi3xx9ap003d4e9mhvvszp7b', name: 'ศูนย์เด็กเล็กน่าอยู่คู่นมแม่ (Day Care)',               bId: 'cmi3vd8o600044e9ms0lj8p4w', floor: 4 },
    { id: 'cmi3xxknj003f4e9mlmvm5quv', name: 'ศูนย์ไตเทียม',                                           bId: 'cmi3vd8o600044e9ms0lj8p4w', floor: 3 },
    { id: 'cmi3xxrxc003h4e9m7sok8qio', name: 'ศูนย์ไตเทียมพรีเมียม',                                   bId: 'cmi3vd8o600044e9ms0lj8p4w', floor: 3 },
    { id: 'cmi3xy3d4003j4e9mcvcy3fci', name: 'หอผู้ปวยชีวาภิบาล 1',                                   bId: 'cmi3vd8o600044e9ms0lj8p4w', floor: 3 },
    { id: 'cmi3xyby2003l4e9mzxrhamu8', name: 'Sleep Lab',                                               bId: 'cmi3vd8o600044e9ms0lj8p4w', floor: 3 },
    { id: 'cmi3xyjeq003n4e9m3f9z1ohz', name: 'หอผู้ปวยเคมีบำบัด',                                     bId: 'cmi3vd8o600044e9ms0lj8p4w', floor: 2 },
    { id: 'cmi3xyq3o003p4e9msmz7bc0i', name: 'คลินิกจิตวิทยา',                                         bId: 'cmi3vd8o600044e9ms0lj8p4w', floor: 2 },
    { id: 'cmi3xyzma003r4e9m8caj2fma', name: 'ห้อง CT, ห้อง MRI',                                      bId: 'cmi3vd8o600044e9ms0lj8p4w', floor: 1 },
    { id: 'cmi3xz7ho003t4e9myoule2jp', name: 'อนุมัติสิทธิ - ส่งต่อ',                                 bId: 'cmi3vd8o600044e9ms0lj8p4w', floor: 1 },
    { id: 'cmi3xzdj8003v4e9mm1e5uzxt', name: 'การเงิน',                                                bId: 'cmi3vd8o600044e9ms0lj8p4w', floor: 1 },
    { id: 'cmipq0os5002g4esd5qf1x5dm', name: 'จุดรับ-ส่งผู้ป่วย',                                       bId: 'cmi3vd8o600044e9ms0lj8p4w', floor: 1 },
    // อาคารพระเทพประสิทธิมล 71 (5ไร่)
    { id: 'cmi3y0pyi003x4e9mvln8iot7', name: 'หอผู้ป่วยอายุกรรมรวม',                                  bId: 'cmi3vi9kc00054e9mzkrwtdht', floor: 3 },
    { id: 'cmqrrius80b0js4dizg1a33fm', name: 'หอผู้ป่วยอายุรกรรมรวม2',                                 bId: 'cmi3vi9kc00054e9mzkrwtdht', floor: 3 },
    { id: 'cmi3y12uf003z4e9mnjydqs90', name: 'หอผู้ป่วย Intermediate Care',                            bId: 'cmi3vi9kc00054e9mzkrwtdht', floor: 2 },
    { id: 'cmqrreo1a0b0es4didnnlyquc', name: 'หอผู้ป่วยระยะฟื้นฟู (IMC)',                              bId: 'cmi3vi9kc00054e9mzkrwtdht', floor: 2 },
    { id: 'cmi3y1c5p00414e9mukcgou9p', name: 'หอผู้ป่วยชีวาภิบาล 2',                                  bId: 'cmi3vi9kc00054e9mzkrwtdht', floor: 1 },
    { id: 'cmi3y1lc900434e9mdzyhalqt', name: 'ICU Palliative',                                          bId: 'cmi3vi9kc00054e9mzkrwtdht', floor: 1 },
    { id: 'cmqrrcrrl0b0cs4diu4ep4yfv', name: 'หอผู้ป่วยวิกฤต (ICU 3)',                                 bId: 'cmi3vi9kc00054e9mzkrwtdht', floor: 1 },
    // อาคารคิลานุปัฏฐาก (5ไร่)
    { id: 'cmi3y1w5p00454e9mdgbura91', name: 'หอผู้ป่วยวชิรคุณาธาร',                                  bId: 'cmi3vipgg00064e9mlkh4qk65', floor: 1 },
    // อาคาร Modular 1 (5ไร่)
    { id: 'cmi3y24uk00474e9mlklqiv5e', name: 'หน่วยกายภาพบำบัด',                                      bId: 'cmi3viv4g00074e9m5a88klfu', floor: 1 },
    { id: 'cmi3y2bnc00494e9mzkri15jo', name: 'หอผู้ป่วยติดเชื้อทางเดินหายใจ',                         bId: 'cmi3viv4g00074e9m5a88klfu', floor: 1 },
    // อาคาร Modular 2 (5ไร่)
    { id: 'cmqrrjwlv0b0ks4dind6uvo8w', name: 'หอผู้ป่วยวชิรคุณาธาร',                                  bId: 'cmqrr5mn00b00s4diad211jne',  floor: 1 },
    // อาคาร Modular 3 (5ไร่)
    { id: 'cmqrrkjyv0b0ns4di2o9gdctt', name: 'หอผู้ป่วยระบบทางเดินหายใจ',                             bId: 'cmqrr6a3l0b03s4di2ntmit0z',  floor: 1 },
    // อาคาร Modular 4 (5ไร่)
    { id: 'cmqrrl5qz0b0os4dimhts1cty', name: 'กายภาพบำบัด',                                           bId: 'cmqrr65qi0b02s4digmoew9ma',  floor: 1 },
  ];

  const lSheet = ss.getSheetByName(SHEET_LOCATION);
  let locCount = 0;
  LOCATIONS.forEach(loc => {
    const floorId = floorMap[loc.bId + '_' + loc.floor];
    if (!floorId) return;
    lSheet.appendRow([loc.id, floorId, loc.name]);
    locCount++;
  });

  clearAllSheetCache();  // seed เขียนชีตตรงๆ — ล้าง cache ทั้งหมด

  return {
    buildings: BUILDINGS.length,
    floors:    Object.keys(floorMap).length,
    locations: locCount,
  };
}
