// ═══════════════════════════════════════════════════════════════════════════
// Kyoto Shi Café — Google Apps Script Backend  v1.0
// วางโค้ดนี้ใน Google Apps Script แล้ว Deploy เป็น Web App
// ═══════════════════════════════════════════════════════════════════════════

// ▶ ขั้นตอนแรก: วาง SHEET_ID ของคุณด้านล่างนี้ (ดูจาก URL ของ Google Sheet)
const SHEET_ID = '1xN_LYVo662oTu1qlMzCkwEiN4v4kmnUya5fn_vhXaLg';

// Sheet names
const SH = {
  EMPLOYEES:    'Employees',
  ATTENDANCE:   'Attendance',
  SCHEDULES:    'Schedules',
  OT_REQUESTS:  'OT_Requests',
  LEAVE:        'Leave_Requests',
};

// ── Helper: ส่งข้อมูลกลับเป็น JSON ──
function res(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function openSS() {
  return SpreadsheetApp.openById(SHEET_ID);
}

// ════════════════════════════════════════════════════════════
// doGet — จัดการทุก request ผ่าน GET (ไม่มีปัญหา CORS)
// ทั้งการอ่านและเขียนข้อมูลใช้ GET เพื่อความเสถียร
// ════════════════════════════════════════════════════════════
function doGet(e) {
  const p = e.parameter || {};
  try {
    switch (p.action) {
      // ── อ่านข้อมูล ──
      case 'login':           return res(login(p.user, p.pass));
      case 'getEmployees':    return res(getEmployees());
      case 'getAttendance':   return res(getAttendance(p.empId, p.month, p.year, p.limit));
      case 'getSchedules':    return res(getSchedules(p.empId));
      case 'getDayRoster':    return res(getDayRoster(p.date));
      // ── Sidework ──
      case 'getSideworkBoard':       return res(getSideworkBoard(p.empId));
      case 'getSideworkTasks':       return res(getSideworkTasks(p.kitchen));
      case 'saveSideworkTask':       return res(saveSideworkTask(p));
      case 'deleteSideworkTask':     return res(deleteSideworkTask(p));
      case 'submitSidework':         return res(submitSidework(p));
      case 'getSideworkLeaderboard': return res(getSideworkLeaderboard());
      case 'getSideworkDashboard':   return res(getSideworkDashboard());
      case 'getPrepList':            return res(getPrepList(p.kitchen));
      case 'savePrepItem':           return res(savePrepItem(p));
      case 'deletePrepItem':         return res(deletePrepItem(p));
      case 'getPrepLog':             return res(getPrepLog(p));
      case 'savePrepLog':            return res(savePrepLog(p));
      case 'getMorningPrep':         return res(getMorningPrep(p.kitchen));
      case 'getShortages':           return res(getShortages(p));
      case 'getPrepStatus':         return res(getPrepStatus());
      case 'saveMorningAck':        return res(saveMorningAck(p));
      case 'savePrepPlan':          return res(savePrepPlan(p));
      case 'getPending':           return res(getPendingApprovals());
      case 'getPendingApprovals':  return res(getPendingApprovals());
      case 'getTodayClock':        return res(getTodayClock(p.empId, p.date));
      case 'getTodayAttendance':   return res(getTodayAttendance(p.date));
      case 'getLeaveBalance':      return res(getLeaveBalance(p.empId, p.year));
      case 'getLeaveHistory':      return res(getLeaveHistory(p.empId, p.year));
      case 'getOT':                return res(getOT(p.empId));
      case 'getOTHistory':         return res(getOTHistory(p.empId, p.month, p.year));
      case 'ping':            return res({ ok: true, message: 'Kyoto Shi API online' });
      case 'migrateSchedules': return res(migrateSchedulesColumns());

      // ── บันทึกข้อมูล ──
      case 'clockIn':         return res(clockIn(p));
      case 'clockOut':        return res(clockOut(p));
      case 'submitSchedule':  return res(submitSchedule({
                                empId: p.empId,
                                month: parseInt(p.month),
                                year:  parseInt(p.year),
                                dates: JSON.parse(decodeURIComponent(p.dates || '{}')),
                              }));
      case 'approveSchedule': return res(approveSchedule({ scheduleId: p.scheduleId, approvedBy: p.approvedBy }));
      case 'rejectSchedule':  return res(rejectSchedule({ scheduleId: p.scheduleId, approvedBy: p.approvedBy }));
      case 'editSchedule':    return res(editSchedule({
                                empId:              p.empId,
                                originalScheduleId: p.originalScheduleId,
                                date:               p.date,
                                month:              parseInt(p.month),
                                year:               parseInt(p.year),
                                dates:              JSON.parse(decodeURIComponent(p.dates || '{}')),
                              }));
      case 'submitOT':        return res(submitOT(p));
      case 'submitLeave':     return res(submitLeave(p));
      case 'approveLeave':    return res(approveLeave({ leaveId: p.leaveId, approve: p.approve === 'true', approvedBy: p.approvedBy }));
      case 'approveOT':       return res(approveOT({ otId: p.otId, approve: p.approve === 'true', approvedBy: p.approvedBy }));
      case 'registerEmployee':return res(registerEmployee(JSON.parse(decodeURIComponent(p.data || '{}'))));
      case 'approveEmployee': return res(approveEmployee({ empId: p.empId, approve: p.approve === 'true', salary: p.salary, type: p.type, note: p.note, ptRate: p.ptRate, ptRateHol: p.ptRateHol }));
      case 'updateEmployee':  return res(updateEmployee(JSON.parse(decodeURIComponent(p.data || '{}'))));

      default:                return res({ ok: false, error: 'Unknown action: ' + p.action });
    }
  } catch (err) {
    return res({ ok: false, error: err.message });
  }
}

// doPost ยังคงรองรับไว้ (สำหรับ client อื่น)
function doPost(e) {
  return doGet(e);
}

// ════════════════════════════════════════════════════════════
// AUTH — เข้าสู่ระบบ
// ════════════════════════════════════════════════════════════
// Columns: [0]id [1]name [2]username [3]password [4]nickname
//          [5]position [6]type [7]salary [8]bank [9]startDate
//          [10]status [11]probationEnd
function login(username, password) {
  if (!username || !password) return { ok: false, error: 'กรุณากรอก Username และ Password' };
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.EMPLOYEES);
  const data = sh.getDataRange().getValues();
  const inputHash = hashPw(password);
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[2]) !== String(username)) continue;
    if (String(row[10]).trim().toLowerCase() !== 'active') continue;
    const stored        = String(row[3]);
    const isHashMatch   = (stored === inputHash);
    const isLegacyMatch = (stored === String(password)); // plaintext stored before hashing was added
    if (isHashMatch || isLegacyMatch) {
      // Auto-migrate legacy plaintext passwords to a hash on first successful login
      if (isLegacyMatch && !isHashMatch) sh.getRange(i + 1, 4).setValue(inputHash);
      const isOwner = row[6] === 'Owner';
      return {
        ok:       true,
        role:     isOwner ? 'owner' : 'emp',
        empId:    row[0],
        name:     row[1],
        nickname: row[4] || row[1],
        position: row[5],
        type:     row[6],
      };
    }
  }
  return { ok: false, error: 'Username หรือ Password ไม่ถูกต้อง' };
}

// ════════════════════════════════════════════════════════════
// EMPLOYEES — พนักงาน
// ════════════════════════════════════════════════════════════
function getEmployees() {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.EMPLOYEES);
  const data = sh.getDataRange().getValues();
  const employees = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    employees.push({
      id:          row[0],
      name:        row[1],
      username:    row[2],
      // password intentionally NOT returned to clients (security)
      nickname:    row[4] || row[1],
      position:    row[5],
      type:        row[6],
      salary:      row[7],
      bank:        row[8],
      startDate:   row[9] ? fmtMaybe(row[9]) : '',
      status:      String(row[10] || '').trim().toLowerCase(),
      probationEnd:row[11] ? fmtMaybe(row[11]) : '',
      phone:       row[12] || '',
      idcard:      row[13] || '',
      address:     row[14] || '',
      dob:         fmtMaybe(row[15]),
      bankAcc:     row[16] || '',
      bankAccName: row[17] || '',
      note:        row[18] || '',
      otrate:      row[19] || '',
      pdpaConsent: row[20] || '',
      kitchen:     row[21] || '',
      ptRate:      row[22] || '',   // Part-time บาท/ชม. วันธรรมดา
      ptRateHol:   row[23] || '',   // Part-time บาท/ชม. วันหยุด/พิเศษ
    });
  }
  return { ok: true, data: employees };
}

function registerEmployee(p) {
  const ss  = openSS();
  const sh  = ss.getSheetByName(SH.EMPLOYEES);
  const data = sh.getDataRange().getValues();
  // Prevent duplicate registrations — reject if this username already exists (any status)
  const uname = String(p.username || '').trim().toLowerCase();
  if (uname) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2] || '').trim().toLowerCase() === uname) {
        return { ok: false, error: 'Username นี้ถูกใช้ไปแล้ว กรุณาเลือกชื่ออื่น' };
      }
    }
  }
  // รหัสไม่ซ้ำ = EMPnnn สูงสุดที่มี + 1 (เดิมใช้ lastRow → ชนกันเมื่อเคยลบแถวพนักงาน)
  let maxN = 0;
  for (let i = 1; i < data.length; i++) { const mm = String(data[i][0] || '').match(/^EMP(\d+)$/); if (mm) maxN = Math.max(maxN, parseInt(mm[1], 10)); }
  const newId = 'EMP' + String(maxN + 1).padStart(3, '0');
  sh.appendRow([
    newId,
    p.name        || '',
    p.username    || '',
    p.password ? hashPw(p.password) : '',   // store password hashed, never plaintext
    p.nickname    || p.name || '',
    p.position    || '',
    p.type        || 'Full-time',
    p.salary      || 0,
    p.bank        || '',          // [8] bank brand (e.g. กสิกรไทย (KBank))
    p.startDate   || '',
    'Pending',
    p.probationEnd|| '',
    p.phone       || '',          // [12]
    p.idcard      || '',          // [13]
    p.address     || '',          // [14]
    p.dob         || '',          // [15]
    p.bankAcc     || '',          // [16] account number
    p.bankAccName || '',          // [17] account holder name
    p.note        || '',          // [18]
    p.otrate      || '',          // [19]
    p.pdpaConsent ? (p.pdpaConsent + ' | ' + fmtDateTime(new Date())) : '',  // [20] PDPA consent version + timestamp
  ]);
  return { ok: true, empId: newId };
}

function approveEmployee(p) {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.EMPLOYEES);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === p.empId) {
      const r = i + 1;
      if (p.approve) {
        // Save the account settings the Owner entered on the approval screen
        if (p.type   !== undefined && p.type   !== '') sh.getRange(r, 7).setValue(p.type);    // [6] type
        if (p.salary !== undefined && p.salary !== '') sh.getRange(r, 8).setValue(p.salary);  // [7] salary
        if (p.ptRate    !== undefined && p.ptRate    !== '') sh.getRange(r, 23).setValue(p.ptRate);    // [22] PT ฿/ชม ธรรมดา
        if (p.ptRateHol !== undefined && p.ptRateHol !== '') sh.getRange(r, 24).setValue(p.ptRateHol); // [23] PT ฿/ชม วันหยุด
        if (p.note   !== undefined && p.note   !== '') sh.getRange(r, 19).setValue(p.note);   // [18] note
        sh.getRange(r, 11).setValue('Active');   // [10] status
      } else {
        sh.getRange(r, 11).setValue('Rejected');
      }
      return { ok: true };
    }
  }
  return { ok: false, error: 'Employee not found' };
}

function setEmployeeStatus(empId, status) {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.EMPLOYEES);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === empId) {
      sh.getRange(i + 1, 11).setValue(status);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Employee not found' };
}

function updateEmployee(p) {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.EMPLOYEES);
  const data = sh.getDataRange().getValues();
  // Accept both naming conventions used by the two frontend forms (register vs. edit)
  const address   = (p.address   !== undefined) ? p.address   : p.addr;
  const startDate = (p.startDate !== undefined) ? p.startDate : p.startdate;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === p.empId) {
      const r = i + 1;
      if (p.name        !== undefined) sh.getRange(r, 2).setValue(p.name);
      if (p.nickname    !== undefined) sh.getRange(r, 5).setValue(p.nickname);
      if (p.position    !== undefined) sh.getRange(r, 6).setValue(p.position);
      if (p.type        !== undefined) sh.getRange(r, 7).setValue(p.type);
      if (p.salary      !== undefined) sh.getRange(r, 8).setValue(p.salary);
      if (p.bank        !== undefined) sh.getRange(r, 9).setValue(p.bank);
      if (startDate     !== undefined) sh.getRange(r, 10).setValue(startDate);
      if (p.status      !== undefined) sh.getRange(r, 11).setValue(p.status);
      if (p.probationEnd!== undefined) sh.getRange(r, 12).setValue(p.probationEnd);
      if (p.phone       !== undefined) sh.getRange(r, 13).setValue(p.phone);
      if (p.idcard      !== undefined) sh.getRange(r, 14).setValue(p.idcard);
      if (address       !== undefined) sh.getRange(r, 15).setValue(address);
      if (p.dob         !== undefined) sh.getRange(r, 16).setValue(p.dob);
      if (p.bankAcc     !== undefined) sh.getRange(r, 17).setValue(p.bankAcc);
      if (p.bankAccName !== undefined) sh.getRange(r, 18).setValue(p.bankAccName);
      if (p.note        !== undefined) sh.getRange(r, 19).setValue(p.note);
      if (p.otrate      !== undefined) sh.getRange(r, 20).setValue(p.otrate);
      if (p.kitchen     !== undefined) sh.getRange(r, 22).setValue(p.kitchen);
      if (p.ptRate      !== undefined) sh.getRange(r, 23).setValue(p.ptRate);     // [22] PT บาท/ชม. ธรรมดา
      if (p.ptRateHol   !== undefined) sh.getRange(r, 24).setValue(p.ptRateHol);  // [23] PT บาท/ชม. วันหยุด/พิเศษ
      return { ok: true };
    }
  }
  return { ok: false, error: 'Employee not found' };
}

// ════════════════════════════════════════════════════════════
// CLOCK IN / OUT — บันทึกเวลา
// ════════════════════════════════════════════════════════════
// Attendance columns:
// [0]id [1]empId [2]date [3]clockIn [4]clockOut
// [5]plannedStart [6]plannedEnd [7]specialDay

function getTodayClock(empId, date) {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.ATTENDANCE);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row     = data[i];
    const rowDate = row[2] ? fmtDate(new Date(row[2])) : '';
    if (String(row[1]) === String(empId) && rowDate === date) {
      return {
        ok: true,
        data: {
          id:           row[0],
          clockIn:      fmtTime(row[3]),
          clockOut:     fmtTime(row[4]),
          plannedStart: fmtTime(row[5]),
          plannedEnd:   fmtTime(row[6]),
          specialDay:   row[7],
        },
      };
    }
  }
  return { ok: true, data: null };
}

function getTodayAttendance(date) {
  const ss      = openSS();
  const attSh   = ss.getSheetByName(SH.ATTENDANCE);
  const empSh   = ss.getSheetByName(SH.EMPLOYEES);
  const schSh   = ss.getSheetByName(SH.SCHEDULES);
  const attData = attSh.getDataRange().getValues();
  const empData = empSh.getDataRange().getValues();
  const schData = schSh.getDataRange().getValues();

  // Build empId → {name, position} map
  const empMap = {};
  for (let i = 1; i < empData.length; i++) {
    const id = String(empData[i][0]);
    empMap[id] = { name: empData[i][1], position: empData[i][6] };
  }

  // Build empId → planStart/planEnd from APPROVED schedule that includes this date
  const planMap = {};
  for (let i = 1; i < schData.length; i++) {
    const r = schData[i];
    if (!r[0]) continue;
    if (String(r[5]).toLowerCase() !== 'approved') continue;
    const dates = r[4] ? JSON.parse(r[4]) : {};
    if (!dates[date]) continue;
    const empId = String(r[1]);
    // ใช้ตารางที่ approve ล่าสุดถ้ามีหลายตัว (override)
    planMap[empId] = {
      start: dates[date].start || '',
      end:   dates[date].end   || '',
    };
  }

  // Filter attendance rows for today
  const seen = {};
  const rows = [];
  for (let i = 1; i < attData.length; i++) {
    const row     = attData[i];
    const rowDate = row[2] ? fmtDate(new Date(row[2])) : '';
    if (rowDate !== date) continue;
    const empId     = String(row[1]);
    const clockIn   = fmtTime(row[3]);
    // Fallback: ถ้า Attendance ไม่มี plannedStart/End ให้ใช้จาก Schedules
    const plannedStart = fmtTime(row[5]) || (planMap[empId]?.start || '');
    const plannedEnd   = fmtTime(row[6]) || (planMap[empId]?.end   || '');
    let late = false;
    if (clockIn && plannedStart) {
      const [ph, pm] = plannedStart.split(':').map(Number);
      const [ch, cm] = clockIn.split(':').map(Number);
      late = (ch * 60 + cm) > (ph * 60 + pm + 15);
    }
    seen[empId] = true;
    rows.push({
      empId,
      name:         empMap[empId]?.name     || empId,
      position:     empMap[empId]?.position || '',
      clockIn,
      clockOut:     fmtTime(row[4]),
      plannedStart,
      plannedEnd,
      late,
    });
  }

  // เพิ่มพนักงานที่มีตาราง approve วันนี้แต่ยังไม่ได้ Clock In (เพื่อ Owner เห็นทุกคน + plan time)
  Object.keys(planMap).forEach(function(empId) {
    if (seen[empId]) return;
    rows.push({
      empId,
      name:         empMap[empId]?.name     || empId,
      position:     empMap[empId]?.position || '',
      clockIn:      '',
      clockOut:     '',
      plannedStart: planMap[empId].start,
      plannedEnd:   planMap[empId].end,
      late:         false,
    });
  });

  return { ok: true, data: rows };
}

function clockIn(p) {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.ATTENDANCE);
  const data = sh.getDataRange().getValues();
  const rain = (p.rain === true || p.rain === 'true' || p.rain === 'TRUE');   // [8] ฝนตก → grace สาย 15 นาที (ปกติ 3)
  // Update existing row if found
  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][2] ? fmtDate(new Date(data[i][2])) : '';
    if (String(data[i][1]) === String(p.empId) && rowDate === p.date) {
      const c = sh.getRange(i + 1, 4);
      c.setNumberFormat('@');            // store as TEXT so "08:25" isn't auto-converted to a time value
      c.setValue(p.time);
      sh.getRange(i + 1, 9).setValue(rain);
      return { ok: true };
    }
  }
  // Create new row — write the time columns as TEXT to avoid Sheets time auto-conversion
  const newRow = sh.getLastRow() + 1;
  const newId  = 'ATT' + Date.now();
  sh.appendRow([ newId, p.empId, p.date, '', '', '', '', p.specialDay || false, rain ]);
  sh.getRange(newRow, 4, 1, 4).setNumberFormat('@');   // cols D-G: clockIn, clockOut, plannedStart, plannedEnd
  sh.getRange(newRow, 4).setValue(p.time);
  sh.getRange(newRow, 6).setValue(p.plannedStart || '');
  sh.getRange(newRow, 7).setValue(p.plannedEnd   || '');
  return { ok: true, id: newId };
}

function clockOut(p) {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.ATTENDANCE);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][2] ? fmtDate(new Date(data[i][2])) : '';
    if (String(data[i][1]) === String(p.empId) && rowDate === p.date) {
      const c = sh.getRange(i + 1, 5);
      c.setNumberFormat('@');            // store clock-out as TEXT too
      c.setValue(p.time);
      return { ok: true };
    }
  }
  return { ok: false, error: 'ไม่พบบันทึก Clock In สำหรับวันนี้' };
}

// ════════════════════════════════════════════════════════════
// SCHEDULES — ตารางงาน
// ════════════════════════════════════════════════════════════
// Columns: [0]id [1]empId [2]month [3]year [4]datesJSON
//          [5]status [6]submittedAt [7]approvedAt [8]approvedBy

function getSchedules(empId) {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.SCHEDULES);
  const data = sh.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (empId && String(row[1]) !== String(empId)) continue;
    result.push({
      id:                 row[0],
      empId:              row[1],
      month:              row[2],
      year:               row[3],
      dates:              row[4] ? JSON.parse(row[4]) : {},
      status:             row[5],
      submittedAt:        row[6],
      approvedAt:         row[7],
      approvedBy:         row[8],
      isEdit:             row[9] === 'TRUE' || row[9] === true,
      originalScheduleId: row[10] || '',
    });
  }
  return { ok: true, data: result };
}

// ── ตารางทีมรายวัน (owner): ใครเข้างานวันที่ X · กี่โมง · สถานะตอกบัตร ──
// รวมตารางที่ "อนุมัติแล้ว" ของทุกคน + สถานะตอกบัตรของวันนั้น ในการเรียกครั้งเดียว
function getDayRoster(date) {
  if (!date) return { ok: false, error: 'missing date' };
  const ss = openSS();
  const toMinR = function (t) { if (!t) return 0; const a = String(t).split(':').map(Number); return (a[0] || 0) * 60 + (a[1] || 0); };
  // employees
  const emps = ss.getSheetByName(SH.EMPLOYEES).getDataRange().getValues();
  const empMap = {};
  for (let i = 1; i < emps.length; i++) {
    const r = emps[i]; if (!r[0]) continue;
    empMap[String(r[0])] = { empId: String(r[0]), name: (r[4] || r[1] || r[0]), position: r[5] || '', status: String(r[10] || '').trim().toLowerCase() };
  }
  // approved schedule entries — both the exact-date plan AND each employee's
  // work-day range this month, so we can surface implicit days off
  // (วันหยุดที่เก็บเป็น "วันที่ไม่มี key" ไม่ใช่ {off:true} — ข้อมูลแบบเก่า)
  const monthPrefix = String(date).slice(0, 7) + '-';   // 'YYYY-MM-'
  const sch = ss.getSheetByName(SH.SCHEDULES).getDataRange().getValues();
  const plan = {}, wkRange = {};
  for (let i = 1; i < sch.length; i++) {
    const r = sch[i]; if (!r[0]) continue;
    if (String(r[5]).toLowerCase() !== 'approved') continue;
    let dd = {}; try { dd = r[4] ? JSON.parse(r[4]) : {}; } catch (e) { dd = {}; }
    const id = String(r[1]);
    if (dd[date]) plan[id] = dd[date];
    Object.keys(dd).forEach(function (k) {
      if (k.indexOf(monthPrefix) !== 0) return;   // เฉพาะเดือนที่กำลังดู
      if (dd[k] && dd[k].off) return;             // วันหยุดไม่ขยายช่วงวันทำงาน
      const cur = wkRange[id] || { min: k, max: k };
      if (k < cur.min) cur.min = k;
      if (k > cur.max) cur.max = k;
      wkRange[id] = cur;
    });
  }
  // attendance for the date
  const att = ss.getSheetByName(SH.ATTENDANCE).getDataRange().getValues();
  const dayAtt = {};
  for (let i = 1; i < att.length; i++) {
    const r = att[i]; if (!r[0]) continue;
    const rd = r[2] ? fmtDate(new Date(r[2])) : '';
    if (rd === date) dayAtt[String(r[1])] = { clockIn: fmtTime(r[3]), clockOut: fmtTime(r[4]), rain: (r[8] === true || r[8] === 'TRUE' || r[8] === 'true') };
  }
  const working = [], off = [];
  Object.keys(empMap).forEach(function (id) {
    const e = empMap[id]; if (e.status === 'inactive') return;
    const pl = plan[id];
    if (!pl || pl.off) {
      // ไม่มีกะวันนั้น → ถือเป็น "หยุด" ถ้าระบุหยุดชัด (pl.off) หรือ
      // มีตารางอนุมัติเดือนนี้และวันนี้อยู่ในช่วงวันที่ลงตารางไว้ (วันหยุดแบบช่องว่าง)
      const rg = wkRange[id];
      if ((pl && pl.off) || (rg && date >= rg.min && date <= rg.max)) {
        off.push({ empId: id, name: e.name, position: e.position });
      }
      return;                                          // นอกช่วงตาราง/ไม่มีตาราง = ไม่แสดง
    }
    const a = dayAtt[id] || {};
    const start = pl.start || '', end = pl.end || '';
    let status = 'planned', late = false;
    if (a.clockIn) { late = !!(start && toMinR(a.clockIn) > toMinR(start) + (a.rain ? 15 : 3)); status = late ? 'late' : 'in'; }
    const otHrs = (start && end) ? Math.max(0, (toMinR(end) - toMinR(start)) / 60 - 9) : 0;
    working.push({ empId: id, name: e.name, position: e.position, start: start, end: end,
      specialDay: !!pl.specialDay, otHrs: Math.round(otHrs * 10) / 10,
      clockIn: a.clockIn || '', clockOut: a.clockOut || '', status: status, late: late });
  });
  working.sort(function (a, b) { return (String(a.start) + a.end).localeCompare(String(b.start) + b.end); });
  return { ok: true, date: date, working: working, off: off,
    counts: { working: working.length, off: off.length, late: working.filter(function (w) { return w.late; }).length } };
}

// ════════════════════════════════════════════════════════════
// SIDEWORK — งานประจำร้าน (บอร์ดทีมต่อครัว · แต้ม · leaderboard)
// Sidework_Tasks: [0]id [1]kitchen [2]name [3]category [4]tier(daily/weekly) [5]points [6]busyInclude [7]active [8]createdAt
// Sidework_Log:   [0]id [1]ts [2]cycleType [3]cycleKey [4]taskId [5]taskName [6]kitchen [7]empId [8]empName [9]points
// reset: daily 19:30 (เริ่ม 06:00) · weekly เสาร์ 19:30 (เริ่มอาทิตย์ 06:00) — คำนวณ logical จากเวลาปัจจุบัน
// ════════════════════════════════════════════════════════════
const SW = { TASKS:'Sidework_Tasks', LOG:'Sidework_Log' };
const SW_OPEN=360, SW_CLOSE=1170;  // 06:00, 19:30 (นาทีของวัน)
const SW_KITCHENS=['บาร์น้ำ','เบเกอรี่','ครัวอาหาร','ส่วนกลาง'];
const SW_HOLIDAYS={'2026-01-01':1,'2026-02-06':1,'2026-04-06':1,'2026-04-13':1,'2026-04-14':1,'2026-04-15':1,'2026-05-01':1,'2026-05-05':1,'2026-05-11':1,'2026-06-03':1,'2026-07-28':1,'2026-08-12':1,'2026-10-13':1,'2026-10-23':1,'2026-12-05':1,'2026-12-10':1,'2026-12-31':1};

function swEnsure_(){
  const ss=openSS();
  let t=ss.getSheetByName(SW.TASKS);
  if(!t){ t=ss.insertSheet(SW.TASKS); t.appendRow(['id','kitchen','name','category','tier','points','busyInclude','active','createdAt','days','deadline','assignee','multi']); t.setFrozenRows(1); }
  else if(t.getLastColumn()<13){ t.getRange(1,10,1,4).setValues([['days','deadline','assignee','multi']]); }  // upgrade schema: 1.2 days · 1.3 once/deadline/assignee · 1.4 multi
  if(!ss.getSheetByName(SW.LOG)){ const l=ss.insertSheet(SW.LOG); l.appendRow(['id','ts','cycleType','cycleKey','taskId','taskName','kitchen','empId','empName','points']); l.setFrozenRows(1); }
  return ss;
}
// 1.2 — daily task มีผลเฉพาะวันที่เลือก (days = CSV ของ dow 0=อา..6=ส · '*'/'' = ทุกวัน)
function swDayMatch_(days,dow){ if(!days||days==='*') return true; return (','+String(days)+',').indexOf(','+dow+',')>=0; }
// name lookup empId -> ชื่อเล่น/ชื่อ
function swNameMap_(){ const d=openSS().getSheetByName(SH.EMPLOYEES).getDataRange().getValues(), m={}; for(let i=1;i<d.length;i++){ if(!d[i][0]) continue; m[String(d[i][0])]=(d[i][4]||d[i][1]||d[i][0]); } return m; }
// active staff (เว้น owner) สำหรับเลือกเพื่อนช่วยงาน (1.4)
function swMates_(){ const d=openSS().getSheetByName(SH.EMPLOYEES).getDataRange().getValues(), out=[];
  for(let i=1;i<d.length;i++){ const r=d[i]; if(!r[0]) continue; if(String(r[10]||'').trim().toLowerCase()==='inactive') continue;
    if(String(r[5]||'').toLowerCase().indexOf('owner')>=0) continue;
    out.push({empId:String(r[0]), name:(r[4]||r[1]||r[0])}); }
  return out; }
function swAddDays_(dateStr,delta){ const p=dateStr.split('-').map(Number); const d=new Date(p[0],p[1]-1,p[2]); d.setDate(d.getDate()+delta); const z=n=>('0'+n).slice(-2); return d.getFullYear()+'-'+z(d.getMonth()+1)+'-'+z(d.getDate()); }
function swCycle_(){
  const n=new Date();
  const dateStr=Utilities.formatDate(n,'Asia/Bangkok','yyyy-MM-dd');
  const mins=parseInt(Utilities.formatDate(n,'Asia/Bangkok','H'),10)*60+parseInt(Utilities.formatDate(n,'Asia/Bangkok','m'),10);
  const dow=parseInt(Utilities.formatDate(n,'Asia/Bangkok','u'),10)%7;  // Mon=1..Sun=7 → Sun=0
  const dailyActive=mins>=SW_OPEN && mins<SW_CLOSE;
  const weeklyClosed=(dow===6 && mins>=SW_CLOSE) || (dow===0 && mins<SW_OPEN);
  return { dateStr:dateStr, mins:mins, dow:dow, dailyKey:dateStr, dailyActive:dailyActive,
    weeklyKey:swAddDays_(dateStr,-dow), weeklyActive:!weeklyClosed, isBusy:(dow===0||dow===6)||!!SW_HOLIDAYS[dateStr] };
}
function swEmpKitchen_(empId){
  const d=openSS().getSheetByName(SH.EMPLOYEES).getDataRange().getValues();
  for(let i=1;i<d.length;i++){ if(String(d[i][0])===String(empId)) return { kitchen:String(d[i][21]||''), name:(d[i][4]||d[i][1]||d[i][0]) }; }
  return { kitchen:'', name:empId };
}
function swReadTasks_(){
  const sh=swEnsure_().getSheetByName(SW.TASKS), last=sh.getLastRow(); if(last<2) return [];
  const r=sh.getRange(2,1,last-1,13).getValues(), out=[];
  for(let i=0;i<r.length;i++){ const x=r[i]; if(!x[0]) continue;
    out.push({id:String(x[0]),kitchen:String(x[1]),name:String(x[2]),category:String(x[3]),tier:String(x[4]),points:Number(x[5])||0,
      busyInclude:(x[6]===true||x[6]==='TRUE'||x[6]==='true'), active:!(x[7]===false||x[7]==='FALSE'||x[7]==='false'),
      days:(String(x[9]||'').trim()||'*'), deadline:String(x[10]||''), assignee:String(x[11]||''), multi:(x[12]===true||x[12]==='TRUE'||x[12]==='true'), _row:i+2}); }
  return out;
}
function swCurrentDone_(){
  const cyc=swCycle_(), sh=openSS().getSheetByName(SW.LOG), last=sh.getLastRow(), done={};
  if(last>=2){ const r=sh.getRange(2,1,last-1,10).getValues();
    for(let i=0;i<r.length;i++){ const x=r[i]; if(!x[4]) continue; const ct=String(x[2]), ck=String(x[3]), tid=String(x[4]);
      const match=(ct==='daily'&&ck===cyc.dailyKey)||(ct==='weekly'&&ck===cyc.weeklyKey)||(ct==='once');  // once = ทำครั้งเดียวถาวร
      if(match){ if(!done[tid]) done[tid]={empName:String(x[8]),empId:String(x[7]),names:[],ids:[]}; done[tid].names.push(String(x[8])); done[tid].ids.push(String(x[7])); } } }
  return done;
}
function swLeaderboard_(){
  const sh=swEnsure_().getSheetByName(SW.LOG), last=sh.getLastRow();
  const month=Utilities.formatDate(new Date(),'Asia/Bangkok','yyyy-MM'); const all={}, mon={};
  if(last>=2){ const r=sh.getRange(2,1,last-1,10).getValues();
    for(let i=0;i<r.length;i++){ const x=r[i]; if(!x[7]) continue; const id=String(x[7]), nm=String(x[8]||id), pt=Number(x[9])||0;
      all[id]=all[id]||{empId:id,name:nm,points:0}; all[id].points+=pt; all[id].name=nm;
      if(String(x[1]).slice(0,7)===month){ mon[id]=mon[id]||{empId:id,name:nm,points:0}; mon[id].points+=pt; mon[id].name=nm; } } }
  const rank=function(o){ const a=Object.keys(o).map(k=>o[k]).sort((x,y)=>y.points-x.points); a.forEach((e,i)=>e.rank=i+1); return a; };
  return { allTime:rank(all), month:rank(mon) };
}

function getSideworkBoard(empId){
  swEnsure_(); const cyc=swCycle_(), me=swEmpKitchen_(empId), done=swCurrentDone_();
  const nowStr=Utilities.formatDate(new Date(),'Asia/Bangkok','yyyy-MM-dd HH:mm');
  // มองเห็น: งานที่มอบหมายให้ฉัน, หรือ (ไม่ได้มอบหมาย และอยู่ครัวฉัน/ส่วนกลาง)
  const all=swReadTasks_().filter(function(t){ if(!t.active) return false;
    if(t.assignee) return String(t.assignee)===String(empId);
    return t.kitchen===me.kitchen || t.kitchen==='ส่วนกลาง'; });
  const canDo=function(t){
    if(t.tier==='daily') return cyc.dailyActive && swDayMatch_(t.days,cyc.dow);
    if(t.tier==='weekly') return cyc.weeklyActive;
    if(t.tier==='once')  return !t.deadline || nowStr<=t.deadline;
    return false; };
  const visible=function(t){ const dn=done[t.id];
    if(t.tier==='daily') return swDayMatch_(t.days,cyc.dow);
    if(t.tier==='once')  return !dn && (!t.deadline || nowStr<=t.deadline);                   // เสร็จ/เลยเวลา = หายไป
    return true; };                                                                            // weekly โชว์เสมอ
  const pack=function(t){ const dn=done[t.id]; return {id:t.id,name:t.name,category:t.category,tier:t.tier,points:t.points,kitchen:t.kitchen,
    done:!!dn, doneByName:dn?dn.names.join(', '):'', mine:dn?dn.ids.indexOf(String(empId))>=0:false,
    multi:t.multi, assignee:t.assignee, deadline:t.deadline, days:t.days, busyInclude:t.busyInclude, active:canDo(t) }; };
  const fil=function(scope){ return all.filter(scope).filter(visible).map(pack); };
  const lb=swLeaderboard_(); const mA=lb.allTime.filter(x=>String(x.empId)===String(empId))[0], mM=lb.month.filter(x=>String(x.empId)===String(empId))[0];
  return { ok:true, data:{ kitchen:me.kitchen, empName:me.name, isBusy:cyc.isBusy, dailyActive:cyc.dailyActive, weeklyActive:cyc.weeklyActive,
    personal: fil(function(t){return !!t.assignee;}),
    own:      fil(function(t){return !t.assignee && t.kitchen===me.kitchen;}),
    central:  fil(function(t){return !t.assignee && t.kitchen==='ส่วนกลาง';}),
    mates: swMates_(),
    myAllTime:mA?mA.points:0, myAllRank:mA?mA.rank:'-', myMonth:mM?mM.points:0, myMonthRank:mM?mM.rank:'-' } };
}
function submitSidework(p){
  swEnsure_(); const cyc=swCycle_(), me=swEmpKitchen_(p.empId);
  const task=swReadTasks_().filter(t=>t.id===String(p.taskId))[0];
  if(!task||!task.active) return { ok:false, error:'ไม่พบงานนี้' };
  if(task.assignee && String(task.assignee)!==String(p.empId)) return { ok:false, error:'งานนี้มอบหมายเฉพาะผู้รับผิดชอบ' };
  const nowStr=Utilities.formatDate(new Date(),'Asia/Bangkok','yyyy-MM-dd HH:mm');
  if(task.tier==='daily'){ if(!cyc.dailyActive) return { ok:false, error:'รอบวันปิดแล้ว เริ่มใหม่ 06:00' };
    if(!swDayMatch_(task.days,cyc.dow)) return { ok:false, error:'งานนี้ไม่ใช่ของวันนี้' }; }
  else if(task.tier==='weekly'){ if(!cyc.weeklyActive) return { ok:false, error:'รอบสัปดาห์ปิดแล้ว เริ่มใหม่อาทิตย์ 06:00' }; }
  else if(task.tier==='once'){ if(task.deadline && nowStr>task.deadline) return { ok:false, error:'เลยกำหนดเวลาแล้ว' }; }
  const ck=task.tier==='daily'?cyc.dailyKey:(task.tier==='weekly'?cyc.weeklyKey:task.id);
  const sh=openSS().getSheetByName(SW.LOG), last=sh.getLastRow();
  if(last>=2){ const r=sh.getRange(2,1,last-1,10).getValues();
    for(let i=0;i<r.length;i++){ const x=r[i]; if(String(x[4])===task.id && String(x[2])===task.tier && String(x[3])===ck) return { ok:false, error:(String(x[8])||'มีคน')+' ทำไปแล้ว' }; } }
  // 1.4 — งานช่วยกัน (multi & ไม่ระบุผู้รับผิดชอบ): คนแรกเลือกเพื่อน → หารแต้มเท่ากัน, log 1 แถวต่อคน
  let helpers=[String(p.empId)];
  if(task.multi && !task.assignee && p.helpers){ try{ const hs=JSON.parse(p.helpers); if(Array.isArray(hs)) helpers=hs.map(String).filter(function(v){return v;}); }catch(e){} }
  if(helpers.indexOf(String(p.empId))<0) helpers.unshift(String(p.empId));
  helpers=helpers.filter(function(v,i){return helpers.indexOf(v)===i;});
  const share=Math.round((task.points/helpers.length)*10)/10, nm=swNameMap_(), ts=fmtDateTime(new Date());
  helpers.forEach(function(hid){ sh.appendRow(['SWL'+Date.now()+'_'+hid, ts, task.tier, ck, task.id, task.name, task.kitchen, hid, (nm[hid]||hid), share]); });
  return { ok:true, points:task.points, share:share, n:helpers.length, doneByName:me.name };
}
function getSideworkTasks(kitchen){
  swEnsure_(); const done=swCurrentDone_(); let ts=swReadTasks_().filter(t=>t.active); if(kitchen) ts=ts.filter(t=>t.kitchen===kitchen);
  const nm=swNameMap_();
  return { ok:true, data: ts.map(function(t){ return {id:t.id,kitchen:t.kitchen,name:t.name,category:t.category,tier:t.tier,points:t.points,busyInclude:t.busyInclude,
    days:t.days,deadline:t.deadline,assignee:t.assignee,assigneeName:(t.assignee?(nm[t.assignee]||t.assignee):''),multi:t.multi,doneThisCycle:!!done[t.id]}; }) };
}
function saveSideworkTask(p){
  const sh=swEnsure_().getSheetByName(SW.TASKS);
  if(!p.name||!p.kitchen) return { ok:false, error:'กรอกชื่อ+ครัว' };
  const tier=(p.tier==='weekly')?'weekly':((p.tier==='once')?'once':'daily');
  const points=Number(p.points)||1, busy=(p.busyInclude===true||p.busyInclude==='true'||p.busyInclude==='TRUE');
  const days=(tier==='daily')?(String(p.days||'*').trim()||'*'):'';
  const deadline=(tier==='once')?String(p.deadline||''):'';
  const assignee=String(p.assignee||'');
  const multi=(!assignee)&&(p.multi===true||p.multi==='true'||p.multi==='TRUE');  // ระบุผู้รับผิดชอบแล้ว = ทำคนเดียว
  if(p.id){ const last=sh.getLastRow(); if(last>=2){ const ids=sh.getRange(2,1,last-1,1).getValues();
    for(let i=0;i<ids.length;i++){ if(String(ids[i][0])===String(p.id)){
      sh.getRange(i+2,2,1,7).setValues([[p.kitchen,p.name,p.category||'',tier,points,busy,true]]);
      sh.getRange(i+2,10,1,4).setValues([[days,deadline,assignee,multi]]);  // col 9 createdAt คงเดิม
      return { ok:true, id:p.id }; } } } }
  const id='SWT'+Date.now();
  sh.appendRow([id,p.kitchen,p.name,p.category||'',tier,points,busy,true,fmtDateTime(new Date()),days,deadline,assignee,multi]);
  return { ok:true, id:id };
}
function deleteSideworkTask(p){
  const ss=swEnsure_(), sh=ss.getSheetByName(SW.TASKS);
  const task=swReadTasks_().filter(t=>t.id===String(p.id))[0];
  if(!task) return { ok:false, error:'ไม่พบงาน' };
  const cyc=swCycle_(), ck=task.tier==='daily'?cyc.dailyKey:(task.tier==='weekly'?cyc.weeklyKey:task.id);
  const log=ss.getSheetByName(SW.LOG), last=log.getLastRow();
  if(last>=2){ const r=log.getRange(2,1,last-1,10).getValues();
    for(let i=0;i<r.length;i++){ const x=r[i]; if(String(x[4])===task.id && String(x[2])===task.tier && String(x[3])===ck) return { ok:false, blocked:true, error:'มีคนทำรอบนี้แล้ว — ลบได้หลังรีเซ็ต' }; } }
  sh.deleteRow(task._row);
  return { ok:true };
}
function getSideworkLeaderboard(){ return { ok:true, data:swLeaderboard_() }; }
function getSideworkDashboard(){
  swEnsure_(); const cyc=swCycle_(), tasks=swReadTasks_().filter(t=>t.active), done=swCurrentDone_(), kits={};
  SW_KITCHENS.forEach(function(k){ kits[k]={kitchen:k,coreTotal:0,coreDone:0}; });
  tasks.forEach(function(t){ if(t.tier!=='daily') return; if(t.assignee) return; if(!swDayMatch_(t.days,cyc.dow)) return; const k=kits[t.kitchen]; if(!k) return; k.coreTotal++; if(done[t.id]) k.coreDone++; });
  return { ok:true, data:{ isBusy:cyc.isBusy, dailyActive:cyc.dailyActive, kitchens:SW_KITCHENS.map(k=>kits[k]), leaderboard:swLeaderboard_() } };
}

// ════════════════════════════════════════════════════════════
// PREP STOCK — ของเตรียมขาย (Phase A)
// ════════════════════════════════════════════════════════════
const PREP = { ITEMS:'Prep_Items', LOGS:'Prep_Logs', SHORT:'Prep_Shortages' };
function prepEnsure_(){
  const ss=openSS();
  if(!ss.getSheetByName(PREP.ITEMS)){ const s=ss.insertSheet(PREP.ITEMS); s.appendRow(['id','kitchen','name','unit','active','createdAt']); s.setFrozenRows(1); }
  if(!ss.getSheetByName(PREP.LOGS)){ const s=ss.insertSheet(PREP.LOGS); s.appendRow(['id','date','kitchen','itemId','itemName','produced','used','left','handover','ready','note','empId','empName','ts']); s.setFrozenRows(1); }
  if(!ss.getSheetByName(PREP.SHORT)){ const s=ss.insertSheet(PREP.SHORT); s.appendRow(['id','date','kitchen','material','note','empId','empName','ts']); s.setFrozenRows(1); }
  return ss;
}
function prepDStr_(v){ if(v instanceof Date){ try{ return Utilities.formatDate(v, openSS().getSpreadsheetTimeZone(), "yyyy-MM-dd"); }catch(e){ return Utilities.formatDate(v, "Asia/Bangkok", "yyyy-MM-dd"); } } return String(v).slice(0,10); }

function prepToday_(){ return Utilities.formatDate(new Date(),'Asia/Bangkok','yyyy-MM-dd'); }
function prepUid_(p){ return p+Date.now()+Math.floor(Math.random()*1000); }
function prepBool_(v){ return v===true||v==='TRUE'||v==='true'; }
function prepDelRows_(sh,date,kitchen){ const last=sh.getLastRow(); if(last<2) return; const r=sh.getRange(2,2,last-1,2).getValues(); for(let i=r.length-1;i>=0;i--){ if(prepDStr_(r[i][0])===prepDStr_(date)&&String(r[i][1])===String(kitchen)) sh.deleteRow(i+2); } }
function getPrepList(kitchen){
  prepEnsure_(); const sh=openSS().getSheetByName(PREP.ITEMS), last=sh.getLastRow(), out=[];
  if(last>=2){ const r=sh.getRange(2,1,last-1,6).getValues();
    for(let i=0;i<r.length;i++){ const x=r[i]; if(!x[0]) continue; if(x[4]===false||x[4]==='FALSE'||x[4]==='false') continue;
      if(kitchen && String(x[1])!==String(kitchen)) continue;
      out.push({id:String(x[0]),kitchen:String(x[1]),name:String(x[2]),unit:String(x[3]||'')}); } }
  return { ok:true, data:out };
}
function savePrepItem(p){
  prepEnsure_(); const sh=openSS().getSheetByName(PREP.ITEMS);
  if(p.id){ const last=sh.getLastRow(); if(last>=2){ const r=sh.getRange(2,1,last-1,1).getValues();
      for(let i=0;i<r.length;i++){ if(String(r[i][0])===String(p.id)){ sh.getRange(i+2,2,1,3).setValues([[p.kitchen||'',p.name||'',p.unit||'']]); return {ok:true,id:p.id}; } } }
    return {ok:false,error:'ไม่พบรายการ'}; }
  if(!p.name||!p.kitchen) return {ok:false,error:'ใส่ชื่อ+ครัว'};
  const id=prepUid_('PI'); sh.appendRow([id,p.kitchen,p.name,p.unit||'',true,fmtDateTime(new Date())]); return {ok:true,id:id};
}
function deletePrepItem(p){
  prepEnsure_(); const sh=openSS().getSheetByName(PREP.ITEMS), last=sh.getLastRow(); if(last<2) return {ok:false,error:'ว่าง'};
  const r=sh.getRange(2,1,last-1,1).getValues();
  for(let i=0;i<r.length;i++){ if(String(r[i][0])===String(p.id)){ sh.getRange(i+2,5).setValue(false); return {ok:true}; } }
  return {ok:false,error:'ไม่พบ'};
}
function getPrepLog(p){
  prepEnsure_(); const date=prepToday_(), kitchen=p.kitchen, ss=openSS();
  const lg=ss.getSheetByName(PREP.LOGS), last=lg.getLastRow(), items={};
  if(last>=2){ const r=lg.getRange(2,1,last-1,14).getValues();
    for(let i=0;i<r.length;i++){ const x=r[i]; if(prepDStr_(x[1])!==date||String(x[2])!==String(kitchen)) continue;
      items[String(x[3])]={produced:Number(x[5])||0,used:Number(x[6])||0,left:Number(x[7])||0,handover:prepBool_(x[8]),ready:!(x[9]===false||x[9]==='FALSE'||x[9]==='false'),note:String(x[10]||'')}; } }
  const sh=ss.getSheetByName(PREP.SHORT), sl=sh.getLastRow(), shorts=[];
  if(sl>=2){ const r=sh.getRange(2,1,sl-1,8).getValues(); for(let i=0;i<r.length;i++){ const x=r[i]; if(prepDStr_(x[1])!==date||String(x[2])!==String(kitchen)) continue; shorts.push({material:String(x[3]),note:String(x[4]||'')}); } }
  return { ok:true, data:{ logged:Object.keys(items).length>0, items:items, shortages:shorts } };
}
function savePrepLog(p){
  prepEnsure_(); const ss=openSS(), date=prepToday_(), kitchen=p.kitchen;
  if(!kitchen) return {ok:false,error:'ไม่มีครัว'};
  const me=swEmpKitchen_(p.empId), empName=me.name||p.empId;
  let items=[], shorts=[];
  try{ items=JSON.parse(p.items||'[]'); }catch(e){}
  try{ shorts=JSON.parse(p.shortages||'[]'); }catch(e){}
  const lg=ss.getSheetByName(PREP.LOGS); prepDelRows_(lg,date,kitchen);
  const now=fmtDateTime(new Date());
  items.forEach(function(it){ lg.appendRow([prepUid_('PL'),date,kitchen,it.itemId||'',it.itemName||'',Number(it.produced)||0,Number(it.used)||0,Number(it.left)||0,(it.handover?true:false),(it.ready===false?false:true),it.note||'',p.empId,empName,now]); });
  const sh=ss.getSheetByName(PREP.SHORT); prepDelRows_(sh,date,kitchen);
  shorts.forEach(function(s){ if(!s.material) return; sh.appendRow([prepUid_('PS'),date,kitchen,s.material,s.note||'',p.empId,empName,now]); });
  return { ok:true };
}
function getMorningPrep(kitchen){
  prepEnsure_(); const today=prepToday_(); const lg=openSS().getSheetByName(PREP.LOGS), last=lg.getLastRow();
  let fromDate=''; const rows=[];
  if(last>=2){ const r=lg.getRange(2,1,last-1,14).getValues();
    for(let i=0;i<r.length;i++){ const x=r[i]; if(kitchen&&String(x[2])!==String(kitchen)) continue; const d=prepDStr_(x[1]); if(d<today && d>fromDate) fromDate=d; }
    if(fromDate){ for(let i=0;i<r.length;i++){ const x=r[i]; if(kitchen&&String(x[2])!==String(kitchen)) continue; if(prepDStr_(x[1])!==fromDate) continue;
      rows.push({itemId:String(x[3]),itemName:String(x[4]),left:Number(x[7])||0,handover:prepBool_(x[8]),ready:!(x[9]===false||x[9]==='FALSE'||x[9]==='false'),note:String(x[10]||'')}); } } }
  const toPrep=rows.filter(function(it){ return it.handover || !it.ready; });
  return { ok:true, data:{ kitchen:kitchen, date:today, fromDate:fromDate, toPrep:toPrep, all:rows } };
}
function getShortages(p){
  prepEnsure_(); const date=(p&&p.date)||prepToday_(); const sh=openSS().getSheetByName(PREP.SHORT), last=sh.getLastRow(), out=[];
  if(last>=2){ const r=sh.getRange(2,1,last-1,8).getValues();
    for(let i=0;i<r.length;i++){ const x=r[i]; if(prepDStr_(x[1])!==date) continue; out.push({kitchen:String(x[2]),material:String(x[3]),note:String(x[4]||''),by:String(x[6]||'')}); } }
  return { ok:true, data:out };
}

function submitSchedule(p) {
  const ss    = openSS();
  const sh    = ss.getSheetByName(SH.SCHEDULES);
  const now   = fmtDateTime(new Date());
  const data  = sh.getDataRange().getValues();
  // Upsert — one schedule per employee per month/year. Resubmitting UPDATES the
  // existing row instead of creating a duplicate pending request.
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(p.empId) &&
        String(data[i][2]) === String(p.month) &&
        String(data[i][3]) === String(p.year)) {
      const r = i + 1;
      sh.getRange(r, 5).setValue(JSON.stringify(p.dates)); // datesJSON
      sh.getRange(r, 6).setValue('Pending');               // status
      sh.getRange(r, 7).setValue(now);                     // submittedAt
      sh.getRange(r, 8).setValue('');                      // approvedAt
      sh.getRange(r, 9).setValue('');                      // approvedBy
      return { ok: true, scheduleId: data[i][0], updated: true };
    }
  }
  const newId = 'SCH' + Date.now();
  sh.appendRow([
    newId,
    p.empId,
    p.month,
    p.year,
    JSON.stringify(p.dates),
    'Pending',
    now,
    '',
    '',
  ]);
  return { ok: true, scheduleId: newId };
}

function approveSchedule(p) {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.SCHEDULES);
  const data = sh.getDataRange().getValues();

  // Find the row being approved
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) !== String(p.scheduleId)) continue;
    const isEdit = data[i][9] === 'TRUE' || data[i][9] === true;
    const origId = String(data[i][10] || '');

    if (isEdit && origId) {
      // Edit request — merge the edited date(s) into the original approved schedule
      const editDates = data[i][4] ? JSON.parse(data[i][4]) : {};
      for (let j = 1; j < data.length; j++) {
        if (String(data[j][0]) !== origId) continue;
        const origDates = data[j][4] ? JSON.parse(data[j][4]) : {};
        Object.keys(editDates).forEach(function(dt) { origDates[dt] = editDates[dt]; });
        sh.getRange(j + 1, 5).setValue(JSON.stringify(origDates));
        break;
      }
    }
    break;
  }

  return setScheduleStatus(p.scheduleId, 'Approved', p.approvedBy);
}

function rejectSchedule(p) {
  return setScheduleStatus(p.scheduleId, 'Rejected', p.approvedBy);
}

function editSchedule(p) {
  if (!p.empId || !p.originalScheduleId || !p.dates)
    return { ok: false, error: 'ข้อมูลไม่ครบถ้วน' };

  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.SCHEDULES);
  const data = sh.getDataRange().getValues();
  const now  = fmtDateTime(new Date());
  let changes = {};
  try { changes = (typeof p.dates === 'string') ? JSON.parse(p.dates) : p.dates; } catch (e) { changes = {}; }
  if (!Object.keys(changes).length) return { ok: false, error: 'ไม่มีวันที่จะแก้' };

  // คำขอแก้ทั้งเดือน = 1 แถว edit ต่อ (พนักงาน · ตารางต้นฉบับ) — มี pending อยู่แล้ว ให้ merge ทุกวันลงแถวเดิม
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (String(row[1]) !== String(p.empId)) continue;
    if (String(row[5]) !== 'Pending') continue;
    if (row[9] !== 'TRUE' && row[9] !== true) continue;
    if (String(row[10]) !== String(p.originalScheduleId)) continue;
    const existingDates = row[4] ? JSON.parse(row[4]) : {};
    Object.keys(changes).forEach(function (k) { existingDates[k] = changes[k]; });
    sh.getRange(i + 1, 5).setValue(JSON.stringify(existingDates));
    sh.getRange(i + 1, 7).setValue(now);
    return { ok: true, scheduleId: row[0], updated: true, days: Object.keys(existingDates).length };
  }

  // สร้าง pending edit request ใหม่ (รวมทุกวันที่ขอแก้)
  const newId = 'SCH' + Date.now();
  sh.appendRow([ newId, p.empId, p.month, p.year, JSON.stringify(changes), 'Pending', now, '', '', 'TRUE', p.originalScheduleId ]);
  return { ok: true, scheduleId: newId, days: Object.keys(changes).length };
}

function setScheduleStatus(scheduleId, status, approvedBy) {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.SCHEDULES);
  const data = sh.getDataRange().getValues();
  const now  = fmtDateTime(new Date());
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === scheduleId) {
      sh.getRange(i + 1, 6).setValue(status);
      sh.getRange(i + 1, 8).setValue(now);
      sh.getRange(i + 1, 9).setValue(approvedBy || 'Owner');
      return { ok: true };
    }
  }
  return { ok: false, error: 'Schedule not found' };
}

// ════════════════════════════════════════════════════════════
// OT REQUESTS — คำขอทำงานล่วงเวลา
// ════════════════════════════════════════════════════════════
// Columns: [0]id [1]empId [2]date [3]requestedHours [4]reason
//          [5]status [6]submittedAt [7]approvedAt [8]approvedBy

function submitOT(p) {
  const ss    = openSS();
  const sh    = ss.getSheetByName(SH.OT_REQUESTS);
  const newId = 'OT' + Date.now();
  const now   = fmtDateTime(new Date());
  sh.appendRow([
    newId,
    p.empId,
    p.date,
    p.requestedHours,
    p.reason || '',
    'Pending',
    now,
    '',
    '',
    p.startTime || '',   // [9]  start HH:MM
    p.endTime   || '',   // [10] end   HH:MM
  ]);
  return { ok: true, otId: newId };
}

// Return all OT requests for a specific employee (all statuses), newest first.
function getOT(empId) {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.OT_REQUESTS);
  const data = sh.getDataRange().getValues();
  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (empId && String(row[1]) !== String(empId)) continue;
    records.push({
      id:             row[0],
      empId:          row[1],
      date:           row[2] ? fmtDate(new Date(row[2])) : '',
      requestedHours: row[3],
      reason:         row[4],
      status:         String(row[5] || '').toLowerCase(),
      submittedAt:    row[6],
      approvedAt:     row[7],
      approvedBy:     row[8],
      startTime:      fmtTime(row[9]),
      endTime:        fmtTime(row[10]),
    });
  }
  records.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; });
  return { ok: true, data: records };
}

// ประวัติ OT ของพนักงาน กรองตามเดือน/ปี (ใช้กับหน้าสถิติ OT ของพนักงาน)
function getOTHistory(empId, month, year) {
  const all = getOT(empId);
  if (!all.ok) return all;
  const mo = month ? parseInt(month) : null;
  const yr = year  ? parseInt(year)  : null;
  const filtered = all.data.filter(function(r) {
    if (!r.date) return false;
    const d = new Date(r.date);
    if (isNaN(d)) return false;
    if (mo && (d.getMonth() + 1) !== mo) return false;
    if (yr && d.getFullYear() !== yr) return false;
    return true;
  });
  return { ok: true, data: filtered };
}

function approveOT(p) {
  return setOTStatus(p.otId, p.approve ? 'Approved' : 'Rejected', p.approvedBy);
}

function setOTStatus(otId, status, approvedBy) {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.OT_REQUESTS);
  const data = sh.getDataRange().getValues();
  const now  = fmtDateTime(new Date());
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === otId) {
      sh.getRange(i + 1, 6).setValue(status);
      sh.getRange(i + 1, 8).setValue(now);
      sh.getRange(i + 1, 9).setValue(approvedBy || 'Owner');
      return { ok: true };
    }
  }
  return { ok: false, error: 'OT request not found' };
}

// ════════════════════════════════════════════════════════════
// ATTENDANCE — ดึงบันทึกเวลา (สำหรับ Export)
// ════════════════════════════════════════════════════════════
function getAttendance(empId, month, year, limit) {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.ATTENDANCE);
  const data = sh.getDataRange().getValues();

  // ── Plan + specialDay map จาก APPROVED schedules: key "empId|date" -> {start,end,special} ──
  const planMap = {};
  const schData = ss.getSheetByName(SH.SCHEDULES).getDataRange().getValues();
  for (let i = 1; i < schData.length; i++) {
    const r = schData[i];
    if (!r[0]) continue;
    if (String(r[5]).toLowerCase() !== 'approved') continue;
    const eId = String(r[1]);
    const dates = r[4] ? JSON.parse(r[4]) : {};
    Object.keys(dates).forEach(function(dt) {
      planMap[eId + '|' + dt] = {
        start:   dates[dt].start || '',
        end:     dates[dt].end   || '',
        special: dates[dt].specialDay === true,
      };
    });
  }

  // ── Approved OT map: key "empId|date" -> [{hours,start,end}] (เก็บเวลาเพื่อจำแนกก่อน/หลัง plan) ──
  const otMap = {};
  const otData = ss.getSheetByName(SH.OT_REQUESTS).getDataRange().getValues();
  for (let i = 1; i < otData.length; i++) {
    const r = otData[i];
    if (!r[0]) continue;
    if (String(r[5]).toLowerCase() !== 'approved') continue;
    const dt = r[2] ? fmtDate(new Date(r[2])) : '';
    if (!dt) continue;
    const hrs = parseFloat(r[3]) || 0;
    if (hrs <= 0) continue;                       // ข้ามคำขอ 0 ชม. (เช่น ขอออกก่อนเวลา)
    const k = String(r[1]) + '|' + dt;
    if (!otMap[k]) otMap[k] = [];
    otMap[k].push({ hours: hrs, start: fmtTime(r[9]), end: fmtTime(r[10]) });
  }

  const toMin = function(t) { const a = t.split(':').map(Number); return a[0] * 60 + a[1]; };

  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (empId && String(row[1]) !== String(empId)) continue;
    const d = row[2] ? new Date(row[2]) : null;
    if (!d) continue;
    if (month && (d.getMonth() + 1) !== parseInt(month)) continue;
    if (year  && d.getFullYear()   !== parseInt(year))  continue;

    const eId     = String(row[1]);
    const dateStr = fmtDate(d);
    const key     = eId + '|' + dateStr;
    const plan    = planMap[key] || {};

    const clockIn      = fmtTime(row[3]);
    const clockOut     = fmtTime(row[4]);
    const plannedStart = fmtTime(row[5]) || plan.start || '';
    const plannedEnd   = fmtTime(row[6]) || plan.end   || '';

    const dow       = d.getDay();
    const isSpecial = (plan.special === true) || dow === 0 || dow === 6;

    const r2 = function(x) { return Math.round(x * 100) / 100; };
    const rain       = (row[8] === true || row[8] === 'TRUE' || row[8] === 'true');
    const completed  = !!(clockIn && clockOut);
    const otReqs     = otMap[key] || [];             // คำขอ OT ที่อนุมัติของวันนั้น

    // ── สาย: clock-in เกิน "จุดเริ่มของช่วงแรก" (กะ หรือ OT ที่จบก่อน/ตรงกับเริ่มกะ = OT ก่อนกะ) ──
    // ปกติ 3 นาที · ฝนตก 15 นาที. OT เย็น (หลังกะ + มีช่วงว่าง) ไม่ตรวจสาย — clock ครั้งเดียวรู้แค่เวลาออก
    let firstStartMin = plannedStart ? toMin(plannedStart) : null;
    otReqs.forEach(function(o) {
      if (o.start && o.end && plannedStart && toMin(o.end) <= toMin(plannedStart)) {
        const s = toMin(o.start);
        if (firstStartMin === null || s < firstStartMin) firstStartMin = s;
      }
    });
    let isLate = false, lateMin = 0, lateKind = '';
    if (clockIn && firstStartMin !== null) {
      const ci = toMin(clockIn);
      if (ci > firstStartMin + (rain ? 15 : 3)) {
        isLate   = true;
        lateMin  = ci - firstStartMin;
        lateKind = (plannedStart && firstStartMin < toMin(plannedStart)) ? 'ot' : 'shift';
      }
    }

    // ── คำนวณ OT (แยกตามแรง 1 / 1.5 / 2) ──
    // มาตรฐาน 9 ชม./วัน (ทำงาน 8 + พัก 1). OT×1 คิดจาก plan ที่ยาวเกิน 9 ชม. (ถึง plan end)
    let planHours = 0;
    if (plannedStart && plannedEnd) planHours = (toMin(plannedEnd) - toMin(plannedStart)) / 60;
    const planExcess = Math.max(0, planHours - 9);   // ชม. plan ที่เกินมาตรฐาน

    let ot1 = 0, ot15 = 0, ot2 = 0;
    const otItems = [];                              // breakdown ต่อคำขอ OT (สำหรับ report)

    // OT จากแผนกะที่ยาวเกิน 9 ชม. (planned OT) — คงเดิม
    if (completed && isSpecial) ot2 += planExcess;
    else if (completed)         ot1 += planExcess;

    // OT จากคำขอ — สรุป = min(ที่ขอ, เวลาจริงที่ clock ทับกับหน้าต่าง OT)
    // clock ครั้งเดียวคร่อมทั้งวัน → clip [clockIn,clockOut] กับ [o.start,o.end] · ช่วงว่างตกไปเอง
    if (clockIn) otReqs.forEach(function(o) {
      let before = false, gapMin = 0;
      if (o.start && o.end && plannedStart && plannedEnd) {
        if (toMin(o.end) <= toMin(plannedStart))      { before = true;  gapMin = toMin(plannedStart) - toMin(o.end); }
        else if (toMin(o.start) >= toMin(plannedEnd))  { before = false; gapMin = toMin(o.start) - toMin(plannedEnd); }
        else { before = ((toMin(o.start) + toMin(o.end)) / 2) < toMin(plannedStart); gapMin = 0; }
      }
      const continuous = gapMin <= 60;               // ห่าง ≤ 60 นาที = ต่อเนื่อง

      let actualHrs = null, summaryHrs = o.hours;     // ยังไม่ clock out → ใช้ที่ขอไปก่อน (ชั่วคราว)
      if (completed && o.start && o.end) {
        const ws = Math.max(toMin(clockIn),  toMin(o.start));
        const we = Math.min(toMin(clockOut), toMin(o.end));
        actualHrs  = r2(Math.max(0, (we - ws) / 60));
        summaryHrs = Math.min(o.hours, actualHrs);   // ทำเกินไม่จ่ายเพิ่ม · ทำขาดหักตามจริง
      }
      summaryHrs = r2(summaryHrs);

      const rate = isSpecial ? 2 : (before ? 1.5 : 1);
      if (rate === 2)        ot2  += summaryHrs;
      else if (rate === 1.5) ot15 += summaryHrs;
      else                   ot1  += summaryHrs;

      otItems.push({
        reqStart:    o.start,
        reqEnd:      o.end,
        reqHours:    r2(o.hours),
        actualHours: actualHrs,          // null = ยังไม่ครบ (ไม่มี clock out)
        summaryHours: summaryHrs,
        rate:        rate,
        kind:        before ? 'before' : 'after',
        continuous:  continuous,
        gapMin:      gapMin,
      });
    });

    ot1 = r2(ot1); ot15 = r2(ot15); ot2 = r2(ot2);
    const otHours    = r2(ot1 + ot15 + ot2);                  // ชม. OT รวม (สรุปแล้ว)
    const otWeighted = r2(ot1 * 1 + ot15 * 1.5 + ot2 * 2);    // แรงรวม (สำหรับคิดเงิน)

    // ── ชั่วโมงในกะ (ปกติ): max(0, เวลาเลิกตาราง − max(clock-in, เวลาเริ่มตาราง)) ──
    // มาสาย=หักชั่วโมง · มาก่อน=ไม่เพิ่ม · ไม่อิง clock-out · ไม่หักพัก · ไม่ติดลบ
    let workedHours = 0;
    if (clockIn && plannedStart && plannedEnd) {
      const startMin = Math.max(toMin(clockIn), toMin(plannedStart));
      workedHours = Math.max(0, (toMin(plannedEnd) - startMin) / 60);
    }
    workedHours = r2(workedHours);

    records.push({
      date:         dateStr,
      clockIn:      clockIn,
      clockOut:     clockOut,
      plannedStart: plannedStart,
      plannedEnd:   plannedEnd,
      workedHours:  workedHours,
      specialDay:   isSpecial ? 'พิเศษ' : 'ปกติ',
      isLate:       isLate,
      lateMin:      lateMin,
      lateKind:     lateKind,
      rain:         rain,
      ot1:          ot1,
      ot15:         ot15,
      ot2:          ot2,
      otHours:      otHours,
      otWeighted:   otWeighted,
      otItems:      otItems,
    });
  }
  // แทรกวันลาที่อนุมัติแล้วเป็นแถว "ลา" (รายงานรายเดือน) — ใช้รายการวันลาที่เก็บไว้
  // (ไม่อิงตารางงาน เพราะวันลาถูกลบออกจากตารางตอน approve แล้ว)
  if (month && year && empId) {
    const lvData = ss.getSheetByName(SH.LEAVE).getDataRange().getValues();
    const LEAVE_LBL = { v: 'ลาพักร้อน', s: 'ลาป่วย', p: 'ลากิจ' };
    const seen = {};
    records.forEach(function(r) { seen[r.date] = true; });
    for (let i = 1; i < lvData.length; i++) {
      const r = lvData[i];
      if (!r[0]) continue;
      if (String(r[1]) !== String(empId)) continue;
      if (String(r[7]).toLowerCase() !== 'approved') continue;
      let dts = [];
      try { dts = r[12] ? JSON.parse(r[12]) : []; } catch (e) { dts = []; }
      if (!dts.length && r[3] && r[4]) {   // fallback เรคคอร์ดเก่า
        const s = new Date(r[3]), e = new Date(r[4]), c = new Date(s);
        while (c <= e) { dts.push(fmtDate(c)); c.setDate(c.getDate() + 1); }
      }
      const label = LEAVE_LBL[String(r[2])] || 'ลา';
      dts.forEach(function(ds) {
        const d = new Date(ds);
        if (isNaN(d)) return;
        if ((d.getMonth() + 1) !== parseInt(month) || d.getFullYear() !== parseInt(year)) return;
        if (seen[ds]) return;
        records.push({
          date: ds, clockIn: '', clockOut: '', plannedStart: '', plannedEnd: '',
          specialDay: label, isLate: false, ot1: 0, ot15: 0, ot2: 0, otHours: 0, otWeighted: 0, isLeave: true,
        });
        seen[ds] = true;
      });
    }
  }

  // Newest first — ensures data[0] is the most recent record
  records.sort(function(a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; });
  if (limit && parseInt(limit) > 0) records.splice(parseInt(limit));
  return { ok: true, data: records };
}

// ════════════════════════════════════════════════════════════
// LEAVE REQUESTS — คำขอวันลา
// ════════════════════════════════════════════════════════════
// Columns: [0]id [1]empId [2]type [3]startDate [4]endDate
//          [5]days [6]reason [7]status [8]submittedAt [9]approvedAt [10]approvedBy [11]medCertUrl [12]leaveDatesJSON

// Leave entitlements per year (calendar year, no carry-over)
const LEAVE_ENTITLEMENT = { v: 6, s: 30, p: 3 }; // vacation, sick, personal

function submitLeave(p) {
  if (!p.empId || !p.type || !p.startDate || !p.endDate)
    return { ok: false, error: 'ข้อมูลไม่ครบถ้วน' };

  const start = new Date(p.startDate);
  const end   = new Date(p.endDate);
  if (end < start) return { ok: false, error: 'วันที่สิ้นสุดต้องไม่ก่อนวันที่เริ่ม' };

  // Validate 3-day advance for vacation and personal leave
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  const diffDays = Math.ceil((start - now) / 86400000);
  if (p.type !== 's' && diffDays < 3)
    return { ok: false, error: 'ลาพักร้อน/ลากิจ ต้องแจ้งล่วงหน้าอย่างน้อย 3 วัน' };

  const ss = openSS();

  // นับวันลา = เฉพาะวันที่อยู่ในตารางงานที่ "อนุมัติแล้ว" (วันทำงานจริง) เท่านั้น
  // วันที่พนักงานไม่ได้ลงตาราง = วันหยุดของเขาอยู่แล้ว ไม่ต้องหักวันลา
  const schData = ss.getSheetByName(SH.SCHEDULES).getDataRange().getValues();
  const workDates = {};
  for (let i = 1; i < schData.length; i++) {
    const r = schData[i];
    if (!r[0]) continue;
    if (String(r[1]) !== String(p.empId)) continue;
    if (String(r[5]).toLowerCase() !== 'approved') continue;
    const dts = r[4] ? JSON.parse(r[4]) : {};
    Object.keys(dts).forEach(function(d) { workDates[d] = true; });
  }
  // เก็บ "รายการวันลาจริง" = วันในช่วงที่อยู่ในตารางงาน (วันเหล่านี้จะถูกลบออกจากตารางตอน approve)
  const leaveDateList = [];
  const cursor = new Date(start);
  while (cursor <= end) {
    const ds = fmtDate(cursor);
    if (workDates[ds]) leaveDateList.push(ds);
    cursor.setDate(cursor.getDate() + 1);
  }
  const days = leaveDateList.length;
  if (days === 0)
    return { ok: false, error: 'ช่วงวันที่เลือกไม่มีวันทำงานในตารางที่อนุมัติ — กรุณาสร้าง/รออนุมัติตารางงานของวันนั้นก่อน' };

  // Check remaining balance
  const year = start.getFullYear();
  const balance = getLeaveBalance(p.empId, year);
  if (balance.ok) {
    const remaining = balance.data[p.type] !== undefined ? balance.data[p.type] : LEAVE_ENTITLEMENT[p.type] || 0;
    if (days > remaining)
      return { ok: false, error: 'วันลาคงเหลือไม่เพียงพอ (เหลือ ' + remaining + ' วัน, ขอ ' + days + ' วัน)' };
  }

  // อัปโหลดใบรับรองแพทย์ขึ้น Google Drive (ถ้าแนบมา) แล้วเก็บลิงก์
  let medCertUrl = '';
  if (p.medCertData) {
    try {
      const folder = getOrCreateFolder_('Kyoto Shi - ใบรับรองแพทย์');
      const bytes  = Utilities.base64Decode(p.medCertData);
      const fname  = (p.empId || 'emp') + '_' + fmtDate(start) + '_' + Date.now() + '_' + (p.medCertName || 'medcert');
      const blob   = Utilities.newBlob(bytes, p.medCertType || 'application/octet-stream', fname);
      const file   = folder.createFile(blob);
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
      medCertUrl = file.getUrl();
    } catch (e) { medCertUrl = ''; }
  }

  const sh    = ss.getSheetByName(SH.LEAVE);
  const newId = 'LV' + Date.now();
  const now2  = fmtDateTime(new Date());
  sh.appendRow([
    newId,
    p.empId,
    p.type,
    p.startDate,
    p.endDate,
    days,
    p.reason || '',
    'Pending',
    now2,
    '',
    '',
    medCertUrl,                      // [11] ลิงก์ใบรับรองแพทย์ใน Google Drive
    JSON.stringify(leaveDateList),   // [12] รายการวันลาจริง (yyyy-MM-dd)
  ]);
  return { ok: true, leaveId: newId, days, medCertUrl: medCertUrl };
}

function approveLeave(p) {
  // เมื่ออนุมัติ → ยกเลิก (ลบ) วันลาออกจากตารางงานที่อนุมัติไว้
  if (p.approve) {
    const ss = openSS();
    const data = ss.getSheetByName(SH.LEAVE).getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === p.leaveId) {
        let dates = [];
        try { dates = data[i][12] ? JSON.parse(data[i][12]) : []; } catch (e) { dates = []; }
        // เรคคอร์ดเก่าที่ไม่มี leaveDatesJSON → ใช้ช่วง start..end
        if (!dates.length && data[i][3] && data[i][4]) {
          const s = new Date(data[i][3]), e = new Date(data[i][4]), c = new Date(s);
          while (c <= e) { dates.push(fmtDate(c)); c.setDate(c.getDate() + 1); }
        }
        if (dates.length) removeDatesFromSchedule_(data[i][1], dates);
        break;
      }
    }
  }
  return setLeaveStatus(p.leaveId, p.approve ? 'Approved' : 'Rejected', p.approvedBy);
}

// ลบวันที่ระบุออกจากตารางงานที่ "อนุมัติแล้ว" ของพนักงาน (ใช้ตอนอนุมัติการลา)
function removeDatesFromSchedule_(empId, dates) {
  const ss = openSS();
  const sh = ss.getSheetByName(SH.SCHEDULES);
  const data = sh.getDataRange().getValues();
  const dateSet = {};
  dates.forEach(function(d) { dateSet[d] = true; });
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (String(data[i][1]) !== String(empId)) continue;
    if (String(data[i][5]).toLowerCase() !== 'approved') continue;
    let obj = {};
    try { obj = data[i][4] ? JSON.parse(data[i][4]) : {}; } catch (e) { obj = {}; }
    let changed = false;
    Object.keys(dateSet).forEach(function(d) { if (obj[d]) { delete obj[d]; changed = true; } });
    if (changed) sh.getRange(i + 1, 5).setValue(JSON.stringify(obj));
  }
}

function setLeaveStatus(leaveId, status, approvedBy) {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.LEAVE);
  const data = sh.getDataRange().getValues();
  const now  = fmtDateTime(new Date());
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === leaveId) {
      sh.getRange(i + 1, 8).setValue(status);
      sh.getRange(i + 1, 10).setValue(now);
      sh.getRange(i + 1, 11).setValue(approvedBy || 'Owner');
      return { ok: true };
    }
  }
  return { ok: false, error: 'Leave request not found' };
}

function getLeaveBalance(empId, year) {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.LEAVE);
  const data = sh.getDataRange().getValues();
  const yr   = parseInt(year) || new Date().getFullYear();

  const used = { v: 0, s: 0, p: 0 };
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (String(row[1]) !== String(empId)) continue;
    if (row[7] !== 'Approved') continue;
    const d = row[3] ? new Date(row[3]) : null;
    if (!d || d.getFullYear() !== yr) continue;
    const type = String(row[2]);
    if (used[type] !== undefined) used[type] += Number(row[5]) || 0;
  }

  const balance = {};
  Object.keys(LEAVE_ENTITLEMENT).forEach(t => {
    balance[t] = Math.max(0, LEAVE_ENTITLEMENT[t] - used[t]);
  });
  return { ok: true, data: balance, used, entitlement: LEAVE_ENTITLEMENT };
}

// ประวัติการลาของพนักงาน (ทุกสถานะ) เรียงใหม่สุดก่อน — กรองตามปีได้
function getLeaveHistory(empId, year) {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.LEAVE);
  const data = sh.getDataRange().getValues();
  const yr   = year ? parseInt(year) : null;
  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (empId && String(row[1]) !== String(empId)) continue;
    const startStr = fmtMaybe(row[3]);
    if (yr) {
      const d = row[3] ? new Date(row[3]) : null;
      if (!d || d.getFullYear() !== yr) continue;
    }
    records.push({
      id:        row[0],
      empId:     row[1],
      type:      row[2],
      startDate: startStr,
      endDate:   fmtMaybe(row[4]),
      days:      row[5],
      reason:    row[6],
      status:    String(row[7] || '').toLowerCase(),  // approved/rejected/pending
      submittedAt: row[8],
      approvedAt:  row[9],
      approvedBy:  row[10],
      medCertUrl:  row[11] || '',
    });
  }
  records.sort(function(a, b) { return a.startDate < b.startDate ? 1 : a.startDate > b.startDate ? -1 : 0; });
  return { ok: true, data: records };
}

// ── Update getPendingApprovals to include Leave ──
function getPendingApprovals() {
  const ss      = openSS();
  const empSh   = ss.getSheetByName(SH.EMPLOYEES);
  const schSh   = ss.getSheetByName(SH.SCHEDULES);
  const otSh    = ss.getSheetByName(SH.OT_REQUESTS);
  const lvSh    = ss.getSheetByName(SH.LEAVE);

  const empData = empSh.getDataRange().getValues();
  const empMap  = {};
  for (let i = 1; i < empData.length; i++) empMap[empData[i][0]] = empData[i][1];

  // Pending schedules (+ แนบค่า "ก่อนแก้" สำหรับคำขอแก้ไข เพื่อโชว์ ก่อน→หลัง)
  const schData = schSh.getDataRange().getValues();
  const schById = {};
  for (let i = 1; i < schData.length; i++) { if (schData[i][0]) { try { schById[String(schData[i][0])] = schData[i][4] ? JSON.parse(schData[i][4]) : {}; } catch (e) { schById[String(schData[i][0])] = {}; } } }
  const pendSch = [];
  for (let i = 1; i < schData.length; i++) {
    if (schData[i][5] === 'Pending') {
      const isEdit = schData[i][9] === 'TRUE' || schData[i][9] === true;
      const origId = String(schData[i][10] || '');
      const editDates = schData[i][4] ? JSON.parse(schData[i][4]) : {};
      const before = {};
      if (isEdit && origId && schById[origId]) { const od = schById[origId]; Object.keys(editDates).forEach(function (k) { before[k] = od[k] || null; }); }
      pendSch.push({
        id: schData[i][0], empId: schData[i][1], name: empMap[schData[i][1]] || schData[i][1],
        month: schData[i][2], year: schData[i][3], dates: editDates, submittedAt: schData[i][6],
        isEdit: isEdit, originalScheduleId: origId, before: before,
      });
    }
  }

  // Pending OT
  const otData = otSh.getDataRange().getValues();
  const pendOT = [];
  for (let i = 1; i < otData.length; i++) {
    if (otData[i][5] === 'Pending') {
      pendOT.push({
        id: otData[i][0], empId: otData[i][1],
        name: empMap[otData[i][1]] || otData[i][1],
        date: otData[i][2] ? fmtDate(new Date(otData[i][2])) : '',
        requestedHours: otData[i][3], reason: otData[i][4],
        submittedAt: otData[i][6],
      });
    }
  }

  // Pending leave
  const lvData  = lvSh.getDataRange().getValues();
  const pendLv  = [];
  for (let i = 1; i < lvData.length; i++) {
    if (lvData[i][7] === 'Pending') {
      const typeMap = { v: 'ลาพักร้อน', s: 'ลาป่วย', p: 'ลากิจ' };
      pendLv.push({
        id: lvData[i][0], empId: lvData[i][1],
        name: empMap[lvData[i][1]] || lvData[i][1],
        type: lvData[i][2], typeName: typeMap[lvData[i][2]] || lvData[i][2],
        startDate: lvData[i][3] ? fmtDate(new Date(lvData[i][3])) : '',
        endDate:   lvData[i][4] ? fmtDate(new Date(lvData[i][4])) : '',
        days: lvData[i][5], reason: lvData[i][6],
        submittedAt: lvData[i][8],
        medCertUrl: lvData[i][11] || '',
      });
    }
  }

  // Pending registrations — return ALL fields the approval screen needs
  const pendReg = [];
  for (let i = 1; i < empData.length; i++) {
    if (empData[i][10] === 'Pending') {
      pendReg.push({
        empId:       empData[i][0],
        name:        empData[i][1] || empData[i][4] || empData[i][2],
        username:    empData[i][2],
        position:    empData[i][5],
        type:        empData[i][6],
        bank:        empData[i][8],   // bank brand
        bankName:    empData[i][8],   // alias used by the approval screen
        startDate:   fmtMaybe(empData[i][9]),
        phone:       empData[i][12] || '',
        idcard:      empData[i][13] || '',
        address:     empData[i][14] || '',
        dob:         fmtMaybe(empData[i][15]),
        bankAcc:     empData[i][16] || '',
        bankAccName: empData[i][17] || '',
      });
    }
  }

  return {
    ok: true,
    data: { schedules: pendSch, ot: pendOT, leave: pendLv, registrations: pendReg },
  };
}

// ════════════════════════════════════════════════════════════
// MIGRATE — อัปเดต Schedules Sheet ให้รองรับ Edit Schedule
// รันครั้งเดียวจาก Apps Script Editor: เลือก migrateSchedulesSheet แล้วกด ▶ Run
// ════════════════════════════════════════════════════════════
// เรียกผ่าน URL: ?action=migrateSchedules
// หรือรันตรงจาก editor ก็ได้: migrateSchedulesSheet()
function migrateSchedulesColumns() {
  const ss  = openSS();
  const sh  = ss.getSheetByName('Schedules');
  if (!sh) return { ok: false, error: 'ไม่พบ Sheet Schedules' };

  const lastCol = sh.getLastColumn();
  const headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const style   = function(cell) { cell.setFontWeight('bold').setBackground('#1A3560').setFontColor('#ffffff'); };
  const added   = [];

  if (headers[9] !== 'isEdit') {
    const c = sh.getRange(1, 10); c.setValue('isEdit'); style(c); sh.setColumnWidth(10, 70);
    added.push('K: isEdit');
  }
  if (headers[10] !== 'originalScheduleId') {
    const c = sh.getRange(1, 11); c.setValue('originalScheduleId'); style(c); sh.setColumnWidth(11, 160);
    added.push('L: originalScheduleId');
  }

  // Leave_Requests — เพิ่ม header approvedBy (K) และ medCertUrl (L)
  const lv = ss.getSheetByName('Leave_Requests');
  if (lv) {
    const lvLast = lv.getLastColumn();
    const lvH = lvLast > 0 ? lv.getRange(1, 1, 1, lvLast).getValues()[0] : [];
    if (lvH[10] !== 'approvedBy') { const c = lv.getRange(1, 11); c.setValue('approvedBy'); style(c); sh.setColumnWidth(11, 120); added.push('Leave K: approvedBy'); }
    if (lvH[11] !== 'medCertUrl') { const c = lv.getRange(1, 12); c.setValue('medCertUrl'); style(c); lv.setColumnWidth(12, 220); added.push('Leave L: medCertUrl'); }
    if (lvH[12] !== 'leaveDates') { const c = lv.getRange(1, 13); c.setValue('leaveDates'); style(c); lv.setColumnWidth(13, 200); added.push('Leave M: leaveDates'); }
  }

  return { ok: true, added: added, message: added.length ? 'เพิ่ม column: ' + added.join(', ') : 'Column ครบแล้ว' };
}

function migrateSchedulesSheet() { // alias สำหรับรันจาก editor
  const r = migrateSchedulesColumns();
  SpreadsheetApp.getUi().alert(r.ok ? '✅ ' + r.message : '❌ ' + r.error);
}

// ════════════════════════════════════════════════════════════
// SETUP — สร้าง Sheets อัตโนมัติ (รันครั้งเดียวตอนเริ่มต้น)
// ════════════════════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const defs = [
    {
      name:    'Employees',
      headers: ['id','name','username','password','nickname','position','type','salary','bank','startDate','status','probationEnd','phone','idcard','address','dob','bankAcc','bankAccName','note','otrate','pdpaConsent'],
    },
    {
      name:    'Attendance',
      headers: ['id','empId','date','clockIn','clockOut','plannedStart','plannedEnd','specialDay'],
    },
    {
      name:    'Schedules',
      headers: ['id','empId','month','year','datesJSON','status','submittedAt','approvedAt','approvedBy','isEdit','originalScheduleId'],
    },
    {
      name:    'OT_Requests',
      headers: ['id','empId','date','requestedHours','reason','status','submittedAt','approvedAt','approvedBy','startTime','endTime'],
    },
    {
      name:    'Leave_Requests',
      headers: ['id','empId','type','startDate','endDate','days','reason','status','submittedAt','approvedAt','approvedBy','medCertUrl','leaveDates'],
    },
  ];

  defs.forEach(function(def) {
    let sh = ss.getSheetByName(def.name);
    if (!sh) sh = ss.insertSheet(def.name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(def.headers);
      sh.getRange(1, 1, 1, def.headers.length)
        .setFontWeight('bold')
        .setBackground('#1A3560')
        .setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
  });

  // Add demo accounts if Employees sheet is empty
  const empSh = ss.getSheetByName('Employees');
  if (empSh.getLastRow() <= 1) {
    empSh.appendRow(['EMP001','เจ้าของร้าน','menbom888','mlb888','เจ้าของ','Owner','Owner',0,'','2024-01-01','Active','','','','','','','','','']);
    empSh.appendRow(['EMP002','นาม สุขใจ','nam_sukjai','emp1234','นาม','เบเกอรี่ (Bakery)','Full-time',13000,'กสิกรไทย (KBank)','2025-01-01','Active','','081-234-5678','1-2345-67890-12-3','123 ถ.สุขุมวิท ต.เนินพระ อ.เมือง ระยอง 21000','1998-05-20','123-4-56789-0','นาม สุขใจ','','75']);
  }

  SpreadsheetApp.getUi().alert('✅ ตั้งค่า Google Sheets สำเร็จ!\nพร้อมใช้งานแล้วครับ');
}

// ════════════════════════════════════════════════════════════
// UTILITIES
// ════════════════════════════════════════════════════════════
// หา/สร้างโฟลเดอร์ใน Google Drive ตามชื่อ (สำหรับเก็บใบรับรองแพทย์)
function getOrCreateFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function fmtDate(d) {
  return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
}

function fmtDateTime(d) {
  return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
}

// Format a cell that may hold either a Date object or a plain string (e.g. 'yyyy-MM-dd')
function fmtMaybe(v) {
  if (!v) return '';
  return (v instanceof Date) ? fmtDate(v) : String(v);
}

// Format a clock-time cell back to "HH:mm". If Sheets auto-converted "08:25" into a
// time value, getValues() returns a Date — convert it back; otherwise return the text.
function fmtTime(v) {
  if (v === '' || v === null || v === undefined) return '';
  return (v instanceof Date) ? Utilities.formatDate(v, 'Asia/Bangkok', 'HH:mm') : String(v);
}

// One-way SHA-256 hash (hex) for storing passwords. Never store/return plaintext.
function hashPw(s) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}


// ═══════════════════════════════════════════════════════════════════════════
// PREP STOCK v2 — "Daily handover loop" PATCH  (July 5 2026)
// ─ วางฟังก์ชันทั้งหมดนี้ต่อท้าย Code.gs ในโปรเจกต์ Apps Script (โปรเจกต์ "Management system")
// ─ แล้วเพิ่ม 3 case ใน switch ของ doGet (ดูบล็อก "// ▶ ADD THESE CASES" ด้านล่าง)
// ─ จากนั้น Deploy → Manage deployments → Edit(ดินสอ) → New version → Deploy (SAME /exec)
//
// เป็น ADDITIVE ล้วน — ไม่แก้ฟังก์ชัน prep เดิม (getPrepList/savePrepLog/getShortages ฯลฯ ทำงานเหมือนเดิม)
// เพิ่ม 2 ชีตใหม่: Prep_Ack (รับกะเช้า) · Prep_Plan (เป้าของพรุ่งนี้)
// เพิ่ม 3 action: getPrepStatus (อ่านรวมทีเดียว—แก้ throttle) · saveMorningAck · savePrepPlan
//
// blocks (วัตถุดิบหมดบล็อกของเตรียมตัวไหน) ไม่ต้องแก้ backend — frontend เก็บใน note ของ Prep_Shortages
// ด้วย convention "⟦blk:ชื่อของเตรียม⟧ ..." แล้ว getPrepStatus ถอดออกมาให้
// ═══════════════════════════════════════════════════════════════════════════

// ▶ ADD THESE CASES  (วางใน switch(p.action) ของ doGet — ในกลุ่ม "อ่านข้อมูล"/"บันทึกข้อมูล")
//      case 'getPrepStatus':   return res(getPrepStatus());
//      case 'saveMorningAck':  return res(saveMorningAck(p));
//      case 'savePrepPlan':    return res(savePrepPlan(p));

const PREP2 = {
  ITEMS:     'Prep_Items',
  LOGS:      'Prep_Logs',
  SHORTAGES: 'Prep_Shortages',
  ACK:       'Prep_Ack',
  PLAN:      'Prep_Plan',
};
const PREP2_KITCHENS = ['บาร์น้ำ', 'เบเกอรี่', 'ครัวอาหาร'];

function prep2Ensure_() {
  const ss = openSS();
  if (!ss.getSheetByName(PREP2.ACK)) {
    const a = ss.insertSheet(PREP2.ACK);
    a.appendRow(['id', 'date', 'kitchen', 'itemId', 'itemName', 'done', 'qty', 'empId', 'empName', 'ts']);
    a.setFrozenRows(1);
  }
  if (!ss.getSheetByName(PREP2.PLAN)) {
    const pl = ss.insertSheet(PREP2.PLAN);
    pl.appendRow(['id', 'date', 'kitchen', 'itemId', 'target', 'empId', 'ts']);
    pl.setFrozenRows(1);
  }
  return ss;
}

// วันที่แบบ string เสมอ (กัน Google Sheets coerce cell เป็น Date object → เทียบ string ไม่ตรง)
function prep2DStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, openSS().getSpreadsheetTimeZone() || 'Asia/Bangkok', 'yyyy-MM-dd');
  return String(v == null ? '' : v).slice(0, 10);
}
function prep2Today_() {
  return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
}
function prep2AddDays_(dateStr, delta) {
  const p = String(dateStr).split('-').map(Number);
  const d = new Date(p[0], p[1] - 1, p[2]); d.setDate(d.getDate() + delta);
  const z = n => ('0' + n).slice(-2);
  return d.getFullYear() + '-' + z(d.getMonth() + 1) + '-' + z(d.getDate());
}
function prep2Rows_(name) {
  const sh = openSS().getSheetByName(name);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
}
function prep2Bool_(v) { return v === true || v === 'TRUE' || v === 'true' || v === 1 || v === '1'; }
// ถอด "⟦blk:xxx⟧ note..." → { blocks:'xxx', note:'note...' }
function prep2ParseBlk_(raw) {
  const s = String(raw == null ? '' : raw);
  const m = s.match(/⟦blk:([^⟧]*)⟧\s*/);
  if (!m) return { blocks: '', note: s.trim() };
  return { blocks: (m[1] || '').trim(), note: s.replace(m[0], '').trim() };
}

// ── SAVE: รับกะเช้า (staff ติ๊ก "เตรียมแล้ว" + จำนวน) — 1 ชุดต่อครัวต่อวัน ─────────
function saveMorningAck(p) {
  const ss = prep2Ensure_();
  const sh = ss.getSheetByName(PREP2.ACK);
  const date = prep2Today_();
  const kitchen = String(p.kitchen || '');
  if (!kitchen) return { ok: false, error: 'no kitchen' };
  let items = [];
  try { items = JSON.parse(p.items || '[]'); } catch (e) { items = []; }
  const emp = swEmpKitchen_(p.empId || '');
  // ลบของวันนี้ครัวนี้ก่อน แล้วเขียนใหม่ (กัน duplicate)
  prep2DelRows_(sh, date, kitchen);
  const ts = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
  let n = 0;
  items.forEach(it => {
    const done = prep2Bool_(it.done);
    if (!done && (it.qty == null || it.qty === '')) return; // ข้ามรายการที่ยังไม่แตะ
    if (done) n++;
    sh.appendRow(['A' + Date.now() + Math.floor(Math.random() * 1000),
      date, kitchen, String(it.itemId || ''), String(it.itemName || ''),
      done, (it.qty === '' || it.qty == null) ? '' : Number(it.qty) || 0,
      String(p.empId || ''), emp.name || '', ts]);
  });
  return { ok: true, ackCount: n };
}

// ── SAVE: เป้าของพรุ่งนี้ (ตั้งตอนปิดร้าน) — 1 ชุดต่อครัวต่อวัน ────────────────────
function savePrepPlan(p) {
  const ss = prep2Ensure_();
  const sh = ss.getSheetByName(PREP2.PLAN);
  const date = prep2Today_();
  const kitchen = String(p.kitchen || '');
  if (!kitchen) return { ok: false, error: 'no kitchen' };
  let items = [];
  try { items = JSON.parse(p.items || '[]'); } catch (e) { items = []; }
  prep2DelRows_(sh, date, kitchen);
  const ts = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
  items.forEach(it => {
    if (it.target === '' || it.target == null) return;
    sh.appendRow(['P' + Date.now() + Math.floor(Math.random() * 1000),
      date, kitchen, String(it.itemId || ''), Number(it.target) || 0, String(p.empId || ''), ts]);
  });
  return { ok: true };
}

// ลบแถวที่ date(col0) + kitchen(col2) ตรงกัน (schema: [id,date,kitchen,...])
function prep2DelRows_(sh, date, kitchen) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const r = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  for (let i = r.length - 1; i >= 0; i--) {
    if (prep2DStr_(r[i][1]) === prep2DStr_(date) && String(r[i][2]) === String(kitchen)) {
      sh.deleteRow(i + 2);
    }
  }
}

// ═══ READ รวมทีเดียว — ใช้ได้ทั้ง owner dashboard / จอเช้า / จอปิดร้าน ═══════════════
// อ่าน 5 ชีตครั้งเดียว ประกอบทุกครัว → ลด fan-out (เดิม owner ยิง 5 request)
function getPrepStatus() {
  prep2Ensure_();
  const today = prep2Today_();
  const yest = prep2AddDays_(today, -1);

  // items ต่อครัว (active เท่านั้น)  schema Prep_Items[0 id,1 kitchen,2 name,3 unit,4 active,5 createdAt]
  const itemRows = prep2Rows_(PREP2.ITEMS);
  const itemsByKit = {}, itemName = {}, itemUnit = {};
  itemRows.forEach(x => {
    if (!x[0]) return;
    if (x[4] === false || x[4] === 'FALSE' || x[4] === 'false') return;
    const k = String(x[1]);
    (itemsByKit[k] = itemsByKit[k] || []).push({ itemId: String(x[0]), name: String(x[2]), unit: String(x[3] || '') });
    itemName[String(x[0])] = String(x[2]);
    itemUnit[String(x[0])] = String(x[3] || '');
  });

  // logs ทั้งหมด  schema Prep_Logs[0 id,1 date,2 kitchen,3 itemId,4 itemName,5 produced,6 used,7 left,8 handover,9 ready,10 note,11 empId,12 empName,13 ts]
  const logRows = prep2Rows_(PREP2.LOGS);
  const logByDateKit = {};           // key = date|kitchen → { byId, ts, empName }
  const loggedDays = {};             // key = kitchen → Set(date)
  logRows.forEach(x => {
    if (!x[0]) return;
    const d = prep2DStr_(x[1]), k = String(x[2]);
    const key = d + '|' + k;
    if (!logByDateKit[key]) logByDateKit[key] = { byId: {}, ts: '', empName: '' };
    logByDateKit[key].byId[String(x[3])] = {
      produced: Number(x[5]) || 0, used: Number(x[6]) || 0, left: Number(x[7]) || 0,
      handover: prep2Bool_(x[8]), ready: !(x[9] === false || x[9] === 'FALSE' || x[9] === 'false'),
      note: String(x[10] || ''), itemName: String(x[4] || '')
    };
    const ts = String(x[13] || '');
    if (ts > logByDateKit[key].ts) { logByDateKit[key].ts = ts; logByDateKit[key].empName = String(x[12] || ''); }
    (loggedDays[k] = loggedDays[k] || {})[d] = true;
  });

  // ack วันนี้  schema Prep_Ack[0 id,1 date,2 kitchen,3 itemId,4 itemName,5 done,6 qty,...]
  const ackRows = prep2Rows_(PREP2.ACK);
  const ackByKit = {};
  ackRows.forEach(x => {
    if (!x[0] || prep2DStr_(x[1]) !== today) return;
    const k = String(x[2]);
    (ackByKit[k] = ackByKit[k] || {})[String(x[3])] = { done: prep2Bool_(x[5]), qty: (x[6] === '' ? '' : Number(x[6]) || 0) };
  });

  // plan  schema Prep_Plan[0 id,1 date,2 kitchen,3 itemId,4 target,...]
  const planRows = prep2Rows_(PREP2.PLAN);
  const planByDateKit = {};
  planRows.forEach(x => {
    if (!x[0]) return;
    const key = prep2DStr_(x[1]) + '|' + String(x[2]);
    (planByDateKit[key] = planByDateKit[key] || {})[String(x[3])] = Number(x[4]) || 0;
  });

  // shortages วันนี้  schema Prep_Shortages[0 id,1 date,2 kitchen,3 material,4 note,...]
  const shRows = prep2Rows_(PREP2.SHORTAGES);
  const shByKit = {}, shortagesAll = [];
  shRows.forEach(x => {
    if (!x[0] || prep2DStr_(x[1]) !== today) return;
    const k = String(x[2]);
    const parsed = prep2ParseBlk_(x[3]);              // block tag ฝังในชื่อวัตถุดิบ (material) — savePrepLog เดิมเก็บคอลัมน์นี้แน่นอน
    const rec = { material: parsed.note || String(x[3]), blocks: parsed.blocks, note: String(x[4] || '') };
    (shByKit[k] = shByKit[k] || []).push(rec);
    shortagesAll.push({ kitchen: k, material: rec.material, blocks: rec.blocks, note: rec.note });
  });

  const kitchens = PREP2_KITCHENS.map(k => {
    const list = itemsByKit[k] || [];
    const todayLog = logByDateKit[today + '|' + k] || { byId: {}, ts: '', empName: '' };
    const logToday = todayLog.byId;
    const logged = Object.keys(logToday).length > 0;

    let sumP = 0, sumU = 0, sumL = 0; const handoverNames = [], notReadyNames = [];
    Object.keys(logToday).forEach(id => {
      const o = logToday[id];
      sumP += o.produced; sumU += o.used; sumL += o.left;
      const nm = o.itemName || itemName[id] || id;
      if (o.handover) handoverNames.push(nm);
      if (o.ready === false) notReadyNames.push(nm);
    });

    // fromDate = วันปิดร้านล่าสุดที่ < วันนี้ (ที่ครัวนี้เคยลง log)
    let fromDate = '';
    Object.keys(loggedDays[k] || {}).forEach(d => { if (d < today && d > fromDate) fromDate = d; });
    const yLogFull = logByDateKit[fromDate + '|' + k] || { byId: {}, ts: '', empName: '' };
    const yLog = yLogFull.byId;
    const yPlan = planByDateKit[fromDate + '|' + k] || {};
    const ack = ackByKit[k] || {};

    // toPrep = ของที่กะปิดร้านส่งต่อ (handover) หรือ เมื่อวานไม่พร้อมขาย
    const shBlocks = {}; (shByKit[k] || []).forEach(s => { if (s.blocks) shBlocks[s.blocks] = true; });
    const toPrep = [];
    Object.keys(yLog).forEach(id => {
      const o = yLog[id];
      if (!(o.handover || o.ready === false)) return;
      const nm = o.itemName || itemName[id] || id;
      const a = ack[id] || {};
      toPrep.push({
        itemId: id, itemName: nm, unit: itemUnit[id] || '',
        target: (yPlan[id] != null ? yPlan[id] : ''), left: o.left, ready: o.ready, note: o.note,
        blocked: !!(shBlocks[id] || shBlocks[nm]),
        done: !!a.done, qty: (a.qty == null ? '' : a.qty)
      });
    });
    const toPrepCount = toPrep.length;
    const ackCount = toPrep.filter(t => t.done).length;
    const anyBlockedOpen = toPrep.some(t => t.blocked && !t.done);
    const prepDone = toPrepCount > 0 && ackCount >= toPrepCount;

    // ไฟสถานะ: แดง=ไม่ส่งกะเมื่อวาน(ที่มีของต้องส่ง) || มีของถูกบล็อกยังไม่จัดการ · เขียว=ไม่มีอะไรต้องเตรียม||เตรียมครบ · เหลือง=กำลังเตรียม
    let readiness;
    if ((!fromDate && list.length > 0) || anyBlockedOpen) readiness = 'risk';
    else if (toPrepCount === 0 || prepDone) readiness = 'ready';
    else readiness = 'follow';

    // dots 7 วันหลังสุด (today-6..today) ครัวนี้ลง log ไหม
    const dots = [];
    for (let i = 6; i >= 0; i--) dots.push(!!(loggedDays[k] || {})[prep2AddDays_(today, -i)]);

    return {
      kitchen: k, listCount: list.length, list: list,
      logged: logged, sendTime: todayLog.ts, sendBy: todayLog.empName,
      sumProduced: sumP, sumUsed: sumU, sumLeft: sumL,
      logToday: logToday, ackToday: ack, planToday: (planByDateKit[today + '|' + k] || {}),
      handoverNames: handoverNames, notReadyNames: notReadyNames,
      shortages: shByKit[k] || [],
      fromDate: fromDate, handoverTime: yLogFull.ts, handoverBy: yLogFull.empName,
      toPrep: toPrep, toPrepCount: toPrepCount, ackCount: ackCount, prepDone: prepDone,
      readiness: readiness, dots: dots
    };
  });

  return { ok: true, data: { date: today, kitchens: kitchens, shortagesAll: shortagesAll } };
}
