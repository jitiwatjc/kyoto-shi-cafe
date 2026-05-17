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
      case 'getAttendance':   return res(getAttendance(p.empId, p.month, p.year));
      case 'getSchedules':    return res(getSchedules(p.empId));
      case 'getPending':      return res(getPendingApprovals());
      case 'getTodayClock':        return res(getTodayClock(p.empId, p.date));
      case 'getTodayAttendance':   return res(getTodayAttendance(p.date));
      case 'getLeaveBalance':      return res(getLeaveBalance(p.empId, p.year));
      case 'ping':            return res({ ok: true, message: 'Kyoto Shi API online' });

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
      case 'submitOT':        return res(submitOT(p));
      case 'submitLeave':     return res(submitLeave(p));
      case 'approveLeave':    return res(approveLeave({ leaveId: p.leaveId, approve: p.approve === 'true', approvedBy: p.approvedBy }));
      case 'approveOT':       return res(approveOT({ otId: p.otId, approve: p.approve === 'true', approvedBy: p.approvedBy }));
      case 'registerEmployee':return res(registerEmployee(JSON.parse(decodeURIComponent(p.data || '{}'))));
      case 'approveEmployee': return res(approveEmployee({ empId: p.empId, approve: p.approve === 'true' }));
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
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[2]) === String(username) &&
        String(row[3]) === String(password) &&
        row[10] === 'Active') {
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
      password:    row[3],
      nickname:    row[4] || row[1],
      position:    row[5],
      type:        row[6],
      salary:      row[7],
      bank:        row[8],
      startDate:   row[9] ? fmtDate(new Date(row[9])) : '',
      status:      row[10],
      probationEnd:row[11] ? fmtDate(new Date(row[11])) : '',
    });
  }
  return { ok: true, data: employees };
}

function registerEmployee(p) {
  const ss  = openSS();
  const sh  = ss.getSheetByName(SH.EMPLOYEES);
  const lastRow = sh.getLastRow();
  const newId   = 'EMP' + String(lastRow).padStart(3, '0');
  sh.appendRow([
    newId,
    p.name        || '',
    p.username    || '',
    p.password    || '',
    p.nickname    || p.name || '',
    p.position    || '',
    p.type        || 'Full-time',
    p.salary      || 0,
    p.bank        || '',
    p.startDate   || '',
    'Pending',
    p.probationEnd|| '',
  ]);
  return { ok: true, empId: newId };
}

function approveEmployee(p) {
  return setEmployeeStatus(p.empId, p.approve ? 'Active' : 'Rejected');
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
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === p.empId) {
      const r = i + 1;
      if (p.name        !== undefined) sh.getRange(r, 2).setValue(p.name);
      if (p.nickname    !== undefined) sh.getRange(r, 5).setValue(p.nickname);
      if (p.position    !== undefined) sh.getRange(r, 6).setValue(p.position);
      if (p.type        !== undefined) sh.getRange(r, 7).setValue(p.type);
      if (p.salary      !== undefined) sh.getRange(r, 8).setValue(p.salary);
      if (p.bank        !== undefined) sh.getRange(r, 9).setValue(p.bank);
      if (p.startDate   !== undefined) sh.getRange(r, 10).setValue(p.startDate);
      if (p.status      !== undefined) sh.getRange(r, 11).setValue(p.status);
      if (p.probationEnd!== undefined) sh.getRange(r, 12).setValue(p.probationEnd);
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
          clockIn:      row[3],
          clockOut:     row[4],
          plannedStart: row[5],
          plannedEnd:   row[6],
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
  const attData = attSh.getDataRange().getValues();
  const empData = empSh.getDataRange().getValues();

  // Build empId → {name, position} map
  const empMap = {};
  for (let i = 1; i < empData.length; i++) {
    const id = String(empData[i][0]);
    empMap[id] = { name: empData[i][1], position: empData[i][6] };
  }

  // Filter attendance rows for today
  const rows = [];
  for (let i = 1; i < attData.length; i++) {
    const row     = attData[i];
    const rowDate = row[2] ? fmtDate(new Date(row[2])) : '';
    if (rowDate !== date) continue;
    const empId     = String(row[1]);
    const clockIn   = row[3] ? String(row[3]) : '';
    const planStart = row[5] ? String(row[5]) : '';
    let late = false;
    if (clockIn && planStart) {
      const [ph, pm] = planStart.split(':').map(Number);
      const [ch, cm] = clockIn.split(':').map(Number);
      late = (ch * 60 + cm) > (ph * 60 + pm + 15);
    }
    rows.push({
      empId,
      name:     empMap[empId]?.name     || empId,
      position: empMap[empId]?.position || '',
      clockIn,
      clockOut: row[4] ? String(row[4]) : '',
      late,
    });
  }
  return { ok: true, data: rows };
}

function clockIn(p) {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.ATTENDANCE);
  const data = sh.getDataRange().getValues();
  // Update existing row if found
  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][2] ? fmtDate(new Date(data[i][2])) : '';
    if (String(data[i][1]) === String(p.empId) && rowDate === p.date) {
      sh.getRange(i + 1, 4).setValue(p.time);
      return { ok: true };
    }
  }
  // Create new row
  const newId = 'ATT' + Date.now();
  sh.appendRow([
    newId,
    p.empId,
    p.date,
    p.time,
    '',
    p.plannedStart || '',
    p.plannedEnd   || '',
    p.specialDay   || false,
  ]);
  return { ok: true, id: newId };
}

function clockOut(p) {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.ATTENDANCE);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][2] ? fmtDate(new Date(data[i][2])) : '';
    if (String(data[i][1]) === String(p.empId) && rowDate === p.date) {
      sh.getRange(i + 1, 5).setValue(p.time);
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
      id:          row[0],
      empId:       row[1],
      month:       row[2],
      year:        row[3],
      dates:       row[4] ? JSON.parse(row[4]) : {},
      status:      row[5],
      submittedAt: row[6],
      approvedAt:  row[7],
      approvedBy:  row[8],
    });
  }
  return { ok: true, data: result };
}

function submitSchedule(p) {
  const ss    = openSS();
  const sh    = ss.getSheetByName(SH.SCHEDULES);
  const newId = 'SCH' + Date.now();
  const now   = fmtDateTime(new Date());
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
  return setScheduleStatus(p.scheduleId, 'Approved', p.approvedBy);
}

function rejectSchedule(p) {
  return setScheduleStatus(p.scheduleId, 'Rejected', p.approvedBy);
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
  ]);
  return { ok: true, otId: newId };
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
// PENDING APPROVALS — รอการอนุมัติ
// ════════════════════════════════════════════════════════════
function getPendingApprovals() {
  const ss      = openSS();
  const empSh   = ss.getSheetByName(SH.EMPLOYEES);
  const schSh   = ss.getSheetByName(SH.SCHEDULES);
  const otSh    = ss.getSheetByName(SH.OT_REQUESTS);

  // Build employee name map
  const empData = empSh.getDataRange().getValues();
  const empMap  = {};
  for (let i = 1; i < empData.length; i++) {
    empMap[empData[i][0]] = empData[i][1];
  }

  // Pending schedules
  const schData  = schSh.getDataRange().getValues();
  const pendSch  = [];
  for (let i = 1; i < schData.length; i++) {
    if (schData[i][5] === 'Pending') {
      pendSch.push({
        id:          schData[i][0],
        empId:       schData[i][1],
        name:        empMap[schData[i][1]] || schData[i][1],
        month:       schData[i][2],
        year:        schData[i][3],
        dates:       schData[i][4] ? JSON.parse(schData[i][4]) : {},
        submittedAt: schData[i][6],
      });
    }
  }

  // Pending OT
  const otData  = otSh.getDataRange().getValues();
  const pendOT  = [];
  for (let i = 1; i < otData.length; i++) {
    if (otData[i][5] === 'Pending') {
      pendOT.push({
        id:             otData[i][0],
        empId:          otData[i][1],
        name:           empMap[otData[i][1]] || otData[i][1],
        date:           otData[i][2] ? fmtDate(new Date(otData[i][2])) : '',
        requestedHours: otData[i][3],
        reason:         otData[i][4],
        submittedAt:    otData[i][6],
      });
    }
  }

  // Pending registrations
  const pendReg = [];
  for (let i = 1; i < empData.length; i++) {
    if (empData[i][10] === 'Pending') {
      pendReg.push({
        empId:    empData[i][0],
        name:     empData[i][1],
        username: empData[i][2],
        position: empData[i][5],
        type:     empData[i][6],
        bank:     empData[i][8],
      });
    }
  }

  return {
    ok: true,
    data: { schedules: pendSch, ot: pendOT, registrations: pendReg },
  };
}

// ════════════════════════════════════════════════════════════
// ATTENDANCE — ดึงบันทึกเวลา (สำหรับ Export)
// ════════════════════════════════════════════════════════════
function getAttendance(empId, month, year) {
  const ss   = openSS();
  const sh   = ss.getSheetByName(SH.ATTENDANCE);
  const data = sh.getDataRange().getValues();
  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (empId && String(row[1]) !== String(empId)) continue;
    const d = row[2] ? new Date(row[2]) : null;
    if (!d) continue;
    if (month && (d.getMonth() + 1) !== parseInt(month)) continue;
    if (year  && d.getFullYear()   !== parseInt(year))  continue;
    records.push({
      date:         fmtDate(d),
      clockIn:      row[3] || '',
      clockOut:     row[4] || '',
      plannedStart: row[5] || '',
      plannedEnd:   row[6] || '',
      specialDay:   row[7] || false,
    });
  }
  return { ok: true, data: records };
}

// ════════════════════════════════════════════════════════════
// LEAVE REQUESTS — คำขอวันลา
// ════════════════════════════════════════════════════════════
// Columns: [0]id [1]empId [2]type [3]startDate [4]endDate
//          [5]days [6]reason [7]status [8]submittedAt [9]approvedAt [10]approvedBy

// Leave entitlements per year (calendar year, no carry-over)
const LEAVE_ENTITLEMENT = { v: 6, s: 30, p: 3 }; // vacation, sick, personal

function submitLeave(p) {
  if (!p.empId || !p.type || !p.startDate || !p.endDate)
    return { ok: false, error: 'ข้อมูลไม่ครบถ้วน' };

  // Validate 3-day advance for vacation and personal leave
  const start  = new Date(p.startDate);
  const now    = new Date();
  now.setHours(0, 0, 0, 0);
  const diffDays = Math.ceil((start - now) / 86400000);
  if (p.type !== 's' && diffDays < 3)
    return { ok: false, error: 'ลาพักร้อน/ลากิจ ต้องแจ้งล่วงหน้าอย่างน้อย 3 วัน' };

  // Count requested days (weekdays only, exclude Thai holidays — simplified)
  const end = new Date(p.endDate);
  let days = 0;
  const cursor = new Date(start);
  while (cursor <= end) {
    const dow = cursor.getDay();
    if (dow !== 0 && dow !== 6) days++;
    cursor.setDate(cursor.getDate() + 1);
  }

  // Check remaining balance
  const year = start.getFullYear();
  const balance = getLeaveBalance(p.empId, year);
  if (balance.ok) {
    const remaining = balance.data[p.type] !== undefined ? balance.data[p.type] : LEAVE_ENTITLEMENT[p.type] || 0;
    if (days > remaining)
      return { ok: false, error: 'วันลาคงเหลือไม่เพียงพอ (เหลือ ' + remaining + ' วัน, ขอ ' + days + ' วัน)' };
  }

  const ss    = openSS();
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
  ]);
  return { ok: true, leaveId: newId, days };
}

function approveLeave(p) {
  return setLeaveStatus(p.leaveId, p.approve ? 'Approved' : 'Rejected', p.approvedBy);
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

  // Pending schedules
  const schData = schSh.getDataRange().getValues();
  const pendSch = [];
  for (let i = 1; i < schData.length; i++) {
    if (schData[i][5] === 'Pending') {
      pendSch.push({
        id: schData[i][0], empId: schData[i][1],
        name: empMap[schData[i][1]] || schData[i][1],
        month: schData[i][2], year: schData[i][3],
        dates: schData[i][4] ? JSON.parse(schData[i][4]) : {},
        submittedAt: schData[i][6],
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
      });
    }
  }

  // Pending registrations
  const pendReg = [];
  for (let i = 1; i < empData.length; i++) {
    if (empData[i][10] === 'Pending') {
      pendReg.push({
        empId: empData[i][0], name: empData[i][1],
        username: empData[i][2], position: empData[i][5],
        type: empData[i][6], bank: empData[i][8],
      });
    }
  }

  return {
    ok: true,
    data: { schedules: pendSch, ot: pendOT, leave: pendLv, registrations: pendReg },
  };
}

// ════════════════════════════════════════════════════════════
// SETUP — สร้าง Sheets อัตโนมัติ (รันครั้งเดียวตอนเริ่มต้น)
// ════════════════════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const defs = [
    {
      name:    'Employees',
      headers: ['id','name','username','password','nickname','position','type','salary','bank','startDate','status','probationEnd'],
    },
    {
      name:    'Attendance',
      headers: ['id','empId','date','clockIn','clockOut','plannedStart','plannedEnd','specialDay'],
    },
    {
      name:    'Schedules',
      headers: ['id','empId','month','year','datesJSON','status','submittedAt','approvedAt','approvedBy'],
    },
    {
      name:    'OT_Requests',
      headers: ['id','empId','date','requestedHours','reason','status','submittedAt','approvedAt','approvedBy'],
    },
    {
      name:    'Leave_Requests',
      headers: ['id','empId','type','startDate','endDate','days','reason','status','submittedAt','approvedAt'],
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
    empSh.appendRow(['EMP001','เจ้าของร้าน','menbom888','mlb888','เจ้าของ','Owner','Owner',0,'','2024-01-01','Active','']);
    empSh.appendRow(['EMP002','นาม สุขใจ','nam_sukjai','emp1234','นาม','เบเกอรี่','Full-time',13000,'KBank 123-4-56789-0','2025-01-01','Active','']);
  }

  SpreadsheetApp.getUi().alert('✅ ตั้งค่า Google Sheets สำเร็จ!\nพร้อมใช้งานแล้วครับ');
}

// ════════════════════════════════════════════════════════════
// UTILITIES
// ════════════════════════════════════════════════════════════
function fmtDate(d) {
  return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
}

function fmtDateTime(d) {
  return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
}
