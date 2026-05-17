# Kyoto Shi Café — Management System · Phase I

ระบบจัดการร้านคาเฟ่แบบครบวงจร สำหรับ Kyoto Shi Café ระยอง

**Live Demo:** [https://jitiwatjc.github.io/kyoto-shi-cafe/](https://jitiwatjc.github.io/kyoto-shi-cafe/)

---

## ขั้นตอนเชื่อมต่อ Google Sheets (Backend Setup)

ระบบใช้ Google Sheets เป็นฐานข้อมูล และ Google Apps Script เป็น API

### ขั้นที่ 1 — สร้าง Google Sheet

1. ไปที่ [sheets.google.com](https://sheets.google.com) → **Blank spreadsheet**
2. ตั้งชื่อ: `Kyoto Shi Cafe DB`
3. คัดลอก **Sheet ID** จาก URL:
   ```
   https://docs.google.com/spreadsheets/d/【SHEET_ID_อยู่ตรงนี้】/edit
   ```

### ขั้นที่ 2 — ติดตั้ง Apps Script

1. ใน Google Sheet → **Extensions → Apps Script**
2. ลบโค้ดเดิมทิ้งทั้งหมด
3. คัดลอกโค้ดจากไฟล์ `Code.gs` วางแทน
4. แก้ไขบรรทัด:
   ```javascript
   const SHEET_ID = 'วาง-SHEET-ID-ของคุณ-ที่นี่';
   ```
5. **Save** (Ctrl+S)

### ขั้นที่ 3 — สร้าง Sheets อัตโนมัติ

1. ใน Apps Script → เลือก function `setupSheets` → **Run**
2. ยืนยัน Permission ที่ขึ้นมา → Allow
3. จะเห็น popup "✅ ตั้งค่า Google Sheets สำเร็จ!"
4. กลับไปดู Google Sheet — จะมี 5 Sheets สร้างขึ้นอัตโนมัติ:
   - `Employees` · `Attendance` · `Schedules` · `OT_Requests` · `Leave_Requests`

### ขั้นที่ 4 — Deploy เป็น Web App

1. ใน Apps Script → **Deploy → New deployment**
2. ตั้งค่า:
   - **Type:** Web app
   - **Execute as:** Me
   - **Who has access:** Anyone
3. คลิก **Deploy** → คัดลอก **Web app URL**

### ขั้นที่ 5 — เชื่อมต่อ Frontend

เปิดไฟล์ `index.html` (หรือ `Kyoto_Shi_Prototype_Phase1.html`) หาบรรทัด:
```javascript
const SCRIPT_URL='YOUR_APPS_SCRIPT_URL';
```
แก้เป็น:
```javascript
const SCRIPT_URL='https://script.google.com/macros/s/AKfy.../exec';
```
(วาง URL ที่ได้จากขั้นที่ 4)

จากนั้นรัน `Deploy_To_GitHub.ps1` อีกครั้งเพื่ออัปเดตเว็บ

---

## บัญชีเริ่มต้น (Demo Accounts)

| บทบาท | Username | Password |
|-------|----------|----------|
| Owner | `menbom888` | `mlb888` |
| พนักงาน | `nam_sukjai` | `emp1234` |

> ⚠️ เปลี่ยน Password ทันทีหลังเริ่มใช้งานจริง

---

## โครงสร้าง Google Sheets

| Sheet | คอลัมน์หลัก |
|-------|-------------|
| `Employees` | id, name, username, password, nickname, position, type, salary, bank, startDate, status, probationEnd |
| `Attendance` | id, empId, date, clockIn, clockOut, plannedStart, plannedEnd, specialDay |
| `Schedules` | id, empId, month, year, datesJSON, status, submittedAt, approvedAt, approvedBy |
| `OT_Requests` | id, empId, date, requestedHours, reason, status, submittedAt, approvedAt, approvedBy |
| `Leave_Requests` | id, empId, type, startDate, endDate, days, reason, status, submittedAt, approvedAt, approvedBy |

---

## Phase Roadmap

| Phase | ระยะเวลา | งาน | สถานะ |
|-------|----------|-----|-------|
| **Phase 1** | ม.ค.–ก.พ. | Auth, RBAC, ข้อมูลพนักงาน, ตาราง, Clock in/out, วันลา | ✅ เสร็จแล้ว |
| Phase 2 | ก.พ.–เม.ย. | OT, KPI, Payroll, Pay Slip PDF, เบิกล่วงหน้า | 🔜 |
| Phase 3 | มี.ค.–พ.ค. | ยอดขาย, Stock, ของเสีย, Incentive | 🔜 |
| Phase 4 | พ.ค.–ก.ค. | บัญชี, ใบกำกับภาษี, กำไร-ขาดทุน | 🔜 |
| Phase 5 | พ.ค.–ส.ค. | Notifications, UAT, Training, Go-live | 🔜 |

---

## GPS Coordinates (ต้องวัดในสถานที่จริง)

ไฟล์ `index.html` → ค้นหา `SHOP_LAT` และ `SHOP_LNG`:
```javascript
const SHOP_LAT = 12.6853;   // ← อัปเดตด้วย Latitude จริงของร้าน
const SHOP_LNG = 101.2782;  // ← อัปเดตด้วย Longitude จริงของร้าน
const GPS_RADIUS_M = 50;    // รัศมี 50 เมตร (ปรับได้)
```

วัดพิกัดโดยไปที่ร้านแล้วเปิด Google Maps → กดค้างที่ตำแหน่งร้าน → อ่านค่า Lat/Lng

---

*Kyoto Shi Café Management System · Phase I · พฤษภาคม 2026*
