// ============================================================
// Code.gs — IEP Sub Tracker | Server-Side Logic
// ============================================================

// ── Sheet Names ──────────────────────────────────────────────
const SHEET_TEACHERS      = "Teachers";
const SHEET_SUBS          = "OnSiteSubs";
const SHEET_IEP_REQUESTS  = "IEPRequests";
const SHEET_MASTER_SCHED  = "MasterSchedule";
const SHEET_SCHEDULES     = "ScheduleTypes";
const SHEET_PERIOD_TIMES  = "PeriodTimes";

// ── Sheet IDs ────────────────────────────────────────────────
const SCHOOL_INFO = {
  'El Cajon Valley': { logoUrl: 'https://braves.guhsd.net/images/logo.png', primaryColor: '#C70011',  sheetID: '' },
  'El Capitan': { logoUrl: 'https://elcapitan.guhsd.net/images/logo.png', primaryColor: '#FFB618',  sheetID: '' },
  'Granite Hills': { logoUrl: 'https://granite.guhsd.net/images/logo.png', primaryColor: '#75B2DD',  sheetID: '' },
  'Grossmont': { logoUrl: 'https://www.foothillers.com/images/logo.png', primaryColor: '#2c66b8',  sheetID: '' },
  'Monte Vista': { logoUrl: 'https://montevista.guhsd.net/images/logo.png', primaryColor: '#AC1818',  sheetID: '' },
  'Mount Miguel': { logoUrl: 'https://mountmiguel.guhsd.net/images/logo.png', primaryColor: '#CF0A2C',  sheetID: '' },
  'Santana': { logoUrl: 'https://santana.guhsd.net/images/logo.png', primaryColor: '#9807e0',  sheetID: '' },
  'Valhalla': { logoUrl: 'https://valhalla.guhsd.net/images/logo.png', primaryColor: '#FF5F00', sheetID: '1FHTin0DKNoh8tMG9eCL4fmitCJWjJo6QxxDN5yp0JWA' },
  'West Hills': { logoUrl: 'https://wolfpack.guhsd.net/images/logo.png', primaryColor: '#1f52c2',  sheetID: '' }
}

// ── Entry Point ──────────────────────────────────────────────
function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("IEP Sub Tracker")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// Allows HTML files to include other HTML files via <?= include('File') ?>
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Spreadsheet Helper ───────────────────────────────────────
function getSpreadsheet() {
  return SpreadsheetApp.openById('1FHTin0DKNoh8tMG9eCL4fmitCJWjJo6QxxDN5yp0JWA');
}

function getOrCreateSheet(name, headers) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground("#1a1a2e")
        .setFontColor("#ffffff")
        .setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map((row, i) => {
    const obj = { _row: i + 2 };
    headers.forEach((h, j) => { obj[h] = row[j]; });
    return obj;
  });
}

// ── Initialization ───────────────────────────────────────────
function initializeSheets() {
  getOrCreateSheet(SHEET_TEACHERS,     ["ID", "Name", "Email", "Active"]);
  getOrCreateSheet(SHEET_SUBS,         ["ID", "Name", "Email", "Active"]);
  getOrCreateSheet(SHEET_IEP_REQUESTS, ["ID", "Date", "TeacherID", "TeacherName", "IEPTime", "Periods", "AssignedSubID", "AssignedSubName", "ScheduleType", "Notes", "Status", "CreatedAt"]);
  getOrCreateSheet(SHEET_MASTER_SCHED, ["TeacherID", "TeacherName", "Period", "CourseName", "RoleType"]);
  getOrCreateSheet(SHEET_SCHEDULES,    ["ID", "Name", "Active"]);
  getOrCreateSheet(SHEET_PERIOD_TIMES, ["ScheduleID", "Period", "StartTime", "EndTime"]);
  initDefaultSchedules();
  return { success: true };
}

function initDefaultSchedules() {
  const sheet = getOrCreateSheet(SHEET_SCHEDULES, ["ID", "Name", "Active"]);
  const existing = sheetToObjects(sheet);
  if (existing.length > 0) return;

  const defaults = [
    ["regular",       "Regular Day",       true],
    ["minimum",       "Minimum Day",       true],
    ["collaboration", "Collaboration Day", true],
    ["late_start",    "Late Start",        true],
  ];
  defaults.forEach(d => sheet.appendRow(d));

  const ptSheet = getOrCreateSheet(SHEET_PERIOD_TIMES, ["ScheduleID", "Period", "StartTime", "EndTime"]);
  const defaultTimes = [
    ["regular", 1, "8:00 AM",  "8:52 AM"],
    ["regular", 2, "8:56 AM",  "9:48 AM"],
    ["regular", 3, "9:52 AM",  "10:44 AM"],
    ["regular", 4, "10:48 AM", "11:40 AM"],
    ["regular", 5, "11:44 AM", "12:36 PM"],
    ["regular", 6, "12:40 PM", "1:32 PM"],
    ["regular", 7, "1:36 PM",  "2:28 PM"],
    ["minimum", 1, "8:00 AM",  "8:35 AM"],
    ["minimum", 2, "8:39 AM",  "9:14 AM"],
    ["minimum", 3, "9:18 AM",  "9:53 AM"],
    ["minimum", 4, "9:57 AM",  "10:32 AM"],
    ["minimum", 5, "10:36 AM", "11:11 AM"],
    ["minimum", 6, "11:15 AM", "11:50 AM"],
    ["minimum", 7, "11:54 AM", "12:29 PM"],
    ["collaboration", 1, "9:00 AM",  "9:45 AM"],
    ["collaboration", 2, "9:49 AM",  "10:34 AM"],
    ["collaboration", 3, "10:38 AM", "11:23 AM"],
    ["collaboration", 4, "11:27 AM", "12:12 PM"],
    ["collaboration", 5, "12:16 PM", "1:01 PM"],
    ["collaboration", 6, "1:05 PM",  "1:50 PM"],
    ["collaboration", 7, "1:54 PM",  "2:39 PM"],
    ["late_start", 1, "9:00 AM",  "9:52 AM"],
    ["late_start", 2, "9:56 AM",  "10:48 AM"],
    ["late_start", 3, "10:52 AM", "11:44 AM"],
    ["late_start", 4, "11:48 AM", "12:40 PM"],
    ["late_start", 5, "12:44 PM", "1:36 PM"],
    ["late_start", 6, "1:40 PM",  "2:32 PM"],
    ["late_start", 7, "2:36 PM",  "3:28 PM"],
  ];
  defaultTimes.forEach(r => ptSheet.appendRow(r));
}

// ── Teacher Functions ────────────────────────────────────────
function getTeachers() {
  const sheet = getOrCreateSheet(SHEET_TEACHERS, ["ID", "Name", "Email", "Active"]);
  return sheetToObjects(sheet).filter(t => t.Active === true || t.Active === "TRUE");
}

function saveTeacher(teacher) {
  const sheet = getOrCreateSheet(SHEET_TEACHERS, ["ID", "Name", "Email", "Active"]);
  if (teacher.ID) {
    const data = sheetToObjects(sheet);
    const found = data.find(t => t.ID === teacher.ID);
    if (found) {
      sheet.getRange(found._row, 1, 1, 4).setValues([[teacher.ID, teacher.Name, teacher.Email, teacher.Active]]);
      return { success: true };
    }
  }
  const id = "T_" + new Date().getTime();
  sheet.appendRow([id, teacher.Name, teacher.Email, true]);
  return { success: true, id };
}

function deleteTeacher(teacherID) {
  const sheet = getOrCreateSheet(SHEET_TEACHERS, ["ID", "Name", "Email", "Active"]);
  const data = sheetToObjects(sheet);
  const found = data.find(t => t.ID === teacherID);
  if (found) {
    sheet.getRange(found._row, 4).setValue(false); // Soft delete
    return { success: true };
  }
  return { success: false, error: "Teacher not found" };
}

// ── On-Site Sub Functions ────────────────────────────────────
function getSubs() {
  const sheet = getOrCreateSheet(SHEET_SUBS, ["ID", "Name", "Email", "Active"]);
  return sheetToObjects(sheet).filter(s => s.Active === true || s.Active === "TRUE");
}

function saveSub(sub) {
  const sheet = getOrCreateSheet(SHEET_SUBS, ["ID", "Name", "Email", "Active"]);
  if (sub.ID) {
    const data = sheetToObjects(sheet);
    const found = data.find(s => s.ID === sub.ID);
    if (found) {
      sheet.getRange(found._row, 1, 1, 4).setValues([[sub.ID, sub.Name, sub.Email, sub.Active]]);
      return { success: true };
    }
  }
  const id = "S_" + new Date().getTime();
  sheet.appendRow([id, sub.Name, sub.Email, true]);
  return { success: true, id };
}

function deleteSub(subID) {
  const sheet = getOrCreateSheet(SHEET_SUBS, ["ID", "Name", "Email", "Active"]);
  const data = sheetToObjects(sheet);
  const found = data.find(s => s.ID === subID);
  if (found) {
    sheet.getRange(found._row, 4).setValue(false);
    return { success: true };
  }
  return { success: false, error: "Sub not found" };
}

// ── Master Schedule Functions ────────────────────────────────
function getMasterSchedule() {
  const sheet = getOrCreateSheet(SHEET_MASTER_SCHED, ["TeacherID", "TeacherName", "Period", "CourseName", "RoleType"]);
  return sheetToObjects(sheet);
}

function getTeacherSchedule(teacherID) {
  const all = getMasterSchedule();
  return all
    .filter(r => r.TeacherID === teacherID)
    .filter(r => {
      const course = (r.CourseName || "").toUpperCase();
      const role   = (r.RoleType  || "").toUpperCase();
      return course !== "PREP" && role !== "CO-TEACHER" && role !== "COTEACHER";
    });
}

function saveMasterScheduleRow(row) {
  const sheet = getOrCreateSheet(SHEET_MASTER_SCHED, ["TeacherID", "TeacherName", "Period", "CourseName", "RoleType"]);
  sheet.appendRow([row.TeacherID, row.TeacherName, row.Period, row.CourseName, row.RoleType]);
  return { success: true };
}

function clearTeacherSchedule(teacherID) {
  const sheet = getOrCreateSheet(SHEET_MASTER_SCHED, ["TeacherID", "TeacherName", "Period", "CourseName", "RoleType"]);
  const data = sheetToObjects(sheet);
  const rows = data.filter(r => r.TeacherID === teacherID).map(r => r._row).sort((a,b) => b - a);
  rows.forEach(r => sheet.deleteRow(r));
  return { success: true };
}

// ── Schedule Type Functions ──────────────────────────────────
function getScheduleTypes() {
  const sheet = getOrCreateSheet(SHEET_SCHEDULES, ["ID", "Name", "Active"]);
  return sheetToObjects(sheet).filter(s => s.Active === true || s.Active === "TRUE");
}

function saveScheduleType(sched) {
  const sheet = getOrCreateSheet(SHEET_SCHEDULES, ["ID", "Name", "Active"]);
  if (sched.ID) {
    const data = sheetToObjects(sheet);
    const found = data.find(s => s.ID === sched.ID);
    if (found) {
      sheet.getRange(found._row, 1, 1, 3).setValues([[sched.ID, sched.Name, sched.Active]]);
      return { success: true };
    }
  }
  const id = "SC_" + new Date().getTime();
  sheet.appendRow([id, sched.Name, true]);
  return { success: true, id };
}

// ── Period Time Functions ────────────────────────────────────
function getPeriodTimes(scheduleID) {
  const sheet = getOrCreateSheet(SHEET_PERIOD_TIMES, ["ScheduleID", "Period", "StartTime", "EndTime"]);
  const all = sheetToObjects(sheet);
  return scheduleID ? all.filter(r => r.ScheduleID === scheduleID) : all;
}

function savePeriodTimes(scheduleID, periods) {
  // periods = [{period, startTime, endTime}, ...]
  const sheet = getOrCreateSheet(SHEET_PERIOD_TIMES, ["ScheduleID", "Period", "StartTime", "EndTime"]);
  const all = sheetToObjects(sheet);
  const existing = all.filter(r => r.ScheduleID === scheduleID).map(r => r._row).sort((a,b) => b - a);
  existing.forEach(r => sheet.deleteRow(r));
  periods.forEach(p => sheet.appendRow([scheduleID, p.period, p.startTime, p.endTime]));
  return { success: true };
}

// ── IEP Request Functions ────────────────────────────────────
function getIEPRequests() {
  const sheet = getOrCreateSheet(SHEET_IEP_REQUESTS,
    ["ID","Date","TeacherID","TeacherName","IEPTime","Periods","AssignedSubID","AssignedSubName","ScheduleType","Notes","Status","CreatedAt"]);
  return sheetToObjects(sheet).map(r => ({
    ...r,
    Periods: typeof r.Periods === "string" ? r.Periods : String(r.Periods || "")
  }));
}

function saveIEPRequest(req) {
  const sheet = getOrCreateSheet(SHEET_IEP_REQUESTS,
    ["ID","Date","TeacherID","TeacherName","IEPTime","Periods","AssignedSubID","AssignedSubName","ScheduleType","Notes","Status","CreatedAt"]);

  const periodsStr = Array.isArray(req.Periods) ? req.Periods.join(",") : (req.Periods || "");

  if (req.ID) {
    const data = sheetToObjects(sheet);
    const found = data.find(r => r.ID === req.ID);
    if (found) {
      sheet.getRange(found._row, 1, 1, 12).setValues([[
        req.ID, req.Date, req.TeacherID, req.TeacherName,
        req.IEPTime, periodsStr,
        req.AssignedSubID || "", req.AssignedSubName || "",
        req.ScheduleType || "", req.Notes || "",
        req.Status || "Pending", req.CreatedAt || new Date().toISOString()
      ]]);
      return { success: true };
    }
  }

  const id = "IEP_" + new Date().getTime();
  const now = new Date().toISOString();
  sheet.appendRow([
    id, req.Date, req.TeacherID, req.TeacherName,
    req.IEPTime, periodsStr,
    req.AssignedSubID || "", req.AssignedSubName || "",
    req.ScheduleType || "regular", req.Notes || "",
    req.Status || "Pending", now
  ]);
  return { success: true, id };
}

function deleteIEPRequest(id) {
  const sheet = getOrCreateSheet(SHEET_IEP_REQUESTS,
    ["ID","Date","TeacherID","TeacherName","IEPTime","Periods","AssignedSubID","AssignedSubName","ScheduleType","Notes","Status","CreatedAt"]);
  const data = sheetToObjects(sheet);
  const found = data.find(r => r.ID === id);
  if (found) {
    sheet.deleteRow(found._row);
    return { success: true };
  }
  return { success: false, error: "Request not found" };
}

// ── Print Data ───────────────────────────────────────────────
function getDailyPrintData(dateStr) {
  const requests = getIEPRequests().filter(r => r.Date === dateStr && r.Status !== "Cancelled");
  const periodTimes = {};
  getPeriodTimes().forEach(pt => {
    if (!periodTimes[pt.ScheduleID]) periodTimes[pt.ScheduleID] = {};
    periodTimes[pt.ScheduleID][pt.Period] = { start: pt.StartTime, end: pt.EndTime };
  });
  return { requests, periodTimes };
}

// ── All Data Bootstrap ───────────────────────────────────────
function getAllData() {
  initializeSheets();
  return {
    teachers:      getTeachers(),
    subs:          getSubs(),
    scheduleTypes: getScheduleTypes(),
    periodTimes:   getPeriodTimes(),
    iepRequests:   getIEPRequests(),
    masterSchedule: getMasterSchedule()
  };
}
