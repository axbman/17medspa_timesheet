// ─────────────────────────────────────────────────────────────
//  TimeClock - Google Apps Script  (Multi-Branch Edition)
//  Paste into Apps Script, then:
//  1. Set MASTER_SPREADSHEET_ID below
//  2. Set DRIVE_ROOT_FOLDER_ID to the parent folder that will
//     hold all branch sub-folders (or leave blank for My Drive)
//  3. Deploy as Web App (Execute as: Me, Anyone can access)
//  4. Run setup() once to create branch tabs
//  5. Run installTriggers() once to enable midnight auto-check
//     and automatic new-month sheet creation
//
//  MASTER SPREADSHEET LAYOUT
//  ─────────────────────────
//  One tab per branch, named exactly:  Branch - Downtown
//  Each tab has two columns:  Name  |  PIN
//
//  DRIVE LAYOUT
//  ────────────
//  <Root Folder>/
//    Downtown/
//      June 2026 - Downtown      ← monthly spreadsheet
//      July 2026 - Downtown
//    Cupertino/
//      June 2026 - Cupertino
//    ...
// ─────────────────────────────────────────────────────────────

const MASTER_SPREADSHEET_ID = '1yzY4MZcfkoifgdFq9jrwnx3JxvqnRpLhtsj93hTZNs4';
const DRIVE_ROOT_FOLDER_ID  = '';          // optional; blank = My Drive root
const BRANCH_TAB_PREFIX     = ''; // tabs are named directly, e.g. "Buena Park"
const SUMMARY_SHEET         = 'Summary';
const LOCATIONS_SHEET       = 'Locations';
const FORM_TITLE            = 'Clock-Out Correction Form';

// Stores the Google Form ID + pre-fill entry IDs in Script Properties
// Run createCorrectionForm() once to set these up.

// ── Web App entry points ──────────────────────────────────────

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'verifyPin') {
      return respond(verifyPin(data.pin));
    }

    if (action === 'punch') {
      const name       = (data.name   || '').trim();
      const type       = (data.type   || '').trim();
      const branch     = (data.branch || '').trim();
      const clientTime = (data.clientTime || '').trim(); // timestamp from employee's phone

      if (!name) return respond({ error: 'Missing employee name.' });
      if (!type) return respond({ error: 'Missing punch type (IN or OUT).' });

      return respond(recordPunch(name, branch, type, clientTime));
    }

    if (action === 'getLocations') {
      return respond(getLocations());
    }

    if (action === 'getStatus') {
      const status = getCurrentStatus((data.name || '').trim(), (data.branch || '').trim());
      return respond({ currentStatus: status.type, lastPunchTime: status.time });
    }

    return respond({ error: 'Unknown action' });
  } catch (err) {
    return respond({ error: err.message });
  }
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── PIN verification ──────────────────────────────────────────
// Searches every "Branch - X" tab.
// Returns { success, name, branch } on match.

function verifyPin(pin) {
  const master   = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const branches = getBranchSheets(master);

  if (branches.length === 0) {
    return { success: false, error: 'No branch tabs found. Run setup() first.' };
  }

  for (const sheet of branches) {
    const branchName = sheet.getName().replace(BRANCH_TAB_PREFIX, '').trim();
    const data       = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const name   = String(data[i][0]).trim();
      const stored = String(data[i][1]).trim();
      if (name && stored && stored === String(pin).trim()) {
        return { success: true, name, branch: branchName };
      }
    }
  }

  return { success: false, error: 'PIN not recognised.' };
}

// ── Branch lookup by employee name ────────────────────────────
// Used by recordPunch when the client doesn't forward branch.

function getBranchForEmployee(employeeName) {
  const master   = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const branches = getBranchSheets(master);
  for (const sheet of branches) {
    const branchName = sheet.getName().replace(BRANCH_TAB_PREFIX, '').trim();
    const data       = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const name = String(data[i][0]).trim();
      if (name && name === employeeName) return branchName;
    }
  }
  return null;
}

// ── Record a punch ────────────────────────────────────────────
// branch is the plain branch name, e.g. "Downtown"

function recordPunch(name, branch, type, clientTime) {
  if (!name) throw new Error('recordPunch: employee name is empty.');
  if (!type) throw new Error('recordPunch: punch type (IN/OUT) is missing.');

  // Auto-resolve branch from master sheet if client didn't send it
  if (!branch) {
    branch = getBranchForEmployee(name);
    if (!branch) throw new Error('Employee "' + name + '" not found in any branch tab. Check the master sheet.');
  }

  // Use client-captured timestamp if available, otherwise fall back to server time
  var now;
  if (clientTime) {
    now = new Date(clientTime);
    if (isNaN(now.getTime())) now = new Date(); // fallback if parse fails
  } else {
    now = new Date();
  }
  const tz   = Session.getScriptTimeZone();
  const date = Utilities.formatDate(now, tz, 'MM/dd/yyyy');
  const time = Utilities.formatDate(now, tz, 'hh:mm:ss a');

  const monthly  = getOrCreateMonthlySheet(now, branch);
  const empTab   = getOrCreateEmployeeTab(monthly, name);
  const lastType = getLastPunchType(empTab, date);

  // Duplicate punch guard
  if (type.toUpperCase() === 'IN' && lastType === 'IN') {
    return { success: false, duplicate: true, type: 'in',
             message: name + ' is already clocked in today.' };
  }
  if (type.toUpperCase() === 'OUT' && lastType === 'OUT') {
    return { success: false, duplicate: true, type: 'out',
             message: name + ' has already clocked out today.' };
  }
  if (type.toUpperCase() === 'OUT' && !lastType) {
    return { success: false, duplicate: true, type: 'out',
             message: name + ' has not clocked in yet today. Please clock in first.' };
  }

  // Write to employee tab using daily-row format
  if (type.toUpperCase() === 'IN') {
    // Clock IN — create a new row for today
    empTab.appendRow([date, time, '', '', '', '', '']);
    // ARRAYFORMULA in D2 auto-calculates hours when clock-out is filled
  } else {
    // Clock OUT — find today's open IN row and fill in the clock-out time.
    // Formulas in D, E, F auto-calculate hours from B and C.
    const data    = empTab.getDataRange().getValues();
    let targetRow = -1;

    for (let i = data.length - 1; i >= 1; i--) {
      const rawDate   = data[i][0];
      const rowDate   = rawDate instanceof Date
        ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'MM/dd/yyyy')
        : String(rawDate).trim();
      const clockOut  = data[i][2];
      const noClockOut = (clockOut === '' || clockOut === null || clockOut === undefined || clockOut === 0);
      const clockIn   = data[i][1];
      const hasClockIn = (clockIn !== '' && clockIn !== null && clockIn !== undefined && clockIn !== 0);

      if (rowDate === date && noClockOut && hasClockIn) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow > 0) {
      // Write clock-out time — ARRAYFORMULA recalculates hours automatically
      empTab.getRange(targetRow, 3).setValue(time);
      Logger.log('recordPunch OUT: wrote clock-out at ' + time + ' for row ' + targetRow);
    } else {
      Logger.log('recordPunch OUT: no open IN row found for ' + date);
      empTab.appendRow([date, '', time, '', '', '', 'No matching clock-in']);
    }
  }

  updateSummary(monthly, name, type, date, time);
  return { success: true, name, branch, type, date, time };
}

// Returns the last punch type ('IN' or 'OUT') for an employee on a given date,
// or null if they have no punches today.
// In the new daily-row format:
//   col B (index 1) = Clock In time   (empty string = not set)
//   col C (index 2) = Clock Out time  (empty string = not set)
// Returns 'IN' if today has a row with Clock In but no Clock Out yet.
// Returns 'OUT' if today has a row with both times filled.
// Returns null if no row for today.
function getLastPunchType(sheet, date) {
  const tz   = Session.getScriptTimeZone();
  const data = sheet.getDataRange().getValues();
  let last   = null;
  for (let i = 1; i < data.length; i++) {
    const rawDate = data[i][0];
    const rowDate = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, tz, 'MM/dd/yyyy')
      : String(rawDate).trim();
    if (rowDate !== date) continue;
    const hasIn  = String(data[i][1]).trim() !== '';
    const hasOut = String(data[i][2]).trim() !== '';
    if (hasIn && !hasOut) last = 'IN';
    if (hasIn && hasOut)  last = 'OUT';
  }
  return last;
}

// Returns { type, time } - the employee's current clock status for today.
// type is 'IN', 'OUT', or null (no punches today).
function getCurrentStatus(name, branch) {
  try {
    const now      = new Date();
    const tz       = Session.getScriptTimeZone();
    const today    = Utilities.formatDate(now, tz, 'MM/dd/yyyy');
    const monthStr = Utilities.formatDate(now, tz, 'MMMM yyyy');
    const period   = getPeriod(now);
    const fileName = monthStr + ' ' + period + ' ' + branch;

    // Use the branch subfolder - same path recordPunch uses
    const branchFolder = getOrCreateBranchFolder(branch);
    const files        = branchFolder.getFilesByName(fileName);
    if (!files.hasNext()) {
      Logger.log('getCurrentStatus: monthly file not found - ' + fileName);
      return { type: null, time: null };
    }

    const monthly  = SpreadsheetApp.open(files.next());
    const empSheet = monthly.getSheetByName(name);
    if (!empSheet) {
      Logger.log('getCurrentStatus: no tab for "' + name + '" in ' + fileName);
      return { type: null, time: null };
    }

    const data = empSheet.getDataRange().getValues();
    let lastType = null;
    let lastTime = null;

    for (let i = 1; i < data.length; i++) {
      const rawDate = data[i][0];
      const rowDate = rawDate instanceof Date
        ? Utilities.formatDate(rawDate, tz, 'MM/dd/yyyy')
        : String(rawDate).trim();
      if (rowDate !== today) continue;

      const rawIn  = data[i][1];
      const rawOut = data[i][2];
      const hasIn  = rawIn  !== '' && rawIn  !== null && rawIn  !== undefined;
      const hasOut = rawOut !== '' && rawOut !== null && rawOut !== undefined;

      if (hasIn && !hasOut) {
        lastType = 'IN';
        lastTime = rawIn instanceof Date
          ? Utilities.formatDate(rawIn, tz, 'hh:mm:ss a')
          : String(rawIn).trim();
      } else if (hasIn && hasOut) {
        lastType = 'OUT';
        lastTime = rawOut instanceof Date
          ? Utilities.formatDate(rawOut, tz, 'hh:mm:ss a')
          : String(rawOut).trim();
      }
    }

    Logger.log('getCurrentStatus: ' + name + ' → ' + lastType + ' at ' + lastTime);
    return { type: lastType, time: lastTime };

  } catch(e) {
    Logger.log('getCurrentStatus error: ' + e.message);
    return { type: null, time: null };
  }
}

// ── Branch helpers ────────────────────────────────────────────

// Non-branch tabs to exclude when scanning for branch tabs
const EXCLUDED_TABS = ['Summary', 'Employees', 'Sheet1', 'Locations'];

// Returns all branch sheets - any tab not in the excluded list
function getBranchSheets(ss) {
  return ss.getSheets().filter(s => !EXCLUDED_TABS.includes(s.getName()));
}

// Returns an array of plain branch names, e.g. ["Buena Park", "Cupertino"]
function getBranchNames(ss) {
  return getBranchSheets(ss).map(s => s.getName().trim());
}

// ── Monthly spreadsheet (branch-aware) ───────────────────────
// File name:  "June 2026 - Downtown"
// Location:   <Root>/Downtown/

// Returns 'P1' for days 1-15, 'P2' for days 16-end of month.
function getPeriod(date) {
  return date.getDate() <= 15 ? 'P1' : 'P2';
}

// Returns the period label for display, e.g. "Jun 1-15" or "Jun 16-30"
function getPeriodLabel(date) {
  const tz      = Session.getScriptTimeZone();
  const month   = Utilities.formatDate(date, tz, 'MMM');
  const year    = Utilities.formatDate(date, tz, 'yyyy');
  const lastDay = new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
  return date.getDate() <= 15
    ? month + ' 1-15 ' + year
    : month + ' 16-' + lastDay + ' ' + year;
}

function getOrCreateMonthlySheet(date, branch) {
  const tz       = Session.getScriptTimeZone();
  const monthStr = Utilities.formatDate(date, tz, 'MMMM yyyy');
  const period   = getPeriod(date);
  const fileName = monthStr + ' ' + period + ' ' + branch; // e.g. "June 2026 P1 Buena Park"

  const branchFolder = getOrCreateBranchFolder(branch);
  const files        = branchFolder.getFilesByName(fileName);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());

  // Create new semi-monthly spreadsheet
  const ss   = SpreadsheetApp.create(fileName);
  const file = DriveApp.getFileById(ss.getId());
  branchFolder.addFile(file);
  try { DriveApp.getRootFolder().removeFile(file); } catch(e) {}

  const defaultSheet = ss.getSheets()[0];
  const periodLabel  = getPeriodLabel(date);

  // Summary tab
  const summary = ss.insertSheet(SUMMARY_SHEET, 0);
  summary.appendRow(['Employee', 'Last Action', 'Date', 'Time', 'Regular Hrs', 'OT Hrs', 'Total Hrs']);
  summary.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
  summary.setFrozenRows(1);
  summary.setColumnWidth(1, 180);
  summary.setColumnWidth(2, 140);
  summary.setColumnWidth(3, 110);
  summary.setColumnWidth(4, 140);
  summary.setColumnWidth(5, 110);
  summary.setColumnWidth(6, 110);
  summary.setColumnWidth(7, 110);

  // Period label in top right for clarity
  summary.getRange(1, 9).setValue('Period: ' + periodLabel).setFontWeight('bold');

  ss.deleteSheet(defaultSheet);
  Logger.log('Created semi-monthly sheet: ' + fileName + ' in folder: ' + branch);
  return ss;
}

// ── Drive folder helpers ──────────────────────────────────────

// Returns (or creates) the root folder that holds all branch sub-folders.
function getRootFolder() {
  // Option 1: explicit root folder ID set by the user
  if (DRIVE_ROOT_FOLDER_ID && DRIVE_ROOT_FOLDER_ID !== '') {
    return DriveApp.getFolderById(DRIVE_ROOT_FOLDER_ID);
  }
  // Option 2: same folder as the master spreadsheet (requires a real ID)
  const validId = MASTER_SPREADSHEET_ID &&
                  MASTER_SPREADSHEET_ID !== 'PASTE_YOUR_SPREADSHEET_ID_HERE';
  if (validId) {
    try {
      const parents = DriveApp.getFileById(MASTER_SPREADSHEET_ID).getParents();
      if (parents.hasNext()) return parents.next();
    } catch(e) {
      Logger.log('Could not resolve master spreadsheet folder: ' + e.message);
    }
  }
  // Option 3: fall back to My Drive root
  return DriveApp.getRootFolder();
}

// Returns (or creates) <root>/<branchName>/
function getOrCreateBranchFolder(branchName) {
  if (!branchName) throw new Error('getOrCreateBranchFolder: branchName is empty or null.');
  const root    = getRootFolder();
  const folders = root.getFoldersByName(branchName);
  if (folders.hasNext()) return folders.next();

  const newFolder = root.createFolder(branchName);
  Logger.log('Created branch folder: ' + branchName);
  return newFolder;
}

// ── Per-employee tab ──────────────────────────────────────────

function getOrCreateEmployeeTab(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    // Col A=Date, B=Clock In, C=Clock Out, D=Regular Hrs, E=OT Hrs, F=Total Hrs, G=Notes
    sheet.appendRow(['Date', 'Clock In', 'Clock Out', 'Regular Hrs', 'OT Hrs', 'Total Hrs', 'Notes']);
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 110);
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 110);
    sheet.setColumnWidth(5, 90);
    sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 180);

    // ARRAYFORMULA in row 2 — auto-calculates hours for ALL rows, including manually added ones.
    // Regular Hrs: capped at 8
    sheet.getRange(2, 4).setFormula(
      '=ARRAYFORMULA(IF((B2:B<>"")*(C2:C<>""),' +
      'ROUND(IF((TIMEVALUE(C2:C)-TIMEVALUE(B2:B))*24>8,8,(TIMEVALUE(C2:C)-TIMEVALUE(B2:B))*24),1),""))'
    );
    // OT Hrs: anything above 8
    sheet.getRange(2, 5).setFormula(
      '=ARRAYFORMULA(IF((B2:B<>"")*(C2:C<>""),' +
      'ROUND(IF((TIMEVALUE(C2:C)-TIMEVALUE(B2:B))*24>8,(TIMEVALUE(C2:C)-TIMEVALUE(B2:B))*24-8,0),1),""))'
    );
    // Total Hrs
    sheet.getRange(2, 6).setFormula(
      '=ARRAYFORMULA(IF((B2:B<>"")*(C2:C<>""),' +
      'ROUND((TIMEVALUE(C2:C)-TIMEVALUE(B2:B))*24,1),""))'
    );
  }
  return sheet;
}

// ── Total hours calculation ───────────────────────────────────

// Returns { reg, ot, total } hours summed from the employee tab for the period.
function calcTotalHours(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return { reg: 0, ot: 0, total: 0 };

  const data = sheet.getDataRange().getValues();
  let reg = 0, ot = 0;

  for (let i = 1; i < data.length; i++) {
    const regVal = parseFloat(data[i][3]);
    const otVal  = parseFloat(data[i][4]);
    if (!isNaN(regVal)) reg += regVal;
    if (!isNaN(otVal))  ot  += otVal;
  }

  return {
    reg:   Math.round(reg   * 10) / 10,
    ot:    Math.round(ot    * 10) / 10,
    total: Math.round((reg + ot) * 10) / 10
  };
}

// ── Summary tab ───────────────────────────────────────────────

function updateSummary(ss, name, type, date, time) {
  let summary = ss.getSheetByName(SUMMARY_SHEET);
  if (!summary) {
    summary = ss.insertSheet(SUMMARY_SHEET, 0);
    summary.appendRow(['Employee', 'Last Action', 'Date', 'Time', 'Regular Hrs', 'OT Hrs', 'Total Hrs']);
    summary.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    summary.setFrozenRows(1);
    summary.setColumnWidth(1, 180);
    summary.setColumnWidth(2, 140);
    summary.setColumnWidth(3, 110);
    summary.setColumnWidth(4, 140);
    summary.setColumnWidth(5, 110);
    summary.setColumnWidth(6, 110);
    summary.setColumnWidth(7, 110);
  }

  const rows     = summary.getDataRange().getValues();
  const status   = type.toLowerCase() === 'in' ? 'CLOCKED IN' : 'CLOCKED OUT';
  const isMissed = time === 'MISSED CLOCK-OUT';
  const hrs      = calcTotalHours(ss, name);

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === name) {
      summary.getRange(i + 1, 2, 1, 6).setValues([[status, date, time, hrs.reg, hrs.ot, hrs.total]]);
      colorRow(summary, i + 1, type, isMissed);
      // Highlight OT cell if non-zero
      return;
    }
  }
  summary.appendRow([name, status, date, time, hrs.reg, hrs.ot, hrs.total]);
  colorRow(summary, summary.getLastRow(), type, isMissed);
}

function colorRow(sheet, row, type, isMissed) {
  const cell = sheet.getRange(row, 2);
  if (type.toLowerCase() === 'in') cell.setBackground('#d9ead3').setFontColor('#38761d');
  else                              cell.setBackground('#f4cccc').setFontColor('#cc0000');
}

// ── Midnight auto-check ───────────────────────────────────────
// Runs for every branch's monthly file.

// ── Google Form for clock-out correction ──────────────────────
// Run createCorrectionForm() once from the editor.
// It creates a Google Form and stores its ID and field entry IDs
// in Script Properties so the email can build pre-filled URLs.

function createCorrectionForm() {
  const form = FormApp.create(FORM_TITLE);
  form.setDescription('Submit your actual clock-out time for a missed day. Your PIN is required for verification.');
  form.setCollectEmail(false);
  form.setConfirmationMessage('Thank you. Your timesheet has been updated.');

  // Add fields — order matters for pre-fill URL
  const nameItem   = form.addTextItem().setTitle('Employee Name').setRequired(true);
  const dateItem   = form.addTextItem().setTitle('Date (MM/DD/YYYY)').setRequired(true);
  const branchItem = form.addTextItem().setTitle('Branch').setRequired(true);
  const pinItem    = form.addTextItem().setTitle('6-Digit PIN').setRequired(true);
  const timeItem   = form.addTextItem().setTitle('Actual Clock-Out Time (e.g. 5:30 PM)').setRequired(true);

  // Store IDs in Script Properties for building pre-fill URLs
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    'CORRECTION_FORM_ID':     form.getId(),
    'CORRECTION_FORM_URL':    form.getPublishedUrl(),
    'ENTRY_NAME':             String(nameItem.getId()),
    'ENTRY_DATE':             String(dateItem.getId()),
    'ENTRY_BRANCH':           String(branchItem.getId()),
  });

  // Install form submit trigger
  ScriptApp.newTrigger('onCorrectionFormSubmit')
    .forForm(form)
    .onFormSubmit()
    .create();

  // Move form to same folder as master sheet
  try {
    const formFile = DriveApp.getFileById(form.getId());
    const root = getRootFolder();
    root.addFile(formFile);
    DriveApp.getRootFolder().removeFile(formFile);
  } catch(e) {}

  Logger.log('Correction form created:');
  Logger.log('  URL:  ' + form.getPublishedUrl());
  Logger.log('  Edit: ' + form.getEditUrl());
  Logger.log('  ID:   ' + form.getId());
  Logger.log('Form submit trigger installed.');
}

// Builds a pre-filled Google Form URL for a specific employee and date.
function buildCorrectionFormUrl(name, date, branch) {
  const props   = PropertiesService.getScriptProperties();
  const formUrl = props.getProperty('CORRECTION_FORM_URL');
  const eName   = props.getProperty('ENTRY_NAME');
  const eDate   = props.getProperty('ENTRY_DATE');
  const eBranch = props.getProperty('ENTRY_BRANCH');

  if (!formUrl || !eName || !eDate || !eBranch) {
    Logger.log('buildCorrectionFormUrl: form not set up. Run createCorrectionForm() first.');
    return '';
  }

  return formUrl +
    '?entry.' + eName   + '=' + encodeURIComponent(name) +
    '&entry.' + eDate   + '=' + encodeURIComponent(date) +
    '&entry.' + eBranch + '=' + encodeURIComponent(branch);
}

// ── Form submission handler ──────────────────────────────────
// Triggered automatically when an employee submits the correction form.

function onCorrectionFormSubmit(e) {
  const responses = e.response.getItemResponses();
  const vals = {};
  responses.forEach(r => { vals[r.getItem().getTitle()] = r.getResponse(); });

  const name         = (vals['Employee Name'] || '').trim();
  const date         = (vals['Date (MM/DD/YYYY)'] || '').trim();
  const branch       = (vals['Branch'] || '').trim();
  const pin          = (vals['6-Digit PIN'] || '').trim();
  const rawTime      = (vals['Actual Clock-Out Time (e.g. 5:30 PM)'] || '').trim();

  Logger.log('Correction form submitted: name=' + name + ' date=' + date + ' branch=' + branch + ' time=' + rawTime);

  // Verify PIN
  const verify = verifyPin(pin);
  if (!verify.success || verify.name !== name) {
    Logger.log('Correction rejected: PIN mismatch for ' + name);
    return;
  }

  // Parse the clock-out time — normalize to HH:MM:SS AM/PM
  const clockOutTime = normalizeTime(rawTime);
  if (!clockOutTime) {
    Logger.log('Correction rejected: could not parse time "' + rawTime + '"');
    return;
  }

  // Find the monthly file
  const tz        = Session.getScriptTimeZone();
  const dateParts = date.split('/');
  if (dateParts.length !== 3) {
    Logger.log('Correction rejected: bad date format "' + date + '"');
    return;
  }
  const dateObj  = new Date(parseInt(dateParts[2]), parseInt(dateParts[0]) - 1, parseInt(dateParts[1]));
  const monthly  = getOrCreateMonthlySheet(dateObj, branch);
  const empSheet = monthly.getSheetByName(name);

  if (!empSheet) {
    Logger.log('Correction rejected: no tab for ' + name + ' in monthly file');
    return;
  }

  // Find the row for that date
  const sheetData = empSheet.getDataRange().getValues();
  let targetRow   = -1;

  for (let i = sheetData.length - 1; i >= 1; i--) {
    const rawDate = sheetData[i][0];
    const rowDate = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, tz, 'MM/dd/yyyy')
      : String(rawDate).trim();
    if (rowDate !== date) continue;

    const hasIn      = String(sheetData[i][1]).trim() !== '';
    const currentOut = String(sheetData[i][2]).trim();
    const rowNotes   = String(sheetData[i][6]).trim();
    const isMissed   = currentOut === '11:59:00 PM' || rowNotes === 'MISSED CLOCK-OUT';

    if (hasIn && (currentOut === '' || currentOut === '0' || isMissed)) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow < 0) {
    Logger.log('Correction rejected: no open/missed row for ' + name + ' on ' + date);
    return;
  }

  // Write corrected clock-out time — ARRAYFORMULA auto-calculates hours
  empSheet.getRange(targetRow, 3).setValue(clockOutTime);
  empSheet.getRange(targetRow, 7).setValue('Corrected by employee');

  SpreadsheetApp.flush(); // force formula evaluation
  const otVal = empSheet.getRange(targetRow, 5).getValue();
  if (otVal > 0) {
    empSheet.getRange(targetRow, 7).setValue('Corrected by employee | ' + Math.round(otVal * 10) / 10 + ' hrs OT');
  }

  const regRounded = empSheet.getRange(targetRow, 4).getValue();
  const otRounded  = otVal;

  updateSummary(monthly, name, 'out', date, clockOutTime);
  Logger.log('Correction applied: ' + name + ' on ' + date + ' at ' + clockOutTime + ' (reg=' + regRounded + ', ot=' + otRounded + ')');
}

// Normalize user-entered time like "5:30 pm", "17:30", "530pm" to "05:30:00 PM"
function normalizeTime(raw) {
  if (!raw) return null;
  var s = raw.trim().toUpperCase();

  // Try "H:MM PM" or "HH:MM PM" or "H:MM:SS PM"
  var match = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)$/);
  if (match) {
    var h = String(parseInt(match[1])).padStart(2, '0');
    var m = match[2];
    var sec = match[3] || '00';
    return h + ':' + m + ':' + sec + ' ' + match[4];
  }

  // Try 24-hour "17:30" or "17:30:00"
  match = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (match) {
    var hr = parseInt(match[1]);
    var ampm = hr >= 12 ? 'PM' : 'AM';
    if (hr > 12) hr -= 12;
    if (hr === 0) hr = 12;
    return String(hr).padStart(2, '0') + ':' + match[2] + ':' + (match[3] || '00') + ' ' + ampm;
  }

  // Try "530pm" or "530 pm"
  match = s.match(/^(\d{1,4})\s*(AM|PM)$/);
  if (match) {
    var num = match[1];
    var mins = num.length <= 2 ? '00' : num.slice(-2);
    var hrs  = num.length <= 2 ? num : num.slice(0, -2);
    return String(parseInt(hrs)).padStart(2, '0') + ':' + mins + ':00 ' + match[2];
  }

  return null; // could not parse
}

// ── Email reminder for employees still clocked in ─────────────
// Runs at 11:50 PM — gives employees ~9 minutes to clock out
// before the auto-logout at 11:55 PM.
// Reads email from col D of each branch tab: Name (A) | PIN (B) | Token (C) | Email (D)

function sendClockOutReminders() {
  const now       = new Date();
  const tz        = Session.getScriptTimeZone();
  const checkDate = Utilities.formatDate(now, tz, 'MM/dd/yyyy');
  const dateNice  = Utilities.formatDate(now, tz, 'EEEE, MMMM d, yyyy');

  const master   = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const branches = getBranchNames(master);
  let sent = 0;

  // Build a map of employee name → email from all branch tabs
  const emailMap = {};
  getBranchSheets(master).forEach(sheet => {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const name  = String(data[i][0]).trim();
      const email = String(data[i][2]).trim(); // col C = Email
      if (name && email && email.indexOf('@') !== -1) {
        emailMap[name] = email;
      }
    }
  });

  branches.forEach(branch => {
    try {
      const monthly = getOrCreateMonthlySheet(now, branch);

      monthly.getSheets().forEach(sheet => {
        const name = sheet.getName();
        if (name === SUMMARY_SHEET) return;

        const data = sheet.getDataRange().getValues();

        for (let i = 1; i < data.length; i++) {
          const rawDate = data[i][0];
          const rowDate = rawDate instanceof Date
            ? Utilities.formatDate(rawDate, tz, 'MM/dd/yyyy')
            : String(rawDate).trim();
          if (rowDate !== checkDate) continue;

          const hasIn  = String(data[i][1]).trim() !== '';
          const hasOut = String(data[i][2]).trim() !== '';

          // Still clocked in — send reminder
          if (hasIn && !hasOut && emailMap[name]) {
            const clockInTime = String(data[i][1]).trim();
            try {
              MailApp.sendEmail({
                to: emailMap[name],
                subject: 'Clock-Out Reminder - ' + dateNice,
                htmlBody:
                  '<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;">' +
                  '<h2 style="color:#0f172a;">Clock-Out Reminder</h2>' +
                  '<p>Hi ' + name + ',</p>' +
                  '<p>You clocked in at <strong>' + clockInTime + '</strong> today (' + dateNice + ') ' +
                  'at the <strong>' + branch + '</strong> location but have not clocked out yet.</p>' +
                  '<p>Please clock out before midnight. If you miss the deadline, ' +
                  'you can submit your actual clock-out time using the form below:</p>' +
                  '<p style="margin:16px 0;">' +
                  '<a href="' + buildCorrectionFormUrl(name, checkDate, branch) +
                  '" style="display:inline-block;padding:12px 24px;background:#0f172a;color:#fff;' +
                  'text-decoration:none;border-radius:8px;font-weight:600;font-size:14px;">' +
                  'Submit Clock-Out Time</a></p>' +
                  '<p style="color:#9a9ea6;font-size:12px;margin-top:24px;">- TimeClock System</p>' +
                  '</div>'
              });
              sent++;
              Logger.log('Reminder sent to ' + name + ' (' + emailMap[name] + ') — ' + branch);
            } catch(emailErr) {
              Logger.log('Failed to send email to ' + name + ': ' + emailErr.message);
            }
          }
        }
      });
    } catch(e) {
      Logger.log('sendClockOutReminders error for branch ' + branch + ': ' + e.message);
    }
  });

  Logger.log('Clock-out reminders sent: ' + sent);
}

function checkMissedClockOuts() {
  // This runs at 11:55 PM - check today's date
  const now       = new Date();
  const tz        = Session.getScriptTimeZone();
  const checkDate = Utilities.formatDate(now, tz, 'MM/dd/yyyy');

  const master   = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const branches = getBranchNames(master);

  branches.forEach(branch => {
    const monthly = getOrCreateMonthlySheet(now, branch);

    monthly.getSheets().forEach(sheet => {
      const name = sheet.getName();
      if (name === SUMMARY_SHEET) return;

      const data = sheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        const rawDate = data[i][0];
        const rowDate = rawDate instanceof Date
          ? Utilities.formatDate(rawDate, tz, 'MM/dd/yyyy')
          : String(rawDate).trim();
        if (rowDate !== checkDate) continue;

        const hasIn  = String(data[i][1]).trim() !== '';
        const hasOut = String(data[i][2]).trim() !== '';

        // Row has clock-in but no clock-out — missed
        if (hasIn && !hasOut) {
          // Fill in 11:59 PM as clock-out — formulas auto-calculate hours
          sheet.getRange(i + 1, 3).setValue('11:59:00 PM');
          // ARRAYFORMULA auto-recalculates hours
          sheet.getRange(i + 1, 7).setValue('MISSED CLOCK-OUT');
          updateSummary(monthly, name, 'out', checkDate, 'MISSED CLOCK-OUT');
          Logger.log('Missed clock-out logged for ' + name + ' (' + branch + ') on ' + checkDate);
        }
      }
    });
  });
}

// ── New-month trigger ─────────────────────────────────────────
// Pre-creates monthly files for every branch on the 1st.

function createNewMonthSheet() {
  const now    = new Date();
  const master = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  // On the 1st: pre-create P1. On the 16th trigger, pre-create P2.
  // We call both so the trigger on the 1st pre-creates P1,
  // and a separate trigger on the 16th pre-creates P2.
  getBranchNames(master).forEach(branch => getOrCreateMonthlySheet(now, branch));
}

// Pre-create P2 files on the 16th of each month
function createP2Sheet() {
  const now    = new Date();
  // Force day to 16 so getPeriod returns P2
  const p2Date = new Date(now.getFullYear(), now.getMonth(), 16);
  const master = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  getBranchNames(master).forEach(branch => getOrCreateMonthlySheet(p2Date, branch));
}

// ── One-time setup ────────────────────────────────────────────
// Creates sample branch tabs if none exist yet.
// Edit the SAMPLE_BRANCHES list (or add tabs manually in the sheet).

function setup() {
  const SAMPLE_BRANCHES = ['Downtown', 'Cupertino']; // ← edit as needed

  const master  = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const existing = getBranchNames(master);

  SAMPLE_BRANCHES.forEach(branch => {
    const tabName = BRANCH_TAB_PREFIX + branch;
    if (!master.getSheetByName(tabName)) {
      const sheet = master.insertSheet(tabName);
      sheet.appendRow(['Name', 'PIN (6 digits)']);
      sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 180);
      sheet.setColumnWidth(2, 140);
      sheet.getRange(2, 2, 100, 1).setNumberFormat('@STRING@');

      // Sample employees - replace with real data
      sheet.appendRow(['Alice Johnson', '123456']);
      sheet.appendRow(['Bob Martinez',  '234567']);

      Logger.log('Created branch tab: ' + tabName);
    } else {
      Logger.log('Branch tab already exists: ' + tabName);
    }

    // Also pre-create the Drive folder
    getOrCreateBranchFolder(branch);
  });

  // Create Locations tab if missing
  if (!master.getSheetByName(LOCATIONS_SHEET)) {
    const locSheet = master.insertSheet(LOCATIONS_SHEET);
    locSheet.appendRow(['Branch', 'Latitude', 'Longitude', 'Radius (meters)']);
    locSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    locSheet.setFrozenRows(1);
    locSheet.setColumnWidth(1, 180);
    locSheet.setColumnWidth(2, 130);
    locSheet.setColumnWidth(3, 130);
    locSheet.setColumnWidth(4, 140);
    // Sample row — edit with real coordinates
    locSheet.appendRow(['Buena Park', 34.04819, -118.26051, 200]);
    locSheet.appendRow(['Cupertino', 37.32299, -122.03218, 200]);
    locSheet.appendRow(['Beverly Center', 34.0754, -118.3772, 200]);
    
    Logger.log('Created Locations tab with sample data');
  }

  Logger.log('Setup complete. Branch tabs and Drive folders are ready.');
  Logger.log('Add employees to each "Branch - X" tab, then deploy the Web App.');
}

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (['checkMissedClockOuts', 'createNewMonthSheet', 'createP2Sheet', 'sendClockOutReminders'].includes(fn)) {
      ScriptApp.deleteTrigger(t);
    }
  });
  // 11:50 PM daily — email reminder to employees still clocked in
  ScriptApp.newTrigger('sendClockOutReminders')
    .timeBased().everyDays(1).atHour(23).nearMinute(50).create();
  // 11:55 PM daily — auto-log missed clock-outs
  ScriptApp.newTrigger('checkMissedClockOuts')
    .timeBased().everyDays(1).atHour(23).nearMinute(55).create();
  // 1st of month — pre-create P1 files
  ScriptApp.newTrigger('createNewMonthSheet')
    .timeBased().onMonthDay(1).atHour(0).create();
  // 16th of month — pre-create P2 files
  ScriptApp.newTrigger('createP2Sheet')
    .timeBased().onMonthDay(16).atHour(0).create();
  Logger.log('Triggers installed: email reminder (11:50 PM), missed clock-out (11:55 PM), P1 (1st), P2 (16th).');
}

// ── DEBUG - run this from the Apps Script editor ─────────────
// Change NAME and BRANCH to match a real employee you've punched.
function debugStatus() {
  const NAME   = 'Alice Johnson';  // ← change to actual employee name in your sheet
  const BRANCH = 'Buena Park';     // ← change to actual branch name

  const tz       = Session.getScriptTimeZone();
  const now      = new Date();
  const today    = Utilities.formatDate(now, tz, 'MM/dd/yyyy');
  const monthStr = Utilities.formatDate(now, tz, 'MMMM yyyy');

  Logger.log('=== debugStatus ===');
  Logger.log('Script timezone : ' + tz);
  Logger.log('Today           : ' + today);
  Logger.log('Month string    : ' + monthStr);

  // 1. Check the branch folder
  const root = getRootFolder();
  Logger.log('Root folder     : ' + root.getName() + ' (' + root.getId() + ')');

  const folderIter = root.getFoldersByName(BRANCH);
  if (!folderIter.hasNext()) {
    Logger.log('ERROR: Branch folder "' + BRANCH + '" not found inside root folder.');
    return;
  }
  const branchFolder = folderIter.next();
  Logger.log('Branch folder   : ' + branchFolder.getName() + ' (' + branchFolder.getId() + ')');

  // 2. List all files in the branch folder
  Logger.log('--- Files in branch folder ---');
  const allFiles = branchFolder.getFiles();
  let fileCount = 0;
  while (allFiles.hasNext()) {
    const f = allFiles.next();
    const fname = f.getName();
    // Print char codes for the dash area to catch encoding issues
    const dashIdx = fname.indexOf(monthStr) + monthStr.length;
    const dashArea = fname.substring(dashIdx, dashIdx + 4);
    const dashCodes = dashArea.split('').map(c => 'U+' + c.charCodeAt(0).toString(16).toUpperCase().padStart(4,'0')).join(' ');
    Logger.log('  [' + fileCount + '] "' + fname + '"  dash chars: ' + dashCodes);
    fileCount++;
  }
  if (fileCount === 0) Logger.log('  (folder is empty)');

  // 3. Build the expected file name and show its dash character codes
  const fileName = monthStr + ' ' + BRANCH;
  const dashIdx2  = fileName.indexOf(monthStr) + monthStr.length;
  const dashArea2 = fileName.substring(dashIdx2, dashIdx2 + 4);
  const dashCodes2 = dashArea2.split('').map(c => 'U+' + c.charCodeAt(0).toString(16).toUpperCase().padStart(4,'0')).join(' ');
  Logger.log('--- Expected file name ---');
  Logger.log('  "' + fileName + '"  dash chars: ' + dashCodes2);

  // 4. Try to find it
  const files = branchFolder.getFilesByName(fileName);
  if (!files.hasNext()) {
    Logger.log('ERROR: File not found by name. Check dash characters above - they must match exactly.');
    return;
  }
  const monthly = SpreadsheetApp.open(files.next());
  Logger.log('Found file      : ' + monthly.getName());

  // 5. List all tabs
  Logger.log('--- Tabs in monthly file ---');
  monthly.getSheets().forEach((s, i) => Logger.log('  [' + i + '] "' + s.getName() + '"'));

  // 6. Find employee tab
  const empSheet = monthly.getSheetByName(NAME);
  if (!empSheet) {
    Logger.log('ERROR: No tab named "' + NAME + '". Check spelling/spaces above.');
    return;
  }

  // 7. Dump today's rows
  Logger.log('--- All rows in "' + NAME + '" tab (raw) ---');
  const data = empSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const raw0 = data[i][0];
    const raw1 = data[i][1];
    const raw2 = data[i][2];
    Logger.log('  row ' + i + ': col0 type=' + typeof raw0 + ' val="' + raw0 + '" | col1="' + raw1 + '" | col2="' + raw2 + '"');
  }
  Logger.log('=== end debugStatus ===');
}

// ── DEBUG 2 - lists everything in the root folder ────────────
function debugRootContents() {
  const root = getRootFolder();
  Logger.log('Root folder: ' + root.getName() + ' (' + root.getId() + ')');

  Logger.log('--- Subfolders ---');
  const folders = root.getFolders();
  let fc = 0;
  while (folders.hasNext()) {
    const f = folders.next();
    Logger.log('  folder: "' + f.getName() + '"');
    fc++;
  }
  if (fc === 0) Logger.log('  (none)');

  Logger.log('--- Files ---');
  const files = root.getFiles();
  let fileCount = 0;
  while (files.hasNext()) {
    Logger.log('  file: "' + files.next().getName() + '"');
    fileCount++;
  }
  if (fileCount === 0) Logger.log('  (none)');

  Logger.log('--- Master sheet branch tabs ---');
  const master = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  getBranchSheets(master).forEach(s => Logger.log('  tab: "' + s.getName() + '" → branch name: "' + s.getName().replace(BRANCH_TAB_PREFIX, '').trim() + '"'));
}

// ── One-time migration ────────────────────────────────────────
// Run once to rename existing monthly files from em dash to hyphen
// e.g. "June 2026 Buena Park" → "June 2026 - Buena Park"
function migrateFileNames() {
  const master   = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const branches = getBranchNames(master);
  const root     = getRootFolder();

  branches.forEach(branch => {
    const branchFolders = root.getFoldersByName(branch);
    if (!branchFolders.hasNext()) {
      Logger.log('No folder for branch: ' + branch);
      return;
    }
    const folder = branchFolders.next();
    const files  = folder.getFiles();
    while (files.hasNext()) {
      const file    = files.next();
      const oldName = file.getName();
      // Replace em dash (U+2014) with ' - '
      const newName = oldName.replace(/-/g, '-').replace(/  /g, ' ');
      if (oldName !== newName) {
        file.setName(newName);
        Logger.log('Renamed: "' + oldName + '" → "' + newName + '"');
      } else {
        Logger.log('No change needed: "' + oldName + '"');
      }
    }
  });
  Logger.log('Migration complete.');
}

// ── Recalculate Summary ───────────────────────────────────────
// Run this from the editor to refresh the Summary tab for every branch's
// current period file. Reads the formula-calculated values from each
// employee tab and updates the Summary.

function recalculateSummary() {
  const now    = new Date();
  const master = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  const branches = getBranchNames(master);

  branches.forEach(function(branch) {
    try {
      var monthly = getOrCreateMonthlySheet(now, branch);
      var summary = monthly.getSheetByName(SUMMARY_SHEET);
      if (!summary) return;

      // Clear existing summary rows (keep header)
      var lastRow = summary.getLastRow();
      if (lastRow > 1) {
        summary.getRange(2, 1, lastRow - 1, 7).clearContent();
        summary.getRange(2, 1, lastRow - 1, 7).clearFormat();
      }

      var summaryRow = 2;
      monthly.getSheets().forEach(function(sheet) {
        var empName = sheet.getName();
        if (empName === SUMMARY_SHEET) return;

        var hrs = calcTotalHours(monthly, empName);

        // Find last punch info
        var data = sheet.getDataRange().getValues();
        var lastAction = '';
        var lastDate   = '';
        var lastTime   = '';

        for (var i = data.length - 1; i >= 1; i--) {
          var hasIn  = String(data[i][1]).trim() !== '';
          var hasOut = String(data[i][2]).trim() !== '';
          if (hasIn && hasOut) {
            lastAction = 'CLOCKED OUT';
            lastDate   = data[i][0] instanceof Date
              ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), 'MM/dd/yyyy')
              : String(data[i][0]).trim();
            lastTime = String(data[i][2]).trim();
            break;
          } else if (hasIn && !hasOut) {
            lastAction = 'CLOCKED IN';
            lastDate   = data[i][0] instanceof Date
              ? Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), 'MM/dd/yyyy')
              : String(data[i][0]).trim();
            lastTime = String(data[i][1]).trim();
            break;
          }
        }

        summary.getRange(summaryRow, 1, 1, 7).setValues([[
          empName, lastAction, lastDate, lastTime, hrs.reg, hrs.ot, hrs.total
        ]]);

        // Color the status cell
        var statusCell = summary.getRange(summaryRow, 2);
        if (lastAction === 'CLOCKED IN') {
          statusCell.setBackground('#d9ead3').setFontColor('#38761d');
        } else if (lastAction === 'CLOCKED OUT') {
          statusCell.setBackground('#f4cccc').setFontColor('#cc0000');
        }
        if (hrs.ot > 0) {
        }

        summaryRow++;
      });

      Logger.log('Recalculated summary for ' + branch + ' — ' + (summaryRow - 2) + ' employees');
    } catch(e) {
      Logger.log('recalculateSummary error for ' + branch + ': ' + e.message);
    }
  });

  Logger.log('All summaries recalculated.');
}

// ── Automated Test Suite ──────────────────────────────────────
// Run runFullTest() from the Apps Script editor.
// It will:
//   1. Create test branch tabs in the master sheet (prefixed TEST-)
//   2. Create Drive folders and semi-monthly files
//   3. Simulate clock-in and clock-out for two test employees
//   4. Simulate an overtime scenario (9 hrs)
//   5. Simulate a missed clock-out
//   6. Print a full pass/fail report to the execution log
//   7. Clean up all test data when done
//
// Your real data is never touched.

const TEST_BRANCH      = 'TEST Branch';
const TEST_EMPLOYEE_1  = 'TEST Alice';
const TEST_EMPLOYEE_2  = 'TEST Bob';
const TEST_PIN_1       = '999001';
const TEST_PIN_2       = '999002';
const TEST_EMAIL_1     = 'test-alice@example.com';
const TEST_EMAIL_2     = 'test-bob@example.com';

function runFullTest() {
  Logger.log('');
  Logger.log('╔══════════════════════════════════════╗');
  Logger.log('║       TIMECLOCK FULL TEST SUITE      ║');
  Logger.log('╚══════════════════════════════════════╝');

  const results = [];
  let passed = 0;
  let failed = 0;

  function assert(label, condition, detail) {
    if (condition) {
      Logger.log('  ✓ PASS — ' + label);
      results.push({ label, pass: true });
      passed++;
    } else {
      Logger.log('  ✗ FAIL — ' + label + (detail ? ' | ' + detail : ''));
      results.push({ label, pass: false, detail });
      failed++;
    }
  }

  // ── STEP 1: Setup test branch in master sheet ──────────────
  Logger.log('');
  Logger.log('── Step 1: Create test branch tab ──');
  try {
    const master   = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    let testSheet  = master.getSheetByName(TEST_BRANCH);
    if (testSheet) {
      master.deleteSheet(testSheet);
      Logger.log('  Removed existing test tab');
    }
    testSheet = master.insertSheet(TEST_BRANCH);
    testSheet.appendRow(['Name', 'PIN', 'Email']);
    testSheet.appendRow([TEST_EMPLOYEE_1, TEST_PIN_1, TEST_EMAIL_1]);
    testSheet.appendRow([TEST_EMPLOYEE_2, TEST_PIN_2, TEST_EMAIL_2]);
    assert('Test branch tab created', master.getSheetByName(TEST_BRANCH) !== null);
    assert('Employee 1 row exists', testSheet.getLastRow() >= 2);
  } catch(e) {
    assert('Test branch tab created', false, e.message);
  }

  // ── STEP 2: PIN verification ───────────────────────────────
  Logger.log('');
  Logger.log('── Step 2: PIN verification ──');
  try {
    const r1 = verifyPin(TEST_PIN_1);
    assert('PIN 1 resolves correct name',  r1.success && r1.name === TEST_EMPLOYEE_1, JSON.stringify(r1));
    assert('PIN 1 resolves correct branch', r1.branch === TEST_BRANCH, 'got: ' + r1.branch);

    const r2 = verifyPin(TEST_PIN_2);
    assert('PIN 2 resolves correct name', r2.success && r2.name === TEST_EMPLOYEE_2, JSON.stringify(r2));

    const rBad = verifyPin('000000');
    assert('Invalid PIN rejected', !rBad.success);
  } catch(e) {
    assert('PIN verification', false, e.message);
  }

  // ── STEP 3: Semi-monthly file creation ────────────────────
  Logger.log('');
  Logger.log('── Step 3: Semi-monthly file creation ──');
  const now     = new Date();
  const tz      = Session.getScriptTimeZone();
  const today   = Utilities.formatDate(now, tz, 'MM/dd/yyyy');
  const period  = getPeriod(now);
  let monthly;
  try {
    monthly = getOrCreateMonthlySheet(now, TEST_BRANCH);
    assert('Monthly file created', monthly !== null);
    assert('File period is correct', monthly.getName().indexOf(period) !== -1, monthly.getName());
    assert('File branch is correct', monthly.getName().indexOf(TEST_BRANCH) !== -1, monthly.getName());
    Logger.log('  File name: ' + monthly.getName());

    const summary = monthly.getSheetByName(SUMMARY_SHEET);
    assert('Summary tab exists', summary !== null);
    const headers = summary.getRange(1, 1, 1, 7).getValues()[0];
    assert('Summary has OT column', headers[5] === 'OT Hrs', 'got: ' + headers[5]);
  } catch(e) {
    assert('Monthly file creation', false, e.message);
  }

  // ── STEP 4: Normal clock-in / clock-out ───────────────────
  Logger.log('');
  Logger.log('── Step 4: Normal clock-in / clock-out ──');
  try {
    const inResult = recordPunch(TEST_EMPLOYEE_1, TEST_BRANCH, 'in');
    assert('Clock-in succeeds', inResult.success, JSON.stringify(inResult));

    const dupIn = recordPunch(TEST_EMPLOYEE_1, TEST_BRANCH, 'in');
    assert('Duplicate clock-in blocked', !dupIn.success && dupIn.duplicate, JSON.stringify(dupIn));

    // Simulate clock-out by directly writing to the tab
    // (we can't wait 8 hours, so we manually set an earlier clock-in time)
    const empTab = getOrCreateEmployeeTab(monthly, TEST_EMPLOYEE_1);
    const data   = empTab.getDataRange().getValues();
    // Find today's row and backdate the clock-in by 3 hours
    for (let i = data.length - 1; i >= 1; i--) {
      const rawDate = data[i][0];
      const rowDate = rawDate instanceof Date
        ? Utilities.formatDate(rawDate, tz, 'MM/dd/yyyy')
        : String(rawDate).trim();
      if (rowDate === today && String(data[i][2]).trim() === '') {
        const fakeIn = new Date(now.getTime() - 3 * 60 * 60 * 1000);
        empTab.getRange(i + 1, 2).setValue(Utilities.formatDate(fakeIn, tz, 'hh:mm:ss a'));
        Logger.log('  Backdated clock-in by 3 hours for test');
        break;
      }
    }

    const outResult = recordPunch(TEST_EMPLOYEE_1, TEST_BRANCH, 'out');
    assert('Clock-out succeeds', outResult.success, JSON.stringify(outResult));

    const dupOut = recordPunch(TEST_EMPLOYEE_1, TEST_BRANCH, 'out');
    assert('Duplicate clock-out blocked', !dupOut.success && dupOut.duplicate, JSON.stringify(dupOut));

    // Verify hours written to tab
    const freshData = empTab.getDataRange().getValues();
    let foundRow = null;
    for (let i = 1; i < freshData.length; i++) {
      const rawDate = freshData[i][0];
      const rowDate = rawDate instanceof Date
        ? Utilities.formatDate(rawDate, tz, 'MM/dd/yyyy')
        : String(rawDate).trim();
      if (rowDate === today) { foundRow = freshData[i]; break; }
    }
    assert('Daily row has clock-in',  foundRow && String(foundRow[1]).trim() !== '', foundRow ? foundRow[1] : 'no row');
    assert('Daily row has clock-out', foundRow && String(foundRow[2]).trim() !== '', foundRow ? foundRow[2] : 'no row');
    assert('Regular hours recorded',  foundRow && parseFloat(foundRow[3]) > 0,       foundRow ? foundRow[3] : 'no row');
    assert('OT hours are zero (3hr shift)', foundRow && parseFloat(foundRow[4]) === 0, foundRow ? foundRow[4] : 'no row');
  } catch(e) {
    assert('Normal clock-in/out', false, e.message);
  }

  // ── STEP 5: Overtime scenario ─────────────────────────────
  Logger.log('');
  Logger.log('── Step 5: Overtime scenario (9 hrs) ──');
  try {
    // Use Employee 2 — clock in, backdate by 9 hours, clock out
    const inOT = recordPunch(TEST_EMPLOYEE_2, TEST_BRANCH, 'in');
    assert('OT employee clock-in succeeds', inOT.success, JSON.stringify(inOT));

    const empTab2 = getOrCreateEmployeeTab(monthly, TEST_EMPLOYEE_2);
    const data2   = empTab2.getDataRange().getValues();
    for (let i = data2.length - 1; i >= 1; i--) {
      const rawDate = data2[i][0];
      const rowDate = rawDate instanceof Date
        ? Utilities.formatDate(rawDate, tz, 'MM/dd/yyyy')
        : String(rawDate).trim();
      if (rowDate === today && String(data2[i][2]).trim() === '') {
        const fakeIn9 = new Date(now.getTime() - 9 * 60 * 60 * 1000);
        empTab2.getRange(i + 1, 2).setValue(Utilities.formatDate(fakeIn9, tz, 'hh:mm:ss a'));
        Logger.log('  Backdated clock-in by 9 hours for OT test');
        break;
      }
    }

    const outOT = recordPunch(TEST_EMPLOYEE_2, TEST_BRANCH, 'out');
    assert('OT employee clock-out succeeds', outOT.success, JSON.stringify(outOT));

    // Check that OT hours > 0 and regular = 8
    const freshData2 = empTab2.getDataRange().getValues();
    let otRow = null;
    for (let i = 1; i < freshData2.length; i++) {
      const rawDate = freshData2[i][0];
      const rowDate = rawDate instanceof Date
        ? Utilities.formatDate(rawDate, tz, 'MM/dd/yyyy')
        : String(rawDate).trim();
      if (rowDate === today) { otRow = freshData2[i]; break; }
    }
    assert('Regular hours capped at 8',    otRow && parseFloat(otRow[3]) === 8,   otRow ? otRow[3] : 'no row');
    assert('OT hours recorded (~1)',        otRow && parseFloat(otRow[4]) > 0,     otRow ? otRow[4] : 'no row');
    assert('Total hours ~9',               otRow && parseFloat(otRow[5]) >= 8.9,  otRow ? otRow[5] : 'no row');
    assert('Notes column mentions OT',     otRow && String(otRow[6]).indexOf('OT') !== -1, otRow ? otRow[6] : 'no row');
  } catch(e) {
    assert('Overtime scenario', false, e.message);
  }

  // ── STEP 6: Status check ──────────────────────────────────
  Logger.log('');
  Logger.log('── Step 6: getCurrentStatus ──');
  try {
    const status1 = getCurrentStatus(TEST_EMPLOYEE_1, TEST_BRANCH);
    assert('Employee 1 status is OUT after clock-out', status1.type === 'OUT', JSON.stringify(status1));
    assert('Employee 1 last time is populated', status1.time !== null && status1.time !== '', JSON.stringify(status1));
  } catch(e) {
    assert('getCurrentStatus', false, e.message);
  }

  // ── STEP 7: P1 / P2 file separation ──────────────────────
  Logger.log('');
  Logger.log('── Step 7: P1 / P2 period separation ──');
  try {
    const p1Date = new Date(now.getFullYear(), now.getMonth(), 1);
    const p2Date = new Date(now.getFullYear(), now.getMonth(), 16);
    const p1File = getOrCreateMonthlySheet(p1Date, TEST_BRANCH);
    const p2File = getOrCreateMonthlySheet(p2Date, TEST_BRANCH);
    assert('P1 file name contains P1', p1File.getName().indexOf('P1') !== -1, p1File.getName());
    assert('P2 file name contains P2', p2File.getName().indexOf('P2') !== -1, p2File.getName());
    assert('P1 and P2 are different files', p1File.getId() !== p2File.getId());
  } catch(e) {
    assert('P1/P2 separation', false, e.message);
  }

  // ── STEP 8: normalizeTime parsing ──────────────────────────
  Logger.log('');
  Logger.log('── Step 8: normalizeTime ──');
  try {
    assert('Parses "5:30 PM"',        normalizeTime('5:30 PM')    === '05:30:00 PM', normalizeTime('5:30 PM'));
    assert('Parses "05:30:00 PM"',     normalizeTime('05:30:00 PM') === '05:30:00 PM', normalizeTime('05:30:00 PM'));
    assert('Parses "17:30"',           normalizeTime('17:30')      === '05:30:00 PM', normalizeTime('17:30'));
    assert('Parses "9:00 AM"',         normalizeTime('9:00 AM')    === '09:00:00 AM', normalizeTime('9:00 AM'));
    assert('Parses "12:00 PM"',        normalizeTime('12:00 PM')   === '12:00:00 PM', normalizeTime('12:00 PM'));
    assert('Parses "530pm"',           normalizeTime('530pm')      === '05:30:00 PM', normalizeTime('530pm'));
    assert('Parses "1:05 am"',         normalizeTime('1:05 am')    === '01:05:00 AM', normalizeTime('1:05 am'));
    assert('Returns null for garbage', normalizeTime('banana')     === null,          String(normalizeTime('banana')));
    assert('Returns null for empty',   normalizeTime('')           === null,          String(normalizeTime('')));
  } catch(e) {
    assert('normalizeTime', false, e.message);
  }

  // ── STEP 9: getLocations ──────────────────────────────────
  Logger.log('');
  Logger.log('── Step 9: getLocations ──');
  try {
    const locs = getLocations();
    assert('getLocations returns success', locs.success === true, JSON.stringify(locs));
    assert('getLocations returns array',   Array.isArray(locs.locations), typeof locs.locations);
    if (locs.locations.length > 0) {
      const first = locs.locations[0];
      assert('Location has name',   typeof first.name === 'string' && first.name !== '', first.name);
      assert('Location has lat',    typeof first.lat === 'number' && !isNaN(first.lat), String(first.lat));
      assert('Location has lng',    typeof first.lng === 'number' && !isNaN(first.lng), String(first.lng));
      assert('Location has radius', typeof first.radius === 'number' && first.radius > 0, String(first.radius));
    } else {
      Logger.log('  (No locations in sheet — skipping field checks. Add rows to Locations tab.)');
    }
  } catch(e) {
    assert('getLocations', false, e.message);
  }

  // ── STEP 10: Correction form URL builder ──────────────────
  Logger.log('');
  Logger.log('── Step 10: buildCorrectionFormUrl ──');
  try {
    const props = PropertiesService.getScriptProperties();
    const formId = props.getProperty('CORRECTION_FORM_ID');
    if (formId) {
      const url = buildCorrectionFormUrl(TEST_EMPLOYEE_1, today, TEST_BRANCH);
      assert('Form URL is not empty',            url !== '', url);
      assert('Form URL contains employee name',  url.indexOf(encodeURIComponent(TEST_EMPLOYEE_1)) !== -1, url);
      assert('Form URL contains date',           url.indexOf(encodeURIComponent(today)) !== -1, url);
      assert('Form URL contains branch',         url.indexOf(encodeURIComponent(TEST_BRANCH)) !== -1, url);
      Logger.log('  URL: ' + url);
    } else {
      Logger.log('  (Correction form not created yet — run createCorrectionForm() first. Skipping.)');
    }
  } catch(e) {
    assert('buildCorrectionFormUrl', false, e.message);
  }

  // ── STEP 11: Missed clock-out simulation ──────────────────
  Logger.log('');
  Logger.log('── Step 11: Missed clock-out flow ──');
  try {
    // Clock in Employee 1 again (fresh row), don't clock out
    // First need to clear the duplicate guard by using a different "date"
    // We'll directly write a row to simulate an open clock-in
    const empTab = getOrCreateEmployeeTab(monthly, TEST_EMPLOYEE_1);
    const fakeClockIn = new Date(now.getTime() - 10 * 60 * 60 * 1000); // 10 hrs ago
    const fakeDate = Utilities.formatDate(fakeClockIn, tz, 'MM/dd/yyyy');
    const fakeTime = Utilities.formatDate(fakeClockIn, tz, 'hh:mm:ss a');
    // Only add if we haven't already (avoid duplicate test rows)
    empTab.appendRow([fakeDate, fakeTime, '', '', '', '', '']);

    // Verify the open row exists
    const openData = empTab.getDataRange().getValues();
    let hasOpenRow = false;
    for (let i = 1; i < openData.length; i++) {
      const rd = openData[i][0] instanceof Date
        ? Utilities.formatDate(openData[i][0], tz, 'MM/dd/yyyy')
        : String(openData[i][0]).trim();
      if (rd === fakeDate && String(openData[i][2]).trim() === '') hasOpenRow = true;
    }
    assert('Open clock-in row exists for missed test', hasOpenRow);

    // Test email map building
    const master2    = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    const emailMap   = {};
    getBranchSheets(master2).forEach(function(sheet) {
      const d = sheet.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) {
        const n = String(d[i][0]).trim();
        const e = String(d[i][2]).trim();
        if (n && e && e.indexOf('@') !== -1) emailMap[n] = e;
      }
    });
    assert('Email found for TEST Alice', emailMap[TEST_EMPLOYEE_1] === TEST_EMAIL_1, emailMap[TEST_EMPLOYEE_1]);
    assert('Email found for TEST Bob',   emailMap[TEST_EMPLOYEE_2] === TEST_EMAIL_2, emailMap[TEST_EMPLOYEE_2]);
  } catch(e) {
    assert('Missed clock-out flow', false, e.message);
  }

  // ── STEP 12: Cleanup ───────────────────────────────────────
  Logger.log('');
  Logger.log('── Step 12: Cleanup ──');
  try {
    // Delete test branch tab from master sheet
    const master    = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
    const testSheet = master.getSheetByName(TEST_BRANCH);
    if (testSheet) master.deleteSheet(testSheet);
    assert('Test branch tab removed', master.getSheetByName(TEST_BRANCH) === null);

    // Delete test Drive folder and all contents
    const root    = getRootFolder();
    const folders = root.getFoldersByName(TEST_BRANCH);
    let folderCount = 0;
    while (folders.hasNext()) {
      const folder = folders.next();
      // Delete all files inside
      const files = folder.getFiles();
      while (files.hasNext()) files.next().setTrashed(true);
      folder.setTrashed(true);
      folderCount++;
    }
    assert('Test Drive folder deleted', folderCount > 0);
    Logger.log('  Deleted ' + folderCount + ' test folder(s)');
  } catch(e) {
    assert('Cleanup', false, e.message);
  }

  // ── FINAL REPORT ──────────────────────────────────────────
  Logger.log('');
  Logger.log('╔══════════════════════════════════════╗');
  Logger.log('║             TEST RESULTS             ║');
  Logger.log('╠══════════════════════════════════════╣');
  Logger.log('║  PASSED: ' + passed + '/' + (passed + failed) + '                          ║'.substring(String(passed + '/' + (passed+failed)).length));
  Logger.log('║  FAILED: ' + failed + '                              ║'.substring(String(failed).length));
  Logger.log('╚══════════════════════════════════════╝');

  if (failed > 0) {
    Logger.log('');
    Logger.log('Failed tests:');
    results.filter(r => !r.pass).forEach(r => Logger.log('  ✗ ' + r.label + (r.detail ? ' — ' + r.detail : '')));
  } else {
    Logger.log('');
    Logger.log('All tests passed. Safe to deploy.');
  }
}

