// ════════════════════════════════════════════════════════════════
// BHDC Downtime Tracker — Google Apps Script
// Paste this entire file into: script.google.com > New Project
// ════════════════════════════════════════════════════════════════

// ── YOUR GOOGLE SHEET ID ─────────────────────────────────────────
// Found in the sheet URL:
// https://docs.google.com/spreadsheets/d/  >>>SHEET_ID_HERE<<<  /edit
const SHEET_ID = "YOUR_GOOGLE_SHEET_ID_HERE";

// ── Sheet tab names (must match exactly) ─────────────────────────
const PUMP_SHEET = "Test Pump Log";
const WCE_SHEET  = "Well Control Equipment Log";

// ════════════════════════════════════════════════════════════════
// DO NOT EDIT BELOW THIS LINE
// ════════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    // FormData sends the JSON string under the key "data"
    const raw  = e.parameter.data || e.postData.contents;
    const data = JSON.parse(raw);
    const ss   = SpreadsheetApp.openById(SHEET_ID);
    const category = data.category || "";

    if (category === "Pressure Test Pump") {
      writePumpLog(ss, data);
    } else if (category === "Well Control Equipment") {
      writeWCELog(ss, data);
    } else {
      return respond(false, "Unknown category: " + category);
    }

    return respond(true, "Entry logged successfully");

  } catch (err) {
    return respond(false, err.message);
  }
}

// ── Write to Test Pump Log ────────────────────────────────────────
function writePumpLog(ss, d) {
  const ws      = ss.getSheetByName(PUMP_SHEET);
  const lastRow = getLastDataRow(ws);
  const entryNo = lastRow - 3;

  ws.getRange(lastRow, 1, 1, 15).setValues([[
    entryNo,       // 1  Entry #
    d.rigNo,       // 2  Rig No.
    d.eqpName,     // 3  Equipment Name
    d.serialNo,    // 4  Serial Number
    d.failDT,      // 5  Failure Date & Time
    d.restartDT,   // 6  Restart Date & Time
    d.dtHours,     // 7  Downtime (hrs)
    d.severity,    // 8  Severity
    d.status,      // 9  Status
    d.failDesc,    // 10 Failure Description
    d.rootCause,   // 11 Root Cause
    d.parts,       // 12 Parts Replaced
    d.repairedBy,  // 13 Repaired By
    d.verifiedBy,  // 14 Verified By
    d.remarks,     // 15 Remarks
  ]]);

  applyRowFormatting(ws, lastRow, 8, d.severity);
}

// ── Write to Well Control Equipment Log ──────────────────────────
function writeWCELog(ss, d) {
  const ws      = ss.getSheetByName(WCE_SHEET);
  const lastRow = getLastDataRow(ws);
  const entryNo = lastRow - 3;

  ws.getRange(lastRow, 1, 1, 16).setValues([[
    entryNo,       // 1  Entry #
    d.rigNo,       // 2  Rig No.
    d.eqpName,     // 3  Equipment Name
    d.serialNo,    // 4  Serial Number
    d.eqpType,     // 5  Equipment Type
    d.failDT,      // 6  Failure Date & Time
    d.restartDT,   // 7  Restart Date & Time
    d.dtHours,     // 8  Downtime (hrs)
    d.severity,    // 9  Severity
    d.status,      // 10 Status
    d.failDesc,    // 11 Failure Description
    d.rootCause,   // 12 Root Cause
    d.parts,       // 13 Parts Replaced
    d.repairedBy,  // 14 Repaired By
    d.verifiedBy,  // 15 Verified By
    d.remarks,     // 16 Remarks
  ]]);

  applyRowFormatting(ws, lastRow, 9, d.severity);
}

// ── Alternating rows + severity colour ───────────────────────────
function applyRowFormatting(ws, row, severityCol, severity) {
  const numCols = ws.getLastColumn();
  ws.getRange(row, 1, 1, numCols)
    .setBackground(row % 2 === 0 ? "#F2F2F2" : "#FFFFFF");

  const sevCell = ws.getRange(row, severityCol);
  if (severity === "Major") {
    sevCell.setBackground("#FCE4D6");
    sevCell.setFontColor("#C0392B");
    sevCell.setFontWeight("bold");
  } else if (severity === "Minor") {
    sevCell.setBackground("#FFF2CC");
    sevCell.setFontColor("#7D6608");
    sevCell.setFontWeight("normal");
  }
}

// ── Next empty data row (data starts at row 4) ───────────────────
function getLastDataRow(ws) {
  const col    = ws.getRange("A:A").getValues();
  let lastRow  = 3; // header is rows 1-3
  for (let i = 3; i < col.length; i++) {
    if (col[i][0] !== "") lastRow = i + 1;
  }
  return lastRow + 1; // next empty row
}

// ── JSON response ────────────────────────────────────────────────
function respond(success, message) {
  return ContentService
    .createTextOutput(JSON.stringify({ success, message }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET handler (health check) ───────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "BHDC Downtime Tracker API is running ✓" }))
    .setMimeType(ContentService.MimeType.JSON);
}
