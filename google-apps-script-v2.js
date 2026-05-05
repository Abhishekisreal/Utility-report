// ─────────────────────────────────────────────────────────────────────────────
// GOOGLE APPS SCRIPT — Utility Shift Report → Single Row Per Submission
// ─────────────────────────────────────────────────────────────────────────────
// SETUP:
// 1. Go to https://script.google.com → New Project → paste this file
// 2. Deploy → New Deployment → Web App → "Anyone" → Deploy
// 3. Copy the Web App URL into your website JS (APPS_SCRIPT_URL variable)
// ─────────────────────────────────────────────────────────────────────────────

// Paste your Google Sheet ID here (from the URL: /spreadsheets/d/YOUR_ID/edit)
// Leave blank to auto-create a new sheet on first run
const SPREADSHEET_ID = "";
const SHEET_NAME     = "Utility Reports";

// Column headers — each submission = ONE ROW
const HEADERS = [
  "Timestamp",                        // A - auto
  "Date",                             // B
  "Shift",                            // C
  "Period Type",                      // D  — "Shift" or "Complete Day"
  "Supervisor",                       // E
  "GCW Make Up RO Water (unit)",      // F
  "RO Water (unit)",                  // G
  "DI Water (unit)",                  // H
  "Argon (unit)",                     // I
  "CDA (SCFM/unit)",                  // J
  "Argon Tank Remaining (Days)",      // K
  "General WW (unit)",                // L
  "Acid WW (unit)",                   // M
  "Electrical Consumption (kWh/unit)",// N
  "Argon Status",                     // O
  "Remarks",                          // P
  "Concerns"                          // Q
];

function doPost(e) {
  try {
    const raw  = e.postData.contents;
    const data = JSON.parse(raw);

    // Open or create spreadsheet
    const ss = SPREADSHEET_ID
      ? SpreadsheetApp.openById(SPREADSHEET_ID)
      : SpreadsheetApp.create("Utility Shift Reports");

    let sheet = ss.getSheetByName(SHEET_NAME);

    // ── First run: create sheet with styled header row ──────────────────────
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);

      // Write headers
      sheet.appendRow(HEADERS);

      // Style header
      const hdr = sheet.getRange(1, 1, 1, HEADERS.length);
      hdr.setBackground("#0f1923");
      hdr.setFontColor("#ffffff");
      hdr.setFontWeight("bold");
      hdr.setFontSize(9);
      hdr.setHorizontalAlignment("center");
      hdr.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

      // Column widths
      sheet.setColumnWidth(1, 150);   // Timestamp
      sheet.setColumnWidth(2, 100);   // Date
      sheet.setColumnWidth(3, 180);   // Shift
      sheet.setColumnWidth(4, 110);   // Period Type
      sheet.setColumnWidth(5, 140);   // Supervisor
      for (let c = 6; c <= 14; c++) sheet.setColumnWidth(c, 130); // Utility cols
      sheet.setColumnWidth(15, 160);  // Argon Status
      sheet.setColumnWidth(16, 200);  // Remarks
      sheet.setColumnWidth(17, 200);  // Concerns

      sheet.setFrozenRows(1);
      sheet.setFrozenColumns(2);      // Freeze Date + Shift so you can scroll right

      Logger.log("Sheet created: " + ss.getUrl());
    }

    // ── Resolve unit label ───────────────────────────────────────────────────
    const isCompleteDay = (data.periodType === "Complete Day");
    const volUnit  = isCompleteDay ? "M³/Day"   : "M³/Shift";
    const flowUnit = isCompleteDay ? "SCFM/Day"  : "SCFM/Shift";
    const kwUnit   = isCompleteDay ? "kWh/Day"   : "kWh/Shift";

    // ── Argon alert ─────────────────────────────────────────────────────────
    let argonStatus = "OK";
    if (data.argdays !== "" && data.argdays !== undefined) {
      const d = parseFloat(data.argdays);
      if      (d <= 2) argonStatus = "CRITICAL — ORDER NOW";
      else if (d <= 7) argonStatus = "Low — Plan Refill";
    }

    // ── Build single data row ────────────────────────────────────────────────
    const row = [
      new Date(),                                           // A Timestamp
      data.date        || "",                               // B Date
      data.shift       || "",                               // C Shift (full name)
      data.periodType  || "",                               // D Period Type
      data.supervisor  || "",                               // E Supervisor
      formatVal(data.gcw)    + " " + volUnit,               // F GCW RO Water
      formatVal(data.ro)     + " " + volUnit,               // G RO Water
      formatVal(data.di)     + " " + volUnit,               // H DI Water
      formatVal(data.argon)  + " " + volUnit,               // I Argon
      formatVal(data.cda)    + " " + flowUnit,              // J CDA
      formatVal(data.argdays)+ " Days",                     // K Argon Tank
      formatVal(data.gww)    + " " + volUnit,               // L General WW
      formatVal(data.aww)    + " " + volUnit,               // M Acid WW
      formatVal(data.elec)   + " " + kwUnit,                // N Electrical
      argonStatus,                                          // O Argon Status
      data.remarks   || "NIL",                              // P Remarks
      data.concerns  || "NIL",                              // Q Concerns
    ];

    sheet.appendRow(row);

    // ── Highlight critical argon rows ────────────────────────────────────────
    const lastRow = sheet.getLastRow();
    if (argonStatus === "CRITICAL — ORDER NOW") {
      sheet.getRange(lastRow, 1, 1, HEADERS.length).setBackground("#FCEBEB");
      sheet.getRange(lastRow, 15).setFontColor("#A32D2D").setFontWeight("bold");
    } else if (argonStatus === "Low — Plan Refill") {
      sheet.getRange(lastRow, 15).setFontColor("#854F0B").setFontWeight("bold");
    }

    // ── Alternate row shading for readability ────────────────────────────────
    if (lastRow % 2 === 0) {
      sheet.getRange(lastRow, 1, 1, HEADERS.length).setBackground("#F9F9F7");
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", row: lastRow, sheetUrl: ss.getUrl() }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log("Error: " + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function formatVal(v) {
  if (v === undefined || v === null || v === "") return "—";
  return String(v).trim();
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ready" }))
    .setMimeType(ContentService.MimeType.JSON);
}
