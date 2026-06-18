/**file name: sheet.gs
/**
 * ============================================================================
 * DOCUMAIL PRO COMPLETE MASTER CORE SCRIPT
 * ============================================================================
 * FILE: sheet.gs - Sheet Initialization, Headers, and Status Columns
 * ============================================================================
 */

// ==========================================
// FUNCTION: GENERATE_DOCUMAIL_TEMPLATE Starts
// ==========================================
function GENERATE_DOCUMAIL_TEMPLATE() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  // 1. DATA EXISTENCE CHECK: Warn user if sheet has pre-existing data
  try {
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    
    if (lastRow > 0 && lastColumn > 0) {
      var response = ui.alert(
        "⚠️ Warning", 
        "There is data in your sheet. Initiating will remove all data, column headers, AND all saved templates. Are you sure?", 
        ui.ButtonSet.YES_NO
      );
      
      if (response !== ui.Button.YES) {
        console.log("Initialization cancelled by user.");
        return; 
      }
    }
  } catch(e) {
    console.log("Non-blocking data check warning error: " + e.message);
  }

  // ERASE ALL SAVED TEMPLATES ON INITIALIZATION
  try {
    var props = PropertiesService.getDocumentProperties();
    props.deleteProperty('documail_templates'); 
    props.setProperty('SIDEBAR_REFRESH_SIGNAL_KEY', "REFRESH_WIPE_" + new Date().getTime());
    console.log("🧹 Saved templates property database successfully purged.");
  } catch (err) {
    console.log("Error clearing templates property database: " + err.message);
  }

  // 2. PROCEED WITH WIPE AND RENDER
  var existingProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < existingProtections.length; i++) {
    existingProtections[i].remove();
  }
  SpreadsheetApp.flush();

  sheet.clear();
  sheet.clearFormats();
  sheet.clearContents();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();

  // Standardizing column blueprint to "Column 1", "Column 2", etc.
  var headers = [
    "Column 1", "Column 2", "Column 3", "Column 4",
    "Recipient Email",
    "Merged Doc Status",
    "Merged Doc ID",
    "Merged Doc URL"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold")
    .setBackground("#E8F0FE")
    .setFontColor("#1A73E8")
    .setWrap(true)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.setFrozenRows(1);

  sheet.getRange("E1").setFontColor("#137333").setBackground("#E6F4EA");
  var emailRange = sheet.getRange("E2:E1000");
  emailRange.setDataValidation(SpreadsheetApp.newDataValidation().requireTextIsEmail().setAllowInvalid(false).build());

  // System columns range (F through H only)
  var systemRange = sheet.getRange("F2:H1000");
  systemRange.setBackground("#f3f3f3");
  systemRange.setFontColor("#888888");

  sheet.getRange("F1").setNote("🔒 SYSTEM COLUMN - Auto-generated. Do not edit.");
  sheet.getRange("G1").setNote("🔒 SYSTEM COLUMN - Auto-generated Document ID");
  sheet.getRange("H1").setNote("🔒 SYSTEM COLUMN - Auto-generated Document URL");

  // =======================================================
  // FIXED: COMBINED COMPREHENSIVE TIP NOTE ATTACHED TO A1
  // =======================================================
  sheet.getRange("A1").setNote(
    "💡 TIP:\n" +
    "1. You can double-click any header (Row 1) to Rename it! (e.g., Change to 'Date', 'City', etc.)\n" +
    "2. Need more columns? Right-click Column E and choose 'Insert column left' to add more fields.\n" +
    "3. You can safely add, delete, or rename columns to the RIGHT of Column A-D or to the LEFT of Column F.\n\n" +
    "⚠️ Columns F-H are system locked."
  );

  // Protect only F through H (PDF system columns)
  var protection = sheet.getRange("F1:H1000").protect();
  protection.setDescription('DocuMail Pro System Columns (PDF) - Locked');
  protection.removeEditors(protection.getEditors());

  // Set tight safety base width profiles (No stretching anymore!)
  sheet.setColumnWidth(1, 110); 
  sheet.setColumnWidth(2, 110); 
  sheet.setColumnWidth(3, 110); 
  sheet.setColumnWidth(4, 110); 
  sheet.setColumnWidth(5, 200); 
  sheet.setColumnWidth(6, 150); 
  sheet.setColumnWidth(7, 150); 
  sheet.setColumnWidth(8, 300); 

  // Auto-resize targets cleanly matching exact text measurements
  for (var col = 1; col <= headers.length; col++) {
    sheet.autoResizeColumn(col);
    if (sheet.getColumnWidth(col) < 100) {
      sheet.setColumnWidth(col, 110);
    }
  }
  
  sheet.setRowHeight(1, 40);

  // Clean, non-technical alert
  ui.alert(
    "✅ Sheet Structure Ready!\n\n" +
    "👉 QUICK GUIDE FOR YOU:\n" +
    "1. You can Rename 'Column 1, 2, 3, 4' to whatever you want (like Date, Invoice No, etc.).\n" +
    "2. You can Delete columns you don't need.\n" +
    "3. You can Add new columns anywhere between Column A and Column E.\n\n" +
    "⚠️ Note: Columns F to H are system locked to protect your generated PDF links."
  );
}

// ==========================================
// Helper Functions
// ==========================================
function GET_ALL_RAW_HEADERS() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastCol = sheet.getLastColumn();
    if (lastCol < 1) lastCol = 1;
    return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  } catch (e) {
    // UPDATED: Clean ghost fallback array matching new layout scheme
    return ["Column 1", "Column 2", "Column 3", "Column 4", "Recipient Email"];
  }
}

function GET_LIVE_SHEET_HEADERS() {
  try {
    var rawHeaders = GET_ALL_RAW_HEADERS();
    var result = [];
    for (var i = 0; i < rawHeaders.length; i++) {
      var hStr = String(rawHeaders[i]).toLowerCase().trim();
      if (hStr === "") continue;
      if (hStr.indexOf("recipient email") !== -1) continue;
      if (hStr.indexOf("merged doc") !== -1) continue;
      // NEW: Exclude ANY column that starts with "sent mail status"
      if (hStr.indexOf("sent mail status") !== -1) continue;
      result.push(rawHeaders[i]);
    }
    return result;
  } catch (e) {
    // UPDATED: Safe fallback defaults matching your new 4 column baseline layout
    return ["Column 1", "Column 2", "Column 3", "Column 4"];
  }
}

function GET_WIZARD_INITIALIZATION_PAYLOAD() {
  return { userHeaders: GET_LIVE_SHEET_HEADERS() };
}

/**
* Gets the column index for a template's status column.
* Creates the column if it doesn't exist (adds to the right of existing data).
* 
* @param {Sheet} sheet - The active sheet
* @param {string} templateName - Name of the email template
* @returns {number} Column index (0-based)
*/
function GET_OR_CREATE_STATUS_COLUMN(sheet, templateName) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var statusColumnName = "Sent Mail Status - " + templateName;

  // Look for existing column
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).toLowerCase().trim() === statusColumnName.toLowerCase()) {
      return i;
    }
  }

  // Column doesn't exist - create it at the end
  var newColIndex = headers.length;
  var newColNumber = newColIndex + 1;
  var headerRange = sheet.getRange(1, newColNumber);

  // Set header value
  headerRange.setValue(statusColumnName);

  // Apply HEADER styling (light blue background, matching standard headers)
  headerRange.setFontWeight("bold")
    .setBackground("#E8F0FE")
    .setFontColor("#1A73E8")
    .setWrap(true)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // Add note/tooltip
  headerRange.setNote("🔒 SYSTEM COLUMN - Auto-generated email status for template: " + templateName);

  // Apply DATA CELL styling (grey background for rows 2-1000)
  var dataRange = sheet.getRange(2, newColNumber, 999, 1);
  dataRange.setBackground("#f3f3f3")
    .setFontColor("#888888");

  // Protect the column (read-only for users)
  var protection = sheet.getRange(1, newColNumber, sheet.getMaxRows(), 1).protect();
  protection.setDescription('DocuMail Pro Status Column - ' + templateName);
  protection.removeEditors(protection.getEditors());

  // Auto-resize the column
  sheet.autoResizeColumn(newColNumber);

  return newColIndex;
}

/**
 * STEP 1: ONLY checks if a template's status column contains any row data.
 * Does NOT delete anything.
 */
function CHECK_TEMPLATE_RECORDS(templateId) {
  var templates = GET_ALL_TEMPLATES();
  var targetTemplate = null;

  for (var i = 0; i < templates.length; i++) {
    if (templates[i].id === templateId) {
      targetTemplate = templates[i];
      break;
    }
  }

  if (!targetTemplate) return { hasRecords: false };

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastColumn = sheet.getLastColumn();

    if (lastColumn > 0) {
      var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
      var statusColumnName = ("Sent Mail Status - " + targetTemplate.name).toLowerCase().trim();
      var targetColIndex = -1;

      for (var c = 0; c < headers.length; c++) {
        if (String(headers[c]).toLowerCase().trim() === statusColumnName) {
          targetColIndex = c + 1;
          break;
        }
      }

      if (targetColIndex !== -1) {
        var lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          var dataRange = sheet.getRange(2, targetColIndex, lastRow - 1, 1).getValues();
          for (var r = 0; r < dataRange.length; r++) {
            if (String(dataRange[r][0]).trim() !== "") {
              return { hasRecords: true }; // Data found!
            }
          }
        }
      }
    }
  } catch (e) {
    console.log("Error checking template rows: " + e.message);
  }

  return { hasRecords: false }; // Safe (empty)
}
//file content end