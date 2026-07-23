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
function HAS_ACTUAL_DATA_ROWS(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol <= 0) return false;

  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
  for (var r = 0; r < data.length; r++) {
    for (var c = 0; c < data[r].length; c++) {
      if (String(data[r][c]).trim() !== "") {
        return true;
      }
    }
  }
  return false;
}

function GENERATE_DOCUMAIL_TEMPLATE() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  // =======================================================
  // SAFETY CHECK: Block if sheet has actual data rows
  // =======================================================
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var actualData = lastRow > 1 ? HAS_ACTUAL_DATA_ROWS(sheet) : false;

  if (lastRow > 1 && !actualData) {
    console.log("⚠️ Found " + (lastRow - 1) + " rows with formulas but no visible data. Proceeding with initialization.");
  }

  if (actualData) {
    ui.alert(
      "❌ Cannot Initialize Sheet Template",
      "This sheet already has data in rows.\n\n" +
      "📊 Rows with data: " + (lastRow - 1) + "\n\n" +
      "To protect your existing data, sheet template initialization is not allowed.\n\n" +
      "Please use a new worksheet or a blank workbook to initialize the template.",
      ui.ButtonSet.OK
    );
    return;
  }

  // =======================================================
  // SAFETY CHECK: Warn if system columns already exist
  // =======================================================
  if (lastColumn > 0) {
    var existingHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    var hasSystemCols = false;
    for (var i = 0; i < existingHeaders.length; i++) {
      var h = String(existingHeaders[i]).toLowerCase().trim();
      if (h.indexOf("merged doc status") !== -1 || h.indexOf("recipient email") !== -1) {
        hasSystemCols = true;
        break;
      }
    }
    if (hasSystemCols) {
      ui.alert(
        "✅ DocuMail PRO Already Configured",
        "This sheet is already configured with DocuMail PRO system columns.\n\nNo changes are needed.",
        ui.ButtonSet.OK
      );
      return;
    }
  }

  // ERASE ALL SAVED TEMPLATES ON INITIALIZATION
  try {
    var props = PropertiesService.getDocumentProperties();
    var key = 'documail_templates_' + sheet.getName();
    props.deleteProperty(key);
    var refreshKey = 'SIDEBAR_REFRESH_SIGNAL_KEY_' + sheet.getName();
    props.setProperty(refreshKey, "REFRESH_WIPE_" + new Date().getTime());
    console.log("🧹 Saved templates property database successfully purged.");
  } catch (err) {
    console.log("Error clearing templates property database: " + err.message);
  }

  // If no columns, create initial headers
  if (lastColumn === 0) {
    var initialHeaders = [
      "Column 1", "Column 2", "Column 3", "Column 4",
      "Recipient Email",
      "Merged Doc Status",
      "Merged Doc ID",
      "Merged Doc URL"
    ];

    var totalCols = initialHeaders.length;
    sheet.getRange(1, 1, 1, totalCols).setValues([initialHeaders]);

    var headerRange = sheet.getRange(1, 1, 1, totalCols);
    headerRange.setFontWeight("bold")
      .setBackground("#E8F0FE")
      .setFontColor("#1A73E8")
      .setWrap(true)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    sheet.setFrozenRows(1);

    // Recipient Email column (column 5) - Email validation
    var emailCol = 5;
    var emailRange = sheet.getRange(2, emailCol, sheet.getMaxRows() - 1, 1);
    emailRange.setDataValidation(SpreadsheetApp.newDataValidation().requireTextIsEmail().setAllowInvalid(false).build());
    sheet.getRange(1, emailCol).setFontColor("#137333").setBackground("#E6F4EA");

    // System columns range (columns 6, 7, 8)
    var systemStartCol = 6;
    var systemRange = sheet.getRange(2, systemStartCol, sheet.getMaxRows() - 1, 3);
    systemRange.setBackground("#f3f3f3")
      .setFontColor("#888888");

    sheet.getRange(1, systemStartCol).setNote("🔒 SYSTEM COLUMN - Auto-generated. Do not edit.");
    sheet.getRange(1, systemStartCol + 1).setNote("🔒 SYSTEM COLUMN - Auto-generated Document ID");
    sheet.getRange(1, systemStartCol + 2).setNote("🔒 SYSTEM COLUMN - Auto-generated Document URL");

    // =======================================================
    // COMPREHENSIVE TIP NOTE ATTACHED TO A1
    // =======================================================
    sheet.getRange("A1").setNote(
      "💡 TIP:\n" +
      "1. You can double-click any header (Row 1) to Rename it! (e.g., Change to 'Date', 'City', etc.)\n" +
      "2. Need more columns? Right-click Column E and choose 'Insert column left' to add more fields.\n" +
      "3. You can safely add, delete, or rename columns to the RIGHT of Column A-D or to the LEFT of Column F.\n\n" +
      "⚠️ Columns F-H are system locked."
    );

    // Protect only F through H (PDF system columns)
    var protection = sheet.getRange(1, systemStartCol, sheet.getMaxRows(), 3).protect();
    protection.setDescription('DocuMail Pro System Columns (PDF) - Locked');
    protection.removeEditors(protection.getEditors());

    // Set column widths
    var columnWidths = [110, 110, 110, 110, 200, 150, 150, 300];
    for (var i = 0; i < columnWidths.length; i++) {
      sheet.setColumnWidth(i + 1, columnWidths[i]);
    }

    // Auto-resize
    for (var col = 1; col <= totalCols; col++) {
      sheet.autoResizeColumn(col);
      if (sheet.getColumnWidth(col) < 100) {
        sheet.setColumnWidth(col, 110);
      }
    }

    sheet.setRowHeight(1, 40);

    ui.alert(
      "✅ Sheet Structure Ready!\n\n" +
      "👉 QUICK GUIDE FOR YOU:\n" +
      "1. You can Rename 'Column 1, 2, 3, 4' to whatever you want (like Date, Invoice No, etc.).\n" +
      "2. You can Delete columns you don't need.\n" +
      "3. You can Add new columns anywhere between Column A and Column E.\n\n" +
      "⚠️ Note: Columns F to H are system locked to protect your generated PDF links."
    );

  } else {
    // =======================================================
    // EXISTING COLUMNS (Headers only, no data rows) - APPEND SYSTEM COLUMNS
    // =======================================================
    var existingHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    var newHeaders = [
      "Recipient Email",
      "Merged Doc Status",
      "Merged Doc ID",
      "Merged Doc URL"
    ];

    var startCol = lastColumn + 1;

    // Create the range for the new headers
    var newRange = sheet.getRange(1, startCol, 1, newHeaders.length);
    newRange.setValues([newHeaders]);

    // Apply styling to these 4 columns
    newRange.setFontWeight("bold")
      .setBackground("#E8F0FE")
      .setFontColor("#1A73E8")
      .setWrap(true)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    // Recipient Email column (first of the 4) - Email validation
    var emailCol = startCol;
    var emailRange = sheet.getRange(2, emailCol, sheet.getMaxRows() - 1, 1);
    emailRange.setDataValidation(SpreadsheetApp.newDataValidation().requireTextIsEmail().setAllowInvalid(false).build());
    sheet.getRange(1, emailCol).setFontColor("#137333").setBackground("#E6F4EA");

    // System columns (3 columns after Recipient Email)
    var systemStartCol = startCol + 1;
    var systemRange = sheet.getRange(2, systemStartCol, sheet.getMaxRows() - 1, 3);
    systemRange.setBackground("#f3f3f3")
      .setFontColor("#888888");

    // Add notes for each system column
    sheet.getRange(1, startCol).setNote("🔒 SYSTEM COLUMN - Recipient Email");
    sheet.getRange(1, startCol + 1).setNote("🔒 SYSTEM COLUMN - Auto-generated. Do not edit.");
    sheet.getRange(1, startCol + 2).setNote("🔒 SYSTEM COLUMN - Auto-generated Document ID");
    sheet.getRange(1, startCol + 3).setNote("🔒 SYSTEM COLUMN - Auto-generated Document URL");

    // =======================================================
    // COMPREHENSIVE TIP NOTE ATTACHED TO A1
    // =======================================================
    sheet.getRange("A1").setNote(
      "💡 TIP:\n" +
      "1. You can double-click any header (Row 1) to Rename it! (e.g., Change to 'Date', 'City', etc.)\n" +
      "2. Need more columns? Right-click Column E and choose 'Insert column left' to add more fields.\n" +
      "3. You can safely add, delete, or rename columns to the RIGHT of Column A-D or to the LEFT of Column F.\n\n" +
      "⚠️ Columns " + String.fromCharCode(64 + systemStartCol) + " to " + String.fromCharCode(64 + startCol + newHeaders.length - 1) + " are system locked."
    );

    // Protect the 4 system columns (entire column)
    var protection = sheet.getRange(1, startCol, sheet.getMaxRows(), newHeaders.length).protect();
    protection.setDescription('DocuMail Pro System Columns - Locked');
    protection.removeEditors(protection.getEditors());

    // Set column widths
    var columnWidths = [200, 150, 150, 300];
    for (var i = 0; i < columnWidths.length; i++) {
      sheet.setColumnWidth(startCol + i, columnWidths[i]);
    }

    // Auto-resize
    for (var col = startCol; col < startCol + newHeaders.length; col++) {
      sheet.autoResizeColumn(col);
      if (sheet.getColumnWidth(col) < 100) {
        sheet.setColumnWidth(col, 110);
      }
    }

    // Set frozen rows if not already
    if (sheet.getFrozenRows() === 0) {
      sheet.setFrozenRows(1);
    }

    ui.alert(
      "✅ New system columns added to the right of existing columns!\n\n" +
      "📊 Existing Columns: " + lastColumn + "\n" +
      "➕ System Columns Added: " + newHeaders.length + "\n" +
      "📋 Total Columns: " + (lastColumn + newHeaders.length) + "\n\n" +
      "⚠️ Note: The 4 new columns are system locked."
    );
  }
}
// ==========================================
// FUNCTION: GENERATE_DOCUMAIL_TEMPLATE Ends
// ==========================================

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
    var foundRecipientEmail = false;
    
    for (var i = 0; i < rawHeaders.length; i++) {
      var hStr = String(rawHeaders[i]).toLowerCase().trim();
      if (hStr === "") continue;
      
      // Skip Recipient Email and everything after it
      if (hStr.indexOf("recipient email") !== -1) {
        foundRecipientEmail = true;
        continue;
      }
      if (foundRecipientEmail) {
        continue; // Skip all columns after Recipient Email
      }
      
      // Skip Merged Doc columns (just in case)
      if (hStr.indexOf("merged doc") !== -1) continue;
      // Skip Sent Mail Status columns
      if (hStr.indexOf("sent mail status") !== -1) continue;
      
      result.push(rawHeaders[i]);
    }
    return result;
  } catch (e) {
    return ["Column 1", "Column 2", "Column 3", "Column 4"];
  }
}

function GET_WIZARD_INITIALIZATION_PAYLOAD() {
  var headers = GET_LIVE_SHEET_HEADERS();
  
  // Find Recipient Email column as default
  var defaultField = "";
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].toLowerCase().indexOf("recipient email") !== -1) {
      defaultField = headers[i];
      break;
    }
  }
  
  return {
    userHeaders: headers,
    defaultField: defaultField
  };
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

      // Check Sent Mail Status column (for EMAIL_ONLY and BOTH)
      if (targetTemplate.type === "EMAIL_ONLY" || targetTemplate.type === "BOTH") {
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
                return { hasRecords: true };
              }
            }
          }
        }
      }

      // Check Merged Doc Status column (for PDF_ONLY and BOTH)
      if (targetTemplate.type === "PDF_ONLY" || targetTemplate.type === "BOTH") {
        for (var c = 0; c < headers.length; c++) {
          if (String(headers[c]).toLowerCase().indexOf("merged doc status") !== -1) {
            var lastRow = sheet.getLastRow();
            if (lastRow > 1) {
              var dataRange = sheet.getRange(2, c + 1, lastRow - 1, 1).getValues();
              for (var r = 0; r < dataRange.length; r++) {
                if (String(dataRange[r][0]).trim() !== "") {
                  return { hasRecords: true };
                }
              }
            }
            break;
          }
        }
      }
    }
  } catch (e) {
    console.log("Error checking template rows: " + e.message);
  }

  return { hasRecords: false };
}