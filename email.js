/**file name: email.gs
/**
 * ============================================================================
 * DOCUMAIL PRO COMPLETE MASTER CORE SCRIPT
 * ============================================================================
 * FILE: email.gs - Email Engine and Quota Management
 * ============================================================================
 */

function executeEmailSend(selectedRows, emailConfig, templateName, tagMappings) { // 👈 ADDED tagMappings
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var allHeaders = GET_ALL_RAW_HEADERS();
  var emailColIdx = -1;

  for (var c = 0; c < allHeaders.length; c++) {
    var hName = String(allHeaders[c]).toLowerCase().trim();
    if (hName.indexOf("recipient email") !== -1) emailColIdx = c;
  }
  if (emailColIdx === -1) emailColIdx = 4;

  var statusColIdx = GET_OR_CREATE_STATUS_COLUMN(sheet, templateName);

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, allHeaders.length).getValues();
  var sent = 0;
  var errors = [];
  
  // Use manual mappings if provided, fallback to standard header-matching
  var mappings = tagMappings || {};

  for (var i = 0; i < data.length; i++) {
    var rowNum = i + 2;
    var found = false;
    for (var s = 0; s < selectedRows.length; s++) {
      if (selectedRows[s] === rowNum) {
        found = true;
        break;
      }
    }
    if (!found) continue;

    var currentStatus = data[i][statusColIdx];
    if (currentStatus && currentStatus.toString().trim() !== "" && currentStatus.toString().indexOf("Failed") === -1) {
      continue;
    }

    var recipient = data[i][emailColIdx];
    if (!recipient || recipient.indexOf("@") === -1) continue;

    var subject = emailConfig.subject || "";
    var body = emailConfig.body || "";

    // =======================================================
    // ✅ FIX: ITERATE OVER THE DICTIONARY MAPPINGS LIVE
    // =======================================================
    if (Object.keys(mappings).length > 0) {
      for (var docTag in mappings) {
        if (mappings.hasOwnProperty(docTag)) {
          var targetHeader = mappings[docTag];
          var headerIdx = allHeaders.indexOf(targetHeader);
          
          if (headerIdx !== -1) {
            var rawValue = data[i][headerIdx];
            var val = "";
            
            if (rawValue instanceof Date) {
              val = FORMAT_DATE_FOR_DISPLAY(rawValue);
            } else if (typeof rawValue === 'number') {
              val = FORMAT_NUMBER_FOR_DISPLAY(rawValue); // Evaluates custom formatting schemas!
            } else {
              val = String(rawValue || "");
            }
            
            var regex = new RegExp("\\{" + escapeRegex(docTag) + "\\}", "g");
            subject = subject.replace(regex, val);
            body = body.replace(regex, val);
          }
        }
      }
    } else {
      // Fallback behavior if no template mapping configurations exist
      for (var h = 0; h < allHeaders.length; h++) {
        var header = allHeaders[h];
        if (header) {
          var rawValue = data[i][h];
          var val = "";
          if (rawValue instanceof Date) {
            val = FORMAT_DATE_FOR_DISPLAY(rawValue);
          } else if (typeof rawValue === 'number') {
            val = FORMAT_NUMBER_FOR_DISPLAY(rawValue);
          } else {
            val = String(rawValue || "");
          }
          var regex = new RegExp("\\{" + escapeRegex(header) + "\\}", "g");
          subject = subject.replace(regex, val);
          body = body.replace(regex, val);
        }
      }
    }

    try {
      GmailApp.sendEmail(recipient, subject, "", { htmlBody: body });
      var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      sheet.getRange(rowNum, statusColIdx + 1).setValue("Sent to " + recipient + " on " + timestamp);
      sent++;
      CHECK_DAILY_QUOTA_ALERT();
    } catch (e) {
      errors.push("Row " + rowNum + ": " + e.message);
      sheet.getRange(rowNum, statusColIdx + 1).setValue("Failed: " + e.message);
    }
  }

  if (errors.length > 0) {
    return "✅ Sent " + sent + " emails.\n⚠️ Errors: " + errors.join(", ");
  }
  return "✅ Successfully sent " + sent + " emails!";
}

function executeEmailSendWithAttachments(emailConfig, folderId, templateName, tagMappings) { // 👈 ADDED tagMappings
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var allHeaders = GET_ALL_RAW_HEADERS();
    var emailColIdx = -1;

    for (var c = 0; c < allHeaders.length; c++) {
        var hName = String(allHeaders[c]).toLowerCase().trim();
        if (hName.indexOf("recipient email") !== -1) emailColIdx = c;
    }
    if (emailColIdx === -1) emailColIdx = 4;

    var statusColIdx = GET_OR_CREATE_STATUS_COLUMN(sheet, templateName);
    var mappings = tagMappings || {};

    var pdfFiles = [];
    if (folderId) {
        var folder = DriveApp.getFolderById(folderId);
        var files = folder.getFiles();
        var currentTime = new Date().getTime();
        while (files.hasNext()) {
            var file = files.next();
            var fileCreated = file.getDateCreated().getTime();
            if (file.getMimeType() === 'application/pdf' && fileCreated > currentTime - 120000) {
                pdfFiles.push(file);
            }
        }
    }

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, allHeaders.length).getValues();
    var sent = 0;
    var errors = [];

    for (var i = 0; i < data.length; i++) {
        var rowNum = i + 2;
        var existingStatus = String(data[i][statusColIdx] || "").trim();
        if (existingStatus !== "") continue;

        var recipient = data[i][emailColIdx];
        if (!recipient || recipient.indexOf("@") === -1) continue;

        var subject = emailConfig.subject || "";
        var body = emailConfig.body || "";

        // =======================================================
        // ✅ FIX: ITERATE OVER THE DICTIONARY MAPPINGS LIVE WITH PDF
        // =======================================================
        if (Object.keys(mappings).length > 0) {
          for (var docTag in mappings) {
            if (mappings.hasOwnProperty(docTag)) {
              var targetHeader = mappings[docTag];
              var headerIdx = allHeaders.indexOf(targetHeader);
              
              if (headerIdx !== -1) {
                var val = data[i][headerIdx];
                if (val instanceof Date) {
                  val = FORMAT_DATE_FOR_DISPLAY(val);
                } else if (typeof val === 'number') {
                  val = FORMAT_NUMBER_FOR_DISPLAY(val);
                } else {
                  val = String(val || "");
                }
                var regex = new RegExp("\\{" + escapeRegex(docTag) + "\\}", "g");
                subject = subject.replace(regex, val);
                body = body.replace(regex, val);
              }
            }
          }
        } else {
          for (var h = 0; h < allHeaders.length; h++) {
              var header = allHeaders[h];
              if (header) {
                  var val = data[i][h];
                  if (val instanceof Date) {
                      val = FORMAT_DATE_FOR_DISPLAY(val);
                  } else if (typeof val === 'number') {
                      val = FORMAT_NUMBER_FOR_DISPLAY(val);
                  } else {
                      val = String(val || "");
                  }
                  var regex = new RegExp("\\{" + escapeRegex(header) + "\\}", "g");
                  subject = subject.replace(regex, val);
                  body = body.replace(regex, val);
              }
          }
        }

        var mailOptions = { htmlBody: body };
        if (emailConfig.replyTo) mailOptions.replyTo = emailConfig.replyTo;
        if (emailConfig.cc) mailOptions.cc = emailConfig.cc;
        if (emailConfig.bcc) mailOptions.bcc = emailConfig.bcc;

        var attachments = [];
        for (var p = 0; p < pdfFiles.length; p++) {
            attachments.push(pdfFiles[p].getBlob());
        }
        if (attachments.length > 0) {
            mailOptions.attachments = attachments;
        }

        try {
            GmailApp.sendEmail(recipient, subject, "", mailOptions);
            var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
            sheet.getRange(rowNum, statusColIdx + 1).setValue("Sent to " + recipient + " on " + timestamp + " (with PDF)");
            sent++;
            CHECK_DAILY_QUOTA_ALERT();
        } catch (e) {
            errors.push("Row " + rowNum + ": " + e.message);
            sheet.getRange(rowNum, statusColIdx + 1).setValue("Failed: " + e.message);
        }
    }

    if (errors.length > 0) {
        return "✅ Sent " + sent + " emails (with PDF attachments).\n⚠️ Errors: " + errors.join(", ");
    }
    return "✅ Successfully sent " + sent + " emails (with PDF attachments)!";
}

// =======================================================
// ✅ FIX: PASS REFRESHED PARAM MAPPINGS VIA CONTROLLER
// =======================================================
function EXECUTE_TEMPLATE_ACTION(params) {
    var mappings = (params.config && params.config.tagMappings) ? params.config.tagMappings : {};
    
    if (params.type === "EMAIL_ONLY") {
        return executeEmailSend(params.selectedRows, params.emailConfig, params.templateName, mappings);
    }
    return "PDF generation coming soon";
}

function CHECK_DAILY_QUOTA_ALERT() {
    try {
        if (MailApp.getRemainingDailyQuota() === 10) {
            var senderInbox = Session.getActiveUser().getEmail();
            if (senderInbox) {
                MailApp.sendEmail(
                    senderInbox,
                    "⚠️ Alert: Your Daily DocuMail Pro Email Quota is Running Low!",
                    "Hello,\n\nYou have already sent your standard volume for the day. There are only 10 emails pending/remaining for today before Google limits your execution thread.\n\nPlease plan your remaining bulk campaigns accordingly.\n\nRegards,\nDocuMail Pro Safety Monitor Engine"
                );
                console.log("🚨 Quota warning email dispatched successfully to: " + senderInbox);
            }
        }
    } catch (qe) {
        console.log("Quota threshold check skipped safely: " + qe.message);
    }
}

function CHECK_SIDEBAR_REFRESH() {
  try {
    var sheetName = SpreadsheetApp.getActiveSheet().getName();
    var key = 'SIDEBAR_REFRESH_SIGNAL_KEY_' + sheetName;
    return PropertiesService.getDocumentProperties().getProperty(key) || "";
  } catch (e) {
    console.log("Error reading layout sync properties: " + e.message);
    return "";
  }
}

function SIGNAL_SIDEBAR_REFRESH() {
  try {
    var sheetName = SpreadsheetApp.getActiveSheet().getName();
    var key = 'SIDEBAR_REFRESH_SIGNAL_KEY_' + sheetName;
    PropertiesService.getDocumentProperties().setProperty(key, "REFRESH_" + new Date().getTime());
    return true;
  } catch (e) {
    return false;
  }
}