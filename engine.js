/**
 * ============================================================================
 * DOCUMAIL PRO COMPLETE MASTER CORE SCRIPT
 * ============================================================================
 * FILE: engine.gs - Document Merge Engine
 * ============================================================================
 */

// ==========================================
// FUNCTION: EXECUTE_DOCUMENT_MERGE_ENGINE Starts
// ==========================================
function EXECUTE_DOCUMENT_MERGE_ENGINE(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var allHeaders = GET_ALL_RAW_HEADERS();

    // Find column indices
    var docStatusColIdx = -1;
    var docIdColIdx = -1;
    var docUrlColIdx = -1;

    for (var c = 0; c < allHeaders.length; c++) {
      var hName = String(allHeaders[c]).toLowerCase();
      if (hName.indexOf("merged doc status") !== -1) docStatusColIdx = c;
      if (hName.indexOf("merged doc id") !== -1) docIdColIdx = c;
      if (hName.indexOf("merged doc url") !== -1) docUrlColIdx = c;
    }

    if (docStatusColIdx === -1) docStatusColIdx = 5;
    if (docIdColIdx === -1) docIdColIdx = 6;
    if (docUrlColIdx === -1) docUrlColIdx = 7;

    // TEMPORARILY REMOVE PROTECTION
    var tempRemovedProtections = [];
    if (!payload.isPreview) {
      var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (var p = 0; p < protections.length; p++) {
        var protection = protections[p];
        var range = protection.getRange();
        var colStart = range.getColumn();
        var colEnd = colStart + range.getWidth() - 1;
        if (colStart <= docUrlColIdx + 1 && colEnd >= docStatusColIdx + 1) {
          tempRemovedProtections.push({
            protection: protection,
            rangeA1: range.getA1Notation()
          });
          protection.remove();
        }
      }
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return "No data rows found.";

    var dataValues = sheet.getRange(2, 1, lastRow - 1, allHeaders.length).getValues();
    var processedCount = 0;

    // Find criteria column
    var criteriaColIdx = allHeaders.indexOf(payload.condField);
    if (criteriaColIdx === -1) criteriaColIdx = 0;

    // Get template file ID
    if (!payload.templateUrl) {
      return "Error: No template document selected. Please edit template and select a document.";
    }

    var templateId = payload.templateUrl.split("/d/")[1].split("/")[0];
    var templateFile = DriveApp.getFileById(templateId);

    // Get destination folder (SAFE VERSION)
    var destinationFolder = DriveApp.getRootFolder();
    if (payload.folderDestination && payload.folderDestination !== "") {
      try {
        var folderUrl = payload.folderDestination;
        var folderId = null;
        if (folderUrl.indexOf("/folders/") > -1) {
          folderId = folderUrl.split("/folders/")[1];
        } else if (folderUrl.indexOf("id=") > -1) {
          folderId = folderUrl.split("id=")[1];
        } else {
          folderId = folderUrl;
        }
        if (folderId) {
          destinationFolder = DriveApp.getFolderById(folderId);
        }
      } catch (e) {
        // Use root folder if error
        destinationFolder = DriveApp.getRootFolder();
      }
    }

    // Process each row
    for (var i = 0; i < dataValues.length; i++) {
      var rowNum = i + 2;
      var rowData = dataValues[i];

      // Check filter condition
      var evalCellText = String(rowData[criteriaColIdx] || "").trim();
      var isTargetEligible = false;

      if (payload.condOperator === "NOT_EMPTY" && evalCellText !== "") {
        isTargetEligible = true;
      } else if (payload.condOperator === "CONTAINS" && evalCellText.toLowerCase().indexOf(String(payload.condValue).toLowerCase().trim()) !== -1) {
        isTargetEligible = true;
      }

      if (!isTargetEligible) continue;

      // For preview mode, only process one row
      if (payload.isPreview && processedCount >= 1) break;

      // Check if PDF already exists (skip if already generated)
      if (!payload.isPreview) {
        var existingStatus = String(rowData[docStatusColIdx] || "").trim();
        if (existingStatus === "Success") {
          processedCount++;
          continue;
        }
      }

      // Create filename from pattern
      var fileName = payload.namePattern || "Generated Document";
      for (var h = 0; h < allHeaders.length; h++) {
        var header = allHeaders[h];
        if (header) {
          var regex = new RegExp("\\{" + escapeRegex(header) + "\\}", "g");
          fileName = fileName.replace(regex, String(rowData[h] || ""));
        }
      }

      fileName = fileName.replace(/[\\/:*?"<>|]/g, "");

      if (payload.isPreview) {
        fileName = "PROV-" + fileName;
      }

      // Copy the template file
      var newDocFile = templateFile.makeCopy(fileName, destinationFolder);
      var newDocId = newDocFile.getId();

      // Open and replace tags
      var newDoc = DocumentApp.openById(newDocId);
      var body = newDoc.getBody();

      // Replace tags using mapped values - PRESERVE FORMATTING
if (payload.tagMappings) {
  for (var docTag in payload.tagMappings) {
    if (payload.tagMappings.hasOwnProperty(docTag)) {
      var sheetColumn = payload.tagMappings[docTag];
      var replacementValue = "";
      var mappedColIdx = allHeaders.indexOf(sheetColumn);
      if (mappedColIdx !== -1) {
        var rawValue = rowData[mappedColIdx];
        if (rawValue instanceof Date) {
          replacementValue = FORMAT_DATE_FOR_DISPLAY(rawValue);
        } else if (typeof rawValue === 'number') {
          replacementValue = FORMAT_NUMBER_FOR_DISPLAY(rawValue);
        } else {
          replacementValue = String(rawValue || "");
        }
      }

      var searchPattern = "{" + docTag + "}";
      var foundElement = body.findText(searchPattern);

      while (foundElement) {
        var element = foundElement.getElement();
        if (element.editAsText) {
          var startOffset = foundElement.getStartOffset();
          var endOffsetInclusive = foundElement.getEndOffsetInclusive();
          
          // Get the formatting from the existing text
          var text = element.asText();
          var fontFamily = text.getFontFamily(startOffset);
          var fontSize = text.getFontSize(startOffset);
          var isBold = text.isBold(startOffset);
          var isItalic = text.isItalic(startOffset);
          var isUnderline = text.isUnderline(startOffset);
          var foregroundColor = text.getForegroundColor(startOffset);
          
          // Delete the tag
          text.deleteText(startOffset, endOffsetInclusive);
          
          // Insert the replacement with same formatting
          text.insertText(startOffset, replacementValue);
          
          // Apply the original formatting to the inserted text
          var newEndOffset = startOffset + replacementValue.length - 1;
          if (fontFamily) text.setFontFamily(startOffset, newEndOffset, fontFamily);
          if (fontSize) text.setFontSize(startOffset, newEndOffset, fontSize);
          if (isBold !== null) text.setBold(startOffset, newEndOffset, isBold);
          if (isItalic !== null) text.setItalic(startOffset, newEndOffset, isItalic);
          if (isUnderline !== null) text.setUnderline(startOffset, newEndOffset, isUnderline);
          if (foregroundColor) text.setForegroundColor(startOffset, newEndOffset, foregroundColor);
        }
        foundElement = body.findText(searchPattern, foundElement);
      }
    }
  }
}

      newDoc.saveAndClose();

      var finalFileId = newDocId;
      var finalFileUrl = newDocFile.getUrl();

      // Convert to PDF if format is PDF
      if (payload.format === "PDF") {
        var pdfBlob = newDocFile.getAs('application/pdf');
        var pdfFileName = fileName + ".pdf";
        newDocFile.setTrashed(true);
        var pdfFile = destinationFolder.createFile(pdfBlob).setName(pdfFileName);
        finalFileId = pdfFile.getId();
        finalFileUrl = pdfFile.getUrl();
      }

      // Update sheet (only for non-preview)
      if (!payload.isPreview) {
        sheet.getRange(rowNum, docStatusColIdx + 1).setValue("Success");
        sheet.getRange(rowNum, docIdColIdx + 1).setValue(finalFileId);
        sheet.getRange(rowNum, docUrlColIdx + 1).setValue(finalFileUrl);
      }

      processedCount++;
    }

    // RESTORE PROTECTIONS
    if (!payload.isPreview) {
      for (var r = 0; r < tempRemovedProtections.length; r++) {
        var prot = tempRemovedProtections[r];
        var newProtection = sheet.getRange(prot.rangeA1).protect();
        newProtection.setDescription(prot.protection.getDescription());
        var editors = prot.protection.getEditors();
        newProtection.addEditors(editors);
        if (prot.protection.canDomainEdit()) {
          newProtection.setDomainEdit(true);
        }
      }
    }

    var mode = payload.isPreview ? "PREVIEW" : "PRODUCTION";
    return mode + " completed! Processed " + processedCount + " record(s). " + (payload.format === "PDF" ? "PDF files" : "Google Docs") + " saved to: " + destinationFolder.getName();

  } catch (err) {
    return "Error: " + err.toString();
  }
}
// ==========================================
// FUNCTION: EXECUTE_DOCUMENT_MERGE_ENGINE Ends
// ==========================================

// ==========================================
// FUNCTION: EXECUTE_DOCUMENT_MERGE_ENGINE_FOR_SINGLE_ROW Starts
// ==========================================
function EXECUTE_DOCUMENT_MERGE_ENGINE_FOR_SINGLE_ROW(payload, singleRowData, rowNum, allHeaders) {
  try {
    // Get template file ID
    var templateUrl = payload.templateUrl;
    var templateId = templateUrl.split("/d/")[1].split("/")[0];
    var templateFile = DriveApp.getFileById(templateId);

    // Get destination folder (with safe check)
    var destinationFolder = DriveApp.getRootFolder();
    if (payload.folderDestination && typeof payload.folderDestination === "string" && payload.folderDestination !== "") {
      try {
        var folderId = payload.folderDestination.split("/folders/")[1] || payload.folderDestination.split("id=")[1];
        if (!folderId) folderId = payload.folderDestination;
        destinationFolder = DriveApp.getFolderById(folderId);
      } catch (e) {
        destinationFolder = DriveApp.getRootFolder();
      }
    }

    var rowData = singleRowData[0];

    // Create filename from pattern
    var fileName = payload.namePattern || "Generated Document";
    for (var h = 0; h < allHeaders.length; h++) {
      var header = allHeaders[h];
      if (header) {
        var regex = new RegExp("\\{" + escapeRegex(header) + "\\}", "g");
        fileName = fileName.replace(regex, String(rowData[h] || ""));
      }
    }

    fileName = fileName.replace(/[\\/:*?"<>|]/g, "");

    if (payload.isPreview) {
      fileName = "PROV-" + fileName;
    }

    // Copy the template file
    var newDocFile = templateFile.makeCopy(fileName, destinationFolder);
    var newDocId = newDocFile.getId();

    // Open and replace tags
    var newDoc = DocumentApp.openById(newDocId);
    var body = newDoc.getBody();

    // Replace tags using mapped values - PRESERVE FORMATTING
    if (payload.tagMappings) {
      for (var docTag in payload.tagMappings) {
        if (payload.tagMappings.hasOwnProperty(docTag)) {
          var sheetColumn = payload.tagMappings[docTag];
          var replacementValue = "";
          var mappedColIdx = allHeaders.indexOf(sheetColumn);
          if (mappedColIdx !== -1) {
            var rawValue = rowData[mappedColIdx];
            if (rawValue instanceof Date) {
              replacementValue = FORMAT_DATE_FOR_DISPLAY(rawValue);
            } else if (typeof rawValue === 'number') {
              replacementValue = FORMAT_NUMBER_FOR_DISPLAY(rawValue);
            } else {
              replacementValue = String(rawValue || "");
            }
          }

          var searchPattern = "{" + docTag + "}";
          var foundElement = body.findText(searchPattern);

          while (foundElement) {
            var element = foundElement.getElement();
            if (element.editAsText) {
              var startOffset = foundElement.getStartOffset();
              var endOffsetInclusive = foundElement.getEndOffsetInclusive();
              
              var text = element.asText();
              var fontFamily = text.getFontFamily(startOffset);
              var fontSize = text.getFontSize(startOffset);
              var isBold = text.isBold(startOffset);
              var isItalic = text.isItalic(startOffset);
              var isUnderline = text.isUnderline(startOffset);
              var foregroundColor = text.getForegroundColor(startOffset);
              
              text.deleteText(startOffset, endOffsetInclusive);
              text.insertText(startOffset, replacementValue);
              
              var newEndOffset = startOffset + replacementValue.length - 1;
              if (fontFamily) text.setFontFamily(startOffset, newEndOffset, fontFamily);
              if (fontSize) text.setFontSize(startOffset, newEndOffset, fontSize);
              if (isBold !== null) text.setBold(startOffset, newEndOffset, isBold);
              if (isItalic !== null) text.setItalic(startOffset, newEndOffset, isItalic);
              if (isUnderline !== null) text.setUnderline(startOffset, newEndOffset, isUnderline);
              if (foregroundColor) text.setForegroundColor(startOffset, newEndOffset, foregroundColor);
            }
            foundElement = body.findText(searchPattern, foundElement);
          }
        }
      }
    }

    // =======================================================
    // CONVERT TO PDF (No watermark)
    // =======================================================
    var finalFileId = newDocId;
    var finalFileUrl = newDocFile.getUrl();

    // Convert to PDF if format is PDF
    if (payload.format === "PDF") {
      var pdfBlob = newDocFile.getAs('application/pdf');
      var pdfFileName = fileName + ".pdf";
      newDocFile.setTrashed(true);
      var pdfFile = destinationFolder.createFile(pdfBlob).setName(pdfFileName);
      finalFileId = pdfFile.getId();
      finalFileUrl = pdfFile.getUrl();
    }

    // Update sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var allSheetHeaders = GET_ALL_RAW_HEADERS();

    var docStatusColIdx = -1;
    var docIdColIdx = -1;
    var docUrlColIdx = -1;

    for (var c = 0; c < allSheetHeaders.length; c++) {
      var hName = String(allSheetHeaders[c]).toLowerCase();
      if (hName.indexOf("merged doc status") !== -1) docStatusColIdx = c;
      if (hName.indexOf("merged doc id") !== -1) docIdColIdx = c;
      if (hName.indexOf("merged doc url") !== -1) docUrlColIdx = c;
    }

    if (docStatusColIdx !== -1) sheet.getRange(rowNum, docStatusColIdx + 1).setValue("Success");
    if (docIdColIdx !== -1) sheet.getRange(rowNum, docIdColIdx + 1).setValue(finalFileId);
    if (docUrlColIdx !== -1) sheet.getRange(rowNum, docUrlColIdx + 1).setValue(finalFileUrl);

    return { success: true, fileId: finalFileId, fileUrl: finalFileUrl };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
// ==========================================
// FUNCTION: EXECUTE_DOCUMENT_MERGE_ENGINE_FOR_SINGLE_ROW Ends
// ==========================================

function EXTRACT_TEMPLATE_TAGS_STREAM(docUrl) {
  try {
    var fileId = docUrl.split("/d/")[1].split("/")[0];
    var docInstance = DocumentApp.openById(fileId);
    var bodyText = docInstance.getBody().getText();
    var tagMatchRegex = /\{([^}]+)\}/g;
    var matches = [];
    var matchItem;
    while ((matchItem = tagMatchRegex.exec(bodyText)) !== null) {
      var cleanToken = matchItem[1].trim();
      if (matches.indexOf(cleanToken) === -1) {
        matches.push(cleanToken);
      }
    }
    return matches;
  } catch (err) {
    return [];
  }
}

// ==========================================
// PREVIEW TEMPLATE FUNCTION - Shows both PDF and Email preview for BOTH type
// ==========================================
// ==========================================
// PREVIEW TEMPLATE FUNCTION - Shows both PDF and Email preview for BOTH type
// ==========================================
function PREVIEW_TEMPLATE(templateId) {
  try {
    var template = GET_TEMPLATE_BY_ID(templateId);
    if (!template) return { name: "Not Found", previewHtml: "<p>Template not found</p>" };

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var allHeaders = GET_ALL_RAW_HEADERS();

    var emailColIdx = -1;
    for (var c = 0; c < allHeaders.length; c++) {
      var hName = String(allHeaders[c]).toLowerCase().trim();
      if (hName.indexOf("recipient email") !== -1) emailColIdx = c;
    }
    if (emailColIdx === -1) emailColIdx = 4;

    var criteriaColIdx = allHeaders.indexOf(template.config.condField);
    if (criteriaColIdx === -1) criteriaColIdx = 0;

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, allHeaders.length).getValues();
    var previewRow = null;

    for (var i = 0; i < data.length; i++) {
      var evalCellText = String(data[i][criteriaColIdx] || "").trim();
      var isMatch = false;
      if (template.config.condOperator === "NOT_EMPTY" && evalCellText !== "") {
        isMatch = true;
      } else if (template.config.condOperator === "CONTAINS" && evalCellText.toLowerCase().indexOf(String(template.config.condValue).toLowerCase().trim()) !== -1) {
        isMatch = true;
      }
      if (isMatch) {
        previewRow = data[i];
        break;
      }
    }

    if (!previewRow) {
      return { name: template.name, previewHtml: "<p>No rows match your filter conditions.</p>" };
    }

    // Build the preview HTML - show file name with tags replaced
    var displayFileName = template.config?.namePattern || "Document";
    for (var h = 0; h < allHeaders.length; h++) {
      var header = allHeaders[h];
      if (header && previewRow[h] !== undefined && previewRow[h] !== null) {
        var val = previewRow[h];
        if (val instanceof Date) {
          val = FORMAT_DATE_FOR_DISPLAY(val);
        } else {
          val = String(val);
        }
        var regex = new RegExp("\\{" + escapeRegex(header) + "\\}", "g");
        displayFileName = displayFileName.replace(regex, val);
      }
    }

    // Email preview with formatted dates
    var recipientEmail = previewRow[emailColIdx] || "example@email.com";
    var emailSubject = template.emailConfig?.subject || "";
    var emailBody = template.emailConfig?.body || "";

    for (var h = 0; h < allHeaders.length; h++) {
      var header = allHeaders[h];
      if (header) {
        var val = previewRow[h];
        if (val instanceof Date) {
          val = FORMAT_DATE_FOR_DISPLAY(val);
        } else {
          val = String(val || "");
        }
        var regex = new RegExp("\\{" + escapeRegex(header) + "\\}", "g");
        emailSubject = emailSubject.replace(regex, val);
        emailBody = emailBody.replace(regex, val);
      }
    }

    var previewHtml = '<div style="font-family: Roboto, sans-serif; padding: 20px;">';
    previewHtml += '<h2 style="color:#1a73e8;">✅ Preview Generated Successfully!</h2>';
    previewHtml += '<hr>';
    previewHtml += '<p><strong>📄 Template Name:</strong> ' + escapeHtml(template.name) + '</p>';
    previewHtml += '<p><strong>📋 Template Type:</strong> ' + (template.type === "PDF_ONLY" ? "PDF Only" : template.type === "EMAIL_ONLY" ? "Email Only" : "PDF & Email") + '</p>';

    if (template.type === "PDF_ONLY" || template.type === "BOTH") {
  // Try to find the generated file
var fileUrl = '';
var fileName = '';
try {
  var folderId = template.config?.folderDestination;
  
  // Search in specified folder
  if (folderId) {
    var folderIdExtracted = folderId.split("/folders/")[1] || folderId.split("id=")[1];
    if (folderIdExtracted) {
      try {
        var folder = DriveApp.getFolderById(folderIdExtracted);
        var files = folder.getFiles();
        var currentTime = new Date().getTime();
        var latestFile = null;
        var latestTime = 0;
        while (files.hasNext()) {
          var file = files.next();
          var fileCreated = file.getDateCreated().getTime();
          if (fileCreated > currentTime - 60000 && file.getName().indexOf("PROV-") !== -1) {
            if (fileCreated > latestTime) {
              latestTime = fileCreated;
              latestFile = file;
            }
          }
        }
        if (latestFile) {
          fileUrl = latestFile.getUrl();
          fileName = latestFile.getName();
        }
      } catch(e) {}
    }
  }
  
  // =======================================================
  // SEARCH ROOT DRIVE IF NOT FOUND
  // =======================================================
  if (!fileUrl) {
    var rootFiles = DriveApp.getFiles();
    var currentTime = new Date().getTime();
    var latestFile = null;
    var latestTime = 0;
    while (rootFiles.hasNext()) {
      var file = rootFiles.next();
      var fileCreated = file.getDateCreated().getTime();
      if (fileCreated > currentTime - 60000 && file.getName().indexOf("PROV-") !== -1) {
        if (fileCreated > latestTime) {
          latestTime = fileCreated;
          latestFile = file;
        }
      }
    }
    if (latestFile) {
      fileUrl = latestFile.getUrl();
      fileName = latestFile.getName();
    }
  }
} catch(e) {}

  previewHtml += '<div style="background:#e8f0fe; padding:12px; border-radius:6px; margin:10px 0;">';
  previewHtml += '<h3 style="color:#1a73e8; margin:0 0 10px 0;">📄 Document Preview</h3>';
  previewHtml += '<p><strong>Document Name:</strong> ' + escapeHtml(displayFileName) + '</p>';
  previewHtml += '<p><strong>Output Format:</strong> ' + (template.config?.format === "PDF" ? "PDF Document" : "Editable Google Doc") + '</p>';
  
  if (fileUrl) {
    previewHtml += '<div style="background:#e6f4ea; padding:12px; border-radius:6px; margin:10px 0; border-left:4px solid #137333;">';
    previewHtml += '<p style="margin:0; color:#137333;"><strong>✅ Preview File Generated!</strong></p>';
    previewHtml += '<p style="margin:5px 0;"><strong>File:</strong> ' + escapeHtml(fileName) + '</p>';
    previewHtml += '<p style="margin:5px 0;"><a href="' + fileUrl + '" target="_blank">📂 Click here to open the preview file</a></p>';
    previewHtml += '<p style="margin:5px 0; color:#5f6368; font-size:12px;">⚠️ This file is for PREVIEW ONLY. DO NOT USE FOR PRODUCTION</p>';
    previewHtml += '</div>';
  } else {
    previewHtml += '<p><em>✅ Preview file generated with PROV- prefix. Check your Google Drive folder.</em></p>';
  }
  
  previewHtml += '</div>';
}

    if (template.type === "EMAIL_ONLY" || template.type === "BOTH") {
      previewHtml += '<div style="background:#e6f4ea; padding:12px; border-radius:6px; margin:10px 0;">';
      previewHtml += '<h3 style="color:#137333; margin:0 0 10px 0;">✉️ Email Preview</h3>';
      previewHtml += '<p><strong>To:</strong> ' + escapeHtml(recipientEmail) + '</p>';
      previewHtml += '<p><strong>Subject:</strong> ' + escapeHtml(emailSubject) + '</p>';
      previewHtml += '<div style="border:1px solid #dadce0; padding:12px; border-radius:6px; margin-top:12px; background:#ffffff;">';
      previewHtml += '<strong>Email Body:</strong><br><br>' + emailBody;
      previewHtml += '</div>';
      previewHtml += '<p style="margin-top:12px; color:#5f6368;"><em>Note: This is a preview only. No email was sent.</em></p>';
      previewHtml += '</div>';
    }

    previewHtml += '<hr>';
    previewHtml += '<div style="margin-top:20px; text-align:center;">';
    previewHtml += '<button onclick="google.script.host.close()" style="background:#1a73e8; color:white; border:none; padding:8px 24px; border-radius:4px; cursor:pointer;">Close</button>';
    previewHtml += '</div>';
    previewHtml += '</div>';

    return { name: template.name, previewHtml: previewHtml };

  } catch (e) {
    return { name: "Error", previewHtml: "<p>ERROR: " + e.toString() + "</p>" };
  }
}

function RUN_TEMPLATE(templateId) {
  var template = GET_TEMPLATE_BY_ID(templateId);
  if (!template) return "Template not found";

  if (template.type === "PDF_ONLY") {
    var config = JSON.parse(JSON.stringify(template.config));
    config.isPreview = false;
    var result = EXECUTE_DOCUMENT_MERGE_ENGINE(config);
    return result;
  }

  if (template.type === "EMAIL_ONLY") {
  var allHeaders = GET_ALL_RAW_HEADERS();
  var emailColIdx = -1;

  for (var c = 0; c < allHeaders.length; c++) {
    var hName = String(allHeaders[c]).toLowerCase().trim();
    if (hName.indexOf("recipient email") !== -1) emailColIdx = c;
  }
  if (emailColIdx === -1) emailColIdx = 4;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var statusColIdx = GET_OR_CREATE_STATUS_COLUMN(sheet, template.name);

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, allHeaders.length).getValues();
  var selectedRows = [];

  for (var i = 0; i < data.length; i++) {
    var statusValue = String(data[i][statusColIdx] || "").trim();
    if (statusValue === "") {
      selectedRows.push(i + 2);
    }
  }   

  if (selectedRows.length === 0) return "No eligible records found for template: " + template.name;
  return executeEmailSend(selectedRows, template.emailConfig, template.name);
}

  if (template.type === "BOTH") {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var allHeaders = GET_ALL_RAW_HEADERS();

    // Find PDF status column indices
    var docStatusColIdx = -1;
    var docIdColIdx = -1;
    var docUrlColIdx = -1;

    for (var c = 0; c < allHeaders.length; c++) {
      var hName = String(allHeaders[c]).toLowerCase();
      if (hName.indexOf("merged doc status") !== -1) docStatusColIdx = c;
      if (hName.indexOf("merged doc id") !== -1) docIdColIdx = c;
      if (hName.indexOf("merged doc url") !== -1) docUrlColIdx = c;
    }

    if (docStatusColIdx === -1) docStatusColIdx = 5;
    if (docIdColIdx === -1) docIdColIdx = 6;
    if (docUrlColIdx === -1) docUrlColIdx = 7;

    // Find filter column
    var criteriaColIdx = allHeaders.indexOf(template.config.condField);
    if (criteriaColIdx === -1) criteriaColIdx = 0;

    // Find email status column
    var emailStatusColIdx = GET_OR_CREATE_STATUS_COLUMN(sheet, template.name);

    // Find email recipient column
    var emailColIdx = -1;
    for (var c = 0; c < allHeaders.length; c++) {
      var hName = String(allHeaders[c]).toLowerCase().trim();
      if (hName.indexOf("recipient email") !== -1) emailColIdx = c;
    }
    if (emailColIdx === -1) emailColIdx = 4;

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, allHeaders.length).getValues();
    var processedCount = 0;
    var pdfSkippedCount = 0;
    var emailSentCount = 0;
    var errors = [];

    for (var i = 0; i < data.length; i++) {
      var rowNum = i + 2;

      // Check filter condition
      var evalCellText = String(data[i][criteriaColIdx] || "").trim();
      var isMatch = false;
      if (template.config.condOperator === "NOT_EMPTY" && evalCellText !== "") {
        isMatch = true;
      } else if (template.config.condOperator === "CONTAINS" && evalCellText.toLowerCase().indexOf(String(template.config.condValue).toLowerCase().trim()) !== -1) {
        isMatch = true;
      }

      if (!isMatch) continue;

      // Check if PDF already exists
      var existingDocStatus = String(data[i][docStatusColIdx] || "").trim();
      var existingDocId = String(data[i][docIdColIdx] || "").trim();
      var existingDocUrl = String(data[i][docUrlColIdx] || "").trim();

      var pdfGenerated = false;
      var pdfFileId = null;
      var pdfFileUrl = null;

      if (existingDocStatus === "Success" && existingDocId !== "") {
        // PDF already exists - reuse it
        pdfGenerated = true;
        pdfFileId = existingDocId;
        pdfFileUrl = existingDocUrl;
        pdfSkippedCount++;
      } else {
        // Generate PDF
        var config = JSON.parse(JSON.stringify(template.config));
        config.isPreview = false;

        // Create a single-row payload for this specific row
        var singleRowData = [data[i]];
        var result = EXECUTE_DOCUMENT_MERGE_ENGINE_FOR_SINGLE_ROW(config, singleRowData, rowNum, allHeaders);

        if (result.success) {
          pdfGenerated = true;
          pdfFileId = result.fileId;
          pdfFileUrl = result.fileUrl;
          processedCount++;
        } else {
          errors.push("Row " + rowNum + ": PDF generation failed - " + result.error);
          continue;
        }
      }

      // Send email with PDF attachment
      var emailStatus = String(data[i][emailStatusColIdx] || "").trim();
      if (emailStatus !== "") {
        continue; // Email already sent for this template
      }

      var recipient = data[i][emailColIdx];
      if (!recipient || recipient.indexOf("@") === -1) continue;

      var subject = template.emailConfig.subject || "";
      var body = template.emailConfig.body || "";

      for (var h = 0; h < allHeaders.length; h++) {
        var header = allHeaders[h];
        if (header) {
          var val = String(data[i][h] || "");
          var regex = new RegExp("\\{" + escapeRegex(header) + "\\}", "g");
          subject = subject.replace(regex, val);
          body = body.replace(regex, val);
        }
      }

      var mailOptions = { htmlBody: body };
      if (template.emailConfig.replyTo) mailOptions.replyTo = template.emailConfig.replyTo;
      if (template.emailConfig.cc) mailOptions.cc = template.emailConfig.cc;
      if (template.emailConfig.bcc) mailOptions.bcc = template.emailConfig.bcc;

      // Attach the PDF file
      if (pdfFileId) {
        try {
          var pdfFile = DriveApp.getFileById(pdfFileId);
          mailOptions.attachments = [pdfFile.getBlob()];
        } catch (e) {
          errors.push("Row " + rowNum + ": Could not attach PDF - " + e.message);
          continue;
        }
      }

      try {
        GmailApp.sendEmail(recipient, subject, "", mailOptions);
        var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        sheet.getRange(rowNum, emailStatusColIdx + 1).setValue("Sent to " + recipient + " on " + timestamp + " (with PDF)");
        emailSentCount++;
        CHECK_DAILY_QUOTA_ALERT(); // <-- Single line injection handles it!
      } catch (e) {
        errors.push("Row " + rowNum + ": " + e.message);
        sheet.getRange(rowNum, emailStatusColIdx + 1).setValue("Failed: " + e.message);
      }
    }

    var resultMsg = "";
    if (processedCount > 0) resultMsg += "✅ New PDFs generated: " + processedCount + "\n";
    if (pdfSkippedCount > 0) resultMsg += "⏭️ Existing PDFs reused: " + pdfSkippedCount + "\n";
    if (emailSentCount > 0) resultMsg += "✅ Emails sent: " + emailSentCount + "\n";
    if (errors.length > 0) resultMsg += "⚠️ Errors: " + errors.join(", ");

    return resultMsg || "No eligible records found.";
  }

  return "Unknown template type";
}

function GET_ELIGIBLE_RECORDS_FOR_TEMPLATE(params) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var allHeaders = GET_ALL_RAW_HEADERS();
  var emailColIdx = -1;

  for (var c = 0; c < allHeaders.length; c++) {
    var hName = String(allHeaders[c]).toLowerCase().trim();
    if (hName.indexOf("recipient email") !== -1) emailColIdx = c;
  }
  if (emailColIdx === -1) emailColIdx = 4;

  // Look for dynamic status column for this template
  var statusColIdx = -1;
  if (params.templateName) {
    var statusColName = "Sent Mail Status - " + params.templateName;
    for (var c = 0; c < allHeaders.length; c++) {
      if (String(allHeaders[c]).toLowerCase().trim() === statusColName.toLowerCase()) {
        statusColIdx = c;
        break;
      }
    }
  }

  var criteriaColIdx = -1;
  for (var c = 0; c < allHeaders.length; c++) {
    if (allHeaders[c] === params.condField) {
      criteriaColIdx = c;
      break;
    }
  }
  if (criteriaColIdx === -1) criteriaColIdx = 0;

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, allHeaders.length).getValues();
  var eligible = [];

  for (var i = 0; i < values.length; i++) {
    var rowNum = i + 2;

    // Skip if already sent for this template
    var alreadySent = false;
    if (statusColIdx !== -1) {
      var existingStatus = String(values[i][statusColIdx] || "").trim();
      if (existingStatus !== "") alreadySent = true;
    }
    if (alreadySent) continue;

    var cell = String(values[i][criteriaColIdx] || "").trim();
    var match = false;
    if (params.condOperator === "NOT_EMPTY") {
      match = (cell !== "");
    } else if (params.condOperator === "CONTAINS") {
      match = (cell.toLowerCase().indexOf(params.condValue.toLowerCase()) !== -1);
    }
    if (match) {
      eligible.push({
        rowNum: rowNum,
        email: values[i][emailColIdx] || "",
        refData: values[i][0] || ""
      });
    }
  }
  return { eligibleRows: eligible };
}

function GET_RECORDS_PREVIEW_PAYLOAD_WITH_TEMPLATE(templateName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var allHeaders = GET_ALL_RAW_HEADERS();
  var template = GET_TEMPLATE_BY_NAME(templateName);

  var emailColIdx = -1;
  for (var c = 0; c < allHeaders.length; c++) {
    if (String(allHeaders[c]).toLowerCase().indexOf("recipient email") !== -1) {
      emailColIdx = c;
      break;
    }
  }
  if (emailColIdx === -1) emailColIdx = 4;

  var statusColName = "Sent Mail Status - " + templateName;
  var statusColIdx = -1;
  for (var c = 0; c < allHeaders.length; c++) {
    if (String(allHeaders[c]).toLowerCase().trim() === statusColName.toLowerCase()) {
      statusColIdx = c;
      break;
    }
  }

  var criteriaColIdx = allHeaders.indexOf(template.config.condField);
  if (criteriaColIdx === -1) criteriaColIdx = 0;

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, allHeaders.length).getValues();
  var eligibleRows = [];

  for (var i = 0; i < data.length; i++) {
    var rowNum = i + 2;

    if (statusColIdx !== -1) {
      var existingStatus = String(data[i][statusColIdx] || "").trim();
      if (existingStatus !== "") continue;
    }

    var cell = String(data[i][criteriaColIdx] || "").trim();
    var match = false;
    if (template.config.condOperator === "NOT_EMPTY") {
      match = (cell !== "");
    } else if (template.config.condOperator === "CONTAINS") {
      match = (cell.toLowerCase().indexOf(template.config.condValue.toLowerCase()) !== -1);
    }

    if (match) {
      var rowData = {};
      for (var h = 0; h < allHeaders.length; h++) {
        rowData[allHeaders[h]] = data[i][h];
      }
      eligibleRows.push({
        rowNum: rowNum,
        email: data[i][emailColIdx] || "",
        refData: data[i][0] || "",
        rowData: rowData,
        selected: true
      });
    }
  }

  var remainingQuota = MailApp.getRemainingDailyQuota();

  return {
    eligibleRows: eligibleRows,
    emailTemplate: template.emailConfig,
    remainingQuota: remainingQuota,
    sheetHeaders: GET_LIVE_SHEET_HEADERS()
  };
}