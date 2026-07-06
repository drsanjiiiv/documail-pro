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
var displayValues = sheet.getRange(2, 1, lastRow - 1, allHeaders.length).getDisplayValues();
    var processedCount = 0;
    var skippedCount = 0;

    // Find criteria column
    var criteriaColIdx = allHeaders.indexOf(payload.condField);
    if (criteriaColIdx === -1) criteriaColIdx = 0;

    // Get template file ID
    if (!payload.templateUrl) {
      return "Error: No template document selected. Please edit template and select a document.";
    }

    var templateId = payload.templateUrl.split("/d/")[1].split("/")[0];
    var templateFile = DriveApp.getFileById(templateId);

    // Get destination folder
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

      // Check if PDF already exists (Strict skip for BOTH Preview and Live Runs)
      var existingStatus = String(rowData[docStatusColIdx] || "").trim();
      if (existingStatus === "Success") {
        skippedCount++;
        continue; // Immediately drop out and evaluate the next row
      }

      // For preview mode, only process one row
      if (payload.isPreview && processedCount >= 1) break;

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

            // ======================================================================
      // 🔥 INJECTION: PROCESS CONDITIONALS BEFORE SWAPPING RAW TAGS
      // ======================================================================
      var rowDataMap = {};
      var displayDataMap = {};
      for (var h = 0; h < allHeaders.length; h++) {
        rowDataMap[allHeaders[h]] = rowData[h];
        displayDataMap[allHeaders[h]] = displayValues[i][h];
      }

      // STEP 1: Process table rows FIRST (before body.setText destroys them)
      if (typeof processConditionalTableRows === 'function') {
        Logger.log("📋 Processing table rows first...");
        processConditionalTableRows(body, rowDataMap, displayDataMap);
      } else {
        Logger.log("⚠️ processConditionalTableRows NOT found!");
      }

      // STEP 2: Process paragraph blocks
      if (typeof processConditionalBlocks === 'function') {
        Logger.log("📋 Processing paragraph blocks...");
        processConditionalBlocks(body, rowDataMap, displayDataMap);
      } else {
        Logger.log("⚠️ processConditionalBlocks NOT found!");
      }
      // ======================================================================
      // ======================================================================

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

    var mode = payload.isPreview ? "PREVIEW" : "RUN";

    // If something was processed, report success clean
    if (processedCount > 0) {
      return mode + " Completed Successfully! 🎉\n\n✅ New Documents Created: " + processedCount +
        (skippedCount > 0 ? "\n⏭️ Skipped Completed Rows: " + skippedCount : "") +
        "\n📂 Saved to Destination Folder: " + destinationFolder.getName();
    }

    if (skippedCount > 0 && processedCount === 0) {
      return "NO_ROWS_ELIGIBLE";
    }

    if (processedCount === 0) {
      return "No eligible records found matching your filter conditions.";
    }

  } catch (err) {
    return "Error: " + err.toString();
  }
}

// ==========================================
// FUNCTION: EXECUTE_DOCUMENT_MERGE_ENGINE_FOR_SINGLE_ROW Starts
// ==========================================
function EXECUTE_DOCUMENT_MERGE_ENGINE_FOR_SINGLE_ROW(payload, singleRowData, rowNum, allHeaders, singleRowDisplayData) {
  if (!singleRowData || singleRowData.length === 0) {
    return { success: false, error: "No data row provided." };
  }
  try {
    var templateUrl = payload.templateUrl;
    var templateId = templateUrl.split("/d/")[1].split("/")[0];
    var templateFile = DriveApp.getFileById(templateId);

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

    var newDocFile = templateFile.makeCopy(fileName, destinationFolder);
    var newDocId = newDocFile.getId();

    var newDoc = DocumentApp.openById(newDocId);
    var body = newDoc.getBody();

    // ======================================================================
    // 🔥 INJECTION: PROCESS CONDITIONALS BEFORE SWAPPING RAW TAGS
    // ======================================================================
    if (typeof processConditionalBlocks === 'function') {
      var rowDataMap = {};
      var displayDataMap = {};
      for (var h = 0; h < allHeaders.length; h++) {
        rowDataMap[allHeaders[h]] = rowData[h];
        displayDataMap[allHeaders[h]] = singleRowDisplayData ? singleRowDisplayData[h] : String(rowData[h] || "");
      }
      processConditionalBlocks(body, rowDataMap, displayDataMap);
      if (typeof processConditionalTableRows === 'function') {
        processConditionalTableRows(body, rowDataMap, displayDataMap);
      }
    }
    // ======================================================================

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

    var finalFileId = newDocId;
    var finalFileUrl = newDocFile.getUrl();

    if (payload.format === "PDF") {
      var pdfBlob = newDocFile.getAs('application/pdf');
      var pdfFileName = fileName + ".pdf";
      newDocFile.setTrashed(true);
      var pdfFile = destinationFolder.createFile(pdfBlob).setName(pdfFileName);
      finalFileId = pdfFile.getId();
      finalFileUrl = pdfFile.getUrl();
    }

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

function EXTRACT_TEMPLATE_TAGS_STREAM(docUrl) {
  try {
    var fileId = docUrl.split("/d/")[1].split("/")[0];
    var docInstance = DocumentApp.openById(fileId);
    var bodyText = docInstance.getBody().getText();
    var tagMatchRegex = /\{([^}]+)\}/g;
    var matches = [];
    var matchItem;

    var validHeaders = GET_LIVE_SHEET_HEADERS();
    var validHeaderSet = {};
    for (var i = 0; i < validHeaders.length; i++) {
      validHeaderSet[validHeaders[i].toLowerCase().trim()] = true;
    }

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

// ============================================================================
// PREVIEW TEMPLATE FUNCTION - PERFECT INTEGRATION WITH SKIP LOGIC
// ============================================================================
function PREVIEW_TEMPLATE(templateId) {
  try {
    var template = GET_TEMPLATE_BY_ID(templateId);
    if (!template) return { name: "Not Found", previewHtml: "<p>Template not found</p>" };

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var allHeaders = GET_ALL_RAW_HEADERS();

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return {
        name: template.name,
        previewHtml: '<div style="font-family: Roboto, sans-serif; padding: 20px;">' +
          '<h2 style="color:#c5221f;">❌ No Data Found</h2><hr><p>No data rows found in the sheet.</p>' +
          '<div style="margin-top:20px; text-align:center;"><button onclick="google.script.host.close()" style="background:#1a73e8; color:white; border:none; padding:8px 24px; border-radius:4px; cursor:pointer;">Close</button></div></div>'
      };
    }

    var emailColIdx = -1;
    for (var c = 0; c < allHeaders.length; c++) {
      var hName = String(allHeaders[c]).toLowerCase().trim();
      if (hName.indexOf("recipient email") !== -1) emailColIdx = c;
    }
    if (emailColIdx === -1) emailColIdx = 4;

    // Find the dynamic merged doc status column
    var docStatusColIdx = -1;
    for (var c = 0; c < allHeaders.length; c++) {
      if (String(allHeaders[c]).toLowerCase().indexOf("merged doc status") !== -1) {
        docStatusColIdx = c;
        break;
      }
    }

    // Dynamic Tracking Column for emails
    var emailStatusColName = "Sent Mail Status - " + template.name;
    var emailStatusColIdx = -1;
    for (var c = 0; c < allHeaders.length; c++) {
      if (String(allHeaders[c]).toLowerCase().trim() === emailStatusColName.toLowerCase()) {
        emailStatusColIdx = c;
        break;
      }
    }

    var criteriaColIdx = allHeaders.indexOf(template.config.condField);
    if (criteriaColIdx === -1) criteriaColIdx = 0;

    var data = sheet.getRange(2, 1, lastRow - 1, allHeaders.length).getValues();
    var previewRow = null;

    // Strict Sync Logic: Iterates over rows exactly as RUN behaves
    for (var i = 0; i < data.length; i++) {
      var evalCellText = String(data[i][criteriaColIdx] || "").trim();
      var isCriteriaMatch = false;

      // 1. Evaluate general filter condition matching rules
      if (template.config.condOperator === "NOT_EMPTY" && evalCellText !== "") {
        isCriteriaMatch = true;
      } else if (template.config.condOperator === "CONTAINS" && evalCellText.toLowerCase().indexOf(String(template.config.condValue).toLowerCase().trim()) !== -1) {
        isCriteriaMatch = true;
      }

      // 2. Strict Skip Check: Bypass based on specific template types
      if (isCriteriaMatch) {
        if (template.type === "PDF_ONLY") {
          if (docStatusColIdx !== -1 && String(data[i][docStatusColIdx] || "").trim() === "Success") {
            continue;
          }
        } else if (template.type === "EMAIL_ONLY") {
          if (emailStatusColIdx !== -1 && String(data[i][emailStatusColIdx] || "").trim() !== "") {
            continue;
          }
        } else if (template.type === "BOTH") {
          var pdfDone = (docStatusColIdx !== -1 && String(data[i][docStatusColIdx] || "").trim() === "Success");
          var emailDone = (emailStatusColIdx !== -1 && String(data[i][emailStatusColIdx] || "").trim() !== "");
          if (pdfDone && emailDone) {
            continue;
          }
        }

        // Found the proper un-processed record row match
        previewRow = data[i];
        break;
      }
    }

    if (!previewRow) {
  return {
    name: template.name,
    previewHtml: '<div style="font-family: Roboto, sans-serif; padding: 24px;">' +
      '<h2 style="color: #f2994a; margin: 0 0 12px 0;">ℹ️ No Rows Eligible</h2>' +
      '<hr style="border: none; border-top: 1px solid #e8eaed; margin: 12px 0;">' +
      '<div style="background: #fff3cd; padding: 16px; border-radius: 6px; margin: 12px 0; border-left: 4px solid #f2994a; color: #856404; line-height: 1.6;">' +
      'All rows matching your filter conditions have already been merged successfully with a status of <strong>Success</strong>.<br><br>' +
      'There are no fresh pending records available to preview layout results against.' +
      '</div>' +
      '<hr style="border: none; border-top: 1px solid #e8eaed; margin: 16px 0;">' +
      '<div style="margin-top: 20px; text-align: center;">' +
      '<button onclick="google.script.host.close()" style="background: #1a73e8; color: white; border: none; padding: 10px 28px; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: 500;">Close</button>' +
      '</div>' +
      '</div>'
  };
}

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

    var recipientEmail = previewRow[emailColIdx] || "example@email.com";
    var emailSubject = template.emailConfig?.subject || "";
    var emailBody = template.emailConfig?.body || "";

    for (var h = 0; h < allHeaders.length; h++) {
      var header = allHeaders[h];
      if (header) {
        var val = previewRow[h];
        if (val instanceof Date) {
          val = FORMAT_DATE_FOR_DISPLAY(val);
        } else if (typeof val === 'number') {
          val = FORMAT_NUMBER_FOR_DISPLAY(val);
        } else {
          val = String(val || "");
        }
        var regex = new RegExp("\\{" + escapeRegex(header) + "\\}", "g");
        emailSubject = emailSubject.replace(regex, val);
        emailBody = emailBody.replace(regex, val);
      }
    }

    var previewHtml = '<div style="font-family: Roboto, sans-serif; padding: 20px;">';
    previewHtml += '<h2 style="color:#1a73e8;">✅ Preview Generated Successfully!</h2><hr>';
    previewHtml += '<p><strong>📄 Template Name:</strong> ' + escapeHtml(template.name) + '</p>';
    previewHtml += '<p><strong>📋 Template Type:</strong> ' + (template.type === "PDF_ONLY" ? "PDF Only" : template.type === "EMAIL_ONLY" ? "Email Only" : "PDF & Email") + '</p>';

    if (template.type === "PDF_ONLY" || template.type === "BOTH") {
      var fileUrl = '';
      var fileName = '';
      try {
        var folderId = template.config?.folderDestination;
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
            } catch (e) { }
          }
        }

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
      } catch (e) { }

      previewHtml += '<div style="background:#e8f0fe; padding:12px; border-radius:6px; margin:10px 0;">';
      previewHtml += '<h3 style="color:#1a73e8; margin:0 0 10px 0;">📄 Document Preview</h3>';
      previewHtml += '<p><strong>Document Name:</strong> ' + escapeHtml(displayFileName) + '</p>';
      previewHtml += '<p><strong>Output Format:</strong> ' + (template.config?.format === "PDF" ? "PDF Document" : "Editable Google Doc") + '</p>';

      if (fileUrl) {
        previewHtml += '<div style="background:#e6f4ea; padding:12px; border-radius:6px; margin:10px 0; border-left:4px solid #137333;">';
        previewHtml += '<p style="margin:0; color:#137333;"><strong>✅ Preview File Generated!</strong></p>';
        previewHtml += '<p style="margin:5px 0;"><strong>File:</strong> ' + escapeHtml(fileName) + '</p>';
        previewHtml += '<p style="margin:5px 0;"><a href="' + fileUrl + '" target="_blank">📂 Click here to open the preview file</a></p>';
        previewHtml += '<p style="margin:8px 0 0 0; color:#c5221f; font-weight:500; border-top:1px solid #dadce0; padding-top:8px;">⚠️ This is a Preview Only file - not for production use</p>';
        previewHtml += '</div>';
      } else {
        previewHtml += '<p><em>`✅ Preview file generated with PROV- prefix. Check your Google Drive folder.`</em></p>';
      }
      previewHtml += '</div>';
    }

    if (template.type === "EMAIL_ONLY" || template.type === "BOTH") {
      previewHtml += '<div style="background:#e6f4ea; padding:12px; border-radius:6px; margin:10px 0;">';
      previewHtml += '<h3 style="color:#137333; margin:0 0 10px 0;">✉️ Email Preview</h3>';
      previewHtml += '<p><strong>To:</strong> ' + escapeHtml(recipientEmail) + '</p>';
      previewHtml += '<p><strong>Subject:</strong> ' + escapeHtml(emailSubject) + '</p>';
      previewHtml += '<div style="border:1px solid #dadce0; padding:12px; border-radius:6px; margin-top:12px; background:#ffffff;"><strong>Email Body:</strong><br><br>' + emailBody + '</div>';
      previewHtml += '</div>';
    }

    previewHtml += '<hr><div style="margin-top:20px; text-align:center;"><button onclick="google.script.host.close()" style="background:#1a73e8; color:white; border:none; padding:8px 24px; border-radius:4px; cursor:pointer;">Close</button></div></div>';
    return { name: template.name, previewHtml: previewHtml };

  } catch (e) {
    return { name: "Error", previewHtml: "<p>ERROR: " + e.toString() + "</p>" };
  }
}

function RUN_TEMPLATE(templateId) {
  var template = GET_TEMPLATE_BY_ID(templateId);
  if (!template) return "Template not found";

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return "No data rows found. Please add data to your sheet before running the template.";
  }

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

    var criteriaColIdx = allHeaders.indexOf(template.config.condField);
    if (criteriaColIdx === -1) criteriaColIdx = 0;

    var statusColIdx = GET_OR_CREATE_STATUS_COLUMN(sheet, template.name);
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, allHeaders.length).getValues();
    var selectedRows = [];

    for (var i = 0; i < data.length; i++) {
      // 1. Evaluate custom criteria filters first
      var evalCellText = String(data[i][criteriaColIdx] || "").trim();
      var isMatch = false;
      if (template.config.condOperator === "NOT_EMPTY" && evalCellText !== "") {
        isMatch = true;
      } else if (template.config.condOperator === "CONTAINS" && evalCellText.toLowerCase().indexOf(String(template.config.condValue).toLowerCase().trim()) !== -1) {
        isMatch = true;
      }

      if (!isMatch) continue;

      // 2. Ensure record hasn't already been processed
      var statusValue = String(data[i][statusColIdx] || "").trim();
      if (statusValue === "") {
        selectedRows.push(i + 2);
      }
    }

    if (selectedRows.length === 0) return "NO_ROWS_ELIGIBLE";
    return executeEmailSend(selectedRows, template.emailConfig, template.name);
  }

  if (template.type === "BOTH") {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var allHeaders = GET_ALL_RAW_HEADERS();

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

    var criteriaColIdx = allHeaders.indexOf(template.config.condField);
    if (criteriaColIdx === -1) criteriaColIdx = 0;

    var emailStatusColIdx = GET_OR_CREATE_STATUS_COLUMN(sheet, template.name);

    var emailColIdx = -1;
    for (var c = 0; c < allHeaders.length; c++) {
      var hName = String(allHeaders[c]).toLowerCase().trim();
      if (hName.indexOf("recipient email") !== -1) emailColIdx = c;
    }
    if (emailColIdx === -1) emailColIdx = 4;

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, allHeaders.length).getValues();
    var displayData = sheet.getRange(2, 1, sheet.getLastRow() - 1, allHeaders.length).getDisplayValues();
    var processedCount = 0;
    var pdfSkippedCount = 0;
    var emailSentCount = 0;
    var errors = [];

    for (var i = 0; i < data.length; i++) {
      var rowNum = i + 2;

      var evalCellText = String(data[i][criteriaColIdx] || "").trim();
      var isMatch = false;
      if (template.config.condOperator === "NOT_EMPTY" && evalCellText !== "") {
        isMatch = true;
      } else if (template.config.condOperator === "CONTAINS" && evalCellText.toLowerCase().indexOf(String(template.config.condValue).toLowerCase().trim()) !== -1) {
        isMatch = true;
      }

      if (!isMatch) continue;

      var existingDocStatus = String(data[i][docStatusColIdx] || "").trim();
      var existingDocId = String(data[i][docIdColIdx] || "").trim();
      var existingDocUrl = String(data[i][docUrlColIdx] || "").trim();

      var pdfGenerated = false;
      var pdfFileId = null;
      var pdfFileUrl = null;

      if (existingDocStatus === "Success" && existingDocId !== "") {
        pdfGenerated = true;
        pdfFileId = existingDocId;
        pdfFileUrl = existingDocUrl;
        pdfSkippedCount++;
      } else {
        var config = JSON.parse(JSON.stringify(template.config));
        config.isPreview = false;

        var singleRowData = [data[i]];
        var singleRowDisplayData = displayData[i];
        var result = EXECUTE_DOCUMENT_MERGE_ENGINE_FOR_SINGLE_ROW(config, singleRowData, rowNum, allHeaders, singleRowDisplayData);

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

      var emailStatus = String(data[i][emailStatusColIdx] || "").trim();
      if (emailStatus !== "") continue;

      var recipient = data[i][emailColIdx];
      if (!recipient || recipient.indexOf("@") === -1) continue;

      var subject = template.emailConfig.subject || "";
      var body = template.emailConfig.body || "";

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

      var mailOptions = { htmlBody: body };
      if (template.emailConfig.replyTo) mailOptions.replyTo = template.emailConfig.replyTo;
      if (template.emailConfig.cc) mailOptions.cc = template.emailConfig.cc;
      if (template.emailConfig.bcc) mailOptions.bcc = template.emailConfig.bcc;

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
        CHECK_DAILY_QUOTA_ALERT();
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

    return resultMsg || "NO_ROWS_ELIGIBLE";
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

  var docStatusColIdx = -1;
  for (var c = 0; c < allHeaders.length; c++) {
    if (String(allHeaders[c]).toLowerCase().indexOf("merged doc status") !== -1) {
      docStatusColIdx = c;
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

    if (docStatusColIdx !== -1) {
      var existingDocStatus = String(data[i][docStatusColIdx] || "").trim();
      if (existingDocStatus === "Success") continue;
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

// ============================================================
// 📋 CONDITIONAL PROCESSING FUNCTIONS - FIXED
// ============================================================

/**
 * Main evaluation router - SINGLE VERSION (remove duplicates)
 */
function evaluateCondition(varValue, operator, targetValue) {
  var currentVal = varValue !== undefined && varValue !== null ? String(varValue).trim() : "";
  var criteriaVal = targetValue !== undefined && targetValue !== null ? String(targetValue).trim() : "";
  var currentValClean = currentVal.replace(/[^0-9.\-]/g, "");
  var criteriaValClean = criteriaVal.replace(/[^0-9.\-]/g, "");
  var isNumeric = currentValClean !== "" && criteriaValClean !== "" && !isNaN(currentValClean) && !isNaN(criteriaValClean);
  operator = operator.toLowerCase().trim();

  Logger.log("  evaluateCondition: currentVal='" + currentVal + "', operator='" + operator + "', criteriaVal='" + criteriaVal + "'");

  if (isNumeric) {
    var numCurrent = parseFloat(currentValClean);
    var numCriteria = parseFloat(criteriaValClean);
    switch (operator) {
      case '==': case '=': return numCurrent === numCriteria;
      case '!=': case '<>': return numCurrent !== numCriteria;
      case '>=': return numCurrent >= numCriteria;
      case '<=': return numCurrent <= numCriteria;
      case '>': return numCurrent > numCriteria;
      case '<': return numCurrent < numCriteria;
      case 'contains': return currentVal.toLowerCase().indexOf(criteriaVal.toLowerCase()) !== -1;
      default: return false;
    }
  }

  switch (operator) {
    case '==': case '=': return currentVal.toLowerCase() === criteriaVal.toLowerCase();
    case '!=': case '<>': return currentVal.toLowerCase() !== criteriaVal.toLowerCase();
    case '>=': return currentVal.toLowerCase() >= criteriaVal.toLowerCase();
    case '<=': return currentVal.toLowerCase() <= criteriaVal.toLowerCase();
    case '>': return currentVal.toLowerCase() > criteriaVal.toLowerCase();
    case '<': return currentVal.toLowerCase() < criteriaVal.toLowerCase();
    case 'contains': return currentVal.toLowerCase().indexOf(criteriaVal.toLowerCase()) !== -1;
    default: return false;
  }
}

/**
 * Process conditional blocks - <<If: ... >> ... <<EndIf>>
 * FIXED: Properly handles both TRUE and FALSE conditions
 */
function processConditionalBlocks(body, rowDataMap, displayDataMap) {
  try {
    Logger.log("🔍 processConditionalBlocks: Starting...");

    var paragraphs = body.getParagraphs();
    var texts = paragraphs.map(function (p) { return p.getText(); });
    var fullText = texts.join("\n");

    var cum = [];
    var running = 0;
    for (var i = 0; i < texts.length; i++) {
      cum.push(running);
      running += texts[i].length + 1;
    }
    function locate(flatOffset) {
      for (var j = cum.length - 1; j >= 0; j--) {
        if (cum[j] <= flatOffset) return { paraIndex: j, localOffset: flatOffset - cum[j] };
      }
      return { paraIndex: 0, localOffset: flatOffset };
    }

    var blockRegex = /<<If:\s*([^>]+)>>([\s\S]*?)<<EndIf>>/gi;
    var blocks = [];
    var match;
    while ((match = blockRegex.exec(fullText)) !== null) {
      var startFlat = match.index;
      blocks.push({
        startFlat: startFlat,
        endFlat: startFlat + match[0].length,
        ifTagEndFlat: startFlat + match[0].indexOf('>>') + 2,
        endIfTagStartFlat: startFlat + match[0].lastIndexOf('<<EndIf>>'),
        condition: match[1].trim()
      });
    }
    Logger.log("  Found " + blocks.length + " blocks");

    var cRegex = /^\s*([^=!><]+?)\s*([=!><]=?|contains)\s*['"“]([^'"”']+)['"”']\s*$/i;

    for (var b = blocks.length - 1; b >= 0; b--) {
      var blk = blocks[b];
      var conditionMet = false;
      var cMatch = cRegex.exec(blk.condition);

      if (cMatch) {
        var leftSide = cMatch[1].trim();
        var operator = cMatch[2].trim();
        var targetValue = cMatch[3].trim();

        var liveValue = leftSide;
        if (leftSide.charAt(0) === '{' && leftSide.charAt(leftSide.length - 1) === '}') {
          var varName = leftSide.slice(1, -1);
          liveValue = rowDataMap.hasOwnProperty(varName) ? rowDataMap[varName] : leftSide;
        }

        conditionMet = evaluateCondition(liveValue, operator, targetValue);
      } else {
        var eqIndex = blk.condition.indexOf("==");
        if (eqIndex !== -1) {
          var leftPart = blk.condition.substring(0, eqIndex).trim();
          var rightPart = blk.condition.substring(eqIndex + 2).trim().replace(/['"“”']/g, '');
          conditionMet = (leftPart === rightPart);
        }
      }
      Logger.log("  Block: '" + blk.condition + "' -> " + conditionMet);

      if (conditionMet) {
        removeFlatRange(paragraphs, blk.endIfTagStartFlat, blk.endIfTagStartFlat + "<<EndIf>>".length, locate);
        removeFlatRange(paragraphs, blk.startFlat, blk.ifTagEndFlat, locate);
      } else {
        removeFlatRange(paragraphs, blk.startFlat, blk.endFlat, locate);
      }
    }

    // FINAL: substitute remaining {VARIABLES} using the sheet's own display text
    Logger.log("📋 Replacing display variables...");
    var dispMap = displayDataMap || {};
    for (var key in rowDataMap) {
      var placeholder = "{" + key + "}";
      var replacementValue = dispMap.hasOwnProperty(key) ? dispMap[key] : String(rowDataMap[key] || "");
      body.replaceText(escapeRegexString(placeholder), replacementValue);
    }

    Logger.log("✅ processConditionalBlocks: Completed");

  } catch (err) {
    Logger.log("❌ Error in processConditionalBlocks: " + err.toString());
  }
}

/**
 * White space handling
 * FIXED: Properly handles blank white sapces leftover
 */

function removeFlatRange(paragraphs, flatStart, flatEnd, locate) {
  var startLoc = locate(flatStart);
  var endLoc = locate(flatEnd);
  var startPara = startLoc.paraIndex;
  var endPara = endLoc.paraIndex;

  if (startPara === endPara) {
    var pText = paragraphs[startPara].editAsText();
    var localStart = startLoc.localOffset;
    var localEnd = endLoc.localOffset - 1;
    if (localEnd >= localStart && pText.getText().length > 0) {
      pText.deleteText(localStart, Math.min(localEnd, pText.getText().length - 1));
    }
    if (pText.getText().trim() === "") {
      try { paragraphs[startPara].removeFromParent(); } catch (e) { /* already gone or last paragraph */ }
    }
  } else {
    var pStart = paragraphs[startPara];
    var pStartText = pStart.editAsText();
    if (pStartText.getText().length > 0 && startLoc.localOffset <= pStartText.getText().length - 1) {
      pStartText.deleteText(startLoc.localOffset, pStartText.getText().length - 1);
    }

    var pEnd = paragraphs[endPara];
    var pEndText = pEnd.editAsText();
    var localEndOffset = endLoc.localOffset - 1;
    if (localEndOffset >= 0 && pEndText.getText().length > 0) {
      pEndText.deleteText(0, Math.min(localEndOffset, pEndText.getText().length - 1));
    }

    for (var k = endPara - 1; k > startPara; k--) {
      try { paragraphs[k].removeFromParent(); } catch (e) { /* already gone */ }
    }
    try { if (pStart.editAsText().getText().trim() === "") pStart.removeFromParent(); } catch (e) {}
    try { if (pEnd.editAsText().getText().trim() === "") pEnd.removeFromParent(); } catch (e) {}
  }
}

/**
 * Process conditional table rows - <<RowIf: ... >> 
 * FIXED: Properly handles both TRUE and FALSE conditions in tables
 */
function processConditionalTableRows(body, rowDataMap, displayDataMap) {
  try {
    Logger.log("🔍 processConditionalTableRows: Starting...");
    var tables = body.getTables();
    if (tables.length === 0) {
      Logger.log("  No tables found, skipping");
      return;
    }

    var dispMap = displayDataMap || {};
    var cRegex = /^\s*([^=!><]+?)\s*([=!><]=?|contains)\s*['"“]([^'"”']+)['"”']\s*$/i;

    for (var t = 0; t < tables.length; t++) {
      var table = tables[t];
      var rowsToRemove = [];
      var numRows = table.getNumRows();
      Logger.log("  Table has " + numRows + " rows");

      for (var r = 0; r < numRows; r++) {
        var row = table.getRow(r);
        var rowText = row.getText();
        var match = rowText.match(/<<RowIf:\s*([^>]+)>>/i);

        if (match) {
          var condition = match[1].trim();
          var conditionMet = false;
          var cMatch = cRegex.exec(condition);

          if (cMatch) {
            var leftSide = cMatch[1].trim();
            var operator = cMatch[2].trim();
            var targetValue = cMatch[3].trim();

            var liveValue = leftSide;
            if (leftSide.charAt(0) === '{' && leftSide.charAt(leftSide.length - 1) === '}') {
              var varName = leftSide.slice(1, -1);
              liveValue = rowDataMap.hasOwnProperty(varName) ? rowDataMap[varName] : leftSide;
            }
            conditionMet = evaluateCondition(liveValue, operator, targetValue);
          }

          if (!conditionMet) {
            rowsToRemove.push(r);
            Logger.log("    Row " + (r + 1) + " marked for REMOVAL (condition FALSE)");
          } else {
            Logger.log("    Row " + (r + 1) + " marked to KEEP - removing tag");
            var numCells = row.getNumCells();
            for (var c = 0; c < numCells; c++) {
              var cell = row.getCell(c);
              var cellText = cell.getText();
              var cleanedText = cellText.replace(/<<RowIf:\s*[^>]+>>/i, '').trim();
              if (cleanedText !== cellText) cell.setText(cleanedText);
            }
          }
        }
      }

      rowsToRemove.sort(function (a, b) { return b - a; });
      for (var i = 0; i < rowsToRemove.length; i++) {
        table.removeRow(rowsToRemove[i]);
      }

      // Substitute remaining {VARIABLES} in surviving rows using display text
      for (var key in rowDataMap) {
        var placeholder = "{" + key + "}";
        var replacementValue = dispMap.hasOwnProperty(key) ? dispMap[key] : String(rowDataMap[key] || "");
        table.replaceText(escapeRegexString(placeholder), replacementValue);
      }
    }

    Logger.log("✅ processConditionalTableRows: Completed");

  } catch (err) {
    Logger.log("❌ Error in processConditionalTableRows: " + err.toString());
  }
}



/**
 * Helper to escape special regular expression characters
 */
function escapeRegexString(str) {
  if (!str) return '';
  return str.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
}