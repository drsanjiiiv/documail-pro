/** file name: code.gs
/**
 * ============================================================================
 * DOCUMAIL PRO COMPLETE MASTER CORE SCRIPT
 * ============================================================================
 * FILE: code.gs - Main UI, Menu, Sidebar, and Dialog Functions
 * ============================================================================
 */

// ==========================================
// SPREADSHEET PLATFORM - ONOPEN (CLEAN)
// ==========================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('DocuMail Pro Platform')
    .addItem('Initialize Sheet Structural Layout', 'GENERATE_DOCUMAIL_TEMPLATE')
    .addItem('Create Doc Template From Sheet', 'CREATE_DOCUMENT_TEMPLATE_MENU')
    .addItem('Open Engine Workspace Control Sidebar', 'INITIALIZE_ADDON_SIDEBAR')
    .addToUi();
}

// ==========================================
// Function (INITIALIZE_ADDON_SIDEBAR) Starts
// ==========================================
function INITIALIZE_ADDON_SIDEBAR() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();

  // =======================================================
  // CHECK IF SYSTEM COLUMNS EXIST
  // =======================================================
  var lastCol = sheet.getLastColumn();
  var hasSystemColumns = false;

  if (lastCol > 0) {
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    for (var i = 0; i < headers.length; i++) {
      var hName = String(headers[i]).toLowerCase().trim();
      if (hName.indexOf("merged doc status") !== -1 ||
        hName.indexOf("recipient email") !== -1) {
        hasSystemColumns = true;
        break;
      }
    }
  }

  // =======================================================
  // IF SYSTEM COLUMNS DON'T EXIST, PROMPT USER
  // =======================================================
  if (!hasSystemColumns) {
    var response = ui.alert(
      "⚠️ DocuMail PRO Not Initialized",
      "DocuMail PRO system columns don't exist in this sheet.\n\n" +
      "Please initialize the sheet structure first to continue.\n\n" +
      "Do you want to initialize it now?",
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      // Call the initialization function
      GENERATE_DOCUMAIL_TEMPLATE();
    }
    return;
  }

  // =======================================================
  // SYSTEM COLUMNS EXIST - PROCEED WITH SIDEBAR
  // =======================================================
  var html = HtmlService.createTemplateFromFile('SidebarView');
  var sidebarUi = html.evaluate()
    .setTitle("DocuMail Pro")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(sidebarUi);
}

// ==========================================
// FUNCTION: CREATE_DOCUMENT_TEMPLATE_MENU
// Copies master template, stores context via Drive metadata, opens cleanly
// ==========================================
function CREATE_DOCUMENT_TEMPLATE_MENU() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var sheetName = sheet.getName();

    // 1. Resolve Parent Directory Framework (Get Grandparent Folder)
    var ssFile = DriveApp.getFileById(ss.getId());
    var parentFolders = ssFile.getParents();
    var currentFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();

    // Move one level above current sheet's folder
    var grandParentFolders = currentFolder.getParents();
    var targetParentFolder = grandParentFolders.hasNext() ? grandParentFolders.next() : DriveApp.getRootFolder();

    var folderName = "DocuMail PRO Templates";
    var templateFolder = null;
    var existingFolders = targetParentFolder.getFoldersByName(folderName);

    if (existingFolders.hasNext()) {
      templateFolder = existingFolders.next();
    } else {
      templateFolder = targetParentFolder.createFolder(folderName);
    }

    // 2. Locate the Master Template Document with the Scripts
    var masterName = "DocuMailPro Master Doc Template";
    var masterFiles = DriveApp.getFilesByName(masterName);
    
    if (!masterFiles.hasNext()) {
      throw new Error("Could not find '" + masterName + "' in your Google Drive. Please ensure the master template file exists and is named precisely.");
    }
    var masterFile = masterFiles.next();

    // 3. Clone the Master File (Clones the file and scripts perfectly into target folder)
    var docFile = masterFile.makeCopy("DocTemplate for " + sheetName, templateFolder);
    var docId = docFile.getId();
    
    // =================================================================
    // SECURE STORAGE: Write the Source Sheet ID directly into the file's description metadata
    // =================================================================
    docFile.setDescription(ss.getId());

    // 4. Set Onboarding Text On Canvas
    var doc = DocumentApp.openById(docId);
    var body = doc.getBody();
    body.clear();
    
    var titleParagraph = body.appendParagraph("📄 DocuMail Pro Master Template Canvas");
    titleParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    
    var descParagraph = body.appendParagraph("\nPlease use the menu item to start designing your automation layout:\n\n👉 Go to: DocuMail Pro > Start Dynamic Doc Template");
    descParagraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);

    doc.saveAndClose();
    var docUrl = doc.getUrl(); // Clean default URL structure without query strings

    // 5. Force Browser to Open the Cloned Document Tab Directly
    var htmlContent = '<html><script>window.open("' + docUrl + '", "_blank");google.script.host.close();</script></html>';
    var interfaceOutput = HtmlService.createHtmlOutput(htmlContent)
        .setWidth(10)
        .setHeight(10);
        
    SpreadsheetApp.getUi().showModalDialog(interfaceOutput, "Launching Template Canvas...");

  } catch (e) {
    SpreadsheetApp.getUi().alert("Error launching template: " + e.message);
  }
}

// ==========================================
// UI Functions
// ==========================================
function OPEN_DOCUMENT_WIZARD_MODAL(templateId) {
  var htmlLayout = HtmlService.createTemplateFromFile('DocConfigView');

  // CRITICAL FIX: Pass templateId to the HTML template
  if (templateId) {
    htmlLayout.existingTemplateId = templateId;
  } else {
    htmlLayout.existingTemplateId = null;
  }

  var dialogWindow = htmlLayout.evaluate().setWidth(960).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(dialogWindow, templateId ? "Edit Template" : "Create New Template");
}

function OPEN_CENTER_TEMPLATE_MODAL() {
  var htmlLayout = HtmlService.createTemplateFromFile('EmailComposerView');
  var dialogWindow = htmlLayout.evaluate().setWidth(980).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(dialogWindow, "Build & Automate Email Templates");
}

function OPEN_RECORDS_PREVIEW_WINDOW(templateName) {
  var htmlLayout = HtmlService.createTemplateFromFile('RecordsPreviewView');
  htmlLayout.templateName = templateName;
  var dialogWindow = htmlLayout.evaluate().setWidth(1000).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(dialogWindow, "Confirm Send Records & Daily Limits Check");
}

function SHOW_PREVIEW_DIALOG(templateId) {
  // =======================================================
  // SHOW LOADING POPUP FIRST
  // =======================================================
  var loadingHtml = '<div style="font-family: Roboto, sans-serif; padding: 40px; text-align: center;">';
  loadingHtml += '<div style="font-size: 48px; margin-bottom: 20px;">⏳</div>';
  loadingHtml += '<h2 style="color: #1a73e8;">Generating Preview...</h2>';
  loadingHtml += '<div style="margin: 20px auto; width: 40px; height: 40px; border: 4px solid #f3f3f3; border-top: 4px solid #1a73e8; border-radius: 50%; animation: spin 1s linear infinite;"></div>';
  loadingHtml += '<p style="color: #5f6368;">Please wait, this may take a few seconds.</p>';
  loadingHtml += '<style>@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }</style>';
  loadingHtml += '</div>';

  var loadingDialog = HtmlService.createHtmlOutput(loadingHtml)
    .setWidth(650)
    .setHeight(550)
    .setTitle('Processing...');
  SpreadsheetApp.getUi().showModalDialog(loadingDialog, 'Processing...');

  // =======================================================
  // PROCESS IN BACKGROUND
  // =======================================================
  try {
    var template = GET_TEMPLATE_BY_ID(templateId);
    if (template && (template.type === "PDF_ONLY" || template.type === "BOTH")) {
      var config = JSON.parse(JSON.stringify(template.config));
      config.isPreview = true;

      // FIX: ENSURE TAG MAPPINGS EXIST
      if (!config.tagMappings || Object.keys(config.tagMappings).length === 0) {
        var allHeaders = GET_ALL_RAW_HEADERS();
        var tags = EXTRACT_TEMPLATE_TAGS_STREAM(config.templateUrl);
        config.tagMappings = {};
        for (var i = 0; i < tags.length; i++) {
          if (allHeaders.indexOf(tags[i]) !== -1) {
            config.tagMappings[tags[i]] = tags[i];
          }
        }
      }

      EXECUTE_DOCUMENT_MERGE_ENGINE(config);
    }
  } catch (e) {
    // Continue - show preview even if file generation fails
  }

  // =======================================================
  // SHOW PREVIEW POPUP
  // =======================================================
  var result = PREVIEW_TEMPLATE(templateId);
  var htmlOutput = HtmlService.createHtmlOutput(result.previewHtml)
    .setWidth(650)
    .setHeight(550)
    .setTitle('Preview: ' + result.name);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Preview: ' + result.name);
}

// =======================================================
// SHOW RUN DIALOG
// =======================================================

// =======================================================
// SHOW RUN DIALOG
// =======================================================

function SHOW_RUN_DIALOG(templateId) {
  // =======================================================
  // SHOW LOADING POPUP FIRST
  // =======================================================
  var loadingHtml = '<div style="font-family: Roboto, sans-serif; padding: 40px; text-align: center;">';
  loadingHtml += '<div style="font-size: 48px; margin-bottom: 20px;">⏳</div>';
  loadingHtml += '<h2 style="color: #1a73e8;">Running Template...</h2>';
  loadingHtml += '<div style="margin: 20px auto; width: 40px; height: 40px; border: 4px solid #f3f3f3; border-top: 4px solid #1a73e8; border-radius: 50%; animation: spin 1s linear infinite;"></div>';
  loadingHtml += '<p style="color: #5f6368;">Please wait, processing your request.</p>';
  loadingHtml += '<style>@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }</style>';
  loadingHtml += '</div>';

  var loadingDialog = HtmlService.createHtmlOutput(loadingHtml)
    .setWidth(650)
    .setHeight(550)
    .setTitle('Processing...');
  SpreadsheetApp.getUi().showModalDialog(loadingDialog, 'Processing...');

  // =======================================================
  // PROCESS IN BACKGROUND
  // =======================================================
  var template = GET_TEMPLATE_BY_ID(templateId);

  if (!template) {
    var errorHtml = '<div style="font-family: Roboto, sans-serif; padding: 20px;">';
    errorHtml += '<h2 style="color:#c5221f;">❌ Template Not Found</h2>';
    errorHtml += '<hr>';
    errorHtml += '<p>Template ID: ' + templateId + '</p>';
    errorHtml += '<p>Please delete this template and recreate it.</p>';
    errorHtml += '<div style="margin-top:20px; text-align:center;">';
    errorHtml += '<button onclick="google.script.host.close()" style="background:#1a73e8; color:white; border:none; padding:8px 24px; border-radius:4px; cursor:pointer;">Close</button>';
    errorHtml += '</div></div>';
    var errorOutput = HtmlService.createHtmlOutput(errorHtml).setWidth(550).setHeight(250);
    SpreadsheetApp.getUi().showModalDialog(errorOutput, 'Error');
    return;
  }

  try {
    // Process mapping configurations for document exports safely
    if (template.type === "PDF_ONLY" || template.type === "BOTH") {
      var config = JSON.parse(JSON.stringify(template.config));
      config.isPreview = false; // Mark this as a live production run loop

      // Fallback auto-mapping logic if dictionary configurations are empty
      if (!config.tagMappings) {
        config.tagMappings = {};
      }

      if (Object.keys(config.tagMappings).length === 0) {
        var allHeaders = GET_ALL_RAW_HEADERS();
        var tags = EXTRACT_TEMPLATE_TAGS_STREAM(config.templateUrl);
        for (var i = 0; i < tags.length; i++) {
          if (allHeaders.indexOf(tags[i]) !== -1) {
            config.tagMappings[tags[i]] = tags[i];
          }
        }

        // Update master template storage if fallback details were written
        template.config.tagMappings = config.tagMappings;
        if (typeof SAVE_TEMPLATE === 'function') {
          SAVE_TEMPLATE(template);
        }
      }
      
      // 🛑 EXTRA RUN REMOVED FROM HERE TO PREVENT DOUBLE EXECUTION
    }
  } catch (e) {
    Logger.log("Execution error inside SHOW_RUN_DIALOG: " + e.toString());

    var exceptionHtml = '<div style="font-family: Roboto, sans-serif; padding: 20px;">';
    exceptionHtml += '<h2 style="color:#c5221f;">❌ Execution Failed</h2>';
    exceptionHtml += '<hr><p>' + e.message + '</p></div>';
    var exceptionOutput = HtmlService.createHtmlOutput(exceptionHtml).setWidth(550).setHeight(250);
    SpreadsheetApp.getUi().showModalDialog(exceptionOutput, 'Error');
    return;
  }

  // =======================================================
  // GET QUOTA & RUN TEMPLATE (ONCE)
  // =======================================================
  var remainingQuota = MailApp.getRemainingDailyQuota();
  var totalQuota = 100;
  var usedQuota = totalQuota - remainingQuota;

  // This call cleanly handles both PDF_ONLY and BOTH workflows perfectly!
  var result = RUN_TEMPLATE(templateId);

  // =======================================================
  // CHECK IF RESULT IS AN ERROR OR SKIP WARNING
  // =======================================================
  var isError = false;
  var errorKeywords = ["No data rows found", "Error:", "Template not found", "Failed", "No rows match", "Execution Aborted"];
  for (var i = 0; i < errorKeywords.length; i++) {
    if (result && result.indexOf(errorKeywords[i]) !== -1) {
      isError = true;
      break;
    }
  }

  // Detect genuine skip status
  var isNoRowsEligible = (typeof result === 'string' && (result === "NO_ROWS_ELIGIBLE" || result.indexOf("All eligible rows already have PDFs generated cleanly") !== -1));

  // =======================================================
  // GET FILE URLS (FOR PDF)
  // =======================================================
  var fileUrls = [];
  if ((template.type === "PDF_ONLY" || template.type === "BOTH") && !isNoRowsEligible && !isError) {
    var folderId = template.config?.folderDestination;
    if (folderId) {
      var folderIdExtracted = folderId.split("/folders/")[1] || folderId.split("id=")[1];
      if (folderIdExtracted) {
        var folder = DriveApp.getFolderById(folderIdExtracted);
        var files = folder.getFiles();
        var latestFiles = [];
        var currentTime = new Date().getTime();

        while (files.hasNext()) {
          var file = files.next();
          var fileCreated = file.getDateCreated().getTime();
          if (fileCreated > currentTime - 120000) {
            latestFiles.push(file);
          }
        }
        latestFiles.sort(function (a, b) {
          return b.getDateCreated().getTime() - a.getDateCreated().getTime();
        });
        for (var i = 0; i < latestFiles.length; i++) {
          fileUrls.push({ name: latestFiles[i].getName(), url: latestFiles[i].getUrl() });
        }
      }
    }
  }

  // =======================================================
  // SHOW RESULT POPUP
  // =======================================================
  var html = '<div style="font-family: Roboto, sans-serif; padding: 20px;">';

  if (isError) {
    html += '<h2 style="color:#c5221f;">❌ Execution Failed</h2>';
  } else if (isNoRowsEligible) {
    html += '<h2 style="color:#f2994a;">ℹ️ No Rows Eligible</h2>';
  } else {
    html += '<h2 style="color:#1a73e8;">✅ Execution Completed!</h2>';
  }

  html += '<hr>';
  html += '<p><strong>📄 Template:</strong> ' + escapeHtml(template.name) + '</p>';
  html += '<p><strong>📋 Type:</strong> ' + (template.type === "PDF_ONLY" ? "PDF Only" : template.type === "EMAIL_ONLY" ? "Email Only" : "PDF & Email") + '</p>';

  if ((template.type === "EMAIL_ONLY" || template.type === "BOTH") && !isNoRowsEligible) {
    html += '<div style="background:#fce8e6; padding:10px; border-radius:6px; margin:10px 0;">';
    html += '<p style="margin:0;"><strong>📊 Daily Email Quota:</strong></p>';
    html += '<p style="margin:5px 0;"><strong>✅ Remaining:</strong> ' + remainingQuota + ' emails</p>';
    html += '<p style="margin:5px 0;"><strong>📤 Used Today:</strong> ' + usedQuota + ' emails</p>';
    html += '<p style="margin:5px 0; color:#5f6368; font-size:12px;">📅 Resets every 24 hours</p>';
    html += '</div>';
  }

  // =======================================================
  // SHOW RESULT DISPLAY CONTAINERS
  // =======================================================
  if (isError) {
    html += '<div style="background:#fce8e6; padding:12px; border-radius:6px; margin:10px 0; border-left:4px solid #c5221f;">';
    html += '<strong style="color:#c5221f;">📝 Error:</strong><br>';
    html += '<span style="color:#c5221f;">' + String(result).replace(/\n/g, '<br>') + '</span>';
    html += '</div>';
  } else if (isNoRowsEligible) {
    html += '<div style="background:#fff3cd; padding:12px; border-radius:6px; margin:10px 0; border-left:4px solid #f2994a; color: #856404;">';
    html += 'All rows matching your filter conditions have already been merged successfully with a status of <strong>Success</strong>.<br><br>';
    html += 'There are no fresh pending records available to preview layout results against.';
    html += '</div>';
  } else {
    html += '<div style="background:#e6f4ea; padding:12px; border-radius:6px; margin:10px 0;">';
    html += '<strong>📝 Result:</strong><br>' + String(result).replace(/\n/g, '<br>');
    html += '</div>';
  }

  // Show file links if any
  if (fileUrls.length > 0) {
    html += '<div style="background:#e8f0fe; padding:12px; border-radius:6px; margin:10px 0;">';
    html += '<strong>📂 Generated Files:</strong><br>';
    for (var i = 0; i < fileUrls.length; i++) {
      html += '<p style="margin:5px 0;"><a href="' + fileUrls[i].url + '" target="_blank">' + escapeHtml(fileUrls[i].name) + '</a></p>';
    }
    html += '</div>';
  }

  html += '<hr>';
  html += '<div style="margin-top:20px; text-align:center;">';
  html += '<button onclick="google.script.host.close()" style="background:#1a73e8; color:white; border:none; padding:8px 24px; border-radius:4px; cursor:pointer;">Close</button>';
  html += '</div>';
  html += '</div>';

  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(650)
    .setHeight(550)
    .setTitle('Execution Result: ' + template.name);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Execution Result: ' + template.name);
}

// =======================================================
// SHOW NOTIFICATION FUNCTIONS
// =======================================================

function SHOW_NOTIFICATION(title, message, type) {
  var icon = type === "success" ? "✅" : (type === "warning" ? "⚠️" : "ℹ️");
  var bgColor = type === "success" ? "#e6f4ea" : (type === "warning" ? "#fce8e6" : "#e8f0fe");
  var borderColor = type === "success" ? "#137333" : (type === "warning" ? "#c5221f" : "#1a73e8");

  var html = '<div style="font-family: Roboto, sans-serif; padding: 20px;">' +
    '<div style="display: flex; align-items: center; margin-bottom: 16px;">' +
    '<span style="font-size: 28px; margin-right: 12px;">' + icon + '</span>' +
    '<h2 style="color: #1a73e8; margin: 0;">' + title + '</h2>' +
    '</div>' +
    '<div style="background: ' + bgColor + '; padding: 16px; border-radius: 6px; margin: 10px 0; border-left: 4px solid ' + borderColor + ';">' +
    '<p style="margin: 0; color: #3c4043; line-height: 1.5;">' + message + '</p>' +
    '</div>' +
    '<div style="margin-top: 24px; text-align: center;">' +
    '<button onclick="google.script.host.close()" style="background: #1a73e8; color: white; border: none; padding: 8px 24px; border-radius: 4px; font-size: 14px; font-family: Roboto, sans-serif; cursor: pointer;">OK</button>' +
    '</div>' +
    '</div>';

  return HtmlService.createHtmlOutput(html)
    .setWidth(450)
    .setHeight(220)
    .setTitle(title);
}

function SHOW_COLUMN_CREATED_NOTIFICATION(templateName) {
  return SHOW_NOTIFICATION(
    "Status Column Created",
    "✅ Column to track Status \"" + templateName + "\" Template has been created in your Google Sheet. It is protected with system styling.",
    "success"
  );
}

function SHOW_COLUMN_EXISTS_NOTIFICATION(templateName) {
  return SHOW_NOTIFICATION(
    "Status Column Already Exists",
    "ℹ️ Column to track Status \"" + templateName + "\" Template already exists in your Google Sheet. No changes were made.",
    "info"
  );
}

function SHOW_TEMPLATE_SAVED_NOTIFICATION(templateName, columnCreated, columnExists) {
  if (columnCreated) {
    return SHOW_COLUMN_CREATED_NOTIFICATION(templateName);
  } else if (columnExists) {
    return SHOW_COLUMN_EXISTS_NOTIFICATION(templateName);
  } else {
    return SHOW_NOTIFICATION(
      "Template Saved",
      "✅ Template \"" + templateName + "\" has been saved successfully.",
      "success"
    );
  }
}

// =======================================================
// AUTO MATCH TAGS FUNCTION 
// =======================================================

function AUTO_MATCH_TAGS(docUrl) {
  try {
    var tags = EXTRACT_TEMPLATE_TAGS_STREAM(docUrl);
    var headers = GET_LIVE_SHEET_HEADERS();
    var matched = [];
    var unmatched = [];

    // Build a lookup map for headers
    var headerMap = {};
    for (var i = 0; i < headers.length; i++) {
      headerMap[headers[i].toLowerCase().trim()] = headers[i];
    }

    for (var i = 0; i < tags.length; i++) {
      var tag = tags[i];
      var matchedHeader = null;
      for (var h = 0; h < headers.length; h++) {
        if (headers[h].toLowerCase().trim() === tag.toLowerCase().trim()) {
          matchedHeader = headers[h];
          break;
        }
      }
      if (matchedHeader) {
        matched.push({ tag: tag, header: matchedHeader });
      } else {
        unmatched.push(tag);
      }
    }
    return { success: true, matched: matched, unmatched: unmatched };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==========================================
// FUNCTION: CREATE_DOCUMENT_TEMPLATE
// Creates a clean, blank Google Doc template
// ==========================================
function CREATE_DOCUMENT_TEMPLATE() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var sheetName = sheet.getName();

    // Get the spreadsheet's parent folder
    var ssFile = DriveApp.getFileById(ss.getId());
    var parentFolders = ssFile.getParents();
    var currentFolder = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();

    // Move one level above (Get the grandparent folder)
    var grandParentFolders = currentFolder.getParents();
    var targetParentFolder = grandParentFolders.hasNext() ? grandParentFolders.next() : DriveApp.getRootFolder();

    // Create or open the DocuMail PRO Templates folder one level above
    var folderName = "DocuMail PRO Templates";
    var templateFolder = null;
    var existingFolders = targetParentFolder.getFoldersByName(folderName);

    if (existingFolders.hasNext()) {
      templateFolder = existingFolders.next();
    } else {
      templateFolder = targetParentFolder.createFolder(folderName);
    }

    // Create a beautifully clean, completely blank document
    var docName = "DocTemplate for " + sheetName;
    var doc = DocumentApp.create(docName);
    var docId = doc.getId();
    var docFile = DriveApp.getFileById(docId);

    // Move the document to the dedicated templates folder
    templateFolder.addFile(docFile);
    DriveApp.getRootFolder().removeFile(docFile);

    // Save and clear any default contents to guarantee a totally blank canvas
    var body = doc.getBody();
    body.clear(); 

    doc.saveAndClose();
    var docUrl = doc.getUrl();

    // =======================================================
    // CRITICAL LINK: SAVE CONTEXT FOR THE DOC SIDEBAR
    // =======================================================
    // Save the Sheet ID globally under the user's account properties.
    // When the Doc sidebar wakes up, it reads this property to fetch your columns.
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('CurrentActiveSheetId', ss.getId());

    console.log("✅ Blank Template created: " + docName);
    return {
      success: true,
      url: docUrl,
      id: docId,
      name: docName,
      folder: templateFolder.getName()
    };

  } catch (e) {
    console.error("Error creating document template: " + e.message);
    return { success: false, error: e.message };
  }
}



//file content end