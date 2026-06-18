/** file name: code.gs
/**
 * ============================================================================
 * DOCUMAIL PRO COMPLETE MASTER CORE SCRIPT
 * ============================================================================
 * FILE: code.gs - Main UI, Menu, Sidebar, and Dialog Functions
 * ============================================================================
 */

// ==========================================
// Function (onOpen) Starts
// ==========================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('DocuMail Pro Platform')
    .addItem('Initialize Sheet Structural Layout', 'GENERATE_DOCUMAIL_TEMPLATE')
    .addItem('Open Engine Workspace Control Sidebar', 'INITIALIZE_ADDON_SIDEBAR')
    .addToUi();
}

// ==========================================
// Function (INITIALIZE_ADDON_SIDEBAR) Starts
// ==========================================
function INITIALIZE_ADDON_SIDEBAR() {
  var html = HtmlService.createTemplateFromFile('SidebarView');
  var sidebarUi = html.evaluate().setTitle("DocuMail Pro").setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(sidebarUi);
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

  // Ensure tag mappings exist for PDF
  if (template.type === "PDF_ONLY" || template.type === "BOTH") {
    if (!template.config.tagMappings || Object.keys(template.config.tagMappings).length === 0) {
      var allHeaders = GET_ALL_RAW_HEADERS();
      var tags = EXTRACT_TEMPLATE_TAGS_STREAM(template.config.templateUrl);
      template.config.tagMappings = {};
      for (var i = 0; i < tags.length; i++) {
        if (allHeaders.indexOf(tags[i]) !== -1) {
          template.config.tagMappings[tags[i]] = tags[i];
        }
      }
      SAVE_TEMPLATE(template);
    }
  }

  // =======================================================
  // GET QUOTA & RUN TEMPLATE (ONCE)
  // =======================================================
  var remainingQuota = MailApp.getRemainingDailyQuota();
  var totalQuota = 100;
  var usedQuota = totalQuota - remainingQuota;

  var result = RUN_TEMPLATE(templateId);
 
  // =======================================================
  // GET FILE URLS (FOR PDF)
  // =======================================================
  var fileUrls = [];
  if (template.type === "PDF_ONLY" || template.type === "BOTH") {
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
  html += '<h2 style="color:#1a73e8;">✅ Execution Completed!</h2>';
  html += '<hr>';
  html += '<p><strong>📄 Template:</strong> ' + escapeHtml(template.name) + '</p>';
  html += '<p><strong>📋 Type:</strong> ' + (template.type === "PDF_ONLY" ? "PDF Only" : template.type === "EMAIL_ONLY" ? "Email Only" : "PDF & Email") + '</p>';

  if (template.type === "EMAIL_ONLY" || template.type === "BOTH") {
    html += '<div style="background:#fce8e6; padding:10px; border-radius:6px; margin:10px 0;">';
    html += '<p style="margin:0;"><strong>📊 Daily Email Quota:</strong></p>';
    html += '<p style="margin:5px 0;"><strong>✅ Remaining:</strong> ' + remainingQuota + ' emails</p>';
    html += '<p style="margin:5px 0;"><strong>📤 Used Today:</strong> ' + usedQuota + ' emails</p>';
    html += '<p style="margin:5px 0; color:#5f6368; font-size:12px;">📅 Resets every 24 hours</p>';
    html += '</div>';
  }

  html += '<div style="background:#e6f4ea; padding:12px; border-radius:6px; margin:10px 0;">';
  html += '<strong>📝 Result:</strong><br>' + String(result).replace(/\n/g, '<br>');
  html += '</div>';

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
//file content end