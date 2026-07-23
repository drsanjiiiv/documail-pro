/** file name: code.gs
/**
 * ============================================================================
 * DOCUMAIL PRO COMPLETE MASTER CORE SCRIPT
 * ============================================================================
 * FILE: code.gs - Main UI, Menu, Sidebar, and Dialog Functions
 * ============================================================================
 */

// ==========================================
// MASTER TEMPLATE CONFIGURATION
// ==========================================
// Before publishing, set this to your "DocuMailPro Master Doc Template" file ID.
// Share the doc with "Anyone with the link can view" so add-on users can access it.
var MASTER_TEMPLATE_FILE_ID = '1RzrM1_w8M_MmtzKpKjeCoglnjHScEOFnZJHRHrBTuZE';

function GET_MASTER_TEMPLATE_ID() {
  var props = PropertiesService.getDocumentProperties();

  // 1. Check PropertiesService first (user-set via menu or auto-detected)
  var stored = props.getProperty('DOCUMAIL_MASTER_TEMPLATE_ID');
  if (stored) return stored;

  // 2. Check hardcoded constant (developer sets before publishing)
  if (MASTER_TEMPLATE_FILE_ID) return MASTER_TEMPLATE_FILE_ID;

  // 3. Fallback: search by name (backward compatible)
  var files = DriveApp.getFilesByName("DocuMailPro Master Doc Template");
  if (files.hasNext()) {
    var file = files.next();
    props.setProperty('DOCUMAIL_MASTER_TEMPLATE_ID', file.getId());
    return file.getId();
  }

  return null;
}

// ==========================================
// SPREADSHEET PLATFORM - ONOPEN (CLEAN)
// ==========================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🚀 DocuMail Pro Sheet')
    .addItem('📄 Initialize DocuMail Pro Sheet Structure', 'GENERATE_DOCUMAIL_TEMPLATE')
    .addItem('📝 Create Dynamic Doc Template From Sheet', 'CREATE_DOCUMENT_TEMPLATE_MENU')
    .addItem('🧠 Open Smart Template Engine', 'INITIALIZE_ADDON_SIDEBAR')
    .addSeparator()
    .addItem('⚙️ Set Master Template', 'SET_MASTER_TEMPLATE')
    .addItem('❓ Help', 'showSheetHelp')
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
// HELP FUNCTION FOR SHEETS MENU
// ==========================================
function showSheetHelp() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    '📚 DocuMail Pro - Help',
    '📄 Initialize DocuMail Pro Sheet Structure\n' +
    'Adds all required system columns to your sheet for tracking:\n' +
    '• Merged Doc Status\n' +
    '• Merged Doc ID\n' +
    '• Merged Doc URL\n' +
    '• Recipient Email (if not exists)\n' +
    'These columns are essential for document generation & tracking.\n\n' +
    '─────────────────────────────\n\n' +
    '📝 Create Dynamic Doc Template From Sheet\n' +
    'Creates a new Google Doc template linked to your current sheet:\n' +
    '• Copies the master template with all scripts\n' +
    '• Automatically links it to this sheet\n' +
    '• Opens the template for you to design\n' +
    '• Variables from your sheet headers can be inserted\n\n' +
    '─────────────────────────────\n\n' +
    '🧠 Open Smart Template Engine\n' +
    'Opens the main sidebar where you can:\n' +
    '• View all your templates\n' +
    '• Create new templates (PDF / Email / Both)\n' +
    '• Edit existing templates\n' +
    '• Preview templates with real data\n' +
    '• Run templates to generate documents & send emails\n\n' +
    '─────────────────────────────\n\n' +
    '💡 Pro Tip: Always initialize your sheet structure first\n' +
    'before creating templates for best results!',
    ui.ButtonSet.OK
  );
}

// ==========================================
// FUNCTION: CREATE_DOCUMENT_TEMPLATE_MENU
// Shows step-by-step animation with dummy progression
// ==========================================
function CREATE_DOCUMENT_TEMPLATE_MENU() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var sheetName = sheet.getName();
    var sheetId = ss.getId();

    // Show the step-by-step animation dialog IMMEDIATELY
    var htmlContent = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body {
      margin: 0;
      padding: 0;
      display: flex;
      justify-content: center;
      align-items: center;
      height: 100vh;
      background: #f8f9fa;
      font-family: 'Roboto', sans-serif;
    }
    .loader-container {
      text-align: left;
      padding: 40px 50px;
      background: white;
      border-radius: 12px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      min-width: 420px;
    }
    .step {
      padding: 10px 0;
      border-bottom: 1px solid #f1f3f4;
      display: flex;
      align-items: center;
      font-size: 14px;
      color: #3c4043;
    }
    .step:last-child {
      border-bottom: none;
    }
    .step .icon {
      width: 28px;
      text-align: center;
      margin-right: 12px;
      font-size: 16px;
    }
    .step.done .icon {
      color: #137333;
    }
    .step.active .icon {
      color: #1a73e8;
    }
    .step.pending .icon {
      color: #dadce0;
    }
    .step.done .text {
      color: #137333;
    }
    .step.active .text {
      color: #1a73e8;
      font-weight: 500;
    }
    .step.pending .text {
      color: #9aa0a6;
    }
    .title {
      font-size: 16px;
      font-weight: 600;
      color: #1a73e8;
      margin-bottom: 16px;
      text-align: center;
    }
    .error-text {
      color: #c5221f;
      font-size: 14px;
      margin-top: 12px;
      text-align: center;
    }
  </style>
</head>
<body>
  <div class="loader-container" id="loader">
    <div class="title">📄 Creating Your Template</div>
    
    <div class="step active" id="step1">
      <span class="icon">⏳</span>
      <span class="text">📊 Fetching headers from your sheet...</span>
    </div>
    
    <div class="step pending" id="step2">
      <span class="icon">⏳</span>
      <span class="text">📝 Creating dynamic template document...</span>
    </div>
    
    <div class="step pending" id="step3">
      <span class="icon">⏳</span>
      <span class="text">🔗 Linking template to your sheet...</span>
    </div>
    
    <div class="step pending" id="step4">
      <span class="icon">⏳</span>
      <span class="text">📄 Copying master template structure...</span>
    </div>
    
    <div class="step pending" id="step5">
      <span class="icon">⏳</span>
      <span class="text">⚙️ Configuring template settings...</span>
    </div>
    
    <div class="step pending" id="step6">
      <span class="icon">⏳</span>
      <span class="text">🚀 Launching template canvas...</span>
    </div>
  </div>
  
  <script>
    var steps = [
      { id: 'step1', done: false },
      { id: 'step2', done: false },
      { id: 'step3', done: false },
      { id: 'step4', done: false },
      { id: 'step5', done: false },
      { id: 'step6', done: false }
    ];
    
    function updateStep(stepIndex, status) {
      var step = document.getElementById(steps[stepIndex].id);
      var icon = step.querySelector('.icon');
      var text = step.querySelector('.text');
      
      step.className = 'step';
      if (status === 'done') {
        step.classList.add('done');
        icon.textContent = '✅';
        steps[stepIndex].done = true;
        if (stepIndex < steps.length - 1) {
          var nextStep = document.getElementById(steps[stepIndex + 1].id);
          nextStep.className = 'step active';
          nextStep.querySelector('.icon').textContent = '⏳';
        }
      } else if (status === 'active') {
        step.classList.add('active');
        icon.textContent = '⏳';
      } else {
        step.classList.add('pending');
        icon.textContent = '⏳';
      }
    }
    
    // DUMMY ANIMATION - Steps complete one by one
    setTimeout(function() { updateStep(0, 'done'); }, 300);
    setTimeout(function() { updateStep(1, 'done'); }, 600);
    setTimeout(function() { updateStep(2, 'done'); }, 900);
    setTimeout(function() { updateStep(3, 'done'); }, 1200);
    setTimeout(function() { updateStep(4, 'done'); }, 1500);
    
    // Open blank popup BEFORE async call (user gesture still active)
    var templateWindow = window.open('', '_blank');
    var popupBlocked = !templateWindow || templateWindow.closed;

    // Start the actual work in background
    google.script.run
      .withSuccessHandler(function(result) {
        if (result.success) {
          updateStep(5, 'done');

          if (!popupBlocked && templateWindow && !templateWindow.closed) {
            templateWindow.location.href = result.url;
          } else {
            document.getElementById('loader').innerHTML =
              '<div style="text-align:center;padding:20px;font-family:sans-serif;">' +
              '<h2 style="color:#137333;">✅ Template Created!</h2>' +
              '<p style="color:#5f6368;margin:16px 0;">Click to open your template:</p>' +
              '<a href="' + result.url + '" target="_blank" style="display:inline-block;padding:12px 32px;background:#1a73e8;color:#fff;text-decoration:none;border-radius:6px;font-size:15px;">📄 Open Template</a>' +
              '<p style="color:#9aa0a6;font-size:12px;margin-top:16px;">(Enable popups for auto-open next time)</p>' +
              '</div>';
            return;
          }

          setTimeout(function() {
            google.script.host.close();
          }, 600);
        } else {
          document.getElementById('loader').innerHTML = 
            '<div class="title" style="color:#c5221f;">❌ Error</div>' +
            '<div class="error-text">' + result.error + '</div>';
        }
      })
      .withFailureHandler(function(error) {
        document.getElementById('loader').innerHTML = 
          '<div class="title" style="color:#c5221f;">❌ Error</div>' +
          '<div class="error-text">' + error.message + '</div>';
      })
      .CREATE_TEMPLATE_IN_BACKGROUND("` + sheetName + `", "` + sheetId + `");
  </script>
</body>
</html>
`;

    var interfaceOutput = HtmlService.createHtmlOutput(htmlContent)
        .setWidth(520)
        .setHeight(400)
        .setTitle("Creating Template...");
        
    SpreadsheetApp.getUi().showModalDialog(interfaceOutput, "Creating Template...");

  } catch (e) {
    SpreadsheetApp.getUi().alert("Error launching template: " + e.message);
  }
}

// ==========================================
// FUNCTION: CREATE_TEMPLATE_IN_BACKGROUND
// All your existing logic - runs in background while animation shows
// ==========================================
function CREATE_TEMPLATE_IN_BACKGROUND(sheetName, sheetId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    
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
    var masterId = GET_MASTER_TEMPLATE_ID();
    if (!masterId) {
      throw new Error("Master template not found. Go to DocuMail Pro Sheet → ⚙️ Set Master Template to configure it.");
    }
    var masterFile = DriveApp.getFileById(masterId);

    // 3. Clone the Master File
    var docFile = masterFile.makeCopy("DocTemplate for " + sheetName, templateFolder);
    var docId = docFile.getId();
    
    // SECURE STORAGE: Write the Source Sheet ID directly into the file's description metadata
    docFile.setDescription(ss.getId());

    // 4. Set Onboarding Text On Canvas
    var doc = DocumentApp.openById(docId);
    var body = doc.getBody();
    body.clear();
    
    var titleParagraph = body.appendParagraph("📄 DocuMail Pro Master Template Canvas");
    titleParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    
    var descParagraph = body.appendParagraph("\n👉 Go to: DocuMail Pro > Initialize DocuMail PRO Template\n\nThis will clear the canvas and link sheet headers and a DocuMail PRO Template Engine will open on side, all the headers will be available as Variables.\n\nYou are free to insert any variable, any number of times. You can also use Variables with conditions, like when a variable shall be visible.\n\nMake sure to choose Paragraph Text for Paragraph & Table Row for Table, if using conditional insert");
    descParagraph.setHeading(DocumentApp.ParagraphHeading.NORMAL);

    doc.saveAndClose();
    var docUrl = doc.getUrl();

    return { success: true, url: docUrl };
    
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==========================================
// FUNCTION: SET_MASTER_TEMPLATE — One-time setup for the master template file
// ==========================================
function SET_MASTER_TEMPLATE() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    '⚙️ Set Master Template',
    'Enter the file ID of your "DocuMailPro Master Doc Template":\n\n' +
    'Tip: The file ID is the long string in the URL after /d/ and before /edit\n\n' +
    'Leave empty and press OK to auto-detect by name.',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;

  var input = result.getResponseText().trim();
  var props = PropertiesService.getDocumentProperties();

  if (input) {
    // Verify the ID is valid
    try {
      var file = DriveApp.getFileById(input);
      props.setProperty('DOCUMAIL_MASTER_TEMPLATE_ID', input);
      ui.alert('✅ Master template set successfully!\n\nFile: ' + file.getName());
    } catch (e) {
      ui.alert('❌ Invalid file ID. Please check and try again.\n\nError: ' + e.message);
    }
  } else {
    // Auto-detect by name
    var files = DriveApp.getFilesByName("DocuMailPro Master Doc Template");
    if (files.hasNext()) {
      var file = files.next();
      props.setProperty('DOCUMAIL_MASTER_TEMPLATE_ID', file.getId());
      ui.alert('✅ Auto-detected and set!\n\nFile: ' + file.getName());
    } else {
      ui.alert('❌ No file named "DocuMailPro Master Doc Template" found in your Drive.');
    }
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
  // SHOW STEP-BY-STEP ANIMATION DIALOG FIRST
  // =======================================================
  var template = GET_TEMPLATE_BY_ID(templateId);
  var templateName = template ? template.name : "Template";
  
  var loadingHtml = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body {
      margin: 0;
      padding: 0;
      display: flex;
      justify-content: center;
      align-items: center;
      height: 100vh;
      background: #f8f9fa;
      font-family: 'Roboto', sans-serif;
    }
    .loader-container {
      text-align: left;
      padding: 40px 50px;
      background: white;
      border-radius: 12px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      min-width: 420px;
    }
    .step {
      padding: 10px 0;
      border-bottom: 1px solid #f1f3f4;
      display: flex;
      align-items: center;
      font-size: 14px;
      color: #3c4043;
    }
    .step:last-child {
      border-bottom: none;
    }
    .step .icon {
      width: 28px;
      text-align: center;
      margin-right: 12px;
      font-size: 16px;
    }
    .step.done .icon {
      color: #137333;
    }
    .step.active .icon {
      color: #1a73e8;
    }
    .step.pending .icon {
      color: #dadce0;
    }
    .step.done .text {
      color: #137333;
    }
    .step.active .text {
      color: #1a73e8;
      font-weight: 500;
    }
    .step.pending .text {
      color: #9aa0a6;
    }
    .title {
      font-size: 16px;
      font-weight: 600;
      color: #1a73e8;
      margin-bottom: 16px;
      text-align: center;
    }
    .subtitle {
      font-size: 12px;
      color: #5f6368;
      text-align: center;
      margin-bottom: 12px;
    }
    .error-text {
      color: #c5221f;
      font-size: 14px;
      margin-top: 12px;
      text-align: center;
    }
  </style>
</head>
<body>
  <div class="loader-container" id="loader">
    <div class="title">👁️ Generating Preview</div>
    <div class="subtitle">` + templateName + `</div>
    
    <div class="step active" id="step1">
      <span class="icon">⏳</span>
      <span class="text">📊 Fetching headers from your sheet...</span>
    </div>
    
    <div class="step pending" id="step2">
      <span class="icon">⏳</span>
      <span class="text">📋 Fetching row data for preview...</span>
    </div>
    
    <div class="step pending" id="step3">
      <span class="icon">⏳</span>
      <span class="text">🔍 Mapping template tags to sheet columns...</span>
    </div>
    
    <div class="step pending" id="step4">
      <span class="icon">⏳</span>
      <span class="text">📄 Processing conditional blocks...</span>
    </div>
    
    <div class="step pending" id="step5">
      <span class="icon">⏳</span>
      <span class="text">⚙️ Generating preview document...</span>
    </div>
    
    <div class="step pending" id="step6">
      <span class="icon">⏳</span>
      <span class="text">🚀 Launching preview window...</span>
    </div>
  </div>
  
  <script>
    var steps = [
      { id: 'step1', done: false },
      { id: 'step2', done: false },
      { id: 'step3', done: false },
      { id: 'step4', done: false },
      { id: 'step5', done: false },
      { id: 'step6', done: false }
    ];
    
    function updateStep(stepIndex, status) {
      var step = document.getElementById(steps[stepIndex].id);
      var icon = step.querySelector('.icon');
      var text = step.querySelector('.text');
      
      step.className = 'step';
      if (status === 'done') {
        step.classList.add('done');
        icon.textContent = '✅';
        steps[stepIndex].done = true;
        if (stepIndex < steps.length - 1) {
          var nextStep = document.getElementById(steps[stepIndex + 1].id);
          nextStep.className = 'step active';
          nextStep.querySelector('.icon').textContent = '⏳';
        }
      } else if (status === 'active') {
        step.classList.add('active');
        icon.textContent = '⏳';
      } else {
        step.classList.add('pending');
        icon.textContent = '⏳';
      }
    }
    
    // DUMMY ANIMATION
    setTimeout(function() { updateStep(0, 'done'); }, 300);
    setTimeout(function() { updateStep(1, 'done'); }, 600);
    setTimeout(function() { updateStep(2, 'done'); }, 900);
    setTimeout(function() { updateStep(3, 'done'); }, 1200);
    setTimeout(function() { updateStep(4, 'done'); }, 1500);
    
    // Start the actual work in background
    google.script.run
      .withSuccessHandler(function(result) {
        updateStep(5, 'done');
        // Show the preview result
        var previewHtml = result.previewHtml;
        document.body.innerHTML = previewHtml;
      })
      .withFailureHandler(function(error) {
        document.getElementById('loader').innerHTML = 
          '<div class="title" style="color:#c5221f;">❌ Error</div>' +
          '<div class="error-text">' + error.message + '</div>';
      })
      .GENERATE_PREVIEW_IN_BACKGROUND("` + templateId + `");
  </script>
</body>
</html>
`;

  var loadingDialog = HtmlService.createHtmlOutput(loadingHtml)
    .setWidth(550)
    .setHeight(450)
    .setTitle('Generating Preview...');
  SpreadsheetApp.getUi().showModalDialog(loadingDialog, 'Generating Preview...');
}

// =======================================================
// SHOW RUN DIALOG - WITH STEP-BY-STEP ANIMATION
// =======================================================

function SHOW_RUN_DIALOG(templateId) {
  // =======================================================
  // SHOW STEP-BY-STEP ANIMATION DIALOG FIRST
  // =======================================================
  var template = GET_TEMPLATE_BY_ID(templateId);
  var templateName = template ? template.name : "Template";
  var templateType = template ? template.type : "UNKNOWN";
  var typeLabel = templateType === "PDF_ONLY" ? "📄 PDF" : templateType === "EMAIL_ONLY" ? "✉️ Email" : "📄✉️ Both";
  
  var loadingHtml = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body {
      margin: 0;
      padding: 0;
      display: flex;
      justify-content: center;
      align-items: center;
      height: 100vh;
      background: #f8f9fa;
      font-family: 'Roboto', sans-serif;
    }
    .loader-container {
      text-align: left;
      padding: 40px 50px;
      background: white;
      border-radius: 12px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      min-width: 420px;
    }
    .step {
      padding: 10px 0;
      border-bottom: 1px solid #f1f3f4;
      display: flex;
      align-items: center;
      font-size: 14px;
      color: #3c4043;
    }
    .step:last-child {
      border-bottom: none;
    }
    .step .icon {
      width: 28px;
      text-align: center;
      margin-right: 12px;
      font-size: 16px;
    }
    .step.done .icon {
      color: #137333;
    }
    .step.active .icon {
      color: #1a73e8;
    }
    .step.pending .icon {
      color: #dadce0;
    }
    .step.done .text {
      color: #137333;
    }
    .step.active .text {
      color: #1a73e8;
      font-weight: 500;
    }
    .step.pending .text {
      color: #9aa0a6;
    }
    .title {
      font-size: 16px;
      font-weight: 600;
      color: #1a73e8;
      margin-bottom: 4px;
      text-align: center;
    }
    .subtitle {
      font-size: 12px;
      color: #5f6368;
      text-align: center;
      margin-bottom: 12px;
    }
    .badge {
      display: inline-block;
      padding: 2px 12px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: 500;
      background: #e8f0fe;
      color: #1a73e8;
    }
    .error-text {
      color: #c5221f;
      font-size: 14px;
      margin-top: 12px;
      text-align: center;
    }
  </style>
</head>
<body>
  <div class="loader-container" id="loader">
    <div class="title">▶️ Running Template</div>
    <div class="subtitle">
      <span class="badge">` + typeLabel + `</span>
      &nbsp; ` + templateName + `
    </div>
    
    <div class="step active" id="step1">
      <span class="icon">⏳</span>
      <span class="text">📊 Fetching headers from your sheet...</span>
    </div>
    
    <div class="step pending" id="step2">
      <span class="icon">⏳</span>
      <span class="text">📋 Scanning for eligible rows...</span>
    </div>
    
    <div class="step pending" id="step3">
      <span class="icon">⏳</span>
      <span class="text">🔍 Mapping template tags to sheet columns...</span>
    </div>
    
    <div class="step pending" id="step4">
      <span class="icon">⏳</span>
      <span class="text">📄 Processing conditional blocks...</span>
    </div>
    
    <div class="step pending" id="step5">
      <span class="icon">⏳</span>
      <span class="text">⚙️ Generating documents & sending emails...</span>
    </div>
    
    <div class="step pending" id="step6">
      <span class="icon">⏳</span>
      <span class="text">✅ Finalizing and updating status...</span>
    </div>
  </div>
  
  <script>
    var steps = [
      { id: 'step1', done: false },
      { id: 'step2', done: false },
      { id: 'step3', done: false },
      { id: 'step4', done: false },
      { id: 'step5', done: false },
      { id: 'step6', done: false }
    ];
    
    function updateStep(stepIndex, status) {
      var step = document.getElementById(steps[stepIndex].id);
      var icon = step.querySelector('.icon');
      var text = step.querySelector('.text');
      
      step.className = 'step';
      if (status === 'done') {
        step.classList.add('done');
        icon.textContent = '✅';
        steps[stepIndex].done = true;
        if (stepIndex < steps.length - 1) {
          var nextStep = document.getElementById(steps[stepIndex + 1].id);
          nextStep.className = 'step active';
          nextStep.querySelector('.icon').textContent = '⏳';
        }
      } else if (status === 'active') {
        step.classList.add('active');
        icon.textContent = '⏳';
      } else {
        step.classList.add('pending');
        icon.textContent = '⏳';
      }
    }
    
    // DUMMY ANIMATION - Steps complete one by one
    setTimeout(function() { updateStep(0, 'done'); }, 300);
    setTimeout(function() { updateStep(1, 'done'); }, 600);
    setTimeout(function() { updateStep(2, 'done'); }, 900);
    setTimeout(function() { updateStep(3, 'done'); }, 1200);
    setTimeout(function() { updateStep(4, 'done'); }, 1500);
    
    // Start the actual work in background
    google.script.run
      .withSuccessHandler(function(result) {
        updateStep(5, 'done');
        // Show the result
        document.body.innerHTML = result;
      })
      .withFailureHandler(function(error) {
        document.getElementById('loader').innerHTML = 
          '<div class="title" style="color:#c5221f;">❌ Error</div>' +
          '<div class="error-text">' + error.message + '</div>';
      })
      .RUN_TEMPLATE_IN_BACKGROUND("` + templateId + `");
  </script>
</body>
</html>
`;

  var loadingDialog = HtmlService.createHtmlOutput(loadingHtml)
    .setWidth(550)
    .setHeight(450)
    .setTitle('Running Template...');
  SpreadsheetApp.getUi().showModalDialog(loadingDialog, 'Running Template...');
}

// ==========================================
// FUNCTION: GENERATE_PREVIEW_IN_BACKGROUND
// Runs preview in background
// ==========================================
function GENERATE_PREVIEW_IN_BACKGROUND(templateId) {
  try {
    // First, process the preview (generate the document)
    var template = GET_TEMPLATE_BY_ID(templateId);
    if (template && (template.type === "PDF_ONLY" || template.type === "BOTH")) {
      var config = JSON.parse(JSON.stringify(template.config));
      config.isPreview = true;

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
    
    // Then get the preview HTML
    var result = PREVIEW_TEMPLATE(templateId);
    return result;
    
  } catch (e) {
    return { name: "Error", previewHtml: "<p>ERROR: " + e.toString() + "</p>" };
  }
}

// ==========================================
// FUNCTION: RUN_TEMPLATE_IN_BACKGROUND
// Runs template in background with file links - MATCHES PREVIEW STYLE
// ==========================================
function RUN_TEMPLATE_IN_BACKGROUND(templateId) {
  try {
    var template = GET_TEMPLATE_BY_ID(templateId);
    
    if (!template) {
      return '<div style="font-family: Roboto, sans-serif; padding: 24px;">' +
             '<h2 style="color: #c5221f; margin: 0 0 12px 0;">❌ Template Not Found</h2>' +
             '<hr style="border: none; border-top: 1px solid #e8eaed; margin: 12px 0;">' +
             '<p style="color: #5f6368; margin: 8px 0;">Template ID: <strong>' + templateId + '</strong></p>' +
             '<p style="color: #5f6368; margin: 8px 0;">Please delete this template and recreate it.</p>' +
             '<hr style="border: none; border-top: 1px solid #e8eaed; margin: 16px 0;">' +
             '<div style="margin-top: 20px; text-align: center;">' +
             '<button onclick="google.script.host.close()" style="background: #1a73e8; color: white; border: none; padding: 10px 28px; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: 500;">Close</button>' +
             '</div></div>';
    }
    
    // Process mapping configurations
    if (template.type === "PDF_ONLY" || template.type === "BOTH") {
      var config = JSON.parse(JSON.stringify(template.config));
      config.isPreview = false;
      
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
        template.config.tagMappings = config.tagMappings;
        if (typeof SAVE_TEMPLATE === 'function') {
          SAVE_TEMPLATE(template);
        }
      }
    }
    
    var result = RUN_TEMPLATE(templateId);
    
    // Get file URLs from the folder
var fileUrls = [];
if ((template.type === "PDF_ONLY" || template.type === "BOTH") && result !== "NO_ROWS_ELIGIBLE" && result && result.indexOf("Error") === -1) {
  var folderId = template.config?.folderDestination;
  if (folderId) {
    var folderIdExtracted = folderId.split("/folders/")[1] || folderId.split("id=")[1];
    if (folderIdExtracted) {
      try {
        var folder = DriveApp.getFolderById(folderIdExtracted);
        var files = folder.getFiles();
        var latestFiles = [];
        var currentTime = new Date().getTime();
        
        while (files.hasNext()) {
          var file = files.next();
          var fileCreated = file.getDateCreated().getTime();
          var fileName = file.getName();
          
          // EXCLUDE preview files (PROV- prefix) and only get files created in last 2 minutes
          if (fileCreated > currentTime - 120000 && fileName.indexOf("PROV-") === -1) {
            latestFiles.push(file);
          }
        }
        latestFiles.sort(function (a, b) {
          return b.getDateCreated().getTime() - a.getDateCreated().getTime();
        });
        for (var i = 0; i < latestFiles.length; i++) {
          fileUrls.push({ name: latestFiles[i].getName(), url: latestFiles[i].getUrl() });
        }
      } catch (e) {
        // Folder access error - ignore
      }
    }
  }
}    
    var html = '<div style="font-family: Roboto, sans-serif; padding: 24px;">';
    
    if (result && result.indexOf("Error") !== -1) {
      html += '<h2 style="color: #c5221f; margin: 0 0 12px 0;">❌ Execution Failed</h2>';
      html += '<hr style="border: none; border-top: 1px solid #e8eaed; margin: 12px 0;">';
      html += '<div style="background: #fce8e6; padding: 16px; border-radius: 6px; margin: 12px 0; border-left: 4px solid #c5221f; line-height: 1.6;">';
      html += '<span style="color: #c5221f;">' + result.replace(/\n/g, '<br>') + '</span>';
      html += '</div>';
      
    } else if (result === "NO_ROWS_ELIGIBLE") {
      html += '<h2 style="color: #f2994a; margin: 0 0 12px 0;">ℹ️ No Rows Eligible</h2>';
      html += '<hr style="border: none; border-top: 1px solid #e8eaed; margin: 12px 0;">';
      html += '<div style="background: #fff3cd; padding: 16px; border-radius: 6px; margin: 12px 0; border-left: 4px solid #f2994a; color: #856404; line-height: 1.6;">';
      html += 'All rows matching your filter conditions have already been merged successfully with a status of <strong>Success</strong>.<br><br>';
      html += 'There are no fresh pending records available to process.';
      html += '</div>';
      
    } else {
      html += '<h2 style="color: #1a73e8; margin: 0 0 12px 0;">✅ Execution Completed!</h2>';
      html += '<hr style="border: none; border-top: 1px solid #e8eaed; margin: 12px 0;">';
      
      // Document Preview section - MATCHES PREVIEW STYLE
      html += '<div style="background: #e8f0fe; padding: 16px; border-radius: 6px; margin: 12px 0;">';
      html += '<h3 style="color: #1a73e8; margin: 0 0 10px 0;">📄 Document Preview</h3>';
      html += '<p style="margin: 0 0 8px 0;"><strong>Template:</strong> ' + escapeHtml(template.name) + '</p>';
      html += '<p style="margin: 0 0 8px 0;"><strong>Type:</strong> ' + (template.type === "PDF_ONLY" ? "PDF Only" : template.type === "EMAIL_ONLY" ? "Email Only" : "PDF & Email") + '</p>';
      html += '<p style="margin: 0;"><strong>Destination:</strong> ' + escapeHtml(template.config?.folderDestination || "Default Folder") + '</p>';
      html += '</div>';
      
      // Result section
      html += '<div style="background: #e6f4ea; padding: 16px; border-radius: 6px; margin: 12px 0; border-left: 4px solid #137333;">';
      html += '<p style="margin: 0; color: #137333;"><strong>✅ Documents Generated Successfully!</strong></p>';
      html += '<p style="margin: 5px 0;"><strong>Result:</strong> ' + result.replace(/\n/g, '<br>') + '</p>';

      // Add daily quota info
      var remainingQuota = MailApp.getRemainingDailyQuota();
      var quotaColor = remainingQuota < 20 ? '#e37400' : '#137333';
      html += '<p style="margin: 8px 0 0 0; color: ' + quotaColor + ';"><strong>📊 Daily Email Quota Remaining:</strong> ' + remainingQuota + ' emails</p>';

      html += '</div>';
      
      // File links section - MATCHES PREVIEW STYLE
      if (fileUrls.length > 0) {
        html += '<div style="background: #e6f4ea; padding: 16px; border-radius: 6px; margin: 12px 0; border-left: 4px solid #137333;">';
        html += '<p style="margin: 0; color: #137333;"><strong>📂 Generated Files:</strong></p>';
        for (var i = 0; i < fileUrls.length; i++) {
          html += '<p style="margin: 5px 0;"><a href="' + fileUrls[i].url + '" target="_blank" style="color: #1a73e8; text-decoration: none;">📄 ' + escapeHtml(fileUrls[i].name) + '</a></p>';
        }
        html += '</div>';
      }
    }
    
    html += '<hr style="border: none; border-top: 1px solid #e8eaed; margin: 16px 0;">';
    html += '<div style="margin-top: 20px; text-align: center;">';
    html += '<button onclick="google.script.host.close()" style="background: #1a73e8; color: white; border: none; padding: 10px 28px; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: 500;">Close</button>';
    html += '</div>';
    html += '</div>';
    
    return html;
    
  } catch (e) {
    return '<div style="font-family: Roboto, sans-serif; padding: 24px;">' +
           '<h2 style="color: #c5221f; margin: 0 0 12px 0;">❌ Error</h2>' +
           '<hr style="border: none; border-top: 1px solid #e8eaed; margin: 12px 0;">' +
           '<p style="color: #5f6368; margin: 8px 0;">' + e.message + '</p>' +
           '<hr style="border: none; border-top: 1px solid #e8eaed; margin: 16px 0;">' +
           '<div style="margin-top: 20px; text-align: center;">' +
           '<button onclick="google.script.host.close()" style="background: #1a73e8; color: white; border: none; padding: 10px 28px; border-radius: 4px; cursor: pointer; font-size: 14px; font-weight: 500;">Close</button>' +
           '</div></div>';
  }
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