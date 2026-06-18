/**file name: utils.gs
/**
 * ============================================================================
 * DOCUMAIL PRO COMPLETE MASTER CORE SCRIPT
 * ============================================================================
 * FILE: utils.gs - Utility and Helper Functions
 * ============================================================================
 */

function escapeHtml(str) {
  if (!str) return '';
  return String(str).replace(/[&<>]/g, function (m) {
    if (m === '&') return '&amp;';
    if (m === '<') return '&lt;';
    if (m === '>') return '&gt;';
    return m;
  });
}

function escapeRegex(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function GET_SYSTEM_OAUTH_TOKEN() {
  try {
    DriveApp.getRootFolder();
    return ScriptApp.getOAuthToken();
  } catch (e) {
    return "";
  }
}

function UPLOAD_LOCAL_FILE_STREAM(fileObject) {
  try {
    var folderName = "DocuMail Pro Local Uploads";
    var folders = DriveApp.getFoldersByName(folderName);
    var targetFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    var contentSplit = fileObject.data.split(',');
    var base64Data = contentSplit.length > 1 ? contentSplit[1] : contentSplit[0];
    var fileBlob = Utilities.newBlob(Utilities.base64Decode(base64Data), fileObject.mimeType, fileObject.name);
    var driveFile = targetFolder.createFile(fileBlob);
    return { success: true, id: driveFile.getId(), name: driveFile.getName() };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function SAVE_TEMPLATE_STATE(configPayload) {
  CacheService.getDocumentCache().put('email_config_state', JSON.stringify(configPayload), 21600);
}

function GET_STORED_TEMPLATE_STATE() {
  var savedState = CacheService.getDocumentCache().get('email_config_state');
  return savedState ? JSON.parse(savedState) : null;
}

function SAVE_DOC_WIZARD_STATE(configPayload) {
  CacheService.getDocumentCache().put('doc_config_state_memory', JSON.stringify(configPayload), 21600);
}

function GET_SAVED_DOC_WIZARD_STATE() {
  var cachePayload = CacheService.getDocumentCache().get('doc_config_state_memory');
  return cachePayload ? JSON.parse(cachePayload) : null;
}

function RECEPTACLE_SIDEBAR_STATE_PASS(sidebarOptionsPayload) {
  CacheService.getDocumentCache().put('sidebar_channel_options', JSON.stringify(sidebarOptionsPayload), 21600);
}

// ==========================================
// NEW: FORMAT_DATE_FOR_DISPLAY
// ==========================================
function FORMAT_DATE_FOR_DISPLAY(dateValue) {
  if (!dateValue) return "";
  
  var date;
  if (dateValue instanceof Date) {
    date = dateValue;
  } else {
    try {
      // Handle string dates like "6/1/26" (MM/DD/YY)
      var str = String(dateValue).trim();
      // Check if it's a date string with slashes
      if (str.indexOf('/') !== -1) {
        var parts = str.split('/');
        if (parts.length === 3) {
          // MM/DD/YY or MM/DD/YYYY
          var month = parseInt(parts[0]) - 1;
          var day = parseInt(parts[1]);
          var year = parseInt(parts[2]);
          if (year < 100) year += 2000;
          date = new Date(year, month, day);
        } else {
          date = new Date(str);
        }
      } else {
        date = new Date(str);
      }
    } catch (e) {
      return String(dateValue);
    }
  }
  
  if (isNaN(date.getTime())) return String(dateValue);
  
  // Format: "Jan 21, 2026"
  var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var month = months[date.getMonth()];
  var day = date.getDate();
  var year = date.getFullYear();
  
  return month + ' ' + day + ', ' + year;
}

// ==========================================
// NEW: FORMAT_NUMBER_FOR_DISPLAY
// ==========================================
function FORMAT_NUMBER_FOR_DISPLAY(value, currencySymbol) {
  if (typeof value !== 'number') return String(value || "");
  currencySymbol = currencySymbol || "₹";
  return currencySymbol + value.toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}
//file content end