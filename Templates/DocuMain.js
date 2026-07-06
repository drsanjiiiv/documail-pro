/**
 * DocuMain.gs (Inside your Document Template Project)
 * Main Controller for the Document Template Designer.
 */

function onOpen() {
  createMenu();
}

/**
 * Creates the custom menu in Google Docs
 */
function createMenu() {
  var ui = DocumentApp.getUi();
  ui.createMenu('🚀 DocuMail Pro Template')
    .addItem('📄 Initialize DocuMail Pro Template', 'INITIALIZE_DOC_DESIGNER_SIDEBAR')
    .addItem('📝 Open Smart Variable Window', 'OPEN_SMART_VARIABLE_WINDOW')
    .addSeparator()
    .addItem('❓ Help', 'showHelp')
    .addToUi();
}

/**
 * Shows help information to the user
 */
function showHelp() {
  var ui = DocumentApp.getUi();
  ui.alert(
    '📚 DocuMail Pro - Help',
    'How to use Smart Variables:\n\n' +
    '1. Click "Start Dynamic Template" to clear the canvas and start fresh\n' +
    '2. Click "Open Smart Variable Window" to open the sidebar\n' +
    '3. Select a variable from the dropdown\n' +
    '4. Choose to insert it "Always" or "Only when..." with conditions\n' +
    '5. Click "Insert" to add it to your template\n\n' +
    'Variables will be replaced with actual data when generating documents.',
    ui.ButtonSet.OK
  );
}

/**
 * Open Smart Variable Window - Opens the sidebar WITHOUT resetting the canvas
 */
function OPEN_SMART_VARIABLE_WINDOW() {
  try {
    var ui = DocumentApp.getUi();

    // Evaluate and render the HTML sidebar panel
    var htmlOutput = HtmlService.createTemplateFromFile('SidebarDocView')
      .evaluate()
      .setTitle('DocuMail Pro - Smart Variables')
      .setWidth(300);

    ui.showSidebar(htmlOutput);

  } catch (error) {
    Logger.log("Error opening sidebar: " + error.toString());
    DocumentApp.getUi().alert("Could not open sidebar: " + error.message);
  }
}

/**
 * Checks for existing custom layout text, prompts user, and opens the sidebar interface.
 */
function INITIALIZE_DOC_DESIGNER_SIDEBAR() {
  try {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var ui = DocumentApp.getUi();

    // 1. Get all text currently sitting on the canvas
    var currentText = body.getText().trim();

    // 2. Define our standard onboarding placeholder phrases
    var line1 = "📄 DocuMail Pro Master Template Canvas";
    var line2 = "Please use the menu item to start designing your automation layout:";

    // Check if the document contains content that isn't our onboarding text
    var hasCustomContent = false;
    if (currentText.length > 0) {
      if (currentText.indexOf(line1) === -1 || currentText.indexOf(line2) === -1) {
        hasCustomContent = true;
      }
    }

    // 3. Trigger Confirmation Box if custom user data is detected
    if (hasCustomContent) {
      var response = ui.alert(
        '⚠️ Confirm Canvas Reset',
        'There is data in your document, it will be erased. Are you sure?',
        ui.ButtonSet.YES_NO
      );

      // If user clicks "NO", stop execution instantly
      if (response !== ui.Button.YES) {
        doc.toast("Operation cancelled. Your existing template was saved.", "🚀 DocuMail Pro");
        return;
      }
    }

    // 4. Wipe the canvas completely clean
    body.clear();
    body.appendParagraph(""); // Google Docs require at least one structural paragraph element

    // 5. Evaluate and render the HTML sidebar panel
    var htmlOutput = HtmlService.createTemplateFromFile('SidebarDocView')
      .evaluate()
      .setTitle('DocuMail Pro - Template Designer')
      .setWidth(300);

    ui.showSidebar(htmlOutput);

  } catch (error) {
    Logger.log("Error initializing document canvas workspace: " + error.toString());
    DocumentApp.getUi().alert("Could not load designer panel: " + error.message);
  }
}

/**
 * Universal token injector engine.
 * Added 'content' parameter to allow custom template content insertion.
 */
function injectTagAtCursor(openingTag, closingTag, content) {
  try {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var cursor = doc.getCursor();

    if (!cursor) {
      throw new Error("Please click your mouse cursor inside the document body.");
    }

    var element = cursor.getElement();
    var offset = cursor.getOffset();
    var textElement = null;
    var parentParagraph = null;

    // Get the current text element and its parent paragraph
    var elementType = element.getType();
    if (elementType === DocumentApp.ElementType.TEXT) {
      textElement = element.asText();
      parentParagraph = textElement.getParent();
      if (offset === 0 && textElement.getText().length > 0) offset = textElement.getText().length;
    } else if (elementType === DocumentApp.ElementType.PARAGRAPH) {
      var paragraph = element.asParagraph();
      parentParagraph = paragraph;
      if (paragraph.getText() === "") {
        textElement = paragraph.appendText(" ");
        offset = 1;
      } else {
        textElement = paragraph.getChild(0).asText();
        offset = textElement.getText().length;
      }
    } else if (elementType === DocumentApp.ElementType.BODY_SECTION || elementType === DocumentApp.ElementType.DOCUMENT) {
      var targetPara = body.getParagraphs()[0] || body.appendParagraph("");
      parentParagraph = targetPara;
      if (targetPara.getText() === "") {
        textElement = targetPara.appendText(" ");
        offset = 1;
      } else {
        textElement = targetPara.getChild(0).asText();
        offset = textElement.getText().length;
      }
    } else {
      var parent = element.getParent();
      while (parent && parent.getType() !== DocumentApp.ElementType.PARAGRAPH) parent = parent.getParent();
      if (parent) {
        parentParagraph = parent.asParagraph();
        if (parentParagraph.getText() === "") {
          textElement = parentParagraph.appendText(" ");
          offset = 1;
        } else {
          textElement = parentParagraph.getChild(0).asText();
          offset = textElement.getText().length;
        }
      } else {
        throw new Error("Please click inside an active text line.");
      }
    }

    var closeStr = closingTag ? closingTag.trim() : "";
    
    // Determine what content to insert
    var contentToInsert;
    if (content !== undefined && content !== null && content.trim() !== "") {
      contentToInsert = content;
    } else {
      contentToInsert = "[ Enter your conditional template content here ]";
    }

    // ============================================================
    // Get the full text and split
    // ============================================================
    var fullText = textElement.getText();
    var beforeText = fullText.substring(0, offset);
    var afterText = fullText.substring(offset);
    
    // Clear the current text element
    textElement.setText("");
    
    // Insert the text before the cursor
    textElement.appendText(beforeText);
    
    // ============================================================
    // GET DEFAULT FORMATTING FROM THE PARAGRAPH
    // ============================================================
    // Get the paragraph's default formatting
    var defaultFontSize = parentParagraph.getFontSize();
    var defaultForegroundColor = parentParagraph.getForegroundColor();
    var defaultFontFamily = parentParagraph.getFontFamily();
    
    // If paragraph has no explicit formatting, use body defaults
    if (!defaultFontSize) defaultFontSize = 11;
    if (!defaultForegroundColor) defaultForegroundColor = "#000000";
    if (!defaultFontFamily) defaultFontFamily = "Arial";
    
    // ============================================================
    // CREATE SEPARATE TEXT ELEMENTS FOR EACH PART
    // ============================================================
    
    // 1. Opening tag - STYLED
    var openingText = parentParagraph.appendText(openingTag);
    openingText.setForegroundColor("#b0b0b0");
    openingText.setFontSize(9);
    openingText.setItalic(true);
    
    // 2. Content - APPLY DEFAULT FORMATTING EXPLICITLY
    var contentText = parentParagraph.appendText(contentToInsert);
    // Explicitly set to default values (not null)
    contentText.setForegroundColor(defaultForegroundColor);
    contentText.setFontSize(defaultFontSize);
    contentText.setFontFamily(defaultFontFamily);
    contentText.setItalic(false);
    contentText.setBold(false);
    contentText.setUnderline(false);
    
    // 3. Closing tag (if exists) - STYLED
    var closingText = null;
    if (closeStr !== "") {
      closingText = parentParagraph.appendText(closeStr);
      closingText.setForegroundColor("#b0b0b0");
      closingText.setFontSize(9);
      closingText.setItalic(true);
    }
    
    // 4. Remaining text after cursor
    // Ensure there's always a default-formatted run at the end, so grey/italic
    // from the closing tag never becomes the "last format" a new line inherits.
    var afterTextNode = parentParagraph.appendText(afterText.length > 0 ? afterText : " ");
    afterTextNode.setForegroundColor(defaultForegroundColor);
    afterTextNode.setFontSize(defaultFontSize);
    afterTextNode.setFontFamily(defaultFontFamily);
    afterTextNode.setItalic(false);
    afterTextNode.setBold(false);
    afterTextNode.setUnderline(false);
    
    // ============================================================
    // SET CURSOR POSITION - At the end of content
    // ============================================================
    
    var cursorPos = doc.newPosition(contentText, contentText.getText().length);
    doc.setCursor(cursorPos);

    return { success: true };

  } catch (error) {
    Logger.log("Core insertion crash tracker: " + error.toString());
    throw new Error(error.message);
  }
}

// ======================================================================
// 📦 Styling of tags - For existing documents (legacy support)
// ======================================================================
function styleLogicTags() {
  var body = DocumentApp.getActiveDocument().getBody();
  
  var tagDefinitions = [
    { tag: "<<If:", length: 5 },
    { tag: "<<EndIf>>", length: 9 },
    { tag: "<<RowIf:", length: 8 }
  ];
  
  tagDefinitions.forEach(function(item) {
    var tag = item.tag;
    var tagLength = item.length;
    var found = body.findText(tag);
    
    while (found) {
      var text = found.getElement().asText();
      var start = found.getStartOffset();
      var end = found.getEndOffsetInclusive();
      var fullText = text.getText();
      
      // Style the tag only
      var tagEnd = Math.min(start + tagLength - 1, fullText.length - 1);
      text.setForegroundColor(start, tagEnd, "#b0b0b0");
      text.setFontSize(start, tagEnd, 9);
      text.setItalic(start, tagEnd, true);
      
      // Reset formatting for content after tag (prevents bleed)
      var contentStart = tagEnd + 1;
      var contentEnd = fullText.length - 1;
      if (contentStart <= contentEnd) {
        text.setForegroundColor(contentStart, contentEnd, null);
        text.setFontSize(contentStart, contentEnd, null);
        text.setItalic(contentStart, contentEnd, false);
        text.setBold(contentStart, contentEnd, false);
        text.setUnderline(contentStart, contentEnd, false);
      }
      
      found = body.findText(tag, found);
    }
  });
}