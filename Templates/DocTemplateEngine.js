/**
 * DocuMain.gs (Inside your Document Template Project)
 * Main Controller for the Document Template Designer.
 */

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
 * Universal token injector engine. Natively avoids empty text element insertion bugs.
 */
function injectTagAtCursor(openingTag, closingTag) {
  try {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var cursor = doc.getCursor();
    
    if (!cursor) {
      throw new Error("Please click your mouse cursor inside the document body where you want to insert this layout token.");
    }
    
    var element = cursor.getElement();
    var offset = cursor.getOffset();
    var textElement = null;
    
    var elementType = element.getType();
    
    if (elementType === DocumentApp.ElementType.TEXT) {
      textElement = element.asText();
      
      // Focus drift protection
      if (offset === 0 && textElement.getText().length > 0) {
        offset = textElement.getText().length;
      }
      
    } else if (elementType === DocumentApp.ElementType.PARAGRAPH) {
      var paragraph = element.asParagraph();
      // SAFE FIX: If paragraph is empty, initialize with a space instead of an empty string
      if (paragraph.getText() === "") {
        textElement = paragraph.appendText(" ");
        offset = 1;
      } else {
        textElement = paragraph.getChild(0).asText();
        offset = textElement.getText().length;
      }
    } else if (elementType === DocumentApp.ElementType.BODY_SECTION || elementType === DocumentApp.ElementType.DOCUMENT) {
      var targetPara = body.getParagraphs()[0] || body.appendParagraph("");
      if (targetPara.getText() === "") {
        textElement = targetPara.appendText(" ");
        offset = 1;
      } else {
        textElement = targetPara.getChild(0).asText();
        offset = textElement.getText().length;
      }
    } else {
      var parent = element.getParent();
      while (parent && parent.getType() !== DocumentApp.ElementType.PARAGRAPH) {
        parent = parent.getParent();
      }
      if (parent) {
        var paragraph = parent.asParagraph();
        if (paragraph.getText() === "") {
          textElement = paragraph.appendText(" ");
          offset = 1;
        } else {
          textElement = paragraph.getChild(0).asText();
          offset = textElement.getText().length;
        }
      } else {
        throw new Error("Please click inside an active text line on the page.");
      }
    }
    
    var closeStr = closingTag ? closingTag.trim() : "";
    
    if (closeStr !== "") {
      // Wrapper block logic
      textElement.insertText(offset, openingTag);
      
      var parentElement = textElement.getParent();
      while (parentElement && parentElement.getType() !== DocumentApp.ElementType.PARAGRAPH && parentElement.getType() !== DocumentApp.ElementType.LIST_ITEM) {
        parentElement = parentElement.getParent();
      }
      
      var container = parentElement.getParent(); 
      var structuralIndex = container.getChildIndex(parentElement);
      
      var closePara = container.insertParagraph(structuralIndex + 1, closeStr);
      var placeholderPara = container.insertParagraph(structuralIndex + 1, "[ Enter your conditional template content here ]");
      placeholderPara.setItalic(true);
      
      var nextPosition = doc.newPosition(placeholderPara.getChild(0), 2);
      doc.setCursor(nextPosition);
      
    } else {
      // Inline token insertion
      textElement.insertText(offset, openingTag);
      var newPosition = doc.newPosition(textElement, offset + openingTag.length);
      doc.setCursor(newPosition);
    }
    
    return { success: true };
    
  } catch (error) {
    Logger.log("Core insertion crash tracker: " + error.toString());
    throw new Error(error.message);
  }
}