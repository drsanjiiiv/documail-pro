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

// ======================================================================
// 🔥 CONDITIONAL PROCESSING FUNCTIONS (ADD THESE TO DocuMain.gs)
// ======================================================================

/**
 * Main evaluation router. Parses variables and values to run correct logical matches.
 * @param {any} varValue The actual data coming from the Google Sheet row.
 * @param {string} operator The comparison operator (==, !=, >=, <=, contains).
 * @param {string} targetValue The criteria rule value set in the template sidebar.
 * @return {boolean} True if the row data matches the rule parameters.
 */
function evaluateCondition(varValue, operator, targetValue) {
  // Sanitize values to strings for safety, trimming trailing whitespaces
  var currentVal = varValue !== undefined && varValue !== null ? String(varValue).trim() : "";
  var criteriaVal = targetValue !== undefined && targetValue !== null ? String(targetValue).trim() : "";

  // If both values represent numerical quantities, cast them to true floats for perfect math calculations
  var isNumeric = !isNaN(currentVal) && !isNaN(criteriaVal) && currentVal !== "" && criteriaVal !== "";
  
  if (isNumeric) {
    var numCurrent = parseFloat(currentVal);
    var numCriteria = parseFloat(criteriaVal);
    
    switch (operator) {
      case '==': return numCurrent === numCriteria;
      case '!=': return numCurrent !== numCriteria;
      case '>=': return numCurrent >= numCriteria;
      case '<=': return numCurrent <= numCriteria;
      case 'contains': return currentVal.toLowerCase().indexOf(criteriaVal.toLowerCase()) !== -1;
      default: return false;
    }
  }

  // Fallback to strict string comparisons if evaluating textual data fields (e.g. "Pass" vs "Fail")
  switch (operator) {
    case '==': return currentVal.toLowerCase() === criteriaVal.toLowerCase();
    case '!=': return currentVal.toLowerCase() !== criteriaVal.toLowerCase();
    case '>=': return currentVal.toLowerCase() >= criteriaVal.toLowerCase();
    case '<=': return currentVal.toLowerCase() <= criteriaVal.toLowerCase();
    case 'contains': return currentVal.toLowerCase().indexOf(criteriaVal.toLowerCase()) !== -1;
    default: return false;
  }
}

/**
 * ============================================================================
 * DOCUMAIL PRO CONDITIONAL BLOCK PROCESSOR
 * ============================================================================
 * Processes <<If: ... >> ... <<EndIf>> structures on the document canvas.
 * 
 * IMPORTANT: This function expects that {VARIABLES} have ALREADY been replaced
 * with actual values by engine.gs. If called with raw variables still present,
 * it will attempt to replace them using rowDataMap.
 * ============================================================================
 */
function processConditionalBlocks(body, rowDataMap) {
  try {
    Logger.log("processConditionalBlocks: Starting with rowDataMap keys: " + Object.keys(rowDataMap || {}).join(', '));
    
    if (!rowDataMap) {
      Logger.log("processConditionalBlocks: No rowDataMap provided, skipping");
      return;
    }
    
    // STEP 1: Check if there are any unprocessed {VARIABLES} in the document
    // This handles the case where engine.gs might not have replaced them yet
    var fullText = body.getText();
    var variableRegex = /\{([^}]+)\}/g;
    var variableMatch;
    var hasUnprocessedVariables = false;
    
    while ((variableMatch = variableRegex.exec(fullText)) !== null) {
      hasUnprocessedVariables = true;
      break;
    }
    
    // If there are unprocessed variables, replace them with values from rowDataMap
    if (hasUnprocessedVariables) {
      Logger.log("processConditionalBlocks: Found unprocessed variables, replacing them first");
      variableRegex.lastIndex = 0; // Reset regex
      var variablesToReplace = new Set();
      
      while ((variableMatch = variableRegex.exec(fullText)) !== null) {
        variablesToReplace.add(variableMatch[1]);
      }
      
      variablesToReplace.forEach(function(varName) {
        var placeholder = "{" + varName + "}";
        var replacementValue = rowDataMap && rowDataMap.hasOwnProperty(varName) ? 
                              String(rowDataMap[varName]) : "";
        body.replaceText(escapeRegexString(placeholder), replacementValue);
        Logger.log("processConditionalBlocks: Replaced " + placeholder + " with '" + replacementValue + "'");
      });
    }
    
    // STEP 2: Now process the conditional blocks with the replaced values
    var textStr = body.getText();
    Logger.log("processConditionalBlocks: Current text preview: " + textStr.substring(0, 200) + "...");
    
    // Matches <<If: ... >> ... <<EndIf>> smoothly across paragraph line breaks
    var blockRegex = /<<If:\s*([^>]+)>>([\s\S]*?)<<EndIf>>/gi;
    var match;
    var modifications = [];
    
    while ((match = blockRegex.exec(textStr)) !== null) {
      var fullMatchText = match[0];
      var conditionExpression = match[1].trim();
      var conditionalContent = match[2];
      
      Logger.log("processConditionalBlocks: Found conditional block: " + conditionExpression);
      
      // Parse the condition expression
      // Handles patterns like: "Deepak == 'Sandeep'" or "{NAME} == 'Sandeep'"
      var syntaxRegex = /([^=!><=contains'\s]+(?:[^=!><=contains']*[^=!><=contains'\s])?)\s*([=!><=contains]+)\s*['"“‘]([^'"”’]+)['"”’]/i;
      var syntaxMatch = syntaxRegex.exec(conditionExpression);
      
      var conditionMet = false;
      
      if (syntaxMatch) {
        var leftSide = syntaxMatch[1].trim();
        var operator = syntaxMatch[2].trim();
        var targetValue = syntaxMatch[3].trim();
        
        Logger.log("processConditionalBlocks: Parsed - leftSide: '" + leftSide + "', operator: '" + operator + "', targetValue: '" + targetValue + "'");
        
        // The left side might be:
        // 1. A raw value like "Deepak" (already replaced by engine.gs)
        // 2. A variable placeholder like "{NAME}" (if replacement didn't happen)
        var liveValue = leftSide;
        
        // If it still looks like a variable placeholder, try to get it from rowDataMap directly
        if (leftSide.startsWith('{') && leftSide.endsWith('}')) {
          var varName = leftSide.slice(1, -1);
          liveValue = rowDataMap && rowDataMap.hasOwnProperty(varName) ? 
                     String(rowDataMap[varName]) : leftSide;
          Logger.log("processConditionalBlocks: Resolved placeholder to: '" + liveValue + "'");
        }
        
        conditionMet = evaluateCondition(liveValue, operator, targetValue);
        Logger.log("processConditionalBlocks: Condition result: " + conditionMet);
      } else {
        Logger.log("processConditionalBlocks: Could not parse condition expression: " + conditionExpression);
      }
      
      if (conditionMet) {
        // Condition is True: Keep the content, remove the conditional wrapper
        Logger.log("processConditionalBlocks: Condition TRUE - keeping content");
        modifications.push({
          match: fullMatchText,
          replacement: conditionalContent
        });
      } else {
        // Condition is False: Remove everything
        Logger.log("processConditionalBlocks: Condition FALSE - removing content");
        modifications.push({
          match: fullMatchText,
          replacement: ""
        });
      }
    }
    
    // Apply all modifications (process from last to first to maintain indices)
    for (var i = modifications.length - 1; i >= 0; i--) {
      var mod = modifications[i];
      body.replaceText(escapeRegexString(mod.match), mod.replacement);
      Logger.log("processConditionalBlocks: Applied modification " + i);
    }
    
    Logger.log("processConditionalBlocks: Completed successfully");
    
  } catch (err) {
    Logger.log("Error evaluating block conditions: " + err.toString());
    console.error("Error evaluating block conditions: " + err.toString());
    throw err;
  }
}

/**
 * Process conditional table rows - handles <<RowIf: ... >> ... <<EndRowIf>> 
 * for table row-level conditions
 */
function processConditionalTableRows(body, rowDataMap) {
  try {
    Logger.log("processConditionalTableRows: Starting");
    
    if (!rowDataMap) {
      Logger.log("processConditionalTableRows: No rowDataMap provided, skipping");
      return;
    }
    
    // First, replace any unprocessed variables in the document
    var fullText = body.getText();
    var variableRegex = /\{([^}]+)\}/g;
    var variableMatch;
    var hasUnprocessedVariables = false;
    
    while ((variableMatch = variableRegex.exec(fullText)) !== null) {
      hasUnprocessedVariables = true;
      break;
    }
    
    if (hasUnprocessedVariables) {
      Logger.log("processConditionalTableRows: Found unprocessed variables, replacing them first");
      variableRegex.lastIndex = 0;
      var variablesToReplace = new Set();
      
      while ((variableMatch = variableRegex.exec(fullText)) !== null) {
        variablesToReplace.add(variableMatch[1]);
      }
      
      variablesToReplace.forEach(function(varName) {
        var placeholder = "{" + varName + "}";
        var replacementValue = rowDataMap && rowDataMap.hasOwnProperty(varName) ? 
                              String(rowDataMap[varName]) : "";
        body.replaceText(escapeRegexString(placeholder), replacementValue);
      });
    }
    
    // Process table row conditions
    var tables = body.getTables();
    for (var t = 0; t < tables.length; t++) {
      var table = tables[t];
      var rows = table.getRows();
      var rowsToRemove = [];
      
      for (var r = 0; r < rows.length; r++) {
        var row = rows[r];
        var rowText = row.getText();
        
        // Check if this row has a condition
        var conditionMatch = rowText.match(/<<RowIf:\s*([^>]+)>>/i);
        if (conditionMatch) {
          var conditionExpression = conditionMatch[1].trim();
          
          // Parse the condition
          var syntaxRegex = /([^=!><=contains'\s]+(?:[^=!><=contains']*[^=!><=contains'\s])?)\s*([=!><=contains]+)\s*['"“‘]([^'"”’]+)['"”’]/i;
          var syntaxMatch = syntaxRegex.exec(conditionExpression);
          
          var conditionMet = false;
          
          if (syntaxMatch) {
            var leftSide = syntaxMatch[1].trim();
            var operator = syntaxMatch[2].trim();
            var targetValue = syntaxMatch[3].trim();
            
            var liveValue = leftSide;
            
            // If it's a variable placeholder, resolve it
            if (leftSide.startsWith('{') && leftSide.endsWith('}')) {
              var varName = leftSide.slice(1, -1);
              liveValue = rowDataMap && rowDataMap.hasOwnProperty(varName) ? 
                         String(rowDataMap[varName]) : leftSide;
            }
            
            conditionMet = evaluateCondition(liveValue, operator, targetValue);
          }
          
          // If condition is false, mark this row for removal
          if (!conditionMet) {
            rowsToRemove.push(r);
          }
        }
      }
      
      // Remove rows from bottom to top to maintain indices
      for (var i = rowsToRemove.length - 1; i >= 0; i--) {
        table.removeRow(rowsToRemove[i]);
        Logger.log("processConditionalTableRows: Removed row " + rowsToRemove[i]);
      }
    }
    
    Logger.log("processConditionalTableRows: Completed successfully");
    
  } catch (err) {
    Logger.log("Error processing conditional table rows: " + err.toString());
    console.error("Error processing conditional table rows: " + err.toString());
  }
}

/**
 * Helper to escape special regular expression characters safely out of raw template strings
 */
function escapeRegexString(str) {
  return str.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
}

// ======================================================================
// 📦 OPTIONAL: Template Preview Functions (If you want to test in document)
// ======================================================================

/**
 * Test function to preview conditional processing with sample data
 * Call this from the script editor to test your template
 */
function TEST_CONDITIONAL_PROCESSING() {
  try {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    
    // Sample data for testing
    var testData = {
      'NAME': 'Deepak',
      'EMAIL': 'deepak@example.com',
      'AMOUNT': '1000',
      'STATUS': 'Active'
    };
    
    Logger.log("Testing conditional processing with data: " + JSON.stringify(testData));
    processConditionalBlocks(body, testData);
    processConditionalTableRows(body, testData);
    
    DocumentApp.getUi().alert('✅ Test completed! Check the document to see the results.');
    
  } catch (error) {
    DocumentApp.getUi().alert('❌ Error: ' + error.toString());
    Logger.log("Test error: " + error.toString());
  }
}

/**
 * Debug function to check what variables are in the document
 */
function DEBUG_CHECK_TEMPLATE_VARIABLES() {
  try {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var text = body.getText();
    
    var variableRegex = /\{([^}]+)\}/g;
    var match;
    var variables = [];
    
    while ((match = variableRegex.exec(text)) !== null) {
      if (variables.indexOf(match[1]) === -1) {
        variables.push(match[1]);
      }
    }
    
    var ui = DocumentApp.getUi();
    if (variables.length === 0) {
      ui.alert('ℹ️ No variables found in the document.');
    } else {
      ui.alert('📋 Variables found:\n\n' + variables.join('\n'));
    }
    
    return variables;
    
  } catch (error) {
    Logger.log("Debug error: " + error.toString());
    return [];
  }
}