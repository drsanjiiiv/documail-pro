/**
 * ConditionEvaluator.gs
 * Safely evaluates logical conditions without risking execution vulnerabilities.
 * THIS FILE MUST BE IN YOUR DOCUMENT TEMPLATE PROJECT (DocuMain.gs project)
 */

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
 * DOCUMAIL PRO CONDITIONAL BLOCK PROCESSOR (FIXED FOR ENGINE.GS)
 * ============================================================================
 * Processes <<If: ... >> ... <<EndIf>> structures on the document canvas.
 * 
 * IMPORTANT: This function expects that {VARIABLES} have ALREADY been replaced
 * with actual values by the time it's called (engine.gs replaces them first).
 * 
 * If called with raw variables still present, it will attempt to replace them.
 * ============================================================================
 */
function processConditionalBlocks(body, rowDataMap) {
  try {
    Logger.log("processConditionalBlocks: Starting with rowDataMap keys: " + Object.keys(rowDataMap).join(', '));
    
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
    Logger.log("processConditionalBlocks: Current text after variable replacement: " + textStr.substring(0, 200) + "...");
    
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