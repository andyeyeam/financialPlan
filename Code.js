function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  var updateSubmenu = ui.createMenu('Update sheet')
      .addItem('Load transactions from Barclays', 'addDumpedTransactions')
      .addItem('Load transactions from Sainsburys', 'addSainsburysTransactions')
      .addItem('Load transactions from American Express', 'addAmexTransactions')
      .addItem('Update levels', 'mapLevel2');

  var mappingsSubmenu = ui.createMenu('Mappings')
      .addItem('Generate', 'autoGenerateLevel2Mappings')
      .addItem('Add', 'addLevel2Mapping')
      .addItem('Add from current', 'addMappingFromCurrent');

  var reportsSubmenu = ui.createMenu('Reports')
      .addItem('Create ABV Sheet', 'createABVSheet')
      .addItem('Actual vs Planned Report', 'createActualVsPlannedReport');

  ui.createMenu('Custom Menu')
      .addSubMenu(updateSubmenu)
      .addSubMenu(mappingsSubmenu)
      .addSubMenu(reportsSubmenu)
      .addToUi();
}

function addDumpedTransactions() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
    var dumpSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Barclays Current Account');
    
    // Get data more efficiently
    var dumpValues = dumpSheet.getDataRange().getValues();
    var currValues = sheet.getDataRange().getValues();
    var outValues = [];

    // Early exit if no dump data
    if (dumpValues.length <= 1) {
      SpreadsheetApp.getUi().alert('No Data', 'No transactions found in Barclays Current Account sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    var lastColumn = sheet.getLastColumn();

    // Create lookup set for faster duplicate checking
    var existingTransactions = new Set();
    for (var i = 1; i < currValues.length; i++) {
      if (currValues[i][0] && currValues[i][2] && currValues[i][8]) {
        var key = String(currValues[i][0]) + '|' + String(currValues[i][2]).replace(/\s+/g, '').toLowerCase() + '|' + Math.abs(currValues[i][8]);
        existingTransactions.add(key);
      }
    }

    // Process transactions with progress indicator
    var totalTransactions = dumpValues.length - 1; // Exclude header row
    var processedCount = 0;
    var addedCount = 0;
    
    // Show initial progress
    SpreadsheetApp.getActiveSpreadsheet().toast('Processing 0 of ' + totalTransactions + ' transactions...', 'Processing Transactions', 3);
    
    for (var dvRow = 1; dvRow < dumpValues.length; dvRow++){
      processedCount++;
      
      // Update progress every 10 transactions or for small batches
      if (processedCount % 10 === 0 || processedCount === totalTransactions || totalTransactions <= 20) {
        SpreadsheetApp.getActiveSpreadsheet().toast('Processing ' + processedCount + ' of ' + totalTransactions + ' transactions...', 'Processing Transactions', 3);
      }
      
      // Create lookup key for this dump transaction
      var dumpKey = String(dumpValues[dvRow][1]) + '|' + String(dumpValues[dvRow][5]).replace(/\s+/g, '').toLowerCase() + '|' + Math.abs(dumpValues[dvRow][3]);
      
      if (!existingTransactions.has(dumpKey)) {
        // Calculate the row number for the new row that will be added
        var newRowNumber = sheet.getLastRow() + 1 + outValues.length;
        
        // Build new row with data and formulas
        var newRow = new Array(lastColumn);
        
        // Column A (0): Date
        newRow[0] = dumpValues[dvRow][1];
        
        // Column B (1): Month formula
        newRow[1] = '=TEXT(A' + newRowNumber + ',"MM")';
        
        // Column C (2): Description
        newRow[2] = dumpValues[dvRow][5];
        
        // Column D (3): Account
        newRow[3] = 'Barclays Current Account';
        
        // Column E (4): Level 0
        newRow[4] = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$C$256,3,0),"")';
        
        // Column F (5): Level 1
        newRow[5] = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$B$250,2,0),"")';
        
        // Column G (6): Level 2
        newRow[6] = '';
        
        // Column H (7): DR/CR formula
        newRow[7] = '=IF(I' + newRowNumber + '>=0,"DR","CR")';
        
        // Column I (8): Amount
        newRow[8] = dumpValues[dvRow][3];
        
        // Column J (9): Absolute Amount formula
        newRow[9] = '=ABS(I' + newRowNumber + ')';
        
        // Column K (10): Item
        newRow[10] = '';
        
        // Column L (11): Subcategory
        if (lastColumn > 11) newRow[11] = dumpValues[dvRow][4];
        
        outValues.push(newRow);
        addedCount++;
      }
    }

    if (outValues.length > 0) {
      // Get the last row that contains data.
      var startRow = sheet.getLastRow() + 1;

      // Append the data to the sheet in one batch operation
      var range = sheet.getRange(startRow, 1, outValues.length, lastColumn);
      range.setValues(outValues);
      
      SpreadsheetApp.getUi().alert('Success', 'Completed processing ' + totalTransactions + ' transactions.\nAdded ' + outValues.length + ' new transactions to the Transactions sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('No New Transactions', 'Completed processing ' + totalTransactions + ' transactions.\nAll transactions from Barclays Current Account already exist in the Transactions sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'An error occurred while processing transactions: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function addSainsburysTransactions() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
    var sainsburysSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sainsburys Bank Credit Card');
    
    // Get data more efficiently
    var sainsburysValues = sainsburysSheet.getDataRange().getValues();
    var currValues = sheet.getDataRange().getValues();
    var outValues = [];

    // Early exit if no data
    if (sainsburysValues.length <= 1) {
      SpreadsheetApp.getUi().alert('No Data', 'No transactions found in Sainsburys Bank Credit Card sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    var lastColumn = sheet.getLastColumn();

    // Create lookup set for faster duplicate checking
    var existingTransactions = new Set();
    for (var i = 1; i < currValues.length; i++) {
      if (currValues[i][0] && currValues[i][2] && currValues[i][8]) {
        var key = String(currValues[i][0]) + '|' + String(currValues[i][2]).replace(/\s+/g, '').toLowerCase() + '|' + Math.abs(currValues[i][8]);
        existingTransactions.add(key);
      }
    }

    // Process Sainsburys transactions with progress indicator
    var totalTransactions = sainsburysValues.length - 1; // Exclude header row
    var processedCount = 0;
    
    // Show initial progress
    SpreadsheetApp.getActiveSpreadsheet().toast('Processing 0 of ' + totalTransactions + ' transactions...', 'Processing Transactions', 3);
    
    for (var svRow = 1; svRow < sainsburysValues.length; svRow++){
      processedCount++;
      
      // Update progress every 10 transactions or for small batches
      if (processedCount % 10 === 0 || processedCount === totalTransactions || totalTransactions <= 20) {
        SpreadsheetApp.getActiveSpreadsheet().toast('Processing ' + processedCount + ' of ' + totalTransactions + ' transactions...', 'Processing Transactions', 3);
      }
      
      // Sainsburys format: Date (0), Description (1), Amount (2), DR/CR (3)
      var transDate = sainsburysValues[svRow][0];
      var transDesc = sainsburysValues[svRow][1];
      var transAmount = sainsburysValues[svRow][2];
      var drCr = sainsburysValues[svRow][3];
      
      // Convert amount - if CR (credit) make it positive, otherwise negative
      var numAmount = 0;
      if (typeof transAmount === 'string') {
        numAmount = parseFloat(transAmount.replace(/[£,]/g, ''));
      } else {
        numAmount = transAmount;
      }
      
      // For credit cards: debits are purchases (negative), credits are payments (positive)
      if (drCr && drCr.toString().trim().toUpperCase() === 'CR') {
        numAmount = Math.abs(numAmount); // Credits are positive
      } else {
        numAmount = -Math.abs(numAmount); // Debits are negative
      }
      
      // Check for duplicates by matching Date, Description, and Amount exactly
      var isDuplicate = false;
      
      for (var k = 1; k < currValues.length; k++) {
        var existingDate = currValues[k][0]; // Column A - Date
        var existingDesc = currValues[k][2];  // Column C - Description
        var existingAmount = currValues[k][8]; // Column I - Amount
        
        // Compare dates (convert to same format for comparison)
        var transDateStr = '';
        var existingDateStr = '';
        
        if (transDate instanceof Date) {
          transDateStr = transDate.toDateString();
        } else {
          transDateStr = new Date(transDate).toDateString();
        }
        
        if (existingDate instanceof Date) {
          existingDateStr = existingDate.toDateString();
        } else {
          existingDateStr = new Date(existingDate).toDateString();
        }
        
        // Check if Date, Description, and Amount all match
        if (transDateStr === existingDateStr &&
            String(transDesc).trim() === String(existingDesc).trim() &&
            numAmount === existingAmount) {
          isDuplicate = true;
          break;
        }
      }
      
      if (!isDuplicate) {
        // Calculate the row number for the new row that will be added
        var newRowNumber = sheet.getLastRow() + 1 + outValues.length;
        
        // Build new row with data and formulas
        var newRow = new Array(lastColumn);
        
        // Column A (0): Date
        newRow[0] = transDate;
        
        // Column B (1): Month formula
        newRow[1] = '=TEXT(A' + newRowNumber + ',"MM")';
        
        // Column C (2): Description
        newRow[2] = transDesc;
        
        // Column D (3): Account
        newRow[3] = 'Sainsburys Bank Credit Card';
        
        // Column E (4): Level 0
        newRow[4] = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$C$256,3,0),"")';
        
        // Column F (5): Level 1
        newRow[5] = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$B$250,2,0),"")';
        
        // Column G (6): Level 2
        newRow[6] = '';
        
        // Column H (7): DR/CR formula
        newRow[7] = '=IF(I' + newRowNumber + '>=0,"DR","CR")';
        
        // Column I (8): Amount
        newRow[8] = numAmount;
        
        // Column J (9): Absolute Amount formula
        newRow[9] = '=ABS(I' + newRowNumber + ')';
        
        // Column K (10): Item
        newRow[10] = '';
        
        // Column L (11): Subcategory
        if (lastColumn > 11) newRow[11] = '';
        
        outValues.push(newRow);
      }
    }

    if (outValues.length > 0) {
      // Get the last row that contains data.
      var startRow = sheet.getLastRow() + 1;

      // Append the data to the sheet in one batch operation
      var range = sheet.getRange(startRow, 1, outValues.length, lastColumn);
      range.setValues(outValues);
      
      SpreadsheetApp.getUi().alert('Success', 'Completed processing ' + totalTransactions + ' transactions.\nAdded ' + outValues.length + ' new transactions to the Transactions sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('No New Transactions', 'Completed processing ' + totalTransactions + ' transactions.\nAll transactions from Sainsburys Bank Credit Card already exist in the Transactions sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'An error occurred while processing transactions: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function addAmexTransactions() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
    var amexSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('American Express Credit Card');
    
    // Get data more efficiently
    var amexValues = amexSheet.getDataRange().getValues();
    var currValues = sheet.getDataRange().getValues();
    var outValues = [];

    // Early exit if no data
    if (amexValues.length <= 1) {
      SpreadsheetApp.getUi().alert('No Data', 'No transactions found in American Express Credit Card sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    var lastColumn = sheet.getLastColumn();

    // Create lookup set for faster duplicate checking
    var existingTransactions = new Set();
    for (var i = 1; i < currValues.length; i++) {
      if (currValues[i][0] && currValues[i][2] && currValues[i][8]) {
        var key = String(currValues[i][0]) + '|' + String(currValues[i][2]).replace(/\s+/g, '').toLowerCase() + '|' + Math.abs(currValues[i][8]);
        existingTransactions.add(key);
      }
    }

    // Process American Express transactions with progress indicator
    var totalTransactions = amexValues.length - 1; // Exclude header row
    var processedCount = 0;
    
    // Show initial progress
    SpreadsheetApp.getActiveSpreadsheet().toast('Processing 0 of ' + totalTransactions + ' transactions...', 'Processing Transactions', 3);
    
    for (var axRow = 1; axRow < amexValues.length; axRow++){
      processedCount++;
      
      // Update progress every 10 transactions or for small batches
      if (processedCount % 10 === 0 || processedCount === totalTransactions || totalTransactions <= 20) {
        SpreadsheetApp.getActiveSpreadsheet().toast('Processing ' + processedCount + ' of ' + totalTransactions + ' transactions...', 'Processing Transactions', 3);
      }
      
      // American Express format: DATE (0), STATUS (1), DESCRIPTION (2), AMOUNT (3)
      var transDate = amexValues[axRow][0];
      var transStatus = amexValues[axRow][1];
      var transDesc = amexValues[axRow][1]; // Use STATUS field for description
      var transAmount = amexValues[axRow][3];
      
      // Convert amount - handle negative values and currency symbols
      var numAmount = 0;
      if (typeof transAmount === 'string') {
        numAmount = parseFloat(transAmount.replace(/[£,]/g, ''));
      } else {
        numAmount = transAmount;
      }
      
      // For American Express credit cards:
      // - "Credit" status transactions are payments (should be positive)
      // - Regular transactions are purchases (should be negative)
      // - Amount already includes the sign in the data
      if (transStatus && transStatus.toString().trim().toUpperCase() === 'CREDIT') {
        numAmount = Math.abs(numAmount); // Credits are positive
      } else {
        // For purchases, if amount is positive, make it negative
        if (numAmount > 0) {
          numAmount = -numAmount;
        }
      }
      
      // Check for duplicates by matching Date, Description, and Amount exactly
      var isDuplicate = false;
      
      for (var k = 1; k < currValues.length; k++) {
        var existingDate = currValues[k][0]; // Column A - Date
        var existingDesc = currValues[k][2];  // Column C - Description
        var existingAmount = currValues[k][8]; // Column I - Amount
        
        // Compare dates (convert to same format for comparison)
        var transDateStr = '';
        var existingDateStr = '';
        
        if (transDate instanceof Date) {
          transDateStr = transDate.toDateString();
        } else {
          transDateStr = new Date(transDate).toDateString();
        }
        
        if (existingDate instanceof Date) {
          existingDateStr = existingDate.toDateString();
        } else {
          existingDateStr = new Date(existingDate).toDateString();
        }
        
        // Check if Date, Description, and Amount all match
        if (transDateStr === existingDateStr &&
            String(transDesc).trim() === String(existingDesc).trim() &&
            numAmount === existingAmount) {
          isDuplicate = true;
          break;
        }
      }
      
      if (!isDuplicate) {
        // Calculate the row number for the new row that will be added
        var newRowNumber = sheet.getLastRow() + 1 + outValues.length;
        
        // Build new row with data and formulas
        var newRow = new Array(lastColumn);
        
        // Column A (0): Date
        newRow[0] = transDate;
        
        // Column B (1): Month formula
        newRow[1] = '=TEXT(A' + newRowNumber + ',"MM")';
        
        // Column C (2): Description
        newRow[2] = transDesc;
        
        // Column D (3): Account
        newRow[3] = 'American Express Credit Card';
        
        // Column E (4): Level 0
        newRow[4] = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$C$256,3,0),"")';
        
        // Column F (5): Level 1
        newRow[5] = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$B$250,2,0),"")';
        
        // Column G (6): Level 2
        newRow[6] = '';
        
        // Column H (7): DR/CR formula
        newRow[7] = '=IF(I' + newRowNumber + '>=0,"DR","CR")';
        
        // Column I (8): Amount
        newRow[8] = numAmount;
        
        // Column J (9): Absolute Amount formula
        newRow[9] = '=ABS(I' + newRowNumber + ')';
        
        // Column K (10): Item
        newRow[10] = '';
        
        // Column L (11): Subcategory
        if (lastColumn > 11) newRow[11] = '';
        
        outValues.push(newRow);
      }
    }

    if (outValues.length > 0) {
      // Get the last row that contains data.
      var startRow = sheet.getLastRow() + 1;

      // Append the data to the sheet in one batch operation
      var range = sheet.getRange(startRow, 1, outValues.length, lastColumn);
      range.setValues(outValues);
      
      SpreadsheetApp.getUi().alert('Success', 'Completed processing ' + totalTransactions + ' transactions.\nAdded ' + outValues.length + ' new transactions to the Transactions sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('No New Transactions', 'Completed processing ' + totalTransactions + ' transactions.\nAll transactions from American Express Credit Card already exist in the Transactions sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'An error occurred while processing transactions: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function mapLevel2 (){

  // Load the map into memory
  var mapValues  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Level 2 Mapping').getDataRange().getValues();

  // Load the transaction into memory and loop round them
  var currValues  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions').getDataRange().getValues();
  var updatedCount = 0;
  
  for (var i = 1; i < currValues.length; i++){      // Start at 1 to skip the transactions header row (row 0)
    // Skip if already has a Level 2 value (column G is index 6)
    if (currValues[i][6] && currValues[i][6].toString().trim() !== "") continue;
    
    for (var j = 1; j < mapValues.length; j++){     // Start at 1 to skip the map header row (row 0)
      if (!currValues[i][2].includes (mapValues[j][0])) continue;
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions').getRange(i + 1, 7).setValue(mapValues[j][1]);
      updatedCount++;
      break; // Stop after first match to avoid multiple updates to the same transaction
    }
  }
  
  // Count remaining blank Level 2 values after updates
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  var updatedValues = sheet.getDataRange().getValues();
  var remainingBlankCount = 0;
  
  for (var i = 1; i < updatedValues.length; i++){
    // Check if Level 2 value (column G, index 6) is blank
    if (!updatedValues[i][6] || updatedValues[i][6].toString().trim() === "") {
      remainingBlankCount++;
    }
  }
  
  // Show completion message with both counts
  var message = 'Updated ' + updatedCount + ' Level 2 values in the Transactions sheet.\n';
  message += remainingBlankCount + ' blank Level 2 values remain.';
  SpreadsheetApp.getUi().alert('Update Levels Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function rowMatch (dRow, cRow) {
  // Date comparison - handle both Date objects and strings
  var dumpDate = dRow[1];
  var transDate = cRow[0];
  
  // Convert dump date to Date object if it's a string
  if (typeof dumpDate === 'string') {
    dumpDate = new Date(dumpDate);
  }
  
  // Convert transaction date to Date object if it's a string
  if (typeof transDate === 'string') {
    transDate = new Date(transDate);
  }
  
  // Compare dates (normalize to same day)
  if (dumpDate.toDateString() !== transDate.toDateString()) return false;
  
  // Amount comparison - use absolute values and fix column index
  // dRow[3] = dump amount, cRow[8] = transaction amount (column I is index 8)
  if (Math.abs(dRow[3]) != Math.abs(cRow[8])) return false;
  
  // Description comparison - normalize whitespace and case
  // dRow[5] = dump memo, cRow[2] = transaction description
  var dumpDesc = String(dRow[5]).replace(/\s+/g, '').toLowerCase().trim();
  var transDesc = String(cRow[2]).replace(/\s+/g, '').toLowerCase().trim();
  if (dumpDesc !== transDesc) return false;
 
  return true;
}

function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  
  // Only process if we're on the Transactions sheet
  if (sheet.getName() !== 'Transactions') return;
  
  // Check if a new row was added (when editing in a previously empty row)
  var editedRow = range.getRow();
  var lastRowWithData = sheet.getLastRow();
  
  // If we're editing beyond the current data range, it's likely a new row
  if (editedRow > lastRowWithData - 1 || isNewRowAdded(sheet, editedRow)) {
    copyFormulasToNewRow(sheet, editedRow);
  }
}

function isNewRowAdded(sheet, editedRow) {
  // Check if this row appears to be newly populated
  var rowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var nonEmptyCount = 0;
  
  for (var i = 0; i < rowData.length; i++) {
    if (rowData[i] !== '') {
      nonEmptyCount++;
    }
  }
  
  // If only a few cells are filled, it's likely a new row being added
  return nonEmptyCount <= 3;
}

function copyFormulasToNewRow(sheet, newRowNumber) {
  if (newRowNumber <= 1) return; // Skip header row
  
  var lastColumn = sheet.getLastColumn();
  
  // Create specific formulas for each column
  for (var col = 1; col <= lastColumn; col++) {
    var targetCell = sheet.getRange(newRowNumber, col);
    var formula = '';
    
    switch (col) {
      case 2: // Column B - Month formula
        formula = '=TEXT(A' + newRowNumber + ',"MM")';
        break;
      case 5: // Column E - Level 0
        formula = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$C$256,3,0),"")';
        break;
      case 6: // Column F - Level 1
        formula = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$B$250,2,0),"")';
        break;
      case 8: // Column H - DR/CR formula
        formula = '=IF(I' + newRowNumber + '>=0,"DR","CR")';
        break;
      case 10: // Column J - Absolute Amount
        formula = '=ABS(I' + newRowNumber + ')';
        break;
    }
    
    // Set formula if one is defined for this column
    if (formula) {
      targetCell.setFormula(formula);
    }
  }
}

function hasFormulas(sheet, rowNumber) {
  var lastColumn = sheet.getLastColumn();
  for (var col = 1; col <= lastColumn; col++) {
    if (sheet.getRange(rowNumber, col).getFormula()) {
      return true;
    }
  }
  return false;
}

function updateFormulaRowReferences(formula, newRowNumber) {
  if (!formula || formula === '') return '';
  
  // Replace row references in the formula, but preserve absolute row references exactly
  // Pattern: Column letters + optional $ (for column) + optional $ (for row) + any row number
  
  var updatedFormula = formula.replace(/([A-Z]+)(\$?)(\$?)(\d+)\b/g, function(match, columnLetters, columnDollar, rowDollar, rowNum) {
    if (rowDollar === '$') {
      // Absolute row reference - copy exactly as is (keep the original row number)
      return match; // Return the entire match unchanged
    } else {
      // Relative row reference - update to new row number
      return columnLetters + columnDollar + newRowNumber;
    }
  });
  
  return updatedFormula;
}

function autoGenerateLevel2Mappings() {
  var transSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  var mappingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Level 2 Mapping');
  
  // Get all transactions data
  var transValues = transSheet.getDataRange().getValues();
  
  // Get existing mappings to avoid duplicates
  var existingMappings = mappingSheet.getDataRange().getValues();
  var existingPatterns = new Set();
  
  for (var i = 1; i < existingMappings.length; i++) {
    if (existingMappings[i][0]) {
      existingPatterns.add(existingMappings[i][0].toLowerCase());
    }
  }
  
  // Analyze transactions with Level 2 values to find patterns
  var patternAnalysis = {};
  
  for (var i = 1; i < transValues.length; i++) {
    var description = transValues[i][2]; // Column C - Description
    var level2 = transValues[i][6]; // Column G - Level 2
    
    if (description && level2) {
      var patterns = extractDescriptionPatterns(description);
      
      for (var j = 0; j < patterns.length; j++) {
        var pattern = patterns[j];
        
        if (!patternAnalysis[pattern]) {
          patternAnalysis[pattern] = {};
        }
        
        if (!patternAnalysis[pattern][level2]) {
          patternAnalysis[pattern][level2] = 0;
        }
        
        patternAnalysis[pattern][level2]++;
      }
    }
  }
  
  // Find reliable patterns (appear multiple times with same Level 2)
  var newMappings = [];
  
  for (var pattern in patternAnalysis) {
    var level2Counts = patternAnalysis[pattern];
    var level2Options = Object.keys(level2Counts);
    
    // Only consider if pattern appears at least 2 times
    var totalCount = 0;
    var dominantLevel2 = '';
    var maxCount = 0;
    
    for (var level2 in level2Counts) {
      totalCount += level2Counts[level2];
      if (level2Counts[level2] > maxCount) {
        maxCount = level2Counts[level2];
        dominantLevel2 = level2;
      }
    }
    
    // Pattern is reliable if it appears at least 2 times and 80% of occurrences map to the same Level 2
    if (totalCount >= 2 && (maxCount / totalCount) >= 0.8 && !existingPatterns.has(pattern.toLowerCase())) {
      newMappings.push([pattern, dominantLevel2]);
    }
  }
  
  // Add new mappings to the Level 2 Mapping sheet
  if (newMappings.length > 0) {
    var startRow = mappingSheet.getLastRow() + 1;
    var range = mappingSheet.getRange(startRow, 1, newMappings.length, 2);
    range.setValues(newMappings);
    
    SpreadsheetApp.getUi().alert('Success', 'Added ' + newMappings.length + ' new Level 2 mappings to the mapping sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('No New Mappings', 'No new reliable patterns were found to add to the mapping sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function extractDescriptionPatterns(description) {
  var patterns = [];
  var cleanDesc = String(description).trim().toUpperCase();
  
  // Extract meaningful substrings that can be used with "contains" logic
  
  // Pattern 1: Well-known merchant/service names (exact substrings)
  var knownMerchants = [
    'SAINSBURYS', 'TESCO', 'ASDA', 'MORRISONS', 'WAITROSE', 'MARKS', 'SPENCER',
    'AMAZON', 'PAYPAL', 'NETFLIX', 'SPOTIFY', 'APPLE', 'GOOGLE', 'MICROSOFT',
    'TFL', 'UBER', 'DELIVEROO', 'JUST EAT', 'MCDONALD', 'KFC', 'SUBWAY',
    'BP', 'SHELL', 'ESSO', 'TEXACO', 'BOOTS', 'SUPERDRUG', 'ARGOS',
    'CURRYS', 'JOHN LEWIS', 'NEXT', 'H&M', 'ZARA', 'PRIMARK'
  ];
  
  for (var i = 0; i < knownMerchants.length; i++) {
    if (cleanDesc.includes(knownMerchants[i])) {
      patterns.push(knownMerchants[i]);
    }
  }
  
  // Pattern 2: Transaction type indicators
  var transactionTypes = [
    { substring: 'FT', pattern: 'FT' },
    { substring: 'DD', pattern: 'DD' }, 
    { substring: 'SO', pattern: 'SO' },
    { substring: 'TRANSFER', pattern: 'TRANSFER' },
    { substring: 'DIRECT DEBIT', pattern: 'DIRECT DEBIT' },
    { substring: 'STANDING ORDER', pattern: 'STANDING ORDER' },
    { substring: 'CARD PAYMENT', pattern: 'CARD PAYMENT' },
    { substring: 'CASH WITHDRAWAL', pattern: 'CASH WITHDRAWAL' }
  ];
  
  for (var i = 0; i < transactionTypes.length; i++) {
    if (cleanDesc.includes(transactionTypes[i].substring)) {
      patterns.push(transactionTypes[i].pattern);
    }
  }
  
  // Pattern 3: Extract clean merchant names from common formats
  // Format: "MERCHANT NAME    LOCATION/CODE"
  var merchantMatch = cleanDesc.match(/^([A-Z][A-Z\s&]{2,20}?)[\s]{2,}/);
  if (merchantMatch) {
    var merchantName = merchantMatch[1].trim();
    // Only add if it's not already covered by known merchants and is meaningful
    if (merchantName.length >= 4 && merchantName.length <= 15) {
      var alreadyHave = false;
      for (var i = 0; i < patterns.length; i++) {
        if (patterns[i].includes(merchantName) || merchantName.includes(patterns[i])) {
          alreadyHave = true;
          break;
        }
      }
      if (!alreadyHave) {
        patterns.push(merchantName);
      }
    }
  }
  
  // Pattern 4: Extract first meaningful word if nothing else found
  if (patterns.length === 0) {
    var words = cleanDesc.split(/\s+/);
    for (var i = 0; i < words.length && i < 2; i++) {
      var word = words[i].replace(/[^A-Z]/g, ''); // Remove non-letters
      if (word.length >= 4 && word.length <= 12) {
        patterns.push(word);
        break;
      }
    }
  }
  
  // Return unique patterns, filtered for meaningful length
  var uniquePatterns = [];
  for (var i = 0; i < patterns.length; i++) {
    var pattern = patterns[i].trim();
    if (pattern.length >= 3 && pattern.length <= 20 && uniquePatterns.indexOf(pattern) === -1) {
      uniquePatterns.push(pattern);
    }
  }
  
  return uniquePatterns;
}

function addLevel2Mapping() {
  try {
    var transSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
    if (!transSheet) {
      SpreadsheetApp.getUi().alert('Transactions Sheet Missing', 'Could not find the Transactions sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Get all transactions data
    var transValues = transSheet.getDataRange().getValues();
    var unmappedTransactions = [];
    
    // Find all unique descriptions with blank Level 2 values
    var seenDescriptions = new Set();
    
    for (var i = 1; i < transValues.length; i++) {
      var description = transValues[i][2]; // Column C - Description
      var level2 = transValues[i][6]; // Column G - Level 2
      var absoluteAmount = transValues[i][9]; // Column J - Absolute Amount
      
      // Check if description exists and Level 2 is blank/empty
      if (description && description.toString().trim() !== '') {
        var descStr = description.toString().trim();
        
        // More comprehensive check for empty Level 2
        var isLevel2Empty = !level2 || 
                           level2 === null || 
                           level2 === undefined || 
                           level2.toString().trim() === '' ||
                           level2.toString().trim() === '0' ||
                           level2.toString().toLowerCase() === 'null';
        
        if (isLevel2Empty) {
          var descKey = descStr;
          if (!seenDescriptions.has(descKey)) {
            seenDescriptions.add(descKey);
            unmappedTransactions.push({
              description: descStr,
              amount: absoluteAmount || 0
            });
          }
        }
      }
    }
    
    if (unmappedTransactions.length === 0) {
      SpreadsheetApp.getUi().alert('No Unmapped Transactions', 'All transactions already have Level 2 values assigned.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Get Level 2 values from Taxonomy sheet
    var taxonomySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Taxonomy');
    if (!taxonomySheet) {
      SpreadsheetApp.getUi().alert('Taxonomy Sheet Missing', 'Could not find the Taxonomy sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var taxonomyValues = taxonomySheet.getRange('A:A').getValues();
    var level2Options = [];
    
    // Extract non-empty values from column A (skip header)
    for (var i = 1; i < taxonomyValues.length; i++) {
      if (taxonomyValues[i][0] && taxonomyValues[i][0].toString().trim() !== '') {
        level2Options.push(taxonomyValues[i][0].toString().trim());
      }
    }
    
    if (level2Options.length === 0) {
      SpreadsheetApp.getUi().alert('No Level 2 Options', 'No Level 2 categories found in the Taxonomy sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Sort Level 2 options alphabetically
    level2Options.sort();
    
    // Create HTML dialog with table
    var html = createLevel2MappingTableDialog(unmappedTransactions, level2Options);
    var htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(900)
        .setHeight(600);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Add Level 2 Mappings');
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'An error occurred: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function createLevel2MappingTableDialog(unmappedTransactions, level2Options) {
  var optionsHtml = '';
  for (var i = 0; i < level2Options.length; i++) {
    optionsHtml += '<option value="' + level2Options[i] + '">' + level2Options[i] + '</option>';
  }
  
  // Generate table rows
  var tableRowsHtml = '';
  for (var i = 0; i < unmappedTransactions.length; i++) {
    var transaction = unmappedTransactions[i];
    var formattedAmount = '';
    
    // Format the amount as currency if it's a number
    if (typeof transaction.amount === 'number') {
      formattedAmount = '£' + Math.abs(transaction.amount).toFixed(2);
    } else if (transaction.amount) {
      formattedAmount = transaction.amount.toString();
    }
    
    tableRowsHtml += `
      <tr>
        <td><input type="text" id="desc_${i}" value="${transaction.description}" /></td>
        <td class="amount-cell">${formattedAmount}</td>
        <td>
          <select id="level2_${i}">
            <option value="">-- Select Level 2 Category --</option>
            ${optionsHtml}
          </select>
        </td>
      </tr>
    `;
  }
  
  var html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        th, td { padding: 8px; border: 1px solid #ddd; text-align: left; }
        th { background-color: #f2f2f2; font-weight: bold; cursor: pointer; position: relative; user-select: none; }
        th:hover { background-color: #e8e8e8; }
        .sort-indicator { position: absolute; right: 8px; top: 50%; transform: translateY(-50%); opacity: 0.5; }
        .sort-asc .sort-indicator::after { content: '▲'; }
        .sort-desc .sort-indicator::after { content: '▼'; }
        input, select { width: 100%; padding: 4px; border: 1px solid #ccc; border-radius: 4px; }
        .amount-cell { text-align: right; font-family: monospace; background-color: #f9f9f9; font-weight: bold; }
        td:nth-child(1) { width: 50%; }
        td:nth-child(2) { width: 15%; }
        td:nth-child(3) { width: 35%; }
        .buttons { text-align: right; margin-top: 20px; }
        button { padding: 8px 16px; margin-left: 8px; border: none; border-radius: 4px; cursor: pointer; }
        .btn-save { background-color: #4CAF50; color: white; }
        .btn-cancel { background-color: #f44336; color: white; }
        .container { max-height: 400px; overflow-y: auto; }
        .spinner { text-align: center; margin: 20px 0; }
        .spinner-icon {
          border: 4px solid #f3f3f3;
          border-top: 4px solid #4CAF50;
          border-radius: 50%;
          width: 30px;
          height: 30px;
          animation: spin 1s linear infinite;
          margin: 0 auto 10px auto;
        }
        .spinner-text { font-size: 14px; color: #666; }
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
        button:disabled { opacity: 0.6; cursor: not-allowed; }
      </style>
    </head>
    <body>
      <h3>Add Level 2 Mappings</h3>
      <p>Edit descriptions and select Level 2 categories for unmapped transactions:</p>
      
      <div class="container">
        <table id="mappingTable">
          <thead>
            <tr>
              <th onclick="sortTable(0)" data-column="0">
                Description Pattern
                <span class="sort-indicator"></span>
              </th>
              <th onclick="sortTable(1)" data-column="1">
                Amount
                <span class="sort-indicator"></span>
              </th>
              <th onclick="sortTable(2)" data-column="2">
                Level 2 Category
                <span class="sort-indicator"></span>
              </th>
            </tr>
          </thead>
          <tbody>
            ${tableRowsHtml}
          </tbody>
        </table>
      </div>
      
      <div class="buttons">
        <button type="button" class="btn-cancel" onclick="google.script.host.close()">Cancel</button>
        <button type="button" class="btn-save" id="saveBtn" onclick="saveMappings()">Save</button>
      </div>
      
      <div id="spinner" class="spinner" style="display: none;">
        <div class="spinner-icon"></div>
        <div class="spinner-text">Saving mappings...</div>
      </div>
      
      <script>
        let sortOrder = {}; // Track sort order for each column
        
        function sortTable(columnIndex) {
          const table = document.getElementById('mappingTable');
          const tbody = table.getElementsByTagName('tbody')[0];
          const rows = Array.from(tbody.getElementsByTagName('tr'));
          const headers = table.getElementsByTagName('th');
          
          // Toggle sort order
          if (!sortOrder[columnIndex]) {
            sortOrder[columnIndex] = 'asc';
          } else if (sortOrder[columnIndex] === 'asc') {
            sortOrder[columnIndex] = 'desc';
          } else {
            sortOrder[columnIndex] = 'asc';
          }
          
          // Clear all header sort classes
          for (let i = 0; i < headers.length; i++) {
            headers[i].classList.remove('sort-asc', 'sort-desc');
          }
          
          // Add sort class to current header
          headers[columnIndex].classList.add('sort-' + sortOrder[columnIndex]);
          
          // Sort rows
          rows.sort((a, b) => {
            let aValue, bValue;
            
            if (columnIndex === 0) {
              // Description - text sort on input value
              aValue = a.cells[columnIndex].getElementsByTagName('input')[0].value.toLowerCase();
              bValue = b.cells[columnIndex].getElementsByTagName('input')[0].value.toLowerCase();
            } else if (columnIndex === 1) {
              // Amount - numeric sort
              aValue = parseFloat(a.cells[columnIndex].textContent.replace(/[£,]/g, '')) || 0;
              bValue = parseFloat(b.cells[columnIndex].textContent.replace(/[£,]/g, '')) || 0;
            } else if (columnIndex === 2) {
              // Level 2 Category - text sort on selected value
              const selectA = a.cells[columnIndex].getElementsByTagName('select')[0];
              const selectB = b.cells[columnIndex].getElementsByTagName('select')[0];
              aValue = selectA.options[selectA.selectedIndex].text.toLowerCase();
              bValue = selectB.options[selectB.selectedIndex].text.toLowerCase();
            }
            
            if (sortOrder[columnIndex] === 'asc') {
              return aValue < bValue ? -1 : aValue > bValue ? 1 : 0;
            } else {
              return aValue > bValue ? -1 : aValue < bValue ? 1 : 0;
            }
          });
          
          // Re-append sorted rows to tbody
          rows.forEach(row => tbody.appendChild(row));
        }

        function saveMappings() {
          var mappings = [];
          var rowCount = ${unmappedTransactions.length};
          
          for (var i = 0; i < rowCount; i++) {
            var description = document.getElementById('desc_' + i).value.trim();
            var level2 = document.getElementById('level2_' + i).value;
            
            if (description && level2) {
              mappings.push({description: description, level2: level2});
            }
          }
          
          if (mappings.length === 0) {
            alert('No complete mappings to save. Please fill in both description and Level 2 category for at least one row.');
            return;
          }
          
          // Show spinner and disable buttons
          showSpinner();
          
          google.script.run
            .withSuccessHandler(function(result) {
              hideSpinner();
              alert('Success: Added ' + result + ' new mappings.');
              google.script.host.close();
            })
            .withFailureHandler(function(error) {
              hideSpinner();
              alert('Error saving mappings: ' + error.message);
            })
            .saveMultipleLevel2Mappings(mappings);
        }
        
        function showSpinner() {
          document.getElementById('spinner').style.display = 'block';
          document.getElementById('saveBtn').disabled = true;
          document.getElementById('saveBtn').textContent = 'Saving...';
          document.querySelector('.btn-cancel').disabled = true;
        }
        
        function hideSpinner() {
          document.getElementById('spinner').style.display = 'none';
          document.getElementById('saveBtn').disabled = false;
          document.getElementById('saveBtn').textContent = 'Save';
          document.querySelector('.btn-cancel').disabled = false;
        }
      </script>
    </body>
    </html>
  `;
  
  return html;
}

function saveLevel2Mapping(description, level2) {
  try {
    var mappingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Level 2 Mapping');
    if (!mappingSheet) {
      throw new Error('Level 2 Mapping sheet not found.');
    }
    
    // Check if this mapping already exists (case-sensitive)
    var existingMappings = mappingSheet.getDataRange().getValues();
    for (var i = 1; i < existingMappings.length; i++) {
      if (existingMappings[i][0] && existingMappings[i][0].toString() === description) {
        throw new Error('A mapping for "' + description + '" already exists.');
      }
    }
    
    // Add the new mapping to the end of the list
    var lastRow = mappingSheet.getLastRow() + 1;
    mappingSheet.getRange(lastRow, 1, 1, 2).setValues([[description, level2]]);
    
    SpreadsheetApp.getUi().alert('Success', 'Mapping added: "' + description + '" → "' + level2 + '"', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    throw error;
  }
}

function saveMultipleLevel2Mappings(mappings) {
  try {
    var mappingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Level 2 Mapping');
    if (!mappingSheet) {
      throw new Error('Level 2 Mapping sheet not found.');
    }
    
    // Get existing mappings to check for duplicates (case-sensitive)
    var existingMappings = mappingSheet.getDataRange().getValues();
    var existingDescriptions = new Set();
    
    for (var i = 1; i < existingMappings.length; i++) {
      if (existingMappings[i][0]) {
        existingDescriptions.add(existingMappings[i][0].toString());
      }
    }
    
    // Filter out duplicates and prepare data for insertion
    var newMappings = [];
    var duplicates = [];
    
    for (var i = 0; i < mappings.length; i++) {
      var description = mappings[i].description.trim();
      var level2 = mappings[i].level2.trim();
      
      if (existingDescriptions.has(description)) {
        duplicates.push(description);
      } else {
        newMappings.push([description, level2]);
        existingDescriptions.add(description); // Prevent duplicates within this batch
      }
    }
    
    // Add new mappings to the sheet
    if (newMappings.length > 0) {
      var startRow = mappingSheet.getLastRow() + 1;
      mappingSheet.getRange(startRow, 1, newMappings.length, 2).setValues(newMappings);
    }
    
    return newMappings.length; // Return count of successfully added mappings
    
  } catch (error) {
    throw error;
  }
}

function addMappingFromCurrent() {
  try {
    // Check if the current sheet is the Transactions sheet
    var activeSheet = SpreadsheetApp.getActiveSheet();
    if (activeSheet.getName() !== 'Transactions') {
      SpreadsheetApp.getUi().alert('Wrong Sheet', 'Please select the Transactions sheet before using this function.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Get the currently selected range
    var activeRange = activeSheet.getActiveRange();
    var currentRow = activeRange.getRow();
    
    // Check if a valid transaction row is selected (not header row)
    if (currentRow <= 1) {
      SpreadsheetApp.getUi().alert('Invalid Row', 'Please select a transaction row (not the header row).', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Get the Description (Column C, index 2) and Level 2 (Column G, index 6) values from current row
    var description = activeSheet.getRange(currentRow, 3).getValue(); // Column C
    var level2 = activeSheet.getRange(currentRow, 7).getValue();       // Column G
    
    // Validate that description exists and is not empty
    if (!description || description.toString().trim() === '') {
      SpreadsheetApp.getUi().alert('Missing Description', 'The selected row does not have a description in Column C.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var descriptionStr = description.toString().trim();
    var level2Str = '';
    
    // If Level 2 is missing, show dialog to let user specify it
    if (!level2 || level2.toString().trim() === '') {
      // Get Level 2 options from Taxonomy sheet
      var taxonomySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Taxonomy');
      if (!taxonomySheet) {
        SpreadsheetApp.getUi().alert('Taxonomy Sheet Missing', 'Could not find the Taxonomy sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
      
      var taxonomyValues = taxonomySheet.getRange('A:A').getValues();
      var level2Options = [];
      
      // Extract non-empty values from column A (skip header)
      for (var i = 1; i < taxonomyValues.length; i++) {
        if (taxonomyValues[i][0] && taxonomyValues[i][0].toString().trim() !== '') {
          level2Options.push(taxonomyValues[i][0].toString().trim());
        }
      }
      
      if (level2Options.length === 0) {
        SpreadsheetApp.getUi().alert('No Level 2 Options', 'No Level 2 categories found in the Taxonomy sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
      
      // Sort options alphabetically
      level2Options.sort();
      
      // Show dialog to select Level 2
      var html = createLevel2SelectDialog(descriptionStr, level2Options);
      var htmlOutput = HtmlService.createHtmlOutput(html)
          .setWidth(500)
          .setHeight(400);
      
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select Level 2 Category');
      return; // Exit here - the dialog will handle the save via callback
      
    } else {
      level2Str = level2.toString().trim();
    }
    
    // If we have both description and level2, proceed with saving
    saveMappingFromCurrent(descriptionStr, level2Str);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'An error occurred: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function createLevel2SelectDialog(description, level2Options) {
  var optionsHtml = '';
  for (var i = 0; i < level2Options.length; i++) {
    optionsHtml += '<option value="' + level2Options[i] + '">' + level2Options[i] + '</option>';
  }
  
  var html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; }
        .form-group { margin: 15px 0; }
        label { display: block; margin-bottom: 5px; font-weight: bold; }
        input, select { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; }
        .description-field { background-color: #f9f9f9; }
        .buttons { text-align: right; margin-top: 20px; }
        button { padding: 8px 16px; margin-left: 8px; border: none; border-radius: 4px; cursor: pointer; }
        .btn-save { background-color: #4CAF50; color: white; }
        .btn-cancel { background-color: #f44336; color: white; }
        .spinner { text-align: center; margin: 20px 0; display: none; }
        .spinner-icon {
          border: 4px solid #f3f3f3;
          border-top: 4px solid #4CAF50;
          border-radius: 50%;
          width: 30px;
          height: 30px;
          animation: spin 1s linear infinite;
          margin: 0 auto 10px auto;
        }
        @keyframes spin {
          0% { transform: rotate(0deg); }
          100% { transform: rotate(360deg); }
        }
        button:disabled { opacity: 0.6; cursor: not-allowed; }
      </style>
    </head>
    <body>
      <h3>Add Level 2 Mapping</h3>
      <p>Create a mapping for the selected transaction:</p>
      
      <div class="form-group">
        <label for="description">Description Pattern:</label>
        <input type="text" id="description" value="${description}" class="description-field" />
      </div>
      
      <div class="form-group">
        <label for="level2">Level 2 Category:</label>
        <select id="level2">
          <option value="">-- Select Level 2 Category --</option>
          ${optionsHtml}
        </select>
      </div>
      
      <div class="buttons">
        <button type="button" class="btn-cancel" onclick="google.script.host.close()">Cancel</button>
        <button type="button" class="btn-save" id="saveBtn" onclick="saveMapping()">Save Mapping</button>
      </div>
      
      <div id="spinner" class="spinner">
        <div class="spinner-icon"></div>
        <div>Saving mapping...</div>
      </div>
      
      <script>
        function saveMapping() {
          var description = document.getElementById('description').value.trim();
          var level2 = document.getElementById('level2').value;
          
          if (!description) {
            alert('Please enter a description pattern.');
            return;
          }
          
          if (!level2) {
            alert('Please select a Level 2 category.');
            return;
          }
          
          // Show spinner and disable buttons
          showSpinner();
          
          google.script.run
            .withSuccessHandler(function(result) {
              hideSpinner();
              alert('Success: Mapping added for "' + description + '" → "' + level2 + '"');
              google.script.host.close();
            })
            .withFailureHandler(function(error) {
              hideSpinner();
              alert('Error saving mapping: ' + error.message);
            })
            .saveMappingFromCurrent(description, level2);
        }
        
        function showSpinner() {
          document.getElementById('spinner').style.display = 'block';
          document.getElementById('saveBtn').disabled = true;
          document.getElementById('saveBtn').textContent = 'Saving...';
          document.querySelector('.btn-cancel').disabled = true;
        }
        
        function hideSpinner() {
          document.getElementById('spinner').style.display = 'none';
          document.getElementById('saveBtn').disabled = false;
          document.getElementById('saveBtn').textContent = 'Save Mapping';
          document.querySelector('.btn-cancel').disabled = false;
        }
      </script>
    </body>
    </html>
  `;
  
  return html;
}

function saveMappingFromCurrent(descriptionStr, level2Str) {
  try {
    // Check if mapping already exists
    var mappingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Level 2 Mapping');
    if (!mappingSheet) {
      throw new Error('Could not find the Level 2 Mapping sheet.');
    }
    
    // Check for existing mapping pair (case-sensitive)
    var existingMappings = mappingSheet.getDataRange().getValues();
    for (var i = 1; i < existingMappings.length; i++) {
      if (existingMappings[i][0] && existingMappings[i][1]) {
        var existingDesc = existingMappings[i][0].toString().trim();
        var existingLevel2 = existingMappings[i][1].toString().trim();
        
        if (existingDesc === descriptionStr && existingLevel2 === level2Str) {
          throw new Error('A mapping already exists for "' + descriptionStr + '" → "' + level2Str + '".');
        }
      }
    }
    
    // Add the new mapping
    var lastRow = mappingSheet.getLastRow() + 1;
    mappingSheet.getRange(lastRow, 1, 1, 2).setValues([[descriptionStr, level2Str]]);
    
    return 'Successfully added mapping: "' + descriptionStr + '" → "' + level2Str + '".';
    
  } catch (error) {
    throw error;
  }
}

function createABVSheet() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Check if ABV sheet already exists and clear it, otherwise create it
    var abvSheet = spreadsheet.getSheetByName('ABV');
    if (abvSheet) {
      // Clear all content and formatting
      abvSheet.clear();
    } else {
      // Create new ABV sheet
      abvSheet = spreadsheet.insertSheet('ABV');
    }
    
    // Get taxonomy data
    var taxonomySheet = spreadsheet.getSheetByName('Taxonomy');
    if (!taxonomySheet) {
      SpreadsheetApp.getUi().alert('Error', 'Taxonomy sheet not found.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var taxonomyValues = taxonomySheet.getDataRange().getValues();
    var uniquePairs = new Set();
    var categoryPairs = [];
    
    // Parse header row to find month columns
    var monthColumns = {};
    var headerRow = taxonomyValues[0];
    var monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 
                     'July', 'August', 'September', 'October', 'November', 'December'];
    
    for (var col = 2; col < headerRow.length; col++) {
      if (headerRow[col] !== null && headerRow[col] !== undefined && headerRow[col] !== '') {
        var headerValue = headerRow[col];
        var monthName = null;
        
        // Check if it's a month integer (1-12)
        if (typeof headerValue === 'number' && headerValue >= 1 && headerValue <= 12) {
          monthName = monthNames[headerValue - 1]; // Convert 1-based to 0-based index
        } 
        // Check if it's a string representation of month integer
        else if (typeof headerValue === 'string') {
          var monthInt = parseInt(headerValue.trim());
          if (!isNaN(monthInt) && monthInt >= 1 && monthInt <= 12) {
            monthName = monthNames[monthInt - 1];
          } else {
            // Map various month text formats to standard month names
            var headerLower = headerValue.toString().trim().toLowerCase();
            var monthMapping = {
              'jan': 'January', 'january': 'January',
              'feb': 'February', 'february': 'February',
              'mar': 'March', 'march': 'March',
              'apr': 'April', 'april': 'April',
              'may': 'May',
              'jun': 'June', 'june': 'June',
              'jul': 'July', 'july': 'July',
              'aug': 'August', 'august': 'August',
              'sep': 'September', 'september': 'September',
              'oct': 'October', 'october': 'October',
              'nov': 'November', 'november': 'November',
              'dec': 'December', 'december': 'December'
            };
            
            if (monthMapping[headerLower]) {
              monthName = monthMapping[headerLower];
            }
          }
        }
        
        if (monthName) {
          monthColumns[monthName] = col;
        }
      }
    }
    
    // Extract unique Level 1 and Level 2 pairs with their planned amounts (skip header row)
    for (var i = 1; i < taxonomyValues.length; i++) {
      var level2 = taxonomyValues[i][0]; // Column A - Level 2
      var level1 = taxonomyValues[i][1]; // Column B - Level 1
      
      if (level2 && level1) {
        var level2Str = level2.toString().trim();
        var level1Str = level1.toString().trim();
        var pairKey = level1Str + '|' + level2Str;
        
        if (!uniquePairs.has(pairKey)) {
          uniquePairs.add(pairKey);
          
          // Get planned amounts for each month
          var plannedAmounts = {};
          for (var monthName in monthColumns) {
            var colIndex = monthColumns[monthName];
            var plannedValue = taxonomyValues[i][colIndex];
            plannedAmounts[monthName] = (plannedValue && typeof plannedValue === 'number') ? plannedValue : 0;
          }
          
          categoryPairs.push({
            level1: level1Str,
            level2: level2Str,
            plannedAmounts: plannedAmounts
          });
        }
      }
    }
    
    if (categoryPairs.length === 0) {
      SpreadsheetApp.getUi().alert('No Data', 'No Level 1 and Level 2 pairs found in the Taxonomy sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Set up headers
    var headers = ['Level 0', 'Level 1', 'Level 2', 'Month', 'Type', 'Amount'];
    abvSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format headers
    var headerRange = abvSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4472C4')
               .setFontColor('#FFFFFF')
               .setFontWeight('bold')
               .setHorizontalAlignment('center');
    
    // Prepare data rows
    var dataRows = [];
    var monthNumbers = ['01', '02', '03', '04', '05', '06', 
                        '07', '08', '09', '10', '11', '12'];
    
    // Generate rows for each category pair
    for (var p = 0; p < categoryPairs.length; p++) {
      var pair = categoryPairs[p];
      
      // Generate 24 rows for each pair (12 months × 2 types)
      for (var month = 1; month <= 12; month++) {
        var monthNumber = monthNumbers[month - 1];
        
        // Get planned amount for this month from taxonomy data - convert month number to name for lookup
        var monthName = ['January', 'February', 'March', 'April', 'May', 'June', 
                        'July', 'August', 'September', 'October', 'November', 'December'][month - 1];
        var plannedAmount = pair.plannedAmounts[monthName] || 0;
        
        // Planned row
        var plannedRowNumber = 2 + dataRows.length; // Current row being added
        dataRows.push([
          `=IFNA(VLOOKUP(C${plannedRowNumber},Taxonomy!$A$2:$C$256,3,0),"")`, // Level 0 lookup
          pair.level1,
          pair.level2,
          monthNumber,
          'Planned',
          plannedAmount
        ]);
        
        // Actual row with SUMIFS formula to aggregate from Transactions
        var actualRowNumber = 2 + dataRows.length; // Current row being added (no +1 needed)
        // Match: Transactions Column B (month as number) with ABV Column D (month as number)
        //        Transactions Column F (Level 1) with ABV Column B (Level 1) 
        //        Transactions Column G (Level 2) with ABV Column C (Level 2)
        //        Sum: Transactions Column J (Absolute Amount)
        var sumFormula = `=SUMIFS(Transactions!J:J,Transactions!B:B,"${monthNumber}",Transactions!F:F,B${actualRowNumber},Transactions!G:G,C${actualRowNumber})`;
        
        dataRows.push([
          `=IFNA(VLOOKUP(C${actualRowNumber},Taxonomy!$A$2:$C$256,3,0),"")`, // Level 0 lookup
          pair.level1,
          pair.level2,
          monthNumber,
          'Actual',
          sumFormula
        ]);
      }
    }
    
    // Write data to sheet
    if (dataRows.length > 0) {
      var dataRange = abvSheet.getRange(2, 1, dataRows.length, headers.length);
      dataRange.setValues(dataRows);
      
      // Format the data area
      dataRange.setBorder(true, true, true, true, true, true);
      
      // Format Amount column as currency
      var amountColumn = abvSheet.getRange(2, 6, dataRows.length, 1);
      amountColumn.setNumberFormat('£#,##0.00');
      
      // Alternate row colors for better readability
      for (var row = 2; row <= dataRows.length + 1; row += 2) {
        abvSheet.getRange(row, 1, 1, headers.length).setBackground('#F8F9FA');
      }
    }
    
    // Auto-resize columns
    abvSheet.autoResizeColumns(1, headers.length);
    
    // Freeze header row
    abvSheet.setFrozenRows(1);
    
    // Set column widths for better appearance
    abvSheet.setColumnWidth(1, 120); // Level 0
    abvSheet.setColumnWidth(2, 150); // Level 1
    abvSheet.setColumnWidth(3, 150); // Level 2
    abvSheet.setColumnWidth(4, 100); // Month
    abvSheet.setColumnWidth(5, 80);  // Type
    abvSheet.setColumnWidth(6, 100); // Amount
    
    var totalRows = dataRows.length;
    var totalPairs = categoryPairs.length;
    
    SpreadsheetApp.getUi().alert(
      'Success', 
      'ABV sheet created successfully!\n\n' +
      'Generated ' + totalRows + ' rows for ' + totalPairs + ' category pairs.\n' +
      'Each pair has 24 rows (12 months × 2 types: Planned and Actual).', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'An error occurred while creating the ABV sheet: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function createActualVsPlannedReport() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    
    // Check if ABV sheet exists
    var abvSheet = spreadsheet.getSheetByName('ABV');
    var shouldRecreateABV = false;
    
    if (!abvSheet) {
      // If ABV sheet doesn't exist, ask if they want to create it
      var createResponse = ui.alert(
        'ABV Sheet Not Found', 
        'The ABV sheet does not exist. Would you like to create it now?', 
        ui.ButtonSet.YES_NO
      );
      
      if (createResponse === ui.Button.YES) {
        createABVSheet();
        abvSheet = spreadsheet.getSheetByName('ABV');
        if (!abvSheet) {
          ui.alert('Error', 'Failed to create ABV sheet.', ui.ButtonSet.OK);
          return;
        }
      } else {
        return; // User chose not to create ABV sheet
      }
    } else {
      // ABV sheet exists, ask if they want to recreate it
      var recreateResponse = ui.alert(
        'Recreate ABV Data?', 
        'Would you like to recreate the ABV data with the latest information before generating the report?\n\nChoose:\n• YES - Refresh ABV data then create report\n• NO - Use existing ABV data for report', 
        ui.ButtonSet.YES_NO
      );
      
      shouldRecreateABV = (recreateResponse === ui.Button.YES);
    }
    
    // Recreate ABV data if requested
    if (shouldRecreateABV) {
      ui.alert('Refreshing Data', 'Recreating ABV sheet with latest data...', ui.ButtonSet.OK);
      createABVSheet();
      
      // Re-get the ABV sheet after recreation
      abvSheet = spreadsheet.getSheetByName('ABV');
      if (!abvSheet) {
        ui.alert('Error', 'Failed to recreate ABV sheet.', ui.ButtonSet.OK);
        return;
      }
    }
    
    // Get ABV data
    var abvData = abvSheet.getDataRange().getValues();
    if (abvData.length <= 1) {
      ui.alert('Error', 'No data found in ABV sheet. Please check your Taxonomy sheet and try recreating the ABV data.', ui.ButtonSet.OK);
      return;
    }
    
    // Build hierarchical data structure
    var reportData = buildReportDataStructure(abvData);
    
    // Create HTML report
    var html = createActualVsPlannedReportHTML(reportData);
    var htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(1200)
        .setHeight(800);
    
    ui.showModalDialog(htmlOutput, 'Actual vs Planned Report');
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'An error occurred while creating the report: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function buildReportDataStructure(abvData) {
  var data = {};
  var monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 
                    'July', 'August', 'September', 'October', 'November', 'December'];
  
  // Extract current year from transactions or use current year
  var currentYear = new Date().getFullYear();
  
  // Initialize Level 0 categories
  var level0Categories = ['Income', 'Expenditure', 'Neutral'];
  
  // Process ABV data (skip header row)
  for (var i = 1; i < abvData.length; i++) {
    var level0 = abvData[i][0];     // Level 0
    var level1 = abvData[i][1];     // Level 1
    var level2 = abvData[i][2];     // Level 2
    var monthValue = abvData[i][3];  // Month (now numeric like "01", "02")
    var type = abvData[i][4];       // Type (Planned/Actual)
    var amount = abvData[i][5];     // Amount
    
    if (!level1 || !level2 || !monthValue || !type) continue;
    
    // Ensure level0 has a value, default to 'Expenditure' if empty
    if (!level0 || level0.toString().trim() === '') {
      level0 = 'Expenditure';
    }
    
    // Convert amount to number if it's a string
    var numAmount = 0;
    if (typeof amount === 'number') {
      numAmount = amount;
    } else if (typeof amount === 'string' && amount.trim() !== '') {
      numAmount = parseFloat(amount.replace(/[£,]/g, '')) || 0;
    }
    
    // Convert month value to number (e.g., "01" -> 1, "02" -> 2)
    var monthNum = parseInt(monthValue);
    if (!monthNum || monthNum < 1 || monthNum > 12) continue;
    
    // Get month name for display purposes
    var monthName = monthNames[monthNum - 1];
    
    // Initialize Level 0 structure
    if (!data[level0]) {
      data[level0] = {
        planned: 0,
        actual: 0,
        variance: 0,
        years: {}
      };
    }
    
    // Initialize year structure under Level 0
    if (!data[level0].years[currentYear]) {
      data[level0].years[currentYear] = {
        planned: 0,
        actual: 0,
        variance: 0,
        months: {},
        level1s: {}
      };
    }
    
    // Initialize month structure under Level 0 > Year
    if (!data[level0].years[currentYear].months[monthNum]) {
      data[level0].years[currentYear].months[monthNum] = {
        name: monthName,
        planned: 0,
        actual: 0,
        variance: 0,
        level1s: {}
      };
    }
    
    // Initialize Level 1 structure under Level 0 > Year
    if (!data[level0].years[currentYear].level1s[level1]) {
      data[level0].years[currentYear].level1s[level1] = {
        planned: 0,
        actual: 0,
        variance: 0,
        months: {},
        level2s: {}
      };
    }
    
    // Initialize Level 1 structure under Level 0 > Year > Month
    if (!data[level0].years[currentYear].months[monthNum].level1s[level1]) {
      data[level0].years[currentYear].months[monthNum].level1s[level1] = {
        planned: 0,
        actual: 0,
        variance: 0,
        level2s: {}
      };
    }
    
    // Initialize Level 1 months under Level 0 > Year > Level 1
    if (!data[level0].years[currentYear].level1s[level1].months[monthNum]) {
      data[level0].years[currentYear].level1s[level1].months[monthNum] = {
        name: monthName,
        planned: 0,
        actual: 0,
        variance: 0,
        level2s: {}
      };
    }
    
    // Initialize Level 2 structures
    if (!data[level0].years[currentYear].level1s[level1].level2s[level2]) {
      data[level0].years[currentYear].level1s[level1].level2s[level2] = {
        planned: 0,
        actual: 0,
        variance: 0,
        months: {}
      };
    }
    
    if (!data[level0].years[currentYear].months[monthNum].level1s[level1].level2s[level2]) {
      data[level0].years[currentYear].months[monthNum].level1s[level1].level2s[level2] = {
        planned: 0,
        actual: 0,
        variance: 0
      };
    }
    
    if (!data[level0].years[currentYear].level1s[level1].months[monthNum].level2s[level2]) {
      data[level0].years[currentYear].level1s[level1].months[monthNum].level2s[level2] = {
        planned: 0,
        actual: 0,
        variance: 0
      };
    }
    
    if (!data[level0].years[currentYear].level1s[level1].level2s[level2].months[monthNum]) {
      data[level0].years[currentYear].level1s[level1].level2s[level2].months[monthNum] = {
        name: monthName,
        planned: 0,
        actual: 0,
        variance: 0
      };
    }
    
    // Add amounts to appropriate structures
    if (type === 'Planned') {
      // Level 0 level
      data[level0].planned += numAmount;
      // Year level
      data[level0].years[currentYear].planned += numAmount;
      // Month level
      data[level0].years[currentYear].months[monthNum].planned += numAmount;
      // Level 1 level
      data[level0].years[currentYear].level1s[level1].planned += numAmount;
      data[level0].years[currentYear].level1s[level1].months[monthNum].planned += numAmount;
      data[level0].years[currentYear].months[monthNum].level1s[level1].planned += numAmount;
      // Level 2 level
      data[level0].years[currentYear].level1s[level1].level2s[level2].planned += numAmount;
      data[level0].years[currentYear].level1s[level1].level2s[level2].months[monthNum].planned += numAmount;
      data[level0].years[currentYear].level1s[level1].months[monthNum].level2s[level2].planned += numAmount;
      data[level0].years[currentYear].months[monthNum].level1s[level1].level2s[level2].planned += numAmount;
    } else if (type === 'Actual') {
      // Level 0 level
      data[level0].actual += numAmount;
      // Year level
      data[level0].years[currentYear].actual += numAmount;
      // Month level
      data[level0].years[currentYear].months[monthNum].actual += numAmount;
      // Level 1 level
      data[level0].years[currentYear].level1s[level1].actual += numAmount;
      data[level0].years[currentYear].level1s[level1].months[monthNum].actual += numAmount;
      data[level0].years[currentYear].months[monthNum].level1s[level1].actual += numAmount;
      // Level 2 level
      data[level0].years[currentYear].level1s[level1].level2s[level2].actual += numAmount;
      data[level0].years[currentYear].level1s[level1].level2s[level2].months[monthNum].actual += numAmount;
      data[level0].years[currentYear].level1s[level1].months[monthNum].level2s[level2].actual += numAmount;
      data[level0].years[currentYear].months[monthNum].level1s[level1].level2s[level2].actual += numAmount;
    }
  }
  
  // Calculate variances (Actual - Planned)
  calculateVariances(data);
  
  return data;
}

function calculateVariances(data) {
  for (var level0 in data) {
    var level0Data = data[level0];
    level0Data.variance = level0Data.actual - level0Data.planned;
    
    // Year variances under Level 0
    for (var year in level0Data.years) {
      var yearData = level0Data.years[year];
      yearData.variance = yearData.actual - yearData.planned;
      
      // Month variances
      for (var monthNum in yearData.months) {
        var monthData = yearData.months[monthNum];
        monthData.variance = monthData.actual - monthData.planned;
        
        // Level 1 variances in month
        for (var level1 in monthData.level1s) {
          var level1Data = monthData.level1s[level1];
          level1Data.variance = level1Data.actual - level1Data.planned;
          
          // Level 2 variances in month/level1
          for (var level2 in level1Data.level2s) {
            var level2Data = level1Data.level2s[level2];
            level2Data.variance = level2Data.actual - level2Data.planned;
          }
        }
      }
      
      // Level 1 variances
      for (var level1 in yearData.level1s) {
        var level1Data = yearData.level1s[level1];
        level1Data.variance = level1Data.actual - level1Data.planned;
        
        // Level 1 month variances
        for (var monthNum in level1Data.months) {
          var monthData = level1Data.months[monthNum];
          monthData.variance = monthData.actual - monthData.planned;
          
          // Level 2 variances in level1/month
          for (var level2 in monthData.level2s) {
            var level2Data = monthData.level2s[level2];
            level2Data.variance = level2Data.actual - level2Data.planned;
          }
        }
        
        // Level 2 variances
        for (var level2 in level1Data.level2s) {
          var level2Data = level1Data.level2s[level2];
          level2Data.variance = level2Data.actual - level2Data.planned;
          
          // Level 2 month variances
          for (var monthNum in level2Data.months) {
            var monthData = level2Data.months[monthNum];
            monthData.variance = monthData.actual - monthData.planned;
          }
        }
      }
    }
  }
}

function createActualVsPlannedReportHTML(reportData) {
  var html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { 
          font-family: Arial, sans-serif; 
          padding: 20px; 
          margin: 0;
          background-color: #f5f5f5;
        }
        .container {
          background-color: white;
          border-radius: 8px;
          padding: 20px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        h2 { 
          color: #333; 
          margin-top: 0;
          text-align: center;
          border-bottom: 2px solid #4472C4;
          padding-bottom: 10px;
        }
        .tree-table {
          width: 100%;
          border-collapse: collapse;
          margin: 20px 0;
        }
        .tree-table th {
          background-color: #4472C4;
          color: white;
          padding: 12px;
          text-align: left;
          font-weight: bold;
          border: 1px solid #ddd;
          position: sticky;
          top: 0;
          z-index: 10;
        }
        .tree-table td {
          padding: 8px 12px;
          border: 1px solid #ddd;
          text-align: right;
        }
        .tree-table td:first-child {
          text-align: left;
        }
        .tree-row {
          background-color: #fff;
        }
        .tree-row:hover {
          background-color: #f0f8ff;
        }
        .level-0 {
          background-color: #e8f0fe;
          font-weight: bold;
          font-size: 16px;
        }
        .level-1 {
          background-color: #f0f4ff;
          font-weight: 600;
          padding-left: 20px;
        }
        .level-2 {
          background-color: #f8f9ff;
          padding-left: 40px;
        }
        .level-3 {
          background-color: #fcfcff;
          padding-left: 60px;
          font-size: 14px;
        }
        .level-4 {
          background-color: #fefeff;
          padding-left: 80px;
          font-size: 13px;
        }
        .expand-collapse {
          cursor: pointer;
          user-select: none;
          color: #4472C4;
          font-weight: bold;
          margin-right: 8px;
          font-size: 12px;
          width: 15px;
          display: inline-block;
        }
        .expand-collapse:hover {
          color: #2851a3;
        }
        .amount {
          font-family: 'Courier New', monospace;
          font-weight: 500;
        }
        .positive {
          color: #228B22;
        }
        .negative {
          color: #DC143C;
        }
        .variance-positive {
          background-color: #d4edda;
          color: #155724;
        }
        .variance-negative {
          background-color: #f8d7da;
          color: #721c24;
        }
        .controls {
          margin-bottom: 20px;
          text-align: center;
        }
        .btn {
          padding: 8px 16px;
          margin: 0 5px;
          border: none;
          border-radius: 4px;
          cursor: pointer;
          font-size: 14px;
        }
        .btn-primary {
          background-color: #4472C4;
          color: white;
        }
        .btn-secondary {
          background-color: #6c757d;
          color: white;
        }
        .btn:hover {
          opacity: 0.8;
        }
        .hidden {
          display: none;
        }
        .summary-stats {
          display: flex;
          justify-content: space-around;
          margin-bottom: 20px;
          background-color: #f8f9fa;
          padding: 15px;
          border-radius: 6px;
        }
        .stat-box {
          text-align: center;
          padding: 10px;
        }
        .stat-label {
          font-size: 12px;
          color: #666;
          text-transform: uppercase;
          margin-bottom: 5px;
        }
        .stat-value {
          font-size: 24px;
          font-weight: bold;
          font-family: 'Courier New', monospace;
        }
        .level0-breakdown {
          display: flex;
          justify-content: space-around;
          margin-bottom: 20px;
          gap: 15px;
        }
        .level0-stat-box {
          background-color: white;
          border-radius: 8px;
          padding: 15px;
          text-align: center;
          box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          border: 2px solid;
          flex: 1;
        }
        .level0-stat-box.income {
          border-color: #28a745;
          background-color: #f8fff9;
        }
        .level0-stat-box.expenditure {
          border-color: #dc3545;
          background-color: #fff8f8;
        }
        .level0-stat-box.neutral {
          border-color: #6c757d;
          background-color: #f8f9fa;
        }
        .level0-title {
          font-size: 16px;
          font-weight: bold;
          margin-bottom: 8px;
          text-transform: uppercase;
        }
        .level0-stat-box.income .level0-title {
          color: #28a745;
        }
        .level0-stat-box.expenditure .level0-title {
          color: #dc3545;
        }
        .level0-stat-box.neutral .level0-title {
          color: #6c757d;
        }
        .level0-details {
          display: flex;
          flex-direction: column;
          gap: 5px;
        }
        .level0-amount {
          font-size: 18px;
          font-weight: bold;
          font-family: 'Courier New', monospace;
        }
        .level0-variance {
          font-size: 14px;
          font-weight: 500;
          font-family: 'Courier New', monospace;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>Actual vs Planned Financial Report</h2>
        ${generateSummaryStats(reportData)}
        <div class="controls">
          <button class="btn btn-primary" onclick="expandAll()">Expand All</button>
          <button class="btn btn-secondary" onclick="collapseAll()">Collapse All</button>
        </div>
        <table class="tree-table">
          <thead>
            <tr>
              <th style="width: 40%;">Category</th>
              <th style="width: 15%;">Planned</th>
              <th style="width: 15%;">Actual</th>
              <th style="width: 15%;">Variance</th>
              <th style="width: 15%;">Variance %</th>
            </tr>
          </thead>
          <tbody>
            ${generateTreeRows(reportData)}
          </tbody>
        </table>
      </div>
      
      <script>
        function toggleExpand(element, targetClass) {
          const isExpanded = element.textContent === '▼';
          const rows = document.querySelectorAll('.' + targetClass);
          
          if (isExpanded) {
            element.textContent = '▶';
            rows.forEach(row => row.classList.add('hidden'));
          } else {
            element.textContent = '▼';
            rows.forEach(row => row.classList.remove('hidden'));
          }
        }
        
        function expandAll() {
          document.querySelectorAll('.expand-collapse').forEach(el => {
            if (el.textContent === '▶') {
              el.click();
            }
          });
        }
        
        function collapseAll() {
          document.querySelectorAll('.expand-collapse').forEach(el => {
            if (el.textContent === '▼') {
              el.click();
            }
          });
        }
        
        function formatCurrency(amount) {
          return '£' + Math.abs(amount).toLocaleString('en-GB', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
          });
        }
        
        function formatVariancePercent(variance, planned) {
          if (planned === 0) return 'N/A';
          const percent = (variance / Math.abs(planned)) * 100;
          return percent.toFixed(1) + '%';
        }
      </script>
    </body>
    </html>
  `;
  
  return html;
}

function generateSummaryStats(reportData) {
  var level0Categories = ['Income', 'Expenditure', 'Neutral'];
  var totalPlanned = 0;
  var totalActual = 0;
  var totalVariance = 0;
  
  // Calculate overall totals
  for (var i = 0; i < level0Categories.length; i++) {
    var level0 = level0Categories[i];
    var level0Data = reportData[level0];
    if (level0Data) {
      totalPlanned += level0Data.planned;
      totalActual += level0Data.actual;
      totalVariance += level0Data.variance;
    }
  }
  
  var totalVariancePercent = totalPlanned !== 0 ? ((totalVariance / Math.abs(totalPlanned)) * 100).toFixed(1) : 'N/A';
  
  var statsHtml = `
    <div class="summary-stats">
      <div class="stat-box">
        <div class="stat-label">Total Planned</div>
        <div class="stat-value amount">£${Math.abs(totalPlanned).toLocaleString('en-GB', {minimumFractionDigits: 2})}</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">Total Actual</div>
        <div class="stat-value amount">£${Math.abs(totalActual).toLocaleString('en-GB', {minimumFractionDigits: 2})}</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">Total Variance</div>
        <div class="stat-value amount ${totalVariance >= 0 ? 'positive' : 'negative'}">
          ${totalVariance >= 0 ? '+' : ''}£${Math.abs(totalVariance).toLocaleString('en-GB', {minimumFractionDigits: 2})}
        </div>
      </div>
      <div class="stat-box">
        <div class="stat-label">Variance %</div>
        <div class="stat-value ${totalVariance >= 0 ? 'positive' : 'negative'}">${totalVariancePercent}%</div>
      </div>
    </div>
    
    <div class="level0-breakdown">
  `;
  
  // Add Level 0 breakdown
  for (var i = 0; i < level0Categories.length; i++) {
    var level0 = level0Categories[i];
    var level0Data = reportData[level0];
    
    if (level0Data) {
      var variancePercent = level0Data.planned !== 0 ? ((level0Data.variance / Math.abs(level0Data.planned)) * 100).toFixed(1) : 'N/A';
      var level0Class = level0.toLowerCase();
      
      statsHtml += `
        <div class="level0-stat-box ${level0Class}">
          <div class="level0-title">${level0}</div>
          <div class="level0-details">
            <span class="level0-amount">£${Math.abs(level0Data.actual).toLocaleString('en-GB', {minimumFractionDigits: 2})}</span>
            <span class="level0-variance ${level0Data.variance >= 0 ? 'positive' : 'negative'}">
              ${level0Data.variance >= 0 ? '+' : ''}£${Math.abs(level0Data.variance).toLocaleString('en-GB', {minimumFractionDigits: 2})}
            </span>
          </div>
        </div>
      `;
    }
  }
  
  statsHtml += `</div>`;
  
  return statsHtml;
}

function generateTreeRows(reportData) {
  var html = '';
  var currentYear = new Date().getFullYear();
  
  // Level 0 sections (Income, Expenditure, Neutral)
  var level0Categories = ['Income', 'Expenditure', 'Neutral'];
  
  for (var l = 0; l < level0Categories.length; l++) {
    var level0 = level0Categories[l];
    var level0Data = reportData[level0];
    
    if (!level0Data) continue;
    
    // Level 0 row (top level section)
    html += generateRowHTML(level0, level0Data, 'level-0', 'level0-' + level0.toLowerCase() + '-years');
    
    // Year rows under Level 0 (hidden by default)
    var sortedYears = Object.keys(level0Data.years).sort();
    for (var y = 0; y < sortedYears.length; y++) {
      var year = sortedYears[y];
      var yearData = level0Data.years[year];
      html += generateRowHTML('  ' + year, yearData, 'level-1 level0-' + level0.toLowerCase() + '-years hidden', 'level0-' + level0.toLowerCase() + '-year-' + year + '-children');
      
      // Month rows under Year (hidden by default)
      var sortedMonths = Object.keys(yearData.months).sort((a, b) => parseInt(a) - parseInt(b));
      for (var i = 0; i < sortedMonths.length; i++) {
        var monthNum = sortedMonths[i];
        var monthData = yearData.months[monthNum];
        html += generateRowHTML('    ' + monthData.name, monthData, 'level-2 level0-' + level0.toLowerCase() + '-year-' + year + '-children hidden', 'level0-' + level0.toLowerCase() + '-year-' + year + '-month-' + monthNum + '-level1s');
        
        // Level 1s in month (hidden by default)
        var sortedLevel1s = Object.keys(monthData.level1s).sort();
        for (var j = 0; j < sortedLevel1s.length; j++) {
          var level1 = sortedLevel1s[j];
          var level1Data = monthData.level1s[level1];
          html += generateRowHTML('      ' + level1, level1Data, 'level-3 level0-' + level0.toLowerCase() + '-year-' + year + '-month-' + monthNum + '-level1s hidden', 'level0-' + level0.toLowerCase() + '-year-' + year + '-month-' + monthNum + '-level1-' + level1.replace(/\s+/g, '-') + '-level2s');
          
          // Level 2s in month/level1 (hidden by default)
          var sortedLevel2s = Object.keys(level1Data.level2s).sort();
          for (var k = 0; k < sortedLevel2s.length; k++) {
            var level2 = sortedLevel2s[k];
            var level2Data = level1Data.level2s[level2];
            html += generateRowHTML('        ' + level2, level2Data, 'level-4 level0-' + level0.toLowerCase() + '-year-' + year + '-month-' + monthNum + '-level1-' + level1.replace(/\s+/g, '-') + '-level2s hidden');
          }
        }
      }
      
      // Level 1 rows under Year (hidden by default)
      var sortedLevel1s = Object.keys(yearData.level1s).sort();
      for (var i = 0; i < sortedLevel1s.length; i++) {
        var level1 = sortedLevel1s[i];
        var level1Data = yearData.level1s[level1];
        html += generateRowHTML('    ' + level1, level1Data, 'level-2 level0-' + level0.toLowerCase() + '-year-' + year + '-children hidden', 'level0-' + level0.toLowerCase() + '-year-' + year + '-level1-' + level1.replace(/\s+/g, '-') + '-children');
        
        // Level 2 rows under Level 1 (hidden by default)
        var sortedLevel2s = Object.keys(level1Data.level2s).sort();
        for (var j = 0; j < sortedLevel2s.length; j++) {
          var level2 = sortedLevel2s[j];
          var level2Data = level1Data.level2s[level2];
          html += generateRowHTML('      ' + level2, level2Data, 'level-3 level0-' + level0.toLowerCase() + '-year-' + year + '-level1-' + level1.replace(/\s+/g, '-') + '-children hidden', 'level0-' + level0.toLowerCase() + '-year-' + year + '-level1-' + level1.replace(/\s+/g, '-') + '-level2-' + level2.replace(/\s+/g, '-') + '-months');
          
          // Level 2 month rows (hidden by default)
          var sortedMonths = Object.keys(level2Data.months).sort((a, b) => parseInt(a) - parseInt(b));
          for (var k = 0; k < sortedMonths.length; k++) {
            var monthNum = sortedMonths[k];
            var monthData = level2Data.months[monthNum];
            html += generateRowHTML('        ' + monthData.name, monthData, 'level-4 level0-' + level0.toLowerCase() + '-year-' + year + '-level1-' + level1.replace(/\s+/g, '-') + '-level2-' + level2.replace(/\s+/g, '-') + '-months hidden');
          }
        }
      }
    }
  }
  
  return html;
}

function generateRowHTML(label, data, cssClass, expandTarget) {
  var variancePercent = data.planned !== 0 ? ((data.variance / Math.abs(data.planned)) * 100).toFixed(1) + '%' : 'N/A';
  var varianceClass = data.variance >= 0 ? 'variance-positive' : 'variance-negative';
  
  var expandButton = '';
  if (expandTarget) {
    expandButton = `<span class="expand-collapse" onclick="toggleExpand(this, '${expandTarget}')">▶</span>`;
  }
  
  return `
    <tr class="tree-row ${cssClass}">
      <td>${expandButton}${label}</td>
      <td class="amount">£${Math.abs(data.planned).toLocaleString('en-GB', {minimumFractionDigits: 2})}</td>
      <td class="amount">£${Math.abs(data.actual).toLocaleString('en-GB', {minimumFractionDigits: 2})}</td>
      <td class="amount ${varianceClass}">${data.variance >= 0 ? '+' : ''}£${Math.abs(data.variance).toLocaleString('en-GB', {minimumFractionDigits: 2})}</td>
      <td class="amount ${varianceClass}">${variancePercent}</td>
    </tr>
  `;
}