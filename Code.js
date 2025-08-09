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
      .addItem('Add', 'addLevel2Mapping');

  ui.createMenu('Custom Menu')
      .addSubMenu(updateSubmenu)
      .addSubMenu(mappingsSubmenu)
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
      if (currValues[i][0] && currValues[i][2] && currValues[i][7]) {
        var key = String(currValues[i][0]) + '|' + String(currValues[i][2]).replace(/\s+/g, '').toLowerCase() + '|' + Math.abs(currValues[i][7]);
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
        newRow[4] = '=IF(F' + newRowNumber + '="Income","Income","Expenditure")';
        
        // Column F (5): Level 1
        newRow[5] = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$B$250,2,0),"")';
        
        // Column G (6): Level 2
        newRow[6] = '';
        
        // Column H (7): Amount
        newRow[7] = dumpValues[dvRow][3];
        
        // Column I (8): Absolute Amount formula
        newRow[8] = '=ABS(H' + newRowNumber + ')';
        
        // Column J (9): Item
        newRow[9] = '';
        
        // Column K (10): Subcategory
        if (lastColumn > 10) newRow[10] = dumpValues[dvRow][4];
        
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
      if (currValues[i][0] && currValues[i][2] && currValues[i][7]) {
        var key = String(currValues[i][0]) + '|' + String(currValues[i][2]).replace(/\s+/g, '').toLowerCase() + '|' + Math.abs(currValues[i][7]);
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
      
      // Create lookup key for this transaction
      var sainsburysKey = String(transDate) + '|' + String(transDesc).replace(/\s+/g, '').toLowerCase() + '|' + Math.abs(numAmount);
      
      if (!existingTransactions.has(sainsburysKey)) {
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
        newRow[4] = '=IF(F' + newRowNumber + '="Income","Income","Expenditure")';
        
        // Column F (5): Level 1
        newRow[5] = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$B$250,2,0),"")';
        
        // Column G (6): Level 2
        newRow[6] = '';
        
        // Column H (7): Amount
        newRow[7] = numAmount;
        
        // Column I (8): Absolute Amount formula
        newRow[8] = '=ABS(H' + newRowNumber + ')';
        
        // Column J (9): Item
        newRow[9] = '';
        
        // Column K (10): Subcategory
        if (lastColumn > 10) newRow[10] = '';
        
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
      if (currValues[i][0] && currValues[i][2] && currValues[i][7]) {
        var key = String(currValues[i][0]) + '|' + String(currValues[i][2]).replace(/\s+/g, '').toLowerCase() + '|' + Math.abs(currValues[i][7]);
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
      
      // Create lookup key for this transaction (using STATUS field)
      var amexKey = String(transDate) + '|' + String(transDesc).replace(/\s+/g, '').toLowerCase() + '|' + Math.abs(numAmount);
      
      if (!existingTransactions.has(amexKey)) {
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
        newRow[4] = '=IF(F' + newRowNumber + '="Income","Income","Expenditure")';
        
        // Column F (5): Level 1
        newRow[5] = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$B$250,2,0),"")';
        
        // Column G (6): Level 2
        newRow[6] = '';
        
        // Column H (7): Amount
        newRow[7] = numAmount;
        
        // Column I (8): Absolute Amount formula
        newRow[8] = '=ABS(H' + newRowNumber + ')';
        
        // Column J (9): Item
        newRow[9] = '';
        
        // Column K (10): Subcategory
        if (lastColumn > 10) newRow[10] = '';
        
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
  // dRow[3] = dump amount, cRow[7] = transaction amount (column H is index 7)
  if (Math.abs(dRow[3]) != Math.abs(cRow[7])) return false;
  
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
        formula = '=IF(F' + newRowNumber + '="Income","Income","Expenditure")';
        break;
      case 6: // Column F - Level 1
        formula = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$B$250,2,0),"")';
        break;
      case 9: // Column I - Absolute Amount
        formula = '=ABS(H' + newRowNumber + ')';
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
      var absoluteAmount = transValues[i][8]; // Column I - Absolute Amount
      
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