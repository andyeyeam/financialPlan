function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu('Custom Menu')
      .addItem('Add dumped Transactions', 'addDumpedTransactions')
      .addItem('Map Level 2', 'mapLevel2')
      .addItem('Auto-Generate Level 2 Mappings', 'autoGenerateLevel2Mappings')
      .addToUi();
}

function addDumpedTransactions() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
    var dumpSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions Dump');
    
    // Get data more efficiently
    var dumpValues = dumpSheet.getDataRange().getValues();
    var currValues = sheet.getDataRange().getValues();
    var outValues = [];

    // Early exit if no dump data
    if (dumpValues.length <= 1) {
      SpreadsheetApp.getUi().alert('No Data', 'No transactions found in Transactions Dump sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    // Read formulas from row 2 (first data row) to use as template
    var templateRow = 2;
    var lastColumn = sheet.getLastColumn();
    var formulas = [];
    
    // Get all formulas in one batch
    var formulaRange = sheet.getRange(templateRow, 1, 1, lastColumn);
    var formulaValues = formulaRange.getFormulas()[0];
    formulas = formulaValues;

    // Create lookup set for faster duplicate checking
    var existingTransactions = new Set();
    for (var i = 1; i < currValues.length; i++) {
      if (currValues[i][0] && currValues[i][2] && currValues[i][7]) {
        var key = String(currValues[i][0]) + '|' + String(currValues[i][2]).replace(/\s+/g, '').toLowerCase() + '|' + Math.abs(currValues[i][7]);
        existingTransactions.add(key);
      }
    }

    // Process dump transactions with progress indicator
    var totalTransactions = dumpValues.length - 1; // Exclude header row
    var processedCount = 0;
    var addedCount = 0;
    
    // Show initial progress
    SpreadsheetApp.getUi().alert('Processing', 'Processing 0 of ' + totalTransactions + ' dumped transactions...', SpreadsheetApp.getUi().ButtonSet.OK);
    
    for (var dvRow = 1; dvRow < dumpValues.length; dvRow++){
      processedCount++;
      
      // Update progress every 10 transactions or for small batches
      if (processedCount % 10 === 0 || processedCount === totalTransactions || totalTransactions <= 20) {
        SpreadsheetApp.getUi().alert('Processing', 'Processing ' + processedCount + ' of ' + totalTransactions + ' dumped transactions...', SpreadsheetApp.getUi().ButtonSet.OK);
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
        
        // Column B (1): Month formula with updated row reference
        newRow[1] = formulas[1] ? updateFormulaRowReferences(formulas[1], newRowNumber) : '';
        
        // Column C (2): Description
        newRow[2] = dumpValues[dvRow][5];
        
        // Column D (3): Account
        newRow[3] = 'Barclays Debit Card';
        
        // Column E (4): Level 0
        newRow[4] = formulas[4] ? updateFormulaRowReferences(formulas[4], newRowNumber) : '';
        
        // Column F (5): Level 1  
        newRow[5] = formulas[5] ? updateFormulaRowReferences(formulas[5], newRowNumber) : '';
        
        // Column G (6): Level 2
        newRow[6] = formulas[6] ? updateFormulaRowReferences(formulas[6], newRowNumber) : '';
        
        // Column H (7): Amount
        newRow[7] = dumpValues[dvRow][3];
        
        // Column I (8): Absolute Amount formula with updated row reference
        newRow[8] = formulas[8] ? updateFormulaRowReferences(formulas[8], newRowNumber) : '';
        
        // Column J (9): Item
        newRow[9] = formulas[9] ? updateFormulaRowReferences(formulas[9], newRowNumber) : '';
        
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
      
      SpreadsheetApp.getUi().alert('Success', 'Completed processing ' + totalTransactions + ' transactions.\nAdded ' + outValues.length + ' new transactions to the sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('No New Transactions', 'Completed processing ' + totalTransactions + ' transactions.\nAll transactions in the dump already exist in the Transactions sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    Logger.log('Error in addDumpedTransactions: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', 'An error occurred while processing transactions: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function mapLevel2 (){

  // Load the map into memory
  var mapValues  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Level 2 Mapping').getDataRange().getValues();

  // Load the transaction into memory and loop round them
  var currValues  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions').getDataRange().getValues();
  for (var i = 1; i < currValues.length; i++){      // Start at 1 to skip the transactions header row (row 0)
    // if (currValues[i][5] != "") continue;           // Ignore if already has a Level 2 value
    for (var j = 1; j < mapValues.length; j++){     // Start at 1 to skip the map header row (row 0)
      if (!currValues[i][2].includes (mapValues[j][0])) continue;
      Logger.log('mapValues ' + mapValues[j][1] + ' Transactions row ' + i);
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions').getRange(i + 1, 7).setValue(mapValues[j][1]);
    }
  }
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
  
  // Find the best source row with formulas (search upward from new row)
  var sourceRow = null;
  
  // First try the row immediately above
  for (var i = newRowNumber - 1; i >= 2; i--) {
    if (hasFormulas(sheet, i)) {
      sourceRow = i;
      break;
    }
  }
  
  // If no formulas found above, search the entire sheet
  if (!sourceRow) {
    for (var i = 2; i <= sheet.getLastRow(); i++) {
      if (i !== newRowNumber && hasFormulas(sheet, i)) {
        sourceRow = i;
        break;
      }
    }
  }
  
  if (!sourceRow) return; // No source row with formulas found
  
  // Copy formulas from source row to new row
  for (var col = 1; col <= lastColumn; col++) {
    var sourceCell = sheet.getRange(sourceRow, col);
    var targetCell = sheet.getRange(newRowNumber, col);
    
    // Copy formula if source has one (don't check if target is empty - just copy)
    if (sourceCell.getFormula()) {
      targetCell.setFormula(sourceCell.getFormula());
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
  
  // Replace row references in the formula, but preserve absolute row references
  // We need to handle two cases:
  // 1. Relative row references (A2, $A2) - update row number
  // 2. Absolute row references ($A$2) - keep original row number (2)
  
  var updatedFormula = formula.replace(/([A-Z]+)(\$?)(\$?)2\b/g, function(match, columnLetters, columnDollar, rowDollar) {
    if (rowDollar === '$') {
      // Absolute row reference - keep original row 2
      return columnLetters + columnDollar + rowDollar + '2';
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
    
    Logger.log('Added ' + newMappings.length + ' new Level 2 mappings');
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