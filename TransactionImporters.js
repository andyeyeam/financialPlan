function addDumpedTransactions() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
    var dumpSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Barclays Current Account');
    
    var dumpValues = dumpSheet.getDataRange().getValues();
    var currValues = sheet.getDataRange().getValues();
    var outValues = [];

    if (dumpValues.length <= 1) {
      SpreadsheetApp.getUi().alert('No Data', 'No transactions found in Barclays Current Account sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    var lastColumn = sheet.getLastColumn();

    var existingTransactions = new Set();
    for (var i = 1; i < currValues.length; i++) {
      if (currValues[i][0] && currValues[i][2] && currValues[i][8]) {
        var key = String(currValues[i][0]) + '|' + String(currValues[i][2]).replace(/\s+/g, '').toLowerCase() + '|' + Math.abs(currValues[i][8]);
        existingTransactions.add(key);
      }
    }

    var totalTransactions = dumpValues.length - 1;
    var processedCount = 0;
    var addedCount = 0;
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Processing 0 of ' + totalTransactions + ' transactions...', 'Processing Transactions', 3);
    
    for (var dvRow = 1; dvRow < dumpValues.length; dvRow++){
      processedCount++;
      
      if (processedCount % 10 === 0 || processedCount === totalTransactions || totalTransactions <= 20) {
        SpreadsheetApp.getActiveSpreadsheet().toast('Processing ' + processedCount + ' of ' + totalTransactions + ' transactions...', 'Processing Transactions', 3);
      }
      
      var dumpKey = String(dumpValues[dvRow][1]) + '|' + String(dumpValues[dvRow][5]).replace(/\s+/g, '').toLowerCase() + '|' + Math.abs(dumpValues[dvRow][3]);
      
      if (!existingTransactions.has(dumpKey)) {
        var newRowNumber = sheet.getLastRow() + 1 + outValues.length;
        
        var newRow = new Array(lastColumn);
        
        newRow[0] = dumpValues[dvRow][1];
        newRow[1] = '=TEXT(A' + newRowNumber + ',"MM")';
        newRow[2] = dumpValues[dvRow][5];
        newRow[3] = 'Barclays Current Account';
        newRow[4] = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$C$256,3,0),"")';
        newRow[5] = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$B$250,2,0),"")';
        newRow[6] = '';
        newRow[7] = '=IF(I' + newRowNumber + '>=0,"DR","CR")';
        newRow[8] = dumpValues[dvRow][3];
        newRow[9] = '=ABS(I' + newRowNumber + ')';
        newRow[10] = '';
        if (lastColumn > 11) newRow[11] = dumpValues[dvRow][4];
        
        outValues.push(newRow);
        addedCount++;
      }
    }

    if (outValues.length > 0) {
      var startRow = sheet.getLastRow() + 1;
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
    
    var sainsburysValues = sainsburysSheet.getDataRange().getValues();
    var currValues = sheet.getDataRange().getValues();
    var outValues = [];

    if (sainsburysValues.length <= 1) {
      SpreadsheetApp.getUi().alert('No Data', 'No transactions found in Sainsburys Bank Credit Card sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    var lastColumn = sheet.getLastColumn();

    var existingTransactions = new Set();
    for (var i = 1; i < currValues.length; i++) {
      if (currValues[i][0] && currValues[i][2] && currValues[i][8]) {
        var key = String(currValues[i][0]) + '|' + String(currValues[i][2]).replace(/\s+/g, '').toLowerCase() + '|' + Math.abs(currValues[i][8]);
        existingTransactions.add(key);
      }
    }

    var totalTransactions = sainsburysValues.length - 1;
    var processedCount = 0;
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Processing 0 of ' + totalTransactions + ' transactions...', 'Processing Transactions', 3);
    
    for (var svRow = 1; svRow < sainsburysValues.length; svRow++){
      processedCount++;
      
      if (processedCount % 10 === 0 || processedCount === totalTransactions || totalTransactions <= 20) {
        SpreadsheetApp.getActiveSpreadsheet().toast('Processing ' + processedCount + ' of ' + totalTransactions + ' transactions...', 'Processing Transactions', 3);
      }
      
      var transDate = sainsburysValues[svRow][0];
      var transDesc = sainsburysValues[svRow][1];
      var transAmount = sainsburysValues[svRow][2];
      var drCr = sainsburysValues[svRow][3];
      
      var numAmount = 0;
      if (typeof transAmount === 'string') {
        numAmount = parseFloat(transAmount.replace(/[£,]/g, ''));
      } else {
        numAmount = transAmount;
      }
      
      if (drCr && drCr.toString().trim().toUpperCase() === 'CR') {
        numAmount = Math.abs(numAmount);
      } else {
        numAmount = -Math.abs(numAmount);
      }
      
      var isDuplicate = false;
      
      for (var k = 1; k < currValues.length; k++) {
        var existingDate = currValues[k][0];
        var existingDesc = currValues[k][2];
        var existingAmount = currValues[k][8];
        
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
        
        if (transDateStr === existingDateStr &&
            String(transDesc).trim() === String(existingDesc).trim() &&
            numAmount === existingAmount) {
          isDuplicate = true;
          break;
        }
      }
      
      if (!isDuplicate) {
        var newRowNumber = sheet.getLastRow() + 1 + outValues.length;
        
        var newRow = new Array(lastColumn);
        
        newRow[0] = transDate;
        newRow[1] = '=TEXT(A' + newRowNumber + ',"MM")';
        newRow[2] = transDesc;
        newRow[3] = 'Sainsburys Bank Credit Card';
        newRow[4] = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$C$256,3,0),"")';
        newRow[5] = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$B$250,2,0),"")';
        newRow[6] = '';
        newRow[7] = '=IF(I' + newRowNumber + '>=0,"DR","CR")';
        newRow[8] = numAmount;
        newRow[9] = '=ABS(I' + newRowNumber + ')';
        newRow[10] = '';
        if (lastColumn > 11) newRow[11] = '';
        
        outValues.push(newRow);
      }
    }

    if (outValues.length > 0) {
      var startRow = sheet.getLastRow() + 1;
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
    
    var amexValues = amexSheet.getDataRange().getValues();
    var currValues = sheet.getDataRange().getValues();
    var outValues = [];

    if (amexValues.length <= 1) {
      SpreadsheetApp.getUi().alert('No Data', 'No transactions found in American Express Credit Card sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    var lastColumn = sheet.getLastColumn();

    var existingTransactions = new Set();
    for (var i = 1; i < currValues.length; i++) {
      if (currValues[i][0] && currValues[i][2] && currValues[i][8]) {
        var key = String(currValues[i][0]) + '|' + String(currValues[i][2]).replace(/\s+/g, '').toLowerCase() + '|' + Math.abs(currValues[i][8]);
        existingTransactions.add(key);
      }
    }

    var totalTransactions = amexValues.length - 1;
    var processedCount = 0;
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Processing 0 of ' + totalTransactions + ' transactions...', 'Processing Transactions', 3);
    
    for (var axRow = 1; axRow < amexValues.length; axRow++){
      processedCount++;
      
      if (processedCount % 10 === 0 || processedCount === totalTransactions || totalTransactions <= 20) {
        SpreadsheetApp.getActiveSpreadsheet().toast('Processing ' + processedCount + ' of ' + totalTransactions + ' transactions...', 'Processing Transactions', 3);
      }
      
      var transDate = amexValues[axRow][0];
      var transStatus = amexValues[axRow][1];
      var transDesc = amexValues[axRow][1];
      var transAmount = amexValues[axRow][3];
      
      var numAmount = 0;
      if (typeof transAmount === 'string') {
        numAmount = parseFloat(transAmount.replace(/[£,]/g, ''));
      } else {
        numAmount = transAmount;
      }
      
      if (transStatus && transStatus.toString().trim().toUpperCase() === 'CREDIT') {
        numAmount = Math.abs(numAmount);
      } else {
        if (numAmount > 0) {
          numAmount = -numAmount;
        }
      }
      
      var isDuplicate = false;
      
      for (var k = 1; k < currValues.length; k++) {
        var existingDate = currValues[k][0];
        var existingDesc = currValues[k][2];
        var existingAmount = currValues[k][8];
        
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
        
        if (transDateStr === existingDateStr &&
            String(transDesc).trim() === String(existingDesc).trim() &&
            numAmount === existingAmount) {
          isDuplicate = true;
          break;
        }
      }
      
      if (!isDuplicate) {
        var newRowNumber = sheet.getLastRow() + 1 + outValues.length;
        
        var newRow = new Array(lastColumn);
        
        newRow[0] = transDate;
        newRow[1] = '=TEXT(A' + newRowNumber + ',"MM")';
        newRow[2] = transDesc;
        newRow[3] = 'American Express Credit Card';
        newRow[4] = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$C$256,3,0),"")';
        newRow[5] = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$B$250,2,0),"")';
        newRow[6] = '';
        newRow[7] = '=IF(I' + newRowNumber + '>=0,"DR","CR")';
        newRow[8] = numAmount;
        newRow[9] = '=ABS(I' + newRowNumber + ')';
        newRow[10] = '';
        if (lastColumn > 11) newRow[11] = '';
        
        outValues.push(newRow);
      }
    }

    if (outValues.length > 0) {
      var startRow = sheet.getLastRow() + 1;
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