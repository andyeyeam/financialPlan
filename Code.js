function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu('Custom Menu')
      .addItem('Add dumped Transactions', 'addDumpedTransactions')
      .addItem('Map Level 2', 'mapLevel2')
      .addToUi();
}

function addDumpedTransactions() {
  var dumpValues = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions Dump').getDataRange().getValues();
  var currValues  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions').getDataRange().getValues();
  var outValues   = new Array();

  for (var dvRow = 1; dvRow < dumpValues.length; dvRow++){
    var match = false;
    for (var cvRow = 1; cvRow < currValues.length; cvRow++){
      if (rowMatch (dumpValues[dvRow], currValues[cvRow])) {
        match = true;
        break;
      }
    }
    if (!match){
      outValues.push([dumpValues[dvRow][1],'',dumpValues[dvRow][5],'Barclays Debit Card','','','',dumpValues[dvRow][3],'','',dumpValues[dvRow][4]]);
    }
  }

  var sheet = SpreadsheetApp.getActiveSheet();
  // Get the last row that contains data.
  var startRow = sheet.getLastRow() + 1;

  // Get the number of rows to append.
  var numRows = outValues.length;

  // Get the number of columns to append.
  var numCols = outValues[0].length;

  // Get the range where you want to append the data.
  var range = sheet.getRange(startRow, 1, numRows, numCols);
  
  // Append the data to the sheet.
  range.setValues(outValues);
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
  if (dRow[1].getTime() !== cRow[0].getTime()) return false; // date
  if (Math.abs(dRow[3]) != Math.abs(cRow[6])) return false; // amount
  if (dRow[5].replace(/\s+/g, '') != cRow[2].replace(/\s+/g, '')) return false; // description
 
  return true;
}