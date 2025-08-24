function rowMatch (dRow, cRow) {
  var dumpDate = dRow[1];
  var transDate = cRow[0];
  
  if (typeof dumpDate === 'string') {
    dumpDate = new Date(dumpDate);
  }
  
  if (typeof transDate === 'string') {
    transDate = new Date(transDate);
  }
  
  if (dumpDate.toDateString() !== transDate.toDateString()) return false;
  
  if (Math.abs(dRow[3]) != Math.abs(cRow[8])) return false;
  
  var dumpDesc = String(dRow[5]).replace(/\s+/g, '').toLowerCase().trim();
  var transDesc = String(cRow[2]).replace(/\s+/g, '').toLowerCase().trim();
  if (dumpDesc !== transDesc) return false;
 
  return true;
}

function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  
  if (sheet.getName() !== 'Transactions') return;
  
  var editedRow = range.getRow();
  var lastRowWithData = sheet.getLastRow();
  
  if (editedRow > lastRowWithData - 1 || isNewRowAdded(sheet, editedRow)) {
    copyFormulasToNewRow(sheet, editedRow);
  }
}

function isNewRowAdded(sheet, editedRow) {
  var rowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var nonEmptyCount = 0;
  
  for (var i = 0; i < rowData.length; i++) {
    if (rowData[i] !== '') {
      nonEmptyCount++;
    }
  }
  
  return nonEmptyCount <= 3;
}

function copyFormulasToNewRow(sheet, newRowNumber) {
  if (newRowNumber <= 1) return;
  
  var lastColumn = sheet.getLastColumn();
  
  for (var col = 1; col <= lastColumn; col++) {
    var targetCell = sheet.getRange(newRowNumber, col);
    var formula = '';
    
    switch (col) {
      case 2:
        formula = '=TEXT(A' + newRowNumber + ',"MM")';
        break;
      case 5:
        formula = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$C$256,3,0),"")';
        break;
      case 6:
        formula = '=IFNA(VLOOKUP(G' + newRowNumber + ',Taxonomy!$A$2:$B$250,2,0),"")';
        break;
      case 8:
        formula = '=IF(I' + newRowNumber + '>=0,"DR","CR")';
        break;
      case 10:
        formula = '=ABS(I' + newRowNumber + ')';
        break;
    }
    
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
  
  var updatedFormula = formula.replace(/([A-Z]+)(\$?)(\$?)(\d+)\b/g, function(match, columnLetters, columnDollar, rowDollar, rowNum) {
    if (rowDollar === '$') {
      return match;
    } else {
      return columnLetters + columnDollar + newRowNumber;
    }
  });
  
  return updatedFormula;
}