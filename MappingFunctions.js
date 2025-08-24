function mapLevel2 (){
  var mapValues  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Level 2 Mapping').getDataRange().getValues();
  var currValues  = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions').getDataRange().getValues();
  var updatedCount = 0;
  
  for (var i = 1; i < currValues.length; i++){
    if (currValues[i][6] && currValues[i][6].toString().trim() !== "") continue;
    
    for (var j = 1; j < mapValues.length; j++){
      if (!currValues[i][2].includes (mapValues[j][0])) continue;
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions').getRange(i + 1, 7).setValue(mapValues[j][1]);
      updatedCount++;
      break;
    }
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  var updatedValues = sheet.getDataRange().getValues();
  var remainingBlankCount = 0;
  
  for (var i = 1; i < updatedValues.length; i++){
    if (!updatedValues[i][6] || updatedValues[i][6].toString().trim() === "") {
      remainingBlankCount++;
    }
  }
  
  var message = 'Updated ' + updatedCount + ' Level 2 values in the Transactions sheet.\n';
  message += remainingBlankCount + ' blank Level 2 values remain.';
  SpreadsheetApp.getUi().alert('Update Levels Complete', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function autoGenerateLevel2Mappings() {
  var transSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  var mappingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Level 2 Mapping');
  
  var transValues = transSheet.getDataRange().getValues();
  
  var existingMappings = mappingSheet.getDataRange().getValues();
  var existingPatterns = new Set();
  
  for (var i = 1; i < existingMappings.length; i++) {
    if (existingMappings[i][0]) {
      existingPatterns.add(existingMappings[i][0].toLowerCase());
    }
  }
  
  var patternAnalysis = {};
  
  for (var i = 1; i < transValues.length; i++) {
    var description = transValues[i][2];
    var level2 = transValues[i][6];
    
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
  
  var newMappings = [];
  
  for (var pattern in patternAnalysis) {
    var level2Counts = patternAnalysis[pattern];
    var level2Options = Object.keys(level2Counts);
    
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
    
    if (totalCount >= 2 && (maxCount / totalCount) >= 0.8 && !existingPatterns.has(pattern.toLowerCase())) {
      newMappings.push([pattern, dominantLevel2]);
    }
  }
  
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
  
  var merchantMatch = cleanDesc.match(/^([A-Z][A-Z\s&]{2,20}?)[\s]{2,}/);
  if (merchantMatch) {
    var merchantName = merchantMatch[1].trim();
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
  
  if (patterns.length === 0) {
    var words = cleanDesc.split(/\s+/);
    for (var i = 0; i < words.length && i < 2; i++) {
      var word = words[i].replace(/[^A-Z]/g, '');
      if (word.length >= 4 && word.length <= 12) {
        patterns.push(word);
        break;
      }
    }
  }
  
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
    
    var transValues = transSheet.getDataRange().getValues();
    var unmappedTransactions = [];
    
    var seenDescriptions = new Set();
    
    for (var i = 1; i < transValues.length; i++) {
      var description = transValues[i][2];
      var level2 = transValues[i][6];
      var absoluteAmount = transValues[i][9];
      
      if (description && description.toString().trim() !== '') {
        var descStr = description.toString().trim();
        
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
    
    var taxonomySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Taxonomy');
    if (!taxonomySheet) {
      SpreadsheetApp.getUi().alert('Taxonomy Sheet Missing', 'Could not find the Taxonomy sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var taxonomyValues = taxonomySheet.getRange('A:A').getValues();
    var level2Options = [];
    
    for (var i = 1; i < taxonomyValues.length; i++) {
      if (taxonomyValues[i][0] && taxonomyValues[i][0].toString().trim() !== '') {
        level2Options.push(taxonomyValues[i][0].toString().trim());
      }
    }
    
    if (level2Options.length === 0) {
      SpreadsheetApp.getUi().alert('No Level 2 Options', 'No Level 2 categories found in the Taxonomy sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    level2Options.sort();
    
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
  
  var tableRowsHtml = '';
  for (var i = 0; i < unmappedTransactions.length; i++) {
    var transaction = unmappedTransactions[i];
    var formattedAmount = '';
    
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
        let sortOrder = {};
        
        function sortTable(columnIndex) {
          const table = document.getElementById('mappingTable');
          const tbody = table.getElementsByTagName('tbody')[0];
          const rows = Array.from(tbody.getElementsByTagName('tr'));
          const headers = table.getElementsByTagName('th');
          
          if (!sortOrder[columnIndex]) {
            sortOrder[columnIndex] = 'asc';
          } else if (sortOrder[columnIndex] === 'asc') {
            sortOrder[columnIndex] = 'desc';
          } else {
            sortOrder[columnIndex] = 'asc';
          }
          
          for (let i = 0; i < headers.length; i++) {
            headers[i].classList.remove('sort-asc', 'sort-desc');
          }
          
          headers[columnIndex].classList.add('sort-' + sortOrder[columnIndex]);
          
          rows.sort((a, b) => {
            let aValue, bValue;
            
            if (columnIndex === 0) {
              aValue = a.cells[columnIndex].getElementsByTagName('input')[0].value.toLowerCase();
              bValue = b.cells[columnIndex].getElementsByTagName('input')[0].value.toLowerCase();
            } else if (columnIndex === 1) {
              aValue = parseFloat(a.cells[columnIndex].textContent.replace(/[£,]/g, '')) || 0;
              bValue = parseFloat(b.cells[columnIndex].textContent.replace(/[£,]/g, '')) || 0;
            } else if (columnIndex === 2) {
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
    
    var existingMappings = mappingSheet.getDataRange().getValues();
    for (var i = 1; i < existingMappings.length; i++) {
      if (existingMappings[i][0] && existingMappings[i][0].toString() === description) {
        throw new Error('A mapping for "' + description + '" already exists.');
      }
    }
    
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
    
    var existingMappings = mappingSheet.getDataRange().getValues();
    var existingDescriptions = new Set();
    
    for (var i = 1; i < existingMappings.length; i++) {
      if (existingMappings[i][0]) {
        existingDescriptions.add(existingMappings[i][0].toString());
      }
    }
    
    var newMappings = [];
    var duplicates = [];
    
    for (var i = 0; i < mappings.length; i++) {
      var description = mappings[i].description.trim();
      var level2 = mappings[i].level2.trim();
      
      if (existingDescriptions.has(description)) {
        duplicates.push(description);
      } else {
        newMappings.push([description, level2]);
        existingDescriptions.add(description);
      }
    }
    
    if (newMappings.length > 0) {
      var startRow = mappingSheet.getLastRow() + 1;
      mappingSheet.getRange(startRow, 1, newMappings.length, 2).setValues(newMappings);
    }
    
    return newMappings.length;
    
  } catch (error) {
    throw error;
  }
}

function addMappingFromCurrent() {
  try {
    var activeSheet = SpreadsheetApp.getActiveSheet();
    if (activeSheet.getName() !== 'Transactions') {
      SpreadsheetApp.getUi().alert('Wrong Sheet', 'Please select the Transactions sheet before using this function.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var activeRange = activeSheet.getActiveRange();
    var currentRow = activeRange.getRow();
    
    if (currentRow <= 1) {
      SpreadsheetApp.getUi().alert('Invalid Row', 'Please select a transaction row (not the header row).', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var description = activeSheet.getRange(currentRow, 3).getValue();
    var level2 = activeSheet.getRange(currentRow, 7).getValue();
    
    if (!description || description.toString().trim() === '') {
      SpreadsheetApp.getUi().alert('Missing Description', 'The selected row does not have a description in Column C.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var descriptionStr = description.toString().trim();
    var level2Str = '';
    
    if (!level2 || level2.toString().trim() === '') {
      var taxonomySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Taxonomy');
      if (!taxonomySheet) {
        SpreadsheetApp.getUi().alert('Taxonomy Sheet Missing', 'Could not find the Taxonomy sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
      
      var taxonomyValues = taxonomySheet.getRange('A:A').getValues();
      var level2Options = [];
      
      for (var i = 1; i < taxonomyValues.length; i++) {
        if (taxonomyValues[i][0] && taxonomyValues[i][0].toString().trim() !== '') {
          level2Options.push(taxonomyValues[i][0].toString().trim());
        }
      }
      
      if (level2Options.length === 0) {
        SpreadsheetApp.getUi().alert('No Level 2 Options', 'No Level 2 categories found in the Taxonomy sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
      
      level2Options.sort();
      
      var html = createLevel2SelectDialog(descriptionStr, level2Options);
      var htmlOutput = HtmlService.createHtmlOutput(html)
          .setWidth(500)
          .setHeight(400);
      
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select Level 2 Category');
      return;
      
    } else {
      level2Str = level2.toString().trim();
    }
    
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
    var mappingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Level 2 Mapping');
    if (!mappingSheet) {
      throw new Error('Could not find the Level 2 Mapping sheet.');
    }
    
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
    
    var lastRow = mappingSheet.getLastRow() + 1;
    mappingSheet.getRange(lastRow, 1, 1, 2).setValues([[descriptionStr, level2Str]]);
    
    return 'Successfully added mapping: "' + descriptionStr + '" → "' + level2Str + '".';
    
  } catch (error) {
    throw error;
  }
}