function createABVSheet() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    var abvSheet = spreadsheet.getSheetByName('ABV');
    if (abvSheet) {
      abvSheet.clear();
    } else {
      abvSheet = spreadsheet.insertSheet('ABV');
    }
    
    var monthlyExpandedSheet = spreadsheet.getSheetByName('Monthly Expanded');
    if (!monthlyExpandedSheet) {
      SpreadsheetApp.getUi().alert('Error', 'Monthly Expanded sheet not found. Please ensure the Monthly Expanded sheet exists with the required data.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var expandedValues = monthlyExpandedSheet.getDataRange().getValues();
    if (expandedValues.length <= 1) {
      SpreadsheetApp.getUi().alert('Error', 'Monthly Expanded sheet is empty. Please ensure it contains data.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    var uniquePairs = new Set();
    var categoryPairs = [];
    
    for (var i = 1; i < expandedValues.length; i++) {
      var level0 = expandedValues[i][0] ? expandedValues[i][0].toString().trim() : '';
      var level1 = expandedValues[i][1] ? expandedValues[i][1].toString().trim() : '';
      var level2 = expandedValues[i][2] ? expandedValues[i][2].toString().trim() : '';
      
      if (level2 && level1) {
        var pairKey = level1 + '|' + level2;
        
        if (!uniquePairs.has(pairKey)) {
          uniquePairs.add(pairKey);
          
          categoryPairs.push({
            level0: level0,
            level1: level1,
            level2: level2
          });
        }
      }
    }
    
    if (categoryPairs.length === 0) {
      SpreadsheetApp.getUi().alert('No Data', 'No Level 1 and Level 2 pairs found in the Monthly Expanded sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    var headers = ['Level 0', 'Level 1', 'Level 2', 'Month', 'Type', 'Amount'];
    abvSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    var headerRange = abvSheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#4472C4')
               .setFontColor('#FFFFFF')
               .setFontWeight('bold')
               .setHorizontalAlignment('center');
    
    var dataRows = [];
    var monthNumbers = ['01', '02', '03', '04', '05', '06', 
                        '07', '08', '09', '10', '11', '12'];
    
    for (var p = 0; p < categoryPairs.length; p++) {
      var pair = categoryPairs[p];
      
      for (var month = 1; month <= 12; month++) {
        var monthNumber = monthNumbers[month - 1];
        
        var plannedAmount = 0;
        for (var i = 1; i < expandedValues.length; i++) {
          var rowLevel0 = expandedValues[i][0] ? expandedValues[i][0].toString().trim() : '';
          var rowLevel1 = expandedValues[i][1] ? expandedValues[i][1].toString().trim() : '';
          var rowLevel2 = expandedValues[i][2] ? expandedValues[i][2].toString().trim() : '';
          var rowMonth = expandedValues[i][3];
          var rowAmount = expandedValues[i][4];
          
          if (rowLevel1 === pair.level1 && 
              rowLevel2 === pair.level2 && 
              rowMonth === month) {
            plannedAmount = (rowAmount && typeof rowAmount === 'number') ? rowAmount : 0;
            pair.level0 = rowLevel0;
            break;
          }
        }
        
        var plannedRowNumber = 2 + dataRows.length;
        dataRows.push([
          pair.level0 || '',
          pair.level1,
          pair.level2,
          monthNumber,
          'Planned',
          plannedAmount
        ]);
        
        var actualRowNumber = 2 + dataRows.length;
        var sumFormula = `=SUMIFS(Transactions!J:J,Transactions!B:B,"${monthNumber}",Transactions!F:F,B${actualRowNumber},Transactions!G:G,C${actualRowNumber})`;
        
        dataRows.push([
          pair.level0 || '',
          pair.level1,
          pair.level2,
          monthNumber,
          'Actual',
          sumFormula
        ]);
      }
    }
    
    if (dataRows.length > 0) {
      var dataRange = abvSheet.getRange(2, 1, dataRows.length, headers.length);
      dataRange.setValues(dataRows);
      
      dataRange.setBorder(true, true, true, true, true, true);
      
      var amountColumn = abvSheet.getRange(2, 6, dataRows.length, 1);
      amountColumn.setNumberFormat('£#,##0.00');
      
      for (var row = 2; row <= dataRows.length + 1; row += 2) {
        abvSheet.getRange(row, 1, 1, headers.length).setBackground('#F8F9FA');
      }
    }
    
    abvSheet.autoResizeColumns(1, headers.length);
    abvSheet.setFrozenRows(1);
    abvSheet.setColumnWidth(1, 120);
    abvSheet.setColumnWidth(2, 150);
    abvSheet.setColumnWidth(3, 150);
    abvSheet.setColumnWidth(4, 100);
    abvSheet.setColumnWidth(5, 80);
    abvSheet.setColumnWidth(6, 100);
    
    var totalRows = dataRows.length;
    var totalPairs = categoryPairs.length;
    
    SpreadsheetApp.getUi().alert(
      'Success', 
      'ABV sheet created successfully!\n\n' +
      'Generated ' + totalRows + ' rows for ' + totalPairs + ' category pairs.\n' +
      'Each pair has 24 rows (12 months × 2 types: Planned and Actual).\n\n' +
      'Planned amounts are now retrieved directly from the Monthly Expanded sheet by matching Level 0, Level 1, Level 2, and Month.', 
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
    
    var abvSheet = spreadsheet.getSheetByName('ABV');
    var shouldRecreateABV = false;
    
    if (!abvSheet) {
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
        return;
      }
    } else {
      var recreateResponse = ui.alert(
        'Recreate ABV Data?', 
        'Would you like to recreate the ABV data with the latest information before generating the report?\n\nChoose:\n• YES - Refresh ABV data then create report\n• NO - Use existing ABV data for report', 
        ui.ButtonSet.YES_NO
      );
      
      shouldRecreateABV = (recreateResponse === ui.Button.YES);
    }
    
    if (shouldRecreateABV) {
      ui.alert('Refreshing Data', 'Recreating ABV sheet with latest data...', ui.ButtonSet.OK);
      createABVSheet();
      
      abvSheet = spreadsheet.getSheetByName('ABV');
      if (!abvSheet) {
        ui.alert('Error', 'Failed to recreate ABV sheet.', ui.ButtonSet.OK);
        return;
      }
    }
    
    var abvData = abvSheet.getDataRange().getValues();
    if (abvData.length <= 1) {
      ui.alert('Error', 'No data found in ABV sheet. Please check your Taxonomy sheet and try recreating the ABV data.', ui.ButtonSet.OK);
      return;
    }
    
    var reportData = buildReportDataStructure(abvData);
    
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
  
  var currentYear = new Date().getFullYear();
  
  var level0Categories = ['Income', 'Expenditure', 'Neutral'];
  
  for (var i = 1; i < abvData.length; i++) {
    var level0 = abvData[i][0];
    var level1 = abvData[i][1];
    var level2 = abvData[i][2];
    var monthValue = abvData[i][3];
    var type = abvData[i][4];
    var amount = abvData[i][5];
    
    if (!level1 || !level2 || !monthValue || !type) continue;
    
    if (!level0 || level0.toString().trim() === '') {
      level0 = 'Expenditure';
    }
    
    var numAmount = 0;
    if (typeof amount === 'number') {
      numAmount = amount;
    } else if (typeof amount === 'string' && amount.trim() !== '') {
      numAmount = parseFloat(amount.replace(/[£,]/g, '')) || 0;
    }
    
    var monthNum = parseInt(monthValue);
    if (!monthNum || monthNum < 1 || monthNum > 12) continue;
    
    var monthName = monthNames[monthNum - 1];
    
    if (!data[level0]) {
      data[level0] = {
        planned: 0,
        actual: 0,
        variance: 0,
        years: {}
      };
    }
    
    if (!data[level0].years[currentYear]) {
      data[level0].years[currentYear] = {
        planned: 0,
        actual: 0,
        variance: 0,
        months: {},
        level1s: {}
      };
    }
    
    if (!data[level0].years[currentYear].months[monthNum]) {
      data[level0].years[currentYear].months[monthNum] = {
        name: monthName,
        planned: 0,
        actual: 0,
        variance: 0,
        level1s: {}
      };
    }
    
    if (!data[level0].years[currentYear].level1s[level1]) {
      data[level0].years[currentYear].level1s[level1] = {
        planned: 0,
        actual: 0,
        variance: 0,
        months: {},
        level2s: {}
      };
    }
    
    if (!data[level0].years[currentYear].months[monthNum].level1s[level1]) {
      data[level0].years[currentYear].months[monthNum].level1s[level1] = {
        planned: 0,
        actual: 0,
        variance: 0,
        level2s: {}
      };
    }
    
    if (!data[level0].years[currentYear].level1s[level1].months[monthNum]) {
      data[level0].years[currentYear].level1s[level1].months[monthNum] = {
        name: monthName,
        planned: 0,
        actual: 0,
        variance: 0,
        level2s: {}
      };
    }
    
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
    
    if (type === 'Planned') {
      data[level0].planned += numAmount;
      data[level0].years[currentYear].planned += numAmount;
      data[level0].years[currentYear].months[monthNum].planned += numAmount;
      data[level0].years[currentYear].level1s[level1].planned += numAmount;
      data[level0].years[currentYear].level1s[level1].months[monthNum].planned += numAmount;
      data[level0].years[currentYear].months[monthNum].level1s[level1].planned += numAmount;
      data[level0].years[currentYear].level1s[level1].level2s[level2].planned += numAmount;
      data[level0].years[currentYear].level1s[level1].level2s[level2].months[monthNum].planned += numAmount;
      data[level0].years[currentYear].level1s[level1].months[monthNum].level2s[level2].planned += numAmount;
      data[level0].years[currentYear].months[monthNum].level1s[level1].level2s[level2].planned += numAmount;
    } else if (type === 'Actual') {
      data[level0].actual += numAmount;
      data[level0].years[currentYear].actual += numAmount;
      data[level0].years[currentYear].months[monthNum].actual += numAmount;
      data[level0].years[currentYear].level1s[level1].actual += numAmount;
      data[level0].years[currentYear].level1s[level1].months[monthNum].actual += numAmount;
      data[level0].years[currentYear].months[monthNum].level1s[level1].actual += numAmount;
      data[level0].years[currentYear].level1s[level1].level2s[level2].actual += numAmount;
      data[level0].years[currentYear].level1s[level1].level2s[level2].months[monthNum].actual += numAmount;
      data[level0].years[currentYear].level1s[level1].months[monthNum].level2s[level2].actual += numAmount;
      data[level0].years[currentYear].months[monthNum].level1s[level1].level2s[level2].actual += numAmount;
    }
  }
  
  calculateVariances(data);
  
  return data;
}

function calculateVariances(data) {
  for (var level0 in data) {
    var level0Data = data[level0];
    level0Data.variance = level0Data.actual - level0Data.planned;
    
    for (var year in level0Data.years) {
      var yearData = level0Data.years[year];
      yearData.variance = yearData.actual - yearData.planned;
      
      for (var monthNum in yearData.months) {
        var monthData = yearData.months[monthNum];
        monthData.variance = monthData.actual - monthData.planned;
        
        for (var level1 in monthData.level1s) {
          var level1Data = monthData.level1s[level1];
          level1Data.variance = level1Data.actual - level1Data.planned;
          
          for (var level2 in level1Data.level2s) {
            var level2Data = level1Data.level2s[level2];
            level2Data.variance = level2Data.actual - level2Data.planned;
          }
        }
      }
      
      for (var level1 in yearData.level1s) {
        var level1Data = yearData.level1s[level1];
        level1Data.variance = level1Data.actual - level1Data.planned;
        
        for (var monthNum in level1Data.months) {
          var monthData = level1Data.months[monthNum];
          monthData.variance = monthData.actual - monthData.planned;
          
          for (var level2 in monthData.level2s) {
            var level2Data = monthData.level2s[level2];
            level2Data.variance = level2Data.actual - level2Data.planned;
          }
        }
        
        for (var level2 in level1Data.level2s) {
          var level2Data = level1Data.level2s[level2];
          level2Data.variance = level2Data.actual - level2Data.planned;
          
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
  
  var level0Categories = ['Income', 'Expenditure', 'Neutral'];
  
  for (var l = 0; l < level0Categories.length; l++) {
    var level0 = level0Categories[l];
    var level0Data = reportData[level0];
    
    if (!level0Data) continue;
    
    html += generateRowHTML(level0, level0Data, 'level-0', 'level0-' + level0.toLowerCase() + '-years');
    
    var sortedYears = Object.keys(level0Data.years).sort();
    for (var y = 0; y < sortedYears.length; y++) {
      var year = sortedYears[y];
      var yearData = level0Data.years[year];
      html += generateRowHTML('  ' + year, yearData, 'level-1 level0-' + level0.toLowerCase() + '-years hidden', 'level0-' + level0.toLowerCase() + '-year-' + year + '-children');
      
      var sortedMonths = Object.keys(yearData.months).sort((a, b) => parseInt(a) - parseInt(b));
      for (var i = 0; i < sortedMonths.length; i++) {
        var monthNum = sortedMonths[i];
        var monthData = yearData.months[monthNum];
        html += generateRowHTML('    ' + monthData.name, monthData, 'level-2 level0-' + level0.toLowerCase() + '-year-' + year + '-children hidden', 'level0-' + level0.toLowerCase() + '-year-' + year + '-month-' + monthNum + '-level1s');
        
        var sortedLevel1s = Object.keys(monthData.level1s).sort();
        for (var j = 0; j < sortedLevel1s.length; j++) {
          var level1 = sortedLevel1s[j];
          var level1Data = monthData.level1s[level1];
          html += generateRowHTML('      ' + level1, level1Data, 'level-3 level0-' + level0.toLowerCase() + '-year-' + year + '-month-' + monthNum + '-level1s hidden', 'level0-' + level0.toLowerCase() + '-year-' + year + '-month-' + monthNum + '-level1-' + level1.replace(/\s+/g, '-') + '-level2s');
          
          var sortedLevel2s = Object.keys(level1Data.level2s).sort();
          for (var k = 0; k < sortedLevel2s.length; k++) {
            var level2 = sortedLevel2s[k];
            var level2Data = level1Data.level2s[level2];
            html += generateRowHTML('        ' + level2, level2Data, 'level-4 level0-' + level0.toLowerCase() + '-year-' + year + '-month-' + monthNum + '-level1-' + level1.replace(/\s+/g, '-') + '-level2s hidden');
          }
        }
      }
      
      var sortedLevel1s = Object.keys(yearData.level1s).sort();
      for (var i = 0; i < sortedLevel1s.length; i++) {
        var level1 = sortedLevel1s[i];
        var level1Data = yearData.level1s[level1];
        html += generateRowHTML('    ' + level1, level1Data, 'level-2 level0-' + level0.toLowerCase() + '-year-' + year + '-children hidden', 'level0-' + level0.toLowerCase() + '-year-' + year + '-level1-' + level1.replace(/\s+/g, '-') + '-children');
        
        var sortedLevel2s = Object.keys(level1Data.level2s).sort();
        for (var j = 0; j < sortedLevel2s.length; j++) {
          var level2 = sortedLevel2s[j];
          var level2Data = level1Data.level2s[level2];
          html += generateRowHTML('      ' + level2, level2Data, 'level-3 level0-' + level0.toLowerCase() + '-year-' + year + '-level1-' + level1.replace(/\s+/g, '-') + '-children hidden', 'level0-' + level0.toLowerCase() + '-year-' + year + '-level1-' + level1.replace(/\s+/g, '-') + '-level2-' + level2.replace(/\s+/g, '-') + '-months');
          
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