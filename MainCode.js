/**
 * Financial Planning Google Apps Script
 * 
 * This is the main entry point for the financial planning application.
 * 
 * Files in this project:
 * - MainCode.js: Main entry point with menu setup
 * - TransactionImporters.js: Functions for importing transactions from different banks
 * - MappingFunctions.js: Functions for managing Level 2 category mappings
 * - SheetEditing.js: Functions for handling sheet edits and formula management
 * - ReportGeneration.js: Functions for generating financial reports
 * 
 * To use this refactored code in Google Apps Script:
 * 1. Create a new Google Apps Script project
 * 2. Replace the default Code.gs with the contents of MainCode.js
 * 3. Add each of the other .js files as separate script files in the project
 * 4. The functions from all files will be available throughout the project
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
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