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