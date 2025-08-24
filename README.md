# Financial Planning Application - Refactored

This Google Apps Script application has been refactored from a single large file into multiple logically organized files for better maintainability and understanding.

## File Structure

### Core Files

- **`MainCode.js`** - Main entry point containing the menu setup (`onOpen` function)
- **`TransactionImporters.js`** - Transaction import functions for different banks
- **`MappingFunctions.js`** - Level 2 category mapping and management functions
- **`SheetEditing.js`** - Sheet editing events and formula management
- **`ReportGeneration.js`** - Report generation and visualization functions

### Legacy Files

- **`Code.js`** - Original monolithic file (preserve as backup)

## Functionality Overview

### Transaction Import (`TransactionImporters.js`)
- `addDumpedTransactions()` - Import from Barclays Current Account
- `addSainsburysTransactions()` - Import from Sainsburys Bank Credit Card
- `addAmexTransactions()` - Import from American Express Credit Card

### Category Mapping (`MappingFunctions.js`)
- `mapLevel2()` - Apply existing mappings to transactions
- `autoGenerateLevel2Mappings()` - Auto-generate mappings from existing data
- `addLevel2Mapping()` - Interactive UI for adding new mappings
- `addMappingFromCurrent()` - Add mapping from currently selected transaction

### Sheet Management (`SheetEditing.js`)
- `onEdit()` - Handle sheet edit events
- `copyFormulasToNewRow()` - Auto-populate formulas for new rows
- `rowMatch()` - Utility for comparing transaction rows

### Reports (`ReportGeneration.js`)
- `createABVSheet()` - Create Actual vs Budget sheet
- `createActualVsPlannedReport()` - Generate comprehensive financial report

## How to Deploy the Refactored Code

### Method 1: Replace Existing Project
1. Open your existing Google Apps Script project
2. Replace the content of `Code.gs` with the content from `MainCode.js`
3. Create new script files for each component:
   - Add new file → Name it `TransactionImporters` → Paste content from `TransactionImporters.js`
   - Add new file → Name it `MappingFunctions` → Paste content from `MappingFunctions.js`
   - Add new file → Name it `SheetEditing` → Paste content from `SheetEditing.js`
   - Add new file → Name it `ReportGeneration` → Paste content from `ReportGeneration.js`
4. Save the project

### Method 2: Create New Project
1. Create a new Google Apps Script project
2. Replace the default `Code.gs` with content from `MainCode.js`
3. Add the other files as described in Method 1
4. Connect to your existing Google Sheets file

## Benefits of Refactoring

### 🎯 **Improved Organization**
- Functions are grouped by logical purpose
- Easier to find specific functionality
- Clear separation of concerns

### 🔧 **Better Maintainability**
- Changes to transaction import logic only affect `TransactionImporters.js`
- Report modifications are isolated to `ReportGeneration.js`
- Mapping features are self-contained in `MappingFunctions.js`

### 📚 **Enhanced Readability**
- Each file is focused on a specific domain
- Reduced cognitive load when working on specific features
- Better code documentation and comments

### 🚀 **Easier Development**
- Multiple developers can work on different files simultaneously
- Feature additions are more straightforward
- Testing specific functionality is simpler

## Google Apps Script Considerations

- **No Import/Export**: Google Apps Script doesn't support ES6 modules, so all functions are available globally
- **Execution Context**: All files run in the same execution context
- **Triggers**: The `onEdit` and `onOpen` functions work the same way across all files
- **Libraries**: All files have access to the same Google Apps Script services

## Migration Notes

- All existing functionality remains identical
- No changes to the user interface or menu structure
- All existing triggers and permissions are preserved
- Sheet structure and data remain unchanged

## Troubleshooting

If you encounter issues after refactoring:

1. **Functions not found**: Ensure all files are properly added to the Google Apps Script project
2. **Menu not appearing**: Check that `onOpen` function is in the main file and project is saved
3. **Execution errors**: Verify that all functions are copied completely without truncation

## Contributing

When making changes:
- Keep related functions in their respective files
- Update this README if adding new files or major functionality
- Test changes thoroughly in a development environment first
- Consider the impact on existing sheets and data

---

*This refactoring maintains 100% backward compatibility while significantly improving code organization and maintainability.*