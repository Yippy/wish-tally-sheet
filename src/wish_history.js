/*
 * Version 3.0.1 made by yippym - 2021-10-22 21:00
 * https://github.com/Yippy/wish-tally-sheet
 */

/**
* Add Formula Character Event Wish History
*/
function addFormulaCharacterEventWishHistory() {
  addFormulaByWishHistoryName(WISH_TALLY_CHARACTER_EVENT_WISH_SHEET_NAME);
}
/**
* Add Formula Permanent Wish History History
*/
function addFormulaPermanentWishHistory() {
  addFormulaByWishHistoryName(WISH_TALLY_PERMANENT_WISH_SHEET_NAME);
}
/**
* Add Formula Weapon Event Wish History
*/
function addFormulaWeaponEventWishHistory() {
  addFormulaByWishHistoryName(WISH_TALLY_WEAPON_EVENT_WISH_SHEET_NAME);
}
/**
* Add Formula Novice Wish History
*/
function addFormulaNoviceWishHistory() {
  addFormulaByWishHistoryName(WISH_TALLY_NOVICE_WISH_SHEET_NAME);
}

/**
* Add Formula for selected Wish History sheet
*/
function addFormulaWishHistory() {
  var sheetActive = SpreadsheetApp.getActiveSpreadsheet();
  var wishHistoryName = sheetActive.getSheetName();
  if (WISH_TALLY_NAME_OF_WISH_HISTORY.indexOf(wishHistoryName) != -1) {
    addFormulaByWishHistoryName(wishHistoryName);
  } else {
    var message = 'Sheet must be called "' + WISH_TALLY_CHARACTER_EVENT_WISH_SHEET_NAME + '" or "' + WISH_TALLY_PERMANENT_WISH_SHEET_NAME + '" or "' + WISH_TALLY_WEAPON_EVENT_WISH_SHEET_NAME + '" or "' + WISH_TALLY_NOVICE_WISH_SHEET_NAME + '"';
    var title = 'Invalid Sheet Name';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function addFormulaByWishHistoryName(name) {
  var sheetSource = SpreadsheetApp.openById(WISH_TALLY_SHEET_SOURCE_ID);
  if (sheetSource) {
    // Add Language
    var wishHistorySource;
    var settingsSheet = getSettingsSheet();
    if (settingsSheet) {
      var languageFound = settingsSheet.getRange(2, 2).getValue();
      wishHistorySource = sheetSource.getSheetByName(WISH_TALLY_WISH_HISTORY_SHEET_NAME+"-"+languageFound);
    }
    if (wishHistorySource) {
      // Found language
    } else {
      // Default
      wishHistorySource = sheetSource.getSheetByName(WISH_TALLY_WISH_HISTORY_SHEET_NAME);
    }
    var sheet = findWishHistoryByName(name,sheetSource);

    var wishHistorySourceNumberOfColumn = wishHistorySource.getLastColumn();
    // Reduce two column due to paste and override
    var wishHistorySourceNumberOfColumnWithFormulas = wishHistorySourceNumberOfColumn - 2;

    var lastRowWithoutTitle = sheet.getMaxRows() - 1;

    var currentOverrideTitleCell = sheet.getRange(1, 2).getValue();
    var sourceOverrideTitleCell = wishHistorySource.getRange(1, 2).getValue();
    if (currentOverrideTitleCell != sourceOverrideTitleCell) {
      // If override column don't exist, populate from source
      var overrideCells = wishHistorySource.getRange(2, 2).getFormula();
      sheet.getRange(2, 2, lastRowWithoutTitle, 1).setValue(overrideCells);
      sheet.getRange(1, 2).setValue(sourceOverrideTitleCell);
      sheet.setColumnWidth(2, wishHistorySource.getColumnWidth(2));
    }
    
    // Get second row formula columns and set current sheet
    var formulaCells = wishHistorySource.getRange(2, 3, 1, wishHistorySourceNumberOfColumnWithFormulas).getFormulas();
    sheet.getRange(2, 3, lastRowWithoutTitle, wishHistorySourceNumberOfColumnWithFormulas).setValue(formulaCells);

    // Get title columns and set current sheet
    var titleCells = wishHistorySource.getRange(1, 3, 1, wishHistorySourceNumberOfColumnWithFormulas).getFormulas();
    sheet.getRange(1, 3, 1, wishHistorySourceNumberOfColumnWithFormulas).setValues(titleCells);

    for (var i = 3; i <= wishHistorySourceNumberOfColumn; i++) {
      // Apply formatting for cells
      var numberFormatCell = wishHistorySource.getRange(2, i).getNumberFormat();
      sheet.getRange(2, i, lastRowWithoutTitle, 1).setNumberFormat(numberFormatCell);
      // Set column width from source
      sheet.setColumnWidth(i, wishHistorySource.getColumnWidth(i));
    }

    // Ensure new row is not the same height as first, if row 2 did not exist
    sheet.autoResizeRows(2, 1);
  } else {
    var message = 'Unable to connect to source';
    var title = 'Error';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

/**
* Check is sheet exist in active spreadsheet, otherwise pull sheet from source
*/
function findWishHistoryByName(name, sheetSource) {
  var wishHistorySheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (wishHistorySheet == null) {
    if (sheetSource == null) {
      sheetSource = SpreadsheetApp.openById(WISH_TALLY_SHEET_SOURCE_ID);
    }
    if (sheetSource) {
      var sheetCopySource = sheetSource.getSheetByName(WISH_TALLY_WISH_HISTORY_SHEET_NAME);
      sheetCopySource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(name);
      wishHistorySheet = SpreadsheetApp.getActive().getSheetByName(name);
      wishHistorySheet.showSheet();
    }
  }
  return wishHistorySheet;
}

/**
* Add sort for selected Wish History sheet
*/
function sortWishHistory() {
  var sheetActive = SpreadsheetApp.getActiveSpreadsheet();
  var wishHistoryName = sheetActive.getSheetName();
  if (WISH_TALLY_NAME_OF_WISH_HISTORY.indexOf(wishHistoryName) != -1) {
    sortWishHistoryByName(wishHistoryName);
  } else {
    var message = 'Sheet must be called "' + WISH_TALLY_CHARACTER_EVENT_WISH_SHEET_NAME + '" or "' + WISH_TALLY_PERMANENT_WISH_SHEET_NAME + '" or "' + WISH_TALLY_WEAPON_EVENT_WISH_SHEET_NAME + '" or "' + WISH_TALLY_NOVICE_WISH_SHEET_NAME + '"';
    var title = 'Invalid Sheet Name';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

/**
* Sort Character Event Wish History
*/
function sortCharacterEventWishHistory() {
  sortWishHistoryByName(WISH_TALLY_CHARACTER_EVENT_WISH_SHEET_NAME);
}

/**
* Sort Permanent Wish History
*/
function sortPermanentWishHistory() {
  sortWishHistoryByName(WISH_TALLY_PERMANENT_WISH_SHEET_NAME);
}

/**
* Sort Weapon Event Wish History
*/
function sortWeaponEventWishHistory() {
  sortWishHistoryByName(WISH_TALLY_WEAPON_EVENT_WISH_SHEET_NAME);
}

/**
* Sort Novice Wish History
*/
function sortNoviceWishHistory() {
  sortWishHistoryByName(WISH_TALLY_NOVICE_WISH_SHEET_NAME);
}

function sortWishHistoryByName(sheetName) {
  var sheet = findWishHistoryByName(sheetName, null);
  if (sheet) {
    if (sheet.getLastColumn() > 6) {
      var range = sheet.getRange(2, 1, sheet.getMaxRows()-1, sheet.getLastColumn());
      range.sort([{column: 5, ascending: true}, {column: 2, ascending: true}, {column: 7, ascending: true}]);
    } else {
      var message = 'Invalid number of columns to sort, run "Refresh Formula" or "Update Items"';
      var title = 'Error';
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
    }
  } else {
    var message = 'Unable to connect to source';
    var title = 'Error';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}