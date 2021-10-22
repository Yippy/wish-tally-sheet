function onOpen( ){
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Wish Tally')
  .addSeparator()
  .addSubMenu(ui.createMenu('Character Event Wish History')
             .addItem('Sort Range', 'sortCharacterEventWishHistory')
             .addItem('Refresh Formula', 'addFormulaCharacterEventWishHistory'))
  .addSubMenu(ui.createMenu('Permanent Wish History')
             .addItem('Sort Range', 'sortPermanentWishHistory')
             .addItem('Refresh Formula', 'addFormulaByWishHistoryName'))
  .addSubMenu(ui.createMenu('Weapon Event Wish History')
             .addItem('Sort Range', 'sortWeaponEventWishHistory')
             .addItem('Refresh Formula', 'addFormulaWeaponEventWishHistory'))
  .addSubMenu(ui.createMenu('Novice Wish History')
             .addItem('Sort Range', 'sortNoviceWishHistory')
             .addItem('Refresh Formula', 'addFormulaNoviceWishHistory'))
  .addSeparator()
  .addItem('Update Items', 'updateItemsList')
  .addItem('About', 'displayAbout')
  .addToUi();
}
  
function displayAbout() {
  var sheetSource = SpreadsheetApp.openById('1mTeEQs1nOViQ-_BVHkDSZgfKGsYiLATe1mFQxypZQWA');
  if (sheetSource) {
    var aboutSource = sheetSource.getSheetByName('About');
    var titleString = aboutSource.getRange("B1").getValue();
    var htmlString = aboutSource.getRange("B2").getValue();
    var widthSize = aboutSource.getRange("B3").getValue();
    var heightSize = aboutSource.getRange("B4").getValue();
    
    
    var htmlOutput = HtmlService
    .createHtmlOutput(htmlString)
    .setWidth(widthSize) //optional
    .setHeight(heightSize); //optional
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, titleString);
  } else {
    var message = 'Unable to connect to source';
    var title = 'Error';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}
/**
* Add Formula Character Event Wish History
*/
function addFormulaCharacterEventWishHistory() {
  addFormulaByWishHistoryName('Character Event Wish History');
}
/**
* Add Formula Permanent Wish History History
*/
function addFormulaPermanentWishHistory() {
  addFormulaByWishHistoryName('Permanent Wish History');
}
/**
* Add Formula Weapon Event Wish History
*/
function addFormulaWeaponEventWishHistory() {
  addFormulaByWishHistoryName('Weapon Event Wish History');
}
/**
* Add Formula Novice Wish History
*/
function addFormulaNoviceWishHistory() {
  addFormulaByWishHistoryName('Novice Wish History');
}

function addFormulaByWishHistoryName(name) {
  var sheetSource = SpreadsheetApp.openById('1mTeEQs1nOViQ-_BVHkDSZgfKGsYiLATe1mFQxypZQWA');
  if (sheetSource) {
    var wishHistorySource = sheetSource.getSheetByName(name);
    var sheet = SpreadsheetApp.getActive().getSheetByName(name);
    
    var numberTitle = wishHistorySource.getRange("B1").getValue();
    var numberFormula = wishHistorySource.getRange("B2").getFormula();
    sheet.getRange(1, 2, 1, 1).setValue(numberTitle);
    sheet.getRange(2, 2, sheet.getLastRow(), 1).setValue(numberFormula);

    var weaponTitle = wishHistorySource.getRange("C1").getValue();
    var weaponFormula = wishHistorySource.getRange("C2").getFormula();
    sheet.getRange(1, 3, 1, 1).setValue(weaponTitle);
    sheet.getRange(2, 3, sheet.getLastRow(), 1).setValue(weaponFormula);

    var dateAndTimeTitle = wishHistorySource.getRange("D1").getValue();
    var dateAndTimeFormula = wishHistorySource.getRange("D2").getFormula();
    sheet.getRange(1, 4, 1, 1).setValue(dateAndTimeTitle);
    sheet.getRange(2, 4, sheet.getLastRow(), 1).setValue(dateAndTimeFormula);

    var itemNameTitle = wishHistorySource.getRange("E1").getValue();
    var itemNameFormula = wishHistorySource.getRange("E2").getFormula();
    sheet.getRange(1, 5, 1, 1).setValue(itemNameTitle);
    sheet.getRange(2, 5, sheet.getLastRow(), 1).setValue(itemNameFormula);

    var itemRarityTitle = wishHistorySource.getRange("F1").getValue();
    var itemRarityFormula = wishHistorySource.getRange("F2").getFormula();
    sheet.getRange(2, 6, sheet.getLastRow(), 1).setValue(itemRarityFormula);
  } else {
    var message = 'Unable to connect to source';
    var title = 'Error';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}
/**
* Sort Character Event Wish History
*/
function sortCharacterEventWishHistory() {
  sortWishHistoryByName('Character Event Wish History');
}

/**
* Sort Permanent Wish History
*/
function sortPermanentWishHistory() {
  sortWishHistoryByName('Permanent Wish History');
}

/**
* Sort Weapon Event Wish History
*/
function sortWeaponEventWishHistory() {
  sortWishHistoryByName('Weapon Event Wish History');
}

/**
* Sort Novice Wish History
*/
function sortNoviceWishHistory() {
  sortWishHistoryByName('Novice Wish History');
}

function sortWishHistoryByName(sheetName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var range = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn());
  range.sort([{column: 4, ascending: true}, {column: 6, ascending: true}]);
}

/**
* Update Item List
*/
function updateItemsList() {
  
  var sheetSource = SpreadsheetApp.openById('1mTeEQs1nOViQ-_BVHkDSZgfKGsYiLATe1mFQxypZQWA');
  // Check source is available
  if (sheetSource) {
    // Remove sheets
    var listOfSheetsToRemove = ["Items","Pity Checker","Results","Changelog"];
    var listOfSheetsToRemoveLength = listOfSheetsToRemove.length;

    for (var i = 0; i < listOfSheetsToRemoveLength; i++) {
      var sheetToRemove = SpreadsheetApp.getActive().getSheetByName(listOfSheetsToRemove[i]);
      if(sheetToRemove) {
        // If exist remove from spreadsheet
        SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheetToRemove);
      }
    }

    var sheetItemSource = sheetSource.getSheetByName('Items');
    sheetItemSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('Items');

    // Refresh spreadsheet
    var listOfSheets = ["Character Event Wish History","Permanent Wish History","Weapon Event Wish History","Novice Wish History"];
    var listOfSheetsLength = listOfSheets.length;
    
    for (var i = 0; i < listOfSheetsLength; i++) {
      addFormulaByWishHistoryName(listOfSheets[i]);
      /*
      var sheetOld = SpreadsheetApp.getActive().getSheetByName(listOfSheets[i]);
      var sheetCopySource = sheetSource.getSheetByName(listOfSheets[i]);
      sheetCopySource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(listOfSheets[i] + " copy");
      var sheetNew = SpreadsheetApp.getActive().getSheetByName(listOfSheets[i] + " copy");
      var sourceRange = sheetNew.getRange("B:F");
      var targetRange = sheetOld.getRange("B:F");
      sourceRange.copyTo(targetRange);
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheetNew);*/
    }
    SpreadsheetApp.flush();
    Utilities.sleep(100);
    var sheetPityCheckerSource = sheetSource.getSheetByName('Pity Checker');
    sheetPityCheckerSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('Pity Checker');
    var sheetResultsSource = sheetSource.getSheetByName('Results');
    sheetResultsSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('Results');
    var sheetChangelogSource = sheetSource.getSheetByName('Changelog');
    sheetChangelogSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('Changelog');
  } else {
    var message = 'Unable to connect to source';
    var title = 'Error';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}