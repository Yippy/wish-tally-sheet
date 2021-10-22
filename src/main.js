/*
 * Version 2.7 made by yippym
 */

var sheetSourceId = '1mTeEQs1nOViQ-_BVHkDSZgfKGsYiLATe1mFQxypZQWA';
var nameOfWishHistorys = ["Character Event Wish History", "Permanent Wish History", "Weapon Event Wish History", "Novice Wish History"];

function onOpen( ){
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Wish Tally')
  .addSeparator()
  .addSubMenu(ui.createMenu('Character Event Wish History')
             .addItem('Sort Range', 'sortCharacterEventWishHistory')
             .addItem('Refresh Formula', 'addFormulaCharacterEventWishHistory'))
  .addSubMenu(ui.createMenu('Permanent Wish History')
             .addItem('Sort Range', 'sortPermanentWishHistory')
             .addItem('Refresh Formula', 'addFormulaPermanentWishHistory'))
  .addSubMenu(ui.createMenu('Weapon Event Wish History')
             .addItem('Sort Range', 'sortWeaponEventWishHistory')
             .addItem('Refresh Formula', 'addFormulaWeaponEventWishHistory'))
  .addSubMenu(ui.createMenu('Novice Wish History')
             .addItem('Sort Range', 'sortNoviceWishHistory')
             .addItem('Refresh Formula', 'addFormulaNoviceWishHistory'))
  .addSeparator()
  .addSubMenu(ui.createMenu('AutoHotkey')
             .addItem('Clear', 'clearAHK')
             .addItem('Convert', 'convertAHK')
             .addItem('Import', 'importAHK')
             .addSeparator()
             .addItem('Generate', 'generateAHK'))
  .addSeparator()
  .addSubMenu(ui.createMenu('Data Management')
             .addItem('Import', 'importDataManagement')
             .addSeparator()
             .addItem('Set Schedule', 'setTriggerDataManagement')
             .addItem('Remove All Schedule', 'removeTriggerDataManagement'))
  .addSeparator()
  .addItem('Quick Update', 'quickUpdate')
  .addItem('Update Items', 'updateItemsList')
  .addItem('Get Latest README', 'displayReadme')
  .addItem('About', 'displayAbout')
  .addToUi();
}

var triggerOnString = "ON";
var triggerOffString = "OFF";

var runAtWhichDay = {
  "Monday": ScriptApp.WeekDay.MONDAY,
  "Tuesday": ScriptApp.WeekDay.TUESDAY,
  "Wednesday": ScriptApp.WeekDay.WEDNESDAY,
  "Thursday": ScriptApp.WeekDay.THURSDAY,
  "Friday": ScriptApp.WeekDay.FRIDAY,
  "Saturday": ScriptApp.WeekDay.SATURDAY,
  "Sunday": ScriptApp.WeekDay.SUNDAY
};

var runAtHour = {
  "Run at 1:00": 1,
  "Run at 2:00": 2,
  "Run at 3:00": 3,
  "Run at 4:00": 4,
  "Run at 5:00": 5,
  "Run at 6:00": 6,
  "Run at 7:00": 7,
  "Run at 8:00": 8,
  "Run at 9:00": 9,
  "Run at 10:00": 10,
  "Run at 11:00": 11,
  "Run at 12:00": 12,
  "Run at 13:00": 13,
  "Run at 14:00": 14,
  "Run at 15:00": 15,
  "Run at 16:00": 16,
  "Run at 17:00": 17,
  "Run at 18:00": 18,
  "Run at 19:00": 19,
  "Run at 20:00": 20,
  "Run at 21:00": 21,
  "Run at 22:00": 22,
  "Run at 23:00": 23,
  "Run at Midnight": 0
};

var runAtEveryHour = {
  "Every hour": 1,
  "Every 2 hours": 2,
  "Every 3 hours": 3,
  "Every 4 hours": 4,
  "Every 5 hours": 5,
  "Every 6 hours": 6,
  "Every 7 hours": 7,
  "Every 8 hours": 8,
  "Every 9 hours": 9,
  "Every 10 hours": 10,
  "Every 11 hours": 11,
  "Every 12 hours": 12,
  "Every 13 hours": 13,
  "Every 14 hours": 14,
  "Every 15 hours": 15,
  "Every 16 hours": 16,
  "Every 17 hours": 17,
  "Every 18 hours": 18,
  "Every 19 hours": 19,
  "Every 20 hours": 20,
  "Every 21 hours": 21,
  "Every 22 hours": 22,
  "Every 23 hours": 23,
  "Every 24 hours": 24
};

function createWeeklyTrigger(functionName, day, hour, minutes) {
  ScriptApp.newTrigger(functionName)
  .timeBased()
  .onWeekDay(day)
  .atHour(hour).nearMinute(minutes)
  .create();
}

function createHourlyTrigger(functionName, hour, minutes) {
  ScriptApp.newTrigger(functionName)
  .timeBased()
  .everyHours(hour).nearMinute(minutes)
  .create();
}

function setTriggerDataManagement() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName('Settings');
  removeTriggers();
  if (settingsSheet) {
    var isScheduleAvailable = false;
    
    // Update Items List trigger
    var updateItemsSchedule = settingsSheet.getRange("E20").getValue();
    
    var gotHour = runAtHour[updateItemsSchedule];
    if (gotHour || gotHour == 0) {
      isScheduleAvailable = true;
      createWeeklyTrigger("updateItemsListTriggered", runAtWhichDay["Sunday"], gotHour, 1);
      settingsSheet.getRange("E21").setValue(updateItemsSchedule);
    } else {
      settingsSheet.getRange("E21").setValue(triggerOffString);
    }

    // Sort trigger
    var sortRangeSchedule = settingsSheet.getRange("E27").getValue();
    var gotRunHourly = runAtEveryHour[sortRangeSchedule];
    if (gotRunHourly) {
      isScheduleAvailable = true;
      createHourlyTrigger("sortRangesTriggered", gotRunHourly, 0);
      settingsSheet.getRange("E28").setValue(sortRangeSchedule);
    } else {
      settingsSheet.getRange("E28").setValue(triggerOffString);
    }

    // scheduledTrigger(1,00, "updateItemsListTriggered");
    if (isScheduleAvailable) {
      settingsSheet.getRange("E16").setValue(triggerOnString);
    } else {
      settingsSheet.getRange("E16").setValue(triggerOffString);
    }
  }
}

function updateItemsListTriggered() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName('Settings');
  settingsSheet.getRange("E22").setValue(new Date());
  updateItemsList();
  settingsSheet.getRange("E23").setValue(new Date());
}

function sortRangesTriggered() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName('Settings');
  settingsSheet.getRange("E29").setValue(new Date());
  sortCharacterEventWishHistory();
  sortPermanentWishHistory();
  sortWeaponEventWishHistory();
  sortNoviceWishHistory();
  settingsSheet.getRange("E30").setValue(new Date());
}

function scheduledTrigger(hours,minutes,functionName){
  var today_D = new Date();
  var year = today_D.getFullYear();
  var month = today_D.getMonth();
  var day = today_D.getDate();
  
  pars = [year,month,day,hours,minutes];
  
  var scheduled_D = new Date(...pars);
  var hours_remain=Math.abs(scheduled_D - today_D) / 36e5;
  ScriptApp.newTrigger(functionName)
  .timeBased()
  .after(hours_remain * 60 *60 * 1000)
  .create()
}

function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (i in triggers) {
    //if ((triggers[i].getHandlerFunction()) == "createStats") {
      ScriptApp.deleteTrigger(triggers[i]);
   // }
  }
}

function removeTriggerDataManagement() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName('Settings');
  removeTriggers();
  var title = "Remove Schedule";
  var message = "All schedule has been removed";
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  settingsSheet.getRange("E16").setValue(triggerOffString);
  settingsSheet.getRange("E21").setValue(triggerOffString);
  settingsSheet.getRange("E28").setValue(triggerOffString);
}

function importDataManagement() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName('Settings');
  var userImportInput = settingsSheet.getRange("D6").getValue();
  var userImportStatus = settingsSheet.getRange("E7").getValue();
  var completeStatus = "COMPLETE";
  var wishHistoryNotDoneStatus = "NOT DONE";
  var wishHistoryDoneStatus = "DONE";
  var wishHistoryMissingStatus = "NOT FOUND";
  var message = "";
  var title = "";
  var statusMessage = "";
  var rowOfStatusWishHistory = 9;
  if (userImportStatus == completeStatus) {
      title = "Error";
      message = "Already done, you only need to run once";
  } else {
    if (userImportInput) {
      // Attempt to load as URL
      var importSource = SpreadsheetApp.openByUrl(userImportInput);
      if (importSource) {
      } else {
        // Attempt to load as ID instead
        importSource = SpreadsheetApp.openById(userImportInput);
      }
      if (importSource) {
        // Go through the available sheet list
        for (var i = 0; i < nameOfWishHistorys.length; i++) {
          var bannerImportSheet = importSource.getSheetByName(nameOfWishHistorys[i]);
          
          var numberOfRows = bannerImportSheet.getMaxRows()-1;
          var range = bannerImportSheet.getRange(2, 1, numberOfRows, 2);

          if (bannerImportSheet && numberOfRows > 0) {
            var bannerSheet = SpreadsheetApp.getActive().getSheetByName(nameOfWishHistorys[i]);

            if (bannerSheet) {
              bannerSheet.getRange(2, 1, numberOfRows, 2).setValues(range.getValues());
              settingsSheet.getRange(rowOfStatusWishHistory+i, 5).setValue(wishHistoryDoneStatus);
            } else {
              settingsSheet.getRange(rowOfStatusWishHistory+i, 5).setValue(wishHistoryMissingStatus);
            }
          } else {
            settingsSheet.getRange(rowOfStatusWishHistory+i, 5).setValue(wishHistoryMissingStatus);
          }
        }
        var sourceSettingsSheet = importSource.getSheetByName("Settings");
        if (sourceSettingsSheet) {
          var sourcePityCheckerSheet = importSource.getSheetByName("Pity Checker");
          if (sourcePityCheckerSheet) {
            savePityCheckerSettings(sourcePityCheckerSheet, settingsSheet);
          }
          if (sourceSettingsSheet.getMaxColumns() >= 8) {
            var version = sourceSettingsSheet.getRange("H1").getValue();
            if (version == "2.7") {
              var pityCheckerIsShow4Star = sourceSettingsSheet.getRange("B18").getValue();
              settingsSheet.getRange("B18").setValue(pityCheckerIsShow4Star == true);
              var pityCheckerIsShow5Star = sourceSettingsSheet.getRange("B19").getValue();
              settingsSheet.getRange("B19").setValue(pityCheckerIsShow5Star == true);
            }
          }
          var pityCheckerSheet = SpreadsheetApp.getActive().getSheetByName('Pity Checker');
          if (pityCheckerSheet) {
            restorePityCheckerSettings(pityCheckerSheet, settingsSheet);
          }
          var offset = sourceSettingsSheet.getRange("B10").getValue();
          if (offset >= -11 && offset <= 12) {
             settingsSheet.getRange("B10").setValue(offset);
          }
          var language = sourceSettingsSheet.getRange("B2").getValue();
          if (language) {
             settingsSheet.getRange("B2").setValue(language);
          }
          var server = sourceSettingsSheet.getRange("B3").getValue();
          if (server) {
             settingsSheet.getRange("B3").setValue(server);
          }
        }
        //Restore Events
        var sourceEventsSheet = importSource.getSheetByName("Events");
        if (sourceEventsSheet) {
          saveEventsSettings(sourceEventsSheet,settingsSheet);
          var eventsSheet = SpreadsheetApp.getActive().getSheetByName('Events');
          if (eventsSheet) {
            restoreEventsSettings(eventsSheet, settingsSheet)
          }
        }
        
        //Restore Results
        var sourceResultsSheet = importSource.getSheetByName("Results");
        if (sourceResultsSheet) {
          saveResultsSettings(sourceResultsSheet, settingsSheet);
          var resultsSheet = SpreadsheetApp.getActive().getSheetByName('Results');
          if (resultsSheet) {
            restoreResultsSettings(resultsSheet, settingsSheet)
          }
        }
        
        title = "Complete";
        message = "Imported all rows in column Paste Value and Override";
        statusMessage = completeStatus;
      } else {
        title = "Error";
        message = "Import From URL or Spreadsheet ID is invalid";
        statusMessage = "Failed";
      }
    } else {
      title = "Error";
      message = "Import From URL or Spreadsheet ID is empty";
      statusMessage = "Failed";
    }

    settingsSheet.getRange("E7").setValue(statusMessage);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
}
/**
 * Return the id of the sheet. (by name)
 *
 * @return The ID of the sheet
 * @customfunction
 */
function GET_SHEET_ID(sheetName) {
    var sheetId = SpreadsheetApp.getActive().getSheetByName(sheetName).getSheetId();
    return sheetId;
}

/**
 * Return the id of this current Spreadsheet
 *
 * @return The ID of the Spreadsheet
 * @customfunction
 */
function GET_SPREADSHEET_ID() {
    var spreadsheetId = SpreadsheetApp.getActive().getId();
    return spreadsheetId;
}

function displayAbout() {
  var sheetSource = SpreadsheetApp.openById(sheetSourceId);
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

function displayReadme() {
  var sheetSource = SpreadsheetApp.openById(sheetSourceId);
  if (sheetSource) {
    // Avoid Exception: You can't remove all the sheets in a document.Details
    var placeHolderSheet = null;
    if (SpreadsheetApp.getActive().getSheets().length == 1) {
      placeHolderSheet = SpreadsheetApp.getActive().insertSheet();
    }
    var sheetToRemove = SpreadsheetApp.getActive().getSheetByName('README');
      if(sheetToRemove) {
        // If exist remove from spreadsheet
        SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheetToRemove);
      }
    var sheetREADMESource;

    // Add Language
    var settingsSheet = SpreadsheetApp.getActive().getSheetByName("Settings");
    if (settingsSheet) {
      var languageFound = settingsSheet.getRange(2, 2).getValue();
      sheetREADMESource = sheetSource.getSheetByName("README"+"-"+languageFound);
    }
    if (sheetREADMESource) {
      // Found language
    } else {
      // Default
      sheetREADMESource = sheetSource.getSheetByName("README");
    }

    sheetREADMESource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('README');

    // Remove placeholder if available
    if(placeHolderSheet) {
      // If exist remove from spreadsheet
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(placeHolderSheet);
    }
    var sheetREADME = SpreadsheetApp.getActive().getSheetByName('README');
    // Refresh Contents Links
    var contentsAvailable = sheetREADME.getRange(13, 1).getValue();
    var contentsStartIndex = 15;
    
    for (var i = 0; i < contentsAvailable; i++) {
      var valueRange = sheetREADME.getRange(contentsStartIndex+i, 1).getValue();
      var formulaRange = sheetREADME.getRange(contentsStartIndex+i, 1).getFormula();
      // Display for user, doesn't do anything
      sheetREADME.getRange(contentsStartIndex+i, 1).setFormula(formulaRange);
 
      // Grab URL RichTextValue from Source
      const range = sheetREADMESource.getRange(contentsStartIndex+i, 1);
      const RichTextValue = range.getRichTextValue().getRuns();
      const res = RichTextValue.reduce((ar, e) => {
        const url = e.getLinkUrl();
        if (url) ar.push(url);
          return ar;
        }, []);
      //  Convert to string
      var resString = res+ "";
      var arrayString = resString.split("=");
      if (arrayString.length > 1) {
        var text = arrayString[2];
        const richText = SpreadsheetApp.newRichTextValue()
          .setText(valueRange)
          .setLinkUrl(["#gid="+GET_SHEET_ID("README")+'range='+text])
          .build();
        sheetREADME.getRange(contentsStartIndex+i, 1).setRichTextValue(richText);
      }
    }
 
    SpreadsheetApp.getActive().setActiveSheet(sheetREADME);
    SpreadsheetApp.getActive().moveActiveSheet(1);
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

/**
* Add Formula for selected Wish History sheet
*/
function addFormulaWishHistory() {
  var sheetActive = SpreadsheetApp.getActiveSpreadsheet();
  var wishHistoryName = sheetActive.getSheetName();
  if (nameOfWishHistorys.indexOf(wishHistoryName) != -1) {
    addFormulaByWishHistoryName(wishHistoryName);
  } else {
    var message = 'Sheet must be called "Character Event Wish History" or "Permanent Wish History" or "Weapon Event Wish History" or "Novice Wish History"';
    var title = 'Invalid Sheet Name';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function addFormulaByWishHistoryName(name) {
  var sheetSource = SpreadsheetApp.openById(sheetSourceId);
  if (sheetSource) {
    // Add Language
    var wishHistorySource;
    var settingsSheet = SpreadsheetApp.getActive().getSheetByName("Settings");
    if (settingsSheet) {
      var languageFound = settingsSheet.getRange(2, 2).getValue();
      wishHistorySource = sheetSource.getSheetByName("Wish History"+"-"+languageFound);
    }
    if (wishHistorySource) {
      // Found language
    } else {
      // Default
      wishHistorySource = sheetSource.getSheetByName("Wish History");
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
      sheetSource = SpreadsheetApp.openById(sheetSourceId);
    }
    if (sheetSource) {
      var sheetCopySource = sheetSource.getSheetByName("Wish History");
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
  if (nameOfWishHistorys.indexOf(wishHistoryName) != -1) {
    sortWishHistoryByName(wishHistoryName);
  } else {
    var message = 'Sheet must be called "Character Event Wish History" or "Permanent Wish History" or "Weapon Event Wish History" or "Novice Wish History"';
    var title = 'Invalid Sheet Name';
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

function savePityCheckerSettings(pityCheckerSheet, settingsSheet) {
  var isDarkMode = pityCheckerSheet.getRange("AV1").getValue();
  var isShowTimer = pityCheckerSheet.getRange("F1").getValue();
  settingsSheet.getRange("G2").setValue(isDarkMode);
  settingsSheet.getRange("H2").setValue(isShowTimer);
}

function saveResultsSettings(resultsSheet, settingsSheet) {
  var isDarkMode = resultsSheet.getRange("B9").getValue();
  settingsSheet.getRange("G3").setValue(isDarkMode);
  var selectionValueRange = resultsSheet.getRange(10,1,1,4).getValues();
  selectionValueRange = String(selectionValueRange).split(",");
  settingsSheet.getRange("H3").setValue(selectionValueRange.join(","));
}

function saveEventsSettings(eventsSheet, settingsSheet) {
  var eventsValueRange = eventsSheet.getRange(2,8, eventsSheet.getMaxRows()-1,1).getValues();
  eventsValueRange = String(eventsValueRange).split(",");
  var eventFormulaRanges = eventsSheet.getRange(2,8, eventsSheet.getMaxRows()-1,1).getFormulas();
  var saveDate = [];
  for (var ii = 0; ii < eventFormulaRanges.length; ii++) {
    var formulaData = eventFormulaRanges[ii];
    if (formulaData == "") {
      var valueData = eventsValueRange[ii];
      if (valueData == "true") {
        saveDate.push("TRUE");
      } else if (valueData == "false") {
        saveDate.push("");
      } else {
        saveDate.push(eventsValueRange[ii]);
      }
    } else {
      saveDate.push("");
    }
  }
  settingsSheet.getRange("G4").setValue(saveDate.join(","));
}

function restoreEventsSettings(sheetEvents, settingsSheet) {
  var saveDate = settingsSheet.getRange("G4").getValue().split(",");
  for (var ii = 0; ii < saveDate.length; ii++) {
    var valueData = saveDate[ii];
    if (valueData == "TRUE") {
      sheetEvents.getRange(2 + ii,8).setValue(true);
    } else if (valueData) {
      if (valueData != "") {
        sheetEvents.getRange(2 + ii,8).setValue(valueData);
      }
    }
  }
}

function restorePityCheckerSettings(sheetPityChecker, settingsSheet) {
  var isDarkMode =  settingsSheet.getRange("G2").getValue();
  if (isDarkMode != sheetPityChecker.getRange("AV1").getValue()) {
    sheetPityChecker.getRange("AV1").setValue(isDarkMode);
  }
  var isShowTimer = settingsSheet.getRange("H2").getValue();
  if (isShowTimer != settingsSheet.getRange("F1").getValue()) {
    sheetPityChecker.getRange("F1").setValue(isShowTimer);
  }

  var itemNameFor4Star = settingsSheet.getRange('B18').getValue();
  var itemNameFor5Star = settingsSheet.getRange('B19').getValue();
  if (itemNameFor4Star) {
    // Character Event Wish
    sheetPityChecker.hideColumns(2)
    sheetPityChecker.showColumns(4)
    // Permanent Wish History
    sheetPityChecker.hideColumns(14)
    sheetPityChecker.showColumns(16)
    // Weapon Wish History
    sheetPityChecker.hideColumns(26)
    sheetPityChecker.showColumns(28)
    // Novice Wish History
    sheetPityChecker.hideColumns(38)
    sheetPityChecker.showColumns(40)
  } else {
    // Character Event Wish
    sheetPityChecker.showColumns(2)
    sheetPityChecker.hideColumns(4)
    // Permanent Wish History
    sheetPityChecker.showColumns(14)
    sheetPityChecker.hideColumns(16)
    // Weapon Wish History
    sheetPityChecker.showColumns(26)
    sheetPityChecker.hideColumns(28)
    // Novice Wish History
    sheetPityChecker.showColumns(38)
    sheetPityChecker.hideColumns(40)
  }
  if (itemNameFor5Star) {
    // Character Event Wish
    sheetPityChecker.hideColumns(8)
    sheetPityChecker.showColumns(10)
    // Permanent Wish History
    sheetPityChecker.hideColumns(20)
    sheetPityChecker.showColumns(22)
    // Weapon Wish History
    sheetPityChecker.hideColumns(32)
    sheetPityChecker.showColumns(34)
    // Novice Wish History
    sheetPityChecker.hideColumns(44)
    sheetPityChecker.showColumns(46)
  } else {
    // Character Event Wish
    sheetPityChecker.showColumns(8)
    sheetPityChecker.hideColumns(10)
    // Permanent Wish History
    sheetPityChecker.showColumns(20)
    sheetPityChecker.hideColumns(22)
    // Weapon Wish History
    sheetPityChecker.showColumns(32)
    sheetPityChecker.hideColumns(34)
    // Novice Wish History
    sheetPityChecker.showColumns(44)
    sheetPityChecker.hideColumns(46)
  }
}

function restoreResultsSettings(sheetResults, settingsSheet) {
  var isDarkMode = settingsSheet.getRange("G3").getValue();
  if (isDarkMode != sheetResults.getRange("B9").getValue()) {
    sheetResults.getRange("B9").setValue(isDarkMode);
  }

  var saveDate = settingsSheet.getRange("H3").getValue().split(",");
  for (var ii = 0; ii < saveDate.length; ii++) {
    var valueData = saveDate[ii];
    sheetResults.getRange(10,1 + ii).setValue(valueData);
  }
}

function quickUpdate() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName("Settings");
  var sheetSource = SpreadsheetApp.openById(sheetSourceId);
  if (settingsSheet) {
    var sheetPityChecker = SpreadsheetApp.getActive().getSheetByName("Pity Checker");
    if (sheetPityChecker) {
      restorePityCheckerSettings(sheetPityChecker, settingsSheet);
      if (sheetSource) {
        var sheetPityCheckerSource = sheetSource.getSheetByName("Pity Checker");
        if (sheetPityCheckerSource) {
          var formula;
          var value;
          // Banner Images
          formula = sheetPityCheckerSource.getRange('A2').getFormula();
          sheetPityChecker.getRange('A2').setFormula(formula);
          formula = sheetPityCheckerSource.getRange('M2').getFormula();
          sheetPityChecker.getRange('M2').setFormula(formula);
          formula = sheetPityCheckerSource.getRange('Y2').getFormula();
          sheetPityChecker.getRange('Y2').setFormula(formula);
          formula = sheetPityCheckerSource.getRange('AK2').getFormula();
          sheetPityChecker.getRange('AK2').setFormula(formula);
          
          // Banner Time
          value = sheetPityCheckerSource.getRange('A3').getValue();
          sheetPityChecker.getRange('A3').setValue(value);
          value = sheetPityCheckerSource.getRange('M3').getValue();
          sheetPityChecker.getRange('M3').setValue(value);
          value = sheetPityCheckerSource.getRange('Y3').getValue();
          sheetPityChecker.getRange('Y3').setValue(value);
        }
      }
    }
  }
}

/**
* Update Item List
*/
function updateItemsList() {
  var sheetSource = SpreadsheetApp.openById(sheetSourceId);
  // Check source is available
  if (sheetSource) {
    // Avoid Exception: You can't remove all the sheets in a document.Details
    var placeHolderSheet = null;
    if (SpreadsheetApp.getActive().getSheets().length == 1) {
      placeHolderSheet = SpreadsheetApp.getActive().insertSheet();
    }
    var settingsSheet = SpreadsheetApp.getActive().getSheetByName("Settings");
    if (settingsSheet) {
      var isLoading = settingsSheet.getRange(5, 7).getValue();
      if (isLoading) {
        var counter = settingsSheet.getRange(5, 8).getValue();
        if (counter > 0) {
          counter++;
          settingsSheet.getRange(5, 8).setValue(counter);
        } else {
          settingsSheet.getRange(5, 8).setValue(1);
        }
        if (counter > 2) {
          // Bypass message - for people with broken update wanting force update
        } else {
          var message = 'Still updating';
          var title = 'Update already started, the number of time you requested is '+counter+'. If you want to force an update due to an error happened during update, proceed in calling "Update Item" one more try.';
          SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
          return;
        }
      } else {
        settingsSheet.getRange(5, 7).setValue(true);
        settingsSheet.getRange(5, 8).setValue(1);
        settingsSheet.getRange("G6").setValue(new Date());
      }
    }
    // Remove sheets
    var listOfSheetsToRemove = ["Items","Events", "Pity Checker","Results","All Wish History", "Constellation"];

    var sheetAvailableSource = sheetSource.getSheetByName("Available");
    var availableRanges = sheetAvailableSource.getRange(2,1, sheetAvailableSource.getMaxRows()-1,1).getValues();
    availableRanges = String(availableRanges).split(",");
    
    // Go through the available sheet list
    for (var i = 0; i < availableRanges.length; i++) {
      listOfSheetsToRemove.push(availableRanges[i]);
    }
 
    var listOfSheetsToRemoveLength = listOfSheetsToRemove.length;
    for (var i = 0; i < listOfSheetsToRemoveLength; i++) {
      var sheetNameToRemove = listOfSheetsToRemove[i];
      var sheetToRemove = SpreadsheetApp.getActive().getSheetByName(sheetNameToRemove);
      if(sheetToRemove) {
        if (settingsSheet) {
          if (sheetNameToRemove == "Pity Checker") {
            savePityCheckerSettings(sheetToRemove, settingsSheet);
          } else if (sheetNameToRemove == "Results") {
            saveResultsSettings(sheetToRemove, settingsSheet);
          } else if (sheetNameToRemove == "Events") {
            saveEventsSettings(sheetToRemove, settingsSheet);
          }
        }

        // If exist remove from spreadsheet
        SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheetToRemove);
      }
    }

    var listOfSheets = ["Character Event Wish History","Permanent Wish History","Weapon Event Wish History","Novice Wish History"];
    var listOfSheetsLength = listOfSheets.length;
    // Check if sheet exist
    for (var i = 0; i < listOfSheetsLength; i++) {
      findWishHistoryByName(listOfSheets[i], sheetSource);
    }

    
    // Add Language
    var sheetItemSource;
    if (settingsSheet) {
      var languageFound = settingsSheet.getRange(2, 2).getValue();
      sheetItemSource = sheetSource.getSheetByName("Items"+"-"+languageFound);
    }
    if (sheetItemSource) {
      // Found language
    } else {
      // Default
      sheetItemSource = sheetSource.getSheetByName("Items");
    }
    sheetItemSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('Items');

    // Refresh spreadsheet
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

    var shouldShowSheet = true;
    if (settingsSheet) {
      if (settingsSheet.getRange("B14").getValue()) {
      } else {
        shouldShowSheet = false;
      }
    }
      
    if (shouldShowSheet) {
      var sheetEventsSource = sheetSource.getSheetByName('Events');
      var sheetEvents = sheetEventsSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('Events');
      
      if (settingsSheet) {
        restoreEventsSettings(sheetEvents, settingsSheet);
      }
      SpreadsheetApp.getActive().setActiveSheet(sheetEvents);
      SpreadsheetApp.getActive().moveActiveSheet(1);
    }

    var sheetPityCheckerSource = sheetSource.getSheetByName('Pity Checker');
    var sheetPityChecker = sheetPityCheckerSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('Pity Checker');

    SpreadsheetApp.getActive().setActiveSheet(sheetPityChecker);
    SpreadsheetApp.getActive().moveActiveSheet(1);

    if (settingsSheet) {
      restorePityCheckerSettings(sheetPityChecker, settingsSheet);
    }
    var sheetAllWishHistorySource = sheetSource.getSheetByName('All Wish History');
    var sheetAllWishHistory = sheetAllWishHistorySource.copyTo(SpreadsheetApp.getActiveSpreadsheet());
    sheetAllWishHistory.setName('All Wish History');
    sheetAllWishHistory.hideSheet();

    // Show Results
    shouldShowSheet = true;
    if (settingsSheet) {
      if (settingsSheet.getRange("B15").getValue()) {
      } else {
        shouldShowSheet = false;
      }
    }
      
    if (shouldShowSheet) {
      // Add Language
      var sheetResultsSource;
      if (settingsSheet) {
        var languageFound = settingsSheet.getRange(2, 2).getValue();
        sheetResultsSource = sheetSource.getSheetByName("Results"+"-"+languageFound);
      }
      if (sheetResultsSource) {
        // Found language
      } else {
        // Default
        sheetResultsSource = sheetSource.getSheetByName("Results");
      }
      var sheetResults = sheetResultsSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('Results');
      
      if (settingsSheet) {
        restoreResultsSettings(sheetResults, settingsSheet);
      }
      SpreadsheetApp.getActive().setActiveSheet(sheetResults);
      SpreadsheetApp.getActive().moveActiveSheet(1);
    }
    // Show Constellation
    shouldShowSheet = true;
    if (settingsSheet) {
      if (settingsSheet.getRange("B16").getValue()) {
      } else {
        shouldShowSheet = false;
      }
    }
    if (shouldShowSheet) {
      // Add Language
      var sheetConstellationSource;
      if (settingsSheet) {
        var languageFound = settingsSheet.getRange(2, 2).getValue();
        sheetConstellationSource = sheetSource.getSheetByName("Constellation"+"-"+languageFound);
      }
      if (sheetConstellationSource) {
        // Found language
      } else {
        // Default
        sheetConstellationSource = sheetSource.getSheetByName("Constellation");
      }
      var sheetConstellation = sheetConstellationSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('Constellation');
      // Refresh Contents Links
      var contentsAvailable = sheetConstellation.getRange(1, 1).getValue();
      var contentsStartIndex = 2;
      
      for (var i = 0; i < contentsAvailable; i++) {
        var valueRange = sheetConstellation.getRange(contentsStartIndex+i, 3).getValue();
        var formulaRange = sheetConstellation.getRange(contentsStartIndex+i, 3).getFormula();
        var textRange = formulaRange.split(",");
        var bookmarkRange = formulaRange.split("=");
        if (textRange.length > 1) {
          textRange = textRange[1].split(")")[0];
        }
        if (bookmarkRange.length > 2) {
          bookmarkRange = bookmarkRange[3].split('"')[0];
        }
        const richText = SpreadsheetApp.newRichTextValue()
            .setText(valueRange)
            .setLinkUrl(["#gid="+GET_SHEET_ID("Constellation")+'range='+bookmarkRange])
            .build();
        sheetConstellation.getRange(contentsStartIndex+i, 3).setRichTextValue(richText);
      }
      
      SpreadsheetApp.getActive().setActiveSheet(sheetConstellation);
      SpreadsheetApp.getActive().moveActiveSheet(1);
    }
    // Put available sheet into current
    var skipRanges = sheetAvailableSource.getRange(2,2, sheetAvailableSource.getMaxRows()-1,1).getValues();
    skipRanges = String(skipRanges).split(",");
    var hiddenRanges = sheetAvailableSource.getRange(2,3, sheetAvailableSource.getMaxRows()-1,1).getValues();
    hiddenRanges = String(hiddenRanges).split(",");
    var settingsOptionRanges = sheetAvailableSource.getRange(2,4, sheetAvailableSource.getMaxRows()-1,1).getValues();
    settingsOptionRanges = String(settingsOptionRanges).split(",");

    for (var i = 0; i < availableRanges.length; i++) {
      var nameOfBanner = availableRanges[i];
      var isSkipString = skipRanges[i];
      var isHiddenString = hiddenRanges[i];
      var settingOptionString = settingsOptionRanges[i];

      var sheetAvailableSelectionSource = sheetSource.getSheetByName(nameOfBanner);
      var storedSheet;
      if (isSkipString == "YES") {
        // skip - disabled by source
      } else {
        if (sheetAvailableSelectionSource) {
          if (settingOptionString == "" || settingOptionString == 0) {
            //Enable without settings
            storedSheet = sheetAvailableSelectionSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(nameOfBanner);
          } else {
            // Check current setting has row
            if (settingOptionString <= settingsSheet.getMaxRows()) {
              var checkEnabledRanges = settingsSheet.getRange(settingOptionString, 2).getValue();
              if (checkEnabledRanges == "YES") {
                storedSheet = sheetAvailableSelectionSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(nameOfBanner);
              } else {
                storedSheet = null;
              }
            } else {
              // Sheet does not have this settings available
              storedSheet = null;
            }
          }
          if (storedSheet) {
            if (isHiddenString == "YES") {
              storedSheet.hideSheet();
            } else {
              storedSheet.showSheet();
            }
          }
        }
      }
    }

    // Remove placeholder if available
    if(placeHolderSheet) {
      // If exist remove from spreadsheet
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(placeHolderSheet);
    }
    // Bring Pity Checker into view
    var sheetPityChecker = SpreadsheetApp.getActive().getSheetByName('Pity Checker');
    SpreadsheetApp.getActive().setActiveSheet(sheetPityChecker);
    
    // Update Settings
    settingsSheet.getRange(5, 7).setValue(false);
    settingsSheet.getRange("H6").setValue(new Date());
  } else {
    var message = 'Unable to connect to source';
    var title = 'Error';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function clearAHK() {
  var autoHotkeySheet = SpreadsheetApp.getActive().getSheetByName("AutoHotkey");
  if (autoHotkeySheet) {
    // Clear Select Banner and date and time
    autoHotkeySheet.getRange(1, 2, 1, 2).clearContent();
    var deleteRows = autoHotkeySheet.getMaxRows()-6;
    if (deleteRows > 0) {
      autoHotkeySheet.deleteRows(6,deleteRows); 
    }
    // Clear all rows
    autoHotkeySheet.getRange(4, 1, autoHotkeySheet.getMaxRows()-3, autoHotkeySheet.getMaxColumns()).clearContent();
  } else {
    var message = 'Enable AutoHotkey in settings, and run "Update Items"';
    var title = 'Error AHK Disabled';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function clearOverrideAHK() {
  var autoHotkeySheet = SpreadsheetApp.getActive().getSheetByName("AutoHotkey");
  if (autoHotkeySheet) {
    var clearRows = autoHotkeySheet.getMaxRows()-3;
    if (clearRows > 0) {
      autoHotkeySheet.getRange(4, 2, clearRows, 3).clearContent();
    }
  } else {
    var message = 'Enable AutoHotkey in settings, and run "Update Items"';
    var title = 'Error AHK Disabled';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function convertAHK() {
  var autoHotkeySheet = SpreadsheetApp.getActive().getSheetByName("AutoHotkey");
  if (autoHotkeySheet) {
    clearOverrideAHK();
    var banner = autoHotkeySheet.getRange(1, 2).getValue();
    var iLastRow = null;
    var lastWishDateAndTimeString = null;
    var lastWishDateAndTime = null;
    
    var bannerSheet = SpreadsheetApp.getActive().getSheetByName(banner);
    if (bannerSheet) {
      var iLastRow = bannerSheet.getRange(2, 5, bannerSheet.getLastRow(), 1).getValues().filter(String).length;
      if (iLastRow && iLastRow != 0 ) {
        iLastRow++;
        lastWishDateAndTimeString = bannerSheet.getRange("E" + iLastRow).getValue();
        if (lastWishDateAndTimeString) {
          autoHotkeySheet.getRange(1,3).setValue("Last wish: "+lastWishDateAndTimeString);
          lastWishDateAndTimeString = lastWishDateAndTimeString.split(" ").join("T");
          lastWishDateAndTime = new Date(lastWishDateAndTimeString+".000Z");
        } else {
          autoHotkeySheet.getRange(1,3).setValue("No previous wishes");
        }
      } else {
        autoHotkeySheet.getRange(1,3).setValue("");
      }
      
      // Ensure all the cells are text format
      autoHotkeySheet.getRange(4,1, autoHotkeySheet.getMaxRows()-3,1).setNumberFormat("@");
      var autoHotkeyRanges = autoHotkeySheet.getRange(4,1, autoHotkeySheet.getMaxRows()-3,1).getValues();
      autoHotkeyRanges = String(autoHotkeyRanges).split(",");
      
      var itemType;
      var itemName;
      var dateAndTime;
      var dateAndTimeString;
      var dateAndTimeStringMod;
      var nextDateAndTime;
      var nextDateAndTimeString;
      var nextDateAndTimeStringMod;
      var overrideCounter = 10;
      var groupIndex;
      var nextGroupIndex;
      var autoHotkeyRangesLength = autoHotkeyRanges.length/3;
      var isMulti = false;
      for(var i = 0; i < autoHotkeyRangesLength; i++) {
        groupIndex = i * 3;
        itemType = autoHotkeyRanges[groupIndex];
        itemName = autoHotkeyRanges[groupIndex+1];
        dateAndTimeString = autoHotkeyRanges[groupIndex+2];
        if (dateAndTimeString) {
          dateAndTimeStringMod = dateAndTimeString.split(" ").join("T");
          dateAndTime = new Date(dateAndTimeStringMod+".000Z");
        } else {
          dateAndTime = null;
        }
        if (overrideCounter == 1) {
          // Check previous
          nextGroupIndex = (i - 1) * 3;
        } else {
          // Check next
          nextGroupIndex = (i + 1) * 3;
        }
        if (nextGroupIndex < autoHotkeyRanges.length) {
          nextDateAndTimeString = autoHotkeyRanges[nextGroupIndex+2];
          if (nextDateAndTimeString) {
            nextDateAndTimeStringMod = nextDateAndTimeString.split(" ").join("T");
            nextDateAndTime = new Date(nextDateAndTimeStringMod+".000Z");
            
            if (nextDateAndTime.getTime() == dateAndTime.getTime()) {
              if (isMulti) {
                //Resume counting
              } else {
                isMulti = true;
                overrideCounter = 10;
              }
            } else {
              isMulti = false;
              // autoHotkeySheet.getRange(2 +i,3).setValue("nothing");
            }
          } else {
            isMulti = false;
          }
        } else {
          //autoHotkeySheet.getRange(2 +i,3).setValue(nextDateAndTime + ":"+ dateAndTime);
        }
        if (itemType && itemName && dateAndTime) {
          if (isMulti) {
            autoHotkeySheet.getRange(4 +i,4).setValue(overrideCounter);
            overrideCounter--;
            if (overrideCounter == 0) {
              // Switch off multi
              isMulti = false;
            }
          } else {
            autoHotkeySheet.getRange(4 +i,4).setValue("");
          }
          autoHotkeySheet.getRange(4 + groupIndex,2,3,1).mergeVertically();
          if (dateAndTime <= lastWishDateAndTime) {
            autoHotkeySheet.getRange(4 + groupIndex,2).setFormula('=CONCATENATE(CHAR(128503)," STOPPED date and time is older than banner")');
            break;
          } else {
            autoHotkeySheet.getRange(4 + i,3).setValue(itemType+itemName+dateAndTimeString);
            autoHotkeySheet.getRange(4 + groupIndex,2).setFormula('=CONCATENATE(CHAR(128505)," Row: '+ (4 + i) + '")');
          }
        } else {
          autoHotkeySheet.getRange(4 + groupIndex,2,3,1).mergeVertically();
          autoHotkeySheet.getRange(4 + groupIndex,2).setFormula('=CONCATENATE(CHAR(128503)," ")');
        }
      }
    } else {
      autoHotkeySheet.getRange(1,3).setValue("Select a valid banner");
    }
  } else {
    var message = 'Enable AutoHotkey in settings, and run "Update Items"';
    var title = 'Error AHK Disabled';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function importAHK() {
  var autoHotkeySheet = SpreadsheetApp.getActive().getSheetByName("AutoHotkey");
  if (autoHotkeySheet) {
    var banner = autoHotkeySheet.getRange(1, 2).getValue();
    var bannerSheet = SpreadsheetApp.getActive().getSheetByName(banner);
    if (bannerSheet) {
      var iLastRow = bannerSheet.getRange(2, 1, bannerSheet.getLastRow(), 1).getValues().filter(String).length;
      if (iLastRow != 0) {
        iLastRow = iLastRow + 1;
      } else {
        iLastRow = 2;
      }

      var iAHKLastRow = autoHotkeySheet.getRange(4, 3, autoHotkeySheet.getLastRow(), 1).getValues().filter(String).length;
      if (iAHKLastRow != 0) {
        //iAHKLastRow++;

        // Used to prevent lag when applying numberformat, must be done before entering data
        var wishHistoryNumberOfColumn = bannerSheet.getLastColumn();
        // Reduce two column due to paste and override
        var wishHistoryNumberOfColumnWithFormulas = wishHistoryNumberOfColumn - 2;

        var lastRowWithoutTitle = bannerSheet.getMaxRows() + iAHKLastRow;
        for (var i = 3; i <= wishHistoryNumberOfColumn; i++) {
          // Apply formatting for cells
          var numberFormatCell = bannerSheet.getRange(2, i).getNumberFormat();
          bannerSheet.getRange(2, i, lastRowWithoutTitle, 1).setNumberFormat(numberFormatCell);
        }

        // pasteValue to banner
        var pasteValue = autoHotkeySheet.getRange(4,3,iAHKLastRow, 2).getValues();
        bannerSheet.getRange(iLastRow,1,iAHKLastRow, 2).setValues(pasteValue);
        //bannerSheet.insertRowAfter(1);
        clearOverrideAHK();
        
       // addFormulaByWishHistoryName(banner); // lags the sheet
        sortWishHistoryByName(banner);
        var message = 'Imported to '+banner;
        var title = 'Complete';
        SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
      } else {
        var message = 'Nothing to import';
        var title = 'Error';
        SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
      }
    } else {
      var message = 'Select banner and run convert again';
      var title = 'Error';
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
    }
  } else {
    var message = 'Enable AutoHotkey in settings, and run "Update Items"';
    var title = 'Error AHK Disabled';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}
function generateAHK() {
  var autoHotkeyScriptSheet = SpreadsheetApp.getActive().getSheetByName("AutoHotkey-Script");
  if (autoHotkeyScriptSheet) {
    var autoHotkeyScriptRanges = autoHotkeyScriptSheet.getRange(7,2, autoHotkeyScriptSheet.getMaxRows()-6,1).getValues();
    autoHotkeyScriptRanges = String(autoHotkeyScriptRanges).split(",");
    
    var selectionString = autoHotkeyScriptSheet.getRange(4,1).getValue();
    if (selectionString) {
      var isFound = false;
      
      var scriptType;
      var SCRIPT_URL;
      
      for(var i = 0; i < autoHotkeyScriptRanges.length; i++) {
        scriptType = autoHotkeyScriptRanges[i];
        if (scriptType == selectionString) {
          SCRIPT_URL = autoHotkeyScriptSheet.getRange(7+i,3).getValue();
          isFound = true;
          break;
        }
      }
      if (isFound) {
        const richText = SpreadsheetApp.newRichTextValue()
          .setText(scriptType+" Link")
          .setLinkUrl([SCRIPT_URL])
          .build();
        autoHotkeyScriptSheet.getRange(7, 1).setRichTextValue(richText);
        var message = 'Script is ready';
        var title = 'Complete';
        SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
      } else {
        var message = 'Script Type selection not valid';
        var title = 'Error';
        SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
        autoHotkeyScriptSheet.getRange(7,1).setValue("ERROR");
      }
    } else {
      var message = 'Script Type selection is empty, check A4';
      var title = 'Error';
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
      autoHotkeyScriptSheet.getRange(7,1).setValue("ERROR");
    }
  } else {
    var message = 'Enable AutoHotkey in settings, and run "Update Items"';
    var title = 'Error AHK Disabled';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
    autoHotkeyScriptSheet.getRange(7,1).setValue("ERROR");
  }
}