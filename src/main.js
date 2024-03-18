/*
 * Version 3.61 made by yippym - 2023-02-57 23:00
 * https://github.com/Yippy/wish-tally-sheet
 */
var dashboardEditRange = [
  "I5", // Status cell
  "AB28", // Document Version
  "AT40", // Current Document Version
  "T29", // Document Status
  "AV1", // Name of drop down 1 (import)
  "AV2", // Name of drop down 2 (auto import)
  "AG14", // Selection
  "AG16", // Subtitle of selection
  "A40" // Method of checking script enabled
];

// Cells that needs Pity Checker
var dashboardRefreshRange = [
  "G15", // Character 5-Star
  "G16", // Character 4-Star
  "K16", // Character Total
  "G20", // Permanent 5-Star
  "G21", // Permanent 4-Star
  "K21", // Permanent Total
  "G25", // Weapon 5-Star
  "G26", // Weapon 4-Star
  "K26", // Weapon Total
  "G30", // Chronicled 5-Star
  "G31", // Chronicled 4-Star
  "K31" // Chronicled Total
];

/**
 * Return the id of the sheet. (by name)
 *
 * @return The ID of the sheet
 * @customfunction
 */
function GET_SHEET_ID(sheetName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  var sheetId = "";
  if (sheet) {
    sheetId = sheet.getSheetId();
  }
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

function reorderSheets(settingsSheet) {
  if (settingsSheet) {
    var sheetsToSort = settingsSheet.getRange(28,2,15,1).getValues();

    for (var i = 0; i < sheetsToSort.length; i++) {
      var sheetName = sheetsToSort[i];
      if (sheetName != "") {
        var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
        if (sheet) {
          SpreadsheetApp.getActive().setActiveSheet(sheet);
          var position = i+1;
          if (position >= SpreadsheetApp.getActive().getNumSheets()) {
            position = SpreadsheetApp.getActive().getNumSheets();
          }
          SpreadsheetApp.getActive().moveActiveSheet(position);
        }
      }
    }
  }
}

function savePityCheckerSettings(pityCheckerSheet, settingsSheet) {
  var isDarkMode = pityCheckerSheet.getRange("AV1").getValue();
  var isShowTimer = pityCheckerSheet.getRange("F1").getValue();
  settingsSheet.getRange("G2").setValue(isDarkMode);
  settingsSheet.getRange("H2").setValue(isShowTimer);
}

function saveResultsSettings(resultsSheet, settingsSheet) {
  var saveRangesString = resultsSheet.getRange("C1").getValue();
  var userSaves = [];
  if (saveRangesString != "") {
    var saveRanges = String(saveRangesString).split(",");
    for (var i = 0; i < saveRanges.length; i++) {
      var saveRange = String(saveRanges[i]).split("=")
      if (saveRange.length == 2) {
        var getSave = resultsSheet.getRange(saveRange[1]).getValue();
        userSaves.push(saveRange[0]+"="+getSave);
      }
    }
  } else {
    // Old Results - Remove Dark Mode saving
    var selectionValueRange = resultsSheet.getRange(10,1,1,4).getValues();
    selectionValueRange = String(selectionValueRange).split(",");
    for (var i = 0; i < selectionValueRange.length; i++) {
      userSaves.push("ItemName"+(i+1)+"="+selectionValueRange[i]);
    }
  }
  settingsSheet.getRange("H3").setValue(userSaves.join(","));
}

function saveEventsSettings(eventsSheet, settingsSheet) {
  var eventsValueRange = eventsSheet.getRange(2,8, eventsSheet.getMaxRows()-1,1).getValues();
  eventsValueRange = String(eventsValueRange).split(",");
  var eventFormulaRanges = eventsSheet.getRange(2,8, eventsSheet.getMaxRows()-1,1).getFormulas();
  var saveData = [];
  for (var ii = 0; ii < eventFormulaRanges.length; ii++) {
    var formulaData = eventFormulaRanges[ii];
    if (formulaData == "") {
      var valueData = eventsValueRange[ii];
      if (valueData == "true") {
        saveData.push("TRUE");
      } else if (valueData == "false") {
        saveData.push("");
      } else {
        saveData.push(eventsValueRange[ii]);
      }
    } else {
      saveData.push("");
    }
  }
  settingsSheet.getRange("G4").setValue(saveData.join(","));
}

function restoreEventsSettings(sheetEvents, settingsSheet) {
  var saveData = settingsSheet.getRange("G4").getValue().split(",");
  var formulas = sheetEvents.getRange(2,8,saveData.length, 1).getFormulas();
  var saveInEvents = [];
  for (var ii = 0; ii < saveData.length; ii++) {
    var valueData = saveData[ii];
    if (valueData == "TRUE") {
      saveInEvents.push([true])
    } else if (valueData) {
      if (valueData != "") {
        saveInEvents.push([valueData])
      } else {
        saveInEvents.push([null])
      }
    } else {
      saveInEvents.push([null])
    }
  }
  // Override values if formulas exist on the cell.
  var data = saveInEvents.map(function(row, i) {
    return row.map(function(col, j) {
      return formulas[i][j] || col;
    });
  });
  sheetEvents.getRange(2,8,data.length,1).setValues(data);
}

function saveCollectionSettings(constellationsSheet, settingsSheet, itemsRange, optionsRange) {
  var maxColumns = constellationsSheet.getMaxColumns();

  var saveData = [];
  var columnValue = constellationsSheet.getRange(1, 2).getValue();

  if (columnValue > 0) {
    var startValue = constellationsSheet.getRange(1, columnValue).getValue();
    var nextValue = constellationsSheet.getRange(1, columnValue+1).getValue();
    var userInputColumnValue = constellationsSheet.getRange(1, columnValue+2).getValue();
    var saveRowsValue = constellationsSheet.getRange(1, columnValue+4).getValue();
    var startSaveRowValue = constellationsSheet.getRange(1, 11).getValue();
    if (startSaveRowValue == ""||  startSaveRowValue <= 0) {
      startSaveRowValue = 16;
    }
    for (var c = startValue; c <= maxColumns; c += nextValue) {
      var nameValue = constellationsSheet.getRange(1, c).getValue();
      if (nameValue != "") {
        var dataUserInput = nameValue;
        var saveValues = constellationsSheet.getRange(startSaveRowValue, c - userInputColumnValue,saveRowsValue,1).getValues();
        var isDataAvailableToSave = false;
        for (var s = 0; s < saveValues.length; s++) {
          var saveValue = saveValues[s];
          if (saveValue != null && saveValue != "" && saveValue != "0") {
            isDataAvailableToSave = true;
            break;
          }
        }
        if (isDataAvailableToSave == true) {
          saveData.push(dataUserInput+"="+saveValues.join("="));
        }
      }
    }
    if (saveData.length > 0) {
      settingsSheet.getRange(itemsRange).setValue(saveData.join(","));
    }
  }
  var contentValue = constellationsSheet.getRange(1, 1).getValue();
  if (contentValue > 0) {
    var lengthValue = constellationsSheet.getRange(contentValue+2, 1).getValue();
    if (lengthValue > 0) {
      saveData = constellationsSheet.getRange(contentValue+3, 1,lengthValue,1).getValues();
      settingsSheet.getRange(optionsRange).setValue(saveData.join(","));
    }
  }
}

function restoreCollectionSettings(constellationsSheet, settingsSheet, itemsRange, optionsRange) {
  var saveData = settingsSheet.getRange(itemsRange).getValue().split(",");
  var saveDict = [];
  var saveDictCounter = 0;
  for (var i = 0; i < saveData.length; i++) {
    var valuesSorting = saveData[i].split("=");
    if (valuesSorting.length > 2) {
      var nameData = valuesSorting[0];
      valuesSorting.splice(0, 1);
      saveDict[nameData] = valuesSorting;
      saveDictCounter++;
    }
  }
  var maxColumns = constellationsSheet.getMaxColumns();
  var columnValue = constellationsSheet.getRange(1, 2).getValue();

  if (columnValue > 0 && saveDictCounter > 0) {
    var startValue = constellationsSheet.getRange(1, columnValue).getValue();
    var nextValue = constellationsSheet.getRange(1, columnValue+1).getValue();
    var userInputColumnValue = constellationsSheet.getRange(1, columnValue+2).getValue();
    var saveRowsValue = constellationsSheet.getRange(1, columnValue+4).getValue();
    var startSaveRowValue = constellationsSheet.getRange(1, 11).getValue();
    if (startSaveRowValue == ""||  startSaveRowValue <= 0) {
      startSaveRowValue = 16;
    }
    for (var c = startValue; c <= maxColumns; c += nextValue) {
      if (saveDictCounter == 0) {
        // No more save data to restore.
        break;
      }
      var nameValue = constellationsSheet.getRange(1, c).getValue();
      if (nameValue != "") {
        var values = saveDict[nameValue];
        if (values) {
          var dataArray = [];
          for (var i = 0; i < values.length; i++) {
            dataArray.push([values[i]]);
          }
          constellationsSheet.getRange(startSaveRowValue, c - userInputColumnValue,dataArray.length,1).setValues(dataArray);
          saveDictCounter--;
        }
      }
    }
  }
  var contentValue = constellationsSheet.getRange(1, 1).getValue();
  if (contentValue > 0) {
    saveData = settingsSheet.getRange(optionsRange).getValue().split(",");
    var lengthValue = constellationsSheet.getRange(contentValue+2, 1).getValue();
    if (lengthValue > 0) {
      for (var i = 0; i < saveData.length; i++) {
        var isToggledOn = true;
        if (saveData[i]=="false") {
          isToggledOn = false;
        }
        if(constellationsSheet.getRange(contentValue+3+i, 1).getValue() != isToggledOn) {
          constellationsSheet.getRange(contentValue+3+i, 1).setValue(isToggledOn)
        }
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
    // Chronicled Wish History
    sheetPityChecker.hideColumns(38)
    sheetPityChecker.showColumns(40)
    // Novice Wish History
    sheetPityChecker.hideColumns(50)
    sheetPityChecker.showColumns(52)
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
    // Chronicled Wish History
    sheetPityChecker.showColumns(38)
    sheetPityChecker.hideColumns(40)
    // Novice Wish History
    sheetPityChecker.showColumns(50)
    sheetPityChecker.hideColumns(52)
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
    // Chronicled Wish History
    sheetPityChecker.hideColumns(44)
    sheetPityChecker.showColumns(46)
    // Novice Wish History
    sheetPityChecker.hideColumns(56)
    sheetPityChecker.showColumns(58)
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
    // Chronicled Wish History
    sheetPityChecker.showColumns(44)
    sheetPityChecker.hideColumns(46)
    // Novice Wish History
    sheetPityChecker.showColumns(56)
    sheetPityChecker.hideColumns(58)
  }
}

function restoreResultsSettings(sheetResults, settingsSheet) {
  var saveRangesString = sheetResults.getRange("C1").getValue();
  var saveRangesDict = [];
  if (saveRangesString != "") {
    var saveRanges = String(saveRangesString).split(",");
    for (var i = 0; i < saveRanges.length; i++) {
      var saveRange = String(saveRanges[i]).split("=")
      if (saveRange.length == 2) {
        saveRangesDict[saveRange[0]] = saveRange[1];
      }
    }
    var saveData = settingsSheet.getRange("H3").getValue().split(",");
    for (var i = 0; i < saveData.length; i++) {
      var valueData = String(saveData[i]).split("=");

      if (valueData.length == 2) {
        if (saveRangesDict.hasOwnProperty(valueData[0])) {
          sheetResults.getRange(saveRangesDict[valueData[0]]).setValue(valueData[1]);
        }
      }
    }
  }
}

var quickUpdateRange = [
  "A2","M2","Y2","AK2","AW2", // Banner Images
  "A3","M3","Y3","AK3" // Banner Time
];

function quickUpdate() {
  var settingsSheet = getSettingsSheet();
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_DASHBOARD_SHEET_NAME);
  if (dashboardSheet) {
    dashboardSheet.getRange(dashboardEditRange[0]).setValue("Quick Update: Running script, please wait.");
    dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("yellow").setFontWeight("bold");
  }
  if (dashboardSheet) {
    if (settingsSheet) {
      var isLoading = settingsSheet.getRange(9, 7).getValue();
      
      if (isLoading) {
        var counter = settingsSheet.getRange(9, 8).getValue();
        if (counter > 0) {
          counter++;
          settingsSheet.getRange(9, 8).setValue(counter);
        } else {
          settingsSheet.getRange(9, 8).setValue(1);
        }
        if (counter > 2) {
          // Bypass message - for people with broken update wanting force update
        } else {
          var message = 'Still updating';
          var title = 'Quick Update already started, the number of time you requested is '+counter+'. If you want to force an quick update due to an error happened during update, proceed in calling "Update Item" one more try.';
          SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
          return;
        }
      } else {
        settingsSheet.getRange(9, 7).setValue(true);
        settingsSheet.getRange(9, 8).setValue(1);
        settingsSheet.getRange("G10").setValue(new Date());
      }

      var changelogSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_CHANGELOG_SHEET_NAME);
      if (changelogSheet) {
        try {
          var sheetSource = getSourceDocument();
          if (sheetSource) {
            // get latest banners
            var sheetPityChecker = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_PITY_CHECKER_SHEET_NAME);
            if (sheetPityChecker) {
              restorePityCheckerSettings(sheetPityChecker, settingsSheet);
              if (sheetSource) {
                var sheetPityCheckerSource = sheetSource.getSheetByName(WISH_TALLY_PITY_CHECKER_SHEET_NAME);
                if (sheetPityCheckerSource) {
                  for (var i = 0; i < quickUpdateRange.length; i++) {
                    var range = quickUpdateRange[i];
                    var formula = sheetPityCheckerSource.getRange(range).getFormula();
                    if(formula) {
                      sheetPityChecker.getRange(range).setFormula(formula);
                    } else {
                      var value = sheetPityCheckerSource.getRange(range).getValue();
                      sheetPityChecker.getRange(range).setValue(value);
                    }
                  }
                }
              }
            }
            // check latest logs to see anything new
            if (dashboardSheet) {
              var sheetAvailableSource = sheetSource.getSheetByName(WISH_TALLY_AVAILABLE_SHEET_NAME);

              var sourceDocumentVersion = sheetAvailableSource.getRange("E1").getValues();
              var currentDocumentVersion = dashboardSheet.getRange(dashboardEditRange[2]).getValues();
              dashboardSheet.getRange(dashboardEditRange[1]).setValue(sourceDocumentVersion);
              if (sourceDocumentVersion>currentDocumentVersion){
                dashboardSheet.getRange(dashboardEditRange[3]).setValue("New Document Available, make a new copy");
                dashboardSheet.getRange(dashboardEditRange[3]).setFontColor("red").setFontWeight("bold");
              } else {
                dashboardSheet.getRange(dashboardEditRange[3]).setValue("Document is up-to-date");
                dashboardSheet.getRange(dashboardEditRange[3]).setFontColor("green").setFontWeight("bold");
              }

              var changesCheckRange = changelogSheet.getRange(2, 1).getValue();
              changesCheckRange = changesCheckRange.split(",");
              var lastDateChangeText;
              var lastDateChangeSourceText;
              var isChangelogTheSame = true;
              
              var sheetChangelogSource = sheetSource.getSheetByName(WISH_TALLY_CHANGELOG_SHEET_NAME);
              for (var i = 0; i < changesCheckRange.length; i++) {
                var checkChangelogSource = sheetChangelogSource.getRange(changesCheckRange[i]).getValue();
                if (checkChangelogSource instanceof Date) {
                  lastDateChangeSourceText = Utilities.formatDate(checkChangelogSource, sheetSource.getSpreadsheetTimeZone(), 'dd-MM-yyyy');
                }
                var checkChangelog = changelogSheet.getRange(changesCheckRange[i]).getValue();
                if (checkChangelog instanceof Date) {
                  lastDateChangeText = Utilities.formatDate(checkChangelog, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'dd-MM-yyyy');
                  if (lastDateChangeSourceText != lastDateChangeText) {
                    isChangelogTheSame = false;
                    break;
                  }
                } else {
                    if (checkChangelogSource != checkChangelog) {
                    isChangelogTheSame = false;
                    break;
                  }
                }
              }
              if (isChangelogTheSame) {
                dashboardSheet.getRange(dashboardEditRange[0]).setValue("Quick Update: There are no changes from source");
                dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("green").setFontWeight("bold");
              } else {
                if (lastDateChangeText == lastDateChangeSourceText) {
                  dashboardSheet.getRange(dashboardEditRange[0]).setValue("Quick Update: Current Changelog has the same date "+ lastDateChangeText+" but isn't the same notes to source. Please run 'Update Items'.");
                } else {
                  dashboardSheet.getRange(dashboardEditRange[0]).setValue("Quick Update: Current Changelog is "+lastDateChangeText+", source is at "+lastDateChangeSourceText+". Please run 'Update Items'.");
                }
                dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("red").setFontWeight("bold");
              }
            }
          } else {
            if (dashboardSheet) {
              dashboardSheet.getRange(dashboardEditRange[0]).setValue("Quick Update: Unable to connect to source, try again next time");
              dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("red").setFontWeight("bold");
            }
          }
        } catch(e) {
          if (dashboardSheet) {
            dashboardSheet.getRange(dashboardEditRange[0]).setValue("Quick Update: Unable to connect to source, try again next time.");
            dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("red").setFontWeight("bold");
          }
        }
      } else {
        if (dashboardSheet) {
          dashboardSheet.getRange(dashboardEditRange[0]).setValue("Quick Update: Missing 'Changelog' sheet in this Document, unable to compare to source. Please run 'Update Items'.");
          dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("red").setFontWeight("bold");
        }
      }
    } else {
      if (dashboardSheet) {
        dashboardSheet.getRange(dashboardEditRange[0]).setValue("Quick Update: Missing 'Settings' sheet in this Document, make a new copy as this Document has important sheet missing.");
        dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("red").setFontWeight("bold");
      }
    }
  }
  var currentSheet = SpreadsheetApp.getActive().getActiveSheet();
  reorderSheets(settingsSheet);
  SpreadsheetApp.getActive().setActiveSheet(currentSheet);
  // Update Settings
  settingsSheet.getRange(9, 7).setValue(false);
  settingsSheet.getRange("H10").setValue(new Date());
}

/**
* Update Item List
*/
function updateItemsList() {
  var settingsSheet = getSettingsSheet();
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_DASHBOARD_SHEET_NAME);
  var updateItemHasFailed = false;
  var errorMessage = "";
  if (dashboardSheet) {
    dashboardSheet.getRange(dashboardEditRange[0]).setValue("Update Items: Running script, please wait.");
    dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("yellow").setFontWeight("bold");
  }
  var sheetSource = getSourceDocument();
  // Check source is available
  if (sheetSource) {
    try {
      // attempt to load sheet from source, to prevent removing sheets first.
      var sheetAvailableSource = sheetSource.getSheetByName(WISH_TALLY_AVAILABLE_SHEET_NAME);
      // Avoid Exception: You can't remove all the sheets in a document.Details
      var placeHolderSheet = null;
      if (SpreadsheetApp.getActive().getSheets().length == 1) {
        placeHolderSheet = SpreadsheetApp.getActive().insertSheet();
      }
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
      var listOfSheetsToRemove = [WISH_TALLY_CHARACTERS_OLD_SHEET_NAME,WISH_TALLY_CHARACTERS_SHEET_NAME,WISH_TALLY_WEAPONS_SHEET_NAME,WISH_TALLY_EVENTS_SHEET_NAME,WISH_TALLY_RESULTS_SHEET_NAME,WISH_TALLY_PITY_CHECKER_SHEET_NAME,WISH_TALLY_ALL_WISH_HISTORY_SHEET_NAME,WISH_TALLY_ITEMS_SHEET_NAME];

      var availableRanges = sheetAvailableSource.getRange(2,1, sheetAvailableSource.getMaxRows()-1,5).getValues();

      if (dashboardSheet) {
        var sourceDocumentVersion = sheetAvailableSource.getRange("E1").getValues();
        var currentDocumentVersion = dashboardSheet.getRange(dashboardEditRange[2]).getValues();
        dashboardSheet.getRange(dashboardEditRange[1]).setValue(sourceDocumentVersion);
        if (sourceDocumentVersion>currentDocumentVersion){
          dashboardSheet.getRange(dashboardEditRange[3]).setValue("New Document Available, make a new copy");
          dashboardSheet.getRange(dashboardEditRange[3]).setFontColor("red").setFontWeight("bold");
        } else {
          dashboardSheet.getRange(dashboardEditRange[3]).setValue("Document is up-to-date");
          dashboardSheet.getRange(dashboardEditRange[3]).setFontColor("green").setFontWeight("bold");
        }
      }
      // Go through the available sheet list
      for (var i = 0; i < availableRanges.length; i++) {
        var nameOfBanner = availableRanges[i][0];
        listOfSheetsToRemove.push(nameOfBanner);
      }

      var listOfSheetsToRemoveLength = listOfSheetsToRemove.length;
      for (var i = 0; i < listOfSheetsToRemoveLength; i++) {
        var sheetNameToRemove = listOfSheetsToRemove[i];
        var sheetToRemove = SpreadsheetApp.getActive().getSheetByName(sheetNameToRemove);
        if(sheetToRemove) {
          if (settingsSheet) {
            if (sheetNameToRemove == WISH_TALLY_PITY_CHECKER_SHEET_NAME) {
              savePityCheckerSettings(sheetToRemove, settingsSheet);
            } else if (sheetNameToRemove == WISH_TALLY_RESULTS_SHEET_NAME) {
              saveResultsSettings(sheetToRemove, settingsSheet);
            } else if (sheetNameToRemove == WISH_TALLY_EVENTS_SHEET_NAME) {
              saveEventsSettings(sheetToRemove, settingsSheet);
            } else if (sheetNameToRemove == WISH_TALLY_CHARACTERS_OLD_SHEET_NAME) {
              saveCollectionSettings(sheetToRemove, settingsSheet,"G7","H7");
            } else if (sheetNameToRemove == WISH_TALLY_CHARACTERS_SHEET_NAME) {
              saveCollectionSettings(sheetToRemove, settingsSheet,"G7","H7");
            } else if (sheetNameToRemove == WISH_TALLY_WEAPONS_SHEET_NAME) {
              saveCollectionSettings(sheetToRemove, settingsSheet,"G8","H8");
            }
          }
          
          // If exist remove from spreadsheet
          SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheetToRemove);
        }
      }
      
      var listOfSheets = [
        WISH_TALLY_CHARACTER_EVENT_WISH_SHEET_NAME,
        WISH_TALLY_PERMANENT_WISH_SHEET_NAME,
        WISH_TALLY_WEAPON_EVENT_WISH_SHEET_NAME,
        WISH_TALLY_NOVICE_WISH_SHEET_NAME,
        WISH_TALLY_CHRONICLED_WISH_SHEET_NAME
      ];
      var listOfSheetsLength = listOfSheets.length;
      // Check if sheet exist
      for (var i = 0; i < listOfSheetsLength; i++) {
        findWishHistoryByName(listOfSheets[i], sheetSource);
      }

      // Add Language
      var sheetItemSource;
      if (settingsSheet) {
        var languageFound = settingsSheet.getRange(2, 2).getValue();
        sheetItemSource = sheetSource.getSheetByName(WISH_TALLY_ITEMS_SHEET_NAME+"-"+languageFound);
      }
      if (sheetItemSource) {
        // Found language
      } else {
        // Default
        sheetItemSource = sheetSource.getSheetByName(WISH_TALLY_ITEMS_SHEET_NAME);
      }
      sheetItemSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(WISH_TALLY_ITEMS_SHEET_NAME);

      var sheetAllWishHistorySource = sheetSource.getSheetByName(WISH_TALLY_ALL_WISH_HISTORY_SHEET_NAME);
      var sheetAllWishHistory = sheetAllWishHistorySource.copyTo(SpreadsheetApp.getActiveSpreadsheet());
      sheetAllWishHistory.setName(WISH_TALLY_ALL_WISH_HISTORY_SHEET_NAME);
      sheetAllWishHistory.hideSheet();

      // Refresh spreadsheet
      for (var i = 0; i < listOfSheetsLength; i++) {
        addFormulaByWishHistoryName(listOfSheets[i], settingsSheet);
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
        if (settingsSheet.getRange(SETTINGS_FOR_OPTIONAL_SHEET[WISH_TALLY_EVENTS_SHEET_NAME]["setting_option"]).getValue()) {
        } else {
          shouldShowSheet = false;
        }
      }
      var sheetEvents;
      if (shouldShowSheet) {
        var sheetEventsSource = sheetSource.getSheetByName(WISH_TALLY_EVENTS_SHEET_NAME);
        sheetEvents = sheetEventsSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(WISH_TALLY_EVENTS_SHEET_NAME);
        
        if (settingsSheet) {
          restoreEventsSettings(sheetEvents, settingsSheet);
        }
      }

      var sheetPityCheckerSource = sheetSource.getSheetByName(WISH_TALLY_PITY_CHECKER_SHEET_NAME);
      var sheetPityChecker = sheetPityCheckerSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(WISH_TALLY_PITY_CHECKER_SHEET_NAME);
      
      
      if (settingsSheet) {
        restorePityCheckerSettings(sheetPityChecker, settingsSheet);
      }
      
      if (dashboardSheet) {
        for (var i = 0; i < dashboardRefreshRange.length; i++) {
          var tempFormula = dashboardSheet.getRange(dashboardRefreshRange[i]).getFormula();
          // Re set formula
          dashboardSheet.getRange(dashboardRefreshRange[i]).setFormula(tempFormula);
        }
      }

      // Show Results
      shouldShowSheet = true;
      if (settingsSheet) {
        if (settingsSheet.getRange(SETTINGS_FOR_OPTIONAL_SHEET[WISH_TALLY_RESULTS_SHEET_NAME]["setting_option"]).getValue()) {
        } else {
          shouldShowSheet = false;
        }
      }
      var sheetResults;
      if (shouldShowSheet) {
        // Add Language
        var sheetResultsSource;
        if (settingsSheet) {
          var languageFound = settingsSheet.getRange(2, 2).getValue();
          sheetResultsSource = sheetSource.getSheetByName(WISH_TALLY_RESULTS_SHEET_NAME+"-"+languageFound);
        }
        if (sheetResultsSource) {
          // Found language
        } else {
          // Default
          sheetResultsSource = sheetSource.getSheetByName(WISH_TALLY_RESULTS_SHEET_NAME);
        }
        sheetResults = sheetResultsSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(WISH_TALLY_RESULTS_SHEET_NAME);
        
        if (settingsSheet) {
          restoreResultsSettings(sheetResults, settingsSheet);
        }
      }
      // Show Constellation
      shouldShowSheet = true;
      if (settingsSheet) {
        if (settingsSheet.getRange(SETTINGS_FOR_OPTIONAL_SHEET[WISH_TALLY_CHARACTERS_SHEET_NAME]["setting_option"]).getValue()) {
        } else {
          shouldShowSheet = false;
        }
      }
      var sheetConstellation;
      if (shouldShowSheet) {
        // Add Language
        var sheetConstellationSource;
        if (settingsSheet) {
          var languageFound = settingsSheet.getRange(2, 2).getValue();
          sheetConstellationSource = sheetSource.getSheetByName(WISH_TALLY_CHARACTERS_SHEET_NAME+"-"+languageFound);
        }
        if (sheetConstellationSource) {
          // Found language
        } else {
          // Default
          sheetConstellationSource = sheetSource.getSheetByName(WISH_TALLY_CHARACTERS_SHEET_NAME);
        }
        if (sheetConstellationSource) {
          sheetConstellation = sheetConstellationSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(WISH_TALLY_CHARACTERS_SHEET_NAME);

          // Refresh Contents Links
          var contentsAvailable = sheetConstellation.getRange(1, 1).getValue();
          var hyperlinkColumn = sheetConstellation.getRange(1, 3).getValue();
          var contentsStartIndex = 2;
          generateRichTextLinks(sheetConstellation, contentsAvailable, contentsStartIndex, hyperlinkColumn, true);

          if (settingsSheet) {
            restoreCollectionSettings(sheetConstellation, settingsSheet,"G7","H7");
          }
          
        }
      }
      // Show Weapons
      shouldShowSheet = true;
      if (settingsSheet) {
        if (settingsSheet.getRange(SETTINGS_FOR_OPTIONAL_SHEET[WISH_TALLY_WEAPONS_SHEET_NAME]["setting_option"]).getValue()) {
        } else {
          shouldShowSheet = false;
        }
      }
      var sheetWeapons;
      if (shouldShowSheet) {
        // Add Language
        var sheetWeaponsSource;
        if (settingsSheet) {
          var languageFound = settingsSheet.getRange(2, 2).getValue();
          sheetWeaponsSource = sheetSource.getSheetByName(WISH_TALLY_WEAPONS_SHEET_NAME+"-"+languageFound);
        }
        if (sheetWeaponsSource) {
          // Found language
        } else {
          // Default
          sheetWeaponsSource = sheetSource.getSheetByName(WISH_TALLY_WEAPONS_SHEET_NAME);
        }
        if (sheetWeaponsSource) {
          sheetWeapons = sheetWeaponsSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(WISH_TALLY_WEAPONS_SHEET_NAME);

          // Refresh Contents Links
          var contentsAvailable = sheetWeapons.getRange(1, 1).getValue();
          var hyperlinkColumn = sheetWeapons.getRange(1, 3).getValue();
          var contentsStartIndex = 2;
          generateRichTextLinks(sheetWeapons, contentsAvailable, contentsStartIndex, hyperlinkColumn, true);

          if (settingsSheet) {
            restoreCollectionSettings(sheetWeapons, settingsSheet,"G8","H8");
          }
          
        }
      }
      // Put available sheet into current
      for (var i = 0; i < availableRanges.length; i++) {
        var nameOfBanner = availableRanges[i][0];
        var isSkipString = availableRanges[i][1];
        var isHiddenString = availableRanges[i][2];
        var settingRowOptionString = availableRanges[i][3];
        var settingColumnOptionString = availableRanges[i][4];

        var sheetAvailableSelectionSource = sheetSource.getSheetByName(nameOfBanner);
        var storedSheet;
        if (isSkipString == "YES") {
          // skip - disabled by source
        } else {
          if (sheetAvailableSelectionSource) {
            if (settingRowOptionString == "" || settingRowOptionString == 0) {
              //Enable without settings
              storedSheet = sheetAvailableSelectionSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(nameOfBanner);
            } else {
              // Check current setting has row
              if (settingRowOptionString <= settingsSheet.getMaxRows()) {
                // For backwards compatibility pre v3.40 script, has only row number in column 'D'.
                var checkEnabledRanges = settingsSheet.getRange(settingRowOptionString, settingColumnOptionString).getValue();
                if (checkEnabledRanges == "YES" || checkEnabledRanges == "TRUE") {
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

      reorderSheets(settingsSheet);
      // Bring Pity Checker into view

      SpreadsheetApp.getActive().setActiveSheet(sheetPityChecker);
      
      // Update Settings
      settingsSheet.getRange(5, 7).setValue(false);
      settingsSheet.getRange("H6").setValue(new Date());
      
    } catch(e) {
      var message = 'Unable to complete "Update Items" due to '+e;
      var title = 'Error';
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
      updateItemHasFailed = true;
      errorMessage = "\nError: "+e;
      settingsSheet.getRange(5, 7).setValue(false);
      settingsSheet.getRange("H6").setValue(new Date());
    }
  } else {
    var message = 'Unable to connect to source';
    var title = 'Error';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
    updateItemHasFailed = true;
    errorMessage = "\n"+message;
    settingsSheet.getRange(5, 7).setValue(false);
    settingsSheet.getRange("H6").setValue(new Date());
  }

  if (dashboardSheet) {
    if (updateItemHasFailed) {
      dashboardSheet.getRange(dashboardEditRange[0]).setValue("Update Items: Update Items has failed, please try again."+errorMessage);
      dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("red").setFontWeight("bold");
    } else {
      dashboardSheet.getRange(dashboardEditRange[0]).setValue("Update Items: Successfully updated the Item list.");
      dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("green").setFontWeight("bold");
    }
  }
}