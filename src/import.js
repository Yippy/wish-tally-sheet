/*
 * Version 3.0.1 made by yippym - 2021-10-22 21:00
 * https://github.com/Yippy/wish-tally-sheet
 */

function importButtonScript() {
  var settingsSheet = getSettingsSheet();
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_DASHBOARD_SHEET_NAME);
  if (!dashboardSheet || !settingsSheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Unable to find '" + WISH_TALLY_DASHBOARD_SHEET_NAME + "' or '" + WISH_TALLY_SETTINGS_SHEET_NAME + "'", "Missing Sheets");
    return;
  }

  var userImportSelection = dashboardSheet.getRange(dashboardEditRange[4]).getValue();
  var autoImportSelection = dashboardSheet.getRange(dashboardEditRange[5]).getValue();
  var importSelectionText = dashboardSheet.getRange(dashboardEditRange[6]).getValue();
  var importSelectionTextSubtitle = dashboardSheet.getRange(dashboardEditRange[7]).getValue();

  var urlInput = null;

  if (importSelectionText === autoImportSelection) {
    urlInput = getCachedAuthKeyInput();
    importSelectionTextSubtitle = "Please note Feedback URL no longer works for Auto Import\n[PC Only]\nDirectory (Double click below):\n%USERPROFILE%/AppData/LocalLow/miHoYo/Genshin Impact/\n\nCheck 'output_log.txt' for URL when visiting your Wish History in game";
  }

  if (urlInput === null) {
    const result = displayUserPrompt(importSelectionText, importSelectionTextSubtitle);
    var button = result.getSelectedButton();
    if (button !== SpreadsheetApp.getUi().Button.OK) {
      return;
    }
    urlInput = result.getResponseText();

    if (importSelectionText === autoImportSelection) {
      setCachedAuthKeyInput(urlInput);
    }
  }

  if (userImportSelection === importSelectionText) {
    settingsSheet.getRange("D6").setValue(urlInput);
    importDataManagement();
  } else {
    importFromAPI(urlInput);
  }
}

function importDataManagement() {
  var settingsSheet = getSettingsSheet();
  var userImportInput = settingsSheet.getRange("D6").getValue();
  var userImportStatus = settingsSheet.getRange("E7").getValue();
  var message = "";
  var title = "";
  var statusMessage = "";
  var rowOfStatusWishHistory = 9;
  if (userImportStatus == IMPORT_STATUS_COMPLETE) {
      title = "Error";
      message = "Already done, you only need to run once";
  } else {
    if (userImportInput) {
      // Attempt to load as URL
      var importSource;
      try {
        importSource = SpreadsheetApp.openByUrl(userImportInput);
      } catch(e) {
      }
      if (importSource) {
      } else {
        // Attempt to load as ID instead
        try {
          importSource = SpreadsheetApp.openById(userImportInput);
        } catch(e) {
        }
      }
      if (importSource) {
        // Go through the available sheet list
        for (var i = 0; i < WISH_TALLY_NAME_OF_WISH_HISTORY.length; i++) {
          var bannerImportSheet = importSource.getSheetByName(WISH_TALLY_NAME_OF_WISH_HISTORY[i]);

          if (bannerImportSheet) {
            var numberOfRows = bannerImportSheet.getMaxRows()-1;
            var range = bannerImportSheet.getRange(2, 1, numberOfRows, 2);

            if (numberOfRows > 0) {
              var bannerSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_NAME_OF_WISH_HISTORY[i]);

              if (bannerSheet) {
                bannerSheet.getRange(2, 1, numberOfRows, 2).setValues(range.getValues());
                settingsSheet.getRange(rowOfStatusWishHistory+i, 5).setValue(IMPORT_STATUS_WISH_HISTORY_COMPLETE);
              } else {
                settingsSheet.getRange(rowOfStatusWishHistory+i, 5).setValue(IMPORT_STATUS_WISH_HISTORY_NOT_FOUND);
              }
            } else {
              settingsSheet.getRange(rowOfStatusWishHistory+i, 5).setValue(IMPORT_STATUS_WISH_HISTORY_EMPTY);
            }
          } else {
            settingsSheet.getRange(rowOfStatusWishHistory+i, 5).setValue(IMPORT_STATUS_WISH_HISTORY_NOT_FOUND);
          }
        }
        var sourceSettingsSheet = importSource.getSheetByName(WISH_TALLY_SETTINGS_SHEET_NAME);
        if (sourceSettingsSheet) {
          var sourcePityCheckerSheet = importSource.getSheetByName(WISH_TALLY_PITY_CHECKER_SHEET_NAME);
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
          var pityCheckerSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_PITY_CHECKER_SHEET_NAME);
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
        var sourceEventsSheet = importSource.getSheetByName(WISH_TALLY_EVENTS_SHEET_NAME);
        if (sourceEventsSheet) {
          saveEventsSettings(sourceEventsSheet,settingsSheet);
          var eventsSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_EVENTS_SHEET_NAME);
          if (eventsSheet) {
            restoreEventsSettings(eventsSheet, settingsSheet)
          }
        }
        
        //Restore Results
        var sourceResultsSheet = importSource.getSheetByName(WISH_TALLY_RESULTS_SHEET_NAME);
        if (sourceResultsSheet) {
          saveResultsSettings(sourceResultsSheet, settingsSheet);
          var resultsSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_RESULTS_SHEET_NAME);
          if (resultsSheet) {
            restoreResultsSettings(resultsSheet, settingsSheet)
          }
        }
        
        //Restore Constellation old name
        var sourceConstellationSheet = importSource.getSheetByName(WISH_TALLY_CHARACTERS_OLD_SHEET_NAME);
        if (sourceConstellationSheet) {
          saveCollectionSettings(sourceConstellationSheet, settingsSheet,"G7","H7");
          // Restore save to the new Characters sheer
          var constellationSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_CHARACTERS_SHEET_NAME);
          if (constellationSheet) {
            restoreCollectionSettings(constellationSheet, settingsSheet,"G7","H7");
          }
        } else {
          // Restore new name Characters
          sourceConstellationSheet = importSource.getSheetByName(WISH_TALLY_CHARACTERS_SHEET_NAME);
          if (sourceConstellationSheet) {
            saveCollectionSettings(sourceConstellationSheet, settingsSheet,"G7","H7");
            var constellationSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_CHARACTERS_SHEET_NAME);
            if (constellationSheet) {
              restoreCollectionSettings(constellationSheet, settingsSheet,"G7","H7");
            }
          }
        }

        //Restore Weapons
        var sourceWeaponsSheet = importSource.getSheetByName(WISH_TALLY_WEAPONS_SHEET_NAME);
        if (sourceWeaponsSheet) {
          saveCollectionSettings(sourceWeaponsSheet, settingsSheet,"G8","H8");
          var weaponsSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_WEAPONS_SHEET_NAME);
          if (weaponsSheet) {
            restoreCollectionSettings(weaponsSheet, settingsSheet,"G8","H8");
          }
        }

        title = "Complete";
        message = "Imported all rows in column Paste Value and Override";
        statusMessage = IMPORT_STATUS_COMPLETE;
      } else {
        title = "Error";
        message = "Import From URL or Spreadsheet ID is invalid";
        statusMessage = IMPORT_STATUS_FAILED;
      }
    } else {
      title = "Error";
      message = "Import From URL or Spreadsheet ID is empty";
      statusMessage = IMPORT_STATUS_FAILED;
    }

    settingsSheet.getRange("E7").setValue(statusMessage);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
}