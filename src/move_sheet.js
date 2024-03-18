/*
 * Version 4.00 made by yippym - 2023-02-57 23:00
 * https://github.com/Yippy/wish-tally-sheet
 */
function moveToSettingsSheet() {
  moveToSheetByName(WISH_TALLY_SETTINGS_SHEET_NAME);
}

function moveToDashboardSheet() {
  moveToSheetByName(WISH_TALLY_DASHBOARD_SHEET_NAME);
}

function moveToCharacterEventWishHistorySheet() {
  moveToSheetByName(WISH_TALLY_CHARACTER_EVENT_WISH_SHEET_NAME);
}

function moveToPermanentWishHistorySheet() {
  moveToSheetByName(WISH_TALLY_PERMANENT_WISH_SHEET_NAME);
}

function moveToWeaponEventWishHistorySheet() {
  moveToSheetByName(WISH_TALLY_WEAPON_EVENT_WISH_SHEET_NAME);
}

function moveToNoviceWishHistorySheet() {
  moveToSheetByName(WISH_TALLY_NOVICE_WISH_SHEET_NAME);
}

function moveToChronicledWishHistorySheet() {
  moveToSheetByName(WISH_TALLY_CHRONICLED_WISH_SHEET_NAME);
}

function moveToChangelogSheet() {
  moveToSheetByName(WISH_TALLY_CHANGELOG_SHEET_NAME);
}

function moveToPityCheckerSheet() {
  moveToSheetByName(WISH_TALLY_PITY_CHECKER_SHEET_NAME);
}

function moveToEventsSheet() {
  moveToSheetByName(WISH_TALLY_EVENTS_SHEET_NAME);
}

function moveToCharactersSheet() {
  moveToSheetByName(WISH_TALLY_CHARACTERS_SHEET_NAME);
}

function moveToWeaponsSheet() {
  moveToSheetByName(WISH_TALLY_WEAPONS_SHEET_NAME);
}

function moveToResultsSheet() {
  moveToSheetByName(WISH_TALLY_RESULTS_SHEET_NAME);
}

function moveToReadmeSheet() {
  moveToSheetByName(WISH_TALLY_README_SHEET_NAME);
}

function moveToCrystalCalculatorSheet() {
  moveToSheetByName(WISH_TALLY_CRYSTAL_CALCULATOR_SHEET_NAME);
}

function moveToSheetByName(nameOfSheet) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(nameOfSheet);
  if (sheet) {
    sheet.activate();
  } else {
    var settingsForOptionalSheet = SETTINGS_FOR_OPTIONAL_SHEET[nameOfSheet];
    if (settingsForOptionalSheet) {
      var settingsSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_SETTINGS_SHEET_NAME);
      if (settingsSheet) {
        var settingOption = settingsForOptionalSheet["setting_option"];
        if (!settingsSheet.getRange(settingOption).getValue()) {
          displayUserAlert("Optional Sheet", nameOfSheet+" has been disabled within Settings, enable this sheet at cell '"+settingOption+"', and run 'Update Items'",  SpreadsheetApp.getUi().ButtonSet.OK)
        }
      }
    }
    title = "Error";
    message = "Unable to find sheet named '"+nameOfSheet+"'.";
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}