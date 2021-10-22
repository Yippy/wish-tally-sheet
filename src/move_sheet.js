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

function moveToSheetByName(nameOfSheet) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(nameOfSheet);
  if (sheet) {
    sheet.activate();
  } else {
    title = "Error";
    message = "Unable to find sheet named '"+nameOfSheet+"'.";
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}