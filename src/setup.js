/*
 * Version 3.51 made by yippym - 2022-12-09 21:00
 * https://github.com/Yippy/wish-tally-sheet
 */
function onInstall(e) {
  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    generateInitialiseToolbar();
  } else {
    onOpen(e);
  }
}

function onOpen(e) {
  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    generateInitialiseToolbar();
  } else {
    var settingsSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_SETTINGS_SHEET_NAME);
    if (!settingsSheet) {
      generateInitialiseToolbar();
    } else {
      getDefaultMenu();
    }
    checkLocaleIsSetCorrectly();
  }
}

function generateInitialiseToolbar() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Wish Tally')
  .addItem('Initialise', 'updateItemsList')
  .addToUi();
}

function displayUserPrompt(titlePrompt, messagePrompt, buttonSet) {
  const ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    titlePrompt,
    messagePrompt,
    buttonSet);
  return result;
}

function displayUserAlert(titleAlert, messageAlert, buttonSet) {
  const ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    titleAlert,
    messageAlert,
    buttonSet);
  return result;
}

/* Ensure Sheets is set to the supported locale due to source document formula */
function checkLocaleIsSetCorrectly() {
  var currentLocale = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale();
  if (currentLocale != WISH_TALLY_SHEET_SUPPORTED_LOCALE) {
    SpreadsheetApp.getActiveSpreadsheet().setSpreadsheetLocale(WISH_TALLY_SHEET_SUPPORTED_LOCALE);
    var message = 'To ensure compatibility with formula from source document, your locale "'+currentLocale+'" has been changed to "'+WISH_TALLY_SHEET_SUPPORTED_LOCALE+'"';
    var title = 'Sheets Locale Changed';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function getDefaultMenu() {
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
  .addSubMenu(ui.createMenu('HoYoLAB Sync')
            .addItem('Characters', 'importHoYoLabCharactersButtonScript')
            .addItem('Weapons', 'importHoYoLabWeaponsButtonScript')
            .addSeparator()
            .addItem('Tutorial', 'displayHoYoLab'))
  .addSeparator()
  .addSubMenu(ui.createMenu('Data Management')
            .addItem('Import', 'importDataManagement')
            .addSeparator()
            .addItem('Set Schedule', 'setTriggerDataManagement')
            .addItem('Remove All Schedule', 'removeTriggerDataManagement')
            .addSeparator()
            .addItem('Auto Import', 'importFromAPI')
            .addItem('Tutorial', 'displayAutoImport')
            )
  .addSeparator()
  .addItem('Quick Update', 'quickUpdate')
  .addItem('Update Items', 'updateItemsList')
  .addItem('Get Latest README', 'displayReadme')
  .addItem('About', 'displayAbout')
  .addToUi();
}

function getSettingsSheet() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_SETTINGS_SHEET_NAME);
  var sheetSource;
  if (!settingsSheet) {
    sheetSource = getSourceDocument();
    var sheetSettingSource = sheetSource.getSheetByName(WISH_TALLY_SETTINGS_SHEET_NAME);
    if (sheetSettingSource) {
      settingsSheet = sheetSettingSource.copyTo(SpreadsheetApp.getActiveSpreadsheet());
      settingsSheet.setName(WISH_TALLY_SETTINGS_SHEET_NAME);
      getDefaultMenu();
    }
  } else {
    settingsSheet.getRange("H1").setValue(WISH_TALLY_SHEET_SCRIPT_VERSION);
  }
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_DASHBOARD_SHEET_NAME);
  if (!dashboardSheet) {
    if (!sheetSource) {
      sheetSource = getSourceDocument();
    }
    var sheetDashboardSource = sheetSource.getSheetByName(WISH_TALLY_DASHBOARD_SHEET_NAME);
    if (sheetDashboardSource) {
      dashboardSheet = sheetDashboardSource.copyTo(SpreadsheetApp.getActiveSpreadsheet());
      dashboardSheet.setName(WISH_TALLY_DASHBOARD_SHEET_NAME);
    }
  } else {
    if (WISH_TALLY_SHEET_SCRIPT_IS_ADD_ON) {
      dashboardSheet.getRange(dashboardEditRange[8]).setFontColor("green").setFontWeight("bold").setHorizontalAlignment("left").setValue("Add-On Enabled");
    } else {
      dashboardSheet.getRange(dashboardEditRange[8]).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("left").setValue("Embedded Script");
    }
  }
  return settingsSheet;
}