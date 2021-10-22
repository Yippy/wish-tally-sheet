/*
 * Version 3.0.1 made by yippym - 2021-10-22 21:00
 * https://github.com/Yippy/wish-tally-sheet
 */

function onOpen( ){
  var ui = SpreadsheetApp.getUi();
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    ui.createMenu('Wish Tally')
    .addItem('Initialise', 'updateItemsList')
    .addToUi();
  } else {
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
              .addItem('Remove All Schedule', 'removeTriggerDataManagement')
              .addSeparator()
              .addItem('Auto Import', 'importFromAPI')
              )
    .addSeparator()
    .addItem('Quick Update', 'quickUpdate')
    .addItem('Update Items', 'updateItemsList')
    .addItem('Get Latest README', 'displayReadme')
    .addItem('About', 'displayAbout')
    .addToUi();
  }
}