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
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_SETTINGS_SHEET_NAME);
  removeTriggers();
  if (settingsSheet) {
    var isScheduleAvailable = false;
    
    // Update Items List trigger
    var updateItemsSchedule = settingsSheet.getRange("E20").getValue();
    
    var gotHour = SCHEDULER_RUN_AT_HOUR[updateItemsSchedule];
    if (gotHour || gotHour == 0) {
      isScheduleAvailable = true;
      createWeeklyTrigger("updateItemsListTriggered", SCHEDULER_RUN_AT_WHICH_DAY["Sunday"], gotHour, 1);
      settingsSheet.getRange("E21").setValue(updateItemsSchedule);
    } else {
      settingsSheet.getRange("E21").setValue(SCHEDULER_TRIGGER_OFF_TEXT);
    }

    // Sort trigger
    var sortRangeSchedule = settingsSheet.getRange("E27").getValue();
    var gotRunHourly = SCHEDULER_RUN_AT_EVERY_HOUR[sortRangeSchedule];
    if (gotRunHourly) {
      isScheduleAvailable = true;
      createHourlyTrigger("sortRangesTriggered", gotRunHourly, 0);
      settingsSheet.getRange("E28").setValue(sortRangeSchedule);
    } else {
      settingsSheet.getRange("E28").setValue(SCHEDULER_TRIGGER_OFF_TEXT);
    }

    // scheduledTrigger(1,00, "updateItemsListTriggered");
    if (isScheduleAvailable) {
      settingsSheet.getRange("E16").setValue(SCHEDULER_TRIGGER_ON_TEXT);
    } else {
      settingsSheet.getRange("E16").setValue(SCHEDULER_TRIGGER_OFF_TEXT);
    }
  }
}

function updateItemsListTriggered() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_SETTINGS_SHEET_NAME);
  settingsSheet.getRange("E22").setValue(new Date());
  updateItemsList();
  settingsSheet.getRange("E23").setValue(new Date());
}

function sortRangesTriggered() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_SETTINGS_SHEET_NAME);
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
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_SETTINGS_SHEET_NAME);
  removeTriggers();
  var title = "Remove Schedule";
  var message = "All schedule has been removed";
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  settingsSheet.getRange("E16").setValue(SCHEDULER_TRIGGER_OFF_TEXT);
  settingsSheet.getRange("E21").setValue(SCHEDULER_TRIGGER_OFF_TEXT);
  settingsSheet.getRange("E28").setValue(SCHEDULER_TRIGGER_OFF_TEXT);
}