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