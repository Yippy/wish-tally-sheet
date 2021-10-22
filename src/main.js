/*
* Version 2.93 made by yippym - 2021-05-12 16:14
 */

/* Add URL here to avoid showing on Sheet */
var urlForAPIByPass = "";
/* (optional) */

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

var bannerSettingsForImport = {
  "Character Event Wish History": {"range_status":"E44","range_toggle":"E37", "gacha_type":301},
  "Permanent Wish History": {"range_status":"E45","range_toggle":"E38", "gacha_type":200},
  "Weapon Event Wish History": {"range_status":"E46","range_toggle":"E39", "gacha_type":302},
  "Novice Wish History": {"range_status":"E47","range_toggle":"E40", "gacha_type":100},
};

var languageSettingsForImport = {
  "English": {"code": "en","full_code":"en-us","4_star":" (4-Star)","5_star":" (5-Star)"},
  "German": {"code": "de","full_code":"de-de","4_star":" (4 Sterne)","5_star":" (5 Sterne)"},
  "French": {"code": "fr","full_code":"fr-fr","4_star":" (4★)","5_star":" (5★)"},
  "Spanish": {"code": "es","full_code":"es-es","4_star":" (4★)","5_star":" (5★)"},
  "Chinese Traditional": {"code": "zh-tw","full_code":"zh-tw","4_star":" (四星)","5_star":" (五星)"},
  "Chinese Simplified": {"code": "zh-cn","full_code":"zh-cn","4_star":" (四星)","5_star":" (五星)"},
  "Indonesian": {"code": "id","full_code":"id-id","4_star":" (4★)","5_star":" (5★)"},
  "Japanese": {"code": "ja","full_code":"ja-jp","4_star":" (★4)","5_star":" (★5)"},
  "Vietnamese": {"code": "vi","full_code":"vi-vn","4_star":" (4 sao)","5_star":" (5 sao)"},
  "Korean": {"code": "ko","full_code":"ko-kr","4_star":" (★4)","5_star":" (★5)"},
  "Portuguese": {"code": "pt","full_code":"pt-pt","4_star":" (4★)","5_star":" (5★)"},
  "Thai": {"code": "th","full_code":"th-th","4_star":" (4 ดาว)","5_star":" (5 ดาว)"},
  "Russian": {"code": "ru","full_code":"ru-ru","4_star":" (4★)","5_star":" (5★)"}
};

var additionalQuery = [
  "authkey_ver=1",
  "sign_type=2",
  "auth_appid=webview_gacha",
  "device_type=pc"
];

var url = "https://hk4e-api-os.mihoyo.com/event/gacha_info/api/getGachaLog";
var urlChina = "https://hk4e-api.mihoyo.com/event/gacha_info/api/getGachaLog";

var errorCodeAuthTimeout = -101;
var errorCodeAuthInvalid = -100;
var errorCodeLanguageCode = -108;
var errorCodeRequestParams = -104;
var errorCodeNotEncountered = true;

function importFromAPI() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName('Settings');
  settingsSheet.getRange("E42").setValue(new Date());
  settingsSheet.getRange("E43").setValue("");

  var urlForAPI = settingsSheet.getRange("D35").getValue();
  if (urlForAPIByPass != "") {
    urlForAPI = urlForAPIByPass;
  }
  urlForAPI = urlForAPI.toString().split("&");
  var foundAuth = "";
  for (var i = 0; i < urlForAPI.length; i++) {
    var queryString = urlForAPI[i].toString().split("=");
    if (queryString.length == 2) {
      if (queryString[0] == "authkey") {
        foundAuth = queryString[1];
        break;
      }
    }
  }
  var bannerName;
  var bannerSheet;
  var bannerSettings;
  if (foundAuth == "") {
    // Display auth key not available
    for (var i = 0; i < nameOfWishHistorys.length; i++) {
      bannerName = nameOfWishHistorys[i];
      bannerSettings = bannerSettingsForImport[bannerName];
      settingsSheet.getRange(bannerSettings['range_status']).setValue("No auth key");
    }
  } else {
    var selectedLanguageCode = settingsSheet.getRange("B2").getValue();
    var selectedServer = settingsSheet.getRange("B3").getValue();
    var languageSettings = languageSettingsForImport[selectedLanguageCode];
    if (languageSettings == null) {
      // Get default language
      languageSettings = languageSettingsForImport["English"];
    }
    var urlForWishHistory;
    if (selectedServer == "China") {
      urlForWishHistory = urlChina;
    } else {
      urlForWishHistory = url;
    }
    urlForWishHistory += "?"+additionalQuery.join("&")+"&authkey="+foundAuth+"&lang="+languageSettings['code'];
    errorCodeNotEncountered = true;
    // Clear status
    for (var i = 0; i < nameOfWishHistorys.length; i++) {
      bannerName = nameOfWishHistorys[i];
      bannerSettings = bannerSettingsForImport[bannerName];
      settingsSheet.getRange(bannerSettings['range_status']).setValue("");
    }
    for (var i = 0; i < nameOfWishHistorys.length; i++) {
      if (errorCodeNotEncountered) {
        bannerName = nameOfWishHistorys[i];
        bannerSettings = bannerSettingsForImport[bannerName];
        var isToggled = settingsSheet.getRange(bannerSettings['range_toggle']).getValue();
        if (isToggled == true) {
          bannerSheet = SpreadsheetApp.getActive().getSheetByName(bannerName);
          if (bannerSheet) {
            checkPages(urlForWishHistory, bannerSheet, bannerName, bannerSettings, languageSettings, settingsSheet);
          } else {
            settingsSheet.getRange(bannerSettings['range_status']).setValue("Missing sheet");
          }
        } else {
          settingsSheet.getRange(bannerSettings['range_status']).setValue("Skipped");
        }
      } else {
        settingsSheet.getRange(bannerSettings['range_status']).setValue("Skipped - Error");
      }
    }
  }
  settingsSheet.getRange("E43").setValue(new Date());
  if (errorCodeNotEncountered) {
    settingsSheet.getRange("D35").setValue("");
  }
}

function checkPages(urlForWishHistory, bannerSheet, bannerName, bannerSettings, languageSettings, settingsSheet) {
  settingsSheet.getRange(bannerSettings['range_status']).setValue("Starting");
  /* Get latest wish from banner */
  var iLastRow = bannerSheet.getRange(2, 5, bannerSheet.getLastRow(), 1).getValues().filter(String).length;
  var lastWishDateAndTimeString;
  var lastWishDateAndTime;
  if (iLastRow && iLastRow != 0 ) {
    iLastRow++;
    lastWishDateAndTimeString = bannerSheet.getRange("E" + iLastRow).getValue();
    if (lastWishDateAndTimeString) {
      settingsSheet.getRange(bannerSettings['range_status']).setValue("Last wish: "+lastWishDateAndTimeString);
      lastWishDateAndTimeString = lastWishDateAndTimeString.split(" ").join("T");
      lastWishDateAndTime = new Date(lastWishDateAndTimeString+".000Z");
    } else {
      iLastRow = 1;
      settingsSheet.getRange(bannerSettings['range_status']).setValue("No previous wishes");
    }
    iLastRow++; // Move last row to new row
  } else {
    iLastRow = 2; // Move last row to new row
    settingsSheet.getRange(bannerSettings['range_status']).setValue("");
  }
  
  var extractWishes = [];
  var page = 1;
  var queryBannerCode = bannerSettings["gacha_type"];
  var numberOfWishPerPage = 6;
  var urlForBanner = urlForWishHistory+"&gacha_type="+queryBannerCode+"&size="+numberOfWishPerPage;
  var failed = 0;
  var is_done = false;
  var end_id = 0;
  
  var checkPreviousDateAndTimeString = "";
  var overrideIndex = 0;
  while (!is_done) {
    settingsSheet.getRange(bannerSettings['range_status']).setValue("Loading page: "+page);
    var response = UrlFetchApp.fetch(urlForBanner+"&page="+page+"&end_id="+end_id);
    var jsonResponse = response.getContentText();
    var jsonDict = JSON.parse(jsonResponse);
    var jsonDictData = jsonDict["data"];
    if (jsonDictData) {
      var listOfWishes = jsonDictData["list"];
      var isDone = false;
      var listOfWishesLength = listOfWishes.length;
      var wish;
      if (listOfWishesLength > 0) {
        for (var i = 0; i < listOfWishesLength; i++) {
          wish = listOfWishes[i];
          var dateAndTimeString = wish['time'];
          var textWish = wish['item_type']+wish['name'];
          /* Mimic the website in showing specific language wording */
          if (wish['rank_type'] == 4) {
            textWish += languageSettings["4_star"];
          } else if (wish['rank_type'] == 5) {
            textWish += languageSettings["5_star"];
          }
          textWish += dateAndTimeString;

          var dateAndTimeStringModified = dateAndTimeString.split(" ").join("T");
          var wishDateAndTime = new Date(dateAndTimeStringModified+".000Z");
          if (checkPreviousDateAndTimeString === dateAndTimeString) {
            if (overrideIndex == 0) {
              var previousWishIndex = extractWishes.length - 1;
              var previousWish = extractWishes[previousWishIndex];
              overrideIndex = 10;
              previousWish[1] = overrideIndex;
              extractWishes[previousWishIndex] = previousWish;
            }
            overrideIndex--;
          } else {
            checkPreviousDateAndTimeString = dateAndTimeString;
            overrideIndex = 0;
          }
          if (lastWishDateAndTime >= wishDateAndTime ) {
            // Banner already got this wish
            is_done = true;
            break;
          } else {
            extractWishes.push([textWish, (overrideIndex > 0 ? overrideIndex:null)]);
          }
        }
        if (numberOfWishPerPage == listOfWishesLength) {
          // There could be more wishes on the next page
          end_id = wish['id'];
          page++;
        } else {
          // If list isn't the size requested, it would mean there is no more wishes.
          is_done = true;
        }
      } else {
        is_done = true;
      }
    } else {
      var message = jsonDict["message"];
      var title ="Error code: "+jsonDict["retcode"];
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
      if (errorCodeAuthTimeout == jsonDict["retcode"]) {
        errorCodeNotEncountered = false;
        is_done = true;
        settingsSheet.getRange(bannerSettings['range_status']).setValue("auth timeout");
      } else if (errorCodeAuthInvalid == jsonDict["retcode"]) {
        errorCodeNotEncountered = false;
        is_done = true;
        settingsSheet.getRange(bannerSettings['range_status']).setValue("auth invalid");
      } else if (errorCodeRequestParams == jsonDict["retcode"]) {
        errorCodeNotEncountered = false;
        is_done = true;
        settingsSheet.getRange(bannerSettings['range_status']).setValue("Change server setting");
      }

      failed++;
      if (failed > 2){
        is_done = true;
      }
    }
  }
  if (failed > 2){
    settingsSheet.getRange(bannerSettings['range_status']).setValue("Failed too many times");
  } else {
    if (errorCodeNotEncountered) {
      if (extractWishes.length > 0) {
        settingsSheet.getRange(bannerSettings['range_status']).setValue("Found: "+extractWishes.length);
        extractWishes.reverse();
        bannerSheet.getRange(iLastRow, 1, extractWishes.length, 2).setValues(extractWishes);
      } else {
        settingsSheet.getRange(bannerSettings['range_status']).setValue("Nothing to add");
      }
    }
  }
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
        
        //Restore Constellation
        var sourceConstellationSheet = importSource.getSheetByName("Constellation");
        if (sourceConstellationSheet) {
          saveCollectionSettings(sourceConstellationSheet, settingsSheet,"G7","H7");
          var constellationSheet = SpreadsheetApp.getActive().getSheetByName('Constellation');
          if (constellationSheet) {
            restoreCollectionSettings(constellationSheet, settingsSheet,"G7","H7");
          }
        }

        //Restore Weapons
        var sourceWeaponsSheet = importSource.getSheetByName("Weapons");
        if (sourceWeaponsSheet) {
          saveCollectionSettings(sourceWeaponsSheet, settingsSheet,"G8","H8");
          var weaponsSheet = SpreadsheetApp.getActive().getSheetByName('Weapons');
          if (weaponsSheet) {
            restoreCollectionSettings(weaponsSheet, settingsSheet,"G8","H8");
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
          .setLinkUrl(["#gid="+sheetREADME.getSheetId()+'range='+text])
          .build();
        sheetREADME.getRange(contentsStartIndex+i, 1).setRichTextValue(richText);
      }
    }
    reorderSheets();
    SpreadsheetApp.getActive().setActiveSheet(sheetREADME);
  } else {
    var message = 'Unable to connect to source';
    var title = 'Error';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function reorderSheets() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName("Settings");
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
  for (var ii = 0; ii < saveData.length; ii++) {
    var valueData = saveData[ii];
    if (valueData == "TRUE") {
      sheetEvents.getRange(2 + ii,8).setValue(true);
    } else if (valueData) {
      if (valueData != "") {
        sheetEvents.getRange(2 + ii,8).setValue(valueData);
      }
    }
  }
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
    for (var c = startValue; c <= maxColumns; c += nextValue) {
      var nameValue = constellationsSheet.getRange(1, c).getValue();
      if (nameValue != "") {
        var dataUserInput = nameValue;
        var saveValues = constellationsSheet.getRange(16, c - userInputColumnValue,saveRowsValue,1).getValues();
        saveData.push(dataUserInput+"="+saveValues.join("="));
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
  for (var i = 0; i < saveData.length; i++) {
    var valuesSorting = saveData[i].split("=");
    if (valuesSorting.length > 2) {
      var nameData = valuesSorting[0];
      valuesSorting.splice(0, 1);
      saveDict[nameData] = valuesSorting;
    }
  }
  var maxColumns = constellationsSheet.getMaxColumns();
  var columnValue = constellationsSheet.getRange(1, 2).getValue();
  if (columnValue > 0) {
    var startValue = constellationsSheet.getRange(1, columnValue).getValue();
    var nextValue = constellationsSheet.getRange(1, columnValue+1).getValue();
    var userInputColumnValue = constellationsSheet.getRange(1, columnValue+2).getValue();
    var saveRowsValue = constellationsSheet.getRange(1, columnValue+4).getValue();
    for (var c = startValue; c <= maxColumns; c += nextValue) {
      var nameValue = constellationsSheet.getRange(1, c).getValue();
      if (nameValue != "") {
        var values = saveDict[nameValue];
        if (values) {
          var dataArray = [];
          for (var i = 0; i < values.length; i++) {
            dataArray.push([values[i]]);
          }
          constellationsSheet.getRange(16, c - userInputColumnValue,dataArray.length,1).setValues(dataArray);
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


var quickUpdateRange = [
  "A2","M2","Y2","AK2", // Banner Images
  "A3","M3","Y3" // Banner Time
];

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
    var currentSheet = SpreadsheetApp.getActive().getActiveSheet();
    reorderSheets();
    SpreadsheetApp.getActive().setActiveSheet(currentSheet);
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
    var listOfSheetsToRemove = ["Items","Events", "Pity Checker","Results","All Wish History", "Constellation", "Weapons"];

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
          } else if (sheetNameToRemove == "Constellation") {
            saveCollectionSettings(sheetToRemove, settingsSheet,"G7","H7");
          } else if (sheetNameToRemove == "Weapons") {
            saveCollectionSettings(sheetToRemove, settingsSheet,"G8","H8");
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
    var sheetEvents;
    if (shouldShowSheet) {
      var sheetEventsSource = sheetSource.getSheetByName('Events');
      sheetEvents = sheetEventsSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('Events');
      
      if (settingsSheet) {
        restoreEventsSettings(sheetEvents, settingsSheet);
      }
    }

    var sheetPityCheckerSource = sheetSource.getSheetByName('Pity Checker');
    var sheetPityChecker = sheetPityCheckerSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('Pity Checker');


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
    var sheetResults;
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
      sheetResults = sheetResultsSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('Results');
      
      if (settingsSheet) {
        restoreResultsSettings(sheetResults, settingsSheet);
      }
    }
    // Show Constellation
    shouldShowSheet = true;
    if (settingsSheet) {
      if (settingsSheet.getRange("B16").getValue()) {
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
        sheetConstellationSource = sheetSource.getSheetByName("Constellation"+"-"+languageFound);
      }
      if (sheetConstellationSource) {
        // Found language
      } else {
        // Default
        sheetConstellationSource = sheetSource.getSheetByName("Constellation");
      }
      if (sheetConstellationSource) {
        sheetConstellation = sheetConstellationSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('Constellation');
        // Refresh Contents Links
        var contentsAvailable = sheetConstellation.getRange(1, 1).getValue();
        var contentsStartIndex = 2;
        var richTextValues = [];
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
          .setLinkUrl(["#gid="+sheetConstellation.getSheetId()+'range='+bookmarkRange])
          .build();
          richTextValues.push([richText]);
        }
        var richTextValuesLength = richTextValues.length;
        if (richTextValuesLength > 0) {
          sheetConstellation.getRange(contentsStartIndex, 3, richTextValuesLength, 1).setRichTextValues(richTextValues);
        }
        if (settingsSheet) {
          restoreCollectionSettings(sheetConstellation, settingsSheet,"G7","H7");
        }
        
      }
    }
    // Show Weapons
    shouldShowSheet = true;
    if (settingsSheet) {
      if (settingsSheet.getRange("B22").getValue()) {
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
        sheetWeaponsSource = sheetSource.getSheetByName("Weapons"+"-"+languageFound);
      }
      if (sheetWeaponsSource) {
        // Found language
      } else {
        // Default
        sheetWeaponsSource = sheetSource.getSheetByName("Weapons");
      }
      if (sheetWeaponsSource) {
        sheetWeapons = sheetWeaponsSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('Weapons');
        // Refresh Contents Links
        var contentsAvailable = sheetWeapons.getRange(1, 1).getValue();
        var contentsStartIndex = 2;
        var richTextValues = [];
        for (var i = 0; i < contentsAvailable; i++) {
          var valueRange = sheetWeapons.getRange(contentsStartIndex+i, 3).getValue();
          var formulaRange = sheetWeapons.getRange(contentsStartIndex+i, 3).getFormula();
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
          .setLinkUrl(["#gid="+sheetWeapons.getSheetId()+'range='+bookmarkRange])
          .build();
          richTextValues.push([richText]);
        }
        var richTextValuesLength = richTextValues.length;
        if (richTextValuesLength > 0) {
          sheetConstellation.getRange(contentsStartIndex, 3, richTextValuesLength, 1).setRichTextValues(richTextValues);
        }
        if (settingsSheet) {
          restoreCollectionSettings(sheetWeapons, settingsSheet,"G8","H8");
        }
        
      }
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
    
    reorderSheets();
    // Bring Pity Checker into view
    
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
        iLastRow = iLastRow + 2;
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