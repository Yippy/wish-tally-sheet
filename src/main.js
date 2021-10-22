/*
* Copyright (C) 2020 yippym
* Version 3.0 made by yippym - 2021-06-28 16:14
 */

var dashboardEditRange = [
  "I5", // Status cell
  "AB28", // Document Version
  "AT40", // Current Document Version
  "T29", // Document Status
  "AV1", // Name of drop down 1 (import)
  "AV2", // Name of drop down 2 (auto import)
  "AG14", // Selection
  "AG18" // URL
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
  "K26" // Weapon Total
];


function importButtonScript() {
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName('Dashboard');
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName('Settings');
  if (dashboardSheet && settingsSheet) {
    var userImportSelection = dashboardSheet.getRange(dashboardEditRange[4]).getValue();
    var importSelectionText = dashboardSheet.getRange(dashboardEditRange[6]).getValue();
    var urlInput = dashboardSheet.getRange(dashboardEditRange[7]).getValue();
    dashboardSheet.getRange(dashboardEditRange[7]).setValue(""); //Clear input
    if (userImportSelection == importSelectionText) {
      settingsSheet.getRange("D6").setValue(urlInput);
      importDataManagement();
    } else {
      settingsSheet.getRange("D35").setValue(urlInput);
      importFromAPI();
    }
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("Unable to find 'Dashboard' or 'Settings'", "Missing Sheets");
  }
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
  settingsSheet.getRange("D35").setValue("");
  if (AUTO_IMPORT_URL_FOR_API_BYPASS != "") {
    urlForAPI = AUTO_IMPORT_URL_FOR_API_BYPASS;
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
    for (var i = 0; i < WISH_TALLY_NAME_OF_WISH_HISTORY.length; i++) {
      bannerName = WISH_TALLY_NAME_OF_WISH_HISTORY[i];
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
    for (var i = 0; i < WISH_TALLY_NAME_OF_WISH_HISTORY.length; i++) {
      bannerName = WISH_TALLY_NAME_OF_WISH_HISTORY[i];
      bannerSettings = bannerSettingsForImport[bannerName];
      settingsSheet.getRange(bannerSettings['range_status']).setValue("");
    }
    for (var i = 0; i < WISH_TALLY_NAME_OF_WISH_HISTORY.length; i++) {
      if (errorCodeNotEncountered) {
        bannerName = WISH_TALLY_NAME_OF_WISH_HISTORY[i];
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
          
          var numberOfRows = bannerImportSheet.getMaxRows()-1;
          var range = bannerImportSheet.getRange(2, 1, numberOfRows, 2);

          if (bannerImportSheet && numberOfRows > 0) {
            var bannerSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_NAME_OF_WISH_HISTORY[i]);

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
        
        //Restore Constellation old name
        var sourceConstellationSheet = importSource.getSheetByName("Constellation");
        if (sourceConstellationSheet) {
          saveCollectionSettings(sourceConstellationSheet, settingsSheet,"G7","H7");
          // Restore save to the new Characters sheer
          var constellationSheet = SpreadsheetApp.getActive().getSheetByName('Characters');
          if (constellationSheet) {
            restoreCollectionSettings(constellationSheet, settingsSheet,"G7","H7");
          }
        } else {
          // Restore new name Characters
          sourceConstellationSheet = importSource.getSheetByName("Characters");
          if (sourceConstellationSheet) {
            saveCollectionSettings(sourceConstellationSheet, settingsSheet,"G7","H7");
            var constellationSheet = SpreadsheetApp.getActive().getSheetByName('Characters');
            if (constellationSheet) {
              restoreCollectionSettings(constellationSheet, settingsSheet,"G7","H7");
            }
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
  var sheetSource = SpreadsheetApp.openById(WISH_TALLY_SHEET_SOURCE_ID);
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
  var sheetSource = SpreadsheetApp.openById(WISH_TALLY_SHEET_SOURCE_ID);
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
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName('Dashboard');
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName('Settings');
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

      var changelogSheet = SpreadsheetApp.getActive().getSheetByName('Changelog');
      if (changelogSheet) {
        try {
          var sheetSource = SpreadsheetApp.openById(WISH_TALLY_SHEET_SOURCE_ID);
          if (sheetSource) {
            // get latest banners
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
            // check latest logs to see anything new
            if (dashboardSheet) {
              var sheetAvailableSource = sheetSource.getSheetByName("Available");
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
              var changesCheckRange = changelogSheet.getRange(2, 1).getValue();
              changesCheckRange = changesCheckRange.split(",");
              var lastDateChangeText;
              var lastDateChangeSourceText;
              var isChangelogTheSame = true;
              
              var sheetChangelogSource = sheetSource.getSheetByName("Changelog");
              for (var i = 0; i < changesCheckRange.length; i++) {
                var checkChangelogSource = sheetChangelogSource.getRange(changesCheckRange[i]).getValue();
                if (checkChangelogSource instanceof Date) {
                  lastDateChangeSourceText = Utilities.formatDate(checkChangelogSource, 'Etc/GMT', 'dd-MM-yyyy');
                }
                var checkChangelog = changelogSheet.getRange(changesCheckRange[i]).getValue();
                if (checkChangelog instanceof Date) {
                  lastDateChangeText = Utilities.formatDate(checkChangelog, 'Etc/GMT', 'dd-MM-yyyy');
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
                dashboardSheet.getRange(dashboardEditRange[0]).setValue("Quick Update: There is no changes from source");
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
  reorderSheets();
  SpreadsheetApp.getActive().setActiveSheet(currentSheet);
  // Update Settings
  settingsSheet.getRange(9, 7).setValue(false);
  settingsSheet.getRange("H10").setValue(new Date());
}

/**
* Update Item List
*/
function updateItemsList() {
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName('Dashboard');
  var updateItemHasFailed = false;
  if (dashboardSheet) {
    dashboardSheet.getRange(dashboardEditRange[0]).setValue("Update Items: Running script, please wait.");
    dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("yellow").setFontWeight("bold");
  }
  var sheetSource = SpreadsheetApp.openById(WISH_TALLY_SHEET_SOURCE_ID);
  // Check source is available
  if (sheetSource) {
    try {
      // attempt to load sheet from source, to prevent removing sheets first.
      var sheetAvailableSource = sheetSource.getSheetByName("Available");
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
      // var listOfSheetsToRemove = ["Items","Events", "Pity Checker","Results","All Wish History", "Constellation", "Weapons"];
      var listOfSheetsToRemove = ["Constellation","Characters","Weapons","Events","Results","Pity Checker","All Wish History","Items"];
      
      var availableRanges = sheetAvailableSource.getRange(2,1, sheetAvailableSource.getMaxRows()-1,1).getValues();
      availableRanges = String(availableRanges).split(",");
      
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
            } else if (sheetNameToRemove == "Characters") {
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
      
      if (dashboardSheet) {
        for (var i = 0; i < dashboardRefreshRange.length; i++) {
          var tempFormula = dashboardSheet.getRange(dashboardRefreshRange[i]).getFormula();
          // Re set formula
          dashboardSheet.getRange(dashboardRefreshRange[i]).setFormula(tempFormula);
        }
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
          sheetConstellation = sheetConstellationSource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName('Characters');
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
      
    } catch(e) {
      var message = 'Unable to connect to source';
      var title = 'Error';
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
      updateItemHasFailed = true;
      settingsSheet.getRange(5, 7).setValue(false);
      settingsSheet.getRange("H6").setValue(new Date());
    }
  } else {
    var message = 'Unable to connect to source';
    var title = 'Error';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
    updateItemHasFailed = true;
    settingsSheet.getRange(5, 7).setValue(false);
    settingsSheet.getRange("H6").setValue(new Date());
  }
  
  if (dashboardSheet) {
    if (updateItemHasFailed) {
      dashboardSheet.getRange(dashboardEditRange[0]).setValue("Update Items: Update Items has failed, please try again.");
      dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("red").setFontWeight("bold");
    } else {
      dashboardSheet.getRange(dashboardEditRange[0]).setValue("Update Items: Successfully updated the Item list.");
      dashboardSheet.getRange(dashboardEditRange[0]).setFontColor("green").setFontWeight("bold");
    }
  }
}