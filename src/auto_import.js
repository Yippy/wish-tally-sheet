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