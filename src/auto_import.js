/*
 * Version 4.00 made by yippym - 2023-02-57 23:00
 * https://github.com/Yippy/wish-tally-sheet
 */
function extractAuthKeyFromInput(userInput) {
  urlForAPI = userInput.toString().split("&");
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
  return foundAuth;
}

function testAuthKeyInputValidity(userInput) {
  var authKey = extractAuthKeyFromInput(userInput);
  if (authKey == "") {
    return false;
  }

  const USING_BANNER = "Permanent Wish History";

  var settingsSheet = getSettingsSheet();
  var queryBannerCode = AUTO_IMPORT_BANNER_SETTINGS_FOR_IMPORT[USING_BANNER].gacha_type;
  var selectedServer = settingsSheet.getRange("B3").getValue();
  var languageSettings = AUTO_IMPORT_LANGUAGE_SETTINGS_FOR_IMPORT[settingsSheet.getRange("B2").getValue()];
  if (languageSettings == null) {
    // Get default language
    languageSettings = AUTO_IMPORT_LANGUAGE_SETTINGS_FOR_IMPORT["English"];
  }
  var url = getWishHistoryUrl(selectedServer, queryBannerCode, languageSettings, 1, authKey);
  responseJson = JSON.parse(UrlFetchApp.fetch(url).getContentText());
  if (responseJson.retcode === 0) {
    return true;
  }
  return false;
}

function getCachedAuthKeyProperty() {
  return  "cachedAuthKey_" + SpreadsheetApp.getActiveSpreadsheet().getId();
}

function invalidateCachedAuthKey() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty(getCachedAuthKeyProperty());
}

function setCachedAuthKeyInput(userInput) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(getCachedAuthKeyProperty(), JSON.stringify({ userInput, timeOfInput: new Date() }));
}

function getCachedAuthKeyInput() {
  const userProperties = PropertiesService.getUserProperties();
  const cachedAuthKey = JSON.parse(userProperties.getProperty(getCachedAuthKeyProperty()));

  if (cachedAuthKey == null) {
    return null;
  }

  const timeOfInput = new Date(cachedAuthKey.timeOfInput);
  const timeDiff = new Date().getTime() - timeOfInput.getTime();
  if (timeDiff > CACHED_AUTHKEY_TIMEOUT) {
    invalidateCachedAuthKey();
    return null;
  }

  if (!testAuthKeyInputValidity(cachedAuthKey.userInput)) {
    invalidateCachedAuthKey();
    return null;
  }

  return cachedAuthKey.userInput;
}


function importFromAPI(urlForAPI) {
  var settingsSheet = getSettingsSheet();
  settingsSheet.getRange("E42").setValue(new Date());
  settingsSheet.getRange("E43").setValue("");

  if (AUTO_IMPORT_URL_FOR_API_BYPASS != "") {
    urlForAPI = AUTO_IMPORT_URL_FOR_API_BYPASS;
  }
  var authKey = extractAuthKeyFromInput(urlForAPI);
  var bannerName;
  var bannerSheet;
  var bannerSettings;
  if (authKey == "") {
    // Display auth key not available
    for (var i = 0; i < WISH_TALLY_NAME_OF_WISH_HISTORY.length; i++) {
      bannerName = WISH_TALLY_NAME_OF_WISH_HISTORY[i];
      bannerSettings = AUTO_IMPORT_BANNER_SETTINGS_FOR_IMPORT[bannerName];
      bannerSettings.setStatusText("No auth key", settingsSheet);
    }
  } else {
    var selectedLanguageCode = settingsSheet.getRange("B2").getValue();
    var selectedServer = settingsSheet.getRange("B3").getValue();
    var languageSettings = AUTO_IMPORT_LANGUAGE_SETTINGS_FOR_IMPORT[selectedLanguageCode];
    if (languageSettings == null) {
      // Get default language
      languageSettings = AUTO_IMPORT_LANGUAGE_SETTINGS_FOR_IMPORT["English"];
    }
    // Clear status
    for (var i = 0; i < WISH_TALLY_NAME_OF_WISH_HISTORY.length; i++) {
      bannerName = WISH_TALLY_NAME_OF_WISH_HISTORY[i];
      bannerSettings = AUTO_IMPORT_BANNER_SETTINGS_FOR_IMPORT[bannerName];
      bannerSettings.setStatusText("", settingsSheet);
    }
    for (var i = 0; i < WISH_TALLY_NAME_OF_WISH_HISTORY.length; i++) {
      bannerName = WISH_TALLY_NAME_OF_WISH_HISTORY[i];
      bannerSettings = AUTO_IMPORT_BANNER_SETTINGS_FOR_IMPORT[bannerName];
      if (bannerSettings.isEnabled(settingsSheet)) {
        bannerSheet = SpreadsheetApp.getActive().getSheetByName(bannerName);
        if (bannerSheet) {
          var success = checkPages(
            bannerSheet,
            bannerName,
            bannerSettings,
            languageSettings,
            selectedServer,
            settingsSheet,
            authKey);
          if (!success) {
            bannerSettings.setStatusText(
              "Stopped Due to Error:\n" + bannerSettings.getStatusText(settingsSheet),
              settingsSheet);
            break;
          }
        } else {
          bannerSettings.setStatusText("Missing sheet", settingsSheet);
        }
      } else {
        bannerSettings.setStatusText("Skipped", settingsSheet);
      }
    }
  }
  settingsSheet.getRange("E43").setValue(new Date());
}

function checkPages(bannerSheet, bannerName, bannerSettings, languageSettings, selectedServer, settingsSheet, authKey) {
  bannerSettings.setStatusText("Starting", settingsSheet);
  /* Get latest wish from banner */
  var iLastRow = bannerSheet.getRange(2, 5, bannerSheet.getLastRow(), 1).getValues().filter(String).length;
  var wishTextString;
  var lastWishDateAndTimeString;
  var lastWishDateAndTime;
  if (iLastRow && iLastRow != 0 ) {
    iLastRow++;
    lastWishDateAndTimeString = bannerSheet.getRange("E" + iLastRow).getValue();
    wishTextString = bannerSheet.getRange("A" + iLastRow).getValue();
    if (lastWishDateAndTimeString) {
      bannerSettings.setStatusText(`Last wish: ${lastWishDateAndTimeString}`, settingsSheet);
      lastWishDateAndTimeString = lastWishDateAndTimeString.split(" ").join("T");
      lastWishDateAndTime = new Date(lastWishDateAndTimeString+".000Z");
    } else {
      iLastRow = 1;
      bannerSettings.setStatusText("No previous wishes", settingsSheet);
    }
    iLastRow++; // Move last row to new row
  } else {
    iLastRow = 2; // Move last row to new row
    bannerSettings.setStatusText("", settingsSheet);
  }
  
  var extractWishes = [];
  var page = 1;
  var queryBannerCode = bannerSettings.gacha_type;
  var numberOfWishPerPage = 6;
  var urlForBanner = getWishHistoryUrl(
    selectedServer,
    queryBannerCode,
    languageSettings,
    numberOfWishPerPage,
    authKey);
  var failed = 0;
  var is_done = false;
  var end_id = 0;
  
  var checkPreviousDateAndTimeString = "";
  var checkPreviousDateAndTime;
  var checkOneSecondOffDateAndTime;
  var overrideIndex = 0;
  var textWish;
  var oldTextWish;
  while (!is_done) {
    bannerSettings.setStatusText("Loading page: "+page, settingsSheet);
    var response = UrlFetchApp.fetch(urlForBanner+"&page="+page+"&end_id="+end_id);
    var jsonResponse = response.getContentText();
    var jsonDict = JSON.parse(jsonResponse);
    var jsonDictData = jsonDict["data"];
    if (jsonDictData) {
      var listOfWishes = jsonDictData["list"];
      var listOfWishesLength = listOfWishes.length;
      var wish;
      if (listOfWishesLength > 0) {
        for (var i = 0; i < listOfWishesLength; i++) {
          wish = listOfWishes[i];
          var dateAndTimeString = wish['time'];
          textWish = wish['item_type']+wish['name'];
          /* Mimic the website in showing specific language wording */
          if (wish['rank_type'] == 4) {
            textWish += languageSettings["4_star"];
          } else if (wish['rank_type'] == 5) {
            textWish += languageSettings["5_star"];
          }
          oldTextWish = textWish+dateAndTimeString;
          var gachaString = "gacha_type_"+wish['gacha_type'];
          var bannerName = "Error New Banner";
          if (gachaString in languageSettings) {
            bannerName = languageSettings[gachaString];
          }
          textWish += bannerName+dateAndTimeString;

          var dateAndTimeStringModified = dateAndTimeString.split(" ").join("T");
          var wishDateAndTime = new Date(dateAndTimeStringModified+".000Z");

          if (overrideIndex == 0 && checkPreviousDateAndTime) {
            /* Check one second difference from previous single wish */
            checkOneSecondOffDateAndTime = new Date(checkPreviousDateAndTime.valueOf());
            checkOneSecondOffDateAndTime.setSeconds(checkOneSecondOffDateAndTime.getSeconds()-1);
            if (checkOneSecondOffDateAndTime.valueOf() == wishDateAndTime.valueOf()) {
              var nextWishIndex = i+1;
              if (nextWishIndex < listOfWishesLength) {
                var nextWish = listOfWishes[nextWishIndex];
                var nextDateAndTimeString = nextWish['time'];
                var nextDateAndTimeStringModified = nextDateAndTimeString.split(" ").join("T");
                var nextWishDateAndTime = new Date(nextDateAndTimeStringModified+".000Z");
                if (checkOneSecondOffDateAndTime.valueOf() == nextWishDateAndTime.valueOf()) {
                  // Due to wish date and time is only second difference, it's therefore a multi. Override previous wish to match.
                  checkPreviousDateAndTimeString = dateAndTimeString;
                  checkPreviousDateAndTime = new Date(wishDateAndTime.valueOf());
                }
              }
            }
          }
          if (checkPreviousDateAndTimeString === dateAndTimeString) {
            // Found matching date and time to previous wish
            if (overrideIndex == 0) {
              // Start multi 10 index
              var previousWishIndex = extractWishes.length - 1;
              var previousWish = extractWishes[previousWishIndex];
              overrideIndex = 10;
              previousWish[1] = overrideIndex;
              extractWishes[previousWishIndex] = previousWish;
            }
            if (overrideIndex == 1) {
              bannerSettings.setStatusText(
                `Error: Multi wish contains 11 within same date and time: ${dateAndTimeString}, found so far: ${extractWishes.length}`,
                settingsSheet);
              return false;
            } else {
              overrideIndex--;
            }
          } else {
            if (overrideIndex > 1) {
              // Resume counting down when override is set more than 1, add a second to checkPreviousDateAndTime
              checkPreviousDateAndTime.setSeconds(checkPreviousDateAndTime.getSeconds()-1);
              if (checkPreviousDateAndTime.valueOf() == wishDateAndTime.valueOf()) {
                // Within 1 second range resuming multi count
                overrideIndex--;
              } else {
                bannerSettings.setStatusText(
                  `Error: Multi wish is incomplete with override ${overrideIndex}@${dateAndTimeString}, found so far: ${extractWishes.length}`,
                  settingsSheet);
                return false;
              }
            } else {
              // Default value for single wishes
              overrideIndex = 0;
            }
            checkPreviousDateAndTimeString = dateAndTimeString;
            checkPreviousDateAndTime = new Date(wishDateAndTime.valueOf());
          }
          if (lastWishDateAndTime >= wishDateAndTime) {
            // Banner already got this wish
            is_done = true;
            break;
          } else {
            extractWishes.push([textWish, (overrideIndex > 0 ? overrideIndex:null)]);
          }
        }
        if (!is_done && numberOfWishPerPage == listOfWishesLength) {
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
      var return_code = jsonDict["retcode"];

      var title ="Error code: "+return_code;
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title);

      switch (return_code) {
        case AUTO_IMPORT_URL_ERROR_CODE_AUTHKEY_DENIED:
          bannerSettings.setStatusText("feedback URL\nNo Longer Works", settingsSheet);
          return false;
        case AUTO_IMPORT_URL_ERROR_CODE_AUTH_TIMEOUT:
          bannerSettings.setStatusText("auth timeout", settingsSheet);
          return false;
        case AUTO_IMPORT_URL_ERROR_CODE_AUTH_INVALID:
          bannerSettings.setStatusText("auth invalid", settingsSheet);
          return false;
        case AUTO_IMPORT_URL_ERROR_CODE_REQUEST_PARAMS:
          bannerSettings.setStatusText("Change server setting", settingsSheet);
          return false;
        default:
          bannerSettings.setStatusText(`Unknown return code: ${return_code}`, settingsSheet);
          failed++;
          if (failed > 2){
            bannerSettings.setStatusText("Failed too many times", bannerSettings, settingsSheet);
            // Preserve legacy behavior which did not treat this the same as other error code
            // cases
            return true;
          }
      }
    }
  }

  if (extractWishes.length > 0) {
    var now = new Date();
    var sixMonthBeforeNow = new Date(now.valueOf());
    sixMonthBeforeNow.setMonth(now.getMonth() - 6);
    var isValid = true;
    var outputString = "Found: "+extractWishes.length;
    if (!lastWishDateAndTime) {
      // fresh history sheet no last date to check
      outputString += ", with wish history being empty"
    } else if (lastWishDateAndTime < sixMonthBeforeNow) {
      // Check if last wish found is more than 6 months, no further validation
      outputString += ", last wish saved was 6 months ago, maybe missing wishes inbetween"
    } else {
      if (wishTextString !== textWish) {
        if (wishTextString !== oldTextWish) {
          // API didn't reach to your last wish stored on the sheet, meaning the API is incomplete
          isValid = false;
          outputString = "Error your recently found wishes did not reach to your last wish, found: "+extractWishes.length+", please try again miHoYo may have sent incomplete wish data.";
        }
      }
    }
    if (isValid) {
      extractWishes.reverse();
      bannerSheet.getRange(iLastRow, 1, extractWishes.length, 2).setValues(extractWishes);
    }
    bannerSettings.setStatusText(outputString, settingsSheet);
  } else {
    bannerSettings.setStatusText("Nothing to add", settingsSheet);
  }

  return true;
}

function getWishHistoryUrl(selectedServer, queryBannerCode, languageSettings, numberOfWishPerPage, authKey) {
    var base_url;
    if (selectedServer == "China") {
      base_url = AUTO_IMPORT_URL_CHINA;
    } else {
      base_url = AUTO_IMPORT_URL;
    }
    return `${base_url}?${AUTO_IMPORT_ADDITIONAL_QUERY.join("&")}&authkey=${authKey}&lang=${languageSettings['code']}&gacha_type=${queryBannerCode}&size=${numberOfWishPerPage}`;
}
