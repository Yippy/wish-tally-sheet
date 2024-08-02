/*
 * Version 4.02 made by yippym - 2024-08-03 00:50
 * https://github.com/Yippy/wish-tally-sheet
 */
const CACHED_LTUID_PROPERTY = "cachedLTUID_" + SpreadsheetApp.getActiveSpreadsheet().getId();
const CACHED_LTOKEN_PROPERTY = "cachedLTOKEN_" + SpreadsheetApp.getActiveSpreadsheet().getId();
const CACHED_UID_PROPERTY = "cachedUID_" + SpreadsheetApp.getActiveSpreadsheet().getId();

function getCachedLtuidInput() {
  const userProperties = PropertiesService.getUserProperties();
  return userProperties.getProperty(CACHED_LTUID_PROPERTY);
}

function setCachedLtuidInput(userInput) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(CACHED_LTUID_PROPERTY, userInput);
}

function getCachedLtokenInput() {
  const userProperties = PropertiesService.getUserProperties();
  return userProperties.getProperty(CACHED_LTOKEN_PROPERTY);
}

function setCachedLtokenInput(userInput) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(CACHED_LTOKEN_PROPERTY, userInput);
}

function getCachedUidInput() {
  const userProperties = PropertiesService.getUserProperties();
  return userProperties.getProperty(CACHED_UID_PROPERTY);
}

function setCachedUidInput(userInput) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(CACHED_UID_PROPERTY, userInput);
}

function isUidCorrect(userInput) {
  var isValid = false;
  if (userInput.match(/^[0-9]+$/) != null) {
    isValid = true;
  }
  return isValid;
}

// Allow flexibility in rolling out new changes, if weapons needs to run seperately for example.
function importHoYoLabWeaponsButtonScript() {
  importHoYoLabCharactersButtonScript();
}

function importHoYoLabCharactersButtonScript() {
  var ltuidInput = getCachedLtuidInput();
  var ltokenInput = getCachedLtokenInput();
  var uidInput = getCachedUidInput();
  if (ltuidInput) {
    var message = `Previous saved:\nHoYoLab UID (ltuid): '`+ ltuidInput+`'`;
    if (ltokenInput) {
      message += `\nHoYoLab Token (ltoken): '`+ltokenInput+`'`;
    }
    if (uidInput) {
      message += `\nGenshin Impact UID: '`+uidInput+`'\n\nPress 'Yes' to sync data.`;
    } else {
      message += `\nGenshin Impact UID: '`+uidInput+`'\n\nPress 'Yes' to use previous saved HoYoLab UID (ltuid).`;
    }
    message += `\nPress 'No' to begin editing HoYoLab UID (ltuid).\nPress 'Cancel' to check tutorial.`;

    const button = displayUserAlert("Import from HoYoLab", message, SpreadsheetApp.getUi().ButtonSet.YES_NO_CANCEL);
    if (button == SpreadsheetApp.getUi().Button.YES) {
      if (uidInput) {
        attemptHoYoLab(ltuidInput, ltokenInput, uidInput);
      } else {
        importHoYoLabCacheLtoken(ltuidInput);
      }
    } else if (button == SpreadsheetApp.getUi().Button.NO) {
      displayHoYoLabLtuid();
    } else if (button == SpreadsheetApp.getUi().Button.CANCEL) {
      displayHoYoLab();
    }
  } else {
    displayHoYoLabLtuid();
  }
}

function displayHoYoLabLtuid() {
  const result = displayUserPrompt("Import from HoYoLab", `Enter HoYoLAB UID (ltuid) to proceed.\n\nPress 'No' to stop.\nPress 'Cancel' to check tutorial.`,SpreadsheetApp.getUi().ButtonSet.YES_NO_CANCEL);
  var button = result.getSelectedButton();
  if (button == SpreadsheetApp.getUi().Button.YES) {
    var ltuidInput = result.getResponseText();
    if (isUidCorrect(ltuidInput)) {
      setCachedLtuidInput(ltuidInput);
      importHoYoLabCacheLtoken(ltuidInput);
    } else {
      displayUserAlert("Import from HoYoLab", "HoYoLAB UID (ltuid) is invalid, you have entered '"+ltuidInput+"'.", SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } else if (button == SpreadsheetApp.getUi().Button.CANCEL) {
    displayHoYoLab();
  }
}

function importHoYoLabCacheLtoken(ltuidInput) {
  var ltokenInput = getCachedLtokenInput();
  if (ltokenInput && ltokenInput.length > 0) {
    const button = displayUserAlert("Import from HoYoLab", `Previous saved\nHoYoLab Token  (ltoken)'`+ ltokenInput+`'\n\nPress 'No' to begin editing HoYoLab Token (ltoken)\nPress 'Yes' to use previous saved HoYoLab Token (ltoken)`, SpreadsheetApp.getUi().ButtonSet.YES_NO);
    if (button == SpreadsheetApp.getUi().Button.YES) {
      importHoYoLabCacheUid(ltuidInput, ltokenInput);
    } else if (button == SpreadsheetApp.getUi().Button.NO) {
      displayHoYoLabLtoken(ltuidInput, ltokenInput);
    }
  } else {
    displayHoYoLabLtoken(ltuidInput, ltokenInput);
  }
}

function displayHoYoLabLtoken(ltuidInput, ltokenInput) {
  const result = displayUserPrompt("Import from HoYoLab", `Enter HoYoLAB Token (ltoken) to proceed.\n`,SpreadsheetApp.getUi().ButtonSet.YES_NO);
  var button = result.getSelectedButton();
  if (button == SpreadsheetApp.getUi().Button.YES) {
    var ltokenInput = result.getResponseText();
    if (ltokenInput.length > 0) {
      setCachedLtokenInput(ltokenInput);
      importHoYoLabCacheUid(ltuidInput, ltokenInput)
    } else {
      displayUserAlert("Import from HoYoLab", "HoYoLAB Token (ltoken) is invalid, you have entered '"+ltokenInput+"'.", SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

function importHoYoLabCacheUid(ltuidInput, ltokenInput) {
  var uidInput = getCachedUidInput();
  if (uidInput && uidInput.length > 0) {
    const button = displayUserAlert("Import from HoYoLab", `Previous saved\nGenshin Impact UID '`+ uidInput+`'\n\nPress 'No' to begin editing Genshin Impact UID\nPress 'Yes' to use previous saved Genshin Impact UID`, SpreadsheetApp.getUi().ButtonSet.YES_NO);
    if (button == SpreadsheetApp.getUi().Button.YES) {
      attemptHoYoLab(ltuidInput, ltokenInput, uidInput);
    } else if (button == SpreadsheetApp.getUi().Button.NO) {
      displayHoYoLabUid(ltuidInput, ltokenInput);
    }
  } else {
    displayHoYoLabUid(ltuidInput, ltokenInput, uidInput);
  }
}

function displayHoYoLabUid(ltuidInput, ltokenInput) {
  const result = displayUserPrompt("Import from HoYoLab", `Enter Genshin Impact UID to proceed.\n`,SpreadsheetApp.getUi().ButtonSet.YES_NO);
  var button = result.getSelectedButton();
  if (button == SpreadsheetApp.getUi().Button.YES) {
    var uidInput = result.getResponseText();
    if (isUidCorrect(uidInput)) {
      setCachedUidInput(uidInput);
      attemptHoYoLab(ltuidInput, ltokenInput, uidInput);
    } else {
      displayUserAlert("Import from HoYoLab", "Genshin Impact UID is invalid, you have entered '"+uidInput+"'.", SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

// thesadru/genshinstats https://github.com/thesadru/genshinstats
function generateDs(salt) {
  var seconds = (new Date().getTime()/1000).toFixed(0);
  let randomletters = makeRandomLetters(6);
  var hash = generateHashMd5(convertStringToUft8("salt="+salt+"&t="+seconds+"&r="+randomletters));
  return seconds+","+randomletters+","+hash;
}

function generateDsChina(salt, body = null, query = null) {
  var seconds = (new Date().getTime()/1000).toFixed(0);
  var randomInt = makeRandomIntBetweenNumbers(100001, 200000);
  var body = (body == null ? "": JSON.stringify(body));
  var query = dictionaryToKeyAndValue(query);
  var hash = generateHashMd5(convertStringToUft8("salt="+salt+"&t="+seconds+"&r="+randomInt+"&b="+body+"&q="+query));
  return seconds+","+randomInt+","+hash;
}

function dictionaryToKeyAndValue(dict) {
  return Object.keys(dict).map( function(key){ return key+"="+dict[key] }).join("&");
}
function attemptHoYoLab(ltuidInput, ltokenInput, uidInput) {
  var serverOptions = UID_SERVER[String(uidInput)[0]];
  var cookie = 'ltuid='+ltuidInput+';ltoken='+ltokenInput+';';
  var settingsSheet = getSettingsSheet();
  var languageSettings = AUTO_IMPORT_LANGUAGE_SETTINGS_FOR_IMPORT[settingsSheet.getRange("B2").getValue()];
  if (languageSettings == null) {
    // Get default language
    languageSettings = AUTO_IMPORT_LANGUAGE_SETTINGS_FOR_IMPORT["English"];
  }
  var ds = null;
  var param = {"role_id": uidInput, "server": serverOptions.code};
  var json = {"role_id": uidInput, "server": serverOptions.code};
  if (serverOptions["is_global"] == true) {
    ds = generateDs(serverOptions.salt);
  } else {
    ds = generateDsChina(serverOptions.salt, json, param);
  }
  var headers = {
    'Cookie': cookie,
    'x-rpc-app_version': serverOptions["x-rpc-app_version"],
    'x-rpc-client_type': serverOptions["x-rpc-client_type"],
    'x-rpc-language': languageSettings["full_code"],
    'ds': ds,
    'json': json
  };
  var opt = {"headers": headers, "method": "post"};
  var url = generateHoYoLabURL(serverOptions.url, param);
  var request = UrlFetchApp.getRequest(url, opt);
  delete request.headers["X-Forwarded-For"];

  var responses =  UrlFetchApp.fetchAll([request]);
  var response = responses.map(function(e) {return JSON.parse(e.getContentText())})[0];
  if (response.retcode !== 0) {
    const button = displayUserAlert("Import from HoYoLab", `Failed to load data from HoYoLab:\n Error code: '`+ response.retcode+`'\nMessage: `+response.message+"\nurl:"+url+"\nds:"+ds+"\n'cookie:"+cookie, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    var title = "Import from HoYoLab";
    var message = "Access granted, attempting to load data. Please wait...";
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);

    var characters = {};
    var weapons = {};
    var avatars = response.data.avatars;
    for (var i = 0; i < avatars.length; i++) {
      var character = avatars[i];
      var characterName = character.name;
      var characterData = {
        "id": character.id,
        "level": character.level,
        "friendship_level": character.fetter,
        "owned": character.actived_constellation_num+1,
      };
      if (character.weapon) {
        var weaponName = character.weapon.name;
        var weaponData =  {
          "name": weaponName,
          "level": character.weapon.level,
          "refinement_level": character.weapon.affix_level
        };
        characterData["weapon"] = weaponData;
        var foundWeapon = weapons[weaponName];
        if (foundWeapon) {
          if (weaponData["refinement_level"] > foundWeapon["refinement_level"]) {
            weapons[weaponName] = weaponData;
          } else if (weaponData["level"] > foundWeapon["level"]) {
            weapons[weaponName] = weaponData;
          }
        } else {
          weapons[weaponName] = weaponData;
        }

      }
      characters[characterName] = characterData;
    }
    processCharacters(characters, weapons);
  }
}

function processCharacters(characters, weapons) {
  var constellationSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_CHARACTERS_SHEET_NAME);
  var sheetsUpdated = [];
  if (constellationSheet) {
    var title = "Import from HoYoLab";
    var message = "Saving data to "+WISH_TALLY_CHARACTERS_SHEET_NAME+" sheet.";
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
    syncCollectionSettings(constellationSheet, characters, WISH_TALLY_CHARACTERS_SHEET_NAME);
    sheetsUpdated.push(WISH_TALLY_CHARACTERS_SHEET_NAME);
  }
  var constellationSheet = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_WEAPONS_SHEET_NAME);
  if (constellationSheet) {
    var title = "Import from HoYoLab";
    var message = "Saving data to "+WISH_TALLY_WEAPONS_SHEET_NAME+" sheet.";
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
    syncCollectionSettings(constellationSheet, weapons, WISH_TALLY_WEAPONS_SHEET_NAME);
    sheetsUpdated.push(WISH_TALLY_WEAPONS_SHEET_NAME);
  }

  var title = "Import from HoYoLab";
  var message = "Nothing was save, due to missing sheets";
  if (sheetsUpdated.length > 0) {
    message = "The following sheets has been sync with HoYoLab:\n"+sheetsUpdated.join(", ")+".";
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
}

function syncCollectionSettings(constellationsSheet, objects, sheetName) {
  var maxColumns = constellationsSheet.getMaxColumns();
  var columnValue = constellationsSheet.getRange(1, 2).getValue();
  var nameRowValue = constellationsSheet.getRange(1, 10).getValue();

  if (columnValue > 0) {
    var startValue = constellationsSheet.getRange(1, columnValue).getValue();
    var nextValue = constellationsSheet.getRange(1, columnValue+1).getValue();
    var userInputColumnValue = constellationsSheet.getRange(1, columnValue+2).getValue();
    var saveRowsValue = constellationsSheet.getRange(1, columnValue+4).getValue();
    var startSaveRowValue = constellationsSheet.getRange(1, 11).getValue();
    for (var c = startValue; c <= maxColumns; c += nextValue) {
      var nameValue = constellationsSheet.getRange(nameRowValue, c).getValue();
      if (nameValue != "") {
        var foundObject = objects[nameValue];
        if (foundObject) {
          var saveValues = constellationsSheet.getRange(startSaveRowValue, c - userInputColumnValue,saveRowsValue,1).getValues();
          var totalValue = constellationsSheet.getRange(startSaveRowValue-1, c - userInputColumnValue).getValue();
          if (sheetName == WISH_TALLY_CHARACTERS_SHEET_NAME) {
            if (10000005 != foundObject.id) {
              var override = saveValues[0][0];
              if (override > 0) {
                //cancel out user override, and try and get true total.
                override = totalValue - override;
              } else {
                override = totalValue;
              }
              if (override >= foundObject.owned) {
                // No need to apple any user override, if total is more than owned anyway.
                override = 0;
              } else {
                override = foundObject.owned - override;
                if (override < 0) {
                  override = 0;
                }
              }
              saveValues[0] = [override];
            }
            saveValues[1] = [foundObject.level]
            if (foundObject.weapon) {
              saveValues[10] = [foundObject.weapon.name]
            }
          } else if (sheetName == WISH_TALLY_WEAPONS_SHEET_NAME) {
            var override = saveValues[0][0];
            if ((override + totalValue) < foundObject.refinement_level) {
              override = foundObject.refinement_level - totalValue;
            } else {
              override = 0;
            }
            saveValues[0] = [override];
            saveValues[1] = [foundObject.level];
            saveValues[2] = [foundObject.refinement_level];
          }
          constellationsSheet.getRange(startSaveRowValue, c - userInputColumnValue,saveValues.length,1).setValues(saveValues);
        }
      }
    }
  }
}

function makeRandomIntBetweenNumbers(min, max) { // min and max included 
  return Math.floor(Math.random() * (max - min + 1) + min)
}

function makeRandomLetters(length) {
  var result = '';
  var characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
  var charactersLength = characters.length;
  for ( var i = 0; i < length; i++ ) {
    result += characters.charAt(Math.floor(Math.random() * charactersLength));
  }
  return result;
}

function generateHoYoLabURL(url, param) {
  return url.concat("?",dictionaryToKeyAndValue(param));
}

function generateHashMd5(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str).reduce(function(str,chr){
    chr = (chr < 0 ? chr + 256 : chr).toString(16);
    return str + (chr.length==1?'0':'') + chr;
  },'');
}

function convertStringToUft8(strInput, outputCharset="UTF-8"){
  return Utilities.newBlob("").setDataFromString(strInput).getDataAsString(outputCharset);
}