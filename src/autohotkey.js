/*
 * Version 3.60 made by yippym - 2023-01-12 01:00
 * https://github.com/Yippy/wish-tally-sheet
 */
function clearAHK() {
  var autoHotkeySheet = SpreadsheetApp.getActive().getSheetByName(AUTOHOTKEY_SHEET_NAME);
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
  var autoHotkeySheet = SpreadsheetApp.getActive().getSheetByName(AUTOHOTKEY_SHEET_NAME);
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
  var autoHotkeySheet = SpreadsheetApp.getActive().getSheetByName(AUTOHOTKEY_SHEET_NAME);
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
  var autoHotkeySheet = SpreadsheetApp.getActive().getSheetByName(AUTOHOTKEY_SHEET_NAME);
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
  var autoHotkeyScriptSheet = SpreadsheetApp.getActive().getSheetByName(AUTOHOTKEY_SCRIPT_SHEET_NAME);
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