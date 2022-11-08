/*
 * Version 3.50 made by yippym - 2022-11-07 21:00
 * https://github.com/Yippy/wish-tally-sheet
 */
function displayModalDiagram(sheet, sheetName, titleRange, htmlRange, widthSizeRange, heightSizeRange) {
  var isModalDisplayed = false;
  if (sheet) {
    var sheetByName = sheet.getSheetByName(sheetName);
    if (sheetByName) {
      var titleString = sheetByName.getRange(titleRange).getValue();
      var htmlString = sheetByName.getRange(htmlRange).getValue();
      var widthSize = sheetByName.getRange(widthSizeRange).getValue();
      var heightSize = sheetByName.getRange(heightSizeRange).getValue();

      var htmlOutput = HtmlService
        .createHtmlOutput(htmlString)
        .setWidth(widthSize) //optional
        .setHeight(heightSize); //optional
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, titleString);
      isModalDisplayed = true;
    }
  }
  if (!isModalDisplayed) {
    var message = 'Unable to connect to source, to find sheet "'+sheetName+'".';
    var title = 'Error';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function displayAbout() {
  displayModalDiagram(SpreadsheetApp.openById(WISH_TALLY_SHEET_SOURCE_REDIRECT_ID), WISH_TALLY_REDIRECT_SOURCE_ABOUT_SHEET_NAME, "B1", "B2", "B3", "B4");
}

function displayMaintenance() {
  displayModalDiagram(SpreadsheetApp.openById(WISH_TALLY_SHEET_SOURCE_REDIRECT_ID), WISH_TALLY_REDIRECT_SOURCE_MAINTENANCE_SHEET_NAME, "B1", "B2", "B3", "B4");
}

function displayBackup() {
  displayModalDiagram(SpreadsheetApp.openById(WISH_TALLY_SHEET_SOURCE_REDIRECT_ID), WISH_TALLY_REDIRECT_SOURCE_BACKUP_SHEET_NAME, "B1", "B2", "B3", "B4");
}

function displayAutoImport() {
  displayModalDiagram(SpreadsheetApp.openById(WISH_TALLY_SHEET_SOURCE_REDIRECT_ID), WISH_TALLY_REDIRECT_SOURCE_AUTO_IMPORT_SHEET_NAME, "B1", "B2", "B3", "B4");
}

function displayReadme() {
  var sheetSource = getSourceDocument();
  if (sheetSource) {
    // Avoid Exception: You can't remove all the sheets in a document.Details
    var placeHolderSheet = null;
    if (SpreadsheetApp.getActive().getSheets().length == 1) {
      placeHolderSheet = SpreadsheetApp.getActive().insertSheet();
    }
    var sheetToRemove = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_README_SHEET_NAME);
      if(sheetToRemove) {
        // If exist remove from spreadsheet
        SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheetToRemove);
      }
    var sheetREADMESource;

    // Add Language
    var settingsSheet = getSettingsSheet();
    if (settingsSheet) {
      var languageFound = settingsSheet.getRange(2, 2).getValue();
      sheetREADMESource = sheetSource.getSheetByName(WISH_TALLY_README_SHEET_NAME+"-"+languageFound);
    }
    if (sheetREADMESource) {
      // Found language
    } else {
      // Default
      sheetREADMESource = sheetSource.getSheetByName(WISH_TALLY_README_SHEET_NAME);
    }

    sheetREADMESource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(WISH_TALLY_README_SHEET_NAME);

    // Remove placeholder if available
    if(placeHolderSheet) {
      // If exist remove from spreadsheet
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(placeHolderSheet);
    }
    var sheetREADME = SpreadsheetApp.getActive().getSheetByName(WISH_TALLY_README_SHEET_NAME);
    // Refresh Contents Links
    var hyperlinkColumn = 1;
    var contentsAvailable = sheetREADME.getRange(13, hyperlinkColumn).getValue();
    var contentsStartIndex = 15;
    generateRichTextLinks(sheetREADME, contentsAvailable, contentsStartIndex, hyperlinkColumn, false);
    reorderSheets();
    SpreadsheetApp.getActive().setActiveSheet(sheetREADME);
  } else {
    var message = 'Unable to connect to source';
    var title = 'Error';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function generateRichTextLinks(sheet, contentsAvailable, contentsStartIndex, hyperlinkColumn, isBookmarkReference) {
  var richTextValues = [];
  var referenceRanges = null;
  var sheetId = sheet.getSheetId();
  if (isBookmarkReference) {
    referenceRanges=  sheet.getRange(contentsStartIndex, hyperlinkColumn+1, contentsAvailable, 1).getValues();
  } else {
    referenceRanges=  sheet.getRange(contentsStartIndex, hyperlinkColumn, contentsAvailable, 1).getRichTextValues();
  }

  var valueRanges = sheet.getRange(contentsStartIndex, hyperlinkColumn, contentsAvailable, 1).getValues();
  for (var i = 0; i < contentsAvailable; i++) {
    var valueRange = valueRanges[i][0];
    var linkURL = "";
    if (isBookmarkReference) {
      // Used in Characters and Weapons
      var bookmarkValue = referenceRanges[i][0];
      if (valueRange.includes("#gid=")) {
        // Meaning that the cell does not have a value
        valueRange = "";
      } else {
        linkURL = "#gid="+sheetId+'range='+bookmarkValue;
      }
    } else {
      // Used in README, grab URL RichTextValue from cell
      const richTextValue = referenceRanges[i][0].getRuns();
      const res = richTextValue.reduce((ar, e) => {
        const url = e.getLinkUrl();
        if (url) ar.push(url);
          return ar;
        }, []);
      //  Convert to string
      var resString = res+ "";
      var arrayString = resString.split("=");

      if (arrayString.length > 1) {
        var text = arrayString[2];
        linkURL = "#gid="+sheetId+'range='+text;
      }

      if (valueRange.includes("#gid=")) {
        valueRange = "";
      }
    }
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(valueRange)
      .setLinkUrl([linkURL])
      .build();

    richTextValues.push([richText]);
  }

  var richTextValuesLength = richTextValues.length;
  if (richTextValues.length > 0) {
    sheet.getRange(contentsStartIndex, hyperlinkColumn, richTextValuesLength, 1).setRichTextValues(richTextValues);
  }
}