/*
 * Version 3.40 made by yippym - 2021-10-22 21:00
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
  displayModalDiagram(SpreadsheetApp.openById(WISH_TALLY_SHEET_SOURCE_REDIRECT_ID), WISH_TALLY_SOURCE_ABOUT_SHEET_NAME, "B1", "B2", "B3", "B4");
}

function displayMaintenance() {
  displayModalDiagram(SpreadsheetApp.openById(WISH_TALLY_SHEET_SOURCE_REDIRECT_ID), WISH_TALLY_REDIRECT_SOURCE_MAINTENANCE_SHEET_NAME, "B1", "B2", "B3", "B4");
}

function displayAutoImport() {
  displayModalDiagram(SpreadsheetApp.openById(WISH_TALLY_SHEET_SOURCE_REDIRECT_ID), WISH_TALLY_REDIRECT_SOURCE_AUTO_IMPORT_SHEET_NAME, "B1", "B2", "B3", "B4");
}

function displayReadme() {
  var sheetSource = SpreadsheetApp.openById(WISH_TALLY_SHEET_SOURCE_ID);
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