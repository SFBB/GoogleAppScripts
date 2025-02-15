function fill() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(12000);

    var app = SpreadsheetApp.getActive();
    var sheet = app.getSheetByName("Phrases");
    var range = sheet.getDataRange();
    var values = range.getValues();
    var startRow = range.getRow();

    var validDict = new Map();
    for (var i=0; i < values.length; i++) {
      var row = startRow + i;
      var phrase = values[i][1].toString().trim();
      var translation = values[i][3].toString().trim();
      var translationLength = translation.toString().length;
      if (phrase == "") {
        continue;
      }
      if (validDict.has(phrase))
      {
        if (validDict.get(phrase).length < translationLength)
        {
          validDict.set(phrase, translation);
        }
      }
      else
      {
        validDict.set(phrase, translation);
      }
    }

    var duplicatedPhrasesStatsSheet = app.getSheetByName("Duplicated Phrases Stats");
    var duplicatedPhrasesStatsSheetRange = duplicatedPhrasesStatsSheet.getDataRange();
    var duplicatedPhrasesStatsSheetValues = duplicatedPhrasesStatsSheetRange.getValues();
    var duplicatedPhrasesStatsSheetStartRow = duplicatedPhrasesStatsSheetRange.getRow();

    for (var i=0; i < duplicatedPhrasesStatsSheetValues.length; i++) {
      var row = duplicatedPhrasesStatsSheetStartRow + i;
      var phrase = duplicatedPhrasesStatsSheetValues[i][1].toString().trim();
      if (phrase == "")
      {
        continue;
      }
      if (validDict.has(phrase) && validDict.get(phrase).toString().length > 0) {
        duplicatedPhrasesStatsSheet.getRange(row, 4).setValue(validDict.get(phrase));
        Utilities.sleep(200);
      }
    }

    lock.releaseLock();
  } catch (e) {
    Logger.log(e.toString());
    Logger.log('Could not obtain lock after 12 seconds.');
  }
}
