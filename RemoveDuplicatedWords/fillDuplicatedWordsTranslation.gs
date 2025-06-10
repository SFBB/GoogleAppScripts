function fill() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(12000);

    var app = SpreadsheetApp.getActive();
    var sheet = app.getSheetByName("Words");
    var range = sheet.getDataRange();
    var values = range.getValues();
    var startRow = range.getRow();

    var validDict = new Map();
    for (var i=0; i < values.length; i++) {
      var row = startRow + i;
      var word = values[i][1].toString().trim();
      var translation = values[i][3].toString().trim();
      var translationLength = translation.toString().length;
      if (word == "") {
        continue;
      }
      if (validDict.has(word))
      {
        if (validDict.get(word).length < translationLength)
        {
          validDict.set(word, translation);
        }
      }
      else
      {
        validDict.set(word, translation);
      }
    }

    var duplicatedWordsStatsSheet = app.getSheetByName("Duplicated Words Stats");
    var duplicatedWordsStatsSheetRange = duplicatedWordsStatsSheet.getDataRange();
    var duplicatedWordsStatsSheetValues = duplicatedWordsStatsSheetRange.getValues();
    var duplicatedWordsStatsSheetStartRow = duplicatedWordsStatsSheetRange.getRow();
    const startFillIndex = Math.max(0, duplicatedWordsStatsSheetValues.length - 300);

    for (var i=startFillIndex; i < duplicatedWordsStatsSheetValues.length; i++) {
      var row = duplicatedWordsStatsSheetStartRow + i;
      var word = duplicatedWordsStatsSheetValues[i][1].toString().trim();
      if (word == "")
      {
        continue;
      }
      if (validDict.has(word) && validDict.get(word).toString().length > 0) {
        duplicatedWordsStatsSheet.getRange(row, 4).setValue(validDict.get(word));
        Utilities.sleep(200);
      }
    }

    lock.releaseLock();
  } catch (e) {
    Logger.log(e.toString());
    Logger.log('Could not obtain lock after 12 seconds.');
  }
}
