function main() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(12000);

    var app = SpreadsheetApp.getActive();
    var sheet = app.getSheetByName("Phrases");
    var range = sheet.getDataRange();
    var values = range.getValues();
    var startRow = range.getRow();

    var validDuplicatedRowNumber = [];
    var validDuplicatedPhraseTranslationMap = new Map();
    var validDuplicatedPhraseTranslationLength = [];
    var duplicatedPhrases = [];
    var visitedPhrases = [];
    var visitedPhraseRowNumber = [];
    var visitedPhraseTranslation = [];
    var visitedPhraseTranslationLength = [];
    for (var i=0; i < values.length; i++) {
      var row = startRow + i;
      var phrase = values[i][1].toString().trim();
      var translation = values[i][3].toString().trim();
      var translationLength = translation.toString().length;
      if (phrase == "") {
        continue;
      }
      var visitedResult = visitedPhrases.indexOf(phrase);
      if (visitedResult != -1) {
        var result = duplicatedPhrases.indexOf(phrase);
        if (result != -1 && translationLength > validDuplicatedPhraseTranslationLength[result]) {
          validDuplicatedRowNumber[result] = row;
          validDuplicatedPhraseTranslationLength[result] = translationLength;
          validDuplicatedPhraseTranslationMap.set(phrase, translation);
        }
        else if (result == -1) {
          duplicatedPhrases.push(phrase);
          validDuplicatedRowNumber.push(visitedPhraseTranslationLength[visitedResult] > translationLength ? visitedPhraseRowNumber[visitedResult] : row);
          validDuplicatedPhraseTranslationLength.push(visitedPhraseTranslationLength[visitedResult] > translationLength ? visitedPhraseTranslationLength[visitedResult] : translationLength);
          validDuplicatedPhraseTranslationMap.set(phrase, visitedPhraseTranslationLength[visitedResult] > translationLength ? visitedPhraseTranslation[visitedResult] : translation)
        }
      }
      visitedPhrases.push(phrase);
      visitedPhraseRowNumber.push(row);
      visitedPhraseTranslation.push(translation);
      visitedPhraseTranslationLength.push(translationLength);
    }

    var duplicatedRows = [];
    for (var i=0; i < values.length; i++) {
      var row = startRow + i;
      var phrase = values[i][1].toString().trim();
      if (duplicatedPhrases.indexOf(phrase) != -1 && validDuplicatedRowNumber.indexOf(row) == -1) {
        duplicatedRows.push(row);
      }
    }

    Logger.log(duplicatedRows.length);
    var removedDuplicatedRowCounter = 0;
    for (var i=0; i < duplicatedRows.length; i++) {
      sheet.deleteRow(duplicatedRows[i] - removedDuplicatedRowCounter);
      removedDuplicatedRowCounter += 1;
      Logger.log("Removed %s, row: %d, we have removed %d phrase%s!", values[duplicatedRows[i] - startRow][1].toString().trim(), duplicatedRows[i], removedDuplicatedRowCounter, removedDuplicatedRowCounter > 1 ? "s" : "");
      Utilities.sleep(200);
    }
    Logger.log(removedDuplicatedRowCounter);



    if (removedDuplicatedRowCounter > 0) {
      var duplicatedPhrasesStatsSheet = app.getSheetByName("Duplicated Phrases Stats");
      var duplicatedPhrasesStatsSheetRange = duplicatedPhrasesStatsSheet.getDataRange();
      var duplicatedPhrasesStatsSheetValues = duplicatedPhrasesStatsSheetRange.getValues();
      var duplicatedPhrasesStatsSheetStartRow = duplicatedPhrasesStatsSheetRange.getRow();
      var duplicatedPhrasesMap = new Map();
      var addedNewDuplicatedRecordsCounter = 0;

      for (var i=0; i < duplicatedPhrases.length; i++) {
        var phrase = duplicatedPhrases[i];
        var bFound = false;
        var cellRow = duplicatedPhrasesStatsSheetValues.length + duplicatedPhrasesStatsSheetStartRow + addedNewDuplicatedRecordsCounter;
        for (var j=0; j < duplicatedPhrasesStatsSheetValues.length; j++) {
          if (duplicatedPhrasesStatsSheetValues[j][1].toString().trim() == phrase)
          {
            cellRow = j + duplicatedPhrasesStatsSheetStartRow;
            bFound = true;
            break;
          }
        }
        if (!bFound) {
          duplicatedPhrasesStatsSheet.getRange(cellRow, 2).setValue(phrase);
          addedNewDuplicatedRecordsCounter += 1;
        }
        var chinese = duplicatedPhrasesStatsSheet.getRange(cellRow, 4).getValue().toString().trim();
        if (validDuplicatedPhraseTranslationMap.get(phrase).toString().length > 0) {
          duplicatedPhrasesStatsSheet.getRange(cellRow, 4).setValue(validDuplicatedPhraseTranslationMap.get(phrase));
          chinese = validDuplicatedPhraseTranslationMap.get(phrase);
        }
        Utilities.sleep(200);
        var counter_ = duplicatedPhrasesStatsSheet.getRange(cellRow, 6).getValue();
        var counter = 1;
        if (typeof counter == "number") {
          counter = counter_ + 1;
        }
        duplicatedPhrasesStatsSheet.getRange(cellRow, 6).setValue(counter);
        duplicatedPhrasesMap.set(phrase, [counter, chinese]);
        
        Utilities.sleep(200);
      }

      var emailBody = `We have removed ${duplicatedPhrasesMap.size} phrase${duplicatedPhrasesMap.size > 1 ? "s" : ""}, ${duplicatedRows.length} row${duplicatedRows.length > 1 ? "s" : ""}!\n\n\n`;
      emailBody += `Those phrases' statistics is here:\n`
      for (var key of duplicatedPhrasesMap.keys()) {
        if (duplicatedPhrasesMap.get(key)[1] != "") {
          emailBody += `\t${key}: ${duplicatedPhrasesMap.get(key)[0]} - ${duplicatedPhrasesMap.get(key)[1]}\n`;
        }
        else
        {
          emailBody += `\t${key}: ${duplicatedPhrasesMap.get(key)[0]}\n`;
        }
      }

      MailApp.sendEmail({
        to: "example@example.com",
        subject: `We have found ${duplicatedPhrasesMap.size} duplicated phrase${duplicatedPhrasesMap.size > 1 ? "s" : ""}`,
        body: emailBody
      });
    }

    lock.releaseLock();
  } catch (e) {
    Logger.log(e.toString());
    Logger.log('Could not obtain lock after 12 seconds.');
  }
}
