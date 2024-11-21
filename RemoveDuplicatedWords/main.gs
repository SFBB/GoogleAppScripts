function main() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(12000);

    var app = SpreadsheetApp.getActive();
    var sheet = app.getSheetByName("Words");
    var range = sheet.getDataRange();
    var values = range.getValues();
    var startRow = range.getRow();

    var validDuplicatedRowNumber = [];
    var validDuplicatedWordTranslationLength = [];
    var duplicatedWords = [];
    var visitedWords = [];
    var visitedWordRowNumber = [];
    var visitedWordTranslationLength = [];
    for (var i=0; i < values.length; i++) {
      var row = startRow + i;
      var word = values[i][1];
      var translation = values[i][3];
      var translationLength = translation.toString().length;
      if (word == "") {
        continue;
      }
      var visitedResult = visitedWords.indexOf(word);
      if (visitedResult != -1) {
        var result = duplicatedWords.indexOf(word);
        if (result != -1 && translationLength > validDuplicatedWordTranslationLength[result]) {
          validDuplicatedRowNumber[result] = row;
          validDuplicatedWordTranslationLength[result] = translationLength;
        }
        else if (result == -1) {
          duplicatedWords.push(word);
          validDuplicatedRowNumber.push(visitedWordTranslationLength[visitedResult] > translationLength ? visitedWordRowNumber[visitedResult] : row);
          validDuplicatedWordTranslationLength.push(visitedWordTranslationLength[visitedResult] > translationLength ? visitedWordTranslationLength[visitedResult] : translationLength);
        }
      }
      visitedWords.push(word);
      visitedWordRowNumber.push(row);
      visitedWordTranslationLength.push(translationLength);
    }

    var duplicatedRows = [];
    for (var i=0; i < values.length; i++) {
      var row = startRow + i;
      var word = values[i][1];
      if (duplicatedWords.indexOf(word) != -1 && validDuplicatedRowNumber.indexOf(row) == -1) {
        duplicatedRows.push(row);
      }
    }

    Logger.log(duplicatedRows.length);
    var removedDuplicatedRowCounter = 0;
    for (var i=0; i < duplicatedRows.length; i++) {
      sheet.deleteRow(duplicatedRows[i] - removedDuplicatedRowCounter);
      removedDuplicatedRowCounter += 1;
      Logger.log("Removed %s, row: %d, we have removed %d word%s!", values[duplicatedRows[i] - startRow][1].toString(), duplicatedRows[i], removedDuplicatedRowCounter, removedDuplicatedRowCounter > 1 ? "s" : "");
      Utilities.sleep(200);
    }
    Logger.log(removedDuplicatedRowCounter);



    if (removedDuplicatedRowCounter > 0) {
      var duplicatedWordsStatsSheet = app.getSheetByName("Duplicated Words Stats");
      var duplicatedWordsStatsSheetRange = duplicatedWordsStatsSheet.getDataRange();
      var duplicatedWordsStatsSheetValues = duplicatedWordsStatsSheetRange.getValues();
      var duplicatedWordsStatsSheetStartRow = duplicatedWordsStatsSheetRange.getRow();
      var duplicatedWordsMap = new Map();

      for (var i=0; i < duplicatedWords.length; i++) {
        var word = duplicatedWords[i];
        var bFound = false;
        var cellRow = duplicatedWordsStatsSheetValues.length + duplicatedWordsStatsSheetStartRow;
        for (var j=0; j < duplicatedWordsStatsSheetValues.length; j++) {
          if (duplicatedWordsStatsSheetValues[j][1] == word)
          {
            cellRow = j + duplicatedWordsStatsSheetStartRow;
            bFound = true;
            break;
          }
        }
        if (!bFound) {
          duplicatedWordsStatsSheet.getRange(cellRow, 2).setValue(word);
        }
        var chinese = duplicatedWordsStatsSheet.getRange(cellRow, 4).getValue();
        var counter_ = duplicatedWordsStatsSheet.getRange(cellRow, 6).getValue();
        var counter = 1;
        if (typeof counter == "number") {
          counter = counter_ + 1;
        }
        duplicatedWordsStatsSheet.getRange(cellRow, 6).setValue(counter);
        duplicatedWordsMap.set(word, [counter, chinese]);
        
        Utilities.sleep(200);
      }

      var emailBody = `We have removed ${duplicatedWordsMap.size} word${duplicatedWordsMap.size > 1 ? "s" : ""}, ${duplicatedRows.length} row${duplicatedRows.length > 1 ? "s" : ""}!\n\n\n`;
      emailBody += `Those words' statistics is here:\n`
      for (var key of duplicatedWordsMap.keys()) {
        if (duplicatedWordsMap.get(key)[1] != "") {
          emailBody += `\t${key}: ${duplicatedWordsMap.get(key)[0]} - ${duplicatedWordsMap.get(key)[1]}\n`;
        }
        else
        {
          emailBody += `\t${key}: ${duplicatedWordsMap.get(key)[0]}\n`;
        }
      }

      MailApp.sendEmail({
        to: "example@exmaple.com",
        subject: `We have found ${duplicatedWordsMap.size} duplicated word${duplicatedWordsMap.size > 1 ? "s" : ""}`,
        body: emailBody
      });
    }

    lock.releaseLock();
  } catch (e) {
    Logger.log(e.toString());
    Logger.log('Could not obtain lock after 12 seconds.');
  }
}
