function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var data = JSON.parse(e.postData.contents);

    // type=saveScore: 채점 결과를 SCORE 시트에 저장 (학생 제출 시 프론트에서 호출)
    if (e.parameter && e.parameter.type === 'saveScore') {
      var scoreSheet = doc.getSheetByName('SCORE');
      if (!scoreSheet) {
        scoreSheet = doc.insertSheet('SCORE');
        scoreSheet.appendRow(['Timestamp', 'StudentID', 'Name', 'Q_ID', 'Score', 'Feedback', 'Type', 'EmailStatus']);
      }
      var headerRange = scoreSheet.getRange(1, 1, 1, 8);
      var headerValues = headerRange.getValues();
      if (!headerValues[0][0] || headerValues[0][0] !== 'Timestamp') {
        scoreSheet.insertRowBefore(1);
        scoreSheet.getRange(1, 1, 1, 8).setValues([['Timestamp', 'StudentID', 'Name', 'Q_ID', 'Score', 'Feedback', 'Type', 'EmailStatus']]);
      }
      scoreSheet.appendRow([
        data.timestamp || '',
        data.sid || '',
        data.sname || '',
        data.qId || '',
        data.score != null ? data.score : '',
        data.feedback || '',
        data.type || '',
        data.emailStatus || ''
      ]);
      return ContentService.createTextOutput(JSON.stringify({ "result": "success", "action": "saved" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 답안 제출: answer 시트에 저장
    var sheet = doc.getSheetByName('answer');
    if (!sheet) {
      sheet = doc.insertSheet('answer');
      sheet.appendRow(['Timestamp', 'StudentID', 'Name', 'Email', 'Q_ID', 'Answer', 'Group']);
    }
    var headerRange = sheet.getRange(1, 1, 1, 7);
    var headerValues = headerRange.getValues();
    if (!headerValues[0][0] || headerValues[0][0] !== 'Timestamp') {
      sheet.insertRowBefore(1);
      sheet.getRange(1, 1, 1, 7).setValues([['Timestamp', 'StudentID', 'Name', 'Email', 'Q_ID', 'Answer', 'Group']]);
    }

    var quizSheet = doc.getSheetByName('Quiz');
    var group = '';
    if (quizSheet) {
      var quizData = quizSheet.getDataRange().getValues();
      for (var i = 1; i < quizData.length; i++) {
        if (quizData[i][0] == data.qId) {
          group = quizData[i][1] || '';
          break;
        }
      }
    }

    sheet.appendRow([
      data.timestamp,
      data.sid,
      data.sname,
      data.email,
      data.qId,
      data.answer,
      group
    ]);

    return ContentService.createTextOutput(JSON.stringify({ "result": "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ "result": "error", "error": err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService.createTextOutput("Smart Grading System Server is Running.");
}
