function doGet() {
  return HtmlService.createTemplateFromFile('RefreshmentTrainingQuestion')
    .evaluate()
    .setTitle('Add Refreshment Training Question')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


function submitRefQuestion(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Refreshment Training Questions');
  var lastRow = sheet.getLastRow();
  var newRow = lastRow + 1;

  
  sheet.getRange(newRow, 2, 1, 6).setFontColor(null);

  
  sheet.getRange(newRow, 2).setValue(data.question);
  sheet.getRange(newRow, 3).setValue(data.option1);
  sheet.getRange(newRow, 4).setValue(data.option2);
  sheet.getRange(newRow, 5).setValue(data.option3);
  sheet.getRange(newRow, 6).setValue(data.option4);
  sheet.getRange(newRow, 7).setValue(data.objective);

  
  var correctAnswerColumn = {
    'option1': 3,
    'option2': 4,
    'option3': 5,
    'option4': 6
  }[data.correctAnswer];

  var correctAnswerCell = sheet.getRange(newRow, correctAnswerColumn);
  correctAnswerCell.setValue('*' + correctAnswerCell.getValue());
  correctAnswerCell.setFontColor('green');

  return 'Question submitted successfully!';
}