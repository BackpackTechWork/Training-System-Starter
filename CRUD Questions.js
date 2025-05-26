function doGet() {
  return HtmlService.createHtmlOutputFromFile('CRUDQuestions')
    .setTitle('Question Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
}

function getQuestions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const postTestSheet = ss.getSheetByName('Post-Test Questions');
  const refreshmentSheet = ss.getSheetByName('Refreshment Training Questions');
  
  // Get Post-Test Questions
  const postTestLastRow = getLastRowWithData(postTestSheet, 'B');
  const postTestQuestions = postTestSheet.getRange('B4:B' + postTestLastRow).getValues();
  const postTestOptions = postTestSheet.getRange('C4:F' + postTestLastRow).getValues(); // Updated to C4:F
  const postTestObjectives = postTestSheet.getRange('G4:G' + postTestLastRow).getValues();
  
  // Get Refreshment Training Questions
  const refreshmentLastRow = getLastRowWithData(refreshmentSheet, 'B');
  const refreshmentQuestions = refreshmentSheet.getRange('B4:B' + refreshmentLastRow).getValues();
  const refreshmentOptions = refreshmentSheet.getRange('C4:F' + refreshmentLastRow).getValues(); // Updated to C4:F
  const refreshmentObjectives = refreshmentSheet.getRange('G4:G' + refreshmentLastRow).getValues();
  
  // Format data for frontend
  const postTestData = formatQuestionData(postTestQuestions, postTestOptions, postTestObjectives);
  const refreshmentData = formatQuestionData(refreshmentQuestions, refreshmentOptions, refreshmentObjectives);
  
  return {
    postTest: postTestData,
    refreshment: refreshmentData
  };
}

function getLastRowWithData(sheet, column) {
  const values = sheet.getRange(column + '1:' + column + sheet.getLastRow()).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== '') {
      return i + 1; // +1 because array is 0-indexed
    }
  }
  return 1;
}

function formatQuestionData(questions, options, objectives) {
  const formattedData = [];
  
  for (let i = 0; i < questions.length; i++) {
    if (questions[i][0] !== '') {
      formattedData.push({
        question: questions[i][0],
        options: options[i],
        objective: objectives[i][0]
      });
    }
  }
  
  return formattedData;
}