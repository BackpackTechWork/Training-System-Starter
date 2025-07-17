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

  const getSheetQuestions = (sheet) => {
    const lastRow = sheet.getLastRow();
    if (lastRow < 4) return [];
    const numRows = lastRow - 3;
    const data = sheet.getRange(4, 2, numRows, 6).getValues();

    return data
      .filter(row => row[0] !== "")
      .map(row => ({
        question: row[0],
        options: row.slice(1, 5),
        objective: row[5]
      }));
  };

  const postTestQuestions = getSheetQuestions(postTestSheet);
  const refreshmentQuestions = getSheetQuestions(refreshmentSheet);

  return {
    postTest: postTestQuestions,
    refreshment: refreshmentQuestions
  };
}


