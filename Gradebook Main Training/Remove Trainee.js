function showAddRemoveTraineeModal() {
  const html = HtmlService.createHtmlOutputFromFile('Add/RemoveTrainee')
    .setWidth(750)  
    .setHeight(800);  
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function getTraineesForRemoval() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Post Test & Pre Test');
  const dataRange = sheet.getRange('B4:D' + sheet.getLastRow());
  const data = dataRange.getValues();
  
  const trainees = data
    .filter(row => row[0] !== '')
    .map((row, index) => {
      return {
        display: `${row[0]} - ${row[1]} - ${row[2]}`, 
        rowIndex: index + 4, 
        name: row[0],
        icPassport: row[1],
        traineeId: row[2]
      };
    });
  
  return trainees;
}

function removeTrainees(selectedTrainees) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSpreadsheetId = ss.getId();
  const sheet = ss.getSheetByName('Post Test & Pre Test');
  
   
  const formulaE4 = sheet.getRange('E4').getFormula();
  const formulaF4 = sheet.getRange('F4').getFormula();
  
  const rowsToDelete = selectedTrainees.map(t => t.rowIndex).sort((a, b) => b - a);
  const isRow4Deleted = rowsToDelete.includes(4);
  
  rowsToDelete.forEach(rowIndex => {
    sheet.deleteRow(rowIndex);
  });
  
   
  if (isRow4Deleted) {
    if (formulaE4) sheet.getRange('E4').setFormula(formulaE4);
    if (formulaF4) sheet.getRange('F4').setFormula(formulaF4);
  }
  
  const otherSS = SpreadsheetApp.openById('YOUR_MAIN_PROJECT_SHEET_ID');
  const otherSheet = otherSS.getSheetByName('Participating Trainees');
  
  const otherDataRange = otherSheet.getRange('B4:F' + otherSheet.getLastRow());
  const otherData = otherDataRange.getValues();
  const otherFormulas = otherSheet.getRange('F4:F' + otherSheet.getLastRow()).getFormulas();
  
  const rowsToDeleteFromOther = [];
  
  selectedTrainees.forEach(trainee => {
    otherData.forEach((row, index) => {
      const hyperlinkFormula = otherFormulas[index][0];
      let spreadsheetIdFromHyperlink = '';
      
      if (hyperlinkFormula) {
        const match = hyperlinkFormula.match(/spreadsheets\/d\/([^\/]+)/);
        if (match && match[1]) {
          spreadsheetIdFromHyperlink = match[1];
        }
      }
      
      if (row[0] === trainee.name && 
          row[1] === trainee.traineeId && 
          row[2] === trainee.icPassport &&
          spreadsheetIdFromHyperlink === currentSpreadsheetId) {
        rowsToDeleteFromOther.push(index + 4); 
      }
    });
  });
  
  rowsToDeleteFromOther.sort((a, b) => b - a).forEach(rowIndex => {
    otherSheet.deleteRow(rowIndex);
  });
  
  return {
    success: true,
    message: `Successfully removed ${selectedTrainees.length} trainee(s) from current sheet and ${rowsToDeleteFromOther.length} matching entries from external sheet`
  };
}



function getTraineesForAdding() {
  const externalSS = SpreadsheetApp.openById('YOUR_MAIN_PROJECT_SHEET_ID');
  const traineeDbSheet = externalSS.getSheetByName('Form Responses 2');
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Post Test & Pre Test');

  if (!traineeDbSheet || !activeSheet) {
    throw new Error('Required sheets not found.');
  }

  const lastRow = traineeDbSheet.getLastRow();
  if (lastRow < 2) return [];

  const nameRange = traineeDbSheet.getRange('B2:B' + lastRow).getValues();
  const icPassportRange = traineeDbSheet.getRange('C2:C' + lastRow).getValues();
  const traineeIdRange = traineeDbSheet.getRange('H2:H' + lastRow).getValues();

  const data = nameRange.map((nameRow, index) => [
    nameRow[0], 
    icPassportRange[index][0], 
    traineeIdRange[index][0]
  ]);

  const activeLastRow = activeSheet.getLastRow();
  const existingData = activeLastRow >= 4 ? activeSheet.getRange('D4:D' + activeLastRow).getValues() : [];
  const existingTraineeIds = new Set(existingData.map(row => row[0]));

  return data
    .filter(row => row[0] && !existingTraineeIds.has(row[2]))  
    .map(row => ({
      display: `${row[0]} - ${row[1]} - ${row[2]}`,
      name: row[0],
      icPassport: row[1],
      traineeId: row[2]
    }));
}





function addTrainees(selectedTrainees) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentSpreadsheetId = ss.getId();
    const currentSpreadsheetName = ss.getName();
    const sheet = ss.getSheetByName('Post Test & Pre Test');

    const externalSS = SpreadsheetApp.openById('YOUR_MAIN_PROJECT_SHEET_ID');
    const participatingSheet = externalSS.getSheetByName('Participating Trainees');
    const formResponsesSheet = externalSS.getSheetByName('Form Responses 2');

    const traineeIdRange = formResponsesSheet.getRange('H2:H' + formResponsesSheet.getLastRow()).getValues();
    const affiliatedHealthcareRange = formResponsesSheet.getRange('F2:F' + formResponsesSheet.getLastRow()).getValues();

    const healthcareMap = {};
    traineeIdRange.forEach((id, index) => {
      if (id[0]) {
        healthcareMap[id[0]] = affiliatedHealthcareRange[index][0];
      }
    });

    let nextRow = 4;
    while (sheet.getRange(`B${nextRow}`).getValue() !== '') {
      nextRow++;
    }

    selectedTrainees.forEach((trainee, index) => {
      const currentRow = nextRow + index;
      sheet.getRange(`B${currentRow}`).setValue(trainee.name);
      sheet.getRange(`C${currentRow}`).setValue(trainee.icPassport);
      sheet.getRange(`D${currentRow}`).setValue(trainee.traineeId);
    });

    let externalNextRow = 4;
    while (participatingSheet.getRange(`B${externalNextRow}`).getValue() !== '') {
      externalNextRow++;
    }

    const hyperlinkFormula = `=HYPERLINK("https://docs.google.com/spreadsheets/d/${currentSpreadsheetId}/edit?usp=drivesdk", "Open Gradebook")`;

    selectedTrainees.forEach((trainee, index) => {
      const currentRow = externalNextRow + index;
      const affiliatedHealthcare = healthcareMap[trainee.traineeId] || '';

      participatingSheet.getRange(`A${currentRow}`).setValue(new Date());
      participatingSheet.getRange(`B${currentRow}`).setValue(trainee.name);
      participatingSheet.getRange(`C${currentRow}`).setValue(trainee.traineeId);
      participatingSheet.getRange(`D${currentRow}`).setValue(trainee.icPassport);
      participatingSheet.getRange(`E${currentRow}`).setValue(currentSpreadsheetName);
      participatingSheet.getRange(`F${currentRow}`).setFormula(hyperlinkFormula);
      participatingSheet.getRange(`H${currentRow}`).setValue(affiliatedHealthcare);
    });

    return {
      success: true,
      message: `Successfully added ${selectedTrainees.length} trainee(s)`
    };
  } catch (error) {
    return {
      success: false,
      message: `Error: ${error.toString()}`
    };
  }
}


