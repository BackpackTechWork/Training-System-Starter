function doGet() {
  const userEmail = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("Settings");
  
  const hasEditAccess = settingsSheet.getRange("L5:L" + settingsSheet.getLastRow())
    .getValues()
    .flat()
    .includes(userEmail);
  
  const template = HtmlService.createTemplateFromFile('CRUDSettings');
  template.hasEditAccess = hasEditAccess;
  
  return template.evaluate()
    .setTitle('Settings Manager')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSettingsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Settings");
  const lastRow = sheet.getRange("F:F").getValues().filter(String).length + 4;
  if (lastRow <= 4) return [];
  const values = sheet.getRange(5, 6, lastRow - 4, 5).getValues();
  return values
    .map((row, index) => ({
      healthcareName: row[0] || "",
      deviceSerialNumber: row[1] || "",
      areaOfSpecialization: row[3] || "",
      kLaserModel: row[4] || "",
      rowIndex: index + 5
    }))
    .filter(item => item.healthcareName || item.deviceSerialNumber || item.areaOfSpecialization || item.kLaserModel);
}


function updateFormHealthcareDropdown() {
  const formId = '1y8gBrxuKcTeppkMekkrK0MzbRRrEmO4MR9apXAkAFpA';
  const form = FormApp.openById(formId);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Settings");
  const lastRow = sheet.getLastRow();
  
  const healthcareNames = sheet.getRange(5, 6, lastRow - 4, 1)
    .getValues()
    .filter(row => row[0] !== "")
    .map(row => row[0]);
  
  const items = form.getItems();
  let healthcareDropdown = null;
  
  for (let i = 0; i < items.length; i++) {
    if (items[i].getTitle() === 'Healthcare') {
      healthcareDropdown = items[i].asListItem();
      break;
    }
  }
  
  if (healthcareDropdown) {
    healthcareDropdown.setChoiceValues(healthcareNames);
    return {
      success: true,
      message: `Updated Healthcare dropdown with ${healthcareNames.length} options`
    };
  } else {
    return {
      success: false,
      message: "Could not find a question titled 'Healthcare' in the form"
    };
  }
}

function getColumnInfo(type) {
  const columnMap = {
    "Healthcare Name": { letter: "F", index: 6, updateForm: true },
    "Device Serial Number": { letter: "G", index: 7, updateForm: false },
    "Area of Specialization": { letter: "I", index: 9, updateForm: false },
    "K-Laser Model": { letter: "J", index: 10, updateForm: false }
  };
  
  const columnInfo = columnMap[type];
  if (!columnInfo) {
    throw new Error("Invalid setting type");
  }
  
  return columnInfo;
}

function addSetting(type, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Settings");
  const columnInfo = getColumnInfo(type);
  
  const columnValues = sheet.getRange(5, columnInfo.index, sheet.getLastRow() - 4, 1).getValues();
  let emptyRowIndex = -1;
  
  for (let i = 0; i < columnValues.length; i++) {
    if (columnValues[i][0] === "") {
      emptyRowIndex = i + 5;
      break;
    }
  }
  
  if (emptyRowIndex === -1) {
    emptyRowIndex = columnValues.length + 5;
  }
  
  sheet.getRange(emptyRowIndex, columnInfo.index).setValue(value);
  
  let formUpdateResult = { success: true };
  if (columnInfo.updateForm) {
    formUpdateResult = updateFormHealthcareDropdown();
  }
  
  return {
    success: true,
    message: `Added "${value}" to ${type} at row ${emptyRowIndex}`,
    formUpdate: columnInfo.updateForm ? formUpdateResult : null
  };
}

function updateSetting(rowIndex, columnType, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Settings");
  const columnInfo = getColumnInfo(columnType);
  
  sheet.getRange(rowIndex, columnInfo.index).setValue(value);
  
  let formUpdateResult = { success: true };
  if (columnInfo.updateForm) {
    formUpdateResult = updateFormHealthcareDropdown();
  }
  
  return {
    success: true,
    message: `Updated ${columnType} at row ${rowIndex} to "${value}"`,
    formUpdate: columnInfo.updateForm ? formUpdateResult : null
  };
}

function deleteSetting(rowIndex, columnType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Settings");
  const columnInfo = getColumnInfo(columnType);
  
  const columnValues = sheet.getRange(5, columnInfo.index, sheet.getLastRow() - 4, 1).getValues();
  let lastRowWithData = 4;
  
  for (let i = columnValues.length - 1; i >= 0; i--) {
    if (columnValues[i][0] !== "") {
      lastRowWithData = i + 5;
      break;
    }
  }
  
  if (parseInt(rowIndex) > lastRowWithData) {
    return {
      success: false,
      message: "No data to delete at this position"
    };
  }
  
  const startRow = parseInt(rowIndex);
  const numRows = lastRowWithData - startRow + 1;
  
  if (numRows <= 1) {
    sheet.getRange(rowIndex, columnInfo.index).clearContent();
  } else {
    const rangeToShift = sheet.getRange(startRow + 1, columnInfo.index, lastRowWithData - startRow, 1);
    const valuesToShift = rangeToShift.getValues();
    
    const targetRange = sheet.getRange(startRow, columnInfo.index, lastRowWithData - startRow, 1);
    targetRange.setValues(valuesToShift.concat([[""]]));
  }
  
  let formUpdateResult = { success: true };
  if (columnInfo.updateForm) {
    formUpdateResult = updateFormHealthcareDropdown();
  }
  
  return {
    success: true,
    message: `Deleted ${columnType} at row ${rowIndex} and shifted values up`,
    formUpdate: columnInfo.updateForm ? formUpdateResult : null
  };
}
