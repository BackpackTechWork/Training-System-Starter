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
  const lastRow = sheet.getLastRow();
  
  const dataRange = sheet.getRange(5, 6, lastRow - 4, 5);
  const values = dataRange.getValues();
  
  const data = [];
  for (let i = 0; i < values.length; i++) {
    const healthcareName = values[i][0];
    const deviceSerialNumber = values[i][1];
    const areaOfSpecialization = values[i][3];
    const kLaserModel = values[i][4];
    
    if (healthcareName || deviceSerialNumber || areaOfSpecialization || kLaserModel) {
      data.push({
        healthcareName: healthcareName || "",
        deviceSerialNumber: deviceSerialNumber || "",
        areaOfSpecialization: areaOfSpecialization || "",
        kLaserModel: kLaserModel || "",
        rowIndex: i + 5
      });
    }
  }
  
  return data;
}

function updateFormHealthcareDropdown() {
  const formId = 'YOUR_TRAINEE REGISTRATION_FORM_ID';
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

function createSpreadsheetEditTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onSpreadsheetEdit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger('onSpreadsheetEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
    
  return "Edit trigger created successfully";
}

function onSpreadsheetEdit(e) {
  if (e.range.getSheet().getName() === "Settings" && 
      e.range.getColumn() === 6 && 
      e.range.getRow() >= 5) {
    
    updateFormHealthcareDropdown();
  }
}