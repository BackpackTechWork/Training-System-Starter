function doGet() {
  return HtmlService.createHtmlOutputFromFile('CRUDTrainingManagement')
    .setTitle('Training Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getAllInitialData() {
  Logger.log("Starting getAllInitialData function");

  return Promise.allSettled([
    Promise.resolve().then(getAllTrainings),
    Promise.resolve().then(getParticipatingTrainees),
    Promise.resolve().then(getDropdownOptions),
    Promise.resolve().then(getCurrentUserEmail)
  ]).then(results => {
    const response = {
      success: true,
      data: {
        trainings: null,
        trainees: null,
        dropdownOptions: null,
        currentUserEmail: null
      },
      errors: {}
    };

    const keys = ['trainings', 'trainees', 'dropdownOptions', 'currentUserEmail'];

    results.forEach((result, index) => {
      const key = keys[index];
      if (result.status === 'fulfilled') {
        response.data[key] = result.value;
      } else {
        Logger.log(`Error getting ${key}: ${result.reason}`);
        response.errors[key] = result.reason.toString();
        response.data[key] = { success: false, error: result.reason.toString() };
      }
    });

    Logger.log("getAllInitialData completed");
    return response;
  }).catch(error => {
    Logger.log("ERROR in getAllInitialData: " + error.toString());
    return {
      success: false,
      error: error.toString(),
      stack: error.stack
    };
  });
}


function getAllTrainings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('All Trainings');
    if (!sheet) return { success: false, error: "'All Trainings' sheet not found" };

    const lastRow = sheet.getLastRow();
    if (lastRow <= 3) return { success: true, data: [], count: 0, message: "No training available" };

    const dataRange = sheet.getRange(4, 1, lastRow - 3, 12);
    const data = dataRange.getValues();
    const formulas = sheet.getRange(4, 9, lastRow - 3, 1).getFormulas();

    const mappedData = data.reduce((acc, row, index) => {
      if (!row[0]) return acc; 

      const gradebookFormula = formulas[index][0];
      const gradebookLink = gradebookFormula
        ? (gradebookFormula.match(/HYPERLINK\("([^"]+)"/i) || [])[1] || row[8]
        : row[8];

      acc.push({
        timestamp: row[0] instanceof Date ? row[0].toISOString() : row[0],
        trainingName: row[1],
        trainer: row[2],
        healthcareCentre: row[3],
        startDateTime: row[4] instanceof Date ? row[4].toISOString() : row[4],
        endDateTime: row[5] instanceof Date ? row[5].toISOString() : row[5],
        deviceSerialNumber: String(row[6] || ''),
        trainingType: row[7],
        gradebookLink: gradebookLink,
        trainingStatus: row[9],
        whatsappLink: row[11],
        rowIndex: index + 4
      });
      return acc;
    }, []);

    return { success: true, data: mappedData, count: mappedData.length };

  } catch (error) {
    return { success: false, error: error.toString(), stack: error.stack };
  }
}


function getParticipatingTrainees() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Participating Trainees');
    if (!sheet) return { success: false, error: "'Participating Trainees' sheet not found" };

    const lastRow = sheet.getLastRow();
    if (lastRow <= 3) return { success: true, data: [], count: 0, message: "No training available" };

    const dataRange = sheet.getRange(4, 1, lastRow - 3, 9).getValues();
    const formulas = sheet.getRange(4, 6, lastRow - 3, 1).getFormulas();

    const mappedData = dataRange.reduce((acc, row, index) => {
      if (!row[0]) return acc; 

      const formula = formulas[index][0];
      const gradebookLink = formula
        ? (formula.match(/HYPERLINK\("([^"]+)"/i) || [])[1] || row[5]
        : row[5];

      acc.push({
        timestamp: row[0] instanceof Date ? row[0].toISOString() : row[0],
        traineeName: row[1],
        traineeId: row[2],
        icPassport: row[3],
        trainingName: row[4],
        gradebookLink: gradebookLink,
        grade: row[6] || "Ungraded",
        affiliatedHealthcare: row[7],
        remarks: row[8] || "",
        rowIndex: index + 4
      });
      return acc;
    }, []);

    return { success: true, data: mappedData, count: mappedData.length };

  } catch (error) {
    return { success: false, error: error.toString(), stack: error.stack };
  }
}



function getDropdownOptions() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("dropdownOptions");


  if (cached) {
    return JSON.parse(cached);
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Settings');
    if (!sheet) return { success: false, error: "'Settings' sheet not found" };

    const lastRow = sheet.getLastRow();


    const values = sheet.getRange(4, 4, lastRow - 3, 9).getValues(); 

    const trainers = values.map(row => row[0]).filter(val => val !== "");
    const healthcareCentres = values.map(row => row[2]).filter(val => val !== "");
    const deviceSerials = values.map(row => row[3]).filter(val => val !== "");
    const authorizedEmails = values.map(row => row[8]).filter(val => val !== "");

    const result = {
      success: true,
      options: {
        trainers,
        healthcareCentres,
        deviceSerials,
        authorizedEmails
      }
    };


    cache.put("dropdownOptions", JSON.stringify(result), 300);

    return result;

  } catch (error) {
    return { success: false, error: error.toString() };
  }
}



function getCurrentUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (error) {
    Logger.log("ERROR in getCurrentUserEmail: " + error.toString());
    return "unknown@example.com";
  }
}

function updateTraining(rowIndex, updatedData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('All Trainings');
    
    if (!sheet) {
      return {
        success: false,
        error: "All Trainings sheet not found"
      };
    }


    if (rowIndex < 4 || rowIndex > sheet.getLastRow()) {
      return {
        success: false,
        error: "Invalid row index"
      };
    }


    const userEmail = Session.getActiveUser().getEmail();


    let canEditStatus = false;
    if (updatedData.trainingStatus) {
      const settingsSheet = ss.getSheetByName('Settings');
      if (settingsSheet) {
        const authorizedEmailsRange = settingsSheet.getRange('L5:L');
        const authorizedEmailsValues = authorizedEmailsRange.getValues();
        const authorizedEmails = authorizedEmailsValues
          .map(row => row[0])
          .filter(value => value !== "");
        
        canEditStatus = authorizedEmails.includes(userEmail);
      }

      if (!canEditStatus) {
        return {
          success: false,
          error: "You are not authorized to edit the training status"
        };
      }
    }


    if (updatedData.trainer !== undefined) {
      sheet.getRange(rowIndex, 3).setValue(updatedData.trainer);
    }

    if (updatedData.healthcareCentre !== undefined) {
      sheet.getRange(rowIndex, 4).setValue(updatedData.healthcareCentre);
    }

    if (updatedData.startDateTime !== undefined) {
      const startDate = new Date(updatedData.startDateTime);
      sheet.getRange(rowIndex, 5).setValue(startDate);
    }

    if (updatedData.endDateTime !== undefined) {
      const endDate = new Date(updatedData.endDateTime);
      sheet.getRange(rowIndex, 6).setValue(endDate);
    }

    if (updatedData.deviceSerialNumber !== undefined) {
      sheet.getRange(rowIndex, 7).setValue(updatedData.deviceSerialNumber);
    }

    if (updatedData.trainingStatus !== undefined && canEditStatus) {
      sheet.getRange(rowIndex, 10).setValue(updatedData.trainingStatus);
    }

    if (updatedData.whatsappLink !== undefined) {
      sheet.getRange(rowIndex, 12).setValue(updatedData.whatsappLink);
    }

    return {
      success: true,
      message: "Training updated successfully"
    };

  } catch (error) {
    Logger.log("ERROR in updateTraining: " + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}
