function doGet() {
  return HtmlService.createHtmlOutputFromFile('CRUDTrainingManagement')
    .setTitle('Training Management')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getAllTrainings() {
  try {
    Logger.log("Starting getAllTrainings function");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log("Active spreadsheet name: " + ss.getName());
    
    const allSheets = ss.getSheets();
    Logger.log("All sheets in the spreadsheet:");
    allSheets.forEach(sheet => Logger.log("- " + sheet.getName()));
    
    const sheet = ss.getSheetByName('All Trainings');
    
    if (!sheet) {
      Logger.log("ERROR: 'All Trainings' sheet not found!");
      return {
        success: false,
        error: "'All Trainings' sheet not found"
      };
    }
    
    Logger.log("Found 'All Trainings' sheet");
    Logger.log("Sheet last row: " + sheet.getLastRow());
    Logger.log("Sheet last column: " + sheet.getLastColumn());
    
    // Check if there's any data beyond the header rows
    if (sheet.getLastRow() <= 3) {
      Logger.log("No training data found in the sheet");
      return {
        success: true,
        data: [],
        count: 0,
        message: "No training available"
      };
    }
    
    // Get timestamps to determine data range
    Logger.log("Attempting to get range A4:A");
    const timestampRange = sheet.getRange('A4:A');
    const timestamps = timestampRange.getValues();
    Logger.log("Raw timestamps data (first 5 rows):");
    for (let i = 0; i < Math.min(5, timestamps.length); i++) {
      Logger.log(`Row ${i+4}: ${JSON.stringify(timestamps[i])}`);
    }
    
    // Filter out empty rows
    const nonEmptyTimestamps = timestamps.filter(row => row[0] !== "");
    Logger.log(`Found ${nonEmptyTimestamps.length} non-empty timestamp rows`);
    
    // Check if there are any non-empty rows
    if (nonEmptyTimestamps.length === 0) {
      Logger.log("No training data found after filtering");
      return {
        success: true,
        data: [],
        count: 0,
        message: "No training available"
      };
    }
    
    // Calculate the last row with data
    const lastRow = nonEmptyTimestamps.length + 3; // +3 because data starts at row 4
    Logger.log(`Calculated lastRow: ${lastRow}`);
    
    // Get all columns of data
    const numRows = lastRow - 3; // Number of rows to fetch
    Logger.log(`Attempting to get range (4, 1, ${numRows}, 12)`); // Updated to include WhatsApp column (L)
    
    const data = sheet.getRange(4, 1, numRows, 12).getValues(); // Updated to include WhatsApp column (L)
    Logger.log(`Retrieved ${data.length} rows of data`);
    Logger.log("First few rows of retrieved data:");
    for (let i = 0; i < Math.min(5, data.length); i++) {
      Logger.log(`Row ${i+4}: ${JSON.stringify(data[i])}`);
    }
    
    // Get the formulas for the gradebook column to extract the actual URLs
    const formulas = sheet.getRange(4, 9, numRows, 1).getFormulas(); // Column 9 is the gradebook column
    
    // Map the data to a structured format with string conversion for dates
    const mappedData = data.map((row, index) => {
      try {
        // Convert date objects to strings to ensure they serialize properly
        const formatValue = (val) => {
          if (val instanceof Date) {
            return val.toISOString();
          }
          return val;
        };
        
        // Extract the actual URL from the HYPERLINK formula if it exists
        let gradebookLink = row[8]; // Default to the display value
        
        if (formulas[index][0]) {
          const formula = formulas[index][0];
          // Extract URL from HYPERLINK formula using regex
          const match = formula.match(/HYPERLINK\("([^"]+)"/i);
          if (match && match[1]) {
            gradebookLink = match[1]; // Use the actual URL instead of the display text
          }
        }
        
        // Ensure deviceSerialNumber is a string
        let deviceSerialNumber = row[6];
        if (deviceSerialNumber !== null && deviceSerialNumber !== undefined) {
          // Convert to string if it's not already
          deviceSerialNumber = String(deviceSerialNumber);
        }
        
        return {
          timestamp: formatValue(row[0]),
          trainingName: row[1],
          trainer: row[2],
          healthcareCentre: row[3],
          startDateTime: formatValue(row[4]),
          endDateTime: formatValue(row[5]),
          deviceSerialNumber: deviceSerialNumber,
          trainingType: row[7],
          gradebookLink: gradebookLink,
          trainingStatus: row[9],
          whatsappLink: row[11], // Updated to column L (index 11)
          rowIndex: index + 4 // Store the actual row index for updating later
        };
      } catch (error) {
        Logger.log(`Error mapping row ${index + 4}: ${error.toString()}`);
        // Return a basic object with the row index to avoid breaking the data structure
        return {
          rowIndex: index + 4,
          error: error.toString(),
          timestamp: row[0] instanceof Date ? row[0].toISOString() : row[0],
          trainingName: row[1] || "",
          deviceSerialNumber: row[6] ? String(row[6]) : ""
        };
      }
    });
    
    Logger.log(`Final mapped data has ${mappedData.length} items`);
    if (mappedData.length > 0) {
      Logger.log("First item in mapped data:");
      Logger.log(JSON.stringify(mappedData[0]));
    }
    
    // Return a plain object that can be easily serialized
    return {
      success: true,
      data: mappedData,
      count: mappedData.length
    };
    
  } catch (error) {
    Logger.log("ERROR in getAllTrainings: " + error.toString());
    Logger.log("Stack trace: " + error.stack);
    
    // Return error information that can be easily serialized
    return {
      success: false,
      error: error.toString(),
      stack: error.stack
    };
  }
}

function getParticipatingTrainees() {
  try {
    Logger.log("Starting getParticipatingTrainees function");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Participating Trainees');
    
    if (!sheet) {
      Logger.log("ERROR: 'Participating Trainees' sheet not found!");
      return {
        success: false,
        error: "'Participating Trainees' sheet not found"
      };
    }
    
    Logger.log("Found 'Participating Trainees' sheet");
    Logger.log("Sheet last row: " + sheet.getLastRow());
    Logger.log("Sheet last column: " + sheet.getLastColumn());
    
    // Check if there's any data beyond the header rows
    if (sheet.getLastRow() <= 3) {
      Logger.log("No trainee data found in the sheet");
      return {
        success: true,
        data: [],
        count: 0,
        message: "No training available"
      };
    }
    
    // Get timestamps to determine data range
    Logger.log("Attempting to get range A4:A");
    const timestampRange = sheet.getRange('A4:A');
    const timestamps = timestampRange.getValues();
    
    // Filter out empty rows
    const nonEmptyTimestamps = timestamps.filter(row => row[0] !== "");
    Logger.log(`Found ${nonEmptyTimestamps.length} non-empty timestamp rows`);
    
    // Check if there are any non-empty rows
    if (nonEmptyTimestamps.length === 0) {
      Logger.log("No trainee data found after filtering");
      return {
        success: true,
        data: [],
        count: 0,
        message: "No training available"
      };
    }
    
    // Calculate the last row with data
    const lastRow = nonEmptyTimestamps.length + 3; // +3 because data starts at row 4
    Logger.log(`Calculated lastRow: ${lastRow}`);
    
    // Get all columns of data
    const numRows = lastRow - 3; // Number of rows to fetch
    Logger.log(`Attempting to get range (4, 1, ${numRows}, 9)`); // Updated to include Remarks (column I)
    
    const data = sheet.getRange(4, 1, numRows, 9).getValues(); // Updated to include Remarks (column I)
    Logger.log(`Retrieved ${data.length} rows of data`);
    
    // Get the formulas for the gradebook column to extract the actual URLs
    const formulas = sheet.getRange(4, 6, numRows, 1).getFormulas(); // Column 6 is the gradebook column
    
    // Map the data to a structured format
    const mappedData = data.map((row, index) => {
      try {
        // Convert date objects to strings to ensure they serialize properly
        const formatValue = (val) => {
          if (val instanceof Date) {
            return val.toISOString();
          }
          return val;
        };
        
        // Extract the actual URL from the HYPERLINK formula if it exists
        let gradebookLink = row[5]; // Default to the display value
        
        if (formulas[index][0]) {
          const formula = formulas[index][0];
          // Extract URL from HYPERLINK formula using regex
          const match = formula.match(/HYPERLINK\("([^"]+)"/i);
          if (match && match[1]) {
            gradebookLink = match[1]; // Use the actual URL instead of the display text
          }
        }
        
        // Set grade to "Ungraded" if missing
        const grade = row[6] ? row[6] : "Ungraded";
        
        return {
          timestamp: formatValue(row[0]),
          traineeName: row[1],
          traineeId: row[2],
          icPassport: row[3],
          trainingName: row[4],
          gradebookLink: gradebookLink,
          grade: grade,
          affiliatedHealthcare: row[7], // Added Affiliated Healthcare
          remarks: row[8] || "", // Added Remarks from column I
          rowIndex: index + 4 // Store the actual row index for updating later
        };
      } catch (error) {
        Logger.log(`Error mapping trainee row ${index + 4}: ${error.toString()}`);
        // Return a basic object with the row index to avoid breaking the data structure
        return {
          rowIndex: index + 4,
          error: error.toString(),
          timestamp: row[0] instanceof Date ? row[0].toISOString() : row[0],
          traineeName: row[1] || "",
          trainingName: row[4] || ""
        };
      }
    });
    
    Logger.log(`Final mapped data has ${mappedData.length} items`);
    
    // Return a plain object that can be easily serialized
    return {
      success: true,
      data: mappedData,
      count: mappedData.length
    };
    
  } catch (error) {
    Logger.log("ERROR in getParticipatingTrainees: " + error.toString());
    Logger.log("Stack trace: " + error.stack);
    
    // Return error information that can be easily serialized
    return {
      success: false,
      error: error.toString(),
      stack: error.stack
    };
  }
}

// Get dropdown options from Settings sheet
function getDropdownOptions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('Settings');
    
    if (!settingsSheet) {
      Logger.log("ERROR: 'Settings' sheet not found!");
      return {
        success: false,
        error: "Settings sheet not found"
      };
    }
    
    // Get trainer options (D4:D)
    const trainerRange = settingsSheet.getRange('D4:D');
    const trainerValues = trainerRange.getValues();
    const trainers = trainerValues
      .map(row => row[0])
      .filter(value => value !== "");
    
    // Get healthcare centre options (F5:F)
    const healthcareCentreRange = settingsSheet.getRange('F5:F');
    const healthcareCentreValues = healthcareCentreRange.getValues();
    const healthcareCentres = healthcareCentreValues
      .map(row => row[0])
      .filter(value => value !== "");
    
    // Get device serial number options (G5:G)
    const deviceSerialRange = settingsSheet.getRange('G5:G');
    const deviceSerialValues = deviceSerialRange.getValues();
    const deviceSerials = deviceSerialValues
      .map(row => row[0])
      .filter(value => value !== "");
    
    // Get authorized emails for status editing (L5:L)
    const authorizedEmailsRange = settingsSheet.getRange('L5:L');
    const authorizedEmailsValues = authorizedEmailsRange.getValues();
    const authorizedEmails = authorizedEmailsValues
      .map(row => row[0])
      .filter(value => value !== "");
    
    return {
      success: true,
      options: {
        trainers: trainers,
        healthcareCentres: healthcareCentres,
        deviceSerials: deviceSerials,
        authorizedEmails: authorizedEmails
      }
    };
  } catch (error) {
    Logger.log("ERROR in getDropdownOptions: " + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

// Get current user's email
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

    // Validate that rowIndex is within range
    if (rowIndex < 4 || rowIndex > sheet.getLastRow()) {
      return {
        success: false,
        error: "Invalid row index"
      };
    }

    // Get current user email
    const userEmail = Session.getActiveUser().getEmail();

    // Check if user is authorized to edit status
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

    // Update fields in All Trainings sheet (excluding training name)
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