function doGet() {
  return HtmlService.createTemplateFromFile('CRUDTraineeDatabase')
    .evaluate()
    .setTitle('Trainee Database')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getInitData() {
  return Promise.all([
    getTraineeData(),
    getDropdownOptions()
  ]).then(([traineeData, dropdownOptions]) => {
    return {
      traineeData,
      dropdownOptions
    };
  });
}


function getTraineeData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Form Responses 2");
  
  const lastRow = getLastRowWithData(sheet, "B");
  if (lastRow < 2) return JSON.stringify([]);
  const dataRange = sheet.getRange("B2:O" + lastRow);
  const data = dataRange.getValues();  
  const trainees = data
    .filter(row => row[0] !== "") 
    .map((row, index) => ({
      rowIndex: index + 2, 
      name: row[0] || "",
      icPassport: row[1] || "",
      traineeId: row[6] || "",
      email: row[2] || "",
      handphone: row[3] || "",
      healthcareCentre: row[4] || "",
      designation: row[5] || "",
      specialization: row[7] || "",
      deviceSerialNumber: row[8] || "",
      firstTraining: row[9] || "",
      latestTraining: row[10] || "",
      recertificationDate: row[11] || "",
      completedTrainings: row[12] || "",
      status: row[13] || ""
    }));

    console.log(trainees)
  return JSON.stringify(trainees);
}

function getLastRowWithData(sheet, column) {
  const values = sheet.getRange(column + "1:" + column + sheet.getMaxRows()).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") {
      return i + 1;
    }
  }
  return 0;
}

function getDropdownOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("Settings");
  
  
  const dataRange = settingsSheet.getRange("F5:I" + settingsSheet.getLastRow()).getValues();
  
  
  const healthcareCentreOptions = [];
  const deviceSerialOptions = [];
  const specializationOptions = [];
  
  dataRange.forEach(row => {
    if (row[0] !== "") healthcareCentreOptions.push(row[0]);
    if (row[1] !== "") deviceSerialOptions.push(row[1]);
    if (row[3] !== "") specializationOptions.push(row[3]);
  });

  return JSON.stringify({
    healthcareCentres: healthcareCentreOptions,
    deviceSerials: deviceSerialOptions,
    specializations: specializationOptions
  });
}

function generateUniqueTraineeId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Form Responses 2");
  
  
  const idColumn = sheet.getRange("H2:H" + sheet.getLastRow()).getValues();
  const existingIds = new Set(idColumn.flat().filter(id => id !== ""));
  
  let newId;
  do {
    
    const randomDigits = Math.floor(100000 + Math.random() * 900000);
    newId = "T" + randomDigits;
  } while (existingIds.has(newId));
  
  return newId;
}

function updateTraineeData(traineeData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Form Responses 2");
    const rowIndex = traineeData.rowIndex;

    if (traineeData.status === "Active" && (!traineeData.traineeId || traineeData.traineeId === "")) {
      traineeData.traineeId = generateUniqueTraineeId();
    }

    const existingRow = sheet.getRange(rowIndex, 2, 1, 14).getValues()[0];

    const updateValues = [
      [
        traineeData.name,
        traineeData.icPassport,
        traineeData.email,
        traineeData.handphone,
        traineeData.healthcareCentre,
        traineeData.designation,
        traineeData.traineeId,
        traineeData.specialization,
        traineeData.deviceSerialNumber,
        existingRow[9],  
        existingRow[10], 
        existingRow[11], 
        rowIndex === 2 
          ? `=ARRAYFORMULA(IF(H2:H="", "", COUNTIFS('Participating Trainees'!C:C, H2:H)))` 
          : "", 
        traineeData.status
      ]
    ];

    const updateRange = sheet.getRange(rowIndex, 2, 1, 14);
    updateRange.setValues(updateValues);

    return JSON.stringify({ 
      success: true, 
      message: "Trainee data updated successfully",
      traineeId: traineeData.traineeId
    });
  } catch (error) {
    console.error("Error updating trainee data:", error);
    return JSON.stringify({ success: false, message: "Error updating trainee data: " + error.toString() });
  }
}

function getTrainingHistory(traineeId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const participatingSheet = ss.getSheetByName("Participating Trainees");
    const allTrainingsSheet = ss.getSheetByName("All Trainings");

    if (!participatingSheet || !allTrainingsSheet) {
      return JSON.stringify({
        success: false,
        message: "Required sheets not found"
      });
    }


    const participatingLastRow = participatingSheet.getLastRow();
    const participatingRange = participatingSheet.getRange("C4:H" + participatingLastRow);
    const participatingValues = participatingRange.getValues();
    const participatingFormulas = participatingRange.getFormulas();

    const allTrainingsLastRow = allTrainingsSheet.getLastRow();
    const allTrainingsRange = allTrainingsSheet.getRange("C4:I" + allTrainingsLastRow);
    const trainingValues = allTrainingsRange.getValues();
    const trainingFormulas = allTrainingsRange.getFormulas();

    const trainingMap = new Map();

    trainingValues.forEach((row, i) => {
      const trainer = row[0];
      const trainingId = row[1];
      const date = row[3];
      const trainingType = row[5];

      const formula = trainingFormulas[i][6];

      let url = null;
      const urlMatch = typeof formula === "string" ? formula.match(/HYPERLINK\("([^"]+)"/) : null;
      if (urlMatch) url = urlMatch[1];

      const formattedDate = date instanceof Date
        ? Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd")
        : "Unknown";

      const details = {
        trainer: trainer || "Unknown",
        trainingType: trainingType || "Unknown",
        date: formattedDate
      };

      if (url) trainingMap.set(url, details);
      if (trainingId) trainingMap.set(trainingId, details);
    });

    const trainingHistory = [];

    participatingValues.forEach((row, i) => {
      const rowTraineeId = row[0];
      if (rowTraineeId === traineeId) {
        const trainingId = row[2];

        const gradebookFormula = participatingFormulas[i][3];
        const grade = row[4];
        const affiliatedHealthcare = row[5];

        let gradebookLinkUrl = null;
        const match = typeof gradebookFormula === "string" ? gradebookFormula.match(/HYPERLINK\("([^"]+)"/) : null;
        if (match) gradebookLinkUrl = match[1];

        const trainingDetails = trainingMap.get(gradebookLinkUrl) || trainingMap.get(trainingId) || {
          trainer: "Not found",
          trainingType: "Not found",
          date: "Unknown"
        };

        trainingHistory.push({
          trainingId: trainingId,
          grade: grade,
          affiliatedHealthcare: affiliatedHealthcare,
          trainer: trainingDetails.trainer,
          trainingType: trainingDetails.trainingType,
          date: trainingDetails.date
        });
      }
    });

    trainingHistory.sort((a, b) => {
      const dateA = new Date(a.date);
      const dateB = new Date(b.date);

      if (!isNaN(dateA.getTime()) && !isNaN(dateB.getTime())) {
        return dateB.getTime() - dateA.getTime();
      }
      if (!isNaN(dateA.getTime())) {
        return -1;
      }
      if (!isNaN(dateB.getTime())) {
        return 1;
      }
      return 0;
    });

    return JSON.stringify({
      success: true,
      data: trainingHistory
    });

  } catch (error) {
    return JSON.stringify({
      success: false,
      message: "Error: " + error.toString()
    });
  }
}
