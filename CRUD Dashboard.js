const ss = SpreadsheetApp.getActiveSpreadsheet();

 
function doGet() {
  try {
    return HtmlService.createTemplateFromFile('CRUDDashboard')
      .evaluate()
      .setTitle('LTS Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    Logger.log('Error in doGet: ' + error.toString());
     
    return HtmlService.createHtmlOutput(
      '<h1>Error Loading Dashboard</h1>' +
      '<p>There was an error loading the dashboard. Please try again later.</p>' +
      '<p>Error details: ' + error.message + '</p>'
    )
    .setTitle('LTS Dashboard - Error')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

 
function getDashboardData() {
  try {
     
    const sheets = {
      formResponses: ss.getSheetByName('Form Responses 2'),
      allTrainings: ss.getSheetByName('All Trainings'),
      settings: ss.getSheetByName('Settings')
    };
    
     
    const dataRanges = {};
    
    if (sheets.formResponses && sheets.formResponses.getLastRow() > 1) {
      dataRanges.formResponses = sheets.formResponses.getRange('B2:O' + sheets.formResponses.getLastRow()).getValues();
    }
    
    if (sheets.allTrainings && sheets.allTrainings.getLastRow() > 3) {
      const lastRow = sheets.allTrainings.getLastRow();
      dataRanges.trainingNames = sheets.allTrainings.getRange('B4:B' + lastRow).getValues().flat();
      dataRanges.trainingStatuses = sheets.allTrainings.getRange('J4:J' + lastRow).getValues().flat();
      
       
      const numRows = Math.min(10, lastRow - 3);
      if (numRows > 0) {
        dataRanges.recentTrainings = sheets.allTrainings.getRange(lastRow - numRows + 1, 2, numRows, 9).getValues();
      }
      
       
      dataRanges.trainingData = sheets.allTrainings.getRange(4, 3, lastRow - 3, 8).getValues();
    }
    
    if (sheets.settings && sheets.settings.getLastRow() > 4) {
      dataRanges.usernames = sheets.settings.getRange('D5:D' + sheets.settings.getLastRow()).getValues().flat().filter(Boolean);
    }
    
     
    return {
      notifications: safeExecute(() => getNotifications(sheets.formResponses, dataRanges.formResponses), []),
      trainings: safeExecute(() => getRecentTrainings(sheets.allTrainings, dataRanges.recentTrainings), []),
      trainingStats: safeExecute(() => getTrainingStats(dataRanges.trainingNames, dataRanges.trainingStatuses), { total: 0, inProgress: 0, completed: 0, canceled: 0, rescheduled: 0 }),
      traineeStats: safeExecute(() => getTraineeStats(sheets.formResponses, dataRanges.formResponses), { total: 0, active: 0, inactive: 0 }),
      trainerStats: safeExecute(() => getTrainerStats(dataRanges.usernames, dataRanges.trainingData), [])
    };
  } catch (error) {
    Logger.log('Error in getDashboardData: ' + error.toString());
    throw new Error('Failed to load dashboard data: ' + error.message);
  }
}

 
function safeExecute(func, defaultValue) {
  try {
    return func();
  } catch (error) {
    Logger.log('Error in ' + func.name + ': ' + error.toString());
    return defaultValue;
  }
}

 
function getNotifications(sheet, data) {
  try {
    if (!sheet || !data || data.length === 0) return [];

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const MS_PER_DAY = 1000 * 60 * 60 * 24;
    const notifications = [];

     
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const name = row[0];
      const recertDate = row[11];

      if (!name || !(recertDate instanceof Date) || isNaN(recertDate.getTime())) continue;

      const normalizedDate = new Date(recertDate);
      normalizedDate.setHours(0, 0, 0, 0);

      const daysLeft = Math.floor((normalizedDate - today) / MS_PER_DAY);

      if (daysLeft <= 60) {
        notifications.push({
          name: name || 'Unknown',
          icPassport: row[1] || 'N/A',
          traineeId: row[6] || 'N/A',
          email: row[2] || 'N/A',
          phone: row[3] || 'N/A',
          healthcareCentre: row[4] || 'N/A',
          designation: row[5] || 'N/A',
          daysLeft,
          recertDate: formatDate(recertDate)
        });
      }
    }

     
    return notifications.sort((a, b) => a.daysLeft - b.daysLeft);
  } catch (error) {
    Logger.log('Error in getNotifications: ' + error.toString());
    throw error;
  }
}

 
function getRecentTrainings(sheet, data) {
  try {
    if (!sheet || !data || data.length === 0) return [];

     
    const trainings = data
      .filter(row => row[0])
      .map(row => {
        const formattedStart = row[3] instanceof Date ? formatDateTime(row[3]) : String(row[3]);
        const formattedEnd = row[4] instanceof Date ? formatDateTime(row[4]) : String(row[4]);

        return {
          name: row[0] || 'Unnamed Training',
          trainer: row[1] || 'Unassigned',
          centre: row[2] || 'N/A',
          startDate: formattedStart || 'N/A',
          endDate: formattedEnd || 'N/A',
          deviceSerial: row[5] || 'N/A',
          type: row[6] || 'N/A',
          status: row[8] || 'Unknown'
        };
      });

    return trainings.reverse();
  } catch (error) {
    Logger.log('Error in getRecentTrainings: ' + error.toString());
    throw error;
  }
}

 
function getTrainingStats(nameCol, statusCol) {
  try {
    if (!nameCol || !statusCol) return { total: 0, inProgress: 0, completed: 0, canceled: 0, rescheduled: 0 };

     
    const stats = {
      total: nameCol.filter(n => n !== '').length,
      inProgress: 0,
      completed: 0,
      canceled: 0,
      rescheduled: 0
    };

     
    for (let i = 0; i < statusCol.length; i++) {
      const status = statusCol[i];
      if (status === 'In Progress') stats.inProgress++;
      else if (status === 'Completed') stats.completed++;
      else if (status === 'Canceled') stats.canceled++;
      else if (status === 'Rescheduled') stats.rescheduled++;
    }

    return stats;
  } catch (error) {
    Logger.log('Error in getTrainingStats: ' + error.toString());
    throw error;
  }
}

 
function getTraineeStats(sheet, data) {
  try {
    if (!sheet || !data || data.length === 0) return { total: 0, active: 0, inactive: 0 };

     
    let total = 0;
    let active = 0;
    let inactive = 0;

     
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[6] !== '') {  
        total++;
        if (row[13] === 'Active') active++;  
        else if (row[13] === 'Inactive') inactive++;
      }
    }

    return { total, active, inactive };
  } catch (error) {
    Logger.log('Error in getTraineeStats: ' + error.toString());
    throw error;
  }
}

 
function getTrainerStats(usernames, trainingData) {
  try {
    if (!usernames || !trainingData) return [];

     
    const trainerStats = {};
    
     
    usernames.forEach(username => {
      trainerStats[username] = { username, completed: 0, inProgress: 0 };
    });
    
     
    for (let i = 0; i < trainingData.length; i++) {
      const row = trainingData[i];
      const username = row[0];
      const status = row[7];
      
      if (username && trainerStats[username]) {
        if (status === 'Completed') trainerStats[username].completed++;
        else if (status === 'In Progress') trainerStats[username].inProgress++;
      }
    }
    
     
    return Object.values(trainerStats);
  } catch (error) {
    Logger.log('Error in getTrainerStats: ' + error.toString());
    throw error;
  }
}

 
function formatDate(date) {
  try {
    if (!date || !(date instanceof Date) || isNaN(date.getTime())) return 'Invalid Date';
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd-MM-yyyy');
  } catch (error) {
    Logger.log('Error in formatDate: ' + error.toString());
    return 'Date Error';
  }
}

function formatDateTime(date) {
  try {
    if (!date || !(date instanceof Date) || isNaN(date.getTime())) return 'Invalid Date';
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm');
  } catch (error) {
    Logger.log('Error in formatDateTime: ' + error.toString());
    return 'DateTime Error';
  }
}