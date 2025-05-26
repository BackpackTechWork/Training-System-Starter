let submissionQueue = [];
let isProcessing = false;
const scriptLock = LockService.getScriptLock();

function doGet() {
  return HtmlService.createTemplateFromFile('TrainingCreation')
    .evaluate()
    .setTitle('Training Creation Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


function getFormOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  const traineeSheet = ss.getSheetByName('Form Responses 2');

  const [healthcareCentres, deviceSerialNumbers] = ['F5:F', 'G5:G'].map(range => 
    settingsSheet.getRange(range).getValues().flat().filter(String)
  );

  const lastRow = traineeSheet.getLastRow();
  const traineeData = traineeSheet.getRange('B2:H' + lastRow).getValues();

  // Instead of concatenating with hyphens, create objects with named properties
  const trainees = traineeData
    .filter(row => row[0] && row[1] && row[6])
    .map(row => ({
      name: row[0],
      icPassport: row[1],
      id: row[6],
      healthcare: row[4] || ''
    }));

  return { healthcareCentres, deviceSerialNumbers, trainees };
}


function getTrainerName() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('Settings');
  const data = settingsSheet.getRange('D5:E').getValues();
  
  const trainer = data.find(row => row[1] === email);
  return trainer ? trainer[0] : 'Unauthorised User';
}


function queueFormSubmission(formData) {
  return new Promise((resolve, reject) => {
    submissionQueue.push({ formData, resolve, reject });
    processQueue();
  });
}

function processQueue() {
  if (isProcessing || submissionQueue.length === 0) return;
  
  // Try to acquire the lock with a timeout of 30 seconds
  try {
    // Only proceed if we can get the lock
    if (scriptLock.tryLock(30000)) {
      isProcessing = true;
      const { formData, resolve, reject } = submissionQueue.shift();

      submitForm(formData)
        .then(result => {
          resolve(result);
        })
        .catch(error => {
          reject(error);
        })
        .finally(() => {
          isProcessing = false;
          // Release the lock
          scriptLock.releaseLock();
          // Process the next item in the queue
          processQueue(); 
        });
    } else {
      // If we couldn't get the lock, wait and try again
      Utilities.sleep(1000);
      processQueue();
    }
  } catch (e) {
    console.error('Lock error:', e);
    // Make sure we release the lock if there's an error
    if (scriptLock.hasLock()) {
      scriptLock.releaseLock();
    }
    isProcessing = false;
  }
}

function submitForm(formData) {
  return new Promise((resolve, reject) => {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const allTrainingsSheet = ss.getSheetByName('All Trainings');
      const participatingTraineesSheet = ss.getSheetByName('Participating Trainees');

      // Get the serial number first
      const lastRow = allTrainingsSheet.getLastRow();
      const serialNumber = (lastRow >= 3 ? lastRow - 2 : 1).toString().padStart(5, '0');
      // Create the TRN+ formatted serial number
      const formattedSerialNumber = "TRN" + serialNumber;

      const gradebookPromise = createGradebook(formattedSerialNumber, formData.participatingTrainees, formData.trainingType);

      gradebookPromise.then(gradebookUrl => {
        const timestamp = new Date(); // Use the same timestamp for all entries
        
        // Create the row for All Trainings sheet
        const allTrainingsRow = [
          timestamp,
          formattedSerialNumber, // Using the formatted serial number as training name
          formData.trainer,
          formData.healthcareCentres.join(', '),
          formatDateTimeForSheet(formData.startDate),
          formatDateTimeForSheet(formData.endDate),
          formData.deviceSerialNumbers.join(', '),
          formData.trainingType,
          `=HYPERLINK("${gradebookUrl}", "Open Gradebook")`,
          "In Progress", // Column J
          serialNumber // Column K - original serial number without TRN+ prefix
        ];

        // Append to All Trainings sheet
        allTrainingsSheet.appendRow(allTrainingsRow);
        
        // Prepare data for Participating Trainees sheet - batch operation
        const participatingTraineesData = [];
        
        // Process each trainee and prepare the data
        formData.participatingTrainees.forEach(trainee => {
          const parts = trainee.split(' - ');
          const name = parts[0];
          const icPassport = parts[1];
          const id = parts[2];
          const healthcare = parts[3] || '';
          
          participatingTraineesData.push([
            timestamp,           
            name,                
            id,
            icPassport,
            formattedSerialNumber, // Using the formatted serial number as training name
            `=HYPERLINK("${gradebookUrl}", "Open Gradebook")`,
            '', // Column G (empty)
            healthcare // Column H (healthcare)
          ]);
        });
        
        // If we have trainee data, write it all at once (much faster than appendRow)
        if (participatingTraineesData.length > 0) {
          const lastTraineeRow = participatingTraineesSheet.getLastRow();
          const startRow = lastTraineeRow + 1;
          participatingTraineesSheet.getRange(
            startRow, 
            1, 
            participatingTraineesData.length, 
            participatingTraineesData[0].length
          ).setValues(participatingTraineesData);
        }

        // Add the formatted serial number to formData for emails
        const updatedFormData = {...formData, trainingName: formattedSerialNumber};

        // Send emails in parallel
        Promise.all([
          sendEmailToTrainer(updatedFormData, gradebookUrl)
        ]).then(() => {
          resolve({ 
            success: true, 
            message: "Form submitted successfully!",
            gradebookUrl: gradebookUrl
          });
        });
      });
    } catch (error) {
      reject(error);
    }
  });
}


function createGradebook(trainingName, trainees, trainingType) {
  return new Promise((resolve, reject) => {
    try {
      let templateId;
      if (trainingType === "Main Training") {
        templateId = 'YOUR_MAIN_TRAINING_TEMPLATE_ID';
      } else if (trainingType === "Refreshment Training") {
        templateId = 'YOUR_REFRESHMENT_TRAINING_TEMPLATE_ID';
      } else {
        throw new Error("Invalid training type. Must be either 'Main Training' or 'Refreshment Training'.");
      }
      
      const templateFile = DriveApp.getFileById(templateId);
      const gradebookFolder = DriveApp.getFolderById('1PqNIjMtTauWB4NrAy-nsuws3OFrgx5rE');
      const newGradebook = templateFile.makeCopy(trainingName, gradebookFolder);
      const gradebookSS = SpreadsheetApp.open(newGradebook);
      const postPreTestSheet = gradebookSS.getSheetByName('Post Test & Pre Test');

      const traineeData = trainees.map(trainee => {
        const parts = trainee.split(' - ');
        const name = parts[0];
        const icPassport = parts[1];
        const id = parts[2];
        return [name, icPassport, id];
      });
      
      postPreTestSheet.getRange(4, 2, traineeData.length, 3).setValues(traineeData);

      resolve(newGradebook.getUrl());
    } catch (error) {
      reject(error);
    }
  });
}


function formatDateTimeForSheet(dateTimeString) {
  const [datePart, timePartWithPeriod] = dateTimeString.split(' ');
  const [day, month, year] = datePart.split('/');
  const [timePart, period] = timePartWithPeriod.split(' ');
  let [hours, minutes] = timePart.split(':');

  hours = (period === 'PM' && hours !== '12') ? parseInt(hours) + 12 :
          (period === 'AM' && hours === '12') ? 0 : parseInt(hours);

  return new Date(year, month - 1, day, hours, parseInt(minutes));
}


function formatDateTimeForEmail(dateTimeString) {
  const [datePart, timePart] = dateTimeString.split(' ');
  const [day, month, year] = datePart.split('/');
  let [hours, minutes] = timePart.split(':');
  const period = timePart.split(' ')[1];
  
  hours = (period === 'PM' && hours !== '12') ? parseInt(hours) + 12 : 
          (period === 'AM' && hours === '12') ? '00' : hours;
  
  const date = new Date(year, month - 1, day, hours, minutes);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy hh:mm a");
}

function sendEmailToTrainer(formData, gradebookUrl) {
  return new Promise((resolve) => {
    const subject = 'New K-Laser Training Created';
    const htmlBody = `
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>New K-Laser Training Created</title>
        <style>
          body {
            font-family: 'Arial', sans-serif;
            color: #333333;
            line-height: 1.6;
            margin: 0;
            padding: 0;
            background-color: #f0f4f8;
          }
          .container {
            max-width: 500px;
            margin: 40px auto;
            background-color: #ffffff;
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.1);
          }
          .header {
            background: linear-gradient(135deg, rgba(87, 63, 215, 0.9), rgba(87, 63, 215, 0.8));
            padding: 20px 0;
            text-align: center;
            color: #ffffff;
          }
          .logo {
            max-width: 120px;
            margin-bottom: 10px;
          }
          .content {
            padding: 30px;
            font-size: 16px;
          }
          .details {
            background-color: #f8f9fa;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 25px;
            border: 1px solid #e9ecef;
          }
          .button {
            display: inline-block;
            padding: 12px 30px;
            background-color: #E48F24;
            color: #ffffff;
            text-decoration: none;
            font-weight: bold;
            border-radius: 6px;
            text-align: center;
          }
          .footer {
            text-align: center;
            font-size: 14px;
            margin-top: 30px;
            color: #888888;
            padding: 20px;
            background-color: #f0f4f8;
            border-top: 1px solid #e9ecef;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <img src="https://drive.google.com/thumbnail?id=11i4Cf-c-i1PJmh-9GfJWmsH4g0rAGFm2&sz=s4000" alt="K-Laser Logo" class="logo">
            <h2>New K-Laser Training Created</h2>
          </div>
          <div class="content">
            <p>Dear ${formData.trainer},</p>
            <p>You have successfully created a new K-Laser Training. Here are the details:</p>
            <div class="details">
              <ul>
                <li><strong>Training ID:</strong> ${formData.trainingName}</li>
                <li><strong>Start Date and Time:</strong> ${formatDateTimeForEmail(formData.startDate)}</li>
                <li><strong>End Date and Time:</strong> ${formatDateTimeForEmail(formData.endDate)}</li>
                <li><strong>Training Type:</strong> ${formData.trainingType}</li>
              </ul>
            </div>
            <p>You can now access the gradebook and generate the training materials by clicking the button below:</p>
                <p style="text-align: center;">
                  <a href="${gradebookUrl}" class="button" style="background-color: #E48F24; color: #FFFFFF; padding: 10px 20px; text-decoration: none; border-radius: 5px; display: inline-block; font-weight: bold;">Access Gradebook</a>
                </p>
            <p>If you have any questions or need further assistance, please feel free to contact the support team.</p>
          </div>
          <div class="footer">
            <p>Best regards,<br>K-Laser Training System</p>
            <p><a href="https://wa.me/+601121194948">Contact Support</a></p>
          </div>
        </div>
      </body>
      </html>
    `;
    
    const trainerEmail = Session.getActiveUser().getEmail();
    
    MailApp.sendEmail({
      to: trainerEmail,
      subject: subject,
      htmlBody: htmlBody
    });
    
    resolve();
  });
}

function handleFormSubmission(formData) {
  return queueFormSubmission(formData);
}