function CreateTrainingMaterials() {
  try {
    showLoadingModal();
    createPreTestForm();
    createPostTestQuiz();
    getCorrectAnswers();

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Post Test & Pre Test');
    if (sheet) {
      sheet.getRange('E4').setFormula(
        '=ARRAYFORMULA(IFERROR(VLOOKUP(C4:C, {\'Pre-Test\'!C:C, \'Pre-Test\'!B:B}, 2, FALSE), IFERROR(VLOOKUP(D4:D, {\'Pre-Test\'!C:C, \'Pre-Test\'!B:B}, 2, FALSE), "")))'
      );

      sheet.getRange('F4').setFormula(
        '=ARRAYFORMULA(IFERROR(VLOOKUP(C4:C, {\'Post-Test\'!C:C, \'Post-Test\'!B:B}, 2, FALSE), IFERROR(VLOOKUP(D4:D, {\'Post-Test\'!C:C, \'Post-Test\'!B:B}, 2, FALSE), "")))'
      );
    } else {
      throw new Error('Sheet "Post Test & Pre Test" not found.');
    }

    closeLoadingModal();
    Logger.log("All steps completed successfully.");
  } catch (error) {
    closeLoadingModal();
    Logger.log("Error occurred during training material creation: " + error.message);
  }
}

function showCompletionModal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let preFormLink = "Pre-Test Uncreated";
  let postFormLink = "Post-Test Uncreated";

  try {
    const preTestSheet = ss.getSheetByName('Pre-Test');
    const preFormUrl = preTestSheet.getFormUrl();
    const preForm = FormApp.openByUrl(preFormUrl);
    preFormLink = preForm.getPublishedUrl();
  } catch (e) {
    Logger.log("Error fetching Pre-Test form: " + e.message);
  }

  try {
    const postTestSheet = ss.getSheetByName('Post-Test');
    const postFormUrl = postTestSheet.getFormUrl();
    const postForm = FormApp.openByUrl(postFormUrl);
    postFormLink = postForm.getPublishedUrl();
  } catch (e) {
    Logger.log("Error fetching Post-Test form: " + e.message);
  }

  const htmlContent = `<!DOCTYPE html>
    <html>
    <head>
      <link href="https: 
      <style>
                :root {
                --primary: #6366F1;
                --primary-dark: #4F46E5;
                --success: #10B981;
                --gray-light: #F3F4F6;
                --gray-medium: #E5E7EB;
                --gray-dark: #6B7280;
                --text-primary: #111827;
                --text-secondary: #4B5563;
              }
              
              body {
                font-family: 'Inter', Arial, sans-serif;
                background-color: rgba(0,0,0,0.5);
                margin: 0;
                padding: 0;
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
              }
              
              .modal {
                background-color: white;
                border-radius: 12px;
                box-shadow: 0 10px 25px rgba(0, 0, 0, 0.1);
                padding: 32px;
                text-align: center;
                width: 480px;
                max-width: 90vw;
                animation: fadeIn 0.3s ease-out;
              }
              
              @keyframes fadeIn {
                from { opacity: 0; transform: translateY(-20px); }
                to { opacity: 1; transform: translateY(0); }
              }
              
              .modal-header {
                margin-bottom: 24px;
              }
              
              .modal-icon {
                width: 64px;
                height: 64px;
                margin: 0 auto 16px;
                display: flex;
                align-items: center;
                justify-content: center;
                background-color: #ECFDF5;
                border-radius: 50%;
              }
              
              .modal-icon svg {
                width: 32px;
                height: 32px;
                color: var(--success);
              }
              
              h2 {
                color: var(--text-primary);
                font-size: 22px;
                font-weight: 600;
                margin: 0 0 8px;
              }
              
              .modal-subtitle {
                color: var(--text-secondary);
                font-size: 14px;
                margin: 0;
              }
              
              .link-container {
                margin: 20px 0;
                text-align: left;
                padding: 16px;
                border: 1px solid var(--gray-medium);
                border-radius: 8px;
                transition: all 0.2s ease;
              }
              
              .link-container:hover {
                border-color: var(--primary);
                box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.1);
              }
              
              .link-label {
                font-weight: 500;
                margin-bottom: 8px;
                color: var(--text-primary);
                font-size: 14px;
                display: flex;
                align-items: center;
              }
              
              .link-label svg {
                width: 16px;
                height: 16px;
                margin-right: 8px;
                color: var(--primary);
              }
              
              .link-url {
                word-break: break-all;
                padding: 12px;
                background-color: var(--gray-light);
                border-radius: 6px;
                margin: 12px 0;
                font-size: 13px;
                color: var(--text-secondary);
                line-height: 1.5;
              }
              
              .button-group {
                display: flex;
                gap: 8px;
              }
              
              button {
                font-family: 'Inter', Arial, sans-serif;
                font-weight: 500;
                border: none;
                padding: 8px 16px;
                border-radius: 6px;
                cursor: pointer;
                font-size: 13px;
                transition: all 0.2s ease;
                display: flex;
                align-items: center;
                justify-content: center;
              }
              
              .copy-button {
                background-color: var(--primary);
                color: white;
                flex-grow: 1;
              }
              
              .copy-button:hover {
                background-color: var(--primary-dark);
              }
              
              .copy-button svg {
                width: 14px;
                height: 14px;
                margin-right: 6px;
              }
              
              .close-button {
                background-color: white;
                color: var(--text-secondary);
                border: 1px solid var(--gray-medium);
                margin-top: 16px;
                width: 100%;
                padding: 10px;
              }
              
              .close-button:hover {
                background-color: var(--gray-light);
              }
              
              .success-message {
                color: var(--success);
                font-size: 12px;
                margin-top: 8px;
                display: none;
                align-items: center;
                justify-content: center;
                font-weight: 500;
              }
              
              .success-message svg {
                width: 14px;
                height: 14px;
                margin-right: 4px;
              }
      </style>
      <script>
        function copyToClipboard(elementId, messageId) {
          const text = document.getElementById(elementId).innerText;
          navigator.clipboard.writeText(text)
            .then(() => {
              const message = document.getElementById(messageId);
              message.style.display = 'flex';
              setTimeout(() => { message.style.display = 'none'; }, 2000);
            })
            .catch(err => console.error('Failed to copy: ', err));
        }
      </script>
    </head>
    <body>
      <div class="modal">
        <div class="modal-header">
          <div class="modal-icon">
            <svg xmlns="http: 
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7" />
            </svg>
          </div>
          <h2>Training Materials Created Successfully</h2>
          <p class="modal-subtitle">Your forms are ready to be shared with participants</p>
        </div>
        
        <div class="link-container">
          <div class="link-label">
            <svg xmlns="http: 
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" />
            </svg>
            Pre-Test Form Link
          </div>
          <div id="preTestLink" class="link-url">${preFormLink}</div>
          <div class="button-group">
            <button class="copy-button" onclick="copyToClipboard('preTestLink', 'preTestSuccess')">
              <svg xmlns="http: 
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 5H6a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2v-1M8 5a2 2 0 002 2h2a2 2 0 002-2M8 5a2 2 0 012-2h2a2 2 0 012 2m0 0h2a2 2 0 012 2v3m2 4H10m0 0l3-3m-3 3l3 3" />
              </svg>
              Copy Link
            </button>
          </div>
          <div id="preTestSuccess" class="success-message">
            <svg xmlns="http: 
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7" />
            </svg>
            Copied to clipboard!
          </div>
        </div>

        <div class="link-container">
          <div class="link-label">
            <svg xmlns="http: 
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" />
            </svg>
            Post-Test Form Link
          </div>
          <div id="postTestLink" class="link-url">${postFormLink}</div>
          <div class="button-group">
            <button class="copy-button" onclick="copyToClipboard('postTestLink', 'postTestSuccess')">
              <svg xmlns="http: 
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 5H6a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2v-1M8 5a2 2 0 002 2h2a2 2 0 002-2M8 5a2 2 0 012-2h2a2 2 0 012 2m0 0h2a2 2 0 012 2v3m2 4H10m0 0l3-3m-3 3l3 3" />
              </svg>
              Copy Link
            </button>
          </div>
          <div id="postTestSuccess" class="success-message">
            <svg xmlns="http: 
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7" />
            </svg>
            Copied to clipboard!
          </div>
        </div>

        <button class="close-button" onclick="google.script.host.close()">Close Window</button>
      </div>
    </body>
    </html>`;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(550)
    .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Form Links');
}


function showLoadingModal() {
  const htmlOutput = HtmlService.createHtmlOutput(
    "<html>" +
      "<head>" +
      "<style>" +
      "body { font-family: Arial, sans-serif; background-color: #EAEAEA; margin: 0; padding: 0; display: flex; justify-content: center; align-items: center; height: 100vh; }" +
      ".modal { background-color: #FFFFFF; border-radius: 10px; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2); padding: 30px; text-align: center; max-width: 400px; width: 100%; }" +
      "h2 { color: #11358B; font-size: 24px; margin-bottom: 20px; }" +
      ".loader { border: 5px solid #EAEAEA; border-top: 5px solid #573FD7; border-radius: 50%; width: 50px; height: 50px; animation: spin 1s linear infinite; margin: 0 auto 20px; }" +
      "@keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }" +
      "p { color: #333; font-size: 16px; margin-bottom: 0; }" +
      "</style>" +
      "</head>" +
      "<body>" +
      '<div class="modal">' +
      "<h2>Creating Training Materials</h2>" +
      '<div class="loader"></div>' +
      "<p>Please wait while materials are being generated...</p>" +
      "</div>" +
      "</body>" +
      "</html>"
  )
    .setWidth(500)
    .setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Loading...");
}

function closeLoadingModal() {
  const htmlOutput = HtmlService.createHtmlOutput("<script>google.script.host.close();</script>");
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Closing...");
}

function createPreTestForm() {
  try {
    const originalFormId = "YOUR_FORM_ID";
    const folderId = "YOUR_FOLDER_ID";
    
    const newFormFile = DriveApp.getFileById(originalFormId).makeCopy("Pre-Test");
    const folder = DriveApp.getFolderById(folderId);
    folder.addFile(newFormFile);
    DriveApp.getRootFolder().removeFile(newFormFile);

    const form = FormApp.openById(newFormFile.getId());
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    
    Utilities.sleep(1000);
    const sheet = ss.getSheets().find(s => /^Form Responses/i.test(s.getName()));
    if (!sheet) throw new Error("Response sheet not found.");

    sheet.setName("Pre-Test");
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(ss.getSheets().length);
  } catch (e) {
    throw new Error(e.message);
  }
}

function createPostTestQuiz() {
  var sourceSpreadsheet = SpreadsheetApp.openById("YOUR_MAIN_PROJECT_SHEET_ID");
  var sheet = sourceSpreadsheet.getSheetByName("Post-Test Questions");
  var data = sheet.getRange("B4:G" + sheet.getLastRow()).getValues();

  var form = FormApp.create("Post-Test Quiz");
  form.setIsQuiz(true);

var idQuestion = form.addTextItem();
  idQuestion.setTitle("Kindly write your NRIC number")
    .setHelpText("Please write your IC number in the following format: XXXXXX-XX-XXXX")  
    .setRequired(true);

  var questionsByObjective = {
    "Mechanism of Photobiomodulation": [],
    "Laser Parameters": [],
    "Laser Safety": [],
    "Product Knowledge": [],
    "Treatment Techniques": []
  };

  data.forEach(function(row) {
    if (row[0] && row[5]) {
      questionsByObjective[row[5]].push(row);
    }
  });

  var selectedQuestions = [];
  var logOutput = [];

  Object.keys(questionsByObjective).forEach(function(objective) {
    var questions = questionsByObjective[objective];
    if (questions.length < 2) {
      throw new Error("Not enough questions for objective: " + objective + ". Only " + questions.length + " available.");
    }
    for (var i = 0; i < 2; i++) {
      var index = Math.floor(Math.random() * questions.length);
      var selectedQuestion = questions.splice(index, 1)[0];
      selectedQuestions.push(selectedQuestion);
      logOutput.push("Objective: " + objective + ", Question: " + selectedQuestion[0]);
    }
  });

  selectedQuestions.forEach(function(question) {
    var quizItem = form.addMultipleChoiceItem();
    quizItem.setTitle(question[0])
      .setChoices([
        quizItem.createChoice(String(question[1]).replace('*', ''), String(question[1]).includes('*')),
        quizItem.createChoice(String(question[2]).replace('*', ''), String(question[2]).includes('*')),
        quizItem.createChoice(String(question[3]).replace('*', ''), String(question[3]).includes('*')),
        quizItem.createChoice(String(question[4]).replace('*', ''), String(question[4]).includes('*'))
      ])
      .setRequired(true)
      .setPoints(1);
  });

  form.setShuffleQuestions(false);

  var currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  form.setDestination(FormApp.DestinationType.SPREADSHEET, currentSpreadsheet.getId());

  Utilities.sleep(10000);
  var responseSheet = currentSpreadsheet.getSheets().find(sheet => sheet.getName().match(/^Form Responses/i));
  
  if (responseSheet) {
    responseSheet.setName("Post-Test");
    var totalSheets = currentSpreadsheet.getSheets().length;
    currentSpreadsheet.setActiveSheet(responseSheet);
    currentSpreadsheet.moveActiveSheet(totalSheets);
  } else {
    throw new Error("Response sheet could not be created or identified.");
  }

  var folderId = "YOUR_FOLDER_ID";
  var formFile = DriveApp.getFileById(form.getId());
  DriveApp.getFolderById(folderId).addFile(formFile);
  DriveApp.getRootFolder().removeFile(formFile);

  Logger.log("Selected Questions:");
  logOutput.forEach(function(log) {
    Logger.log(log);
  });

  Logger.log("Quiz URL: " + form.getPublishedUrl());
}


function getCorrectAnswers() {
   
  const externalSheetId = 'YOUR_MAIN_PROJECT_SHEET_ID';
  
  try {
     
    const externalSpreadsheet = SpreadsheetApp.openById(externalSheetId);
    const sourceSheet = externalSpreadsheet.getSheetByName('Post-Test Questions');
    
    if (!sourceSheet) {
      throw new Error('Source sheet "Post-Test Questions" not found');
    }
    
     
    const questionsRange = sourceSheet.getRange('B4:B' + sourceSheet.getLastRow());
    const questions = questionsRange.getValues();
    
    const optionCRange = sourceSheet.getRange('C4:C' + sourceSheet.getLastRow());
    const optionC = optionCRange.getValues();
    
    const optionDRange = sourceSheet.getRange('D4:D' + sourceSheet.getLastRow());
    const optionD = optionDRange.getValues();
    
    const optionERange = sourceSheet.getRange('E4:E' + sourceSheet.getLastRow());
    const optionE = optionERange.getValues();
    
    const optionFRange = sourceSheet.getRange('F4:F' + sourceSheet.getLastRow());
    const optionF = optionFRange.getValues();
    
    const objectivesRange = sourceSheet.getRange('G4:G' + sourceSheet.getLastRow());
    const objectives = objectivesRange.getValues();
    
     
    const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = currentSpreadsheet.getSheetByName('Correct Answers');
    
    if (!targetSheet) {
       
      targetSheet = currentSpreadsheet.insertSheet('Correct Answers');
       
      targetSheet.getRange('B3').setValue('Question');
      targetSheet.getRange('C3').setValue('Correct Answer');
      targetSheet.getRange('D3').setValue('Objective');
    }
    
     
    const dataToWrite = [];
    
    for (let i = 0; i < questions.length; i++) {
      if (!questions[i][0]) continue;  
      
      const question = questions[i][0];
      let correctAnswer = '';
      
       
      if (optionC[i][0] && optionC[i][0].toString().includes('*')) {
        correctAnswer = optionC[i][0].toString().replace('*', '').trim();
      } else if (optionD[i][0] && optionD[i][0].toString().includes('*')) {
        correctAnswer = optionD[i][0].toString().replace('*', '').trim();
      } else if (optionE[i][0] && optionE[i][0].toString().includes('*')) {
        correctAnswer = optionE[i][0].toString().replace('*', '').trim();
      } else if (optionF[i][0] && optionF[i][0].toString().includes('*')) {
        correctAnswer = optionF[i][0].toString().replace('*', '').trim();
      }
      
      const objective = objectives[i][0] || '';
      
      dataToWrite.push([question, correctAnswer, objective]);
    }
    
     
    if (dataToWrite.length > 0) {
      targetSheet.getRange(4, 2, dataToWrite.length, 3).setValues(dataToWrite);
      return `Successfully processed ${dataToWrite.length} questions.`;
    } else {
      return 'No data found to process.';
    }
    
  } catch (error) {
    Logger.log('Error: ' + error.message);
    return 'Error: ' + error.message;
  }
}