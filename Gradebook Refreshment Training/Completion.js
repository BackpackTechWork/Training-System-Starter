function masterFunction() {
  showAnimatedModal();
  updateModalContent("Compiling marks...");
  gradeObjectives();
  updateModalContent("Sending data to portal...");
  updateSheets();
  updateModalContent("The Training has been completed.", true);
}

function showAnimatedModal() {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <style>
          body{font-family:Arial,sans-serif;background-color:#EAEAEA;margin:0;padding:0;display:flex;justify-content:center;align-items:center;height:100vh}
          .modal{background-color:#FFF;border-radius:10px;box-shadow:0 4px 8px rgba(0,0,0,0.2);padding:30px;text-align:center;max-width:400px;width:100%}
          h2{color:#11358B;font-size:24px;margin-bottom:15px}
          p{color:#333;font-size:16px;margin-bottom:25px}
          .button{background-color:#573FD7;border:none;color:#FFF;padding:12px 24px;text-align:center;text-decoration:none;display:inline-block;font-size:16px;margin:4px 2px;cursor:pointer;border-radius:25px;transition:background-color 0.3s ease}
          .button:hover{background-color:#E48F24}
          .loading{display:inline-block;width:50px;height:50px;border:3px solid rgba(0,0,0,.3);border-radius:50%;border-top-color:#573FD7;animation:spin 1s ease-in-out infinite}
          @keyframes spin{to{transform:rotate(360deg)}}
        </style>
      </head>
      <body>
        <div class="modal">
          <h2>Processing</h2>
          <p id="status">Initializing...</p>
          <div class="loading"></div>
          <button id="closeButton" class="button" onclick="google.script.host.close()" style="display:none">Close</button>
        </div>
        <script>
          function updateStatus(c){document.getElementById('status').textContent=c.message;if(c.showCloseButton){document.querySelector('.loading').style.display='none';document.getElementById('closeButton').style.display='inline-block'}}
          function checkForUpdates(){google.script.run.withSuccessHandler(function(c){updateStatus(c);if(!c.showCloseButton)setTimeout(checkForUpdates,1000)}).getModalContent()}
          checkForUpdates();
        </script>
      </body>
    </html>
  `).setWidth(500).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Training Progress");
}

function updateModalContent(message, showCloseButton = false) {
  PropertiesService.getScriptProperties().setProperty('modalContent', JSON.stringify({message, showCloseButton}));
}

function getModalContent() {
  const content = PropertiesService.getScriptProperties().getProperty('modalContent');
  return content ? JSON.parse(content) : {message: "Initializing...", showCloseButton: false};
}




function gradeObjectives() {
    return new Promise((resolve, reject) => {
        try {
            const lock = LockService.getScriptLock();
            if (!lock.tryLock(30000)) {
                console.warn("Another function is currently running. Exiting gradeObjectives.");
                return reject("Lock in use.");
            }

            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const requiredSheets = ["Objective Grade (Post-Test)", "Post-Test", "Correct Answers"];
            const missingSheets = requiredSheets.filter((sheetName) => !ss.getSheetByName(sheetName));

            if (missingSheets.length > 0) {
                console.error(`Error: The following sheets are missing: ${missingSheets.join(", ")}`);
                return reject("Missing required sheets.");
            }

            const gradeSheet = ss.getSheetByName("Objective Grade (Post-Test)");
            const postTestSheet = ss.getSheetByName("Post-Test");
            const correctAnswersSheet = ss.getSheetByName("Correct Answers");

            const studentInfo = gradeSheet.getRange("C4:D" + gradeSheet.getLastRow()).getValues();
            const correctAnswers = correctAnswersSheet
                .getRange("B4:D" + correctAnswersSheet.getLastRow())
                .getValues()
                .map((row) => [row[0], typeof row[1] === "string" ? row[1].replace(/^\*/, "") : row[1], row[2]]);

            const responses = postTestSheet.getRange("C2:M" + postTestSheet.getLastRow()).getValues();

            const objectives = {
                "Mechanism of Photobiomodulation": 5,
                "Laser Parameters": 6,
                "Laser Safety": 7,
                "Product Knowledge": 8,
                "Treatment Techniques": 9,
            };

            const gradesToSet = [];

            studentInfo.forEach((student, index) => {
                const [passportIc, traineeId] = student;
                const studentResponse = responses.find((r) => r[0] === traineeId || r[0] === passportIc);

                console.log(`\nGrading student: Passport/ID - ${passportIc || traineeId}`);

                const objectiveScores = {};
                for (const objective in objectives) {
                    objectiveScores[objective] = { correct: 0, total: 0 };
                }

                if (studentResponse) {
                    for (let i = 1; i < studentResponse.length; i++) {
                        const question = postTestSheet.getRange(1, i + 3).getValue();
                        const answer = studentResponse[i];
                        const correctAnswer = correctAnswers.find((ca) => ca[0] === question);

                        if (correctAnswer) {
                            const objective = correctAnswer[2];
                            objectiveScores[objective].total++;

                            if (answer === correctAnswer[1]) {
                                objectiveScores[objective].correct++;
                            }
                        }
                    }
                }

                for (const objective in objectives) {
                    const score = objectiveScores[objective];
                    const percentage = score.total > 0 ? (score.correct / score.total) * 100 : 0;
                    const roundedPercentage = Math.round(percentage * 10) / 10;

                    gradesToSet.push({ row: index + 4, col: objectives[objective], value: roundedPercentage });
                }
            });

            gradesToSet.forEach((grade) => {
                gradeSheet.getRange(grade.row, grade.col).setValue(grade.value);
            });

            SpreadsheetApp.flush();
            lock.releaseLock();
            resolve(); // Resolving the promise after function execution
        } catch (error) {
            console.error("An error occurred:", error.message);
            reject(error.message);
        }
    });
}


function updateSheets() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSpreadsheetId = activeSpreadsheet.getId();
  var everythingTogetherSheet = activeSpreadsheet.getSheetByName("Everything Together");

  if (!everythingTogetherSheet) return;

  var lastRow = everythingTogetherSheet.getLastRow();
  var dataRange = everythingTogetherSheet.getRange("B4:F" + lastRow).getValues();
  var remarksRange = everythingTogetherSheet.getRange("T4:T" + lastRow).getValues();

  var traineeMap = {};
  var remarksMap = {};

  for (var j = 0; j < dataRange.length; j++) {
    var name = dataRange[j][0];
    var icPassport = dataRange[j][1];
    var traineeId = dataRange[j][2];
    var totalGrade = dataRange[j][4];
    var remarks = remarksRange[j][0];

    if (!traineeId) continue; // Skip if traineeId is blank

    var key = name + "|" + traineeId + "|" + icPassport;
    traineeMap[key] = totalGrade;
    remarksMap[key] = remarks;
  }

  var mainSpreadsheet = SpreadsheetApp.openById("1Rg5HvoPB69vJfZBfqUtLZiN-xFeWirqrsUPa0mcOKOM");
  var participatingTraineesSheet = mainSpreadsheet.getSheetByName("Participating Trainees");
  var allTrainingsSheet = mainSpreadsheet.getSheetByName("All Trainings");
  var traineeDatabaseSheet = mainSpreadsheet.getSheetByName("Form Responses 2");

  if (!participatingTraineesSheet || !allTrainingsSheet || !traineeDatabaseSheet) return;

  var ptLastRow = participatingTraineesSheet.getLastRow();
  var ptData = participatingTraineesSheet.getRange("B4:G" + ptLastRow).getValues();
  var ptFormulas = participatingTraineesSheet.getRange("F4:F" + ptLastRow).getFormulas();

  var gradesToUpdate = [];
  var remarksToUpdate = [];
  var updateRows = [];

  for (var i = 0; i < ptData.length; i++) {
    var hyperlinkFormula = ptFormulas[i][0];
    if (hyperlinkFormula.indexOf(activeSpreadsheetId) !== -1) {
      var name = ptData[i][0];
      var traineeId = ptData[i][1];
      var icPassport = ptData[i][2];
      var key = name + "|" + traineeId + "|" + icPassport;

      if (traineeMap[key] !== undefined) {
        gradesToUpdate.push([traineeMap[key]]);
        remarksToUpdate.push([remarksMap[key] || ""]);
        updateRows.push(i + 4); // Row number in sheet
      }
    }
  }

  for (var i = 0; i < updateRows.length; i++) {
    participatingTraineesSheet.getRange("G" + updateRows[i]).setValue(gradesToUpdate[i][0]);
    participatingTraineesSheet.getRange("I" + updateRows[i]).setValue(remarksToUpdate[i][0]);
  }

  var atLastRow = allTrainingsSheet.getLastRow();
  var atHyperlinks = allTrainingsSheet.getRange("I4:I" + atLastRow).getFormulas();

  for (var i = 0; i < atHyperlinks.length; i++) {
    var hyperlinkFormula = atHyperlinks[i][0];
    if (hyperlinkFormula.indexOf(activeSpreadsheetId) !== -1) {
      allTrainingsSheet.getRange("J" + (i + 4)).setValue("Completed");
      break;
    }
  }

  var tdLastRow = traineeDatabaseSheet.getLastRow();
  var tdData = traineeDatabaseSheet.getRange("H2:N" + tdLastRow).getValues();

  // Use Date objects directly
  var todayDate = new Date();
  var futureDate = new Date();
  futureDate.setFullYear(todayDate.getFullYear() + 2);

  var traineeIdSet = new Set();
  for (var j = 0; j < dataRange.length; j++) {
    var traineeId = dataRange[j][2];
    if (traineeId) {
      traineeIdSet.add(traineeId);
    }
  }

  for (var i = 0; i < tdData.length; i++) {
    var rowValues = tdData[i];
    if (rowValues.every(cell => cell === "")) continue; // Skip blank rows

    var traineeId = rowValues[0];
    if (traineeId && traineeIdSet.has(traineeId)) {
      var row = i + 2;
      var kCell = rowValues[7];
      var lCell = rowValues[8];

      if (!kCell && !lCell) {
        traineeDatabaseSheet.getRange("K" + row + ":M" + row).setValues([[todayDate, todayDate, futureDate]]);
      } else if (!lCell) {
        traineeDatabaseSheet.getRange("L" + row).setValue(todayDate);
        traineeDatabaseSheet.getRange("M" + row).setValue(futureDate);
      } else {
        traineeDatabaseSheet.getRange("M" + row).setValue(futureDate);
      }
    }
  }
}