function showTrainingPackageModal() {
  const ui = HtmlService.createHtmlOutputFromFile('TrainingPackage')
    .setWidth(800)
    .setHeight(700)
    .setTitle('Generate Training Package');
  
  SpreadsheetApp.getUi().showModalDialog(ui, 'Generate Training Package');
}


function getExternalSheetData() {
  try {
    const externalSheet = SpreadsheetApp.openById('YOUR_MAIN_PROJECT_SHEET_ID');
    const settingsSheet = externalSheet.getSheetByName('Settings');
    
    
    const data = {
      hospitals: settingsSheet.getRange('F5:F').getValues().filter(row => row[0] !== '').map(row => row[0]),
      models: settingsSheet.getRange('J5:J').getValues().filter(row => row[0] !== '').map(row => row[0])
    };
    
    return data;
  } catch (error) {
    console.error('Error fetching external sheet data:', error);
    return { hospitals: [], models: [] };
  }
}


function getHospitalNames() {
  return getExternalSheetData().hospitals;
}

function getKLaserModels() {
  return getExternalSheetData().models;
}

function getTrainees() {
  try {
    const currentSheet = SpreadsheetApp.getActiveSpreadsheet();
    const trainingSheet = currentSheet.getSheetByName('Post Test & Pre Test');
    
    if (!trainingSheet) {
      return [];
    }
    
    
    const data = trainingSheet.getRange('B4:C').getValues();
    
    
    return data
      .filter(row => row[0] !== '')
      .map(row => ({
        name: row[0],
        id: row[1]
      }));
    
  } catch (error) {
    console.error('Error fetching trainees:', error);
    return [];
  }
}

function extractIdFromHyperlink(hyperlinkFormula) {
  if (!hyperlinkFormula) return null;
  
  const formula = hyperlinkFormula.toString();
  const urlMatch = formula.match(/HYPERLINK\s*\(\s*"([^"]+)"/i);
  if (!urlMatch || urlMatch.length < 2) return null;
  
  const url = urlMatch[1];
  const idMatch = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return idMatch && idMatch.length > 1 ? idMatch[1] : null;
}

function findTrainingData() {
  try {
    const currentFileId = SpreadsheetApp.getActiveSpreadsheet().getId();
    const externalSheet = SpreadsheetApp.openById('YOUR_MAIN_PROJECT_SHEET_ID');
    const allTrainingsSheet = externalSheet.getSheetByName('All Trainings');
    
    if (!allTrainingsSheet) {
      throw new Error('All Trainings sheet not found in external file');
    }
    
    
    const data = allTrainingsSheet.getRange('E4:K').getValues();
    const formulas = allTrainingsSheet.getRange('I4:I').getFormulas();
    
    
    for (let i = 0; i < formulas.length; i++) {
      const hyperlinkFormula = formulas[i][0];
      const extractedId = extractIdFromHyperlink(hyperlinkFormula);
      
      if (extractedId && extractedId === currentFileId) {
        let increment = data[i][6];
        
        
        if (typeof increment === 'number') {
          increment = increment.toString().padStart(5, '0');
        } else if (typeof increment === 'string' && !isNaN(increment)) {
          increment = increment.padStart(5, '0');
        } else {
          increment = '00000'; 
        }

        return {
          startDate: data[i][0], 
          endDate: data[i][1],   
          increment: increment  
        };
      }
    }
    
    return null; 
  } catch (error) {
    Logger.log('Error in findTrainingData: ' + error.message);
    return null;
  }
}


function getDocumentReference(affiliatedCompany) {
  try {
    const trainingData = findTrainingData();
    const currentYear = new Date().getFullYear();
    return `${affiliatedCompany}/TRN/${currentYear}/${trainingData.increment || 1}`;
  } catch (error) {
    console.error('Error getting document reference:', error);
    throw error;
  }
}


function formatDateRange() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetId = ss.getId();
  
  const externalFileId = "YOUR_MAIN_PROJECT_SHEET_ID"; 
  const externalSS = SpreadsheetApp.openById(externalFileId);
  const allTrainingsSheet = externalSS.getSheetByName("All Trainings");
  
  if (!allTrainingsSheet) {
    console.log("Sheet 'All Trainings' not found");
    return "Sheet 'All Trainings' not found";
  }
  
  
  const linksRange = allTrainingsSheet.getRange("I4:I").getFormulas(); 
  const datesRange = allTrainingsSheet.getRange("E4:F").getValues();
  
  
  for (let i = 0; i < linksRange.length; i++) {
    const linkFormula = linksRange[i][0];
    const match = linkFormula.match(/HYPERLINK\("https:\/\/docs\.google\.com\/spreadsheets\/d\/([^"]+)\/edit/);
    
    if (match && match[1] === sheetId) { 
      const startDate = datesRange[i][0];
      const endDate = datesRange[i][1];
      
      if (!startDate || !endDate || startDate === "" || endDate === "") {
        return "Date not available";
      }
      
      const formattedStart = formatGoogleSheetDateTime(startDate);
      const formattedEnd = formatGoogleSheetDateTime(endDate);
      
      return formattedStart === formattedEnd ? formattedStart : `${formattedStart} - ${formattedEnd}`;
    }
  }
  
  return "Matching row not found";
}


function formatGoogleSheetDateTime(value) {
  let date;
  
  if (typeof value === "number") {
    
    date = new Date(Date.UTC(1899, 11, 30) + value * 86400000); 
  } else if (typeof value === "string") {
    
    const parts = value.match(/(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2})/);
    if (parts) {
      date = new Date(parts[3], parts[2] - 1, parts[1], parts[4], parts[5]);
    } else {
      return "Invalid date";
    }
  } else {
    date = value;
  }
  
  if (date instanceof Date && !isNaN(date.getTime())) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  }
  return "Invalid date";
}


function formatPercentage(value) {
  if (value === undefined || value === null || value === "") {
    return "0%";
  }
  
  if (typeof value === 'string' && value.includes('%')) {
    return value;
  }
  
  const numValue = Number(value);
  if (isNaN(numValue)) {
    return value;
  }
  
  return Math.round(numValue) + "%";  
}


function createTraineeDataMap(sheet, idColumn, startRow) {
  const map = new Map();
  const idRange = sheet.getRange(`${idColumn}${startRow}:${idColumn}`).getValues();

  for (let i = 0; i < idRange.length; i++) {
    const rawId = idRange[i][0];
    const id = String(rawId).trim(); 
    if (id !== '') {
      map.set(id, i + startRow);
    }
  }

  return map;
}


function getPerformanceDescriptor(percentage) {
  if (percentage > 80) return "Outstanding";
  if (percentage > 60) return "Above Average";
  if (percentage > 40) return "Average";
  if (percentage > 20) return "Below Average";
  return "Needs Improvement";
}


const OBJECTIVE_COMMENTS = {
  "Mechanism": [
    "No understanding; cannot explain or apply concepts.",
    "Limited understanding; struggles to explain or apply principles.",
    "Basic understanding; requires significant guidance to apply concepts.",
    "Understands core principles and demonstrates moderate ability to apply them practically.",
    "Fully grasps concepts and mechanisms, able to explain the process in detail and apply knowledge to various scenarios."
  ],
  "Parameters": [
    "Lacks any comprehension of parameters or their importance.",
    "Poor grasp of parameter adjustments; often requires intervention.",
    "Knows general parameters but struggles with precise adjustments.",
    "Good understanding but requires minor support in complex parameter adjustments.",
    "Excellent understanding of all parameters (wavelength, power, frequency) and their practical application."
  ],
  "Safety": [
    "Unaware of safety measures, posing potential risks.",
    "Limited awareness of safety protocols, increasing risk of errors.",
    "Understands basic safety but may overlook critical aspects.",
    "Knows most safety measures but may miss minor details.",
    "Demonstrates mastery of all safety protocols and can guide others effectively."
  ],
  "Product": [
    "Completely unfamiliar with product features.",
    "Superficial understanding of product features and benefits.",
    "Familiar with general product features but lacks depth.",
    "Solid product understanding with minor gaps in advanced features.",
    "Profound understanding of product features, functionality, and benefits."
  ],
  "Techniques": [
    "Unable to perform even basic techniques correctly.",
    "Inconsistent and lacks proficiency in applying techniques.",
    "Executes basic techniques but needs improvement in consistency and accuracy.",
    "Performs techniques well but lacks fluidity or advanced adaptability.",
    "Highly skilled in applying techniques with precision and adaptability to patient needs."
  ],
  "Total": [
    "Shows no understanding of photobiomodulation, laser parameters, safety protocols, product features, or treatment techniques, requiring comprehensive development in all areas.",
    "Demonstrates limited understanding in photobiomodulation, laser parameters, safety, product knowledge, and treatment techniques, requiring significant improvement in all areas.",
    "Has a basic understanding of photobiomodulation, laser parameters, safety, product features, and treatment techniques, but requires further development in application, precision, and consistency.",
    "Has a solid understanding of photobiomodulation, laser parameters, safety, product features, and treatment techniques, with room for improvement in advanced application and adaptability.",
    "Exhibits expert knowledge in photobiomodulation, laser parameters, safety, product features, and treatment techniques, with a strong ability to guide others effectively."
  ]
};

function getObjectiveComment(objective, percentage) {
  
  percentage = parseFloat(percentage) || 0;
  
  
  const comments = OBJECTIVE_COMMENTS[objective] || [];
  if (comments.length === 0) return "";
  
  
  let index = 0;
  if (percentage > 80) index = 4;
  else if (percentage > 60) index = 3;
  else if (percentage > 40) index = 2;
  else if (percentage > 20) index = 1;
  
  return comments[index];
}


const imageCache = {};

function getImageAsBase64(imageUrl) {
  
  if (imageCache[imageUrl]) {
    return imageCache[imageUrl];
  }
  
  try {
    const response = UrlFetchApp.fetch(imageUrl);
    const blob = response.getBlob();
    const base64String = Utilities.base64Encode(blob.getBytes());
    const contentType = blob.getContentType();
    const dataUrl = `data:${contentType};base64,${base64String}`;
    
    
    imageCache[imageUrl] = dataUrl;
    
    return dataUrl;
  } catch (error) {
    console.error('Error fetching image:', error);
    return null;
  }
}


function convertHtmlToPdf(htmlContent, fileName) {
  const blob = Utilities.newBlob(htmlContent, 'text/html', fileName + '.html');
  const pdf = blob.getAs('application/pdf');
  return DriveApp.createFile(pdf).setName(fileName + '.pdf');
}


function generateIndividualReportHtml(trainee, formData, etData, ogData) {
  
  const headerUrl = formData.affiliatedCompany === "Company 1" 
    ? "https://drive.google.com/uc?id=18WKmAt3S4XkWVz3Z4mdN-Vkd13mama4f" 
    : "https://drive.google.com/uc?id=1yzzbZNGEQ6ov5iEmNYgNphZfQsydim8y";
  
  const headerImageBase64 = getImageAsBase64(headerUrl);
  
  
  const preTestMark = parseFloat(etData[2]) || 0;
  const preTestPercentage = (preTestMark / 10) * 100;
  
  const postTestMark = parseFloat(etData[3]) || 0;
  const postTestPercentage = (postTestMark / 10) * 100;
  
  const handsOnMark = parseFloat(etData[14]) || 0;
  const handsOnPercentage = (handsOnMark / 40) * 100;
  
  const totalMark = parseFloat(etData[0]) || 0;
  
  
  const remarks = etData[19] || "";
  
  
  const mechMark = parseFloat(ogData[0]) || 0;
  const paramMark = parseFloat(ogData[1]) || 0;
  const safetyMark = parseFloat(ogData[2]) || 0;
  const productMark = parseFloat(ogData[3]) || 0;
  const techMark = parseFloat(ogData[4]) || 0;
  const objTotalMark = (mechMark + paramMark + safetyMark + productMark + techMark) / 5;
  
  
  const handsOnAspects = [
    { label: "Able to understand and explain mechanism of laser", index: 4 },
    { label: "Able to understand and describe the risk of laser hazard", index: 5 },
    { label: "Comply to applicable regulations and administrative control to minimize laser hazards", index: 6 },
    { label: "Demonstrate appropriate safety precautions before and during performing laser therapy", index: 7 },
    { label: "Exhibit professional, legal and ethical practice of laser therapy", index: 8 },
    { label: "Demonstrate the ability to prevent risk and danger", index: 9 },
    { label: "Able to choose accurate protocol of laser treatment", index: 10 },
    { label: "Able to maintain safety handling care of equipment", index: 11 },
    { label: "Able to apply laser correctly using the right treatment technique", index: 12 },
    { label: "Able to supervise and analyse patient's response and treatment outcome", index: 13 }
  ];
  
  
  const handsOnRowsArray = [];
  for (let i = 0; i < handsOnAspects.length; i++) {
    const aspect = handsOnAspects[i];
    const mark = parseFloat(etData[aspect.index]) || 0;
    const percentage = (mark / 4) * 100;
    const performanceClass = getPerformanceClass(percentage);
    
    handsOnRowsArray.push(`
      <tr>
        <td>${aspect.label}</td>
        <td class="text-center">
          <div class="percentage-badge percentage-${performanceClass}">
            ${Math.round(percentage)}%
          </div>
        </td>
        <td class="performance-${performanceClass}">${getPerformanceDescriptor(percentage)}</td>
      </tr>
    `);
  }
  
  
  const handsOnRows = handsOnRowsArray.join('');
  
  
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>${trainee.name} - Individual Report</title>
      <style>
        @page {
          margin: 0.5in;
          size: portrait;
        }
        body {
          font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
          margin: 0;
          padding: 0;
          color: #333;
          line-height: 1.5;
          font-size: 10pt;
          background-color: #fff;
        }
        .header {
          text-align: center;
          margin-bottom: 15px;
          padding: 10px;
        }
        .header img {
          width: 1000px;
          max-width: 100%;
          height: auto;
        }
        .container {
          padding: 0 15px;
        }
        .info-card {
          background-color: #f8f9fa;
          border-radius: 8px;
          padding: 15px;
          margin-bottom: 20px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.05);
          display: flex;
          flex-wrap: wrap;
          justify-content: space-between;
        }
        .info-item {
          margin-bottom: 5px;
          flex: 1 0 45%;
        }
        h2 {
          color: #2c3e50;
          border-bottom: 2px solid #3498db;
          padding-bottom: 5px;
          margin-top: 25px;
          margin-bottom: 15px;
          font-size: 14pt;
          font-weight: 600;
        }
        table {
          width: 100%;
          border-collapse: separate;
          border-spacing: 0;
          margin-bottom: 25px;
          border-radius: 8px;
          overflow: hidden;
          box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        th, td {
          padding: 10px 12px;
          text-align: left;
          font-size: 9pt;
        }
        th {
          background-color: #3498db;
          color: white;
          font-weight: 600;
          border: none;
        }
        td {
          border-top: 1px solid #eee;
        }
        tr:nth-child(odd) {
          background-color: #fff;
        }
        tr:nth-child(even) {
          background-color: #f8f9fa;
        }
        tr:hover {
          background-color: #f1f8ff;
        }
        .total-row {
          font-weight: bold;
          background-color: #e8f4fc !important;
        }
        .total-row td {
          border-top: 2px solid #3498db;
          padding-top: 12px;
          padding-bottom: 12px;
        }
        .remarks-row {
          background-color: #f1f8ff !important;
          font-style: italic;
        }
        .remarks-row td {
          border-top: 1px solid #3498db;
          padding-top: 10px;
          padding-bottom: 10px;
        }
        .performance-outstanding {
          color: #27ae60;
          font-weight: 600;
        }
        .performance-above {
          color: #2980b9;
          font-weight: 600;
        }
        .performance-average {
          color: #f39c12;
          font-weight: 600;
        }
        .performance-below {
          color: #e67e22;
          font-weight: 600;
        }
        .performance-needs {
          color: #c0392b;
          font-weight: 600;
        }
        .text-center {
          text-align: center;
        }
        .percentage-badge {
          display: inline-block;
          padding: 4px 8px;
          border-radius: 12px;
          font-weight: 600;
          font-size: 8pt;
          color: white;
          min-width: 40px;
          text-align: center;
        }
        .percentage-outstanding {
          background-color: #27ae60;
        }
        .percentage-above {
          background-color: #2980b9;
        }
        .percentage-average {
          background-color: #f39c12;
        }
        .percentage-below {
          background-color: #e67e22;
        }
        .percentage-needs {
          background-color: #c0392b;
        }
      </style>
    </head>
    <body>
      <div class="header">
        ${headerImageBase64 ? `<img src="${headerImageBase64}" alt="Company Logo">` : '<h1>Company Logo</h1>'}
      </div>
      
      <div class="container">
        <div class="info-card">
          <div class="info-item"><strong>Ref:</strong> ${formData.refNumber || getDocumentReference(formData.affiliatedCompany)}</div>
          <div class="info-item"><strong>Date:</strong> ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMMM yyyy")}</div>
          <div class="info-item"><strong>Name:</strong> ${trainee.name}</div>
          <div class="info-item"><strong>IC/Passport:</strong> ${trainee.id}</div>
          <div class="info-item"><strong>Hospital:</strong> ${formData.hospitalName}</div>
        </div>
        
        <h2>Training Result</h2>
        <table>
          <tr>
            <th width="50%">Grading Aspects</th>
            <th width="20%" class="text-center">Marks (%)</th>
            <th width="30%">Performance Descriptor</th>
          </tr>
          <tr>
            <td>Pre-Test</td>
            <td class="text-center">
              <div class="percentage-badge percentage-${getPerformanceClass(preTestPercentage)}">
                ${Math.round(preTestPercentage)}%
              </div>
            </td>
            <td class="performance-${getPerformanceClass(preTestPercentage)}">${getPerformanceDescriptor(preTestPercentage)}</td>
          </tr>
          <tr>
            <td>Post-Test</td>
            <td class="text-center">
              <div class="percentage-badge percentage-${getPerformanceClass(postTestPercentage)}">
                ${Math.round(postTestPercentage)}%
              </div>
            </td>
            <td class="performance-${getPerformanceClass(postTestPercentage)}">${getPerformanceDescriptor(postTestPercentage)}</td>
          </tr>
          <tr>
            <td>Hands on</td>
            <td class="text-center">
              <div class="percentage-badge percentage-${getPerformanceClass(handsOnPercentage)}">
                ${Math.round(handsOnPercentage)}%
              </div>
            </td>
            <td class="performance-${getPerformanceClass(handsOnPercentage)}">${getPerformanceDescriptor(handsOnPercentage)}</td>
          </tr>
          <tr class="total-row">
            <td>Total</td>
            <td class="text-center">
              <div class="percentage-badge percentage-${getPerformanceClass(totalMark)}">
                ${Math.round(totalMark)}%
              </div>
            </td>
            <td class="performance-${getPerformanceClass(totalMark)}">${getPerformanceDescriptor(totalMark)}</td>
          </tr>
          ${remarks ? `
          <tr class="remarks-row">
            <td><strong>Remarks:</strong></td>
            <td colspan="2">${remarks}</td>
          </tr>` : ''}
        </table>
        
        <h2>Understanding of Objectives</h2>
        <table>
          <tr>
            <th width="30%">Objectives</th>
            <th width="15%" class="text-center">Marks (%)</th>
            <th width="55%">Comment</th>
          </tr>
          <tr>
            <td>Mechanism of Photobiomodulation</td>
            <td class="text-center">
              <div class="percentage-badge percentage-${getPerformanceClass(mechMark)}">
                ${Math.round(mechMark)}%
              </div>
            </td>
            <td>${getObjectiveComment("Mechanism", mechMark)}</td>
          </tr>
          <tr>
            <td>Laser Parameters</td>
            <td class="text-center">
              <div class="percentage-badge percentage-${getPerformanceClass(paramMark)}">
                ${Math.round(paramMark)}%
              </div>
            </td>
            <td>${getObjectiveComment("Parameters", paramMark)}</td>
          </tr>
          <tr>
            <td>Laser Safety</td>
            <td class="text-center">
              <div class="percentage-badge percentage-${getPerformanceClass(safetyMark)}">
                ${Math.round(safetyMark)}%
              </div>
            </td>
            <td>${getObjectiveComment("Safety", safetyMark)}</td>
          </tr>
          <tr>
            <td>Product Knowledge</td>
            <td class="text-center">
              <div class="percentage-badge percentage-${getPerformanceClass(productMark)}">
                ${Math.round(productMark)}%
              </div>
            </td>
            <td>${getObjectiveComment("Product", productMark)}</td>
          </tr>
          <tr>
            <td>Treatment Techniques</td>
            <td class="text-center">
              <div class="percentage-badge percentage-${getPerformanceClass(techMark)}">
                ${Math.round(techMark)}%
              </div>
            </td>
            <td>${getObjectiveComment("Techniques", techMark)}</td>
          </tr>
          <tr class="total-row">
            <td>Total</td>
            <td class="text-center">
              <div class="percentage-badge percentage-${getPerformanceClass(objTotalMark)}">
                ${Math.round(objTotalMark)}%
              </div>
            </td>
            <td>${getObjectiveComment("Total", objTotalMark)}</td>
          </tr>
        </table>
        
        <h2>Hands on Results</h2>
        <table>
          <tr>
            <th width="60%">Hands on Aspects</th>
            <th width="15%" class="text-center">Marks (%)</th>
            <th width="25%">Performance Descriptor</th>
          </tr>
          ${handsOnRows}
          <tr class="total-row">
            <td>Total</td>
            <td class="text-center">
              <div class="percentage-badge percentage-${getPerformanceClass(handsOnPercentage)}">
                ${Math.round(handsOnPercentage)}%
              </div>
            </td>
            <td class="performance-${getPerformanceClass(handsOnPercentage)}">${getPerformanceDescriptor(handsOnPercentage)}</td>
          </tr>
        </table>
      </div>
    </body>
    </html>
  `;
  
  return html;
}


function getPerformanceClass(percentage) {
  if (percentage > 80) return "outstanding";
  if (percentage > 60) return "above";
  if (percentage > 40) return "average";
  if (percentage > 20) return "below";
  return "needs";
}


function generateGroupReportHtml(formData, selectedTrainees, everythingData, objectiveData) {
  const headerUrl = formData.affiliatedCompany === "Company 1"
    ? "https://drive.google.com/uc?id=18WKmAt3S4XkWVz3Z4mdN-Vkd13mama4f" 
    : "https://drive.google.com/uc?id=1yzzbZNGEQ6ov5iEmNYgNphZfQsydim8y";

  const headerImageBase64 = getImageAsBase64(headerUrl);

  const performanceMap = new Map();
  for (let i = 0; i < everythingData.ids.length; i++) {
    performanceMap.set(everythingData.ids[i][0], everythingData.performance[i][0]);
  }

  const remarksMap = new Map();
  if (everythingData.remarks) {
    for (let i = 0; i < everythingData.ids.length; i++) {
      const id = everythingData.ids[i][0];
      if (id) {
        remarksMap.set(id, everythingData.remarks[i][0] || "");
      }
    }
  }

  const trainingResultsRowsArray = [];
  const objectivesRowsArray = [];
  const handsOnRowsArray = [];

  for (let i = 0; i < selectedTrainees.length; i++) {
    const trainee = selectedTrainees[i];
    const performanceDesc = (trainee.performance || "Needs Improvement").trim();
    const performanceClass = getPerformanceClassFromText(performanceDesc);
    const remarks = remarksMap.get(trainee.id) || "";

    
    trainingResultsRowsArray.push(`
      <tr>
        <td>${trainee.name}</td>
        <td>${trainee.id}</td>
        <td class="performance-${performanceClass}">${performanceDesc}</td>
        <td class="remarks-cell">${remarks}</td>
      </tr>
    `);

    
    objectivesRowsArray.push(`
      <tr>
        <td>${trainee.name}</td>
        <td>${trainee.id}</td>
        <td class="text-center">${formatPercentage(trainee.photo || 0)}</td>
        <td class="text-center">${formatPercentage(trainee.laserParams || 0)}</td>
        <td class="text-center">${formatPercentage(trainee.laserSafety || 0)}</td>
        <td class="text-center">${formatPercentage(trainee.productKnowledge || 0)}</td>
        <td class="text-center">${formatPercentage(trainee.treatmentTechniques || 0)}</td>
      </tr>
    `);

    
    handsOnRowsArray.push(`
      <tr>
        <td>${trainee.name}</td>
        <td>${trainee.id}</td>
        <td class="text-center">${formatPercentage((trainee.mechanism || 0) / 4 * 100)}</td>
        <td class="text-center">${formatPercentage((trainee.laserHazard || 0) / 4 * 100)}</td>
        <td class="text-center">${formatPercentage((trainee.regulations || 0) / 4 * 100)}</td>
        <td class="text-center">${formatPercentage((trainee.safetyPrecautions || 0) / 4 * 100)}</td>
        <td class="text-center">${formatPercentage((trainee.professionalPractice || 0) / 4 * 100)}</td>
        <td class="text-center">${formatPercentage((trainee.preventRisk || 0) / 4 * 100)}</td>
        <td class="text-center">${formatPercentage((trainee.protocol || 0) / 4 * 100)}</td>
        <td class="text-center">${formatPercentage((trainee.equipmentCare || 0) / 4 * 100)}</td>
        <td class="text-center">${formatPercentage((trainee.treatmentTechnique || 0) / 4 * 100)}</td>
        <td class="text-center">${formatPercentage((trainee.patientResponse || 0) / 4 * 100)}</td>
      </tr>
    `);
  }

  const trainingResultsRows = trainingResultsRowsArray.join('');
  const objectivesRows = objectivesRowsArray.join('');
  const handsOnRows = handsOnRowsArray.join('');

  const docRef = getDocumentReference(formData.affiliatedCompany);
  const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy");

  
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>Group Report</title>
      <style>
        @page {
          size: landscape;
          margin: 0.5in;
        }
        body {
          font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
          margin: 0;
          padding: 0;
          color: #333;
          line-height: 1.5;
          font-size: 10pt;
          background-color: #fff;
        }
        .header {
          text-align: center;
          margin-bottom: 15px;
          padding: 10px;
        }
        .header img {
          max-width: 100%;
          height: auto;
        }
        .container {
          padding: 0 15px;
        }
        .info-card {
          background-color: #f8f9fa;
          border-radius: 8px;
          padding: 15px;
          margin-bottom: 20px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.05);
          display: flex;
          justify-content: space-between;
          align-items: center;
        }
        .title {
          font-size: 16pt;
          font-weight: bold;
          color: #2c3e50;
        }
        .date-ref {
          text-align: right;
          font-size: 9pt;
        }
        h2 {
          color: #2c3e50;
          text-align: center;
          font-size: 14pt;
          margin: 20px 0 15px 0;
          border-bottom: 2px solid #3498db;
          padding-bottom: 5px;
          font-weight: 600;
        }
        table {
          width: 100%;
          border-collapse: separate;
          border-spacing: 0;
          margin-bottom: 25px;
          font-size: 9pt;
          border-radius: 8px;
          overflow: hidden;
          box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        th, td {
          padding: 8px 6px;
          text-align: left;
        }
        th {
          background-color: #3498db;
          color: white;
          font-weight: 600;
          border: none;
        }
        td {
          border-top: 1px solid #eee;
        }
        tr:nth-child(odd) {
          background-color: #fff;
        }
        tr:nth-child(even) {
          background-color: #f8f9fa;
        }
        tr:hover {
          background-color: #f1f8ff;
        }
        .hands-on-table th, .hands-on-table td {
          font-size: 8pt;
          padding: 6px 4px;
        }
        .performance-outstanding {
          color: #27ae60;
          font-weight: 600;
        }
        .performance-above {
          color: #2980b9;
          font-weight: 600;
        }
        .performance-average {
          color: #f39c12;
          font-weight: 600;
        }
        .performance-below {
          color: #e67e22;
          font-weight: 600;
        }
        .performance-needs {
          color: #c0392b;
          font-weight: 600;
        }
        .text-center {
          text-align: center;
        }
        .remarks-cell {
          font-style: italic;
          color: #555;
          max-width: 200px;
        }
      </style>
    </head>
    <body>
      <div class="header">
        ${headerImageBase64 ? `<img src="${headerImageBase64}" alt="Company Logo">` : '<h1>Company Logo</h1>'}
      </div>
      
      <div class="container">
        <div class="info-card">
          <div class="title">Group Report</div>
          <div class="date-ref">
            <strong>Date:</strong> ${currentDate}<br>
            <strong>Ref:</strong> ${docRef}
          </div>
        </div>
        
        <h2>Training Results</h2>
        <table>
          <tr>
            <th width="30%">Name</th>
            <th width="20%">IC/Passport</th>
            <th width="20%">Performance Descriptor</th>
            <th width="30%">Remarks</th>
          </tr>
          ${trainingResultsRows}
        </table>
        
        <h2>Understanding of Objectives</h2>
        <table>
          <tr>
            <th width="20%">Name</th>
            <th width="15%">IC/Passport</th>
            <th width="13%" class="text-center">Mechanism of Photobiomodulation</th>
            <th width="13%" class="text-center">Laser Parameters</th>
            <th width="13%" class="text-center">Laser Safety</th>
            <th width="13%" class="text-center">Product Knowledge</th>
            <th width="13%" class="text-center">Treatment Techniques</th>
          </tr>
          ${objectivesRows}
        </table>
        
        <h2>Hands-on Results</h2>
        <table class="hands-on-table">
          <tr>
            <th>Name</th>
            <th>IC/Passport</th>
            <th class="text-center">Mechanism</th>
            <th class="text-center">Hazards</th>
            <th class="text-center">Regulations</th>
            <th class="text-center">Safety</th>
            <th class="text-center">Ethics</th>
            <th class="text-center">Risk Prevention</th>
            <th class="text-center">Protocol</th>
            <th class="text-center">Equipment</th>
            <th class="text-center">Technique</th>
            <th class="text-center">Patient Response</th>
          </tr>
          ${handsOnRows}
        </table>
      </div>
    </body>
    </html>
  `;


  return html;
}



function getPerformanceClassFromText(descriptor) {
  if (!descriptor) return "needs";
  
  if (descriptor.includes("Outstanding")) return "outstanding";
  if (descriptor.includes("Above Average")) return "above";
  if (descriptor.includes("Average") && !descriptor.includes("Above") && !descriptor.includes("Below")) return "average";
  if (descriptor.includes("Below Average")) return "below";
  return "needs";
}


function generateTrainingLetterHtml(formData, selectedTrainees) {
  
  const headerUrl = formData.affiliatedCompany === "Company 1" 
    ? "https://drive.google.com/uc?id=18WKmAt3S4XkWVz3Z4mdN-Vkd13mama4f" 
    : "https://drive.google.com/uc?id=1yzzbZNGEQ6ov5iEmNYgNphZfQsydim8y";
  
  const headerImageBase64 = getImageAsBase64(headerUrl);
  
  const docRef = getDocumentReference(formData.affiliatedCompany);
  const formattedDates = formatDateRange();
  const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy");
  const companyName = formData.affiliatedCompany === "Company 1" ? "Company 1" : "Company 2";
  
  
  const html = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>In House Training Letter</title>
  <style>
    @page {
      margin: 0.3in 0.5in 0.3in 0.5in;
      size: A4 portrait;
    }
    body {
      font-family: 'Calibri', 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
      margin: 0;
      padding: 0;
      color: #1a1a1a;
      line-height: 1.5;
      font-size: 10pt;
      background-color: #fff;
    }
    .header {
      text-align: center;
      margin: 0 0 15px 0;
      padding: 0;
    }
    .header img {
      max-width: 600px;
      height: auto;
    }
    .container {
      padding: 0 15px;
      max-width: 210mm;
      margin: 0 auto;
    }
    .letterhead {
      display: flex;
      justify-content: space-between;
      margin-bottom: 15px;
      gap: 20px;
    }
    .address-block {
      padding: 15px;
      max-width: 60%;
      font-size: 10pt;
    }
    .date-ref-block {
      text-align: right;
      padding: 15px;
      font-size: 10pt;
    }
    .recipient-block {
      margin-bottom: 12px;
      padding: 0;
    }
    .recipient-block p {
      margin: 3px 0;
    }
    .subject {
      font-weight: bold;
      font-size: 11pt;
      margin: 15px 0;
      color: #1a1a1a;
      text-align: center;
      background-color: #e8f4fc;
      padding: 10px;
      border-radius: 8px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.08);
      letter-spacing: 0.3px;
    }
    p {
      text-align: justify;
      margin: 8px 0;
      padding: 0 5px;
    }
    .details-block {
      background-color: #f9f9f9;
      border-radius: 8px;
      padding: 15px;
      margin: 15px 0;
      box-shadow: 0 2px 4px rgba(0,0,0,0.08);
    }
    .details-item {
      margin: 6px 0;
      font-size: 9.5pt;
    }
    .signature-block {
      margin: 25px 0 10px 0;
      padding: 0 5px;
    }
    .signature-line {
      border-top: 1.5px solid #000;
      width: 200px;
      margin-top: 25px;
    }
    .signatory {
      margin-top: 5px;
    }
    .signatory p {
      margin: 2px 0;
      line-height: 1.3;
    }
    .bold {
      font-weight: bold;
    }
    .italic {
      font-style: italic;
    }
    .contact-block {
      background-color: #f9f9f9;
      border-radius: 8px;
      padding: 12px 15px;
      margin: 8px 0;
      box-shadow: 0 2px 4px rgba(0,0,0,0.08);
      font-size: 9pt;
    }
    .contact-item {
      margin: 4px 0;
    }
    .contact-item a {
      color: #2980b9;
      text-decoration: none;
    }
    .contact-item a:hover {
      text-decoration: underline;
    }
    .motto {
      font-style: italic;
      text-align: center;
      color: #444;
      margin: 8px 0;
      font-size: 9pt;
      letter-spacing: 0.3px;
    }
    .footer {
      font-size: 8pt;
      color: #444;
      border-top: 1px solid #ddd;
      padding-top: 8px;
      text-align: center;
    }
    @media print {
      body {
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
      }
      .address-block, .date-ref-block, .details-block, .contact-block {
        break-inside: avoid;
      }
      .signature-block {
        page-break-inside: avoid;
      }
    }
  </style>
</head>
<body>
  <div class="header">
    ${headerImageBase64 ? `<img src="${headerImageBase64}" alt="Company Logo">` : '<h1>Company Logo</h1>'}
  </div>
  
  <div class="container">
    <div class="letterhead">
      <div class="address-block">
        <strong>${formData.hospitalName}</strong><br>
        ${formData.address.split(',').join('<br>')}
      </div>
      <div class="date-ref-block">
        <strong>Date:</strong> ${currentDate}<br>
        <strong>Ref:</strong> ${docRef}
      </div>
    </div>
    
    <div class="recipient-block">
      <p><strong>Attn to:</strong> Mr/Mrs ${formData.recipientName}</p>
      <p><strong>Phone num:</strong> ${formData.recipientPhone}</p>
    </div>
    
    <div class="subject">CONFIRMATION OF IN-HOUSE TRAINING FOR KLASER DEVICE</div>
    
    <p>Dear ${formData.recipientName},</p>
    
    <p>We are pleased to confirm that the following staff members have successfully attended the in-house training for the KLaser device. Details of the training is as below:</p>
    
    <div class="details-block">
      <div class="details-item"><strong>Date and Time:</strong> ${formattedDates}</div>
      <div class="details-item"><strong>KLaser Model:</strong> ${formData.kLaserModel}</div>
      <div class="details-item"><strong>Training Type:</strong> Refreshment Training</div>
      <div class="details-item"><strong>Group Report:</strong> please refer to Attachment 1</div>
    </div>
    
    <p>The training was conducted by ${companyName}, covering safety, technical, and clinical aspects of the device. The participants have demonstrated a satisfactory understanding of these key areas.</p>
    
    <p>As a result of this training, these staff members are now qualified to perform laser treatments using the KLaser (${formData.kLaserModel}) device.</p>
    
    <p>Should you require any further information or clarification, please do not hesitate to contact us. We appreciate your cooperation and look forward to continued collaboration.</p>
    
    <p>Thank you.</p>
    
    <p>Yours Sincerely,</p>
    
    <div class="signature-block">
      <div class="signature-line"></div>
      <div class="signatory">
        <p class="bold">Example Person</p>
        <p class="bold">Example Position,</p>
        <p class="bold">${companyName}</p>
      </div>
    </div>
    
    <div class="contact-block">
      <div class="contact-item">
        <strong>Address:</strong> example company address
      </div>
      <div class="contact-item">
        <strong>(E)</strong> <a href="mailto:example@example.com">example@example.com</a> / <a href="mailto:example@example.com">example@example.com</a>
      </div>
      <div class="contact-item">
        <strong>(W)</strong> <a href="example.com">example.com</a>, <a href="example.com">example.com</a>, <a href="example.com">www.example.com</a>
      </div>
      <div class="contact-item"><strong>(C)</strong> example number</div>
    </div>
    
    <div class="motto">example motto</div>
    
    <div class="footer">
      <strong>Malaysian address:</strong> example company address
    </div>
  </div>
</body>
</html>
  `;
  
  return html;
}


function generateIndividualReports(formData, selectedTrainees) {
  
  const individualReportsFolder = DriveApp.createFolder("Individual Reports");
  const individualReportFiles = [];
  
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const everythingTogetherSheet = ss.getSheetByName("Everything Together");
  const objectiveGradeSheet = ss.getSheetByName("Objective Grade (Post-Test)");
  
  if (!everythingTogetherSheet || !objectiveGradeSheet) {
    throw new Error("Required sheets not found");
  }
  
  
  const etTraineeMap = createTraineeDataMap(everythingTogetherSheet, 'C', 4);
  const ogTraineeMap = createTraineeDataMap(objectiveGradeSheet, 'D', 4);
  
  
  const etRangeData = everythingTogetherSheet.getRange('E4:T').getValues(); 
  const ogRangeData = objectiveGradeSheet.getRange('E4:I').getValues();
  
  
  const batchSize = 10; 
  for (let batchIndex = 0; batchIndex < selectedTrainees.length; batchIndex += batchSize) {
    const batch = selectedTrainees.slice(batchIndex, batchIndex + batchSize);
    
    batch.forEach(trainee => {
      try {
        
        const traineeRowET = etTraineeMap.get(trainee.id);
        const traineeRowOG = ogTraineeMap.get(trainee.id);
        
        if (!traineeRowET || !traineeRowOG) {
          console.error(`Trainee data not found for ${trainee.name} (${trainee.id})`);
          return;
        }
        
        
        const etData = etRangeData[traineeRowET - 4]; 
        const ogData = ogRangeData[traineeRowOG - 4]; 
        
        
        const html = generateIndividualReportHtml(trainee, formData, etData, ogData);
        
        
        const pdfFile = convertHtmlToPdf(html, `${trainee.name}_${trainee.id}`);
        const pdfCopy = pdfFile.makeCopy(`${trainee.name}_${trainee.id}.pdf`, individualReportsFolder);
        individualReportFiles.push(pdfCopy);
        
        
        pdfFile.setTrashed(true);
      } catch (error) {
        console.error(`Error generating report for trainee ${trainee.name}: ${error.message}`);
      }
    });
    
    
    if (batchIndex + batchSize < selectedTrainees.length) {
      Utilities.sleep(100); 
    }
  }
  
  return {
    folder: individualReportsFolder,
    files: individualReportFiles
  };
}

function generateGroupReport(formData, selectedTrainees) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const everythingTogetherSheet = ss.getSheetByName('Everything Together');
    const objectiveGradeSheet = ss.getSheetByName('Objective Grade (Post-Test)');

    if (!everythingTogetherSheet || !objectiveGradeSheet) {
      throw new Error('Required sheets not found. Please ensure "Everything Together" and "Objective Grade (Post-Test)" sheets exist.');
    }

    const startRow = 4;
    const lastRowET = everythingTogetherSheet.getLastRow();
    const lastRowOG = objectiveGradeSheet.getLastRow();

    if (lastRowET < startRow || lastRowOG < startRow) {
      throw new Error("No sufficient data found in required sheets.");
    }

    
    const everythingData = {
      names: everythingTogetherSheet.getRange(`B${startRow}:B${lastRowET}`).getValues(),
      ids: everythingTogetherSheet.getRange(`C${startRow}:C${lastRowET}`).getValues().map(r => [String(r[0]).trim()]),
      performance: everythingTogetherSheet.getRange(`F${startRow}:F${lastRowET}`).getValues(),
      mechanism: everythingTogetherSheet.getRange(`I${startRow}:I${lastRowET}`).getValues(),
      laserHazard: everythingTogetherSheet.getRange(`J${startRow}:J${lastRowET}`).getValues(),
      regulations: everythingTogetherSheet.getRange(`K${startRow}:K${lastRowET}`).getValues(),
      safetyPrecautions: everythingTogetherSheet.getRange(`L${startRow}:L${lastRowET}`).getValues(),
      professionalPractice: everythingTogetherSheet.getRange(`M${startRow}:M${lastRowET}`).getValues(),
      preventRisk: everythingTogetherSheet.getRange(`N${startRow}:N${lastRowET}`).getValues(),
      protocol: everythingTogetherSheet.getRange(`O${startRow}:O${lastRowET}`).getValues(),
      equipmentCare: everythingTogetherSheet.getRange(`P${startRow}:P${lastRowET}`).getValues(),
      treatmentTechnique: everythingTogetherSheet.getRange(`Q${startRow}:Q${lastRowET}`).getValues(),
      patientResponse: everythingTogetherSheet.getRange(`R${startRow}:R${lastRowET}`).getValues(),
      remarks: everythingTogetherSheet.getRange(`T${startRow}:T${lastRowET}`).getValues()
    };

    
    const objectiveData = {
      ids: objectiveGradeSheet.getRange(`D${startRow}:D${lastRowOG}`).getValues().map(r => [String(r[0]).trim()]),
      photo: objectiveGradeSheet.getRange(`E${startRow}:E${lastRowOG}`).getValues(),
      laserParams: objectiveGradeSheet.getRange(`F${startRow}:F${lastRowOG}`).getValues(),
      laserSafety: objectiveGradeSheet.getRange(`G${startRow}:G${lastRowOG}`).getValues(),
      productKnowledge: objectiveGradeSheet.getRange(`H${startRow}:H${lastRowOG}`).getValues(),
      treatmentTechniques: objectiveGradeSheet.getRange(`I${startRow}:I${lastRowOG}`).getValues()
    };

    
    const selectedData = selectedTrainees.map(selected => {
      const traineeId = String(selected.id).trim();
      const indexET = everythingData.ids.findIndex(row => row[0] === traineeId);
      const indexOG = objectiveData.ids.findIndex(row => row[0] === traineeId);

      if (indexET === -1) throw new Error(`Trainee ID ${traineeId} not found in 'Everything Together' sheet.`);
      if (indexOG === -1) throw new Error(`Trainee ID ${traineeId} not found in 'Objective Grade (Post-Test)' sheet.`);

      return {
        id: traineeId,
        name: everythingData.names[indexET]?.[0] || "",
        performance: everythingData.performance[indexET]?.[0] || "",
        mechanism: everythingData.mechanism[indexET]?.[0] || "",
        laserHazard: everythingData.laserHazard[indexET]?.[0] || "",
        regulations: everythingData.regulations[indexET]?.[0] || "",
        safetyPrecautions: everythingData.safetyPrecautions[indexET]?.[0] || "",
        professionalPractice: everythingData.professionalPractice[indexET]?.[0] || "",
        preventRisk: everythingData.preventRisk[indexET]?.[0] || "",
        protocol: everythingData.protocol[indexET]?.[0] || "",
        equipmentCare: everythingData.equipmentCare[indexET]?.[0] || "",
        treatmentTechnique: everythingData.treatmentTechnique[indexET]?.[0] || "",
        patientResponse: everythingData.patientResponse[indexET]?.[0] || "",
        remarks: everythingData.remarks[indexET]?.[0] || "",

        photo: objectiveData.photo[indexOG]?.[0] || "",
        laserParams: objectiveData.laserParams[indexOG]?.[0] || "",
        laserSafety: objectiveData.laserSafety[indexOG]?.[0] || "",
        productKnowledge: objectiveData.productKnowledge[indexOG]?.[0] || "",
        treatmentTechniques: objectiveData.treatmentTechniques[indexOG]?.[0] || ""
      };
    });

    console.log('Selected Data before PDF generation:', JSON.stringify(selectedData, null, 2));

    
    const html = generateGroupReportHtml(formData, selectedData, everythingData, objectiveData);
    const pdfFile = convertHtmlToPdf(html, "Attachment 1");

    return pdfFile;

  } catch (error) {
    console.error("Error generating group report:", error);
    throw new Error("Failed to generate group report: " + error.message);
  }
}



function generateTrainingLetter(formData) {
  try {
    const currentSheet = SpreadsheetApp.getActiveSpreadsheet();
    const fileName = currentSheet.getName() + "_Training_Package";

    let selectedTrainees = [];
    if (typeof formData.selectedTrainees === 'string') {
      try {
        selectedTrainees = JSON.parse(formData.selectedTrainees);
      } catch (e) {
        console.error("Error parsing selectedTrainees JSON:", e);
        throw new Error("Invalid trainee data format. Please try again.");
      }
    } else if (Array.isArray(formData.selectedTrainees)) {
      selectedTrainees = formData.selectedTrainees;
    } else {
      throw new Error("Invalid trainee data format. Please select trainees and try again.");
    }

    
    const html = generateTrainingLetterHtml(formData, selectedTrainees);

    
    const letterFile = convertHtmlToPdf(html, "In House Training Letter_" + formData.hospitalName);

    
    const reportFile = generateGroupReport(formData, selectedTrainees);

    
    const individualReports = generateIndividualReports(formData, selectedTrainees);

    
    const tempFolder = DriveApp.createFolder("Temp_" + fileName);

    
    DriveApp.getFileById(letterFile.getId()).makeCopy("In House Training Letter_" + formData.hospitalName + ".pdf", tempFolder);
    DriveApp.getFileById(reportFile.getId()).makeCopy("Attachment 1.pdf", tempFolder);

    
    const individualReportsFolderInTemp = tempFolder.createFolder("Individual Reports");

    
    const batchSize = 20;
    for (let i = 0; i < individualReports.files.length; i += batchSize) {
      const batch = individualReports.files.slice(i, i + batchSize);
      batch.forEach(file => {
        file.makeCopy(file.getName(), individualReportsFolderInTemp);
      });

      
      if (i + batchSize < individualReports.files.length) {
        Utilities.sleep(100);
      }
    }

    
    const allFiles = getAllFilesInFolder(tempFolder);

    
    const zipBlob = Utilities.zip(allFiles, fileName + ".zip");
    const zipFile = DriveApp.createFile(zipBlob);

    
    const targetFolder = DriveApp.getFolderById('YOUR_DRIVE_FOLDER_ID');

    
    const movedZipFile = zipFile.moveTo(targetFolder);

    
    letterFile.setTrashed(true);
    reportFile.setTrashed(true);
    tempFolder.setTrashed(true);

    
    individualReports.files.forEach(file => {
      file.setTrashed(true);
    });
    individualReports.folder.setTrashed(true);

    return movedZipFile.getUrl();
  } catch (error) {
    console.error("Error generating training letter:", error);
    throw new Error("Failed to generate training package: " + error.message);
  }
}



function getAllFilesInFolder(folder) {
  const files = [];
  
  
  const fileIterator = folder.getFiles();
  while (fileIterator.hasNext()) {
    const file = fileIterator.next();
    files.push(file.getBlob().setName(file.getName()));
  }
  
  
  const folderIterator = folder.getFolders();
  while (folderIterator.hasNext()) {
    const subFolder = folderIterator.next();
    const subFolderFiles = subFolder.getFiles();
    
    
    while (subFolderFiles.hasNext()) {
      const file = subFolderFiles.next();
      files.push(file.getBlob().setName(subFolder.getName() + "/" + file.getName()));
    }
  }
  
  return files;
}


function processForm(formData) {
  try {
    
    if (!formData.affiliatedCompany || !formData.hospitalName || !formData.kLaserModel || 
        !formData.address || !formData.recipientName || !formData.recipientPhone || 
        !formData.selectedTrainees) {
      throw new Error('Please fill in all required fields and select at least one trainee.');
    }
    
    
    try {
      findTrainingData();
    } catch (error) {
      throw new Error('This training session is not registered in the All Trainings sheet. Please add it first before generating a package.');
    }
    
    
    const zipUrl = generateTrainingLetter(formData);
    
    return {
      success: true,
      message: 'Training package generated successfully!',
      url: zipUrl
    };
  } catch (error) {
    console.error('Error processing form:', error);
    return {
      success: false,
      message: 'Error: ' + error.message
    };
  }
}