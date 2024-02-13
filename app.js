const { GoogleSpreadsheet } = require("google-spreadsheet");
const { JWT } = require("google-auth-library");

const fs = require("fs");
const moment = require("moment");

// Importing the "math-round" library for rounding numbers
const round = require("math-round");

// Google Sheets document and sheet identifiers
const spreadsheetId = "1N2LOsXmlou12SEzFwiMY35MOTp_qI_L4KIuOSMwW26s";
const sheetId = 0;

async function main() {
  // Reading credentials from the JSON file
  const credentialsRaw = require("./credentials.json");
  const credentials = JSON.parse(JSON.stringify(credentialsRaw));

  // Authenticating using a service account JWT
  const serviceAccountAuth = new JWT({
    email: credentials.client_email,
    key: credentials.private_key,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  // Initializing the Google Sheets document
  const doc = new GoogleSpreadsheet(spreadsheetId, serviceAccountAuth);

  // Loading information about the spreadsheet
  await doc.loadInfo();

  // Accessing the specific sheet within the document
  const sheet = doc.sheetsByIndex[sheetId];

  // Loading data from the sheet
  await sheet.loadCells("A4:H27");

  // Creating a log file name based on the current date and time
  const log = `log_register_${moment().format("DDMMYYYY_HHmmss")}.txt`;

  // Opening the log file in append mode
  const logStream = fs.createWriteStream(log, { flags: "a" });

  // Iterating over the students data
  for (let i = 3; i < 27; i++) {
    // Adjusting the loop start and end as needed
    const registration = sheet.getCell(i, 0).value;
    const name = sheet.getCell(i, 1).value;
    const p1 = parseFloat(sheet.getCell(i, 3).value);
    const p2 = parseFloat(sheet.getCell(i, 4).value);
    const p3 = parseFloat(sheet.getCell(i, 5).value);
    const absences = parseInt(sheet.getCell(i, 2).value);

    // Calculating the average of the three scores and multiplying by ten to return the average between 0 and 10
    const average = (p1 + p2 + p3) / (3 * 10);

    // Determining the student's situation
    let situation = "";

    // Setting default score for final approval
    var finalApprovalScore = 0;

    if (absences > 0.25 * 60) {
      // Student failed due to excessive absences
      situation = "Reprovado por Falta";

      // Seting the grade to 0 (zero) as the student is already failed
      sheet.getCell(i, 7).value = finalApprovalScore;
    } else if (average < 5) {
      // Student failed due to low average
      situation = "Reprovado por Nota";
    } else if (average >= 7) {
      // Student passed
      situation = "Aprovado";

      // Set the grade to 0 (zero) as the student already passed
      sheet.getCell(i, 7).value = finalApprovalScore;
    } else {
      // Student is in the final exam situation

      // Calculating the final passing grade and rounding to the next whole number in cases where the result has decimal places
      finalApprovalScore = Math.round(Math.max((5 * 2 - average) * 10, 0));

      // Updating cell with the score required for final approval
      sheet.getCell(i, 7).value = finalApprovalScore;

      // Set the situation as "Exame Final"
      situation = "Exame Final";
    }

    // Updating cell with the student's situation
    sheet.getCell(i, 6).value = situation;

    // Registering the student's information
    const registerLog = `${moment().format(
      "DD-MM-YYYY HH:mm:ss"
    )} - ${registration} - ${name}: Média "${average.toFixed(
      1
    )}", Faltas "${absences}", Situação "${situation}", Nota para Aprovação Final "${finalApprovalScore}"\n`;

    logStream.write(registerLog);
  }

  // Closing log flux
  logStream.end();

  // Saving the changes to the spreadsheet
  await sheet.saveUpdatedCells();
}

// Calling the main function to execute the script
main();
