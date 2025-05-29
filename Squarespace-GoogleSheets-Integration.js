// App Script

function checkForexpired() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  const fileExpirationRange = sheet.getRange(2, 6, lastRow - 1, 1);
  const fileStatusRange = sheet.getRange(2, 10, lastRow - 1, 1);
  const currentDate = new Date();

  let fileExpirationRangeValues = fileExpirationRange.getValues();
  let fileStatusRangeValues = fileStatusRange.getValues();

  for (i = 0; i < fileExpirationRangeValues.length; i++) {
    let expirationDate = new Date(fileExpirationRangeValues[i][0]);
    // Change status of expired entries to expired
    if (expirationDate < currentDate) {
      fileStatusRangeValues[i][0] = "Expired";
    }
  }

  fileStatusRange.setValues(fileStatusRangeValues);
}

function filterUnrecognizedSubmissions() {
  // Get Sheets
  const formResponsesSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetById(); // Replace with your Form Responses Sheet ID
  const unrecognizedSubmissionsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetById(); // Replace with your Unrecognized Submissions Sheet ID
  const approvedSubmittersSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetById(); // Replace with your Approved Submitters Sheet ID
  const approvedDomainsSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetById(); // Replace with your Approved Domains Sheet ID

  // Get Data
  const formResponsesData = formResponsesSheet.getDataRange().getValues();
  const approvedDomains = approvedDomainsSheet
    .getDataRange()
    .getValues()
    .flat();
  var numOfRows = approvedSubmittersSheet.getMaxRows() - 1;
  var approvedEmails = approvedSubmittersSheet
    .getRange(2, 3, numOfRows)
    .getValues()
    .flat();

  // Loop through Form Response Sheet Row by Row

  for (i = formResponsesData.length - 1; i > 0; i--) {
    let submissionEmail = formResponsesData[i][5];
    let domainIndex = submissionEmail.indexOf("@") + 1;
    let submissionDomain = submissionEmail.substring(domainIndex);

    if (
      !approvedEmails.includes(submissionEmail) &&
      !approvedDomains.includes(submissionDomain)
    ) {
      unrecognizedSubmissionsSheet.appendRow([
        formResponsesData[i][0],
        formResponsesData[i][1],
        formResponsesData[i][2],
        formResponsesData[i][3],
        formResponsesData[i][4],
        formResponsesData[i][5],
        formResponsesData[i][6],
        formResponsesData[i][7],
      ]);

      formResponsesSheet.deleteRow(i + 1);
    }
  }
}

function getDomains() {
  var approvedSubmittersSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetById(); // Replace with your Approved Submitters Sheet ID
  var approvedDomains =
    SpreadsheetApp.getActiveSpreadsheet().getSheetById(); // Replace with your Approved Domains Sheet ID
  var numOfRows = approvedSubmittersSheet.getMaxRows() - 1;

  var approvedEmails = approvedSubmittersSheet
    .getRange(2, 3, numOfRows)
    .getValues()
    .flat();
  var domains = [];

  approvedEmails.forEach((email) => {
    let startingIndex = email.indexOf("@") + 1;
    domains.push(email.substring(startingIndex));
  });

  domains.forEach((domain) => {
    approvedDomains.appendRow([domain]);
  });
}


function sendTemplateEmailOfResponse() {
  // Get Form
  const form = FormApp.openById(''); // Replace with your Form ID

  // Get Most Recent Form Response
  const formResponses = form.getResponses();
  const latestResponse = formResponses[formResponses.length - 1];

  // Get Item Ressponses from Most Recent Response
  const itemResponses = latestResponse.getItemResponses();

  // Create Email Body
  let body =  `<b>${itemResponses[0].getItem().getTitle()}:</b>      ${itemResponses[0].getResponse()} <br>` +
                `<b>${itemResponses[1].getItem().getTitle()}:</b>      ${itemResponses[1].getResponse()} <br>` +
                `<b>${itemResponses[2].getItem().getTitle()}:</b>     ${itemResponses[2].getResponse()} <br>` +
                `<b>${itemResponses[3].getItem().getTitle()}:</b>     ${itemResponses[3].getResponse()} <br>` +
                `<b>${itemResponses[4].getItem().getTitle()}:</b>     ${itemResponses[4].getResponse()} <br>`;          

  if (itemResponses[5])
  {
    body += `<b>${itemResponses[5].getItem().getTitle()}:</b>     ${itemResponses[5].getResponse()} <br>`;
  }               

  if (itemResponses[6])
  {
    body += `<b>${itemResponses[6].getItem().getTitle()}:</b>     <a href="https://drive.google.com/open?id=${itemResponses[6].getResponse()}">File Link</a> <br>`;
  }

  // Send Email
  MailApp.sendEmail({
    to: '', // Replace with the recipient's email address
    subject: 'Form Response',
    htmlBody: body
  });
}

// Custom JS for Squarespace Code Block
// This code uses Handlebars.js and Sheetrock.js to fetch and display data from a Google Spreadsheet
var mySpreadsheet = ""; // Replace with your Google Spreadsheet URL

var alertTemplate = Handlebars.compile($("#alert-template").html());

$("#alerts").sheetrock({
  url: mySpreadsheet,
  rowTemplate: alertTemplate,
  query: "select A,B,C,D,E,F,H,I where J = 'Approved'",
});

Handlebars.registerHelper("notExpired", function (expirationDateString) {
  let currentDate = new Date();
  let expirationDate = new Date(expirationDateString);
  return expirationDate > currentDate;
});

Handlebars.registerHelper("formatSubmitDate", function (submitDate) {
  return submitDate.substring(0, 10);
});


