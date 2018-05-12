// Documentation from Google developers site: https://developers.google.com/apps-script/articles/expense_report_approval
//  Open the Script Editor from the expense report spreadsheet (not the approvals spreadsheet). You should be able to see a script. 
// The first two lines of code are constants (APPROVALS_SPREADSHEET_ID and APPROVAL_FORM_URL) that have to be updated
// before you can run the tutorial script.
// Set the constant APPROVALS_SPREADSHEET_ID to the id of the approvals spreadsheet (not the expense report spreadsheet) you created in . 
// For instance, if the the URL of the spreadsheet is https://spreadsheets.google.com/a/yourdomain.com/ccc?key=rdkapM1ai4DGQ56B08z45g, then the id is rdkapM1ai4DGQ56B08z45g.
// Set the constant APPROVAL_FORM_URL to the URL of the approval form. Just go the the approvals 
// spreadsheet (not the Expense report spreadsheet) and click on the menu 'Form', 'Go to live form' to open the form and copy its URL.


// teeyong updated below
//https://docs.google.com/spreadsheets/d/1W0jz_FJzadWUWmYasA5PVOhNeXpr10IaxGbynSDyVO4/edit#gid=336994326

var APPROVALS_SPREADSHEET_ID = "1W0jz_FJzadWUWmYasA5PVOhNeXpr10IaxGbynSDyVO4"

// teeyong updated below
//https://docs.google.com/forms/d/e/1FAIpQLScLrHEgBx0KE9VIqR0y8odFZnnMPGswlav2eqga_2QgXgrzCw/viewform?usp=sf_link

var APPROVAL_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLScLrHEgBx0KE9VIqR0y8odFZnnMPGswlav2eqga_2QgXgrzCw/viewform?usp=pp_url"

var TREASURER_EMAIL = "teeyonglim@hotmail.com";
var STATE_MANAGER_EMAIL = "MANAGER_EMAIL";
var STATE_APPROVED = "APPROVED";
var STATE_DENIED = "DENIED";
var COLUMN_STATE = 9;   //This is the column with STATE info
var COLUMN_COMMENT = 5;  //This is the column with comments, apparently not used
var COLUMN_TIMESTAMP = 1;
var COLUMN_DEACONEMAIL = 8; //This is the column with Deacon Email address. 
var COLUMN_RECEIPT = 7; // This is the column with link to receipt

// Main tutorial function:
// For each row (expense report):
//   - if it's new, email the report to a manager(deacon) for approval
//   - if it has recently been accepted or denied by a manager (deacon), email the results to the member
//   - otherwise (expense reports that have already been fully processed or old expense reports 
//     that still have not been approved or rejected), do nothing
// Ideally, this function would be run every time the Approvals Spreadsheet or the Expense Report
// Spreadsheet are updated (via a Form submission) or regularly (once a day).

function onReportOrApprovalSubmit() {
  // This is the Expense Report Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];

  var currentTimeStamp = new Date();
  
  // Also open the Approvals Spreadsheet
  var approvalsSpreadsheet = SpreadsheetApp.openById(APPROVALS_SPREADSHEET_ID);
  var approvalsSheet = approvalsSpreadsheet.getSheets()[0];

  // Fetch all the data from the Expense Report Spreadsheet
  // getRowsData was reused from Reading Spreadsheet Data using JavaScript Objects tutorial
  var data = getRowsData(sheet);

  // Fetch all the data from the Approvals Spreadsheet
  var approvalsData = getRowsData(approvalsSheet);

  // For every expense report
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
   
    row.rowNumber = i + 2;
    if (!row.state) {
      // This is a new Expense Report.
      // Email the deacon to request his approval.
      // Lookup deacon email address based on department selected.
      row.deaconEmailAddress = lookupDeaconEmail(row);
      sendReportToManager(row);
      // Update the state of the report to avoid email sending multiple emails
      // to managers about the same report.
      sheet.getRange(row.rowNumber, COLUMN_STATE).setValue(row.state);
      //update the deaconEmailAddress cell with lookup email. Essentially overriding the value the user entered.
      sheet.getRange(row.rowNumber, COLUMN_DEACONEMAIL).setValue(row.deaconEmailAddress);
    } else if (row.state == STATE_MANAGER_EMAIL) {
      // This expense report has already been submitted to a manager for approval.
      // Check if the manager has accepted or rejected the report in the Approval Spreadsheet.
      for (var j = 0; j < approvalsData.length; ++j) {
        var approval = approvalsData[j];
        if (row.rowNumber != approval.expenseReportId) {
          continue;
        }
        // Email the member to notify him/her of the Deacon's decision about the Check or Expense Reimbursement Request.
        sendApprovalResults(row, approval);
        sendApprovalResultsTreasurer(row, approval);
        
        // Update the state of the report to APPROVED or DENIED
        sheet.getRange(row.rowNumber, COLUMN_STATE).setValue(row.state);
        break;
      }
     
     
      // Check to see how long the state has remained in STATE_MANAGER_EMAIL
      // After 3 days, the manager will be sent the approval request reminder 
      
      if (elapseTimeInDays(row)>2){
         // reset the time stamp so that manager won't be receiving constant stream of approval reminders.
         // Remember that reminders will not stop until approval is given or denied.
         // obviously, we can make it simple and just use a counter to keep state of reminders as REMINDER_ONE; REMINDER_TWO and so on.
        sheet.getRange(row.rowNumber, COLUMN_TIMESTAMP).setValue(currentTimeStamp);
      
        sendReportToManagerReminder(row);  //remind manager to approve
     
      }
    }
  }
}

// Sends an email to requester to communicate the deacon's decision on a given Check/Expense Reimbursement Request.
function sendApprovalResults(row, approval) {
  
  var approvedOrRejected = (approval.approveExpenseReport == "Yes") ? "approved" : "rejected";
  
  var mailsubject = "mailto:teeyonglim@hotmail.com?subject=Expense%20Report%20ID%20" + row.rowNumber;
  //original code had a strange way of specifying the approver. Approval email is actually used to pre-fill the approval form.
  //I changed it to use the requester email. So, iinstead of approval.emailAdress, I changed it to row.deaconEmailAddress.
  
  var message = "<HTML><BODY>"
    + "<P>" + row.department + " deacon has " + approvedOrRejected + " your check/expense reimbursement request ."
    + "<P>Amount: $" + row.amount
    + "<P>Description: " + row.descriptionOfItem
    + "<P>Report Id: " + row.rowNumber
    + "<P>Deacon's comment: " + (approval.comments || "")
  // + '<P>If your request is approved, please submit the receipt to our Treasurer <a href="'+mailsubject +'">here</a>'
    + "<P>" + "<br/>"
    + "</HTML></BODY>";
  MailApp.sendEmail(row.emailAddress, "GCCSD Check/Expense Approval Results", "", {htmlBody: message});
  if (approval.approveExpenseReport == "Yes") {
    row.state = STATE_APPROVED;
  } else {
    row.state = STATE_DENIED;
  }
}

// Sends an email to an employee to communicate the manager's decision on a given Expense Report.
function sendApprovalResultsTreasurer(row, approval) {
  
   var receiptUrl = row.pleaseAttachReceipt;
   var fileID;
   var file;
  
   if(receiptUrl!=undefined)
   {
     fileID = getIdFromURL(receiptUrl);
  
     file = DriveApp.getFileById(fileID);
   }
  
  var approvedOrRejected = (approval.approveExpenseReport == "Yes") ? "approved" : "rejected";
   
  //original code had a strange way of specifying the approver. Approval email is actually used to pre-fill the approval form.
  //I changed it to use the requester email. So, instead of approval.emailAdress, I changed it to row.deaconEmailAddress.

  var message = "<HTML><BODY>"
    + "<P>" + row.department + " deacon has " + approvedOrRejected + " a Check/Expense Reimbursement Request from " + row.nameOfRequester
    + "<P>Amount: $" + row.amount
    + "<P>Description: " + row.descriptionOfItem
    + "<P>Report Id: " + row.rowNumber
    + "<P>Deacon's comment: " + (approval.comments || "")
   // + '<P>To view the receipt uploaded by requester, click <A HREF="' +row.pleaseAttachReceipt+ '">here</A>' 
    + "<P>" + "</br>"
    + "<P>" + "</br>"
    + "</HTML></BODY>";
    var subject = "GCCSD Check/Expense Reimbursement id " + row.rowNumber + " " + approvedOrRejected;
  
  // Attach the receipt if available for Treasurer to review and to issue check
  
  if(receiptUrl!=undefined)
  {
  MailApp.sendEmail(TREASURER_EMAIL, subject, "", {htmlBody: message,attachments:[file.getAs(MimeType.JPEG)]});
  }
  else
  {
    MailApp.sendEmail(row.emailAddress, "GCCSD Check/Expense Approval Results", "", {htmlBody: message});

  }
  
  if (approval.approveExpenseReport == "Yes") {
    row.state = STATE_APPROVED;
  } else {
    row.state = STATE_DENIED;
  }
}


// https://docs.google.com/forms/d/e/1FAIpQLScLrHEgBx0KE9VIqR0y8odFZnnMPGswlav2eqga_2QgXgrzCw/viewform?usp=pp_url&entry.336166770=teeyonglim@hotmail.com&entry.1001484094=2&entry.89196159
// row.emailAddress is changed to row.nameofRequester
// Sends an email to a Deacon to request his approval of a member's check/Reimbursement Request.
// Currently, below function supports JPG format for receipt. It is attached to the email as an attachment.
// If the user did not attach a receipt, then don't try to get fileID and just send plain email.


function sendReportToManager(row) {
  
   var receiptUrl = row.pleaseAttachReceipt;
   var fileID;
   var file;
  
  
  if (receiptUrl!= undefined && check_url(receiptUrl) )
  {
  
   fileID = getIdFromURL(receiptUrl);
  
   file = DriveApp.getFileById(fileID);
  }
   
  var message = "<HTML><BODY>"
    + "<P>" + row.nameOfRequester + " has requested your approval for a check or Expense Reimbursement."
    + "<P>" + "Amount: $" + row.amount
    + "<P>" + "Description: " + row.descriptionOfItem
    + "<P>" + "Report Id: " + row.rowNumber
    + "<P>" + "<br/>" 
 //  + '<P>To view the receipt uploaded by requester, click <A HREF="' + row.pleaseAttachReceipt + '">here</A>'
    + '<P>Please approve or reject the expense report <A HREF="' + APPROVAL_FORM_URL + '&entry.336166770='+row.emailAddress + '&entry.1001484094='+ row.rowNumber +'">here</A>.'
    + "<P>" + "<br/>" 
    + "<P>" + "<br/>" 
    + "</HTML></BODY>";
 
  if(receiptUrl!=undefined && check_url(receiptUrl))
  {
   
    MailApp.sendEmail(row.deaconEmailAddress, "GCCSD Check/Expense Approval Request", "", {htmlBody: message,attachments:[file.getAs(MimeType.JPEG)]});
  }
  else
  {
    message = "<HTML><BODY>"
    + "<P>" + row.nameOfRequester + " has requested your approval for a check or Expense Reimbursement."
    + "<P>" + "Amount: $" + row.amount
    + "<P>" + "Description: " + row.descriptionOfItem
    + "<P>" + "Report Id: " + row.rowNumber
    + "<P>" + "<br/>"
    + "<P>" + "Note that Requester did not upload a receipt during the submission process"
    + "<P>" + "<br/>" 
    + '<P>Please approve or reject the expense report <A HREF="' + APPROVAL_FORM_URL + '&entry.336166770='+row.emailAddress + '&entry.1001484094='+ row.rowNumber +'">here</A>.'
    + "<P>" + "<br/>" 
    + "<P>" + "<br/>" 
    + "</HTML></BODY>";
   
    MailApp.sendEmail(row.emailAddress, "GCCSD Check/Expense Approval Request", "", {htmlBody: message});
  }
  
  row.state = STATE_MANAGER_EMAIL;
}

function sendReportToManagerReminder(row) {
  
   var receiptUrl = row.pleaseAttachReceipt;
   var fileID;
   var file;
  
    if(receiptUrl!=undefined && check_url(receiptUrl))
    {
      fileID = getIdFromURL(receiptUrl);
  
      file = DriveApp.getFileById(fileID);
    }
  
  var message = "<HTML><BODY>"
    + "<P>" + "This is a reminder that you have a pending check/Reimbursement approval request." 
    + "<P>" + row.nameOfRequester + " has requested your approval for a check or Expense Reimbursement."
    + "<P>" + "Amount: $" + row.amount
    + "<P>" + "Description: " + row.descriptionOfItem
    + "<P>" + "Report Id: " + row.rowNumber
    + "<P>" + "<br/>"
    + '<P>Please approve or reject the expense report <A HREF="' + APPROVAL_FORM_URL + '&entry.336166770='+row.emailAddress + '&entry.1001484094='+ row.rowNumber +'">here</A>.'
    + "<P>" + "<br/>" 
    + "<P>" + "<br/>" 
    + "</HTML></BODY>";
  
  if(receiptUrl!=undefined && check_url(receiptUrl))
  {
    
    MailApp.sendEmail(row.deaconEmailAddress, "GCCSD Check/Expense Approval Reminder", "", {htmlBody: message,attachments:[file.getAs(MimeType.JPEG)]});
  }
  else
  {
    message = "<HTML><BODY>"
    + "<P>" + "This is a reminder that you have a pending check/Reimbursement approval request." 
    + "<P>" + row.nameOfRequester + " has requested your approval for a check or Expense Reimbursement."
    + "<P>" + "Amount: $" + row.amount
    + "<P>" + "Description: " + row.descriptionOfItem
    + "<P>" + "Report Id: " + row.rowNumber
    + "<P>" +  "<br/>"
    + "<P>" + "Note that Requester did not upload a receipt during the submission process"
    + "<P>" + "<br/>"
    + '<P>Please approve or reject the expense report <A HREF="' + APPROVAL_FORM_URL + '&entry.336166770='+row.emailAddress + '&entry.1001484094='+ row.rowNumber +'">here</A>.'
    + "<P>" +  "<br/>"
    + "<P>" +  "<br/>"
    + "</HTML></BODY>";
    MailApp.sendEmail(row.emailAddress, "GCCSD Check/Expense Approval Reminder", "", {htmlBody: message});
  }
  
  row.state = STATE_MANAGER_EMAIL;
}

function elapseTimeInDays(row){
  
  // This is a fuaction to figure out how long a form has been submitted in days.
  var currentTime= new Date();
  
//  Logger.log(row.timestamp.getTime());
//  Logger.log(currentTime.getTime());
//  Logger.log(currentTime.getTime()- row.timestamp.getTime());
  
  var days = parseInt((currentTime.getTime()- row.timestamp.getTime())/(24*3600*1000));
  Logger.log(days);
  return days;
  
}
                    
                    
//Function to lookup Deacon Email address based on Department selected.
function lookupDeaconEmail(row){

var email ="";

switch (row.department)  {
        case "Business Administration":
            email = "kfan@pacbell.net";
            break;
        case "Chairman":
            email = "slongyu@yahoo.com";
            break;
        case "Christian Education":
            email = "ymzhangsd@gmail.com";
            break;
        case "Fellowship":
            email = "jxy17@hotmail.com";
            break;
        case "Mission":
            email = "gcjiang@yahoo.com";
            break;
        case "Secretary":
            email = "ymyu6626@yahoo.com";
            break;
        case "Treasurer":
             email = "teeyonglim@hotmail.com";
             break;
        case "Worship":
             email = "jfeng66@yahoo.com";
        default:
            email = ""; //we cannot possibly have no selection of department
            break;
    }
    
    return email;
 }
    


/////////////////////////////////////////////////////////////////////////////////
// Code reused from Reading Spreadsheet Data using JavaScript Objects tutorial //
/////////////////////////////////////////////////////////////////////////////////

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//       This argument is optional and it defaults to all the cells except those in the first row
//       or all the cells below columnHeadersRowIndex (if defined).
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range; 
// Returns an Array of objects.
function getRowsData(sheet, range, columnHeadersRowIndex) {
  var headersIndex = columnHeadersRowIndex || range ? range.getRowIndex() - 1 : 1;
  var dataRange = range || 
    sheet.getRange(headersIndex + 1, 1, sheet.getMaxRows() - headersIndex, sheet.getMaxColumns());
  var numColumns = dataRange.getEndColumn() - dataRange.getColumn() + 1;
  var headersRange = sheet.getRange(headersIndex, dataRange.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects(dataRange.getValues(), normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings. 
// Empty Strings are returned for all Strings that could not be successfully normalized.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    keys.push(normalizeHeader(headers[i]));
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit(char) {
  return char >= '0' && char <= '9';
}

// This code fetches the Google and YouTube logos, inlines them in an email
// and sends the email 
// The link to Google is broken, so I had to comment it out
// The following code works as expected
function inlineImage() {
  // var googleLogoUrl = "http://www.google.com/intl/en_com/images/srpr/logo3w.png";
 
     var youtubeLogoUrl =
         "https://developers.google.com/youtube/images/YouTube_logo_standard_white.png";

 //  var googleLogoBlob = UrlFetchApp
 //                         .fetch(googleLogoUrl)
 //                         .getBlob()
 //                         .setName("googleLogoBlob");
   var youtubeLogoBlob = UrlFetchApp
                           .fetch(youtubeLogoUrl)
                           .getBlob()
                           .setName("youtubeLogoBlob");
   MailApp.sendEmail({
     to: "teeyonglim@hotmail.com",
     subject: "Logos",
     htmlBody: "inline Youtube Logo <img src = 'cid:youtubeLogo'>",
   //  htmlBody: "inline Google Logo<img src='cid:googleLogo'> images! <br>" +
    //           "inline YouTube Logo <img src='cid:youtubeLogo'>",
     inlineImages:
       {
       //  googleLogo: googleLogoBlob,
         youtubeLogo: youtubeLogoBlob
       }
   });
 }

// This code fetches the image from a link, inlines them in an email
// and sends the email
// row contains the link in COLUMN_RECEIPT
// Following code doesn't work, still trying to figure it out. 
// Just keeping it in the file to make sure that we don't use inlineImages to attach to email. 

function inlineReceipt(row) {
 //  var receiptUrl = row.pleaseAttachReceipt;
  var receiptUrl= "https://drive.google.com/open?id=18_R-KqbVC-QHYHne7W2pGIVh-tqw3Mom";
   var receiptBlob = UrlFetchApp
                          .fetch(receiptUrl)
                          .getBlob()
                          .setName("receiptBlob");
  
   MailApp.sendEmail({
     to: "teeyonglim@hotmail.com",
     subject: "Receipt Test",
     htmlBody: "inline receipt <img src='cid:receipt'> <br>",
              
     inlineImages:
       {
         receipt: receiptBlob
        
       }
   });
 }

//The following function gets FileID from a Google File URL.
// It is taken from https://stackoverflow.com/questions/16840038/easiest-way-to-get-file-id-from-url-on-google-apps-script


function getIdFromURL(url) {
  var id = "";
  var parts = url.split(/^(([^:\/?#]+):)?(\/\/([^\/?#]*))?([^?#]*)(\?([^#]*))?(#(.*))?/);
  if (url.indexOf('?id=') >= 0){
     id = (parts[6].split("=")[1]).replace("&usp","");
     return id;
   } else {
   id = parts[5].split("/");
   //Using sort to get the id as it is the longest element. 
   var sortArr = id.sort(function(a,b){return b.length - a.length});
   id = sortArr[0];
   return id;
   }
 }

//Following code taken from Google documentation: https://developers.google.com/apps-script/reference/mail/mail-app
//

function sendEmailwithAttachments(){
// Send an email with two attachments: a file from Google Drive (as a PDF) and an HTML file.
 var file = DriveApp.getFileById('1234567890abcdefghijklmnopqrstuvwxyz');
 var blob = Utilities.newBlob('Insert any HTML content here', 'text/html', 'my_document.html');
 MailApp.sendEmail('mike@example.com', 'Attachment example', 'Two files are attached.', {
     name: 'Automatic Emailer Script',
     attachments: [file.getAs(MimeType.PDF), blob]
 });
}


//below function checks to see if a url is valid URL
function check_url(url) {
  var response = UrlFetchApp.fetch(url)
  
  if( response.getResponseCode() == 200 ) {
    return true
  } else {
    return false
  }
    
}
