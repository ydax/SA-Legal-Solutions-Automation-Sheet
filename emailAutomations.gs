////////////////////////////////////////////////////////////////////////////////////
/////////////// DEPLOYING TEMPLATED EMAILS BASED ON SHEET ACTIVITIES ///////////////
////////////////////////////////////////////////////////////////////////////////////

/** Sends an email confirmation to deposition orderer
@params {multiple} strings Arguments passed from the getDepositionData functions originating in the New Depositions sidebars
*/
function sendConfirmationToOrderer(orderedBy, ordererEmail, caseStyle, depoDate, witness, depoHour, depoMinute, amPm, locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, services, courtReporter, videographer, pip, confirmationCC, videoPlatform) {
  try {
    // Re-format date, time, location
    var date = formatDateForEmail(depoDate);
    var depoTime = depoHour + ':' + depoMinute + ' ' + amPm;
    var depoLocation = locationFirm + ', ' + locationAddress1 + ', ' + locationAddress2 + ', ' + locationCity + ' ' + locationState + ' ' + locationZip;
    // If depo is via video conference, re-format depoLocation
    if (videoPlatform.length > 2) {
      depoLocation = 'via ' + videoPlatform;
    };
    var firstName = firstNameOnly(orderedBy);
    
    // Convert courtReporter, videographer, and PIP into bool
    if (courtReporter !== '') {
      var reporter = 'Yes';
    } else {
      var reporter = 'No';
    };
    
    if (videographer !== '') {
      var video = 'Yes';
    } else {
      var video = 'No';
    };
    
    if (pip == true) {
      var includesPip = 'Yes';
    } else {
      var includesPip = 'No';
    };
    
    // Creates a PDF version of the scheduling confirmation.
    var pdfUrl = createPDFConfirmation(orderedBy, caseStyle, date, witness, depoTime, depoLocation, services, courtReporter, videographer, pip);
    
    // Dynamically generates HTML for the scheduling confirmation email to be sent to orderer.
    var template = GmailApp.getDraft('r-3755280940389063773').getMessage().getBody();
    
    var pdf = DriveApp.getFileById(getIdFromUrl(pdfUrl));
    var blob = pdf.getBlob().getAs('application/pdf').setName('SA Legal Solutions | Confirmation of ' + witness + ' Deposition on ' + depoDate + '.pdf');
    
    template = template.replace(/firstName/, firstNameOnly(orderedBy));
    template = template.replace(/witnessName/, witness);
    template = template.replace(/caseStyle/, caseStyle);
    template = template.replace(/witnessName/, witness);
    template = template.replace(/depoDate/, date);
    template = template.replace(/depoTime/, depoTime);
    template = template.replace(/depoLocation/, depoLocation);
    template = template.replace(/serviceDescription/, services);
    template = template.replace(/courtReporter/, courtReporter);
    template = template.replace(/videographerInfo/, videographer);
    template = template.replace(/pipInfo/, pip);
    
    // Sends the confirmation email.
    if (confirmationCC.length > 1) {
      GmailApp.sendEmail(
        ordererEmail, 
        'Confirmation of ' + witness + ' Deposition | ' + caseStyle + ' | ' + date, 
        'Hello ' + firstName + ',\n\nThanks for sending this assignment to SA Legal Solutions. Our understanding of your requested resources & services are detailed below:\n• Case: ' + caseStyle + '\n• Witness: ' + witness + '\n• Date: ' + date + '\n• Time: ' + depoTime + '\n• Location: ' + depoLocation + '\n• Services: ' + services + '\n• Court reporter? ' + reporter + '\n• Videographer? ' + video + '\n• Picture-in-Picture? ' + includesPip + '\n\nA PDF version of this scheduling confirmation is available for your convenience and records here: ' + pdfUrl + '\n\nIf any changes are necessary, please let us know. Thanks for your business!\n\nSA Legal Solutions | Litigation Support Specialists\nPhone: 210-591-1791\nAddress: 3201 Cherry Ridge, B 208-3, SATX 78230\nWebsite: www.salegalsolutions.com\nEmail: depos@salegalsolutions.com', 
        {
        attachments: [blob],
        htmlBody: template,
        name: 'SA Legal Solutions',
        bcc: 'shannonk@salegalsolutions.com',
        cc: confirmationCC
        });
    } else {
      GmailApp.sendEmail(
        ordererEmail, 
        'Confirmation of ' + witness + ' Deposition | ' + caseStyle + ' | ' + date, 
        'Hello ' + firstName + ',\n\nThanks for sending this assignment to SA Legal Solutions. Our understanding of your requested resources & services are detailed below:\n• Case: ' + caseStyle + '\n• Witness: ' + witness + '\n• Date: ' + date + '\n• Time: ' + depoTime + '\n• Location: ' + depoLocation + '\n• Services: ' + services + '\n• Court reporter? ' + reporter + '\n• Videographer? ' + video + '\n• Picture-in-Picture? ' + includesPip + '\n\nA PDF version of this scheduling confirmation is available for your convenience and records here: ' + pdfUrl + '\n\nIf any changes are necessary, please let us know. Thanks for your business!\n\nSA Legal Solutions | Litigation Support Specialists\nPhone: 210-591-1791\nAddress: 3201 Cherry Ridge, B 208-3, SATX 78230\nWebsite: www.salegalsolutions.com\nEmail: depos@salegalsolutions.com', 
      {
        attachments: [blob],
        htmlBody: template,
        name: 'SA Legal Solutions',
        bcc: 'shannonk@salegalsolutions.com'
      });
    }
  } catch (error) {
    Logger.log(error);
  }
};

/** Re-sends confirmation emails. */
function resendConfirmationEmail () {
  var ss = SpreadsheetApp.getActive();
  
  /** Ensures that user is highlighting a row (other than row 1) in Schedule a depo Sheet. */
  var currentSheet = ss.getActiveSheet().getName();
  if (currentSheet !== 'Schedule a depo') {
    SpreadsheetApp.getActiveSpreadsheet().toast('⚠️️ To use this tool, you must be on the \"Schedule a depo\" Sheet.');
    return;
  };
  var currentRow = ss.getActiveRange().getRow();
  if(currentRow == 1) {
    SpreadsheetApp.getActiveSpreadsheet().toast('⚠️️ Please highlight a row other than row 1.');
    return;
  };
  
  /** Gets current row and gathers data from it. */
  SpreadsheetApp.getActiveSpreadsheet().toast('🏎️💨 Automation initiated.');
  var deposSheet = ss.getSheetByName(currentSheet);
  
  var rowData = deposSheet.getRange(currentRow, 2, 1, deposSheet.getLastColumn()).getValues()[0];

  var witness = rowData[1];
  var orderedBy = rowData[2];
  var ordererEmail = rowData[3];
  var caseStyle = rowData[4];
  var depoTime = rowData[5];
  var firm = rowData[6];
  var attorney = rowData[7];
  var firmAddress1 = rowData[8];
  var firmAddress2 = rowData[9];
  var city = rowData[10];
  var state = rowData[11];
  var zip = rowData[12];
  var attorneyPhone = rowData[13];
  var attorneyEmail = rowData[14];
  var locationFirm = rowData[15];
  var locationAddress1 = rowData[16];
  var locationAddress2 = rowData[17];
  var locationCity = rowData[18];
  var locationState = rowData[19];
  var locationZip = rowData[20];
  var services = rowData[22];
  var courtReporter = rowData[23];
  var videographer = rowData[24];
  var pip = rowData[25];
  var copyAttorney = rowData[26];
  var copyFirm = rowData[27];
  var copyAddress1 = rowData[28];
  var copyAddress2 = rowData[29];
  var copyCity = rowData[30];
  var copyState = rowData[31];
  var copyZip = rowData[32];
  var copyPhone = rowData[33];
  var copyEmail = rowData[34];
  
  // Instantiates variables related to video conferencing logic in sendConfirmationToOrderer.
  let videoPlatform
  if (locationFirm.substring(0, 3) === 'via') {
    videoPlatform = locationFirm.substring(4, locationFirm.length);
  } else {
    videoPlatform = '';
  }
  var confirmationCC = '';
  
  // Restructures time variables to format expected by sendConfirmationToOrderer function.
  var depoHourMatch = depoTime.match(/.*:/)[0];
  var depoHour = depoHourMatch.substring(0, depoHourMatch.length - 1);
  var depoMinuteMatch = depoTime.match(/:.*/)[0];
  var depoMinute = depoMinuteMatch.substring(1, 3);
  var amPm = depoTime.substring(depoTime.length - 2, depoTime.length);
  
  // Generates date format expected by sendConfirmationToOrderer function.
  var unformattedDepoDate = dateFromHour(depoTime, currentRow).toString();
  var depoDate = unformattedDepoDate.substring(0, 10);
  
  sendConfirmationToOrderer(orderedBy, ordererEmail, caseStyle, depoDate, witness, depoHour, depoMinute, amPm, locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, services, courtReporter, videographer, pip, confirmationCC, videoPlatform);
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Confirmation email successfully re-sent.');
};

/** Sends a status update with a list of outstanding videos to be processed. 
1. Send a status update with a list of outstanding videos to be processed. 
Shannon is wanting an email on M / W 9am that goes to the videographer, Scott (on all of them), and Shannon. 
Only depos in the past. Only deps that have 'Yes' in the Videographer colum in L.
Needs to tell them the status of the data in columns M, N, and O.
Cluster by Zack, Manny, or Scott in Videographers, anyone else say "Other Depos Waiting to be Processed" w/ data in colums A - C.
*/
function sendVideoStatusEmail() {
  const ss = SpreadsheetApp.getActive();
  const currentList = ss.getSheetByName('Current List');
  
  // Gets the rows in Current List representing depos prior to today
  const currentListRows = currentList.getRange(2, 1, currentList.getLastRow(), 17).getValues();
  const now = getMMDDDate();
  let notProcessed = [];
  for (var i = 0; i < currentListRows.length; i++) {
    ///////////////////////////
    // FINDING RELEVANT ROWS //
    ///////////////////////////
    let date = currentListRows[i][0];
    let processedStatus = currentListRows[i][12];
    let hasVideo = currentListRows[i][7];
    let monthValue = parseInt(date.substring(5, 7)); // Because current list format is YYYY-MM-DD
    let nowMonthValue = parseInt(now.substring(0, 2));
    let dayValue = parseInt(date.substring(8, 10)); // Because current list format is YYYY-MM-DD
    let nowDayValue = parseInt(now.substring(3, 5));
    
    if (monthValue < nowMonthValue && processedStatus !== 'Processed' && hasVideo === 'Yes') {
      notProcessed.push(currentListRows[i])
    };

    if (monthValue === nowMonthValue && dayValue <= nowDayValue && processedStatus !== 'Processed' && hasVideo === 'Yes') {
      notProcessed.push(currentListRows[i])
    };
  };
  
  ///////////////////////
  // STRUCTURING EMAIL //
  ///////////////////////
  
  /** Sorts data array from current list by Videographer name */
  var sortedArray = notProcessed.sort(function(a, b) {
    
    var nameA = a[11];
    var nameB = b[11];
    
    if (nameA < nameB) {
      return -1;
    };
    if (nameA > nameB) {
      return 1;
    };
    
    return 0;
  });

  
  let body = 'Howdy. Here\'s a video processing status update for you.\n\n\nASSIGNED DEPOS WAITING TO BE PROCESSED:\n\n';
  sortedArray.forEach(function(depo) {
    if (depo[11] === 'Manny' || depo[11] === 'Sam' || depo[11] === 'Scott' || depo[11] === 'Zach') {
      let witness = depo[1];
      let client = depo[2];
      let teamMember= depo[11];
      let videoStatus = depo[12];
      let exhibitsStatus = depo[13];
      let paperworkStatus = depo[14];
      body = body + teamMember + ' | ' + witness + ' for ' + client + '\nVideo Status: ' + videoStatus + ' , Exhibits status: ' + exhibitsStatus + ' , Paperwork status: ' + paperworkStatus + '\n\n';
    };    
  });
  
  body = body + '\n\OTHER DEPOS WAITING TO BE PROCESSED:\n'
  
  sortedArray.forEach(function(depo) {
    if (depo[11] !== 'Manny' || depo[11] !== 'Sam' || depo[11] !== 'Scott' || depo[11] !== 'Zach') {
      let date = depo[0];
      let witness = depo[1];
      let client = depo[2];
      body = body + '• ' + witness + ' for ' + client + ' on ' + date + '\n';
    };
  });
  
  body = body + '\n\nNote: Any depos not marked "Processed" in column L and "Yes" in column H that are before today in the Current List will appear in this automated report.';
  const subject = 'Video Processing Status Update for ' + now;
  
  // Sends email
  GmailApp.sendEmail('shannonk@salegalsolutions.com', subject, body, { cc: 'swoody@salegalsolutions.com, zmata@salegalsolutions.com, mvasquez@salegalsolutions.com, shedemann@salegalsolutions.com', name: 'SALS Automations' });
};

////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////// UTILITIES /////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////

/** Converts YYYY-MM-DD date to MM/DD/YY */
function formatDateForEmail (date) {
  var month = date.substring(5, 7);
  var day = date.substring(8, 10);
  var year = date.substring(2, 4);
  var formattedDate = month + '/' + day + '/' + year;
  return formattedDate
}

/** Converts order full name into first name only.
@param {orderer} string Full name of orderer (e.g. Blake Smith).
@return {firstName} string First name of orderer (e.g. Blake).
*/
function firstNameOnly(orderer) {
  var firstName = orderer.match(/\S+/)[0];
  return firstName;
};

/** Generates a deposition confirmation PDF to include in confirmation email
@params {multiple} strings Deposition information that will be included in the confirmation email.
@return {pdfUrl} string URL (file hosted on Google Drive) where the confirmation PDF can be found.
*/
function createPDFConfirmation (orderedBy, caseStyle, depoDate, witness, depoTime, depoLocation, services, courtReporter, videographer, pip) {
  SpreadsheetApp.getActiveSpreadsheet().toast('📝 Started Creating Confirmation PDF');
  
  // setup
  var template = DocumentApp.openByUrl('https://docs.google.com/document/d/1hOhzKWj2l49toceMLSlMkxL4y9qp0hc5m6Rm43DgbVE/edit');
  var templateId = '1hOhzKWj2l49toceMLSlMkxL4y9qp0hc5m6Rm43DgbVE';
  var automatedConfirmationsFolderId = '1nC2FQXyiEt5SIJQC-S2JF37kfFtHtK3f';
  
  // Generates the Google Doc version of the confirmation PDF.
  var certFileName = 'SA Legal Solutions | Confirmation of ' + witness + ' Deposition on ' + depoDate ;
  var folder = DriveApp.getFolderById(automatedConfirmationsFolderId);
  var generatedDocCertUrl = DriveApp.getFileById(templateId).makeCopy(certFileName, folder).getUrl();
  
  // Generates the URL of the newly-generated Google Docs version of the confirmation PDF (without copying the template fresh).
  var newUrl = '';
  var files = DriveApp.getFilesByName(certFileName);
  while (files.hasNext()) {
    var file = files.next();
    newUrl = file.getUrl();
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('✔️ New Confirmation PDF Template Created 📝');
  
  // Adds deposition information to template.  
  var confirmationBody = DocumentApp.openByUrl(newUrl).getBody();
  
  confirmationBody.replaceText('firstName', firstNameOnly(orderedBy));
  confirmationBody.replaceText('witnessName', witness);
  confirmationBody.replaceText('caseStyle', caseStyle);
  confirmationBody.replaceText('witnessName', witness);
  confirmationBody.replaceText('depoDate', depoDate);
  confirmationBody.replaceText('depoTime', depoTime);
  confirmationBody.replaceText('depoLocation', depoLocation);
  confirmationBody.replaceText('serviceDescription', services);
  confirmationBody.replaceText('courtReporter', courtReporter);
  confirmationBody.replaceText('videographerInfo', videographer);
  confirmationBody.replaceText('pipInfo', pip);
  
  DocumentApp.openByUrl(newUrl).saveAndClose();
  
  // Converts the Google Doc version to PDF and updates sharing settings.
  var pdfUrl = convertToPDF(newUrl).slice(0, -13);
  var pdfId = getIdFromUrl(pdfUrl);
  moveFile(pdfId, automatedConfirmationsFolderId);
  folder.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  SpreadsheetApp.getActiveSpreadsheet().toast('✅ Confirmation PDF Creation Successful');
  
  // remove the doc version of the generated cert
  DriveApp.getFileById(getIdFromUrl(newUrl)).setTrashed(true);
  
  return pdfUrl;
};

/** Converts the Google Doc version of confirmation to PDF and returns the URL.
@param {newUrl} string Google Doc URL of newly-created confirmation.
@return {pdfUrl} URL (Google Drive link) of dynamically-generated deposition confirmation PDF.
*/
function convertToPDF(newUrl) {
  var docVersionOfPdf = DocumentApp.openByUrl(newUrl);
  var docblob = docVersionOfPdf.getAs('application/pdf');
  /* Add the PDF extension */
  docblob.setName(docVersionOfPdf.getName() + ".pdf");
  var pdfVersion = DriveApp.createFile(docblob);
  var pdfVersionURL = pdfVersion.getUrl();
  return pdfVersionURL;
}

/** Moves a file in in the depos Google Drive.
@params {sourceFileId, targetFolderId} strings Google Drive file and folder ids.
*/
function moveFile(sourceFileId, targetFolderId) {
  var file = DriveApp.getFileById(sourceFileId);
  file.getParents().next().removeFile(file);
  DriveApp.getFolderById(targetFolderId).addFile(file);
}

/** Gets a file ID from a Google Drive file URL.
@return string The ID of a Google Drive file, extracted from the URL.
*/
function getIdFromUrl(url) { 
  return url.match(/[-\w]{25,}/); 
};

/** Displays list of draft ids in depos@salegalsolutions.com mailbox. */
function seeDrafts () {
  var drafts = GmailApp.getDrafts();
  drafts.forEach(function(draft) {
    var id = draft.getId();
    Logger.log(id);
  });
};

/** Returns a date formatted MM-DD to match format of Current List */
function getMMDDDate() {
  let today = new Date().toISOString();
  return today.substring(5, 10);
};

