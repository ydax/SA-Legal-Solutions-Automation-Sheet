////////////////////////////////////////////////////////////////////////////////////
/////////////// DEPLOYING TEMPLATED EMAILS BASED ON SHEET ACTIVITIES ///////////////
////////////////////////////////////////////////////////////////////////////////////

/** Sends an email confirmation to deposition orderer
@params {multiple} strings Arguments passed from the getDepositionData functions originating in the New Depositions sidebars
*/
function sendConfirmationToOrderer(orderedBy, ordererEmail, caseStyle, depoDate, witness, depoHour, depoMinute, amPm, locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, services, courtReporter, videographer, pip) {
  // Re-format date, time, location
  var date = formatDateForEmail(depoDate);
  var depoTime = depoHour + ':' + depoMinute + ' ' + amPm;
  var depoLocation = locationFirm + ', ' + locationAddress1 + ', ' + locationAddress2 + ', ' + locationCity + ' ' + locationState + ' ' + locationZip;
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

  // Sends a scheduling confirmation to orderer
  GmailApp.sendEmail(
    ordererEmail, 
    // Confirmation of Witness name | Doe v. Doe | Date
    'Confirmation of ' + witness + ' Deposition | ' + caseStyle + ' | ' + date, 
    'Hello ' + firstName + ',\n\nThanks for sending this assignment to SA Legal Solutions. Our understanding of your requested resources & services are detailed below:\n• Case: ' + caseStyle + '\n• Witness: ' + witness + '\n• Date: ' + date + '\n• Time: ' + depoTime + '\n• Location: ' + depoLocation + '\n• Services: ' + services + '\n• Court reporter? ' + reporter + '\n• Videographer? ' + video + '\n• Picture-in-Picture? ' + includesPip + '\n\nIf any changes are necessary, please let us know. Thanks for your business!\n\nSA Legal Solutions | Litigation Support Specialists\nPhone: 210-591-1791\nAddress: 3201 Cherry Ridge, B 208-3, SATX 78230\nWebsite: www.salegalsolutions.com\nEmail: depos@salegalsolutions.com', 
    {
    name: 'SA Legal Solutions',
    bcc: 'shannonk@salegalsolutions.com'
    }
  );
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

function createPDFConfirmation (orderedBy, ordererEmail, caseStyle, depoDate, witness, depoTime, depoLocation, services, courtReporter, videographer, pip) {
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
  
  Logger.log( pdfUrl)
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

