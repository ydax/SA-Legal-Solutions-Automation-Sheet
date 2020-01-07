////////////////////////////////////////////////////////////////////////////////////
///////// AUTOMATED MANIPULATION OF DEPOSITION DATA WITHIN THE SPREADSHEET /////////
////////////////////////////////////////////////////////////////////////////////////

/** Collects data from new orderer deposition sidebar
* @params {multiple} strings, bool Values from the form deployed through Google Sheet.
* @return Sequential array of values.
*/
function getNewDepositionData(orderedBy,orderedByEmail, witnessName, caseStyle, depoDate, depoHour, depoMinute, amPm, firm, attorney, attorneyEmail, attorneyPhone, firmAddress1, firmAddress2, city, state, zip, locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, locationPhone, services, courtReporter, videographer, pip) {
  // Updates progress to user through the sidebar UI
  SpreadsheetApp.getActiveSpreadsheet().toast('🚀️ Automation initiated');
  
  // Checks for orderedByEmail, if blank, exits script and alerts user
  if (orderedByEmail == '') {
    SpreadsheetApp.getActiveSpreadsheet().toast('⚠️️ Error: orderer email address was not included. Please add it.');
    return;
  };
  
  // Concatenates deposition time-related variables for print formatting
  var depoTime = depoHour + ':' + depoMinute + ' ' + amPm;
  
  // Converts PIP boolean value into "yes" or "no" string
  if (pip === true) {
    pip = 'Yes';
  } else {
    pip = 'No';
  };
  
  // Begins construction of deposition information array
  var newScheduledDepo = ['Scheduled', depoDate, witnessName, orderedBy, orderedByEmail, caseStyle, depoTime, firm, attorney, firmAddress1, firmAddress2, city, state, zip, attorneyPhone, attorneyEmail, locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, locationPhone, services, courtReporter, videographer, pip];

  // Formats the array for Google Sheets setValue() method, calls printing function
  var formattedArray = [newScheduledDepo];
  printNewDeposition(formattedArray);
  SpreadsheetApp.getActiveSpreadsheet().toast('➕️ Depo added to Schedule a depo sheet');
  
  // Adds deposition information to Video Worksheet
  updateVideoWorksheet(locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, depoDate, witnessName, caseStyle, depoTime, courtReporter, firm, attorney, attorneyEmail, firmAddress1, firmAddress2, city, state, zip, orderedBy);
  SpreadsheetApp.getActiveSpreadsheet().toast('🎥 Video Worksheet updated');
  
  // Adds deposition information to CR Worksheet
  updateCRWorksheet(locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, depoDate, witnessName, caseStyle, depoTime, courtReporter, firm, attorney, attorneyEmail, firmAddress1, firmAddress2, city, state, zip, attorneyPhone, orderedBy);
  SpreadsheetApp.getActiveSpreadsheet().toast('✍️ CR Worksheet updated');

  // Adds deposition information to Confirmation of Scheduling  
  updateConfirmationOfScheduling(locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, depoDate, witnessName, caseStyle, depoTime, courtReporter, firm, attorney, firmAddress1, firmAddress2, city, state, zip, attorneyPhone, orderedBy, videographer, pip);
  SpreadsheetApp.getActiveSpreadsheet().toast('🗓 Confirmation of Scheduling updated');
  
  // Adds the deposition to the Services calendar and logs it for internal record keeping.
  var event = addEvent(orderedBy, witnessName, caseStyle, depoDate, depoHour, depoMinute, amPm, firm, attorney, firmAddress1, firmAddress2, city, state, zip, locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, services, courtReporter, videographer, pip);
  addOrderToLog(orderedBy, firm);
  SpreadsheetApp.getActiveSpreadsheet().toast('📅 Deposition added to Services calendar');
  
  // Sends a confirmation email to orderer
  sendConfirmationToOrderer(orderedBy, orderedByEmail, caseStyle, depoDate, witnessName, depoHour, depoMinute, amPm, locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, services, courtReporter, videographer, pip);
  SpreadsheetApp.getActiveSpreadsheet().toast('📧 Confirmation email sent to orderer');
};

/** Collects data from repeat orderer deposition sidebar
* @params {multiple} strings, bool Values from the form deployed through Google Sheet.
* @return Sequential array of values.
*/
function getRepeatDepositionData(previousOrderer, witnessName, caseStyle, depoDate, depoHour, depoMinute, amPm, locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, locationPhone, services, courtReporter, videographer, pip) {
  // Updates progress to user through the sidebar UI
  SpreadsheetApp.getActiveSpreadsheet().toast('🚀️ Automation initiated');
  
  // Concatenates deposition time-related variables for print formatting
  var depoTime = depoHour + ':' + depoMinute + ' ' + amPm;
  
  // Gets email address from previous orderer, exits the process if not included on the "Schedule a depo" Sheet
  var ordererEmail = emailFromOrderer(previousOrderer);
  if (ordererEmail == '') {
    SpreadsheetApp.getActiveSpreadsheet().toast('⚠️️ Error: orderer email address is not included in column E of "Schedule a depo" sheet. Please add it.');
    return;
  };
  
  // Begins construction of deposition information array
  var newScheduledDepo = ['Scheduled', depoDate, witnessName, previousOrderer, ordererEmail, caseStyle, depoTime];
  
  // Gets firm and attorney information from previous orderer, pushes it into the newScheduledDepo array
  var infoFromPreviousOrderer = firmInformationFromOrderer(previousOrderer);
  Logger.log(infoFromPreviousOrderer);
  for (var i = 0; i < infoFromPreviousOrderer.length; i++) {
    newScheduledDepo.push(infoFromPreviousOrderer[i]);
  };
  SpreadsheetApp.getActiveSpreadsheet().toast('📙️ Found attorney and firm info');
  
  // Convert PIP boolean value into "yes" or "no" string
  if (pip === true) {
    pip = 'Yes';
  } else {
    pip = 'No';
  };

  newScheduledDepo.push(locationFirm); 
  newScheduledDepo.push(locationAddress1); 
  newScheduledDepo.push(locationAddress2); 
  newScheduledDepo.push(locationCity); 
  newScheduledDepo.push(locationState); 
  newScheduledDepo.push(locationZip); 
  newScheduledDepo.push(locationPhone); 
  newScheduledDepo.push(services); 
  newScheduledDepo.push(courtReporter); 
  newScheduledDepo.push(videographer); 
  newScheduledDepo.push(pip); 
  
  // Formats the array for Google Sheets setValue() method, calls printing function
  var formattedArray = [newScheduledDepo];
  printNewDeposition(formattedArray);
  SpreadsheetApp.getActiveSpreadsheet().toast('➕️ Depo added to Schedule a depo sheet');
  
  // Adds deposition information to Video Worksheet
  updateVideoWorksheet(locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, depoDate, witnessName, caseStyle, depoTime, courtReporter, infoFromPreviousOrderer[0], infoFromPreviousOrderer[1], infoFromPreviousOrderer[8], infoFromPreviousOrderer[2], infoFromPreviousOrderer[3], infoFromPreviousOrderer[4], infoFromPreviousOrderer[5], infoFromPreviousOrderer[6], previousOrderer);
  SpreadsheetApp.getActiveSpreadsheet().toast('🎥 Video Worksheet updated');
  
  // Adds deposition information to CR Worksheet
  updateCRWorksheet(locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, depoDate, witnessName, caseStyle, depoTime, courtReporter, infoFromPreviousOrderer[0], infoFromPreviousOrderer[1], infoFromPreviousOrderer[8], infoFromPreviousOrderer[2], infoFromPreviousOrderer[3], infoFromPreviousOrderer[4], infoFromPreviousOrderer[5], infoFromPreviousOrderer[6], infoFromPreviousOrderer[7], previousOrderer);
  SpreadsheetApp.getActiveSpreadsheet().toast('✍️ CR Worksheet updated');
  
  // Adds deposition information to Confirmation of Scheduling  
  //                            locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, depoDate, witnessName, caseStyle, depoTime, courtReporter,  firm,                       attorney,                   firmAddress1,                firmAddress2,               city,                      state,                     zip, attorneyPhone, orderedBy, videographer, pip) {
  updateConfirmationOfScheduling(locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, depoDate, witnessName, caseStyle, depoTime, courtReporter, infoFromPreviousOrderer[0], infoFromPreviousOrderer[1], infoFromPreviousOrderer[2], infoFromPreviousOrderer[3], infoFromPreviousOrderer[4], infoFromPreviousOrderer[5], infoFromPreviousOrderer[6], infoFromPreviousOrderer[7], previousOrderer, videographer, pip);
  SpreadsheetApp.getActiveSpreadsheet().toast('🗓 Confirmation of Scheduling updated');
  
  // Adds the deposition to the Services calendar and logs it.
  addEvent(previousOrderer, witnessName, caseStyle, depoDate, depoHour, depoMinute, amPm, infoFromPreviousOrderer[0], infoFromPreviousOrderer[1], infoFromPreviousOrderer[2], infoFromPreviousOrderer[3], infoFromPreviousOrderer[4], infoFromPreviousOrderer[5], infoFromPreviousOrderer[6], locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, services, courtReporter, videographer, pip);
  addOrderToLog(previousOrderer, infoFromPreviousOrderer[0]);
  SpreadsheetApp.getActiveSpreadsheet().toast('📅 Deposition added to Services calendar');
  
  // Sends a confirmation email to orderer
  sendConfirmationToOrderer(previousOrderer, ordererEmail, caseStyle, depoDate, witnessName, depoHour, depoMinute, amPm, locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, services, courtReporter, videographer, pip);
  SpreadsheetApp.getActiveSpreadsheet().toast('📧 Confirmation email sent to orderer');
};

////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////// UTILITIES /////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////

/** Return a clean array of previous deposition orderers 
* @return {array} Previous deposition orderers without duplicates, sorted alphabetically (by First Name).
*/
function getPreviousOrderers () {
  var ss = SpreadsheetApp.getActive()
  var depoSheet = ss.getSheetByName('Schedule a depo');
  var lastDepoSheetRow = depoSheet.getLastRow();
  var rawOrdererData = depoSheet.getRange(2, 4, lastDepoSheetRow, 1).getValues();
  
  // Creates a 2d array of previous orderers.
  var firstLevelArray = [];
  rawOrdererData.forEach(function(element) {
    firstLevelArray.push(element[0]);
  });
  
  /** Removes all elements that are empty strings from an array
  */
  function isntEmpty (element) {
  return element != '';
  };
  
  // Filter out empty strings, remove duplicate elements, and sort the array
  var firstLevelEmptiesRemoved = firstLevelArray.filter(isntEmpty);
  
  var uniqueArray = firstLevelEmptiesRemoved.filter(function(elem, index, self) {
    return index === self.indexOf(elem);
  });
  
  var sortedUniqueArray = uniqueArray.sort();
  
  return sortedUniqueArray;
};

/** Updates the Video Worksheet with the most recently-entered deposition information.
@params {depositionInformation} strings Deposition information received from the sidebar.
*/
function updateVideoWorksheet(locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, depoDate, witnessName, caseStyle, depoTime, courtReporter, firm, attorney, attorneyEmail, firmAddress1, firmAddress2, city, state, zip, orderedBy) {
  var videoSheet = SpreadsheetApp.getActive().getSheetByName('Video Worksheet');
  
  // Sets values inside video worksheet.
  videoSheet.getRange('B9').setValue(locationFirm);
  videoSheet.getRange('B10').setValue(locationAddress1);
  videoSheet.getRange('B11').setValue(locationAddress2);
  videoSheet.getRange('B12').setValue(locationCity);
  videoSheet.getRange('C12').setValue(locationState);
  videoSheet.getRange('D12').setValue(locationZip);
  videoSheet.getRange('F9').setValue(depoDate);
  videoSheet.getRange('F10').setValue(witnessName);
  videoSheet.getRange('F11').setValue(caseStyle);
  videoSheet.getRange('F14').setValue(depoTime);
  videoSheet.getRange('B13').setValue(courtReporter);
  videoSheet.getRange('B22').setValue(firm);
  videoSheet.getRange('B20').setValue(attorney);
  videoSheet.getRange('D21').setValue(attorneyEmail);
  videoSheet.getRange('B24').setValue(firmAddress1);
  videoSheet.getRange('B25').setValue(firmAddress2);
  videoSheet.getRange('A26').setValue(city);
  videoSheet.getRange('B26').setValue(state);
  videoSheet.getRange('C26').setValue(zip);
  videoSheet.getRange('H55').setValue(orderedBy);
};

/** Updates the CR Worksheet with the most recently-entered deposition information.
@params {depositionInformation} strings Deposition information received from the sidebar.
*/
function updateCRWorksheet(locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, depoDate, witnessName, caseStyle, depoTime, courtReporter, firm, attorney, attorneyEmail, firmAddress1, firmAddress2, city, state, zip, attorneyPhone, orderedBy) {
  var crSheet = SpreadsheetApp.getActive().getSheetByName('CR Worksheet');
  
  // Sets values inside CR Worksheet.
  crSheet.getRange('B7').setValue(locationFirm);
  crSheet.getRange('B8').setValue(locationAddress1);
  crSheet.getRange('B9').setValue(locationAddress2);
  crSheet.getRange('B10').setValue(locationCity);
  crSheet.getRange('C10').setValue(locationState);
  crSheet.getRange('D10').setValue(locationZip);
  crSheet.getRange('F7').setValue(depoDate);
  crSheet.getRange('F8').setValue(witnessName);
  crSheet.getRange('F9').setValue(caseStyle);
  crSheet.getRange('F11').setValue(depoTime);
  crSheet.getRange('B11').setValue(courtReporter);
  crSheet.getRange('D20').setValue(firm);
  crSheet.getRange('B19').setValue(attorney);
  crSheet.getRange('E21').setValue(attorneyEmail);
  crSheet.getRange('A22').setValue(firmAddress1);
  crSheet.getRange('C22').setValue(firmAddress2);
  crSheet.getRange('A23').setValue(city);
  crSheet.getRange('B23').setValue(state);
  crSheet.getRange('C23').setValue(zip);
  crSheet.getRange('C21').setValue(attorneyPhone);
  crSheet.getRange('H57').setValue(orderedBy);
};

/** Updates the Confirmation of Scheduling with the most recently-entered deposition information.
@params {depositionInformation} strings Deposition information received from the sidebar.
*/
function updateConfirmationOfScheduling(locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, depoDate, witnessName, caseStyle, depoTime, courtReporter, firm, attorney, firmAddress1, firmAddress2, city, state, zip, attorneyPhone, orderedBy, videographer, pip) {
  var confSheet = SpreadsheetApp.getActive().getSheetByName('Confirmation of Scheduling');
  
  // Sets values inside CR Worksheet.
  confSheet.getRange('C18').setValue(locationFirm);
  confSheet.getRange('C19').setValue(locationAddress1);
  confSheet.getRange('C20').setValue(locationAddress2);
  confSheet.getRange('C21').setValue(locationCity);
  confSheet.getRange('D21').setValue(locationState);
  confSheet.getRange('E21').setValue(locationZip);
  confSheet.getRange('G16').setValue(depoDate);
  confSheet.getRange('G20').setValue(witnessName);
  confSheet.getRange('C16').setValue(caseStyle);
  confSheet.getRange('G18').setValue(depoTime);
  confSheet.getRange('C22').setValue(courtReporter);
  confSheet.getRange('C8').setValue(firm);
  confSheet.getRange('G22').setValue(attorney);
  confSheet.getRange('G8').setValue(firmAddress1);
  confSheet.getRange('G9').setValue(firmAddress2);
  confSheet.getRange('G10').setValue(city);
  confSheet.getRange('H10').setValue(state);
  confSheet.getRange('I10').setValue(zip);
  confSheet.getRange('E11').setValue(attorneyPhone);
  confSheet.getRange('C10').setValue(orderedBy);
  confSheet.getRange('E22').setValue(videographer);
  confSheet.getRange('D28').setValue(pip);
};

/** Prints an array to the final row of the "Schedule a depo" sheet
@param {array} 1d array ordered to align with the columns in "Schedule a depo."
*/
function printNewDeposition (array) {
  var ss = SpreadsheetApp.getActive();
  var scheduleSheet = ss.getSheetByName('Schedule a depo');
  
  // Create an empty row for the new deposition at the top of the sheet, shift others down by 1, print to the new row
  scheduleSheet.insertRowBefore(2);
  scheduleSheet.getRange(2, 1, 1, 27).setValues(array);
};

/** Takes the most recently-scheduled depo by an orderer and returns an array with the lawyer and firm information.
@param {orderer} string The previous orderer's name as selected from the New Deposition from a Previous Orderer sidebar dropdown menu.
*/
function firmInformationFromOrderer (orderer) {
  var ss = SpreadsheetApp.getActive();
  var scheduleSheet = ss.getSheetByName('Schedule a depo');
  
  // Gets an array of row arrays that match orderer name
  var allScheduledRows = scheduleSheet.getRange(2, 1, scheduleSheet.getLastRow(), scheduleSheet.getLastColumn()).getValues();
  var attyAndFirmInformation = [];
  for (var i = 0; i < allScheduledRows.length; i++) {
    if (allScheduledRows[i][3] === orderer) {
      // allScheduledRows[i][n] because columns 7 - 15 contain the desired information on the "Schedule a depo" sheet
      attyAndFirmInformation.push(allScheduledRows[i][7]);
      attyAndFirmInformation.push(allScheduledRows[i][8]);
      attyAndFirmInformation.push(allScheduledRows[i][9]);
      attyAndFirmInformation.push(allScheduledRows[i][10]);
      attyAndFirmInformation.push(allScheduledRows[i][11]);
      attyAndFirmInformation.push(allScheduledRows[i][12]);
      attyAndFirmInformation.push(allScheduledRows[i][13]);
      attyAndFirmInformation.push(allScheduledRows[i][14]);
      attyAndFirmInformation.push(allScheduledRows[i][15]);
      break;
    };
  };
  return attyAndFirmInformation;
};


/** Gets the email address of a previous orderer from the most recently-scheduled depo from them
@param {orderer} string The previous orderer's name as selected from the New Deposition from a Previous Orderer sidebar dropdown menu.
*/
function emailFromOrderer (orderer) {
  var ss = SpreadsheetApp.getActive();
  var scheduleSheet = ss.getSheetByName('Schedule a depo');
  
  // Gets an array of row arrays that match orderer name
  var allScheduledRows = scheduleSheet.getRange(2, 1, scheduleSheet.getLastRow(), scheduleSheet.getLastColumn()).getValues();
  var ordererEmail = '';
  for (var i = 0; i < allScheduledRows.length; i++) {
    if (allScheduledRows[i][3] === orderer) {
      // allScheduledRows[i][n] because columns 7 - 15 contain the desired information on the "Schedule a depo" sheet
      ordererEmail = allScheduledRows[i][4];
      break;
    };
  };
  
  return ordererEmail;
};











