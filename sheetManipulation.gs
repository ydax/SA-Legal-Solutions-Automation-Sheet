/** Collects data from new orderer deposition sidebar
* @params {multiple} strings, bool Values from the form deployed through Google Sheet.
* @return Sequential array of values.
*/
function getNewDepositionData(orderedBy,orderedByEmail, witnessName, caseStyle, depoDate, depoHour, depoMinute, amPm, firm, attorney, attorneyEmail, attorneyPhone, firmAddress1, firmAddress2, city, state, zip, locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, locationPhone, services, courtReporter, videographer, pip) {
  Logger.log(orderedBy); 
  Logger.log(orderedByEmail); 
  Logger.log(witnessName); 
  Logger.log(caseStyle); 
  Logger.log(depoDate); 
  Logger.log(depoHour);
  Logger.log(depoMinute);
  Logger.log(amPm);
  Logger.log(firm); 
  Logger.log(attorney); 
  Logger.log(attorneyEmail); 
  Logger.log(attorneyPhone); 
  Logger.log(firmAddress1); 
  Logger.log(firmAddress2); 
  Logger.log(city); 
  Logger.log(state); 
  Logger.log(zip); 
  Logger.log(locationAddress1); 
  Logger.log(locationAddress2); 
  Logger.log(locationCity); 
  Logger.log(locationState); 
  Logger.log(locationZip); 
  Logger.log(locationPhone); 
  Logger.log(services); 
  Logger.log(courtReporter); 
  Logger.log(videographer); 
  Logger.log(pip); 
  return 'Success';
};

/** Collects data from repeat orderer deposition sidebar
* @params {multiple} strings, bool Values from the form deployed through Google Sheet.
* @return Sequential array of values.
*/
function getRepeatDepositionData(previousOrderer, witnessName, caseStyle, depoDate, depoHour, depoMinute, amPm, locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, locationPhone, services, courtReporter, videographer, pip) {
  SpreadsheetApp.getActiveSpreadsheet().toast('🚀️ Automation initiated');
  var depoTime = depoHour + ':' + depoMinute + ' ' + amPm;
  var newScheduledDepo = [previousOrderer, witnessName, caseStyle, depoTime];
  
  /** Gets firm and attorney information from previous orderer, pushes it into the newScheduledDepo array */
  var infoFromPreviousOrderer = firmInformationFromOrderer(previousOrderer);
  for (var i = 0; i < infoFromPreviousOrderer.length; i++) {
    newScheduledDepo.push(infoFromPreviousOrderer[i]);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('✔️ Found attorney and firm info');

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
  
  Logger.log(newScheduledDepo);
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
  
  var firstLevelEmptiesRemoved = firstLevelArray.filter(isntEmpty);
  
  var uniqueArray = firstLevelEmptiesRemoved.filter(function(elem, index, self) {
    return index === self.indexOf(elem);
  });
  
  var sortedUniqueArray = uniqueArray.sort();
  
  // Array cleaning functions
  function isntEmpty (element) {
  return element != '';
  };
  
  return sortedUniqueArray;
};

/** Prints an array to the final row of the "Schedule a depo" sheet -- TODO
@param {array} 1d array ordered to align with the columns in "Schedule a depo."
*/
function printNewDeposition (array) {
  var ss = SpreadsheetApp.getActive();
  var scheduleSheet = ss.getSheetByName('Schedule a depo');
  
  // Create an empty row for the new deposition at the top of the sheet, shift others down by 1
  scheduleSheet.insertRowBefore(2);
  
  
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
      // allScheduledRows[i][n] because columns 6 - 14 contain the desired information on the "Schedule a depo" sheet
      attyAndFirmInformation.push(allScheduledRows[i][6]);
      attyAndFirmInformation.push(allScheduledRows[i][7]);
      attyAndFirmInformation.push(allScheduledRows[i][8]);
      attyAndFirmInformation.push(allScheduledRows[i][9]);
      attyAndFirmInformation.push(allScheduledRows[i][10]);
      attyAndFirmInformation.push(allScheduledRows[i][11]);
      attyAndFirmInformation.push(allScheduledRows[i][12]);
      attyAndFirmInformation.push(allScheduledRows[i][13]);
      attyAndFirmInformation.push(allScheduledRows[i][14]);
      break;
    };
  };
  return attyAndFirmInformation;
};














