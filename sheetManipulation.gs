/** Collects data from new deposition sidebars--both from new orderers and repeat orderers
* @params {multiple} strings, bool Values from the form deployed through Google Sheet.
* @return Sequential array of values.
*/
function getDepositionData(orderedBy,orderedByEmail, witnessName, caseStyle, depoDate, depoTime, firm, attorney, attorneyEmail, attorneyPhone, firmAddress1, firmAddress2, city, state, zip, locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, locationPhone, services, courtReporter, videographer, pip) {
  Logger.log(orderedBy); 
  Logger.log(orderedByEmail); 
  Logger.log(witnessName); 
  Logger.log(caseStyle); 
  Logger.log(depoDate); 
  Logger.log(depoTime); 
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
