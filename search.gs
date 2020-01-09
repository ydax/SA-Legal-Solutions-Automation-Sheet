////////////////////////////////////////////////////////////////////////////////////
/////////////////////// CUSTOM DEPOSITION SEARCH FUNCTIONS /////////////////////////
////////////////////////////////////////////////////////////////////////////////////


/** Queries the "Schedule a depo" Sheet and creates string to be displayed in modal for user */
function searchByDate() {
  // Prompts the user to enter a search query, captures their response
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("📅 Enter the search date", "Enter a date in MM/DD/YY format\n\n✅ 04/05/2020\n✅ 4/5/21\n✅ 12-05-19\n\n", ui.ButtonSet.OK);
  if (response.getSelectedButton() == ui.Button.OK) {
    var date = response.getResponseText();
  };
  
  // Store the search query in the "Infrastructure" Sheet
  SpreadsheetApp.getActive().getSheetByName('Infrastructure').getRange(8, 2).setValue(date);
  
  displayDateSearch();
};

/** Queries the "Schedule a depo" Sheet and displays results in modal window */
function searchByWitness() {
  // Prompts the user to enter a search query, captures their response
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("👤 Enter witness name", "Name must be an exact match to what\'s\nwritten in Column C of the Schedule a depo Sheet.\n\n", ui.ButtonSet.OK);
  if (response.getSelectedButton() == ui.Button.OK) {
    var witness = response.getResponseText();
  };
  
  // Store the search query in the "Infrastructure" Sheet
  SpreadsheetApp.getActive().getSheetByName('Infrastructure').getRange(8, 2).setValue(witness);
  
  // Generate modal with search results
  displayWitnessSearch();
};

/** Queries the "Schedule a depo" Sheet and creates string to be displayed in modal for user */
function searchByCase() {
  // Prompts the user to enter a search query, captures their response
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("⚖️ Enter your case style search", "Your search query must match the Case Style in Column F\nof the Schedule a depo Sheet exactly\n\n✅ Tim Duncan v. Michael Jordan\n\n", ui.ButtonSet.OK);
  if (response.getSelectedButton() == ui.Button.OK) {
    var query = response.getResponseText();
  };
  
  // Store the search query in the "Infrastructure" Sheet
  SpreadsheetApp.getActive().getSheetByName('Infrastructure').getRange(9, 2).setValue(query);
  
  displayCaseSearch();
};

/** Returns a string with witness name search results from "Schedule a depo" Sheet
@return string HTML-formatted string containing results of search query
*/
function searchWitness () {
  var query = SpreadsheetApp.getActive().getSheetByName('Infrastructure').getRange(8, 2).getValue();
  var results = [];
  var ss = SpreadsheetApp.getActive();
  var scheduleSheet = ss.getSheetByName('Schedule a depo');
  
  // Grabs data from the "Schedule a depo" Sheet, iterates over each row, and pushes results to an array
  var allScheduledData = scheduleSheet.getRange(2, 1, scheduleSheet.getLastRow(), scheduleSheet.getLastColumn()).getValues();
  for (var i = 0; i < allScheduledData.length; i++) {
    if (allScheduledData[i][2] == query) {
      results.push(allScheduledData[i]);
    };
  };

  
  // Determines whether results were found, prints response on Search Sheet 
  if (results.length == 0) {
    return '⚠️ No results found for ' + query + '. Make sure the name is written exactly as in the "Schedule a depo" Sheet.'
  } else {
    // Store found search elements in variables
    results = results[0];
    var rawDate = results[1].toString();
    var date = rawDate.substring(0, 15);
    var witness = results[2];
    var firm = results[7];
    var orderer = results[3];
    var ordererEmail = results[4];
    var caseStyle = results[5];
    var services = results[23];
    var courtReporter = results[24];
    var videographer = results[25];
    var pip = results[26];
    var depoCity = results[19];
    var depoState = results[20];
    
    // Generate a string with the desired variables for display through modal
    return '1️⃣ Key Information:<br>Deposition date: <em>' + date + '</em><br>Witness: <em>' + witness + '</em><br>Firm: <em> '+ firm + '</em><br>Ordered by: <em>' + orderer + '</em><br><br>2️⃣ Additional Information:<br>' + '</em>Orderer email: <em>' + ordererEmail + '</em><br>Case Style: <em>' + caseStyle + '</em><br>Services: <em>' + services + '</em><br>Court reporter? <em>' + courtReporter + '</em><br>Videographer? <em>' + videographer + '</em><br>PIP? <em>' + pip + '</em><br>City: <em>' + depoCity + '</em><br>State: <em>' + depoState + '</em>';
  };
};


/** Returns an array of date search results from "Schedule a depo" Sheet
@return string HTML-formatted string that can be injected into a modal
*/
function searchDate () {
  // Grabs data from the "Schedule a depo" Sheet, iterates over each row, and pushes results to an array
  var results = [];
  var ss = SpreadsheetApp.getActive();
  var scheduleSheet = ss.getSheetByName('Schedule a depo');
  var allScheduledData = scheduleSheet.getRange(2, 1, scheduleSheet.getLastRow(), scheduleSheet.getLastColumn()).getValues();
  
  // Get query from the "Infrastructure" sheet
  var query = SpreadsheetApp.getActive().getSheetByName('Infrastructure').getRange(8, 2).getValue().toString();
  
  // Parse date query into month, day, and year
  var queryMonth = monthToMm(query.substring(4, 7));
  var queryDay = query.substring(8, 10);
  var queryYear = query.substring(11, 15);
  
  // Grabs data from the "Schedule a depo" Sheet, iterates over each row, and pushes results to an array
  var allScheduledData = scheduleSheet.getRange(2, 1, scheduleSheet.getLastRow(), scheduleSheet.getLastColumn()).getValues();
  for (var i = 0; i < allScheduledData.length; i++) {
    var isMatch = matchDate(queryMonth, queryDay, queryYear, allScheduledData[i][1]);
    if (isMatch == true) {
      results.push(allScheduledData[i]);
    };
  };
  
  var resultString = ''; 
  
  // Determines whether results were found, prints response on Search Sheet 
  if (results.length == 0) {
    var rawDate = query.toString();
    var date = rawDate.substring(0, 15);
    return '⚠️ No depositions found matching ' + date + '. Make sure the date entered is in month/day/year format.';
  } else {

    // Format the string contribution for each result    
    for (var i = 0; i < results.length; i++) {
      var resultCount = i + 1;
      resultString += '<strong>↩ Result ' + resultCount + '</strong><br>';
      var rawDate = results[i][1].toString();
      var date = rawDate.substring(0, 15);
      var witness = results[i][2];
      var orderer = results[i][3];
      var firm = results[i][7];
      var ordererEmail = results[i][4];
      var caseStyle = results[i][5];
      var services = results[i][23];
      var courtReporter = results[i][24];
      var videographer = results[i][25];
      var pip = results[i][26];
      var depoCity = results[i][19];
      var depoState = results[i][20];
      resultString += '1️⃣ Key Information:<br>Witness: <em>' + witness + '</em><br>Firm: <em> '+ firm + '</em><br>Ordered by: <em>' + orderer + '</em><br><br>2️⃣ Additional Information:<br>' + '</em>Orderer email: <em>' + ordererEmail + '</em><br>Case Style: <em>' + caseStyle + '</em><br>Services: <em>' + services + '</em><br>Court reporter? <em>' + courtReporter + '</em><br>Videographer? <em>' + videographer + '</em><br>PIP? <em>' + pip + '</em><br>City: <em>' + depoCity + '</em><br>State: <em>' + depoState + '</em><br><br>';
    };
  };
  return resultString;
};

/** Returns an array of case style search results from "Schedule a depo" Sheet
@return string HTML-formatted string that can be injected into a modal
*/
function searchCase () {
 // Grabs data from the "Schedule a depo" Sheet, iterates over each row, and pushes results to an array
  var results = [];
  var ss = SpreadsheetApp.getActive();
  var scheduleSheet = ss.getSheetByName('Schedule a depo');
  var allScheduledData = scheduleSheet.getRange(2, 1, scheduleSheet.getLastRow(), scheduleSheet.getLastColumn()).getValues();
  
  // Get case style query consumed from user input and added to the "Infrastructure" sheet
  var query = SpreadsheetApp.getActive().getSheetByName('Infrastructure').getRange(9, 2).getValue().toString();
  
  // Grabs data from the "Schedule a depo" Sheet, iterates over each row, and pushes results to an array
  var allScheduledData = scheduleSheet.getRange(2, 1, scheduleSheet.getLastRow(), scheduleSheet.getLastColumn()).getValues();
  for (var i = 0; i < allScheduledData.length; i++) {
    if(allScheduledData[i][5] == query) {
      results.push(allScheduledData[i]);
    };
  };
  
  var resultString = ''; 
  
  // Determines whether results were found, prints response on Search Sheet 
  if (results.length == 0) {
    var rawDate = query.toString();
    var date = rawDate.substring(0, 15);
    return '⚠️ No depositions found matching ' + query + '. Make sure that you entered the case style query exactly as it\'s written in the Schedule a depo Sheet.';
  } else {

    // Format the string contribution for each result    
    for (var i = 0; i < results.length; i++) {
      var resultCount = i + 1;
      resultString += '<strong>↩ Result ' + resultCount + '</strong><br>';
      var rawDate = results[i][1].toString();
      var date = rawDate.substring(0, 15);
      var witness = results[i][2];
      var orderer = results[i][3];
      var firm = results[i][7];
      var ordererEmail = results[i][4];
      var caseStyle = results[i][5];
      var services = results[i][23];
      var courtReporter = results[i][24];
      var videographer = results[i][25];
      var pip = results[i][26];
      var depoCity = results[i][19];
      var depoState = results[i][20];
      resultString += '1️⃣ Key Information:<br>Deposition Date: <em>' + date + '</em><br>Witness: <em>' + witness + '</em><br>Firm: <em> '+ firm + '</em><br>Ordered by: <em>' + orderer + '</em><br><br>2️⃣ Additional Information:<br>' + '</em>Orderer email: <em>' + ordererEmail + '</em><br>Services: <em>' + services + '</em><br>Court reporter? <em>' + courtReporter + '</em><br>Videographer? <em>' + videographer + '</em><br>PIP? <em>' + pip + '</em><br>City: <em>' + depoCity + '</em><br>State: <em>' + depoState + '</em><br><br>';
    };
  };
  return resultString;
};


////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////// DISPLAY FUNCTIONS /////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////
function displayDateSearch() {
  var template = HtmlService.createTemplateFromFile('displayDateSearch');
  var rawDate = SpreadsheetApp.getActive().getSheetByName('Infrastructure').getRange(8, 2).getValue().toString();
  var date = rawDate.substring(0, 15);
  var headerString = '🔎 Search results for ' + date;
  template.automationInfo = searchDate();
  var html = template.evaluate().setTitle('🔎 Search by Date Results')
    .setWidth(800)
    .setHeight(800);
  SpreadsheetApp.getUi() 
  .showModalDialog(html, headerString);
};

function displayWitnessSearch() {
  var template = HtmlService.createTemplateFromFile('displayWitnessSearch');
  var headerString = '🔎 Search results for ' + SpreadsheetApp.getActive().getSheetByName('Infrastructure').getRange(8, 2).getValue();
  template.automationInfo = searchWitness();
  var html = template.evaluate().setTitle('🔎 Search by Witness Results')
    .setWidth(800)
    .setHeight(800);
  SpreadsheetApp.getUi() 
  .showModalDialog(html, headerString);
};

function displayCaseSearch() {
  var template = HtmlService.createTemplateFromFile('displayCaseSearch');
  var caseStyle = SpreadsheetApp.getActive().getSheetByName('Infrastructure').getRange(9, 2).getValue()
  var headerString = '🔎 Search results for ' + caseStyle;
  template.automationInfo = searchCase();
  var html = template.evaluate().setTitle('🔎 Search by Case Results')
    .setWidth(800)
    .setHeight(800);
  SpreadsheetApp.getUi() 
  .showModalDialog(html, headerString);
};

////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////// UTILITIES /////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////
/** Quickly searches for matching date in MM/DD/YYYY format
@params {month, day, year} strings Strings of integers
@param {comparisonDate} object Date object used in comparisons
@return bool
*/
function matchDate (month, day, year, comparisonDate) {
  // Converts date object into a string, then formats into comparisonDay (DD), comparisonMonth (MM), and comparisonYear (YYYY) strings of integers
  var stringifiedComparisonDate = comparisonDate.toString();
  var comparisonMonthRaw = stringifiedComparisonDate.substring(4, 7);
  var comparisonMonth = monthToMm(comparisonMonthRaw);
  var comparisonDay = stringifiedComparisonDate.substring(8, 10);
  var comparisonYear = stringifiedComparisonDate.substring(11, 15);
  
  // Return bool based on full matches only
  if (month === comparisonMonth && day === comparisonDay && year === comparisonYear) {
    return true;
  } else {
    return false;
  };
};

/** Converts three-character month into two-integer string 
@param {month} string Three-character month created from date substring (e.g. "Jan").
@return {MM} string Two-integer month string (e.g. "01").
*/
function monthToMm (month) {
  switch (month) {
    case 'Jan':
      return '01';
      break;
    case 'Feb':
      return '02';
      break;    
    case 'Mar':
      return '03';
      break;    
    case 'Apr':
      return '04';
      break;    
    case 'May':
      return '05';
      break;    
    case 'Jun':
      return '06';
      break;    
    case 'Jul':
      return '07';
      break;    
    case 'Aug':
      return '08';
      break;    
    case 'Sep':
      return '09';
      break;    
    case 'Oct':
      return '10';
      break;    
    case 'Nov':
      return '11';
      break;    
    case 'Dec':
      return '12';
      break;
    default:
      Logger.log('Sorry, we are out of months.');
  };
};

