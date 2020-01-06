////////////////////////////////////////////////////////////////////////////////////
///////////////// INTERACTIONS BETWEEN SHEET AND SERVICES CALENDAR /////////////////
////////////////////////////////////////////////////////////////////////////////////

/** Adds an event to the "Services" calendar and the eventId to the Scheudle a depo Sheet.
@params {multiple} Received from getNewDepositionData in the sheetManipulation module.
*/
function addEvent(orderedBy, witnessName, caseStyle, depoDate, depoHour, depoMinute, amPm, firm, attorney, firmAddress1, firmAddress2, city, state, zip, locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, services, courtReporter, videographer, pip) {
  var SACal = CalendarApp.getCalendarById('salegalsolutions.com_17vfv1akbq03ro6jvtsre0rv84@group.calendar.google.com');
  var ss = SpreadsheetApp.getActive();
  var depoSheet = ss.getSheetByName('Schedule a depo');
  
  // Create event title and description
  var title = '(' + services + ')' + ' ' + firm + ' - ' + witnessName;
  var depoTime = depoHour + ':' + depoMinute + ' ' + amPm;
  var depoLocation = locationFirm + ', ' + locationAddress1 + ' ' + locationAddress2 + ', ' + locationCity + ' ' + locationState + ' ' + locationZip;
  var description = 'Witness Name: ' + witnessName + '\nCase Style: ' + caseStyle + '\nOrdered by: ' + orderedBy + '\n\nCSR: ' +courtReporter + '\nVideographer: ' + videographer + '\nPIP: ' + pip + '\n\nLocation: ' + '\n' + depoLocation + '\n\nOur client:\n' + attorney + '\n' + firm + '\n' + firmAddress1 + ' ' + firmAddress2 + '\n' + city + ' ' + state + ' ' + zip;

  // Add the deposition event to the Services calendar
  var formattedDate = toStringDate(depoDate);
  var formattedHours = to24Format(depoHour, depoMinute, amPm);
  var formattedDateAndHour = formattedDate + ' ' + formattedHours;
  
  var event = SACal.createEvent(title, 
    new Date(formattedDateAndHour),
    new Date(formattedDateAndHour),{
      description: description,
      location: depoLocation
    });
  
  // Add eventId to the Schedule a depo Sheet.
  depoSheet.getRange(2, 37).setValue(event.getId());
};

/** Checks for manual edits to deposition times / dates, updates the calendar event if necessary
@param {Event} e Event object created when an edit is made.
@dev Called by onEdit(e) trigger
*/
function manuallyUpdateCalendar(e) {
  var SACal = CalendarApp.getCalendarById('salegalsolutions.com_17vfv1akbq03ro6jvtsre0rv84@group.calendar.google.com');
  var ss = SpreadsheetApp.getActive();
  var depoSheet = ss.getSheetByName('Schedule a depo');
  
  // Check to see if the edit was made on the "Schedule a depo" sheet
  var sheetName = e.source.getSheetName();
  if (sheetName === 'Schedule a depo') {
    
    // If yes, get information about the edit made
    var editRow = e.range.getRow();
    var editColumn = e.range.getColumn();
    
    // Routing the edit if it was made to the event time. 7 because Start Time is in Column G.
    if (editColumn === 7) {
      // Get information used in date modification
      var oldValue = e.oldValue;
      var newValue = e.value;
      var eventId = depoSheet.getRange(editRow, 37).getValue();
      
      // See if the newly-added time is in HH:MM AM format. If not, reject the change.
      var lastThreeCharacters = newValue.slice(-3);
      var validationCheck = false;
      if (lastThreeCharacters === ' PM' || lastThreeCharacters === ' AM') {
        validationCheck = true;
      };
      if (validationCheck !== true) {
        SpreadsheetApp.getUi().alert('⚠️ Incorrect Format. Please add the new time in Hour:Minute AM/PM format. Note that AM or PM must be capitalized. Your edit to row ' + editRow + ', column ' + editColumn + ' was not saved.')
        depoSheet.getRange(editRow, editColumn).setValue(oldValue);
      } else {
        // Constructs the new time (date obj).
        var newTime = new Date(dateFromHour(newValue, editRow));
        
        // Tries to update Services calendar, alerts user with result.
        try {
          // Deletes old event and adds a new one at the correct time.
          var title = SACal.getEventById(eventId).getTitle();
          var description = SACal.getEventById(eventId).getDescription();
          var location = SACal.getEventById(eventId).getLocation();
          SACal.getEventById(eventId).deleteEvent();
          var event = SACal.createEvent(title, newTime, newTime,{ description: description, location: location });
          
          // Add new eventId to the Schedule a depo Sheet.
          depoSheet.getRange(editRow, 37).setValue(event.getId());
          ss.toast('✅ Services Calendar Updated Successfully');
        } catch (error) {
          SpreadsheetApp.getUi().alert('⚠️ Unable to update Services calendar with this change. The updated deposition time you entered in row ' + editRow + ', column ' + editColumn + ' is NOT reflected on the Services calendar. Please update it manually.');
          addToDevLog('In event time onEdit function: ' + error);
        };
      };
      
    }
    
    // Routing if it was made to event date. 2 because Date is in Column B.
    else if (editColumn === 2) {
      // Gets raw information used in date modification.
      var newValue = e.value;
      var newUnformattedDate = floatToCSTDate(newValue);
      var eventId = depoSheet.getRange(editRow, 37).getValue();
      
      var newTime = new Date(dateFromDate(newUnformattedDate, editRow));
      Logger.log(newTime);
        
      // Tries to update Services calendar, alerts user with result.
      try {
        // Deletes old event and adds a new one at the correct time.
        var title = SACal.getEventById(eventId).getTitle();
        var description = SACal.getEventById(eventId).getDescription();
        var location = SACal.getEventById(eventId).getLocation();
        SACal.getEventById(eventId).deleteEvent();
        var event = SACal.createEvent(title, newTime, newTime,{ description: description, location: location });
        
        // Add new eventId to the Schedule a depo Sheet.
        depoSheet.getRange(editRow, 37).setValue(event.getId());
        ss.toast('✅ Services Calendar Updated Successfully');
      } catch (error) {
        SpreadsheetApp.getUi().alert('⚠️ Unable to update Services calendar with this change. The updated deposition date you entered in row ' + editRow + ', column ' + editColumn + ' is NOT reflected on the Services calendar. Please update it manually.');
        addToDevLog('In event date onEdit function: ' + error);
      };
    };
  };
};

////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////// UTILITIES /////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////

/** Converts hour and am/pm into 24-hour formatted time 
@param {depoHour} string Integer from 1-12 entered by SA Legal Solutions from sidebar.
@param {depoMinute} string Integer from 0-60 entered by SA Legal Solutions from sidebar.
@param (amPm} string AM or PM as entered by SA Legal Solutions from sidebar.
@return Hour in 24-hour format plus central time zone (e.g. 16:30:00 CST).
*/
function to24Format (depoHour, depoMinute, amPm) {
  var hour = parseInt(depoHour, 10);
  if (amPm === 'PM') {
    hour += 12;
  };
  var formattedTime = hour + ':' + depoMinute + ':00 CST';
  return formattedTime;
};

/** Converts YYYY-MM-DD into string date format
@param {depoDate} Date in YYYY-MM-DD format.
@return string Date formatted in full string format (e.g. January 2, 2020).
*/
function toStringDate (depoDate) {
  // Note: incoming format 2020-01-30
  
  // Seperate YYYY, MM, and DD into variables
  var year = depoDate.substring(0, 4);
  var month = depoDate.substring(5, 7);
  var day = depoDate.substring(8, 10);
  
  // Convert month from a two-digit intiger into a full month string
  switch (month) {
    case '01':
      month = 'January';
      break;
    case '02':
      month = 'February';
      break;    
    case '03':
      month = 'March';
      break;    
    case '04':
      month = 'April';
      break;    
    case '05':
      month = 'May';
      break;    
    case '06':
      month = 'June';
      break;    
    case '07':
      month = 'July';
      break;    
    case '08':
      month = 'August';
      break;    
    case '09':
      month = 'September';
      break;    
    case '10':
      month = 'October';
      break;    
    case '11':
      month = 'November';
      break;    
    case '12':
      month = 'December';
      break;
    default:
      Logger.log('Sorry, we are out of months.');
  };
  
  return month + ' ' + day + ', ' + year;
};

/** Converts value from float number to UTC date format.
@param {floatValue} number The integer created when Google Sheets returns the value of a cell with a date or time in it.
See more: https://stackoverflow.com/questions/38815858/google-apps-script-convert-float-to-date
*/
function floatToCSTDate (floatValue) {
  // Converts float value into UTC date (+6 in the 4th argument because Google defaults this Script to CST, and we need GMT for the correct result).
  var rawValue = new Date(Date.UTC(1899, 11, 30, 6, 0, floatValue * 86400)).toString();
  
  return rawValue;
};

/** Generates date obj from hour and onEdit(e) info 
@param {hour} string Hour data formatted HH:MM AM.
@param {editRow} number The row a user has just edited.
@return {date} object A date object reprsenting the new deposition time.
*/
function dateFromHour(hour, editRow) {
  var ss = SpreadsheetApp.getActive();
  var depoSheet = ss.getSheetByName('Schedule a depo');
  
  // Needed Format: 1995-12-17T03:24:00
  var formattedHour = amPmTo24(hour);
  var unformattedDate = depoSheet.getRange(editRow, 2).getValue().toString();
  
  // Destructures and formats date. Incoming format: Mon Jan 06 2020 00:00:00 GMT-0600 (CST)
  var unformattedMonth = unformattedDate.substring(4, 7);
  var formattedMonth = monthToMm(unformattedMonth);
  var day = unformattedDate.substring(8, 10);
  var year = unformattedDate.substring(11, 15);
  
  // Creates date in Needed Format.
  var formattedDate = year + '-' + formattedMonth + '-' + day + 'T' + formattedHour + ':00';
  
  return formattedDate;
};

/** Generates date obj from date and onEdit(e) info 
@param {date} string Stringified date object.
@param {editRow} number The row a user has just edited.
@return {date} object A date object reprsenting the new deposition time.
*/
function dateFromDate (unformattedDate, editRow) {
  var ss = SpreadsheetApp.getActive();
  var depoSheet = ss.getSheetByName('Schedule a depo');

  // Gets hour info and converts it to 24 hour format.
  var hour = depoSheet.getRange(editRow, 7).getValue();
  var formattedHour = amPmTo24(hour);
  
  // Destructures and formats date. Incoming format: Mon Jan 06 2020 00:00:00 GMT-0600 (CST)
  var unformattedMonth = unformattedDate.substring(4, 7);
  var formattedMonth = monthToMm(unformattedMonth);
  var day = unformattedDate.substring(8, 10);
  var year = unformattedDate.substring(11, 15);
  
  // Creates date in Needed Format.
  var formattedDate = year + '-' + formattedMonth + '-' + day + 'T' + formattedHour + ':00';
  
  return formattedDate;
};

/** Converts Hour:Minute AM/PM into 24-hour format 
@param {originalFormat} string Time represented in Hour:Minute AM/PM (e.g. 2:30 PM).
@return {newFormat} string Time represented in 24 hour format (14:30).
*/
function amPmTo24 (originalFormat) {
  
  // Identify the length of the string to determine parsing details.
  var stringLength = originalFormat.length;
  
  // Parse elements of the Hour:Minute AM/PM format.
  switch (stringLength) {
    case 8:
      var hour = parseInt(originalFormat.substring(0, 2));
      var minute = originalFormat.substring(3, 5);
      var amPm = originalFormat.slice(-2);
      break;
    case 7:
      var hour = parseInt(originalFormat.substring(0, 1));
      var minute = originalFormat.substring(2, 4);
      var amPm = originalFormat.slice(-2);
      break;
    default:
      Logger.log('There are no more acceptible lengths');
  };
  
  // Adds 12 hours only for PM times.
  if (amPm === 'PM') {
    hour += 12;
  };
  
  // Adds a zero for integers 1-9.
  hour = hour.toString();
  if (hour.length < 2) {
    hour = '0' + hour;
  }
  
  // Constructs 24-hour format.
  var newFormat = hour + ':' + minute;

  return newFormat;
};

/** Gives developer visibility into accessible Google Calendars */
function seeCalendars () {
  var allCalendars = CalendarApp.getAllCalendars();
  allCalendars.forEach(function(calendar) {
    var id = calendar.getId();
    Logger.log(id);
  });
};

/** Adds each event's eventId to column AK in the "Schedule a depo" Sheet by matching the event's automatically-generated title
Note: This can usually execute ~400 rows per run before exceeding maximum allowable execution time.
*/
function addIds() {
  var SACal = CalendarApp.getCalendarById('salegalsolutions.com_17vfv1akbq03ro6jvtsre0rv84@group.calendar.google.com');
  var ss = SpreadsheetApp.getActive();
  var depoSheet = ss.getSheetByName('Schedule a depo');
  
  // Calibrate returned Google Calendar date range and get events.
  var endTime = new Date('2022-12-01');
  var startTime = new Date ('2018-12-01');
  var events = SACal.getEvents(startTime, endTime);
  
  /** Iterate all rows in Schedule a depo Sheet, search for title matches, and add eventIds.
  Note: change starting value of i to start iteration at a different row.
  */
  for (var i = 2; i < 3; i++) {
    var eventTitle = createTitleString(i);
    events.forEach(function(event) {
      if (event.getTitle() == eventTitle) {
        var eventId = event.getId();
        depoSheet.getRange(i, 37).setValue(eventId);
      };
    });
  };
};

/** Returns a Title string that should match a Calendar Event Title
@param {row} number A row number generated from an iterator in the addIds function.
@return {eventTitle} string The title of an event that would have been automatically generated and added to a specific Google Calendar event.
*/
function createTitleString(row) {
  var ss = SpreadsheetApp.getActive();
  var depoSheet = ss.getSheetByName('Schedule a depo');
  var services = depoSheet.getRange(row, 24).getValue();
  var firmName = depoSheet.getRange(row, 8).getValue();
  var witness = depoSheet.getRange(row, 3).getValue();
  var eventTitle = '(' + services + ')' + firmName + ' - ' + witness;
  return eventTitle;
};