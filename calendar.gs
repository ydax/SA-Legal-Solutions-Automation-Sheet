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
      var oldValue = floatToCSTDate(e.oldValue);
      var newValue = floatToCSTDate(e.value);
      var oldTime = oldValue.substring(16, 25);
      var newTime = newValue.substring(16, 25);
      var eventId = depoSheet.getRange(editRow, 37).getValue();
      var event = SACal.getEventById(eventId).getTitle();
      Logger.log(event);
    }
    
    // Routing if it was made to event date. 2 because Date is in Column B.
    else if (editColumn === 2) {
      Logger.log('Date');
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

/** Converts value from float number to CST date format
@param {floatValue} number The integer created when Google Sheets returns the value of a cell with a date or time in it.
See more: https://stackoverflow.com/questions/38815858/google-apps-script-convert-float-to-date
*/
function floatToCSTDate (floatValue) {
  // Converts float value into date CST date (hence -6)
  var rawValue = new Date(Date.UTC(1899, 11, 30, -6, 0, floatValue * 86400)).toString();
  
  return rawValue;
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
  for (var i = 932; i < 1080; i++) {
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