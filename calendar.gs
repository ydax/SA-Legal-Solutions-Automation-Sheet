function testAddEvent() {
  addEvent('Davis Jones', 'Charles Burkes', 'Burkes vs. Jones', '2020-01-02', '3', '30', 'PM', 'Jones & Jones', 'Sarah Jones', '2510 Quarry Road', '', 'Austin', 'TX', '78703', 'Aggie Law Firm', '1011 Wonder World Drive', '#1909', 'San Marcos', 'TX', '78666', 'CR + Video', 'Ludell Jones', 'Jamie F.', 'Yes');
}

/** Adds an event to the "Services" calendar.
@params {multiple} Received from getNewDepositionData in the sheetManipulation module.
*/
function addEvent(orderedBy, witnessName, caseStyle, depoDate, depoHour, depoMinute, amPm, firm, attorney, firmAddress1, firmAddress2, city, state, zip, locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, services, courtReporter, videographer, pip) {
  var SACal = CalendarApp.getCalendarById('salegalsolutions.com_17vfv1akbq03ro6jvtsre0rv84@group.calendar.google.com');
  
  // Create event title and description
  var title = '(' + services + ')' + ' ' + firm + ' - ' + witnessName;
  var depoTime = depoHour + ':' + depoMinute + ' ' + amPm;
  var depoLocation = locationFirm + ', ' + locationAddress1 + ' ' + locationAddress2 + ', ' + locationCity + ' ' + locationState + ' ' + locationZip;
  var description = 'Witness Name: ' + witnessName + '\nCase Style: ' + caseStyle + '\nOrdered by: ' + orderedBy + '\n\nCSR: ' +courtReporter + '\nVideographer: ' + videographer + '\nPIP: ' + pip + '\n\nLocation: ' + '\n' + depoLocation + '\n\nOur client:\n' + attorney + '\n' + firm + '\n' + firmAddress1 + ' ' + firmAddress2 + '\n' + city + ' ' + state + ' ' + zip;

  // Add the deposition event to the Services calendar
  var formattedDate = toStringDate(depoDate);
  var formattedHours = to24Format(depoHour, depoMinute, amPm);
  var formattedDateAndHour = formattedDate + ' ' + formattedHours;
  
  Logger.log(formattedDateAndHour);
  
  var event = SACal.createEvent(title, 
    new Date(formattedDateAndHour),
    new Date(formattedDateAndHour),{
      description: description,
      location: depoLocation
    });
  Logger.log('Event ID: ' + event.getId());
};

function seeCalendars () {
  var allCalendars = CalendarApp.getAllCalendars();
  allCalendars.forEach(function(calendar) {
    var id = calendar.getId();
    Logger.log(id);
  });
};

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

