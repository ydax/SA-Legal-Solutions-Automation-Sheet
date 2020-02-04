/** SA Legal Solutions Automation Station Codebase
* GitHub repo github.com/ydax/SA-Legal-Solutions-Automation-Sheet
* @dev Davis Jones | github.com/ydax | davis@eazl.co
* SA Legal Solutions POC Blake Boyd | bboyd@salegalsolutions.com
* Color Palette
  Use              HEX     MaterializeCSS
  Primary          #c62828 red darken-3
  Cell Background  #ffebee red lighten-5
  Confirmation     #00796b teal darken-2
  Error            #e65100 orange darken-4
  Primary -1       #e53935 red darken-1
  Primary -2       #e57373 red lighten-2
  Primary +1       #b71c1c red darken-4
*/

////////////////////////////////////////////////////////////////////////////////////
////////////// CREATION OF SPREADSHEET MENU PLUS USER INTERFACE CALLS //////////////
////////////////////////////////////////////////////////////////////////////////////

/** Creates the SA Legal Solutions menu.
@param {e} object Sheet load event object.
*/
function onOpen (e) {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("⚖️ SA Legal Services")
  .addSubMenu(SpreadsheetApp.getUi().createMenu("📝 Add Deposition(s)")
              .addItem("🔁 Repeat Orderer", "initiateRepeatOrdererModal")
              .addItem("🆕 New Orderer", "initiateNewOrdererModal"))
  .addSubMenu(SpreadsheetApp.getUi().createMenu("🔎 Search")
              .addItem("📅 By Date", "searchByDate")
              .addItem("👤 By Witness", "searchByWitness")
              .addItem("⚖️ By Case", "searchByCase"))
  .addItem("📋 Update Worksheets", "updateWorksheetsByRow")
  .addItem("📧 Re-send Conf. Email (by Row)", "launchConfEmailModal")
  .addToUi();
};

// Launches the new orderer sidebar (triggered from modal).
function launchNewOrdererSidebar() {
  var template = HtmlService.createTemplateFromFile('newOrderer');
  template.copyAttys = getCopyAttorneys();
  template.previousLocations = getPreviousLocations();
  var html = template.evaluate().setTitle('🆕 New Deposition from a New Orderer');
  SpreadsheetApp.getUi().showSidebar(html);
};

// Launches the new orderer deposition modal.
function initiateNewOrdererModal() {
  var html = HtmlService.createHtmlOutputFromFile('newOrdererM')
    .setWidth(350)
    .setHeight(105);
  SpreadsheetApp.getUi() 
    .showModalDialog(html, '👥 Getting previous orderers...');
};

// Initiates the repeat orderer modal.
function initiateRepeatOrdererModal() {
  var html = HtmlService.createHtmlOutputFromFile('repeatOrdererM')
    .setWidth(350)
    .setHeight(105);
  SpreadsheetApp.getUi() 
    .showModalDialog(html, '👥 Getting previous orderers...');
};

// Creates the repeat orderer sidebar (triggered from modal).
function launchRepeatOrdererSidebar() {
  var template = HtmlService.createTemplateFromFile('repeatOrderer');
  template.orderers = getPreviousOrderers();
  template.previousLocations = getPreviousLocations();
  template.copyAttys = getCopyAttorneys();
  var html = template.evaluate().setTitle('🔁 New Deposition from a Repeat Orderer');
  SpreadsheetApp.getUi().showSidebar(html);
};

// Launches resending confirmation email modal.
function launchConfEmailModal() {
  var html = HtmlService.createHtmlOutputFromFile('resendConfM')
    .setWidth(350)
    .setHeight(105);
  SpreadsheetApp.getUi() 
    .showModalDialog(html, '🤖 Resending confirmation...');
};

// Runs an unimportant function when user isn't depos@salegalsolutions.com
function checkLogin () {
  Logger.log(Session.getActiveUser().getEmail());
};


////////////////////////////////////////////////////////////////////////////////////
//////////////////////////// APPLICATION DEVELOPMENT LOG ///////////////////////////
////////////////////////////////////////////////////////////////////////////////////
/** 
--- Version 1.1 Modifications, Started Thursday, January 9th ---
X • Add a Search by Case Style function that returns all depositions for that case, with the same information as the Search by Date function
X • Enable synching between the Schedule a depo Sheet's deposition location and services information and the Services Calendar > ALL fields in the calendar event are editable from the sheet now
X • Make the confirmation email optional in the sidebars > this will default come from the logged in user of the Sheet
X • Have the confirmation emails come from depos@salegalsolutions.com, and bcc shannonk@salegalsolutions.com
X • Expand the sidebars to add Copy Attorney information: firm name, attorney, their address (Columns B:J on the Schedule a depo Sheet)
X • Add the internal record-keeping fields on the CR Worksheet and Video Worksheet at top
X • On the addition of a new deposition, automatically populate Columns A:K on the Current List Sheet and set Status (Column A) to Current. Reduce Status options to "Current" and "Cancelled" only.
X • On deposition Cancel in the Schedule a depo Sheet: remove Calendar event from Services, add CANCELED in front of the title, and add it to the Cancelled Calendar, and remove it from the Current List Sheet
X • Enable date changes from Schedule a depo Sheet to reflect on the Current List as well
X • On date and time changes made to the Schedule a depo Sheet, auto-populate the worksheets again
X • If the logged in user isn’t depos@salegalsolutions.com, remove the automation options
X • Add templated pdf to be sent with new deposition confirmations
X • Enable HTML editing of confirmation email template


--- Version 1.0, Started on Friday, December 20th 2019 ---
REQUIREMENTS
• Streamline the process of entering a new deposition. I plan to do this by building a custom interface that can be activated within the Google Sheet, which can then be used to add new depositions ordered by (1) repeat orderers and (2) new clients, and which will trigger automated population of the Automations Sheet, as well as a calendar event and an automatically-generated confirmation email to the orderer.
• Adding custom search functionality that enables the SA Legal Solutions team to search by deposition date, orderer, and ordering firm.
• Develop automatic sync between the date, time, location, and witness information on the "Schedule a depo" sheet for each deposition row and the deposition schedule on Blake's Google Calendar.
• Add a column to the "Schedule a depo" sheet that tracks the status of a scheduled deposition, and enable users to push data from depositions from the "Schedule a depo" Sheet to the "Current List" Sheet.
• Create automated reporting for Blake that generates a summary of the week's deposition activity, and sends that summary to Blake via email, weekly.

WORKFLOW
X 1. Modify Sheet Structure: Add Query Sheet w/ results section, modify Confirmation of Scheduling to have a status (done w/ data validation)
X 2. Add in dev logging functions via properties
X 3. Add new deposition creation methods, verify they work (incl. population of templates)
X 3.1 Modify Sheet, sidebar template, and back-end functions to include orderer email address
X 4. Add querying features
X 5. Add in automatically-generated email feature for new depos
X 6. Add in automatic calendar population for new depos (incl. tag)
X 7. Add onChange fcn that looks for changes in Schedule a Depo columns, and change calendar event if needed
X 8. Add data push functionality from Schedule a Depo to Current List
X 9. Create automatic reporting for Blake
X 10. Create documentation


NOTES
- Drop down for new clients vs. existing client w/ ordered by field and ordered by email address (this is who the confirmation email goes to)
  - Existing client needs to have the Ordered By as a field
- Location needs to populate the address 1, city, state, zip, but the location address 2 needs to be manual)
- Services (column V) will be manual, so will column W
- Search page on the front
   - Searching
       - By Date: Witness, Ordered By, Firm
       - By Witness: Date, Ordered By, Firm
- Enable any changes made in the sheet to reflect in the calendar, too (I think this needs to be an onChange trigger)
- Blake Email (once per week)
   - Count of who schedules Ordered By + Firm Name
   - Total of how many depos were scheduled for the week


--- Version 1.1, Started on Tuesday, January 28th 2020 ---
REQUIREMENTS:
X 1. Add feature that enables user to select a row, then force populate the worksheets
X 2. Add dropdown to sidebars where previous copy attorneys can be selected
X 3. Enable mirroring of date, location, services, and time from Schedule a depo to Current List tabs on update by users
   > date already works
   > time isn't on the current list
   > location isn't on the current list
   > services aren't on the current list
X 4. Add a dropdown to sidebars with previous locations
X 5. Hunt and remove bug causing Blake to get an extra report on Saturday
*/























