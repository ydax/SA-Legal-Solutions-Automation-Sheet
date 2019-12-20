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
*/

function myFunction() {
  
}


/** 
REQUIREMENTS
• Streamline the process of entering a new deposition. I plan to do this by building a custom interface that can be activated within the Google Sheet, which can then be used to add new depositions ordered by (1) repeat orderers and (2) new clients, and which will trigger automated population of the Automations Sheet, as well as a calendar event and an automatically-generated confirmation email to the orderer.
• Adding custom search functionality that enables the SA Legal Solutions team to search by deposition date, orderer, and ordering firm.
• Develop automatic sync between the date, time, location, and witness information on the "Schedule a depo" sheet for each deposition row and the deposition schedule on Blake's Google Calendar.
• Add a column to the "Schedule a depo" sheet that tracks the status of a scheduled deposition, and enable users to push data from depositions from the "Schedule a depo" Sheet to the "Current List" Sheet.
• Create automated reporting for Blake that generates a summary of the week's deposition activity, and sends that summary to Blake via email, weekly.

WORKFLOW
1. Modify Sheet Structure: Add Query Sheet w/ results section, modify Confirmation of Scheduling to have a status (done w/ data validation)
2. Add in dev logging functions via properties
3. Add new deposition creation methods, verify they work (incl. population of templates)
4. Add querying features
5. Add in automatically-generated email feature for new depos
6. Add in automatic calendar population for new depos (incl. tag)
7. Add onChange fcn that looks for changes in Schedule a Depo columns, and change calendar event if needed
8. Add data push functionality from Schedule a Depo to Current List
9. Create automatic reporting for Blake



NOTES
- Drop down for new clients vs. existing client w/ ordered by field and ordered by email address (this is who the confirmation email goes to)
  - Existing client needs to have the Ordered By as a field
- Location needs to populate the address 1, city, state, zip, but the location address 2 needs to be manual)
- Services (column V) will be manual, so will column W X
- Search page on the front
   - Searching
       - By Date: Name, Ordered By, Firm
       - By Witness: Date, Ordered By, Firm
- Enable any changes made in the sheet to reflect in the calendar, too (I think this needs to be an onChange trigger)
- Blake Email (once per week)
   - Count of who schedules Ordered By + Firm Name
   - Total of how many depos were scheduled for the week
*/