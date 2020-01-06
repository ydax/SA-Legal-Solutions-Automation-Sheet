////////////////////////////////////////////////////////////////////////////////////
/////////////// DEPLOYING TEMPLATED EMAILS BASED ON SHEET ACTIVITIES ///////////////
////////////////////////////////////////////////////////////////////////////////////

/** Sends an email confirmation to deposition orderer
@params {multiple} strings Arguments passed from the getDepositionData functions originating in the New Depositions sidebars
*/
function sendConfirmationToOrderer(ordererEmail, caseStyle, depoDate, witness, depoHour, depoMinute, amPm, locationFirm, locationAddress1, locationAddress2, locationCity, locationState, locationZip, services, courtReporter, videographer, pip) {
  // Re-format date, time, location
  var date = formatDateForEmail(depoDate);
  var depoTime = depoHour + ':' + depoMinute + ' ' + amPm;
  var depoLocation = locationFirm + ', ' + locationAddress1 + ', ' + locationAddress2 + ', ' + locationCity + ' ' + locationState + ' ' + locationZip;
  
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
    caseStyle + ' Deposition Notice | ' + date, 
    'Hello,\n\nThanks for sending this assignment to SA Legal Solutions. Our understanding of your requested resources & services are detailed below:\n• Case: ' + caseStyle + '\n• Witness: ' + witness + '\n• Date: ' + date + '\n• Time: ' + depoTime + '\n• Location: ' + depoLocation + '\n• Services: ' + services + '\n• Court reporter? ' + reporter + '\n• Videographer? ' + video + '\n• Picture-in-Picture? ' + includesPip + '\n\nIf any changes are necessary, please let us know. Thanks for your business!\n\nSA Legal Solutions | Litigation Support Specialists\nPhone: 210-591-1791\nAddress: 3201 Cherry Ridge, B 208-3, SATX 78230\nWebsite: www.salegalsolutions.com\nEmail: depos@salegalsolutions.com\n\n\n---\nThis e-mail message and any documents attached to it are confidential and may contain information that is protected from disclosure by various federal and state laws, including the HIPAA privacy rule (45 C.F.R., Part 164). This information is intended to be used solely by the entity or individual to whom this message is addressed. If you are not the intended recipient, be advised that any use, dissemination, forwarding, printing, or copying of this message without the sender\'s written permission is strictly prohibited and may be unlawful. Accordingly, if you have received this message in error, please notify the sender immediately by return e-mail or call 210-591-1791, and then delete this message.', 
    {
    name: 'SA Legal Solutions'
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