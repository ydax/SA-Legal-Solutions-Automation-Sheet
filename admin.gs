////////////////////////////////////////////////////////////////////////////////////
/////////////////// GIVING BLAKE VISIBILITY INTO CLIENT BEHAVIOR ///////////////////
////////////////////////////////////////////////////////////////////////////////////

/** Logs order behavior.
@param {orderer} string The orderer of a deposition.
@param {firm} string The firm the orderer works for.
*/
function addOrderToLog(orderer, firm) {
  var scriptProps = PropertiesService.getScriptProperties();
  var keys = scriptProps.getKeys();
  var alreadyExists = false;
  var key = '';
  
  // Structures the orderer / firm information so that the App can see if there are already orderes from this during this reporting period.
  var ordererInfoFromApp = orderer + ' (' + firm + ')';
  
  // If keys.length > 0, look for previous orders and take action.
  if (keys.length > 0) {
  
    // Cycle through values, see if there's already an instance of this orderer + firm.
    for (var i = 0; i < keys.length ; i++) {
      var ordererInfoFromProps = parseOrdererInfo(keys[i]);
      // If there's already an instance, increase the count in props and store it
      if (ordererInfoFromApp === ordererInfoFromProps) {
        alreadyExists = true;
        key = keys[i];
      };
    };

    // If previous orders exist, add one to the count and replace the value.    
    if (alreadyExists === true) {
        var count = parseOrdererCount(key);
        count++;
        var newValue = '#O#' + orderer + '#F#' + firm + '#C#' + count;
        scriptProps.setProperty(key, newValue);
        
        // If no existing instance, create a new record
    } else {
      // Create order key string using #O# as the pattern for the first three characters concatenated with the now ISO string.
      var now = new Date().toISOString();
      var key = '#K#' + now;
      
      // Create value string using #O#ordererName#F#firmName#C#count pattern and store it.
      var value = '#O#' + orderer + '#F#' + firm + '#C#' + '1';
      scriptProps.setProperty(key, value);
    };

  // If keys.length = 0, record the first order instance for this reporting period.        
  } else {
    // Create order key string using #O# as the pattern for the first three characters concatenated with the now ISO string.
    var now = new Date().toISOString();
    var key = '#K#' + now;
    
    // Create value string using #O#ordererName#F#firmName#C#count pattern and store it.
    var value = '#O#' + orderer + '#F#' + firm + '#C#' + 1;
    scriptProps.setProperty(key, value);
  }; 
};

/** Cycles through Script Properties data, structures an email for Blake, sends it, and clears the datastore. */
function sendOrderActivityReport() {
  var props = PropertiesService.getScriptProperties();
  var keys = props.getKeys();
  var reportText = 'Here\'s the ordering activity for this week:\n\nCount   Orderer\n';
  
  try {
    if (keys.length != 0) {
      // Structure email with top orderers first. 50 because the App assumes nobody is going to order more than 50x in a week.
      for (var i = 35; i > 0; i--) {
        for (var j = 0; j < keys.length; j++) {
          if(parseOrdererCount(keys[j]) === i) {
            reportText += i + '           ' + parseOrdererInfo(keys[j]) + '\n';
          };
        };  
      };
    };
    
    var date = toStringDate(new Date().toISOString());
    
    // Sends the email to Blake.
    GmailApp.sendEmail('bboyd@salegalsolutions.com', 'Order Activity Report for Week Prior to ' + date, reportText, { name: 'SALS Reporting Bot'});
    
    // Resets the Script Properties.
    deleteScriptProps();
    
  } catch (error) {
    addToDevLog('Error inside sendOrderActivityReport: ' + error);
  };
};


////////////////////////////////////////////////////////////////////////////////////
//////// STORING AND RETREIVING PREVIOUS ORDERS, COPY ATTYS, AND LOCATIONS /////////
////////////////////////////////////////////////////////////////////////////////////

function updateInfrastructure() {
  storeCopyAttys()
  storePreviousLocations()
  storePreviousOrderers()
}

// Saves a stringified array of filtered copy attys in Infrastructure sheet
function storeCopyAttys() {
  const stringCell = SpreadsheetApp.getActive().getSheetByName('Infrastructure').getRange(14, 2)
  const deposSheet = SpreadsheetApp.getActive().getSheetByName('Schedule a depo')
  
  /** Loops through schedule a depo and does magic to create an array of copy attys ready to be stringified and stored */
  const rawCopyAttys = deposSheet.getRange(2, 28, deposSheet.getLastRow(), 9).getValues()
  const namesArray = []
  const filteredArray = []
  
  rawCopyAttys.forEach(function(atty) {
    if (atty[0] !== '') {
      if (!namesArray.some(name => name === atty[0])) {
        namesArray.push(atty[0])
        filteredArray.push(atty)
      }
    }
  })
  
  // Sorts the array of copy attorney arrays alphabetically by name
  var sortedCopyAttys = filteredArray.sort(function(a, b) {

    var nameA = a[0];
    var nameB = b[0];
    
    if (nameA < nameB) {
      return -1;
    };
    if (nameA > nameB) {
      return 1;
    };
    
    return 0;
  });
  
  // Stores the array in 
  let value = JSON.stringify(sortedCopyAttys)
  stringCell.setValue(value)
}

// Gets array of copy attys / called by sidebar initiator functions used to add depos
function returnCopyAttys() {
  const stringCell = SpreadsheetApp.getActive().getSheetByName('Infrastructure').getRange(14, 2).getValue().toString()
  return JSON.parse(stringCell)
}

// Saves a stringified array of filtered previous locations in Infrastructure sheet
function storePreviousLocations() {
  const stringCell = SpreadsheetApp.getActive().getSheetByName('Infrastructure').getRange(15, 2)
  const deposSheet = SpreadsheetApp.getActive().getSheetByName('Schedule a depo')
  
  /** Loops through schedule a depo and does magic to create an array of copy attys ready to be stringified and stored */
  var rawLocations = deposSheet.getRange(2, 17, deposSheet.getLastRow(), 7).getValues();
  const namesArray = []
  const filteredArray = []
  
  rawLocations.forEach(function(location) {
    if (location[0] !== '') {
      if (!namesArray.some(name => name === location[0])) {
        namesArray.push(location[0])
        filteredArray.push(location)
      }
    }
  })
  
  // Sorts the array of copy attorney arrays alphabetically by name
  var sortedLocations = filteredArray.sort(function(a, b) {

    var nameA = a[0];
    var nameB = b[0];
    
    if (nameA < nameB) {
      return -1;
    };
    if (nameA > nameB) {
      return 1;
    };
    
    return 0;
  });
  
  // Stores the array in 
  let value = JSON.stringify(sortedLocations)
  stringCell.setValue(value)
}

// Gets array of locations / called by sidebar initiator functions used to add depos
function returnLocations() {
  const stringCell = SpreadsheetApp.getActive().getSheetByName('Infrastructure').getRange(15, 2).getValue().toString()
  return JSON.parse(stringCell)
}

// Stores a clean (no dupes, sorted by First Name) array of previous  orderers as a JSON string in infrastructure for fast recall
function storePreviousOrderers () {
  var ss = SpreadsheetApp.getActive()
  const stringCell = SpreadsheetApp.getActive().getSheetByName('Infrastructure').getRange(16, 2)
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
  
  stringCell.setValue(JSON.stringify(sortedUniqueArray))
};

// Gets array of previous orderers / called by sidebar initiator functions used to add depos
function returnPreviousOrderers() {
  const stringCell = SpreadsheetApp.getActive().getSheetByName('Infrastructure').getRange(16, 2).getValue().toString()
  return JSON.parse(stringCell)
}


////////////////////////////////////////////////////////////////////////////////////
////////////////// RECORDING ANY APPLICATION ERRORS FOR DEVELOPER //////////////////
////////////////////////////////////////////////////////////////////////////////////

/** Send any errors to a developer log which is stored in this Sheet's Properties
* @param {message} string The error message generated by the Script.
*/
function addToDevLog(message) {
  let date = new Date().toString()
  const devSheet = SpreadsheetApp.getActive().getSheetByName('Developer')
  devSheet.insertRowBefore(2)
  devSheet.getRange(2, 1).setValue(date)
  devSheet.getRange(2, 2).setValue(message)
};


////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////// UTILITIES /////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////

/** Parses Script Property value into orderer name and firm name.
@param {key} string Key to a Script Property.
@return {ordererInfo} string Order's full name, followed by their firm in parenthesis.
Note--a value looks like this: #O#Davis Jones#F#Jones Law#C#1
*/
function parseOrdererInfo(key) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  var ordererName = value.match(/#O#.*(?=#F)/)[0].slice(3);
  var firmName = value.match(/#F#.*(?=#C)/)[0].slice(3);
  var ordererInfo = ordererName + ' (' + firmName + ')';
  return ordererInfo;
};

/** Parses Script Property value to find and return order count.
@param {key} string Key to a Script Property.
@return {count} number The number of times this orderer has ordered a deposition during this reporting period.
Note--a value looks like this: #O#Davis Jones#F#Jones Law#C#1
*/
function parseOrdererCount(key) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  var count = parseInt(value.match(/#C#.*/)[0].slice(3));
  return count;
};

/** Provides visibility into Script Properties keys. */
function seeScriptPropsKeys() {
  var keys = PropertiesService.getScriptProperties().getKeys();
  Logger.log(keys);
};

/** Provides visibility into Script Properties values. */
function seeScriptPropsValues() {
  var keys = PropertiesService.getScriptProperties().getKeys();
  keys.forEach(function(key) {
    var value = PropertiesService.getScriptProperties().getProperty(key);
    Logger.log(value);
  });
};

/** Wipes Script Properties. */
function deleteScriptProps() {
  var keys = PropertiesService.getScriptProperties().deleteAllProperties();
};


