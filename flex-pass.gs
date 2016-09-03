  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = sheet.getRange(2,1,sheet.getLastRow()-1, sheet.getLastColumn());
  var values = range.getValues();

  
  var myData = getRowsData(sheet);


// This function will get the calendar invite data by ID (which was printed to spreadsheet)
// then parse through the data to extract guest names and response status.
// Ultimately, we run a conditional where the calendar event will be deleted if a response is no.
// New question: how will this function be triggered?
// I'd like for it to be triggered every time a response is entered, but it might have to be time driven?
  // Alert all parties of the deletion event, indicating who stopped the event.

function calEventResponses() {
  var guestStatusArray = [];
  
  for (var i = 0, count = myData.length; i < count; i++){
    var calId = sheet.getRange((i+2), 9, 1).getValue();
    var calEvent = CalendarApp.getEventSeriesById(calId);
    var eventGuests = calEvent.getGuestList();
    
    var guestStatusByEvent = [];
    for (var j = 0, guestCount = eventGuests.length; j<guestCount; j++){
      // Logger.log(eventGuests[j].getName());
      // Logger.log(eventGuests[j].getGuestStatus());
      var guestStatus = eventGuests[j].getGuestStatus();
        // Logger.log(eventGuests[j].getName());
        // Logger.log(eventGuests[j].getGuestStatus());
        guestStatusByEvent.push(guestStatus);
      
      if (guestStatus == "NO"){
        Logger.log(eventGuests[j].getName());
        Logger.log(eventGuests[j].getGuestStatus());
        calEvent.deleteEventSeries();
      }
      else{
        guestStatusByEvent.push(guestStatus); 
      }
 
    } // End of inner for-loop
    
    guestStatusArray.push(guestStatusByEvent);

  } // end of outer for-loop
}



function onFlexPassSubmit() {
  
  // Substitute teacher name with teacher email address
  for (var i = 0, count = myData.length; i < count; i++){
    switch(myData[i].releasingTeacher){
      case "Auble": 
        myData[i].releasingTeacherEmail = 'techteam@mustangsla.org';
        break;
      case "Lanus": 
        myData[i].releasingTeacherEmail = 'cqlanus@gmail.com';
        break;
      case "Salcedo": 
        myData[i].releasingTeacherEmail = 'dla@mustangsla.org';
        break;
      default: 
        Logger.log('Try again');
    }
    
    switch(myData[i].receivingTeacher){
      case "Auble": 
        myData[i].receivingTeacherEmail = 'techteam@mustangsla.org';
        break;
      case "Lanus": 
        myData[i].receivingTeacherEmail = 'cqlanus@gmail.com';
        break;
      case "Salcedo": 
        myData[i].receivingTeacherEmail = 'dla@mustangsla.org';
        break;
      default: 
        Logger.log('Try again');
  }
  
    if (myData[i].homeworkTeacher){
      switch(myData[i].homeworkTeacher){
      case "Auble": 
        myData[i].homeworkTeacherEmail = 'techteam@mustangsla.org';
        break;
      case "Lanus": 
        myData[i].homeworkTeacherEmail = 'cqlanus@gmail.com';
        break;
      default: 
        myData[i].homeworkTeacher = 'none';
  }
    }
  }


  // Loop through the array of objects, creating an object full of details in preparation for our calendar invite
  for (var i = 0, count = myData.length; i < count; i++){
    var options = {
        description: "This represents a Flex Pass for " + myData[i].fullName + ". He wishes to leave " + myData[i].releasingTeacher + "'s classroom and study in " + myData[i].receivingTeacher + "'s classroom. Please accept this invitation to give permission.",
        location: "Resource Room",
        guests: myData[i].releasingTeacherEmail + ", " + myData[i].receivingTeacherEmail,
        sendInvites: false
      }
    
    // Write code that alerts Jimmy if the computer lab is needed, and alerts the teacher who assigned the work.
    if (myData[i].receivingTeacher == "Salcedo"){
      options.guests += ", " + myData[i].homeworkTeacherEmail;
      options.description += " This student has also indicated the need to use a computer. "
      options.description += myData[i].homeworkTeacher + ", please accept this invitation to indicate that you approve this request."
    }
 
  
    // Identify the last column and target the next column to add contents to
    var cell = sheet.getRange((i+2), 1, 1, 1);
    var cellColor = cell.getBackground();

    // This block of code sends calendar invites if the Flex Pass form is newly submitted, 
    // then colors a marker cell red to indicate the calendar invite has been sent
    if (cellColor != "#ff0000"){
      var title = "Flex Pass: " + myData[i].fullName;
      var date = myData[i].flexDate;
      options.sendInvites = true;
      var calEvent = CalendarApp.createAllDayEvent(title, date, options)
      cell.setBackground("red");
      
     // This adds the array of EventGuest objects to the myData array of objects
      var x = myData[i].calEvent = calEvent;
      var y = x.getId();
      var newCell = sheet.getRange((i+2), 9, 1)
      newCell.setValue(y);

    } // End if statement      
  } // End outer for-loop
} // End function



  // getRowsData iterates row by row in the input range and returns an array of objects.
  // Each object contains all the data for a given row, indexed by its normalized column name.
  // Arguments:
  //   - sheet: the sheet object that contains the data to be processed
  //   - range: the exact range of cells where the data is stored
  //       This argument is optional and it defaults to all the cells except those in the first row
  //       or all the cells below columnHeadersRowIndex (if defined).
  //   - columnHeadersRowIndex: specifies the row number where the column names are stored.
  //       This argument is optional and it defaults to the row immediately above range;
  // Returns an Array of objects.
  
  //* @param {sheet} sheet with data to be pulled from.
  // * @param {range} range where the data is in the sheet, headers are above
  //* @param {row} 
  
  function getRowsData(sheet, range, columnHeadersRowIndex) {
    if (sheet.getLastRow() < 2){
      return [];
    }
    var headersIndex = columnHeadersRowIndex || (range ? range.getRowIndex() - 1 : 1);
    Logger.log(headersIndex);
    var dataRange = range ||
      sheet.getRange(headersIndex+1, 1, sheet.getLastRow() - headersIndex, sheet.getLastColumn());
    var numColumns = dataRange.getLastColumn() - dataRange.getColumn() + 1;
    var headersRange = sheet.getRange(headersIndex, dataRange.getColumn(), 1, numColumns);
    var headers = headersRange.getValues()[0];
    return getObjects_(dataRange.getValues(), normalizeHeaders(headers));
  }
  
  // For every row of data in data, generates an object that contains the data. Names of
  // object fields are defined in keys.
  // Arguments:
  //   - data: JavaScript 2d array
  //   - keys: Array of Strings that define the property names for the objects to create
  function getObjects_(data, keys) {
    var objects = [];
    var timeZone = Session.getScriptTimeZone();
    
    for (var i = 0; i < data.length; ++i) {
      var object = {};
      var hasData = false;
      for (var j = 0; j < data[i].length; ++j) {
        var cellData = data[i][j];
        if (isCellEmpty_(cellData)) {
          object[keys[j]] = '';
          continue;
        }
        object[keys[j]] = cellData;
        hasData = true;
      }
      if (hasData) {
        objects.push(object);
      }
    }
    return objects;
  }
  
  
  // Returns an Array of normalized Strings.
  // Empty Strings are returned for all Strings that could not be successfully normalized.
  // Arguments:
  //   - headers: Array of Strings to normalize
  function normalizeHeaders(headers) {
    var keys = [];
    for (var i = 0; i < headers.length; ++i) {
      keys.push(normalizeHeader(headers[i]));
    }
    return keys;
  }
  
  // Normalizes a string, by removing all alphanumeric characters and using mixed case
  // to separate words. The output will always start with a lower case letter.
  // This function is designed to produce JavaScript object property names.
  // Arguments:
  //   - header: string to normalize
  // Examples:
  //   "First Name" -> "firstName"
  //   "Market Cap (millions) -> "marketCapMillions
  //   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
  function normalizeHeader(header) {
    var key = "";
    var upperCase = false;
    for (var i = 0; i < header.length; ++i) {
      var letter = header[i];
      if (letter == " " && key.length > 0) {
        upperCase = true;
        continue;
      }
      if (!isAlnum_(letter)) {
        continue;
      }
      if (key.length == 0 && isDigit_(letter)) {
        continue; // first character must be a letter
      }
      if (upperCase) {
        upperCase = false;
        key += letter.toUpperCase();
      } else {
        key += letter.toLowerCase();
      }
    }
    return key;
  }
  
  // Returns true if the cell where cellData was read from is empty.
  // Arguments:
  //   - cellData: string
  function isCellEmpty_(cellData) {
    return typeof(cellData) == "string" && cellData == "";
  }
  
  // Returns true if the character char is alphabetical, false otherwise.
  function isAlnum_(char) {
    return char >= 'A' && char <= 'Z' ||
      char >= 'a' && char <= 'z' ||
        isDigit_(char);
  }
  
  // Returns true if the character char is a digit, false otherwise.
  function isDigit_(char) {
    return char >= '0' && char <= '9';
  }
