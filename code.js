function getTest() {
  var test = false;
  return test;
}

function getCalendar() {
  if (getTest()) return "TestCalendars";
  else return "actualCalendars";
}

function getDataSheet() {
  if (getTest()) return "TestSemester";
  else return "CurrentSemester";
}

////////////////////////////////////////////////////////////////////////
// menu bar creation
////////////////////////////////////////////////////////////////////////

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  Logger.log('open working')
  formatSheet();

  // Or DocumentApp or FormApp.
  ui.createMenu('Scheduler')
    .addSubMenu(ui.createMenu('Update')
    .addItem('Update All', 'updateAll')
      .addSubMenu(ui.createMenu('Select Rooms')
        .addItem('ROOM 109 HIGH BAY', 'updateOne')
        .addItem('ROOM 111 SCREENING ROOM', 'updateTwo')
        .addItem('ROOM 126 ACTING', 'updateThree')
        .addItem('ROOM 129 Writing/Producing', 'updateFour')
        .addItem('ROOM 133 PRODUCTION 1 & SOUND', 'updateFive')
        .addItem('ROOM 134 POST 1 LAB', 'updateSix')
        .addItem('ROOM 135 WRITING PRODUCING LAB', 'updateSeven')
        .addItem('ROOM 136 ADVANCED LAB', 'updateEight')
        .addItem('ROOM 142 CONFERENCE ROOM', 'updateNine')
        .addItem('ROOM 147 SOUND LAB', 'updateTen')
        .addItem('ROOM 151 FISH BOWL', 'updateEleven')
        .addItem('ROOM 121 STUDIO', 'updateTwelve')
        .addItem('ROOM 50s CAFE', 'updateThirteen')
        .addItem('ROOM BUNGALOW', 'updateFourteen')
        .addItem('Canceled Classes', 'updateFifteen'))
      .addItem('Update Selected Rows', 'updateActive'))
    .addSubMenu(ui.createMenu('Clear Calendars')
      .addItem('Clear All Calendars', 'clearCalendars')
      .addSubMenu(ui.createMenu('Clear Select Rooms')
      .addItem('ROOM 109 HIGH BAY', 'clearOne')
      .addItem('ROOM 111 SCREENING ROOM', 'clearTwo')
      .addItem('ROOM 126 ACTING', 'clearThree')
      .addItem('ROOM 129 Writing/Producing', 'clearFour')
      .addItem('ROOM 133 PRODUCTION 1 & SOUND', 'clearFive')
      .addItem('ROOM 134 POST 1 LAB', 'clearSix')
      .addItem('ROOM 135 WRITING PRODUCING LAB', 'clearSeven')
      .addItem('ROOM 136 ADVANCED LAB', 'clearEight')
      .addItem('ROOM 142 CONFERENCE ROOM', 'clearNine')
      .addItem('ROOM 147 SOUND LAB', 'clearTen')
      .addItem('ROOM 151 FISH BOWL', 'clearEleven')
      .addItem('ROOM 121 STUDIO', 'clearTwelve')
      .addItem('ROOM 50s CAFE', 'clearThirteen')
      .addItem('ROOM BUNGALOW', 'clearFourteen')
      .addItem('Canceled Classes', 'clearFifteen')))
    .addSubMenu(ui.createMenu('Validate')
      // .addItem('Reset All', 'resetAll')
      // .addSeparator()
      .addItem('Validate Calendars', 'validateCalendars'))
    .addToUi();
};

////////////////////////////////////////////////////////////////////////
// formats active sheet, applies frozen row to top so column subjects can be seen when scrolling.
////////////////////////////////////////////////////////////////////////
function formatSheet(){

  var mySS = SpreadsheetApp.getActiveSpreadsheet();
  var myS = SpreadsheetApp.getActiveSheet();
  myS.setFrozenRows(1);
  var colToHide = myS.getRange(1, 16, 1, 5)
  myS.hideColumn(colToHide)
};

////////////////////////////////////////////////////////////////////////
// outputs single row to assigned calendar
////////////////////////////////////////////////////////////////////////

function outputCourse (x){
  var mySS = SpreadsheetApp.getActiveSpreadsheet();
  var myS = mySS.getActiveSheet();
  var course = myS.getRange(x, 1, 1, 24)
  var calendarCol = myS.getRange(x, 21);
  var eventIdCol = myS.getRange(x, 22);
  //return console.log(course.getValues())
  var range = course.getValues();
  var destCalendar = range[0][20];

  // creates event start date
  var origStartDate = range[0][6];
  var startDay = origStartDate.getDate();
  var startMonth = origStartDate.getMonth();
  var startYear = origStartDate.getFullYear();

  // creates event start time
  var startTime = range[0][8];

  //var startHour = startTime.getHours();
  //var startMin = startTime.getMinutes();

  var startHour = startTime.substring(0, 2);
  startHour = startHour.replace(/O/g,"0");
  var startMin = startTime.substring(2, 4);
  startMin = startMin.replace(/O/g,"0");
  var suffix = startTime.substring(4,6);

  if (suffix == "PM" && startHour != "12") {
  var t1 = parseInt(startHour,10);
  t1+= 12;
  startHour = t1.toString();
  }

  var newStart = new Date(startYear, startMonth, startDay, startHour, startMin, 0, 0);
  Logger.log(newStart);
  Logger.log('offset: ' + newStart.getTimezoneOffset());

  // creates event end time
  var endTime = range[0][9];
  //var endHour = (endTime.getHours());
  //var endMin = endTime.getMinutes();

  var endHour = endTime.substring(0, 2);
  endHour = endHour.replace(/O/g,"0");
  var endMin = endTime.substring(2, 4);
  endMin = endMin.replace(/O/g,"0");
  var suffix = endTime.substring(4,6);

  if (suffix == "PM" && endHour != "12") {
  var t1 = parseInt(endHour,10);
  t1+= 12;
  endHour = t1.toString();
  }

  var newEnd = new Date(startYear, startMonth, startDay, endHour, endMin, 0, 0);

  // creates date reoccurance end
  var origFinalDate = range[0][7]
  var roomNumber = range[0][12];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(getCalendar());
  var cal = getRoomCalendar(sheet, roomNumber)

  Logger.log(cal);

  // returns calendar name
  // var courseName = function () {
  //   var roomNum =  range[0][12];
  //   var calSheet = getSheetByName(actualCalendars);
  //   var calSheetRoomNum = calSheet.getRange(1,1,17)
  // }

  // Logger.log(courseName());

  // gathers weekday reoccurance
  var wkdays = "";
      switch (range[0][10])
      {
        case "MW": wkdays = [CalendarApp.Weekday.MONDAY,CalendarApp.Weekday.WEDNESDAY]; break;
        case "TR": wkdays = [CalendarApp.Weekday.TUESDAY,CalendarApp.Weekday.THURSDAY]; break;
        case "M": wkdays = [CalendarApp.Weekday.MONDAY]; break;
        case "T": wkdays = [CalendarApp.Weekday.TUESDAY]; break;
        case "W": wkdays = [CalendarApp.Weekday.WEDNESDAY]; break;
        case "R": wkdays = [CalendarApp.Weekday.THURSDAY]; break;
        case "F": wkdays = [CalendarApp.Weekday.FRIDAY]; break;
      }

  var eventTitle = range[0][0] + ' ' + range[0][1] + ' ' + range[0][2] + ' ' + range[0][3] + ' -' + range[0][13] + '-';

  if (cal) {

    // outputs calendar event
    var eventSeries = CalendarApp.getCalendarsByName(cal.getName())[0].createEventSeries(
      eventTitle,
      newStart,
      newEnd,
      CalendarApp.newRecurrence().addWeeklyRule()
          .onlyOnWeekdays(wkdays)
          .until(origFinalDate),
      {location: 'Room ' + range[0][12]}
    );

    calendarCol.setValue(cal.getName());
    eventIdCol.setValue(eventSeries.getId());

    Logger.log(eventSeries.getId());
  }
};

////////////////////////////////////////////////////////////////////////
// UPDATE ALL FUNCTION
////////////////////////////////////////////////////////////////////////

function updateAll() {
  clearCalendars();
  Utilities.sleep(5 * 1000)

  var mySS = SpreadsheetApp.getActiveSpreadsheet();
  var myS = SpreadsheetApp.getActiveSheet();
  var arr = myS.getRange(1, 1,150).getValues();
  var newArr = arr.join().split(',').filter(Boolean);

  var y;
  for (y = 0; y < newArr.length; y++){
    var checkArr = newArr[y].length;
    if (checkArr === 3) {
      outputCourse(y + 1);
    } else {
      console.log(y + ' is not a course')
    }
  };

  //console.log(newArr.length)
  //var result = arr.filter(i => i.toString().length === 3);
  }

////////////////////////////////////////////////////////////////////////s
// UPDATE ACTIVE ROWS
////////////////////////////////////////////////////////////////////////

function updateActive() {
  var mySS = SpreadsheetApp.getActiveSpreadsheet();
  var myS = mySS.getActiveSheet();
  var rangeList = myS.getActiveRangeList().getRanges();
  

  for (var i = 0; i<rangeList.length; i++){
    var range = rangeList[i].getRowIndex();
    outputCourse(range);
  };

  //Logger.log(myS.getActiveRangeList().getRanges())

};

////////////////////////////////////////////////////////////////////////
// UPDATE SELECT ROOMS
////////////////////////////////////////////////////////////////////////
function updateRoom(roomNum) {
  clearSelectRooms(roomNum);
  var mySS = SpreadsheetApp.getActiveSpreadsheet();
  var myS = mySS.getActiveSheet();
  var roomArr = myS.getRange(1,13,150).getValues();
  var newRoomArr = roomArr.join().split(',').filter(Boolean);

  var v;
  for (v = 0; v < newRoomArr.length; v++) {
    var checkRoomArr = newRoomArr[v]
    if (checkRoomArr == roomNum) {
     outputCourse(v+1);
    } else {
     Logger.log(checkRoomArr + 'nope')
    }
  };


}

function updateOne (){ return updateRoom(109) };
function updateTwo (){ return updateRoom(111) };
function updateThree (){ return updateRoom(126) };
function updateFour (){ return updateRoom(129) };
function updateFive (){ return updateRoom(133) };
function updateSix (){ return updateRoom(134) };
function updateSeven (){ return updateRoom(135) };
function updateEight (){ return updateRoom(136) };
function updateNine (){ return updateRoom(142) };
function updateTen (){ return updateRoom(147) };
function updateEleven (){ return updateRoom(151) };
function updateTwelve (){ return updateRoom(121) };
function updateThirteen (){ return updateRoom('CAN') };
function updateFourteen (){ return updateRoom("50's") };
function updateFifteen (){ return updateRoom('Bungalow') };

////////////////////////////////////////////////////////////////////////
// Clear SELECT ROOMS
////////////////////////////////////////////////////////////////////////

function clearSelectRooms(roomNum) {
  var mySS = SpreadsheetApp.getActiveSpreadsheet();
  var myS = mySS.getActiveSheet();
  var roomArr = myS.getRange(1,13,150).getValues();
  var newRoomArr = roomArr.join().split(',').filter(Boolean);

  var v;
  for (v = 0; v < newRoomArr.length; v++) {
    var checkRoomArr = newRoomArr[v]
    Logger.log('checkRoomArr: ' + checkRoomArr);

    if (checkRoomArr == roomNum) {
     var getId = myS.getRange((v+1),22).getValues().join().split(',').filter(Boolean)
     var getCalName = myS.getRange((v+1),21).getValues().join().split(',').filter(Boolean)
     if (getId == '' || getCalName == ''){
       Logger.log('NO CALENDAR EVENT')
     } else {
     var getCal = CalendarApp.getCalendarsByName(getCalName)[0].getEventSeriesById(getId);
     getCal.deleteEventSeries();
     myS.getRange((v+1),22).clearContent();
     myS.getRange((v+1),21).clearContent();
     Logger.log(getCal);
    }
    } else {
     Logger.log(checkRoomArr + 'nope')
    }
  };
};

  function clearOne (){ return clearSelectRooms(109) };
  function clearTwo (){ return clearSelectRooms(111) };
  function clearThree (){ return clearSelectRooms(126) };
  function clearFour (){ return clearSelectRooms(129) };
  function clearFive (){ return clearSelectRooms(133) };
  function clearSix (){ return clearSelectRooms(134) };
  function clearSeven (){ return clearSelectRooms(135) };
  function clearEight (){ return clearSelectRooms(136) };
  function clearNine (){ return clearSelectRooms(142) };
  function clearTen (){ return clearSelectRooms(147) };
  function clearEleven (){ return clearSelectRooms(151) };
  function clearTwelve (){ return clearSelectRooms(121) };
  function clearThirteen (){ return clearSelectRooms('CAN') };
  function clearFourteen (){ return clearSelectRooms("50's") };
  function clearFifteen (){ return clearSelectRooms('Bungalow') };
/*
 * LEGACY CODE
 */


// ***********************************************************
// Clear Calendars
// ***********************************************************
function clearCalendars(spreadsheetName, calendarSheetName, roomID) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var classesSheet = SpreadsheetApp.getActiveSheet();
  var calendarColumn = getDataColumn("Calendar");
  var eventIDColumn = getDataColumn("EventID");
  var titleColumn = getDataColumn("CourseName");
  var roomColumn = getDataColumn("Room");
  var resultColumn = getDataColumn("Result");
  var rows = classesSheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  var startRow = 1;  // First row of data to process (row 0 is the headers)
  var result = "";
  var resultList = [];
  var results = 0;

  for (var r=startRow; r < numRows - startRow + 1; r++) {
    var row = values[r];
    var rowRoom = row[roomColumn];
    var calendarName = row[calendarColumn];
    var eventId = row[eventIDColumn];
    var classTitle = row[titleColumn];
    var deleteDetail = "";

      Logger.log("About to process row "+r+" = rowRoom "+ rowRoom);

      if (rowRoom) {
        Logger.log ( "This row's room ( "+ rowRoom +" ) is the room we are clearing ( "+ roomID +" )");
        Logger.log("This room: rowRoom = " + rowRoom + " calendarName = " + calendarName + " eventId = " + eventId + " classTitle = " + classTitle);

        if (( calendarName == "" ) || ( eventId == "" )) { // no event series to delete
          Logger.log("row "+r+", room "+rowRoom+" skipped, no existing event ("+calendarName+'/'+eventId+")");
        } else {
          // deleteEvents, check for fail, clear calendarID and eventSeries
          try {
              Logger.log( "Try: calendarName = "+calendarName+" eventId = "+eventId);
              deleteEventFromNamedCalendar(eventId, calendarName);
              Logger.log("SUCESS: row "+ r +", room "+ rowRoom +" ("+ calendarName +'/'+ eventId +"):" + result);
            deleteDetail = "deleted";
          // log failure and continue
          } catch (err) {
            Logger.log ("*************** CATCH error ****************");
            Logger.log("FAIL: row "+ r +", room "+ rowRoom +" ("+ calendarName +'/'+ eventId +"):" + err);
            Logger.log ("*************** CATCH error ****************");
            Logger.log (" ERROR deleting "+ classTitle +" from "+calendarName+" - check calendar and delete by hand if necessary.");

            deleteDetail = "delete failed";
          }
            // clear calendar/event info from line - regardless of delete success/fail - it only seems to fail when the eventID is bad
            // so just issue a warning
            // spreadsheet ranges to write data are 1-n, not zero based, so modify row and column indices
            var destEventID = classesSheet.getRange(r+1,eventIDColumn+1);
            var destCalendar = classesSheet.getRange(r+1,calendarColumn+1);
            var destResult = classesSheet.getRange(r+1, resultColumn+1)

            destEventID.clear();
            destCalendar.clear();
            destResult.setValue("Room" + rowRoom + " " + deleteDetail);
          } // end try/catch delete events
        } // end is this the room
  } // end check every row in the data table

  result = " Finished clearing " + roomID;
  return result;
}

// ***********************************************************
// deleteEvents functions
// ***********************************************************
function deleteEventFromNamedCalendar(eventId, calName) {
  Logger.log("DELETING EVENT: (" + eventId + ") from calendar (" + calName + ")");

  var calendar = CalendarApp.getOwnedCalendarsByName(calName)[0];

  if (!calendar) {
    Logger.log("FAILED to access " + calName);
    return false;
  }

  var eventSeries = calendar.getEventSeriesById(eventId)

  console.log(eventSeries);

  return eventSeries.deleteEventSeries();

  // return deleteEventFromOpenCalendar(eventId, calendar);
  // TBD do I need to close the calendar somehow? I think not but...
}

function deleteEventFromOpenCalendar(eventId, calendar) {
  if (!calendar) {
    Logger.log ("defoc: Invalid calendar passed to deleteEventFromOpenCalendar");
    return false;
  }

  // TBD this in a try block? or check/report error differently?
  var eventSeries = calendar.getEventSeriesById(eventId);

  Logger.log('eventSeries: ' + eventSeries);

  if (!eventSeries) {
    Logger.log("defoc: Event " + eventId + " not found in " + calendar.getName());
    return false;
  }

  // TBD test this is called from within a try/catch - so nested; make sure no problem here
  try {
    eventSeries.deleteEventSeries();
  } catch (err) {
    Logger.log("defoc: **** FAIL CATCH **** Error " + err + " deleting event " + eventId + " from " + calendar.getName());
    return false;
  }

  Logger.log("defoc: Event " + eventId + " removed from " + calendar.getName());

  return true;
}

// ***********************************************************
// getDataColumn - get data column index for particular data
// ***********************************************************

function getDataColumn (columnName) {
  // TBD modify scheduleRooms function to use this
  var name = [ "Subject","Course","Section","CourseName","unused4","unused5","StartDate","StopDate","StartTime","StopTime","Days","unused11",
             "Room","InstructorLast","InstructorFirst","unused15","unused16","unused17","unused18","unused19","Calendar","EventID","Action",
             "Result","RoomNotes" ];
  // loop, check name= columnName, return index
  for (var i = 0; i < name.length; i++) {
    if (name[i] == columnName) return i;
  }

  return -1; // not found
}





function getRoomList() {
  Logger.log('Running get room list')
  var pageName = getDataSheet();
  Logger.log('Page Name:' + pageName)
  var columnNumber = getDataColumn("Room");
  Logger.log('Column Number:' + columnNumber)

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var list  = ss.getSheetByName(pageName);
  var values = list.getDataRange().getValues();
  var allValues = [];

  // get all room numbers into an array
  var length = values.length;
  for (var i = 0; i < length-1; i++) {
    allValues[i] =  values[i+1][columnNumber];
    //Logger.log("Row " + i+1 + " Contains value " + allValues[i]);
  }
  //Logger.log("sheet: " + pageName + " column: " + columnNumber);
  //Logger.log(allValues);

  // transfer only the unique values to another array
  var uniqueValues = [];
  var exists, x, y;

  for(x=0;x < allValues.length;x++){
    exists = false;
    for(y=0;y<uniqueValues.length;y++){
      if(allValues[x].toString() === uniqueValues[y].toString()){
        exists=true;
        break;
      }
    }
    //if (allValues[x].toString() == '') Logger.log("Blank line in room list will be omitted");
    if(!exists && allValues[x].toString() != '') uniqueValues.push(allValues[x]);
  }

  // sort unique values and return just those
  Logger.log("Before sort: "+uniqueValues);
  uniqueValues.sort();
  // eliminating blank lines above so no need to shift (sally 14.Jan.2014)
  //Logger.log("Before shift: "+uniqueValues);
  //uniqueValues.shift(); // remove the top value (always blank line)
      //Logger.log("unique values shifted");
  Logger.log("Returning: "+uniqueValues);

  Logger.log("JSON Stringify");

  console.log(uniqueValues);

  return uniqueValues;
};

// ***********************************************************
// getRoomCalendar function
// ***********************************************************
function getRoomCalendar(calendarSheet, roomNumber) {
  Logger.log("getRoomCalendar: find calendar for room " + roomNumber);
  Logger.log(calendarSheet);

  // get the room/calendar mapping
  var roomRows = calendarSheet.getDataRange();
  var roomData = roomRows.getValues();
  var startRoomRow = 2;
  var numRoomRows = roomRows.getNumRows();

  // get the room number and calendar name (we assume data is validated, if its not and the calendar is
  //   there, great, otherwise we return an invalid calendar and calling code checks that
  var calendarFound = false;

  for (var i = 1; i < numRoomRows; i++) {
    var calendarRoom = roomData[i][0];
    var calendarName = roomData[i][1];

    if (calendarRoom == roomNumber) {
      calendarFound = true;
      Logger.log("                 " + calendarName + " for " + calendarRoom + " found");
      break;
    }
  }

  if (calendarFound) {
    return CalendarApp.getOwnedCalendarsByName(calendarName)[0];
  } else {
    Logger.log("                 "+ roomNumber + " not found in calendar sheet");
    return null;
  }
}

// ***********************************************************
// ValidateCalendars function
// ***********************************************************
function validateCalendars()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(getCalendar());
  var rows = sheet.getDataRange();
  var data = rows.getValues();

  var startRow = 2;
  var numRows = rows.getNumRows()-startRow;
  var result = "Room calendars in "+getCalendar()+" are valid and accessible.";
  var errorCount = 0;

  //var dataSheet = getDataSheet();
  // var roomColumn = getDataColumn("Room");
  var roomList = getRoomList();
  Logger.log("validate calendars: "+roomList.length +" rooms in data");
  var roomListHasCalendar = new Array();
  for (var r=0; r<roomList.length; r++)
  {
    roomListHasCalendar[r] = false;
  }

  //TODO modify following code to check the rooms against the roomlist and report missing ones


  // for each row (room/calendar pair) in RoomCalendars, check whether we can access the calendar
  for (var i=startRow; i < numRows + startRow + 1; i++)
  {
    // get the room number and calendar name
    var room = data[i-1][0];
    var calendarName = data[i-1][1];
    Logger.log('room: ' +  room);
    Logger.log('calendarName: ' +  calendarName);
    // test to make sure its a valid calendar
    var cal = CalendarApp.getOwnedCalendarsByName(calendarName)[0];

    // if its valid update RoomCalendars validation column
    var dest = sheet.getRange(i,3);
    if (cal) {
      dest.setValue(room+" OK");
      Logger.log("Calendar for room "+room+" ("+calendarName+") validated.");
      for (var n=0; n<roomList.length; n++)
      {
        if (roomList[n] == room)
        {
          roomListHasCalendar[n] = true;
        }
      }
    } else {
      dest.setValue(room + ' FAILED');
      Logger.log("Calendar for room "+room+" ("+calendarName+") FAILED validation.");
      if (errorCount == 0) {
        result = "Error accessing calendar in "+getCalendar()+", FAILED: ";
        errorCount++;
      }
      else {
        result += ", ";
      }
      result += "room "+room+" ("+calendarName+")";
    }
  }
  // TBD how do I actually make new lines in my label output?
  var first=true;
  for (var n=0; n<roomList.length; n++)
  {
    if (!roomListHasCalendar[n]) {
      if (first) { first=false; result += " ERROR: room(s) ";}
      result += " " + roomList[n];
    }
  }
  if (!first) result += " not found or not accessible.";
  Logger.log(result);
  return result;
}

// ***********************************************************
// getRoomList - returns array of unique items in column of spreadsheet page
// ***********************************************************
function getRoomList()
{
  Logger.log('Running get room list')
  var pageName = getDataSheet();
  Logger.log('Page Name:' + pageName)
  var columnNumber = 1
  Logger.log('Column Number:' + columnNumber)

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var list  = ss.getSheetByName(pageName);
  var values = list.getDataRange().getValues();
  var allValues = [];

  // get all room numbers into an array
  var length = values.length;
  for (var i = 0; i < length-1; i++) {
    allValues[i] =  values[i+1][columnNumber];
    //Logger.log("Row " + i+1 + " Contains value " + allValues[i]);
  }
  //Logger.log("sheet: " + pageName + " column: " + columnNumber);
  //Logger.log(allValues);

  // transfer only the unique values to another array
  var uniqueValues = [];
  var exists, x, y;

  for(x=0;x < allValues.length;x++){
    exists = false;
    for(y=0;y<uniqueValues.length;y++){
      if(allValues[x].toString() === uniqueValues[y].toString()){
        exists=true;
        break;
      }
    }
    //if (allValues[x].toString() == '') Logger.log("Blank line in room list will be omitted");
    if(!exists && allValues[x].toString() != '') uniqueValues.push(allValues[x]);
  }

  // sort unique values and return just those
  Logger.log("Before sort: "+uniqueValues);
  uniqueValues.sort();
  // eliminating blank lines above so no need to shift (sally 14.Jan.2014)
  //Logger.log("Before shift: "+uniqueValues);
  //uniqueValues.shift(); // remove the top value (always blank line)
      //Logger.log("unique values shifted");
  Logger.log("Returning: "+uniqueValues);

  Logger.log("JSON Stringify");

  return uniqueValues;
};