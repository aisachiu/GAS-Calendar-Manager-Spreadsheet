// Calendar Manager - A Google Spreadsheet to help manage your calendars, made especially for AIS
//
// By: Andrew Chiu (achiu@ais.edu.hk; twitter: @chew_ed)
// Credits:
// Google Loading gif using CSS3 taken from Nicolas Saubi's CodePen: https://codepen.io/gonnarule/pen/avmBao
//


var sleepTime = 2300; //sleep for how long each time?
var sleepInterval = 18; //sleep every how many calendar entries?

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Calendar Manager')
      .addItem('Import Calendar into new sheet', 'showImportDialog')
      .addItem('Delete Marked Events on Calendar', 'showDeleteMarkedDialog')
      .addSeparator()
      .addItem('Create Calendar Template Sheet', 'makeTemplateForExport')
      .addItem('Make Calendar events from this sheet', 'showExportDialog')
      .addToUi();
}

function showImportDialog() {

  var htmlt = HtmlService.createTemplateFromFile('picker');
  htmlt.myCals = getCalendars();
  Logger.log(htmlt.myCals);
  var html = htmlt.evaluate();
  SpreadsheetApp.getUi() 
      .showModalDialog(html,"choose calendar and dates");
 // Logger.log("HTML return = %s", html.getContent());     // What does html contain?
}

function showExportDialog() {

  var htmlt = HtmlService.createTemplateFromFile('Exportpicker');
  htmlt.myCals = getCalendars();
  Logger.log(htmlt.myCals);
  var html = htmlt.evaluate();
  SpreadsheetApp.getUi() 
      .showModalDialog(html,"choose calendar");
 // Logger.log("HTML return = %s", html.getContent());     // What does html contain?
}

function showDeleteMarkedDialog() {

  var htmlt = HtmlService.createTemplateFromFile('deleteMarked');
  var html = htmlt.evaluate();
  SpreadsheetApp.getUi() 
      .showModalDialog(html,"choose calendar and dates");
 // Logger.log("HTML return = %s", html.getContent());     // What does html contain?
}


function makeTemplateForExport(){
  var mySpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  mySpreadsheet.getSheetByName("CreateCalendarEventsTemplate").copyTo(mySpreadsheet);
}


function getCalendars() {
  //Get user's calendars
  var myCalendars = [];
  var myCals = CalendarApp.getAllCalendars();
  var myDefCal = CalendarApp.getDefaultCalendar().getId();
  for(var c = 0; c < myCals.length; c++){
   myCalendars.push({name: myCals[c].getName(), id: myCals[c].getId()});
  }
  myCalendars.sort(function(a,b){
                    if (a.name < b.name)
                      return -1;
                    if (a.name > b.name)
                      return 1;
                    return 0;});
  return myCalendars;
}


function putCalEventsOnSS(myCalID, startDateTxt, endDateTxt) {
  
  //Logger.log(JSON.stringify(arguments));
  // To send error messages, throw an exception.
  // e.g. if (invalid) throw new error("Invalid date")
  
  
  //Open Calendar
  var myCal = CalendarApp.getCalendarById(myCalID);
  
  //Get all events within date range
  var myEvts = myCal.getEvents(new Date(startDateTxt), new Date(endDateTxt));
  
  var EventsArray = new Array();
  EventsArray.push(["Title", "Creator", "Start", "End", "Event ID", "Calendar ID"]);
  
  for (var i=0; i<myEvts.length;i++){
    //var myLink = "https://www.googleapis.com/calendar/v3/calendars/"+ myCalID1 + "/events/" + myEvts[i].getId();
    EventsArray.push([  myEvts[i].getTitle(), myEvts[i].getCreators().toString(),myEvts[i].getStartTime(),myEvts[i].getEndTime(),myEvts[i].getId(), myCalID]);
  }
  //write to Spreadsheet
  var mySheet = SpreadsheetApp.getActive().insertSheet();
  var writeData = mySheet.getRange(1, 1, EventsArray.length, EventsArray[0].length).setValues(EventsArray);
  
  
  Logger.log(EventsArray);
}

function lookForDuplicates(){
  var mySs = SpreadsheetApp.getActiveSheet();
  var myData =mySs.getDataRange().getValues();
  
  var tempHeader = myData[0];
  tempHeader.push("# of duplicates", "Duplicate", "Duplicate Of");
  var firstFound = new Array();
  firstFound.push(myData[0]);
  
  for (var i = 1; i < myData.length; i++){ //loop each value in sheet
    var found=false;
    var j = 1;
    var duplicateOf = ""
    while (!found && j < firstFound.length){
      if((firstFound[j][0]==myData[i][0]) && (firstFound[j][2]==myData[i][2])&&(firstFound[j][3]==myData[i][3])){
        found = true;
        duplicateOf = firstFound[j][4];
        firstFound[j][8]++;
      }
      j++;
    }  
    if (!found) {firstFound.push(myData[i]); }
  }
    
    //write to Spreadsheet
  var mySheet = SpreadsheetApp.getActive().insertSheet("Duplicates from " & mySs.getSheetName());
  var writeData = mySheet.getRange(1, 1, firstFound.length, firstFound[0].length).setValues(firstFound);

}


//***
//function deleteEventsOnSheet()
//
//Script to delete all events on the calendar from a spreadsheet. Assumes event ID in column E / 4, calendar ID in Col F / 5
//***
function deleteEventsOnSheet(){
  var mySs = SpreadsheetApp.getActiveSheet();
  var myData =mySs.getDataRange().getValues();
  
  for (var i = 1; i < myData.length; i++){ //loop each value in sheet
    try{
      CalendarApp.getCalendarById(myData[i][5]).getEventSeriesById(myData[i][4]).deleteEventSeries();
    } catch(e) {
      Logger.log([i,e]);
    }
  }
}

//***
//function deleteMarkedEvents()
//
//Script to delete events on the calendar from a spreadsheet. Assumes event ID in column E / 4, calendar ID in Col F / 5, Mark for deletion on Col G / 6 indicated for deletion by value "1"
//***
function deleteMarkedEvents(){
  var mySs = SpreadsheetApp.getActiveSheet();
  var myData =mySs.getDataRange().getValues();
  var myActions = [[myData[0][0],myData[0][1],myData[0][2],myData[0][3],myData[0][4],"Action"]];
  for (var i = 1; i < myData.length; i++){ //loop each value in sheet
    try{
      Logger.log([i, myData[i][6]]);
      if(myData[i][6]===1){
        CalendarApp.getCalendarById(myData[i][5]).getEventSeriesById(myData[i][4]).deleteEventSeries();
        myActions.push([myData[i][0], myData[i][1], myData[i][2], myData[i][3], myData[i][4], "Deleted"]);
      }
    } catch(e) {
      Logger.log([i,e]);
    }
    if(i%sleepInterval ==0) { Utilities.sleep(sleepTime); } //pause the script momentarily to avoid GCalendar overload error
  }
  var mySheet = SpreadsheetApp.getActive().insertSheet("Deleted " & new Date());
  var writeData = mySheet.getRange(1, 1, myActions.length, myActions[0].length).setValues(myActions);

}


//***
//function putSSEventsOnCalendar
//
//Script to write spreadsheet events onto calendar
//***
function putSSEventsOnCalendar(calID){
  var mySs = SpreadsheetApp.getActiveSheet();
  var myData =mySs.getDataRange().getValues();
  var myCal = CalendarApp.getCalendarById(calID);
  
  
  for (var i = 1; i < myData.length; i++){ //loop each value in sheet
    var action = 0;
    var myEventID = "";
    if(myData[i][1] == "") {continue;} //if title is blank, skip the row.
    if(myData[i][3] == "") { action = 1}; //if endDate is blank, assume it is an all-day event
    try{
      if(action = 0){
        myEventID = myCal.createEvent(myData[i][1], new Date(myData[i][2]), new Date(myData[i][3]),
                                              {description: myData[i][4], location: myData[i][5], guests: myData[i][6]}).getEventSeries().getId(); //save eventID to write to sheet
  //Create normal event
      }
      if(action = 1){ //create all day event
        myEventID = myCal.createAllDayEvent(myData[i][1], new Date(myData[i][2]),
                                              {description: myData[i][4], location: myData[i][5], guests: myData[i][6]}).getEventSeries().getId(); //save eventID to write to sheet
      }
    } catch(e) {
      myEventID = e;
    }
    mySs.getRange(i+1, 1).setValue(myEventID); //write eventID to sheet
    if(i%sleepInterval ==0) { Utilities.sleep(sleepTime); } //pause the script momentarily to avoid GCalendar overload error.
  }
}



// -----
// include - include files in HTML
// -----
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}
