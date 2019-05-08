function onOpen() {
  /** 
  * Activate Script menu item must be called first because the 
  * createEventsFromScript function uses Oauth and permissions 
  * dialog will not come up when the function is called directly 
  * from a menu
  **/
  var menu = [{name: 'Activate Script', functionName: 'activateMyScript'},
              {name: 'Add Teacher Schedule to Google Calendar', functionName: 'createEventsFromSheet'},
              {name: 'Remove Teacher Schedules from Google Calendar', functionName: 'removeEventsFromSheet'}
             ];
  SpreadsheetApp.getActive().addMenu('Teacher Schedules', menu);
}

function activateMyScript(){
  Browser.msgBox('Teacher Schedule has been activated!');
}

/**
* creates calendar event for each instance of a class for the teacher
* @param {string} className  Name of class
* @param {string} sectionNumber    Section number of individual class
* @param {string} startDateTime date/time for when class starts in this format (yyyy-MM-dd'T'HH:mm:ssZ)
* @param {string} endDateTime date/time for when class ends in this format (yyyy-MM-dd'T'HH:mm:ssZ)
* @param {string} classLocation  room number for class
* @return {string} eventID   returns the eventID of the created event
**/

function createEvent(className, sectionNumber, startDateTime, endDateTime, classLocation) {
  var calendarId = 'primary';
  var formattedStartDate = Utilities.formatDate(startDateTime, "GMT+5:30", "yyyy-MM-dd'T'HH:mm:ssZ");
  var formattedEndDate = Utilities.formatDate(endDateTime, "GMT+5:30", "yyyy-MM-dd'T'HH:mm:ssZ");
  
  var event = {
    summary: className+sectionNumber,
    location: classLocation,
    start: {
      dateTime: formattedStartDate
    },
    end: {
      dateTime: formattedEndDate
    },
    reminders:{
      useDefault: false
    },
      visibility: "public",
    // Red background. Use Calendar.Colors.get() for the full list.reminders.overrides[]
    colorId: 11
  };
  event = Calendar.Events.insert(event, calendarId);
  var eventID = event.id;
  Logger.log('Event ID: ' + event.id);
  //Logger.log('Event summary: ' + event['start']['dateTime'] );
  return eventID;
}


/**
* createEventsFromSheet
* 
* Get data from sheet that has the events for all staff
* get events that only belong to the user
* 
* 
**/
function createEventsFromSheet(){
  var scriptUserEmail = Session.getActiveUser().getEmail();
  Logger.log(scriptUserEmail);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('TeacherSchedule');
  var data = sheet.getDataRange().getValues();
  for (var i=1; i<data.length; i++){
    var values = data[i];
    if (values[0] == scriptUserEmail){
      var className = values[1];
      var sectionNumber = values[2];
      var startDateTime = values[3];
      var endDateTime = values[4];
      var classLocation = values[5];
      
      var eventID = createEvent(className, sectionNumber, startDateTime, endDateTime, classLocation);
      sheet.getRange("G"+(i+1)).setValue(eventID);
    }
  } 
}

/**
* createEventsFromApi
* 
* Get data from API that has the events for staff
* based on the email address of the user running the script
* 
* 
**/
function createEventsFromApi(){
  var scriptUserEmail = Session.getActiveUser().getEmail();
  Logger.log(scriptUserEmail);
  var query = scriptUserEmail;
  var url = 'URL TO API GOES HERE'
  + '&scriptUserEmail=' + encodeURIComponent(query);

  var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
  var json = response.getContentText();
  var data = JSON.parse(json);
  
  
  for (var i=1; i<data.length; i++){
    var values = data[i];
    if (values[0] == scriptUserEmail){
      var className = values[1];
      var sectionNumber = values[2];
      var startDateTime = values[3];
      var endDateTime = values[4];
      var classLocation = values[5];
      
      var eventID = createEvent(className, sectionNumber, startDateTime, endDateTime, classLocation);
      sheet.getRange("G"+(i+1)).setValue(eventID);
    }
  } 
}

/**
* removes calendar events from Google sheet
* removes eventID from each entry in the google sheet
**/
function removeEventsFromSheet(){
  var scriptUserEmail = Session.getActiveUser().getEmail();
  var calendarId = 'primary';
  Logger.log(scriptUserEmail);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('TeacherSchedule');
  var data = sheet.getDataRange().getValues();
  
  for (var i=1; i<data.length; i++){
    var values = data[i];
    if (values[0] == scriptUserEmail){
      var calendarEventID = values[6];
      removeEvent = Calendar.Events.remove(calendarId, calendarEventID);
      //sheet.getRange("G"+(i+1)).setValue("");
      sheet.getRange("G"+(i+1)).clear()
      
    }
  }
}