//Adds "Teacher Schedules" custom Menu
function onOpen() {
  
  var menu = [{name: 'Remove Teacher Schedules from Google Calendar', functionName: 'removeEventsFromSheet'}];
  SpreadsheetApp.getActive().addMenu('Teacher Schedules', menu);
}

/** 
  * Activate Script button must be called first because the 
  * createEventsFromScript function uses Oauth and permissions 
  * dialog will not come up when the "createEventsFromSheet" 
  * function is called directly from a menu for a regular user
  **/
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

function createEvent(className, startDateTime, endDateTime, classLocation) {
  //uncomment this in production
  var calendarId = 'primary';

  var formattedStartDate = Utilities.formatDate(startDateTime, "GMT+5:30", "yyyy-MM-dd'T'HH:mm:ssZ");
  var formattedEndDate = Utilities.formatDate(endDateTime, "GMT+5:30", "yyyy-MM-dd'T'HH:mm:ssZ");
  
  var event = {
    summary: className,
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
    //this allows administrators to see teacher's schedule even if their
    //primary is marked as private 
     visibility: "public",
    // Red background. Use Calendar.Colors.get() for the full list.reminders.overrides[]
    colorId: 11
  };
  event = Calendar.Events.insert(event, calendarId);
  var eventID = event.getId();
  return eventID;
}

/**
* Creates individual events for each class (one of every class, every day
* that you have a class).  This function pulls the class name from the 
* google sheet.  If the class name is blank, it will not create an event
* for the period.  For each event created, the eventID gets saved to a 
* logfile that you can use to delete all events at a later point.  All
* sheets in the logfile get saved for each user in a sheet with their 
* email address as the name.

**/

function createEventsFromSheet(){
  var scriptUserEmail = Session.getActiveUser().getEmail();
  //Logger.log(scriptUserEmail);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Make Class Schedules');
  
  var data = sheet.getDataRange().getValues();
  //periodNames[0] is for sem1 class names, periodNames[1] is for sem2 class names
  var periodNames = new Array(2);
  periodNames = [new Array(8), new Array(8)];
  //periodLocation[0] is for sem1 room numbers, periodLocation[1] is for sem2 room numbers
  var periodLocation = new Array(2);
  periodLocation = [new Array(8), new Array(8)];
  
  /**
  * put the sheetid of the file you would like to save the 
  * log file to here
  **/
  //var logSs = SpreadsheetApp.openById("1dkXQ1Nv7FUNmWq3yIknLH_u9hl96gjtm5qprOqK6Gwk");
  //check to see if log file for the user exists
  var logSheet = ss.getSheetByName(scriptUserEmail);
  if(!logSheet){
    var logSheet = ss.insertSheet(0).setName(scriptUserEmail).hideSheet();
  }
  //puts the header row on the log file
  logSheet.appendRow(["Teacher Email", "Class Name", "Event Start Date-Time", "Event End Date-Time", "Class Location", "EventID"]);
  
  //add the class names and room numbers to an array
  for (var i=1; i<data.length; i++){
    var values = data[i];
    periodNames[0][i-1] = values[1];
    periodNames[1][i-1] = values[3];
    periodLocation[0][i-1] = values[2];
    periodLocation[1][i-1] = values[4];
  }
    /**
    * we are using an 8 period tumbeling schedule
    * with 4 blocks each day.  The templates used to
    * create the events are created are in 
    * CreateScheduleTemplates.gs.  The for loop below
    * will need to be changed if you have a different
    * number of periods.
    **/
    for (var i=0;i<9;i++){
      var periodSheet = ss.getSheetByName("P"+(i+1));
      var periodData = periodSheet.getDataRange().getValues();
      
      for ( var j=1; j<periodData.length; j++){
        var periodValues = periodData[j];
        
        /**
        * pulling semester info from the template files.
        * we are only using 2 semesters for our terms
        * if you are different, you will need to change the below to 
        * suit your needs.
        * This determines which column from "Make Class Schedules" is 
        * being used to pull the event name (from sem1 or sem2).
        **/
        if (periodValues[3] == "Semester 1"){
          var className = periodNames[0][i];
        
          var startDateTime = new Date(periodValues[1]);
          var endDateTime = new Date(periodValues[2]);
          var classLocation = periodLocation[0][i];
         
        }
        //It must be semester 2
        else{   
          if (periodNames[1][i] == ""){
            var className = "";
          }
          else{
            var className = periodNames[1][i];
          }
        
          var startDateTime = new Date(periodValues[1]);
          var endDateTime = new Date(periodValues[2]);
          var classLocation = periodLocation[1][i];
          
        }
        
        /**
        * Creates an event for each entry in each template file that 
        * matches up with a period entry in "Make Class Schedules"
        * It does not create an event if there is no name in the 
        * "Make Class Schedules" for a period.  A log enty is created
        * for every event created so that events can be removed if
        * it need to be changed later        
        **/
        if(className != ""){
          var eventID = createEvent(className, startDateTime, endDateTime, classLocation);
          logSheet.appendRow([scriptUserEmail, className, startDateTime, endDateTime, classLocation,eventID])
        }
      }
    } 
  
}

/**
* removes calendar events from Google sheet
* removes eventID from each entry in the google sheet
* for each user that has run the script.  Each user has
* their own log sheet with their eventIDs that were
* save when they ran the script.
* The function is run from the "Teacher Schedules"
* custom menu.
**/
function removeEventsFromSheet(){
  var scriptUserEmail = Session.getActiveUser().getEmail();
  var calendarId = 'primary';
  //replace with the spreadsheetID of your logsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // gets the sheetname for the logsheet for the user.  It will be the user's email address
  var sheet = ss.getSheetByName(scriptUserEmail);
  var data = sheet.getDataRange().getValues();
  
  for (var i=1; i<data.length; i++){
    var eventId = data[i][5];
    var removeEvent = Calendar.Events.remove(calendarId, eventId);
  }
  //deletes sheet after all events have been deleted
  ss.deleteSheet(sheet)
}

/**
* used when testing to clear out the calendar of all the events that are going to
* a testing calendar.  Not used in production
**/
function deleteAllEvents(){
    //replace with your calendar ID
    var cal = CalendarApp.getCalendarById("aes.ac.in_t1ra2v8sad3gt8hsbpuugldpd4@group.calendar.google.com");
    var events = cal.getEvents(new Date("July 31, 2018 00:00:00"), new Date("June 1, 2019 00:00:00 +530"));
    for(var i=0;i<events.length;i++){
      var ev = events[i];
      ev.deleteEvent();
      //added this when I started getting the below error
      //Service invoked too many times in a short time: Calendar
      Utilities.sleep(250)
    }
}
