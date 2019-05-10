//gets the names of the days of the week
function getDayOfWeek(A) {
  var a = new Date();
  var weekdays = new Array(7);
  weekdays[0] = "Sunday";
  weekdays[1] = "Monday";
  weekdays[2] = "Tuesday";
  weekdays[3] = "Wednesday";
  weekdays[4] = "Thursday";
  weekdays[5] = "Friday";
  weekdays[6] = "Saturday";
  var dayOfWeek = weekdays[A.getDay()];
  return dayOfWeek
}

/**
* checks the var currentDay to see if it is a school holiday.
* Holidays are in the "School Holidays" hidden sheet. 
* If it is, it returns 1, if not returns 0
* @param {string}  currentDay
**/
function isHoliday(currentDay){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("School Holidays");
  var data = sheet.getDataRange().getValues();
  for (var i=1; i<data.length; i++){
    var values = data[i];
    
    if (Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd") == Utilities.formatDate(values[1],"GMT+5:30", "yyyy-MM-dd")){
      var dateIsHoliday = 1;
      return dateIsHoliday;
    }
  }
  var dateIsHoliday = 0;
  return dateIsHoliday;
}

/**
* create blank templates for each period in your schedule.
* if you have a different number of periods, change the number 
* in the for loop.  We are using 8 periods.
**/
function createScheduleTemplates (){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssUrl = ss.getUrl();
  var sheet = ss.getSheetByName("Schedule Templates");
  sheet.appendRow(["Sheet Name","Sheet ID","Sheet Link"]);
  
  for (var i=1; i<=8; i++){
    var sheetCount = ss.getNumSheets();
    var newSheet = ss.insertSheet(sheetCount).setName("P" + i).hideSheet();
    var newSheetID = newSheet.getSheetId();
    var newSheetName = newSheet.getName();
    var newSheetLink = ssUrl + "#gid=" + newSheetID;
    sheet.appendRow([newSheetName,newSheetID,newSheetLink]);
    newSheet.appendRow(["Class Name", "StartDateTime", "EndDateTime", "Term"])
  }
}

//used for testing
function deleteScheduleTemplates(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssUrl = ss.getUrl();
  var sheet = ss.getSheetByName("Schedule Templates");
  var data = sheet.getDataRange().getValues();
  for (var i=1; i<data.length; i++){
    var values = data[i];
    var removeSheet = ss.getSheetByName(values[0]);
    ss.deleteSheet(removeSheet);
  }
  sheet.getDataRange().clear();
}

function populateTemplates(){
  //clears previous entries from template
  clearTemplates();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssUrl = ss.getUrl();
  var p1Sheet = ss.getSheetByName("P1");
  var p2Sheet = ss.getSheetByName("P2");
  var p3Sheet = ss.getSheetByName("P3");
  var p4Sheet = ss.getSheetByName("P4");
  var p5Sheet = ss.getSheetByName("P5");
  var p6Sheet = ss.getSheetByName("P6");
  var p7Sheet = ss.getSheetByName("P7");
  var p8Sheet = ss.getSheetByName("P8");
  
  /**
  * first day of your regular of each term (Sem1 & Sem2 for us).  
  * We are in GMT+5:30.  Change for your timezone
  * to fit your needs.
  **/
  var firstDay = [new Date('August 07, 2018 00:00:00 +0530'),new Date('January 01, 2019 00:00:00 +0530')];
  //last day of each term
  var lastDay = [new Date('December 21, 2018 15:30:00 +0530'),new Date('May 31, 2019 15:30:00 +0530')];
    //iterate each semester
    for(var i = 0;i<2;i++){
      //initialize to day schedule.  We are on a 1-8 day schedule
      var classDay = 1;
      var semester = "Semester " + (i+1);
      //initialize first day of each semester
      var currentDay = firstDay[i];
      //sets if it is the list of holidays for IF statement below
      var dateIsHoliday = isHoliday(currentDay);
      //sets the day of the week for IF statement below
      var dayOfWeek = getDayOfWeek(currentDay);
      while(currentDay < lastDay[i]){
        //make sure it is not weekend or holiday
        if(dayOfWeek !== "Saturday" && dayOfWeek !== "Sunday"  && !dateIsHoliday){
          /**
          * we have an early release day with adifferent
          * schedule on wednesday.  If you do not,
          you can remove this IF and set the timings for
          * the block once
          **/
          if (currentDay.getDay() == 3){
            var blockOneStart = "T08:30:00";
            var blockOneFinish = "T09:40:00";
            var blockTwoStart = "T10:10:00";
            var blockTwoFinish = "T11:20:00";
            var blockThreeStart = "T12:00:00";
            var blockThreeFinish = "T13:10:00";
            var blockFourStart = "T13:20:00";
            var blockFourFinish = "T14:30:00";
          }
          else{
            var blockOneStart = "T08:30:00";
            var blockOneFinish = "T09:55:00";
            var blockTwoStart = "T10:25:00";
            var blockTwoFinish = "T11:50:00";
            var blockThreeStart = "T12:35:00";
            var blockThreeFinish = "T14:00:00";
            var blockFourStart = "T14:10:00";
            var blockFourFinish = "T15:35:00";
          }
          /**
          * Gets the day (1-8) and adds an entry to the period
          * template sheets (P1-P8) for each class for each day.  
          * We are on a 8 day tumbling schedule, adjust as needed.
          * Also ass whis senester the entry is in.  This is used
          * in "createEventsFromSheet()" function in the "AddSchedulToCalendar"
          * script.
          **/
          switch(classDay){
            case 1:
              p1Sheet.appendRow(["P1",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockOneStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockOneFinish, semester])
              p2Sheet.appendRow(["P2",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockTwoStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockTwoFinish, semester])
              p3Sheet.appendRow(["P3",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockThreeStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockThreeFinish, semester])
              p4Sheet.appendRow(["P4",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockFourStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockFourFinish, semester])
              break;
            case 2:
              p5Sheet.appendRow(["P5",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockOneStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockOneFinish, semester])
              p6Sheet.appendRow(["P6",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockTwoStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockTwoFinish, semester])
              p7Sheet.appendRow(["P7",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockThreeStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockThreeFinish, semester])
              p8Sheet.appendRow(["P8",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockFourStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockFourFinish, semester])
              break;
            case 3:
              p2Sheet.appendRow(["P2",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockOneStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockOneFinish, semester])
              p3Sheet.appendRow(["P3",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockTwoStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockTwoFinish, semester])
              p4Sheet.appendRow(["P4",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockThreeStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockThreeFinish, semester])
              p1Sheet.appendRow(["P1",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockFourStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockFourFinish, semester])
              break;
            case 4:
              p6Sheet.appendRow(["P6",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockOneStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockOneFinish, semester])
              p7Sheet.appendRow(["P7",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockTwoStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockTwoFinish, semester])
              p8Sheet.appendRow(["P8",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockThreeStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockThreeFinish, semester])
              p5Sheet.appendRow(["P5",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockFourStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockFourFinish,semester])
              break;
            case 5:
              p3Sheet.appendRow(["P3",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockOneStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockOneFinish, semester])
              p4Sheet.appendRow(["P4",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockTwoStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockTwoFinish, semester])
              p1Sheet.appendRow(["P1",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockThreeStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockThreeFinish, semester])
              p2Sheet.appendRow(["P2",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockFourStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockFourFinish, semester])
              break;
            case 6:
              p7Sheet.appendRow(["P7",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockOneStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockOneFinish, semester])
              p8Sheet.appendRow(["P8",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockTwoStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockTwoFinish, semester])
              p5Sheet.appendRow(["P5",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockThreeStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockThreeFinish, semester])
              p6Sheet.appendRow(["P6",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockFourStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockFourFinish, semester])
              break;
            case 7:
              p4Sheet.appendRow(["P4",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockOneStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockOneFinish, semester])
              p1Sheet.appendRow(["P1",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockTwoStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockTwoFinish, semester])
              p2Sheet.appendRow(["P2",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockThreeStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockThreeFinish, semester])
              p3Sheet.appendRow(["P3",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockFourStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockFourFinish, semester])
              break;
            case 8:
              p8Sheet.appendRow(["P8",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockOneStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockOneFinish, semester])
              p5Sheet.appendRow(["P5",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockTwoStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockTwoFinish, semester])
              p6Sheet.appendRow(["P6",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockThreeStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockThreeFinish, semester])
              p7Sheet.appendRow(["P7",Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockFourStart,Utilities.formatDate(currentDay,"GMT+5:30", "yyyy-MM-dd")+blockFourFinish, semester])
              break;   
          }
          //increments the day, unless it is day 8 and the starts over
          if (classDay < 8){
            classDay++;
          }
          else{
            classDay = 1;
          }
        }
        //increments to the next day
        var nextDay = new Date(currentDay.setDate(currentDay.getDate()+1));
        currentDay = nextDay;
        //checks if currentDay is holiday or weekend
        dateIsHoliday = isHoliday(currentDay);
        dayOfWeek = getDayOfWeek(currentDay);
     }
  } 
}

//not used in production.  Clears out P1-P8 templates
function clearTemplates(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  for (var i=0;i<8;i++){
    var periodSheet = ss.getSheetByName("P"+(i+1));
    var range = periodSheet.getRange(2, 1, periodSheet.getLastRow(), 4);
    range.clear();
  }
    

}

