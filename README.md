# addTeacherSchedulesToGoogleCalendar
These scripts take input on Google sheet with a teacher's schedule and add entries to the user's primary calendar.  The script uses the Advanced Google Services for Calendar and must be enabled.  Instruction for enabling the can be found [here](https://developers.google.com/apps-script/guides/services/advanced).

## Before Use

A [copy of the underlying sheet and script](https://docs.google.com/spreadsheets/d/1s9rIxoga-0W8XAH5jFNvIobYAUnFOjWSeEM-l2VHo1g/copy) can be found here.  There is one visible sheet called "Make class schedules" that the scripts are tied to.  Teachers should put in their schedules for each block they have classes in.
There is a hidden sheet called "School Holidays" that the days off for each semester that will have to be changed.  This ensures that the script that creates the templates does not create events on days off from school.

In the script called "CreateScheduleTemplates" there is a function called "populateTemplates".  The first and last days of each term (for us it is semester 1 & 2) needs to be but in. You will have to also change the timings for you blocks.  We are on a 8 day tumbling rotation with 4 block per day.  We also have an early release on wednesdays.  The script will have to be changed to reflect your schedule.  Once the changes have been made the "populateTemplates" function will need to be run.  This function populates each day's template with the timings for each class for the entire school year.  When the templates have been populated, the sheet can be shared with teacher to use.

You will need to create spreadsheet to store the log files in.  The ID for this sheet need to be put in lines 87 & 181 of AddScheduleToCalendar.gs

## Teacher Use
The script needs to be activated by each user before it can create schedules for them.  I think this is because I am using the Advanced Google Services for Calendar API and the OAuth calls it makes.  Once the script has been activated by pressing the green button (which is tied to the activateMyScript() function on AddScheduleToCalendar.gs), the teacher should put their schedule into the periods for each term.  If a class is a year-long class it will need to be put in each term.

Once their classes have been put in, they can run the script by clicking the Red button.  The button is tied to createEventsFromSheet() function on AddScheduleToCalendar.gs.  The script then creates a calendar event for each entry in the day template where the teacher has added a class name in the sheet.  Each event created is then logged to a seperate spreadsheet so that the events can be removed programatically if need be. The log entries are stored in a sheet created that is the user's email address.  

If the teacher would like to remove the events at a later time, the can run the removeEventsFromSheet() function by choosing "Remove Teacher Schedules from Google Calendar" in the Teachers Schedules menu.  This will load the data from the teacher's log sheet and remove each event listed in the sheet.  The logsheet deletes itself after completion.
