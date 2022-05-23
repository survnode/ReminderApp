/*--------------------------------------------------------------------------------------
 * SENDS EMAIL REMINDERS AFTER THE NUMBER OF DAYS SPECIFIED BY THE USER.
 *--------------------------------------------------------------------------------------
 *
 * 
 * 
 Licensed under the Apache License, Version 2.0 (the "License"); you may not
 use this file except in compliance with the License.  You may obtain a copy
 of the License at

     https://www.apache.org/licenses/LICENSE-2.0

 Unless required by applicable law or agreed to in writing, software
 distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
 WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.  See the
 License for the specific language governing permissions and limitations under
 the License.
 */

function onOpen() {
 createCustomMenu()
} 

/******************************************************************************************
 * This is the primary function that sends emails. It loops through each due date and sends and email if today matches the due date.
 * The function runs on daily trigger and calls the clearJMTrigger() function which deletes the daily trigger after a specified number
 * of days have passed since the maximum due date.
 * 
 */
function JM_Reminder() {

  const mainEmail = 'your email address'; //the primary email address to receive reminders
  const ccEmail = 'cc email (optional)'; //the email address that will be CC'd for each reminder.
  const delTrigDays = 30; //the number of days after which an inactive trigger should be deleted

  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var numRows = sheet.getLastRow()-1;   // Number of rows to process
  var numOfColumns = sheet.getLastColumn();  // Number of columns to process
  var dataRange = sheet.getRange(startRow, 1, numRows, numOfColumns); //create a row/column range
  var data = dataRange.getValues(); 
  
  var dt=new Date();
  var dv=new Date(dt.getFullYear(),dt.getMonth(),dt.getDate()).valueOf(); //value of today's date

  for (var i=0;i<data.length;i++) {
    var row = data[i];

    var action   = row[1]; // action
    var activity   = row[2]; // activity
    var dayspast = row[3]; // days to reminder column, from date when file was received 
    var sentto   = row[4]; // cc address column  
    var d=new Date(row[5]); // due date/reminder date column
    var r= new Date(row[0]); //date received column
    
    //convert dates to local time 
    var remindDateFormat = Utilities.formatDate(d, 'Africa/Nairobi', "EEE, d MMM yyyy"); 
    var receivedDateFormat = Utilities.formatDate(r, 'Africa/Nairobi', "EEE, d MMM yyyy");

    var htmlOutput = HtmlService.createHtmlOutputFromFile('message1'); // HTML file containing email contents
    var message = htmlOutput.getContent()

    //replace HTML variables 
      message = message.replace("%action", action);
      message = message.replace("%activity", activity);
      message = message.replace("%receivedDateFormat", receivedDateFormat);
      message = message.replace("%dayspast", dayspast);
      message = message.replace("%remindDateFormat", remindDateFormat);
    
    var lv = new Date(d.getFullYear(),d.getMonth(),d.getDate()).valueOf(); // reminder date

    if (lv == dv) { // if today equal reminder date

      var subject =  // subject of the email
        'REMINDER:-  ' + action; //+ ' - ' + remindDateFormat
      
      MailApp.sendEmail({   // send email to:
      to:mainEmail,         // main user
      cc:ccEmail+","+sentto, // main cc account and any account in the cc address column 
      subject: subject,
      htmlBody: message
      });
     // row+1 because array are 0 indexed while range is 1 indexed 
    dataRange.getCell(i+1, numOfColumns).setBackground("#00FA9A"); // highlight due date cell
    dataRange.getCell(i+1, numOfColumns).setHorizontalAlignment("center"); // align center
    dataRange.getCell(i+1, numOfColumns).setFontWeight("bold"); // style text bold
    dataRange.getCell(i+1, numOfColumns).setNote("*Reminder_Sent!*"); // add a note
    } 
  }
  clearJMTrigger(delTrigDays);   //clear the daily reminder trigger x number of days (default is 30 days) after the last email for the maximum date was sent  
}

/******************************************************************************************
 * This is an auxiliary function that complements the main function. It serves the Administrator by alerting them
 * when a new sheet is created. It runs the following functions:
 * - cloneRecordsSpreadsheet() 
 * - deleteTriggersByName().
 *
 */
function newSheet_Rem() {
    
   
   const delTrigDays = 30; //the number of days after which an inactive trigger should be deleted

    var dateTime  =new Date(); // find first day of each month
    var localTime = convertTZ(dateTime,'Africa/Nairobi');
    var nowTime = new Date(localTime);
    var nowDay = nowTime.getDate();

    var d = new Date(); 
    d.setDate(nowTime.getDate() - 1);// yesterday
    
    var secondDateFormat = Utilities.formatDate(nowTime, 'Africa/Nairobi', "MMMM yyyy").toUpperCase();
    var secDateFormat = Utilities.formatDate(nowTime, 'Africa/Nairobi', "MMMM, yyyy");
    var yestDateFormat = Utilities.formatDate(d, 'Africa/Nairobi', "MMMM yyyy").toUpperCase();


    var mailaddress = Session.getEffectiveUser().getEmail();  // get admin of the app

    var tutorialUrl = "https://drive.google.com/uc?id=xxxxxxx"; //public URL for .gif tutorial file
  
    var tutorialBlob = UrlFetchApp
                            .fetch(tutorialUrl, {muteHttpExceptions: true })
                            .getBlob()
                            .setName("Tutorial");  // required to append .gif to email

    var htmlOutput = HtmlService.createHtmlOutputFromFile('message2'); // HTML file containing email contents
    var message = htmlOutput.getContent()

    //replace HTML variables 
      message = message.replace("%secDateFormat", secDateFormat);
      message = message.replace("%yestDateFormat", yestDateFormat);
      message = message.replace("%delTrigDays", delTrigDays);
      message = message.replace("%secondDateFormat ", secondDateFormat);
      message = message.replace("%yestDateFormat", yestDateFormat);
      message = message+"<img src='cid:tutorial' style='width:835px; height:411px;'/>"; // add .gif file

    if (nowDay == 1) // if today is the first day of the month
      {     
       cloneRecordsSpreadsheet();  // make a copy of template worksheet
        
        // notify administrator that a copy has been made and attach a .gif file demo on how to activate it
        var subject =
        'NEW WORKSHEET ALERT: - ' + secondDateFormat+' has been created!';
       
        MailApp.sendEmail(mailaddress, subject, "",
                    { htmlBody: message,
                      inlineImages:
                      {
                        tutorial: tutorialBlob,
                      }
                    });
    
       deleteTriggersByName("newSheet_Rem"); //Remove this function's trigger since it only required to run once per worksheet/month
      }
  
 
  Logger.log("done!");
  
}

/******************************************************************************************
 * Deletes any triggers of the main JM_Reminder() after a specified number of days. 
 *
 */
  function clearJMTrigger(tot_days) {
    
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var sheet = ss.getSheetByName("RECORDS");  
  var range = sheet.getDataRange(); 
  var values = range.getDisplayValues();
  var lastRow = range.getLastRow();
  var max = 0;

 for (i = 0; i < lastRow; i++)  //iterate through the reminder date column and find the maximum date
  {  
    if (new Date(values[i][15]) > max)
    max = new Date(values[i][15]);
  }  
  
   var today = new Date();
   var mdate = new Date(max);
   mdate.setDate(mdate.getDate() + tot_days); // add specific number of days to the max date

   var dv=new Date(today.getFullYear(),today.getMonth(),today.getDate()).valueOf(); //value of today's date
   var ld=new Date(mdate.getFullYear(),mdate.getMonth(),mdate.getDate()).valueOf(); //value of delete date

if (dv == ld) // if today is equal to delete date then delete reminder trigger
      {  
       deleteTriggersByName("JM_Reminder")
      } 

}
/******************************************************************************************
 * This function creates a copy of a specified template file by the file ID.
 * 
 */
function cloneRecordsSpreadsheet() {
  
  const template = DriveApp.getFileById("1UXfr-xxxxx"); //spreadheet ID to duplicate each month. This should contain the template of your reminders
  

  var dateTime  =new Date(); // find first day of each month
  var localTime = convertTZ(dateTime,'Africa/Nairobi');
  var nowTime = new Date(localTime);
 
  var secondDateFormat = Utilities.formatDate(nowTime, 'Africa/Nairobi', "MMMM yyyy").toUpperCase();
   
    
  var editors = SpreadsheetApp.getActiveSpreadsheet().getEditors(); // get editors of current file
  
  var newCopy = template.makeCopy(); //copy template
  newCopy.setName(secondDateFormat); //rename template

  for (var i = 0; i<editors.length;i++) { //loop through editors of current file and add them to new file
    newCopy.addEditor(editors[i])
  }
  newCopy.setShareableByEditors(false);
 
} 


/******************************************************************************************
 * Custom menu function to initialize authorization.
 * 
 */
function authorizeApp() {
var count = ScriptApp.getProjectTriggers().length;
if(!isAdmin())  
   {
    SpreadsheetApp.getUi().alert("Admin access required");
   }  
else if (count !== null)
        {
         SpreadsheetApp.getActive().toast("Application is already authorized.");
        }  
Utilities.sleep(1000);
}

/******************************************************************************************
 * Custom menu function to create triggers.
 * 
 */
function activateRems() {

  if(isAdmin())  //restrict trigger creation to admin(s) only
   {
     if (findTriggerbyFunction("JM_Reminder") === false)
        {
         createJMRemTrigger();  
        }
     if (findTriggerbyFunction("newSheet_Rem") === false)
        {
         createNSRemTrigger();  
        }
      else {
            SpreadsheetApp.getActive().toast("Reminders are already active. Only one activation per sheet is required.");}
            }     
   else {
         SpreadsheetApp.getUi().alert("Admin access required"); 
        }
}

/******************************************************************************************
 * Custom menu function to delete triggers by calling deleteTriggers() function.
 * 
 */
function deleteRems() {

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('⛔ NOTE: You are about to disable reminders. Any pending reminder emails will not be sent in future.'+ 
                          ' Are you sure you wish to continue?', ui.ButtonSet.YES_NO);
  
  if(isAdmin())  //restrict trigger deletion to admin(s) only
    {
      
      if(response == ui.Button.YES) 
            {
             deleteAllTriggers(); 
             SpreadsheetApp.getActive().toast("Reminders have been cleared.")
            }   
          else if(response == ui.Button.NO)      
                {
                  return;
                 }
    }             
   else {
         ui.alert("Admin access required"); 
        }
}

/******************************************************************************************
 * Function to create main function trigger. Trigger should fire between 9 and 10 AM daily
 * 
 */
function createJMRemTrigger() {
   ScriptApp.newTrigger("JM_Reminder")
            .timeBased()
            .inTimezone("Africa/Nairobi") //specify TZ
            .atHour(9)  // hour to fire
            .everyDays(1) // on a daily basis
            .create();
}

/******************************************************************************************
 * Function to create auxiliary function trigger for creating new file. 
 * Trigger should fire between Midnight and 1AM daily
 * 
 */
 function createNSRemTrigger() {
   ScriptApp.newTrigger("newSheet_Rem")
            .timeBased()
            .inTimezone("Africa/Nairobi")
            .atHour(0)
            .everyDays(1) //.onMonthDay(1)
            .create();
}

/******************************************************************************************
 * Boolean function to iterate through triggers and find a specific handler function's trigger.
 * TRUE if function's trigger is present and FALSE if trigger is absent
 * 
 */
function findTriggerbyFunction(theFunction) {

var triggers = ScriptApp.getProjectTriggers();
var findTrigger = false;
triggers.forEach(function (trigger) {
  if(trigger.getHandlerFunction() === theFunction)
    findTrigger = true;
});
return findTrigger;
}

/******************************************************************************************
 * Boolean function to determine whether current sheet user has Admin privileges to functions.
 * 
 */
function isAdmin() {
var admins = []; //array that will hold admninistrator(s) of the app, set to dynamic to allow easy add/removal
admins.push(""); // add admins email addresses without the '@gmail.com' eg. admin instead of admin@gmail.com

var adminStatus = false;
var user =  Session.getActiveUser().getEmail();
var id = user.substring(0, user.indexOf("@"));

for (var i = 0; i < admins.length; i++){ 
    if (admins[i] === id) 
    {
      adminStatus = true;
    }
};
 return adminStatus;  
}

/******************************************************************************************
 * This function iterates through all project triggers and deletes any trigger with a specified
 * handler function.
 * 
 */
function deleteTriggersByName(name){
var triggers = ScriptApp.getProjectTriggers();
for (var i = 0; i < triggers.length; i++){ 
    if (triggers[i].getHandlerFunction().indexOf(name) != -1) 
    {
      ScriptApp.deleteTrigger(triggers[i]);
    }
}}

/******************************************************************************************
 * Simple trigger to call custom menu.
 * 
 */
 function createCustomMenu() {

     var ui = SpreadsheetApp.getUi(); 
         ui.createMenu('☸ Admin Settings')
           .addItem('📧  Authorize Application', 'authorizeApp')
           .addSeparator()
           .addItem('🔃  Activate Reminders', 'activateRems')
           .addItem('♻  Clear Reminders', 'deleteRems')
           .addToUi();
     
}
  
/******************************************************************************************
 * Deletes all the triggers.
 *
 */
function deleteAllTriggers(){
  var triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(function(trigger){

    try{
      ScriptApp.deleteTrigger(trigger);
    } catch(e) {
      throw e.message;
    };

    Utilities.sleep(1000);

  });

};

/******************************************************************************************
 *Converts current time to specific timezone according to locales. See list in link below
 *https://gist.github.com/diogocapela/12c6617fc87607d11fd62d2a4f42b02a
 *
 */
function convertTZ(date, tzString) {
    return new Date((typeof date === "string" ? new Date(date) : date).toLocaleString("en-US", {timeZone: tzString}));   
 }
