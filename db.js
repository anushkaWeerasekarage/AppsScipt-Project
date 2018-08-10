/**
**to update the company name
**works on a timely trigger
**everyhour check the 'Applicant ID' updated last hour
**fetch the right company name and upadate
**/

function updateCompanyName() {
  
  var changeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("1. Member General Information");
  var time = changeSheet.getRange(2, 1, changeSheet.getMaxRows(), 1).getValues();
  var MILLIS_PER_HOUR = 1000 * 60 * 60 * 1;
  var hourBefore = new Date(new Date() - MILLIS_PER_HOUR);
  var i, id, name;
  Logger.log(time);
  try {
    for(i = 0; i < time.length; i++) {
      Logger.log(time[i]);
      if(time[i][0] > hourBefore) {
        Logger.log("inside if");
        id = changeSheet.getRange(i+2, 4).getValue();
        //name = changeSheet.getRange(i+2, 6).getValue()+ " " +changeSheet.getRange(i+2, 7).getValue()
        changeSheet.getRange(i+2, 14).setValue(fetchCompanyName(id));
        
      }
    }
  }
  catch(err) {
    Logger.log(err);
  }
}

/**
to check the id and fetch the company name
**/
function fetchCompanyName(id) {
  var lookInSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("2.Team General Information");
  var values = lookInSheet.getRange(2, 4, lookInSheet.getMaxRows(), 1).getValues();
  var i, comName;
  
  for(i = 0; i < values.length; i++) {
    if(values[i][0] == id) {
      comName = lookInSheet.getRange(i+2, 9).getValue();
      break;
    }
  }
  return comName;
}

/**
** a time based trigger to track all the changes made in the database sheets
**/
function keepLogs() {
  
  var logsheet = SpreadsheetApp.openById("1GEc9QmgArThudCAiL9I3kcRvM0Onorwc_luDB_ExGHI").getSheetByName('Sheet1');
  var row = (logsheet.getLastRow()?logsheet.getLastRow()+1:2);
  var date = new Date();
  var MILLIS_PER_HOUR = 1000 * 60 * 60 * 1;
  var hourBefore = new Date(date.getTime() - MILLIS_PER_HOUR);
  //Logger.log(hourBefore);
  
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var i, values, reverseValues, timeValues, j, k, colLength;
  
  var time, sheetName, url, timeToCompare;
  
  try {
    loop1:for(i = 0; i < sheets.length; i++) {
      if(sheets[i].getRange('A1').getValue() == 'Timestamp') {
        //Logger.log('inside first if');
        //read all the data
        values = sheets[i].getRange(2, 1, sheets[i].getLastRow(), sheets[i].getLastColumn()).getValues();
        colLength = sheets[i].getLastColumn();
        //reverse the array
        reverseValues = values.reverse();
        //Logger.log(reverseValues);
       
        //timeValues = sheets[i].getRange(2, 1, sheets[i].getLastRow(), 1).getValues();
        loop2:for(j = 0; j < reverseValues.length; j++) {
          //Logger.log("inside loop2");
          //Logger.log(reverseValues[j][0]);
            timeToCompare = reverseValues[j][0];
            //Logger.log(timeToCompare);
            if(timeToCompare >= hourBefore) {
              //Logger.log('inside second if');
              time = timeToCompare;
              sheetName = sheets[i].getName();
              url = sheets[i].getParent().getUrl();
              logsheet.getRange(row, 1).setValue(date);
              logsheet.getRange(row, 2).setValue(time);
              logsheet.getRange(row, 3).setValue(sheetName);
              logsheet.getRange(row, 4).setValue(url);
              row++;
            }
          
        }
         
      }
    }
  }
  catch(err) {
    Logger.log(err);
  }
}

/*
** onEdit() function
** sens an email to the applicant when a supervisor is assigned
*/

function onEditSheet(e) {
  //Logger.log("inside onEdit");
  var range = e.range;
  var col = 2;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("2.Team General Information");
  var row, supName, applicantEmail, applicantName;
  var ui, response;
  var user = e.user;
  var ccMail;
  var user = e.user;
  var oldV = e.oldValue;
  var newV = e.value;
  
  if(range.getSheet().getName() == "2.Team General Information" && range.getColumn() == col) {
     //Logger.log('inside if');
     supName = e.value;
     ccMail = supEmail(supName)+', tngrant@ryerson.ca';
     row = range.getRow();
     ui = SpreadsheetApp.getUi();
     response = ui.alert('Are you sure you want to assign supervisor ' +supName+ '?', ui.ButtonSet.YES_NO);
     
     if(response == ui.Button.YES) {
       applicantEmail = sheet.getRange(row, 7).getValue();
       applicantName = sheet.getRange(row, 5).getValue();
       //Logger.log(applicantEmail +','+ supName +','+ applicantName);
       sendMail(applicantEmail, applicantName, supName, ccMail);
       FormSubmissionScript.logsOnEdit(range, user, e.oldValue, supName);
     }
     else {
       range.setValue(e.oldValue);
     }
  }
  
  logsOnEdit(range, user, oldV, newV);
}

function logsOnEdit(range, user, oldV, newV) {
  
  Logger.log("inside logs");
  var date = new Date();
  //var range = e.range;
  var r = range.getA1Notation();
  var sheet = range.getSheet().getName();
  //var user = e.user;
  //var oldV = e.oldValue;
  //var newV = e.value;
  
  var databaseSheet = SpreadsheetApp.openById('1GEc9QmgArThudCAiL9I3kcRvM0Onorwc_luDB_ExGHI').getSheetByName('Sheet2');
  //Logger.log(date +","+ sheet +","+ r +","+ oldV +","+ newV +","+ user);
  databaseSheet.appendRow([date, sheet, r, oldV, newV, user]);
 
  
}

/*
** email notification to applicant on assigning of the supervisor
*/

function sendMail(email, aName, sName, ccMail) {
  try {
    //Logger.log('send mail');
    //Logger.log(MailApp.getRemainingDailyQuota());
    MailApp.sendEmail(email, "New supervisor assigned", "Dear " +aName+ ",\n\n\n  Your new supervisor is: " +sName+ "\n\n Thank you \n\n\n Best, \n iBoost Team.", {cc: ccMail});
  }
  catch(err) {
    Logger.log(err);
  }
}

function supEmail(supName) {
  
  var mail = '';
  if(supName != undefined && supName == 'JP Silva') {
    mail = 'jpsilva@ryerson.ca';
  }
  else if(supName != undefined && supName == 'Rafik Loutfy') {
    mail = 'rloutfy@ryerson.ca';
  }
  else if(supName != undefined && supName == 'Stephen Pumple') {
    mail = 'spumple@ryerson.ca';
  }
  else if(supName != undefined && supName == 'Tarek Sadek') {
    mail = 'tsadek@ryerson.ca';
  }
  else {
    Logger.log('not applicable');
  }
  
  return mail;
}

function test(range, user, oldV, newV) {
  FormSubmissionScript.logsOnEdit(range, user, oldV, newV);
  //MailApp.sendEmail("writetoraminderpal@gmail.com", "subject", "body", {cc: "w.anushka@gmail.com"});
}