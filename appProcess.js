var spreadsheet = SpreadsheetApp.openById("1sSDDu0wUzuPatvencCnaQO7Mu_BPLFMCf-XZgU-Jp1g"); //database spreadsheet
var imageLogo = DriveApp.getFileById('1HOUBFBGYLfUM6_HhKlBqS9mwXgrCj_ow').getBlob();
//var logoBlob = UrlFetchApp.fetch(imageLogo).getBlob().setName("logoBlob");
/*
 **trigger for Application form
 **check both 'status' and the 'database_status' of the application
 **if both status are true update the database
*/

function onEditSheet(e) {
  //Logger.log("onedit trigger");
  //Logger.log(e.source.getActiveRange());
  var eRange = e.source.getActiveRange();
  var user = e.user;
  var oldV = e.oldValue;
  var newV = e.value;
  var reuestSheet = SpreadsheetApp.getActive().getSheetByName("1.Application Forms"); 
  var memSheet = SpreadsheetApp.getActive().getSheetByName("3. Register New Member");
  var status = memSheet.getLastColumn() - 1;
  var sheet = spreadsheet.getSheetByName("2.Team General Information"); 
  var sheet2 = spreadsheet.getSheetByName("1. Member General Information"); 
  var lastCol = reuestSheet.getLastColumn(); 
  var statusCol = lastCol - 1;
  var nextRowClose = (sheet.getLastRow()?sheet.getLastRow()+1:2); //73 is the starting row with data.
  //Logger.log(nextRowClose);
  var nextRowClose2 = (sheet2.getLastRow()?sheet2.getLastRow()+1:2);
  var values, id, idList, i, memId, memIdList;
  var ui, response, email, name;
  var arr = [];
  /*
    **Application acceptance
  */
  if(eRange.getSheet().getName()=="1.Application Forms" && eRange.getColumn()== statusCol) {
    //Logger.log("inside application");
    if(eRange.getValue()=="Accept") {
      values = reuestSheet.getRange(eRange.getRow(), 2, 1, 14).getValues(); //may have to update with the changes of column positions of APPROVE/REJECT_Form Collection, 1.Application Forms
      //Logger.log(values);
      // Display a dialog box with a message and "Yes" and "No" buttons.
      ui = SpreadsheetApp.getUi();
      response = ui.alert('Are you sure you want to accept this application?', ui.ButtonSet.YES_NO);
      
      // Process the user's response.
      if (response == ui.Button.YES) {
        //Logger.log('The user clicked "Yes."');
        id = values[0][1];
        email = values[0][4];
        name = values[0][2];
        sheet.getRange(nextRowClose, 3, 1, 14).setValues(values); //may have to update with the changes of column psitions of iDB, 2.Team General Information sheet
        sheet.getRange(nextRowClose, 1).setValue(new Date());
        reuestSheet.getRange(eRange.getRow(), lastCol).setValue("sent");
        acceptNotification(email, name, id);
      } 
      else if(response == ui.Button.NO){
        //Logger.log('The user clicked "No" or the dialog\'s close button.');
        reuestSheet.getRange(eRange.getRow(), statusCol).setValue("Pending");
      }
      else {
        Logger.log("User clicked the close button");
      }
    }
    
    else if(eRange.getValue()=="Reject") {
      ////'Applicant ID' of the edited row
      id = reuestSheet.getRange(eRange.getRow(), 3).getValue();
      ////all the 'Applicant ID's in database sheet
      idList = sheet.getRange(2, 4, sheet.getMaxRows(), 1).getValues();
      ////check the id exists in the database sheet
      for(i = 0; i < idList.length; i++) {
        //Logger.log("inside if");
        if(idList[i] == id) {
          sheet.deleteRow(i+2);
          reuestSheet.getRange(eRange.getRow(), lastCol).clear();
          break;
        }
      }
    }
     
  }
    /*
      **Member Registraion
    */
 else if(eRange.getSheet().getName()=="3. Register New Member"  &&  eRange.getColumn()==status) {
    if(eRange.getValue()=="Accept") {
        //Logger.log("got the sheet");
        values = memSheet.getRange(eRange.getRow(), 3, 1, status - 2).getValues();
        Logger.log(values);
        //// Display a dialog box with a message and "Yes" and "No" buttons.
        ui = SpreadsheetApp.getUi();
        response = ui.alert('Are you sure you want to accept this application?', ui.ButtonSet.YES_NO);
        
        //// Process the user's response.
        if (response == ui.Button.YES) {
          //Logger.log('The user clicked "Yes."');
          name = values[0][1];
          email = values[0][3];
          
          sheet2.getRange(nextRowClose2, 1).setValue(new Date());
          sheet2.getRange(nextRowClose2, 4, 1, 10).setValues(values);
          
          memSheet.getRange(eRange.getRow(), memSheet.getLastColumn()).setValue("sent");
          welcomeMem(name,email);
        } 
        else {
          //Logger.log('The user clicked "No" or the dialog\'s close button.');
          memSheet.getRange(eRange.getRow(), status).setValue("Pending");
        }
      }
      
      else if(eRange.getValue()=="Reject") {
        //Logger.log("inside reject");
        
        memId = memSheet.getRange(eRange.getRow(), 3).getValue();
        memIdList =  sheet2.getRange(2, 4,  sheet2.getMaxRows(), 1).getValues();
        
        for(i = 0; i < memIdList.length; i++) {
          if(memId == memIdList[i]) {
            sheet2.deleteRow(i+2);
            memSheet.getRange(eRange.getRow(), memSheet.getLastColumn()).clear();
            break;
          }
        }
        
      }
     
    }
  logsOnEdit(eRange, user, oldV, newV);
}

/*
**** on form submission trigger, for the forms,
2. Register New Team
3. Register New Member
*/

function onFormSubmit(event) {
  Logger.log("onformsubmit trigger");
  var eRange = event.range;
  var appSheet = SpreadsheetApp.getActive().getSheetByName("1.Application Forms");
  var requestSheet = SpreadsheetApp.getActive().getSheetByName("2. Register New Team");
  var newMemSheet = SpreadsheetApp.getActive().getSheetByName("3. Register New Member");
  
  ////database spreadsheet
  //var spreadsheet = SpreadsheetApp.openById("1sSDDu0wUzuPatvencCnaQO7Mu_BPLFMCf-XZgU-Jp1g"); 
  var sheet = spreadsheet.getSheetByName("2.Team General Information");
  var sheet2 = spreadsheet.getSheetByName("1. Member General Information");
  var nextRow = (sheet2.getLastRow()?sheet2.getLastRow()+1:2);; 
  var values, id, idList, i, row;
  var arr = [];
  var comName, memName;
  
  try {
    if(eRange.getSheet().getName()=="2. Register New Team") {
      values = event.values;
     Logger.log(values);
      id = values[2];
      ////remove first two values (Timestamp/ For id's/Team ID Code)
      values.splice(0, 2);
      ////create a two dimensional array
      arr.push(values.splice(0,10));
      Logger.log(arr);
      idList = sheet.getRange(2, 4, sheet.getMaxRows(), 1).getValues();
      Logger.log(idList);
      ////check the id exists in the database sheet
      
      for(i = 0; i < idList.length; i++) {
        if(idList[i] == id) {
          Logger.log('inside if');
          row = i + 2; ////change here
          sheet.getRange(row, 17, 1, 10).setValues(arr);
          sheet.getRange(row, 1).setValue(new Date());
          requestSheet.getRange(eRange.getRow(),14).setValue("sent");
          // geting email notification 
          id = requestSheet.getRange(eRange.getRow(),3).getValue();
          //Logger.log(id);
          comName = companyName(id);
          //Logger.log(comName);
          newTeamNotify(id, comName); //send the email to iboost notify of new team registered
          break;
        }
      }
 
    }
    else if(eRange.getSheet().getName()=="3. Register New Member") {
      
      values = event.values;
      Logger.log(values);
      id = values[2];
      memName = values[3];
      comName = companyName(id);
      
      newMemNotify(id, comName, memName); //send the email to iboost
      newMemSheet.getRange(eRange.getRow(), newMemSheet.getLastColumn()-1).setValue("Pending"); //update the status of the sheet
      
    }
    
    else if(eRange.getSheet().getName() == "1.Application Forms") {
      
      appSheet.getRange(eRange.getRow(), appSheet.getLastColumn()-1).setValue("Pending"); //update the status of the sheet
    }
    
  }
  catch(err) {
    Logger.log(err);
  }
}


/**
  **function to notify the applicant that application is accepted
**/

function acceptNotification(email, name, appId) {
  //email = 'writetoraminderpal@gmail.com';
  //name = 'name';
  //appId = '222';
  
  var html = "<p>Dear "+name+","+"</p><p>We are happy to inform you that your iBoost application has been accepted."+
    "</p><p> Please follow these next steps to officially register your team. Ensure you read through the information carefully, ALSO PLEASE SECURELY SAVE YOUR iBoost Team ID. You will need this throughout your time at iBoost"+
    
      "</p><p>Your iBoost Team ID: "+appId+"</p><p>Please click <a href='www.iboostzone.com/onboarding/registration'>here</a> to register your team at iBoost.<br><br><br><br><p>Best,</p><br><img src='cid:logo' height='50' width='150'>";
  var template = HtmlService.createHtmlOutput(html).getContent();
  //Logger.log(template)
  MailApp.sendEmail(email, "iBoost Application Review","", {htmlBody: template, inlineImages:{logo: imageLogo}});
   
}

/**
  **function to notify iBoost that new team's been registered
**/

function newTeamNotify(com_id, comName){
  //com_id = '22';
  //comName = 'aaa';
  MailApp.sendEmail({to:"iboostzone@gmail.com,tngrant@ryerson.ca, jpsilva@ryerson.ca" , subject: "New Team Registered", htmlBody: "The Following Team has Registered: <br><br><br>  Company Name: " + comName + "<br><br> iBoost Team ID : " + com_id+"<br><br><br><br> Best, <br><br> <img src='cid:logo' height='50' width='150'>", inlineImages:{logo: imageLogo}});
}

/**
  **function to notify iBoost that new member's been registered
**/

function newMemNotify(com_id, comName, memName ){
 
  //com_id = "df25656";
  //comName = "abc";
  //memName = "mnk";
  var url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  //Logger.log(url);

 var h = "<p> A new member has registered under the following company:"+
 "</p><p>Company Name: " + comName + "</p><p> iBoost Team ID: " + com_id + "</p><p> Member Name: " + memName +
 "</p><p> Please click <a href="+url+"> Here </a> to accept them into iBoost.<br><br><br><br><p>Best,</p><br><img src='cid:logo' height='50' width='150'>";
 var temp = HtmlService.createHtmlOutput(h).getContent();
 MailApp.sendEmail("iboostzone@gmail.com, tngrant@ryerson.ca, jpsilva@ryerson.ca", "A New Member Has Registered","", {htmlBody: temp, inlineImages:{logo: imageLogo}});
}

/**
   *** send auto welcome Email to new Registered member 
**/
function welcomeMem(memName,memEmail){

//memName ="rami";
//memEmail ="writetoraminderpal@gmail.com";

var htm = "<p>Hi " + memName +",</p><p> You are now a registered member with iBoost!"+
"</p><p> Please complete the last step in the enrolment process <a href='www.iboostzone.com/onboarding/welcome'>HERE</a></p>"
+ "</p><p> If you have any question or concerns regarding this process, feel free to email us.</p><br><br><br><br><p>Best,</p><br><img src='cid:logo' height='50' width='150'>"
;
var templt = HtmlService.createHtmlOutput(htm).getContent();
MailApp.sendEmail(memEmail, "Welcome to iBoost","", {htmlBody:templt, inlineImages:{logo: imageLogo}});
}

function companyName(id) {
  //Logger.log("data");
  //id = "iB43214.8655";
   var sheet = spreadsheet.getSheetByName("2.Team General Information");
   var values = sheet.getRange(2, 4, sheet.getMaxRows(), 1).getValues();
  //Logger.log(values);
   var i, r, name;
  //startup name column
   var col = 8;
  
  for(i = 0; i < values.length; i++) {
    if(values[i] == id) {
      r = i + 2;
      break;
    }
  }
  name = sheet.getRange(r, col).getValue();
  //Logger.log(name);
  return name;
}

function logs() {
  DatabaseScript.keepLogs();
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
function test() {
  MailApp.sendEmail({
    to: "w.anushka@gmail.com",
    subject: "Logos",
    htmlBody: "inline Google Logo<img src='cid:logo'> images! <br>",
    inlineImages:
      {
        logo: imageLogo
      }
  });
}

