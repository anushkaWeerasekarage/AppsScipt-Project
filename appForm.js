function onFormSubmit(event) {
  Logger.log('inside trigger');
  var formRes = event.response;
  //Logger.log(formRes);
  var desId = event.source.getDestinationId();
  var itemRes = formRes.getItemResponses();
  var i, email;
  var docCopy,fileId, doc, docBody;
  var pdf, url, sheet;
  var link = 'https://docs.google.com/spreadsheets/d/1XlLXLrQo1ls2iTAW-WrOBkumQXONEvgudixZkXk5cuc/edit?usp=sharing';
  var subject = "New Application Form";
  var body = "<p>You have recieved a new application form.</p><p> Please click <a href="+link+"> Here </a> to accept them into iBoost.<br><br><br><br></p>";
  var date = new Date(); 
  
  //DriveApp.getFileById("1DD7B0bB1iJcFfFMB-p8Z7VVWdzuW7DluS_9lrR6-ego").setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
  docCopy = DriveApp.getFileById("1DD7B0bB1iJcFfFMB-p8Z7VVWdzuW7DluS_9lrR6-ego").makeCopy("New Application Form");
  fileId = docCopy.getId();
  Logger.log(fileId);
  doc = DocumentApp.openById(fileId);
  docBody = doc.getBody();
  email = formRes.getRespondentEmail();
  Logger.log(email);
  
  for(i = 0; i < itemRes.length; i++) {
    //Logger.log("in for");
   //Logger.log(itemRes[i].getResponse());
    /*
    if(itemRes[i] == '') {
      itemRes[i] == "undefined";
    }
    */
  }
     ////all the MCQ s should be required(there should be an answer)
     docBody.replaceText("appTime", date);   
     docBody.replaceText("emailAdd", email);
     docBody.replaceText("fName", itemRes[0].getResponse());
     docBody.replaceText("lName", itemRes[1].getResponse());
     docBody.replaceText("affiliationStatus", itemRes[2].getResponse());
     docBody.replaceText("startupName", itemRes[3].getResponse());
     docBody.replaceText("websiteAdd", itemRes[4].getResponse());
     docBody.replaceText("fDate", itemRes[5].getResponse());
     docBody.replaceText("incoStatus", itemRes[6].getResponse());
     docBody.replaceText("dateOfInco", itemRes[7].getResponse());
     docBody.replaceText("ecoSec", itemRes[8].getResponse());
     docBody.replaceText("noOfEmp", itemRes[9].getResponse());
     docBody.replaceText("mRevenue", itemRes[10].getResponse());
     docBody.replaceText("endusers", itemRes[11].getResponse());
     docBody.replaceText("cusProb", itemRes[12].getResponse());
  
     docBody.replaceText("exSol", itemRes[13].getResponse());
     docBody.replaceText("targetCus", itemRes[14].getResponse());
     docBody.replaceText("endUser", itemRes[15].getResponse());
     docBody.replaceText("insightSoFar", itemRes[16].getResponse());
     docBody.replaceText("lookForEndUser", itemRes[17].getResponse());
     docBody.replaceText("dataCollected", itemRes[18].getResponse());
     docBody.replaceText("whoAreYou", itemRes[19].getResponse());
     docBody.replaceText("howProb", itemRes[20].getResponse());
     docBody.replaceText("whySolProb", itemRes[21].getResponse());
     docBody.replaceText("doneSoFar", itemRes[22].getResponse());
     docBody.replaceText("productDev", itemRes[23].getResponse());
     docBody.replaceText("firstVersion", itemRes[24].getResponse());
     docBody.replaceText("nextVersion", itemRes[25].getResponse());
     docBody.replaceText("whatTech", itemRes[26].getResponse());
     docBody.replaceText("prodDetails", itemRes[27].getResponse());
     docBody.replaceText("moreProd", itemRes[28].getResponse());
     docBody.replaceText("serviceExpect", itemRes[29].getResponse());
     docBody.replaceText("hotDesks", itemRes[30].getResponse());
     docBody.replaceText("howYouHear", itemRes[31].getResponse());
     docBody.replaceText("addAnything", itemRes[32].getResponse());
     
  
     doc.setName("Application Form of "+itemRes[0].getResponse()+" "+itemRes[1].getResponse());
     doc.saveAndClose();
  
     pdf = DriveApp.getFileById(fileId).getAs("application/pdf");
     MailApp.sendEmail({
     to: "tngrant@ryerson.ca, jpsilva@ryerson.ca",
      //to: "w.anushka@gmail.com",
     subject: subject,
     htmlBody: body, attachments: pdf
     });
     //MailApp.sendEmail("tngrant@ryerson.ca, jpsilva@ryerson.ca", subject, body, {html: body, attachments: pdf});
     sendPdfToDb(pdf);
     DriveApp.getFileById(fileId).setTrashed(true);
  
}
/*
function test() {
   var doc = DriveApp.getFileById("1gYYufQy8ewSQqCcGv1JNdCRcGj2FxN9Nvwkqmn17ZFM");
   var pdf = doc.getAs('application/pdf');
   var subject = "New Application Form";
   var body = "You have recieved a new application form";
   MailApp.sendEmail("w.anushka@gmail.com, writetoraminderpal@gmail.com", subject, body, {html: body, attachments: pdf});
   DriveApp.getFileById("1gYYufQy8ewSQqCcGv1JNdCRcGj2FxN9Nvwkqmn17ZFM").setTrashed(true);
  
}
*/

function sendPdfToDb(doc) {
  Logger.log('inside sendPdf');
  //var doc = DriveApp.getFileById('1TYoC0t_gZsfMS6S3ZeUxbsWvaxGuQfFdAjLxmxtkvDw').getAs('application/pdf');
  //var pdf = doc.getBlob().setContentType('application/pdf');
  var url = DriveApp.getFolderById('1MVbJai8JnczmLGGSUEx2-rtQDXSuHEP5').createFile(doc).getUrl();
  var sheet = SpreadsheetApp.openById('1XlLXLrQo1ls2iTAW-WrOBkumQXONEvgudixZkXk5cuc').getSheetByName('1.Application Forms');
  
  for(var i = 3; i < sheet.getMaxRows(); i++) {
    if(sheet.getRange(i,2).getValue() == "") {
      sheet.getRange(i, 2).setValue(url);
      break;
    }
  }
}

function test() {
  /*
    var docCopy = DriveApp.getFileById("1DD7B0bB1iJcFfFMB-p8Z7VVWdzuW7DluS_9lrR6-ego").makeCopy("New Application Form");
    Logger.log(docCopy.getId());
       MailApp.sendEmail("w.anushka@gmail.com", "subject", "body");

 */
     var pdf = DriveApp.getFileById('1Lrj57nUwLXEGC0iuA_mw_fVtc4yNI1dfmHBDIi8W-dI').getAs("application/pdf");
     var subject = "New Application Form";
     var link = 'https://docs.google.com/spreadsheets/d/1XlLXLrQo1ls2iTAW-WrOBkumQXONEvgudixZkXk5cuc/edit?usp=sharing';

     var body = "<p>You have recieved a new application form.</p><p> Please click <a href="+link+"> Here </a> to accept them into iBoost.<br><br><br><br></p>";

     MailApp.sendEmail({
     to: "writetoraminderpal@gmail.com",
     subject: subject,
     htmlBody: body, attachments: pdf
     });
     //MailApp.sendEmail("tngrant@ryerson.ca, jpsilva@ryerson.ca", subject, body, {html: body, attachments: pdf});
}