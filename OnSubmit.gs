function emailtoAccounts(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary");
  
  var approvals = sheet.getRange("A2:A").getvalues();
  var emailsAccounting = sheet.getRange("B2:B").getValues();
  var emailsStaff = sheet.getRange("C2:C").getValues();
  
  
  

}

function onSubmit(){
  
  var allDetails; 
  var appEmail; //email address of the approver
  var reqEmail; //email address of the requester
  var poNum; //Purchase Order #
  var branch; //Branch detail of the purchase order
  var carDetail; //Car details of the purchase order
  var poDetail; //Financial details of the purchase order
  var poTotal; //GST inclusive total of the purchase order
  
  CreateNamedRange();
  
  SortFormResponses();
  
  UpdateRecordNumber();
  
  UpdatePONumber();
  
  allDetails = buildUrls();
  appEmail = allDetails[0];
  reqEmail = allDetails[1];
  poNum = allDetails[2];
  branch = allDetails[3];
  carDetail = allDetails[4];
  poDetail = allDetails[5];
  poTotal = allDetails[6];
  
  EmailAuthoriser(appEmail, reqEmail, poNum, branch, carDetail, poDetail, poTotal);
  
  
}


function SortFormResponses() {
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  var range = sheet.getRange("Responses");

   
  range.sort({column:1, ascending: false}); 
  
  var sourceRange = sheet.getRange("AH3");
  sourceRange.copyTo(sheet.getRange("AH2"));
  sourceRange = sheet.getRange("AI3");
  sourceRange.copyTo(sheet.getRange("AI2"));
  sourceRange = sheet.getRange("AJ3");
  sourceRange.copyTo(sheet.getRange("AJ2"));
  sourceRange = sheet.getRange("AE3");
  sourceRange.copyTo(sheet.getRange("AE2"));
  
}

function buildUrls() {
  
  var template = "https://docs.google.com/a/glv.co.nz/forms/d/e/1FAIpQLSdPiwHBAU9g0gqXycRw0br55ZvVHDcGsyAUFA01Uh6zS8XjsA/viewform?usp=pp_url&entry.381516357=##APPROVER##&entry.250270077=##REQUESTER##&entry.968813716=##PONUM##&entry.1921129323=##BRANCH##&entry.333459014=##CARDETAIL##&entry.1660368047=##PODETAIL##&entry.1262264700=##TOTAL##";
    
  var ss = SpreadsheetApp.getActive().getSheetByName("Data");  
  
  //Collect the Approver's Email from the "Data" sheet and add to URL
  var appEmail = ss.getRange("D2").getValue();
  var newUrl = template.replace('##APPROVER##', appEmail);
  
  //Collect the Requesting Staff's Email from the "Data" sheet and add to URL
  var reqEmail = ss.getRange("b2").getValue();
  newUrl = newUrl.replace('##REQUESTER##', reqEmail);
  
  //Collect the Purchase Order# from "Data" sheet and add to the URL  
  var poNum = ss.getRange("AF2").getValue();
  newUrl = newUrl.replace('##PONUM##', poNum);
  
  //Collect the Branch from "Data" sheet and add to the URL  
  var branch = ss.getRange("C2").getValue();
  var htmlBranch = branch.toString().replace(/\s/g, "+");
  newUrl = newUrl.replace('##BRANCH##', htmlBranch);
    
  //Collect the Purchase Order Details from the "Data" sheet and add to the URL
  var poDetail = ss.getRange("AH2").getValue();
  var htmlpoDetail = poDetail.toString().replace(/\s/g, "+");
  newUrl = newUrl.replace('##PODETAIL##', htmlpoDetail);
    
  //Collect the Purchase Order GST incl. Total
  var totalPO = ss.getRange("AI2").getValue();
  newUrl = newUrl.replace('##TOTAL##', totalPO);
   
  ss = SpreadsheetApp.getActive().getSheetByName("Summary");
  
  //Collect the Car Details from "Summary" sheet and add to the URL  
  var carDetail = ss.getRange("I2").getValue();
  var htmlcarDetail = carDetail.toString().replace(/\s/g, "+");
  newUrl = newUrl.replace('##CARDETAIL##', htmlcarDetail);
  
  
  //Enter the completed URL into the "URL" column ("AG") of the "Data" sheet
  ss = SpreadsheetApp.getActive().getSheetByName("Data");
  var formUrl = ss.getRange("AG2");
  formUrl.setValue(newUrl);
  
  
  
  return [appEmail, reqEmail, poNum, branch, carDetail, poDetail, totalPO];
  
}


function CreateNamedRange() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Data");
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
 
  ss.setNamedRange("Responses", sheet.getRange(2,1,lastRow-1, lastColumn));
  ss.setNamedRange("DataTopRow", sheet.getRange(2,1, 1,31));
  ss.setNamedRange("Approved", sheet.getRange(2,31,lastRow-1, 1));
  ss.setNamedRange("RecordNum", sheet.getRange(2,30,lastRow-1, 1));
  ss.setNamedRange("LastRecord", sheet.getRange(3,30));
  ss.setNamedRange("NewRecord", sheet.getRange(2,30));
  
  ss.setNamedRange("PONum", sheet.getRange(2,32,lastRow-1, 1));
  ss.setNamedRange("LastPO", sheet.getRange(3,32));
  ss.setNamedRange("NewPO", sheet.getRange(2,32));
  
  
  sheet = ss.getSheetByName("Summary");
  lastRow = ss.getLastRow();
  lastColumn = ss.getLastColumn();
  
  ss.setNamedRange("SummaryTopRow", sheet.getRange(2,1, lastRow-1, lastColumn));
  
  
}


function UpdateRecordNumber(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Data");
  var NewRecordRange = sheet.getRange("NewRecord");
  var LastRecordRange = sheet.getRange("LastRecord");
  var LastRecordNum = LastRecordRange.getValue();
  var NewRecordNum = LastRecordNum + 1;
  NewRecordRange.setValue(NewRecordNum);
  
}

function UpdatePONumber(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Data");
  var NewPORange = sheet.getRange("NewPO");
  var LastPORange = sheet.getRange("LastPO");
  var LastPONum = LastPORange.getValue();
  var NewPONum = LastPONum + 1;
  NewPORange.setValue(NewPONum);
  
}


function EmailAuthoriser(appEmail, reqEmail, poNum, branch, carDetail, poDetail, poTotal){
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  var subject = "Purchase Order Waiting - Click the link to approve";
  var toEmail = sheet.getRange("D2").getValue();
  var emailBody = 'Hi,<br/>'+'<br/>A new purchase order is waiting for your approval.<br/>'+'<br/>Please click the link below to approve or decline it.<br/><br/><br/>Thank you.<br/><br/><br/>';
  var podetailBody = 'The PO is from:  '+reqEmail+'<br/>The Purchase Order# is: '+poNum+'<br/>The Branch is:  '+branch+'<br/>Vehicle: '+carDetail+'<br/>The details are:  '+poDetail+'<br/>The GST inclusive total is : $'+poTotal+'<br/><br/>';
    
  var formUrl = sheet.getRange("AG2").getValue();
  var html = HtmlService.createHtmlOutput(formUrl).getContent();
  
  MailApp.sendEmail(
    toEmail,         // recipient
    subject,                  // subject 
    emailBody, {                        // body
      htmlBody: emailBody+podetailBody+html             // advanced options
    }
    
  ); 
  
  
}