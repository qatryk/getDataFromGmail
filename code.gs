//get data from gmail v.0.1

var SHEET_NAME = "SHEET_NAME";
var QUERY="label:YOURLABEL label:unread";

function getEmails(q){
  var emails = [];
  var thds = GmailApp.search(q);
  for(var i in thds){
    var msgs = thds[i].getMessages();
    for(var j in msgs){
      emails.push([msgs[j].getBody().replace(/<.*?>/g, '\n').replace(/^\s*\n/gm, '').replace(/^\s*/gm, '').replace(/\s*\n/gm, '\n')]);
    }
  }
  GmailApp.markThreadsRead(thds);
  return emails;
}

function appendData(sheet, email_array){
  for (var i in email_array){
    var line = email_array[i][0].split(/[\r\n]+/g);
    var row=sheet.getLastRow() + 1;
    var empty=0;
    for (var k=1;empty==0;){
      if (sheet.getRange(1, k).getValue()==""){
        empty=1;}
      else{
        var done=0;
        for (var j =0; j <line.length; j++){
          if (sheet.getRange(1, k).getValue()==line[j].replace(":","") && done==0){
            sheet.getRange(row, k).setValue(line[j+1]);
            done=1;
          }
        }
        k++;
      } 
    }
  }
}
  
function run(){
  var email_array = getEmails(QUERY);
  if(email_array) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if(!sheet) sheet = ss.insertSheet(SHEET_NAME);
    appendData(sheet, email_array);
  }
}

function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Get Data from Gmail",
    functionName : "run"
  }];
  sheet.addMenu("Get Data", entries);
};
