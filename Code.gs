/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */

//Globals
var input = 'ldap';
var types = ['redmine', 'Gitlab', 'LDAP'];
var typesColumns = ["A1:A","C1:C","E1:E"];
var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");

//Locate the last line occupied
function findLine(inp){
  var inputType = findColumn(inp);
  var sheet = SpreadsheetApp.getActiveSheet();
  //all values from collumn 
  var aVals = sheet.getRange(inputType).getValues();
  //number of used lines
  var last = aVals.filter(String).length;  
  var ui = SpreadsheetApp.getUi();
  //ui.alert(Alast); https://script.google.com/a/smdh.org/macros/d/MGMZ_wZpazxY3hT2SOBwFgi7RkztxW3XY/gwt/clear.cache.gif
  return last;
}

//Locate the right collumn
function findColumn(inp){
  //var out = typesColumns[types.indexOf(inp)];
  return typesColumns[types.indexOf(inp)];
}

//Always knows the right things
function writeTable(type, data, texto){
  var sheet = SpreadsheetApp.getActive();
  var lin = (findLine(type)+1).toString();
  //sheet.toast(col+a, type);
  var col1=findColumn(type)[0];
  var col2=alphabet[alphabet.indexOf(col1)+1];
  sheet.getRange(col1+lin).setValue(data);
  sheet.getRange(col2+lin).setValue(texto);
  sheet.getRange('w1').setValue('');
  var stop='';
  return;
}

function readRows() {
  var input = 'redmine';
  var types = ['redmine', 'gitlab', 'ldap'];
  var columns = ["A1:A","C1:C","E1:E"];
  var inputType = columns[types.indexOf(input)];
  var sheet = SpreadsheetApp.getActiveSheet();
  //all values from collumn 
  var Avals = sheet.getRange(inputType).getValues();
  //number of used lines
  var Alast = Avals.filter(String).length;
  var rows = sheet.getRange(3,1,1,2);
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  var ui = SpreadsheetApp.getUi();
  
  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
    Logger.log(row);
  }
  ui.alert(Alast);
};

//get attachments, get contend of attached files, returning a vector with email subject and attached files as text
function getAttachment(){
   // Logs information about any attachments in the first 100 inbox threads.
  var threads = GmailApp.getInboxThreads(0, 20);
  var msgs = GmailApp.getMessagesForThreads(threads);
  var out = [];  
  
  for (var i = 0 ; i < msgs.length; i++) {
    if (msgs[i][0].isUnread()){
      for (var j = 0; j < msgs[i].length; j++) {
        var attachments = msgs[i][j].getAttachments();
        for (var k = 0; k < attachments.length; k++) {
          Logger.log('Message "%s" contains the attachment "%s" (%s bytes)',
                     msgs[i][j].getSubject(), attachments[k].getName(), attachments[k].getDataAsString());
          //msgs[i][j].markRead();
          out.push(msgs[i][j].getSubject());
          out.push(attachments[k].getDataAsString());
        }
      }
    }
  }
  //var sheet = SpreadsheetApp.getActive();
  //sheet.toast('toast', out);
  return out;
}

function mailAttachParser(emails){
  var tes=[];
  var out ='';
  var aux=0;

  if (emails.search('Errors')!=-1){
      while (emails.search('Errors')!=-1)
      {
        tes.push(emails.substr(emails.search('Errors')+7,1));
        aux=aux+parseInt(emails.substr(emails.search('Errors')+7,1));
        emails=emails.substr(emails.search('Errors')+8,emails.length-emails.search('Errors')+8);
      }
      out=aux.toString();
    }
    else{
      out = 'Problem with backup';
    }
  return out;
}

function mailParser(){
  var emails = getAttachment();
  var body = [];
  var data = [];
  var type = [];
  var out = [data,type,body];
  
  for (var i = 0 ; i < emails.length; i+=2) {
    type.push(emails[i].split('-')[0]);
    data.push(emails[i].split('-')[1]);
    body.push(mailAttachParser(emails[i+1]));
  }
  return out;
}

function main(){
  var data='23/23';
  var texto='0errors';
  //writeTable(input,data,texto);
  toast ='';
}

function readMailSetTable(){
  var mailContents = mailParser();
  var body = '';
  var data = '';
  var type = '';
  for (var i = 0 ; i < mailContents[0].length; i++) {
    body = mailContents[2][i];
    data = mailContents[0][i];
    type = mailContents[1][i];
    writeTable(type,data,body);
    var s='';
  }
  
  
}

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Read Data ZZZZ",
    functionName : "readRows"
  },
   {
    name : "readMailSetTable",
    functionName : "readMailSetTable"
  }];
  spreadsheet.addMenu("Script Center Menu", entries);
};