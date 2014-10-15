/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
 
//Globals
var input = 'ldap';
var types = ['redmine', 'gitlab', 'ldap'];
var typesColumns = ["A1:A","C1:C","E1:E"];
var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");


//Locate the last line occupied
function findLine(inp){
  var inputType = findColumn(inp);

  var sheet = SpreadsheetApp.getActiveSheet();
  //all values from collumn 
  var aVals = sheet.getRange(inputType).getValues();
  //number of used lines
  var last = Avals.filter(String).length;  

  var ui = SpreadsheetApp.getUi();
  //ui.alert(Alast); 
  return last;
}

//Locate the right collumn
function findColumn(inp){
  return typesColumns[types.indexOf(inp)];
}
 
//Always knows the right things
function writer(type, data, texto){
  var sheet = SpreadsheetApp.getActive();
  var lin = (findLine(type)+1).toString();
  //sheet.toast(col+a, type);
  var col1=findColumn(type)[0];
  var col2=alphabet[alphabet.indexOf(col1)+1];
  sheet.getRange(col1+lin).setValue(data);
  sheet.getRange(col2+lin).setValue(texto);
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
  
  //find collumn 
  //for (var i=0; i < types.length ; i++ ) {
  //  ui.alert(types[i]);  
  //}
  
};


function getattachment(){
   // Logs information about any attachments in the first 100 inbox threads.
  var threads = GmailApp.getInboxThreads(0, 5);
  var msgs = GmailApp.getMessagesForThreads(threads);
  for (var i = 0 ; i < msgs.length; i++) {
    if (msgs[i][0].isUnread()){
      for (var j = 0; j < msgs[i].length; j++) {
        var attachments = msgs[i][j].getAttachments();
        for (var k = 0; k < attachments.length; k++) {
          Logger.log('Message "%s" contains the attachment "%s" (%s bytes)',
                     msgs[i][j].getSubject(), attachments[k].getName(), attachments[k].getDataAsString());
          msgs[i][j].markRead();
        }
      }
    }
  }
  var ui = SpreadsheetApp.getUi();
  ui.alert(Logger.getLog());
  
}

function main(){
  data='23/23';
  texto='0errors';
  writer(input,data,texto);
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
  }];
  spreadsheet.addMenu("Script Center Menu", entries);
};
