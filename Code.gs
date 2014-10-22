 /*
  *Linux command:
  *mutt -s "Gitlab-$(date +%D)" logger@dominio -a nomelog.log < /dev/null
 */

//Globals
var input = 'ldap';
var types = ['Redmine', 'Gitlab', 'LDAP'];
var typesColumns = ["A1:A","C1:C","E1:E"];
var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
var adminMail = ['teste@smdh.org'];
var mailReport = 1; //0 not send - 1 send
var markMailAsRead = 1; //0 not mark - 1 to mark

bgSucess = '#87b798';
bgError = '#c04000';
bgData = '#fde7be';

//function to report backup problem to admin
function sendMail(subject,message){
  for (var i = 0; i < adminMail.length; i++) {
    MailApp.sendEmail(adminMail[i], subject, message);
  }
}

//Locate the last line occupied
function findLine(inp){
  var inputType = findColumn(inp);
  var sheet = SpreadsheetApp.getActiveSheet();
  //all values from collumn 
  var aVals = sheet.getRange(inputType).getValues();
  //number of used lines
  var last = aVals.filter(String).length;  
  return last;
}

//Locate the right collumn to start writing
function findColumn(inp){
  return typesColumns[types.indexOf(inp)];
}

//Always knows the right things
function writeTable(type, data, texto, note){
  var sheet = SpreadsheetApp.getActive();
  var lin = (findLine(type)+1).toString();
  //sheet.toast(col+a, type);
  var col1=findColumn(type)[0];
  var col2=alphabet[alphabet.indexOf(col1)+1];
  sheet.getRange(col1+lin).setValue(data);
  sheet.getRange(col2+lin).setValue(texto);
  sheet.getRange(col2+lin).setNote(note);
  sheet.getRange('w1').setValue('');
  
  sheet.getRange(col1+lin).setBackground(bgData);
  if (texto.search('0') != -1){
    sheet.getRange(col2+lin).setBackground(bgSucess);
  }
  else{
    sheet.getRange(col2+lin).setBackground(bgError);
    if (mailReport == 1){
      sendMail(type+' '+texto,note)
    }
  }
  return;
}

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
          if (markMailAsRead==1){
            msgs[i][j].markRead();
          }
          out.push(msgs[i][j].getSubject());
          out.push(attachments[k].getDataAsString());
        }
      }
    }
  }
  return out;
}

//this function parse the attachments, in order to discovery if some error happened during the backups
//bassically the function seeks for errors inside the logs, if the term "errors" is not present, the backup failed because conectivity 
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

/** This function just returns vectors eith the important parts of the emails
 * out = [data,type,body,note];
 */
function mailParser(){
  var emails = getAttachment();
  var body = [];
  var data = [];
  var type = [];
  var note = [];
  var out = [data,type,body,note];
  
  for (var i = 0 ; i < emails.length; i+=2) {
    type.push(emails[i].split('-')[0]);
    data.push(emails[i].split('-')[1]);
    body.push(mailAttachParser(emails[i+1]));
    note.push(emails[i+1]);
  }
  return out;
}

//this function will be called from the spreeadshet interface, mainly it will call other functions to write the values inside the table
function readMailSetTable(){
  var sheet = SpreadsheetApp.getActive();
  sheet.toast('Reading mails', 'Now');  
  var mailContents = mailParser();
  var body = '';
  var data = '';
  var type = '';
  var note = '';
  for (var i = 0 ; i < mailContents[0].length; i++) {
    body = mailContents[2][i];
    data = mailContents[0][i];
    type = mailContents[1][i];
    note = mailContents[3][i];
    writeTable(type,data,body,note);
    var s='';
  }
  sheet.toast('No more mails to read', 'Now');  
}

//this function will be called from the spreeadshet interface
function changeGreen(){
  var sheet = SpreadsheetApp.getActive();
  sheet.toast('Changing color to green', 'Now');    
  sheet.getActiveCell().setBackground(bgSucess);
}

//this function will be called from the spreeadshet interface
function changeRed(){
  var sheet = SpreadsheetApp.getActive();
  sheet.toast('Changing color to red', 'Now');    
  sheet.getActiveCell().setBackground(bgError);
}

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "readMailSetTable",
    functionName : "readMailSetTable"
  },
                 {
    name : "Change to green",
    functionName : "changeGreen"
  },
                {
    name : "Change to red",
    functionName : "changeRed"
  }];
  spreadsheet.addMenu("SMDH LOGGER", entries);
};