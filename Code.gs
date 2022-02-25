//global variables
var gradeLevel = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Select Module/Agenda").getRange("A1").getValue();
var currentModule = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Select Module/Agenda").getRange("h17").getValue();
var ss = SpreadsheetApp.getActiveSpreadsheet();
var agendaSs = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1wEKUp_k5LYTUfl161YurLs5J529e3JeZtJMzMh_wgsc/edit#gid=338644476");


// *** UI FUNCTIONS *** 
function onEdit(dd){
  var ws = ss.getActiveSheet();
  var range = ws.getRange("N6");
  var dd = range.getValue(); 

  if(dd === "Send Email Reminder"){
    showMessage1();
    range.clearContent();
  }else if(dd === "Send Email Summary to Admin"){
    showMessage2();
    range.clearContent();
  }else if(dd === "Save PLC Notes"){
    showDialog();
    range.clearContent();
  }else if(dd === "Retrieve Previous PLC Notes"){
    showDialog3();
    range.clearContent();
  }else if(dd === "Clear Current PLC Notes"){
    showDialog4();
    range.clearContent();
  }
}


function showMessage1(){
  var ui = SpreadsheetApp.getUi();
  var buttonPressed = ui.alert("Are you sure you want to send the email reminder?",ui.ButtonSet.YES_NO);
  if(buttonPressed === ui.Button.YES){
    sendEmails();
  }
}

function showMessage2(){
  var ui = SpreadsheetApp.getUi();
  var buttonPressed = ui.alert("Are you sure you want to send an email summary to admin?",ui.ButtonSet.YES_NO);
  if(buttonPressed === ui.Button.YES){
     sendAdminEmails();
  }
}

function showMessage3(){
  var ui = SpreadsheetApp.getUi();
  var buttonPressed = ui.alert("Would you like to save your PLC notes?",ui.ButtonSet.YES_NO);
  if(buttonPressed === ui.Button.YES){
    saveProgressMonitor();
  }
}

function showMessage4(){
  var ui = SpreadsheetApp.getUi();
  var buttonPressed = ui.alert("Are you sure you want to retrieve a previous PLC? Remember to make sure that you have selected the correct Lesson.",ui.ButtonSet.YES_NO);
  if(buttonPressed === ui.Button.YES){
    retrievePMData();
  }
}

function showMessage4post(){
  var ui = SpreadsheetApp.getUi();
  var buttonPressed = ui.alert("Are you sure you want to retrieve a previous PLC? Remember to make sure that you have selected the correct Lesson.",ui.ButtonSet.YES_NO);
  if(buttonPressed === ui.Button.YES){
    retrievePostData();
  }
}

  function showMessage5post(){
  var ui = SpreadsheetApp.getUi();
  var buttonPressed = ui.alert("You will be clearing all fields. Have you saved your PLC notes yet?",ui.ButtonSet.YES_NO_CANCEL);
  if(buttonPressed === ui.Button.YES){
     clearPostNotes();
      }else if(buttonPressed === ui.Button.NO){
      showMessage3();
    }
}

 function showMessage5(){
  var ui = SpreadsheetApp.getUi();
  var buttonPressed = ui.alert("You will be clearing all fields. Have you saved your PLC notes yet?",ui.ButtonSet.YES_NO_CANCEL);
  if(buttonPressed === ui.Button.YES){
     clearCurrentPLCNotes();
      }else if(buttonPressed === ui.Button.NO){
      showMessage3();
    }
}

function showMessage6(){
  var ui = SpreadsheetApp.getUi();
  var buttonPressed = ui.alert("Would you like to save your PLC notes for you post-assessment?",ui.ButtonSet.YES_NO_CANCEL);
  if(buttonPressed === ui.Button.YES){
    savePostNotes();
  }
}

function showDialog(){
  var html = HtmlService.createHtmlOutputFromFile('Page')
      .setWidth(300)
      .setHeight(300);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, 'Save PLC Notes');
}

function showDialog2(){
  var html = HtmlService.createHtmlOutputFromFile('Page2')
      .setWidth(300)
      .setHeight(300);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, 'Welcome to Washington Data Teams!');
}

function showDialog3(){
  var html = HtmlService.createHtmlOutputFromFile('Page3')
      .setWidth(300)
      .setHeight(300);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, 'Retreive PLC Notes');
}

function showDialog4(){
  var html = HtmlService.createHtmlOutputFromFile('Page4')
      .setWidth(300)
      .setHeight(300);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, 'Clear PLC Notes');
}





//unit map
 function mainMessage6(){
  var ui = SpreadsheetApp.getUi();
  var buttonPressed = ui.alert("Would you like to import the unit map for "+currentModule+ " ?",ui.ButtonSet.YES_NO);
  if(buttonPressed === ui.Button.YES){
    getUnitMap();
    Browser.msgBox("The unit map  for "+currentModule+" was imported");
    ss.getSheetByName("Unit Map").activate();
    }else if(buttonPressed === ui.Button.NO){
      Browser.msgBox("The unit map for "+currentModule+" was NOT imported");
    }
 }

//archive cfa data
function mainMessage7(){
  var ui = SpreadsheetApp.getUi();
    var buttonPressed = ui.alert("Would you like to archive the CFA data for "+currentModule+" ? Please know this process could take several minutes.",ui.ButtonSet.YES_NO_CANCEL);
    if(buttonPressed === ui.Button.YES){
    dumpData();
    Browser.msgBox("The CFA Data for "+currentModule+" was archived");
    ss.getSheetByName("CFA Data").activate();
    }else if(buttonPressed === ui.Button.NO){
      Browser.msgBox("The CFA data for "+currentModule+" was NOT archived");
    }
}

//retrieve CFA data
function mainMessage8(){
  var ui = SpreadsheetApp.getUi();
  var dset = ss.getSheetByName("Select Module/Agenda").getRange("E12").getValue();
  var buttonPressed = ui.alert("Would you like to retrieve CFA data for "+dset+" ? Remember that if you make any changes to this data you will need to archive it again.",ui.ButtonSet.YES_NO_CANCEL); 
      if(buttonPressed === ui.Button.YES){
        pullAndEdit();
        Browser.msgBox("The CFA data for "+currentModule+" was retrieved and is now available in the CFA data tab");
        ss.getSheetByName("CFA Data").activate();
      }else if(buttonPressed === ui.Button.NO){
      Browser.msgBox("The CFA data for "+currentModule+" was NOT retrieved");
    }
}

//clear CFA data
function mainMessage9(){
  var ui = SpreadsheetApp.getUi();
  var buttonPressed = ui.alert("You have selected to clear your current CFA data. Have you already archived your data for "+currentModule+" ?",ui.ButtonSet.YES_NO_CANCEL); 
        if(buttonPressed === ui.Button.YES){
          clearData();
          Browser.msgBox("CFA data for the current module was cleared");
          ss.getSheetByName("CFA Data").activate();
        }else if(buttonPressed === ui.Button.NO){
          mainMessage7();
        }

}

function mainMessage10(){
  var ui = SpreadsheetApp.getUi();
  var buttonPressed = ui.alert("Would you like to import the agenda for today?",ui.ButtonSet.YES_NO_CANCEL)
if(buttonPressed === ui.Button.YES){
          getAgenda();
            Browser.msgBox("Agenda was imported");
  }else if(buttonPressed === ui.Button.NO){
    Browser.msgBox("Agenda was NOT imported");
  }
}

function mainMessage11(){
  var ui = SpreadsheetApp.getUi();
  var buttonPressed = ui.alert("Would you like clear the current unit map?",ui.ButtonSet.YES_NO_CANCEL)
if(buttonPressed === ui.Button.YES){
           clearMess();
            Browser.msgBox("Current unit map was cleared");
            ss.getSheetByName("Unit Map").activate();
  }else if(buttonPressed === ui.Button.NO){
    Browser.msgBox("Unit map not cleared");
  }
}


function onEdit3(dd){
  var ws = ss.getSheetByName("Select Module/Agenda");
  var range = ws.getRange("E5");
  var dd= range.getValue();
 
 if(dd === "Import Unit Map"){
   mainMessage6();
   range.clearContent();
 }else if(dd === "Archive Data"){
   mainMessage7();
   range.clearContent();
   ss.getSheetByName("CFA Data").activate();
 }
 else if(dd === "Retrieve Data"){
   mainMessage8();
   range.clearContent();
   ss.getSheetByName("CFA Data").activate();
 }else if(dd === "Clear Data"){
   mainMessage9();
   range.clearContent();
 }else if(dd === "Import Agenda"){
   mainMessage10();
   range.clearContent();
 }else if(dd === "Create Assessment Agreements"){
   ss.getSheetByName("Assessment Agreements").activate();
   range.clearContent();
 }else if(dd === "Clear Unit Map"){
   mainMessage11();
   range.clearContent();
 }
}

// *** END UI FUNCTIONS ***




// *** UNIT MAP FUNCTIONS ***
function getUnitMap(){
  //Create search key using module name 
  var moduleName = ss.getSheetByName("Select Module/Agenda").getRange("j17").getValue();
  var searchFor = 'title contains "'+moduleName+' Unit Map"'; 
  Logger.log(searchFor);

  //Use search key to find file in google drive folder 
  var dApp = DriveApp;
  var files = dApp.searchFiles(searchFor);
  var fileIds = [];
 // var fileUrls = [];


 while(files.hasNext()){
  var file = files.next();
  var fileId = file.getId();
  var fileMime = file.getMimeType();
  fileIds.push(fileId);
}

Logger.log(fileIds);

//Write data from searched for file to original spreadsheet 
var srcSs = SpreadsheetApp.openById(fileIds[0]);
var srcSheet = srcSs.getSheetByName("Curriculum Map");
var srcValues = srcSheet.getRange(20,1,srcSheet.getLastRow(),srcSheet.getLastColumn()).getValues(); 
var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Unit Map");
    targetSheet.getRange(20,1,srcSheet.getLastRow(),srcSheet.getLastColumn()).setValues(srcValues); 
}

function clearMess() {
 ss.getSheetByName("Unit Map").getRange("A22:M30").clearContent();
}

// *** END UNIT MAP FUNCTIONS ***


//***  Get Agenda ***/

function getAgenda(){
  var gl = ss.getSheetByName("Select Module/Agenda").getRange("k13").getValue();
  var date = ss.getSheetByName("Select Module/Agenda").getRange("g5").getValue();
  var aName = gl+"Agenda";
  var agenda = agendaSs.getSheetByName(aName); 
  var dates = agenda.getRange(1,1,1,agenda.getLastColumn()).getValues()[0];
  var ind = dates.indexOf(date)+1;
  var text = agenda.getRange(1,ind,12).getValues();
  Logger.log(ind);
  Logger.log(text);
  ss.getSheetByName("Select Module/Agenda").getRange(17,5,12).setValues(text);
  
  
  
}

// *** CFA Standards Dropdown **** 

function onEdit2(){
var cfaWs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CFA Data");
var ddWs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Standards"); 
var ac = cfaWs.getActiveCell();
var row = ac.getRow();
var col = ac.getColumn();

if(row===4 && col>6 && cfaWs.getName()==="CFA Data"){
      ac.offset(1,0).clearContent().clearDataValidations();

      var domains = ddWs.getRange(1,5,1,8).getValues();
      var domainsIndex = domains[0].indexOf(ac.getValue())+5;
      Logger.log(domains);

      if(domainsIndex != 0){
        var valRange = ddWs.getRange(2,domainsIndex,ddWs.getLastRow()); 
        var valRule = SpreadsheetApp.newDataValidation().requireValueInRange(valRange).build();
        ac.offset(1,0).setDataValidation(valRule);
      }


    }


}



// *** EMAIL FUNCTIONS ***
/*function emailTrigger(){
  var ws = ss.getActiveSheet();
  var remindDate = ws.getRange("N7").getValue();
  var formattedDate = new Date(remindDate)
 // Logger.log(formattedDate);

 ScriptApp.newTrigger("sendEmails")
    .timeBased()
    .at(formattedDate)
    .create();
}
function cancelTimeTrigger(){
  
  var triggers = ScriptApp.getProjectTriggers();
  
  for(var i = 0; i < triggers.length; i++){
    if(triggers[i].getTriggerSource() == ScriptApp.TriggerSource.CLOCK){
      ScriptApp.deleteTrigger(triggers[i]);
    };
  };
}*/




function sendEmails(){

  var emails = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Import");
  var active = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = emails.getLastRow(); 

  var templateText = emails.getRange("F2").getValue(); 
  var staffNames = emails.getRange("D7").getValue();

  for(var i = 2; i <7;i++){
      if(emails.getRange(i,5).getValue()===true){
        var currentTitle = emails.getRange(i,1).getValue();
        var currentName = emails.getRange(i,2).getValue(); 
        var currentEmail = emails.getRange(i,3).getValue();
        var error = active.getRange("G36").getValue();
        var currentCis = active.getRange("A40").getValue();
        var nextCfa = active.getRange("I6").getValue();
        var nextCfaDate = active.getRange("K6").getDisplayValue();

      var messageBody = templateText.replace("{title}",currentTitle).replace("{name}",currentName).replace("{error}",error).replace("{cis}",currentCis).replace("{assessment2}",nextCfa).replace("{date}",nextCfaDate);

      var easyDate = new Date().toDateString();

      var subjectLine = "Common Instructional Strategy for week of "+ easyDate;
      Logger.log(messageBody);
      MailApp.sendEmail(currentEmail,subjectLine,messageBody);
      }
  }
 Browser.msgBox("An email was sent to "+staffNames);
}


//Email summary to admin
function sendAdminEmails(){

  var emails = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Import");
  var active = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lr = emails.getLastRow(); 
  var templateText = emails.getRange("G2").getValue(); 
  
  for(var i = 11; i <lr;i++){
      if(emails.getRange(i,5).getValue()===true){
          var currentTitle = emails.getRange(i,1).getValue();
          var currentName = emails.getRange(i,2).getValue(); 
          var currentEmail = emails.getRange(i,3).getValue();
          var error = active.getRange("G36").getValue();
          var currentCis = active.getRange("A40").getValue();
          var currentCfa = active.getRange("G2").getValue();
          var nextCfa = active.getRange("I6").getValue();
          var nextCfaDate = active.getRange("K6").getDisplayValue();
          var teacherNames = emails.getRange("D7").getValue();

          var messageBody = templateText
          .replace("{title}",currentTitle)
          .replace("{name}",currentName)
          .replace("{gradel}",gradeLevel)
          .replace("{module}",currentModule)
          .replace("{error}",error)
          .replace("{cis}",currentCis)
          .replace("{assessment1}",currentCfa)
          .replace("{assessment2}",nextCfa)
          .replace("{date}",nextCfaDate)
          .replace("{together1}",teacherNames);

          Logger.log(messageBody);

          var easyDate = new Date().toDateString();

          var subjectLine = gradeLevel+" Common Instructional Strategy for week of "+ easyDate;
        MailApp.sendEmail(currentEmail,subjectLine,messageBody);
      }//end if statement
  }//end for loop
    var adminNames = emails.getRange("D16").getValue();
   Browser.msgBox("An email was sent to "+adminNames);
}

/*
//email modular function
function emailTemplate(object,templateText){
    var messageBody = templateText;
    for(let[key,val]of Object.entries(object)){
        messageBody.replace("{"+key+"}",val);
    };
    return messageBody;
}*/
// *** END EMAIL FUNCTIONS ***





// *** DATABASE FUNCTIONS *** 
function onOpen(){
  showDialog2()

var ui = SpreadsheetApp.getUi();
/*var moreUi = ui.createMenu("Archive Functions");
  moreUi.addItem("Archive Data", "dumpData").addToUi();
  moreUi.addItem("Retrieve Data", "pullAndEdit").addToUi();
  moreUi.addItem("Retrieve Progress Monitor", "retrievePmData").addToUi();
  moreUi.addItem("Clear Data", "clearData").addToUi();*/
  
  
var moreUi2 = ui.createMenu("Test Functions");
  moreUi2.addItem("Generate Data", "randData").addToUi();
  moreUi2.addItem("Empty Archive", "emptyArchive").addToUi();
  moreUi2.addItem("Empty PM Archive", "emptyPMArchive").addToUi();
  moreUi2.addItem("Empty Post Archive", "emptyPostArchive").addToUi();

 
  
}

//Archive funcitons 
function dumpData() {
  
  //archive sheet data 
  var ss =  SpreadsheetApp.getActiveSpreadsheet();
  var entrySheet =ss.getSheetByName("CFA Data");
  var targetSheet = ss.getSheetByName("CFA Archive");
  var selectSheet = ss.getSheetByName("Select Module/Agenda"); 
  var lr = entrySheet.getLastRow();
  var lc = entrySheet.getLastColumn();
  var entryRange = entrySheet.getRange(4,1,lr,lc);
  var targetRange = targetSheet.getRange(1,1,lr,lc);
  entryRange.copyValuesToRange(targetSheet,1,lc,1,lr);
 // targetSheet.getRange("A1").moveTo(targetSheet.getRange("A1"));

   //create named range from sheet data
 var nrName = nameInputPrompt()
 ss.setNamedRange(nrName, targetRange);

 //create reference to named range in other tab 
 var setSheet = ss.getSheetByName("CFA Data Sets");
 var nameOutput = setSheet.getRange("C1");
 nameOutput.setValue(nrName);
 nameOutput.insertCells(SpreadsheetApp.Dimension.ROWS);
  
 // move archived data down to make room for more data 
  targetRange.insertCells(SpreadsheetApp.Dimension.ROWS);
  targetSheet.getRange(1,1,85,targetSheet.getLastColumn()).clearDataValidations();
  
  //clear active CFA data 
  entrySheet.getRange(7,7,lr,lc).clearContent(); 
 
}

function nameInputPrompt(){
   var ui = SpreadsheetApp.getUi(); 
  var input = ui.prompt("Please give your CFA data set or PLC a name. Remember to use underscores instead of spaces i.e. 'Module_1_CFA_Data or Lesson_1' instead of 'Module 1 CFA Data' or 'Lesson 1'").getResponseText();
  return input;
}

function pullAndEdit(){
   var ss =  SpreadsheetApp.getActiveSpreadsheet();
    var entrySheet =ss.getSheetByName("Select Module/Agenda");
    var dataSheet = ss.getSheetByName("CFA Data");
    var lr = dataSheet.getLastRow();
    var lc = dataSheet.getLastColumn();
    var nrName = entrySheet.getRange("E12").getValue(); 
    ss.getRangeByName(nrName).copyValuesToRange(dataSheet, 1, lc, 4, lr);
}


//Test functions
function randData(){
 var entrySheet = ss.getSheetByName("CFA Data");
 var lc = entrySheet.getRange("A1").getValue();
 var dataRange = entrySheet.getRange(7,7,106,lc);
 for (var x = 1; x <= dataRange.getWidth(); x++) {
    for (var y = 1; y <= dataRange.getHeight(); y++) {
      var number = Math.floor(Math.random() * 4) + 0;
      dataRange.getCell(y, x).setValue(number);
    }
  }

}

function emptyArchive(){
  var ws =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CFA Archive");
  var range = ws.getRange(1,1,ws.getLastRow(),ws.getLastColumn())
  range.clearContent();
  range.clearDataValidations();
} 

function emptyPostArchive(){
  var ss =  SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheetByName("Post Archive");
  var targetRange = targetSheet.getRange("1:1000");
  targetRange.clear();
  targetRange.clearDataValidations(); 

}

function emptyPMArchive(){
  var ss =  SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheetByName("PM Archive");
  var targetRange = targetSheet.getRange("1:1000");
  targetRange.clear();
  targetRange.clearDataValidations(); 

}




function clearData(){
  var range =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CFA Data").getRange("G7:Ay106");
  range.clearContent();
  range.clearDataValidations();
} 



//*** PLC ARCHIVE FUNCTIONALITY *** 

function saveProgressMonitor(){
  var ws1 = ss.getActiveSheet();
  var ws2 = ss.getSheetByName("PM Archive");

  var sourceRange = ws1.getRange("A18:N60"); 
  var targetRange = ws2.getRange("A1:N43");

  sourceRange.copyTo(targetRange);

  //create named range from sheet data
 var nrNameRange = ws1.getRange("G2");
 var nrName = nrNameRange.getValue();
 ss.setNamedRange(nrName, targetRange);

 //list of named ranges
 var setSheet = ss.getSheetByName("CFA Data Sets");
 var nameOutput = setSheet.getRange("E1");
 nameOutput.setValue(nrName);
 nameOutput.insertCells(SpreadsheetApp.Dimension.ROWS);

 // move down to make room for more data 
  targetRange.insertCells(SpreadsheetApp.Dimension.ROWS);
  
  //clear active CFA data 
  ws1.getRange("C18").clearContent();
  ws1.getRange("A20:L26").clearContent(); 
  ws1.getRange("A28:L34").clearContent();
  ws1.getRange("G36").clearContent();
  ws1.getRange("A40").clearContent();
  ws1.getRange("A46").clearContent();
  ws1.getRange("a49").clearContent();
  ws1.getRange("a52").clearContent();
  ws1.getRange("b55").clearContent();
  ws1.getRange("j55").clearContent();
  ws1.getRange("a56").clearContent();
  ws1.getRange("k56").clearContent();
  ws1.getRange("a60").clearContent();
  ws1.getRange("a50").clearContent();  

  ws1.getRange("m17:N37").clearContent();
  ws1.getRange("m39:N50").clearContent();
  ws1.getRange("m52:N60").clearContent();
  nrNameRange.clearContent();

  Browser.msgBox("Your PLC notes have been saved under the named range "+nrName);
}

function savePostNotes(){
  var ws1 = ss.getSheetByName("Post-Assessment");
  var ws2 = ss.getSheetByName("Post Archive");

  var sourceRange = ws1.getRange("A18:N71"); 
  var targetRange = ws2.getRange("A1:N55");

  sourceRange.copyTo(targetRange);

  //create named range from sheet data
 var nrNameRange = ws1.getRange("G2");
 var nrName = currentModule +"_"+nrNameRange.getValue();
 ss.setNamedRange(nrName, targetRange);

 //list of named ranges
 var setSheet = ss.getSheetByName("CFA Data Sets");
 var nameOutput = setSheet.getRange("G1");
 nameOutput.setValue(nrName);
 nameOutput.insertCells(SpreadsheetApp.Dimension.ROWS);

 // move down to make room for more data 
  targetRange.insertCells(SpreadsheetApp.Dimension.ROWS);
  
  //clear active CFA data 
  ws1.getRange("C18").clearContent();
  ws1.getRange("A20:L26").clearContent(); 
  ws1.getRange("A28:L34").clearContent();
  ws1.getRange("G36").clearContent();
  ws1.getRange("A40").clearContent();
  ws1.getRange("A46").clearContent();
  ws1.getRange("a49").clearContent();
  ws1.getRange("a52").clearContent();
  ws1.getRange("a53").clearContent();
  ws1.getRange("c53").clearContent(); 
  ws1.getRange("i53").clearContent(); 
  ws1.getRange("a61").clearContent();
  ws1.getRange("a63").clearContent();
  ws1.getRange("a65").clearContent();
  ws1.getRange("a67").clearContent();
  ws1.getRange("a69").clearContent(); 
  

  ws1.getRange("m18:N37").clearContent();
  ws1.getRange("m39:N71").clearContent();
  //nrNameRange.clearContent();

  Browser.msgBox("Your PLC notes have been saved under the named range "+nrName);
}


function clearCurrentPLCNotes(){
  //clear active CFA data 

  var ws1 = ss.getActiveSheet();

  ws1.getRange("C18").clearContent();
  ws1.getRange("A20:L26").clearContent(); 
  ws1.getRange("A28:L34").clearContent();
  ws1.getRange("G36").clearContent();
  ws1.getRange("A40").clearContent();
  ws1.getRange("A46").clearContent();
  ws1.getRange("a49").clearContent();
  ws1.getRange("a52").clearContent();
  ws1.getRange("b55").clearContent();
  ws1.getRange("j55").clearContent();
  ws1.getRange("a56").clearContent();
  ws1.getRange("k56").clearContent();
  ws1.getRange("a60").clearContent();
  ws1.getRange("a50").clearContent();
 // ws1.getRange("G2").clearContent();

  ws1.getRange("m17:N37").clearContent();
  ws1.getRange("m39:N50").clearContent();
  ws1.getRange("m52:N60").clearContent();
  nrNameRange.clearContent();
}

function clearPostNotes(){
  //clear active CFA data 

  var ws1 = ss.getActiveSheet();

  ws1.getRange("C18").clearContent();
  ws1.getRange("A20:L26").clearContent(); 
  ws1.getRange("A28:L34").clearContent();
  ws1.getRange("G36").clearContent();
  ws1.getRange("A40").clearContent();
  ws1.getRange("A46").clearContent();
  ws1.getRange("a49").clearContent();
  ws1.getRange("a52").clearContent();
  ws1.getRange("b55").clearContent();
  ws1.getRange("j55").clearContent();
  ws1.getRange("a56").clearContent();
  ws1.getRange("k56").clearContent();
  ws1.getRange("a60").clearContent();
  ws1.getRange("a50").clearContent();
  //ws1.getRange("G2").clearContent();

  ws1.getRange("c53").clearContent();
  ws1.getRange("i53").clearContent(); 

  ws1.getRange("a61").clearContent();
  ws1.getRange("a63").clearContent();
  ws1.getRange("a65").clearContent();
  ws1.getRange("a67").clearContent();
  ws1.getRange("a69").clearContent(); 
 

  ws1.getRange("m18:N37").clearContent();
  ws1.getRange("m39:N71").clearContent();
 // nrNameRange.clearContent();
}


function retrievePostData(){
    var active =ss.getActiveSheet();
    var nrName = active.getRange("L6").getValue();  
    ss.getRangeByName(nrName).copyValuesToRange(active,1,14,18,71);
    var splitString = nrName.toString().split("_");
    var cfaName = splitString[2]+"_"+splitString[3];
   // active.getRange("G2").setValue(cfaName);
    active.getRange("L6").clearContent();
    
 // Logger.log(cfaName);

}

function retrievePMData(){
    var active =ss.getActiveSheet();
    var nrName = active.getRange("L6").getValue();  
   ss.getRangeByName(nrName).copyValuesToRange(active,1,14,18,60);
    var splitString = nrName.toString().split("_")
    var cfaName = splitString[0]+"_"+splitString[1];
    active.getRange("G2").setValue(cfaName);
    active.getRange("L6").clearContent();
    
  Logger.log(cfaName);

}


function onEdit4(){
 var ws = ss.getSheetByName("Select Module/Agenda"); 
 var range = ws.getRange("G17"); 
 var dd = range.getValue(); 

 if( dd === "Step 1: Negative 4"){
   ss.getSheetByName("Step 1: Negative 4").activate();
 }else if (dd === "Step 1: Negative 6"){
   ss.getSheetByName("Step 1: Negative 6").activate();
 }else if(dd === "Hide Step 1 Checklists"){
   ss.getSheetByName("Step 1: Negative 4").hideSheet();
   ss.getSheetByName("Step 1: Negative 6").hideSheet();
 }

 range.clearContent();
}

// *** More Ui *** 

function activateSelect(){
  ss.getSheetByName("Select Module/Agenda").activate();
  ss.toast("Select which module you will be cycling around, manage saved data, and import your unit map and agenda","Select Module/Agenda",8);
}

function activateCFA(){
  ss.getSheetByName("CFA Data").activate();
  ss.toast("Enter your CFA data on this tab","CFA Data",8);
}

function activatePre(){
  ss.getSheetByName("Pre-Assessment").activate();
  ss.toast("Cycle around your pre-assessment for this module","Pre-Assessment",8);
}

function activatePm(){
  ss.getSheetByName("Progress Monitor").activate();
  ss.toast("Cycle around a recent progress monitor of your choice","Progress Monitor",8);
}

function activatePost(){
  ss.getSheetByName("Post-Assessment").activate();
  ss.toast("Cycle around the post-assessment for this module","Post-Assessment",8);
}

function activateAnalytics(){
  ss.getSheetByName("Analytics").activate();
  ss.toast("See visualizations of data for an assessment of your choice","Analytics",8);
}


