function setup() {
  createSettingsSheet();
  createOtherCasesSheet();
  createIgnoreCasesSheet();
  createTriggers();
  initializeApp();
  msg = "app initialized";
  myLogger(msg);
  SpreadsheetApp.getActive().setActiveSheet(ss.getSheetByName('settings'));
}



function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(' ☑ Docket Monitor')
  .addItem('setup', 'setup')
  .addItem('update case list', 'getAttyCases')
  .addItem('run docket monitor', 'dmProcessLock')
  .addItem('reset', 'resetDM')
  .addItem('help', 'help')
  .addToUi();
}



function help() {  
  try {
    var content = UrlFetchApp.fetch('https://www.dynamicpractices.com/apps/docketMonitor/help.html', fetchOptions);
    var ss = SpreadsheetApp.getActive();
    var html = HtmlService.createHtmlOutput(String(content))
    .setTitle("help")
    .setWidth(400)
    .setHeight(400);
    ss.show(html);
  } catch (err) { }
}


function createTriggers() {
  //remove all existing triggers
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
  ScriptApp.newTrigger("onOpen").forSpreadsheet(SpreadsheetApp.getActive()).onOpen().create();
  ScriptApp.newTrigger("getAttyCases").timeBased().atHour(8).everyDays(1).create();
  ScriptApp.newTrigger("dmProcessLock").timeBased().atHour(10).everyDays(1).create();
  ScriptApp.newTrigger("dmProcessLock").timeBased().atHour(14).everyDays(1).create();
}


function createSettingsSheet() {
  try {
    var ss = SpreadsheetApp.getActive();
    
    var file = DriveApp.getFileById(ss.getId());
    var fileName = file.getName().toLowerCase();
    if (fileName.indexOf('untitled') >= 0) { file.setName('Docket Monitor'); }
    
    var sheet = ss.getSheetByName( 'settings' );
    if (sheet === null) {
      sheet = ss.getSheetByName( 'Sheet1' );
      if (sheet === null) {
        sheet = ss.insertSheet();
        Utilities.sleep(500);
        sheet.setName( 'settings' );
        Utilities.sleep(500);
      } else {
        Utilities.sleep(500);
        sheet.setName( 'settings' );
        Utilities.sleep(500);
      }
    } else {
      var sheetDup = sheet.copyTo(ss);
      Utilities.sleep(500);
      sheetDup.setName( 'settings' + ' as on ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM-dd-yyyy 'at' h:mm a") );
      Utilities.sleep(500);
    }
    
    sheet.clear();
    Utilities.sleep(500);
    sheet.setName( 'settings' );
    Utilities.sleep(500);
    
    
    // set the template values
    var labels = [];
    labels.push(["Attorney name","",'(see "help" from the Docket Monitor menu)']);
    labels.push(["Attorney email","",""]);
    labels.push(["Attorney bar number","", ""]);
    labels.push(["Exclude closed cases","FALSE",""]);
    
    sheet.getRange(1,1,(labels.length),(labels[0].length)).setValues(labels);
    
    sheet.getRange('C1:C4').mergeVertically();
    sheet.getRange('C1:C4').setVerticalAlignment("middle");
    
    
    // delete empty rows and columns
    var maxRows = sheet.getMaxRows();
    var lastRow = sheet.getLastRow();
    if ( (maxRows !== 1) && (maxRows !== lastRow) ) {
      sheet.deleteRows(lastRow+1, maxRows-lastRow);
      Utilities.sleep(500);
    }
    
    var maxCols = sheet.getMaxColumns();
    var lastCol = sheet.getLastColumn();
    if ( (maxCols !== 1) && (maxCols !== lastCol) ) {
      sheet.deleteColumns(lastCol+1, maxCols-lastCol);
      Utilities.sleep(500);
    }
    
    
    // set column width for each column
    sheet.setColumnWidth(1, 195);
    sheet.setColumnWidth(2, 475);
    sheet.setColumnWidth(3, 300);
    
    
    // set background and font colors
    var rangeRef, range;
    
    rangeRef = 'A1:A4';
    range = sheet.getRange(rangeRef);
    range.setFontColor("white");
    range.setBackground("black");
    range.setFontSize(11);
    
    
    rangeRef = 'B1:B4';
    range = sheet.getRange(rangeRef);
    range.setBackground('#d9edf7');
    range.setFontSize(11);
    range.setHorizontalAlignment("center");
    
    
    rangeRef = 'C1:C4';
    range = sheet.getRange(rangeRef);
    range.setBackgroundRGB(221,221,221);
    range.setFontSize(11);
    range.setFontStyle("italic");
    range.setHorizontalAlignment("center");
    
    
    rangeRef = 'B1:B1';
    range = sheet.getRange(rangeRef);
    range.setBackground('#dff0d8');
    
    
    // data validation
    rangeRef = 'B4:B4';
    range = sheet.getRange(rangeRef);
    var vals = ["FALSE", "TRUE"];
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(vals, true).build();
    range.setDataValidation(rule);
    
    
  } catch (err) {
    msg = "Error: " + err.message + "\n";
    msg += "Script: " + err.fileName + "\n";
    msg += "Line: " + err.lineNumber;
    myLogger(msg);
  }
}



function createOtherCasesSheet() {
  
  try {
    
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName( 'other cases' );
    if (sheet === null) {
      sheet = ss.insertSheet();
      Utilities.sleep(500);
      sheet.setName( 'other cases' );
      Utilities.sleep(500);
    } else { 
      sheet.clear();
      Utilities.sleep(500);
    }    
    
    // set the template values
    var labels = [];
    labels.push(["Case number"]);
    
    sheet.getRange(1,1,(labels.length),(labels[0].length)).setValues(labels);    
    
    // delete empty rows and columns
    var maxRows = sheet.getMaxRows();
    var lastRow = sheet.getLastRow();
    if ( (maxRows !== 1) && (maxRows !== lastRow) ) {
      sheet.deleteRows(lastRow+19, maxRows-(lastRow+19));
      Utilities.sleep(500);
    } else {
      sheet.deleteRows(10, maxRows-10);
      Utilities.sleep(500);
    }
    
    var maxCols = sheet.getMaxColumns();
    var lastCol = sheet.getLastColumn();
    sheet.deleteColumns(lastCol+1, maxCols-lastCol);
    
    // set column width
    sheet.setColumnWidth(1, 200);
    
    // set background and font colors
    var rangeRef, range;
    
    rangeRef = 'A1';
    range = sheet.getRange(rangeRef);
    range.setFontColor("white");
    range.setBackground("black");
    range.setFontSize(11);
    range.setHorizontalAlignment("center");    
    
  } catch (err) {
    msg = "Error: " + err.message + "\n";
    msg += "Script: " + err.fileName + "\n";
    msg += "Line: " + err.lineNumber;
    myLogger(msg);
  }
  
}



function createIgnoreCasesSheet() {
  
  try {
    
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName( 'ignore cases' );
    if (sheet === null) {
      sheet = ss.insertSheet();
      Utilities.sleep(500);
      sheet.setName( 'ignore cases' );
      Utilities.sleep(500);
    } else { 
      sheet.clear();
      Utilities.sleep(500);
    }    
    
    // set the template values
    var labels = [];
    labels.push(["Case number"]);
    
    sheet.getRange(1,1,(labels.length),(labels[0].length)).setValues(labels);    
    
    // delete empty rows and columns
    var maxRows = sheet.getMaxRows();
    var lastRow = sheet.getLastRow();
    if ( (maxRows !== 1) && (maxRows !== lastRow) ) {
      sheet.deleteRows(lastRow+19, maxRows-(lastRow+19));
      Utilities.sleep(500);
    } else {
      sheet.deleteRows(10, maxRows-10);
      Utilities.sleep(500);
    }
    
    var maxCols = sheet.getMaxColumns();
    var lastCol = sheet.getLastColumn();
    sheet.deleteColumns(lastCol+1, maxCols-lastCol);
    
    // set column width
    sheet.setColumnWidth(1, 200);
    
    // set background and font colors
    var rangeRef, range;
    
    rangeRef = 'A1';
    range = sheet.getRange(rangeRef);
    range.setFontColor("white");
    range.setBackground("black");
    range.setFontSize(11);
    range.setHorizontalAlignment("center");    
    
  } catch (err) {
    msg = "Error: " + err.message + "\n";
    msg += "Script: " + err.fileName + "\n";
    msg += "Line: " + err.lineNumber;
    myLogger(msg);
  }
  
}



function initializeApp() {
  
  try {
    
    var ss = SpreadsheetApp.getActive();
    var appSettings = ss.getSheetByName("appSettings");
    if (appSettings === null) {
      appSettings = ss.insertSheet();
      Utilities.sleep(500);
      appSettings.hideSheet();
      Utilities.sleep(500);
      appSettings.setName("appSettings");
      Utilities.sleep(500);
      appSettings.appendRow(["send_updates", "TRUE"]);
      Utilities.sleep(500);
      appSettings.appendRow(["new_case_updates", "TRUE"]);
      Utilities.sleep(500);
      appSettings.appendRow(["email_error_reports", "FALSE"]);
      Utilities.sleep(500);
    }
    var appSettingsValues = appSettings.getDataRange().getValues();
    
    var App, AppFolder, AppFolderId, AppSubFolderName, AppSubFolder, AppSubFolderId;
    
    App = 'Docket Monitor';
    AppSubFolderName = 'Docket Monitor case files';
    
    var prop, value;
    // loop through settings
    for (var p = 0; p < appSettingsValues.length; p++) {
      prop = appSettingsValues[p][0];
      value = appSettingsValues[p][1];
      
      // determine if settings are already defined
      if ( prop.indexOf('AppFolderId') >= 0 ) {
        AppFolderId = value;
        AppFolder = DriveApp.getFolderById(AppFolderId);
        continue;
      }
      if ( prop.indexOf('AppSubFolderId') >= 0 ) {
        AppSubFolderId = value;
        AppSubFolder = DriveApp.getFolderById(AppSubFolderId);
        continue;
      }
    }
    
    if ( !AppFolderId ) {
      
      AppFolder = DriveApp.getRootFolder().createFolder(App);
      AppFolderId = AppFolder.getId();
      appSettings.appendRow(["AppFolderId", AppFolderId]);
      AppFolder = DriveApp.getFolderById(AppFolderId);
      var ssId = ss.getId();
      var file = DriveApp.getFileById(ssId);
      // add spreadsheet to app folder
      AppFolder.addFile(file);
      // remove spreadsheet from root folder
      DriveApp.getRootFolder().removeFile(file);
      
    }
    
    if ( !AppSubFolderId ) {
      
      AppSubFolder = AppFolder.createFolder(AppSubFolderName);
      AppSubFolderId = AppSubFolder.getId();
      appSettings.appendRow(["AppSubFolderId", AppSubFolderId]);
      AppSubFolder = DriveApp.getFolderById(AppSubFolderId);
      
    }  
    
  } catch (err) {
    msg = "Error: " + err.message + "\n";
    msg += "Script: " + err.fileName + "\n";
    msg += "Line: " + err.lineNumber;
    myLogger(msg);
  }
    
}


function getSettings(property) {
  var ss = SpreadsheetApp.getActive();
  var appSettings, appSettingsValues;
  appSettings = ss.getSheetByName("appSettings");
  if (appSettings === null) {
    appSettings = ss.insertSheet();
    Utilities.sleep(500);
    appSettings.hideSheet();
    Utilities.sleep(500);
    appSettings.setName("appSettings");
    Utilities.sleep(500);
  }
  
  appSettingsValues = appSettings.getDataRange().getValues();
  
  var prop, value;
  for (var p = 0; p < appSettingsValues.length; p++) {
    prop = appSettingsValues[p][0];
    value = appSettingsValues[p][1];
    if ( ( property.indexOf('App') >= 0 ) || ( property.indexOf('fol') >= 0 ) ) {
      if ( prop.indexOf(property) >= 0 ) { return value; }
    } else {
      if ( prop === property ) { return value; }
    }
  }
  return false;
}

function removeSettings(property) {
  var ss = SpreadsheetApp.getActive();
  var appSettings = ss.getSheetByName("appSettings");
  if (appSettings === null) {
    appSettings = ss.insertSheet();
    Utilities.sleep(500);
    appSettings.hideSheet();
    Utilities.sleep(500);
    appSettings.setName("appSettings");
    Utilities.sleep(500);
  }
  var appSettingsValues, appSettingsValuesString;
  var cache = CacheService.getUserCache();
  appSettingsValuesString = cache.get("appSettingsValues");
  if ( typeof appSettingsValuesString === "undefined" || !appSettingsValuesString ) {
    appSettingsValues = appSettings.getDataRange().getValues();
  } else {
    appSettingsValues = JSON.parse(appSettingsValuesString);
  }
  
  var prop, value;
  for ( var p = (appSettingsValues.length - 1); p >= 0; --p ) {          
    prop = appSettingsValues[p][0];
    value = appSettingsValues[p][1];
    if (appSettingsValues[p][0].indexOf(property) > -1) {
      appSettingsValues.splice(p, 1);
      break;
    }
  }
  
  appSettingsValuesString = JSON.stringify(appSettingsValues);
  try {
    cache.put("appSettingsValues", appSettingsValuesString, (5 * 60));
  } catch (err) {
    msg = "\nMessage: " + err.message;
    msg += "\nScript: " + err.fileName;
    msg += "\nLine: " + err.lineNumber;
    myLogger(msg);
  }
  appSettings.clearContents();
  var numRows = appSettingsValues.length;
  var numCols = appSettingsValues[0].length;
  var newRange = appSettings.getRange(1, 1, numRows, numCols);
  newRange.setValues(appSettingsValues);
  
  return false;
}


function removeInvalidSheets() {
  
  try {
    
    var script_name = 'removeInvalidSheets';
    msg = script_name + ' running';
    myLogger(msg);
    
    var ss = SpreadsheetApp.getActive();
    var sheet, sheetName;
    
    var allSheets = ss.getSheets();
    
    for ( var sheetNum = (allSheets.length - 1); sheetNum >= 0; --sheetNum ) {  
      
      sheet = allSheets[sheetNum];
      sheetName = sheet.getName();
      
      if (sheetName === "log") { continue; }
      if (sheetName === "other cases") { continue; }
      if (sheetName === "ignore cases") { continue; }
      if (sheet.isSheetHidden()) { continue; }
      if ( sheetName.indexOf("etting") >= 0 ) { continue; }
      
      if ( sheetName.indexOf("folUp") >= 0 ) {
        ss.deleteSheet(ss.getSheetByName(sheetName));
        Utilities.sleep(500);
        msg = "sheet deleted: " + sheetName;
        myLogger(msg);
      }
      
    }
    
  } catch (err) {
    msg = "\nMessage: " + err.message;
    msg += "\nScript: " + err.fileName;
    msg += "\nLine: " + err.lineNumber;
    myLogger(msg);
  }
  
}


function resetDM() {
  
  //remove fol up sheet
  removeInvalidSheets();
  
  
  //remove all existing triggers
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
  msg = "triggers deleted";
  myLogger(msg);
  createTriggers();
  
  
  //remove fol up prop
  var ss = SpreadsheetApp.getActive();
  var appSettings = ss.getSheetByName("appSettings");
  var appSettingsValues = appSettings.getDataRange().getValues();
  
  for ( var p = (appSettingsValues.length - 1); p >= 0; --p ) {
    if (appSettingsValues[p][0].indexOf('fol') > -1) {
      appSettingsValues.splice(p, 1);
      break;
    }
  }
  
  appSettings.clearContents();
  var numRows = appSettingsValues.length;
  var numCols = appSettingsValues[0].length;
  var newRange = appSettings.getRange(1, 1, numRows, numCols);
  newRange.setValues(appSettingsValues);

}


var ss = SpreadsheetApp.getActive();
var sheet, sheetName, numRows, numCols, dataRange, values, r, k;

var trigger_delay_mins = 5;

var max_running_time_mins = 5.1;
var max_running_time = 1000 * 60 * max_running_time_mins;
var timeLimitIsNear, currTime;

var fetchOptions =
    {
      muteHttpExceptions: true,
      validateHttpsCertificates: false,
      followRedirects: true
    };

var caseCount, caseAddedCount, caseRemCount, caseUpdateCount, caseAttachmentCount;
caseCount = caseAddedCount = caseRemCount = caseUpdateCount = caseAttachmentCount = 0;

var subject, msg, body, logText;



function getAttyCases() {
  
  try {
    
    var script_start = (new Date()).getTime();
    var script_name = 'getAttyCases';
    msg = script_name + ' running';
    myLogger(msg);
    
    var ss = SpreadsheetApp.getActive();
    var appSettings = ss.getSheetByName("appSettings");
    if (appSettings === null) {
      appSettings = ss.insertSheet();
      Utilities.sleep(500);
      appSettings.hideSheet();
      Utilities.sleep(500);
      appSettings.setName("appSettings");
      Utilities.sleep(500);
    }
    
    var settings = getColumnsData_(ss.getSheetByName( 'settings' ))[0];
    var atty = settings.attorneyName;
    var attyEmail = settings.attorneyEmail;
    var barNo = settings.attorneyBarNumber;
    barNo = 'AR' + String(barNo);
    barNo = String(barNo).replace('ARAR', 'AR');
    barNo = String(barNo).replace('AR19', 'AR');
    
    var reEmail = /^([\w-]+(?:\.[\w-]+)*)@((?:[\w-]+\.)*\w[\w-]{0,66})\.([a-z]{2,6}(?:\.[a-z]{2})?)/i;  
    var reBarNo = /^(AR)([0-9]{3,8})$/i;
    if ( (typeof atty === "undefined") || !atty || (atty === 'AR') || (!reEmail.test(attyEmail)) || (!reBarNo.test(barNo)) ) {
      help();
      return;
    }
    
    //loop through every sheet and remove extraneous sheets
    removeInvalidSheets();
    
    // delete any follow up remaining from previous day
    removeSettings('follow_up_cases');
    
    // generate case list
    processAtty(barNo, atty, attyEmail);
    
    ss.setActiveSheet(ss.getSheetByName('settings'));
    
    var caseCount, caseAddedCount, caseRemCount, caseUpdateCount, caseAttachmentCount;
    caseCount = caseAddedCount = caseRemCount = caseUpdateCount = caseAttachmentCount = 0;
    wrapUp(script_name, script_start, caseCount, caseAddedCount, caseRemCount, caseUpdateCount, caseAttachmentCount);
    
  } catch (err) {
    msg = "\nMessage: " + err.message;
    msg += "\nScript: " + err.fileName;
    msg += "\nLine: " + err.lineNumber;
    myLogger(msg);
  }
  
}


function dmProcessLock() {
  
  try {
    msg = 'dmProcessLock running';
    myLogger(msg);
    
    var lock = LockService.getPublicLock();
    
    if (lock.tryLock(max_running_time-(1000*60))) {
      
      msg = "lock acquisition";
      myLogger(msg);
      
      try {
        dmPrimary();
      } catch (err) {
        msg = "Error: " + err.message + "\n";
        msg += "Script: " + err.fileName + "\n";
        msg += "Line: " + err.lineNumber;
        myLogger(msg);
      }
      
    } else {
      
      msg = "error: lock acquisition fail (queueProcess locked)";
      myLogger(msg);
      
      //create a trigger
      currTime = (new Date()).getTime();
      var waitTime = (1000 * 60 * trigger_delay_mins);
      ScriptApp.newTrigger("dmProcessLock").timeBased().at((new Date(currTime + waitTime))).create();
      
      msg = "new trigger scheduled for " + new Date(currTime + waitTime);
      myLogger(msg);
      
    }
    
    lock.releaseLock();
    msg = "lock released" + "\n" + ":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::";
    msg += ":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::";
    myLogger(msg);
    
  } catch (err) {
    msg = "Error: " + err.message + "\n";
    msg += "Script: " + err.fileName + "\n";
    msg += "Line: " + err.lineNumber;
    myLogger(msg);
  }  
}



function dmPrimary() {
  
  try {
    
    var script_start = (new Date()).getTime();
    var script_name = 'dmPrimary';
    msg = script_name + ' running';
    myLogger(msg);
    
    var ss = SpreadsheetApp.getActive();
    var settings = getColumnsData_(ss.getSheetByName( 'settings' ))[0];
    var adminEmail = settings.attorneyEmail;
    
    var atty = settings.attorneyName;
    var attyEmail = settings.attorneyEmail;
    var barNo = settings.attorneyBarNumber;
    barNo = 'AR' + String(barNo);
    barNo = String(barNo).replace('ARAR', 'AR');
    barNo = String(barNo).replace('AR19', 'AR');
    
    var reEmail = /^([\w-]+(?:\.[\w-]+)*)@((?:[\w-]+\.)*\w[\w-]{0,66})\.([a-z]{2,6}(?:\.[a-z]{2})?)/i;  
    var reBarNo = /^(AR)([0-9]{3,8})$/i;
    if ( (typeof atty === "undefined") || !atty || (atty === 'AR') || (!reEmail.test(attyEmail)) || (!reBarNo.test(barNo)) ) {
      help();
      return;
    }
    
    // stop the script on Saturday and Sunday
    if (((new Date().getDay()) === 6) || ((new Date().getDay()) === 0)) return;
    
    var appUrl = SpreadsheetApp.getActive().getUrl();
    
    var caseNo, rCheck;
    var allSheets = ss.getSheets();
    var sheet, case_list_sheet;
    var case_list = [];
    var case_list_remaining = [];
    var follow_up_prop, follow_up_info, follow_up_sheet_name, follow_up_trigger_id, values;
    
    var send_updates = getSettings("send_updates");
    if ( send_updates === false ) {
      var subject = "docket montior case update emails are disabled";
      var body = "docket montior case update emails are disabled!";
      MailApp.sendEmail({name: 'Docket Monitor', to: adminEmail, subject: subject, body: body + "\nSheet:\n" + appUrl});
    }
    
    //determine if follow up is necessary
    follow_up_prop = getSettings('follow_up_cases');
    if ( follow_up_prop ) {
      
      script_name = 'dmPrimary (follow up run)';
      msg = script_name;
      myLogger(msg);
      
      follow_up_info = follow_up_prop.split('|');
      follow_up_sheet_name = String(follow_up_info[ 0 ]);
      follow_up_trigger_id = String(follow_up_info[ 1 ]);
      
      if ( !follow_up_sheet_name || !follow_up_trigger_id ) { 
        msg = "follow up setting error: setting is missing\n";
        msg += "follow_up_sheet_name: " + follow_up_sheet_name + "\n";
        msg += "follow_up_trigger_id: " + follow_up_trigger_id + "\n";
        myLogger(msg);
        return false;
      }
      
      case_list_sheet = ss.getSheetByName(follow_up_sheet_name);
      if ( case_list_sheet === null ) { 
        msg = "follow up setting error: case_list_sheet is missing\n";
        msg += "follow_up_sheet_name: " + follow_up_sheet_name + "\n";
        myLogger(msg);
        return false;
      }
      
      msg = "current follow up setting: " + follow_up_prop + "\n";
      msg += "follow_up_sheet_name: " + follow_up_sheet_name + "\n";
      msg += "follow_up_trigger_id: " + follow_up_trigger_id;
      myLogger(msg);
      
      sheet = ss.getSheetByName(follow_up_sheet_name);
      values = sheet.getDataRange().getValues();
      case_list_remaining = values;      
      
      
      // delete the current follow up sheet
      try {
        Utilities.sleep(500);
        ss.deleteSheet(sheet);
        Utilities.sleep(500);
        msg = "current follow up sheet deleted: " + follow_up_sheet_name;
        myLogger(msg);
      } catch (err) {
        msg = "error - failed to delete the follow up sheet: " + follow_up_sheet_name;
        msg += "\nMessage: " + err.message;
        msg += "\nScript: " + err.fileName;
        msg += "\nLine: " + err.lineNumber;
        myLogger(msg);
        return;
      }
      
      
      //loop over all triggers to find and remove the follow up trigger
      var allTriggers = ScriptApp.getProjectTriggers();
      for (var i = 0; i < allTriggers.length; i++) {
        if (allTriggers[i].getUniqueId() === follow_up_trigger_id) {
          ScriptApp.deleteTrigger(allTriggers[i]);
          msg = "current follow up trigger deleted: " + follow_up_trigger_id;
          myLogger(msg);
          break;
        }
      }
      
      
      // delete the follow_up setting
      removeSettings('follow_up_cases' + follow_up_sheet_name);
      msg = "current follow up setting deleted: follow_up_cases" + follow_up_sheet_name;
      myLogger(msg);
      
      
      
      // process that array
      msg = "processing follow up cases";
      myLogger(msg);
      for ( r = (values.length - 1); r >= 0; --r ) {
        
        if ( values[r][0].length < 1 ) { continue; }
        
        caseCount++;
        caseNo = values[r][0];
        
        if ( caseNo === 'Case number' ) {
          rCheck = "finished";
        } else {
          rCheck = processCase( caseNo, atty, attyEmail );
        }
        
        // remove the case from the case_list_remaining array
        if ( rCheck === "finished" ) {
          for ( k = (case_list_remaining.length - 1); k >= 0; --k ) {          
            if (case_list_remaining[k][0].indexOf(caseNo) > -1) {
              case_list_remaining.splice(k, 1);
              break;
            }          
          }
        }
        
        currTime = (new Date()).getTime();
        timeLimitIsNear = (currTime - script_start >= max_running_time);
        
        if ( timeLimitIsNear ) { break; }
        
      }
      
      if ( timeLimitIsNear || (case_list_remaining.length > 0) ) {
        
        caseRemCount = case_list_remaining.length;
        followUp( case_list_remaining );
        Utilities.sleep(500);
        
      }
      
      wrapUp(script_name, script_start, caseCount, caseAddedCount, caseRemCount, caseUpdateCount, caseAttachmentCount);
      
    } else {
      
      //no follow up cases from last run
      
      //loop through every sheet and push every case to an array
      for (var sheetNum = 0; sheetNum < allSheets.length; sheetNum++) {
        
        sheet = allSheets[sheetNum];
        
        if (sheet.getName().indexOf("folUp") >= 0) {
          msg = "error - extraneous follow-up sheet remaining: " + sheet.getName();
          myLogger(msg);
          ss.deleteSheet(sheet);
          continue;
        }
        
        if ( (sheet.getName().indexOf("AR") < 0) && (sheet.getName().indexOf("ther cases") < 0) ) { continue; }
        
        // iterate through every row/case on the sheet and add it to the case_list array
        if (sheet.getLastRow() === 0) continue;
        values = sheet.getDataRange().getValues();
        
        for ( r = 0; r < values.length; r++) {
          
          if ( values[r][0].length > 0 ) {
            
            caseNo = values[r][0];            
            case_list.push( [caseNo, barNo, atty, attyEmail] );
            
          }
        }
      }
      
      
      // add the array to a sheet called case_list_sheet
      case_list_sheet = ss.getSheetByName("case_list_sheet");
      if (case_list_sheet !== null) {
        ss.deleteSheet(ss.getSheetByName("case_list_sheet"));
        Utilities.sleep(500);
      }
      
      case_list_sheet = ss.getSheetByName("case_list_sheet");
      if (case_list_sheet === null) {
        case_list_sheet = ss.insertSheet();
        Utilities.sleep(500);
        case_list_sheet.setName("case_list_sheet");
        Utilities.sleep(500);
      }
      
      sheet = ss.getSheetByName("case_list_sheet");
      Utilities.sleep(500);
      var listRange = sheet.getRange(1, 1, case_list.length, case_list[0].length);
      listRange.setValues(case_list);
      
      Utilities.sleep(500);
      
      currTime = (new Date()).getTime();
      timeLimitIsNear = (currTime - script_start >= max_running_time);
      if ( timeLimitIsNear ) {
        myLogger("timeLimitIsNear fired before processing a case");
        followUp( case_list );
        Utilities.sleep(500);
        ss.deleteSheet(case_list_sheet);
        Utilities.sleep(500);
        wrapUp(script_name, script_start, caseCount, caseAddedCount, caseRemCount, caseUpdateCount, caseAttachmentCount);
        return false;
      }
      
      
      sheet = ss.getSheetByName('case_list_sheet');
      if (sheet.getLastRow() === 0) return false;
      values = sheet.getDataRange().getValues();
      case_list_remaining = values;
      
      
      // remove the case list sheet
      ss.deleteSheet(ss.getSheetByName("case_list_sheet"));
      Utilities.sleep(500);
      
      currTime = '';
      timeLimitIsNear = '';
      caseNo = '';
      barNo = '';
      atty = '';
      attyEmail = '';
      
      // process the cases
      for ( r = (values.length - 1); r >= 0; --r ) {
        
        if ( values[r][0].length < 1 ) { continue; }
        
        caseCount++;
        caseNo = values[r][0];
        barNo = values[r][1];
        atty = values[r][2];
        attyEmail = values[r][3];
        
        if ( caseNo === 'Case number' ) {
          rCheck = "finished";
        } else {
          rCheck = processCase( caseNo, atty, attyEmail );
        }
        
        // remove the case from the case_list_remaining array
        if ( rCheck === "finished" ) {
          for ( k = (case_list_remaining.length - 1); k >= 0; --k ) {          
            if (case_list_remaining[k][0].indexOf(caseNo) > -1) {
              case_list_remaining.splice(k, 1);
              break;
            }          
          }
        }
        
        currTime = (new Date()).getTime();
        timeLimitIsNear = (currTime - script_start >= max_running_time);
        
        if ( timeLimitIsNear ) { break; }
        
      }
      
      if ( timeLimitIsNear || (case_list_remaining.length > 0) ) {
        
        caseRemCount = case_list_remaining.length;        
        followUp( case_list_remaining );        
        Utilities.sleep(500);
        
      }
      
      //ss.setActiveSheet(ss.getSheetByName('settings'));
      case_list_sheet = '';
      sheet = '';
      dataRange = '';
      values = '';
      case_list_remaining = '';
      
      wrapUp(script_name, script_start, caseCount, caseAddedCount, caseRemCount, caseUpdateCount, caseAttachmentCount);
      
    }
    
  } catch (err) {
    msg = "\nMessage: " + err.message;
    msg += "\nScript: " + err.fileName;
    msg += "\nLine: " + err.lineNumber;
    myLogger(msg);
    
	var settings = getColumnsData_(SpreadsheetApp.getActive().getSheetByName( 'settings' ))[0];
	var adminEmail = settings.attorneyEmail;
	var appUrl = SpreadsheetApp.getActive().getUrl();
	var email_error_reports = getSettings("email_error_reports");
	if ( email_error_reports === true ) { MailApp.sendEmail({name: 'Docket Monitor', to: adminEmail, subject: 'docketMonitor error report', body: msg + "\nSheet:\n" + appUrl}); }
  }
  
}


function processAtty( barNo, atty, attyEmail ) {
  
  try {
    
    var script_start = (new Date()).getTime();
    
    var ss = SpreadsheetApp.getActive();
    var attySheet = ss.getSheetByName(barNo);
    
    if (attySheet === null) {
      attySheet = ss.insertSheet();
      Utilities.sleep(500);
      attySheet.setName(barNo);
      Utilities.sleep(500);
    }
    
    attySheet = ss.getSheetByName(barNo);
    
    
    // fetch the URL for atty's CourtConnect page
    var getDataURL = 'https://caseinfo.aoc.arkansas.gov/cconnect/PROD/public/ck_public_qry_cpty.cp_personcase_srch_details?backto=P&id_code=' + barNo;
    getDataURL = encodeURI(getDataURL);
    var getDataURLbase = getDataURL;
    var fetch, response;
    var arrMatch = [];
    var caseTemp, caseStatus;
    var caseList = [];
    var c = 1;
    var nextTest = 1;
    var retry = 0;
    var rStr = "unfinished";
    
    var rePattern = new RegExp("case_id=(.*?)&(.*?)<br><b>status:(.*?)<\/i>","gi");    
    
    // fetch a full list of cases by iterating through every page of results
    // search html for case numbers using a regex pattern
    while (nextTest > 0) {
      
      fetch = UrlFetchApp.fetch(String(getDataURL), fetchOptions);
      if ( fetch.getResponseCode() !== 200 ) {
        msg = "response code " + fetch.getResponseCode() + " for " + getDataURL + "";
        myLogger(msg);
        return;
      }
      response = fetch.getContentText();
      
      if ( (!response) || (response === "") ) {
        msg = "no response (" + getDataURL + ")";
        myLogger(msg);
      }
      
      while ( (!response) || (response === "") ) {
        retry++;
        if ( retry > 4 ) {
          msg = "no response-- giving up";
          myLogger(msg);
          return rStr;
        }
        msg = "no response, retrying fetch";
        myLogger(msg);
        Utilities.sleep(500);
        response = UrlFetchApp.fetch(String(getDataURL), fetchOptions).getContentText();
      }
      
      // determine if another page of results exists
      // if not, set nextTest to 0, which will break the loop after this iteration
      if ( (response.indexOf("Next") < 0) && (response.indexOf("next") < 0) ) { 
        nextTest = 0;
        if (c < 2) { 
          msg = "only a single page of results found for " + atty;
          myLogger(msg);
        }
      }
      
      if ( (response.indexOf("Case:") < 0) ) { 
        break;
      }
      
      while (arrMatch = rePattern.exec( response )) {
        
        // push each match to array of case numbers
        caseTemp = arrMatch[ 0 ];
        caseTemp = caseTemp.replace(/case_id=/i, '');
        caseTemp = caseTemp.replace(/&([^]+)/i, '');
        
        var settings = getColumnsData_(ss.getSheetByName( 'settings' ))[0];
        var exclude_closed_cases = settings.excludeClosedCases;
        if ( !exclude_closed_cases ) {
          
          caseList.push( [caseTemp, barNo, atty, attyEmail] );
          
        } else {
          
          caseStatus = arrMatch[ 0 ];
          caseStatus = caseStatus.replace(/&nbsp;/gi, '');
          caseStatus = caseStatus.replace(/(.*?)\/b>/i, '');
          caseStatus = caseStatus.replace(/<\/i>/i, '');
          if ( !caseStatus ) { caseStatus = "UNLISTED"; }
          
          if ( caseStatus.toLowerCase() !== "closed" ) {
            caseList.push( [caseTemp, barNo, atty, attyEmail] );
          }
          
        }
        
        currTime = (new Date()).getTime();
        timeLimitIsNear = (currTime - script_start >= max_running_time);
        if ( timeLimitIsNear ) {
          msg = "reached time limit before finishing case list update for " + atty;
          myLogger(msg);
          //break;
          return rStr;
        }
        
      }
      
      // if another page of results exists, fetch the URL
      c++;
      getDataURL = getDataURLbase + "&PageNo=" + c;
      
    }
    
    if ( (typeof caseList[0] === "undefined") || (caseList.length === 0) ) {
      msg = "error - no cases found for " + atty;
      myLogger(msg);
      
      return rStr;
    }
    
    // remove duplicate case numbers
    var newData = [];
    var data = caseList;
    for( var i in data ){
      var row = data[i];
      var duplicate = false;
      for( var j in newData ){
        if(row.join() === newData[j].join()){ duplicate = true; }
      }
      if(!duplicate){ newData.push(row); }
    }
    
    if ( (newData.length === 0) || (typeof newData[0] === "undefined") ) {
      msg = "error - no cases found for " + atty;
      myLogger(msg);
      return rStr;
    }
    
    var numRows, numCols, listRange;
    
    numRows = attySheet.getLastRow();
    
    var minRows = (numRows > 0) ? (numRows / 2) : 0;
    
    if ( newData.length > minRows ) {
      
      attySheet.clearContents();
      
      numRows = newData.length;
      numCols = newData[0].length;
      listRange = attySheet.getRange(1, 1, numRows, numCols);
      listRange.setValues(newData);
      
      msg = "case list updated for " + atty + " (" + barNo + ")" + " [" + numRows + " cases]";
      myLogger(msg);
      
      rStr = "finished";
      
    } else {
      
      msg = "error - insufficient cases found for " + atty + "\n";
      msg += getDataURLbase + "\n";
      msg += "(original records: " + numRows + "    new records: " + newData.length + ")";
      myLogger(msg);
      
    }
    
    return rStr;
    
  } catch (err) {
    msg = "\nMessage: " + err.message;
    msg += "\nScript: " + err.fileName;
    msg += "\nLine: " + err.lineNumber;
    myLogger(msg);
    
	var settings = getColumnsData_(SpreadsheetApp.getActive().getSheetByName( 'settings' ))[0];
	var adminEmail = settings.attorneyEmail;
	var appUrl = SpreadsheetApp.getActive().getUrl();
	var email_error_reports = getSettings("email_error_reports");
	if ( email_error_reports === true ) { MailApp.sendEmail({name: 'Docket Monitor', to: adminEmail, subject: 'docketMonitor error report', body: msg + "\nSheet:\n" + appUrl}); }
  }
}




function processCase( caseNo, atty, attyEmail ) {
  
  try {
    
    if ( caseNo.length > 0 ) {
      
      var rStr = "unfinished";
      
      if ( caseNo === 'Case number' ) {
        rStr = "finished";
        return rStr;
      }
      
      var ss = SpreadsheetApp.getActive();
      
      // check ignore list for this case
      var ignoreSheet = ss.getSheetByName('ignore cases');
      if (ignoreSheet !== null) {
        var ignore = ignoreSheet.getRange(1, 1, ignoreSheet.getLastRow(), ignoreSheet.getLastColumn()).getValues();
        if (String(ignore).indexOf(caseNo) >= 0) {
          rStr = "finished";
          return rStr;
        }
      }
      
      var settings = getColumnsData_(ss.getSheetByName( 'settings' ))[0];
      
      var appSettings = ss.getSheetByName("appSettings");
      if (appSettings === null) {
        appSettings = ss.insertSheet();
        Utilities.sleep(500);
        appSettings.hideSheet();
        Utilities.sleep(500);
        appSettings.setName("appSettings");
        Utilities.sleep(500);
      }
      var AppSubFolderId = getSettings('AppSubFolderId');
      var AppSubFolder = DriveApp.getFolderById(AppSubFolderId);      
      
      var getDataURL = 'https://caseinfo.aoc.arkansas.gov/cconnect/PROD/public/ck_public_qry_doct.cp_dktrpt_docket_report?backto=P&case_id=' + caseNo;
      
      var response = UrlFetchApp.fetch(getDataURL, fetchOptions);
      var siteStatus = response.getResponseCode();
      var caseResponse = response.getContentText();
      
      // determine if http response is valid    
      if ( siteStatus !== 200 ) {
        msg = caseNo + ": error - server response was " + siteStatus + " [" + getDataURL + "] (" + atty + ")";
        myLogger(msg);
        return rStr;
      }
      
      // determine if url is valid
      var temp = '';
      var tempR = /(no case was found|no record was found|no records found|no case found)/i;
      temp = String(caseResponse).toLowerCase().match(tempR);
      if (temp) {
        msg = caseNo + ": error - no case was found [" + getDataURL + "] (" + atty + ")";
        myLogger(msg);
        rStr = "finished";
        return rStr;
      }
      
      temp = '';
      tempR = /(case description)/i;
      temp = String(caseResponse).toLowerCase().match(tempR);
      var tempRE = /(under maintenance|planned maintenance|scheduled maintenance|routine maintenance|page offline|downtime)/i;
      var tempE = String(caseResponse).toLowerCase().match(tempRE);
      if (!temp && !tempE) {
        msg = caseNo + ": server error - !case description && !under maintenance [" + getDataURL + "] (" + atty + ")";
        myLogger(msg);
        return rStr;
      }
      
      temp = '';
      tempR = /(under maintenance|planned maintenance|scheduled maintenance|routine maintenance|page offline|downtime)/i;
      temp = String(caseResponse).toLowerCase().match(tempR);
      if (temp) {
        msg = caseNo + ": site seems to be under maintenance [" + getDataURL + "] (" + atty + ")";
        myLogger(msg);
        return rStr;
      }
      
      
      // determine if the case docket was stored to a document in a previous run
      var caseFile = '';
      var caseFileId;
      caseFileId = getSettings(caseNo);
      
      if ( caseFileId ) {
        try {
          caseFile = DriveApp.getFileById(caseFileId);
          if (caseFile.isTrashed()) { caseFile = ''; }
        } catch (err) {
          caseFile = '';
          msg = caseNo + ": case file error (" + atty + ")";
          msg += "\nMessage: " + err.message;
          msg += "\nScript: " + err.fileName;
          msg += "\nLine: " + err.lineNumber;
          myLogger(msg);
          return rStr;
        }
      }
      
      
      // if a case docket document does not exist, create it, and skip to next iteration
      var nCaseText;
      if ( !caseFileId || !caseFile ) {
        caseAddedCount++;
        var date = Utilities.formatDate(new Date(), "America/Chicago", "MM/dd/yyyy");
        caseFile = AppSubFolder.createFile(caseNo, caseResponse, MimeType.PLAIN_TEXT);
        caseFileId = caseFile.getId();
        appSettings.appendRow([caseNo,caseFileId,date]);
        
        nCaseText = processHTML(String(caseResponse), String(caseResponse).length);
        nCaseText = processText(String(nCaseText), String(nCaseText).length);
        
        var new_case_updates = getSettings("new_case_updates");
        
        if ( new_case_updates === true ) {
          subject = 'case added successfully';
          body = caseNo + ' was added to your docket monitor.\n' + getDataURL + "\n----------------------------------------\n";
          body += nCaseText + "\n----------------------------------------\n(" + caseNo + " docket: " + getDataURL + ")" + "\nAttorney: " + atty;  
          MailApp.sendEmail({name: 'Docket Monitor', to: attyEmail, subject: subject, body: body});
        }
        msg = caseNo + ": case added successfully (" + atty + ")";
        myLogger(msg);
        rStr = "finished";
        return rStr;
      }
      
      
      // if the case docket document does exist, return the data inside it for comparison
      var pCaseText = '';
      var blob = DriveApp.getFileById(caseFileId);
      var txtBlob = blob.getBlob();
      pCaseText = txtBlob.getDataAsString();
      
      // compare the existing case docket document to the http response
      // if the response is different, replace the existing case docket document, and generate an update email
      if (caseResponse !== pCaseText) {
        
        caseUpdateCount++;
        
        var send_updates = getSettings("send_updates");
        
        // check for case docs to attach to update email
        var attach = [];
        if ( send_updates === true ) {
          attach = processAttachments(caseNo, atty, pCaseText, caseResponse);
        }
        
        // highlight the new text
        pCaseText = processHTML(String(pCaseText), String(pCaseText).length);
        pCaseText = processText(String(pCaseText), String(pCaseText).length);
        
        nCaseText = processHTML(String(caseResponse), String(caseResponse).length);
        nCaseText = processText(String(nCaseText), String(nCaseText).length);
        
        var dmp = new diff_match_patch();
        //dmp.Diff_Timeout = '10';
        var d = dmp.diff_main(pCaseText, nCaseText);
        dmp.diff_cleanupSemantic(d);
        var ds = dmp.diff_prettyHtml(d);
        ds = ds.replace(/&para;/g, '');
        
        subject = caseNo + ' Docket Monitor Update';
        body = ds + "<br /><hr>(" + caseNo + " docket: " + getDataURL + ")" + "<br />Attorney: " + atty;
        
        if ( attach.length > 0 ) {
          if ( send_updates === true ) { MailApp.sendEmail({name: 'Docket Monitor', to: attyEmail, subject: subject, htmlBody: body, attachments:attach}); }
        } else {
          if ( send_updates === true ) { MailApp.sendEmail({name: 'Docket Monitor', to: attyEmail, subject: subject, htmlBody: body}); }
        }
        
        // save the new content to the docket file
        caseFile = DriveApp.getFileById(caseFileId);
        caseFile.setContent(caseResponse);  
        msg = caseNo + ": docket file updated (" + atty + ")";
        myLogger(msg);
        rStr = "finished";
        return rStr;
        
      }
      
      rStr = "finished";
      return rStr;
      
    }
    
  } catch (err) {
    msg = "\nMessage: " + err.message;
    msg += "\nScript: " + err.fileName;
    msg += "\nLine: " + err.lineNumber;
    myLogger(msg);
    
	var settings = getColumnsData_(SpreadsheetApp.getActive().getSheetByName( 'settings' ))[0];
	var adminEmail = settings.attorneyEmail;
	var appUrl = SpreadsheetApp.getActive().getUrl();
	var email_error_reports = getSettings("email_error_reports");
	if ( email_error_reports === true ) { MailApp.sendEmail({name: 'Docket Monitor', to: adminEmail, subject: 'docketMonitor error report', body: msg + "\nSheet:\n" + appUrl}); }
  }
  
}



function processAttachments( caseNo, atty, pCaseText, caseResponse ) {
  
  try {
    
    var ss = SpreadsheetApp.getActive();
    var appUrl = SpreadsheetApp.getActive().getUrl();
    var settings = getColumnsData_(ss.getSheetByName( 'settings' ))[0];
    var adminEmail = settings.attorneyEmail;
    
    var attach = [];
    var match;
    var pdf = '';
    var cAttCount = 0;
    
    var d = diffString_( pCaseText, caseResponse );
    d = String(d).replace(/&quot;/gi, '"');
    d = String(d).replace(/%20/gi, ' ');
    d = String(d).replace(/&lt;/gi, '<');
    d = String(d).replace(/&gt;/gi, '>');
    d = String(d).replace(/\n<\/ins><ins>/gi, '</ins>\n<ins>');
    d = String(d).replace(/<\/ins><ins>/gi, '');
    
    var docPattern = new RegExp("<ins><?T?D?>?<a *?href=\"http(.*?)image(.*?)<\/a>","gim");
    var tDate = Utilities.formatDate(new Date(), "America/Chicago", "yyyyMMdd");
    while (arrMatch = docPattern.exec( d )) {
      
      cAttCount++;
      if ( cAttCount > 25 ) {
        msg = caseNo + ": attachment notice - excessive number of attachments detected (" + atty + ")";
        myLogger(msg);
        break;
      }
      
      match = arrMatch[ 0 ];
      match = String(match).replace(/%20/gi, ' ');
      
      var pdfURL = match;
      pdfURL = pdfURL.replace(/(.*?)href=\"(.*?)\" *?target([^]+)/i, '$2');
      pdfURL = pdfURL.replace(/images\/dms\/ck_image.present\?/i, 'IMAGES/DMS/ck_image.present2?');
      
      if ( pdfURL.indexOf('href="') !== -1 ) {
        pdfURL = pdfURL.split('href="')[1];
        if ( pdfURL.indexOf('"') !== -1 ) {
          pdfURL = pdfURL.split('"')[0];
        }
      }
      
      if ( pdfURL.indexOf('192.168') >= 0 ) {
        msg = caseNo + ": attachment error (" + atty + ")\nlocal ip listed for server";
        myLogger(msg);
        continue;
      }
      
      var pdfLinkText = match;
      pdfLinkText = pdfLinkText.replace(/(.*?)blank\">(.*?)<\/a>/i, '$2');
      pdfLinkText = toTitleCase_(pdfLinkText);
      
      var pdfName = caseNo + '-' + cAttCount + "-" + pdfLinkText + '-' + tDate + '.pdf';
      
      try {
        var blob = UrlFetchApp.fetch(pdfURL, fetchOptions);
        
        if ( blob.getResponseCode() !== 200 ) {
          msg = "attachment response code " + blob.getResponseCode() + " for " + pdfURL;
          myLogger(msg);
        } else {
          var pdfBlob = blob.getBlob().setContentType("application/pdf");
          pdf = pdfBlob.getAs("application/pdf").getBytes();
        }
      } catch (err) {
        msg = "attachment error\n" + "link:\n" + pdfURL + "\n(" + atty + ")\n";
        msg += "\nMessage: " + err.message;
        msg += "\nScript: " + err.fileName;
        msg += "\nLine: " + err.lineNumber;
        myLogger(msg);
        
        var email_error_reports = getSettings("email_error_reports");
        if ( email_error_reports === true ) {
          if ( msg.indexOf("Address unavailable") < 0 ) {
            subject = "docket monitor attachment error";
            MailApp.sendEmail({name: 'Docket Monitor', to: adminEmail, subject: subject, body: msg + "\nSheet:\n" + appUrl});
          }
        }
      }
      
      if ( (typeof pdf === "undefined") || !pdf || (typeof pdf.length === "undefined") || (pdf.length <= 0) || (String(blob).toLowerCase().indexOf('pdf') < 0) ) {
        
        msg = caseNo + ": pdf size error (dead or unresp link) - " + pdfLinkText + " [" + pdfURL + "] (" + atty + ")";
        myLogger(msg);
        
      } else {
        
        try {
          caseAttachmentCount++;
          attach.push({fileName:pdfName, content:pdf, mimeType:'application/pdf'});
          msg = caseNo + ": " + pdfLinkText + " [" + pdfURL + "] (" + atty + ")";
          myLogger(msg);
        } catch (err) {
          msg = "attachment push error\n" + "link:\n" + pdfURL + "\n(" + atty + ")\n" + "attachment link text:\n" + match;
          msg += "\nMessage: " + err.message;
          msg += "\nScript: " + err.fileName;
          msg += "\nLine: " + err.lineNumber;
          myLogger(msg);
          
          msg += "\n\nattachment text:\n" + d + "\n\n";
          subject = "docket monitor attachment push error";
          if ( msg.indexOf("Address unavailable") < 0 ) {
            MailApp.sendEmail({name: 'Docket Monitor', to: adminEmail, subject: subject, body: msg + "\nSheet:\n" + appUrl});
          }
        }
        
      }
      
    }
    
  } catch (err) {
    
    msg = caseNo + ": attachment error (" + atty + ")";
    msg += "\nMessage: " + err.message;
    msg += "\nScript: " + err.fileName;
    msg += "\nLine: " + err.lineNumber;
    myLogger(msg);
    
	var settings = getColumnsData_(SpreadsheetApp.getActive().getSheetByName( 'settings' ))[0];
	var adminEmail = settings.attorneyEmail;
	var appUrl = SpreadsheetApp.getActive().getUrl();
	var email_error_reports = getSettings("email_error_reports");
	if ( email_error_reports === true ) { MailApp.sendEmail({name: 'Docket Monitor', to: adminEmail, subject: 'docketMonitor error report', body: msg + "\nSheet:\n" + appUrl}); }
    
  }
  
  return attach;
  
}



function followUp( case_list_remaining ) {
  try {
    
    var script_name = "followUp";
    msg = script_name + ' running';
    myLogger(msg);
    
    var ss = SpreadsheetApp.getActive();
    var appSettings = ss.getSheetByName("appSettings");
    if (appSettings === null) {
      appSettings = ss.insertSheet();
      Utilities.sleep(500);
      appSettings.hideSheet();
      Utilities.sleep(500);
      appSettings.setName("appSettings");
      Utilities.sleep(500);
    }
    var dateTime = Utilities.formatDate(new Date(), "America/Chicago", "yyyyMMddHHmmss");
    var tempName = String(dateTime + 'folUp');
    
    //create a sheet & name it
    var case_list_rem_sheet = ss.getSheetByName(tempName);
    if (case_list_rem_sheet !== null) {
      Utilities.sleep(500);
      ss.deleteSheet(ss.getSheetByName(tempName));
      Utilities.sleep(500);
    }
    
    case_list_rem_sheet = ss.insertSheet();
    Utilities.sleep(500);
    case_list_rem_sheet.setName(tempName);
    Utilities.sleep(500);
    
    var follow_up_sheet_name;
    
    try {
      follow_up_sheet_name = case_list_rem_sheet.getName();
      Utilities.sleep(500);
    } catch (err) {
      msg = "follow_up_sheet_name error\n";
      msg += "follow_up_sheet_name = case_list_rem_sheet.getName() failed";
      myLogger(msg);
    }
    
    if (!follow_up_sheet_name) {
      case_list_rem_sheet = ss.getSheetByName(tempName);
      Utilities.sleep(500);
      follow_up_sheet_name = tempName;
    }
    
    // add the remaining cases to the sheet
    var lastRow = case_list_rem_sheet.getLastRow();
    var listRange = case_list_rem_sheet.getRange((lastRow + 1), 1, case_list_remaining.length, case_list_remaining[0].length);
    listRange.setValues(case_list_remaining);
    
    msg = "new follow up sheet populated: " + follow_up_sheet_name;
    myLogger(msg);
    
    //create a trigger
    currTime = (new Date()).getTime();
    var waitTime = (1000 * 60 * trigger_delay_mins);
    var follow_up_trigger = ScriptApp.newTrigger("dmProcessLock").timeBased().at((new Date(currTime + waitTime))).create();
    var follow_up_trigger_id = follow_up_trigger.getUniqueId();
    
    msg = "new follow up trigger scheduled for " + new Date(currTime + waitTime);
    myLogger(msg);
    
    //set script settings
    var tempProp = String(follow_up_sheet_name + '|' + follow_up_trigger_id);
    appSettings.appendRow(["follow_up_cases" + tempName, tempProp]);
    
  } catch (err) {
    msg = "\nMessage: " + err.message;
    msg += "\nScript: " + err.fileName;
    msg += "\nLine: " + err.lineNumber;
    myLogger(msg);
    
	var settings = getColumnsData_(SpreadsheetApp.getActive().getSheetByName( 'settings' ))[0];
	var adminEmail = settings.attorneyEmail;
	var appUrl = SpreadsheetApp.getActive().getUrl();
	var email_error_reports = getSettings("email_error_reports");
	if ( email_error_reports === true ) { MailApp.sendEmail({name: 'Docket Monitor', to: adminEmail, subject: 'docketMonitor error report', body: msg + "\nSheet:\n" + appUrl}); }
  }
}


function wrapUp( script_name, script_start, caseCount, caseAddedCount, caseRemCount, caseUpdateCount, caseAttachmentCount ) {
  
  try {
    
    msg = 'wrapUp running';
    myLogger(msg);
    
    var ss = SpreadsheetApp.getActive();
    var settings = getColumnsData_(ss.getSheetByName( 'settings' ))[0];
    var adminEmail = settings.attorneyEmail;
    
    var script_end = (new Date()).getTime();
    var script_dur = (((script_end - script_start) / 1000) / 60);
    
    var appUrl = SpreadsheetApp.getActive().getUrl();
    
    var reportDiv = "\n" + "==========================================";
    var bodyDiv = ":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::";
    bodyDiv += ":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::";
    
    msg = "";
    msg += script_name + " initiated";
    msg += " on " + Utilities.formatDate(new Date(script_start), "America/Chicago", "yyyy-MM-dd");
    msg += " at " + Utilities.formatDate(new Date(script_start), "America/Chicago", "h:mm a") + "\n";
    msg += 'Script duration: ' + (script_dur.toFixed(1)) + ' minutes' + reportDiv;
    if ( caseCount > 0) {
      msg += '\nStats:\n~ ' + caseCount + ' dockets analyzed (' + (caseCount / script_dur).toFixed(2) + " cases per minute)\n";
      if ( caseAddedCount > 0) { msg += "~ " + caseAddedCount + ' new cases added\n'; }
      msg += "~ " + caseUpdateCount + ' dockets updated\n';
      msg += "~ " + caseAttachmentCount + ' attachments discovered\n';
      msg += "~ " + caseRemCount + ' dockets remaining' + reportDiv;
    }
    msg = msg.replace(/1 dockets/ig, '1 docket');
    msg = msg.replace(/1 attachments/ig, '1 attachment');
    msg = msg.replace(/1 cases/ig, '1 case');
    
    myLogger(msg);
    
    var send_logs = getSettings("send_logs");
    if ( send_logs === true ) {
      var logText = Logger.getLog();
      msg = "\nScript log:" + "\n" + logText + bodyDiv;
      MailApp.sendEmail({name: 'Docket Monitor', to: adminEmail, subject: "docket monitor log", body: msg + "\nSheet:\n" + appUrl});
    }
    Logger.clear();
    
  } catch (err) {
    msg = "\nMessage: " + err.message;
    msg += "\nScript: " + err.fileName;
    msg += "\nLine: " + err.lineNumber;
    myLogger(msg);
    
    var settings = getColumnsData_(SpreadsheetApp.getActive().getSheetByName( 'settings' ))[0];
    var adminEmail = settings.attorneyEmail;
    var appUrl = SpreadsheetApp.getActive().getUrl();
    var email_error_reports = getSettings("email_error_reports");
    if ( email_error_reports === true ) { MailApp.sendEmail({name: 'Docket Monitor', to: adminEmail, subject: 'docketMonitor error report', body: msg + "\nSheet:\n" + appUrl}); }
  }
  
}

/////////////////////////
// peripheral functions

function myLogger(msg) {
  try {
    var logDiv = "\n" + "--------------------------------------------------------------------";
    var reportDiv = "==========================================";
    var bodyDiv = "\n" + ":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::";
    bodyDiv += ":::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::";
    
    var logSS = SpreadsheetApp.getActive();
    var logSheet = logSS.getSheetByName( 'log' );
    if (logSheet === null) {
      logSheet = logSS.insertSheet();
      Utilities.sleep(500);
      logSheet.hideSheet();
      Utilities.sleep(500);
      logSheet.setName( 'log' );
      Utilities.sleep(500);
    }
    if (logSheet.getLastRow() >= 500) {
      var data = logSheet.getRange(475, 1, logSheet.getLastRow(), logSheet.getLastColumn()).getValues();
      logSheet.clearContents();
      logSheet.getRange(1,1,(data.length),(data[0].length)).setValues(data);
      Utilities.sleep(500);
    }
    var dt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd h:mm a':\n'");
    logSheet.appendRow(["INFO " + String(dt) + msg + logDiv]);
  } catch (err) { return; }
}



function toTitleCase_(str) {
  return str.replace(/\w\S*/g, function(txt){return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();});
}



function processHTML(html, count) {
  
  html = html.replace(/<(?:.|\n)*?>/gm, '');
  html = html.replace(/^\s+|\s+$/g, '');
  
  //return html.substring(0, count);
  return html;
  
}


function processText(text, count) {
  
  //entities
  text = text.replace(/(&?nbsp;?|&?amp;?)/gi, '\n');
  text = text.replace(/(&lt;\/TABLE&gt; *)/gi, '\n');
  text = text.replace(/(<\/TABLE> ?)/gi, '\n');
  
  //white space
  text = text.replace(/  /g, " ");
  text = text.replace(/(    )/g, "\n");
  text = text.replace(/(\n)( +)/g, "$1");
  text = text.replace(/(\n)(\n+)/gi, "$1");
  
  //text = text.replace(/( \n   \n\n)/gi, "\n\n");
  //text = text.replace(/(\n)(\n{2,})/gi, "\n\n");
  //text = text.replace(/([a-z])      ([0-9])/gi, "$1 $2");
  text = text.replace(/( {2,}\n)/gi, '\n');
  //text = text.replace(/( {7,})/gi, '\n');
  text = text.replace(/( +,)/gi, ",");
  
  //dates and times
  text = text.replace(/(\d{1,2}\/\d{1,2}\/\d{2,4})/gi, "\n$1");
  text = text.replace(/(January|February|March|April|May|June|July|August|September|October|November|December)(\n{1,})(\d{2})/gi, "$1 $3");
  //text = text.replace(/(\d{1,2}\/\d{1,2}\/\d{2,4})( at )(\d{1,2}:\d{1,2}) (AM|PM) /gi, "$1$2$3 $4\n");
  text = text.replace(/(2)(\d{3})(\d{2}\:)/gi, "$1$2 $3");
  text = text.replace(/(\d{1,2}\/\d{1,2}\/\d{2,4})([\n]|[\s]{1,})(\d{2}:\d{2})/gi, "$1 at $3");
                      
  //misc text
  text = text.replace(/(\- Not an Official Document|Report Selection Criteria)/gi, "");
  text = text.replace(/(Event\nDate\/Time\nRoom\nLocation\nJudge *)/, "");
  text = text.replace(/(Seq #\nAssoc\nEnd Date\nType\nID\nName\n)/, "");  
  text = text.replace(/(Filing Date\nDescription\nName\nMonetary\n)/, "");
  
  text = text.replace(/(\n)(-)/g, " $2 ");
  //text = text.replace(/\n\n(\d)\n/g, "\n$1\n");
  //text = text.replace(/(\nnone\.?)/gi, "");
  text = text.replace(/(Aliases:)(\n)([A-Z])/gi, "$1 $3");  
  text = text.replace(/(\n)(\n)?(Aliases:)(\n| )?(none)?/gi, "\n$2");
  
  text = text.replace(/(\n)(Entry)(:?)(\n| )?(none)(\.)?/gi, "");
  text = text.replace(/(\n)(Images)(:?)(\n| )?(none|no images)(\.)?/gi, "");
  //text = text.replace(/(No Images)(\.| )?/gi, "");
  
  text = text.replace(/(Entry)(?!:)/gi,"Entry:");
  text = text.replace(/(Entry:)(\n)/gi, "$1 ");
  text = text.replace(/(Entry:) (Images)/gi, "$1\n$2");
  text = text.replace(/(Images)(?!:)/gi,"Images:");
  //text = text.replace(/(Images:)(\n)/gi, "$1 ");
  text = text.replace(/(Images:)(\n)([A-Z])/gi, "$1 $3");
  
  text = text.replace(/(Case ID:|Filing Date:|Court:|Location:|Type:|Status:)( *\n)/gi, "$1 ");
  text = text.replace(/(Violation Date:|Violation Time:)( *\n)/gi, "$1 \n");
  text = text.replace(/(Sentence)(:)*(\n| )*(No Sentence Info)( found)?(\.)?(\n)?/gi, "Sentence:");
  text = text.replace(/(Milestone Tracks)(:)*(\n| )*(No Milestone Tracks)( found)?(\.)?/gi, "");
  text = text.replace(/(Violation: *)\n(\d) *\n/gi, "$1 $2\n");
  text = text.replace(/(No docket entries found\.)/gi, "No docket entries found.");
  
  //simple plaintext formatting
  text = text.replace(/(\n)(\n{2,})/gi, "\n\n");
  //text = text.replace(/(Case Event Schedule)([\s\n]+)(\w)/gi, "$1\n$3");
  text = text.replace(/(Case Description|Case Event Schedule|Case Parties|Violations|Docket Entries)/gi, "\n\n*** $1 ***");
  //text = text.replace(/(Sentence)(\nName)/gi, "\n\n*** $1 ***$2");  
  text = text.replace("*** Case Description ***", "*** CASE DESCRIPTION ***");
  text = text.replace("*** Case Event Schedule ***", "*** CASE EVENT SCHEDULE ***");
  text = text.replace("*** Case Parties ***", "*** CASE PARTIES ***");
  text = text.replace("*** Violations ***", "*** VIOLATIONS ***");
  text = text.replace("*** Docket Entries ***", "*** DOCKET ENTRIES ***");
  text = text.replace(/(\*\*\* )(Case Event Schedule|Docket Entries)( \*\*\*\n)(\n)/gi, "$1$2$3");
  
  text = text.replace(/(Images: ?\n\n)(\*\*\* Case Event Schedule \*\*\*)/gi, "\n\n$2");
  text = text.replace(/(\n\n\n\n)(\*\*\* Violations \*\*\*)/gi, "\n\n\n$2");
  text = text.replace(/\n\n(\d)\n/g, "\n$1\n");
  text = text.replace(/\n(\d{1,2}\/\d{1,2}\/\d{2,4})([^]+)(Case Parties)/gi, "$1$2$3");
  text = text.replace(/^(Violation: [0-9]{1,}\n)/gim, "\n$1");
  
  return text;
}



/////////////////////////

/*
 * Javascript Diff Algorithm
 *  By John Resig (http://ejohn.org/)
 *  Modified by Chu Alan "sprite"
 *
 * Released under the MIT license.
 *
 * More Info:
 *  http://ejohn.org/projects/javascript-diff-algorithm/
 */

(function(){function diff_match_patch(){this.Diff_Timeout=1;this.Diff_EditCost=4;this.Match_Threshold=0.5;this.Match_Distance=1E3;this.Patch_DeleteThreshold=0.5;this.Patch_Margin=4;this.Match_MaxBits=32}
diff_match_patch.prototype.diff_main=function(a,b,c,d){"undefined"==typeof d&&(d=0>=this.Diff_Timeout?Number.MAX_VALUE:(new Date).getTime()+1E3*this.Diff_Timeout);if(null==a||null==b)throw Error("Null input. (diff_main)");if(a==b)return a?[[0,a]]:[];"undefined"==typeof c&&(c=!0);var e=c,f=this.diff_commonPrefix(a,b);c=a.substring(0,f);a=a.substring(f);b=b.substring(f);var f=this.diff_commonSuffix(a,b),g=a.substring(a.length-f);a=a.substring(0,a.length-f);b=b.substring(0,b.length-f);a=this.diff_compute_(a,
b,e,d);c&&a.unshift([0,c]);g&&a.push([0,g]);this.diff_cleanupMerge(a);return a};
diff_match_patch.prototype.diff_compute_=function(a,b,c,d){if(!a)return[[1,b]];if(!b)return[[-1,a]];var e=a.length>b.length?a:b,f=a.length>b.length?b:a,g=e.indexOf(f);return-1!=g?(c=[[1,e.substring(0,g)],[0,f],[1,e.substring(g+f.length)]],a.length>b.length&&(c[0][0]=c[2][0]=-1),c):1==f.length?[[-1,a],[1,b]]:(e=this.diff_halfMatch_(a,b))?(f=e[0],a=e[1],g=e[2],b=e[3],e=e[4],f=this.diff_main(f,g,c,d),c=this.diff_main(a,b,c,d),f.concat([[0,e]],c)):c&&100<a.length&&100<b.length?this.diff_lineMode_(a,b,
d):this.diff_bisect_(a,b,d)};
diff_match_patch.prototype.diff_lineMode_=function(a,b,c){var d=this.diff_linesToChars_(a,b);a=d.chars1;b=d.chars2;d=d.lineArray;a=this.diff_main(a,b,!1,c);this.diff_charsToLines_(a,d);this.diff_cleanupSemantic(a);a.push([0,""]);for(var e=d=b=0,f="",g="";b<a.length;){switch(a[b][0]){case 1:e++;g+=a[b][1];break;case -1:d++;f+=a[b][1];break;case 0:if(1<=d&&1<=e){a.splice(b-d-e,d+e);b=b-d-e;d=this.diff_main(f,g,!1,c);for(e=d.length-1;0<=e;e--)a.splice(b,0,d[e]);b+=d.length}d=e=0;g=f=""}b++}a.pop();return a};
diff_match_patch.prototype.diff_bisect_=function(a,b,c){for(var d=a.length,e=b.length,f=Math.ceil((d+e)/2),g=f,h=2*f,j=Array(h),i=Array(h),k=0;k<h;k++)j[k]=-1,i[k]=-1;j[g+1]=0;i[g+1]=0;for(var k=d-e,q=0!=k%2,r=0,t=0,p=0,w=0,v=0;v<f&&!((new Date).getTime()>c);v++){for(var n=-v+r;n<=v-t;n+=2){var l=g+n,m;m=n==-v||n!=v&&j[l-1]<j[l+1]?j[l+1]:j[l-1]+1;for(var s=m-n;m<d&&s<e&&a.charAt(m)==b.charAt(s);)m++,s++;j[l]=m;if(m>d)t+=2;else if(s>e)r+=2;else if(q&&(l=g+k-n,0<=l&&l<h&&-1!=i[l])){var u=d-i[l];if(m>=
u)return this.diff_bisectSplit_(a,b,m,s,c)}}for(n=-v+p;n<=v-w;n+=2){l=g+n;u=n==-v||n!=v&&i[l-1]<i[l+1]?i[l+1]:i[l-1]+1;for(m=u-n;u<d&&m<e&&a.charAt(d-u-1)==b.charAt(e-m-1);)u++,m++;i[l]=u;if(u>d)w+=2;else if(m>e)p+=2;else if(!q&&(l=g+k-n,0<=l&&(l<h&&-1!=j[l])&&(m=j[l],s=g+m-l,u=d-u,m>=u)))return this.diff_bisectSplit_(a,b,m,s,c)}}return[[-1,a],[1,b]]};
diff_match_patch.prototype.diff_bisectSplit_=function(a,b,c,d,e){var f=a.substring(0,c),g=b.substring(0,d);a=a.substring(c);b=b.substring(d);f=this.diff_main(f,g,!1,e);e=this.diff_main(a,b,!1,e);return f.concat(e)};
diff_match_patch.prototype.diff_linesToChars_=function(a,b){function c(a){for(var b="",c=0,f=-1,g=d.length;f<a.length-1;){f=a.indexOf("\n",c);-1==f&&(f=a.length-1);var r=a.substring(c,f+1),c=f+1;(e.hasOwnProperty?e.hasOwnProperty(r):void 0!==e[r])?b+=String.fromCharCode(e[r]):(b+=String.fromCharCode(g),e[r]=g,d[g++]=r)}return b}var d=[],e={};d[0]="";var f=c(a),g=c(b);return{chars1:f,chars2:g,lineArray:d}};
diff_match_patch.prototype.diff_charsToLines_=function(a,b){for(var c=0;c<a.length;c++){for(var d=a[c][1],e=[],f=0;f<d.length;f++)e[f]=b[d.charCodeAt(f)];a[c][1]=e.join("")}};diff_match_patch.prototype.diff_commonPrefix=function(a,b){if(!a||!b||a.charAt(0)!=b.charAt(0))return 0;for(var c=0,d=Math.min(a.length,b.length),e=d,f=0;c<e;)a.substring(f,e)==b.substring(f,e)?f=c=e:d=e,e=Math.floor((d-c)/2+c);return e};
diff_match_patch.prototype.diff_commonSuffix=function(a,b){if(!a||!b||a.charAt(a.length-1)!=b.charAt(b.length-1))return 0;for(var c=0,d=Math.min(a.length,b.length),e=d,f=0;c<e;)a.substring(a.length-e,a.length-f)==b.substring(b.length-e,b.length-f)?f=c=e:d=e,e=Math.floor((d-c)/2+c);return e};
diff_match_patch.prototype.diff_commonOverlap_=function(a,b){var c=a.length,d=b.length;if(0==c||0==d)return 0;c>d?a=a.substring(c-d):c<d&&(b=b.substring(0,c));c=Math.min(c,d);if(a==b)return c;for(var d=0,e=1;;){var f=a.substring(c-e),f=b.indexOf(f);if(-1==f)return d;e+=f;if(0==f||a.substring(c-e)==b.substring(0,e))d=e,e++}};
diff_match_patch.prototype.diff_halfMatch_=function(a,b){function c(a,b,c){for(var d=a.substring(c,c+Math.floor(a.length/4)),e=-1,g="",h,j,n,l;-1!=(e=b.indexOf(d,e+1));){var m=f.diff_commonPrefix(a.substring(c),b.substring(e)),s=f.diff_commonSuffix(a.substring(0,c),b.substring(0,e));g.length<s+m&&(g=b.substring(e-s,e)+b.substring(e,e+m),h=a.substring(0,c-s),j=a.substring(c+m),n=b.substring(0,e-s),l=b.substring(e+m))}return 2*g.length>=a.length?[h,j,n,l,g]:null}if(0>=this.Diff_Timeout)return null;
var d=a.length>b.length?a:b,e=a.length>b.length?b:a;if(4>d.length||2*e.length<d.length)return null;var f=this,g=c(d,e,Math.ceil(d.length/4)),d=c(d,e,Math.ceil(d.length/2)),h;if(!g&&!d)return null;h=d?g?g[4].length>d[4].length?g:d:d:g;var j;a.length>b.length?(g=h[0],d=h[1],e=h[2],j=h[3]):(e=h[0],j=h[1],g=h[2],d=h[3]);h=h[4];return[g,d,e,j,h]};
diff_match_patch.prototype.diff_cleanupSemantic=function(a){for(var b=!1,c=[],d=0,e=null,f=0,g=0,h=0,j=0,i=0;f<a.length;)0==a[f][0]?(c[d++]=f,g=j,h=i,i=j=0,e=a[f][1]):(1==a[f][0]?j+=a[f][1].length:i+=a[f][1].length,e&&(e.length<=Math.max(g,h)&&e.length<=Math.max(j,i))&&(a.splice(c[d-1],0,[-1,e]),a[c[d-1]+1][0]=1,d--,d--,f=0<d?c[d-1]:-1,i=j=h=g=0,e=null,b=!0)),f++;b&&this.diff_cleanupMerge(a);this.diff_cleanupSemanticLossless(a);for(f=1;f<a.length;){if(-1==a[f-1][0]&&1==a[f][0]){b=a[f-1][1];c=a[f][1];
d=this.diff_commonOverlap_(b,c);e=this.diff_commonOverlap_(c,b);if(d>=e){if(d>=b.length/2||d>=c.length/2)a.splice(f,0,[0,c.substring(0,d)]),a[f-1][1]=b.substring(0,b.length-d),a[f+1][1]=c.substring(d),f++}else if(e>=b.length/2||e>=c.length/2)a.splice(f,0,[0,b.substring(0,e)]),a[f-1][0]=1,a[f-1][1]=c.substring(0,c.length-e),a[f+1][0]=-1,a[f+1][1]=b.substring(e),f++;f++}f++}};
diff_match_patch.prototype.diff_cleanupSemanticLossless=function(a){function b(a,b){if(!a||!b)return 6;var c=a.charAt(a.length-1),d=b.charAt(0),e=c.match(diff_match_patch.nonAlphaNumericRegex_),f=d.match(diff_match_patch.nonAlphaNumericRegex_),g=e&&c.match(diff_match_patch.whitespaceRegex_),h=f&&d.match(diff_match_patch.whitespaceRegex_),c=g&&c.match(diff_match_patch.linebreakRegex_),d=h&&d.match(diff_match_patch.linebreakRegex_),i=c&&a.match(diff_match_patch.blanklineEndRegex_),j=d&&b.match(diff_match_patch.blanklineStartRegex_);
return i||j?5:c||d?4:e&&!g&&h?3:g||h?2:e||f?1:0}for(var c=1;c<a.length-1;){if(0==a[c-1][0]&&0==a[c+1][0]){var d=a[c-1][1],e=a[c][1],f=a[c+1][1],g=this.diff_commonSuffix(d,e);if(g)var h=e.substring(e.length-g),d=d.substring(0,d.length-g),e=h+e.substring(0,e.length-g),f=h+f;for(var g=d,h=e,j=f,i=b(d,e)+b(e,f);e.charAt(0)===f.charAt(0);){var d=d+e.charAt(0),e=e.substring(1)+f.charAt(0),f=f.substring(1),k=b(d,e)+b(e,f);k>=i&&(i=k,g=d,h=e,j=f)}a[c-1][1]!=g&&(g?a[c-1][1]=g:(a.splice(c-1,1),c--),a[c][1]=
h,j?a[c+1][1]=j:(a.splice(c+1,1),c--))}c++}};diff_match_patch.nonAlphaNumericRegex_=/[^a-zA-Z0-9]/;diff_match_patch.whitespaceRegex_=/\s/;diff_match_patch.linebreakRegex_=/[\r\n]/;diff_match_patch.blanklineEndRegex_=/\n\r?\n$/;diff_match_patch.blanklineStartRegex_=/^\r?\n\r?\n/;
diff_match_patch.prototype.diff_cleanupEfficiency=function(a){for(var b=!1,c=[],d=0,e=null,f=0,g=!1,h=!1,j=!1,i=!1;f<a.length;){if(0==a[f][0])a[f][1].length<this.Diff_EditCost&&(j||i)?(c[d++]=f,g=j,h=i,e=a[f][1]):(d=0,e=null),j=i=!1;else if(-1==a[f][0]?i=!0:j=!0,e&&(g&&h&&j&&i||e.length<this.Diff_EditCost/2&&3==g+h+j+i))a.splice(c[d-1],0,[-1,e]),a[c[d-1]+1][0]=1,d--,e=null,g&&h?(j=i=!0,d=0):(d--,f=0<d?c[d-1]:-1,j=i=!1),b=!0;f++}b&&this.diff_cleanupMerge(a)};
diff_match_patch.prototype.diff_cleanupMerge=function(a){a.push([0,""]);for(var b=0,c=0,d=0,e="",f="",g;b<a.length;)switch(a[b][0]){case 1:d++;f+=a[b][1];b++;break;case -1:c++;e+=a[b][1];b++;break;case 0:1<c+d?(0!==c&&0!==d&&(g=this.diff_commonPrefix(f,e),0!==g&&(0<b-c-d&&0==a[b-c-d-1][0]?a[b-c-d-1][1]+=f.substring(0,g):(a.splice(0,0,[0,f.substring(0,g)]),b++),f=f.substring(g),e=e.substring(g)),g=this.diff_commonSuffix(f,e),0!==g&&(a[b][1]=f.substring(f.length-g)+a[b][1],f=f.substring(0,f.length-
g),e=e.substring(0,e.length-g))),0===c?a.splice(b-d,c+d,[1,f]):0===d?a.splice(b-c,c+d,[-1,e]):a.splice(b-c-d,c+d,[-1,e],[1,f]),b=b-c-d+(c?1:0)+(d?1:0)+1):0!==b&&0==a[b-1][0]?(a[b-1][1]+=a[b][1],a.splice(b,1)):b++,c=d=0,f=e=""}""===a[a.length-1][1]&&a.pop();c=!1;for(b=1;b<a.length-1;)0==a[b-1][0]&&0==a[b+1][0]&&(a[b][1].substring(a[b][1].length-a[b-1][1].length)==a[b-1][1]?(a[b][1]=a[b-1][1]+a[b][1].substring(0,a[b][1].length-a[b-1][1].length),a[b+1][1]=a[b-1][1]+a[b+1][1],a.splice(b-1,1),c=!0):a[b][1].substring(0,
a[b+1][1].length)==a[b+1][1]&&(a[b-1][1]+=a[b+1][1],a[b][1]=a[b][1].substring(a[b+1][1].length)+a[b+1][1],a.splice(b+1,1),c=!0)),b++;c&&this.diff_cleanupMerge(a)};diff_match_patch.prototype.diff_xIndex=function(a,b){var c=0,d=0,e=0,f=0,g;for(g=0;g<a.length;g++){1!==a[g][0]&&(c+=a[g][1].length);-1!==a[g][0]&&(d+=a[g][1].length);if(c>b)break;e=c;f=d}return a.length!=g&&-1===a[g][0]?f:f+(b-e)};
diff_match_patch.prototype.diff_prettyHtml=function(a){for(var b=[],c=/&/g,d=/</g,e=/>/g,f=/\n/g,g=0;g<a.length;g++){var h=a[g][0],j=a[g][1],j=j.replace(c,"&amp;").replace(d,"&lt;").replace(e,"&gt;").replace(f,"&para;<br>");switch(h){case 1:b[g]='<ins style="background:#e6ffe6;">'+j+"</ins>";break;case -1:b[g]='<del style="background:#ffe6e6;">'+j+"</del>";break;case 0:b[g]="<span>"+j+"</span>"}}return b.join("")};
diff_match_patch.prototype.diff_text1=function(a){for(var b=[],c=0;c<a.length;c++)1!==a[c][0]&&(b[c]=a[c][1]);return b.join("")};diff_match_patch.prototype.diff_text2=function(a){for(var b=[],c=0;c<a.length;c++)-1!==a[c][0]&&(b[c]=a[c][1]);return b.join("")};diff_match_patch.prototype.diff_levenshtein=function(a){for(var b=0,c=0,d=0,e=0;e<a.length;e++){var f=a[e][0],g=a[e][1];switch(f){case 1:c+=g.length;break;case -1:d+=g.length;break;case 0:b+=Math.max(c,d),d=c=0}}return b+=Math.max(c,d)};
diff_match_patch.prototype.diff_toDelta=function(a){for(var b=[],c=0;c<a.length;c++)switch(a[c][0]){case 1:b[c]="+"+encodeURI(a[c][1]);break;case -1:b[c]="-"+a[c][1].length;break;case 0:b[c]="="+a[c][1].length}return b.join("\t").replace(/%20/g," ")};
diff_match_patch.prototype.diff_fromDelta=function(a,b){for(var c=[],d=0,e=0,f=b.split(/\t/g),g=0;g<f.length;g++){var h=f[g].substring(1);switch(f[g].charAt(0)){case "+":try{c[d++]=[1,decodeURI(h)]}catch(j){throw Error("Illegal escape in diff_fromDelta: "+h);}break;case "-":case "=":var i=parseInt(h,10);if(isNaN(i)||0>i)throw Error("Invalid number in diff_fromDelta: "+h);h=a.substring(e,e+=i);"="==f[g].charAt(0)?c[d++]=[0,h]:c[d++]=[-1,h];break;default:if(f[g])throw Error("Invalid diff operation in diff_fromDelta: "+
f[g]);}}if(e!=a.length)throw Error("Delta length ("+e+") does not equal source text length ("+a.length+").");return c};diff_match_patch.prototype.match_main=function(a,b,c){if(null==a||null==b||null==c)throw Error("Null input. (match_main)");c=Math.max(0,Math.min(c,a.length));return a==b?0:a.length?a.substring(c,c+b.length)==b?c:this.match_bitap_(a,b,c):-1};
diff_match_patch.prototype.match_bitap_=function(a,b,c){function d(a,d){var e=a/b.length,g=Math.abs(c-d);return!f.Match_Distance?g?1:e:e+g/f.Match_Distance}if(b.length>this.Match_MaxBits)throw Error("Pattern too long for this browser.");var e=this.match_alphabet_(b),f=this,g=this.Match_Threshold,h=a.indexOf(b,c);-1!=h&&(g=Math.min(d(0,h),g),h=a.lastIndexOf(b,c+b.length),-1!=h&&(g=Math.min(d(0,h),g)));for(var j=1<<b.length-1,h=-1,i,k,q=b.length+a.length,r,t=0;t<b.length;t++){i=0;for(k=q;i<k;)d(t,c+
k)<=g?i=k:q=k,k=Math.floor((q-i)/2+i);q=k;i=Math.max(1,c-k+1);var p=Math.min(c+k,a.length)+b.length;k=Array(p+2);for(k[p+1]=(1<<t)-1;p>=i;p--){var w=e[a.charAt(p-1)];k[p]=0===t?(k[p+1]<<1|1)&w:(k[p+1]<<1|1)&w|((r[p+1]|r[p])<<1|1)|r[p+1];if(k[p]&j&&(w=d(t,p-1),w<=g))if(g=w,h=p-1,h>c)i=Math.max(1,2*c-h);else break}if(d(t+1,c)>g)break;r=k}return h};
diff_match_patch.prototype.match_alphabet_=function(a){for(var b={},c=0;c<a.length;c++)b[a.charAt(c)]=0;for(c=0;c<a.length;c++)b[a.charAt(c)]|=1<<a.length-c-1;return b};
diff_match_patch.prototype.patch_addContext_=function(a,b){if(0!=b.length){for(var c=b.substring(a.start2,a.start2+a.length1),d=0;b.indexOf(c)!=b.lastIndexOf(c)&&c.length<this.Match_MaxBits-this.Patch_Margin-this.Patch_Margin;)d+=this.Patch_Margin,c=b.substring(a.start2-d,a.start2+a.length1+d);d+=this.Patch_Margin;(c=b.substring(a.start2-d,a.start2))&&a.diffs.unshift([0,c]);(d=b.substring(a.start2+a.length1,a.start2+a.length1+d))&&a.diffs.push([0,d]);a.start1-=c.length;a.start2-=c.length;a.length1+=
c.length+d.length;a.length2+=c.length+d.length}};
diff_match_patch.prototype.patch_make=function(a,b,c){var d;if("string"==typeof a&&"string"==typeof b&&"undefined"==typeof c)d=a,b=this.diff_main(d,b,!0),2<b.length&&(this.diff_cleanupSemantic(b),this.diff_cleanupEfficiency(b));else if(a&&"object"==typeof a&&"undefined"==typeof b&&"undefined"==typeof c)b=a,d=this.diff_text1(b);else if("string"==typeof a&&b&&"object"==typeof b&&"undefined"==typeof c)d=a;else if("string"==typeof a&&"string"==typeof b&&c&&"object"==typeof c)d=a,b=c;else throw Error("Unknown call format to patch_make.");
if(0===b.length)return[];c=[];a=new diff_match_patch.patch_obj;for(var e=0,f=0,g=0,h=d,j=0;j<b.length;j++){var i=b[j][0],k=b[j][1];!e&&0!==i&&(a.start1=f,a.start2=g);switch(i){case 1:a.diffs[e++]=b[j];a.length2+=k.length;d=d.substring(0,g)+k+d.substring(g);break;case -1:a.length1+=k.length;a.diffs[e++]=b[j];d=d.substring(0,g)+d.substring(g+k.length);break;case 0:k.length<=2*this.Patch_Margin&&e&&b.length!=j+1?(a.diffs[e++]=b[j],a.length1+=k.length,a.length2+=k.length):k.length>=2*this.Patch_Margin&&
e&&(this.patch_addContext_(a,h),c.push(a),a=new diff_match_patch.patch_obj,e=0,h=d,f=g)}1!==i&&(f+=k.length);-1!==i&&(g+=k.length)}e&&(this.patch_addContext_(a,h),c.push(a));return c};diff_match_patch.prototype.patch_deepCopy=function(a){for(var b=[],c=0;c<a.length;c++){var d=a[c],e=new diff_match_patch.patch_obj;e.diffs=[];for(var f=0;f<d.diffs.length;f++)e.diffs[f]=d.diffs[f].slice();e.start1=d.start1;e.start2=d.start2;e.length1=d.length1;e.length2=d.length2;b[c]=e}return b};
diff_match_patch.prototype.patch_apply=function(a,b){if(0==a.length)return[b,[]];a=this.patch_deepCopy(a);var c=this.patch_addPadding(a);b=c+b+c;this.patch_splitMax(a);for(var d=0,e=[],f=0;f<a.length;f++){var g=a[f].start2+d,h=this.diff_text1(a[f].diffs),j,i=-1;if(h.length>this.Match_MaxBits){if(j=this.match_main(b,h.substring(0,this.Match_MaxBits),g),-1!=j&&(i=this.match_main(b,h.substring(h.length-this.Match_MaxBits),g+h.length-this.Match_MaxBits),-1==i||j>=i))j=-1}else j=this.match_main(b,h,g);
if(-1==j)e[f]=!1,d-=a[f].length2-a[f].length1;else if(e[f]=!0,d=j-g,g=-1==i?b.substring(j,j+h.length):b.substring(j,i+this.Match_MaxBits),h==g)b=b.substring(0,j)+this.diff_text2(a[f].diffs)+b.substring(j+h.length);else if(g=this.diff_main(h,g,!1),h.length>this.Match_MaxBits&&this.diff_levenshtein(g)/h.length>this.Patch_DeleteThreshold)e[f]=!1;else{this.diff_cleanupSemanticLossless(g);for(var h=0,k,i=0;i<a[f].diffs.length;i++){var q=a[f].diffs[i];0!==q[0]&&(k=this.diff_xIndex(g,h));1===q[0]?b=b.substring(0,
j+k)+q[1]+b.substring(j+k):-1===q[0]&&(b=b.substring(0,j+k)+b.substring(j+this.diff_xIndex(g,h+q[1].length)));-1!==q[0]&&(h+=q[1].length)}}}b=b.substring(c.length,b.length-c.length);return[b,e]};
diff_match_patch.prototype.patch_addPadding=function(a){for(var b=this.Patch_Margin,c="",d=1;d<=b;d++)c+=String.fromCharCode(d);for(d=0;d<a.length;d++)a[d].start1+=b,a[d].start2+=b;var d=a[0],e=d.diffs;if(0==e.length||0!=e[0][0])e.unshift([0,c]),d.start1-=b,d.start2-=b,d.length1+=b,d.length2+=b;else if(b>e[0][1].length){var f=b-e[0][1].length;e[0][1]=c.substring(e[0][1].length)+e[0][1];d.start1-=f;d.start2-=f;d.length1+=f;d.length2+=f}d=a[a.length-1];e=d.diffs;0==e.length||0!=e[e.length-1][0]?(e.push([0,
c]),d.length1+=b,d.length2+=b):b>e[e.length-1][1].length&&(f=b-e[e.length-1][1].length,e[e.length-1][1]+=c.substring(0,f),d.length1+=f,d.length2+=f);return c};
diff_match_patch.prototype.patch_splitMax=function(a){for(var b=this.Match_MaxBits,c=0;c<a.length;c++)if(!(a[c].length1<=b)){var d=a[c];a.splice(c--,1);for(var e=d.start1,f=d.start2,g="";0!==d.diffs.length;){var h=new diff_match_patch.patch_obj,j=!0;h.start1=e-g.length;h.start2=f-g.length;""!==g&&(h.length1=h.length2=g.length,h.diffs.push([0,g]));for(;0!==d.diffs.length&&h.length1<b-this.Patch_Margin;){var g=d.diffs[0][0],i=d.diffs[0][1];1===g?(h.length2+=i.length,f+=i.length,h.diffs.push(d.diffs.shift()),
j=!1):-1===g&&1==h.diffs.length&&0==h.diffs[0][0]&&i.length>2*b?(h.length1+=i.length,e+=i.length,j=!1,h.diffs.push([g,i]),d.diffs.shift()):(i=i.substring(0,b-h.length1-this.Patch_Margin),h.length1+=i.length,e+=i.length,0===g?(h.length2+=i.length,f+=i.length):j=!1,h.diffs.push([g,i]),i==d.diffs[0][1]?d.diffs.shift():d.diffs[0][1]=d.diffs[0][1].substring(i.length))}g=this.diff_text2(h.diffs);g=g.substring(g.length-this.Patch_Margin);i=this.diff_text1(d.diffs).substring(0,this.Patch_Margin);""!==i&&
(h.length1+=i.length,h.length2+=i.length,0!==h.diffs.length&&0===h.diffs[h.diffs.length-1][0]?h.diffs[h.diffs.length-1][1]+=i:h.diffs.push([0,i]));j||a.splice(++c,0,h)}}};diff_match_patch.prototype.patch_toText=function(a){for(var b=[],c=0;c<a.length;c++)b[c]=a[c];return b.join("")};
diff_match_patch.prototype.patch_fromText=function(a){var b=[];if(!a)return b;a=a.split("\n");for(var c=0,d=/^@@ -(\d+),?(\d*) \+(\d+),?(\d*) @@$/;c<a.length;){var e=a[c].match(d);if(!e)throw Error("Invalid patch string: "+a[c]);var f=new diff_match_patch.patch_obj;b.push(f);f.start1=parseInt(e[1],10);""===e[2]?(f.start1--,f.length1=1):"0"==e[2]?f.length1=0:(f.start1--,f.length1=parseInt(e[2],10));f.start2=parseInt(e[3],10);""===e[4]?(f.start2--,f.length2=1):"0"==e[4]?f.length2=0:(f.start2--,f.length2=
parseInt(e[4],10));for(c++;c<a.length;){e=a[c].charAt(0);try{var g=decodeURI(a[c].substring(1))}catch(h){throw Error("Illegal escape in patch_fromText: "+g);}if("-"==e)f.diffs.push([-1,g]);else if("+"==e)f.diffs.push([1,g]);else if(" "==e)f.diffs.push([0,g]);else if("@"==e)break;else if(""!==e)throw Error('Invalid patch mode "'+e+'" in: '+g);c++}}return b};diff_match_patch.patch_obj=function(){this.diffs=[];this.start2=this.start1=null;this.length2=this.length1=0};
diff_match_patch.patch_obj.prototype.toString=function(){var a,b;a=0===this.length1?this.start1+",0":1==this.length1?this.start1+1:this.start1+1+","+this.length1;b=0===this.length2?this.start2+",0":1==this.length2?this.start2+1:this.start2+1+","+this.length2;a=["@@ -"+a+" +"+b+" @@\n"];var c;for(b=0;b<this.diffs.length;b++){switch(this.diffs[b][0]){case 1:c="+";break;case -1:c="-";break;case 0:c=" "}a[b+1]=c+encodeURI(this.diffs[b][1])+"\n"}return a.join("").replace(/%20/g," ")};
this.diff_match_patch=diff_match_patch;this.DIFF_DELETE=-1;this.DIFF_INSERT=1;this.DIFF_EQUAL=0;})()

function escape_(s) {
    var n = s;
    n = n.replace(/&/g, "&amp;");
    n = n.replace(/</g, "&lt;");
    n = n.replace(/>/g, "&gt;");
    n = n.replace(/"/g, "&quot;");

    return n;
}

function diffString_( o, n ) {
  o = o.replace(/\s+$/, '');
  n = n.replace(/\s+$/, '');

  var out = diff_(o == "" ? [] : o.split(/\s+/), n == "" ? [] : n.split(/\s+/) );
  var str = "";

  var oSpace = o.match(/\s+/g);
  if (oSpace == null) {
    oSpace = ["\n"];
  } else {
    oSpace.push("\n");
  }
  var nSpace = n.match(/\s+/g);
  if (nSpace == null) {
    nSpace = ["\n"];
  } else {
    nSpace.push("\n");
  }

  if (out.n.length == 0) {
      for (var i = 0; i < out.o.length; i++) {
        str += '<del>' + escape_(out.o[i]) + oSpace[i] + "</del>";
      }
  } else {
    if (out.n[0].text == null) {
      for (n = 0; n < out.o.length && out.o[n].text == null; n++) {
        str += '<del>' + escape_(out.o[n]) + oSpace[n] + "</del>";
      }
    }

    for ( var i = 0; i < out.n.length; i++ ) {
      if (out.n[i].text == null) {
        str += '<ins>' + escape_(out.n[i]) + nSpace[i] + "</ins>";
      } else {
        var pre = "";

        for (n = out.n[i].row + 1; n < out.o.length && out.o[n].text == null; n++ ) {
          pre += '<del>' + escape_(out.o[n]) + oSpace[n] + "</del>";
        }
        str += " " + out.n[i].text + nSpace[i] + pre;
      }
    }
  }
  
  return str;
}

function diff_( o, n ) {
  var ns = new Object();
  var os = new Object();
  
  for ( var i = 0; i < n.length; i++ ) {
    if ( ns[ n[i] ] == null )
      ns[ n[i] ] = { rows: new Array(), o: null };
    ns[ n[i] ].rows.push( i );
  }
  
  for ( var i = 0; i < o.length; i++ ) {
    if ( os[ o[i] ] == null )
      os[ o[i] ] = { rows: new Array(), n: null };
    os[ o[i] ].rows.push( i );
  }
  
  for ( var i in ns ) {
    if ( ns[i].rows.length == 1 && typeof(os[i]) != "undefined" && os[i].rows.length == 1 ) {
      n[ ns[i].rows[0] ] = { text: n[ ns[i].rows[0] ], row: os[i].rows[0] };
      o[ os[i].rows[0] ] = { text: o[ os[i].rows[0] ], row: ns[i].rows[0] };
    }
  }
  
  for ( var i = 0; i < n.length - 1; i++ ) {
    if ( n[i].text != null && n[i+1].text == null && n[i].row + 1 < o.length && o[ n[i].row + 1 ].text == null && 
         n[i+1] == o[ n[i].row + 1 ] ) {
      n[i+1] = { text: n[i+1], row: n[i].row + 1 };
      o[n[i].row+1] = { text: o[n[i].row+1], row: i + 1 };
    }
  }
  
  for ( var i = n.length - 1; i > 0; i-- ) {
    if ( n[i].text != null && n[i-1].text == null && n[i].row > 0 && o[ n[i].row - 1 ].text == null && 
         n[i-1] == o[ n[i].row - 1 ] ) {
      n[i-1] = { text: n[i-1], row: n[i].row - 1 };
      o[n[i].row-1] = { text: o[n[i].row-1], row: i - 1 };
    }
  }
  
  return { o: o, n: n };
}



// The rest of this code is currently (c) Google Inc.
 
// setRowsData fills in one row of data per object defined in the objects Array.
// For every Column, it checks if data objects define a value for it.
// Arguments:
//   - sheet: the Sheet Object where the data will be written
//   - objects: an Array of Objects, each of which contains data for a row
//   - optHeadersRange: a Range of cells where the column headers are defined. This
//     defaults to the entire first row in sheet.
//   - optFirstDataRowIndex: index of the first row where data should be written. This
//     defaults to the row immediately below the headers.
function setRowsData_(sheet, objects, optHeadersRange, optFirstDataRowIndex) {
  var headersRange = optHeadersRange || sheet.getRange(1, 1, 1, sheet.getMaxColumns());
  var firstDataRowIndex = optFirstDataRowIndex || headersRange.getRowIndex() + 1;
  var headers = normalizeHeaders_(headersRange.getValues()[0]);
 
  var data = [];
  for (var i = 0; i < objects.length; ++i) {
    var values = []
    for (j = 0; j < headers.length; ++j) {
      var header = headers[j];
      values.push(header.length > 0 && objects[i][header] ? objects[i][header] : "");
    }
    data.push(values);
  }
 
  var destinationRange = sheet.getRange(firstDataRowIndex, headersRange.getColumnIndex(), 
                                        objects.length, headers.length);
  destinationRange.setValues(data);
}

// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//       This argument is optional and it defaults to all the cells except those in the first row
//       or all the cells below columnHeadersRowIndex (if defined).
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range; 
// Returns an Array of objects.
function getRowsData_(sheet, range, columnHeadersRowIndex) {
  var headersIndex = columnHeadersRowIndex || 1;
  var dataRange = range || 
    sheet.getRange(headersIndex + 1, 1, sheet.getMaxRows() - headersIndex, sheet.getMaxColumns());
  var numColumns = dataRange.getEndColumn() - dataRange.getColumn() + 1;
  var headersRange = sheet.getRange(headersIndex, dataRange.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects_(dataRange.getValues(), normalizeHeaders_(headers));
}

// Given a JavaScript 2d Array, this function returns the transposed table.
// Arguments:
//   - data: JavaScript 2d Array
// Returns a JavaScript 2d Array
// Example: arrayTranspose_([[1,2,3],[4,5,6]]) returns [[1,4],[2,5],[3,6]].
function arrayTranspose_(data) {
  if (data.length == 0 || data[0].length == 0) {
    return null;
  }
 
  var ret = [];
  for (var i = 0; i < data[0].length; ++i) {
    ret.push([]);
  }
 
  for (var i = 0; i < data.length; ++i) {
    for (var j = 0; j < data[i].length; ++j) {
      ret[j][i] = data[i][j];
    }
  }
 
  return ret;
}
 
// getColumnsData iterates column by column in the input range and returns an array of objects.
// Each object contains all the data for a given column, indexed by its normalized row name.
// [Modified by mhawksey]
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//       This argument is optional and it defaults to all the cells except those in the first column
//       or all the cells below rowHeadersIndex (if defined).
//   - columnHeadersIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range; 
// Returns an Array of objects.
function getColumnsData_(sheet, range, columnHeadersIndex) {
  var headersIndex = columnHeadersIndex || range ? range.getColumnIndex() - 1 : 1;
  var dataRange = range || 
    sheet.getRange(1, headersIndex + 1, sheet.getMaxRows(), sheet.getMaxColumns()- headersIndex);
  var numRows = dataRange.getLastRow() - dataRange.getRow() + 1;
  var headersRange = sheet.getRange(dataRange.getRow(),headersIndex,numRows,1);
  var headers = arrayTranspose_(headersRange.getValues())[0];
  return getObjects_(arrayTranspose_(dataRange.getValues()), normalizeHeaders_(headers));
}
 
// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects_(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty_(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}
 
// Returns an Array of normalized Strings. 
// Empty Strings are returned for all Strings that could not be successfully normalized.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders_(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    keys.push(normalizeHeader_(headers[i]));
  }
  return keys;
}
 
// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader_(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    //if (!isAlnum_(letter)) { // I removed this because result identifiers have '_' in name
    //  continue;
    //}
    if (key.length == 0 && isDigit_(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}
 
// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty_(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}
 
// Returns true if the character char is alphabetical, false otherwise.
function isAlnum_(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit_(char);
}
 
// Returns true if the character char is a digit, false otherwise.
function isDigit_(char) {
  return char >= '0' && char <= '9';
}
 
function alltrim_(str) {
  if (str != null) return str.replace(/^\s+|\s+$/g, '');
}
