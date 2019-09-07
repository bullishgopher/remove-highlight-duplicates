var DIALOG_TITLE = 'Remove And Highlight Duplicates';

/**
 * For selected range type
 * 
 * 0 : custom select
 * 1 : auto select
 */

//var range_select = 0;
//PropertiesService.getUserProperties().setProperty('range_select', 0);
//var myvalue = PropertiesService.getUserProperties().getProperty('mykey');

/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Remove and Highlight', 'showDialog')
      .addToUi();
}

function onMyTrigger(){
  //Logger.log("onMyTrigger");
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  //Logger.log(range);
  //Logger.log(range.getColumn());
  var range_arr = [range.getColumn(), range.getRow(), range.getLastColumn(), range.getLastRow()];
  return range_arr;
}
/*function onChange(e) {
  Logger.log("onChange");
}
function onEdit(e) {
  Logger.log("onEdit");
}*/
/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a dialog. The dialog structure is described in the Dialog.html
 * project file.
 */
function showDialog() {
  var ui = HtmlService.createTemplateFromFile('Dialog')
      .evaluate()
      .setWidth(600)
      .setHeight(600)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(ui, DIALOG_TITLE);
}

/**
 * Returns the value in the active cell.
 *
 * @return {String} The value of the active cell.
 */
function getActiveValue() {
  // Retrieve and return the information requested by the sidebar.
  var cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  return cell.getValue();
}

/**
 * Replaces the active cell value with the given value.
 *
 * @param {Number} value A reference number to replace with.
 */
function setActiveValue(value) {
  // Use data collected from sidebar to manipulate the sheet.
  var cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  cell.setValue(value);
}

/**
 * Executes the specified action (create a new sheet, copy the active sheet, or
 * clear the current sheet).
 *
 * @param {String} action An identifier for the action to take.
 */
function modifySheets(action) {
  // Use data collected from dialog to manipulate the spreadsheet.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = ss.getActiveSheet();
  if(Array.isArray(action))
  {/*
    if (action[0] == "sort") {
      if (action[1] == 0)
      {
        return "None";
      }
      else
      {
        return sortSpreadSheetA(action[1]);
      }
    } else */
    if (action[0] == "remove") {
      remove_duplicate(action);
    } else if (action[0] == "remove_populate") {
      remove_duplicate_and_copy(action);
    } else if (action[0] == "highlight") {
      highlight_duplicate(action);
    } else if (action[0] == "highlight_populate")
    {
      highlight_duplicate_and_copy(action);
    } else if (action[0] == "show_status_dlg"){
    if (action[1] == "Can't get titles")
    {
      var ui = HtmlService.createTemplateFromFile('StatusDialog')
      .evaluate()
      .append(
        '<div class="error-message"><div class="block"><p id="err_content_p">' + action[2] + '</p></div><div class="block" id="dialog-button-bar"><button onclick="google.script.host.close()">OK</button></div></div>'
      )
      .setWidth(400)
      .setHeight(100)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModelessDialog(ui, action[1]);
      return true
    }
    var ui = HtmlService.createTemplateFromFile('StatusDialog')
      .evaluate()
      .append(
        '<div class="error-message"><div class="block"><p id="err_content_p">' + action[2] + '</p></div><div class="block" id="dialog-button-bar"><button id="dialog-ok-button">OK</button></div></div>'
      )
      .setWidth(400)
      .setHeight(100)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModelessDialog(ui, action[1]);
    }
    
  }
  if (action == "get_auto_range") {
    //range_select = 1;
    PropertiesService.getUserProperties().setProperty('range_select', 1);
    return getAutoRange();
  } else if (action == "get_current_range"){
    return getCurrentRange();
  } else if (action == "get_current_range_by_clicking_table_icon"){
    var ui = HtmlService.createTemplateFromFile('RangeSelectDialog')
      .evaluate()
      .setWidth(400)
      .setHeight(200)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModelessDialog(ui, "Range Select");
    //SpreadsheetApp.getUi().showModalDialog(ui, "Range Select");
  } else if (action == "get_titles") {
    return getTitles();
  } else if (action == "get_titles_and_first_row"){
    return getTitlesAndFristRow();
  } else if (action == "cancel_dialog"){
    PropertiesService.getUserProperties().setProperty('range_select', 0);
  } else if (action == "ok_dialog2"){
    var ui = HtmlService.createTemplateFromFile('Dialog')
      .evaluate()
      .setWidth(600)
      .setHeight(600)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(ui, DIALOG_TITLE);
  } else if (action == "cancel_dialog2"){
    var ui = HtmlService.createTemplateFromFile('Dialog')
      .evaluate()
      .setWidth(600)
      .setHeight(600)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(ui, DIALOG_TITLE);
  } else if (action == "on_timer"){
    return onMyTrigger();
  } else if (action == "on_ok_err_dlg"){
    var ui = HtmlService.createTemplateFromFile('Dialog')
      .evaluate()
      .setWidth(600)
      .setHeight(600)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(ui, DIALOG_TITLE);
  }
}
function remove_duplicate(action) {
  //var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var range;
  var range_select = PropertiesService.getUserProperties().getProperty('range_select');
  if (range_select == 1)
  {
    range = sheet.getRange(2,1, sheet.getLastRow(), sheet.getLastColumn());
  }
  else{
    range = sheet.getActiveRange();
  }
  var data = range.getValues();

  var rowNum = range.getRow();
  var columnNum = range.getColumn();
  var columnLength = data[0].length;

  var uniqueData = [];
  
  var firstCol = columnNum;
  var lastCol = range.getLastColumn();
  //var dataLen = data.length;
  //var duplicateData = [];

  // iterate through each 'row' of the selected range
  // x is
  // y is
  var x = 0;
  var y = data.length;
  var duplicateRange=[];
  var dup_rge_cnt=0;
  // when row is
  while (x < y) {
    Logger.log(x + " " + y);
    var row = data[x];
    var is_empty = 0;
    for ( var stosic = 1; stosic <= columnLength; stosic++)
    {
      if (row[stosic-1]==0)
      {
        is_empty += 1;
      }
    }
    if (is_empty == columnLength)
    {
      x--;
      y--;
      continue;
    }
    for ( var stosic2 = 1; stosic2 <= lastCol; stosic2++)
    {
      Logger.log(row[stosic2-1]);
      if (action[stosic2]==0)
      {
        Logger.log(row[stosic2-1]);
        row[stosic2-1] = "";
      }
    }
    var duplicate = false;

    // iterate through the uniqueData array to see if 'row' already exists
    for (var j = 0; j < uniqueData.length; j++) {
      if (row.join() == uniqueData[j].join()) {
        // if there is a duplicate, delete the 'row' from the sheet and add it to the duplicateData array
        duplicate = true;
        duplicateRange[dup_rge_cnt] = rowNum + x;
        //duplicateRange[dup_rge_cnt] = sheet.getRange(
        //  rowNum + x,
       //   columnNum,
        //  1,
        //  columnLength
        //);
        dup_rge_cnt++;
        //duplicateRange.deleteCells(SpreadsheetApp.Dimension.ROWS);
        //duplicateData.push(row);

        // rows shift up by one when duplicate is deleted
        // in effect, it skips a line
        // so we need to decrement x to stay in the same line
        x--;
        y--;
        //range = sheet.getActiveRange();
        //data = range.getValues();
        // return;
      }
    }

    // if there are no duplicates, add 'row' to the uniqueData array
    if (!duplicate) {
      uniqueData.push(row);
    }
    x++;
  }
  if (dup_rge_cnt)
  {
    for(var kk=0;kk<dup_rge_cnt;kk++)
    {
      //duplicateRange[kk].deleteCells(SpreadsheetApp.Dimension.ROWS);
      sheet.deleteRow(duplicateRange[kk]);
    }
  }
}

function remove_duplicate_and_copy(action) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var range;
  var range_select = PropertiesService.getUserProperties().getProperty('range_select');
  if (range_select == 1)
  {
    range = sheet.getRange(2,1, sheet.getLastRow(), sheet.getLastColumn());
  }
  else{
    range = sheet.getActiveRange();
  }
  var data = range.getValues();

  var rowNum = range.getRow();
  var columnNum = range.getColumn();
  var columnLength = data[0].length;

  var uniqueData = [];
  var duplicateData = [];
  
  var firstCol = range.getColumn();
  var lastCol = range.getLastColumn();
  var dataLen = data.length;
  // iterate through each 'row' of the selected range
  // x is
  // y is
  var x = 0;
  var y = data.length;
  
  
  var duplicateRange=[];
  var dup_rge_cnt=0;
  
  // when row is
  while (x < y) {
    var row = data[x];
    var is_empty=0;
    for ( var stosic = firstCol; stosic <= lastCol; stosic++)
    {
      if (row[stosic-1]==0)
      {
        is_empty += 1;
      }
    }
    if (is_empty == lastCol - firstCol + 1)
    {
      x--;
      y--;
      continue;
    }
    for ( var stosic = firstCol; stosic <= lastCol; stosic++)
    {
      if (action[stosic]==0)
      {
        row[stosic-1] = "";
      }
    }
    var duplicate = false;

    // iterate through the uniqueData array to see if 'row' already exists
    for (var j = 0; j < uniqueData.length; j++) {
      if (row.join() == uniqueData[j].join()) {
        // if there is a duplicate, delete the 'row' from the sheet and add it to the duplicateData array
        duplicate = true;
        var duplicateRange2 = sheet.getRange(
          rowNum + x,
          columnNum,
          1,
          columnLength
        );
        duplicateRange[dup_rge_cnt] = rowNum + x;
        dup_rge_cnt++;
        //duplicateRange.deleteCells(SpreadsheetApp.Dimension.ROWS);
        var tmp = duplicateRange2.getValues();
        duplicateData.push(tmp[0]);

        // rows shift up by one when duplicate is deleted
        // in effect, it skips a line
        // so we need to decrement x to stay in the same line
        x--;
        y--;
        //range = sheet.getActiveRange();
        //data = range.getValues();
        // return;
      }
    }

    // if there are no duplicates, add 'row' to the uniqueData array
    if (!duplicate) {
      uniqueData.push(row);
    }
    x++;
  }
  // remove all rows duplicates
  if (dup_rge_cnt)
  {
    for(var kk=0;kk<dup_rge_cnt;kk++)
    {
      //duplicateRange[kk].deleteCells(SpreadsheetApp.Dimension.ROWS);
      sheet.deleteRow(duplicateRange[kk]);
    }
  }
  // create a new sheet with the duplicate data
  if (duplicateData) {
    var newSheet = spreadsheet.insertSheet();
    var duplicateDataLen = duplicateData.length;
    var header = duplicateDataLen + ' duplicates found';
    newSheet.setName('Duplicates v' + (spreadsheet.getSheets().length - 1));
    newSheet.appendRow([header]);

    for (var k = 0; k < duplicateDataLen; k++) {
      newSheet.appendRow(duplicateData[k]);
    }
  }
}

function highlight_duplicate(action) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range;
  var range_select = PropertiesService.getUserProperties().getProperty('range_select');
  if (range_select == 1)
  {
    range = sheet.getRange(2,1, sheet.getLastRow(), sheet.getLastColumn());
  }
  else{
    range = sheet.getActiveRange();
  }
  var data = range.getValues();

  var rowNum = range.getRow();
  var columnNum = range.getColumn();
  var columnLength = data[0].length;

  var uniqueData = [];
  var firstCol = range.getColumn();
  var lastCol = range.getLastColumn();
  var dataLen = data.length;
  // iterate through each 'row' of the selected range
  for (var i = 0; i < dataLen; i++) {
    var row = data[i];
    var is_empty=0;
    for ( var stosic = firstCol; stosic <= lastCol; stosic++)
    {
      if (row[stosic-1]==0)
      {
        is_empty += 1;
      }
    }
    if (is_empty == lastCol - firstCol + 1)
    {
      continue;
    }
    for ( var stosic = firstCol; stosic <= lastCol; stosic++)
    {
      if (action[stosic]==0)
      {
        row[stosic-1] = "";
      }
    }
    var duplicate = false;

    // iterate through the uniqueData array to see if 'row' already exists
    for (var j = 0; j < uniqueData.length; j++) {
      if (row.join() == uniqueData[j].join()) {
        // if there is a duplicate, highlight the 'row' from the sheet
        duplicate = true;
        var duplicateRange = sheet.getRange(
          rowNum + i,
          columnNum,
          1,
          columnLength
        );
        duplicateRange.setBackground('yellow');
      }
    }

    // if there are no duplicates, add 'row' to the uniqueData array
    if (!duplicate) {
      uniqueData.push(row);
    }
  }
}

function highlight_duplicate_and_copy(action){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var range;
  var range_select = PropertiesService.getUserProperties().getProperty('range_select');
  if (range_select == 1)
  {
    range = sheet.getRange(2,1, sheet.getLastRow(), sheet.getLastColumn());
  }
  else{
    range = sheet.getActiveRange();
  }
  var data = range.getValues();
  
  var rowNum = range.getRow();
  var columnNum = range.getColumn();
  var columnLength = data[0].length;

  var uniqueData = [];
  var duplicateData = [];
  
  var firstCol = range.getColumn();
  var lastCol = sheet.getLastColumn();
  var dataLen = data.length;
  // iterate through each 'row' of the selected range
  for (var i = 0; i < dataLen; i++) {
    var row = data[i];
    var is_empty=0;
    for ( var stosic = firstCol; stosic <= lastCol; stosic++)
    {
      if (row[stosic-1]==0)
      {
        is_empty += 1;
      }
    }
    if (is_empty == lastCol - firstCol + 1)
    {
      continue;
    }
    for ( var stosic = firstCol; stosic <= lastCol; stosic++)
    {
      if (action[stosic]==0)
      {
        row[stosic-1] = "";
      }
    }
    var duplicate = false;
    var duplicateRange;
    // iterate through the uniqueData array to see if 'row' already exists
    for (var j = 0; j < uniqueData.length; j++) {
      if (row.join() == uniqueData[j].join()) {
        // if there is a duplicate, highlight the 'row' from the sheet
        duplicate = true;
        duplicateRange = sheet.getRange(
          rowNum + i,
          columnNum,
          1,
          columnLength
        );
        duplicateRange.setBackground('yellow');
      }
    }

    // if there are no duplicates, add 'row' to the uniqueData array
    if (!duplicate) {
      uniqueData.push(row);
    }
    else
    {
      var tmp = duplicateRange.getValues();
      duplicateData.push(tmp[0]);
    }
    
    
  }
  
  // create a new sheet with the duplicate data
  if (duplicateData) {
    var newSheet = spreadsheet.insertSheet();
    var header = duplicateData.length + ' duplicates found';
    newSheet.setName('Duplicates v' + (spreadsheet.getSheets().length - 1));
    newSheet.appendRow([header]);

    for (var k = 0; k < duplicateData.length; k++) {
      newSheet.appendRow(duplicateData[k]);
    }
  }
}
//For Descending:
function sortSpreadSheetD(){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    //Change Form Responses 1 to your sheetname
    var sheet = ss.getSheetByName("Form Responses 1");
    var range = sheet.getRange(2,1, sheet.getLastRow(), sheet.getLastColumn());
    range.sort({column: 1, ascending: false});
}

//For Ascending:
function sortSpreadSheetA(col){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    //Change Form Responses 1 to your sheetname
    //var sheet = ss.getSheetByName("Form Responses 1");
    var sheet = ss.getActiveSheet();
    var range;
  
    var range_select = PropertiesService.getUserProperties().getProperty('range_select');
    if (range_select == 1)
    {
      range = sheet.getRange(2,1, sheet.getLastRow(), sheet.getLastColumn());
    }
    else {
      range = sheet.getActiveRange();
    }
    //range.sort({column: 1, ascending: true});
    //Logger.log(range_select);
    //Logger.log(parseInt(col, 10));
    range.sort({column: col, ascending: true}); 
} 

function getAutoRange(){
   var ss = SpreadsheetApp.getActiveSpreadsheet();
    //Change Form Responses 1 to your sheetname
    //var sheet = ss.getSheetByName("Form Responses 1");
   var sheet = ss.getActiveSheet();
   var range = sheet.getRange(1,1, sheet.getLastRow(), sheet.getLastColumn());
   //range.sort({column: 1, ascending: true});
   var range_arr = [sheet.getLastColumn(), sheet.getLastRow()];
   return range_arr;
}

function getCurrentRange(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  //Logger.log(range);
  //Logger.log(range.getColumn());
  var range_arr = [range.getColumn(), range.getRow(), range.getLastColumn(), range.getLastRow()];
  return range_arr;
}

function getTitles(){
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var sheet = ss.getActiveSheet();
   var titles = [];
  for (var i = 1; i <= sheet.getLastColumn(); i++) {
    var title_cell_value = sheet.getRange(1, i).getValue();
    titles[i-1] = title_cell_value;
  }
   return titles;
}

function getTitlesAndFristRow(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
   var sheet = ss.getActiveSheet();
  var range = sheet.getRange(1,1, sheet.getLastRow(), sheet.getLastColumn());
   var titles = [];
   var lastCol = sheet.getLastColumn();
  for (var i = 1; i <= lastCol; i++) {
    var title_cell_value = sheet.getRange(1, i).getValue();
    titles[i-1] = title_cell_value;
  }
  for (var j = 1; j <= lastCol; j++) {
    var temp_str="";
    var first_row_cell_value = range.getCell(2, j).getValue();
    temp_str = temp_str + first_row_cell_value;
    titles[lastCol + j - 1] = temp_str;
  }
   return titles;
}