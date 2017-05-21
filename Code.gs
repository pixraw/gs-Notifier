/**
* @author Ludovic Hautier
* @copyright GPLv3
* @description Send emails to dest. on edit, about a column and a regex
*
* @todo Unhardcode messages 
* @todo Add document name in subject or email body
* @todo Unhardcode the watched sheet (get by name)
*/

var notifier = new Notifier();

function onEdit(e) {
  var regex = new RegExp(notifier.getRegex(), "g");
  //if the cell value is in specified column, and is an email
  if ( e.range.getColumn() == notifier.getNumColumn() && regex.test(e.value) ) {
    var dest = notifier.getDest();
    for (i in dest){
      MailApp.sendEmail(dest[i], "Notifier : something changed", "Hi,\n\rPlease update : "+ e.value +"\n\rThanks.");
    }
  }
}

function onInstall(e){
 onOpen(e); 
}

/**
* @description Create the notifier sheet if it doesn't exist then fill with "basic" parameters
*/
function onOpen() {
  if (notifier.getNotifierSheet() == null){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(notifier.name);
    //set titles
    var range = sheet.getRange("A1:D1");
    var titles = [["Dest.", "NumDest", "NumColumn", "RegEx"]];
    range.setValues(titles);
  
    //set formulas
    range = sheet.getRange("B2:D2");
    var formulas = [["=COUNTA(A:A)-1", "=COLUMN('Feuille 1'!D:D)", "=\".*@.*[.].*\""]];
    range.setFormulas(formulas);
  }
  notifier = new Notifier(); //create then refresh notifier
}

/**
* @description Notifier class, properties added on the fly.
*/
function Notifier(){
  //init sheet
  this.name = "Notifier";
  this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(this.name);
  if (this.sheet == null) {
    Logger.log("Notifier sheet error");
  } else {
    //complete the properties if the sheet exists
    //get numDest (B)
    var num = this.sheet.getSheetValues(2,2,1,1);
    if (num == "" || num == 0) {
      Logger.log("numDest error");
      num = 50;
    }
    var values = this.sheet.getSheetValues(2, 1, num, 1); //get Dest. (A)
  
    //set dest
    //if an email is present, add to finalValues (strip the [] elems)
    //needed if num = 50
    this.dest = [];
    for (i in values){
      if (values[i] != "")
        this.dest.push(values[i]);
    } 
    //set numColumn
    this.numColumn = this.sheet.getSheetValues(2,3,1,1); //get numColumn (C)
    if (this.numColumn == null) {
      Logger.log("numColumn error");
    }
    this.regex = this.sheet.getSheetValues(2,4,1,1); //get regex (D)
    if (this.numColumn == null) {
      Logger.log("regex error");
    }
  }
  //Getters
  this.getNotifierSheet = function(){
    return this.sheet;
  } 
  this.getDest = function(){
    return this.dest;
  }
  this.getNumColumn = function(){
    return this.numColumn;
  }
  this.getRegex = function(){
    return this.regex;
  }
}
  
