var properties = PropertiesService.getScriptProperties();

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Raffle')
      .addItem('Create tickets', 'createUniqueTickets')
      .addItem('Pick winner', 'pickWinner')
      .addItem('Reset winners', 'resetWinners')
      .addToUi();
}

// change these to relevant columns
var colName = 0;
var colSurname = 1;
var colEmail = 3;
var colTickets = 5;

var sname = 'Unique tickets';

function pickWinner() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getSheetByName(sname);
  var last = sheet.getLastRow();
  var winner = Math.ceil(Math.random()*last);
  var winCount = properties.getProperty('winCount') || 1;

  var existing = ss.getRange('A'+winner).getBackground()
  if (existing == '#ff0000') {
    pickWinner();
    // TODO: takes a while if lots of winnersand will crash if all are selected.
    return;
  }
  ss.getRange('A'+winner+':B'+winner).setBackground('#ff0000');

  var cell = sheet.getRange('A'+winner);
  ss.setCurrentCell(cell);

  ss.getRange('C'+winner).setValue(winCount);

  properties.setProperty('winCount', parseInt(winCount)+1);
}

function resetWinners(){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getSheetByName(sname);
  var last = sheet.getLastRow();

  ss.getRange('A1:B'+last).setBackground('#ffffff');
  ss.getRange('C1:C'+last).clear();

  properties.deleteProperty('winCount');
}

function createUniqueTickets() {

  var dataSheet = SpreadsheetApp.getActive().getSheets()[0];
  var data = dataSheet.getDataRange().getValues();
  var uniques = [];
  data.forEach(function (row) {
    
    var name = row[colName] + ' ' + row[colSurname];
    var email = row[colEmail];
    var tickets = row[colTickets];

    for (var i=0; i<tickets; i++) {
      uniques.push([name, email]);
    }

    Logger.log('Added: ' + name +'('+ email + ') tickets:'+tickets);
  });


  var ss = SpreadsheetApp.getActiveSpreadsheet(); 

  var usheet = ss.getSheetByName(sname);
  if (usheet) {
    ss.deleteSheet(usheet);
  }

  var ns = ss.insertSheet(sname);
  
  var range = ns.getRange('A1:B'+uniques.length);
  range.setValues(uniques);

}
