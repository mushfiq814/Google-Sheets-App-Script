function createTicketNames() {
  var endRow = 79;

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Monday Feb 26");
  var array = sh.getRange(3,1,endRow-2,12).getValues();
  for (var i=0; i< endRow-2; i++) {
    sh.getRange(i+3, 17).setValue(array[i][0] + " - " + array[i][8] + " - " + array[i][9] + " - " + array[i][11]);
  }
}

function splitIntoColumns() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Live Sheet");
  var shRecur = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Recurring Appointments");
  var shBackup = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Backup for Recurring");
  var shEdit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Editing Sheet");
  var array= [];
  var lastRow = shEdit.getLastRow();
  var dataArray = shEdit.getRange(1,1,lastRow,1).getValues();
  if (lastRow%3==0) {
    for (var x=0; x<lastRow/3; x++ ) {
      array[x] = [];
      for (var y=0; y<3; y++) {
        array[x][y] = dataArray[3*x+y];
      }
    }
    shEdit.getRange(2, 9, array.length, array[0].length).setValues(array);

    shEdit.getRange(2, 11, array.length, 1).copyTo(shEdit.getRange(2, 4, array.length, 1));
    shEdit.getRange(2, 10, array.length, 1).copyTo(shEdit.getRange(2, 7, array.length, 1));

    for (var i=0; i<array.length; i++) {
      shEdit.getRange(i+2, 5, 1, 1).setFormula("=SPLIT($I" + (i+2) + ", \"–\")");
    }
    //addDate();
  }
}

function addDate() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Live Sheet");
  var shRecur = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Recurring Appointments");
  var shBackup = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Backup for Recurring");
  var shEdit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Editing Sheet");
  var lastRow = shEdit.getLastRow();
  shEdit.deleteColumns(1,3);
  shEdit.getRange(2, 9, lastRow-1, 2).setValues(shEdit.getRange(2, 2, lastRow-1, 2).getValues());
  shEdit.deleteColumns(6,3);
  lastRow = shEdit.getLastRow();
  var numDaysSince1899 = 43148;
  //shEdit.getRange(2, 2, lastRow-1, 2).copyTo(shEdit.getRange(2, 6, lastRow-1, 2));
  var array = shEdit.getRange(2, 6, lastRow-1, 2).getValues();
  for (var i=0; i<array.length; i++) {
    for (var j=0; j<2; j++) {
      array[i][j]=numDaysSince1899+array[i][j];
    }
  }
  shEdit.getRange(2, 2, array.length, array[0].length).setValues(array);
  shEdit.getRange(2, 2, lastRow-1, 2).setNumberFormat("mm/dd/yyyy hh:mmam/pm");
}

function createRepeatedApps() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Live Sheet");
  var shRecur = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Recurring Appointments");
  var shBackup = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Backup for Recurring");
  var shEdit = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Editing Sheet");
  var startCopy = 287, endCopy = 296;
  var array = shBackup.getRange(startCopy,1,endCopy-startCopy,15).getValues();
  var lastRow = shRecur.getLastRow();
  for (var i=0; i<=array.length; i++) {
    lastRow = shRecur.getLastRow();
    shRecur.getRange(lastRow+1, 1, 1, 15).setValues([array[i]]);
    for (var j=0; j<15; j++) {
      lastRow = shRecur.getLastRow();
      shRecur.getRange(lastRow+1, 1, 1, 15).setValues([array[i]]);
      shRecur.getRange(lastRow+1, 8).setFormula("=H"+lastRow+"+7");
      shRecur.getRange(lastRow+1, 9).setFormula("=I"+lastRow+"+7");
    }
  }
}
