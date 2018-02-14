function temp() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Live Sheet");
  var shRecur = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Recurring Appointments");
  var shFormat = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Formatting Sheet");
  var shCancel = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cancelled Appointments");
  var shNew = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Backup for Recurring");
  var lastRow = sh.getLastRow();
  var lastColumn = sh.getLastColumn();

  Logger.log("lastRow: " + lastRow);

}

//  for (var i=6; i<=50; i++) {
//    var lastRow = shRecur.getLastRow();
//    var array = shNew.getRange(i,1,1,15).getValues();
//  }
//
//    shRecur.getRange(lastRow+1, 1, 1, 15).setValues(array);
//
//    for (var j=0; j<15; j++) {
//      var lastRow = shRecur.getLastRow();
//      shRecur.getRange(lastRow+1, 1, 1, 15).setValues(array);
//      shRecur.getRange(lastRow+1, 8).setFormula("=H"+lastRow+"+7");
//      shRecur.getRange(lastRow+1, 9).setFormula("=I"+lastRow+"+7");
//      if (even==0) {shRecur.getRange(lastRow+1, 1, 1, 15).setBackground("#fce5cd"); even=1;}
//      else if (even==1) {shRecur.getRange(lastRow+1, 1, 1, 15).setBackground("#fff2cc"); even=0;}
//    }
//  }
//
//  var lastRow = shRecur.getLastRow();
//  var array = shNew.getRange(5,1,1,15).getValues();
//  shRecur.getRange(lastRow+1, 1, 1, 15).setValues(array);
//  for (var j=0; j<15; j++) {
//    var lastRow = shRecur.getLastRow();
//    shRecur.getRange(lastRow+1, 1, 1, 15).setValues(array);
//    shRecur.getRange(lastRow+1, 8).setFormula("=H"+lastRow+"+7");
//    shRecur.getRange(lastRow+1, 9).setFormula("=I"+lastRow+"+7");
//    if (count==0) {shRecur.getRange(lastRow+1, 1, 1, 15).setBackground("#fce5cd"); count=1;}
//    else if (count==1) {shRecur.getRange(lastRow+1, 1, 1, 15).setBackground("#fff2cc"); count=0;}
//  }
