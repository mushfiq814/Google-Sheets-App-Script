/*************************************************************************
SQUARE APPOINTMENTS HISTORY TO GOOGLE SHEETS REFORMATTER V1.0
**************************************************************************

Author:   Mushfiq Mahmud
Company:  Disciplined Minds Tutoring LLC, Tampa, FL
Created:  January 2018
Language: JavaScript

*************************************************************************/

function reFormatter(){

  /******************************************
  VARIABLES
  *******************************************/

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Live Sheet");
  var shRecur = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Recurring Appointments");
  var shFormat = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Formatting Sheet");
  var shCancel = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Cancelled Appointments");

  var lastRow = sh.getLastRow();                // last Row variable
  var lastColumn = sh.getLastColumn();          // last Column variable
  var maxRows = sh.getMaxRows();                // maximum rows
  var rowHeight = 30;                           // desired row height in pixels
  var strStartDateCol = "D";                    // Start Date/Time Column
  var strEndDateCol = "E";                      // End Date/Time Column
  var strClientNameCol = "G";                   // Client Name Column
  var strStaffNameCol = "I";                    // Staff Name Column
  var strBusNoteCol = "L";                      // Business Notes Column

  var date = Utilities.formatDate(new Date(), "GMT-5", "MM/dd/yyyy"); // current Date

  var strFullDate="";                           // Date in mm/dd/yyyy hh:mmam/pm
  var strOnlyDate="";                           // Date in mm/dd/yyyy

  var range = sh.getRange("A2:I" + lastRow);

  var recurLastRow = shRecur.getLastRow();
  var recurLastColumn = shRecur.getLastColumn();
  sh.getRange(lastRow+1, 1, recurLastRow, recurLastColumn).setValues(shRecur.getRange(2,1,recurLastRow, recurLastColumn).getValues());

  // Sort by Date
  lastRow = sh.getLastRow(); // update lastRow
  lastColumn = sh.getLastColumn(); // update lastColumn
  sh.getRange(2, 1, lastRow, lastColumn).sort([{column: 7, ascending: true}]);

  /******************************************
  Original Order:
  A (1): location_type
  B (2): location_name
  C (3): status
  D (4): created_at
  E (5): address
  F (6): service
  G (7): start
  H (8): end
  I (9): client_name
  J (10): client_email
  K (11): client_phone
  L (12): staff
  M (13): note_from_client
  N (14): note_from_business
  *******************************************/

  // Delete columns
  sh.deleteColumns(11);   // client_phone
  sh.deleteColumns(10);   // client_email
  sh.deleteColumns(5);    // address
  sh.deleteColumns(4);    // created_at
  sh.deleteColumns(1);    // location_type

  sh.autoResizeColumn(1);
  sh.autoResizeColumn(2);
  sh.autoResizeColumn(3);
  sh.autoResizeColumn(4);
  sh.autoResizeColumn(5);
  sh.autoResizeColumn(6);
  sh.autoResizeColumn(7);
  sh.autoResizeColumn(8);
  sh.autoResizeColumn(9);

  sh.setFrozenRows(1);
  sh.setRowHeight(1, 60);

  lastRow = sh.getLastRow(); // update lastRow
  lastColumn = sh.getLastColumn(); // update lastColumn
  sh.getRange(1,1, lastRow, lastColumn).clearDataValidations().clear({formatOnly:true});

  /******************************************
  Current Order after deletion:
  A (1): location
  B (2): status
  C (3): service
  D (4): start
  E (5): end
  F (6): client_name
  G (7): staff
  H (8): note_from_client
  I (9): note_from_business

  Desired Order:
  A (1): location
  B (2): status
  C (3): client_name
  D (4): service
  E (5): duration
  F (6): staff
  G (7): note_from_business
  H (8): start
  I (9): end
  J (10): note_from_client
  *******************************************/

  /*************************************************
  1. Delete entries that are not today
  *************************************************/

  sh.getRange(strStartDateCol + "2:" + strStartDateCol + lastRow).setNumberFormat("mm/dd/yyyy hh:mmam/pm");
  sh.getRange(strEndDateCol + "2:" + strEndDateCol + lastRow).setNumberFormat("mm/dd/yyyy hh:mmam/pm");
  sh.getRange(strStartDateCol + "2:" + strStartDateCol + lastRow).setNumberFormat("@");

            /***************************************
             !!! MORE EFFICIENT DELETING METHOD !!!
            ***************************************/

  lastRow = sh.getLastRow(); // update lastRow
  lastColumn = sh.getLastColumn(); // update lastColumn

  var dataArray = sh.getRange(2,1,lastRow,lastColumn).getValues();

  var todayAppCount = 0, todayAppStartRow = 0;
  var cancelledAppCount = 0, cancelledAppStartRow = 0;

  for(var i=0, dLen=dataArray.length; i<dLen; i++) {
    strOnlyDate = dataArray[i][3].substring(0,dataArray[i][3].indexOf(' '));
    if(strOnlyDate == date) {
      todayAppCount++;
    }
  }

  for(var j=0; j<dataArray.length; j++) {
    strOnlyDate = dataArray[j][3].substring(0,dataArray[j][3].indexOf(' '));
    if(strOnlyDate != date) {
      todayAppStartRow++;
    } else {
      break;
    }
  }

  var onlyToday = dataArray.splice(todayAppStartRow, todayAppCount);

  /**************************************************************************************************/

//  for(var i=0, dLen=onlyToday.length; i<dLen; i++) {
//    if(onlyToday[i][2] != "accepted") {
//      cancelledAppCount++;
//    }
//  }
//
//  for(var j=0; j<onlyToday.length; j++) {
//    if(onlyToday == "accepted") {
//      cancelledAppStartRow++;
//    } else {
//      break;
//    }
//  }

  /**************************************************************************************************/

  sh.getRange(2,1,sh.getMaxRows()-1, sh.getMaxColumns()).clear();
  shCancel.getRange(2, 1, shCancel.getMaxRows()-1, shCancel.getMaxColumns()).clearDataValidations().clear();
//  var onlyCancelled = onlyToday.splice(cancelledAppStartRow, cancelledAppCount);

  sh.getRange(2, 1, onlyToday.length, onlyToday[0].length).setValues(onlyToday);
//  shCancel.getRange(2, 1, onlyCancelled.length, onlyCancelled[0].length).setValues(onlyCancelled);

  lastRow = sh.getLastRow(); // update lastRow
  lastColumn = sh.getLastColumn(); // update lastColumn

  /*************************************************
  2. Arrange columns according to Desired Order
  *************************************************/

  // move Client Name to 3rd column
  sh.insertColumnAfter(2);
  sh.getRange(strClientNameCol + "1:" + strClientNameCol + lastRow).moveTo(sh.getRange("C1"));

  // move Staff Name to 5th column
  sh.insertColumnAfter(4);
  sh.getRange(strStaffNameCol + "1:" + strStaffNameCol + lastRow).moveTo(sh.getRange("E1"));

  // move Note from Business  to 6th column
  sh.insertColumnAfter(5);
  sh.getRange(strBusNoteCol + "1:" + strBusNoteCol + lastRow).moveTo(sh.getRange("F1"));

  // delete empty columns
  sh.deleteColumns(12);
  sh.deleteColumns(10);
  sh.deleteColumns(9);

  /*************************************************
  3. Add Duration Column
  *************************************************/

  lastRow = sh.getLastRow(); // update lastRow
  lastColumn = sh.getLastColumn(); // update lastColumn
  sh.insertColumnAfter(4);
  sh.getRange("E1").setValue("duration");

  for (i=2; i<=lastRow; i++) {
    sh.getRange("E" + i).setFormula("=I" + i + "-H" + i);
  }
  sh.getRange("E2:E" + lastRow).setNumberFormat("h:mm").setHorizontalAlignment("center");

  /*************************************************
  4. Increase Row height
  *************************************************/

  for (i=2; i<=lastRow; i++) {
    sh.setRowHeight(i, rowHeight);
  }

  /*************************************************
  5. Sort by Location, then Status, then Duration, then Client
  *************************************************/

  lastRow = sh.getLastRow(); // update lastRow
  lastColumn = sh.getLastColumn(); // update lastColumn
  sh.getRange(2, 8, lastRow-1, 1).setNumberFormat("mm/dd/yyyy hh:mmam/pm");

  // Sort by Location, then Status, then Client Name
  sh.getRange(2, 1, lastRow, lastColumn).sort([{column: 1, ascending: true}, {column: 2, ascending: true}, {column: 5, ascending: true}, {column: 8, ascending: true}, {column: 3, ascending: true}]);


  /*************************************************
  6. Add columns for Tickets created?, Card on File
  Present? and Envoy Sign In? Processes?
  *************************************************/

  var numColAdded = 4;
  lastRow = sh.getLastRow(); // update lastRow
  lastColumn = sh.getLastColumn(); // update lastColumn

  sh.insertColumnsAfter(2,4);
  sh.setColumnWidth(3, 40);
  sh.setColumnWidth(4, 40);
  sh.setColumnWidth(5, 40);
  sh.setColumnWidth(6, 40);
  sh.getRange("C1").setValue("Tickets?");
  sh.getRange("D1").setValue("Card On File?");
  sh.getRange("E1").setValue("Envoy?");
  sh.getRange("F1").setValue("Processed?");
  var dataValidRule = SpreadsheetApp.newDataValidation().requireValueInList(['Y', 'N'], true).build();
  sh.getRange(2, 3, lastRow, numColAdded).setDataValidation(dataValidRule);

  shFormat.getRange("B3").copyTo(sh.getRange(2,1, lastRow-1, lastColumn), {formatOnly:true});
  shFormat.getRange("B4").copyTo(sh.getRange(2,1, lastRow-1, lastColumn), {formatOnly:true});
  shFormat.getRange("C1").copyTo(sh.getRange(2,3, lastRow-1, 1), {formatOnly:true});
  shFormat.getRange("C2").copyTo(sh.getRange(2,3, lastRow-1, 1), {formatOnly:true});
  shFormat.getRange("C1").copyTo(sh.getRange(2,4, lastRow-1, 1), {formatOnly:true});
  shFormat.getRange("C2").copyTo(sh.getRange(2,4, lastRow-1, 1), {formatOnly:true});
  shFormat.getRange("C1").copyTo(sh.getRange(2,5, lastRow-1, 1), {formatOnly:true});
  shFormat.getRange("C2").copyTo(sh.getRange(2,5, lastRow-1, 1), {formatOnly:true});
  shFormat.getRange("C1").copyTo(sh.getRange(2,6, lastRow-1, 1), {formatOnly:true});
  shFormat.getRange("C2").copyTo(sh.getRange(2,6, lastRow-1, 1), {formatOnly:true});

  /*************************************************
  7. Alternating colors
  *************************************************/

  lastColumn = sh.getLastColumn();
  sh.getRange(1, 1, 1, lastColumn).setBackground("#2c7fb4").setFontColor("#FFF").setFontWeight("bold"); // header color

  for (i=2; i<=lastRow; i+=2) {
    sh.getRange(i, 1, 1, lastColumn).setBackground("#FFF");
  }
  for (j=3; j<=lastRow; j+=2) {
    sh.getRange(j, 1, 1, lastColumn).setBackground("#b4e3f6");
  }

  /*************************************************
  8. Resize Columns
  *************************************************/

  sh.autoResizeColumn(1);
  sh.autoResizeColumn(2);
  sh.autoResizeColumn(7);
  sh.autoResizeColumn(8);
  sh.autoResizeColumn(9);
  sh.autoResizeColumn(10);
  sh.autoResizeColumn(12);
  sh.autoResizeColumn(13);

  lastRow = sh.getLastRow(); // update lastRow
  lastColumn = sh.getLastColumn(); // update lastColumn
  sh.getRange(1,3,lastRow).setWrap(true);
  sh.getRange(1,4,lastRow).setWrap(true);
  sh.getRange(1,5,lastRow).setWrap(true);
  sh.getRange(1,6,lastRow).setWrap(true);
  sh.getRange(1,11,lastRow).setWrap(true);

  /*************************************************
  9. Highlight Cancelled Appointments
  *************************************************/

  lastRow = sh.getLastRow(); // update lastRow
  lastColumn = sh.getLastColumn(); // update lastColumn

//  sh.getRange(1, 1, 1, lastColumn).copyTo(shCancel.getRange(1, 1, 1, lastColumn));
//  for (i=2; i<=lastRow; i++) {
//    if (sh.getRange(i, 2).getValue()!="accepted") {
//      sh.getRange(i, 1, 1, lastColumn).setBackground("#e06666").copyTo(shCancel.getRange(i,1, 1, lastColumn));
//    }
//  }

  /*************************************************
  10. Delete Empty Rows at the end
  *************************************************/

  lastRow = sh.getLastRow(); // update lastRow
  lastColumn = sh.getLastColumn(); // update lastColumn
  maxRows = sh.getMaxRows(); // update maxRow

  sh.deleteRows(lastRow + 1, maxRows - lastRow);

}


/************************************************************
UNUSED CODE
************************************************************/

//  // iterate through array (backwards) to calculate the rowIndex
//  for(var j=dLen=dataArray.length-1; j>=0; j--) {
//    if(dataArray[j][2] != "Done") {
//      cSecond++;
//    } else {
//      break;
//    }
//  }


//  var editedCell = sh.getActiveRange().getColumnIndex();
//
//  var now = new Date();
//  var twoDaysFromNow = new Date(now.getTime() + (48 * 60 * 60 * 1000));
//
//  var cal = CalendarApp.getCalendarById("mushfiq8194@gmail.com");
//  var calName = cal.getTitle();
//  var events = cal.getEvents(now, twoDaysFromNow);
//
//  var str1, str2, str3, str4, str5;
//
//  for (j=1; j<=events.length; j++) {
//    str1="A";
//    str1+=j;
//    sh.getRange(str1).setValue(events[0].getDescription());
//    str2="B";
//    str2+=j;
//    sh.getRange(str2).setValue(events[0].getLocation());
//  }
//    .getTitle()
//    .getStartTime()
//    .getEndTime()
//    .getDescription()
//    .getLocation()
