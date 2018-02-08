/*************************************************************************
SQUARE APPOINTMENTS HISTORY TO GOOGLE SHEETS REFORMATTER V1.0
**************************************************************************

Author:   Mushfiq Mahmud
Company:  Disciplined Minds Tutoring LLC, Tampa, FL
Created:  January 2018
Language: JavaScript

*************************************************************************/

//function onEdit(e){
//  reFormatter();
//}

function reFormatter(){

  /******************************************
  VARIABLES
  *******************************************/

  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var lastRow = sh.getLastRow();                // last Row variable
  var lastColumn = sh.getLastColumn();          // last Column variable
  var rowHeight = 30;                           // desired row height in pixels
  var strStartDateCol = "D";                    // Start Date/Time Column
  var strEndDateCol = "E";                     // End Date/Time Column
  var strClientNameCol = "G";                   // Client Name Column
  var strStaffNameCol = "I";                    // Staff Name Column
  var strBusNoteCol = "L";                      // Business Notes Column

  var date = Utilities.formatDate(new Date(), "GMT-5", "MM/dd/yyyy"); // current Date

  var strFullDate="";                           // Date in mm/dd/yyyy hh:mmam/pm
  var strOnlyDate="";                           // Date in mm/dd/yyyy

  var range = sh.getRange("A2:I" + lastRow);

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

  for (i=lastRow; i>1; i--) {
    strFullDate = sh.getRange(strStartDateCol + i).getValue();
    strOnlyDate = strFullDate.substring(0,strFullDate.indexOf(' '));

    if (strOnlyDate!=date) {
      sh.deleteRow(i);
    }
  }

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

  sh.insertColumnAfter(4);
  sh.getRange("E1").setValue("Duration");

  for (i=2; i<=lastRow; i++) {
    sh.getRange("E" + i).setFormula("=I" + i + "-H" + i);
  }
  sh.getRange("E2:E" + lastRow).setNumberFormat("h:mm").setHorizontalAlignment("center");

  /*************************************************
  4. Increase Row height
  *************************************************/

  for (i=1; i<=lastRow; i++) {
    sh.setRowHeight(i, rowHeight);
  }

  /*************************************************
  5. Sort by Location, then Status, then Duration, then Client
  *************************************************/

  lastColumn = sh.getLastColumn(); // update last Column
  lastRow = sh.getLastRow(); // update last Row
  sh.getRange(2, 1, lastRow, lastColumn).sort([{column: 1, ascending: true}, {column: 2, ascending: true}, {column: 5, ascending: true}, {column: 3, ascending: true}, {column: 6, ascending: true}]); // Sort by Location, then Status, then Client Name

  /*************************************************
  6. Alternating colors
  *************************************************/

  lastColumn = sh.getLastColumn();
  sh.getRange(1, 1, 1, lastColumn).setBackground("#8989eb"); // header color

  for (i=2; i<=lastRow; i+=2) {
    sh.getRange(i, 1, 1, lastColumn).setBackground("#FFF");
  }
  for (j=3; j<=lastRow; j+=2) {
    sh.getRange(j, 1, 1, lastColumn).setBackground("#e8e7fc");
  }

  /*************************************************
  7. Resize Columns
  *************************************************/

  sh.autoResizeColumn(1);
  sh.autoResizeColumn(2);
  sh.autoResizeColumn(3);
  sh.autoResizeColumn(4);
  sh.autoResizeColumn(5);
  sh.autoResizeColumn(6);
  sh.autoResizeColumn(8);
  sh.autoResizeColumn(9);
  sh.autoResizeColumn(10);

  lastRow = sh.getLastRow(); // update lastRow
  lastColumn = sh.getLastColumn(); // update lastColumn
  sh.getRange(1,7,lastRow).setWrap(true);

  /*************************************************
  8. Highlight Cancelled Appointments
  *************************************************/

  for (i=2; i<=lastRow; i++) {
    if (sh.getRange(i, 2).getValue()!="accepted") {
      sh.getRange(i, 1, 1, lastColumn).setBackground("#e06666");
    }
  }

  /*************************************************
  9. Completion Dialog Box
  *************************************************/

  // TBW

}

/************************************************
UNUSED CODE:
*************************************************

var editedCell = sh.getActiveRange().getColumnIndex();

*************************************************
DATE FORMATS:
*************************************************

Year          yyyy  1996
Year          yy    96
Month         MMMMM July
Month         MMM   Jul
Month         MM    07
Month         M     7
Day           dd    09
Day           d     9

Hour(0-23)    H     0
Hour(am/pm)   h     12

Am/pm marker  a     PM
Day name      E     Tuesday; Tue

************************************************/
