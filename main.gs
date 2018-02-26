/*************************************************************************
SQUARE APPOINTMENTS HISTORY TO GOOGLE SHEETS REFORMATTER V2.0
**************************************************************************

Author:   Mushfiq Mahmud
Company:  Disciplined Minds Tutoring LLC, Tampa, FL
Created:  January 2018; Updated: February 2018

*************************************************************************/

function createArrays(nonCancel) {

  var boolNonCancel = nonCancel;

  /*************************************************************************************
                                        VARIABLES
  *************************************************************************************/
  var sh = SpreadsheetApp.getActiveSpreadsheet();
  var shMain = sh.getSheetByName("Live Sheet");
  var shRecur = sh.getSheetByName("Recurring Appointments");
  var shFormat = sh.getSheetByName("Formatting Sheet (DO NOT EDIT)");
  var shCancel = sh.getSheetByName("Cancelled Appointments");

  var lastRow = shMain.getLastRow();
  var lastColumn = shMain.getLastColumn();
  var recurLastRow = shRecur.getLastRow();
  var recurLastColumn = shRecur.getLastColumn();

  var date = Utilities.formatDate(new Date(), "GMT-5", "MM/dd/yyyy"); // current Date
  date = "02/21/2018";

  var columns = {
    locationType:1,
    locationName:2,
    status:3,
    createdAt:4,
    address:5,
    service:6,
    start:7,
    end:8,
    clientName:9,
    clientEmail:10,
    clientPhone:11,
    staff:12,
    noteFromClient:13,
    noteFromBusiness:14,
  };

  // var column = { ticketNumber:"", locationName:"", status:"", ticketsCreated:"", cardOnFilePresent:"", envoySignIn:"", specialBilling:"", processed:"", clientName:"", service:"", duration:"", staff:"", noteFromBusiness:"", start:"", end:"", noteFromClient:"" locationType:"", createdAt:"", address:"", clientEmail:"", clientPhone:"" };

  /*************************************************************************************
  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!DELETE ME!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
  *************************************************************************************/
  var shBackupData = sh.getSheetByName("Backup Data");
  var array = shBackupData.getRange(1,1,shBackupData.getLastRow(),shBackupData.getLastColumn()).getValues();
  shMain.getRange(1,1,shMain.getMaxRows(), shMain.getMaxColumns()).clear();
  shMain.getRange(1,1,shBackupData.getLastRow(),shBackupData.getLastColumn()).setValues(array);

  /*************************************************************************************
                                          STARTUP
  *************************************************************************************/

  // Clear Data Validations and formatting
  shMain.getRange(1,1,shMain.getMaxRows(),shMain.getMaxColumns()).clearDataValidations().clearFormat();

  // Add first row
  shMain.insertRowBefore(1);

  // Get values from recurring sheet
  lastRow = shMain.getLastRow(); // update lastRow
  lastColumn = shMain.getLastColumn(); // update lastColumn
  recurLastRow = shRecur.getLastRow(); // update recurLastRow
  recurLastColumn = shRecur.getLastColumn(); // update recurLastColumn
  shMain.getRange(lastRow+1, 1, recurLastRow-1, recurLastColumn).setValues(shRecur.getRange(2,1,recurLastRow-1, recurLastColumn).getValues());

  // Sort by date
  lastRow = shMain.getLastRow(); // update lastRow
  lastColumn = shMain.getLastColumn(); // update lastColumn
  shMain.getRange(3,1,lastRow-2,lastColumn).sort([{column: columns["start"], ascending: true}]);

  // Auto resize all columns
  for (var i=1; i<=14; i++) {
    shMain.autoResizeColumn(i);
  }

  // Set number format for columns (G,H) OR (7,8) to "mm/dd/yyyy hh:mmam/pm"
  lastRow = shMain.getLastRow(); // update lastRow
  lastColumn = shMain.getLastColumn(); // update lastColumn
  shMain.getRange(3, columns["start"], lastRow-2, 2).setNumberFormat("mm/dd/yyyy hh:mmam/pm");

  // Set number format for columns (G) OR (7) to "@"
  shMain.getRange(3, columns["start"], lastRow-2, 1).setNumberFormat("@");

  /*************************************************************************************
                                    MEMORY
  *************************************************************************************/

  // Create allEntriesByDate[] with all values in sheet (except Header row)
  var allEntriesByDate = shMain.getRange(3, 1, lastRow-2, 14).getValues();

  // Calculate how many are today and which row has today's first appointment
  var todayAppCount = 0, todayAppStartRow = 0;

  for (var i=0, dLen=allEntriesByDate.length; i<dLen; i++) {
    var strOnlyDate = allEntriesByDate[i][Math.floor(columns["start"])-1].substring(0,allEntriesByDate[i][Math.floor(columns["start"])-1].indexOf(' '));
    if (strOnlyDate == date) {
      todayAppCount++;
    }
  }

  for (var j=0; j<allEntriesByDate.length; j++) {
    var strOnlyDate = allEntriesByDate[j][Math.floor(columns["start"])-1].substring(0,allEntriesByDate[j][Math.floor(columns["start"])-1].indexOf(' '));
    if (strOnlyDate != date) {
      todayAppStartRow++;
    } else {
      break;
    }
  }

  // Create onlyTodayByStartTime[] with only today's appointments
  var onlyTodayByStartTime = allEntriesByDate.splice(todayAppStartRow, todayAppCount);

  // Clear whole sheet (except Header row)
  shMain.getRange(3,1,shMain.getMaxRows()-2, shMain.getMaxColumns()).clear();

  // Paste onlyTodayByStartTime[] into sheet
  shMain.getRange(3, 1, onlyTodayByStartTime.length, onlyTodayByStartTime[0].length).setValues(onlyTodayByStartTime);

	// Sort by status
  lastRow = shMain.getLastRow(); // update lastRow
  lastColumn = shMain.getLastColumn(); // update lastColumn
  shMain.getRange(3,1,lastRow-2,lastColumn).sort([{column: Math.floor(columns["status"]), ascending: true}]);

  // Create onlyTodayByStatus[] with all remaining values
  var onlyTodayByStatus = shMain.getRange(3, 1, lastRow-2, 14).getValues();

  // Clear whole sheet (except Header row)
  shMain.getRange(3,1,shMain.getMaxRows()-2, shMain.getMaxColumns()).clear();

  // Calculate how many are cancelled and which row has the first cancelled appointment
  var cancelledAppCount = 0, cancelledAppStartRow = 0;

   for (var i=0, dLen=onlyTodayByStatus.length; i<dLen; i++) {
     if (onlyTodayByStatus[i][Math.floor(columns["status"])-1] != "accepted") {
       cancelledAppCount++;
     }
   }

   for (var j=0; j<onlyTodayByStatus.length; j++) {
     if (onlyTodayByStatus[j][Math.floor(columns["status"])-1] == "accepted") {
       cancelledAppStartRow++;
     } else {
       break;
     }
   }

  // Create onlyCancelled[] with only today's cancelled appointments
  // Create onlyNonCancelled[] with only today's non-cancelled appointments
  var onlyCancelled = onlyTodayByStatus.splice(cancelledAppStartRow, cancelledAppCount);
  var onlyNonCancelled = onlyTodayByStatus;

  // If flag is true, return onlyNonCancelled[], otherwise, return onlyCancelled[]
  if (boolNonCancel) {
    return onlyNonCancelled;
  }
  else {
    return onlyCancelled;
  }

}

function getNonCancelledForToday() {

  var sh = SpreadsheetApp.getActiveSpreadsheet();
  var shNonCancel = sh.getSheetByName("Live Sheet");
  var shEdit = sh.getSheetByName("Editing Sheet");
  var shFormat = sh.getSheetByName("Formatting Sheet (DO NOT EDIT)");

  // Run createArrays(true) and Save returned result into onlyNonCancelled[]
  var onlyNonCancelled = createArrays(true);

  // Clear Live Sheet (except Header row)
  shNonCancel.getRange(3,1,shNonCancel.getMaxRows()-2, shNonCancel.getMaxColumns()).clear().clearDataValidations();

  // Paste values into sheet
  shEdit.getRange(2, 1, onlyNonCancelled.length, onlyNonCancelled[0].length).setValues(onlyNonCancelled);
  var editLastRow = shEdit.getLastRow();
  shEdit.getRange(2, 2, editLastRow-1, 1).copyTo(shNonCancel.getRange(3, 2, editLastRow-1, 1));   // "Location"           onlyNonCancelled[][1]  to (B) OR (2)
  shEdit.getRange(2, 3, editLastRow-1, 1).copyTo(shNonCancel.getRange(3, 3, editLastRow-1, 1));   // "Status"             onlyNonCancelled[][2]  to (C) OR (3)
  shEdit.getRange(2, 6, editLastRow-1, 1).copyTo(shNonCancel.getRange(3, 10, editLastRow-1, 1));  // "Service"            onlyNonCancelled[][5]  to (J) OR (10)
  shEdit.getRange(2, 7, editLastRow-1, 1).copyTo(shNonCancel.getRange(3, 14, editLastRow-1, 1));  // "Start"              onlyNonCancelled[][6]  to (N) OR (14)
  shEdit.getRange(2, 8, editLastRow-1, 1).copyTo(shNonCancel.getRange(3, 15, editLastRow-1, 1));  // "End"                onlyNonCancelled[][7]  to (O) OR (15)
  shEdit.getRange(2, 9, editLastRow-1, 1).copyTo(shNonCancel.getRange(3, 9, editLastRow-1, 1));   // "Student"            onlyNonCancelled[][8]  to (I) OR (9)
  shEdit.getRange(2, 12, editLastRow-1, 1).copyTo(shNonCancel.getRange(3, 12, editLastRow-1, 1)); // "Tutor"              onlyNonCancelled[][11] to (L) OR (12)
  shEdit.getRange(2, 13, editLastRow-1, 1).copyTo(shNonCancel.getRange(3, 16, editLastRow-1, 1)); // "Note from Client"   onlyNonCancelled[][12] to (P) OR (16)
  shEdit.getRange(2, 14, editLastRow-1, 1).copyTo(shNonCancel.getRange(3, 13, editLastRow-1, 1)); // "Note from Business" onlyNonCancelled[][13] to (M) OR (13)

  // Put Column Names as follows:
  shNonCancel.getRange(2, 1).setValue("Ticket #");
  shNonCancel.getRange(2, 2).setValue("Location");
  shNonCancel.getRange(2, 3).setValue("Status");
  shNonCancel.getRange(2, 4).setValue("Tickets Created?");
  shNonCancel.getRange(2, 5).setValue("Card on File Present?");
  shNonCancel.getRange(2, 6).setValue("Envoy Sign-In?");
  shNonCancel.getRange(2, 7).setValue("Special Billing?");
  shNonCancel.getRange(2, 8).setValue("Processed?");
  shNonCancel.getRange(2, 9).setValue("Student");
  shNonCancel.getRange(2, 10).setValue("Service");
  shNonCancel.getRange(2, 11).setValue("duration");
  shNonCancel.getRange(2, 12).setValue("Tutor");
  shNonCancel.getRange(2, 13).setValue("Note from Business");
  shNonCancel.getRange(2, 14).setValue("Start");
  shNonCancel.getRange(2, 15).setValue("End");
  shNonCancel.getRange(2, 16).setValue("Note from Client");
  shNonCancel.getRange(2, 17).setValue("Ticket Name");

  // Do Data Validation
  var lastRow = shNonCancel.getLastRow(); // get lastRow
  var lastColumn = shNonCancel.getLastColumn(); // get lastColumn
  var ticketsDataValidRule = SpreadsheetApp.newDataValidation().requireValueInList(['Y', 'N'], true).build();
  var cardOnFileDataValidRule = SpreadsheetApp.newDataValidation().requireValueInList(['Y', 'N pays with cash/check', 'N'], true).build();
  var envoyDataValidRule = SpreadsheetApp.newDataValidation().requireValueInList(['Y', 'N but student was here', 'N'], true).build();
  var specialBillingDataValidRule = SpreadsheetApp.newDataValidation().requireValueInList(['Y Prepaid', 'Y Monthly', 'Y Biweekly', 'Y Weekly', 'N'], true).build();
  var processedDataValidRule = SpreadsheetApp.newDataValidation().requireValueInList(['Y CC on file', 'Y check/cash', 'Y on time cancel', 'Y late cancel/no-show', 'N no payment info', 'N special billing'], true).build();
  shNonCancel.getRange(3, 4, lastRow, 1).setDataValidation(ticketsDataValidRule);
  shNonCancel.getRange(3, 5, lastRow, 1).setDataValidation(cardOnFileDataValidRule);
  shNonCancel.getRange(3, 6, lastRow, 1).setDataValidation(envoyDataValidRule);
  shNonCancel.getRange(3, 7, lastRow, 1).setDataValidation(specialBillingDataValidRule);
  shNonCancel.getRange(3, 8, lastRow, 1).setDataValidation(processedDataValidRule);

  // Do Conditional Formatting
  lastRow = shNonCancel.getLastRow(); // update lastRow
  lastColumn = shNonCancel.getLastColumn(); // update lastColumn
  shFormat.getRange("C2").copyTo(shNonCancel.getRange(3,4,lastRow-2,1) , {formatOnly: true}); // Tickets 'Y'
  shFormat.getRange("C3").copyTo(shNonCancel.getRange(3,4,lastRow-2,1) , {formatOnly: true}); // Tickets 'N'
  shFormat.getRange("D2").copyTo(shNonCancel.getRange(3,5,lastRow-2,1) , {formatOnly: true}); // Card on File 'Y'
  shFormat.getRange("D3").copyTo(shNonCancel.getRange(3,5,lastRow-2,1) , {formatOnly: true}); // Card on File 'N pays with cash/check'
  shFormat.getRange("D4").copyTo(shNonCancel.getRange(3,5,lastRow-2,1) , {formatOnly: true}); // Card on File 'N'
  shFormat.getRange("E2").copyTo(shNonCancel.getRange(3,6,lastRow-2,1) , {formatOnly: true}); // Envoy 'Y'
  shFormat.getRange("E3").copyTo(shNonCancel.getRange(3,6,lastRow-2,1) , {formatOnly: true}); // Envoy 'N but student was here'
  shFormat.getRange("E4").copyTo(shNonCancel.getRange(3,6,lastRow-2,1) , {formatOnly: true}); // Envoy 'N'
  shFormat.getRange("F2").copyTo(shNonCancel.getRange(3,7,lastRow-2,1) , {formatOnly: true}); // Special Billing 'Y Prepaid'
  shFormat.getRange("F3").copyTo(shNonCancel.getRange(3,7,lastRow-2,1) , {formatOnly: true}); // Special Billing 'Y Monthly'
  shFormat.getRange("F4").copyTo(shNonCancel.getRange(3,7,lastRow-2,1) , {formatOnly: true}); // Special Billing 'Y Biweekly'
  shFormat.getRange("F5").copyTo(shNonCancel.getRange(3,7,lastRow-2,1) , {formatOnly: true}); // Special Billing 'Y Weekly'
  shFormat.getRange("F6").copyTo(shNonCancel.getRange(3,7,lastRow-2,1) , {formatOnly: true}); // Special Billing 'N'
  // Full row Formats
  shFormat.getRange("F9").copyTo(shNonCancel.getRange(3,1,lastRow-2,lastColumn) , {formatOnly: true}); // Processed 'Y CC on file'
  shFormat.getRange("F10").copyTo(shNonCancel.getRange(3,1,lastRow-2,lastColumn) , {formatOnly: true}); // Processed 'Y check/cash'
  shFormat.getRange("F11").copyTo(shNonCancel.getRange(3,1,lastRow-2,lastColumn) , {formatOnly: true}); // Processed 'Y on time cancel'
  shFormat.getRange("F12").copyTo(shNonCancel.getRange(3,1,lastRow-2,lastColumn) , {formatOnly: true}); // Processed 'Y late cancel/no-show'
  shFormat.getRange("F13").copyTo(shNonCancel.getRange(3,1,lastRow-2,lastColumn) , {formatOnly: true}); // Processed 'N no payment info'
  shFormat.getRange("F14").copyTo(shNonCancel.getRange(3,1,lastRow-2,lastColumn) , {formatOnly: true}); // Processed 'N special billing'

  // Create Ticket numbers


  // Create Ticket names


  // Run finishing
  //finishing(true);

}

function finishing(nonCancel) {

  if (nonCancel) {
    var shWorking = SpreadsheetApp.getSheetByName("onlyNonCancelled");
  }
  else {
    var shWorking = SpreadsheetApp.getSheetByName("onlyCancelled");
  }

  var lastRow = shWorking.getLastRow();
  var lastColumn = shWorking.getLastColumn();

  // Increase Row height
  var rowHeight = 30; // desired row height in pixelst
  for (i=3; i<=lastRow; i++) {
    shWorking.setRowHeight(i, rowHeight);
  }

  // Sort by Location, then Status, then Duration, then Client
  shWorking.getRange(3, 1, lastRow-2, lastColumn).sort([{column: 1, ascending: true}, {column: 2, ascending: true}, {column: 5, ascending: true}, {column: 8, ascending: true}, {column: 3, ascending: true}]);

  // Add Column for Prices, and new Row at the bottom for total price
  // NEED TO CREATE DATABASE FOR SERVICES WITH PRICES

  // Alternating colors
  lastRows = shWorking.getLastRow(); // update lastRow
  lastColumn = shWorking.getLastColumn(); // update lastColumn
  shWorking.getRange(1, 1, 2, lastColumn).setBackground("#2c7fb4").setFontColor("#FFF").setFontWeight("bold"); // header color

  for (i=3; i<=lastRow; i+=2) {
    shWorking.getRange(i, 1, 1, lastColumn).setBackground("#FFF"); // white
  }
  for (j=4; j<=lastRow; j+=2) {
    shWorking.getRange(j, 1, 1, lastColumn).setBackground("#b4e3f6"); // light blue
  }

  // Resize Columns
  for (var i=1; i<=lastColumn; i++) {
    shWorking.autoResizeColumn(i);
  }

  // Delete Empty Rows and Columns at the end
  lastRow = shWorking.getLastRow(); // update lastRow
  lastColumn = shWorking.getLastColumn(); // update lastColumn
  maxRows = shWorking.getMaxRows(); // update maxRows
  maxColumns = shWorking.getMaxColumns(); // update maxColumns

  shWorking.deleteRows(lastRow + 1, maxRows - lastRow);
  shWorking.deleteColumns(lastColumn + 1, maxColumns - lastColumn);

  // Change Ticket column format to "Plain Text" and Left Align


}
