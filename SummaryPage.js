function populateSheetNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName("Summary");
  var sheets = ss.getSheets();
  
  // Clear out old data in column B from row 4 downward
  summarySheet.getRange("B4:B").clearContent();
  
  var row = 4;
  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName();
    // Skip the Summary sheet so it doesn't list itself
    if (sheetName === "Summary") {
      continue;
    }
    summarySheet.getRange(row, 2).setValue(sheetName);
    row++;
  }
}

function updateSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName("Summary");
  
  // 1) First, refresh the list of sheet names in Column B
  populateSheetNames();
  
  // 2) Now, update Columns C–I based on each sheet’s data
  var lastRow = summarySheet.getLastRow();
  var today = new Date();
  
  for (var row = 4; row <= lastRow; row++) {
    var sheetName = summarySheet.getRange(row, 2).getValue(); // Column B
    
    if (!sheetName) {
      continue;
    }
    
    var loanSheet = ss.getSheetByName(sheetName);
    if (!loanSheet) {
      continue;
    }
    
    var dataRange = loanSheet.getDataRange();
    var values = dataRange.getValues();
    
    // Find the row with the most recent date <= today
    var mostRecentDateIndex = -1;
    var mostRecentDate = new Date(1900, 0, 1);
    
    // Data rows start at row 8 in the sheet (index 7 in the array)
    for (var i = 7; i < values.length; i++) {
      var dateCell = values[i][3]; // Column D => index 3
      if (dateCell instanceof Date && dateCell <= today && dateCell > mostRecentDate) {
        mostRecentDate = dateCell;
        mostRecentDateIndex = i;
      }
    }
    
    if (mostRecentDateIndex !== -1) {
      // 1) Last Due Date (Column C)
      var lastDueDate = mostRecentDate;
      
      // 2) Outstanding Principal => Column P (index 15)
      var outstandingPrincipal = values[mostRecentDateIndex][15];
      
      // 3) Accumulated Interest => Column O (index 14)
      var accumulatedInterest = values[mostRecentDateIndex][14];
      
      // 4) Sum Principal Paid => Column J (index 9)
      //    Sum Interest Paid => Column L (index 11)
      var principalPaid = 0;
      var interestPaid = 0;
      
      // 5) Sum Principal Due => Column I (index 8) [ADJUST IF NEEDED]
      //    Sum Interest Due  => Column K (index 10) [ADJUST IF NEEDED]
      var principalDueSum = 0;
      var interestDueSum = 0;
      
      // Loop from row 8 to the row with the most recent date
      for (var j = 7; j <= mostRecentDateIndex; j++) {
        var pPaidVal = values[j][9];   // Column J => Principal Paid
        var iPaidVal = values[j][11];  // Column L => Interest Paid
        var pDueVal  = values[j][8];   // Column I => Principal Due
        var iDueVal  = values[j][10];  // Column K => Interest Due
        
        if (typeof pPaidVal === "number") {
          principalPaid += pPaidVal;
        }
        if (typeof iPaidVal === "number") {
          interestPaid += iPaidVal;
        }
        if (typeof pDueVal === "number") {
          principalDueSum += pDueVal;
        }
        if (typeof iDueVal === "number") {
          interestDueSum += iDueVal;
        }
      }
      
      // 6) Past Due Principal & Past Due Interest
      //    If the difference is negative, store 0.
      var pastDuePrincipal = Math.max(0, principalDueSum - principalPaid);
      var pastDueInterest  = Math.max(0, interestDueSum - interestPaid);
      
      // Write results to Summary:
      // Column C: Last Due Date
      // Column D: Outstanding Principal
      // Column E: Accumulated Interest
      // Column F: Total Principal Paid
      // Column G: Total Interest Paid
      // Column H: Past Due Principal
      // Column I: Past Due Interest
      
      summarySheet.getRange(row, 3).setValue(lastDueDate);
      summarySheet.getRange(row, 4).setValue(outstandingPrincipal);
      summarySheet.getRange(row, 5).setValue(accumulatedInterest);
      summarySheet.getRange(row, 6).setValue(principalPaid);
      summarySheet.getRange(row, 7).setValue(interestPaid);
      summarySheet.getRange(row, 8).setValue(pastDuePrincipal);
      summarySheet.getRange(row, 9).setValue(pastDueInterest);
    } else {
      // If no date is found, clear columns C–I
      summarySheet.getRange(row, 3, 1, 7).clearContent();
    }
  }
}

function createLoanSummaryMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Summary Tools')
    .addItem('Populate Sheet Names', 'populateSheetNames')
    .addItem('Update Summary', 'updateSummary')
    .addToUi();
}