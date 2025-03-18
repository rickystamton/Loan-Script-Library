/**
 * Wrapper for the Loan Schedule Library.
 * Replace "LoanLib" with your library's identifier.
 */

// Triggered when the spreadsheet is opened; adds the custom menu.
function onOpen(e) {
  LoanScriptLibrary.onOpen(e);
}

// Called (via the custom menu) to generate the loan schedule.
function generateLoanSchedule() {
  LoanScriptLibrary.generateLoanSchedule();
}

// Called (via the custom menu) to insert an unscheduled payment row.
function insertUnscheduledPaymentRow() {
  LoanScriptLibrary.insertUnscheduledPaymentRow();
}

// Called (via the custom menu) to recalculate loan fields.
function recalcAll() {
  LoanScriptLibrary.recalcAll();
}

// Triggered on edits in the spreadsheet.
function onEdit(e) {
  LoanScriptLibrary.onEdit(e);
}

function populateSheetNames() {
  LoanScriptLibrary.populateSheetNames();
}

function updateSummary() {
  LoanScriptLibrary.updateSummary();
}

function onOpen() {
  // Call the functions from your libraries that create the custom menus
  LoanScriptLibrary.createLoanSummaryMenu();
  LoanScriptLibrary.createLoanScheduleMenu();
}

function setupTriggers() {
  LoanScriptLibrary.createOnEditTrigger();
}