# LoanScript.js – Loan Schedule Library Core

## Overview

LoanScript.js contains the core logic for generating and managing loan amortization schedules in Google Sheets using Google Apps Script. It defines configuration constants, helper functions for date and schedule calculations, classes to construct and recalculate loan schedules, and global functions that tie everything together. This script is designed to be used as a library in a Google Sheets project, handling calculations for both fully amortizing and interest-only loans, fees, and balance updates as payments are made.

## Constants

- **SHEET_CONFIG**: A configuration object defining sheet layout and input/output locations.
  - **START_ROW / END_ROW**: The row range in the sheet where the loan schedule is output (defaults 8 to 500).
  - **COLUMNS**: Column indices for various schedule fields (e.g., `PERIOD=2` for column B, `DUE_DATE=4` for column D, etc.) to know which columns store period number, due date, interest, principal, balances, etc.
  - **INPUTS**: Cell references for input parameters on the sheet (row 4 by default). For example:
    - `PRINCIPAL`: `'D4'` (loan principal)
    - `INTEREST_RATE`: `'E4'` (annual rate)
    - `CLOSING_DATE`: `'F4'` (loan start date)
    - `TERM_MONTHS`: `'G4'`
    - `PRORATE`: `'H4'`
    - `PAYMENT_FREQ`: `'I4'`
    - `DAY_COUNT`: `'J4'`
    - `DAYS_PER_YEAR`: `'K4'`
    - `PREPAID_INTEREST_DATE`: `'L4'`
    - `AMORTIZE`: `'M4'`
    - `ORIG_FEE_PCT`: `'N4'`
    - `EXIT_FEE_PCT`: `'O4'`
    - `LOCK_INPUTS`: `'Q4'`
- **IPMT_COL / PPMT_COL**: Column indices (26 = Z, 27 = AA) used internally to temporarily store Excel-style IPMT/PPMT formula results for amortizing loan calculations.

## Helper Functions

### `daysBetween(startDate, endDate)`

- **Description**: Returns the integer number of days between two dates (ignoring the time of day). It calculates the difference in milliseconds, divides by 86,400,000 (milliseconds per day), and rounds to the nearest whole number.
- **Parameters**:
  - `startDate (Date)`: The starting date.
  - `endDate (Date)`: The ending date.
- **Returns**: *(Number)* The number of days between `startDate` and `endDate`. Returns 0 if either parameter is not a valid Date.
- **Example**:

  ```js
  // Difference between Jan 1, 2025 and Jan 15, 2025:
  var diff = daysBetween(new Date('2025-01-01'), new Date('2025-01-15'));
  Logger.log(diff);  // Outputs: 14
  ```

### `daysBetweenInclusive(startDate, endDate)`

- **Description**: Returns the inclusive day count between two dates (i.e. `daysBetween(startDate, endDate)` + 1). Use this when both the start and end dates should be counted.
- **Parameters**:
  - `startDate (Date)`: The starting date.
  - `endDate (Date)`: The ending date.
- **Returns**: *(Number)* The inclusive number of days.
- **Example**:

  ```js
  // Inclusive days between Jan 1 and Jan 15, 2025:
  var diffInc = daysBetweenInclusive(new Date('2025-01-01'), new Date('2025-01-15'));
  Logger.log(diffInc);  // Outputs: 15
  ```

### `getLastDayOfMonth(dateObj)`

- **Description**: Returns a new Date set to the last day of the month of the given `dateObj`.
- **Parameters**:
  - `dateObj (Date)`: A date within the target month.
- **Returns**: *(Date)* A date representing the last day of `dateObj`’s month.
- **Example**:

  ```js
  var d = new Date('2025-03-14');
  var endOfMonth = getLastDayOfMonth(d);
  Logger.log(endOfMonth);  // Outputs: Mon Mar 31 2025 ...
  ```

### `getLastDayAfterAddingMonths(dateObj, monthsToAdd)`

- **Description**: Calculates the last day of the month that is `monthsToAdd` months after the month of the given date. Handles year transitions.
- **Parameters**:
  - `dateObj (Date)`: The base date.
  - `monthsToAdd (Number)`: The number of months to advance.
- **Returns**: *(Date)* The last day of the resulting month.
- **Example**:

  ```js
  var d = new Date('2025-03-14');
  Logger.log(getLastDayAfterAddingMonths(d, 1));  // Outputs: Wed Apr 30 2025 ...
  Logger.log(getLastDayAfterAddingMonths(d, 12)); // Outputs: Tue Mar 31 2026 ...
  ```

### `oneDayAfter(dateObj)`

- **Description**: Returns a new Date that is exactly one day after the given date. The time is normalized to the same hour as in `dateObj`.
- **Parameters**:
  - `dateObj (Date)`: The reference date.
- **Returns**: *(Date)* The date representing the next calendar day.
- **Example**:

  ```js
  var d = new Date('2025-05-31');
  Logger.log(oneDayAfter(d));  // Outputs: Sun Jun 01 2025 ...
  ```

### `isEdgeDay(dateObj)`

- **Description**: Checks if the given date falls on a “month edge” (day is 1 or greater than 28). This affects loan schedule prorating logic.
- **Parameters**:
  - `dateObj (Date)`: The date to check (typically the loan closing date).
- **Returns**: *(Boolean)* `true` if the day is 1 or >28; otherwise, `false`.

### `isUnscheduledRow(rowArr)`

- **Description**: Determines if a given row from the schedule represents an "unscheduled payment" (a manually inserted payment outside the regular schedule).
- **Parameters**:
  - `rowArr (Array)`: An array of cell values representing one row of the schedule (columns B through R).
- **Returns**: *(Boolean)* `true` if the row does not have an integer period number but does have a payment date, otherwise `false`.

### `findLastScheduledEnd(schedule, rowIndex)`

- **Description**: Finds the last scheduled period end date before the given row index by searching upward in the schedule.
- **Parameters**:
  - `schedule (Array of Arrays)`: The full schedule data.
  - `rowIndex (Number)`: The index from which to search backward.
- **Returns**: The period end date of the nearest prior scheduled row, or `null` if none is found.

### `calcUnpaidDays(periodStart, periodEnd, prepaidUntil)`

- **Description**: Calculates the number of days in a period not covered by prepaid interest.
- **Parameters**:
  - `periodStart (Date)`: The start date of the period.
  - `periodEnd (Date)`: The end date of the period.
  - `prepaidUntil (Date|null)`: The date up to which interest is prepaid.
- **Returns**: *(Number)* The count of days in the period that are not prepaid.

### `getOverlapDays(rowIndex, schedule, params)`

- **Description**: Determines the number of days in a period that overlap with prepaid interest.
- **Parameters**:
  - `rowIndex (Number)`: The index of the period row.
  - `schedule (Array of Arrays)`: The schedule data.
  - `params (Object)`: The loan parameters object.
- **Returns**: *(Number)* The count of non-prepaid days for that row.

### `getAllInputs(sheet)`

- **Description**: Reads all the loan input values from the specified Google Sheet (expects inputs in the cells defined by `SHEET_CONFIG.INPUTS`) and constructs a parameters object. It enforces certain rules and computes derived values such as the financed origination fee and exit fee.
- **Parameters**:
  - `sheet (Sheet)`: The Google Sheets sheet object containing the loan inputs.
- **Returns**: *(Object)* An object with all necessary loan parameters for schedule generation, including derived values like `financedFee`, `financedPrepaidInterest`, and `exitFee`.

### `getTotalPeriods(params)`

- **Description**: Determines how many periods the amortization schedule will have based on the loan parameters and whether the first period is prorated.
- **Parameters**:
  - `params (Object)`: The parameters object returned by `getAllInputs`.
- **Returns**: *(Number)* The total count of periods to generate.

### `calcPeriodEndDate_Prorate(closingDate, i)`

- **Description**: Calculates the period end date for the i-th period when the first period is prorated.
- **Parameters**:
  - `closingDate (Date)`: The loan closing date.
  - `i (Number)`: The index of the period (0 for the first, 1 for the second, etc.).
- **Returns**: *(Date)* The end date for period `i`.
- **Example**:

  ```js
  // For a loan closing on March 15, 2025 with prorated first period:
  calcPeriodEndDate_Prorate(closingDate, 0);  // Returns March 31, 2025.
  calcPeriodEndDate_Prorate(closingDate, 1);  // Returns April 30, 2025.
  ```

### `calcPeriodEndDate_NoProrate(closingDate, periodNum)`

- **Description**: Calculates the period end date for a given period when the first period is not prorated.  
  - For the first period:
    - If the closing date is the 1st, the period ends on the last day of the same month.
    - If the closing date is after the 28th, the period ends on the last day of the following month.
    - Otherwise, the period ends the day before the closing date’s monthly anniversary.
  - For subsequent periods, similar logic applies based on the closing day.
- **Parameters**:
  - `closingDate (Date)`: The loan closing date.
  - `periodNum (Number)`: The period number (1 for first period, 2 for second, etc.).
- **Returns**: *(Date)* The end date for the given period.
- **Example**:

  ```js
  // For a loan closing on March 15, 2025 without prorating:
  calcPeriodEndDate_NoProrate(closingDate, 1);  // Returns April 14, 2025.
  // For a loan closing on Jan 30, 2025, the first period end might be Feb 29, 2025.
  ```

---

## Classes

### LoanScheduleGenerator

- **Description**: Generates the initial amortization schedule based on the sheet’s input parameters. It reads inputs, clears any existing schedule, builds a new schedule array, writes it to the sheet, applies formatting, and then calls balance recalculation.
- **Constructor**:  
  ```js
  new LoanScheduleGenerator(sheet)
  ```
  - **Parameters**:
    - `sheet (Sheet)`: The Google Sheet where the schedule will be generated.
- **Methods**:
  - **`generateSchedule()`**  
    Generates the complete loan schedule by:
    - Reading loan inputs via `getAllInputs`.
    - Clearing the existing schedule.
    - Building schedule data (using `buildScheduleData`).
    - Writing the schedule to the sheet.
    - Applying formatting (via `applyFormatting`).
    - Instantiating a `BalanceManager` to recalc balances.
  - **`clearOldSchedule()`**  
    Clears old schedule data from the output range.
  - **`buildScheduleData(params)`**  
    Constructs the amortization schedule as an array of rows. Each row includes period number, period end date, due date, days in period, payment placeholders, and notes. Also handles special fee notes.
  - **`applyFormatting(numRows)`**  
    Applies Google Sheets formatting to date, numeric, and text columns in the schedule.

### BalanceManager

- **Description**: Handles recalculation of the schedule by allocating payments to interest and principal, updating balances, and managing unscheduled or extra payments.
- **Constructor**:  
  ```js
  new BalanceManager(sheet)
  ```
  - **Parameters**:
    - `sheet (Sheet)`: The Google Sheet for the loan schedule.
- **Methods**:
  - **`recalcAll()`**  
    Recalculates the entire schedule by:
    - Reading schedule data.
    - Separating scheduled and unscheduled rows.
    - Calculating scheduled interest and principal using helper formulas.
    - Iterating through periods to update running balances.
    - Writing the updated values back to the sheet.
  - **`buildIpmtPpmtResults(schedule, lastUsedCount, params)`**  
    Internally writes and reads financial formulas (IPMT/PPMT) from hidden columns to compute scheduled interest and principal portions.

### RowManager

- **Description**: Manages manual modifications to the schedule—specifically, handling insertion of unscheduled payment rows.
- **Constructor**:  
  ```js
  new RowManager(sheet)
  ```
  - **Parameters**:
    - `sheet (Sheet)`: The Google Sheet for the loan schedule.
- **Method**:
  - **`handleInsertedRow(insertedRow)`**  
    Prepares a newly inserted row as an unscheduled payment row by:
    - Assigning a fractional period number.
    - Copying the previous row’s ending balances.
    - Initializing other fields (Period End, Due Date, Days, etc.) with default or blank values.

---

## Global Functions (Library Interface)

- **`generateLoanSchedule()`**  
  Clears the old schedule and generates a new amortization schedule on the active sheet using `LoanScheduleGenerator`.

- **`recalcAll()`**  
  Recalculates all interest, principal, and balance fields on the active sheet’s loan schedule by invoking `BalanceManager.recalcAll()`.

- **`insertUnscheduledPaymentRow()`**  
  Inserts a new row into the active sheet’s schedule for an unscheduled payment. The row is initialized via `RowManager.handleInsertedRow()`, then recalculates balances.

- **`onEdit(e)`**  
  Trigger function that runs on every user edit. It:
  - Regenerates the schedule if an input (row 4) is edited (and inputs are unlocked).
  - Recalculates balances if a cell in the schedule output area (rows 8+) is modified.
  - Ignores edits on the "Lock Inputs" cell or the "Summary" sheet.
  
- **`createOnEditTrigger()`**  
  Creates an installable onEdit trigger for the spreadsheet to ensure `onEdit(e)` is called.

- **`createLoanScheduleMenu()`**  
  Adds a custom menu (e.g., "Loan Tools") to the Google Sheets UI with items to generate the schedule, insert an unscheduled payment row, and recalculate balances.

---

# SummaryPage.js – Loan Summary Sheet Script

## Overview

SummaryPage.js manages a "Summary" sheet that aggregates key information from multiple loan sheets. It provides functions to:
- Populate a list of loan sheet names.
- Update summary statistics (e.g., outstanding principals, last payment dates).
- Add a custom menu for summary-related actions.

## Functions

### `populateSheetNames()`

- **Description**:  
  Scans the spreadsheet for all sheets (tabs) and populates the "Summary" sheet with the names of each loan sheet (excluding the "Summary" sheet itself). The sheet names are listed in column B starting from row 4.
- **Parameters**: None.
- **Usage**:  
  Run via a custom menu item ("Summary Tools → Populate Sheet Names") to refresh the list.

### `updateSummary()`

- **Description**:  
  Updates summary metrics for each loan listed on the "Summary" sheet. It:
  - Calls `populateSheetNames()` first.
  - For each loan sheet listed in column B, retrieves data from its loan schedule:
    - **Last Due Date**: (Column C of Summary)
    - **Outstanding Principal**: (Column D)
    - **Accumulated Interest**: (Column E)
    - **Total Principal Paid**: (Column F)
    - **Total Interest Paid**: (Column G)
    - **Past Due Principal**: (Column H)
    - **Past Due Interest**: (Column I)
  - Clears summary data if a loan sheet has no valid schedule or data.
- **Parameters**: None.
- **Usage**:  
  Invoked via a custom menu item ("Summary Tools → Update Summary").

### `createLoanSummaryMenu()`

- **Description**:  
  Adds a custom menu labeled "Summary Tools" to the Google Sheets UI with options for:
  - Populating sheet names.
  - Updating summary metrics.
- **Parameters**: None.
- **Usage**:  
  Called on spreadsheet open (e.g., via an `onOpen` trigger) to ensure the menu is available.

---

# LoanScriptWrapper.js – Google Sheets Script Wrapper

## Overview

LoanScriptWrapper.js is a thin wrapper intended to be placed in the Google Sheet’s Apps Script project. It connects the user interface (menus and triggers) with the functions provided by the Loan Script Library. The wrapper defines functions that forward calls to the corresponding library functions. When using the library by its project key, ensure you replace the placeholder identifier (e.g., `LoanScriptLibrary`) with the actual name assigned to the library.

## Functions

### `onOpen(e)`

- **Description**:  
  A trigger that runs when the spreadsheet is opened. It adds custom menus for both loan schedule actions and summary actions by calling:
  - `LoanScriptLibrary.createLoanSummaryMenu()`
  - `LoanScriptLibrary.createLoanScheduleMenu()`
- **Parameters**:
  - `e (Event)`: The onOpen event object.
- **Usage**:  
  Automatically invoked on spreadsheet open.

### `generateLoanSchedule()`

- **Description**:  
  Calls the library’s `generateLoanSchedule()` function to create or refresh the loan schedule on the active sheet.
- **Parameters**: None.
- **Usage**:  
  Typically triggered via the menu item "Loan Tools → Generate Schedule".

### `insertUnscheduledPaymentRow()`

- **Description**:  
  Inserts a new unscheduled payment row into the active sheet’s schedule by calling `LoanScriptLibrary.insertUnscheduledPaymentRow()`.
- **Parameters**: None.
- **Usage**:  
  Triggered via the menu option "Add Unscheduled Payment".

### `recalcAll()`

- **Description**:  
  Recalculates the loan schedule on the active sheet by calling `LoanScriptLibrary.recalcAll()`.
- **Parameters**: None.
- **Usage**:  
  Invoked via the menu item "Recalculate Balances" or when a payment is recorded.

### `onEdit(e)`

- **Description**:  
  Forwards the edit event to the library’s `onEdit(e)` function. This ensures that any user edit in the sheet triggers the centralized library logic.
- **Parameters**:
  - `e (Event)`: The edit event object.
- **Usage**:  
  Automatically invoked on every edit when installed as a trigger.

### `populateSheetNames()`

- **Description**:  
  Calls `LoanScriptLibrary.populateSheetNames()` to update the list of loan sheet names on the Summary sheet.
- **Parameters**: None.
- **Usage**:  
  Invoked via the menu option "Summary Tools → Populate Sheet Names".

### `updateSummary()`

- **Description**:  
  Calls `LoanScriptLibrary.updateSummary()` to refresh summary data for all loans.
- **Parameters**: None.
- **Usage**:  
  Invoked via the menu option "Summary Tools → Update Summary".

### `setupTriggers()`

- **Description**:  
  Installs the necessary triggers by calling `LoanScriptLibrary.createOnEditTrigger()`. This should be run once to ensure that the onEdit trigger is correctly set up.
- **Parameters**: None.
- **Usage**:  
  Run manually (or via a one-time menu action) after the library is installed.

---
