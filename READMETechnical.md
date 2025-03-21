# LoanScript.js – Loan Schedule Library Core

## Overview:
LoanScript.js contains the core logic for generating and managing loan amortization schedules in Google Sheets using Google Apps Script. It defines configuration constants, helper functions for date and schedule calculations, classes to construct and recalculate loan schedules, and global functions that tie everything together. This script is designed to be used as a library in a Google Sheets project, handling calculations for loan schedules (both fully amortizing and interest-only loans), fees, and balance updates as payments are made.

## Constants
SHEET_CONFIG: A configuration object defining sheet layout and input/output locations. Key fields include:
- **START_ROW / END_ROW:** The row range in the sheet where the loan schedule is output (defaults 8 to 500).
- **COLUMNS:** Column indices for various schedule fields (e.g., PERIOD=2 for column B, DUE_DATE=4 for column D, etc.), allowing the script to know which columns store period number, due date, interest, principal, balances, etc.
- **INPUTS:** Cell references for input parameters on the sheet (row 4 by default). For example, PRINCIPAL: 'D4' (loan principal), INTEREST_RATE: 'E4' (annual rate), CLOSING_DATE: 'F4' (loan start date), TERM_MONTHS: 'G4', PRORATE: 'H4', PAYMENT_FREQ: 'I4', DAY_COUNT: 'J4', DAYS_PER_YEAR: 'K4', PREPAID_INTEREST_DATE: 'L4', AMORTIZE: 'M4', ORIG_FEE_PCT: 'N4', EXIT_FEE_PCT: 'O4', and LOCK_INPUTS: 'Q4'. These are the inputs the user provides in the sheet.
- **IPMT_COL / PPMT_COL:** Column indices (26 = Z, 27 = AA) used internally to temporarily store Excel-style IPMT/PPMT formula results for amortizing loan calculations. These help compute scheduled interest and principal for each period.

## Helper Functions

### daysBetween(startDate, endDate)
**Description:** Returns the integer number of days between two dates, ignoring the time of day. It calculates the difference in milliseconds and divides by 86,400,000 (ms per day), then rounds to the nearest whole number.  
**Parameters:**
- startDate (Date): The starting date.
- endDate (Date): The ending date.  
**Returns:** (Number) The number of days between startDate and endDate. If either parameter is not a valid Date object, it returns 0.  
**Example:**

```js
// Difference between Jan 1, 2025 and Jan 15, 2025:
var diff = daysBetween(new Date('2025-01-01'), new Date('2025-01-15'));
Logger.log(diff);  // Outputs: 14
```

### daysBetweenInclusive(startDate, endDate)
**Description:** Returns the inclusive day count between two dates, effectively daysBetween(startDate, endDate) + 1. Use this when both the start and end dates should be counted in the interval.  
**Parameters:**
- startDate (Date): The starting date.
- endDate (Date): The ending date.  
**Returns:** (Number) The number of days between the dates, inclusive of both start and end.  
**Example:**

```js
// Inclusive days between Jan 1 and Jan 15, 2025:
var diffInc = daysBetweenInclusive(new Date('2025-01-01'), new Date('2025-01-15'));
Logger.log(diffInc);  // Outputs: 15
```

### getLastDayOfMonth(dateObj)
**Description:** Given a Date object, returns a new Date set to the last day of that same month. This is useful for finding period end dates when scheduling monthly periods.  
**Parameters:**
- dateObj (Date): Any date within the target month.  
**Returns:** (Date) A date representing the last day of dateObj’s month (with the same year).  
**Example:**

```js
var d = new Date('2025-03-14');
var endOfMonth = getLastDayOfMonth(d);
Logger.log(endOfMonth);  // Outputs: Mon Mar 31 2025 ...
```

### getLastDayAfterAddingMonths(dateObj, monthsToAdd)
**Description:** Calculates the last day of the month that is monthsToAdd months after the month of the given date. For example, if dateObj is March 14, 2025 and monthsToAdd is 1, this returns April 30, 2025. This handles year transitions as well.  
**Parameters:**
- dateObj (Date): The base date.
- monthsToAdd (Number): How many months to advance from dateObj’s month.  
**Returns:** (Date) The last day of the resulting month after adding the specified months.  
**Example:**

```js
var d = new Date('2025-03-14');
Logger.log( getLastDayAfterAddingMonths(d, 1) );  // Outputs: Wed Apr 30 2025 ...
Logger.log( getLastDayAfterAddingMonths(d, 12) ); // Last day 12 months later: Tue Mar 31 2026 ...
```

### oneDayAfter(dateObj)
**Description:** Returns a new Date that is exactly one day after the given date. Time is not preserved (the result is normalized to the next day at the same hour as dateObj).  
**Parameters:**
- dateObj (Date): The reference date.  
**Returns:** (Date) A new date representing the next calendar day after dateObj.  
**Example:**

```js
var d = new Date('2025-05-31');
Logger.log( oneDayAfter(d) );  // Outputs: Sun Jun 01 2025 ...
```

### isEdgeDay(dateObj)
**Description:** Checks if a given date falls on a “month edge” that affects loan schedule prorating logic. It returns true if the day of the month is 1 or greater than 28. This is used to decide if the first period should be treated normally (no prorate) even if user asked for prorate, because closing on day 29, 30, 31, or 1 can be handled as full first month.  
**Parameters:**
- dateObj (Date): The date to check (usually the loan closing date).  
**Returns:** (Boolean) True if the date’s day is 1 or >28 (29th, 30th, 31st), otherwise false.  
**Usage:** If this function returns true for the loan’s closing date, the script will internally force prorateFirst to "No" (no prorated first period) to avoid extremely short first periods.

### isUnscheduledRow(rowArr)
**Description:** Determines if a given row from the schedule is an "unscheduled payment" row (a manually inserted payment outside the regular schedule). It checks the row data to see if the period identifier is not an integer but the payment date is present.  
**Parameters:**
- rowArr (Array): An array of cell values representing one row of the schedule (columns B through R).  
**Returns:** (Boolean) True if the row does not have an integer period number but does have a payment date (indicating it's an extra/unscheduled payment row added by the user). False for regular scheduled rows.  
**Note:** This is used internally when recalculating to separate regular schedule rows from extra payment rows.

### findLastScheduledEnd(schedule, rowIndex)
**Description:** In the context of recalculation, finds the last scheduled period end date before a given row index. It searches upward in the schedule array from rowIndex to find a prior row that is part of the original schedule (as opposed to an unscheduled row).  
**Parameters:**
- schedule (Array of Arrays): The full schedule data as a 2D array (each sub-array is a row of values).
- rowIndex (Number): The index in schedule from which to search backwards.  
**Returns:** The period end date (from column C) of the nearest prior scheduled row. If none is found, returns null.  
**Usage:** This is used to calculate interest accrual for unscheduled payments by knowing where the last scheduled period ended.

### calcUnpaidDays(periodStart, periodEnd, prepaidUntil)
**Description:** Calculates the number of days in a period that are not covered by prepaid interest. If a loan has interest prepaid up to a certain date, this function computes how many days of the period [periodStart, periodEnd] still accrue interest normally.  
**Parameters:**
- periodStart (Date): The start date of the period (often one day after the last period’s end).
- periodEnd (Date): The end date of the period.
- prepaidUntil (Date or null): The date up to which interest is prepaid (from loan inputs). May be null if no prepaid interest.  
**Returns:** (Number) The count of days in [periodStart, periodEnd] that are not covered by prepaid interest. If prepaidUntil is null or before periodStart, it returns the full days between periodStart and periodEnd. If prepaidUntil extends into this period, the overlap is excluded.  
**Usage:** Used when building the schedule to adjust interest calculations for loans where some interest was prepaid.

### getOverlapDays(rowIndex, schedule, params)
**Description:** Determines how many days of a particular period overlap with prepaid interest. It is a helper used during recalculation when handling unscheduled payments that might occur mid-period. Essentially, it uses calcUnpaidDays for a specific row in the schedule.  
**Parameters:**
- rowIndex (Number): The index of the row (period) in question.
- schedule (Array of Arrays): The full schedule data up to current point.
- params (Object): The loan parameters object (from getAllInputs) which includes prepaidUntil date, etc.  
**Returns:** (Number) Days in that row’s period that are not prepaid (similar to calcUnpaidDays result for that row).  
**Note:** Typically used internally when recalculating interest for a period that had an extra payment.

### getAllInputs(sheet)
**Description:** Reads all the loan input values from the specified Google Sheet (expects inputs in the cells defined by SHEET_CONFIG.INPUTS) and constructs a parameters object. It also enforces certain rules and computes derived values like origination fee amount and exit fee.  
**Parameters:**
- sheet (Sheet): The Google Sheets sheet object containing the loan inputs in row 4.  
**Returns:** (Object) An object with all necessary loan parameters for schedule generation. Important properties in this object include:
- **principal (Number):** The initial principal (will be adjusted to include financed fees if any).
- **closingDate (Date):** The loan closing date.
- **annualRate (Number):** Annual interest rate (e.g., 0.05 for 5%).
- **paymentFreq (String):** Payment frequency, e.g., "Monthly" or "Single Period".
- **dayCountMethod (String):** "Actual" or "Periodic" for interest calculation.
- **daysPerYear (Number):** Days in a year for interest calculation (e.g., 365 or 360).
- **termMonths (Number):** Loan term in months.
- **prorateFirst (String):** "Yes" or "No" indicating if the first period is prorated. This function will set it to "No" if the closing date is an edge day (1 or >28) to override user choice for practicality.
- **amortizeYN (String):** "Yes" for amortizing loan, "No" for interest-only. If paymentFreq is "Single Period", this is forced to "No" (interest-only) because single period loans can’t amortize monthly.
- **prepaidIntDate (Date|null):** If a prepaid interest date is provided and valid, the Date object up to which interest is prepaid, otherwise null.
- **origFeePct (Number):** Origination fee percentage (e.g., 0.02 for 2%). If blank, treated as 0.
- **exitFeePct (Number):** Exit fee percentage of original principal.
- **origFeePctString (String):** The origination fee percentage in a readable string format (taken from cell display, e.g., "2%" if provided). Used for notes.  
**Derived values:** After reading initial inputs, this function calculates:
- **financedFee (Number):** The dollar amount of origination fee financed into the loan. This equals principal * origFeePct. The principal is increased by this amount (i.e., the fee is added on top of the original principal).
- **financedPrepaidInterest (Number):** If prepaidIntDate is provided, the amount of interest prepaid from closing date up to that date. The principal is increased by this amount as well (prepaid interest is treated as if it’s added to the loan balance).  
After adjusting principal for financed amounts, it calculates exitFee (Number) as exitFeePct * originalPrincipal (note: exit fee is based on the original principal before fees are added). This value will be used to populate the final period’s fee due.  
**Usage:** Call this at the start of schedule generation or recalculation to get a consistent set of parameters. It also updates the sheet if it had to override any inputs (for example, it will set the Prorate cell to "No" if isEdgeDay returned true, or ensure Day Count is "Actual" and Amortize "No" for single period loans).

### getTotalPeriods(params)
**Description:** Determines how many periods (rows) the amortization schedule will have, based on the loan parameters and whether the first period is prorated.  
**Parameters:**
- params (Object): The parameters object returned by getAllInputs. It looks at params.termMonths and params.prorateFirst.  
**Returns:** (Number) The total count of periods to generate. If prorateFirst === "Yes", it returns termMonths + 1 (an extra initial period), otherwise it returns termMonths as-is.  
**Example:** If a 12-month loan has prorateFirst="Yes", this will return 13 (the first prorated partial month plus 12 full months). If prorating is "No", it returns 12.

### calcPeriodEndDate_Prorate(closingDate, i)
**Description:** Calculates the period end date for the i-th period of a loan if the first period is being prorated. This logic is used when prorateFirst is "Yes".  
**Parameters:**
- closingDate (Date): The loan closing date (start date of the loan).
- i (Number): The index of the period (0 for the first, 1 for second, etc.).  
**Returns:** (Date) The end date for period i. If i === 0 (the first, prorated period), it returns the last day of the closing date’s month. If i > 0, it returns the last day of the month that is i months after the closing date’s month.  
**Example:** For a loan closing on March 15, 2025 with prorated first period, calcPeriodEndDate_Prorate(closingDate, 0) returns March 31, 2025. calcPeriodEndDate_Prorate(closingDate, 1) returns April 30, 2025 (end of second period, which is the first full month after the prorated month).

### calcPeriodEndDate_NoProrate(closingDate, periodNum)
**Description:** Calculates the period end date for a given period number when the first period is not prorated (prorateFirst = "No"). This function has special logic for the first period depending on the day of the month the loan closed.  
**Parameters:**
- closingDate (Date): The loan closing date.
- periodNum (Number): The period number (1 for first period, 2 for second, etc.).  
**Returns:** (Date) The end date for the given period. For periodNum === 1 (first period):
- If closing date is the 1st of a month, first period ends on the last day of that same month.
- If closing date is after the 28th, the first period ends on the last day of the following month (ensuring at least a full month duration for period 1).
- Otherwise (e.g., closing on the 15th), the first period ends the day before the closing date’s monthly anniversary. For example, closing on March 15 → first period ends on April 14.  
For subsequent periods (periodNum > 1):
- If closing day was 1: each period ends on the last day of the respective month (period 2 ends last day of next month, etc.).
- If closing day was > 28: each period end is simply the last day after adding periodNum months (similar to above logic for each period).
- Otherwise (normal case): It creates a date on the same day as closing date but periodNum months out, then subtracts one day to get the end date. (This effectively makes each period run from the 15th to the 14th of next month in the example above.)  
**Example:** If a loan closes on March 15, 2025 without prorating, calcPeriodEndDate_NoProrate(closingDate, 1) -> April 14, 2025 (first period). periodNum=2 -> May 14, 2025, and so on. If a loan closes on Jan 30, 2025, first period end will be Feb 29, 2025 (last day of following month, since 30th is >28th).

## Classes

The library defines several classes to encapsulate the generation and recalculation of the loan schedule. Users of the library typically do not need to instantiate these classes directly; instead, they are used internally by the global functions (like generateLoanSchedule() or recalcAll()). However, understanding their behavior can be useful for advanced customization or debugging.

### LoanScheduleGenerator

**Description:** This class is responsible for generating the initial amortization schedule based on the current sheet’s input parameters. It reads inputs, clears any existing schedule, builds a new schedule array, writes it to the sheet, applies formatting, and then calls a balance recalculation to initialize all balance fields.

**Constructor**

```js
new LoanScheduleGenerator(sheet)
```

**Parameters:**
- sheet (Sheet): The Google Sheets sheet object representing the loan sheet where the schedule will be generated. Typically this is the active sheet containing the loan inputs and schedule area.  

**Description:** Creates a LoanScheduleGenerator instance tied to a specific sheet. Usually only used internally; the global function generateLoanSchedule() will handle instantiation.

**generateSchedule()**  
**Description:** Generates a complete loan schedule on the sheet. This is the main method to call for schedule creation. It performs the following steps:
- Reads all loan inputs from the sheet (getAllInputs).
- Clears any existing schedule content from the sheet (to avoid mixing old and new data).
- Builds the schedule data array by iterating through each period (buildScheduleData).
- Writes the new schedule data to the sheet (populating rows starting at SHEET_CONFIG.START_ROW).
- Applies number/date formatting to the inserted schedule range for readability (applyFormatting).
- Instantiates a BalanceManager and calls recalcAll() to calculate initial balances, interest, and principal allocations for each period (especially important for amortizing loans where principal/interest breakdown is needed).

**Parameters:** None (it uses the sheet provided at construction).  
**Returns:** Nothing directly. It updates the sheet with the new schedule.  
**Usage Example:**

```js
// Assuming 'sheet' is a Google Sheet with loan inputs set up:
var generator = new LoanScheduleGenerator(sheet);
generator.generateSchedule();
// This will populate the schedule on the sheet based on current inputs.
```

**clearOldSchedule()**  
**Description:** Helper method that clears out the old schedule data from the sheet before a new schedule is written. It clears the content in the output range (rows START_ROW through END_ROW, columns B through R by default). This ensures no remnants of a previous schedule remain.  
**Parameters:** None. (It uses the sheet bound to the instance.)  
**Returns:** None.  
**Note:** This method is called internally by generateSchedule() before writing new data.

**buildScheduleData(params)**  
**Description:** Constructs the amortization schedule as an array of rows (arrays). It uses the loan parameters to generate each period's row. The row data includes period number, period end date, due date, days in period, and placeholders for payments, interest, fees, and balances which start at 0. Specifically:
- **Period Number (Col B):** For prorated schedules, the first row might be period 0 or 1 (depending on implementation, but effectively a special period). Otherwise periods start at 1. Unscheduled manual rows (not generated here) will use fractional period numbers like 0.5, 1.5 (set by RowManager).
- **Period End (Col C):** Calculated via calcPeriodEndDate_Prorate or calcPeriodEndDate_NoProrate depending on prorateFirst.
- **Due Date (Col D):** Simply one day after the period end (so payment is due the day after the period accrual ends).
- **Days (Col E):** The approximate number of days in the period. If it’s the first prorated period, it’s the inclusive days between closing date and period end. If using 30-day periods (dayCountMethod = "Periodic"), it sets 30 for each full month. If using actual days, it calculates actual days between last period’s end and this period’s end.
- **Paid On (Col F):** Initialized empty (user will fill actual payment date).
- **Total Due/Paid, Principal Due/Paid, Interest Due/Paid, Fees Due/Paid (Cols G through N):** All initialized to 0 for now. Total Due is 0 until interest/principal calculation, which happens in recalculation phase.
- **Interest Balance, Principal Balance, Total Balance (Cols O, P, Q):** These running balances will be calculated later by the BalanceManager. Set initially to 0 here.
- **Notes (Col R):** Initially empty for each period, but the first period’s note may be set to indicate any prepaid interest or origination fee added, and the last period’s note might indicate an exit fee.

The function also handles adding special notes and fees: if an origination fee or prepaid interest was financed into the principal, it appends a note in the first period’s Notes column explaining the addition. If an exit fee is present, it places the fee amount in the final period’s Fees Due and a note in the final Notes column indicating the exit fee.

**Parameters:**
- params (Object): Loan parameters (from getAllInputs).  
**Returns:** (Array of Arrays) The schedule data ready to be written to the sheet. Each sub-array corresponds to one row (period) with columns B through R values as described above. If no periods (e.g., termMonths = 0), it could return an empty array.  
**Usage:** This method is used internally. After obtaining the rows array from this, generateSchedule() writes it to the sheet.

**applyFormatting(numRows)**  
**Description:** Applies Google Sheets formatting to the newly written schedule data for consistency and readability. It formats date columns, numeric columns, and text columns appropriately:
- Columns C (Period End), D (Due Date), and F (Paid On) are set to date format MM/dd/yyyy.
- Column E (Days) is formatted as an integer (no decimal).
- Columns G through N (monetary amounts for due/paid) are formatted as currency ($#,##0.00).
- Columns O through Q (balances) are also formatted as currency.
- Column R (Notes) is formatted as plain text.
  
**Parameters:**
- numRows (Number): The number of schedule rows that were written (so formatting can be applied only to that range). If numRows <= 0, the function does nothing.  
**Returns:** None. It modifies cell formats on the sheet.  
**Usage:** Called internally right after writing the schedule values.

### BalanceManager

Description: This class handles recalculation of the schedule, particularly allocating payments to interest and principal, updating balances, and handling unscheduled payments or prepayments. After a schedule is generated (or when payments are recorded/edited), **BalanceManager.recalcAll()** will update each period’s due, paid, and balance fields according to the payments made.

**Constructor**

```js
new BalanceManager(sheet)
```

Parameters:

  * **sheet** (Sheet): The Google Sheets sheet object for the loan schedule to recalc.  

Description: Creates a BalanceManager tied to a specific sheet. Internally it also references `SHEET_CONFIG` for column indices and boundaries.

#### recalcAll()

Description: Recalculates the entire loan schedule on the sheet. This is typically called after the schedule is generated, or whenever a payment entry is modified (via the onEdit trigger or manually via a menu command). **Major steps performed:**

  * **Read Schedule Data:** Reads all rows of the schedule output range (columns B through R, from `START_ROW` down to `END_ROW`).
  * **Find Last Row of Data:** Determines how many of those rows are actually in use (continuous from the start) by finding the first blank Period cell.
  * **Separate Scheduled vs Unscheduled Rows:** Iterates through each used row. If the row has a valid period number and period end date, it's a *scheduled* row; if it has no period number but does have a payment date, it's treated as an *unscheduled payment* row (likely inserted by the user for an extra payment). Two lists are built: one for `scheduledRows` and one for `unscheduledRows`.
  * **Sort Rows (for calculation order):** Scheduled rows are sorted by due date, and unscheduled (extra payment) rows are sorted by actual payment date. This ensures payments are applied in chronological order when updating balances.
  * **Calculate Amortization (if needed):** If the loan is amortizing (not interest-only), the script uses spreadsheet formulas to calculate the scheduled interest (IPMT) and principal (PPMT) portions for each period. This is done by writing IPMT and PPMT formulas to hidden columns (e.g. Z and AA) for each period with the loan parameters (principal, rate, etc.), retrieving their results, then clearing those helper cells. The results are stored in memory for use in the next steps. (If the loan is interest-only, these values are not needed since scheduled principal due will be 0 until the final period.)
  * **Prepare Re-amortization Tracking:** Creates an array `hasReAmortized` to mark if a period’s schedule was re-amortized due to an unscheduled payment (used in complex scenarios of multiple prepayments). Initially all values are **false**.
  * **Initialize Running Balances:** Sets up running totals for principal and interest. The starting principal is the loan principal (including any financed fees) from inputs; starting accrued interest is 0.
  * **Iterate Through Periods:** For each period in chronological order (processing unscheduled payments in between as they occur):
    * **Scheduled periods:** Calculate interest accrued for the period = `runningPrincipal * periodicRate` (depending on day-count method). For interest-only loans, principal due is 0 (except possibly in the final period). For amortizing loans, use the precomputed `scheduledPr` and `scheduledInt` from PPMT/IPMT for that period if no prior prepayment affected the schedule. If a prior unscheduled payment (prepayment) occurred, the remaining balance is lower; the algorithm reduces the interest due for this period accordingly and increases the principal due by the difference (the total payment due remains equal to the originally scheduled amount). By default, future scheduled payment amounts remain as initially calculated (resulting in the loan being paid off early if extra payments were made).
    * **Apply any payments:** If an unscheduled payment row is encountered (or if the user entered an actual payment on a scheduled row), allocate that payment to fees, then interest, then principal. Underpayments/overpayments are handled: if Total Paid is less than Total Due, the shortfall remains as unpaid interest (accruing to next period); if Total Paid is greater, the extra amount reduces principal ahead of schedule.
    * **Update balances:** The interest balance (`runningInterest`) carries over any unpaid interest. The principal balance (`runningPrincipal`) is reduced by any principal paid. The code ensures principal never goes below 0 (floors at 0).
    * **If loan pays off early:** If a prepayment (or combination of payments) fully pays off the remaining principal **before** the end of the term, the script will zero out any subsequent scheduled periods (setting their Principal Due, Interest Due, and Total Due to 0) since the loan is now fully repaid, and it breaks out of the loop.
  * **Write back calculated values:** Updates the schedule’s cells with the newly calculated amounts and balances for each period. This includes Total Due (col G = interest due + principal due + fees due), Principal Paid/Interest Paid (if an actual payment was entered), and the Interest Balance (col O), Principal Balance (col P), and Total Balance (col Q) for each period after applying payments.

**Note:** By default, extra payments on amortizing loans will shorten the loan (you’ll pay off earlier), while the scheduled payment amounts remain unchanged. If you want to re-amortize the remaining loan after a prepayment (i.e. adjust future payment amounts to the new balance), you can use the `recastLoan()` function (described below) to recalculate the schedule for all future periods.

Parameters: None (uses the sheet provided in the constructor and reads inputs via `getAllInputs` internally).  
Returns: None. The sheet’s schedule is updated in place.  
Usage: This method is called internally by triggers or menu actions. For example, if a user records a payment or inserts an unscheduled payment row, the script will invoke `BalanceManager.recalcAll()` to update the schedule. If needed, one could manually call `new BalanceManager(sheet).recalcAll()` to recalc a sheet’s loan balances after editing payments.

#### buildIpmtPpmtResults(schedule, lastUsedCount, params)

Description: *(Internal helper method)* Uses Google Sheets financial formulas to calculate the scheduled interest and principal portions for each period of a fully amortizing loan. It dynamically writes formulas into hidden cells to leverage the spreadsheet’s **IPMT**/**PPMT** functions, then retrieves the results. Specifically:

  * Constructs IPMT formulas of the form `=-IPMT(rate, periodNum, totalPeriods, principal)` for each period, and PPMT formulas of the form `=-PPMT(rate, periodNum, totalPeriods, principal)`. (The negative sign is used to get positive values, since IPMT/PPMT normally return negatives for outgoing payments.)
  * Writes these formulas to the sheet (in hidden helper columns, e.g. Z and AA) for all periods, then immediately reads the calculated values with `getValues()`.
  * Clears the helper formulas from those cells (to avoid leaving any artifacts on the sheet).
  * Returns two arrays: one for interest portions and one for principal portions, each of size `[lastUsedCount x 1]`. These arrays contain the scheduled interest and principal for each period (or empty strings for periods where amortization doesn’t apply, such as interest-only periods). These results are then mapped to their corresponding period indices in memory for use during recalculation.

Parameters:

  * **schedule** (Array of Arrays): The raw schedule data (not heavily used except for counting periods).
  * **lastUsedCount** (Number): How many rows of the schedule are active (number of periods).
  * **params** (Object): Loan parameters (uses `params.monthlyRate` for the interest rate per period, `params.termMonths` for total periods, etc.).  

Returns: `[ipmtVals, ppmtVals]` — two 2D arrays (each of dimensions lastUsedCount × 1), containing the interest and principal portions for each period. (For periods where no amortization applies, these may be empty strings.)  
**Note:** This method is called inside **recalcAll()** for amortizing loans to get the payment breakdown. It is not typically called on its own.

### RowManager

Description: This class helps manage manual modifications to the schedule, specifically when the user inserts a new row in the schedule to record an unscheduled payment. The **RowManager.handleInsertedRow()** method will initialize the new row with appropriate values and adjust the period numbering as needed.

**Constructor**

```js
new RowManager(sheet)
```

Parameters:

  * **sheet** (Sheet): The Google Sheets sheet object for the loan schedule.  

Description: Creates a RowManager instance for a given sheet.

#### handleInsertedRow(insertedRow)

Description: Prepares a newly inserted row (at position `insertedRow`) in the schedule to represent an unscheduled payment. When the user inserts a blank row in the schedule area, this function should be called to set it up. It performs the following:

  * **Assign Period Number:** Determines an appropriate period identifier for the inserted row that lies between the previous and next period. If the adjacent rows have period numbers, it averages them (e.g. if inserted between period 1 and 2, the new period becomes 1.5). If inserted at the very top or bottom of the schedule, it uses 0.5 or lastPeriod+0.5 accordingly. The period number is placed in column B of the new row.
  * **Copy Balances:** Copies the ending Interest Balance (col O) and Principal Balance (col P) from the row above (the previous row) into the new row as the starting balances for the unscheduled payment period. It also sets Total Balance (col Q) as the sum of those, ensuring the new row starts with the correct carry-over balances.
  * **Initialize Other Fields:** Sets the new row’s Period End, Due Date, and Days in Period to blank (because this row isn’t a scheduled period with its own due date). Sets Paid On to blank (the user will fill in the actual payment date). It sets Total Due (col G) to 0 (since this unscheduled row doesn’t have a pre-computed payment due; it will be filled after the payment is entered), Total Paid to 0, and Principal Due and Interest Due to blank (not applicable for an unscheduled row). Fees Due is set to blank as well. Principal Paid, Interest Paid, and Fees Paid are initialized to 0. The Notes column is left empty.

Parameters:

  * **insertedRow** (Number): The sheet row number where a new row has been inserted (1-indexed, e.g., 9 if inserted after sheet row 8). This should be within the schedule area (`START_ROW` to `END_ROW`).  

Returns: None. This function directly populates the cells in the newly inserted row on the sheet.  
Usage: Typically triggered by a custom menu item or via an onEdit trigger when detecting a specific user action. For example, if a user wants to record an extra payment between period 1 and 2, they would insert a blank row in the sheet (say, new sheet row 9 if row 8 was period 1 and row 10 was period 2). After insertion, the script should run `RowManager.handleInsertedRow(9)` to label it as period 1.5 and carry over the balances from the previous period.

## Global Functions (Library Interface)

The following functions are exposed globally by the library and serve as the interface for the Google Sheets script to use. They typically create instances of the above classes or coordinate the overall process. Each can be called either by custom menu items, buttons, or triggers in the Google Sheet.

### generateLoanSchedule()
**Description:** Clears the old schedule (if any) and generates a new amortization schedule on the active sheet, based on the input parameters in that sheet’s row 4. Internally, this function obtains the active sheet and uses LoanScheduleGenerator to produce the schedule and initialize balances. It is the primary function to create or refresh a loan’s schedule.  
**Parameters:** None. (Operates on the currently active spreadsheet and sheet).  
**Returns:** None. After execution, the active sheet’s schedule area (rows 8 onward) will be filled with the new schedule.  
**Usage Example:**

```js
// In the Google Sheets script context:
generateLoanSchedule();
// This will read inputs from the active sheet (row 4) and generate the schedule.
```

Typically, this is invoked via a menu item (e.g., "Generate Schedule") or can be called manually from the script editor for testing.

### recalcAll()
**Description:** Recalculates all interest, principal, and balance fields for the active sheet’s loan schedule. This should be run whenever payment information is added or changed. Internally, it creates a BalanceManager for the active sheet and calls its recalcAll() method to update the schedule. This ensures that any payments (including unscheduled ones) are properly applied and the remaining balances and accrued interest are correct.  
**Parameters:** None. (Uses the active sheet by default).  
**Returns:** None. The sheet’s schedule will be updated in place.  
**Usage:** Usually triggered by a menu action like "Recalculate Balances" or automatically via onEdit. For example, if the user records a payment in the sheet, onEdit will catch and recalc balances.

### insertUnscheduledPaymentRow()
**Description:** Inserts a new row in the active sheet’s schedule at the currently selected position to record an unscheduled payment, and initializes it. When called, it will take the currently active cell’s row as the insertion point. It inserts a row after that, then uses RowManager to set up the new row (assigning a fractional period number and carrying down balances). After inserting and initializing the row, it calls recalcAll() to update the schedule with this new row factored in.  
**Parameters:** None (operates based on active sheet and active cell selection).  
**Returns:** None. A new row will appear in the sheet’s schedule with period like “1.5” (for example) and zeroed out fields ready for input. Balances from the prior row will be carried into it.  
**Usage:** This is typically tied to a custom menu item like "Add Unscheduled Payment Row". The user would select a cell in the schedule (usually a row after which they want to add the payment) and trigger this function. The script will handle the rest (insertion, setup, recalculation).  
**Example:** Suppose you want to add an extra payment after period 5. You click any cell in period 5’s row and then run insertUnscheduledPaymentRow(). The script will insert a new row after period 5, label it as 5.5, copy down balances, and recalc the sheet.

### onEdit(e)
**Description:** A trigger function that runs whenever the user edits the spreadsheet (if a trigger is installed or for simple trigger in a bound script context). This function handles dynamic updates: if the user edits certain key cells, it will automatically regenerate or recalc the schedule. Specifically:
- If an input in row 4 (the loan parameters) is edited and the "Lock Inputs" (Q4) is set to "No", it will automatically call generateLoanSchedule() to regenerate the schedule with the new inputs. If inputs are locked (Q4 = "Yes"), then editing row 4 is not allowed – the script will immediately revert the change and show an alert informing the user that inputs are locked (and need to be unlocked to edit).
- If the user edits any cell in the schedule output area (rows 8 and below) in one of the following columns: Paid On (F), Total Paid (H), Principal Paid (J), Interest Paid (L), Fees Due (M), or Fees Paid (N), the script will trigger a recalculation by calling recalcAll(). These are the editable fields that affect balances. For example, entering an actual payment date or amount, or marking a fee due as applied, will prompt the schedule to update accordingly.
- If the user edits the "Lock Inputs" cell (Q4) itself, the script ignores it (no action on toggling the lock except to enforce it on other edits).
- The script also ignores edits on the "Summary" sheet (to avoid interference if the summary is present).  
**Parameters:**
- e (Event): The edit event object passed by the trigger, which includes e.range (the cell range that was edited), among other properties.  
**Returns:** None directly. (It performs actions on the sheet or shows alerts.)  
**Behavior Details:** This function is intended to be installed as an onEdit trigger (either simple or installable). If used as a simple trigger in a bound script, it runs automatically but cannot use certain services like SpreadsheetApp.getUi() in case of errors; if used as an installable trigger, it can show alerts. In this code, onEdit is likely meant to be an installable trigger (since it calls SpreadsheetApp.getUi().alert() on errors or lock conditions).  
**Usage:** The user generally does not call onEdit(e) manually. Instead, it is set up as a trigger so that user actions in the sheet invoke it. For example, if the user types a new value for interest rate in E4, onEdit will catch that and regenerate the schedule. If the user records a payment in H10 (Total Paid for period 2), onEdit will catch and recalc balances.

### createOnEditTrigger()
**Description:** Creates an installable onEdit trigger for the current spreadsheet bound to the library’s onEdit function. This is used because simple triggers in library code may not fire by default when used as a library. By running this function once, it programmatically ensures that an onEdit trigger is set up.  
**Parameters:** None.  
**Returns:** None. (On success, the trigger is created; on subsequent runs, multiple triggers might be created if not careful, so it’s typically run just once.)  
**Usage:** This can be executed manually (or via a one-time menu action) to set up the onEdit trigger. For example, the wrapper script provides a setupTriggers() function that calls this. When run, Google will ask for authorization to set up triggers if not already granted. After running, any edit on the sheet will fire the library’s onEdit as described above. Normally, you would call createOnEditTrigger() once after setting up the sheet for the first time.

### createLoanScheduleMenu()

Description: Adds a custom menu to the Google Sheets UI for loan schedule actions. The menu is typically labeled "Loan Tools" (or similar) and contains items to generate the schedule, insert an unscheduled payment row, recalculate balances, and recast the remaining schedule. This function uses the SpreadsheetApp UI service to create the menu and link each item to the corresponding function. For example, menu items like **"Generate Schedule" → generateLoanSchedule**, **"Add Unscheduled Payment" → insertUnscheduledPaymentRow**, **"Recalculate Balances" → recalcAll()**, and **"Recast Loan" → recastLoan** are added under the "Loan Tools" menu.

Parameters: None.  
Returns: None. The menu is added to the spreadsheet’s interface.  
Usage: This function should be called when the spreadsheet is opened. In practice, the wrapper’s **onOpen** trigger calls `LoanScriptLibrary.createLoanScheduleMenu()` to build the menu for the user. (If implementing without the provided wrapper, a bound script’s onOpen could call this library function to achieve the same result.)

# SummaryPage.js – Loan Summary Sheet Script

## Overview:
SummaryPage.js contains functions to manage a "Summary" sheet that aggregates key information from multiple loan sheets. This script is optional but useful if you have several loan tabs and want an overview of all loans (e.g., outstanding principals, last payment dates, etc.). It provides functions to populate a list of loan sheet names and update summary statistics for each loan, as well as a menu to trigger these actions easily.

## Functions

### populateSheetNames()
**Description:** Scans the spreadsheet for all sheets (tabs) and populates the "Summary" sheet with the names of each loan sheet. It skips the "Summary" sheet itself to avoid listing it. By default, it will list the sheet names in column B of the Summary sheet, starting from row 4 downward (clearing any previous contents in that range first). This establishes a list of loans to be referenced for summary calculations.  
**Parameters:** None. (The function assumes there is a sheet named "Summary" where the data will go.)  
**Returns:** None. It writes the list of sheet names into the Summary sheet.  
**Usage Example:** After adding a new loan sheet or renaming sheets, run populateSheetNames() to refresh the list. This can be invoked via the custom menu "Summary Tools -> Populate Sheet Names". Each loan sheet name will appear in the Summary sheet, one per row starting at B4.

### updateSummary()
**Description:** Updates the summary metrics for each loan listed on the Summary sheet. It performs the following steps:
- Calls populateSheetNames() first to ensure the sheet list (Column B) is up to date.
- For each sheet name listed in column B (from row 4 down to the last filled row in that column), it opens that sheet and gathers data from its loan schedule. Specifically, it finds the most recent period up to today (the latest due date that is on or before today’s date) and uses that row as a point of reference. From that row, it retrieves:
  - Last Due Date (to Column C of Summary): the due date of that most recent period.
  - Outstanding Principal (to Column D): the principal balance from that period (Column P on the loan sheet).
  - Accumulated Interest (to Column E): the interest balance from that period (Column O on the loan sheet).
  - Total Principal Paid (to Column F): the sum of all principal paid values up to and including that period.
  - Total Interest Paid (to Column G): the sum of all interest paid values up to and including that period.
  - Past Due Principal (to Column H): any principal that was scheduled (due) up to that period but not yet paid (i.e., Principal Due sum minus Principal Paid sum up to that point; if negative, it reports 0). This represents principal that should have been paid by now but wasn’t.
  - Past Due Interest (to Column I): similarly, interest that was due but remains unpaid as of that period (Interest Due sum minus Interest Paid sum, floored at 0).
- If a particular sheet has no valid data (for example, no schedule or no date <= today), it clears columns C–I for that row to indicate no data.  
**Parameters:** None.  
**Returns:** None. The Summary sheet will be updated with all the above metrics for each loan.  
**Usage:** Typically invoked via the custom menu "Summary Tools -> Update Summary". You might run this periodically or after recording new payments to refresh the portfolio overview. For example, after marking a payment on one of the loan sheets, go to Summary and click "Update Summary" to see the new outstanding balance and paid amounts reflected.

### createLoanSummaryMenu()
**Description:** Adds a custom menu to the spreadsheet UI for summary-related actions. The menu is labeled "Summary Tools". It provides at least two options: "Populate Sheet Names" (linked to populateSheetNames()) and "Update Summary" (linked to updateSummary()). Selecting these from the menu will run the corresponding function.  
**Parameters:** None.  
**Returns:** None. It creates the menu in the Google Sheets UI.  
**Usage:** Should be called on spreadsheet open to ensure the menu is available. In practice, the wrapper’s onOpen trigger calls LoanScriptLibrary.createLoanSummaryMenu() to add these options. If the Summary sheet is being used, this menu provides an easy way for users to refresh the summary data without going into the script editor.

# LoanScriptWrapper.js – Google Sheets Script Wrapper

## Overview:
LoanScriptWrapper.js is a thin wrapper script intended to be placed in the Google Sheet’s Apps Script project. Its main role is to connect the user interface (menus and triggers) with the functions provided by the Loan Script Library. The wrapper defines functions that call the corresponding library functions. This allows you to use the library in your sheet by adding this file and the library without having to write custom code. If you use the library by its project key, you might need to replace the placeholder library identifier (e.g., LoanScriptLibrary) with the actual name you assign to the library in your project. Each function in this wrapper either calls a function from the Loan Script Library or sets up triggers/menus that call those functions.

## Functions

### onOpen(e)

Description: A simple trigger that runs when the spreadsheet is opened. It adds custom menus for both loan schedule actions and summary actions by calling the library’s menu creation functions. Specifically, it calls `LoanScriptLibrary.createLoanSummaryMenu()` and `LoanScriptLibrary.createLoanScheduleMenu()` to add the **"Summary Tools"** and **"Loan Tools"** menus to the Google Sheets UI.  

Parameters:

  * **e** (Event): The onOpen event object (not used in this implementation, included only for completeness).  

Returns: None.  
Usage: This function is automatically invoked when the spreadsheet is opened (as long as it is set up as an onOpen trigger or the function exists in a bound script with this name). Through it, the user will see custom menu items like "Generate Schedule", "Add Unscheduled Payment Row", "Recalculate Balances", **"Recast Loan"**, "Populate Sheet Names", "Update Summary". No manual call by the user is needed; it runs on every open to ensure menus are present.

### generateLoanSchedule()

Description: Menu/utility function that calls the library’s `generateLoanSchedule()` to create or refresh the loan schedule on the active sheet. This is the function triggered by the "Generate Schedule" menu item.  

Parameters: None.  
Returns: None.  
Usage: The user clicks **Loan Tools → Generate Schedule** in the spreadsheet, which triggers this function. (It can also be run directly from the script editor for testing.) It requires that the Loan Script Library is added and accessible as `LoanScriptLibrary`. When invoked, it simply executes `LoanScriptLibrary.generateLoanSchedule()`, which generates a new schedule on the active sheet based on the input parameters.

### insertUnscheduledPaymentRow()

Description: Calls the library’s `insertUnscheduledPaymentRow()` function to add a new unscheduled payment row to the active sheet’s schedule. This is tied to the "Add Unscheduled Payment" menu option.  

Parameters: None.  
Returns: None.  
Usage: Triggered via the menu when the user wants to insert an extra payment row. It internally runs `LoanScriptLibrary.insertUnscheduledPaymentRow()`, which will handle inserting the row and initializing it as described above. After this function executes, a new row will appear in the schedule (with a period number like 1.5, etc.) ready for the user to enter the payment details. Balances from the prior period are carried into the new row.

### recalcAll()

Description: Calls the library’s `recalcAll()` function to recalculate the loan schedule on the active sheet. This corresponds to the "Recalculate Balances" menu item.  

Parameters: None.  
Returns: None.  
Usage: Invoked by the user through the menu or via an onEdit trigger. It will execute `LoanScriptLibrary.recalcAll()`, causing the active loan sheet’s balances and payment allocations to be refreshed. For example, if the user updates a payment amount or date in the schedule, this function (via the menu or trigger) will recompute all affected interest, principal, and balance fields.

### recastLoan()

Description: Calls the library’s `recastLoan()` function to re-amortize the remaining loan schedule on the active sheet. This corresponds to the "Recast Loan" menu command under "Loan Tools".  

Parameters: None.  
Returns: None.  
Usage: Typically invoked by the user via the menu after making a lump-sum prepayment on an amortizing loan. When the user selects **Loan Tools → Recast Loan**, this function runs and calls `LoanScriptLibrary.recastLoan()`. The library will determine the current outstanding principal and number of periods remaining, then recalculate the payment schedule for all future periods so that the loan is fully paid off by the original end date. After this function executes, each remaining scheduled period’s **Principal Due**, **Interest Due**, and **Total Due** are updated on the sheet to reflect the new amortized payment plan (usually resulting in lower periodic payments going forward, since the principal was reduced by the prepayment).

### onEdit(e)

Description: A simple trigger that runs when the user edits the spreadsheet. This wrapper function passes the event to the library’s onEdit handler. Essentially, it ensures that any edit events in the spreadsheet invoke the central library logic (which handles locking inputs, auto-regeneration, and auto-recalc as needed).  

Parameters:

  * **e** (Event): The edit event object.  

Returns: None.  
Usage: This is automatically invoked on every edit (when installed as an onEdit trigger). In this wrapper, it simply calls `LoanScriptLibrary.onEdit(e)`. By forwarding the event, it allows the library’s centralized logic to respond to user edits. The user should not need to call this manually.

### populateSheetNames()

Description: Calls the library’s `populateSheetNames()` function to list all loan sheet names on the Summary sheet. This is connected to the "Populate Sheet Names" menu item under "Summary Tools".  

Parameters: None.  
Returns: None.  
Usage: Invoked via the menu by the user (or can be run manually). It executes `LoanScriptLibrary.populateSheetNames()`, causing the Summary sheet to be updated with the current loan sheet names. Typically used after adding or renaming loan sheets, so the summary list stays up-to-date.

### updateSummary()

Description: Calls the library’s `updateSummary()` function to refresh the summary data for all loans. This is tied to the "Update Summary" menu item under "Summary Tools".  

Parameters: None.  
Returns: None.  
Usage: When the user selects **Summary Tools → Update Summary** from the menu, this function runs and simply calls `LoanScriptLibrary.updateSummary()`. The result is that the Summary sheet’s metrics (columns C–I listing balances, payments, etc. for each loan) are recalculated for each listed loan.

### setupTriggers()

Description: Installs the necessary triggers by calling the library’s `createOnEditTrigger()` function. Running this will create an *installable* onEdit trigger for the spreadsheet if one doesn’t exist, ensuring that the library’s onEdit function is properly called on user edits.  

Parameters: None.  
Returns: None.  
Usage: This function can be run once (for example, via the script editor or a one-time menu action) to set up the trigger after the library is installed. In the provided wrapper code, this isn’t automatically called on open; it’s provided as a utility. The user or developer may need to run `setupTriggers()` manually (or set up the trigger via the Apps Script interface). Once executed, the installable onEdit trigger will persist and call the library’s onEdit handler for each edit on any loan sheet.

**Note:** Ensure that the identifier `LoanScriptLibrary` in all the above calls matches the name of the library as added to your project. If you import the library under a different identifier, you will need to update the function calls in this wrapper (e.g., use your chosen name instead of `LoanScriptLibrary`). The wrapper functions themselves contain no logic besides forwarding to the library, so they rely on the library being added with the correct identifier.
```