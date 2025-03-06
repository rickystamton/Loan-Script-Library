# Loan Script Library

## Project Overview

The Loan Script Library is a **Google Apps Script** solution for **Google Sheets** that generates detailed loan amortization schedules and helps track loan repayments. It is designed to manage loans directly within a Google Sheet – calculating periodic interest and principal, handling fees, and updating balances as repayments are made. With this script, you can input loan parameters into a sheet and automatically produce a schedule of payments, interest accruals, and outstanding balances. It supports both fully amortizing loans and interest-only structures, and includes features for prorated first periods, origination fees, exit fees, and more. All data is stored in Google Sheets, allowing you to review and modify loan details easily.

## Features

- **Amortization Schedule Generation**: Automatically creates a period-by-period loan schedule with columns for due dates, days in period, payment amounts, interest due, principal due, fees due, and remaining balances. The schedule supports monthly payments or a single lump-sum payment at maturity.
- **Interest Calculation Options**: Supports both actual day count and 30-day periodic calculations. You can specify the day count method and set the days-per-year (e.g., 365 or 360) to control how interest is accrued.
- **Interest-Only vs Amortizing Loans**: Toggle the *Amortize* parameter to generate either interest-only schedules (interest due each period and principal at the end) or fully amortizing schedules with equal periodic payments.
- **Origination and Exit Fees**: Automatically includes an origination fee (financed into the principal) and an exit fee (due in the final period) in the schedule.
- **Repayment Tracking**: Provides columns to record the actual payment date and amount for each period. The script adjusts outstanding balances based on these inputs, handling underpayments, extra payments, or prepayments.
- **Google Sheets Integration**: Designed to run entirely within Google Sheets with automatic formatting for dates and currency.
- **User-Editable Parameters**: Loan inputs can be changed and the schedule regenerated to simulate different scenarios or update with real payment data.
- **No External Services Required**: All calculations are done within the Google Apps Script environment, keeping your data secure within your spreadsheet.

## Google Sheets Integration

This script assumes a specific layout in your Google Sheet for each loan. Loan parameters are read from **row 4** and the amortization schedule is written starting from **row 8**.

### Loan Input Assignments (Row 4)
- **B4 – Loan Name**: A name or identifier for the loan.
- **C4 – Borrower Name**: Name of the borrower (optional).
- **D4 – Principal**: The initial loan principal amount.
- **E4 – Interest Rate**: Annual interest rate (numeric, e.g. `0.05` for 5%).
- **F4 – Closing Date**: The start date of the loan or disbursement date.
- **G4 – Term Months**: The loan term in months (e.g., `12` for one year).
- **H4 – Prorate**: `"Yes"` or `"No"`. If `"Yes"`, a prorated first period is created.
- **I4 – Payment Frequency**: Currently supports `"Monthly"` or `"Single Period"`.
- **J4 – Day Count**: `"Actual"` (uses actual days) or `"Periodic"` (assumes 30-day months).
- **K4 – Days Per Year**: Number of days in a year for interest calculations (e.g., `365` or `360`).
- **L4 – Prepaid Interest Date**: *(Optional)* Date up to which interest is prepaid.
- **M4 – Amortize**: `"Yes"` for fully amortizing loans, `"No"` for interest-only.
- **N4 – Origination Fee %**: *(Optional)* Origination fee as a percentage (e.g., `0.02` for 2%).
- **O4 – Exit Fee %**: *(Optional)* Exit fee as a percentage.
- **Q4 – Lock Inputs**: *(Optional)* A flag to lock input values (use with Google Sheets protection if desired).

## Loan Schedule Output

After running the script, the amortization schedule is populated starting from **row 8**. It is expected that row 7 (or earlier) contains the column headers. The schedule uses the following columns:

- **Period (Col B)**: The payment period number.
- **Period End (Col C)**: End date of the period.
- **Due Date (Col D)**: Payment due date.
- **Days (Col E)**: Number of days in the period.
- **Paid On (Col F)**: *(User input)* The actual payment date.
- **Total Due (Col G)**: Total amount due (interest, principal, fees).
- **Total Paid (Col H)**: *(User input)* The amount actually paid.
- **Principal Due (Col I)**: Scheduled principal due.
- **Principal Paid (Col J)**: Actual principal paid.
- **Interest Due (Col K)**: Interest due for the period.
- **Interest Paid (Col L)**: Actual interest paid.
- **Fees Due (Col M)**: Any fees due.
- **Fees Paid (Col N)**: Fees paid.
- **Interest Balance (Col O)**: Unpaid accrued interest.
- **Principal Balance (Col P)**: Remaining principal balance.
- **Total Balance (Col Q)**: Sum of principal and interest balances.
- **Notes (Col R)**: Additional remarks (e.g., indicating fees added).

## Installation & Usage

### Installing the Script
1. **Open Your Google Sheet**: Create or open a sheet where you want to manage loans.
2. **Access Apps Script**: Navigate to **Extensions > Apps Script**.
3. **Add the Script**: Create a new script file and copy-paste the contents of `LoanScript.js` from this repository.
4. *(Optional)* **Add Summary Functionality**: Create another script file (e.g., `SummaryPage.js`) for the optional summary features.
5. **Save the Project**.

### Setting Up Your Sheet
- **Loan Sheets**: Create a separate worksheet (tab) for each loan.
- **Input Row**: Set up **row 4** in each loan sheet with the input parameters as described above.
- *(Optional)* **Summary Sheet**: If using the summary feature, create a sheet named **Summary** where key metrics will be aggregated.

### Generating a Loan Schedule
After entering the inputs on a loan sheet:
1. **Via Apps Script Editor**:
   - Add a wrapper function:
     ```javascript
     function generateLoanSchedule() {
       // Generate schedule for the active sheet
       var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
       var generator = new LoanScheduleGenerator(sheet);
       generator.generateSchedule();
     }
     ```
   - Save the script and run `generateLoanSchedule()`. Grant the necessary permissions.
2. **Via Custom Menu (Optional)**:
   - Add an `onOpen` trigger to create a custom menu:
     ```javascript
     function onOpen() {
       SpreadsheetApp.getUi().createMenu('Loan Tools')
         .addItem('Generate Loan Schedule', 'generateLoanSchedule')
         .addItem('Recalculate Loan Balances', 'recalcLoanSchedule')
         .addToUi();
     }
     ```
   - Also add a recalculation function:
     ```javascript
     function recalcLoanSchedule() {
       var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
       var bal = new BalanceManager(sheet);
       bal.recalcAll();
     }
     ```
   - Reload your spreadsheet to see the **Loan Tools** menu.

 # Recording and Tracking Repayments

As time passes, you can update the schedule with actual payment data:

- **Paid On (Col F)**: Enter the actual payment date.
- **Total Paid (Col H)**: Enter the amount paid.

After updating, run the recalculation function (`recalcLoanSchedule()`) to update:
- **Interest Paid (Col L)**
- **Principal Paid (Col J)**
- **Principal Balance (Col P)**
- **Interest Balance (Col O)**

*Note:* You can add unscheduled payments by inserting a new row (leave the Period column blank) and entering the payment details. The script will recognize and apply these as prepayments.

# Using the Summary Sheet (Optional)

If you manage multiple loan sheets, a **Summary** sheet can provide an overview:

- **Populate Sheet Names**: Lists all loan sheet names (starting from cell B4).
- **Update Summary**: Aggregates data from each loan sheet into key metrics:
  - **Last Due Date (Col C)**
  - **Outstanding Principal (Col D)**
  - **Accumulated Interest (Col E)**
  - **Principal Paid (Col F)**
  - **Interest Paid (Col G)**
  - **Past Due Principal (Col H)**
  - **Past Due Interest (Col I)**

Use the **Summary Tools** menu (added via an onOpen trigger) to populate and update this summary.

# Customization

The Loan Script Library is designed to be flexible:

- **User Inputs and Schedule**: Modify the loan parameters in **row 4** and regenerate the schedule.
- **Payment Amounts & Dates**: Edit the **Paid On** and **Total Paid** columns to simulate different payment scenarios.
- **Adding/Removing Periods**: Adjust the **Term Months** or insert unscheduled payment rows to change the schedule.
- **Column Configuration**: Update the `SHEET_CONFIG` in `LoanScript.js` if you need a different layout.
- **Formatting**: Change number and date formats in the script’s formatting function.
- **Interest Calculation Methods**: Extend the logic for alternative frequencies or compounding if needed.
- **Locking Inputs**: Use the **Lock Inputs** flag (Q4) in conjunction with Google Sheets protection features.

Always test any customizations on a copy of your data to ensure the script works as expected.

# Example Usage

**Scenario:**
- **Loan Amount**: $10,000
- **Annual Interest Rate**: 5% (`0.05`)
- **Term**: 12 months
- **Closing Date**: January 15, 2025
- **Payment Frequency**: Monthly
- **Prorate First Period**: Yes
- **Day Count**: Actual (with 365 days/year)
- **Amortize**: No (interest-only)
- **Origination Fee**: 2% (added to principal)
- **Exit Fee**: 1% (due in final period)

**Process:**
- **Origination Fee**: 2% of $10,000 is added, increasing the principal to $10,200.
- **Prorated Period**: A short first period (Period 0) is generated (Jan 15–Jan 31, 2025).
- **Monthly Calculations**: Interest is computed for each period. In the final period, the principal plus exit fee is due.
- **Payment Tracking**: Record payments as they occur. If a payment is missed or partial, update the **Paid On** and **Total Paid** fields, then run recalculation to adjust the remaining balances.

# Contributing

This project is **not open** for external contributions. It is provided as a library for personal or internal use. You are welcome to fork the repository and modify the code for your own needs. Bug reports or suggestions via GitHub issues are accepted, but there is no formal process for contributions.

# License

This project is licensed under the **MIT License**. You are free to use, modify, and distribute the code as permitted by the license. The software is provided "as is" without warranty of any kind, and the author is not liable for any claims or damages arising from its use.

---

By using the Loan Script Library in your Google Sheets, you can streamline loan tracking and amortization calculations. Enjoy efficient loan scheduling!  
