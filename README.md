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
3. **Add the Script**: Create a new script file and copy-paste the contents of `LoanScriptWrapper.js` from this repository.
4. **Add the Library**: In Apps Script click on the + sign next to Libraries. Then paste this script ID into the Look Up box "1lfXARnO_beVQDAfErbGQUajgL7vuYWM9BGB9JmtR-JOrlPIlHSa4OgTP". Choose the most recent version and push "Add".
5. **Save the Project**.

### Setting Up Your Sheet
- **Loan Sheets**: Create a separate worksheet (tab) for each loan.
- **Input Row**: Set up **row 4** in each loan sheet with the input parameters as described above.
- *(Optional)* **Summary Sheet**: If using the summary feature, create a sheet named **Summary** where key metrics will be aggregated.

### Generating a Loan Schedule

After entering the inputs on a loan sheet (see Installation & Usage above), the script will generate an amortization schedule based on the type of loan specified by your inputs. The **Payment Frequency** (`I4`) and **Amortize** (`M4`) inputs determine which of the four main schedule structures is produced:

- **Single Period Loan** – one period covering the entire loan term, with a single payment at maturity.
- **Monthly Periodic Loan** – monthly periods with interest-only payments (principal due at end), using a 30-day month convention (30/360 day count).
- **Monthly Actual Loan** – monthly periods with interest-only payments (principal due at end), using actual day counts for each period (actual/365).
- **Monthly Amortized Loan** – monthly periods with fully amortizing payments (principal and interest in each payment), using either 30-day or actual day count as specified.

Each structure is generated by the script according to the inputs. Below is a detailed breakdown of each loan type, the formulas used, and an example of the schedule output that the script would generate for a sample loan.

#### Single Period Loans (One-Time Payment at Maturity)

In a **Single Period Loan**, the entire loan is one period long – there are no interim payments. The borrower pays all interest and principal in one lump sum at the end of the term. To generate this schedule, set **Payment Frequency** to `"Single Period"` (the script will automatically treat it as interest-only and use actual day count for accuracy). The schedule will have a single row representing the loan’s maturity. The **Due Date** is the end of the term, and the **Days** in the period is the total days from loan start to that date. The **Interest Due** is calculated on the full principal over that period, and **Principal Due** is the full loan principal (plus any fees) due at maturity. No payments are scheduled before the due date.

**Formulas:**  
For a single period loan, interest accrues on the principal for the length of the term. If *P* is the principal, *r* is the annual interest rate, and *d* is the number of days in the term, the interest due at maturity is calculated as: `Interest = P × r × (d / DaysPerYear)`

The script uses the actual day count for *d* (actual days between disbursement and maturity) and your specified *DaysPerYear* (often 365) for accuracy. The total payment at maturity will thus be the sum of principal and interest (plus any fees). For example, a \$1,000 loan at 10% annual interest for 6 months would accrue roughly half a year’s interest.

![Single Period Loan Example](path/to/single_period_example.png)  
*Example:* **Single Period Loan** – A \$1,000 loan at 10% annual interest, disbursed Jan 1, 2023, and due July 1, 2023. The schedule shows one period spanning 181 days, with \$49.59 interest due and \$1,000 principal due at the end (total payment \$1,049.59). No intermediate periods or payments exist, as everything is paid in one lump sum.

#### Monthly Loans with Periodic Day Count (30/360 Interest-Only)

A **Monthly Periodic Loan** uses monthly periods and assumes a 30-day month/360-day year convention for interest calculations. This structure is typically **interest-only**, meaning the borrower pays interest each month and the entire principal at the end. To generate this, set **Payment Frequency** to `"Monthly"`, **Day Count** to `"Periodic"` (30-day months), and **Amortize** to `"No"` (interest-only). The script will create a row for each monthly period up to the term. Each period’s **Interest Due** is computed as if the month had 30 days, so the monthly interest is constant each period. The **Principal Due** for all periods except the last is zero (since principal is not scheduled to be paid during the term), and in the final period the Principal Due equals the full loan principal (plus any exit fee). The **Payment Due** each month is equal to that period’s interest, and the final payment includes the last interest amount plus the principal.

**Formulas:**  
With a 30/360 convention, the monthly interest factor is typically *r/12* (using a 360-day year). For principal *P* and annual rate *r*, each month’s interest due is roughly: `Interest = P × (r / 12)`

For example, if *P* = \$1,000 and *r* = 10%, each month’s interest is about \$8.33. Over 6 months, the borrower would pay interest-only of ~\$8.33 per month, and in month 6 also pay the \$1,000 principal. (If an origination fee was financed into principal or an exit fee is due, those would appear as added principal in the first or last period, respectively.)

![Monthly Periodic Loan Example](path/to/monthly_periodic_example.png)  
*Example:* **Monthly Periodic Loan (Interest-Only)** – \$1,000 principal, 10% annual interest, 6 monthly periods (30/360 basis). The schedule shows 6 monthly periods of 30 days each. Interest due is \$8.33 every month (constant because of the 30-day month assumption). **Payment Due** is \$8.33 each month, covering interest-only, and the **Principal Due** is \$0.00 until the final period. In period 6, the interest due is \$8.33 and Principal due is \$1,000, so the final **Payment Due** is \$1,008.33 (principal + interest). This reflects an interest-only schedule with principal repaid at the end.

#### Monthly Loans with Actual Day Count (Actual/365 Interest-Only)

A **Monthly Actual Loan** is similar to the above but uses actual calendar days for interest calculations instead of a fixed 30-day assumption. This means monthly interest amounts can vary slightly depending on the number of days in each period. To generate this, set **Payment Frequency** to `"Monthly"`, **Day Count** to `"Actual"`, and **Amortize** to `"No"`. The script will create a row for each monthly period, and compute each period’s **Interest Due** based on the exact days between payments. Typically, February will have fewer days of interest than January, so interest due is not identical each month. Like the periodic case, principal is not due until the final period (interest-only structure), so each period’s **Payment Due** is the interest for that month, and the last payment includes all principal. Using actual days and a specified DaysPerYear (e.g., 365) yields a slightly more precise accrual of interest day-by-day.

**Formulas:**  
For each period, interest is calculated as: `Interest = P × r × (d / DaysPerYear)` where *d* is the actual number of days in that period. The value of *d* might be 31 for long months, 28 for February, etc. Thus, months with more days incur a bit more interest. The principal remains unchanged during the term (since we’re not amortizing or paying it monthly), and is paid in full at the end. The overall interest paid may differ slightly from the periodic method if the term doesn’t divide evenly into 30-day blocks.

![Monthly Actual Loan Example](path/to/monthly_actual_example.png)  
*Example:* **Monthly Actual Day Count Loan (Interest-Only)** – \$1,000 at 10% annual, 6 monthly periods using actual days (assuming Jan–Jun 2023). The schedule shows varying **Interest Due** each month based on days: January (31 days) approximately \$8.49, February (28 days) approximately \$7.67, etc. The **Payment Due** each period equals that interest amount. **Principal Due** is \$0.00 for periods 1–5 and \$1,000 in period 6. The final payment in period 6 is about \$1,008.22 (which includes \$1,000 principal + \$8.22 interest for June’s 30 days). This illustrates how actual day counts cause slight month-to-month differences in interest, while principal is paid at the end.

#### Fully Amortizing Monthly Loans (Equal Payments of Principal & Interest)

For a **Monthly Amortized Loan**, set **Payment Frequency** to `"Monthly"` and **Amortize** to `"Yes"` (the Day Count can be Actual or Periodic, depending on how you want interest accrued). In a fully amortizing loan, each scheduled payment includes both interest and principal so that the loan is paid off by the end of the term. The script calculates an equal **Payment Due** amount for each period that covers the accruing interest and also pays down principal over time. This is typically done using the standard amortization formula for the payment amount. Internally, the script determines the periodic interest rate (using either 30-day or actual day count accrual for each period) and then finds the payment amount *A* such that the present value of all payments equals the loan principal (ensuring full payoff). All period rows are then filled with their respective **Interest Due** and **Principal Due** such that interest is paid and principal reduces to zero at term end. If an actual day count is used, the interest portion is adjusted for each period’s actual days, but the total payment can be kept equal by slight adjustments.

**Formulas:**  
The fixed payment for an amortizing loan is calculated using the annuity formula:

\[
A = P \times \frac{i(1+i)^n}{(1+i)^n - 1}
\]

where *P* is the principal, *i* is the periodic interest rate (for example, monthly rate = annual rate/12 for a 30/360 loan), and *n* is the total number of periods. For each period, the interest due is computed as: `Interest = Current Balance × i` and the principal due is the difference: `Principal = Payment - Interest`

In an amortizing schedule, **Principal Due** starts smaller and increases over time as the interest portion decreases (with a fixed total payment). By the final period, the last payment’s interest component is very small and the principal component fully extinguishes the remaining balance. The script ensures that with **Amortize** set to "Yes", the **Total Balance** (principal remaining) reaches zero at the end of the term, and it automatically recalculates the schedule if any extra payments are made early so that the remaining payments still amortize the loan fully.

![Fully Amortizing Loan Example](path/to/amortizing_loan_example.png)  
*Example:* **Fully Amortizing Loan** – \$1,000 at 10% annual, 6 monthly payments (Amortize = Yes, 30/360 basis for simplicity). The schedule shows a fixed **Payment Due** of \$171.56 each month. Initially, interest is approximately \$8.33 (for period 1) and principal due is \$163.23 (the remainder of the \$171.56 payment). Each period, the interest due decreases as the principal balance drops, and the principal portion increases. By period 6, the interest due is only about \$1.42 and principal due is \$170.14, which pays off the remaining balance. The loan is fully paid by the end, with the total of payments equaling the principal plus interest.

Each of these examples corresponds to the schedule output you would see on the sheet after running the **Generate Loan Schedule** function. The script reads your inputs and constructs the appropriate schedule with the correct formulas applied for interest and principal in each period. By understanding which loan type your inputs produce, you can interpret the schedule (and even make manual adjustments or additional payments) with confidence in how the calculations are handled by the Loan Script Library.

### Loan Script Wrapper

#### Overview

The **Loan Script Wrapper** is a helper script that provides easy-to-use functions and menu items to execute the core Loan Script Library features from within your Google Sheet. Instead of calling the library classes directly each time, the wrapper exposes simple functions (e.g. `generateLoanSchedule()` and `recalcLoanSchedule()`) that internally call the appropriate library objects and invoke their methods on the active sheet. The wrapper also handles Google Sheets UI integration – it defines an `onOpen()` trigger that adds a custom menu so you can run these functions with a click. In short, the wrapper acts as a bridge between the spreadsheet and the Loan Script Library, making it straightforward to generate schedules and refresh loan data without editing code each time.

#### Using the Wrapper

To use the wrapper properly, follow these guidelines:

#### 1. Include the Wrapper Script

- **Add the Wrapper to Your Project:**  
  Add the wrapper to your Apps Script project along with adding the library code under libraries. You can either manually create the wrapper functions as shown above or simply copy the provided `LoanScriptWrapper.js` file into your project. This file already contains the implementations for `generateLoanSchedule`, `recalcLoanSchedule`, `onOpen`, and other helper functions so you don’t have to write them yourself.
  
- **Save Your Script Project:**  
  Once added, save your script project.

#### 2. Open/Refresh the Spreadsheet

- **Trigger the onOpen Function:**  
  After adding the wrapper, reload your Google Sheet (or reopen it) to trigger the `onOpen` function. The wrapper’s `onOpen` will automatically create a **Loan Tools** menu (labeled “Loan Schedule Tools” in the latest version) in your spreadsheet’s menu bar.

- **Access the Custom Menu:**  
  This custom menu includes options such as:
  - Generate Loan Schedule
  - Recalculate Schedule
  - Insert Unscheduled Payment Row
  - Set Up Triggers
  
  All options are linked to the corresponding wrapper functions. You should see this menu appear after the sheet is opened, allowing you to run the loan script features without going back to the script editor.

#### 3. Generate a Loan Schedule

- **Navigate to Your Loan Sheet:**  
  Make sure your loan sheet has inputs set up (typically in row 4).

- **Run the Schedule Generator:**  
  You can generate the loan schedule by either:
  - Clicking the “Generate Loan Schedule” item from the custom menu, or
  - Running the `generateLoanSchedule()` function from the Apps Script editor.

  The wrapper will take the active sheet and call the library’s schedule generator to populate the loan’s payment schedule on that sheet.  
  _(On first run, you may be prompted to authorize the script to run – grant the necessary permissions.)_  
  After this step, the sheet will be filled with the calculated schedule of periods, payments, interest, balances, etc.

#### 4. Record Payments and Recalculate

- **Update the Schedule After Recording Payments:**  
  As you record actual repayments over time, use the wrapper’s recalculation function to update the schedule. Enter the payment details (e.g. Paid On date and amounts paid) in the schedule and the `onEdit()` function should automatically update the loan schedule. If it does not automatically update, you can either select “Recalculate Schedule” from the Loan Tools menu or run `recalcLoanSchedule()` in the script.

- **Recompute Balances:**  
  The wrapper will call the library’s balance manager to recompute interest accruals and remaining balances based on the payments entered. This ensures your sheet reflects all payments made and adjusts the interest and principal balances accordingly.  

#### 5. Optional – Set Up Automated Triggers

- **Automate Functions:**  
  For convenience, the wrapper provides a “Set Up Triggers” function accessible via the custom menu. Running this will configure Google Apps Script triggers to automate the running of certain functions at specified times or events.

- **Examples of Automated Tasks:**  
  For example, the wrapper can install an onEdit trigger to run `onEdit()` (or other maintenance tasks) on your loan sheets. This is useful if you want your loan schedule to update automatically (e.g., updating interest calculations daily or refreshing the summary sheet regularly) without manual intervention.

- **Trigger Management:**  
  Once you select **Set Up Triggers**, the script will create the necessary trigger(s) behind the scenes. If needed, you can later modify or remove these triggers via **Extensions > Apps Script > Triggers** in your Google Sheet’s script project.

### 6. Additional Feature: Insert Unscheduled Payment Row

- **Log Extra Payments Easily:**  
  The wrapper’s menu includes an “Insert Unscheduled Payment Row” option. This function inserts a new row in the schedule (with the proper formatting and formulas) for an unscheduled payment. You can then fill in the **Paid On** date and amounts.
  
- **Post-Insertion Action:**  
  After inserting an unscheduled payment row, remember to run the recalculation to integrate that payment into the loan balances. Using this menu item ensures the extra payment is added in the correct format, saving you from manually adjusting formulas or references when handling prepayments.

---

By using the Loan Script Wrapper as described above, users can seamlessly interact with the Loan Script Library’s capabilities. The wrapper abstracts the complex function calls into one-click menu actions or simple function calls, making the loan management process in Google Sheets much more user-friendly.

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

 ## Recording and Tracking Repayments

As time passes, record actual payments in your schedule by updating key fields:

- **Paid On (Col F)**  
  Enter the actual payment date. For scheduled payment rows (those with a period number), this records when the due payment was made (leave it blank if a payment hasn’t occurred yet). For unscheduled payments (new rows with a blank Period), this date indicates when the extra payment occurred.

- **Principal Paid (Col J) / Interest Paid (Col L) / Fees Paid (Col N)**  
  Enter the amounts of the payment applied to principal, interest, and any fees, respectively.  
- **Note:** Do not enter a combined “Total Paid” amount. The script will automatically calculate Total Paid (Col H) as the sum of these values.

After inputting the payment details, run the recalculation function (`recalcLoanSchedule()`). This will refresh the schedule to reflect the payments:

- **Payment Due (Col G)**  
  Recalculated for each period based on the accrued interest and any fees for that period.

- **Total Paid (Col H)**  
  Automatically updated to equal Principal + Interest + Fees Paid for that row.

- **Principal Balance (Col P)**  
  Adjusted to the remaining principal after the payment.

- **Interest Balance (Col O)**  
  Updated to any interest that has accrued but remains unpaid (carried forward if the payment didn’t cover all interest due).

    **Note:** To record an unscheduled payment (prepayment), insert a new row in the schedule (leave the Period column blank) and fill in the Paid On date along with the Principal/Interest/Fees Paid for that entry. The script will recognize these as prepayments and incorporate them (in chronological order by Paid On date) when recalculating balances.
---

## Handling Unscheduled Payments in Different Loan Scenarios

Unscheduled payment rows (inserted with the **Period** column blank) let you record extra payments outside the regular schedule. The script will recognize these entries and apply them as prepayments or out-of-sequence payments. Below we explain how unscheduled **Interest Paid**, **Principal Paid**, and **Fees Paid** affect the loan’s balances and future payments in various scenarios, and the role of the **Paid On** date in each case.


### 1. Single Period Loans (One-Time Payment at Maturity)

- **Interest Paid:**  
  If you enter an interest amount in an unscheduled row for a single-period loan, it immediately reduces the accrued interest balance. The code subtracts any Interest Paid from the running interest due. This means you’re paying off interest before the maturity date, so less (or no) interest remains to be paid at final maturity. In effect, interest paid early stops interest from accumulating further on that portion.

- **Principal Paid:**  
  An unscheduled principal payment lowers the outstanding principal immediately​. The script will apply that payment to reduce the loan balance, so the remaining principal due at maturity is smaller. After this prepayment, interest will only accrue on the reduced principal going forward, which directly lowers the final payment due (both principal and interest). In short, any principal paid early is a direct reduction of what you owe later.

- **Fees Paid:**  
  Entering a fee amount in an unscheduled row will cut down the outstanding fees balance right away​. For example, if the loan had an exit fee due at the end, paying it (or part of it) early means that much less fee will be due at maturity. This doesn’t affect interest or principal calculations, but it does reduce the total amount you’ll have to pay in the end.

- **Paid On (Date):**  
  Adding a Paid On date to the unscheduled row is crucial in a single-period loan. This tells the script exactly when the extra payment happened. The engine will accrue interest from the loan start (or last payment) up to that date, then apply your payment on that day​. In practice, it splits the loan’s one big period into two segments: before and after the unscheduled payment. The interest for the first segment (up to the Paid On date) is calculated and can be settled by the unscheduled payment, and the remaining time until maturity will accrue interest on the now-lower balance. Without a Paid On date, the payment wouldn’t be anchored in time – the script wouldn’t know when to apply it – so it would effectively be ignored in the interest calculations until the end.

---

### 2. Monthly Loans with Periodic Day Count (30/360) – Non-Amortizing/Interest-Only

- **Interest Paid:**  
  In a monthly loan using a 30/360 convention, interest is typically calculated on a fixed 30-day period. If you record an Interest Paid in an unscheduled row, the script subtracts that amount from any interest balance immediately. This is usually only relevant if interest from a previous period was unpaid – the unscheduled payment would then clear that owed interest. Paying interest outside the normal schedule doesn’t change future scheduled interest charges (those will still accrue on the remaining principal each period), but it does ensure no past interest is carrying forward. Essentially, you’re catching up on interest so that the interest balance is brought to $0$, preventing accumulation of unpaid interest.

- **Principal Paid:**  
  An unscheduled Principal Paid on a monthly periodic loan will reduce the principal balance as of the Paid On date. The script handles this by prorating the interest for that partial period based on the payment date. It calculates the fraction of the month that has passed up to the unscheduled payment date​ and accrues interest for that portion of the period. Then it applies the principal payment, dropping the running principal balance immediately. For the remainder of that month (and all following months), interest will accrue on the lower principal. In practical terms, future interest due each period will be less than originally scheduled because the principal has been curtailed mid-period. In an interest-only setup (non-amortizing), the scheduled payment each month is just the interest – since the principal is now smaller, the interest due in subsequent months decreases. The principal prepaid will also directly reduce the final principal due (if the loan requires a lump-sum principal payoff at the end).

- **Fees Paid:**  
  If the loan involves fees (for example, an origination fee financed into balance or an exit fee due later), an unscheduled Fees Paid will cut down the outstanding fee balance immediately. This doesn’t affect interest accrual or the core amortization, but it means that any fee due in the future is partially/fully paid now. For instance, if a $500 fee is due at loan end, and you pay $200 in an unscheduled row, the remaining fee due later will be $300. The fee payment is applied as soon as the Paid On date indicates, reducing the Fees Balance tracked by the script.

- **Paid On (Date):**  
  In monthly 30/360 loans, the Paid On date in an unscheduled row tells the model exactly when the extra payment occurred during the month. The script will compute interest from the last period’s end date up to this date as a portion of the 30-day period. It then applies the payment on that date and adjusts the balances (principal/interest/fees) accordingly. After the payment, interest for the rest of the period is accrued on the new, lower balance. This means the period’s interest due is correctly “split” into two parts: pre- and post-payment. If you do not provide a date, the script cannot place the payment in the timeline, so it would assume the payment happens after all scheduled periods (making it ineffective for reducing interest in the interim). Always include the actual payment date for unscheduled entries to ensure the calculation reflects the timing of that prepayment.

---

### 3. Monthly Loans with Actual Day Count (365/Actual days) – Non-Amortizing/Interest-Only

- **Interest Paid:**  
  For a monthly loan that uses actual day counts, interest accrues day by day. When you add an unscheduled Interest Paid entry, the script immediately deducts that from any accumulated interest balance. If there was unpaid interest from a prior period (say you missed or underpaid a scheduled interest payment), this extra payment will reduce or clear that outstanding interest. Paying interest early (before the normal due date) in an actual day-count loan isn’t common since interest is typically paid as it accrues each period, but the option exists to remove any interest that’s hanging out in the balance. Once paid, that interest stops accruing (the script’s running interest balance is set to zero or lower) so you won’t be charged interest on it later (note: the script doesn’t compound interest on overdue interest – it keeps it separate as an interest balance).

- **Principal Paid:**  
  An unscheduled Principal Paid in a monthly actual loan directly reduces the principal on the date of payment. The script calculates the exact number of days between the last payment date (or loan start) and the unscheduled payment’s date to find how much interest accrued on the old principal in that interval. It adds that interest to the interest balance, then applies your principal payment, cutting down the outstanding principal immediately. Going forward from that date, the principal is smaller, so the interest that will accrue each day (and thus each future period’s interest due) is reduced. In an interest-only structure, this means your upcoming scheduled interest payments will be lower. And if a final principal payoff is expected at maturity, that amount will be reduced by the unscheduled principal you paid early.

- **Fees Paid:**  
  Unscheduled Fees Paid in an actual day-count loan work the same way as in the periodic case – the fee balance is decreased as soon as that payment is applied. If the loan had a fee due later, paying some or all of it now (on a given date) will remove that portion from the outstanding fees. There’s no effect on interest or principal calculations aside from reducing the total obligations. It simply means when the fee would have been due, you owe less (because you’ve already paid part of it).

- **Paid On (Date):**  
  The Paid On date in an unscheduled entry is equally important for actual day-count loans. It pins the extra payment to a specific day, allowing the script to correctly compute accrued interest to that point. When a date is provided, the code will accrue interest from the last period up to that exact day (using the actual number of days) and add it to running interest. Then the payment is applied on that date, which immediately adjusts the balances (reducing principal/interest/fees as specified). From the next day onward, interest accrues on the new principal. In effect, the schedule acknowledges the early payment: the current period’s interest due is reduced and the remaining balance is lower. Without a Paid On date, the extra payment wouldn’t be inserted into the timeline – the script wouldn’t know when to apply it – so it would not affect the interest calculation for any period (it might only show up as an additional payment at the end, not saving any interest in the interim). Always use the actual payment date for unscheduled payments so the calculation reflects the timing correctly.

---

### 4. Fully Amortizing Monthly Loans (Amortize = "Yes") – Actual or Periodic Day Count

- **Interest Paid:**  
  In a fully amortizing loan, each scheduled payment is supposed to cover that period’s interest in full (plus some principal). Therefore, an unscheduled Interest Paid is usually only needed if there was interest that went unpaid in a prior period. If you do input an Interest Paid in an unscheduled row, the script will immediately apply it to reduce the interest balance. This could happen if, for example, a scheduled payment was missed or short-paid and interest accrued into the next period – an extra payment can then be made to pay off that lingering interest. Once applied, the running interest is decreased, ensuring that no old interest remains to hinder the amortization. In normal cases (payments made in full), there wouldn’t be an interest balance to pay outside the schedule, so this field is less commonly used for amortizing scenarios except to correct a deficit.

- **Principal Paid:**  
  An unscheduled Principal Paid has a significant effect on an amortizing loan. This is essentially a prepayment of principal on top of the regular installment. The script will apply the principal reduction immediately on the Paid On date, lowering the outstanding balance mid-schedule. It also recalculates the interest up to that date so that the borrower is charged interest only for the time the original principal was still in the loan. After that date, the remaining principal is smaller, so the interest for the rest of the period (and future periods) will be computed on that lower balance. By default, the amortization plan assumes a fixed payment amount each period, but a prepayment means that plan is no longer optimal. The library handles this by re-amortizing the future schedule to account for the extra principal payment. In other words, it will recalculate the remaining payment amounts or allocations so that the loan still fully pays off by the end of the term with the new reduced balance. All subsequent scheduled rows are updated with new Principal Due and Interest Due values reflecting this recalculation. The net effect is that your future monthly payments could decrease or your loan will finish earlier (depending on how the schedule is structured), because you've paid extra principal. The script ensures the adjustment is made such that no negative or “extra” payments occur at the end – the loan is simply paid off sooner or with less due each period. (Internally, a flag is set so that once the schedule is re-amortized, it uses the new values instead of the original amortization plan.)

- **Fees Paid:**  
  For amortizing loans, fees (like origination or exit fees) might be included as part of the balance or as separate due amounts in certain periods. An unscheduled Fees Paid will immediately reduce any outstanding fees just as in other scenarios. If the fee was scheduled at a future date (e.g., an exit fee at maturity), paying some of it early will decrease the fee balance and thus lower the amount due when that fee’s period comes. This doesn’t directly alter the amortization of principal and interest – it primarily affects the Fees Due column and the total remaining balance. However, since the total balance (principal + interest + fees) is tracked, an early fee payment will reduce the overall balance figure. The schedule will show the fee as partially paid in the unscheduled row and the fee due at the final period would effectively be smaller (even though the scheduled fee due in that final period might remain the same on paper, the fee balance carried into that period will be lower due to your prepayment).

- **Paid On (Date):**  
  In a fully amortizing loan, the Paid On date for an unscheduled payment is what allows the script to insert the prepayment into the amortization schedule timeline. When you provide a date, the calculation will include interest accrual up to that day and apply the payment there, just like in the other scenarios. The difference in an amortizing loan is what happens next: the script will recompute the remaining amortization from that point forward. Practically, the steps are: interest is accrued from the last payment date to the unscheduled payment date (so you pay interest only for those days on the old balance), then the extra payment reduces the principal, and then the remaining future payments are adjusted based on the new principal. The Paid On date ensures this all happens at the correct time. If you didn’t specify a date, the extra payment couldn’t be placed in the schedule properly – the script would treat it as if it came after the last period, which means it wouldn’t shorten the loan or reduce any intermediate interest. So, for an amortizing loan, the Paid On date is essential to reap the benefit of paying off principal early: with it, the schedule re-calculates and you see lower balances and updated payment amounts going forward. Always include the actual payment date for unscheduled principal prepayments; this way, the schedule will show the effect of that payment (less interest accrued after that date and a new amortization schedule for the remaining term).

- **Future Payments Due:**  
  After an unscheduled payment in an amortizing loan, the future payment schedule is adjusted. Initially, amortizing loans have a fixed periodic payment. Once a prepayment happens, the script’s re-amortization will typically result in a new (lower) payment amount for the remaining periods or the same payment amount but fewer periods (the library by default recalculates the remaining payments to fully amortize by the original end date, often yielding a lower monthly payment). For example, if you had 24 months left and you prepaid a chunk of principal, the code will recalculate what the new equal payment should be for the next 24 months to pay off the now-smaller balance. Those recalculated amounts are written into the schedule for future periods. This means that the “Payment Due” column (total due each period) may change after the unscheduled payment. In contrast, in interest-only scenarios, there is no fixed installment to recalc – instead the interest due each period simply drops in proportion to the lower balance. But in amortizing scenarios, the entire payment structure is updated so that the loan remains on track to finish at the same time, just with less money paid in interest. The key takeaway is: an extra principal payment in an amortizing loan will reduce what you owe in the future, and the script reflects that by lowering either the upcoming payment amounts or the number of payments needed.

---

### Summary of Unscheduled Payment Effects

- **Interest Paid (unscheduled):**  
  Always reduces the accumulated interest balance immediately. This prevents interest from remaining unpaid. It’s most impactful if interest was accrued from missed payments or in between scheduled dates; once paid, that interest is no longer owed and won’t appear in future due amounts. It does not prepay future interest – it only pays off interest that has accrued up to the payment date (or was scheduled up to that point).

- **Principal Paid (unscheduled):**  
  Always lowers the outstanding principal right away. Future interest calculations use the new lower principal, so interest due each period will drop going forward. In interest-only loans, this simply reduces the interest portion of future payments (since principal isn’t due until the end, but now the end balance is smaller). In amortizing loans, this triggers a re-amortization of remaining payments, typically reducing the periodic payment amount or the loan term. Either way, you save on interest overall by paying down principal early.

- **Fees Paid (unscheduled):**  
  Immediately reduces any outstanding or future fees. This lowers the total balance of the loan but doesn’t affect interest accrual on principal (fees are handled separately). It means when a fee would have been due, that required amount is smaller (because you paid part of it in advance). Always record the date for fee payments as well to place them correctly (especially if interest might accrue on overdue fees in a custom scenario, though by default the script treats fees separately from interest).

- **Paid On date:**  
  The Paid On date is critical for every unscheduled entry. It ensures the payment is integrated into the timeline of the loan. The script processes unscheduled payments in chronological order alongside scheduled periods. By providing the date, you let the script accrue interest correctly up to that point and then apply the payment. The result is an accurate reflection of the loan’s state after the payment: interest is calculated only for the time it was actually outstanding, and principal/fee balances are updated when they should be. If you omit the Paid On date, the script will not know when to apply the payment – effectively the payment would either be ignored in the interim calculations or treated as happening at the end, yielding no benefit in reducing interest or balances until the loan’s end. Always include a Paid On date for unscheduled payments to get the intended outcome in the amortization schedule.

## Using the Summary Sheet (Optional)

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

## Customization

The Loan Script Library is designed to be flexible:

- **User Inputs and Schedule**: Modify the loan parameters in **row 4** and regenerate the schedule.
- **Payment Amounts & Dates**: Edit the **Paid On** and **Total Paid** columns to simulate different payment scenarios.
- **Adding/Removing Periods**: Adjust the **Term Months** or insert unscheduled payment rows to change the schedule.
- **Column Configuration**: Update the `SHEET_CONFIG` in `LoanScript.js` if you need a different layout.
- **Formatting**: Change number and date formats in the script’s formatting function.
- **Interest Calculation Methods**: Extend the logic for alternative frequencies or compounding if needed.
- **Locking Inputs**: Use the **Lock Inputs** flag (Q4) in conjunction with Google Sheets protection features.

Always test any customizations on a copy of your data to ensure the script works as expected.

## Example Usage

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

## Contributing

This project is **not open** for external contributions. It is provided as a library for personal or internal use. You are welcome to fork the repository and modify the code for your own needs. Bug reports or suggestions via GitHub issues are accepted, but there is no formal process for contributions.

## License

This project is licensed under the **MIT License**. You are free to use, modify, and distribute the code as permitted by the license. The software is provided "as is" without warranty of any kind, and the author is not liable for any claims or damages arising from its use.

---

By using the Loan Script Library in your Google Sheets, you can streamline loan tracking and amortization calculations. Enjoy efficient loan scheduling!  
