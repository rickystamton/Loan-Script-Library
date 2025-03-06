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

# Handling Unscheduled Payments in Different Loan Scenarios

Unscheduled payment rows (inserted with the **Period** column blank) let you record extra payments outside the regular schedule. The script will recognize these entries and apply them as prepayments or out-of-sequence payments. Below we explain how unscheduled **Interest Paid**, **Principal Paid**, and **Fees Paid** affect the loan’s balances and future payments in various scenarios, and the role of the **Paid On** date in each case.

---

## 1. Single Period Loans (One-Time Payment at Maturity)

- **Interest Paid:**  
  If you enter an interest amount in an unscheduled row for a single-period loan, it immediately reduces the accrued interest balance. The code subtracts any Interest Paid from the running interest due​ GITHUB.COM. This means you’re paying off interest before the maturity date, so less (or no) interest remains to be paid at final maturity. In effect, interest paid early stops interest from accumulating further on that portion.

- **Principal Paid:**  
  An unscheduled principal payment lowers the outstanding principal immediately​ GITHUB.COM. The script will apply that payment to reduce the loan balance, so the remaining principal due at maturity is smaller. After this prepayment, interest will only accrue on the reduced principal going forward, which directly lowers the final payment due (both principal and interest). In short, any principal paid early is a direct reduction of what you owe later.

- **Fees Paid:**  
  Entering a fee amount in an unscheduled row will cut down the outstanding fees balance right away​ GITHUB.COM. For example, if the loan had an exit fee due at the end, paying it (or part of it) early means that much less fee will be due at maturity. This doesn’t affect interest or principal calculations, but it does reduce the total amount you’ll have to pay in the end.

- **Paid On (Date):**  
  Adding a Paid On date to the unscheduled row is crucial in a single-period loan. This tells the script exactly when the extra payment happened. The engine will accrue interest from the loan start (or last payment) up to that date, then apply your payment on that day​ GITHUB.COM. In practice, it splits the loan’s one big period into two segments: before and after the unscheduled payment. The interest for the first segment (up to the Paid On date) is calculated and can be settled by the unscheduled payment, and the remaining time until maturity will accrue interest on the now-lower balance. Without a Paid On date, the payment wouldn’t be anchored in time – the script wouldn’t know when to apply it – so it would effectively be ignored in the interest calculations until the end.

---

## 2. Monthly Loans with Periodic Day Count (30/360) – Non-Amortizing/Interest-Only

- **Interest Paid:**  
  In a monthly loan using a 30/360 convention, interest is typically calculated on a fixed 30-day period. If you record an Interest Paid in an unscheduled row, the script subtracts that amount from any interest balance immediately​ GITHUB.COM. This is usually only relevant if interest from a previous period was unpaid – the unscheduled payment would then clear that owed interest. Paying interest outside the normal schedule doesn’t change future scheduled interest charges (those will still accrue on the remaining principal each period), but it does ensure no past interest is carrying forward. Essentially, you’re catching up on interest so that the interest balance is brought to $0$, preventing accumulation of unpaid interest.

- **Principal Paid:**  
  An unscheduled Principal Paid on a monthly periodic loan will reduce the principal balance as of the Paid On date​ GITHUB.COM. The script handles this by prorating the interest for that partial period based on the payment date. It calculates the fraction of the month that has passed up to the unscheduled payment date​ GITHUB.COM and accrues interest for that portion of the period. Then it applies the principal payment, dropping the running principal balance immediately. For the remainder of that month (and all following months), interest will accrue on the lower principal. In practical terms, future interest due each period will be less than originally scheduled because the principal has been curtailed mid-period. In an interest-only setup (non-amortizing), the scheduled payment each month is just the interest – since the principal is now smaller, the interest due in subsequent months decreases. The principal prepaid will also directly reduce the final principal due (if the loan requires a lump-sum principal payoff at the end).

- **Fees Paid:**  
  If the loan involves fees (for example, an origination fee financed into balance or an exit fee due later), an unscheduled Fees Paid will cut down the outstanding fee balance immediately​ GITHUB.COM. This doesn’t affect interest accrual or the core amortization, but it means that any fee due in the future is partially/fully paid now. For instance, if a $500 fee is due at loan end, and you pay $200 in an unscheduled row, the remaining fee due later will be $300. The fee payment is applied as soon as the Paid On date indicates, reducing the Fees Balance tracked by the script.

- **Paid On (Date):**  
  In monthly 30/360 loans, the Paid On date in an unscheduled row tells the model exactly when the extra payment occurred during the month. The script will compute interest from the last period’s end date up to this date as a portion of the 30-day period​ GITHUB.COM. It then applies the payment on that date and adjusts the balances (principal/interest/fees) accordingly​ GITHUB.COM. After the payment, interest for the rest of the period is accrued on the new, lower balance. This means the period’s interest due is correctly “split” into two parts: pre- and post-payment. If you do not provide a date, the script cannot place the payment in the timeline, so it would assume the payment happens after all scheduled periods (making it ineffective for reducing interest in the interim). Always include the actual payment date for unscheduled entries to ensure the calculation reflects the timing of that prepayment.

---

## 3. Monthly Loans with Actual Day Count (365/Actual days) – Non-Amortizing/Interest-Only

- **Interest Paid:**  
  For a monthly loan that uses actual day counts, interest accrues day by day. When you add an unscheduled Interest Paid entry, the script immediately deducts that from any accumulated interest balance​ GITHUB.COM. If there was unpaid interest from a prior period (say you missed or underpaid a scheduled interest payment), this extra payment will reduce or clear that outstanding interest. Paying interest early (before the normal due date) in an actual day-count loan isn’t common since interest is typically paid as it accrues each period, but the option exists to remove any interest that’s hanging out in the balance. Once paid, that interest stops accruing (the script’s running interest balance is set to zero or lower) so you won’t be charged interest on it later (note: the script doesn’t compound interest on overdue interest – it keeps it separate as an interest balance).

- **Principal Paid:**  
  An unscheduled Principal Paid in a monthly actual loan directly reduces the principal on the date of payment​ GITHUB.COM. The script calculates the exact number of days between the last payment date (or loan start) and the unscheduled payment’s date to find how much interest accrued on the old principal in that interval​ GITHUB.COM. It adds that interest to the interest balance, then applies your principal payment, cutting down the outstanding principal immediately. Going forward from that date, the principal is smaller, so the interest that will accrue each day (and thus each future period’s interest due) is reduced. In an interest-only structure, this means your upcoming scheduled interest payments will be lower. And if a final principal payoff is expected at maturity, that amount will be reduced by the unscheduled principal you paid early.

- **Fees Paid:**  
  Unscheduled Fees Paid in an actual day-count loan work the same way as in the periodic case – the fee balance is decreased as soon as that payment is applied​ GITHUB.COM. If the loan had a fee due later, paying some or all of it now (on a given date) will remove that portion from the outstanding fees. There’s no effect on interest or principal calculations aside from reducing the total obligations. It simply means when the fee would have been due, you owe less (because you’ve already paid part of it).

- **Paid On (Date):**  
  The Paid On date in an unscheduled entry is equally important for actual day-count loans. It pins the extra payment to a specific day, allowing the script to correctly compute accrued interest to that point. When a date is provided, the code will accrue interest from the last period up to that exact day (using the actual number of days) and add it to running interest​ GITHUB.COM. Then the payment is applied on that date, which immediately adjusts the balances (reducing principal/interest/fees as specified). From the next day onward, interest accrues on the new principal. In effect, the schedule acknowledges the early payment: the current period’s interest due is reduced and the remaining balance is lower. Without a Paid On date, the extra payment wouldn’t be inserted into the timeline – the script wouldn’t know when to apply it – so it would not affect the interest calculation for any period (it might only show up as an additional payment at the end, not saving any interest in the interim). Always use the actual payment date for unscheduled payments so the calculation reflects the timing correctly.

---

## 4. Fully Amortizing Monthly Loans (Amortize = "Yes") – Actual or Periodic Day Count

- **Interest Paid:**  
  In a fully amortizing loan, each scheduled payment is supposed to cover that period’s interest in full (plus some principal). Therefore, an unscheduled Interest Paid is usually only needed if there was interest that went unpaid in a prior period. If you do input an Interest Paid in an unscheduled row, the script will immediately apply it to reduce the interest balance​ GITHUB.COM. This could happen if, for example, a scheduled payment was missed or short-paid and interest accrued into the next period – an extra payment can then be made to pay off that lingering interest. Once applied, the running interest is decreased, ensuring that no old interest remains to hinder the amortization. In normal cases (payments made in full), there wouldn’t be an interest balance to pay outside the schedule, so this field is less commonly used for amortizing scenarios except to correct a deficit.

- **Principal Paid:**  
  An unscheduled Principal Paid has a significant effect on an amortizing loan. This is essentially a prepayment of principal on top of the regular installment. The script will apply the principal reduction immediately on the Paid On date, lowering the outstanding balance mid-schedule​ GITHUB.COM. It also recalculates the interest up to that date so that the borrower is charged interest only for the time the original principal was still in the loan​ GITHUB.COM. After that date, the remaining principal is smaller, so the interest for the rest of the period (and future periods) will be computed on that lower balance. By default, the amortization plan assumes a fixed payment amount each period, but a prepayment means that plan is no longer optimal. The library handles this by re-amortizing the future schedule to account for the extra principal payment. In other words, it will recalculate the remaining payment amounts or allocations so that the loan still fully pays off by the end of the term with the new reduced balance​ GITHUB.COM. All subsequent scheduled rows are updated with new Principal Due and Interest Due values reflecting this recalculation. The net effect is that your future monthly payments could decrease or your loan will finish earlier (depending on how the schedule is structured), because you've paid extra principal. The script ensures the adjustment is made such that no negative or “extra” payments occur at the end – the loan is simply paid off sooner or with less due each period. (Internally, a flag is set so that once the schedule is re-amortized, it uses the new values instead of the original amortization plan​ GITHUB.COM.)

- **Fees Paid:**  
  For amortizing loans, fees (like origination or exit fees) might be included as part of the balance or as separate due amounts in certain periods. An unscheduled Fees Paid will immediately reduce any outstanding fees just as in other scenarios​ GITHUB.COM. If the fee was scheduled at a future date (e.g., an exit fee at maturity), paying some of it early will decrease the fee balance and thus lower the amount due when that fee’s period comes. This doesn’t directly alter the amortization of principal and interest – it primarily affects the Fees Due column and the total remaining balance. However, since the total balance (principal + interest + fees) is tracked, an early fee payment will reduce the overall balance figure​ GITHUB.COM. The schedule will show the fee as partially paid in the unscheduled row and the fee due at the final period would effectively be smaller (even though the scheduled fee due in that final period might remain the same on paper, the fee balance carried into that period will be lower due to your prepayment).

- **Paid On (Date):**  
  In a fully amortizing loan, the Paid On date for an unscheduled payment is what allows the script to insert the prepayment into the amortization schedule timeline. When you provide a date, the calculation will include interest accrual up to that day and apply the payment there, just like in the other scenarios. The difference in an amortizing loan is what happens next: the script will recompute the remaining amortization from that point forward. Practically, the steps are: interest is accrued from the last payment date to the unscheduled payment date (so you pay interest only for those days on the old balance), then the extra payment reduces the principal, and then the remaining future payments are adjusted based on the new principal. The Paid On date ensures this all happens at the correct time. If you didn’t specify a date, the extra payment couldn’t be placed in the schedule properly – the script would treat it as if it came after the last period, which means it wouldn’t shorten the loan or reduce any intermediate interest. So, for an amortizing loan, the Paid On date is essential to reap the benefit of paying off principal early: with it, the schedule re-calculates and you see lower balances and updated payment amounts going forward​ GITHUB.COM. Always include the actual payment date for unscheduled principal prepayments; this way, the schedule will show the effect of that payment (less interest accrued after that date and a new amortization schedule for the remaining term).

- **Future Payments Due:**  
  After an unscheduled payment in an amortizing loan, the future payment schedule is adjusted. Initially, amortizing loans have a fixed periodic payment. Once a prepayment happens, the script’s re-amortization will typically result in a new (lower) payment amount for the remaining periods or the same payment amount but fewer periods (the library by default recalculates the remaining payments to fully amortize by the original end date, often yielding a lower monthly payment). For example, if you had 24 months left and you prepaid a chunk of principal, the code will recalculate what the new equal payment should be for the next 24 months to pay off the now-smaller balance​ GITHUB.COM. Those recalculated amounts are written into the schedule for future periods. This means that the “Payment Due” column (total due each period) may change after the unscheduled payment. In contrast, in interest-only scenarios, there is no fixed installment to recalc – instead the interest due each period simply drops in proportion to the lower balance. But in amortizing scenarios, the entire payment structure is updated so that the loan remains on track to finish at the same time, just with less money paid in interest. The key takeaway is: an extra principal payment in an amortizing loan will reduce what you owe in the future, and the script reflects that by lowering either the upcoming payment amounts or the number of payments needed.

---

## Summary of Unscheduled Payment Effects

- **Interest Paid (unscheduled):**  
  Always reduces the accumulated interest balance immediately​ GITHUB.COM. This prevents interest from remaining unpaid. It’s most impactful if interest was accrued from missed payments or in between scheduled dates; once paid, that interest is no longer owed and won’t appear in future due amounts. It does not prepay future interest – it only pays off interest that has accrued up to the payment date (or was scheduled up to that point).

- **Principal Paid (unscheduled):**  
  Always lowers the outstanding principal right away​ GITHUB.COM. Future interest calculations use the new lower principal, so interest due each period will drop going forward. In interest-only loans, this simply reduces the interest portion of future payments (since principal isn’t due until the end, but now the end balance is smaller). In amortizing loans, this triggers a re-amortization of remaining payments, typically reducing the periodic payment amount or the loan term. Either way, you save on interest overall by paying down principal early.

- **Fees Paid (unscheduled):**  
  Immediately reduces any outstanding or future fees​ GITHUB.COM. This lowers the total balance of the loan but doesn’t affect interest accrual on principal (fees are handled separately). It means when a fee would have been due, that required amount is smaller (because you paid part of it in advance). Always record the date for fee payments as well to place them correctly (especially if interest might accrue on overdue fees in a custom scenario, though by default the script treats fees separately from interest).

- **Paid On date:**  
  The Paid On date is critical for every unscheduled entry. It ensures the payment is integrated into the timeline of the loan. The script processes unscheduled payments in chronological order​ GITHUB.COM alongside scheduled periods. By providing the date, you let the script accrue interest correctly up to that point and then apply the payment. The result is an accurate reflection of the loan’s state after the payment: interest is calculated only for the time it was actually outstanding, and principal/fee balances are updated when they should be. If you omit the Paid On date, the script will not know when to apply the payment – effectively the payment would either be ignored in the interim calculations or treated as happening at the end, yielding no benefit in reducing interest or balances until the loan’s end. Always include a Paid On date for unscheduled payments to get the intended outcome in the amortization schedule.

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
