// tests/loanHelpers.test.js

const { separateRows, applyUnscheduledPaymentsForPeriod, calculateDueAmounts } = require('../LoanHelpers.js');

// jest setup file or top of test file:
require('gas-mock-globals');  // This will automatically define SpreadsheetApp, etc.
// Ensure flush is defined
if (typeof SpreadsheetApp.flush !== 'function') {
  SpreadsheetApp.flush = jest.fn();
}

describe('separateRows', () => {
  test('splits scheduled and unscheduled rows and sorts them by date', () => {
    // Create sample schedule data
    const dateA = new Date(2025, 0, 10); // Jan 10, 2025
    const dateB = new Date(2025, 0, 15); // Jan 15, 2025
    const payDateA = new Date(2025, 0, 5);  // Jan 5, 2025
    const payDateB = new Date(2025, 0, 12); // Jan 12, 2025
    const allRows = [
      [1, dateB, "", "", "", null],           // Scheduled row (Period 1, End Jan 15)
      [2, dateA, "", "", "", null],           // Scheduled row (Period 2, End Jan 10)
      ["", null, "", "", payDateB, null, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], // Unscheduled row (Paid On Jan 12)
      ["", null, "", "", payDateA, null, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]  // Unscheduled row (Paid On Jan 5)
    ];
    const count = 4;
    const { scheduledRows, unscheduledRows } = separateRows(allRows, count);
    // Should identify two scheduled and two unscheduled rows
    expect(scheduledRows.length).toBe(2);
    expect(unscheduledRows.length).toBe(2);
    // Scheduled rows sorted by Period End date (earliest first)
    expect(scheduledRows[0].rowIndex).toBe(1); // row with Jan 10
    expect(scheduledRows[1].rowIndex).toBe(0); // row with Jan 15
    // Unscheduled rows sorted by Paid On date (earliest first)
    expect(unscheduledRows[0].rowIndex).toBe(3); // row with Jan 5
    expect(unscheduledRows[1].rowIndex).toBe(2); // row with Jan 12
  });
});

describe('applyUnscheduledPaymentsForPeriod', () => {
  test('accrues interest over full period when there are no unscheduled payments', () => {
    const periodStart = new Date(2025, 0, 1);
    const periodEnd = new Date(2025, 0, 31);
    const params = {
      paymentFreq: "Monthly",
      dayCountMethod: "Actual",
      perDiemRate: 0.01 // 1% per day for easy calculation
    };
    const unscheduledRows = [];
    const result = applyUnscheduledPaymentsForPeriod(
      1, periodStart, periodEnd, params,
      unscheduledRows, 0,
      1000, 0, 0
    );
    // 31 days at 1% per day on principal 1000 => interestAccrued ~ 310
    expect(result.runningPrincipal).toBe(1000);
    expect(result.runningInterest).toBeCloseTo(310, 5);
    expect(result.runningFees).toBe(0);
    expect(result.unschedIndex).toBe(0);
    expect(result.interestAccrued).toBeCloseTo(310, 5);
    expect(result.unscheduledPrincipalPaid).toBe(0);
  });

  test('applies an unscheduled payment within a periodic interest period', () => {
    const periodStart = new Date(2025, 0, 1);
    const periodEnd = new Date(2025, 0, 30); // 30-day period
    const params = {
      paymentFreq: "Monthly",
      dayCountMethod: "Periodic",
      monthlyRate: 0.03 // 3% per month (dailyPeriodicRate = 0.001)
    };
    // One unscheduled payment on mid-period (Jan 15, 2025)
    const paidOnDate = new Date(2025, 0, 15);
    const unschedRow = new Array(16).fill(0);
    unschedRow[0] = "";         // Period blank
    unschedRow[4] = paidOnDate; // Paid On date
    unschedRow[8] = 200;        // Principal Paid (J)
    unschedRow[10] = 5;         // Interest Paid (L)
    unschedRow[11] = 0;         // Fees Due (M)
    unschedRow[12] = 0;         // Fees Paid (N)
    const unscheduledRows = [{ rowIndex: 0, rowData: unschedRow }];
    const result = applyUnscheduledPaymentsForPeriod(
      2, periodStart, periodEnd, params,
      unscheduledRows, 0,
      1000, 0, 0
    );
    // Principal should reduce by the unscheduled principal payment
    expect(result.runningPrincipal).toBe(800);
    // Total interest accrued over the period (30 days at 0.1% per day on principal) ~ 30
    expect(result.interestAccrued).toBeCloseTo(30, 5);
    // Running interest after subtracting paid interest (5) should be interestAccrued - 5
    expect(result.runningInterest).toBeCloseTo(25, 5);
    expect(result.runningFees).toBe(0);
    expect(result.unschedIndex).toBe(1);
    expect(result.unscheduledPrincipalPaid).toBe(200);
    // The unscheduled row should be updated with total paid and new balances
    const uData = unscheduledRows[0].rowData;
    expect(uData[6]).toBe(205);            // Total Paid (H) = 200 + 5 + 0
    expect(uData[13]).toBeCloseTo(25, 5);  // Interest balance (O) after payment
    expect(uData[14]).toBe(800);          // Principal balance (P) after payment
    expect(uData[15]).toBeCloseTo(825, 5); // Total balance (Q) after payment
  });

  test('applies an unscheduled payment in an actual day-count period', () => {
    const periodStart = new Date(2025, 0, 1);
    const periodEnd = new Date(2025, 0, 10); // 10-day period
    const params = {
      paymentFreq: "Monthly",
      dayCountMethod: "Actual",
      perDiemRate: 0.002 // 0.2% per day
    };
    // Unscheduled payment on Jan 5, 2025
    const paidOnDate = new Date(2025, 0, 5);
    const unschedRow = new Array(16).fill(0);
    unschedRow[0] = "";
    unschedRow[4] = paidOnDate;
    unschedRow[8] = 500; // Principal Paid
    unschedRow[10] = 2;  // Interest Paid
    unschedRow[11] = 0;
    unschedRow[12] = 0;
    const unscheduledRows = [{ rowIndex: 0, rowData: unschedRow }];
    const result = applyUnscheduledPaymentsForPeriod(
      1, periodStart, periodEnd, params,
      unscheduledRows, 0,
      1000, 0, 0
    );
    // Principal should drop by 500 due to unscheduled payment
    expect(result.runningPrincipal).toBe(500);
    // After payment on Jan 5, no interest remains at that moment (interest paid covered it)
    const uData = unscheduledRows[0].rowData;
    expect(uData[13]).toBe(0);   // Interest balance (O) after payment
    expect(uData[14]).toBe(500); // Principal balance (P) after payment
    expect(uData[15]).toBe(500); // Total balance (Q) after payment
    // Interest accrues from Jan 6 to Jan 10 on remaining principal (5 days at 0.2%/day on 500) = 5
    expect(result.interestAccrued).toBeCloseTo(5, 5);
    expect(result.runningInterest).toBeCloseTo(5, 5);
    expect(result.unschedIndex).toBe(1);
    expect(result.unscheduledPrincipalPaid).toBe(500);
  });
});

describe('calculateDueAmounts', () => {
  test('handles single-period loan: only final period has all interest and principal due', () => {
    const params = { paymentFreq: "Single Period", termMonths: 2, principal: 1000 };
    // Non-final period (period 1 of 2) – nothing due yet
    let result = calculateDueAmounts(1, 0, params, 50, 0, 0, false, false, false, []);
    expect(result.newInterestDue).toBe(0);
    expect(result.newPrincipalDue).toBe(0);
    // Final period (period 2 of 2) – all accrued interest and full principal due
    result = calculateDueAmounts(2, 1, params, 50, 0, 0, false, false, false, []);
    expect(result.newInterestDue).toBe(50);
    expect(result.newPrincipalDue).toBe(1000);
  });

  test('handles interest-only loan: interest due each period, principal due at final period', () => {
    const params = { paymentFreq: "Monthly", amortizeYN: "No", termMonths: 3, principal: 500 };
    // Intermediate period of interest-only loan
    let result = calculateDueAmounts(1, 0, params, 20, 0, 0, false, false, false, []);
    expect(result.newInterestDue).toBe(20);
    expect(result.newPrincipalDue).toBe(0);
    // Final period of interest-only loan
    result = calculateDueAmounts(3, 2, params, 30, 0, 0, false, false, false, []);
    expect(result.newInterestDue).toBe(30);
    expect(result.newPrincipalDue).toBe(500);
  });

  test('uses scheduled payment split when no prepayments in an amortizing loan period', () => {
    const params = { paymentFreq: "Monthly", amortizeYN: "Yes", termMonths: 5 };
    const scheduledInt = 30;
    const scheduledPr = 70;
    const result = calculateDueAmounts(2, 1, params, 40, scheduledInt, scheduledPr, false, false, false, []);
    // Should output original amortized values (interest and principal due unchanged)
    expect(result.newInterestDue).toBe(scheduledInt);
    expect(result.newPrincipalDue).toBe(scheduledPr);
  });

  test('adjusts interest/principal due when an extra payment is made in the period', () => {
    const params = { paymentFreq: "Monthly", amortizeYN: "Yes", termMonths: 5 };
    const scheduledInt = 30;
    const scheduledPr = 70;
    const scheduledPayment = scheduledInt + scheduledPr; // 100
    // Scenario 1: interest accrued less than scheduled payment
    let result = calculateDueAmounts(3, 2, params, 20, scheduledInt, scheduledPr, false, true, false, []);
    expect(result.newInterestDue).toBe(20);
    expect(result.newPrincipalDue).toBe(80); // remaining part of 100
    // Scenario 2: interest accrued exceeds scheduled payment
    result = calculateDueAmounts(3, 2, params, 120, scheduledInt, scheduledPr, false, true, false, []);
    expect(result.newInterestDue).toBe(scheduledPayment); // capped at 100
    expect(result.newPrincipalDue).toBe(0);
  });

  test('adjusts payment split if a prior period had an extra payment (no new extra this period)', () => {
    const params = { paymentFreq: "Monthly", amortizeYN: "Yes", termMonths: 5 };
    const scheduledInt = 25;
    const scheduledPr = 75;
    const scheduledPayment = scheduledInt + scheduledPr; // 100
    const result = calculateDueAmounts(4, 3, params, 40, scheduledInt, scheduledPr, true, false, false, []);
    // Should behave like an extra payment happened previously:
    // interest due is accrued (40, capped to 100 if needed), principal due fills remaining scheduled payment (60)
    expect(result.newInterestDue).toBe(40);
    expect(result.newPrincipalDue).toBe(60);
  });

  test('uses stored due values if the period was already re-amortized from a prior prepayment', () => {
    const params = { paymentFreq: "Monthly", amortizeYN: "Yes", termMonths: 5 };
    // Simulate that this period was re-amortized and rowData contains adjusted due values
    const rowData = [];
    rowData[7] = 80;  // Stored Principal Due (col I)
    rowData[9] = 15;  // Stored Interest Due (col K)
    const result = calculateDueAmounts(5, 4, params, 50, 30, 70, false, false, true, rowData);
    expect(result.newInterestDue).toBe(15);
    expect(result.newPrincipalDue).toBe(80);
  });
});
