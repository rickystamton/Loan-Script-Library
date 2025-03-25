// tests/loanCalculations.test.js

const { LoanScheduleGenerator, BalanceManager } = require('../LoanScript.js');

// jest setup file or top of test file:
require('gas-mock-globals');  // This will automatically define SpreadsheetApp, etc.
// Ensure flush is defined
if (typeof SpreadsheetApp.flush !== 'function') {
  SpreadsheetApp.flush = jest.fn();
}

process.env.TZ = 'America/Denver';  // Set timezone for consistent date handling

// Reuse the MockSheet and MockRange from the previous test (they could be moved to a shared helper module)
class MockRange {
  constructor(sheet, startRow, startCol, numRows = 1, numCols = 1) {
    this.sheet = sheet;
    this.startRow = startRow;
    this.startCol = startCol;
    this.numRows = numRows;
    this.numCols = numCols;
  }
  getValue() {
    return this.sheet.grid[this.startRow][this.startCol];
  }
  getDisplayValue() {
    const value = this.getValue();
    if (typeof value === 'number' && value < 1 && value > 0) {
      return (value * 100).toFixed(2) + '%';
    }
    return value instanceof Date ? value.toLocaleDateString('en-US') : String(value || "");
  }
  getValues() {
    const values = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowArr = [];
      for (let c = 0; c < this.numCols; c++) {
        rowArr.push(this.sheet.grid[this.startRow + r][this.startCol + c]);
      }
      values.push(rowArr);
    }
    return values;
  }
  setValue(val) {
    this.sheet.grid[this.startRow][this.startCol] = val;
  }
  setValues(arr) {
    // If writing IPMT/PPMT formulas, compute results
    if (this.startCol === 26 || this.startCol === 27) {  // Z or AA column (IPMT or PPMT)
      const isIpmt = (this.startCol === 26);
      const formulas = arr;
      const lastUsedCount = formulas.length;
      // Find first non-empty formula to extract parameters
      let exampleFormula = "";
      for (const row of formulas) {
        if (row[0] && typeof row[0] === 'string' && row[0].startsWith('=-')) {
          exampleFormula = row[0];
          break;
        }
      }
      if (exampleFormula) {
        // Parse formula like "=-IPMT(rate,period,nper,principal)"
        const parts = exampleFormula.match(/=?-[IP]{1,2}PMT\(\s*([^,\s]+)\s*,\s*([^,\s]+)\s*,\s*([^,\s]+)\s*,\s*([^,\s\)]+)/);
        if (!parts) {
          throw new Error("Failed to parse formula: " + exampleFormula);
        }
        const rate = parseFloat(parts[1]);
        const nper = parseInt(parts[3], 10);
        const principal = parseFloat(parts[4]);
        // Calculate equal payment amount for amortizing loan
        const payment = rate !== 0 
          ? (principal * rate) / (1 - Math.pow(1 + rate, -nper)) 
          : principal / nper;
        // Generate amortization schedule for interest (IPMT) and principal (PPMT) portions
        let balance = principal;
        const interestPortions = [];
        const principalPortions = [];
        for (let period = 1; period <= nper; period++) {
          const interest = rate * balance;
          const principalPortion = payment - interest;
          balance -= principalPortion;
          interestPortions.push(interest);
          principalPortions.push(principalPortion);
        }
        // Fill sheet.grid with calculated results corresponding to each formula row
        const offset = lastUsedCount - nper;  // if there's a prorate period 0, offset = 1
        for (let i = 0; i < lastUsedCount; i++) {
          const val = formulas[i][0];
          if (typeof val === 'string' && val.startsWith('=-')) {
            const periodNum = i + 1 - offset;
            // Use computed values (note: IPMT formulas are negated, but we calculated positive interest)
            this.sheet.grid[this.startRow + i][26] = interestPortions[periodNum - 1] || 0;
            this.sheet.grid[this.startRow + i][27] = principalPortions[periodNum - 1] || 0;
          } else {
            // No formula (unscheduled row or period0): set 0
            this.sheet.grid[this.startRow + i][26] = 0;
            this.sheet.grid[this.startRow + i][27] = 0;
          }
        }
      }
      return;
    }
    // Normal setValues for schedule or input ranges
    for (let r = 0; r < arr.length; r++) {
      for (let c = 0; c < arr[0].length; c++) {
        this.sheet.grid[this.startRow + r][this.startCol + c] = arr[r][c];
      }
    }
  }
  clearContent() {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        this.sheet.grid[this.startRow + r][this.startCol + c] = "";
      }
    }
  }
  offset(rowOffset, colOffset, numRows = 1, numCols = 1) {
    return new MockRange(
      this.sheet,
      this.startRow + rowOffset,
      this.startCol + colOffset,
      numRows,
      numCols
    );
  }
}

class MockSheet {
  constructor() {
    const rows = 501;
    const cols = 28;
    this.grid = Array.from({ length: rows }, () => Array(cols).fill(""));
  }
  getRange(a1OrRow, col, numRows, numCols) {
    if (typeof a1OrRow === 'string') {
      const match = a1OrRow.match(/([A-Z]+)(\d+)/);
      const colLetters = match[1];
      const row = parseInt(match[2], 10);
      let colNum = 0;
      for (let i = 0; i < colLetters.length; i++) {
        colNum = colNum * 26 + (colLetters.charCodeAt(i) - 64);
      }
      return new MockRange(this, row, colNum);
    } else {
      return new MockRange(this, a1OrRow, col, numRows, numCols);
    }
  }
}

// Helper to set input values in the MockSheet
function setLoanInputs(sheet, params) {
  sheet.getRange('B4').setValue(params.loanName || "Test Loan");
  sheet.getRange('D4').setValue(params.principal);
  sheet.getRange('E4').setValue(params.annualRate);
  sheet.getRange('F4').setValue(params.closingDate);
  sheet.getRange('G4').setValue(params.termMonths);
  sheet.getRange('H4').setValue(params.prorateFirst);
  sheet.getRange('I4').setValue(params.paymentFreq);
  sheet.getRange('J4').setValue(params.dayCountMethod);
  sheet.getRange('K4').setValue(params.daysPerYear);
  sheet.getRange('L4').setValue(params.prepaidIntDate || "");
  sheet.getRange('M4').setValue(params.amortizeYN);
  sheet.getRange('N4').setValue(params.origFeePct || 0);
  sheet.getRange('O4').setValue(params.exitFeePct || 0);
}

describe('Loan repayment calculations', () => {

  test('Fully amortizing loan produces equal total payments and zero balance at end', () => {
    const sheet = new MockSheet();
    // Loan parameters for a fully amortizing loan
    setLoanInputs(sheet, {
      loanName: "Amortizing Loan",
      principal: 100000,
      annualRate: 0.05,            // 5% annual interest
      closingDate: new Date('2023-01-01'),
      termMonths: 12,
      prorateFirst: "No",
      paymentFreq: "Monthly",
      dayCountMethod: "Periodic",  // use 30/360 for consistent monthly interest
      daysPerYear: 360,
      prepaidIntDate: "",         // none
      amortizeYN: "Yes",
      origFeePct: 0,
      exitFeePct: 0
    });
    // Generate schedule and recalc balances
    const gen = new LoanScheduleGenerator(sheet);
    const balMgr = new BalanceManager(sheet);
    // Bypass actual Google Sheets-specific formatting calls
    gen.applyFormatting = () => {};
    // Run schedule generation and recalculation
    gen.generateSchedule();
    balMgr.recalcAll();
    // Retrieve the computed schedule from the mock sheet
    const range = sheet.getRange(8, 2, 500 - 8 + 1, 17);  // full schedule area
    const allRows = range.getValues();
    // Filter out the generated rows (stop at first blank period cell)
    const schedule = allRows.filter(row => row[0] !== "" && row[0] !== null);
    const numPeriods = schedule.length;
    expect(numPeriods).toBe(12);
    // Check that Total Due (Col G) is the same (or nearly the same) each period
    const totalDueValues = schedule.map(row => row[5]);
    // All payments should be equal (within a few cents of each other due to rounding)
    const firstPayment = totalDueValues[0];
    const lastPayment = totalDueValues[totalDueValues.length - 1];
    expect(firstPayment).toBeCloseTo(lastPayment, 2);
    // Check that interest due (Col K) decreases each period and principal due (Col I) increases
    for (let i = 1; i < schedule.length; i++) {
      const prevInterestDue = schedule[i-1][9];
      const currInterestDue = schedule[i][9];
      expect(currInterestDue).toBeLessThan(prevInterestDue + 1e-6);  // strictly decreasing
      const prevPrincipalDue = schedule[i-1][7];
      const currPrincipalDue = schedule[i][7];
      expect(currPrincipalDue).toBeGreaterThan(prevPrincipalDue - 1e-6);
    }
    // Sum of Principal Due over all periods should equal the initial principal
    const totalPrincipalPaid = schedule.reduce((sum, row) => sum + row[7], 0);
    expect(totalPrincipalPaid).toBeCloseTo(100000, 2);
  });

  test('Interest-only loan has interest due each period and principal due at maturity', () => {
    const sheet = new MockSheet();
    // Loan parameters for an interest-only loan (principal repaid at end)
    setLoanInputs(sheet, {
      loanName: "Interest-Only Loan",
      principal: 50000,
      annualRate: 0.06,            // 6% annual interest
      closingDate: new Date('2023-01-01'),
      termMonths: 6,
      prorateFirst: "No",
      paymentFreq: "Monthly",
      dayCountMethod: "Periodic",  // 30/360 for simplicity
      daysPerYear: 360,
      prepaidIntDate: "",
      amortizeYN: "No",            // interest-only
      origFeePct: 0,
      exitFeePct: 0
    });
    const gen = new LoanScheduleGenerator(sheet);
    const balMgr = new BalanceManager(sheet);
    gen.applyFormatting = () => {};
    gen.generateSchedule();
    balMgr.recalcAll();
    // Get the schedule output
    const allRows = sheet.getRange(8, 2, 500 - 8 + 1, 17).getValues();
    const schedule = allRows.filter(row => row[0] !== "" && row[0] !== null);
    const numPeriods = schedule.length;
    expect(numPeriods).toBe(6);
    // In an interest-only loan, each period's Principal Due (col I) should be 0 except final, and Interest Due (col K) should be constant each period (since no principal amortization).
    const principalDueVals = schedule.map(r => r[7]);
    const interestDueVals = schedule.map(r => r[9]);
    // Principal due is zero for all but last period
    for (let i = 0; i < numPeriods - 1; i++) {
      expect(principalDueVals[i]).toBeCloseTo(0, 10);
    }
    // Last period principal due equals original principal
    expect(principalDueVals[numPeriods - 1]).toBeCloseTo(50000, 2);
    // Interest due each period (for 6% annual on 30/360, monthly interest = 0.5% of principal = 250)
    const expectedMonthlyInterest = 50000 * 0.06 / 12;  // = 250
    for (let i = 0; i < numPeriods - 1; i++) {
      expect(interestDueVals[i]).toBeCloseTo(expectedMonthlyInterest, 2);
    }
    // If no interest was paid until the end, the final period's Interest Due should include all accrued interest.
    // However, in our case with paymentFreq "Monthly", the script treats interest as due each period, so final Interest Due should just be one period's interest (if prior interest was not paid, it would reflect in Interest Balance).
    const finalInterestDue = interestDueVals[numPeriods - 1];
    expect(finalInterestDue).toBeCloseTo(expectedMonthlyInterest, 2);
    // Check that the interest balance column (col O) accumulates unpaid interest:
    // Since we did not mark any interest as paid, interestBalance in final row should equal sum of all interest (which would be 5 * 250 = 1250) if interest was left unpaid.
    const interestBalanceVals = schedule.map(r => r[13]);
    const finalInterestBalance = interestBalanceVals[numPeriods - 1];
    const sumInterestAccrued = expectedMonthlyInterest * (numPeriods - 1);
    expect(finalInterestBalance).toBeCloseTo(sumInterestAccrued, 2);
    // Total balance (col Q) of final period should equal principal + accrued interest (since nothing paid)
    const finalTotalBalance = schedule[numPeriods - 1][15];
    expect(finalTotalBalance).toBeCloseTo(50000 + sumInterestAccrued, 2);
  });

});