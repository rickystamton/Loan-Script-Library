// jest setup file or top of test file:
require('gas-mock-globals');  // This will automatically define SpreadsheetApp, etc.

// tests/scheduleGeneration.test.js

const { LoanScheduleGenerator, getTotalPeriods } = require('../LoanScript.js');

// Simple mock for Google Sheets Sheet and Range
class MockRange {
  constructor(sheet, startRow, startCol, numRows = 1, numCols = 1) {
    this.sheet = sheet;
    this.startRow = startRow;
    this.startCol = startCol;
    this.numRows = numRows;
    this.numCols = numCols;
  }
  getValue() {
    // Single-cell value
    return this.sheet.grid[this.startRow][this.startCol];
  }
  getDisplayValue() {
    const value = this.getValue();
    // Simulate percentage display formatting for numeric rates/percents
    if (typeof value === 'number' && value < 1 && value > 0) {
      return (value * 100).toFixed(2) + '%';
    }
    // Dates or other values - use default string conversion
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
    // Single-cell set
    this.sheet.grid[this.startRow][this.startCol] = val;
  }
  setValues(arr) {
    // Write a 2D array of values starting at this range
    for (let r = 0; r < arr.length; r++) {
      for (let c = 0; c < arr[0].length; c++) {
        this.sheet.grid[this.startRow + r][this.startCol + c] = arr[r][c];
      }
    }
  }
  clearContent() {
    // Clear the range content (set to empty string)
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        this.sheet.grid[this.startRow + r][this.startCol + c] = "";
      }
    }
  }
  offset(rowOffset, colOffset, numRows = 1, numCols = 1) {
    // Return a new range shifted by the given offset
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
    // Initialize a 2D grid for rows 1-500 and columns A(1) through AA(27)
    const rows = 501;
    const cols = 28;
    this.grid = Array.from({ length: rows }, () => Array(cols).fill(""));
  }
  getRange(a1OrRow, col, numRows, numCols) {
    if (typeof a1OrRow === 'string') {
      // A1 notation parsing (e.g., "D4")
      const rangeStr = a1OrRow;
      // Simple A1 parser for single-cell references and one-row ranges (no ":" handling)
      const match = rangeStr.match(/([A-Z]+)(\d+)/);
      const colLetters = match[1];
      const row = parseInt(match[2], 10);
      // Compute column number from letters
      let colNum = 0;
      for (let i = 0; i < colLetters.length; i++) {
        colNum = colNum * 26 + (colLetters.charCodeAt(i) - 64);
      }
      return new MockRange(this, row, colNum);
    } else {
      // numeric parameters
      const row = a1OrRow;
      const startCol = col;
      const nRows = numRows || 1;
      const nCols = numCols || 1;
      return new MockRange(this, row, startCol, nRows, nCols);
    }
  }
}

// Helper to set loan input values into the mock sheet (row 4 inputs)
function setLoanInputs(sheet, { loanName, principal, annualRate, closingDate, termMonths, prorateFirst, paymentFreq, dayCountMethod, daysPerYear, prepaidIntDate, amortizeYN, origFeePct, exitFeePct }) {
  // Following the README assignments for row 4 inputs:
  sheet.getRange('B4').setValue(loanName || "Test Loan");
  sheet.getRange('D4').setValue(principal);
  sheet.getRange('E4').setValue(annualRate);
  sheet.getRange('F4').setValue(closingDate);
  sheet.getRange('G4').setValue(termMonths);
  sheet.getRange('H4').setValue(prorateFirst);
  sheet.getRange('I4').setValue(paymentFreq);
  sheet.getRange('J4').setValue(dayCountMethod);
  sheet.getRange('K4').setValue(daysPerYear);
  sheet.getRange('L4').setValue(prepaidIntDate || "");  // blank if none
  sheet.getRange('M4').setValue(amortizeYN);
  sheet.getRange('N4').setValue(origFeePct || 0);
  sheet.getRange('O4').setValue(exitFeePct || 0);
}

// Tests for core schedule generation logic
describe('Loan schedule generation', () => {

  test('getTotalPeriods computes an extra period when first period is prorated', () => {
    const paramsNoProrate = { termMonths: 12, prorateFirst: "No" };
    const paramsProrate = { termMonths: 12, prorateFirst: "Yes" };
    expect(getTotalPeriods(paramsNoProrate)).toBe(12);
    expect(getTotalPeriods(paramsProrate)).toBe(13);
  });

  test('buildScheduleData generates correct periods and dates with and without prorate', () => {
    const sheet = new MockSheet();
    // Scenario 1: No prorate
    const params1 = {
      termMonths: 3,
      prorateFirst: "No",
      closingDate: new Date('2023-01-15'),  // mid-month start
      dayCountMethod: "Actual"
    };
    const gen1 = new LoanScheduleGenerator(sheet);
    const schedule1 = gen1.buildScheduleData(params1);
    expect(schedule1.length).toBe(3);                     // 3 periods (no extra)
    expect(schedule1[0][0]).toBe(1);                      // first period number is 1
    expect(schedule1[1][0]).toBe(2);
    expect(schedule1[2][0]).toBe(3);
    // Check that period end dates advance monthly and first period is partial month (Jan15->Feb14)
    expect(schedule1[0][1]).toBeInstanceOf(Date);
    expect(schedule1[1][1]).toBeInstanceOf(Date);
    // Period 1 should end around mid-Feb (for Jan 15 start, no prorate logic sets first period end to Feb 14)
    expect(schedule1[0][1].getMonth()).toBe(1);  // February (month index 1)
    expect(schedule1[0][1].getDate()).toBe(14);
    // Scenario 2: Prorated first period
    const params2 = {
      termMonths: 3,
      prorateFirst: "Yes",
      closingDate: new Date('2023-01-15'),
      dayCountMethod: "Actual"
    };
    const gen2 = new LoanScheduleGenerator(sheet);
    const schedule2 = gen2.buildScheduleData(params2);
    expect(schedule2.length).toBe(4);                     // termMonths+1 periods
    expect(schedule2[0][0]).toBe(0);                      // first period labeled 0
    expect(schedule2[1][0]).toBe(1);
    // First period (0) should end Jan 31 (prorated to month-end)
    expect(schedule2[0][1].getMonth()).toBe(0);           // January (month index 0)
    expect(schedule2[0][1].getDate()).toBe(31);
    // Second period (periodNum 1 in schedule) should then end end-of-Feb
    expect(schedule2[1][1].getMonth()).toBe(1);           // February
    expect(schedule2[1][1].getDate()).toBe(28);           // Feb 28 (2023 not leap year)
  });

  test('Origination fee, prepaid interest, and exit fee are reflected in first and last schedule entries', () => {
    const sheet = new MockSheet();
    // Prepare params: small loan with origination fee, prepaid interest, and exit fee
    const principal = 1000;
    const origFeePct = 0.05;   // 5% origination
    const exitFeePct = 0.02;   // 2% exit
    const annualRate = 0.1;    // 10% interest (for day count calculation, not critical here)
    const closingDate = new Date('2023-01-01');
    // Set inputs in the sheet (getAllInputs will read these)
    setLoanInputs(sheet, {
      loanName: "Fee Test Loan",
      principal: principal,
      annualRate: annualRate,
      closingDate: closingDate,
      termMonths: 2,
      prorateFirst: "No",
      paymentFreq: "Monthly",
      dayCountMethod: "Actual",
      daysPerYear: 365,
      prepaidIntDate: new Date('2023-01-15'),  // prepaid interest up to Jan 15, 2023
      amortizeYN: "Yes",  // (amortize doesn't matter for fee placement)
      origFeePct: origFeePct,
      exitFeePct: exitFeePct
    });
    // Generate schedule data
    const gen = new LoanScheduleGenerator(sheet);
    const params = gen.sheet ? null : null; // (we'll let generateSchedule call getAllInputs internally)
    gen.clearOldSchedule = () => {};        // override formatting/clearing to focus on data
    gen.applyFormatting = () => {};
    // Use generateSchedule to populate the sheet data
    gen.generateSchedule();  
    // The schedule should now be written to the sheet's grid (rows 8+)
    const generated = sheet.getRange(8, 2,  /* B8 */  sheet.grid.length - 8, 17).getValues();
    // Find actual used rows (stop at first empty Period cell)
    let usedRows = 0;
    for (let i = 0; i < generated.length; i++) {
      if (generated[i][0] === "" || generated[i][0] === null) break;
      usedRows++;
    }
    const schedule = generated.slice(0, usedRows);
    expect(schedule.length).toBe(2);             // termMonths=2 (no prorate) -> 2 periods
    // First row (period 1) Notes should mention origination fee and prepaid interest
    const firstNotes = schedule[0][16];
    expect(firstNotes).toMatch(/Origination Fee added to Principal/);
    expect(firstNotes).toMatch(/Prepaid Interest/);
    // It should include the percentage string "5%" and a currency amount for prepaid interest
    expect(firstNotes).toMatch(/5% Origination Fee/);
    expect(firstNotes).toMatch(/\$[0-9,.]+ of Prepaid Interest/);
    // Last row (period 2) should have an exit fee due and a note
    const lastRow = schedule[schedule.length - 1];
    expect(lastRow[11]).toBeCloseTo(principal * exitFeePct, 2);  // FeesDue = 2% of original principal = 20
    expect(lastRow[16]).toContain("Exit Fee");
  });

});
