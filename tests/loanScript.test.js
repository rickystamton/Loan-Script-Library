// tests/loanScript.test.js

const { getTotalPeriods, RowManager } = require('../LoanScript.js');

// jest setup file or top of test file:
require('gas-mock-globals');  // This will automatically define SpreadsheetApp, etc.
// Ensure flush is defined
if (typeof SpreadsheetApp.flush !== 'function') {
  SpreadsheetApp.flush = jest.fn();
}

describe('getTotalPeriods', () => {
  test('returns termMonths + 1 when first period is prorated', () => {
    const params = { termMonths: 12, prorateFirst: 'Yes' };
    expect(getTotalPeriods(params)).toBe(13);
  });

  test('returns termMonths when first period is not prorated', () => {
    const params = { termMonths: 12, prorateFirst: 'No' };
    expect(getTotalPeriods(params)).toBe(12);
  });
});

describe('RowManager.handleInsertedRow', () => {
  // Create a fake Sheet object to simulate Google Sheets API behavior
  class FakeRange {
    constructor(sheet, startRow, startCol, numRows = 1, numCols = 1) {
      this.sheet = sheet;
      this.startRow = startRow;
      this.startCol = startCol;
      this.numRows = numRows;
      this.numCols = numCols;
    }
    getValue() {
      // Only valid for 1x1 ranges
      if (this.numRows !== 1 || this.numCols !== 1) {
        throw new Error('getValue called on multi-cell range');
      }
      return this.sheet.getCellValue(this.startRow, this.startCol);
    }
    setValue(val) {
      if (this.numRows !== 1 || this.numCols !== 1) {
        throw new Error('setValue called on multi-cell range');
      }
      this.sheet.setCellValue(this.startRow, this.startCol, val);
    }
    getValues() {
      const values = [];
      for (let r = 0; r < this.numRows; r++) {
        const rowArr = [];
        for (let c = 0; c < this.numCols; c++) {
          rowArr.push(this.sheet.getCellValue(this.startRow + r, this.startCol + c));
        }
        values.push(rowArr);
      }
      return values;
    }
    setValues(valuesMatrix) {
      for (let r = 0; r < this.numRows; r++) {
        for (let c = 0; c < this.numCols; c++) {
          this.sheet.setCellValue(
            this.startRow + r,
            this.startCol + c,
            valuesMatrix[r][c]
          );
        }
      }
    }
    clearContent() {
      for (let r = 0; r < this.numRows; r++) {
        for (let c = 0; c < this.numCols; c++) {
          this.sheet.setCellValue(this.startRow + r, this.startCol + c, '');
        }
      }
    }
  }

  class FakeSheet {
    constructor() {
      // Data stored as a map of "row,col" -> value
      this.data = {};
    }
    getRange(row, col, numRows = 1, numCols = 1) {
      return new FakeRange(this, row, col, numRows, numCols);
    }
    // Helper to set a cell value in the internal data store
    setCellValue(row, col, value) {
      this.data[`${row},${col}`] = value;
    }
    // Helper to get a cell value from the internal data store
    getCellValue(row, col) {
      const key = `${row},${col}`;
      return Object.prototype.hasOwnProperty.call(this.data, key)
        ? this.data[key]
        : null;
    }
  }

  test('inserts a new unscheduled row with period halfway between surrounding periods', () => {
    const sheet = new FakeSheet();
    const rm = new RowManager(sheet);
    // Simulate existing scheduled rows:
    // Row 9 (previous row) with period 1, balances set
    sheet.setCellValue(9, 2, 1);    // Period (col B) on previous row
    sheet.setCellValue(9, 15, 100); // Interest balance (col O) on previous row
    sheet.setCellValue(9, 16, 900); // Principal balance (col P) on previous row
    // Row 11 (next row) with period 2
    sheet.setCellValue(11, 2, 2);   // Period (col B) on next row
    const insertedRow = 10;
    rm.handleInsertedRow(insertedRow);
    // Validate new row (row 10) values
    const newPeriod = sheet.getCellValue(10, 2);
    const newIntBal = sheet.getCellValue(10, 15);
    const newPrinBal = sheet.getCellValue(10, 16);
    const newTotalBal = sheet.getCellValue(10, 17);
    // Period should be halfway between 1 and 2 => 1.5
    expect(newPeriod).toBe(1.5);
    // Balances should carry over from previous row
    expect(newIntBal).toBe(100);
    expect(newPrinBal).toBe(900);
    expect(newTotalBal).toBe(1000);
    // Other fields in new row should be blank or zero as appropriate
    expect(sheet.getCellValue(10, 3)).toBe('');  // Period End (col C)
    expect(sheet.getCellValue(10, 4)).toBe('');  // Due Date (col D)
    expect(sheet.getCellValue(10, 5)).toBe('');  // Days (col E)
    expect(sheet.getCellValue(10, 6)).toBe('');  // Paid On (col F)
    expect(sheet.getCellValue(10, 7)).toBe(0);   // Total Due (col G)
    expect(sheet.getCellValue(10, 8)).toBe(0);   // Total Paid (col H)
    expect(sheet.getCellValue(10, 9)).toBe('');  // Principal Due (col I)
    expect(sheet.getCellValue(10, 10)).toBe(0);  // Principal Paid (col J)
    expect(sheet.getCellValue(10, 11)).toBe(''); // Interest Due (col K)
    expect(sheet.getCellValue(10, 12)).toBe(0);  // Interest Paid (col L)
    expect(sheet.getCellValue(10, 13)).toBe(''); // Fees Due (col M)
    expect(sheet.getCellValue(10, 14)).toBe(0);  // Fees Paid (col N)
    expect(sheet.getCellValue(10, 18)).toBe(''); // Notes (col R)
  });

  test('inserts a new unscheduled row at end of schedule with period incremented by 0.5', () => {
    const sheet = new FakeSheet();
    const rm = new RowManager(sheet);
    // Simulate last existing scheduled row
    sheet.setCellValue(10, 2, 5);    // Last scheduled period = 5
    sheet.setCellValue(10, 15, 50);  // Interest balance
    sheet.setCellValue(10, 16, 500); // Principal balance
    const insertedRow = 11;
    rm.handleInsertedRow(insertedRow);
    // New row period should be 5.5 (prev period 5, no next period defined)
    expect(sheet.getCellValue(11, 2)).toBe(5.5);
    // Balances should carry from previous row
    expect(sheet.getCellValue(11, 15)).toBe(50);
    expect(sheet.getCellValue(11, 16)).toBe(500);
    expect(sheet.getCellValue(11, 17)).toBe(550);
  });

  test('inserts a new unscheduled row at start of schedule with period 0.5 when no prior period exists', () => {
    const sheet = new FakeSheet();
    const rm = new RowManager(sheet);
    // Insert at the very start (no previous scheduled row)
    const insertedRow = 8; // Start of schedule (SHEET_CONFIG.START_ROW)
    rm.handleInsertedRow(insertedRow);
    // With no previous period, expect new period 0.5
    expect(sheet.getCellValue(8, 2)).toBe(0.5);
    // Balances should default to 0
    expect(sheet.getCellValue(8, 15)).toBe(0);
    expect(sheet.getCellValue(8, 16)).toBe(0);
    expect(sheet.getCellValue(8, 17)).toBe(0);
  });
});
