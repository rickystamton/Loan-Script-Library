/*********************************************************************************
 * FAST LOAN SCHEDULE SCRIPT
 *********************************************************************************/

// ---------------------
// 1) SHEET CONFIG
// ---------------------
const SHEET_CONFIG = {
  START_ROW: 8,
  END_ROW: 500,
  COLUMNS: {
    PERIOD:        2,  // B
    PERIOD_END:    3,  // C
    DUE_DATE:      4,  // D
    DAYS:          5,  // E
    PAID_ON:       6,  // F
    TOTAL_DUE:     7,  // G
    TOTAL_PAID:    8,  // H
    PRINCIPAL_DUE: 9,  // I
    PRINCIPAL_PD:  10, // J
    INTEREST_DUE:  11, // K
    INTEREST_PD:   12, // L
    FEES_DUE:      13, // M
    FEES_PD:       14, // N
    INT_BAL:       15, // O
    PRIN_BAL:      16, // P
    TOTAL_BAL:     17, // Q
    NOTES:         18  // R
  },
  INPUTS: {
    LOAN_NAME:               'B4',
    BORROWER_NAME:           'C4',
    PRINCIPAL:               'D4',
    INTEREST_RATE:           'E4',
    CLOSING_DATE:            'F4',
    TERM_MONTHS:             'G4',
    PRORATE:                 'H4',
    PAYMENT_FREQ:            'I4',
    DAY_COUNT:               'J4',
    DAYS_PER_YEAR:           'K4',
    PREPAID_INTEREST_DATE:   'L4',
    AMORTIZE:                'M4',
    ORIG_FEE_PCT:            'N4',
    EXIT_FEE_PCT:            'O4',
    LOCK_INPUTS:             'Q4'
  },
  // For IPMT/PPMT storage:
  IPMT_COL: 26, // Z
  PPMT_COL: 27  // AA
};

// ---------------------
// 2) HELPER FUNCTIONS
// ---------------------

/**
 * Returns the integer number of days between two dates (ignoring time of day).
 * The result is (endDate - startDate) / 86400000, rounded.
 */
function daysBetween(startDate, endDate) {
  if (!(startDate instanceof Date) || !(endDate instanceof Date)) return 0;
  const msPerDay = 24 * 3600 * 1000;
  const s = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  const e = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
  return Math.round((e - s) / msPerDay);
}

/**
 * Returns daysBetween(...) + 1, i.e., inclusive day count.
 */
// [REFINED] Centralize “inclusive” day counting.
function daysBetweenInclusive(startDate, endDate) {
  return daysBetween(startDate, endDate) + 1;
}

/**
 * Get the last day of the same month as dateObj.
 */
function getLastDayOfMonth(dateObj) {
  return new Date(dateObj.getFullYear(), dateObj.getMonth() + 1, 0);
}

/**
 * Get the last day of the month after adding `monthsToAdd`.
 * For example, if dateObj=3/14/2024 and monthsToAdd=1 => returns 4/30/2024.
 */
function getLastDayAfterAddingMonths(dateObj, monthsToAdd) {
  const year = dateObj.getFullYear();
  const monthIndex = dateObj.getMonth() + monthsToAdd + 1;
  const newYear = year + Math.floor(monthIndex / 12);
  const newMonth = monthIndex % 12;
  return new Date(newYear, newMonth, 0);
}

/**
 * Returns a new Date that is exactly +1 day of the given `dateObj`.
 */
// [REFINED] Common single “+1 day” logic
function oneDayAfter(dateObj) {
  const d = new Date(dateObj);
  d.setDate(d.getDate() + 1);
  return d;
}

/**
 * Checks whether the closingDate is an "edge" day for forcing No prorate (day=1 or day>28).
 */
function isEdgeDay(dateObj) {
  const d = dateObj.getDate();
  return (d === 1 || d > 28);
}

/**
 * Return true if row is an unscheduled row (i.e., B-col is not an integer or period-end is blank).
 */
function isUnscheduledRow(rowArr) {
  const periodNum = rowArr[0];
  const periodEnd = rowArr[1];
  if (!periodEnd) return true;
  if (!Number.isInteger(periodNum)) return true;
  return false;
}

/**
 * Finds the last scheduled row's 'Period End Date' going backward from rowIndex in schedule.
 */
function findLastScheduledEnd(schedule, rowIndex) {
  for (let i = rowIndex; i >= 0; i--) {
    const periodNum = schedule[i][0];
    const endVal    = schedule[i][1];
    if (Number.isInteger(periodNum) && endVal instanceof Date && !isNaN(endVal)) {
      return endVal;
    }
  }
  return null;
}

/**
 * Calculates the inclusive day count from periodStart..periodEnd,
 * minus any 'prepaidUntil' portion (if applicable).
 */
function calcUnpaidDays(periodStart, periodEnd, prepaidUntil) {
  if (!(periodEnd instanceof Date) || isNaN(periodEnd)) {
    return 0;
  }
  const raw = daysBetweenInclusive(periodStart, periodEnd);
  if (!prepaidUntil) {
    return raw;
  }
  if (periodEnd < prepaidUntil) {
    return 0;
  }
  if (periodStart >= prepaidUntil) {
    return raw;
  }
  const afterPrepaid = new Date(prepaidUntil);
  return daysBetweenInclusive(afterPrepaid, periodEnd);
}

/**
 * Gets the overlapping day count for schedule[r] using the logic in the original code.
 */
function getOverlapDays(r, schedule, params) {
  const rowArr    = schedule[r];
  const periodNum = rowArr[0];
  const periodEnd = rowArr[1];

  if (!Number.isInteger(periodNum) || !periodEnd) {
    return 0; // unscheduled row or blank date
  }

  if (params.paymentFreq === "Single Period") {
    // Single Period:
    if (r === 0) {
      return calcUnpaidDays(params.closingDate, periodEnd, params.prepaidUntil);
    }
    const lastEnd = findLastScheduledEnd(schedule, r - 1);
    if (lastEnd) {
      const ps = oneDayAfter(lastEnd);
      return calcUnpaidDays(ps, periodEnd, params.prepaidUntil);
    }
    // fallback
    return calcUnpaidDays(params.closingDate, periodEnd, params.prepaidUntil);
  }

  // Monthly:
  if (params.dayCountMethod === "Periodic") {
    // If first row and "Prorate = Yes," do a true difference.
    if (r === 0 && params.prorateFirst === "Yes") {
      return calcUnpaidDays(params.closingDate, periodEnd, params.prepaidUntil);
    }
    // If no prepaid date, just return 30.
    if (!params.prepaidUntil) {
      return 30;
    }
    // Otherwise, compute actual unpaid days, but cap at 30.
    const lastEnd = findLastScheduledEnd(schedule, r - 1) || params.closingDate;
    const periodStart = oneDayAfter(lastEnd);
    const rawDays = calcUnpaidDays(periodStart, periodEnd, params.prepaidUntil);
    return Math.min(rawDays, 30);
  } else {
    // dayCountMethod === "Actual"
    const lastEnd = findLastScheduledEnd(schedule, r - 1);
    if (lastEnd) {
      const ps = oneDayAfter(lastEnd);
      return calcUnpaidDays(ps, periodEnd, params.prepaidUntil);
    }
    return calcUnpaidDays(params.closingDate, periodEnd, params.prepaidUntil);
  }
}

/**
 * Build the set of user inputs from row 4 (taking into account forced changes to certain fields).
 */
function getAllInputs(sheet) {
  const feeCell        = sheet.getRange(SHEET_CONFIG.INPUTS.ORIG_FEE_PCT);
  const exitFeePctCell = sheet.getRange(SHEET_CONFIG.INPUTS.EXIT_FEE_PCT); 
  const datePrepaidInt = sheet.getRange(SHEET_CONFIG.INPUTS.PREPAID_INTEREST_DATE).getValue();

  const inputs = {
    principal      : sheet.getRange(SHEET_CONFIG.INPUTS.PRINCIPAL).getValue(),
    closingDate    : new Date(sheet.getRange(SHEET_CONFIG.INPUTS.CLOSING_DATE).getValue()),
    annualRate     : sheet.getRange(SHEET_CONFIG.INPUTS.INTEREST_RATE).getValue(),
    paymentFreq    : sheet.getRange(SHEET_CONFIG.INPUTS.PAYMENT_FREQ).getValue(),
    dayCountMethod : sheet.getRange(SHEET_CONFIG.INPUTS.DAY_COUNT).getValue(),
    daysPerYear    : sheet.getRange(SHEET_CONFIG.INPUTS.DAYS_PER_YEAR).getValue(),
    termMonths     : sheet.getRange(SHEET_CONFIG.INPUTS.TERM_MONTHS).getValue(),
    prorateFirst   : sheet.getRange(SHEET_CONFIG.INPUTS.PRORATE).getValue(),
    amortizeYN     : sheet.getRange(SHEET_CONFIG.INPUTS.AMORTIZE).getValue(),

    prepaidIntDate : (datePrepaidInt instanceof Date && !isNaN(datePrepaidInt))
                     ? new Date(datePrepaidInt)
                     : null,

    origFeePct       : feeCell.getValue() || 0,
    origFeePctString : feeCell.getDisplayValue() || "",
    exitFeePct       : exitFeePctCell.getValue() || 0
  };

  // [REFINED] Force "No" prorate if the day=1 or day>28
  if (isEdgeDay(inputs.closingDate)) {
    inputs.prorateFirst = "No";
    sheet.getRange(SHEET_CONFIG.INPUTS.PRORATE).setValue("No");
  }

  // If single period => dayCount="Actual", amortize="No"
  if (inputs.paymentFreq === "Single Period") {
    inputs.dayCountMethod = "Actual";
    sheet.getRange(SHEET_CONFIG.INPUTS.DAY_COUNT).setValue("Actual");

    inputs.amortizeYN = "No";
    sheet.getRange(SHEET_CONFIG.INPUTS.AMORTIZE).setValue("No");
  }

  // Some rates:
  inputs.perDiemRate  = inputs.daysPerYear ? inputs.annualRate / inputs.daysPerYear : 0;
  inputs.monthlyRate  = inputs.annualRate / 12;

  // 1) Add orig fee to principal
  let financedFee = 0;
  if (inputs.origFeePct > 0) {
    financedFee = inputs.principal * inputs.origFeePct;
    inputs.principal += financedFee;
  }

  // 2) Prepaid interest
  let financedPrepaidInterest = 0;
  if (inputs.prepaidIntDate && inputs.daysPerYear && inputs.annualRate) {
    const dayCount = daysBetweenInclusive(inputs.closingDate, inputs.prepaidIntDate);
    if (dayCount > 0) {
      const fractionOfYear = dayCount / inputs.daysPerYear;
      const numerator   = inputs.principal * inputs.annualRate * fractionOfYear;
      const denominator = 1 - (inputs.annualRate * fractionOfYear);
      if (denominator !== 0) {
        financedPrepaidInterest = numerator / denominator;
        inputs.principal += financedPrepaidInterest;
      }
    }
  }

  inputs.financedFee             = financedFee; 
  inputs.financedPrepaidInterest = financedPrepaidInterest;

  // We'll define a "prepaidUntil" date
  if (inputs.prepaidIntDate) {
    inputs.prepaidUntil = oneDayAfter(inputs.prepaidIntDate);
  } else {
    inputs.prepaidUntil = null;
  }

  // EXIT FEE (based on original principal, not financed)
  const originalPrincipal = sheet.getRange(SHEET_CONFIG.INPUTS.PRINCIPAL).getValue();
  inputs.exitFee = inputs.exitFeePct * originalPrincipal;

  return inputs;
}

/**
 * Returns the total number of rows to generate in the schedule (including 0 if prorated).
 */
function getTotalPeriods(params) {
  return (params.prorateFirst === "Yes")
    ? params.termMonths + 1
    : params.termMonths;
}

/**
 * Calculates the Period End Date for the i‐th row if `prorateFirst=Yes`.
 */
function calcPeriodEndDate_Prorate(closingDate, i) {
  if (i === 0) {
    return getLastDayOfMonth(closingDate);
  } else {
    return getLastDayAfterAddingMonths(closingDate, i);
  }
}

/**
 * Calculates the Period End Date for the given `periodNum` if `prorateFirst=No`.
 */
function calcPeriodEndDate_NoProrate(closingDate, periodNum) {
  const day = closingDate.getDate();
  if (periodNum === 1) {
    if (day === 1) {
      // If the closing date is the 1st, the first period ends at the last day of the same month.
      return getLastDayOfMonth(closingDate);
    } else if (day > 28) {
      // If the closing day is >28, the first period ends at the last day of the following month.
      return getLastDayAfterAddingMonths(closingDate, 1);
    } else {
      // Otherwise, use the “day-1” logic (e.g., a closing date of 15 will end on the 14th of next month).
      const m = new Date(closingDate.getFullYear(), closingDate.getMonth() + 1, day);
      m.setDate(m.getDate() - 1);
      return m;
    }
  } else {
    // For subsequent periods, you might want to continue a similar logic.
    if (day === 1) {
      // For a closing day of 1, each period's end is the last day of the month for the month in question.
      const baseDate = new Date(closingDate.getFullYear(), closingDate.getMonth() + periodNum - 1, 1);
      return getLastDayOfMonth(baseDate);
    } else if (day > 28) {
      return getLastDayAfterAddingMonths(closingDate, periodNum);
    } else {
      const t = new Date(closingDate.getFullYear(), closingDate.getMonth() + periodNum, day);
      t.setDate(t.getDate() - 1);
      return t;
    }
  }
}

// ---------------------
// 3) SCHEDULE GENERATOR
// ---------------------
class LoanScheduleGenerator {
  constructor(sheet) {
    this.sheet = sheet;
  }

  generateSchedule() {
    const params = getAllInputs(this.sheet);
    this.clearOldSchedule();
    const data = this.buildScheduleData(params);

    if (data.length > 0) {
      this.sheet
        .getRange(SHEET_CONFIG.START_ROW, 2, data.length, data[0].length)
        .setValues(data);
    }
    this.applyFormatting(data.length);

    // Recalc final balances
    const bal = new BalanceManager(this.sheet);
    bal.recalcAll();
  }

  clearOldSchedule() {
    const numRows = SHEET_CONFIG.END_ROW - SHEET_CONFIG.START_ROW + 1;
    this.sheet.getRange(SHEET_CONFIG.START_ROW, 2, numRows, 17).clearContent();
  }

  buildScheduleData(params) {
    const totalPeriods = getTotalPeriods(params);
    const rows = [];

    for (let i = 0; i < totalPeriods; i++) {
      const periodNum = (params.prorateFirst === "Yes") ? i : i + 1;
      let periodEnd;

      // Period End date
      if (params.prorateFirst === "Yes") {
        periodEnd = calcPeriodEndDate_Prorate(params.closingDate, i);
      } else {
        periodEnd = calcPeriodEndDate_NoProrate(params.closingDate, periodNum);
      }

      // Due date = +1 day
      const dueDate = oneDayAfter(periodEnd);

      // Approx “Days in period” for display
      let approxDays = 0;
      if (i === 0 && params.prorateFirst === "Yes") {
        approxDays = daysBetweenInclusive(params.closingDate, periodEnd);
      } else if (params.dayCountMethod === "Periodic") {
        approxDays = 30;
      } else {
        // dayCountMethod===Actual or first row no‐prorate
        if (i === 0) {
          approxDays = daysBetween(params.closingDate, periodEnd);
        } else {
          let prevEnd;
          if (params.prorateFirst === "Yes") {
            prevEnd = calcPeriodEndDate_Prorate(params.closingDate, i - 1);
          } else {
            prevEnd = calcPeriodEndDate_NoProrate(params.closingDate, periodNum - 1);
          }
          approxDays = daysBetween(prevEnd, periodEnd);
        }
      }
      if (approxDays < 0) approxDays = 0;

      rows.push([
        periodNum,    // B => [0]
        periodEnd,    // C => [1]
        dueDate,      // D => [2]
        approxDays,   // E => [3]
        "",           // F => [4] (PaidOn)
        0,            // G => [5] (TotalDue)
        0,            // H => [6] (TotalPaid)
        0,            // I => [7] (PrincipalDue)
        0,            // J => [8] (PrincipalPaid)
        0,            // K => [9] (InterestDue)
        0,            // L => [10](InterestPaid)
        0,            // M => [11](FeesDue)
        0,            // N => [12](FeesPaid)
        0,            // O => [13](InterestBalance)
        0,            // P => [14](PrincipalBalance)
        0,            // Q => [15](TotalBalance)
        ""            // R => [16](Notes)
      ]);
    }

    // Build the note for Orig/Prepaid (exclude exit fee from this note)
    const fee = params.financedFee;
    const pre = params.financedPrepaidInterest;
    let noteFirst = "";
    if (fee > 0 && pre > 0) {
      const preFmt = pre.toLocaleString("en-US", { style: "currency", currency: "USD" });
      noteFirst = `(${preFmt} of Prepaid Interest + ${params.origFeePctString} Origination Fee added to Principal.)`;
    } else if (fee > 0) {
      noteFirst = `(${params.origFeePctString} Origination Fee added to Principal.)`;
    } else if (pre > 0) {
      const preFmt = pre.toLocaleString("en-US", { style: "currency", currency: "USD" });
      noteFirst = `(${preFmt} of Prepaid Interest added to Principal.)`;
    }
    if (noteFirst && rows.length > 0) {
      rows[0][16] = noteFirst; // place in R of first row
    }

    // Place EXIT FEE in final scheduled row’s FeesDue
    if (params.exitFee > 0 && rows.length > 0) {
      const lastRowIndex = rows.length - 1;
      rows[lastRowIndex][11] = params.exitFee; // M=FeesDue
      // Also put a note in the final row
      const exitFeeFmt = params.exitFee.toLocaleString("en-US", { style: "currency", currency: "USD" });
      rows[lastRowIndex][16] = `(${exitFeeFmt} Exit Fee)`;
    }

    return rows;
  }

  applyFormatting(numRows) {
    if (numRows <= 0) return;
    const sr = SHEET_CONFIG.START_ROW;
    const sh = this.sheet;

    // Date columns: C, D, F
    sh.getRange(sr, 3, numRows, 1).setNumberFormat("MM/dd/yyyy");
    sh.getRange(sr, 4, numRows, 1).setNumberFormat("MM/dd/yyyy");
    sh.getRange(sr, 6, numRows, 1).setNumberFormat("MM/dd/yyyy");

    // E => integer
    sh.getRange(sr, 5, numRows, 1).setNumberFormat("0");

    // G..N => currency
    sh.getRange(sr, 7, numRows, 8).setNumberFormat("$#,##0.00");

    // O..Q => currency
    sh.getRange(sr, 15, numRows, 3).setNumberFormat("$#,##0.00");

    // R => text
    sh.getRange(sr, 18, numRows, 1).setNumberFormat("@");
  }
}

// ---------------------
// 4) BALANCE MANAGER
// ---------------------
class BalanceManager {
  constructor(sheet) {
    this.sheet = sheet;
    this.cfg = SHEET_CONFIG;
  }

  recalcAll() {
  const params = getAllInputs(this.sheet);

  // 1) Read entire schedule area (B..R)
  const range = this.sheet.getRange(
    this.cfg.START_ROW,
    2, // Column B
    this.cfg.END_ROW - this.cfg.START_ROW + 1,
    17
  );
  let allRows = range.getValues();

  // 2) Determine how many rows are “in use”
  let lastUsedRowIndex = 0;
  for (let i = 0; i < allRows.length; i++) {
    const periodVal = allRows[i][0]; // col B in array
    if (periodVal === "" && periodVal !== 0) break;
    lastUsedRowIndex++;
  }
  if (lastUsedRowIndex === 0) return;

  // 3) Separate “scheduled” rows vs. “unscheduled”
  const scheduledRows = [];
  const unscheduledRows = [];
  for (let i = 0; i < lastUsedRowIndex; i++) {
    const rowArr    = allRows[i];
    const periodVal = rowArr[0];  // B
    const periodEnd = rowArr[1];  // C
    const paidOnVal = rowArr[4];  // F

    const isScheduled = (
      Number.isInteger(periodVal) &&
      periodEnd instanceof Date &&
      !isNaN(periodEnd)
    );
    const isUnscheduled = (
      !Number.isInteger(periodVal) &&
      paidOnVal instanceof Date &&
      !isNaN(paidOnVal)
    );
    if (isScheduled) {
      scheduledRows.push({ rowIndex: i, rowData: rowArr });
    } else if (isUnscheduled) {
      unscheduledRows.push({ rowIndex: i, rowData: rowArr });
    }
  }

  // 4) Sort scheduled rows by PeriodEnd date
  scheduledRows.sort((a,b) => {
    const dateA = a.rowData[1]; // col C
    const dateB = b.rowData[1];
    return dateA - dateB;
  });

  // 5) Sort unscheduled rows by PaidOn date
  unscheduledRows.sort((a,b) => {
    const dateA = a.rowData[4]; // col F
    const dateB = b.rowData[4];
    return dateA - dateB;
  });

  // 6) Build IPMT/PPMT results for monthly amortizing
  const [ipmtVals, ppmtVals] = this.buildIpmtPpmtResults(
    allRows,
    lastUsedRowIndex,
    params
  );
  // Map them by rowIndex
  const ipmtMap = {};
  const ppmtMap = {};
  for (let i = 0; i < lastUsedRowIndex; i++) {
    ipmtMap[i] = ipmtVals[i][0] || 0;
    ppmtMap[i] = ppmtVals[i][0] || 0;
  }

  // Track which rows have been re‐amortized
  const hasReAmortized = new Array(lastUsedRowIndex).fill(false);

  // 7) Running balances
  let runningPrincipal = params.principal;
  let runningInterest  = 0;
  let runningFees      = 0;

  // 8) Iterate over scheduled rows in ascending date order
  let unschedIndex = 0;
  let lastEndDate  = params.closingDate;

  const isSinglePeriod   = (params.paymentFreq === "Single Period");
  const isMonthly        = (params.paymentFreq === "Monthly");
  const isPeriodicMethod = (params.dayCountMethod === "Periodic");
  const dailyRate        = params.perDiemRate; // = annualRate / daysPerYear
   // Example: for "Periodic 360", we treat each month as 30 days:
   // monthlyInterestFactor = annualRate * (30/daysPerYear) or simply (annualRate / 12)
  const monthlyInterestFactor = (params.daysPerYear)
    ? (params.annualRate * 30 / params.daysPerYear)
    : params.monthlyRate;
  // The 'dailyPeriodicRate' is the monthly factor / 30
  const dailyPeriodicRate = monthlyInterestFactor / 30;

  scheduledRows.forEach(schObj => {
    const rowIndex = schObj.rowIndex;
    const rowArr   = schObj.rowData;
    const periodNum = rowArr[0];   // col B
    const periodEnd = rowArr[1];   // col C
    //New code Below: /////////////////////////////////////////////////////////////////////////////////////////
    let totalActualDaysThisPeriod = 30;
    let monthly30DaysUsedSoFar    = 0;

    if (
      isPeriodicMethod &&
      isMonthly &&
      Number.isInteger(periodNum) &&
      periodNum >= 1
    ) {
      // For a standard monthly row (periodNum >= 1), figure out how many
      // *calendar* days are in this entire scheduled period (for scaling to 30).
      //
      // Usually your code sets subStart = oneDayAfter(lastEndDate), but
      // lastEndDate is the prior row’s C‐col. So we define the same here:
      const scheduledPeriodStart = oneDayAfter(lastEndDate);

      let rawDays = calcUnpaidDays(scheduledPeriodStart, periodEnd, params.prepaidUntil);
      if (rawDays < 1) {
        rawDays = 30; // fallback to avoid 0 or negative
      }
      totalActualDaysThisPeriod = rawDays;
    }
    //New code above /////////////////////////////////////////////////////////////////////////////////////////
    // We'll accumulate interest for sub-intervals in this variable
    let interestAccruedThisRow = 0;
    // Track how many "days" we've used so far in this monthly row:
    let daysUsedSoFarThisRow = 0;

    // Sub‐interval from (lastEndDate+1) up to periodEnd, 
    // accounting for any unscheduled payments
    let subStart;
      if (Number.isInteger(periodNum) && periodNum === 0 && params.prorateFirst === "Yes") {
        // For the very first scheduled row (Period 0 + prorate=Yes),
        // start from the actual closingDate instead of closingDate+1.
        subStart = lastEndDate;
      } else {
        subStart = oneDayAfter(lastEndDate);
      }
      if (
        isPeriodicMethod && 
        isMonthly && 
        Number.isInteger(periodNum) && 
        periodNum >= 1
      ) {
        // Only snap subStart to lastEndDate if subStart has not 
        // already advanced beyond it (due to unscheduled payments).
        if (subStart < lastEndDate) {
          subStart = lastEndDate;
        }
      }

    while (
      unschedIndex < unscheduledRows.length &&
      unscheduledRows[unschedIndex].rowData[4] /* col F */ <= periodEnd
    ) {
      const uObj    = unscheduledRows[unschedIndex];
      const uRowArr = uObj.rowData;
      const paidOn  = uRowArr[4]; // col F

      if (
        isPeriodicMethod &&
        isMonthly &&
        Number.isInteger(periodNum) &&
        periodNum >= 1
      ) {
        if (paidOn >= subStart && runningPrincipal > 1e-6) {
          // Actual calendar days from subStart..paidOn
          let actualSubDays = calcUnpaidDays(subStart, paidOn, params.prepaidUntil);
          if (actualSubDays < 0) actualSubDays = 0;

          // Fraction of the entire month
          const fractionOfMonth = actualSubDays / totalActualDaysThisPeriod;

          // Scale to 30 days
          let scaledSubDays = 30 * fractionOfMonth;

          // Don’t exceed leftover in the 30-day month
          const leftover = 30 - monthly30DaysUsedSoFar;
          if (scaledSubDays > leftover) scaledSubDays = leftover;
          if (scaledSubDays < 0) scaledSubDays = 0;

          // Accrue interest
          if (scaledSubDays > 0) {
            const iNow = runningPrincipal * dailyPeriodicRate * scaledSubDays;
            runningInterest        += iNow;
            interestAccruedThisRow += iNow;
            monthly30DaysUsedSoFar += scaledSubDays;
          }
        }

        // Now subtract the user’s principal/interest/fees in this unscheduled row as normal:
        const principalPdU = uRowArr[8]  || 0; // col J
        const interestPdU  = uRowArr[10] || 0; // col L
        const feesDueU     = uRowArr[11] || 0; // col M
        const feesPdU      = uRowArr[12] || 0; // col N

        runningFees       += feesDueU;
        runningInterest    = Math.max(0, runningInterest - interestPdU);
        runningPrincipal   = Math.max(0, runningPrincipal - principalPdU);
        runningFees        = Math.max(0, runningFees - feesPdU);

        // Refresh G/H in the unscheduled row, etc. (same as your existing code):
        const oldPrinDueU = uRowArr[7] || 0;
        const oldIntDueU  = uRowArr[9] || 0;
        const totalDueU   = oldPrinDueU + oldIntDueU + feesDueU;
        const totalPdU    = principalPdU + interestPdU + feesPdU;
        uRowArr[5] = totalDueU; // G
        uRowArr[6] = totalPdU;  // H

        // Balances:
        uRowArr[13] = runningInterest; 
        uRowArr[14] = runningPrincipal;
        uRowArr[15] = runningInterest + runningPrincipal + runningFees;

        // Bump subStart to the day after the unscheduled payment:
        subStart = oneDayAfter(paidOn);
      } 
      else {
        // "Actual" daily or SinglePeriod or periodNum=0
        if (paidOn >= subStart && runningPrincipal > 1e-6) {
          const partialDays = calcUnpaidDays(subStart, paidOn, params.prepaidUntil);
          if (partialDays > 0) {
            const iNow = runningPrincipal * dailyRate * partialDays;
            runningInterest += iNow;
          }
        }

        // Subtract user’s payments
        const principalPdU = uRowArr[8]  || 0; // col J
        const interestPdU  = uRowArr[10] || 0; // col L
        const feesDueU     = uRowArr[11] || 0; // col M
        const feesPdU      = uRowArr[12] || 0; // col N

        runningFees += feesDueU;
        runningInterest  = Math.max(0, runningInterest  - interestPdU);
        runningPrincipal = Math.max(0, runningPrincipal - principalPdU);
        runningFees      = Math.max(0, runningFees      - feesPdU);

        // Update G,H for unsched row
        const oldPrinDueU = uRowArr[7] || 0;
        const oldIntDueU  = uRowArr[9] || 0;
        const totalDueU   = oldPrinDueU + oldIntDueU + feesDueU;
        const totalPdU    = principalPdU + interestPdU + feesPdU;
        uRowArr[5] = totalDueU;
        uRowArr[6] = totalPdU;

        // Balances
        uRowArr[13] = runningInterest;
        uRowArr[14] = runningPrincipal;
        uRowArr[15] = runningInterest + runningPrincipal + runningFees;

        // Advance subStart
        subStart = oneDayAfter(paidOn);
      }

      unschedIndex++;
    }

    // Now handle the scheduled row’s interest/principal
    let newInterestDue  = 0;
    let newPrincipalDue = 0;

    // If principal > 0, subStart <= periodEnd => add the final chunk for this row
    if (runningPrincipal > 1e-6 && subStart <= periodEnd) {
      if (
        isPeriodicMethod &&
        isMonthly &&
        Number.isInteger(periodNum) &&
        periodNum >= 1
      ) {
        let actualSubDays = calcUnpaidDays(subStart, periodEnd, params.prepaidUntil);
        if (actualSubDays < 0) actualSubDays = 0;

        const fractionOfMonth = actualSubDays / totalActualDaysThisPeriod;
        let scaledSubDays     = 30 * fractionOfMonth;

        const leftover = 30 - monthly30DaysUsedSoFar;
        if (scaledSubDays > leftover) scaledSubDays = leftover;
        if (scaledSubDays < 0) scaledSubDays = 0;

        if (scaledSubDays > 0 && runningPrincipal > 1e-6) {
          const i2 = runningPrincipal * dailyPeriodicRate * scaledSubDays;
          runningInterest += i2;
          newInterestDue  += i2;
          monthly30DaysUsedSoFar += scaledSubDays;
        }
      } else {
        // "Actual" daily or SinglePeriod or periodNum=0 => partial day approach
        const partialDays = calcUnpaidDays(subStart, periodEnd, params.prepaidUntil);
        if (partialDays > 0 && runningPrincipal > 1e-6) {
          const iNow = runningPrincipal * dailyRate * partialDays;
          runningInterest += iNow;
          newInterestDue  += iNow;
        }
      }
    }

    // SINGLE PERIOD => only the final row is "Due"
    if (isSinglePeriod) {
      const isFinal = (Number.isInteger(periodNum) && periodNum === params.termMonths);
      if (isFinal) {
        newInterestDue  = runningInterest;
        newPrincipalDue = runningPrincipal;
      } else {
        // For intermediate rows, show 0 "due," so interest just accumulates in col O
        newInterestDue  = 0;
        newPrincipalDue = 0;
      }
    }
    // MONTHLY with amortize=Yes => override with IPMT/PPMT if within 1..termMonths
    else if (
      isMonthly &&
      params.amortizeYN === "Yes" &&
      Number.isInteger(periodNum) &&
      periodNum >= 1 &&
      periodNum <= params.termMonths
    ) {
      // If not re‐amortized
      if (!hasReAmortized[rowIndex]) {
        // Subtract the newlyAccrued from runningInterest, then add IPMT
        runningInterest = Math.max(0, runningInterest - newInterestDue);

        let iVal = ipmtMap[rowIndex] || 0;
        runningInterest += iVal;
        newInterestDue   = iVal;

        let pVal = ppmtMap[rowIndex] || 0;
        newPrincipalDue  = pVal;
      } else {
        // Already re‐amortized => take row's I/K
        runningInterest = Math.max(0, runningInterest - newInterestDue);
        newInterestDue  = rowArr[9] || 0;
        newPrincipalDue = rowArr[7] || 0;
        runningInterest += newInterestDue;
      }
    }
    // MONTHLY interest‐only => final row gets leftover interest/principal
    else if (
      isMonthly &&
      params.amortizeYN === "No" &&
      Number.isInteger(periodNum) &&
      periodNum === params.termMonths
    ) {
      newInterestDue  = runningInterest;
      newPrincipalDue = runningPrincipal;
    }

    // Write these due amounts into row I,K
    rowArr[7] = newPrincipalDue; // col I
    rowArr[9] = newInterestDue;  // col K

    // Add typed fees M to runningFees
    const thisFeesDue = rowArr[11] || 0;
    runningFees += thisFeesDue;

    // G = principalDue + interestDue + feesDue
    rowArr[5] = newPrincipalDue + newInterestDue + thisFeesDue;

    // Subtract user payments in this scheduled row
    const principalPd = rowArr[8]  || 0; // col J
    const interestPd  = rowArr[10] || 0; // col L
    const feesPd      = rowArr[12] || 0; // col N

    runningInterest  = Math.max(0, runningInterest  - interestPd);
    runningPrincipal = Math.max(0, runningPrincipal - principalPd);
    runningFees      = Math.max(0, runningFees      - feesPd);

    // H = totalPaid
    rowArr[6] = principalPd + interestPd + feesPd;

    // Re‐amortize future if needed
    if (
      isMonthly &&
      !isSinglePeriod &&
      !isUnscheduledRow(rowArr) &&
      params.amortizeYN === "Yes" &&
      principalPd > 0
    ) {
      const leftoverPrincipal = runningPrincipal;
      const leftoverStart     = rowIndex + 1;
      let lastFutureRow       = leftoverStart;
      while (lastFutureRow < lastUsedRowIndex) {
        const pVal = allRows[lastFutureRow][0]; // col B
        if (!pVal && pVal !== 0) break; // blank => end
        lastFutureRow++;
      }
      const leftoverCount = params.termMonths - Math.floor(periodNum);
      if (leftoverCount > 0) {
        this.reAmortizeFutureRows(
          allRows,
          leftoverStart,
          lastFutureRow,
          leftoverCount,
          leftoverPrincipal,
          params,
          hasReAmortized
        );
      }
    }

    // Write final O,P,Q
    rowArr[13] = runningInterest;
    rowArr[14] = runningPrincipal;
    rowArr[15] = runningInterest + runningPrincipal + runningFees;

    // Done with this scheduled row
    schObj.rowData = rowArr;
    lastEndDate    = periodEnd;
  });

  // 9) (Optional) Handle leftover unscheduled after final row, if desired.

  // 10) Bulk write updated data back
  for (let s = 0; s < scheduledRows.length; s++) {
    const ix = scheduledRows[s].rowIndex;
    allRows[ix] = scheduledRows[s].rowData;
  }
  for (let u = 0; u < unscheduledRows.length; u++) {
    const ix = unscheduledRows[u].rowIndex;
    allRows[ix] = unscheduledRows[u].rowData;
  }

  range
    .offset(0, 0, lastUsedRowIndex, 17)
    .setValues(allRows.slice(0, lastUsedRowIndex));

  SpreadsheetApp.flush();
}

  buildIpmtPpmtResults(schedule, lastUsedCount, params) {
    if (params.amortizeYN === "No") {
      const empties = new Array(lastUsedCount).fill([0]);
      return [empties, empties];
    }
    const ipmtFormulas = [];
    const ppmtFormulas = [];
    const monthlyRate  = params.monthlyRate;
    const nper         = params.termMonths;
    const principal    = params.principal;

    for (let r = 0; r < lastUsedCount; r++) {
      const periodNum = schedule[r][0];
      if (
        Number.isInteger(periodNum) &&
        params.paymentFreq === "Monthly" &&
        periodNum >= 1 &&
        periodNum <= nper
      ) {
        ipmtFormulas.push([`=-IPMT(${monthlyRate},${periodNum},${nper},${principal})`]);
        ppmtFormulas.push([`=-PPMT(${monthlyRate},${periodNum},${nper},${principal})`]);
      } else {
        ipmtFormulas.push([""]);
        ppmtFormulas.push([""]);
      }
    }

    const sheet = this.sheet;
    const ipmtRange = sheet.getRange(SHEET_CONFIG.START_ROW, SHEET_CONFIG.IPMT_COL, lastUsedCount, 1);
    const ppmtRange = sheet.getRange(SHEET_CONFIG.START_ROW, SHEET_CONFIG.PPMT_COL, lastUsedCount, 1);
    ipmtRange.setValues(ipmtFormulas);
    ppmtRange.setValues(ppmtFormulas);
    SpreadsheetApp.flush();

    // read computed values
    const ipmtVals = ipmtRange.getValues();
    const ppmtVals = ppmtRange.getValues();

    // clear the formulas
    ipmtRange.clearContent();
    ppmtRange.clearContent();
    SpreadsheetApp.flush();

    return [ipmtVals, ppmtVals];
  }

  reAmortizeFutureRows(schedule, startRow, endRow, leftoverCount, leftoverPrincipal, params, hasReAmortized) {
    if (leftoverCount <= 0 || leftoverPrincipal <= 0.000001) return;

    const monthlyRate = params.monthlyRate;
    let newIpmtFormulas = [];
    let newPpmtFormulas = [];
    let periodIndex = 1;

    for (let r = startRow; r < endRow; r++) {
      const pNum = schedule[r][0];
      if (leftoverPrincipal <= 0.000001) {
        newIpmtFormulas.push([""]);
        newPpmtFormulas.push([""]);
        continue;
      }
      if (Number.isInteger(pNum)) {
        if (periodIndex <= leftoverCount) {
          newIpmtFormulas.push([
            `=-IPMT(${monthlyRate},${periodIndex},${leftoverCount},${leftoverPrincipal})`
          ]);
          newPpmtFormulas.push([
            `=-PPMT(${monthlyRate},${periodIndex},${leftoverCount},${leftoverPrincipal})`
          ]);
          periodIndex++;
        } else {
          newIpmtFormulas.push([""]);
          newPpmtFormulas.push([""]);
        }
      } else {
        newIpmtFormulas.push([""]);
        newPpmtFormulas.push([""]);
      }
    }

    const sheet = this.sheet;
    const ipmtRange = sheet.getRange(SHEET_CONFIG.START_ROW + startRow, SHEET_CONFIG.IPMT_COL, endRow - startRow, 1);
    const ppmtRange = sheet.getRange(SHEET_CONFIG.START_ROW + startRow, SHEET_CONFIG.PPMT_COL, endRow - startRow, 1);
    ipmtRange.setValues(newIpmtFormulas);
    ppmtRange.setValues(newPpmtFormulas);
    SpreadsheetApp.flush();

    const newIpmtVals = ipmtRange.getValues();
    const newPpmtVals = ppmtRange.getValues();
    ipmtRange.clearContent();
    ppmtRange.clearContent();
    SpreadsheetApp.flush();

    periodIndex = 1;
    for (let r = startRow; r < endRow; r++) {
      const rowArr = schedule[r];
      const pNum = rowArr[0];
      if (Number.isInteger(pNum)) {
        if (periodIndex <= leftoverCount && leftoverPrincipal > 0.000001) {
          const iVal = newIpmtVals[r - startRow][0] || 0;
          const pVal = newPpmtVals[r - startRow][0] || 0;
          const feesDue = rowArr[11] || 0;

          rowArr[7] = pVal;                      // I
          rowArr[9] = iVal;                      // K
          rowArr[5] = pVal + iVal + feesDue;     // G

          hasReAmortized[r] = true;
          periodIndex++;
        } else {
          rowArr[7] = 0; // I
          rowArr[9] = 0; // K
          rowArr[5] = 0; // G
          hasReAmortized[r] = true;
        }
      }
    }
  }
}

// ---------------------
// 5) ROW MANAGER
// ---------------------
class RowManager {
  constructor(sheet) {
    this.sheet = sheet;
  }

  handleInsertedRow(insertedRow) {
    const cfg = SHEET_CONFIG.COLUMNS;

    let prevPeriod = null;
    if (insertedRow - 1 >= SHEET_CONFIG.START_ROW) {
      prevPeriod = this.sheet.getRange(insertedRow - 1, cfg.PERIOD).getValue();
    }
    let nextPeriod = null;
    if (insertedRow + 1 <= SHEET_CONFIG.END_ROW) {
      nextPeriod = this.sheet.getRange(insertedRow + 1, cfg.PERIOD).getValue();
    }

    let newPeriod;
    if (nextPeriod != null && !isNaN(nextPeriod) && prevPeriod != null && !isNaN(prevPeriod)) {
      newPeriod = prevPeriod + (nextPeriod - prevPeriod) / 2;
    } else if (prevPeriod != null && !isNaN(prevPeriod)) {
      newPeriod = prevPeriod + 0.5;
    } else {
      newPeriod = 0.5;
    }

    // Copy balances from prior row
    let prevIntBal = 0, prevPrinBal = 0;
    if (insertedRow - 1 >= SHEET_CONFIG.START_ROW) {
      prevIntBal  = this.sheet.getRange(insertedRow - 1, cfg.INT_BAL).getValue() || 0;
      prevPrinBal = this.sheet.getRange(insertedRow - 1, cfg.PRIN_BAL).getValue() || 0;
    }

        const rowValues = [
      newPeriod,                 // B => PERIOD
      "",                        // C => PERIOD_END
      "",                        // D => DUE_DATE
      "",                        // E => DAYS
      "",                        // F => PAID_ON
      0,                         // G => TOTAL_DUE
      0,                         // H => TOTAL_PAID
      "",                        // I => PRINCIPAL_DUE
      0,                         // J => PRINCIPAL_PD
      "",                        // K => INTEREST_DUE
      0,                         // L => INTEREST_PD
      "",                        // M => FEES_DUE
      0,                         // N => FEES_PD
      prevIntBal,                // O => INT_BAL
      prevPrinBal,               // P => PRIN_BAL
      prevIntBal + prevPrinBal,  // Q => TOTAL_BAL
      ""                         // R => NOTES
    ];

    this.sheet.getRange(insertedRow, 2, 1, rowValues.length).setValues([rowValues]);
  }
}

// ---------------------
// 6) UI & TRIGGERS
// ---------------------
function createLoanScheduleMenu(){
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Loan Schedule Tools')
    .addItem('Insert Unscheduled Payment Row', 'insertUnscheduledPaymentRow')
    .addItem('Generate Loan Schedule', 'generateLoanSchedule')
    .addItem('Set Up Triggers', 'setupTriggers')
    .addToUi();
}

function generateLoanSchedule() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet) return;
  if (sheet.getName() === 'Summary') {
    return;
  }
  const gen = new LoanScheduleGenerator(sheet);
  gen.generateSchedule();
}

function insertUnscheduledPaymentRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!sheet) return;

  const sel = sheet.getActiveRange();
  if (!sel) return;

  const r = sel.getRow();
  if (r < SHEET_CONFIG.START_ROW || r > SHEET_CONFIG.END_ROW) {
    SpreadsheetApp.getUi().alert('Please select a row in the schedule area (B8:Q500).');
    return;
  }

  sheet.insertRowAfter(r);
  const rm = new RowManager(sheet);
  rm.handleInsertedRow(r + 1);

  // Then recalc
  const bal = new BalanceManager(sheet);
  bal.recalcAll();
}

function createOnEditTrigger() {
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

/**
 * onEdit trigger: 
 *  - If user edits row 4 (and it's locked), revert.
 *  - If user edits row 4 (and it's not locked), regenerate schedule.
 *  - If user edits any schedule row in columns H/J/L/M/N, recalc balances.
 */
function onEdit(e) {
  try {
    const sheet = e.range.getSheet();

    // Ignore edits made in the 'Summary' sheet.
    if (sheet.getName() === 'Summary') {
      return;
    }

    const r = e.range.getRow();
    const c = e.range.getColumn();

    // "Lock Inputs" cell
    const lockVal = sheet.getRange(SHEET_CONFIG.INPUTS.LOCK_INPUTS).getValue();  // Q4
    const lockCol = sheet.getRange(SHEET_CONFIG.INPUTS.LOCK_INPUTS).getColumn();

    // 1) Edits to row 4 (input row):
    if (r === 4) {
      if (c === lockCol) {
        return; // editing Q4 itself => do nothing
      }
      // Only regenerate if the edit is in columns D..P (i.e., 4..16).
      if (c < 4 || c > 16) {
        return;
      }
      // If locked => revert
      if (lockVal === "Yes") {
        e.range.setValue(e.oldValue);
        SpreadsheetApp.getUi().alert(
          "Inputs are locked. Set Q4 to 'No' to edit these fields."
        );
        return;
      } else {
        // If not locked => regenerate
        generateLoanSchedule();
        return;
      }
    }

    // 2) Edits in schedule area => recalc
    if (r >= SHEET_CONFIG.START_ROW && r <= SHEET_CONFIG.END_ROW) {
      const recalcCols = [
        SHEET_CONFIG.COLUMNS. PAID_ON ,      // F
        SHEET_CONFIG.COLUMNS.TOTAL_PAID,    // H=8
        SHEET_CONFIG.COLUMNS.PRINCIPAL_PD,  // J=10
        SHEET_CONFIG.COLUMNS.INTEREST_PD,   // L=12
        SHEET_CONFIG.COLUMNS.FEES_DUE,      // M=13
        SHEET_CONFIG.COLUMNS.FEES_PD        // N=14
      ];
      if (recalcCols.includes(c)) {
        const bal = new BalanceManager(sheet);
        bal.recalcAll();
      }
    }
  } catch (err) {
    SpreadsheetApp.getUi().alert("onEdit error: " + err.message);
  }
}