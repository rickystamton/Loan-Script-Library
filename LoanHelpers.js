// LoanHelpers.js – Helper functions for loan schedule calculations (interest, principal, prepayments)

// Calculate inclusive day count between two dates, minus any prepaid interest period.
function daysBetweenInclusive(startDate, endDate) {
  const msPerDay = 24 * 60 * 60 * 1000;
  const start = new Date(startDate);
  const end = new Date(endDate);
  // Normalize times to midnight:
  start.setHours(0,0,0,0);
  end.setHours(0,0,0,0);
  if (end < start) return 0;
  const diffDays = Math.floor((end - start) / msPerDay);
  return diffDays + 1;
}
function computeAccrualDays(periodStart, periodEnd, prepaidUntil) {
  if (!(periodEnd instanceof Date) || isNaN(periodEnd)) {
    return 0;
  }
  const rawDays = daysBetweenInclusive(periodStart, periodEnd);
  if (!prepaidUntil) return rawDays;
  // If interest is prepaid through a certain date, exclude that portion
  if (periodEnd < prepaidUntil) return 0;
  if (periodStart >= prepaidUntil) return rawDays;
  const afterPrepaid = new Date(prepaidUntil);
  return daysBetweenInclusive(afterPrepaid, periodEnd);
}

/**
 * Separate scheduled and unscheduled rows from the full schedule.
 * @param {any[][]} allRows - The full schedule values (array of rows).
 * @param {number} count - Number of used rows in the schedule.
 * @returns {{ scheduledRows: Array, unscheduledRows: Array }}
 */
function separateRows(allRows, count) {
  const scheduledRows = [];
  const unscheduledRows = [];
  for (let i = 0; i < count; i++) {
    const row = allRows[i];
    const periodVal = row[0];       // Period (col B)
    const periodEnd = row[1];       // Period End Date (col C)
    const paidOnVal = row[4];       // Paid On Date (col F)
    const isScheduled = Number.isInteger(periodVal) && periodEnd instanceof Date && !isNaN(periodEnd);
    const isUnscheduled = (!Number.isInteger(periodVal)) && paidOnVal instanceof Date && !isNaN(paidOnVal);
    if (isScheduled) {
      scheduledRows.push({ rowIndex: i, rowData: row });
    } else if (isUnscheduled) {
      unscheduledRows.push({ rowIndex: i, rowData: row });
    }
  }
  // Sort scheduled by due date, unscheduled by paid date (chronologically)
  scheduledRows.sort((a, b) => a.rowData[1] - b.rowData[1]);
  unscheduledRows.sort((a, b) => a.rowData[4] - b.rowData[4]);
  return { scheduledRows, unscheduledRows };
}

/**
 * Process all unscheduled payments that occur on or before the given period's end date.
 * Accrues interest up to each unscheduled payment, applies the payment to balances, and returns updated balances and totals.
 * @param {number} periodNum – The period number (if scheduled period) or identifier.
 * @param {Date} periodStart – The start date of the period (for interest accrual).
 * @param {Date} periodEnd – The end date of the scheduled period.
 * @param {Object} params – Loan parameters (including dayCountMethod, paymentFreq, annualRate, monthlyRate, perDiemRate, prepaidUntil, daysPerYear).
 * @param {Array} unscheduledRows – Array of unscheduled payment row objects (with rowData).
 * @param {number} startUnschedIndex – Index in unscheduledRows to start processing from.
 * @param {number} runningPrincipal – Current remaining principal balance at period start.
 * @param {number} runningInterest – Current accrued interest balance at period start.
 * @param {number} runningFees – Current accrued fees balance at period start.
 * @returns {{ runningPrincipal: number, runningInterest: number, runningFees: number, unschedIndex: number, interestAccrued: number, unscheduledPrincipalPaid: number }}
 */
function applyUnscheduledPaymentsForPeriod(periodNum, periodStart, periodEnd, params, unscheduledRows, startUnschedIndex, runningPrincipal, runningInterest, runningFees) {
  let unschedIndex = startUnschedIndex;
  let interestAccrued = 0;
  let unscheduledPrincipalPaid = 0;
  // Flags for interest calculation method:
  const isMonthly = (params.paymentFreq === "Monthly");
  const isPeriodic = (params.dayCountMethod === "Periodic");
  // Interest rate factors:
  const dailyRate = params.perDiemRate;  // actual daily interest rate
  let dailyPeriodicRate = dailyRate;
  if (isPeriodic && isMonthly) {
    // 30/360 method: derive daily periodic rate from annualRate & daysPerYear
    const monthlyInterestFactor = params.daysPerYear 
      ? (params.annualRate * 30 / params.daysPerYear) 
      : params.monthlyRate;
    dailyPeriodicRate = monthlyInterestFactor / 30;
  }
  // Determine total days in this period (for 30/360 calculations)
  let totalActualDays = 30;
  if (isPeriodic && isMonthly && Number.isInteger(periodNum) && periodNum >= 1) {
    let rawDays = computeAccrualDays(periodStart, periodEnd, params.prepaidUntil);
    if (rawDays < 1) rawDays = 30;
    totalActualDays = rawDays;
  }
  let monthly30DaysUsed = 0;
  let subStart = new Date(periodStart);  // starting point for interest accrual within the period

  // Process each unscheduled payment up to periodEnd
  while (
    unschedIndex < unscheduledRows.length &&
    unscheduledRows[unschedIndex].rowData[4] <= periodEnd
  ) {
    const uRow = unscheduledRows[unschedIndex];
    const paidOn = uRow.rowData[4]; // Paid On date of unscheduled payment
    // Accrue interest from subStart up to the unscheduled payment date (for 30/360 partial period interest)
    if (isPeriodic && isMonthly && Number.isInteger(periodNum) && periodNum >= 1) {
      if (paidOn >= subStart && runningPrincipal > 1e-6) {
        const actualSubDays = computeAccrualDays(subStart, paidOn, params.prepaidUntil);
        const fractionalMonth = (actualSubDays > 0 ? actualSubDays : 0) / totalActualDays;
        let scaledDays = 30 * fractionalMonth;
        const remainingDays = 30 - monthly30DaysUsed;
        if (scaledDays > remainingDays) scaledDays = remainingDays;
        if (scaledDays < 0) scaledDays = 0;
        if (scaledDays > 0) {
          const interestPortion = runningPrincipal * dailyPeriodicRate * scaledDays;
          runningInterest += interestPortion;
          interestAccrued += interestPortion;
          monthly30DaysUsed += scaledDays;
        }
      }
    } else {
      // For actual day-count conventions, interest from subStart to paidOn will be accrued in final step below.
      // (No intermediate accrual here to avoid double-counting in actual/365 mode.)
    }

    // Apply the unscheduled payment amounts to balances
    const principalPaidU = uRow.rowData[8] || 0;  // col J: Principal Paid (unscheduled row)
    const interestPaidU = uRow.rowData[10] || 0;  // col L: Interest Paid
    const feesDueU = uRow.rowData[11] || 0;       // col M: Fees Due (if any)
    const feesPaidU = uRow.rowData[12] || 0;      // col N: Fees Paid
    runningFees += feesDueU;
    runningInterest = Math.max(0, runningInterest - interestPaidU);
    runningPrincipal = Math.max(0, runningPrincipal - principalPaidU);
    runningFees = Math.max(0, runningFees - feesPaidU);
    if (principalPaidU > 0) {
      unscheduledPrincipalPaid += principalPaidU;
    }

    // Update the unscheduled row's totals and balance columns (H, O, P, Q)
    uRow.rowData[6] = principalPaidU + interestPaidU + feesPaidU;    // col H: Total Paid for unscheduled row
    uRow.rowData[13] = runningInterest;                             // col O: Interest balance after payment
    uRow.rowData[14] = runningPrincipal;                            // col P: Principal balance after payment
    uRow.rowData[15] = runningInterest + runningPrincipal + runningFees; // col Q: Total balance after payment

    // Move subStart to the day after this unscheduled payment
    subStart = new Date(paidOn);
    subStart.setDate(subStart.getDate() + 1);
    unschedIndex++;
  }

  // Accrue interest from the last subStart (after final unscheduled payment or period start) up to periodEnd
  if (runningPrincipal > 1e-6 && subStart <= periodEnd) {
    if (isPeriodic && isMonthly && Number.isInteger(periodNum) && periodNum >= 1) {
      const remainingDays = computeAccrualDays(subStart, periodEnd, params.prepaidUntil);
      const fractionalMonth = (remainingDays > 0 ? remainingDays : 0) / totalActualDays;
      let scaledDays = 30 * fractionalMonth;
      const leftover = 30 - monthly30DaysUsed;
      if (scaledDays > leftover) scaledDays = leftover;
      if (scaledDays < 0) scaledDays = 0;
      if (scaledDays > 0) {
        const interestEnd = runningPrincipal * dailyPeriodicRate * scaledDays;
        runningInterest += interestEnd;
        interestAccrued += interestEnd;
        monthly30DaysUsed += scaledDays;
      }
    } else {
      // Actual day-count or single-period: accrue interest for all days from subStart to periodEnd
      const partialDays = computeAccrualDays(subStart, periodEnd, params.prepaidUntil);
      if (partialDays > 0 && runningPrincipal > 1e-6) {
        const interestEnd = runningPrincipal * dailyRate * partialDays;
        runningInterest += interestEnd;
        interestAccrued += interestEnd;
      }
    }
  }

  return {
    runningPrincipal,
    runningInterest,
    runningFees,
    unschedIndex,
    interestAccrued,
    unscheduledPrincipalPaid
  };
}

/**
 * Calculate the Principal Due and Interest Due for a scheduled period, taking into account any extra payments.
 * Uses original amortization amounts for baseline and adjusts if prepayments occurred.
 * @param {number} periodNum – The period number (for scheduled periods).
 * @param {number} rowIndex – The index of the current row in the original schedule array.
 * @param {Object} params – Loan parameters (includes termMonths, amortizeYN, paymentFreq, etc.).
 * @param {number} interestAccrued – The total interest accrued during this period.
 * @param {number} scheduledInt – The originally scheduled interest due for this period (from amortization schedule).
 * @param {number} scheduledPr – The originally scheduled principal due for this period.
 * @param {boolean} hadExtraPaymentBefore – True if a prior period had an extra payment (flag to indicate re-amortization in effect).
 * @param {boolean} extraPaymentThisPeriod – True if an extra principal payment occurred in this same period.
 * @param {boolean} wasReAmortized – True if this period was already re-amortized (if prior reAmortize logic was applied).
 * @param {Array} rowData – Reference to the current row’s data array (to use prior stored due values if re-amortized).
 * @returns {{ newPrincipalDue: number, newInterestDue: number }}
 */
function calculateDueAmounts(periodNum, rowIndex, params, interestAccrued, scheduledInt, scheduledPr, hadExtraPaymentBefore, extraPaymentThisPeriod, wasReAmortized, rowData) {
  let newInterestDue = 0;
  let newPrincipalDue = 0;
  const isSinglePeriod = (params.paymentFreq === "Single Period");
  const isMonthlyAmortizing = (params.paymentFreq === "Monthly" && params.amortizeYN === "Yes");
  const isInterestOnly = (params.paymentFreq === "Monthly" && params.amortizeYN === "No");

  if (isSinglePeriod) {
    // Single lump-sum loan: all interest and principal due at final period only
    const isFinalPeriod = (Number.isInteger(periodNum) && periodNum === params.termMonths);
    if (isFinalPeriod) {
      newInterestDue = interestAccrued;
      newPrincipalDue = params.principal;  // remaining principal due at end
    } else {
      newInterestDue = 0;
      newPrincipalDue = 0;
    }
  } else if (isMonthlyAmortizing && Number.isInteger(periodNum) && periodNum >= 1 && periodNum <= params.termMonths) {
    // Amortizing monthly loan:
    if (extraPaymentThisPeriod) {
      // Extra principal was paid this period – reduce interest due for this period
      // Remove all accrued interest (it will be recalculated as actualInterest)
      newInterestDue = interestAccrued;
      // Cap interest due to not exceed scheduled payment (edge case where accrued > payment)
      const scheduledPayment = scheduledInt + scheduledPr;
      if (newInterestDue > scheduledPayment) {
        newInterestDue = scheduledPayment;
      }
      newPrincipalDue = scheduledPayment - newInterestDue;
      // Mark that an extra payment has occurred (in BalanceManager, extraPaidOccurred will be set)
    } else if (hadExtraPaymentBefore) {
      // A previous period had an extra payment, but none this period – use adjusted amortization
      // (Recalculate current period’s split based on actual accrued interest vs original payment)
      newInterestDue = interestAccrued;
      const scheduledPayment = scheduledInt + scheduledPr;
      if (newInterestDue > scheduledPayment) {
        newInterestDue = scheduledPayment;
      }
      newPrincipalDue = scheduledPayment - newInterestDue;
    } else if (!wasReAmortized) {
      // No prepayment impact – use original amortization values
      newInterestDue = scheduledInt;
      newPrincipalDue = scheduledPr;
    } else {
      // Already re-amortized by a prior prepayment (this scenario would apply if reAmortizeFutureRows was used)
      newInterestDue = rowData[9] || 0;  // use the last stored Interest Due
      newPrincipalDue = rowData[7] || 0; // use the last stored Principal Due
    }
  } else if (isInterestOnly && Number.isInteger(periodNum)) {
    // Interest-only loan logic:
    if (periodNum === params.termMonths) {
      // Final period of interest-only: all remaining interest and principal due
      newInterestDue = interestAccrued;
      newPrincipalDue = params.principal;  // all principal due at maturity
    } else {
      // Intermediate interest-only period: interest due is whatever accrued, principal due is 0
      newInterestDue = interestAccrued;
      newPrincipalDue = 0;
    }
  } else {
    // Default case (for any periods that don’t fall into above categories)
    newInterestDue = interestAccrued;
    newPrincipalDue = 0;
  }

  return { newPrincipalDue, newInterestDue };
}

// Export the helpers for usage in Node tests or import in other scripts
// (In Google Apps Script, these will be available globally once this file is included)
if (typeof module !== 'undefined' && module.exports) {
  // Node.js export
  module.exports = { separateRows, applyUnscheduledPaymentsForPeriod, calculateDueAmounts };
} else {
  // Apps Script: assign functions to a global LoanHelpers object
  if (typeof LoanHelpers === 'undefined') {
    this.LoanHelpers = {};  // `this` refers to global in Apps Script
  }
  LoanHelpers.separateRows = separateRows;
  LoanHelpers.applyUnscheduledPaymentsForPeriod = applyUnscheduledPaymentsForPeriod;
  LoanHelpers.calculateDueAmounts = calculateDueAmounts;
}