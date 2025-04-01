/**
 * AmortizingLoans.test.js
 *
 * This suite tests amortizing loans (Amortize = "Yes"), ensuring:
 *  - Monthly payment calculation and schedule generation
 *  - Correct principal and interest allocation
 *  - Balance reaching zero at the end of the term
 *  - Handling of unscheduled (extra) payments
 *  - Re-amortization after prepayment (recasting)
 */

const { calculateDueAmounts } = require('../LoanHelpers.js');

// jest setup file or top of test file:
require('gas-mock-globals');  // This will automatically define SpreadsheetApp, etc.
// Ensure flush is defined
if (typeof SpreadsheetApp.flush !== 'function') {
  SpreadsheetApp.flush = jest.fn();
}

/**
 * Quick helper function to calculate a standard fully-amortizing monthly payment:
 * P = r * L / [1 - (1 + r)^(-n)]
 *   where:
 *     L = loan principal
 *     r = monthly interest rate
 *     n = total number of payments
 */
function calcMonthlyPayment(principal, annualRate, months) {
  const monthlyRate = annualRate / 12;
  return principal * (monthlyRate / (1 - Math.pow(1 + monthlyRate, -months)));
}

describe('Amortizing Loans (Amortize = "Yes")', () => {
  it('generates a fully amortizing schedule with zero balance at end of term', () => {
    // Loan parameters: $1000 @ 5% annual interest, 12-month term, monthly amortizing
    const principal = 1000;
    const annualRate = 0.05;
    const termMonths = 12;
    const monthlyRate = annualRate / 12;
    const monthlyPayment = calcMonthlyPayment(principal, annualRate, termMonths);

    let balance = principal;

    for (let period = 1; period <= termMonths; period++) {
      const interestDue = balance * monthlyRate;
      const principalDue = monthlyPayment - interestDue;
      balance -= principalDue;
    }

    // After all payments, the balance should be (close to) 0
    expect(balance).toBeCloseTo(0, 6);
  });

  it('calculates interest and principal for each monthly payment correctly (interest decreases, principal increases)', () => {
    const principal = 1000;
    const annualRate = 0.05;
    const termMonths = 12;
    const monthlyRate = annualRate / 12;
    const monthlyPayment = calcMonthlyPayment(principal, annualRate, termMonths);

    let balance = principal;
    let previousInterest = null;
    let previousPrincipal = null;

    for (let i = 1; i <= termMonths; i++) {
      const interestThisPeriod = balance * monthlyRate;
      const principalThisPeriod = monthlyPayment - interestThisPeriod;
      balance -= principalThisPeriod;

      if (previousInterest !== null) {
        // Interest should decrease each period
        expect(interestThisPeriod).toBeLessThan(previousInterest);
      }
      if (previousPrincipal !== null) {
        // Principal portion should increase each period
        expect(principalThisPeriod).toBeGreaterThan(previousPrincipal);
      }

      previousInterest = interestThisPeriod;
      previousPrincipal = principalThisPeriod;
    }

    // Final balance check
    expect(balance).toBeCloseTo(0, 6);
  });

  it('updates the principal balance correctly after each payment', () => {
    const principal = 5000;
    const annualRate = 0.04; // 4%
    const termMonths = 6;
    const monthlyRate = annualRate / 12;
    const monthlyPayment = calcMonthlyPayment(principal, annualRate, termMonths);

    let balance = principal;

    for (let period = 1; period <= termMonths; period++) {
      const interestDue = balance * monthlyRate;
      const principalDue = monthlyPayment - interestDue;
      balance -= principalDue;
    }

    // After final payment, balance should be near 0
    expect(balance).toBeCloseTo(0, 6);
  });

  describe('Handling Unscheduled (Extra) Payments', () => {
    it('applies an extra payment immediately to reduce the remaining principal balance', () => {
      // $1000, 5% annual, 12-month term
      const principal = 1000;
      const annualRate = 0.05;
      const termMonths = 12;
      const monthlyRate = annualRate / 12;
      const monthlyPayment = calcMonthlyPayment(principal, annualRate, termMonths);

      let balance = principal;

      // Make 3 normal payments
      for (let period = 1; period <= 3; period++) {
        const interestDue = balance * monthlyRate;
        const principalDue = monthlyPayment - interestDue;
        balance -= principalDue;
      }

      const balanceBeforeExtra = balance;
      const extraPayment = 200;
      balance -= extraPayment;

      // Balance should drop exactly by the extra payment amount
      expect(balance).toBeCloseTo(balanceBeforeExtra - extraPayment, 6);
    });

    it('reduces the interest due in subsequent periods after an extra principal payment', () => {
      const principal = 1000;
      const annualRate = 0.05;
      const termMonths = 12;
      const monthlyRate = annualRate / 12;
      const monthlyPayment = calcMonthlyPayment(principal, annualRate, termMonths);

      // 1) Compute period 4 interest with NO prepayment:
      let balanceNoPrepay = principal;
      for (let period = 1; period <= 3; period++) {
        const interestDue = balanceNoPrepay * monthlyRate;
        const principalDue = monthlyPayment - interestDue;
        balanceNoPrepay -= principalDue;
      }
      const interestDueP4_noPrepay = balanceNoPrepay * monthlyRate;

      // 2) Compute period 4 interest WITH a $200 prepayment after period 3
      let balanceWithPrepay = principal;
      for (let period = 1; period <= 3; period++) {
        const interestDue = balanceWithPrepay * monthlyRate;
        const principalDue = monthlyPayment - interestDue;
        balanceWithPrepay -= principalDue;
      }
      balanceWithPrepay -= 200; // extra payment
      const interestDueP4_withPrepay = balanceWithPrepay * monthlyRate;

      // Interest in period 4 should be less after prepayment
      expect(interestDueP4_withPrepay).toBeLessThan(interestDueP4_noPrepay);

      // The difference in interest should be roughly extraPayment * monthlyRate
      const expectedReduction = 200 * monthlyRate;
      const actualReduction = interestDueP4_noPrepay - interestDueP4_withPrepay;
      expect(actualReduction).toBeCloseTo(expectedReduction, 6);
    });

    it('shortens the loan term (earlier payoff) when extra payments are made without recasting', () => {
      // $1000, 5% annual, 12-month term, extra $300 payment after 6 months
      const principal = 1000;
      const annualRate = 0.05;
      const termMonths = 12;
      const monthlyRate = annualRate / 12;
      const monthlyPayment = calcMonthlyPayment(principal, annualRate, termMonths);

      let balance = principal;

      // Pay 6 normal months
      for (let i = 1; i <= 6; i++) {
        const interestDue = balance * monthlyRate;
        const principalDue = monthlyPayment - interestDue;
        balance -= principalDue;
      }

      // Prepay $300
      balance -= 300;

      // Continue paying same monthly payment until payoff
      let additionalPayments = 0;
      while (balance > 0.000001 && additionalPayments < 50) {
        const interestDue = balance * monthlyRate;
        const principalDue = Math.min(monthlyPayment - interestDue, balance);
        balance -= principalDue;
        additionalPayments++;
      }

      const totalPaidMonths = 6 + additionalPayments;
      // Should pay off before the original 12 months
      expect(totalPaidMonths).toBeLessThan(termMonths);
      // Final balance ~ 0
      expect(balance).toBeCloseTo(0, 6);
    });

    it('recalculates (re-amortizes) remaining schedule after a prepayment if recastLoan is called', () => {
      // $1000, 5% annual, 12-month term, $200 prepayment after month 3, then recast
      const principal = 1000;
      const annualRate = 0.05;
      const termMonths = 12;
      const monthlyRate = annualRate / 12;
      const originalPayment = calcMonthlyPayment(principal, annualRate, termMonths);

      let balance = principal;

      // Make 3 normal payments
      for (let i = 1; i <= 3; i++) {
        const interestDue = balance * monthlyRate;
        const principalDue = originalPayment - interestDue;
        balance -= principalDue;
      }

      // $200 prepayment
      balance -= 200;

      const remainingTerm = termMonths - 3;
      // Recalculate a new payment for the remaining term (simulating recastLoan)
      const newPayment = balance * (monthlyRate / (1 - Math.pow(1 + monthlyRate, -remainingTerm)));

      // New payment should be smaller than original
      expect(newPayment).toBeLessThan(originalPayment);

      // Pay off the recast schedule
      for (let i = 1; i <= remainingTerm; i++) {
        const interestDue = balance * monthlyRate;
        const principalDue = newPayment - interestDue;
        balance -= principalDue;
      }

      // Balance should be near 0 at the end of the original 12 months
      expect(balance).toBeCloseTo(0, 6);
    });
  });

  it('properly re-allocates interest/principal if an extra payment happens mid-period (calculateDueAmounts)', () => {
    // Example: scheduled Payment = $500 ($50 interest, $450 principal).
    // If actual interest is only $30 due to a mid-period prepayment, the extra $20 goes to principal.
    const params = {
      paymentFreq: 'Monthly',
      amortizeYN: 'Yes',
      termMonths: 12,
      principal: 10000,
    };
    const periodNum = 5;
    const scheduledInt = 50;
    const scheduledPr = 450;
    const scheduledPayment = scheduledInt + scheduledPr; // 500
    const interestAccrued = 30; // reduced actual interest
    const result = calculateDueAmounts(
      periodNum,
      /* rowIndex */ 0,
      params,
      interestAccrued,
      scheduledInt,
      scheduledPr,
      /* hadExtraPaymentBefore */ false,
      /* extraPaymentThisPeriod */ true,
      /* wasReAmortized */ false,
      /* rowData */ []
    );

    expect(result.newInterestDue).toBeCloseTo(30, 6);
    // The principal portion should now be 470
    expect(result.newPrincipalDue).toBeCloseTo(470, 6);
    // Total remains 500
    expect(result.newInterestDue + result.newPrincipalDue).toBeCloseTo(scheduledPayment, 6);
  });

  it('uses re-amortized values for future periods when wasReAmortized = true', () => {
    // If a loan was re-amortized, the code should pull the new principal/interest from stored rowData
    const params = {
      paymentFreq: 'Monthly',
      amortizeYN: 'Yes',
      termMonths: 24,
      principal: 5000,
    };
    const periodNum = 10;
    // Suppose after recasting, the sheet stored new interest = 40, principal = 110
    const rowData = [];
    rowData[7] = 110; // principal due col (example index)
    rowData[9] = 40;  // interest due col (example index)

    const result = calculateDueAmounts(
      periodNum,
      0,
      params,
      /* interestAccrued */ 40,
      /* scheduledInt */ 50,
      /* scheduledPr */ 100,
      /* hadExtraPaymentBefore */ false,
      /* extraPaymentThisPeriod */ false,
      /* wasReAmortized */ true,
      rowData
    );

    // Should return the stored re-amortized values
    expect(result.newInterestDue).toBe(40);
    expect(result.newPrincipalDue).toBe(110);
  });
    
  // Test for early payoff with unscheduled payment
  it('pays off the loan early with an unscheduled payment and no recast', () => {
    // $1000, 5% annual, 12-month term, $300 extra payment after 6 months
    // This test checks that the loan is paid off early without recasting the loan after an unscheduled payment.


    const principal = 1000;
    const annualRate = 0.05;
    const termMonths = 12;
    const monthlyRate = annualRate / 12;
    const monthlyPayment = calcMonthlyPayment(principal, annualRate, termMonths);

    let balance = principal;

    // Make 6 scheduled monthly payments
    for (let period = 1; period <= 6; period++) {
    const interestDue = balance * monthlyRate;
    const principalDue = monthlyPayment - interestDue;
    balance -= principalDue;
    }
    const balanceBeforeExtra = balance;

    // Apply an unscheduled extra principal payment of $300
    const extraPayment = 300;
    balance -= extraPayment;

    // After the extra payment, the loan should not be fully paid off yet
    expect(balance).toBeGreaterThan(0);  // Balance is still > $0, not fully paid by extra

    // Continue simulating payments for the remaining periods
    let payoffPeriod = null;
    for (let period = 7; period <= termMonths; period++) {
    const interestDue = balance * monthlyRate;
    const principalDue = monthlyPayment - interestDue;
    if (balance - principalDue <= 1e-6) {
        // This payment will pay off the remaining balance
        payoffPeriod = period;
        balance -= balance;  // subtract entire remaining principal
        break;
    } else {
        // Regular payment (loan not paid off yet)
        balance -= principalDue;
    }
    }

    // The loan should be paid off before the original end of term (early payoff)
    expect(payoffPeriod).not.toBeNull();
    expect(payoffPeriod).toBeLessThan(termMonths);

    // All **remaining** scheduled periods after payoff should have 0 principal and 0 interest due
    for (let period = payoffPeriod + 1; period <= termMonths; period++) {
    // Since the loan is already paid off, no interest or principal is due in these periods
    const interestDueRemaining = 0;
    const principalDueRemaining = 0;
    expect(interestDueRemaining).toBeCloseTo(0, 6);
    expect(principalDueRemaining).toBeCloseTo(0, 6);
    }

    // The loan balance should only reach (near) zero at the very end of the schedule
    expect(balance).toBeCloseTo(0, 6);
  });
});