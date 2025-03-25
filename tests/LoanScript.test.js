// tests/LoanScript.test.js
const loanScript = require('../LoanScript.js');

test('double() returns the number multiplied by 2', () => {
  // Arrange
  const input = 5;
  const expectedOutput = 10;
  // Act
  const result = loanScript.double(input);
  // Assert
  expect(result).toBe(expectedOutput);
});
