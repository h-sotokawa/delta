// Execute the test function
try {
  testSaferNullChecking();
  console.log('\n--- Running printer test ---');
  testPrinterOtherSaferNullChecking();
} catch (error) {
  console.error('Error:', error);
}
