/**
 * Global teardown for Playwright tests
 * Cleans up after OOXML slides validation tests
 */

const fs = require('fs');
const path = require('path');

async function globalTeardown(config) {
  console.log('Cleaning up OOXML Slides Playwright testing environment...');
  
  // Generate test summary
  const resultsDir = path.join(__dirname, '..', 'test-results');
  const summaryFile = path.join(resultsDir, 'test-summary.json');
  
  const summary = {
    timestamp: new Date().toISOString(),
    testType: 'OOXML Slides Validation',
    environment: 'Playwright',
    notes: [
      'Tests validate OOXML-generated slides in web environments',
      'Screenshots captured for visual validation',
      'Tests cover Google Slides and PowerPoint Online compatibility'
    ]
  };
  
  try {
    fs.writeFileSync(summaryFile, JSON.stringify(summary, null, 2));
    console.log('Test summary written to:', summaryFile);
  } catch (error) {
    console.error('Error writing test summary:', error);
  }
  
  console.log('Teardown complete');
  
  return Promise.resolve();
}

module.exports = globalTeardown;