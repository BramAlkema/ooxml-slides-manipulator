/**
 * OOXMLCore Test Runner
 * Simple test runner that can be called via web app
 */

/**
 * Run OOXMLCore comprehensive unit tests
 * Returns simplified test results
 */
function testOOXMLCore() {
  ConsoleFormatter.header('🧪 OOXMLCore Test Runner');
  
  try {
    // Check if OOXMLCore exists
    if (typeof OOXMLCore === 'undefined') {
      return {
        error: 'OOXMLCore class not found',
        passed: false,
        totalTests: 0
      };
    }
    
    // Run basic smoke tests
    var results = {
      totalTests: 0,
      passed: 0,
      failed: 0,
      tests: []
    };
    
    // Test 1: Constructor validation
    try {
      var mockBlob = Utilities.newBlob('test content', 'application/octet-stream');
      var core = new OOXMLCore(mockBlob, 'pptx');
      results.tests.push({ name: 'Constructor', passed: true });
      results.passed++;
    } catch (error) {
      results.tests.push({ name: 'Constructor', passed: false, error: error.message });
      results.failed++;
    }
    results.totalTests++;
    
    // Test 2: Format validation
    try {
      var mockBlob = Utilities.newBlob('test content', 'application/octet-stream');
      var core = new OOXMLCore(mockBlob, 'docx');
      var config = core.getFormatInfo();
      var passed = config.type === 'docx' && config.name === 'Word';
      results.tests.push({ name: 'Format Config', passed });
      if (passed) results.passed++; else results.failed++;
    } catch (error) {
      results.tests.push({ name: 'Format Config', passed: false, error: error.message });
      results.failed++;
    }
    results.totalTests++;
    
    // Test 3: Error handling
    try {
      new OOXMLCore(null, 'pptx');
      results.tests.push({ name: 'Error Handling', passed: false, error: 'Should have thrown error' });
      results.failed++;
    } catch (error) {
      var passed = error.message.includes('OOXML_CORE_001');
      results.tests.push({ name: 'Error Handling', passed });
      if (passed) results.passed++; else results.failed++;
    }
    results.totalTests++;
    
    // Test 4: Metadata tracking
    try {
      var mockBlob = Utilities.newBlob('test content', 'application/octet-stream');
      var core = new OOXMLCore(mockBlob, 'xlsx');
      var metadata = core.getMetadata();
      var passed = metadata.originalSize > 0 && metadata.fileCount === 0;
      results.tests.push({ name: 'Metadata', passed });
      if (passed) results.passed++; else results.failed++;
    } catch (error) {
      results.tests.push({ name: 'Metadata', passed: false, error: error.message });
      results.failed++;
    }
    results.totalTests++;
    
    ConsoleFormatter.summary('Test Results', {
      passed: results.passed,
      failed: results.failed,
      total: results.totalTests,
      successRate: `${((results.passed/results.totalTests) * 100).toFixed(1)}%`
    });
    
    return {
      success: true,
      summary: `${results.passed}/${results.totalTests} tests passed`,
      totalTests: results.totalTests,
      passed: results.passed,
      failed: results.failed,
      allTestsPassed: results.failed === 0,
      details: results.tests
    };
    
  } catch (error) {
    ConsoleFormatter.error('OOXMLCore test runner failed', error);
    return {
      error: 'Test runner failed: ' + error.message,
      passed: false,
      totalTests: 0
    };
  }
}

/**
 * Quick validation for CI/CD
 */
function validateOOXMLCore() {
  var results = testOOXMLCore();
  return {
    passed: results.allTestsPassed || false,
    message: results.summary || results.error
  };
}