/**
 * CloudPPTXService Test Runner
 * Simple test runner that can be called via web app
 */

/**
 * Run CloudPPTXService comprehensive unit tests
 * Returns simplified test results
 */
function testCloudPPTXService() {
  console.log('🧪 Starting CloudPPTXService test runner...');
  
  try {
    // Check if CloudPPTXService exists
    if (typeof CloudPPTXService === 'undefined') {
      return {
        error: 'CloudPPTXService class not found',
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
    
    // Test 1: Service info
    try {
      var info = CloudPPTXService.getServiceInfo();
      var passed = info && info.version && info.capabilities;
      results.tests.push({ name: 'Service Info', passed });
      if (passed) results.passed++; else results.failed++;
    } catch (error) {
      results.tests.push({ name: 'Service Info', passed: false, error: error.message });
      results.failed++;
    }
    results.totalTests++;
    
    // Test 2: Input validation (test the private method directly for sync testing)
    try {
      CloudPPTXService._validateExtractInput(null);
      results.tests.push({ name: 'Input Validation', passed: false, error: 'Should have thrown error' });
      results.failed++;
    } catch (error) {
      var passed = error.message.includes('CLOUD_FUNCTION_005');
      results.tests.push({ name: 'Input Validation', passed });
      if (passed) results.passed++; else results.failed++;
    }
    results.totalTests++;
    
    // Test 3: Configuration access
    try {
      var config = CloudPPTXService._CONFIG;
      var passed = config && config.cloudFunctionUrl && config.retryAttempts;
      results.tests.push({ name: 'Configuration', passed });
      if (passed) results.passed++; else results.failed++;
    } catch (error) {
      results.tests.push({ name: 'Configuration', passed: false, error: error.message });
      results.failed++;
    }
    results.totalTests++;
    
    // Test 4: Error codes access
    try {
      var errors = CloudPPTXService._ERROR_CODES;
      var passed = errors && errors.CLOUD_FUNCTION_001 && errors.CLOUD_FUNCTION_007;
      results.tests.push({ name: 'Error Codes', passed });
      if (passed) results.passed++; else results.failed++;
    } catch (error) {
      results.tests.push({ name: 'Error Codes', passed: false, error: error.message });
      results.failed++;
    }
    results.totalTests++;
    
    console.log(`✅ CloudPPTXService tests completed: ${results.passed}/${results.totalTests} passed`);
    
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
    console.error('❌ CloudPPTXService test runner failed:', error);
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
function validateCloudPPTXService() {
  var results = testCloudPPTXService();
  return {
    passed: results.allTestsPassed || false,
    message: results.summary || results.error
  };
}