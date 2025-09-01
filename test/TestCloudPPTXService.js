/**
 * UNIT TESTS FOR CloudPPTXService_v2.js
 * 
 * PURPOSE:
 * Comprehensive test suite for the Cloud Function bridge service
 * Tests network scenarios, error handling, fallback mechanisms, and performance
 * 
 * TESTING STRATEGY:
 * - Mock UrlFetchApp for network simulation
 * - Mock Utilities for fallback testing
 * - Test retry logic and error recovery
 * - Validate error codes and messages
 * - Test various failure scenarios
 * 
 * AI CONTEXT:
 * This test suite validates the CloudPPTXService without external dependencies.
 * Focuses on error handling, network resilience, and fallback behavior.
 */

class TestCloudPPTXService {

  /**
   * Main test runner for CloudPPTXService v2
   * 
   * @returns {Object} Test results with detailed breakdown
   */
  static async runAllTests() {
    ConsoleFormatter.header('🧪 CloudPPTXService Test Suite');
    
    const results = {
      totalTests: 0,
      passed: 0,
      failed: 0,
      details: [],
      startTime: new Date().toISOString(),
      endTime: null,
      duration: 0
    };
    
    const startTime = Date.now();
    
    // Test categories
    const testSuites = [
      'testInputValidation',
      'testExtraction',
      'testCompression', 
      'testErrorHandling',
      'testFallbackMechanisms',
      'testHealthCheck',
      'testRetryLogic',
      'testServiceInfo'
    ];
    
    for (const suiteName of testSuites) {
      try {
        ConsoleFormatter.section(`📋 ${suiteName}`);
        const suiteResults = await this[suiteName]();
        
        results.totalTests += suiteResults.length;
        suiteResults.forEach(test => {
          if (test.passed) {
            results.passed++;
          } else {
            results.failed++;
          }
          results.details.push({
            suite: suiteName,
            ...test
          });
        });
        
      } catch (error) {
        results.failed++;
        results.details.push({
          suite: suiteName,
          test: 'Suite Execution',
          passed: false,
          error: error.message,
          stack: error.stack
        });
      }
    }
    
    const endTime = Date.now();
    results.endTime = new Date().toISOString();
    results.duration = endTime - startTime;
    
    ConsoleFormatter.summary('Test Results', {
      passed: results.passed,
      failed: results.failed,
      total: results.totalTests,
      duration: `${results.duration}ms`,
      successRate: `${((results.passed/results.totalTests) * 100).toFixed(1)}%`
    });
    
    return results;
  }

  /**
   * Test input validation
   */
  static async testInputValidation() {
    const tests = [];
    
    // Test 1: Null blob validation
    tests.push(await this._runTest('Null blob rejection', async () => {
      try {
        await CloudPPTXService.extractFiles(null);
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('CLOUD_FUNCTION_005'), 'Should use input validation error code');
      }
    }));
    
    // Test 2: Invalid blob validation
    tests.push(await this._runTest('Invalid blob rejection', async () => {
      try {
        await CloudPPTXService.extractFiles({ notABlob: true });
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('CLOUD_FUNCTION_005'), 'Should use input validation error code');
      }
    }));
    
    // Test 3: Empty files validation
    tests.push(await this._runTest('Empty files validation', async () => {
      try {
        await CloudPPTXService.compressFiles({});
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('CLOUD_FUNCTION_005'), 'Should reject empty files');
      }
    }));
    
    // Test 4: Invalid files object validation
    tests.push(await this._runTest('Invalid files object validation', async () => {
      try {
        await CloudPPTXService.compressFiles(null);
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('CLOUD_FUNCTION_005'), 'Should reject null files');
      }
    }));
    
    return tests;
  }

  /**
   * Test extraction functionality
   */
  static async testExtraction() {
    const tests = [];
    
    // Mock UrlFetchApp for successful extraction
    const originalUrlFetch = globalThis.UrlFetchApp;
    globalThis.UrlFetchApp = {
      fetch: (url, options) => ({
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify({
          success: true,
          files: {
            '[Content_Types].xml': '<?xml version="1.0"?>',
            '_rels/.rels': '<?xml version="1.0"?>',
            'test.xml': '<test>content</test>'
          },
          fileCount: 3
        })
      })
    };
    
    // Test 1: Successful extraction
    tests.push(await this._runTest('Successful extraction', async () => {
      const mockBlob = this._createMockBlob('test content');
      const result = await CloudPPTXService.extractFiles(mockBlob);
      
      this._assert(typeof result === 'object', 'Should return object');
      this._assert(result['[Content_Types].xml'], 'Should contain Content_Types.xml');
      this._assert(result['_rels/.rels'], 'Should contain relationships');
      this._assert(Object.keys(result).length === 3, 'Should return correct file count');
    }));
    
    // Test 2: Invalid response handling
    globalThis.UrlFetchApp.fetch = () => ({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        success: false,
        error: 'Test error'
      })
    });
    
    tests.push(await this._runTest('Invalid response handling', async () => {
      const mockBlob = this._createMockBlob('test content');
      try {
        await CloudPPTXService.extractFiles(mockBlob, { useFallback: false });
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('CLOUD_FUNCTION_007'), 'Should use compression error code');
      }
    }));
    
    // Restore original
    globalThis.UrlFetchApp = originalUrlFetch;
    
    return tests;
  }

  /**
   * Test compression functionality
   */
  static async testCompression() {
    const tests = [];
    
    // Mock UrlFetchApp for successful compression
    const originalUrlFetch = globalThis.UrlFetchApp;
    globalThis.UrlFetchApp = {
      fetch: (url, options) => ({
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify({
          success: true,
          pptxData: 'dGVzdCBkYXRh', // base64 for 'test data'
          size: 9,
          fileCount: 2
        })
      })
    };
    
    // Test 1: Successful compression
    tests.push(await this._runTest('Successful compression', async () => {
      const files = {
        'test1.xml': '<test>content1</test>',
        'test2.xml': '<test>content2</test>'
      };
      
      const result = await CloudPPTXService.compressFiles(files);
      
      this._assert(result && result.getBytes, 'Should return blob');
      this._assert(result.getBytes().length > 0, 'Should have content');
    }));
    
    // Test 2: Compression failure handling
    globalThis.UrlFetchApp.fetch = () => ({
      getResponseCode: () => 200,
      getContentText: () => JSON.stringify({
        success: false,
        error: 'Compression failed'
      })
    });
    
    tests.push(await this._runTest('Compression failure handling', async () => {
      const files = { 'test.xml': '<test>content</test>' };
      
      try {
        await CloudPPTXService.compressFiles(files, { useFallback: false });
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('CLOUD_FUNCTION_007'), 'Should use compression error code');
      }
    }));
    
    // Restore original
    globalThis.UrlFetchApp = originalUrlFetch;
    
    return tests;
  }

  /**
   * Test error handling and recovery
   */
  static async testErrorHandling() {
    const tests = [];
    
    // Test 1: Network timeout error
    const originalUrlFetch = globalThis.UrlFetchApp;
    globalThis.UrlFetchApp = {
      fetch: () => {
        throw new Error('Request timeout');
      }
    };
    
    tests.push(await this._runTest('Network timeout handling', async () => {
      const mockBlob = this._createMockBlob('test content');
      
      try {
        await CloudPPTXService.extractFiles(mockBlob, { useFallback: false });
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('CLOUD_FUNCTION_004'), 'Should use network error code');
      }
    }));
    
    // Test 2: HTTP error codes
    globalThis.UrlFetchApp.fetch = () => ({
      getResponseCode: () => 500,
      getContentText: () => 'Internal Server Error'
    });
    
    tests.push(await this._runTest('HTTP 500 error handling', async () => {
      const mockBlob = this._createMockBlob('test content');
      
      try {
        await CloudPPTXService.extractFiles(mockBlob, { useFallback: false });
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('CLOUD_FUNCTION_004'), 'Should use network error code');
      }
    }));
    
    // Test 3: Invalid JSON response
    globalThis.UrlFetchApp.fetch = () => ({
      getResponseCode: () => 200,
      getContentText: () => 'Invalid JSON'
    });
    
    tests.push(await this._runTest('Invalid JSON response handling', async () => {
      const mockBlob = this._createMockBlob('test content');
      
      try {
        await CloudPPTXService.extractFiles(mockBlob, { useFallback: false });
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('CLOUD_FUNCTION_007'), 'Should handle JSON parsing errors');
      }
    }));
    
    // Restore original
    globalThis.UrlFetchApp = originalUrlFetch;
    
    return tests;
  }

  /**
   * Test fallback mechanisms
   */
  static async testFallbackMechanisms() {
    const tests = [];
    
    // Mock failed UrlFetchApp and successful Utilities
    const originalUrlFetch = globalThis.UrlFetchApp;
    const originalUtilities = globalThis.Utilities;
    
    globalThis.UrlFetchApp = {
      fetch: () => {
        throw new Error('Network unavailable');
      }
    };
    
    globalThis.Utilities = {
      unzip: (blob) => [
        {
          getName: () => 'test.xml',
          getDataAsString: () => '<test>content</test>',
          isGoogleType: () => false
        }
      ],
      zip: (blobs, filename) => this._createMockBlob('compressed data'),
      newBlob: (data, type, name) => this._createMockBlob(data),
      base64Encode: (data) => 'base64data',
      base64Decode: (str) => new Uint8Array([116, 101, 115, 116]) // 'test'
    };
    
    // Test 1: Fallback extraction
    tests.push(await this._runTest('Fallback extraction', async () => {
      const mockBlob = this._createMockBlob('test content');
      
      const result = await CloudPPTXService.extractFiles(mockBlob, { useFallback: true });
      
      this._assert(typeof result === 'object', 'Should return object');
      this._assert(result['test.xml'] === '<test>content</test>', 'Should extract via fallback');
    }));
    
    // Test 2: Fallback compression
    tests.push(await this._runTest('Fallback compression', async () => {
      const files = { 'test.xml': '<test>content</test>' };
      
      const result = await CloudPPTXService.compressFiles(files, { useFallback: true });
      
      this._assert(result && result.getBytes, 'Should return blob via fallback');
    }));
    
    // Test 3: Fallback failure
    globalThis.Utilities.unzip = () => {
      throw new Error('Fallback failed');
    };
    
    tests.push(await this._runTest('Fallback failure handling', async () => {
      const mockBlob = this._createMockBlob('test content');
      
      try {
        await CloudPPTXService.extractFiles(mockBlob, { useFallback: true });
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('CLOUD_FUNCTION_006'), 'Should use fallback error code');
      }
    }));
    
    // Restore originals
    globalThis.UrlFetchApp = originalUrlFetch;
    globalThis.Utilities = originalUtilities;
    
    return tests;
  }

  /**
   * Test health check functionality
   */
  static async testHealthCheck() {
    const tests = [];
    
    // Test 1: Healthy service
    const originalUrlFetch = globalThis.UrlFetchApp;
    globalThis.UrlFetchApp = {
      fetch: (url, options) => ({
        getResponseCode: () => 200,
        getContentText: () => 'OK'
      })
    };
    
    tests.push(await this._runTest('Healthy service detection', async () => {
      const health = await CloudPPTXService.healthCheck();
      
      this._assert(health.available === true, 'Should detect healthy service');
      this._assert(typeof health.responseTime === 'number', 'Should measure response time');
      this._assert(health.statusCode === 200, 'Should return status code');
      this._assert(health.timestamp, 'Should include timestamp');
    }));
    
    // Test 2: Unhealthy service
    globalThis.UrlFetchApp.fetch = () => ({
      getResponseCode: () => 500,
      getContentText: () => 'Server Error'
    });
    
    tests.push(await this._runTest('Unhealthy service detection', async () => {
      const health = await CloudPPTXService.healthCheck();
      
      this._assert(health.available === false, 'Should detect unhealthy service');
      this._assert(health.statusCode === 500, 'Should return error status code');
    }));
    
    // Test 3: Network failure
    globalThis.UrlFetchApp.fetch = () => {
      throw new Error('Network error');
    };
    
    tests.push(await this._runTest('Network failure detection', async () => {
      const health = await CloudPPTXService.healthCheck();
      
      this._assert(health.available === false, 'Should detect network failure');
      this._assert(health.error, 'Should include error message');
    }));
    
    // Restore original
    globalThis.UrlFetchApp = originalUrlFetch;
    
    return tests;
  }

  /**
   * Test retry logic (simplified for testing)
   */
  static async testRetryLogic() {
    const tests = [];
    
    // Test 1: Retry on network error (mock implementation)
    tests.push(await this._runTest('Retry logic validation', async () => {
      let attemptCount = 0;
      
      const originalUrlFetch = globalThis.UrlFetchApp;
      globalThis.UrlFetchApp = {
        fetch: (url, options) => {
          attemptCount++;
          if (attemptCount < 2) {
            throw new Error('Network timeout');
          }
          return {
            getResponseCode: () => 200,
            getContentText: () => JSON.stringify({ success: true, files: {} })
          };
        }
      };
      
      // Mock Utilities.sleep to avoid actual delays
      const originalSleep = globalThis.Utilities?.sleep;
      if (globalThis.Utilities) {
        globalThis.Utilities.sleep = () => {}; // No-op for testing
      }
      
      try {
        const mockBlob = this._createMockBlob('test content');
        await CloudPPTXService.extractFiles(mockBlob, { useFallback: false });
        
        this._assert(attemptCount >= 2, 'Should retry on network error');
        
      } finally {
        globalThis.UrlFetchApp = originalUrlFetch;
        if (globalThis.Utilities && originalSleep) {
          globalThis.Utilities.sleep = originalSleep;
        }
      }
    }));
    
    return tests;
  }

  /**
   * Test service information
   */
  static async testServiceInfo() {
    const tests = [];
    
    // Test 1: Service info structure
    tests.push(await this._runTest('Service info structure', async () => {
      const info = CloudPPTXService.getServiceInfo();
      
      this._assert(info.version, 'Should have version');
      this._assert(info.cloudFunctionUrl, 'Should have Cloud Function URL');
      this._assert(info.capabilities, 'Should have capabilities');
      this._assert(info.supportedFormats, 'Should have supported formats');
      this._assert(Array.isArray(info.supportedFormats), 'Supported formats should be array');
      this._assert(info.supportedFormats.includes('pptx'), 'Should support pptx');
    }));
    
    return tests;
  }

  // UTILITY METHODS FOR TESTING

  /**
   * Create a mock blob for testing
   * @private
   */
  static _createMockBlob(content) {
    return {
      getBytes: () => new TextEncoder().encode(content),
      setContentType: (type) => this,
      getContentType: () => 'application/octet-stream'
    };
  }

  /**
   * Run a single test with error handling
   * @private
   */
  static async _runTest(testName, testFunction) {
    const startTime = Date.now();
    
    try {
      await testFunction();
      const duration = Date.now() - startTime;
      ConsoleFormatter.testResult(testName, true, null, duration);
      return {
        test: testName,
        passed: true,
        duration: duration
      };
    } catch (error) {
      const duration = Date.now() - startTime;
      ConsoleFormatter.testResult(testName, false, error.message, duration);
      return {
        test: testName,
        passed: false,
        duration: duration,
        error: error.message,
        stack: error.stack
      };
    }
  }

  /**
   * Assert utility for tests
   * @private
   */
  static _assert(condition, message) {
    if (!condition) {
      throw new Error(`Assertion failed: ${message}`);
    }
  }
}

/**
 * Web app compatible test runner
 */
function testCloudPPTXService() {
  return TestCloudPPTXService.runAllTests();
}

/**
 * Quick validation test for CI/CD
 */
function validateCloudPPTXService() {
  return TestCloudPPTXService.runAllTests().then(results => ({
    passed: results.failed === 0,
    totalTests: results.totalTests,
    passedTests: results.passed,
    failedTests: results.failed,
    duration: results.duration
  }));
}