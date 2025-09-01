/**
 * Unit Tests for FFlatePPTXService
 * 
 * Tests all functionality of the new fflate-based PPTX service including:
 * - File extraction and compression
 * - Text/binary file handling
 * - Error scenarios and validation
 * - Conversion utilities
 * - Performance and edge cases
 */

function testFFlatePPTXService() {
  ConsoleFormatter.header('🧪 FFlatePPTXService Unit Tests');
  
  const results = {
    passed: 0,
    failed: 0,
    total: 0,
    startTime: Date.now(),
    testSuites: {}
  };

  // Test suites
  const testSuites = {
    'Validation Tests': testValidation,
    'Conversion Utilities': testConversionUtilities, 
    'File Type Detection': testFileTypeDetection,
    'Extract Operation': testExtractOperation,
    'Compress Operation': testCompressOperation,
    'Error Handling': testErrorHandling,
    'Performance Tests': testPerformance,
    'Legacy Compatibility': testLegacyCompatibility
  };

  // Run all test suites
  for (const [suiteName, suiteFunction] of Object.entries(testSuites)) {
    ConsoleFormatter.section(`📋 ${suiteName}`);
    const suiteResults = suiteFunction();
    results.testSuites[suiteName] = suiteResults;
    results.passed += suiteResults.passed;
    results.failed += suiteResults.failed;
    results.total += suiteResults.total;
  }

  const endTime = Date.now();
  const duration = endTime - results.startTime;

  // Print summary
  ConsoleFormatter.summary('Test Results', {
    passed: results.passed,
    failed: results.failed,
    total: results.total,
    duration: `${duration}ms`,
    successRate: `${((results.passed/results.total) * 100).toFixed(1)}%`
  });

  if (results.failed > 0) {
    ConsoleFormatter.section('❌ Failed Test Details');
    for (const [suiteName, suiteResults] of Object.entries(results.testSuites)) {
      if (suiteResults.failed > 0) {
        ConsoleFormatter.status('FAIL', suiteName, `${suiteResults.failed} failures`);
      }
    }
  }

  return results;
}

function testValidation() {
  const tests = [];

  // Test extract input validation
  tests.push(() => {
    try {
      FFlatePPTXService.extractFiles(null);
      return { success: false, message: 'Should have thrown for null blob' };
    } catch (error) {
      return error.message.includes('OOXML blob is required') 
        ? { success: true } 
        : { success: false, message: `Wrong error: ${error.message}` };
    }
  });

  tests.push(() => {
    try {
      FFlatePPTXService.extractFiles({ invalid: 'object' });
      return { success: false, message: 'Should have thrown for invalid blob' };
    } catch (error) {
      return error.message.includes('Invalid blob object')
        ? { success: true }
        : { success: false, message: `Wrong error: ${error.message}` };
    }
  });

  // Test compress input validation
  tests.push(() => {
    try {
      FFlatePPTXService.compressFiles(null);
      return { success: false, message: 'Should have thrown for null files' };
    } catch (error) {
      return error.message.includes('Files object is required')
        ? { success: true }
        : { success: false, message: `Wrong error: ${error.message}` };
    }
  });

  tests.push(() => {
    try {
      FFlatePPTXService.compressFiles({});
      return { success: false, message: 'Should have thrown for empty files' };
    } catch (error) {
      return error.message.includes('No files to compress')
        ? { success: true }
        : { success: false, message: `Wrong error: ${error.message}` };
    }
  });

  tests.push(() => {
    try {
      FFlatePPTXService.compressFiles({ 'test.xml': 'content' });
      return { success: false, message: 'Should have thrown for missing Content_Types' };
    } catch (error) {
      return error.message.includes('Missing required file: [Content_Types].xml')
        ? { success: true }
        : { success: false, message: `Wrong error: ${error.message}` };
    }
  });

  return runTestSuite('Validation Tests', tests);
}

function testConversionUtilities() {
  const tests = [];

  // Test string to Uint8Array conversion
  tests.push(() => {
    const testString = 'Hello World! 🎉';
    const result = FFlatePPTXService.stringToUint8Array(testString);
    
    if (!(result instanceof Uint8Array)) {
      return { success: false, message: 'Result should be Uint8Array' };
    }
    
    if (result.length === 0) {
      return { success: false, message: 'Result should not be empty' };
    }
    
    return { success: true };
  });

  // Test Uint8Array to string conversion
  tests.push(() => {
    const testString = 'Hello World! 🎉';
    const uint8Array = FFlatePPTXService.stringToUint8Array(testString);
    const result = FFlatePPTXService.uint8ArrayToString(uint8Array);
    
    return result === testString
      ? { success: true }
      : { success: false, message: `Expected "${testString}", got "${result}"` };
  });

  // Test XML content conversion
  tests.push(() => {
    const xmlContent = '<?xml version="1.0"?><root><item>test</item></root>';
    const uint8Array = FFlatePPTXService.stringToUint8Array(xmlContent);
    const result = FFlatePPTXService.uint8ArrayToString(uint8Array);
    
    return result === xmlContent
      ? { success: true }
      : { success: false, message: 'XML conversion failed' };
  });

  // Test empty string conversion
  tests.push(() => {
    const emptyString = '';
    const uint8Array = FFlatePPTXService.stringToUint8Array(emptyString);
    const result = FFlatePPTXService.uint8ArrayToString(uint8Array);
    
    return result === emptyString && uint8Array.length === 0
      ? { success: true }
      : { success: false, message: 'Empty string conversion failed' };
  });

  return runTestSuite('Conversion Utilities', tests);
}

function testFileTypeDetection() {
  const tests = [];

  // Test text file detection by extension
  tests.push(() => {
    const textFiles = [
      'slide1.xml',
      'theme.xml', 
      '_rels/.rels',
      'app.xml',
      'core.xml',
      'custom.xml'
    ];
    
    for (const filename of textFiles) {
      if (!FFlatePPTXService._isTextFile(filename)) {
        return { success: false, message: `${filename} should be detected as text file` };
      }
    }
    
    return { success: true };
  });

  // Test binary file detection
  tests.push(() => {
    const binaryFiles = [
      'image1.png',
      'image2.jpg', 
      'media/image1.jpeg',
      'audio.wav',
      'video.mp4',
      'font.ttf'
    ];
    
    for (const filename of binaryFiles) {
      if (FFlatePPTXService._isTextFile(filename)) {
        return { success: false, message: `${filename} should be detected as binary file` };
      }
    }
    
    return { success: true };
  });

  // Test special file detection
  tests.push(() => {
    const specialFiles = [
      '[Content_Types].xml',
      '_rels/.rels'
    ];
    
    for (const filename of specialFiles) {
      if (!FFlatePPTXService._isTextFile(filename)) {
        return { success: false, message: `${filename} should be detected as text file` };
      }
    }
    
    return { success: true };
  });

  return runTestSuite('File Type Detection', tests);
}

function testExtractOperation() {
  const tests = [];

  // Mock fflate for testing
  const mockFflate = {
    unzipSync: (data) => ({
      'ppt/slides/slide1.xml': new Uint8Array([60, 63, 120, 109, 108]), // "<?xml"
      'ppt/media/image1.png': new Uint8Array([137, 80, 78, 71]), // PNG header
      '[Content_Types].xml': new Uint8Array([60, 63, 120, 109, 108]),
      '_rels/.rels': new Uint8Array([60, 63, 120, 109, 108])
    })
  };

  // Test successful extraction with mocked fflate
  tests.push(() => {
    const originalFflate = global.fflate;
    global.fflate = mockFflate;
    
    try {
      const mockBlob = {
        getBytes: () => new Array(100).fill(0) // Mock PPTX data
      };
      
      const result = FFlatePPTXService.extractFiles(mockBlob);
      
      if (typeof result !== 'object') {
        return { success: false, message: 'Result should be an object' };
      }
      
      if (!result['[Content_Types].xml']) {
        return { success: false, message: 'Should contain Content_Types.xml' };
      }
      
      // Check text file conversion
      if (typeof result['ppt/slides/slide1.xml'] !== 'string') {
        return { success: false, message: 'XML files should be converted to strings' };
      }
      
      // Check binary file preservation
      if (!(result['ppt/media/image1.png'] instanceof Uint8Array)) {
        return { success: false, message: 'Binary files should remain as Uint8Array' };
      }
      
      return { success: true };
      
    } catch (error) {
      return { success: false, message: `Unexpected error: ${error.message}` };
    } finally {
      global.fflate = originalFflate;
    }
  });

  return runTestSuite('Extract Operation', tests);
}

function testCompressOperation() {
  const tests = [];

  // Mock fflate for testing
  const mockFflate = {
    zipSync: (files, options) => new Uint8Array([80, 75, 3, 4]) // ZIP header
  };

  // Test successful compression
  tests.push(() => {
    const originalFflate = global.fflate;
    const originalUtilities = global.Utilities;
    
    global.fflate = mockFflate;
    global.Utilities = {
      newBlob: (data, type, name) => ({ data, type, name })
    };
    
    try {
      const files = {
        '[Content_Types].xml': '<?xml version="1.0"?>',
        '_rels/.rels': '<?xml version="1.0"?>',
        'ppt/slides/slide1.xml': '<slide></slide>',
        'ppt/media/image1.png': new Uint8Array([137, 80, 78, 71])
      };
      
      const result = FFlatePPTXService.compressFiles(files);
      
      if (!result) {
        return { success: false, message: 'Should return a blob' };
      }
      
      if (!result.type.includes('presentation')) {
        return { success: false, message: 'Should have correct MIME type' };
      }
      
      return { success: true };
      
    } catch (error) {
      return { success: false, message: `Unexpected error: ${error.message}` };
    } finally {
      global.fflate = originalFflate;
      global.Utilities = originalUtilities;
    }
  });

  return runTestSuite('Compress Operation', tests);
}

function testErrorHandling() {
  const tests = [];

  // Test missing fflate library
  tests.push(() => {
    const originalFflate = global.fflate;
    delete global.fflate;
    
    try {
      const mockBlob = {
        getBytes: () => new Array(100).fill(0)
      };
      
      FFlatePPTXService.extractFiles(mockBlob);
      return { success: false, message: 'Should have thrown for missing fflate' };
      
    } catch (error) {
      const success = error.message.includes('fflate library not available');
      return success 
        ? { success: true }
        : { success: false, message: `Wrong error: ${error.message}` };
    } finally {
      global.fflate = originalFflate;
    }
  });

  return runTestSuite('Error Handling', tests);
}

function testPerformance() {
  const tests = [];

  // Test conversion performance
  tests.push(() => {
    const startTime = Date.now();
    const testString = 'x'.repeat(10000); // 10KB string
    
    for (let i = 0; i < 100; i++) {
      const uint8Array = FFlatePPTXService.stringToUint8Array(testString);
      const result = FFlatePPTXService.uint8ArrayToString(uint8Array);
    }
    
    const duration = Date.now() - startTime;
    
    // Should complete 100 conversions in under 1 second
    return duration < 1000
      ? { success: true }
      : { success: false, message: `Too slow: ${duration}ms` };
  });

  return runTestSuite('Performance Tests', tests);
}

function testLegacyCompatibility() {
  const tests = [];

  // Test legacy method aliases
  tests.push(() => {
    const hasUnzipAlias = typeof FFlatePPTXService.unzipPPTX === 'function';
    const hasZipAlias = typeof FFlatePPTXService.zipPPTX === 'function';
    
    return hasUnzipAlias && hasZipAlias
      ? { success: true }
      : { success: false, message: 'Legacy aliases missing' };
  });

  return runTestSuite('Legacy Compatibility', tests);
}

/**
 * Run a test suite and return results
 * @param {string} suiteName - Name of the test suite
 * @param {Function[]} tests - Array of test functions
 * @returns {Object} Test results
 */
function runTestSuite(suiteName, tests) {
  const results = {
    passed: 0,
    failed: 0,
    total: tests.length,
    failures: []
  };

  for (let i = 0; i < tests.length; i++) {
    const test = tests[i];
    const startTime = Date.now();
    
    try {
      const result = test();
      const duration = Date.now() - startTime;
      
      if (result.success) {
        results.passed++;
        ConsoleFormatter.testResult(`Test ${i + 1}`, true, null, duration);
      } else {
        results.failed++;
        results.failures.push(`Test ${i + 1}: ${result.message}`);
        ConsoleFormatter.testResult(`Test ${i + 1}`, false, result.message, duration);
      }
    } catch (error) {
      results.failed++;
      results.failures.push(`Test ${i + 1}: ${error.message}`);
      ConsoleFormatter.status('ERROR', `Test ${i + 1}`, error.message);
    }
  }

  const successRate = ((results.passed / results.total) * 100).toFixed(1);
  ConsoleFormatter.status('INFO', suiteName, `${results.passed}/${results.total} passed (${successRate}%)`);

  return results;
}

/**
 * Validate FFlatePPTXService implementation
 * This function can be called to verify the service works correctly
 */
function validateFFlatePPTXService() {
  ConsoleFormatter.header('🔍 FFlatePPTXService Validation');

  // Check class exists
  if (typeof FFlatePPTXService === 'undefined') {
    ConsoleFormatter.status('FAIL', 'Class Check', 'FFlatePPTXService class not found');
    return false;
  }

  // Check required methods
  const requiredMethods = [
    'extractFiles',
    'compressFiles', 
    'stringToUint8Array',
    'uint8ArrayToString',
    'getServiceInfo'
  ];

  for (const method of requiredMethods) {
    if (typeof FFlatePPTXService[method] !== 'function') {
      ConsoleFormatter.status('FAIL', 'Method Check', `Missing method: ${method}`);
      return false;
    }
  }

  // Check legacy aliases
  if (typeof FFlatePPTXService.unzipPPTX !== 'function') {
    ConsoleFormatter.status('FAIL', 'Legacy Alias Check', 'Missing legacy alias: unzipPPTX');
    return false;
  }

  if (typeof FFlatePPTXService.zipPPTX !== 'function') {
    ConsoleFormatter.status('FAIL', 'Legacy Alias Check', 'Missing legacy alias: zipPPTX');
    return false;
  }

  // Check service info
  try {
    const info = FFlatePPTXService.getServiceInfo();
    if (!info.service || !info.version) {
      ConsoleFormatter.status('FAIL', 'Service Info Check', 'Invalid service info');
      return false;
    }
  } catch (error) {
    ConsoleFormatter.status('ERROR', 'Service Info Check', error.message);
    return false;
  }

  ConsoleFormatter.success('FFlatePPTXService validation passed');
  return true;
}