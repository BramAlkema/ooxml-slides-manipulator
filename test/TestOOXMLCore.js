/**
 * UNIT TESTS FOR OOXMLCore_v2.js
 * 
 * PURPOSE:
 * Comprehensive test suite for the universal OOXML manipulation engine
 * Tests all public methods, error conditions, and edge cases
 * 
 * TESTING STRATEGY:
 * - Mock FFlatePPTXService for isolated unit testing
 * - Test both success paths and error conditions
 * - Validate all error codes and messages
 * - Test all supported formats (pptx, docx, xlsx)
 * 
 * AI CONTEXT:
 * This test suite validates the OOXMLCore functionality without external dependencies.
 * Run via web app endpoint or manual execution. Tests are designed for automated CI/CD.
 */

class TestOOXMLCore {

  /**
   * Main test runner - executes all tests and returns detailed results
   * 
   * @returns {Object} Test results with passed/failed counts and details
   */
  static async runAllTests() {
    ConsoleFormatter.header('🧪 OOXMLCore Test Suite');
    
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
      'testConstructor',
      'testFormatValidation', 
      'testExtraction',
      'testFileOperations',
      'testCompression',
      'testErrorHandling',
      'testMetadata',
      'testAllFormats'
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
   * Test constructor validation and initialization
   */
  static async testConstructor() {
    const tests = [];
    
    // Test 1: Valid construction
    tests.push(await this._runTest('Valid PPTX construction', async () => {
      const mockBlob = this._createMockBlob('valid pptx content');
      const core = new OOXMLCore(mockBlob, 'pptx');
      
      this._assert(core.formatType === 'pptx', 'Format type should be pptx');
      this._assert(core.isExtracted === false, 'Should not be extracted initially');
      this._assert(core.isModified === false, 'Should not be modified initially');
      this._assert(core.metadata.originalSize > 0, 'Should track original size');
    }));
    
    // Test 2: Invalid blob
    tests.push(await this._runTest('Invalid blob rejection', async () => {
      try {
        new OOXMLCore(null, 'pptx');
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('OOXML_CORE_001'), 'Should use correct error code');
      }
    }));
    
    // Test 3: Invalid format type
    tests.push(await this._runTest('Invalid format type rejection', async () => {
      const mockBlob = this._createMockBlob('content');
      try {
        new OOXMLCore(mockBlob, 'invalid');
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('OOXML_CORE_002'), 'Should use correct error code');
      }
    }));
    
    // Test 4: Options handling
    tests.push(await this._runTest('Options handling', async () => {
      const mockBlob = this._createMockBlob('content');
      const core = new OOXMLCore(mockBlob, 'pptx', { 
        validateFormat: true, 
        enableLogging: false 
      });
      
      this._assert(core.options.validateFormat === true, 'Should set validateFormat');
      this._assert(core.options.enableLogging === false, 'Should set enableLogging');
    }));
    
    return tests;
  }

  /**
   * Test format validation and configuration
   */
  static async testFormatValidation() {
    const tests = [];
    
    // Test all supported formats
    const formats = ['pptx', 'docx', 'xlsx'];
    
    for (const format of formats) {
      tests.push(await this._runTest(`${format.toUpperCase()} format config`, async () => {
        const mockBlob = this._createMockBlob('content');
        const core = new OOXMLCore(mockBlob, format);
        const config = core.getFormatInfo();
        
        this._assert(config.type === format, `Type should be ${format}`);
        this._assert(config.name, 'Should have name');
        this._assert(config.extension === `.${format}`, `Extension should be .${format}`);
        this._assert(config.contentType.includes(format), 'Content type should include format');
        this._assert(config.mainDocument, 'Should have main document path');
      }));
    }
    
    return tests;
  }

  /**
   * Test extraction functionality
   */
  static async testExtraction() {
    const tests = [];
    
    // Mock FFlatePPTXService for testing
    const originalService = globalThis.FFlatePPTXService;
    globalThis.FFlatePPTXService = {
      unzipPPTX: async (blob) => {
        return {
          '[Content_Types].xml': '<?xml version="1.0"?>',
          '_rels/.rels': '<?xml version="1.0"?>',
          'ppt/presentation.xml': '<presentation></presentation>',
          'ppt/slides/slide1.xml': '<slide></slide>'
        };
      }
    };
    
    // Test 1: Successful extraction
    tests.push(await this._runTest('Successful extraction', async () => {
      const mockBlob = this._createMockBlob('pptx content');
      const core = new OOXMLCore(mockBlob, 'pptx');
      
      await core.extract();
      
      this._assert(core.isExtracted === true, 'Should be marked as extracted');
      this._assert(core.files.size === 4, 'Should have extracted 4 files');
      this._assert(core.metadata.fileCount === 4, 'Should track file count');
      this._assert(core.metadata.extractedAt, 'Should track extraction time');
    }));
    
    // Test 2: Double extraction safety
    tests.push(await this._runTest('Double extraction safety', async () => {
      const mockBlob = this._createMockBlob('pptx content');
      const core = new OOXMLCore(mockBlob, 'pptx');
      
      await core.extract();
      const firstCount = core.files.size;
      
      await core.extract(); // Should be safe to call again
      
      this._assert(core.files.size === firstCount, 'File count should remain same');
    }));
    
    // Test 3: Extraction failure handling
    tests.push(await this._runTest('Extraction failure handling', async () => {
      globalThis.FFlatePPTXService.unzipPPTX = async () => {
        throw new Error('Extraction failed');
      };
      
      const mockBlob = this._createMockBlob('invalid content');
      const core = new OOXMLCore(mockBlob, 'pptx');
      
      try {
        await core.extract();
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('OOXML_CORE_003'), 'Should use correct error code');
        this._assert(core.isExtracted === false, 'Should not be marked as extracted');
      }
    }));
    
    // Restore original service
    globalThis.FFlatePPTXService = originalService;
    
    return tests;
  }

  /**
   * Test file operations (get, set, list)
   */
  static async testFileOperations() {
    const tests = [];
    
    // Setup mock service
    globalThis.FFlatePPTXService = {
      unzipPPTX: async () => ({
        'test1.xml': '<test>content1</test>',
        'test2.xml': '<test>content2</test>'
      })
    };
    
    // Test 1: File access before extraction
    tests.push(await this._runTest('File access before extraction error', async () => {
      const mockBlob = this._createMockBlob('content');
      const core = new OOXMLCore(mockBlob, 'pptx');
      
      try {
        core.getFile('test.xml');
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('OOXML_CORE_006'), 'Should use correct error code');
      }
    }));
    
    // Test 2: Get existing file
    tests.push(await this._runTest('Get existing file', async () => {
      const mockBlob = this._createMockBlob('content');
      const core = new OOXMLCore(mockBlob, 'pptx');
      await core.extract();
      
      const content = core.getFile('test1.xml');
      this._assert(content === '<test>content1</test>', 'Should return correct content');
    }));
    
    // Test 3: Get non-existent file
    tests.push(await this._runTest('Get non-existent file', async () => {
      const mockBlob = this._createMockBlob('content');
      const core = new OOXMLCore(mockBlob, 'pptx');
      await core.extract();
      
      const content = core.getFile('nonexistent.xml');
      this._assert(content === null, 'Should return null for non-existent file');
    }));
    
    // Test 4: Set file content
    tests.push(await this._runTest('Set file content', async () => {
      const mockBlob = this._createMockBlob('content');
      const core = new OOXMLCore(mockBlob, 'pptx');
      await core.extract();
      
      core.setFile('test1.xml', '<modified>content</modified>');
      
      this._assert(core.isModified === true, 'Should be marked as modified');
      this._assert(core.metadata.modifiedAt, 'Should track modification time');
      this._assert(core.getFile('test1.xml') === '<modified>content</modified>', 'Should have updated content');
    }));
    
    // Test 5: Add new file
    tests.push(await this._runTest('Add new file', async () => {
      const mockBlob = this._createMockBlob('content');
      const core = new OOXMLCore(mockBlob, 'pptx');
      await core.extract();
      
      const originalCount = core.metadata.fileCount;
      core.setFile('new.xml', '<new>content</new>');
      
      this._assert(core.metadata.fileCount === originalCount + 1, 'Should increment file count');
    }));
    
    // Test 6: Invalid file content type
    tests.push(await this._runTest('Invalid file content type', async () => {
      const mockBlob = this._createMockBlob('content');
      const core = new OOXMLCore(mockBlob, 'pptx');
      await core.extract();
      
      try {
        core.setFile('test.xml', 123);
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('OOXML_CORE_004'), 'Should use correct error code');
      }
    }));
    
    // Test 7: List files
    tests.push(await this._runTest('List files', async () => {
      const mockBlob = this._createMockBlob('content');
      const core = new OOXMLCore(mockBlob, 'pptx');
      await core.extract();
      
      const files = core.listFiles();
      this._assert(Array.isArray(files), 'Should return array');
      this._assert(files.length === 2, 'Should list all files');
      this._assert(files.includes('test1.xml'), 'Should include test1.xml');
      this._assert(files.includes('test2.xml'), 'Should include test2.xml');
    }));
    
    return tests;
  }

  /**
   * Test compression functionality
   */
  static async testCompression() {
    const tests = [];
    
    // Mock FFlatePPTXService
    globalThis.FFlatePPTXService = {
      unzipPPTX: async () => ({
        'test.xml': '<test>content</test>'
      }),
      zipPPTX: async (files) => {
        const mockBlob = Utilities.newBlob('compressed content', 'application/zip');
        return mockBlob;
      }
    };
    
    // Test 1: Successful compression
    tests.push(await this._runTest('Successful compression', async () => {
      const mockBlob = this._createMockBlob('content');
      const core = new OOXMLCore(mockBlob, 'pptx');
      await core.extract();
      
      const compressedBlob = await core.compress();
      
      this._assert(compressedBlob, 'Should return blob');
      this._assert(compressedBlob.getContentType().includes('presentationml'), 'Should have correct MIME type');
    }));
    
    // Test 2: Compression without extraction
    tests.push(await this._runTest('Compression without extraction error', async () => {
      const mockBlob = this._createMockBlob('content');
      const core = new OOXMLCore(mockBlob, 'pptx');
      
      try {
        await core.compress();
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('OOXML_CORE_006'), 'Should use correct error code');
      }
    }));
    
    // Test 3: Compression failure handling
    tests.push(await this._runTest('Compression failure handling', async () => {
      globalThis.FFlatePPTXService.zipPPTX = async () => {
        throw new Error('Compression failed');
      };
      
      const mockBlob = this._createMockBlob('content');
      const core = new OOXMLCore(mockBlob, 'pptx');
      await core.extract();
      
      try {
        await core.compress();
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('OOXML_CORE_005'), 'Should use correct error code');
      }
    }));
    
    return tests;
  }

  /**
   * Test error handling and edge cases
   */
  static async testErrorHandling() {
    const tests = [];
    
    // Test 1: FFlatePPTXService unavailable
    tests.push(await this._runTest('FFlatePPTXService unavailable', async () => {
      const originalService = globalThis.FFlatePPTXService;
      delete globalThis.FFlatePPTXService;
      
      try {
        const mockBlob = this._createMockBlob('content');
        const core = new OOXMLCore(mockBlob, 'pptx');
        await core.extract();
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('FFlatePPTXService not available'), 'Should detect missing service');
      } finally {
        globalThis.FFlatePPTXService = originalService;
      }
    }));
    
    // Test 2: Malformed error propagation
    tests.push(await this._runTest('Error propagation with original error', async () => {
      globalThis.FFlatePPTXService = {
        unzipPPTX: async () => {
          const originalError = new Error('Original failure');
          originalError.code = 'SPECIFIC_CODE';
          throw originalError;
        }
      };
      
      const mockBlob = this._createMockBlob('content');
      const core = new OOXMLCore(mockBlob, 'pptx');
      
      try {
        await core.extract();
        throw new Error('Should have thrown error');
      } catch (error) {
        this._assert(error.message.includes('OOXML_CORE_003'), 'Should wrap with correct error code');
        this._assert(error.originalError, 'Should preserve original error');
        this._assert(error.originalError.message === 'Original failure', 'Should preserve original message');
      }
    }));
    
    return tests;
  }

  /**
   * Test metadata and utility methods
   */
  static async testMetadata() {
    const tests = [];
    
    // Setup mock service
    globalThis.FFlatePPTXService = {
      unzipPPTX: async () => ({ 'test.xml': '<test></test>' })
    };
    
    // Test 1: Metadata tracking
    tests.push(await this._runTest('Metadata tracking', async () => {
      const mockBlob = this._createMockBlob('test content');
      const core = new OOXMLCore(mockBlob, 'pptx');
      
      const initialMetadata = core.getMetadata();
      this._assert(initialMetadata.originalSize > 0, 'Should track original size');
      this._assert(initialMetadata.fileCount === 0, 'Should start with 0 files');
      this._assert(!initialMetadata.extractedAt, 'Should not have extraction time initially');
      
      await core.extract();
      
      const extractedMetadata = core.getMetadata();
      this._assert(extractedMetadata.fileCount === 1, 'Should update file count');
      this._assert(extractedMetadata.extractedAt, 'Should track extraction time');
    }));
    
    // Test 2: Format info
    tests.push(await this._runTest('Format info', async () => {
      const mockBlob = this._createMockBlob('content');
      const core = new OOXMLCore(mockBlob, 'docx');
      
      const formatInfo = core.getFormatInfo();
      this._assert(formatInfo.type === 'docx', 'Should return correct type');
      this._assert(formatInfo.name === 'Word', 'Should return correct name');
      this._assert(formatInfo.extension === '.docx', 'Should return correct extension');
    }));
    
    // Test 3: Modification tracking
    tests.push(await this._runTest('Modification tracking', async () => {
      const mockBlob = this._createMockBlob('content');
      const core = new OOXMLCore(mockBlob, 'pptx');
      await core.extract();
      
      this._assert(core.hasModifications() === false, 'Should not have modifications initially');
      
      core.setFile('test.xml', '<modified></modified>');
      
      this._assert(core.hasModifications() === true, 'Should detect modifications');
    }));
    
    return tests;
  }

  /**
   * Test all supported formats
   */
  static async testAllFormats() {
    const tests = [];
    
    const formats = [
      { type: 'pptx', mainDoc: 'ppt/presentation.xml' },
      { type: 'docx', mainDoc: 'word/document.xml' },
      { type: 'xlsx', mainDoc: 'xl/workbook.xml' }
    ];
    
    // Mock service for all formats
    globalThis.FFlatePPTXService = {
      unzipPPTX: async (blob) => {
        const files = {};
        files['[Content_Types].xml'] = '<?xml version="1.0"?>';
        files['_rels/.rels'] = '<?xml version="1.0"?>';
        return files;
      }
    };
    
    for (const format of formats) {
      tests.push(await this._runTest(`${format.type.toUpperCase()} format end-to-end`, async () => {
        const mockBlob = this._createMockBlob(`${format.type} content`);
        const core = new OOXMLCore(mockBlob, format.type);
        
        // Test construction
        this._assert(core.formatType === format.type, `Should be ${format.type} format`);
        
        const config = core.getFormatInfo();
        this._assert(config.mainDocument === format.mainDoc, 'Should have correct main document path');
        
        // Test extraction
        await core.extract();
        this._assert(core.isExtracted === true, 'Should be extracted');
        
        // Test file operations
        core.setFile(format.mainDoc, `<${format.type}>modified</${format.type}>`);
        this._assert(core.hasModifications() === true, 'Should track modifications');
        
        const content = core.getFile(format.mainDoc);
        this._assert(content.includes('modified'), 'Should have modified content');
      }));
    }
    
    return tests;
  }

  // UTILITY METHODS FOR TESTING

  /**
   * Create a mock blob for testing
   * @private
   */
  static _createMockBlob(content) {
    return Utilities.newBlob(content, 'application/octet-stream');
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
 * Runs all OOXMLCore tests and returns JSON results
 */
function testOOXMLCore() {
  return TestOOXMLCore.runAllTests();
}

/**
 * Quick validation test for CI/CD
 * Returns simple pass/fail status
 */
function validateOOXMLCore() {
  return TestOOXMLCore.runAllTests().then(results => ({
    passed: results.failed === 0,
    totalTests: results.totalTests,
    passedTests: results.passed,
    failedTests: results.failed,
    duration: results.duration
  }));
}