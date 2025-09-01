/**
 * Extension Framework Test Runner
 * Tests the extensible API framework and example extensions
 */

/**
 * Test the Extension Framework functionality
 * Returns comprehensive test results
 */
function testExtensionFramework() {
  console.log('🧪 Starting Extension Framework test suite...');
  
  try {
    var results = {
      totalTests: 0,
      passed: 0,
      failed: 0,
      tests: []
    };
    
    // Test 1: Extension Framework availability
    try {
      var framework = typeof ExtensionFramework !== 'undefined';
      results.tests.push({ name: 'Extension Framework Available', passed: framework });
      if (framework) results.passed++; else results.failed++;
    } catch (error) {
      results.tests.push({ name: 'Extension Framework Available', passed: false, error: error.message });
      results.failed++;
    }
    results.totalTests++;
    
    // Test 2: Base Extension availability
    try {
      var baseExtension = typeof BaseExtension !== 'undefined';
      results.tests.push({ name: 'Base Extension Available', passed: baseExtension });
      if (baseExtension) results.passed++; else results.failed++;
    } catch (error) {
      results.tests.push({ name: 'Base Extension Available', passed: false, error: error.message });
      results.failed++;
    }
    results.totalTests++;
    
    // Test 3: Extension registration
    try {
      class TestExtension extends BaseExtension {
        async _customExecute(input) { return { success: true }; }
      }
      
      var registered = ExtensionFramework.register('TestExtension', TestExtension, {
        type: 'CONTENT',
        description: 'Test extension'
      });
      
      results.tests.push({ name: 'Extension Registration', passed: registered });
      if (registered) results.passed++; else results.failed++;
    } catch (error) {
      results.tests.push({ name: 'Extension Registration', passed: false, error: error.message });
      results.failed++;
    }
    results.totalTests++;
    
    // Test 4: Extension listing
    try {
      var extensions = ExtensionFramework.list();
      var hasExtensions = Array.isArray(extensions) && extensions.length > 0;
      results.tests.push({ name: 'Extension Listing', passed: hasExtensions, count: extensions.length });
      if (hasExtensions) results.passed++; else results.failed++;
    } catch (error) {
      results.tests.push({ name: 'Extension Listing', passed: false, error: error.message });
      results.failed++;
    }
    results.totalTests++;
    
    // Test 5: Extension types validation
    try {
      var types = ExtensionFramework._EXTENSION_TYPES;
      var hasTypes = types && types.THEME && types.VALIDATION && types.CONTENT;
      results.tests.push({ name: 'Extension Types', passed: hasTypes });
      if (hasTypes) results.passed++; else results.failed++;
    } catch (error) {
      results.tests.push({ name: 'Extension Types', passed: false, error: error.message });
      results.failed++;
    }
    results.totalTests++;
    
    // Test 6: Template system
    try {
      var template = ExtensionFramework.getTemplate('THEME');
      var hasTemplate = template && template.template && template.example;
      results.tests.push({ name: 'Template System', passed: hasTemplate });
      if (hasTemplate) results.passed++; else results.failed++;
    } catch (error) {
      results.tests.push({ name: 'Template System', passed: false, error: error.message });
      results.failed++;
    }
    results.totalTests++;
    
    // Test 7: OOXMLSlides integration
    try {
      var slidesExtensible = typeof OOXMLSlides !== 'undefined';
      results.tests.push({ name: 'OOXMLSlides Available', passed: slidesExtensible });
      if (slidesExtensible) results.passed++; else results.failed++;
    } catch (error) {
      results.tests.push({ name: 'OOXMLSlides Available', passed: false, error: error.message });
      results.failed++;
    }
    results.totalTests++;
    
    // Test 8: Extension Templates availability
    try {
      var templates = typeof ExtensionTemplates !== 'undefined';
      results.tests.push({ name: 'Extension Templates Available', passed: templates });
      if (templates) results.passed++; else results.failed++;
    } catch (error) {
      results.tests.push({ name: 'Extension Templates Available', passed: false, error: error.message });
      results.failed++;
    }
    results.totalTests++;
    
    console.log(`✅ Extension Framework tests completed: ${results.passed}/${results.totalTests} passed`);
    
    return {
      success: true,
      summary: `${results.passed}/${results.totalTests} tests passed`,
      totalTests: results.totalTests,
      passed: results.passed,
      failed: results.failed,
      allTestsPassed: results.failed === 0,
      details: results.tests,
      framework: {
        available: results.tests.some(t => t.name === 'Extension Framework Available' && t.passed),
        extensionCount: results.tests.find(t => t.name === 'Extension Listing')?.count || 0,
        templatesAvailable: results.tests.some(t => t.name === 'Extension Templates Available' && t.passed)
      }
    };
    
  } catch (error) {
    console.error('❌ Extension Framework test runner failed:', error);
    return {
      error: 'Test runner failed: ' + error.message,
      passed: false,
      totalTests: 0
    };
  }
}

/**
 * Test example extensions specifically
 */
function testExampleExtensions() {
  console.log('🧪 Testing example extensions...');
  
  try {
    var results = {
      totalTests: 0,
      passed: 0,
      failed: 0,
      tests: []
    };
    
    // Test BrandColors extension
    try {
      var brandColors = typeof BrandColorsExtension !== 'undefined';
      results.tests.push({ name: 'BrandColors Extension', passed: brandColors });
      if (brandColors) results.passed++; else results.failed++;
    } catch (error) {
      results.tests.push({ name: 'BrandColors Extension', passed: false, error: error.message });
      results.failed++;
    }
    results.totalTests++;
    
    // Test BrandCompliance extension
    try {
      var brandCompliance = typeof BrandComplianceExtension !== 'undefined';
      results.tests.push({ name: 'BrandCompliance Extension', passed: brandCompliance });
      if (brandCompliance) results.passed++; else results.failed++;
    } catch (error) {
      results.tests.push({ name: 'BrandCompliance Extension', passed: false, error: error.message });
      results.failed++;
    }
    results.totalTests++;
    
    // Test extension registration in framework
    try {
      var extensions = ExtensionFramework.list();
      var hasBrandColors = extensions.some(ext => ext.name === 'BrandColors');
      var hasBrandCompliance = extensions.some(ext => ext.name === 'BrandCompliance');
      
      results.tests.push({ name: 'Example Extensions Registered', passed: hasBrandColors && hasBrandCompliance });
      if (hasBrandColors && hasBrandCompliance) results.passed++; else results.failed++;
    } catch (error) {
      results.tests.push({ name: 'Example Extensions Registered', passed: false, error: error.message });
      results.failed++;
    }
    results.totalTests++;
    
    return {
      success: true,
      summary: `${results.passed}/${results.totalTests} extension tests passed`,
      totalTests: results.totalTests,
      passed: results.passed,
      failed: results.failed,
      details: results.tests
    };
    
  } catch (error) {
    return {
      error: 'Extension test failed: ' + error.message,
      passed: false,
      totalTests: 0
    };
  }
}

/**
 * Combined test runner for web app
 */
function testExtensionSystem() {
  console.log('🧪 Running complete Extension System test suite...');
  
  var frameworkResults = testExtensionFramework();
  var extensionResults = testExampleExtensions();
  
  var combined = {
    success: frameworkResults.success && extensionResults.success,
    summary: `Framework: ${frameworkResults.summary}, Extensions: ${extensionResults.summary}`,
    framework: frameworkResults,
    extensions: extensionResults,
    totalTests: frameworkResults.totalTests + extensionResults.totalTests,
    passed: frameworkResults.passed + extensionResults.passed,
    failed: frameworkResults.failed + extensionResults.failed,
    allTestsPassed: frameworkResults.allTestsPassed && extensionResults.allTestsPassed,
    timestamp: new Date().toISOString()
  };
  
  console.log(`📊 Extension System Test Summary:`);
  console.log(`   Framework: ${frameworkResults.passed}/${frameworkResults.totalTests} passed`);
  console.log(`   Extensions: ${extensionResults.passed}/${extensionResults.totalTests} passed`);
  console.log(`   Overall: ${combined.passed}/${combined.totalTests} passed`);
  
  return combined;
}

/**
 * Quick validation for CI/CD
 */
function validateExtensionSystem() {
  var results = testExtensionSystem();
  return {
    passed: results.allTestsPassed,
    message: results.summary,
    details: {
      frameworkAvailable: results.framework.framework?.available || false,
      extensionCount: results.framework.framework?.extensionCount || 0,
      templatesAvailable: results.framework.framework?.templatesAvailable || false
    }
  };
}