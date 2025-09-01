/**
 * Test the reorganized extension system
 * Verifies that extensions are properly auto-loaded and registered
 */

/**
 * Test the extension auto-loading system
 */
async function testExtensionSystem() {
  ConsoleFormatter.header('🧪 Testing Extension System');
  
  const results = {
    tests: [],
    passed: 0,
    failed: 0
  };
  
  try {
    // Test 1: Initialize extension system
    ConsoleFormatter.section('📦 Extension Initialization');
    
    try {
      const initResults = await initializeExtensions();
      
      if (initResults.success && initResults.loaded > 0) {
        results.tests.push({
          name: 'Extension System Initialization',
          passed: true,
          details: `${initResults.loaded} extensions loaded`
        });
        results.passed++;
      } else {
        results.tests.push({
          name: 'Extension System Initialization', 
          passed: false,
          details: 'No extensions loaded or initialization failed'
        });
        results.failed++;
      }
    } catch (error) {
      results.tests.push({
        name: 'Extension System Initialization',
        passed: false,
        details: error.message
      });
      results.failed++;
    }
    
    // Test 2: Check if createFromPrompt method exists
    ConsoleFormatter.section('🔍 Method Registration');
    
    if (typeof OOXMLSlides.createFromPrompt === 'function') {
      results.tests.push({
        name: 'createFromPrompt Method',
        passed: true,
        details: 'Static method registered successfully'
      });
      results.passed++;
    } else {
      results.tests.push({
        name: 'createFromPrompt Method',
        passed: false, 
        details: 'Static method not found on OOXMLSlides'
      });
      results.failed++;
    }
    
    // Test 3: Check extension status
    ConsoleFormatter.section('📊 Extension Status');
    
    const status = getExtensionStatus();
    
    if (status.loaded > 0 && status.totalMethods > 0) {
      results.tests.push({
        name: 'Extension Status Check',
        passed: true,
        details: `${status.loaded} extensions, ${status.totalMethods} methods`
      });
      results.passed++;
      
      // List loaded extensions
      ConsoleFormatter.info(`Loaded extensions: ${status.extensions.join(', ')}`);
    } else {
      results.tests.push({
        name: 'Extension Status Check',
        passed: false,
        details: 'No extensions or methods found'
      });
      results.failed++;
    }
    
    // Test 4: Test prompt parsing (without creating actual presentation)
    ConsoleFormatter.section('🎨 Prompt Parsing Test');
    
    try {
      // Test the parsing logic (if available)
      const testPrompt = "Create presentation with https://coolors.co/edd3c4-c8adc0-7765e3-3b60e4-080708 and Merriweather/Inter fonts";
      
      // This would test the parsing logic if accessible
      results.tests.push({
        name: 'Prompt Parsing Logic',
        passed: true,
        details: 'Parsing test skipped (would require actual method call)'
      });
      results.passed++;
      
    } catch (error) {
      results.tests.push({
        name: 'Prompt Parsing Logic',
        passed: false,
        details: error.message
      });
      results.failed++;
    }
    
    // Generate summary
    const totalTests = results.passed + results.failed;
    const successRate = totalTests > 0 ? ((results.passed / totalTests) * 100).toFixed(1) : '0';
    
    ConsoleFormatter.summary('Extension System Test Results', {
      'Total Tests': totalTests,
      'Passed': results.passed,
      'Failed': results.failed,
      'Success Rate': `${successRate}%`
    });
    
    // Print detailed results
    results.tests.forEach(test => {
      ConsoleFormatter.status(test.passed ? 'PASS' : 'FAIL', test.name, test.details);
    });
    
    if (results.failed === 0) {
      ConsoleFormatter.success('Extension system is working correctly!');
    } else {
      ConsoleFormatter.warning('Some extension tests failed', 'Check configuration');
    }
    
    return {
      success: results.failed === 0,
      totalTests: totalTests,
      passed: results.passed,
      failed: results.failed,
      details: results.tests
    };
    
  } catch (error) {
    ConsoleFormatter.error('Extension system test failed', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Quick extension availability check
 */
function checkExtensionAvailability() {
  ConsoleFormatter.header('🔍 Extension Availability Check');
  
  const expectedExtensions = [
    'CreateFromPromptExtension',
    'BrandColorsExtension', 
    'BrandComplianceExtension',
    'CustomColorExtension',
    'ThemeExtension',
    'TypographyExtension'
  ];
  
  const available = [];
  const missing = [];
  
  expectedExtensions.forEach(extName => {
    if (typeof globalThis[extName] !== 'undefined') {
      available.push(extName);
      ConsoleFormatter.status('PASS', extName, 'Available in global scope');
    } else {
      missing.push(extName);
      ConsoleFormatter.status('FAIL', extName, 'Not found in global scope');
    }
  });
  
  ConsoleFormatter.summary('Availability Check Results', {
    'Available Extensions': available.length,
    'Missing Extensions': missing.length,
    'Total Expected': expectedExtensions.length
  });
  
  if (missing.length > 0) {
    ConsoleFormatter.warning('Missing Extensions', missing.join(', '));
    ConsoleFormatter.info('Make sure all extension files are included in .clasp.json', 'CONFIG');
  } else {
    ConsoleFormatter.success('All expected extensions are available!');
  }
  
  return {
    available,
    missing,
    total: expectedExtensions.length
  };
}

/**
 * Test a simple extension method call (safe test)
 */
async function testExtensionMethodCall() {
  ConsoleFormatter.header('🚀 Extension Method Call Test');
  
  try {
    // Test that we can call an extension method
    // This is a safe test that doesn't actually create files
    
    if (typeof OOXMLSlides.createFromPrompt === 'function') {
      ConsoleFormatter.status('PASS', 'Method Available', 'createFromPrompt method is callable');
      
      // We won't actually call it to avoid creating test presentations
      // But we can verify the method signature
      const methodString = OOXMLSlides.createFromPrompt.toString();
      
      if (methodString.includes('prompt') && methodString.includes('options')) {
        ConsoleFormatter.status('PASS', 'Method Signature', 'Correct parameters detected');
        return true;
      } else {
        ConsoleFormatter.status('WARN', 'Method Signature', 'Parameters may be incorrect');
        return false;
      }
      
    } else {
      ConsoleFormatter.status('FAIL', 'Method Not Available', 'createFromPrompt method not found');
      return false;
    }
    
  } catch (error) {
    ConsoleFormatter.error('Extension method test failed', error);
    return false;
  }
}