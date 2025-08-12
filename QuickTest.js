/**
 * Quick Test - Simple test file for Google Apps Script environment
 * Run these functions to verify the library is working properly
 * 
 * 🚀 QUICK START: Run oneClickSetup() first!
 */

/**
 * 🎯 ONE-CLICK SETUP - Run this first!
 * Automatically handles API enablement and configuration
 */
function oneClickSetup() {
  console.log('🎯 One-Click OOXML Slides Setup');
  console.log('=' * 32);
  
  try {
    // Check if already set up
    console.log('🔍 Checking current status...');
    
    if (testDriveAPIAccess()) {
      console.log('✅ Already configured! Running quick test...');
      return runQuickTests();
    }
    
    console.log('🔧 Setting up APIs automatically...');
    
    // Try smart API enablement
    const apiResult = smartEnableAPIs();
    
    if (apiResult) {
      console.log('🎉 Setup completed! Running validation tests...');
      return runQuickTests();
    } else {
      console.log('⚠️ Manual setup may be required');
      showQuickSetupHelp();
      return false;
    }
    
  } catch (error) {
    console.error('❌ Setup failed:', error);
    showQuickSetupHelp();
    return false;
  }
}

/**
 * Show quick setup help
 */
function showQuickSetupHelp() {
  console.log('');
  console.log('🆘 QUICK SETUP HELP');
  console.log('=' * 20);
  console.log('');
  console.log('If you see authorization prompts:');
  console.log('✅ Click "Review permissions"');
  console.log('✅ Sign in to your Google account');  
  console.log('✅ Click "Allow" for all permissions');
  console.log('✅ Run oneClickSetup() again');
  console.log('');
  console.log('Alternative quick methods:');
  console.log('🔧 authorizeDirectly() - Force authorization');
  console.log('🧠 smartEnableAPIs() - Advanced API enablement');
  console.log('📋 showManualSetupInstructions() - Manual steps');
}

/**
 * Test 1: Library Loading Test
 * Tests if all classes can be instantiated
 */
function testLibraryLoading() {
  try {
    console.log('🧪 Testing library loading...');
    
    // Test if classes are available
    const tests = [
      { name: 'OOXMLSlides', test: () => typeof OOXMLSlides !== 'undefined' },
      { name: 'OOXMLParser', test: () => typeof OOXMLParser !== 'undefined' },
      { name: 'ThemeEditor', test: () => typeof ThemeEditor !== 'undefined' },
      { name: 'SlideManager', test: () => typeof SlideManager !== 'undefined' },
      { name: 'FileHandler', test: () => typeof FileHandler !== 'undefined' },
      { name: 'Validators', test: () => typeof Validators !== 'undefined' },
      { name: 'ErrorHandler', test: () => typeof ErrorHandler !== 'undefined' }
    ];
    
    let passed = 0;
    tests.forEach(({ name, test }) => {
      try {
        const result = test();
        console.log(`${result ? '✅' : '❌'} ${name}: ${result ? 'Available' : 'Missing'}`);
        if (result) passed++;
      } catch (error) {
        console.log(`❌ ${name}: Error - ${error.message}`);
      }
    });
    
    console.log(`📊 Library Loading: ${passed}/${tests.length} classes available`);
    
    if (passed === tests.length) {
      console.log('🎉 All library classes loaded successfully!');
      return true;
    } else {
      console.log('⚠️ Some library classes are missing');
      return false;
    }
    
  } catch (error) {
    console.error('❌ Library loading test failed:', error);
    return false;
  }
}

/**
 * Test 2: Validation Functions Test
 * Tests the validation utilities without requiring file access
 */
function testValidation() {
  try {
    console.log('🧪 Testing validation functions...');
    
    // Test color validation
    const colorTests = [
      { color: '#FF0000', expected: true },
      { color: '#f00', expected: true },
      { color: 'FF0000', expected: true },
      { color: 'invalid', expected: false },
      { color: '#GGGGGG', expected: false }
    ];
    
    let colorsPassed = 0;
    colorTests.forEach(({ color, expected }) => {
      const result = Validators.isValidHexColor(color);
      console.log(`Color ${color}: ${result} (expected: ${expected})`);
      if (result === expected) colorsPassed++;
    });
    
    // Test file ID validation
    const fileIdTests = [
      { id: '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms', expected: true },
      { id: 'short', expected: false },
      { id: 'invalid file id with spaces', expected: false }
    ];
    
    let fileIdsPassed = 0;
    fileIdTests.forEach(({ id, expected }) => {
      const result = Validators.isValidFileId(id);
      console.log(`File ID validation: ${result} (expected: ${expected})`);
      if (result === expected) fileIdsPassed++;
    });
    
    console.log(`📊 Validation Tests: ${colorsPassed + fileIdsPassed}/${colorTests.length + fileIdTests.length} passed`);
    
    if (colorsPassed === colorTests.length && fileIdsPassed === fileIdTests.length) {
      console.log('✅ All validation tests passed!');
      return true;
    } else {
      console.log('⚠️ Some validation tests failed');
      return false;
    }
    
  } catch (error) {
    console.error('❌ Validation test failed:', error);
    return false;
  }
}

/**
 * Test 3: FileHandler Access Test
 * Tests if Google Drive API access is configured properly
 */
function testDriveAccess() {
  try {
    console.log('🧪 Testing Drive API access...');
    
    // Test if Drive API is available
    const driveTest = typeof DriveApp !== 'undefined';
    console.log(`DriveApp available: ${driveTest}`);
    
    if (driveTest) {
      // Test basic Drive operations (that don't require specific files)
      const fileHandler = new FileHandler();
      console.log('✅ FileHandler instantiated successfully');
      
      // Test file name sanitization
      const sanitizedName = Validators.sanitizeFileName('Test<>File|Name?.pptx');
      console.log(`File name sanitization: "${sanitizedName}"`);
      
      console.log('✅ Drive access test passed');
      return true;
    } else {
      console.log('❌ DriveApp not available - check API permissions');
      return false;
    }
    
  } catch (error) {
    console.error('❌ Drive access test failed:', error);
    console.log('💡 Make sure to enable Drive API in Advanced Google Services');
    return false;
  }
}

/**
 * Test 4: OOXMLSlides Class Test
 * Tests if the main class can be instantiated (without file operations)
 */
function testOOXMLSlidesClass() {
  try {
    console.log('🧪 Testing OOXMLSlides class...');
    
    // Test static methods
    const presets = OOXMLSlides.Presets;
    console.log(`Presets available: ${Object.keys(presets).join(', ')}`);
    
    // Test if we can access the preset functions
    const hasPresets = typeof presets.modernTheme === 'function';
    console.log(`Modern theme preset: ${hasPresets ? 'Available' : 'Missing'}`);
    
    if (hasPresets) {
      console.log('✅ OOXMLSlides class test passed');
      return true;
    } else {
      console.log('❌ OOXMLSlides class test failed');
      return false;
    }
    
  } catch (error) {
    console.error('❌ OOXMLSlides class test failed:', error);
    return false;
  }
}

/**
 * Run All Quick Tests
 * Runs all tests and provides a summary
 */
function runQuickTests() {
  console.log('🚀 Starting Quick Tests for OOXML Slides Library');
  console.log('=' * 50);
  
  const testResults = [];
  
  // Run all tests
  testResults.push({ name: 'Library Loading', passed: testLibraryLoading() });
  console.log('');
  
  testResults.push({ name: 'Validation Functions', passed: testValidation() });
  console.log('');
  
  testResults.push({ name: 'Drive Access', passed: testDriveAccess() });
  console.log('');
  
  testResults.push({ name: 'OOXMLSlides Class', passed: testOOXMLSlidesClass() });
  console.log('');
  
  // Summary
  console.log('📊 QUICK TEST RESULTS');
  console.log('=' * 30);
  
  let totalPassed = 0;
  testResults.forEach(result => {
    console.log(`${result.passed ? '✅' : '❌'} ${result.name}`);
    if (result.passed) totalPassed++;
  });
  
  console.log(`\nOverall: ${totalPassed}/${testResults.length} tests passed`);
  
  if (totalPassed === testResults.length) {
    console.log('🎉 All quick tests passed! Library is ready to use.');
    console.log('\n💡 Next steps:');
    console.log('1. Update TEST_CONFIG in basic-test.js with your file IDs');
    console.log('2. Run runAllTests() for comprehensive testing');
    console.log('3. Try the examples in usage-examples.js');
  } else {
    console.log('⚠️ Some tests failed. Check the output above for details.');
    console.log('\n🔧 Troubleshooting:');
    console.log('1. Ensure Drive API is enabled in Advanced Google Services');
    console.log('2. Check that all library files were uploaded correctly');
    console.log('3. Verify OAuth scopes are properly configured');
  }
  
  return { passed: totalPassed, total: testResults.length, results: testResults };
}

/**
 * Quick Setup Check
 * Provides setup instructions and checks
 */
function setupCheck() {
  console.log('🔧 OOXML Slides Library Setup Check');
  console.log('=' * 40);
  
  console.log('\n📋 Required Setup Steps:');
  console.log('1. ✅ Library files uploaded via Clasp');
  console.log('2. ⚠️  Enable Drive API in Advanced Google Services:');
  console.log('   - Go to Resources > Advanced Google Services');
  console.log('   - Enable Google Drive API v3');
  console.log('   - Also enable it in Google Cloud Console');
  console.log('3. ⚠️  OAuth Scopes configured in appsscript.json:');
  console.log('   - https://www.googleapis.com/auth/drive');
  console.log('   - https://www.googleapis.com/auth/drive.file');
  console.log('   - https://www.googleapis.com/auth/presentations');
  
  console.log('\n🎯 To start using the library:');
  console.log('1. Run runQuickTests() to verify everything works');
  console.log('2. Get a PowerPoint file ID from Google Drive');
  console.log('3. Try: OOXMLSlides.fromFile("your-file-id").getTheme()');
  
  console.log('\n📚 Documentation:');
  console.log('- Check usage-examples.js for complete examples');
  console.log('- See basic-test.js for comprehensive test suite');
  console.log('- GitHub: https://github.com/your-repo/ooxml-slides-manipulator');
}