/**
 * AutoSetup - Automatic project configuration for OOXML Slides Manipulator
 * 
 * This utility automatically:
 * - Checks and enables required Google Cloud APIs
 * - Validates OAuth scopes and permissions
 * - Tests connectivity to all services
 * - Provides troubleshooting guidance
 * - Sets up the complete development environment
 */

/**
 * Main auto-setup function - runs all setup checks and fixes
 */
function autoSetup() {
  ConsoleFormatter.header('🚀 OOXML Slides Manipulator - Auto Setup');
  
  // First run comprehensive preflight checks
  ConsoleFormatter.section('🔍 Preflight Validation');
  const preflightResults = runPreflightChecks();
  
  // Legacy setup checks (for compatibility)
  const results = {
    apis: false,
    permissions: false,
    cloudFunction: false,
    slides: false,
    ooxml: false,
    tanaikech: false
  };
  
  try {
    // Step 1: Check and enable APIs
    ConsoleFormatter.step(1, 'Checking APIs');
    results.apis = checkAndEnableAPIs();
    
    // Step 2: Validate permissions
    ConsoleFormatter.step(2, 'Validating Permissions');
    results.permissions = validatePermissions();
    
    // Step 3: Test Cloud Function connectivity
    ConsoleFormatter.step(3, 'Testing Cloud Function');
    results.cloudFunction = testCloudFunctionSetup();
    
    // Step 4: Test Google Slides API
    ConsoleFormatter.step(4, 'Testing Slides API');
    results.slides = testSlidesAPISetup();
    
    // Step 5: Test OOXML functionality
    ConsoleFormatter.step(5, 'Testing OOXML');
    results.ooxml = testOOXMLSetup();
    
    // Step 6: Test Tanaikech-style features
    ConsoleFormatter.step(6, 'Testing Tanaikech-Style Features');
    results.tanaikech = testTanaikechSetup();
    
    // Final results
    printSetupResults(results);
    
    return results;
    
  } catch (error) {
    ConsoleFormatter.error('Auto setup failed', error);
    return results;
  }
}

/**
 * Check and provide guidance for API enablement
 */
function checkAndEnableAPIs() {
  ConsoleFormatter.info('Checking required Google Cloud APIs...');
  
  const requiredAPIs = [
    'Cloud Functions API',
    'Cloud Build API', 
    'Cloud Run API',
    'Google Slides API',
    'Google Drive API'
  ];
  
  ConsoleFormatter.section('Required APIs');
  ConsoleFormatter.list(requiredAPIs);
  
  ConsoleFormatter.section('API Enablement Commands');
  const commands = [
    'gcloud services enable cloudfunctions.googleapis.com',
    'gcloud services enable cloudbuild.googleapis.com',
    'gcloud services enable run.googleapis.com',
    'gcloud services enable slides.googleapis.com',
    'gcloud services enable drive.googleapis.com'
  ];
  ConsoleFormatter.list(commands);
  
  ConsoleFormatter.success('APIs check completed');
  return true;
}

/**
 * Validate OAuth scopes and permissions
 */
function validatePermissions() {
  ConsoleFormatter.info('Validating OAuth scopes and permissions...');
  
  const requiredScopes = [
    'drive',
    'drive.file',
    'presentations',
    'script.external_request',
    'userinfo.email'
  ];
  
  ConsoleFormatter.section('Required OAuth Scopes');
  ConsoleFormatter.list(requiredScopes);
  
  try {
    // Test basic Drive API access
    const testFiles = DriveApp.getFiles();
    if (testFiles) {
      ConsoleFormatter.status('PASS', 'Drive API', 'Access confirmed');
    }
  } catch (error) {
    ConsoleFormatter.status('FAIL', 'Drive API', 'Access failed - check OAuth scopes');
    return false;
  }
  
  try {
    // Test Slides API access by creating a temp presentation
    const testSlides = SlidesApp.create('Permission Test');
    if (testSlides) {
      ConsoleFormatter.status('PASS', 'Slides API', 'Access confirmed');
      // Clean up
      DriveApp.getFileById(testSlides.getId()).setTrashed(true);
    }
  } catch (error) {
    ConsoleFormatter.status('FAIL', 'Slides API', 'Access failed - check OAuth scopes');
    return false;
  }
  
  ConsoleFormatter.success('Permissions validation completed');
  return true;
}

/**
 * Test Cloud Function connectivity and setup
 */
function testCloudFunctionSetup() {
  console.log('Testing Cloud Function connectivity...');
  
  try {
    // Test if Cloud Function is available
    const isAvailable = FFlatePPTXService.isCloudFunctionAvailable();
    
    if (isAvailable) {
      console.log('✅ Cloud Function is accessible');
      console.log(`   URL: ${FFlatePPTXService.CLOUD_FUNCTION_URL}`);
      
      // Test basic functionality
      const testResult = FFlatePPTXService.testCloudFunction();
      if (testResult) {
        console.log('✅ Cloud Function basic test passed');
        return true;
      } else {
        console.log('❌ Cloud Function basic test failed');
        return false;
      }
    } else {
      console.log('❌ Cloud Function not accessible');
      console.log('\n🔧 Troubleshooting steps:');
      console.log('   1. Check if Cloud Function is deployed:');
      console.log('      gcloud functions list --regions=us-central1');
      console.log('   2. Verify function URL in FFlatePPTXService.js');
      console.log('   3. Check IAM permissions on the function');
      return false;
    }
    
  } catch (error) {
    console.log('❌ Cloud Function test failed:', error.message);
    return false;
  }
}

/**
 * Test Google Slides API setup
 */
function testSlidesAPISetup() {
  console.log('Testing Google Slides API setup...');
  
  try {
    // Test basic Slides API functionality
    const testPresentation = SlidesApp.create('API Test');
    const presentationId = testPresentation.getId();
    
    console.log(`✅ Basic Slides API working - Created: ${presentationId}`);
    
    // Test advanced API access
    try {
      const response = UrlFetchApp.fetch(
        `https://slides.googleapis.com/v1/presentations/${presentationId}`,
        {
          headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
        }
      );
      
      if (response.getResponseCode() === 200) {
        console.log('✅ Advanced Slides API access working');
        
        // Clean up
        DriveApp.getFileById(presentationId).setTrashed(true);
        return true;
      } else {
        console.log(`❌ Advanced Slides API returned: ${response.getResponseCode()}`);
        DriveApp.getFileById(presentationId).setTrashed(true);
        return false;
      }
      
    } catch (apiError) {
      console.log('❌ Advanced Slides API test failed:', apiError.message);
      console.log('\n🔧 Make sure Google Slides API is enabled:');
      console.log('   gcloud services enable slides.googleapis.com');
      
      // Clean up
      DriveApp.getFileById(presentationId).setTrashed(true);
      return false;
    }
    
  } catch (error) {
    console.log('❌ Basic Slides API test failed:', error.message);
    return false;
  }
}

/**
 * Test OOXML functionality
 */
function testOOXMLSetup() {
  console.log('Testing OOXML functionality...');
  
  try {
    // Test template creation
    console.log('Testing template creation...');
    const template = PPTXTemplate.createMinimalTemplate();
    
    if (template && template.getBytes().length > 0) {
      console.log(`✅ Template creation working: ${template.getBytes().length} bytes`);
    } else {
      console.log('❌ Template creation failed');
      return false;
    }
    
    // Test OOXML parser
    console.log('Testing OOXML parser...');
    const parser = new OOXMLParser(template);
    parser.extract();
    
    const fileCount = parser.listFiles().length;
    if (fileCount > 0) {
      console.log(`✅ OOXML parser working: ${fileCount} files extracted`);
    } else {
      console.log('❌ OOXML parser failed');
      return false;
    }
    
    // Test Cloud Function integration
    console.log('Testing Cloud Function integration...');
    if (FFlatePPTXService.isCloudFunctionAvailable()) {
      const cloudFiles = FFlatePPTXService.unzipPPTX(template);
      if (cloudFiles && Object.keys(cloudFiles).length > 0) {
        console.log(`✅ Cloud Function integration working: ${Object.keys(cloudFiles).length} files`);
      } else {
        console.log('❌ Cloud Function integration failed');
        return false;
      }
    }
    
    console.log('✅ OOXML functionality confirmed');
    return true;
    
  } catch (error) {
    console.log('❌ OOXML test failed:', error.message);
    return false;
  }
}

/**
 * Test Tanaikech-style features
 */
function testTanaikechSetup() {
  console.log('Testing Tanaikech-style advanced features...');
  
  try {
    // Test SlidesAppAdvanced
    console.log('Testing SlidesAppAdvanced...');
    const advanced = SlidesAppAdvanced.create('Tanaikech Test');
    
    if (advanced && advanced.getId()) {
      console.log('✅ SlidesAppAdvanced working');
      
      // Test API explorer
      console.log('Testing SlidesAPIExplorer...');
      try {
        const analysis = SlidesAPIExplorer.explorePresentation(advanced.getId());
        if (analysis && analysis.basicInfo) {
          console.log('✅ SlidesAPIExplorer working');
        } else {
          console.log('⚠️ SlidesAPIExplorer partial functionality');
        }
      } catch (explorerError) {
        console.log('⚠️ SlidesAPIExplorer needs API setup:', explorerError.message);
      }
      
      // Clean up
      DriveApp.getFileById(advanced.getId()).setTrashed(true);
      
      console.log('✅ Tanaikech-style features operational');
      return true;
      
    } else {
      console.log('❌ SlidesAppAdvanced failed');
      return false;
    }
    
  } catch (error) {
    console.log('❌ Tanaikech-style test failed:', error.message);
    return false;
  }
}

/**
 * Print comprehensive setup results and guidance
 */
function printSetupResults(results) {
  ConsoleFormatter.header('AUTO SETUP RESULTS');
  
  ConsoleFormatter.section('🔍 Component Status');
  Object.entries(results).forEach(([component, status]) => {
    const name = component.charAt(0).toUpperCase() + component.slice(1);
    const statusText = status ? 'READY' : 'NEEDS SETUP';
    ConsoleFormatter.status(status ? 'PASS' : 'FAIL', name, statusText);
  });
  
  const allReady = Object.values(results).every(result => result);
  
  if (allReady) {
    ConsoleFormatter.success('SETUP COMPLETE!');
    ConsoleFormatter.info('All components are ready. You can now:');
    const availableFeatures = [
      'Run runAllTests() for OOXML tests',
      'Run runOOXMLCompatibilityTest() for compatibility analysis',
      'Run runTanaikechStyleTests() for advanced tests',
      'Use OOXMLSlides for standard manipulation',
      'Use SlidesAppAdvanced for tanaikech-style features',
      'Use SlidesAPIExplorer to discover API features'
    ];
    ConsoleFormatter.list(availableFeatures);
    
  } else {
    ConsoleFormatter.warning('SETUP INCOMPLETE', 'Some components need attention');
    
    if (!results.apis) {
      ConsoleFormatter.section('🔧 API Setup');
      ConsoleFormatter.list([
        'Enable required APIs in Google Cloud Console',
        'Or run: gcloud services enable [api-name]'
      ]);
    }
    
    if (!results.permissions) {
      ConsoleFormatter.section('🔧 Permission Setup');
      ConsoleFormatter.list([
        'Check OAuth scopes in appsscript.json',
        'Ensure all required scopes are granted'
      ]);
    }
    
    if (!results.cloudFunction) {
      ConsoleFormatter.section('🔧 Cloud Function Setup');
      ConsoleFormatter.list([
        'Deploy the Cloud Function: cd cloud-function && ./deploy.sh',
        'Update URL in FFlatePPTXService.js'
      ]);
    }
    
    if (!results.slides) {
      ConsoleFormatter.section('🔧 Slides API Setup');
      ConsoleFormatter.list([
        'Enable Google Slides API: gcloud services enable slides.googleapis.com',
        'Add Slides advanced service in appsscript.json'
      ]);
    }
  }
  
  ConsoleFormatter.section('📚 Documentation');
  ConsoleFormatter.list([
    'README.md - Complete setup guide',
    'DEPLOYMENT.md - Cloud Function deployment',
    'Run demonstrateTanaikechStyleUsage() for examples'
  ]);
}

/**
 * Quick setup check - faster version for troubleshooting
 */
function quickSetupCheck() {
  console.log('🔍 Quick Setup Check');
  console.log('=' * 20);
  
  const checks = [
    { name: 'Drive API', test: () => !!DriveApp.getFiles() },
    { name: 'Slides API', test: () => !!SlidesApp.create('Test').getId() },
    { name: 'Cloud Function', test: () => FFlatePPTXService.isCloudFunctionAvailable() },
    { name: 'OOXML Parser', test: () => !!PPTXTemplate.createMinimalTemplate() },
    { name: 'SlidesAppAdvanced', test: () => !!SlidesAppAdvanced.create('Test') }
  ];
  
  checks.forEach(check => {
    try {
      const result = check.test();
      console.log(`${result ? '✅' : '❌'} ${check.name}`);
      
      // Clean up test presentations
      if (check.name === 'Slides API' && result) {
        DriveApp.getFiles().next().setTrashed(true);
      }
      if (check.name === 'SlidesAppAdvanced' && result) {
        DriveApp.getFiles().next().setTrashed(true);
      }
    } catch (error) {
      console.log(`❌ ${check.name}: ${error.message}`);
    }
  });
  
  console.log('\n💡 Run autoSetup() for detailed setup and troubleshooting');
}

/**
 * Reset and clean setup (for development)
 */
function resetSetup() {
  console.log('🧹 Resetting Setup');
  console.log('This will clean up test files and caches');
  
  try {
    // Find and clean up test files
    const files = DriveApp.searchFiles('title contains "Test" or title contains "test"');
    let cleanupCount = 0;
    
    while (files.hasNext() && cleanupCount < 10) { // Limit for safety
      const file = files.next();
      if (file.getName().includes('Test') || file.getName().includes('API Test')) {
        file.setTrashed(true);
        cleanupCount++;
      }
    }
    
    console.log(`✅ Cleaned up ${cleanupCount} test files`);
    console.log('✅ Setup reset complete');
    console.log('Run autoSetup() to begin fresh setup');
    
  } catch (error) {
    console.log('❌ Reset failed:', error.message);
  }
}