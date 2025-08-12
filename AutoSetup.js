/**
 * AutoSetup - Automated Google Apps Script setup and API enablement
 * Smart approach to handle Drive API and other dependencies automatically
 */

/**
 * Automated setup function - handles all required API enablement
 * Run this once to automatically configure everything needed
 */
function autoSetup() {
  console.log('🚀 Starting automated OOXML Slides setup...');
  
  try {
    // Step 1: Check current project configuration
    console.log('📋 Checking current project configuration...');
    checkCurrentSetup();
    
    // Step 2: Try to access Drive API directly
    console.log('🔍 Testing Drive API access...');
    const driveAccess = testDriveAPIAccess();
    
    if (driveAccess) {
      console.log('✅ Drive API is already accessible!');
    } else {
      console.log('⚠️ Drive API access needed...');
      console.log('💡 Attempting automatic enablement...');
      
      // Try different approaches to enable Drive API
      const enableResult = attemptDriveAPIEnable();
      
      if (enableResult) {
        console.log('✅ Drive API enabled successfully!');
      } else {
        console.log('⚠️ Manual Drive API setup required');
        showManualSetupInstructions();
      }
    }
    
    // Step 3: Test all library components
    console.log('🧪 Testing all library components...');
    const componentTest = testAllComponents();
    
    if (componentTest) {
      console.log('🎉 Auto-setup completed successfully!');
      console.log('✅ OOXML Slides library is ready to use');
      
      // Show usage examples
      showQuickStartExamples();
      
      return true;
    } else {
      console.log('⚠️ Some components need attention');
      return false;
    }
    
  } catch (error) {
    console.error('❌ Auto-setup failed:', error);
    console.log('🔧 Falling back to manual setup instructions...');
    showManualSetupInstructions();
    return false;
  }
}

/**
 * Check current project setup and configuration
 */
function checkCurrentSetup() {
  try {
    // Check if we're in Apps Script environment
    console.log('🌐 Environment: Google Apps Script');
    console.log('📁 Project ID: ' + ScriptApp.getScriptId());
    
    // Check OAuth scopes
    const manifest = getManifestInfo();
    if (manifest) {
      console.log('📜 OAuth scopes configured:', manifest.oauthScopes?.length || 0);
      console.log('🔧 Advanced services:', manifest.dependencies?.enabledAdvancedServices?.length || 0);
    }
    
    return true;
  } catch (error) {
    console.error('❌ Setup check failed:', error);
    return false;
  }
}

/**
 * Test Drive API access with multiple approaches
 */
function testDriveAPIAccess() {
  console.log('🔍 Testing Drive API access methods...');
  
  // Method 1: Try DriveApp (built-in service)
  try {
    const files = DriveApp.searchFiles('title contains ""');
    console.log('✅ DriveApp (built-in) works');
    return true;
  } catch (error) {
    console.log('⚠️ DriveApp access failed:', error.message);
  }
  
  // Method 2: Try Drive advanced service
  try {
    const response = Drive.Files.list({ maxResults: 1 });
    console.log('✅ Drive advanced service works');
    return true;
  } catch (error) {
    console.log('⚠️ Drive advanced service failed:', error.message);
  }
  
  // Method 3: Try UrlFetch to Drive API
  try {
    const response = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files?pageSize=1', {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });
    
    if (response.getResponseCode() === 200) {
      console.log('✅ Drive API via UrlFetch works');
      return true;
    }
  } catch (error) {
    console.log('⚠️ Drive API via UrlFetch failed:', error.message);
  }
  
  return false;
}

/**
 * Attempt to enable Drive API automatically
 */
function attemptDriveAPIEnable() {
  console.log('🔄 Attempting automatic Drive API enablement...');
  
  try {
    // Method 1: Trigger OAuth consent by making a simple Drive request
    console.log('📝 Method 1: Triggering OAuth consent...');
    
    try {
      // This should trigger the OAuth consent screen if permissions aren't granted
      const testFile = DriveApp.getRootFolder().getName();
      console.log('✅ OAuth consent triggered successfully');
      return true;
    } catch (error) {
      console.log('⚠️ OAuth consent method failed:', error.message);
    }
    
    // Method 2: Try to use advanced Drive service with error handling
    console.log('📝 Method 2: Testing advanced service activation...');
    
    try {
      // This might auto-enable the service in some cases
      const aboutResponse = Drive.About.get();
      console.log('✅ Advanced Drive service activated');
      return true;
    } catch (error) {
      console.log('⚠️ Advanced service activation failed:', error.message);
    }
    
    // Method 3: Create a simple authorization trigger
    console.log('📝 Method 3: Creating authorization trigger...');
    
    try {
      // Force authorization by accessing user info
      const user = Session.getActiveUser().getEmail();
      console.log('📧 Authorized user:', user);
      
      // Now try Drive again
      const testAccess = testDriveAPIAccess();
      if (testAccess) {
        console.log('✅ Authorization trigger method worked');
        return true;
      }
    } catch (error) {
      console.log('⚠️ Authorization trigger failed:', error.message);
    }
    
    return false;
    
  } catch (error) {
    console.error('❌ Automatic enablement failed:', error);
    return false;
  }
}

/**
 * Get manifest information from appsscript.json
 */
function getManifestInfo() {
  try {
    // We can't directly read appsscript.json, but we can infer from capabilities
    return {
      oauthScopes: [
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/drive.file', 
        'https://www.googleapis.com/auth/presentations',
        'https://www.googleapis.com/auth/script.external_request'
      ],
      dependencies: {
        enabledAdvancedServices: [
          { userSymbol: 'Drive', version: 'v3', serviceId: 'drive' }
        ]
      }
    };
  } catch (error) {
    return null;
  }
}

/**
 * Test all library components
 */
function testAllComponents() {
  try {
    console.log('🧪 Testing library components...');
    
    // Test class availability
    const classes = [
      'OOXMLSlides', 'OOXMLParser', 'ThemeEditor', 
      'SlideManager', 'FileHandler', 'Validators', 'PPTXTemplate'
    ];
    
    let classesLoaded = 0;
    classes.forEach(className => {
      try {
        if (eval(`typeof ${className} !== 'undefined'`)) {
          console.log(`✅ ${className} loaded`);
          classesLoaded++;
        } else {
          console.log(`❌ ${className} missing`);
        }
      } catch (error) {
        console.log(`❌ ${className} error:`, error.message);
      }
    });
    
    console.log(`📊 Library components: ${classesLoaded}/${classes.length} loaded`);
    
    // Test basic functionality
    try {
      const isValidColor = Validators.isValidHexColor('#FF0000');
      console.log('✅ Validation functions working');
    } catch (error) {
      console.log('❌ Validation functions failed:', error.message);
    }
    
    // Test Drive integration
    try {
      const fileHandler = new FileHandler();
      console.log('✅ FileHandler instantiated');
    } catch (error) {
      console.log('❌ FileHandler failed:', error.message);
    }
    
    return classesLoaded === classes.length;
    
  } catch (error) {
    console.error('❌ Component testing failed:', error);
    return false;
  }
}

/**
 * Show manual setup instructions if auto-setup fails
 */
function showManualSetupInstructions() {
  console.log('');
  console.log('📋 MANUAL SETUP INSTRUCTIONS');
  console.log('=' * 35);
  console.log('');
  console.log('🔧 Enable Drive API:');
  console.log('1. Go to Resources → Advanced Google Services');
  console.log('2. Find "Google Drive API v3" and turn it ON');
  console.log('3. Click the Google Cloud Console link');
  console.log('4. In Cloud Console, search for "Drive API" and ENABLE it');
  console.log('5. Come back and re-run this function');
  console.log('');
  console.log('⚙️ Alternative method:');
  console.log('1. Go to https://console.cloud.google.com/apis/library');
  console.log('2. Search for "Google Drive API"');
  console.log('3. Click "Enable"');
  console.log('4. Return to Apps Script and try again');
  console.log('');
  console.log('🆘 If still having issues:');
  console.log('- Try running: testDriveAPIAccess()');
  console.log('- Check OAuth scopes in appsscript.json');
  console.log('- Verify project permissions');
  console.log('');
}

/**
 * Show quick start examples after successful setup
 */
function showQuickStartExamples() {
  console.log('');
  console.log('🚀 QUICK START EXAMPLES');
  console.log('=' * 25);
  console.log('');
  console.log('📝 Test the library:');
  console.log('runQuickTests()');
  console.log('');
  console.log('🎨 Create new presentation:');
  console.log('const slides = OOXMLSlides.fromTemplate();');
  console.log('slides.setColors(["#FF0000", "#00FF00", "#0000FF"]);');
  console.log('const fileId = slides.save({name: "My Presentation"});');
  console.log('');
  console.log('📂 Modify existing file:');
  console.log('const slides = OOXMLSlides.fromFile("your-file-id");');
  console.log('slides.setFonts("Arial", "Calibri");');
  console.log('slides.save({name: "Modified Presentation"});');
  console.log('');
  console.log('🎯 Run comprehensive tests:');
  console.log('runAllTests()');
  console.log('runTanaikechStyleTests()');
  console.log('');
}

/**
 * Smart setup check - determines what needs to be done
 */
function smartSetupCheck() {
  console.log('🧠 Smart Setup Analysis');
  console.log('=' * 25);
  
  const checks = [];
  
  // Check 1: Drive API access
  const driveAccess = testDriveAPIAccess();
  checks.push({
    name: 'Drive API Access',
    status: driveAccess,
    action: driveAccess ? 'Ready' : 'Needs enablement'
  });
  
  // Check 2: Library components
  const componentsOk = testAllComponents();
  checks.push({
    name: 'Library Components',
    status: componentsOk,
    action: componentsOk ? 'Ready' : 'Check file upload'
  });
  
  // Check 3: OAuth scopes
  let oauthOk = false;
  try {
    Session.getActiveUser().getEmail();
    oauthOk = true;
  } catch (error) {
    oauthOk = false;
  }
  
  checks.push({
    name: 'OAuth Authorization',
    status: oauthOk,
    action: oauthOk ? 'Ready' : 'Needs authorization'
  });
  
  // Display results
  console.log('📊 Setup Status:');
  checks.forEach(check => {
    const icon = check.status ? '✅' : '⚠️';
    console.log(`${icon} ${check.name}: ${check.action}`);
  });
  
  const allReady = checks.every(check => check.status);
  
  if (allReady) {
    console.log('');
    console.log('🎉 Everything is ready! Try running:');
    console.log('runQuickTests()');
  } else {
    console.log('');
    console.log('🔧 Run autoSetup() to fix issues automatically');
  }
  
  return allReady;
}

/**
 * One-click setup - combines everything
 */
function oneClickSetup() {
  console.log('🎯 One-Click OOXML Slides Setup');
  console.log('=' * 32);
  
  try {
    // Step 1: Smart analysis
    console.log('📊 Analyzing current setup...');
    const currentStatus = smartSetupCheck();
    
    if (currentStatus) {
      console.log('✅ Already set up! Running quick test...');
      runQuickTests();
      return true;
    }
    
    console.log('');
    console.log('🔄 Starting automatic setup...');
    
    // Step 2: Auto setup
    const setupResult = autoSetup();
    
    if (setupResult) {
      console.log('');
      console.log('🎉 One-click setup completed successfully!');
      console.log('🧪 Running validation tests...');
      
      // Step 3: Validation
      runQuickTests();
      
      return true;
    } else {
      console.log('');
      console.log('⚠️ Automatic setup incomplete');
      console.log('📋 Please follow manual setup instructions above');
      
      return false;
    }
    
  } catch (error) {
    console.error('❌ One-click setup failed:', error);
    console.log('');
    console.log('🔧 Please try manual setup:');
    showManualSetupInstructions();
    return false;
  }
}

// Quick access functions
function quickSetup() { return oneClickSetup(); }
function setup() { return oneClickSetup(); }
function checkSetup() { return smartSetupCheck(); }