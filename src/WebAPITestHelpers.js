/**
 * Web API Test Helper Functions
 * 
 * These functions are designed to be called via the Web App API
 * for testing with Playwright screenshots.
 */

/**
 * Create a presentation from prompt via API
 * This function will be called by Playwright tests
 */
async function createPresentationFromPrompt(prompt) {
  try {
    ConsoleFormatter.header('🎨 API: Creating Presentation from Prompt');
    ConsoleFormatter.info(`Prompt: "${prompt}"`);
    
    // Initialize extensions first
    await initializeExtensions();
    
    // Create presentation using the extension
    const slides = await OOXMLSlides.createFromPrompt(prompt);
    
    const result = {
      success: true,
      message: 'Presentation created successfully from prompt',
      presentationId: slides.fileId,
      presentationUrl: `https://docs.google.com/presentation/d/${slides.fileId}/edit`,
      publicUrl: `https://docs.google.com/presentation/d/${slides.fileId}/pub`,
      timestamp: new Date().toISOString()
    };
    
    ConsoleFormatter.success(`Presentation created: ${result.presentationId}`);
    return result;
    
  } catch (error) {
    ConsoleFormatter.error('Failed to create presentation from prompt', error);
    return {
      success: false,
      error: error.message,
      stack: error.stack
    };
  }
}

/**
 * Test the extension system and return status
 */
async function testExtensionSystemAPI() {
  try {
    ConsoleFormatter.header('🧩 API: Testing Extension System');
    
    // Run the extension system test
    const testResults = await testExtensionSystem();
    
    // Get extension status
    const extensionStatus = getExtensionStatus();
    
    return {
      success: testResults.success,
      testResults: testResults,
      extensionStatus: extensionStatus,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    ConsoleFormatter.error('Extension system API test failed', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Run system validation and return results
 */
async function runSystemValidationAPI() {
  try {
    ConsoleFormatter.header('🔍 API: Running System Validation');
    
    // Run preflight checks
    const preflightResults = runPreflightChecks();
    
    // Run extension system test
    const extensionTest = await testExtensionSystem();
    
    // Run migration test
    const migrationTest = testCloudPPTXServiceMigration();
    
    const overallSuccess = preflightResults && extensionTest.success && migrationTest.success;
    
    return {
      success: overallSuccess,
      preflight: preflightResults,
      extensions: extensionTest,
      migration: migrationTest,
      timestamp: new Date().toISOString(),
      summary: {
        preflightPassed: !!preflightResults,
        extensionsPassed: extensionTest.passed,
        migrationPassed: migrationTest.passed,
        overallStatus: overallSuccess ? 'ALL_SYSTEMS_OPERATIONAL' : 'ISSUES_DETECTED'
      }
    };
    
  } catch (error) {
    ConsoleFormatter.error('System validation API failed', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Create a test presentation with specific theme
 */
async function createTestPresentationWithTheme(colors, fonts) {
  try {
    ConsoleFormatter.header('🎭 API: Creating Test Presentation with Theme');
    
    // Create a new presentation
    const presentation = SlidesApp.create(`API Test Presentation - ${new Date().toISOString().split('T')[0]}`);
    const presentationId = presentation.getId();
    
    ConsoleFormatter.status('PASS', 'Base Creation', `Created: ${presentationId}`);
    
    // Initialize OOXMLSlides
    const slides = new OOXMLSlides(presentationId, {
      createBackup: false
    });
    
    await slides.load();
    
    // Apply custom colors if provided
    if (colors) {
      await slides.applyCustomColors(colors);
      ConsoleFormatter.status('PASS', 'Colors Applied', `Applied custom color scheme`);
    }
    
    // Apply fonts if provided
    if (fonts) {
      await slides.applyFontPairing(fonts);
      ConsoleFormatter.status('PASS', 'Fonts Applied', `Applied font pairing: ${fonts.heading}/${fonts.body}`);
    }
    
    await slides.save();
    
    const result = {
      success: true,
      presentationId: presentationId,
      presentationUrl: presentation.getUrl(),
      editUrl: `https://docs.google.com/presentation/d/${presentationId}/edit`,
      pubUrl: `https://docs.google.com/presentation/d/${presentationId}/pub`,
      colors: colors,
      fonts: fonts,
      timestamp: new Date().toISOString()
    };
    
    ConsoleFormatter.success(`Test presentation created with theme`);
    return result;
    
  } catch (error) {
    ConsoleFormatter.error('Failed to create test presentation with theme', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test FFlatePPTXService functionality
 */
function testFFlatePPTXServiceAPI() {
  try {
    ConsoleFormatter.header('⚡ API: Testing FFlatePPTXService');
    
    // Run the FFlatePPTXService tests
    const testResults = testFFlatePPTXService();
    
    return {
      success: testResults.success || testResults.allTestsPassed,
      results: testResults,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    ConsoleFormatter.error('FFlatePPTXService API test failed', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Get comprehensive system status
 */
function getSystemStatusAPI() {
  try {
    ConsoleFormatter.header('📊 API: Getting System Status');
    
    const status = {
      timestamp: new Date().toISOString(),
      services: {
        fflatePPTXService: typeof FFlatePPTXService !== 'undefined',
        ooxmlCore: typeof OOXMLCore !== 'undefined',
        ooxmlSlides: typeof OOXMLSlides !== 'undefined',
        driveAPI: typeof DriveApp !== 'undefined',
        slidesAPI: typeof SlidesApp !== 'undefined'
      },
      extensions: getExtensionStatus(),
      runtime: {
        hasV8: typeof Map !== 'undefined',
        hasClasses: true,
        hasArrowFunctions: true
      }
    };
    
    // Count available services
    const availableServices = Object.values(status.services).filter(available => available).length;
    const totalServices = Object.keys(status.services).length;
    
    status.summary = {
      servicesAvailable: availableServices,
      totalServices: totalServices,
      serviceHealth: availableServices / totalServices,
      extensionsLoaded: status.extensions.loaded || 0,
      overallHealth: (availableServices === totalServices && status.extensions.loaded > 0) ? 'HEALTHY' : 'DEGRADED'
    };
    
    return {
      success: true,
      status: status
    };
    
  } catch (error) {
    ConsoleFormatter.error('System status API failed', error);
    return {
      success: false,
      error: error.message
    };
  }
}