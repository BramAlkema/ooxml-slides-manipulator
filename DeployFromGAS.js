/**
 * DEPLOY PPTX SERVICE FROM GOOGLE APPS SCRIPT
 * 
 * This script deploys the PPTX processing service directly from GAS
 * using the built-in OOXMLDeployment class. No command line needed!
 * 
 * SETUP:
 * 1. Copy this code to your Google Apps Script project
 * 2. Set your GCP_PROJECT_ID in Script Properties
 * 3. Run the deployment functions below
 */

// ========== STEP 1: SET YOUR PROJECT ID ==========

/**
 * First, set your Google Cloud Project ID
 * Run this once to configure your project
 */
function setupProjectId() {
  // Replace with your actual GCP project ID
  const YOUR_PROJECT_ID = 'your-project-id-here';
  
  // Save to Script Properties
  PropertiesService.getScriptProperties().setProperty('GCP_PROJECT_ID', YOUR_PROJECT_ID);
  
  console.log(`✅ Project ID set to: ${YOUR_PROJECT_ID}`);
  console.log('Next step: Run showPreflightChecks()');
}

// ========== STEP 2: PREFLIGHT CHECKS ==========

/**
 * Show the preflight checks sidebar
 * This verifies billing, APIs, and settings
 */
function showPreflightChecks() {
  console.log('🔍 Opening preflight checks sidebar...');
  
  try {
    // This opens a sidebar in Google Sheets/Docs
    OOXMLDeployment.showGcpPreflight();
    
    console.log('✅ Preflight sidebar opened');
    console.log('Complete the checks in the sidebar, then run deployToUSFreeTier()');
    
  } catch (error) {
    console.error('❌ Error showing preflight:', error.message);
    console.log('Make sure OOXMLDeployment class is included in your project');
  }
}

// ========== STEP 3: DEPLOY TO US FREE TIER ==========

/**
 * Deploy the PPTX processing service to US free tier
 * Run this after completing preflight checks
 */
async function deployToUSFreeTier() {
  console.log('🚀 Starting deployment to US free tier...');
  
  try {
    // Deploy with US region for free tier
    const serviceUrl = await OOXMLDeployment.initAndDeploy({
      region: 'us-central1',  // US free tier region
      skipBillingCheck: false // Set to true if you're sure billing is enabled
    });
    
    console.log('✅ Deployment successful!');
    console.log(`🔗 Service URL: ${serviceUrl}`);
    
    // Save the URL for later use
    PropertiesService.getScriptProperties().setProperty('CLOUD_SERVICE_URL', serviceUrl);
    
    // Verify deployment
    const status = OOXMLDeployment.getDeploymentStatus();
    console.log('📊 Deployment status:', JSON.stringify(status, null, 2));
    
    console.log('\n🎉 Your PPTX service is now running on US free tier!');
    console.log('Next step: Run testDeployedService()');
    
    return serviceUrl;
    
  } catch (error) {
    console.error('❌ Deployment failed:', error.message);
    
    if (error.message.includes('Billing')) {
      console.log('💳 Please enable billing first via showPreflightChecks()');
    } else if (error.message.includes('PROJECT_ID')) {
      console.log('🔧 Please run setupProjectId() first');
    } else {
      console.log('Check the error message above and try again');
    }
  }
}

// ========== STEP 4: TEST THE DEPLOYED SERVICE ==========

/**
 * Test the deployed PPTX processing service
 */
async function testDeployedService() {
  console.log('🧪 Testing deployed PPTX service...');
  
  try {
    // Test 1: Health check
    const health = await OOXMLJsonService.healthCheck();
    console.log('✅ Health check:', JSON.stringify(health, null, 2));
    
    // Test 2: Service info
    const info = OOXMLJsonService.getServiceInfo();
    console.log('📋 Service info:', JSON.stringify(info, null, 2));
    
    // Test 3: Create a simple test
    console.log('\n🔧 Testing PPTX processing...');
    
    // Create a test presentation
    const testPresentation = createTestPresentation();
    const testFileId = testPresentation.getId();
    
    // Unwrap the test presentation
    const manifest = await OOXMLJsonService.unwrap(testFileId);
    console.log(`✅ Unwrapped ${manifest.entries.length} files from PPTX`);
    
    // Make a simple modification
    let modifiedCount = 0;
    manifest.entries.forEach(entry => {
      if (entry.type === 'xml' && entry.path.includes('slide')) {
        entry.text = entry.text.replace('Test Slide', 'Modified Slide');
        modifiedCount++;
      }
    });
    console.log(`✅ Modified ${modifiedCount} slide files`);
    
    // Rewrap the presentation
    const outputFileId = await OOXMLJsonService.rewrap(manifest, {
      filename: 'test-output-pptx.pptx'
    });
    console.log(`✅ Created modified PPTX: ${outputFileId}`);
    
    console.log('\n🎉 All tests passed! Your PPTX service is working!');
    
    return {
      health,
      info,
      testFileId: outputFileId
    };
    
  } catch (error) {
    console.error('❌ Test failed:', error.message);
    console.log('Make sure the service is deployed first with deployToUSFreeTier()');
  }
}

// ========== HELPER FUNCTIONS ==========

/**
 * Create a simple test presentation
 */
function createTestPresentation() {
  const presentation = SlidesApp.create('Test Presentation ' + new Date().toISOString());
  
  // Add a test slide
  const slide = presentation.getSlides()[0];
  slide.getShapes()[0].getText().setText('Test Slide\nCreated by GAS');
  
  presentation.saveAndClose();
  
  console.log('📝 Created test presentation:', presentation.getId());
  return DriveApp.getFileById(presentation.getId());
}

/**
 * Check current deployment status
 */
function checkDeploymentStatus() {
  try {
    const status = OOXMLDeployment.getDeploymentStatus();
    console.log('📊 Current deployment status:');
    console.log(JSON.stringify(status, null, 2));
    
    if (status.deployed) {
      console.log('✅ Service is deployed and running');
      console.log(`🔗 URL: ${status.serviceUrl}`);
    } else {
      console.log('⚠️ Service is not deployed yet');
      console.log('Run deployToUSFreeTier() to deploy');
    }
    
    return status;
    
  } catch (error) {
    console.error('❌ Error checking status:', error.message);
  }
}

/**
 * Quick deployment (all steps in one)
 * Only use this if you're confident everything is set up
 */
async function quickDeployToFreeTier() {
  console.log('⚡ Quick deployment to US free tier...');
  
  try {
    // Check if project ID is set
    const projectId = PropertiesService.getScriptProperties().getProperty('GCP_PROJECT_ID');
    if (!projectId || projectId === 'your-project-id-here') {
      throw new Error('Please run setupProjectId() first with your actual project ID');
    }
    
    console.log(`📋 Using project: ${projectId}`);
    
    // Deploy directly (assumes billing is enabled)
    const serviceUrl = await OOXMLDeployment.initAndDeploy({
      region: 'us-central1',
      skipBillingCheck: true // Skip for quick deployment
    });
    
    console.log('✅ Quick deployment successful!');
    console.log(`🔗 Service URL: ${serviceUrl}`);
    
    // Save URL
    PropertiesService.getScriptProperties().setProperty('CLOUD_SERVICE_URL', serviceUrl);
    
    // Quick test
    const health = await OOXMLJsonService.healthCheck();
    console.log('✅ Service is healthy:', health.status);
    
    console.log('\n🎉 PPTX service deployed to US free tier!');
    return serviceUrl;
    
  } catch (error) {
    console.error('❌ Quick deployment failed:', error.message);
    console.log('Try the step-by-step approach instead:');
    console.log('1. setupProjectId()');
    console.log('2. showPreflightChecks()');
    console.log('3. deployToUSFreeTier()');
    console.log('4. testDeployedService()');
  }
}

/**
 * Clean up / remove the deployed service
 * USE WITH CAUTION - This will delete your service
 */
function removeDeployedService() {
  console.log('⚠️ WARNING: This will remove your deployed service');
  console.log('Uncomment the code below if you really want to proceed');
  
  // Uncomment these lines to actually remove the service:
  /*
  try {
    const result = OOXMLDeployment.cleanup({
      removeService: true,
      removeBucket: false,
      clearProperties: false
    });
    
    console.log('Service removal result:', JSON.stringify(result, null, 2));
    
    if (result.success) {
      console.log('✅ Service removed successfully');
    } else {
      console.log('⚠️ Some cleanup operations failed');
    }
  } catch (error) {
    console.error('❌ Removal failed:', error.message);
  }
  */
}

// ========== RUN THESE IN ORDER ==========
// 1. setupProjectId()       - Set your GCP project ID
// 2. showPreflightChecks()  - Verify billing and APIs
// 3. deployToUSFreeTier()   - Deploy the service
// 4. testDeployedService()  - Test that it works

// Or use quickDeployToFreeTier() if you're confident everything is set up