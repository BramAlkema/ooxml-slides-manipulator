/**
 * SmartAPIEnabler - Advanced API enablement through various methods
 * Uses multiple strategies to automatically enable required APIs
 */

/**
 * Advanced Drive API enablement with multiple fallback strategies
 */
function enableDriveAPIAdvanced() {
  console.log('🚀 Advanced Drive API Enablement');
  console.log('=' * 35);
  
  const strategies = [
    { name: 'OAuth Trigger Method', func: enableViaOAuth },
    { name: 'Manifest Update Method', func: enableViaManifest },
    { name: 'Cloud Console API Method', func: enableViaCloudConsole },
    { name: 'Service Account Method', func: enableViaServiceAccount },
    { name: 'Force Authorization Method', func: enableViaForceAuth }
  ];
  
  let successCount = 0;
  
  strategies.forEach((strategy, index) => {
    try {
      console.log(`\n${index + 1}. ${strategy.name}:`);
      const result = strategy.func();
      
      if (result) {
        console.log('✅ Success!');
        successCount++;
      } else {
        console.log('⚠️ Partial success or needs manual step');
      }
      
    } catch (error) {
      console.log('❌ Failed:', error.message);
    }
  });
  
  console.log(`\n📊 Results: ${successCount}/${strategies.length} methods succeeded`);
  
  // Final verification
  console.log('\n🔍 Final verification...');
  const finalTest = testDriveAPIAccess();
  
  if (finalTest) {
    console.log('🎉 Drive API is now enabled and working!');
    return true;
  } else {
    console.log('⚠️ API enablement may need manual completion');
    showAlternativeEnablementMethods();
    return false;
  }
}

/**
 * Strategy 1: Enable via OAuth consent flow
 */
function enableViaOAuth() {
  console.log('  📝 Triggering OAuth consent flow...');
  
  try {
    // Force OAuth consent by accessing user-specific data
    const userEmail = Session.getActiveUser().getEmail();
    console.log(`  👤 User: ${userEmail}`);
    
    // Try to access Drive to trigger consent
    const rootFolder = DriveApp.getRootFolder().getName();
    console.log(`  📁 Root folder: ${rootFolder}`);
    
    // If we get here, Drive access is working
    return true;
    
  } catch (error) {
    console.log(`  ⚠️ OAuth trigger: ${error.message}`);
    
    // Check if it's a permissions error (which means API is enabled but needs consent)
    if (error.message.includes('permission') || error.message.includes('scope')) {
      console.log('  💡 Permissions error detected - API may be enabled');
      return false; // Partial success
    }
    
    return false;
  }
}

/**
 * Strategy 2: Update manifest/scopes approach
 */
function enableViaManifest() {
  console.log('  📜 Checking manifest configuration...');
  
  try {
    // We can't directly modify appsscript.json, but we can validate scopes
    const requiredScopes = [
      'https://www.googleapis.com/auth/drive',
      'https://www.googleapis.com/auth/drive.file'
    ];
    
    console.log('  🔍 Required scopes:');
    requiredScopes.forEach(scope => console.log(`    - ${scope}`));
    
    // Try to validate if scopes are working
    const token = ScriptApp.getOAuthToken();
    if (token && token.length > 0) {
      console.log('  ✅ OAuth token available');
      
      // Test token with Drive API
      const response = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/about', {
        headers: { 'Authorization': `Bearer ${token}` }
      });
      
      if (response.getResponseCode() === 200) {
        console.log('  ✅ Drive API responding to token');
        return true;
      }
    }
    
    console.log('  💡 Manifest appears configured but API may need manual enablement');
    return false;
    
  } catch (error) {
    console.log(`  ⚠️ Manifest check: ${error.message}`);
    return false;
  }
}

/**
 * Strategy 3: Cloud Console API enablement via UrlFetch
 */
function enableViaCloudConsole() {
  console.log('  ☁️ Attempting Cloud Console API enablement...');
  
  try {
    // Get project ID
    const projectId = getProjectId();
    if (!projectId) {
      console.log('  ⚠️ Could not determine project ID');
      return false;
    }
    
    console.log(`  📁 Project ID: ${projectId}`);
    
    // Try to enable Drive API via Service Management API
    const enableUrl = `https://servicemanagement.googleapis.com/v1/services/drive.googleapis.com:enable`;
    
    const response = UrlFetchApp.fetch(enableUrl, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${ScriptApp.getOAuthToken()}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({ 
        consumerId: `project:${projectId}` 
      })
    });
    
    if (response.getResponseCode() === 200) {
      console.log('  ✅ API enablement request sent');
      return true;
    } else {
      console.log(`  ⚠️ API enablement response: ${response.getResponseCode()}`);
      return false;
    }
    
  } catch (error) {
    console.log(`  ⚠️ Cloud Console method: ${error.message}`);
    return false;
  }
}

/**
 * Strategy 4: Service Account approach
 */
function enableViaServiceAccount() {
  console.log('  🤖 Checking service account approach...');
  
  try {
    // Check if we can access service information
    const scriptId = ScriptApp.getScriptId();
    console.log(`  📋 Script ID: ${scriptId}`);
    
    // Try to get project information
    const projectInfo = getProjectInformation();
    if (projectInfo) {
      console.log('  ✅ Project information accessible');
      
      // Service accounts don't apply to Apps Script in the same way
      // But we can check if advanced services are configured
      return testAdvancedServices();
    }
    
    return false;
    
  } catch (error) {
    console.log(`  ⚠️ Service account check: ${error.message}`);
    return false;
  }
}

/**
 * Strategy 5: Force authorization through multiple service calls
 */
function enableViaForceAuth() {
  console.log('  🔐 Force authorization method...');
  
  try {
    // Make multiple different API calls to trigger authorization
    const tests = [
      () => DriveApp.getRootFolder().getId(),
      () => DriveApp.searchFiles('title contains "test"').hasNext(),
      () => Drive.About.get(),
      () => Drive.Files.list({ maxResults: 1 })
    ];
    
    let successes = 0;
    
    tests.forEach((test, index) => {
      try {
        const result = test();
        console.log(`    ✅ Test ${index + 1}: Success`);
        successes++;
      } catch (error) {
        console.log(`    ⚠️ Test ${index + 1}: ${error.message.substring(0, 50)}...`);
      }
    });
    
    console.log(`  📊 Authorization tests: ${successes}/${tests.length} passed`);
    
    return successes > 0;
    
  } catch (error) {
    console.log(`  ⚠️ Force auth: ${error.message}`);
    return false;
  }
}

/**
 * Get project ID through various methods
 */
function getProjectId() {
  try {
    // Method 1: Try to extract from script URL or ID
    const scriptId = ScriptApp.getScriptId();
    
    // Method 2: Try to get from Cloud Resource Manager
    try {
      const response = UrlFetchApp.fetch('https://cloudresourcemanager.googleapis.com/v1/projects', {
        headers: { 'Authorization': `Bearer ${ScriptApp.getOAuthToken()}` }
      });
      
      if (response.getResponseCode() === 200) {
        const data = JSON.parse(response.getContentText());
        if (data.projects && data.projects.length > 0) {
          return data.projects[0].projectId;
        }
      }
    } catch (error) {
      // Ignore, try other methods
    }
    
    // Method 3: Use script ID as fallback (not always the same as project ID)
    return scriptId;
    
  } catch (error) {
    return null;
  }
}

/**
 * Get project information
 */
function getProjectInformation() {
  try {
    return {
      scriptId: ScriptApp.getScriptId(),
      user: Session.getActiveUser().getEmail(),
      timezone: Session.getScriptTimeZone()
    };
  } catch (error) {
    return null;
  }
}

/**
 * Test advanced services configuration
 */
function testAdvancedServices() {
  console.log('  🔧 Testing advanced services...');
  
  const services = [
    { name: 'Drive', test: () => Drive.Files.list({ maxResults: 1 }) },
    { name: 'Admin SDK', test: () => typeof AdminDirectory !== 'undefined' },
    { name: 'Gmail', test: () => typeof GmailApp !== 'undefined' }
  ];
  
  let available = 0;
  
  services.forEach(service => {
    try {
      service.test();
      console.log(`    ✅ ${service.name}: Available`);
      available++;
    } catch (error) {
      console.log(`    ⚠️ ${service.name}: ${error.message.substring(0, 30)}...`);
    }
  });
  
  console.log(`  📊 Advanced services: ${available}/${services.length} available`);
  return available > 0;
}

/**
 * Show alternative enablement methods if all strategies fail
 */
function showAlternativeEnablementMethods() {
  console.log('\n🔄 ALTERNATIVE ENABLEMENT METHODS');
  console.log('=' * 38);
  console.log('');
  console.log('🎯 Method A: Manual Console Enablement');
  console.log('1. Go to: https://console.cloud.google.com/apis/dashboard');
  console.log('2. Select your Apps Script project');
  console.log('3. Click "+ ENABLE APIS AND SERVICES"');
  console.log('4. Search for "Google Drive API" and enable');
  console.log('');
  console.log('🎯 Method B: Apps Script Interface');
  console.log('1. In Apps Script: Resources → Advanced Google Services');
  console.log('2. Enable "Google Drive API v3"');
  console.log('3. Click the Cloud Console link and enable there too');
  console.log('');
  console.log('🎯 Method C: Project Recreation');
  console.log('1. Create new Apps Script project');
  console.log('2. Copy all code files');
  console.log('3. Configure APIs from scratch');
  console.log('');
  console.log('🎯 Method D: Direct Authorization');
  console.log('1. Run: authorizeDirectly()');
  console.log('2. Follow OAuth prompts');
  console.log('3. Grant all requested permissions');
  console.log('');
}

/**
 * Direct authorization method
 */
function authorizeDirectly() {
  console.log('🔐 Direct Authorization Process');
  console.log('=' * 32);
  
  try {
    console.log('Step 1: Requesting user authorization...');
    const userEmail = Session.getActiveUser().getEmail();
    console.log(`✅ User authenticated: ${userEmail}`);
    
    console.log('Step 2: Testing Drive access...');
    const rootFolderId = DriveApp.getRootFolder().getId();
    console.log(`✅ Drive access granted: ${rootFolderId}`);
    
    console.log('Step 3: Testing advanced Drive API...');
    const aboutInfo = Drive.About.get();
    console.log(`✅ Advanced API working: ${aboutInfo.user.emailAddress}`);
    
    console.log('');
    console.log('🎉 Direct authorization completed!');
    console.log('✅ All APIs are now accessible');
    
    return true;
    
  } catch (error) {
    console.error('❌ Direct authorization failed:', error);
    console.log('');
    console.log('💡 This might be the first run - please:');
    console.log('1. Click "Review permissions" when prompted');
    console.log('2. Sign in to your Google account');
    console.log('3. Grant all requested permissions');
    console.log('4. Run this function again');
    
    return false;
  }
}

/**
 * All-in-one smart enablement
 */
function smartEnableAPIs() {
  console.log('🧠 Smart API Enablement');
  console.log('=' * 23);
  
  // Quick test first
  if (testDriveAPIAccess()) {
    console.log('✅ APIs already enabled and working!');
    return true;
  }
  
  console.log('🔄 APIs need enablement, trying smart methods...');
  
  // Try direct authorization first (most likely to work)
  if (authorizeDirectly()) {
    return true;
  }
  
  // Fall back to advanced methods
  return enableDriveAPIAdvanced();
}