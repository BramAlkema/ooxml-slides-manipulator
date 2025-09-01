/**
 * PreflightChecks - Comprehensive system validation for OOXML deployment
 * 
 * Validates both FFlatePPTXService (client-side) and FFlatePPTXService (Cloud Function)
 * along with all required permissions, APIs, and configurations.
 * 
 * Updated for the new dual-service architecture.
 */

/**
 * Main preflight check function - validates entire system
 */
function runPreflightChecks() {
  ConsoleFormatter.header('🔍 OOXML System Preflight Checks', '=', 41);
  
  const startTime = Date.now();
  const results = {
    runtime: checkRuntimeEnvironment(),
    services: checkServices(),
    cloudFunction: checkCloudFunction(), 
    permissions: checkPermissions(),
    apis: checkAPIs(),
    network: checkNetworkConnectivity(),
    storage: checkStorageAccess(),
    iam: checkIAMConfiguration()
  };
  
  // Generate summary
  const duration = Date.now() - startTime;
  const summary = generatePreflightSummary(results, duration);
  
  ConsoleFormatter.header('\n📋 PREFLIGHT SUMMARY', '=', 50);
  console.log(summary);
  
  return results;
}

/**
 * Check runtime environment and basic capabilities
 */
function checkRuntimeEnvironment() {
  ConsoleFormatter.section('🔧 Runtime Environment Checks');
  
  const checks = [];
  
  // Check GAS runtime version
  try {
    const runtimeInfo = {
      hasV8: typeof Map !== 'undefined',
      hasClasses: true, // GAS V8 supports classes
      hasArrowFunctions: true, // If this executes, arrow functions work
      hasSpreadOperator: true // GAS V8 supports spread operator
    };
    
    checks.push({
      name: 'V8 Runtime',
      status: runtimeInfo.hasV8 ? 'PASS' : 'FAIL',
      details: runtimeInfo.hasV8 ? 'Modern JS features available' : 'Legacy runtime detected'
    });
    
    checks.push({
      name: 'Class Support',
      status: runtimeInfo.hasClasses ? 'PASS' : 'FAIL', 
      details: runtimeInfo.hasClasses ? 'ES6 classes supported' : 'Classes not available'
    });
    
  } catch (error) {
    checks.push({
      name: 'Runtime Detection',
      status: 'ERROR',
      details: error.message
    });
  }
  
  // Check available GAS utilities
  const utilities = [
    { name: 'Utilities', ref: () => typeof Utilities !== 'undefined' },
    { name: 'DriveApp', ref: () => typeof DriveApp !== 'undefined' },
    { name: 'UrlFetchApp', ref: () => typeof UrlFetchApp !== 'undefined' },
    { name: 'SlidesApp', ref: () => typeof SlidesApp !== 'undefined' }
  ];
  utilities.forEach(util => {
    const available = util.ref();
    checks.push({
      name: `${util.name} API`,
      status: available ? 'PASS' : 'FAIL',
      details: available ? 'Available' : 'Not accessible'
    });
  });
  
  checks.forEach(check => {
    ConsoleFormatter.status(check.status, check.name, check.details);
  });
  
  return checks;
}

/**
 * Check both FFlatePPTXService and FFlatePPTXService
 */
function checkServices() {
  ConsoleFormatter.section('📦 Service Availability Checks');
  
  const checks = [];
  
  // Check FFlatePPTXService
  try {
    if (typeof FFlatePPTXService !== 'undefined') {
      // Test service info
      const serviceInfo = FFlatePPTXService.getServiceInfo();
      checks.push({
        name: 'FFlatePPTXService',
        status: 'PASS',
        details: `v${serviceInfo.version} - ${serviceInfo.type}`
      });
      
      // Test conversion utilities
      try {
        const testString = 'Hello World';
        const uint8 = FFlatePPTXService.stringToUint8Array(testString);
        const converted = FFlatePPTXService.uint8ArrayToString(uint8);
        
        checks.push({
          name: 'Text Conversion',
          status: converted === testString ? 'PASS' : 'FAIL',
          details: converted === testString ? 'String ↔ Uint8Array working' : 'Conversion failed'
        });
      } catch (error) {
        checks.push({
          name: 'Text Conversion',
          status: 'ERROR',
          details: error.message
        });
      }
      
      // Check for fflate dependency
      checks.push({
        name: 'fflate Library',
        status: typeof fflate !== 'undefined' ? 'PASS' : 'WARN',
        details: typeof fflate !== 'undefined' ? 'Available' : 'Not found - will cause runtime errors'
      });
      
    } else {
      checks.push({
        name: 'FFlatePPTXService',
        status: 'FAIL',
        details: 'Service class not found'
      });
    }
  } catch (error) {
    checks.push({
      name: 'FFlatePPTXService',
      status: 'ERROR',
      details: error.message
    });
  }
  
  // Check FFlatePPTXService
  try {
    if (typeof FFlatePPTXService !== 'undefined') {
      const serviceInfo = FFlatePPTXService.getServiceInfo();
      checks.push({
        name: 'FFlatePPTXService',
        status: 'PASS',
        details: `v${serviceInfo.version} - ${serviceInfo.type}`
      });
    } else {
      checks.push({
        name: 'FFlatePPTXService',
        status: 'FAIL',
        details: 'Service class not found'
      });
    }
  } catch (error) {
    checks.push({
      name: 'FFlatePPTXService',
      status: 'ERROR',
      details: error.message
    });
  }
  
  // Check OOXMLCore
  try {
    if (typeof OOXMLCore !== 'undefined') {
      checks.push({
        name: 'OOXMLCore',
        status: 'PASS',
        details: 'Abstraction layer available'
      });
    } else {
      checks.push({
        name: 'OOXMLCore',
        status: 'FAIL',
        details: 'Core class not found'
      });
    }
  } catch (error) {
    checks.push({
      name: 'OOXMLCore',
      status: 'ERROR',
      details: error.message
    });
  }
  
  checks.forEach(check => {
    ConsoleFormatter.status(check.status, check.name, check.details);
  });
  
  return checks;
}

/**
 * Check Cloud Function connectivity and IAM
 */
function checkCloudFunction() {
  ConsoleFormatter.section('☁️ Cloud Function Checks');
  
  const checks = [];
  
  try {
    // Get the configured Cloud Function URL
    const config = FFlatePPTXService._CONFIG;
    const functionUrl = config.cloudFunctionUrl;
    
    checks.push({
      name: 'Function URL Config',
      status: functionUrl ? 'PASS' : 'FAIL',
      details: functionUrl ? functionUrl : 'No URL configured'
    });
    
    if (functionUrl) {
      // Test basic connectivity
      try {
        const testPayload = { files: {} };
        const response = UrlFetchApp.fetch(functionUrl + '/zip', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          payload: JSON.stringify(testPayload),
          muteHttpExceptions: true
        });
        
        const statusCode = response.getResponseCode();
        const responseText = response.getContentText();
        
        if (statusCode === 200) {
          try {
            const result = JSON.parse(responseText);
            checks.push({
              name: 'Function Connectivity',
              status: result.success ? 'PASS' : 'FAIL',
              details: result.success ? 'Function responding correctly' : 'Function error: ' + result.error
            });
          } catch (parseError) {
            checks.push({
              name: 'Function Connectivity',
              status: 'WARN',
              details: `Got response but invalid JSON: ${responseText.substring(0, 100)}`
            });
          }
        } else if (statusCode === 403) {
          checks.push({
            name: 'Function Connectivity',
            status: 'FAIL',
            details: 'HTTP 403 - IAM policy issue (not public)'
          });
          
          checks.push({
            name: 'IAM Public Access',
            status: 'FAIL',
            details: 'Function not publicly accessible - needs allUsers invoker role'
          });
        } else {
          checks.push({
            name: 'Function Connectivity',
            status: 'FAIL',
            details: `HTTP ${statusCode}: ${responseText.substring(0, 100)}`
          });
        }
      } catch (networkError) {
        checks.push({
          name: 'Function Connectivity',
          status: 'ERROR',
          details: `Network error: ${networkError.message}`
        });
      }
    }
    
  } catch (error) {
    checks.push({
      name: 'Cloud Function Check',
      status: 'ERROR',
      details: error.message
    });
  }
  
  checks.forEach(check => {
    ConsoleFormatter.status(check.status, check.name, check.details);
  });
  
  return checks;
}

/**
 * Check OAuth permissions and scopes
 */
function checkPermissions() {
  ConsoleFormatter.section('🔐 Permission Checks');
  
  const checks = [];
  
  // Test Drive access
  try {
    const testFolder = DriveApp.getRootFolder();
    checks.push({
      name: 'Drive API Access',
      status: 'PASS',
      details: 'Can access Google Drive'
    });
  } catch (error) {
    checks.push({
      name: 'Drive API Access',
      status: 'FAIL',
      details: `Drive access denied: ${error.message}`
    });
  }
  
  // Test Slides access  
  try {
    // Try to create a test presentation
    const testPresentation = Slides.Presentations.create({
      title: 'OOXML Preflight Test - ' + new Date().getTime()
    });
    
    if (testPresentation && testPresentation.presentationId) {
      checks.push({
        name: 'Slides API Access',
        status: 'PASS',
        details: 'Can create presentations'
      });
      
      // Clean up test presentation
      try {
        DriveApp.getFileById(testPresentation.presentationId).setTrashed(true);
      } catch (cleanupError) {
        console.log(`  ℹ️  Cleanup note: ${cleanupError.message}`);
      }
    } else {
      checks.push({
        name: 'Slides API Access',
        status: 'FAIL',
        details: 'Cannot create presentations'
      });
    }
  } catch (error) {
    checks.push({
      name: 'Slides API Access',
      status: 'FAIL', 
      details: `Slides access denied: ${error.message}`
    });
  }
  
  // Test external requests (for Cloud Function)
  try {
    const response = UrlFetchApp.fetch('https://httpbin.org/json', { muteHttpExceptions: true });
    const statusCode = response.getResponseCode();
    
    checks.push({
      name: 'External HTTP Access',
      status: statusCode === 200 ? 'PASS' : 'FAIL',
      details: statusCode === 200 ? 'Can make external requests' : `HTTP ${statusCode} error`
    });
  } catch (error) {
    checks.push({
      name: 'External HTTP Access',
      status: 'FAIL',
      details: `External requests blocked: ${error.message}`
    });
  }
  
  checks.forEach(check => {
    ConsoleFormatter.status(check.status, check.name, check.details);
  });
  
  return checks;
}

/**
 * Check required Google Cloud APIs
 */
function checkAPIs() {
  ConsoleFormatter.section('📡 API Status Checks');
  
  const checks = [];
  
  // These checks are indirect since we can't directly query API status from GAS
  const apiTests = [
    {
      name: 'Drive API v3',
      test: () => DriveApp.getRootFolder().getName(),
      details: 'Required for file operations'
    },
    {
      name: 'Slides API v1', 
      test: () => Slides.Presentations !== undefined,
      details: 'Required for presentation manipulation'
    }
  ];
  
  apiTests.forEach(apiTest => {
    try {
      apiTest.test();
      checks.push({
        name: apiTest.name,
        status: 'PASS',
        details: apiTest.details + ' - Working'
      });
    } catch (error) {
      checks.push({
        name: apiTest.name,
        status: 'FAIL',
        details: `${apiTest.details} - Error: ${error.message}`
      });
    }
  });
  
  checks.forEach(check => {
    ConsoleFormatter.status(check.status, check.name, check.details);
  });
  
  return checks;
}

/**
 * Check network connectivity to required services
 */
function checkNetworkConnectivity() {
  ConsoleFormatter.section('🌐 Network Connectivity Checks');
  
  const checks = [];
  
  const endpoints = [
    {
      name: 'Google APIs',
      url: 'https://www.googleapis.com',
      required: true
    },
    {
      name: 'Cloud Function',
      url: FFlatePPTXService._CONFIG.cloudFunctionUrl,
      required: false // Can work without it using FFlatePPTXService
    }
  ];
  
  endpoints.forEach(endpoint => {
    if (!endpoint.url) {
      checks.push({
        name: endpoint.name,
        status: 'SKIP',
        details: 'URL not configured'
      });
      return;
    }
    
    try {
      const response = UrlFetchApp.fetch(endpoint.url, {
        method: 'HEAD',
        muteHttpExceptions: true
      });
      
      const statusCode = response.getResponseCode();
      const isSuccess = statusCode >= 200 && statusCode < 400;
      
      checks.push({
        name: endpoint.name,
        status: isSuccess ? 'PASS' : (endpoint.required ? 'FAIL' : 'WARN'),
        details: `HTTP ${statusCode} - ${isSuccess ? 'Reachable' : 'Unreachable'}`
      });
    } catch (error) {
      checks.push({
        name: endpoint.name,
        status: endpoint.required ? 'FAIL' : 'WARN',
        details: `Network error: ${error.message}`
      });
    }
  });
  
  checks.forEach(check => {
    ConsoleFormatter.status(check.status, check.name, check.details);
  });
  
  return checks;
}

/**
 * Check storage and file access capabilities
 */
function checkStorageAccess() {
  ConsoleFormatter.section('💾 Storage Access Checks');
  
  const checks = [];
  
  try {
    // Test creating a temporary file
    const testBlob = Utilities.newBlob('test content', 'text/plain', 'preflight-test.txt');
    const tempFile = DriveApp.createFile(testBlob);
    
    if (tempFile) {
      checks.push({
        name: 'File Creation',
        status: 'PASS',
        details: 'Can create files in Drive'
      });
      
      // Test file modification
      tempFile.setContent('modified content');
      checks.push({
        name: 'File Modification',
        status: 'PASS',
        details: 'Can modify existing files'
      });
      
      // Cleanup
      tempFile.setTrashed(true);
      checks.push({
        name: 'File Cleanup',
        status: 'PASS',
        details: 'Can delete files'
      });
    }
  } catch (error) {
    checks.push({
      name: 'Storage Operations',
      status: 'FAIL',
      details: `Storage error: ${error.message}`
    });
  }
  
  checks.forEach(check => {
    ConsoleFormatter.status(check.status, check.name, check.details);
  });
  
  return checks;
}

/**
 * Check IAM configuration for Cloud Function
 */
function checkIAMConfiguration() {
  ConsoleFormatter.section('🔑 IAM Configuration Checks');
  
  const checks = [];
  
  try {
    const functionUrl = FFlatePPTXService._CONFIG.cloudFunctionUrl;
    
    if (!functionUrl) {
      checks.push({
        name: 'Cloud Function IAM',
        status: 'SKIP',
        details: 'No Cloud Function configured'
      });
      return checks;
    }
    
    // Test anonymous access (should work if IAM is set correctly)
    try {
      const testResponse = UrlFetchApp.fetch(functionUrl + '/zip', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        payload: JSON.stringify({ files: {} }),
        muteHttpExceptions: true
      });
      
      const statusCode = testResponse.getResponseCode();
      
      if (statusCode === 200) {
        checks.push({
          name: 'Public Access IAM',
          status: 'PASS',
          details: 'Function is publicly accessible (allUsers has run.invoker)'
        });
      } else if (statusCode === 403) {
        checks.push({
          name: 'Public Access IAM',
          status: 'FAIL',
          details: 'Function requires authentication - missing allUsers role'
        });
        
        // Provide remediation steps
        checks.push({
          name: 'IAM Fix Required',
          status: 'ACTION',
          details: 'Run: gcloud run services add-iam-policy-binding [service] --member=allUsers --role=roles/run.invoker'
        });
      } else {
        checks.push({
          name: 'Public Access IAM',
          status: 'WARN',
          details: `Unexpected HTTP ${statusCode} - check function status`
        });
      }
    } catch (error) {
      checks.push({
        name: 'Public Access IAM',
        status: 'ERROR',
        details: `IAM test failed: ${error.message}`
      });
    }
    
  } catch (error) {
    checks.push({
      name: 'IAM Configuration',
      status: 'ERROR',
      details: error.message
    });
  }
  
  checks.forEach(check => {
    const icon = check.status === 'PASS' ? '✅' : 
                 check.status === 'FAIL' ? '❌' : 
                 check.status === 'ACTION' ? '🔧' : '⚠️';
    console.log(`  ${icon} ${check.name}: ${check.details}`);
  });
  
  return checks;
}

/**
 * Generate comprehensive preflight summary
 */
function generatePreflightSummary(results, duration) {
  let summary = '';
  let totalChecks = 0;
  let passedChecks = 0;
  let failedChecks = 0;
  let warnings = 0;
  let actions = 0;
  
  // Count results across all categories
  Object.values(results).forEach(category => {
    if (Array.isArray(category)) {
      category.forEach(check => {
        totalChecks++;
        switch (check.status) {
          case 'PASS':
            passedChecks++;
            break;
          case 'FAIL':
          case 'ERROR':
            failedChecks++;
            break;
          case 'WARN':
            warnings++;
            break;
          case 'ACTION':
            actions++;
            break;
        }
      });
    }
  });
  
  summary += `\n📊 Overall Status: `;
  if (failedChecks === 0 && actions === 0) {
    summary += `✅ READY FOR DEPLOYMENT\n`;
  } else if (failedChecks > 0) {
    summary += `❌ DEPLOYMENT BLOCKED\n`;
  } else {
    summary += `⚠️ DEPLOYMENT WITH WARNINGS\n`;
  }
  
  summary += `\n📈 Check Results:\n`;
  summary += `   ✅ Passed: ${passedChecks}/${totalChecks}\n`;
  summary += `   ❌ Failed: ${failedChecks}/${totalChecks}\n`;
  summary += `   ⚠️  Warnings: ${warnings}/${totalChecks}\n`;
  summary += `   🔧 Actions Required: ${actions}/${totalChecks}\n`;
  summary += `\n⏱️  Duration: ${duration}ms\n`;
  
  if (failedChecks > 0 || actions > 0) {
    summary += `\n🔧 Required Actions:\n`;
    Object.values(results).forEach(category => {
      if (Array.isArray(category)) {
        category.forEach(check => {
          if (check.status === 'FAIL' || check.status === 'ERROR' || check.status === 'ACTION') {
            summary += `   • ${check.name}: ${check.details}\n`;
          }
        });
      }
    });
  }
  
  return summary;
}

/**
 * Quick preflight check - minimal validation
 */
function quickPreflightCheck() {
  console.log('⚡ Quick Preflight Check');
  console.log('-'.repeat(22));
  
  const critical = [
    () => typeof FFlatePPTXService !== 'undefined',
    () => typeof FFlatePPTXService !== 'undefined', 
    () => typeof OOXMLCore !== 'undefined',
    () => typeof DriveApp !== 'undefined',
    () => typeof Slides !== 'undefined'
  ];
  
  const results = critical.map((check, i) => {
    try {
      return check();
    } catch (error) {
      return false;
    }
  });
  
  const passed = results.filter(r => r).length;
  const total = results.length;
  
  console.log(`${passed}/${total} critical checks passed`);
  
  if (passed === total) {
    console.log('✅ Basic system ready');
    return true;
  } else {
    console.log('❌ Critical issues detected - run full preflight');
    return false;
  }
}