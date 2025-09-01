/**
 * SystemValidator - Entry point for all system validation
 * 
 * Provides simple functions to validate the entire OOXML system
 * including both FFlatePPTXService and FFlatePPTXService deployments.
 */

/**
 * Validate the entire OOXML system
 * This is the main function to call for comprehensive validation
 */
function validateSystem() {
  ConsoleFormatter.header('🔍 OOXML System Validation');
  
  // Run preflight checks
  const preflightResults = runPreflightChecks();
  
  // Run service-specific tests
  console.log('\n🧪 Running Service Tests...');
  const serviceResults = validateServices();
  
  // Generate overall report
  const overallStatus = generateValidationReport(preflightResults, serviceResults);
  
  return {
    preflight: preflightResults,
    services: serviceResults,
    overall: overallStatus
  };
}

/**
 * Validate both FFlatePPTXService and FFlatePPTXService
 */
function validateServices() {
  const results = {
    fflateService: null,
    cloudService: null,
    ooxmlCore: null
  };
  
  // Test FFlatePPTXService
  ConsoleFormatter.section('📦 FFlatePPTXService Validation');
  
  try {
    if (typeof validateFFlatePPTXService === 'function') {
      const isValid = validateFFlatePPTXService();
      results.fflateService = {
        available: true,
        valid: isValid,
        details: isValid ? 'All validation checks passed' : 'Validation issues detected'
      };
    } else {
      results.fflateService = {
        available: false,
        valid: false,
        details: 'validateFFlatePPTXService function not found'
      };
    }
  } catch (error) {
    results.fflateService = {
      available: true,
      valid: false,
      details: `Validation error: ${error.message}`
    };
  }
  
  // Test FFlatePPTXService
  ConsoleFormatter.section('☁️ FFlatePPTXService Validation');
  
  try {
    if (typeof validateFFlatePPTXService === 'function') {
      const isValid = validateFFlatePPTXService();
      results.cloudService = {
        available: true,
        valid: isValid,
        details: isValid ? 'All validation checks passed' : 'Validation issues detected'
      };
    } else {
      results.cloudService = {
        available: false,
        valid: false,
        details: 'validateFFlatePPTXService function not found'
      };
    }
  } catch (error) {
    results.cloudService = {
      available: true,
      valid: false,
      details: `Validation error: ${error.message}`
    };
  }
  
  // Test OOXMLCore
  ConsoleFormatter.section('🏗️ OOXMLCore Validation');
  
  try {
    if (typeof validateOOXMLCore === 'function') {
      const isValid = validateOOXMLCore();
      results.ooxmlCore = {
        available: true,
        valid: isValid,
        details: isValid ? 'All validation checks passed' : 'Validation issues detected'
      };
    } else {
      results.ooxmlCore = {
        available: false,
        valid: false,
        details: 'validateOOXMLCore function not found'
      };
    }
  } catch (error) {
    results.ooxmlCore = {
      available: true,
      valid: false,
      details: `Validation error: ${error.message}`
    };
  }
  
  // Print results
  Object.entries(results).forEach(([service, result]) => {
    if (result) {
      const icon = result.valid ? '✅' : '❌';
      const status = result.available ? 'Available' : 'Missing';
      console.log(`  ${icon} ${service}: ${status} - ${result.details}`);
    }
  });
  
  return results;
}

/**
 * Generate overall validation report
 */
function generateValidationReport(preflightResults, serviceResults) {
  ConsoleFormatter.header('📋 VALIDATION SUMMARY');
  
  // Count preflight results
  let preflightPassed = 0;
  let preflightTotal = 0;
  let preflightFailed = 0;
  
  if (preflightResults) {
    Object.values(preflightResults).forEach(category => {
      if (Array.isArray(category)) {
        category.forEach(check => {
          preflightTotal++;
          if (check.status === 'PASS') {
            preflightPassed++;
          } else if (check.status === 'FAIL' || check.status === 'ERROR') {
            preflightFailed++;
          }
        });
      }
    });
  }
  
  // Count service results
  const servicesValid = Object.values(serviceResults).filter(r => r && r.valid).length;
  const servicesTotal = Object.values(serviceResults).filter(r => r && r.available).length;
  
  // Overall status
  const preflightHealthy = preflightFailed === 0;
  const servicesHealthy = servicesValid >= 2; // At least 2 services should be working
  const systemHealthy = preflightHealthy && servicesHealthy;
  
  const status = {
    overall: systemHealthy ? 'HEALTHY' : 'ISSUES_DETECTED',
    preflight: {
      passed: preflightPassed,
      total: preflightTotal,
      failed: preflightFailed,
      healthy: preflightHealthy
    },
    services: {
      valid: servicesValid,
      total: servicesTotal,
      healthy: servicesHealthy
    }
  };
  
  // Print summary
  console.log(`\n🎯 Overall Status: ${systemHealthy ? '✅ SYSTEM HEALTHY' : '❌ ISSUES DETECTED'}`);
  console.log(`\n📊 Details:`);
  console.log(`   Preflight: ${preflightPassed}/${preflightTotal} passed (${preflightFailed} failed)`);
  console.log(`   Services: ${servicesValid}/${servicesTotal} valid`);
  
  if (!systemHealthy) {
    console.log(`\n🔧 Recommendations:`);
    if (!preflightHealthy) {
      console.log(`   • Fix preflight failures before deployment`);
    }
    if (!servicesHealthy) {
      console.log(`   • Ensure at least FFlatePPTXService OR FFlatePPTXService is working`);
      console.log(`   • Check service deployment and configuration`);
    }
  }
  
  return status;
}

/**
 * Quick system health check
 */
function quickHealthCheck() {
  ConsoleFormatter.section('⚡ Quick Health Check', '-');
  
  const criticalServices = [
    { name: 'FFlatePPTXService', check: () => typeof FFlatePPTXService !== 'undefined' },
    { name: 'FFlatePPTXService', check: () => typeof FFlatePPTXService !== 'undefined' },
    { name: 'OOXMLCore', check: () => typeof OOXMLCore !== 'undefined' },
    { name: 'Drive API', check: () => typeof DriveApp !== 'undefined' },
    { name: 'Slides API', check: () => typeof Slides !== 'undefined' }
  ];
  
  const results = criticalServices.map(service => {
    try {
      const available = service.check();
      console.log(`  ${available ? '✅' : '❌'} ${service.name}`);
      return available;
    } catch (error) {
      console.log(`  ❌ ${service.name} (Error: ${error.message})`);
      return false;
    }
  });
  
  const healthy = results.filter(r => r).length;
  const total = results.length;
  
  console.log(`\n📊 ${healthy}/${total} services healthy`);
  
  if (healthy === total) {
    console.log('✅ System appears healthy');
  } else if (healthy >= 3) {
    console.log('⚠️ System partially functional');
  } else {
    console.log('❌ System has critical issues');
  }
  
  return healthy / total;
}

/**
 * Test specific deployment component
 */
function testDeploymentComponent(component) {
  console.log(`🔧 Testing ${component}...`);
  
  switch (component.toLowerCase()) {
    case 'fflate':
    case 'fflatepptxservice':
      return testFFlatePPTXService();
      
    case 'cloud':
    case 'cloudfunction':
    case 'cloudpptxservice':
      return testFFlatePPTXService();
      
    case 'ooxml':
    case 'ooxmlcore':
      return testOOXMLCore();
      
    case 'iam':
      return checkIAMConfiguration();
      
    case 'network':
      return checkNetworkConnectivity();
      
    default:
      console.log(`❌ Unknown component: ${component}`);
      return false;
  }
}

/**
 * Test FFlatePPTXService specifically
 */
function testFFlatePPTXService() {
  try {
    if (typeof FFlatePPTXService === 'undefined') {
      console.log('❌ FFlatePPTXService not available');
      return false;
    }
    
    // Test service info
    const info = FFlatePPTXService.getServiceInfo();
    console.log(`✅ Service: ${info.service} v${info.version}`);
    
    // Test conversion utilities
    const testString = 'Hello OOXML!';
    const uint8 = FFlatePPTXService.stringToUint8Array(testString);
    const converted = FFlatePPTXService.uint8ArrayToString(uint8);
    
    if (converted === testString) {
      console.log('✅ String conversion working');
      return true;
    } else {
      console.log('❌ String conversion failed');
      return false;
    }
  } catch (error) {
    console.log(`❌ FFlatePPTXService error: ${error.message}`);
    return false;
  }
}

/**
 * Test FFlatePPTXService specifically
 */
function testFFlatePPTXService() {
  try {
    if (typeof FFlatePPTXService === 'undefined') {
      console.log('❌ FFlatePPTXService not available');
      return false;
    }
    
    // Test service info
    const info = FFlatePPTXService.getServiceInfo();
    console.log(`✅ Service: ${info.service} v${info.version}`);
    
    // Test cloud function connectivity (basic)
    const config = FFlatePPTXService._CONFIG;
    if (!config.cloudFunctionUrl) {
      console.log('❌ No Cloud Function URL configured');
      return false;
    }
    
    console.log(`✅ Cloud Function URL: ${config.cloudFunctionUrl}`);
    
    // Could add actual connectivity test here
    return true;
  } catch (error) {
    console.log(`❌ FFlatePPTXService error: ${error.message}`);
    return false;
  }
}

/**
 * Test OOXMLCore specifically
 */
function testOOXMLCore() {
  try {
    if (typeof OOXMLCore === 'undefined') {
      console.log('❌ OOXMLCore not available');
      return false;
    }
    
    console.log('✅ OOXMLCore available');
    
    // Could add functionality test here
    return true;
  } catch (error) {
    console.log(`❌ OOXMLCore error: ${error.message}`);
    return false;
  }
}