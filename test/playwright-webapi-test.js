/**
 * Playwright Web API Test with Screenshots
 * 
 * Tests the OOXML system by calling the web API and capturing screenshots
 * of the results in Google Slides.
 */

const { chromium } = require('playwright');
const path = require('path');

// Configuration
const config = {
  // Replace with your actual Google Apps Script web app URL
  webAppUrl: 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec',
  
  // Test scenarios
  testScenarios: [
    {
      name: 'extension-system-test',
      functionName: 'testExtensionSystem',
      description: 'Test the extension auto-loading system'
    },
    {
      name: 'fflate-service-test', 
      functionName: 'testFFlatePPTXService',
      description: 'Test FFlatePPTXService functionality'
    },
    {
      name: 'create-from-prompt-test',
      functionName: 'testCreateFromPrompt',
      description: 'Test presentation creation from prompt'
    },
    {
      name: 'preflight-checks',
      functionName: 'runPreflightChecks',
      description: 'Run comprehensive system validation'
    }
  ],
  
  // Screenshot settings
  screenshots: {
    path: './test/screenshots/',
    fullPage: true,
    quality: 90
  }
};

async function runWebAPITests() {
  console.log('🎭 Starting Playwright Web API Tests...');
  console.log('=' .repeat(50));
  
  let browser;
  let results = {
    passed: 0,
    failed: 0,
    screenshots: [],
    errors: []
  };
  
  try {
    // Launch browser
    browser = await chromium.launch({ 
      headless: false, // Set to false to see the browser
      slowMo: 1000 // Slow down for visibility
    });
    
    const context = await browser.newContext({
      viewport: { width: 1920, height: 1080 }
    });
    
    const page = await context.newPage();
    
    console.log('🚀 Browser launched, running test scenarios...');
    
    for (const scenario of config.testScenarios) {
      console.log(`\n📋 Testing: ${scenario.description}`);
      
      try {
        // Call the web API
        const apiResult = await callWebAPI(scenario.functionName);
        console.log(`✅ API call successful for ${scenario.name}`);
        
        // Take screenshot of the API response
        await takeAPIResponseScreenshot(page, scenario, apiResult);
        
        // If the API returned a presentation URL, screenshot it
        if (apiResult.presentationUrl || apiResult.result?.presentationUrl) {
          await screenshotPresentation(page, scenario, apiResult);
        }
        
        results.passed++;
        
      } catch (error) {
        console.log(`❌ Test failed: ${scenario.name} - ${error.message}`);
        results.failed++;
        results.errors.push({
          scenario: scenario.name,
          error: error.message
        });
        
        // Take error screenshot
        await takeErrorScreenshot(page, scenario, error);
      }
    }
    
    // Generate test report with screenshots
    await generateTestReport(page, results);
    
  } catch (error) {
    console.error('💥 Test suite failed:', error);
    results.errors.push({ general: error.message });
    
  } finally {
    if (browser) {
      await browser.close();
    }
  }
  
  // Print final results
  console.log('\n📊 Test Results Summary:');
  console.log('=' .repeat(30));
  console.log(`✅ Passed: ${results.passed}`);
  console.log(`❌ Failed: ${results.failed}`);
  console.log(`📸 Screenshots: ${results.screenshots.length}`);
  console.log(`📁 Screenshot path: ${config.screenshots.path}`);
  
  if (results.errors.length > 0) {
    console.log('\n🚨 Errors:');
    results.errors.forEach(error => {
      console.log(`   • ${error.scenario || 'General'}: ${error.error}`);
    });
  }
  
  return results;
}

/**
 * Call the Google Apps Script Web API
 */
async function callWebAPI(functionName, params = {}) {
  console.log(`  🌐 Calling API: ${functionName}`);
  
  const payload = {
    fn: functionName,
    args: Object.values(params)
  };
  
  try {
    const response = await fetch(config.webAppUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(payload)
    });
    
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }
    
    const result = await response.json();
    
    if (result.error) {
      throw new Error(`API Error: ${result.error}`);
    }
    
    return result;
    
  } catch (error) {
    throw new Error(`API call failed: ${error.message}`);
  }
}

/**
 * Take screenshot of API response
 */
async function takeAPIResponseScreenshot(page, scenario, apiResult) {
  try {
    // Create a simple HTML page showing the API response
    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>API Response: ${scenario.name}</title>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; background: #f5f5f5; }
          .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; }
          .header { color: #333; border-bottom: 2px solid #4285f4; padding-bottom: 10px; }
          .success { color: #0f9d58; }
          .error { color: #ea4335; }
          .result { background: #f8f9fa; padding: 15px; margin: 15px 0; border-radius: 4px; }
          pre { background: #263238; color: #eeff41; padding: 15px; border-radius: 4px; overflow-x: auto; }
        </style>
      </head>
      <body>
        <div class="container">
          <h1 class="header">🧪 ${scenario.description}</h1>
          <p><strong>Function:</strong> ${scenario.functionName}</p>
          <p><strong>Status:</strong> <span class="success">✅ SUCCESS</span></p>
          <div class="result">
            <h3>API Response:</h3>
            <pre>${JSON.stringify(apiResult, null, 2)}</pre>
          </div>
        </div>
      </body>
      </html>
    `;
    
    await page.setContent(htmlContent);
    await page.waitForLoadState('networkidle');
    
    const screenshotPath = path.join(config.screenshots.path, `${scenario.name}-api-response.png`);
    await page.screenshot({ 
      path: screenshotPath,
      fullPage: config.screenshots.fullPage 
    });
    
    console.log(`  📸 API response screenshot: ${screenshotPath}`);
    
  } catch (error) {
    console.log(`  ⚠️ Failed to screenshot API response: ${error.message}`);
  }
}

/**
 * Take screenshot of created presentation
 */
async function screenshotPresentation(page, scenario, apiResult) {
  try {
    const presentationUrl = apiResult.presentationUrl || apiResult.result?.presentationUrl;
    
    if (!presentationUrl) {
      console.log('  ℹ️ No presentation URL to screenshot');
      return;
    }
    
    console.log(`  🎭 Taking screenshot of presentation: ${presentationUrl}`);
    
    // Navigate to the presentation
    await page.goto(presentationUrl);
    await page.waitForLoadState('networkidle');
    
    // Wait for slides to load
    await page.waitForTimeout(3000);
    
    const screenshotPath = path.join(config.screenshots.path, `${scenario.name}-presentation.png`);
    await page.screenshot({ 
      path: screenshotPath,
      fullPage: config.screenshots.fullPage 
    });
    
    console.log(`  📸 Presentation screenshot: ${screenshotPath}`);
    
  } catch (error) {
    console.log(`  ⚠️ Failed to screenshot presentation: ${error.message}`);
  }
}

/**
 * Take screenshot of error state
 */
async function takeErrorScreenshot(page, scenario, error) {
  try {
    const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>Test Error: ${scenario.name}</title>
        <style>
          body { font-family: Arial, sans-serif; padding: 20px; background: #ffebee; }
          .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; }
          .header { color: #d32f2f; border-bottom: 2px solid #f44336; padding-bottom: 10px; }
          .error { background: #ffcdd2; padding: 15px; margin: 15px 0; border-radius: 4px; }
          pre { background: #263238; color: #ff5722; padding: 15px; border-radius: 4px; }
        </style>
      </head>
      <body>
        <div class="container">
          <h1 class="header">❌ Test Failed: ${scenario.description}</h1>
          <p><strong>Function:</strong> ${scenario.functionName}</p>
          <div class="error">
            <h3>Error Details:</h3>
            <pre>${error.message}</pre>
          </div>
        </div>
      </body>
      </html>
    `;
    
    await page.setContent(htmlContent);
    await page.waitForLoadState('networkidle');
    
    const screenshotPath = path.join(config.screenshots.path, `${scenario.name}-error.png`);
    await page.screenshot({ 
      path: screenshotPath,
      fullPage: config.screenshots.fullPage 
    });
    
    console.log(`  📸 Error screenshot: ${screenshotPath}`);
    
  } catch (screenshotError) {
    console.log(`  ⚠️ Failed to take error screenshot: ${screenshotError.message}`);
  }
}

/**
 * Generate comprehensive test report
 */
async function generateTestReport(page, results) {
  try {
    const timestamp = new Date().toISOString();
    
    const reportHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>OOXML Web API Test Report</title>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin: 0; padding: 20px; background: #f5f5f5; }
          .container { max-width: 1400px; margin: 0 auto; }
          .header { background: #4285f4; color: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
          .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px; }
          .stat-card { background: white; padding: 20px; border-radius: 8px; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
          .stat-value { font-size: 2em; font-weight: bold; margin-bottom: 10px; }
          .passed { color: #0f9d58; }
          .failed { color: #ea4335; }
          .total { color: #4285f4; }
          .section { background: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
          .error-list { background: #ffebee; padding: 15px; border-radius: 4px; margin-top: 15px; }
          .footer { text-align: center; color: #666; margin-top: 30px; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h1>🧪 OOXML Web API Test Report</h1>
            <p>Generated: ${timestamp}</p>
          </div>
          
          <div class="summary">
            <div class="stat-card">
              <div class="stat-value passed">${results.passed}</div>
              <div>Tests Passed</div>
            </div>
            <div class="stat-card">
              <div class="stat-value failed">${results.failed}</div>
              <div>Tests Failed</div>
            </div>
            <div class="stat-card">
              <div class="stat-value total">${results.passed + results.failed}</div>
              <div>Total Tests</div>
            </div>
            <div class="stat-card">
              <div class="stat-value total">${results.screenshots.length}</div>
              <div>Screenshots</div>
            </div>
          </div>
          
          <div class="section">
            <h2>📊 Test Results</h2>
            <p><strong>Success Rate:</strong> ${((results.passed / (results.passed + results.failed)) * 100).toFixed(1)}%</p>
            
            ${results.errors.length > 0 ? `
            <div class="error-list">
              <h3>❌ Errors Encountered:</h3>
              <ul>
                ${results.errors.map(error => `
                  <li><strong>${error.scenario || 'General'}:</strong> ${error.error}</li>
                `).join('')}
              </ul>
            </div>
            ` : '<p style="color: #0f9d58;">✅ All tests passed successfully!</p>'}
          </div>
          
          <div class="section">
            <h2>📸 Screenshots</h2>
            <p>Screenshots saved to: <code>${config.screenshots.path}</code></p>
            <ul>
              <li>API response screenshots show the JSON results from each function call</li>
              <li>Presentation screenshots show any Google Slides created during testing</li>
              <li>Error screenshots capture any failures for debugging</li>
            </ul>
          </div>
          
          <div class="footer">
            <p>Generated by OOXML Slides Manipulator Test Suite</p>
          </div>
        </div>
      </body>
      </html>
    `;
    
    await page.setContent(reportHtml);
    await page.waitForLoadState('networkidle');
    
    const reportPath = path.join(config.screenshots.path, `test-report-${Date.now()}.png`);
    await page.screenshot({ 
      path: reportPath,
      fullPage: true 
    });
    
    console.log(`\n📋 Test report screenshot: ${reportPath}`);
    
  } catch (error) {
    console.log(`⚠️ Failed to generate test report: ${error.message}`);
  }
}

/**
 * Create a simple test function to call via Web API
 */
function createTestCreateFromPrompt() {
  return `
    // This function gets called via the Web API
    async function testCreateFromPrompt() {
      try {
        // Initialize extensions first
        await initializeExtensions();
        
        // Test creating a presentation from prompt
        const slides = await OOXMLSlides.createFromPrompt(
          "Create presentation with https://coolors.co/edd3c4-c8adc0-7765e3-3b60e4-080708 and Merriweather/Inter fonts"
        );
        
        return {
          success: true,
          message: "Presentation created successfully from prompt",
          presentationId: slides.fileId,
          presentationUrl: \`https://docs.google.com/presentation/d/\${slides.fileId}/edit\`
        };
        
      } catch (error) {
        return {
          success: false,
          error: error.message
        };
      }
    }
  `;
}

// Main execution
if (require.main === module) {
  // Create screenshots directory
  const fs = require('fs');
  if (!fs.existsSync(config.screenshots.path)) {
    fs.mkdirSync(config.screenshots.path, { recursive: true });
  }
  
  runWebAPITests()
    .then(results => {
      console.log('\n🎉 Test suite completed!');
      process.exit(results.failed > 0 ? 1 : 0);
    })
    .catch(error => {
      console.error('💥 Test suite crashed:', error);
      process.exit(1);
    });
}

module.exports = { runWebAPITests };