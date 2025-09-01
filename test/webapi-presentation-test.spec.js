const { test, expect } = require('@playwright/test');

/**
 * Web API to Google Slides Integration Test
 * 
 * Tests the complete flow:
 * 1. Call GAS Web App API to create presentation
 * 2. Screenshot the generated presentation using /edit URL
 * 3. Screenshot the /pub URL for public view
 * 4. Verify presentation content
 */

test.describe('Web API Presentation Integration', () => {
  
  // Your actual Google Apps Script Web App URL
  const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwsD1RQSgoRsF3tqGka2YZmhYlIMmDYKmnIlA5c9YqeiJ6OtpzgRw5C1e7teKacOEkw/exec';
  
  test.beforeEach(async ({ page }) => {
    test.setTimeout(180000); // 3 minutes for presentation creation + screenshots
  });

  test('Create presentation via API and screenshot results', async ({ page }) => {
    console.log('🚀 Starting Web API presentation creation test...');
    
    // Step 1: Call the Web API to create a presentation
    console.log('📞 Calling Web API to create presentation...');
    
    const apiResponse = await page.request.post(WEB_APP_URL, {
      data: {
        fn: 'createPresentationFromPrompt',
        args: [
          "Create presentation with https://coolors.co/edd3c4-c8adc0-7765e3-3b60e4-080708 and Merriweather/Inter fonts"
        ]
      }
    });
    
    expect(apiResponse.ok()).toBeTruthy();
    const result = await apiResponse.json();
    
    console.log('📋 API Response:', result);
    
    // Verify API call was successful
    expect(result.success).toBeTruthy();
    expect(result.presentationId).toBeTruthy();
    
    const presentationId = result.presentationId;
    console.log(`✅ Presentation created with ID: ${presentationId}`);
    
    // Step 2: Screenshot the presentation in edit mode
    console.log('📸 Taking screenshot of presentation in edit mode...');
    
    const editUrl = `https://docs.google.com/presentation/d/${presentationId}/edit`;
    await page.goto(editUrl);
    
    // Wait for presentation to load
    await page.waitForSelector('.punch-viewer-content', { timeout: 60000 });
    await page.waitForTimeout(3000); // Allow additional load time
    
    // Take screenshot of edit view
    await page.screenshot({ 
      path: `test/screenshots/api-created-presentation-edit-${Date.now()}.png`,
      fullPage: true 
    });
    
    // Verify basic presentation structure
    await expect(page.locator('.punch-viewer-content')).toBeVisible();
    await expect(page.locator('.punch-filmstrip')).toBeVisible();
    
    console.log('✅ Edit view screenshot captured');
    
    // Step 3: Screenshot the presentation in public view
    console.log('📸 Taking screenshot of presentation in public view...');
    
    const pubUrl = `https://docs.google.com/presentation/d/${presentationId}/pub`;
    await page.goto(pubUrl);
    
    // Wait for public view to load
    await page.waitForSelector('.punch-present-read-only-content, .punch-viewer-content', { timeout: 30000 });
    await page.waitForTimeout(2000);
    
    // Take screenshot of public view
    await page.screenshot({ 
      path: `test/screenshots/api-created-presentation-pub-${Date.now()}.png`,
      fullPage: true 
    });
    
    console.log('✅ Public view screenshot captured');
    
    // Step 4: Verify custom colors are applied (if visible in DOM)
    // Check for any elements that might contain our custom colors
    const colorElements = await page.locator('[style*="edd3c4"], [style*="c8adc0"], [style*="7765e3"], [style*="3b60e4"], [style*="080708"]').count();
    
    if (colorElements > 0) {
      console.log(`✅ Found ${colorElements} elements with custom colors`);
    } else {
      console.log('ℹ️  Custom colors not detected in DOM (may be applied via OOXML theme)');
    }
  });
  
  test('Test system validation via API and screenshot', async ({ page }) => {
    console.log('🔍 Testing system validation via API...');
    
    // Call preflight checks via API
    const apiResponse = await page.request.post(WEB_APP_URL, {
      data: {
        fn: 'runPreflightChecks',
        args: []
      }
    });
    
    expect(apiResponse.ok()).toBeTruthy();
    const result = await apiResponse.json();
    
    console.log('📋 Preflight Results:', result);
    
    // Create HTML page showing the results
    const resultsHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>OOXML System Validation Results</title>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin: 20px; background: #f5f5f5; }
          .container { max-width: 1200px; margin: 0 auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
          .header { color: #1a73e8; border-bottom: 3px solid #1a73e8; padding-bottom: 15px; margin-bottom: 30px; }
          .section { margin: 20px 0; padding: 20px; border-radius: 8px; }
          .pass { background: #e8f5e8; border-left: 4px solid #34a853; }
          .fail { background: #fce8e6; border-left: 4px solid #ea4335; }
          .warn { background: #fff4e5; border-left: 4px solid #fbbc04; }
          pre { background: #f8f9fa; padding: 15px; border-radius: 4px; overflow-x: auto; font-size: 12px; }
          .status-icon { font-size: 1.2em; margin-right: 8px; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h1>🔍 OOXML System Validation Results</h1>
            <p>Generated: ${new Date().toISOString()}</p>
          </div>
          
          <div class="section ${result.success ? 'pass' : 'fail'}">
            <h2><span class="status-icon">${result.success ? '✅' : '❌'}</span>Overall Status</h2>
            <p><strong>Status:</strong> ${result.success ? 'PASSED' : 'FAILED'}</p>
          </div>
          
          <div class="section">
            <h2>📊 Detailed Results</h2>
            <pre>${JSON.stringify(result, null, 2)}</pre>
          </div>
        </div>
      </body>
      </html>
    `;
    
    await page.setContent(resultsHtml);
    await page.waitForLoadState('networkidle');
    
    // Screenshot the results
    await page.screenshot({ 
      path: `test/screenshots/system-validation-results-${Date.now()}.png`,
      fullPage: true 
    });
    
    console.log('✅ System validation screenshot captured');
  });
  
  test('Test extension system via API', async ({ page }) => {
    console.log('🧩 Testing extension system via API...');
    
    // Test extension initialization
    const initResponse = await page.request.post(WEB_APP_URL, {
      data: {
        fn: 'initializeExtensions',
        args: []
      }
    });
    
    expect(initResponse.ok()).toBeTruthy();
    const initResult = await initResponse.json();
    
    console.log('🔧 Extension initialization result:', initResult);
    
    // Test extension status
    const statusResponse = await page.request.post(WEB_APP_URL, {
      data: {
        fn: 'getExtensionStatus',
        args: []
      }
    });
    
    expect(statusResponse.ok()).toBeTruthy();
    const statusResult = await statusResponse.json();
    
    console.log('📊 Extension status:', statusResult);
    
    // Create extension status visualization
    const extensionHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>Extension System Status</title>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin: 20px; background: #f5f5f5; }
          .container { max-width: 1000px; margin: 0 auto; background: white; padding: 30px; border-radius: 12px; }
          .header { color: #1a73e8; border-bottom: 2px solid #1a73e8; padding-bottom: 15px; }
          .stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin: 30px 0; }
          .stat-card { background: #f8f9fa; padding: 20px; border-radius: 8px; text-align: center; }
          .stat-value { font-size: 2em; font-weight: bold; color: #1a73e8; }
          .extension-list { background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0; }
          .extension-item { background: white; padding: 10px; margin: 10px 0; border-radius: 4px; border-left: 4px solid #34a853; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h1>🧩 Extension System Status</h1>
            <p>Generated: ${new Date().toISOString()}</p>
          </div>
          
          <div class="stats">
            <div class="stat-card">
              <div class="stat-value">${statusResult.loaded || 0}</div>
              <div>Extensions Loaded</div>
            </div>
            <div class="stat-card">
              <div class="stat-value">${statusResult.totalMethods || 0}</div>
              <div>Methods Available</div>
            </div>
            <div class="stat-card">
              <div class="stat-value">${initResult.success ? '✅' : '❌'}</div>
              <div>Initialization</div>
            </div>
          </div>
          
          <div class="extension-list">
            <h3>📦 Loaded Extensions</h3>
            ${(statusResult.extensions || []).map(ext => 
              `<div class="extension-item">📦 ${ext}</div>`
            ).join('')}
          </div>
          
          <div style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin-top: 20px;">
            <h3>🔧 Initialization Details</h3>
            <pre>${JSON.stringify(initResult, null, 2)}</pre>
          </div>
          
          <div style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin-top: 20px;">
            <h3>📊 Status Details</h3>
            <pre>${JSON.stringify(statusResult, null, 2)}</pre>
          </div>
        </div>
      </body>
      </html>
    `;
    
    await page.setContent(extensionHtml);
    await page.waitForLoadState('networkidle');
    
    await page.screenshot({ 
      path: `test/screenshots/extension-system-status-${Date.now()}.png`,
      fullPage: true 
    });
    
    console.log('✅ Extension system screenshot captured');
    
    // Verify extensions are working
    expect(statusResult.loaded).toBeGreaterThan(0);
    expect(statusResult.totalMethods).toBeGreaterThan(0);
  });
  
  test('Test FFlatePPTXService migration', async ({ page }) => {
    console.log('🔄 Testing CloudPPTXService to FFlatePPTXService migration...');
    
    const migrationResponse = await page.request.post(WEB_APP_URL, {
      data: {
        fn: 'testCloudPPTXServiceMigration',
        args: []
      }
    });
    
    expect(migrationResponse.ok()).toBeTruthy();
    const migrationResult = await migrationResponse.json();
    
    console.log('📋 Migration test results:', migrationResult);
    
    // Create migration status visualization
    const migrationHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <title>CloudPPTXService Migration Status</title>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; margin: 20px; background: #f5f5f5; }
          .container { max-width: 1000px; margin: 0 auto; background: white; padding: 30px; border-radius: 12px; }
          .header { color: #1a73e8; border-bottom: 2px solid #1a73e8; padding-bottom: 15px; }
          .summary { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin: 30px 0; }
          .stat-card { background: #f8f9fa; padding: 20px; border-radius: 8px; text-align: center; }
          .stat-value { font-size: 2em; font-weight: bold; }
          .passed { color: #34a853; }
          .failed { color: #ea4335; }
          .test-result { background: #f8f9fa; padding: 15px; margin: 10px 0; border-radius: 8px; }
          .test-pass { border-left: 4px solid #34a853; }
          .test-fail { border-left: 4px solid #ea4335; }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h1>🔄 CloudPPTXService Migration Status</h1>
            <p>Migration from CloudPPTXService to FFlatePPTXService</p>
            <p>Generated: ${new Date().toISOString()}</p>
          </div>
          
          <div class="summary">
            <div class="stat-card">
              <div class="stat-value passed">${migrationResult.passed || 0}</div>
              <div>Tests Passed</div>
            </div>
            <div class="stat-card">
              <div class="stat-value failed">${migrationResult.failed || 0}</div>
              <div>Tests Failed</div>
            </div>
            <div class="stat-card">
              <div class="stat-value">${migrationResult.totalTests || 0}</div>
              <div>Total Tests</div>
            </div>
          </div>
          
          <div>
            <h3>📊 Test Results</h3>
            ${(migrationResult.details || []).map(test => `
              <div class="test-result ${test.passed ? 'test-pass' : 'test-fail'}">
                <strong>${test.passed ? '✅' : '❌'} ${test.name}</strong><br>
                <small>${test.details}</small>
              </div>
            `).join('')}
          </div>
          
          <div style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin-top: 20px;">
            <h3>🔧 Raw Results</h3>
            <pre>${JSON.stringify(migrationResult, null, 2)}</pre>
          </div>
        </div>
      </body>
      </html>
    `;
    
    await page.setContent(migrationHtml);
    await page.waitForLoadState('networkidle');
    
    await page.screenshot({ 
      path: `test/screenshots/migration-status-${Date.now()}.png`,
      fullPage: true 
    });
    
    console.log('✅ Migration status screenshot captured');
    
    // Verify migration was successful
    expect(migrationResult.success).toBeTruthy();
  });
});