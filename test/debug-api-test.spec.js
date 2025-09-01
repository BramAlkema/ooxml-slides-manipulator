const { test, expect } = require('@playwright/test');

/**
 * Debug API Test - Check if enhanced functions are working
 */

test.describe('API Debug Test', () => {
  
  const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbwsD1RQSgoRsF3tqGka2YZmhYlIMmDYKmnIlA5c9YqeiJ6OtpzgRw5C1e7teKacOEkw/exec';
  
  test.beforeEach(async ({ page }) => {
    test.setTimeout(60000);
  });

  test('Debug font pair API response', async ({ page }) => {
    console.log('🔍 Debug: Testing API response structure...');
    
    const apiResponse = await page.request.post(WEB_APP_URL, {
      data: {
        fn: 'createPresentationFromPrompt',
        args: ["Create presentation with Merriweather/Inter fonts"]
      }
    });
    
    expect(apiResponse.ok()).toBeTruthy();
    const response = await apiResponse.json();
    
    console.log('📋 Full API Response:');
    console.log(JSON.stringify(response, null, 2));
    
    // Check response structure
    if (response.result) {
      console.log('✅ Response has result wrapper');
      const result = response.result;
      
      console.log('📊 Result properties:', Object.keys(result));
      
      if (result.fontPair) {
        console.log('✅ FontPair data present:', result.fontPair);
      } else {
        console.log('❌ FontPair data missing');
      }
      
      if (result.colorPalette) {
        console.log('✅ ColorPalette data present:', result.colorPalette);
      } else {
        console.log('❌ ColorPalette data missing');
      }
      
      if (result.message) {
        console.log('📝 Message:', result.message);
      }
      
    } else if (response.error) {
      console.log('❌ API returned error:', response.error);
    } else {
      console.log('⚠️ Unexpected response structure');
    }
  });
  
  test('Test ping function', async ({ page }) => {
    console.log('🏓 Testing basic ping function...');
    
    const apiResponse = await page.request.post(WEB_APP_URL, {
      data: {
        fn: 'ping',
        args: ['DebugTest']
      }
    });
    
    expect(apiResponse.ok()).toBeTruthy();
    const response = await apiResponse.json();
    
    console.log('🏓 Ping Response:', JSON.stringify(response, null, 2));
    
    expect(response.result).toBeTruthy();
    expect(response.result.message).toBe('pong:DebugTest');
  });
  
  test('Test font parsing function directly', async ({ page }) => {
    console.log('🧪 Testing font parsing function directly...');
    
    const apiResponse = await page.request.post(WEB_APP_URL, {
      data: {
        fn: 'testFontParsing',
        args: ["Create presentation with https://coolors.co/edd3c4-c8adc0-7765e3-3b60e4-080708 and Merriweather/Inter fonts"]
      }
    });
    
    expect(apiResponse.ok()).toBeTruthy();
    const response = await apiResponse.json();
    
    console.log('🧪 Font Parsing Test Response:', JSON.stringify(response, null, 2));
    
    if (response.result) {
      const result = response.result;
      
      if (result.success) {
        console.log('✅ Font parsing test succeeded');
        
        if (result.fontPair) {
          console.log('✅ FontPair parsed:', result.fontPair);
        } else {
          console.log('❌ FontPair parsing failed');
        }
        
        if (result.colorPalette) {
          console.log('✅ ColorPalette parsed:', result.colorPalette);
        } else {
          console.log('❌ ColorPalette parsing failed');
        }
      } else {
        console.log('❌ Font parsing test failed:', result.error);
        if (result.stack) {
          console.log('📚 Error stack:', result.stack);
        }
      }
    } else if (response.error) {
      console.log('❌ API error:', response.error);
    }
  });
});