const { test, expect } = require('@playwright/test');

/**
 * Manual Screenshot Test - Get screenshot of working presentation
 */

test.describe('Manual Screenshot Test', () => {
  
  test.beforeEach(async ({ page }) => {
    test.setTimeout(120000); // 2 minutes for screenshot
  });

  test('Take screenshot of existing presentation', async ({ page }) => {
    console.log('📸 Taking screenshot of existing presentation...');
    
    // Use the presentation ID from the previous test
    const presentationId = '1MCHfRQz1r-wOSlSvAkSZ18h6w0VoaNvJzL8Q2MyKQxE';
    
    // Try public URL first (no authentication needed)
    const pubUrl = `https://docs.google.com/presentation/d/${presentationId}/pub`;
    console.log(`📖 Opening public URL: ${pubUrl}`);
    
    await page.goto(pubUrl);
    
    // Wait for content to load (more flexible selectors)
    await page.waitForLoadState('networkidle');
    await page.waitForTimeout(5000); // Allow time for full rendering
    
    // Take screenshot of the public view
    await page.screenshot({ 
      path: `test/screenshots/working-presentation-${Date.now()}.png`,
      fullPage: true 
    });
    
    console.log('✅ Public view screenshot captured');
    
    // Also try edit URL for more detailed view
    try {
      const editUrl = `https://docs.google.com/presentation/d/${presentationId}/edit`;
      console.log(`📝 Opening edit URL: ${editUrl}`);
      
      await page.goto(editUrl);
      await page.waitForLoadState('networkidle');
      await page.waitForTimeout(5000);
      
      // Take screenshot of edit view
      await page.screenshot({ 
        path: `test/screenshots/working-presentation-edit-${Date.now()}.png`,
        fullPage: true 
      });
      
      console.log('✅ Edit view screenshot captured');
      
    } catch (error) {
      console.log('⚠️ Edit view not accessible (may require authentication):', error.message);
    }
    
    console.log('✅ Screenshots completed successfully');
  });
});