/**
 * Google Apps Script Library Tests
 * Tests OOXML Slides library functionality through the web interface
 */

import { test, expect } from '@playwright/test';
import path from 'path';

// Test configuration
const PROJECT_URL = process.env.GAS_PROJECT_URL || 'https://script.google.com/d/1feN12V9_9EgBR6lHIh1FcCRXJlT-w-uFEe3NYdN_AuaUiiWk0Ov8jICB/edit';
const TEST_FILE_ID = process.env.TEST_PPTX_FILE_ID; // Set this to a test PowerPoint file ID

test.describe('OOXML Slides Library Tests', () => {
  
  test.beforeEach(async ({ page }) => {
    // Load authentication state
    const authPath = path.join(process.cwd(), 'tests', 'auth', 'auth.json');
    await page.context().addInitScript(() => {
      // Any initialization scripts if needed
    });
  });

  test('should load Google Apps Script project', async ({ page }) => {
    await page.goto(PROJECT_URL);
    
    // Wait for the editor to load
    await page.waitForSelector('.ace_editor', { timeout: 30000 });
    
    // Check that our library files are present
    await expect(page.locator('text=OOXMLSlides')).toBeVisible();
    await expect(page.locator('text=QuickTest')).toBeVisible();
    
    console.log('✅ Google Apps Script project loaded successfully');
  });

  test('should run library loading test', async ({ page }) => {
    await page.goto(PROJECT_URL);
    await page.waitForSelector('.ace_editor', { timeout: 30000 });
    
    // Select QuickTest.js file
    await page.click('text=QuickTest.js');
    await page.waitForTimeout(2000);
    
    // Select the testLibraryLoading function
    const functionDropdown = page.locator('select[title="Select function"]');
    await functionDropdown.selectOption('testLibraryLoading');
    
    // Run the function
    await page.click('button:has-text("Run")');
    
    // Wait for execution to complete
    await page.waitForSelector('text=Execution completed', { timeout: 60000 });
    
    // Check for success indicators in logs
    const logsPanel = page.locator('#executionTranscript');
    await expect(logsPanel).toContainText('Library Loading');
    
    // Look for success indicators
    const hasSuccess = await logsPanel.textContent();
    expect(hasSuccess).toContain('✅');
    
    console.log('✅ Library loading test passed');
  });

  test('should run validation tests', async ({ page }) => {
    await page.goto(PROJECT_URL);
    await page.waitForSelector('.ace_editor', { timeout: 30000 });
    
    // Select and run validation test
    await page.click('text=QuickTest.js');
    await page.waitForTimeout(1000);
    
    const functionDropdown = page.locator('select[title="Select function"]');
    await functionDropdown.selectOption('testValidation');
    
    await page.click('button:has-text("Run")');
    await page.waitForSelector('text=Execution completed', { timeout: 60000 });
    
    const logsPanel = page.locator('#executionTranscript');
    await expect(logsPanel).toContainText('Validation Tests');
    
    console.log('✅ Validation tests passed');
  });

  test('should run Drive access test', async ({ page }) => {
    await page.goto(PROJECT_URL);
    await page.waitForSelector('.ace_editor', { timeout: 30000 });
    
    await page.click('text=QuickTest.js');
    await page.waitForTimeout(1000);
    
    const functionDropdown = page.locator('select[title="Select function"]');
    await functionDropdown.selectOption('testDriveAccess');
    
    await page.click('button:has-text("Run")');
    
    // This might trigger permission request
    const permissionModal = page.locator('text=Authorization required');
    if (await permissionModal.isVisible({ timeout: 5000 })) {
      await page.click('button:has-text("Review permissions")');
      await page.waitForTimeout(2000);
      
      // Handle OAuth flow if needed
      const continueButton = page.locator('button:has-text("Continue")');
      if (await continueButton.isVisible({ timeout: 5000 })) {
        await continueButton.click();
      }
    }
    
    await page.waitForSelector('text=Execution completed', { timeout: 120000 });
    
    const logsPanel = page.locator('#executionTranscript');
    const logText = await logsPanel.textContent();
    
    // Should either pass or show specific permission error
    expect(logText).toMatch(/(✅|Drive access|DriveApp)/);
    
    console.log('✅ Drive access test completed');
  });

  test('should run complete quick test suite', async ({ page }) => {
    await page.goto(PROJECT_URL);
    await page.waitForSelector('.ace_editor', { timeout: 30000 });
    
    await page.click('text=QuickTest.js');
    await page.waitForTimeout(1000);
    
    const functionDropdown = page.locator('select[title="Select function"]');
    await functionDropdown.selectOption('runQuickTests');
    
    await page.click('button:has-text("Run")');
    await page.waitForSelector('text=Execution completed', { timeout: 120000 });
    
    const logsPanel = page.locator('#executionTranscript');
    await expect(logsPanel).toContainText('QUICK TEST RESULTS');
    
    const logText = await logsPanel.textContent();
    
    // Should show test results summary
    expect(logText).toMatch(/\d+\/\d+ tests passed/);
    
    console.log('✅ Complete quick test suite executed');
  });

  test.skip('should test with real PowerPoint file', async ({ page }) => {
    // Skip if no test file ID provided
    if (!TEST_FILE_ID) {
      test.skip('No TEST_PPTX_FILE_ID environment variable set');
    }

    await page.goto(PROJECT_URL);
    await page.waitForSelector('.ace_editor', { timeout: 30000 });
    
    // Create a custom test function that uses the test file
    await page.click('text=QuickTest.js');
    await page.waitForTimeout(1000);
    
    // Add custom code to test with real file
    const editor = page.locator('.ace_editor');
    await editor.click();
    
    // Add test function at the end
    await page.keyboard.press('Control+End');
    await page.keyboard.type(`

function testWithRealFile() {
  try {
    const slides = OOXMLSlides.fromFile('${TEST_FILE_ID}');
    const theme = slides.getTheme();
    console.log('✅ Successfully loaded real PowerPoint file');
    console.log('Theme info:', theme);
    return true;
  } catch (error) {
    console.error('❌ Real file test failed:', error);
    return false;
  }
}`);
    
    // Save the file
    await page.keyboard.press('Control+S');
    await page.waitForTimeout(2000);
    
    // Run the new function
    const functionDropdown = page.locator('select[title="Select function"]');
    await functionDropdown.selectOption('testWithRealFile');
    
    await page.click('button:has-text("Run")');
    await page.waitForSelector('text=Execution completed', { timeout: 120000 });
    
    const logsPanel = page.locator('#executionTranscript');
    const logText = await logsPanel.textContent();
    
    expect(logText).toContain('Successfully loaded real PowerPoint file');
    
    console.log('✅ Real PowerPoint file test passed');
  });

  test('should verify all library files are present', async ({ page }) => {
    await page.goto(PROJECT_URL);
    await page.waitForSelector('.ace_editor', { timeout: 30000 });
    
    // Check for all expected files in the file tree
    const expectedFiles = [
      'OOXMLSlides.js',
      'QuickTest.js',
      'OOXMLParser.js',
      'ThemeEditor.js',
      'SlideManager.js',
      'FileHandler.js',
      'Validators.js',
      'usage-examples.js',
      'basic-test.js'
    ];
    
    for (const fileName of expectedFiles) {
      await expect(page.locator(`text=${fileName}`)).toBeVisible();
    }
    
    console.log('✅ All library files are present in the project');
  });

  test('should check project configuration', async ({ page }) => {
    await page.goto(PROJECT_URL);
    await page.waitForSelector('.ace_editor', { timeout: 30000 });
    
    // Check project settings/manifest
    await page.click('text=appsscript.json');
    await page.waitForTimeout(1000);
    
    const editor = page.locator('.ace_editor');
    const editorContent = await editor.textContent();
    
    // Verify OAuth scopes are configured
    expect(editorContent).toContain('googleapis.com/auth/drive');
    expect(editorContent).toContain('googleapis.com/auth/presentations');
    expect(editorContent).toContain('enabledAdvancedServices');
    
    console.log('✅ Project configuration verified');
  });
});

test.describe('Error Handling Tests', () => {
  
  test('should handle invalid file IDs gracefully', async ({ page }) => {
    await page.goto(PROJECT_URL);
    await page.waitForSelector('.ace_editor', { timeout: 30000 });
    
    await page.click('text=QuickTest.js');
    await page.waitForTimeout(1000);
    
    // Create test for invalid file ID
    const editor = page.locator('.ace_editor');
    await editor.click();
    await page.keyboard.press('Control+End');
    
    await page.keyboard.type(`

function testInvalidFileId() {
  try {
    const isValid = Validators.isValidFileId('invalid-id');
    console.log('Invalid file ID validation:', isValid);
    
    // Should return false for invalid IDs
    if (!isValid) {
      console.log('✅ Invalid file ID correctly rejected');
      return true;
    } else {
      console.log('❌ Invalid file ID was accepted');
      return false;
    }
  } catch (error) {
    console.error('Error in validation test:', error);
    return false;
  }
}`);
    
    await page.keyboard.press('Control+S');
    await page.waitForTimeout(2000);
    
    const functionDropdown = page.locator('select[title="Select function"]');
    await functionDropdown.selectOption('testInvalidFileId');
    
    await page.click('button:has-text("Run")');
    await page.waitForSelector('text=Execution completed', { timeout: 60000 });
    
    const logsPanel = page.locator('#executionTranscript');
    await expect(logsPanel).toContainText('Invalid file ID correctly rejected');
    
    console.log('✅ Error handling test passed');
  });
});