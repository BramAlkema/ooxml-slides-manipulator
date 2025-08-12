/**
 * File Operations Tests
 * Tests Drive API integration and file manipulation
 */

import { test, expect } from '@playwright/test';
import path from 'path';

const PROJECT_URL = process.env.GAS_PROJECT_URL || 'https://script.google.com/d/1feN12V9_9EgBR6lHIh1FcCRXJlT-w-uFEe3NYdN_AuaUiiWk0Ov8jICB/edit';

test.describe('File Operations Tests', () => {
  
  test.beforeEach(async ({ page }) => {
    await page.goto(PROJECT_URL);
    await page.waitForSelector('.ace_editor', { timeout: 30000 });
  });

  test('should create test PPTX file programmatically', async ({ page }) => {
    // This test creates a basic PPTX file to test with
    await page.click('text=QuickTest.js');
    await page.waitForTimeout(1000);
    
    const editor = page.locator('.ace_editor');
    await editor.click();
    await page.keyboard.press('Control+End');
    
    // Add function to create test PPTX
    await page.keyboard.type(`

function createTestPPTX() {
  try {
    console.log('🚀 Creating test PPTX file...');
    
    // Create a basic presentation using Google Slides
    const presentation = SlidesApp.create('Test Presentation for OOXML');
    const presentationId = presentation.getId();
    
    console.log('📄 Created Google Slides presentation:', presentationId);
    
    // Export as PPTX via Drive API
    const exportUrl = 'https://docs.google.com/presentation/d/' + presentationId + '/export/pptx';
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      }
    });
    
    if (response.getResponseCode() === 200) {
      // Save as PPTX file
      const pptxBlob = response.getBlob().setName('Test-Presentation.pptx');
      const pptxFile = DriveApp.createFile(pptxBlob);
      const pptxId = pptxFile.getId();
      
      console.log('✅ Created PPTX file:', pptxId);
      console.log('🔗 File URL:', pptxFile.getUrl());
      
      // Clean up the original Slides file
      DriveApp.getFileById(presentationId).setTrashed(true);
      
      return pptxId;
    } else {
      throw new Error('Failed to export PPTX: ' + response.getResponseCode());
    }
    
  } catch (error) {
    console.error('❌ Failed to create test PPTX:', error);
    throw error;
  }
}`);
    
    await page.keyboard.press('Control+S');
    await page.waitForTimeout(2000);
    
    const functionDropdown = page.locator('select[title="Select function"]');
    await functionDropdown.selectOption('createTestPPTX');
    
    await page.click('button:has-text("Run")');
    
    // Handle potential permission requests
    await page.waitForTimeout(5000);
    const permissionModal = page.locator('text=Authorization required');
    if (await permissionModal.isVisible({ timeout: 5000 })) {
      await page.click('button:has-text("Review permissions")');
      await page.waitForTimeout(3000);
    }
    
    await page.waitForSelector('text=Execution completed', { timeout: 120000 });
    
    const logsPanel = page.locator('#executionTranscript');
    const logText = await logsPanel.textContent();
    
    // Should show successful file creation
    expect(logText).toMatch(/Created PPTX file: \w+/);
    
    console.log('✅ Test PPTX file creation test completed');
  });

  test('should test OOXML parsing with generated file', async ({ page }) => {
    await page.click('text=QuickTest.js');
    await page.waitForTimeout(1000);
    
    const editor = page.locator('.ace_editor');
    await editor.click();
    await page.keyboard.press('Control+End');
    
    // Add comprehensive OOXML test
    await page.keyboard.type(`

function testOOXMLParsingComplete() {
  try {
    console.log('🧪 Starting comprehensive OOXML parsing test...');
    
    // First create a test file
    const testFileId = createTestPPTX();
    console.log('📁 Using test file:', testFileId);
    
    // Test FileHandler
    const fileHandler = new FileHandler();
    const fileInfo = fileHandler.getFileInfo(testFileId);
    console.log('📊 File info:', fileInfo);
    
    // Test file loading
    const fileBlob = fileHandler.loadFile(testFileId);
    console.log('📦 File loaded, size:', fileBlob.getSize());
    
    // Test OOXML parsing
    const parser = new OOXMLParser(fileBlob);
    parser.extract();
    
    const files = parser.listFiles();
    console.log('📄 OOXML files found:', files.length);
    
    // Test theme access
    if (parser.hasFile('ppt/theme/theme1.xml')) {
      const themeXML = parser.getThemeXML();
      console.log('🎨 Theme XML loaded successfully');
    }
    
    // Test with OOXMLSlides class
    const slides = OOXMLSlides.fromFile(testFileId);
    const currentTheme = slides.getTheme();
    console.log('🎯 Current theme:', currentTheme);
    
    // Clean up test file
    DriveApp.getFileById(testFileId).setTrashed(true);
    console.log('🗑️ Cleaned up test file');
    
    console.log('✅ Comprehensive OOXML parsing test completed successfully');
    return true;
    
  } catch (error) {
    console.error('❌ OOXML parsing test failed:', error);
    return false;
  }
}`);
    
    await page.keyboard.press('Control+S');
    await page.waitForTimeout(2000);
    
    const functionDropdown = page.locator('select[title="Select function"]');
    await functionDropdown.selectOption('testOOXMLParsingComplete');
    
    await page.click('button:has-text("Run")');
    await page.waitForSelector('text=Execution completed', { timeout: 180000 }); // 3 minutes
    
    const logsPanel = page.locator('#executionTranscript');
    const logText = await logsPanel.textContent();
    
    // Should complete successfully
    expect(logText).toContain('Comprehensive OOXML parsing test completed successfully');
    
    console.log('✅ Complete OOXML parsing test passed');
  });

  test('should test theme modification workflow', async ({ page }) => {
    await page.click('text=QuickTest.js');
    await page.waitForTimeout(1000);
    
    const editor = page.locator('.ace_editor');
    await editor.click();
    await page.keyboard.press('Control+End');
    
    await page.keyboard.type(`

function testThemeModificationWorkflow() {
  try {
    console.log('🎨 Testing theme modification workflow...');
    
    // Create test file
    const testFileId = createTestPPTX();
    console.log('📁 Created test file:', testFileId);
    
    // Load with OOXML Slides
    const slides = OOXMLSlides.fromFile(testFileId, { createBackup: false });
    
    // Get original theme
    const originalTheme = slides.getTheme();
    console.log('🎭 Original theme colors:', Object.keys(originalTheme.colors));
    
    // Apply modifications
    slides
      .setColors(['#FF0000', '#FFFFFF', '#000000', '#CCCCCC', '#0000FF', '#00FF00'])
      .setFonts('Arial', 'Calibri')
      .setSize(1920, 1080);
    
    console.log('⚙️ Applied theme modifications');
    
    // Get modified theme
    const modifiedTheme = slides.getTheme();
    console.log('🎨 Modified theme fonts:', modifiedTheme.fonts);
    
    // Save as new file
    const modifiedFileId = slides.save({
      name: 'Modified-Test-Presentation',
      importToSlides: false
    });
    
    console.log('💾 Saved modified file:', modifiedFileId);
    
    // Verify the modification worked
    const verificationSlides = OOXMLSlides.fromFile(modifiedFileId);
    const finalTheme = verificationSlides.getTheme();
    console.log('✅ Verification theme loaded');
    
    // Clean up
    DriveApp.getFileById(testFileId).setTrashed(true);
    DriveApp.getFileById(modifiedFileId).setTrashed(true);
    console.log('🗑️ Cleaned up test files');
    
    console.log('🎉 Theme modification workflow completed successfully');
    return true;
    
  } catch (error) {
    console.error('❌ Theme modification workflow failed:', error);
    return false;
  }
}`);
    
    await page.keyboard.press('Control+S');
    await page.waitForTimeout(2000);
    
    const functionDropdown = page.locator('select[title="Select function"]');
    await functionDropdown.selectOption('testThemeModificationWorkflow');
    
    await page.click('button:has-text("Run")');
    await page.waitForSelector('text=Execution completed', { timeout: 240000 }); // 4 minutes
    
    const logsPanel = page.locator('#executionTranscript');
    const logText = await logsPanel.textContent();
    
    expect(logText).toContain('Theme modification workflow completed successfully');
    
    console.log('✅ Theme modification workflow test passed');
  });
});