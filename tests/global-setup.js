/**
 * Global setup for Playwright tests
 * Handles authentication and initial configuration
 */

import { chromium } from '@playwright/test';
import fs from 'fs';
import path from 'path';

async function globalSetup() {
  console.log('🚀 Starting global setup for OOXML Slides tests');
  
  // Create auth directory if it doesn't exist
  const authDir = path.join(process.cwd(), 'tests', 'auth');
  if (!fs.existsSync(authDir)) {
    fs.mkdirSync(authDir, { recursive: true });
  }

  // Launch browser for authentication
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext();
  const page = await context.newPage();

  try {
    // Navigate to Google Apps Script
    await page.goto('https://script.google.com');
    
    // Check if already authenticated
    const isAuthenticated = await page.isVisible('text=New project', { timeout: 5000 }).catch(() => false);
    
    if (!isAuthenticated) {
      console.log('🔐 Authentication required. Please sign in manually...');
      
      // Wait for manual authentication
      await page.waitForURL('**/home', { timeout: 300000 }); // 5 minutes timeout
      
      console.log('✅ Authentication successful');
    } else {
      console.log('✅ Already authenticated');
    }

    // Save authentication state
    await context.storageState({ path: path.join(authDir, 'auth.json') });
    console.log('💾 Authentication state saved');

    // Get project information if available
    const projectUrl = process.env.GAS_PROJECT_URL || 'https://script.google.com/d/1feN12V9_9EgBR6lHIh1FcCRXJlT-w-uFEe3NYdN_AuaUiiWk0Ov8jICB/edit';
    
    // Navigate to our project to verify it exists
    await page.goto(projectUrl);
    
    const projectExists = await page.isVisible('text=OOXML', { timeout: 10000 }).catch(() => false);
    
    if (projectExists) {
      console.log('✅ Project found and accessible');
    } else {
      console.log('⚠️ Project not found - tests may fail');
    }

  } catch (error) {
    console.error('❌ Global setup failed:', error);
    throw error;
  } finally {
    await browser.close();
  }

  console.log('🎉 Global setup completed successfully');
}

export default globalSetup;