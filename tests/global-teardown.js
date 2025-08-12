/**
 * Global teardown for Playwright tests
 * Cleanup and final reporting
 */

import fs from 'fs';
import path from 'path';

async function globalTeardown() {
  console.log('🧹 Starting global teardown');

  try {
    // Clean up temporary test files if any
    const tempDir = path.join(process.cwd(), 'tests', 'temp');
    if (fs.existsSync(tempDir)) {
      fs.rmSync(tempDir, { recursive: true, force: true });
      console.log('🗑️ Cleaned up temporary files');
    }

    // Generate test summary
    const reportDir = path.join(process.cwd(), 'playwright-report');
    if (fs.existsSync(reportDir)) {
      console.log('📊 Test reports generated in playwright-report/');
    }

    console.log('✅ Global teardown completed');

  } catch (error) {
    console.error('❌ Global teardown failed:', error);
  }
}

export default globalTeardown;