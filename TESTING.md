# Testing Guide for OOXML Slides Library

This document explains how to test the OOXML Slides library using both manual and automated approaches.

## 🧪 Testing Strategy

The library uses a multi-layered testing approach:

1. **Unit Tests** - Internal Google Apps Script testing
2. **Integration Tests** - Playwright-driven browser automation
3. **End-to-End Tests** - Complete workflow testing
4. **Manual Testing** - Interactive verification

## 🚀 Quick Start Testing

### 1. Manual Testing in Google Apps Script

```javascript
// In Google Apps Script console
runQuickTests(); // Basic functionality
runAllTests();   // Comprehensive testing
```

### 2. Automated Testing with Playwright

```bash
# Install dependencies
npm install

# Install Playwright browsers
npm run install-playwright

# Run tests
npm test                 # All tests
npm run test:headed     # With visible browser
npm run test:ui         # Interactive UI mode
```

## 🔧 Setup for Testing

### Prerequisites

1. **Google Apps Script Project** - Set up via Clasp
2. **Test PowerPoint File** - Upload a .pptx to Google Drive
3. **Environment Variables** - Copy `.env.example` to `.env`

### Environment Configuration

```bash
# Copy example environment file
cp .env.example .env

# Edit with your values
GAS_PROJECT_URL=https://script.google.com/d/YOUR_PROJECT_ID/edit
TEST_PPTX_FILE_ID=your_google_drive_pptx_file_id_here
```

## 📋 Test Categories

### 1. Library Loading Tests
- Verify all classes are available
- Check dependency resolution
- Validate core functionality

```javascript
// Google Apps Script
testLibraryLoading()
```

### 2. Validation Tests
- Color format validation
- File ID format checking
- Input sanitization

```javascript
// Google Apps Script
testValidation()
```

### 3. Drive API Tests
- File access permissions
- Google Drive integration
- OAuth scope verification

```javascript
// Google Apps Script
testDriveAccess()
```

### 4. OOXML Parsing Tests
- ZIP file extraction
- XML parsing and manipulation
- File reconstruction

```javascript
// Google Apps Script
test2_OOXMLParsing()
```

### 5. Theme Modification Tests
- Color palette changes
- Font pair modifications
- Theme preset application

```javascript
// Google Apps Script
test3_ThemeEditing()
```

### 6. File Operations Tests
- PPTX file creation
- Theme modifications
- File saving and export

```javascript
// Google Apps Script
test4_ModificationWorkflow()
```

## 🎭 Playwright Test Suites

### Browser Automation Tests

```bash
# Run specific test files
npx playwright test gas-library.spec.js      # Library tests
npx playwright test file-operations.spec.js  # File operation tests

# Run with specific browser
npx playwright test --project=chromium
npx playwright test --project=firefox
npx playwright test --project=webkit
```

### Test Features

1. **Authentication Handling** - Manages Google OAuth flow
2. **Project Navigation** - Automates Google Apps Script interface
3. **Function Execution** - Runs tests and captures results
4. **Error Detection** - Identifies and reports failures
5. **Cleanup** - Removes test files automatically

## 🔍 Test Results and Debugging

### Viewing Results

```bash
# Generate and view HTML report
npm run test:report

# Debug mode (step-by-step)
npm run test:debug
```

### Test Artifacts

- **Screenshots** - Captured on test failures
- **Videos** - Recorded for failed tests
- **Traces** - Detailed execution traces
- **Reports** - HTML and JUnit formats

### Common Test Patterns

```javascript
// Playwright test structure
test('should test library functionality', async ({ page }) => {
  // Navigate to Google Apps Script
  await page.goto(PROJECT_URL);
  
  // Wait for editor to load
  await page.waitForSelector('.ace_editor');
  
  // Select and run test function
  await page.selectOption('select', 'testFunction');
  await page.click('button:has-text("Run")');
  
  // Verify results
  await page.waitForSelector('text=Execution completed');
  const logs = await page.locator('#executionTranscript').textContent();
  expect(logs).toContain('✅');
});
```

## 🚨 Troubleshooting Tests

### Common Issues

1. **Authentication Failures**
   ```bash
   # Clear auth state and retry
   rm -rf tests/auth/
   npm test
   ```

2. **Permission Errors**
   - Enable Drive API in Google Apps Script
   - Check OAuth scopes in appsscript.json
   - Re-authorize permissions

3. **File Not Found**
   - Verify TEST_PPTX_FILE_ID in .env
   - Ensure file is accessible to your Google account
   - Check file is actual .pptx format

4. **Timeout Issues**
   ```bash
   # Increase timeout for slow operations
   PLAYWRIGHT_TIMEOUT=120000 npm test
   ```

### Debug Commands

```bash
# Run single test with debug info
npx playwright test --debug gas-library.spec.js

# Record new test
npx playwright codegen https://script.google.com

# Show trace viewer
npx playwright show-trace trace.zip
```

## 📊 Continuous Integration

### GitHub Actions

The project includes automated testing on:
- **Push to main/develop** - Full test suite
- **Pull requests** - Regression testing
- **Daily schedule** - Health checks

### Test Matrix

- **Browsers**: Chromium, Firefox, WebKit
- **Platforms**: Ubuntu (Linux)
- **Node versions**: 18.x

### CI Environment Variables

Set these in GitHub repository secrets:
```
GAS_PROJECT_URL
TEST_PPTX_FILE_ID
GOOGLE_CLIENT_EMAIL
GOOGLE_PRIVATE_KEY
```

## 🎯 Test Coverage

### Current Coverage Areas

- ✅ Library loading and initialization
- ✅ Input validation and sanitization
- ✅ Google Drive API integration
- ✅ OOXML parsing and manipulation
- ✅ Theme editing functionality
- ✅ File operations workflow
- ✅ Error handling and recovery

### Planned Coverage

- 🔄 Performance benchmarking
- 🔄 Memory usage testing
- 🔄 Large file handling
- 🔄 Concurrent operations
- 🔄 Edge case validation

## 📈 Performance Testing

### Benchmark Tests

```javascript
// Google Apps Script performance test
function benchmarkThemeModification() {
  const start = new Date();
  
  const slides = OOXMLSlides.fromFile(TEST_FILE_ID);
  slides.setColors(['#FF0000', '#00FF00', '#0000FF']);
  const fileId = slides.save();
  
  const duration = new Date() - start;
  console.log(`Theme modification took: ${duration}ms`);
}
```

### Metrics to Monitor

- **File parsing time** - OOXML extraction speed
- **Theme modification time** - Color/font changes
- **File building time** - ZIP reconstruction
- **Memory usage** - Peak memory consumption
- **API call count** - Drive API efficiency

## 🔒 Security Testing

### Security Checks

- **Input sanitization** - Prevent injection attacks
- **File validation** - Ensure safe file processing
- **Permission verification** - OAuth scope compliance
- **Secret management** - No hardcoded credentials

### Security Test Commands

```bash
# Run security scan
npm audit

# Check for secrets in code
grep -r "api_key\|secret\|password" --include="*.js" .
```

## 🎉 Test Best Practices

1. **Isolation** - Each test creates/cleans up its own files
2. **Idempotency** - Tests can run multiple times safely
3. **Clear naming** - Descriptive test and function names
4. **Error handling** - Graceful failure and cleanup
5. **Documentation** - Comments explain complex test logic

## 📝 Writing New Tests

### Google Apps Script Test Template

```javascript
function testNewFeature() {
  try {
    console.log('🧪 Testing new feature...');
    
    // Setup
    const testData = setupTestData();
    
    // Execute
    const result = executeFeature(testData);
    
    // Verify
    if (validateResult(result)) {
      console.log('✅ Test passed');
      return true;
    } else {
      console.log('❌ Test failed');
      return false;
    }
    
  } catch (error) {
    console.error('❌ Test error:', error);
    return false;
  } finally {
    // Cleanup
    cleanupTestData();
  }
}
```

### Playwright Test Template

```javascript
test('should test new feature', async ({ page }) => {
  await page.goto(PROJECT_URL);
  await page.waitForSelector('.ace_editor');
  
  // Your test steps here
  
  const result = await page.locator('#executionTranscript').textContent();
  expect(result).toContain('expected output');
});
```

This comprehensive testing strategy ensures the OOXML Slides library works reliably across different environments and use cases.