# Console Formatting Style Guide

## Overview
Use `ConsoleFormatter` for all console output to ensure consistent, professional formatting across the OOXML system.

## Basic Usage

### Headers
```javascript
// Replace this:
console.log('🚀 My Function Name');
console.log('=' + '='.repeat(20));

// With this:
ConsoleFormatter.header('🚀 My Function Name');
```

### Sections  
```javascript
// Replace this:
console.log('\n📦 Service Checks');
console.log('-'.repeat(15));

// With this:
ConsoleFormatter.section('📦 Service Checks');
```

### Status Messages
```javascript
// Replace this:
const icon = status === 'PASS' ? '✅' : '❌';
console.log(`  ${icon} ${name}: ${details}`);

// With this:
ConsoleFormatter.status(status, name, details);
```

### Test Results
```javascript
// Replace this:
console.log(`✅ Test passed: ${testName} (${duration}ms)`);

// With this:
ConsoleFormatter.testResult(testName, true, details, duration);
```

## Standard Patterns

### Function Entry Point
```javascript
function myTestFunction() {
  ConsoleFormatter.header('🧪 My Test Function');
  
  // function body...
  
  return results;
}
```

### Test Suite Structure
```javascript
function runTestSuite() {
  ConsoleFormatter.header('🧪 Test Suite Name');
  
  const testSuites = {
    'Validation Tests': testValidation,
    'Integration Tests': testIntegration
  };

  for (const [suiteName, suiteFunction] of Object.entries(testSuites)) {
    ConsoleFormatter.section(`📋 ${suiteName}`);
    const results = suiteFunction();
    
    results.forEach(result => {
      ConsoleFormatter.status(result.status, result.name, result.details);
    });
  }
}
```

### Progress Reporting
```javascript
// Show progress
ConsoleFormatter.progress(currentItem, totalItems, 'Processing files');

// Show summary
ConsoleFormatter.summary('Test Results', {
  passed: 25,
  failed: 2, 
  total: 27,
  duration: '1.2s'
});
```

### Error Handling
```javascript
// Replace this:
console.error('❌ Error:', error.message);

// With this:
ConsoleFormatter.error('Operation failed', error, 'Additional context');
```

### Warnings and Info
```javascript
// Warnings
ConsoleFormatter.warning('Deprecated feature used', 'Use newMethod() instead');

// Info messages  
ConsoleFormatter.info('Configuration loaded', 'SYSTEM');

// Success messages
ConsoleFormatter.success('All tests passed!');
```

## Status Values

Use these standard status values with `ConsoleFormatter.status()`:

- **PASS** - ✅ Test/check passed
- **FAIL** - ❌ Test/check failed  
- **WARN** - ⚠️ Warning condition
- **INFO** - ℹ️ Informational
- **ERROR** - 🚨 Error condition
- **ACTION** - 🔧 Action required
- **SKIP** - ⏭️ Skipped

## Files to Update

### Already Updated ✅
- `examples/PreflightChecks.js` - Updated section headers and status logging
- `examples/SystemValidator.js` - Updated headers and sections

### Still Need Updates ❓
- `test/TestFFlatePPTXService.js` - Test suite headers
- `test/TestOOXMLCore.js` - Test output formatting
- `test/TestCloudPPTXService.js` - Test result formatting
- `examples/AutoSetup.js` - Setup process headers
- `examples/AcidTestFramework.js` - Test framework output
- `src/TestOOXMLCoreRunner.js` - Test runner output
- All other files with console.log usage

## Migration Pattern

### Before (Inconsistent)
```javascript
console.log('=== Test Results ===');
console.log('Tests passed: ' + passed);
console.log('Tests failed: ' + failed);
console.log('✅ Success!');
```

### After (Consistent)
```javascript
ConsoleFormatter.summary('Test Results', {
  passed: passed,
  failed: failed,
  total: passed + failed
});

if (failed === 0) {
  ConsoleFormatter.success('All tests passed!');
}
```

## Benefits

✅ **Consistent**: Same formatting across all functions  
✅ **Professional**: Clean, structured output  
✅ **Maintainable**: Single place to change formatting  
✅ **Readable**: Clear visual hierarchy  
✅ **Extensible**: Easy to add new format types  

## Next Steps

1. Update remaining files to use ConsoleFormatter
2. Replace all manual console.log formatting
3. Use consistent status values and icons
4. Test the updated formatting in Google Apps Script
5. Document any project-specific formatting patterns

The ConsoleFormatter is now deployed and ready to use across the entire OOXML system!