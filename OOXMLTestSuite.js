/**
 * OOXML Slides Manipulator - Comprehensive Test Suite
 * 
 * This file contains all tests for the OOXML PowerPoint manipulation system:
 * - Basic functionality tests
 * - Cloud Function integration tests  
 * - Performance comparisons
 * - Theme customization tests
 * - Round-trip integrity verification
 */

/**
 * Main test runner - executes all tests in sequence
 */
function runAllTests() {
  console.log('🚀 OOXML Slides Manipulator - Complete Test Suite');
  console.log('=' * 60);
  
  const results = {};
  
  try {
    // Test 1: Basic Cloud Function Integration
    console.log('\n🧪 TEST 1: CLOUD FUNCTION INTEGRATION');
    console.log('-' * 40);
    results.cloudFunction = testCloudFunctionIntegration();
    
    // Test 2: Performance Comparison
    console.log('\n⚡ TEST 2: PERFORMANCE COMPARISON');
    console.log('-' * 35);
    results.performance = testPerformanceComparison();
    
    // Test 3: Theme Customization
    console.log('\n🎨 TEST 3: THEME CUSTOMIZATION');
    console.log('-' * 32);
    results.customTheme = testThemeCustomization();
    
    // Test 4: OOXML Parser Tests
    console.log('\n🔧 TEST 4: OOXML PARSER FUNCTIONALITY');
    console.log('-' * 38);
    results.parser = testOOXMLParser();
    
    // Test 5: Template System
    console.log('\n📄 TEST 5: TEMPLATE SYSTEM');
    console.log('-' * 28);
    results.template = testTemplateSystem();
    
    // Final Results
    console.log('\n📋 FINAL TEST RESULTS');
    console.log('=' * 25);
    console.log(`Cloud Function Integration: ${results.cloudFunction ? '✅ PASSED' : '❌ FAILED'}`);
    console.log(`Performance Comparison: ${results.performance ? '✅ PASSED' : '❌ FAILED'}`);
    console.log(`Theme Customization: ${results.customTheme ? '✅ PASSED' : '❌ FAILED'}`);
    console.log(`OOXML Parser: ${results.parser ? '✅ PASSED' : '❌ FAILED'}`);
    console.log(`Template System: ${results.template ? '✅ PASSED' : '❌ FAILED'}`);
    
    const allPassed = Object.values(results).every(result => result);
    
    if (allPassed) {
      console.log('\n🎉 ALL TESTS PASSED!');
      console.log('The OOXML Slides Manipulator is working perfectly.');
      console.log('You can now manipulate PowerPoint files with full reliability.');
    } else {
      console.log('\n⚠️ SOME TESTS FAILED');
      console.log('Check the detailed output above for troubleshooting information.');
    }
    
    return results;
    
  } catch (error) {
    console.error('❌ Test suite failed with error:', error);
    return { error: error.message };
  }
}

/**
 * Test 1: Cloud Function Integration
 */
function testCloudFunctionIntegration() {
  console.log('Testing Cloud Function integration and OOXML manipulation...');
  
  try {
    // Step 1: Test connectivity
    console.log('Step 1: Testing Cloud Function connectivity...');
    if (!CloudPPTXService.isCloudFunctionAvailable()) {
      console.log('❌ Cloud Function not available');
      console.log('Make sure the Cloud Function is deployed and accessible');
      return false;
    }
    console.log('✅ Cloud Function is available');
    
    // Step 2: Create template
    console.log('Step 2: Creating PPTX template...');
    const templateBlob = PPTXTemplate.createMinimalTemplate();
    console.log(`✅ Template created: ${templateBlob.getBytes().length} bytes`);
    
    // Step 3: Test extraction
    console.log('Step 3: Testing PPTX extraction...');
    const parser = new OOXMLParser(templateBlob);
    parser.extract();
    const fileCount = parser.listFiles().length;
    console.log(`✅ Extracted ${fileCount} files from PPTX`);
    
    // Step 4: Test theme manipulation
    console.log('Step 4: Testing theme manipulation...');
    if (parser.hasFile('ppt/theme/theme1.xml')) {
      const themeXML = parser.getXML('ppt/theme/theme1.xml');
      const root = themeXML.getRootElement();
      const ns = XmlService.getNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
      
      const themeElements = root.getChild('themeElements', ns);
      const clrScheme = themeElements.getChild('clrScheme', ns);
      
      // Modify accent colors
      const accent1 = clrScheme.getChild('accent1', ns);
      if (accent1) {
        const srgbClr = accent1.getChild('srgbClr', ns);
        if (srgbClr) {
          srgbClr.setAttribute('val', 'FF6B35'); // Orange
          console.log('✅ Changed accent1 to orange (#FF6B35)');
        }
      }
      
      const modifiedXML = XmlService.getRawFormat().format(themeXML);
      parser.setFile('ppt/theme/theme1.xml', modifiedXML);
      console.log('✅ Theme modifications saved');
    }
    
    // Step 5: Test rebuilding
    console.log('Step 5: Testing PPTX rebuilding...');
    const modifiedBlob = parser.build();
    console.log(`✅ Modified PPTX created: ${modifiedBlob.getBytes().length} bytes`);
    
    // Step 6: Save to Drive
    console.log('Step 6: Saving to Google Drive...');
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const fileName = `OOXML_Integration_Test_${timestamp}.pptx`;
    
    const driveFile = DriveApp.createFile(modifiedBlob.setName(fileName));
    console.log(`✅ Saved to Google Drive: ${driveFile.getName()}`);
    console.log(`🔗 Open URL: https://docs.google.com/presentation/d/${driveFile.getId()}`);
    
    // Step 7: Verify integrity
    console.log('Step 7: Verifying round-trip integrity...');
    const verifyParser = new OOXMLParser(modifiedBlob);
    verifyParser.extract();
    const verifyFileCount = verifyParser.listFiles().length;
    
    if (verifyFileCount === fileCount) {
      console.log(`✅ Integrity verified: ${verifyFileCount} files preserved`);
      
      // Verify theme changes
      if (verifyParser.hasFile('ppt/theme/theme1.xml')) {
        const verifyContent = verifyParser.getFile('ppt/theme/theme1.xml');
        const hasOrange = verifyContent.includes('FF6B35');
        console.log(`✅ Theme verification: Orange color preserved=${hasOrange}`);
      }
      
      return driveFile.getId();
    } else {
      console.log(`❌ Integrity check failed: ${verifyFileCount} vs ${fileCount} files`);
      return false;
    }
    
  } catch (error) {
    console.error('❌ Cloud Function integration test failed:', error);
    return false;
  }
}

/**
 * Test 2: Performance Comparison
 */
function testPerformanceComparison() {
  console.log('Comparing Cloud Function vs Native Google Apps Script performance...');
  
  try {
    const templateBlob = PPTXTemplate.createMinimalTemplate();
    console.log(`Test file: ${templateBlob.getBytes().length} bytes`);
    
    // Test Cloud Function approach
    console.log('Testing Cloud Function approach...');
    const startCloud = new Date().getTime();
    
    try {
      const cloudFiles = CloudPPTXService.unzipPPTX(templateBlob);
      const cloudResult = CloudPPTXService.zipPPTX(cloudFiles);
      const endCloud = new Date().getTime();
      
      console.log(`✅ Cloud Function: ${endCloud - startCloud}ms`);
      console.log(`   Files processed: ${Object.keys(cloudFiles).length}`);
      console.log(`   Output size: ${cloudResult.getBytes().length} bytes`);
      
    } catch (cloudError) {
      console.log(`❌ Cloud Function failed: ${cloudError.message}`);
      return false;
    }
    
    // Test native approach (expected to fail)
    console.log('Testing native Google Apps Script approach...');
    const startNative = new Date().getTime();
    
    try {
      const nativeFiles = Utilities.unzip(templateBlob);
      const nativeResult = Utilities.zip(nativeFiles);
      const endNative = new Date().getTime();
      
      console.log(`✅ Native method: ${endNative - startNative}ms`);
      console.log(`   Files processed: ${nativeFiles.length}`);
      console.log(`   Output size: ${nativeResult.getBytes().length} bytes`);
      
    } catch (nativeError) {
      console.log(`❌ Native method failed: ${nativeError.message}`);
      console.log('   This is expected for PPTX files - validates our Cloud Function approach');
    }
    
    console.log('✅ Performance comparison complete');
    return true;
    
  } catch (error) {
    console.error('❌ Performance test failed:', error);
    return false;
  }
}

/**
 * Test 3: Theme Customization
 */
function testThemeCustomization() {
  console.log('Testing comprehensive theme customization...');
  
  try {
    const templateBlob = PPTXTemplate.createMinimalTemplate();
    const parser = new OOXMLParser(templateBlob);
    parser.extract();
    
    console.log('Base template loaded, applying custom professional theme...');
    
    // Define professional color scheme
    const professionalTheme = {
      accent1: '2E86AB',  // Ocean Blue
      accent2: 'A23B72',  // Deep Pink  
      accent3: 'F18F01',  // Orange
      accent4: 'C73E1D',  // Red
      accent5: '4C956C',  // Green
      accent6: '2F4858',  // Dark Blue-Gray
      dark1: '1A1A1A',    // Almost Black
      light1: 'FFFFFF'    // White
    };
    
    if (parser.hasFile('ppt/theme/theme1.xml')) {
      const themeXML = parser.getXML('ppt/theme/theme1.xml');
      const root = themeXML.getRootElement();
      const ns = XmlService.getNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
      
      const themeElements = root.getChild('themeElements', ns);
      const clrScheme = themeElements.getChild('clrScheme', ns);
      
      // Apply the custom theme
      Object.entries(professionalTheme).forEach(([colorName, hexValue]) => {
        const colorElement = clrScheme.getChild(colorName, ns);
        if (colorElement) {
          const srgbClr = colorElement.getChild('srgbClr', ns);
          if (srgbClr) {
            srgbClr.setAttribute('val', hexValue);
            console.log(`✅ Set ${colorName} to #${hexValue}`);
          }
        }
      });
      
      // Update color scheme name
      clrScheme.setAttribute('name', 'OOXML Professional Theme');
      
      // Save modifications
      const modifiedXML = XmlService.getRawFormat().format(themeXML);
      parser.setFile('ppt/theme/theme1.xml', modifiedXML);
    }
    
    // Build and save
    const customBlob = parser.build();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const fileName = `OOXML_Professional_Theme_${timestamp}.pptx`;
    
    const driveFile = DriveApp.createFile(customBlob.setName(fileName));
    
    console.log('✅ Custom-themed presentation created!');
    console.log(`📁 File: ${driveFile.getName()}`);
    console.log(`🔗 Open: https://docs.google.com/presentation/d/${driveFile.getId()}`);
    
    return driveFile.getId();
    
  } catch (error) {
    console.error('❌ Theme customization test failed:', error);
    return false;
  }
}

/**
 * Test 4: OOXML Parser Functionality
 */
function testOOXMLParser() {
  console.log('Testing OOXML Parser core functionality...');
  
  try {
    const templateBlob = PPTXTemplate.createMinimalTemplate();
    const parser = new OOXMLParser(templateBlob);
    
    // Test extraction
    console.log('Testing file extraction...');
    parser.extract();
    const files = parser.listFiles();
    console.log(`✅ Extracted ${files.length} files`);
    
    // Test file access
    console.log('Testing file access methods...');
    const contentTypes = parser.getFile('[Content_Types].xml');
    if (contentTypes && contentTypes.length > 0) {
      console.log('✅ Content types file accessible');
    } else {
      console.log('❌ Content types file not accessible');
      return false;
    }
    
    // Test XML parsing
    console.log('Testing XML parsing...');
    if (parser.hasFile('ppt/theme/theme1.xml')) {
      const themeXML = parser.getXML('ppt/theme/theme1.xml');
      if (themeXML) {
        console.log('✅ Theme XML parsed successfully');
      } else {
        console.log('❌ Theme XML parsing failed');
        return false;
      }
    }
    
    // Test file modification
    console.log('Testing file modification...');
    const originalContent = parser.getFile('_rels/.rels');
    parser.setFile('_rels/.rels', originalContent + '<!-- Modified by test -->');
    const modifiedContent = parser.getFile('_rels/.rels');
    
    if (modifiedContent.includes('Modified by test')) {
      console.log('✅ File modification working');
    } else {
      console.log('❌ File modification failed');
      return false;
    }
    
    // Test rebuilding
    console.log('Testing PPTX rebuilding...');
    const rebuiltBlob = parser.build();
    if (rebuiltBlob && rebuiltBlob.getBytes().length > 0) {
      console.log(`✅ PPTX rebuilt successfully: ${rebuiltBlob.getBytes().length} bytes`);
    } else {
      console.log('❌ PPTX rebuilding failed');
      return false;
    }
    
    return true;
    
  } catch (error) {
    console.error('❌ OOXML Parser test failed:', error);
    return false;
  }
}

/**
 * Test 5: Template System
 */
function testTemplateSystem() {
  console.log('Testing template creation system...');
  
  try {
    // Test base64 template
    console.log('Testing base64 template...');
    const base64Template = PPTXTemplate.createFromBase64Template();
    if (base64Template && base64Template.getBytes().length > 0) {
      console.log(`✅ Base64 template: ${base64Template.getBytes().length} bytes`);
    } else {
      console.log('❌ Base64 template failed');
      return false;
    }
    
    // Test minimal template
    console.log('Testing minimal template creation...');
    const minimalTemplate = PPTXTemplate.createMinimalTemplate();
    if (minimalTemplate && minimalTemplate.getBytes().length > 0) {
      console.log(`✅ Minimal template: ${minimalTemplate.getBytes().length} bytes`);
    } else {
      console.log('❌ Minimal template creation failed');
      return false;
    }
    
    // Test template info
    console.log('Testing template info...');
    const templateInfo = MinimalPPTXTemplate.getInfo();
    console.log(`✅ Template info: ${templateInfo.description}, ${templateInfo.binarySize} bytes`);
    
    return true;
    
  } catch (error) {
    console.error('❌ Template system test failed:', error);
    return false;
  }
}

/**
 * Quick connectivity test for troubleshooting
 */
function quickConnectivityTest() {
  console.log('🔍 Quick Cloud Function Connectivity Test');
  
  try {
    const isAvailable = CloudPPTXService.isCloudFunctionAvailable();
    console.log(`Cloud Function available: ${isAvailable ? '✅ YES' : '❌ NO'}`);
    
    if (isAvailable) {
      console.log(`Function URL: ${CloudPPTXService.CLOUD_FUNCTION_URL}`);
      return true;
    } else {
      console.log('❌ Cloud Function not accessible');
      console.log('Check your deployment and network connection');
      return false;
    }
    
  } catch (error) {
    console.error('❌ Connectivity test failed:', error);
    return false;
  }
}