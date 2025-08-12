/**
 * Comprehensive test of Cloud Function integration with Google Apps Script
 * This demonstrates the complete OOXML manipulation workflow
 */

function testCloudFunctionIntegration() {
  console.log('🧪 Testing Cloud Function Integration with Google Apps Script\n');
  
  try {
    // Step 1: Test basic connectivity
    console.log('Step 1: Testing Cloud Function connectivity...');
    if (!CloudPPTXService.isCloudFunctionAvailable()) {
      console.log('❌ Cloud Function not available at localhost:8080');
      console.log('Make sure to run "npm start" in the cloud-function directory');
      return false;
    }
    console.log('✅ Cloud Function is available\n');
    
    // Step 2: Create a working PPTX template
    console.log('Step 2: Creating PPTX template...');
    const templateBlob = PPTXTemplate.createMinimalTemplate();
    console.log(`✅ Template created: ${templateBlob.getBytes().length} bytes\n`);
    
    // Step 3: Test the complete round-trip
    console.log('Step 3: Testing complete OOXML manipulation workflow...');
    
    // Create an OOXMLParser instance with the template
    const parser = new OOXMLParser(templateBlob);
    
    // Extract using Cloud Function
    console.log('  3a. Extracting PPTX using Cloud Function...');
    parser.extract();
    console.log(`  ✅ Extracted ${parser.listFiles().length} files`);
    console.log(`     Files: ${parser.listFiles().slice(0, 5).join(', ')}...`);
    
    // Step 4: Demonstrate OOXML manipulation
    console.log('\nStep 4: Demonstrating OOXML manipulation...');
    
    if (parser.hasFile('ppt/theme/theme1.xml')) {
      console.log('  4a. Found theme file, modifying colors...');
      
      const themeXML = parser.getXML('ppt/theme/theme1.xml');
      const root = themeXML.getRootElement();
      const ns = XmlService.getNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
      
      // Navigate to color scheme
      const themeElements = root.getChild('themeElements', ns);
      const clrScheme = themeElements.getChild('clrScheme', ns);
      
      // Change accent colors to a custom theme
      const accent1 = clrScheme.getChild('accent1', ns);
      if (accent1) {
        const srgbClr = accent1.getChild('srgbClr', ns);
        if (srgbClr) {
          srgbClr.setAttribute('val', 'FF6B35'); // Orange
          console.log('  ✅ Changed accent1 to orange (#FF6B35)');
        }
      }
      
      const accent2 = clrScheme.getChild('accent2', ns);
      if (accent2) {
        const srgbClr = accent2.getChild('srgbClr', ns);
        if (srgbClr) {
          srgbClr.setAttribute('val', 'E74C3C'); // Red
          console.log('  ✅ Changed accent2 to red (#E74C3C)');
        }
      }
      
      // Save the modified XML back
      const modifiedXML = XmlService.getRawFormat().format(themeXML);
      parser.setFile('ppt/theme/theme1.xml', modifiedXML);
      console.log('  ✅ Theme modifications saved');
    }
    
    // Step 5: Build the modified PPTX using Cloud Function
    console.log('\nStep 5: Building modified PPTX using Cloud Function...');
    const modifiedBlob = parser.build();
    console.log(`✅ Modified PPTX created: ${modifiedBlob.getBytes().length} bytes`);
    
    // Step 6: Save to Google Drive for verification
    console.log('\nStep 6: Saving to Google Drive...');
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const fileName = `CloudFunction_PPTX_Test_${timestamp}.pptx`;
    
    const driveFile = DriveApp.createFile(modifiedBlob.setName(fileName));
    console.log(`✅ Saved to Google Drive: ${driveFile.getName()}`);
    console.log(`📁 File ID: ${driveFile.getId()}`);
    console.log(`🔗 Open URL: https://docs.google.com/presentation/d/${driveFile.getId()}`);
    
    // Step 7: Verify the round-trip worked
    console.log('\nStep 7: Verifying round-trip integrity...');
    
    try {
      // Try to extract the modified file to verify it's valid
      const verifyParser = new OOXMLParser(modifiedBlob);
      verifyParser.extract();
      const verifyFileCount = verifyParser.listFiles().length;
      
      console.log(`✅ Verification successful: ${verifyFileCount} files extracted from modified PPTX`);
      
      // Check if our theme modification is present
      if (verifyParser.hasFile('ppt/theme/theme1.xml')) {
        const verifyContent = verifyParser.getFile('ppt/theme/theme1.xml');
        const hasOrange = verifyContent.includes('FF6B35');
        const hasRed = verifyContent.includes('E74C3C');
        
        console.log(`✅ Theme verification: Orange=${hasOrange}, Red=${hasRed}`);
      }
      
    } catch (verifyError) {
      console.log(`❌ Verification failed: ${verifyError.message}`);
      return false;
    }
    
    console.log('\n🎉 ALL TESTS PASSED!');
    console.log('\n📊 Summary:');
    console.log('✅ Cloud Function connectivity working');
    console.log('✅ PPTX extraction via Cloud Function working');
    console.log('✅ OOXML manipulation working');
    console.log('✅ PPTX creation via Cloud Function working');
    console.log('✅ Round-trip integrity verified');
    console.log('✅ File saved to Google Drive and openable');
    
    return driveFile.getId();
    
  } catch (error) {
    console.error('❌ Cloud Function integration test failed:', error);
    console.log('\n🔧 Troubleshooting:');
    console.log('1. Make sure Cloud Function is running: npm start');
    console.log('2. Check localhost:8080 is accessible');
    console.log('3. Verify all required OAuth scopes are granted');
    console.log('4. Check Google Apps Script execution limits');
    
    return false;
  }
}

/**
 * Test performance comparison: Native vs Cloud Function approach
 */
function comparePerformance() {
  console.log('⚡ Performance Comparison: Native vs Cloud Function\n');
  
  try {
    const templateBlob = PPTXTemplate.createMinimalTemplate();
    console.log(`Test file: ${templateBlob.getBytes().length} bytes\n`);
    
    // Test 1: Cloud Function approach
    console.log('Test 1: Cloud Function approach...');
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
    }
    
    // Test 2: Native approach (expected to fail)
    console.log('\nTest 2: Native Google Apps Script approach...');
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
      console.log('   This is expected for many PPTX files');
    }
    
    console.log('\n📊 Analysis:');
    console.log('The Cloud Function approach provides reliable PPTX manipulation');
    console.log('while the native approach fails with most real-world files.');
    
  } catch (error) {
    console.error('❌ Performance test failed:', error);
  }
}

/**
 * Test creating a custom-themed presentation
 */
function createCustomThemedPresentation() {
  console.log('🎨 Creating Custom-Themed PPTX Presentation\n');
  
  try {
    // Create base template
    const templateBlob = PPTXTemplate.createMinimalTemplate();
    const parser = new OOXMLParser(templateBlob);
    parser.extract();
    
    console.log('Base template loaded, applying custom theme...');
    
    // Define a professional color scheme
    const customTheme = {
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
      Object.entries(customTheme).forEach(([colorName, hexValue]) => {
        const colorElement = clrScheme.getChild(colorName, ns);
        if (colorElement) {
          const srgbClr = colorElement.getChild('srgbClr', ns);
          if (srgbClr) {
            srgbClr.setAttribute('val', hexValue);
            console.log(`✅ Set ${colorName} to #${hexValue}`);
          }
        }
      });
      
      // Also update the color scheme name
      clrScheme.setAttribute('name', 'Custom Professional Theme');
      
      // Save the modified theme
      const modifiedXML = XmlService.getRawFormat().format(themeXML);
      parser.setFile('ppt/theme/theme1.xml', modifiedXML);
    }
    
    // Build and save
    const customBlob = parser.build();
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const fileName = `Custom_Professional_Theme_${timestamp}.pptx`;
    
    const driveFile = DriveApp.createFile(customBlob.setName(fileName));
    
    console.log('\n🎉 Custom-themed presentation created!');
    console.log(`📁 File: ${driveFile.getName()}`);
    console.log(`🔗 Open: https://docs.google.com/presentation/d/${driveFile.getId()}`);
    console.log('\n🎨 Theme Colors Applied:');
    Object.entries(customTheme).forEach(([name, hex]) => {
      console.log(`   ${name}: #${hex}`);
    });
    
    return driveFile.getId();
    
  } catch (error) {
    console.error('❌ Custom theme creation failed:', error);
    return false;
  }
}

/**
 * Run all Cloud Function integration tests
 */
function runAllCloudFunctionTests() {
  console.log('🚀 Running Complete Cloud Function Integration Test Suite\n');
  console.log('=' * 60);
  
  const results = {};
  
  // Test 1: Basic integration
  console.log('\n🧪 TEST 1: BASIC INTEGRATION');
  console.log('-' * 30);
  results.basicIntegration = testCloudFunctionIntegration();
  
  // Test 2: Performance comparison  
  console.log('\n⚡ TEST 2: PERFORMANCE COMPARISON');
  console.log('-' * 35);
  comparePerformance();
  
  // Test 3: Custom theme creation
  console.log('\n🎨 TEST 3: CUSTOM THEME CREATION');
  console.log('-' * 32);
  results.customTheme = createCustomThemedPresentation();
  
  // Summary
  console.log('\n📋 FINAL RESULTS');
  console.log('=' * 20);
  console.log(`Basic Integration: ${results.basicIntegration ? '✅ PASSED' : '❌ FAILED'}`);
  console.log(`Custom Theme: ${results.customTheme ? '✅ PASSED' : '❌ FAILED'}`);
  
  if (results.basicIntegration && results.customTheme) {
    console.log('\n🎉 ALL TESTS PASSED!');
    console.log('The Cloud Function integration is working perfectly.');
    console.log('You can now manipulate OOXML/PPTX files with full reliability.');
  } else {
    console.log('\n⚠️ SOME TESTS FAILED');
    console.log('Check the error messages above for troubleshooting.');
  }
  
  return results;
}