/**
 * Test integration between Google Apps Script and Cloud Function
 * Demonstrates PPTX manipulation using the CloudPPTXService
 */

function testCloudIntegration() {
  console.log('🧪 Testing Cloud Function integration...\n');
  
  try {
    // Test 1: Basic connectivity
    console.log('Test 1: Cloud Function connectivity...');
    const isAvailable = CloudPPTXService.testCloudFunction();
    if (!isAvailable) {
      console.log('❌ Cloud Function not available. Make sure it\'s running on localhost:8080');
      return;
    }
    console.log('✅ Cloud Function connectivity test passed!\n');
    
    // Test 2: Create PPTX template and manipulate it
    console.log('Test 2: Template creation and manipulation...');
    
    // Create a template using our existing system
    const templateBlob = PPTXTemplate.createMinimalTemplate();
    console.log(`Created template: ${templateBlob.getBytes().length} bytes`);
    
    // Use OOXMLParser with Cloud Function integration
    const parser = new OOXMLParser(templateBlob);
    
    // Extract using Cloud Function
    parser.extract();
    console.log(`Extracted ${parser.listFiles().length} files:`);
    console.log(parser.listFiles().join(', '));
    
    // Test 3: Modify theme colors
    console.log('\nTest 3: Theme color manipulation...');
    
    if (parser.hasFile('ppt/theme/theme1.xml')) {
      const themeXML = parser.getXML('ppt/theme/theme1.xml');
      console.log('Found theme XML, modifying colors...');
      
      // Find and update accent colors
      const root = themeXML.getRootElement();
      const ns = XmlService.getNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
      const clrScheme = root.getChild('themeElements', ns).getChild('clrScheme', ns);
      
      // Change accent1 to bright red
      const accent1 = clrScheme.getChild('accent1', ns);
      if (accent1) {
        const srgbClr = accent1.getChild('srgbClr', ns);
        if (srgbClr) {
          srgbClr.setAttribute('val', 'FF0000'); // Bright red
          console.log('✅ Changed accent1 to bright red');
        }
      }
      
      // Save the modified XML back
      const modifiedXML = XmlService.getRawFormat().format(themeXML);
      parser.setFile('ppt/theme/theme1.xml', modifiedXML);
    }
    
    // Test 4: Build modified PPTX using Cloud Function
    console.log('\nTest 4: Building modified PPTX...');
    const modifiedBlob = parser.build();
    console.log(`✅ Built modified PPTX: ${modifiedBlob.getBytes().length} bytes`);
    
    // Test 5: Save to Google Drive for verification
    console.log('\nTest 5: Saving to Google Drive...');
    const file = DriveApp.createFile(modifiedBlob.setName('Modified_PPTX_Test.pptx'));
    console.log(`✅ Saved to Google Drive: ${file.getName()}`);
    console.log(`   File ID: ${file.getId()}`);
    console.log(`   You can open it at: https://docs.google.com/presentation/d/${file.getId()}`);
    
    console.log('\n🎉 All cloud integration tests passed!');
    console.log('The PPTX manipulation system is working with Cloud Function support.');
    
  } catch (error) {
    console.error('❌ Cloud integration test failed:', error);
    console.log('\nTroubleshooting tips:');
    console.log('1. Make sure the Cloud Function is running on localhost:8080');
    console.log('2. Check that the Cloud Function code is deployed correctly');
    console.log('3. Verify network connectivity between Apps Script and the Cloud Function');
  }
}

/**
 * Test advanced PPTX manipulation scenarios
 */
function testAdvancedManipulation() {
  console.log('🔬 Testing advanced PPTX manipulation...\n');
  
  try {
    // Test creating a PPTX with custom theme
    console.log('Creating PPTX with custom theme...');
    
    const customTemplate = OOXMLSlides.fromTemplate()
      .setColorTheme({
        accent1: 'FF6B35', // Orange
        accent2: '004E89', // Blue  
        accent3: '1A5E63', // Teal
        dark1: '2C3E50',   // Dark gray
        light1: 'FFFFFF'   // White
      })
      .setFonts({
        majorFont: 'Roboto',
        minorFont: 'Open Sans'
      });
    
    const customBlob = customTemplate.build();
    console.log(`✅ Created custom themed PPTX: ${customBlob.getBytes().length} bytes`);
    
    // Save the custom template
    const customFile = DriveApp.createFile(customBlob.setName('Custom_Theme_PPTX.pptx'));
    console.log(`✅ Saved custom theme PPTX: https://docs.google.com/presentation/d/${customFile.getId()}`);
    
  } catch (error) {
    console.error('❌ Advanced manipulation test failed:', error);
  }
}

/**
 * Compare performance: Cloud Function vs Native Methods
 */
function testPerformanceComparison() {
  console.log('⚡ Performance comparison: Cloud Function vs Native Methods\n');
  
  try {
    const templateBlob = PPTXTemplate.createMinimalTemplate();
    
    // Test Cloud Function approach
    console.log('Testing Cloud Function approach...');
    const startCloud = new Date().getTime();
    
    try {
      const cloudFiles = CloudPPTXService.unzipPPTX(templateBlob);
      const cloudRebuild = CloudPPTXService.zipPPTX(cloudFiles);
      const endCloud = new Date().getTime();
      
      console.log(`✅ Cloud Function: ${endCloud - startCloud}ms`);
      console.log(`   Extracted ${Object.keys(cloudFiles).length} files`);
      console.log(`   Final size: ${cloudRebuild.getBytes().length} bytes`);
      
    } catch (cloudError) {
      console.log(`❌ Cloud Function failed: ${cloudError.message}`);
    }
    
    // Test native approach
    console.log('\nTesting native approach...');
    const startNative = new Date().getTime();
    
    try {
      const nativeFiles = Utilities.unzip(templateBlob);
      const nativeRebuild = Utilities.zip(nativeFiles);
      const endNative = new Date().getTime();
      
      console.log(`✅ Native method: ${endNative - startNative}ms`);
      console.log(`   Extracted ${nativeFiles.length} files`);
      console.log(`   Final size: ${nativeRebuild.getBytes().length} bytes`);
      
    } catch (nativeError) {
      console.log(`❌ Native method failed: ${nativeError.message}`);
      console.log('This is expected for real PPTX files');
    }
    
  } catch (error) {
    console.error('❌ Performance test failed:', error);
  }
}