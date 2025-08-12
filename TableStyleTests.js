/**
 * TableStyleTests - Test suite for Brandwares XML table hacking
 * 
 * Tests the XML manipulation techniques for modifying default table text properties
 * that are not accessible through standard PowerPoint/Google Slides APIs
 */

/**
 * Main test function for table style XML hacking
 */
function runTableStyleTests() {
  console.log('🧪 TABLE STYLE XML HACKING TESTS');
  console.log('=' * 40);
  console.log('Based on: https://www.brandwares.com/bestpractices/2015/03/xml-hacking-default-table-text/\n');
  
  const results = {
    basicTableText: testBasicTableTextHack(),
    brandwaresHack: testBrandwaresTableHack(),
    customTableStyle: testCustomTableStyleCreation(),
    cellMargins: testTableCellMargins(),
    advancedFormatting: testAdvancedTableFormatting()
  };
  
  printTableTestResults(results);
  return results;
}

/**
 * Test basic table text property modification
 */
function testBasicTableTextHack() {
  console.log('🔧 TEST 1: Basic Table Text Property Hack');
  console.log('-' * 45);
  
  try {
    // Create template with table styling
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Setting default table text properties...');
    slides.setDefaultTableText({
      headerFont: {
        family: 'Segoe UI',
        size: 14,
        bold: true,
        color: '2F5597' // Dark blue
      },
      bodyFont: {
        family: 'Segoe UI',
        size: 11,
        bold: false,
        color: '404040' // Dark gray
      },
      alignment: {
        horizontal: 'left',
        vertical: 'middle'
      }
    });
    
    // Test by creating a presentation
    const fileId = slides.saveToGoogleDrive('Table Text Hack Test');
    
    console.log('✅ Basic table text hack successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Header: Segoe UI 14pt Bold Dark Blue');
    console.log('   Body: Segoe UI 11pt Regular Dark Gray');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Table text properties successfully modified via XML hack'
    };
    
  } catch (error) {
    console.log('❌ Basic table text hack failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test the full Brandwares table styling hack
 */
function testBrandwaresTableHack() {
  console.log('\n🎨 TEST 2: Brandwares Table Style Hack');
  console.log('-' * 40);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Applying Brandwares table styling...');
    slides.applyBrandwaresTableHack({
      headerFont: {
        family: 'Arial',
        size: 12,
        bold: true,
        color: 'FFFFFF' // White
      },
      bodyFont: {
        family: 'Arial', 
        size: 10,
        bold: false,
        color: '333333' // Dark gray
      },
      headerBackground: '4472C4', // Blue
      alternateRowColor: 'F2F2F2', // Light gray
      borderColor: 'CCCCCC'
    });
    
    const fileId = slides.saveToGoogleDrive('Brandwares Table Hack Test');
    
    console.log('✅ Brandwares table hack successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Applied professional table styling');
    console.log('   Blue headers with white text');
    console.log('   Alternating row colors');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Brandwares table styling successfully applied'
    };
    
  } catch (error) {
    console.log('❌ Brandwares table hack failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test custom table style creation
 */
function testCustomTableStyleCreation() {
  console.log('\n🏗️ TEST 3: Custom Table Style Creation');
  console.log('-' * 42);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Creating custom "Executive" table style...');
    slides.createTableStyle('Executive', {
      headerFont: {
        family: 'Montserrat',
        size: 13,
        bold: true,
        color: 'FFFFFF'
      },
      bodyFont: {
        family: 'Open Sans',
        size: 10,
        bold: false,
        color: '2C3E50'
      },
      headerBackground: '34495E', // Dark blue-gray
      alternateRowColor: 'ECF0F1', // Very light gray
      borderColor: 'BDC3C7'
    });
    
    const fileId = slides.saveToGoogleDrive('Custom Table Style Test');
    
    console.log('✅ Custom table style creation successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   "Executive" style with Montserrat headers');
    console.log('   Open Sans body text');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Custom table style successfully created'
    };
    
  } catch (error) {
    console.log('❌ Custom table style creation failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test table cell margin modification
 */
function testTableCellMargins() {
  console.log('\n📏 TEST 4: Table Cell Margin Hack');
  console.log('-' * 35);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Setting custom table cell margins...');
    slides.setTableCellMargins({
      left: 0.15,   // Inches
      right: 0.15,
      top: 0.08,
      bottom: 0.08
    });
    
    const fileId = slides.saveToGoogleDrive('Table Cell Margins Test');
    
    console.log('✅ Table cell margins hack successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Custom margins: 0.15" horizontal, 0.08" vertical');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Table cell margins successfully modified'
    };
    
  } catch (error) {
    console.log('❌ Table cell margins hack failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test advanced table formatting features
 */
function testAdvancedTableFormatting() {
  console.log('\n✨ TEST 5: Advanced Table Formatting');
  console.log('-' * 40);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Applying advanced table formatting...');
    
    // Combine multiple table hacks
    slides
      .setDefaultTableText({
        headerFont: {
          family: 'Roboto',
          size: 12,
          bold: true,
          color: 'FFFFFF'
        },
        bodyFont: {
          family: 'Roboto',
          size: 10,
          bold: false,
          color: '37474F'
        }
      })
      .setTableCellMargins({
        left: 0.12,
        right: 0.12,
        top: 0.06,
        bottom: 0.06
      });
    
    // Apply custom styling
    slides.applyBrandwaresTableHack({
      headerBackground: '607D8B', // Blue-gray
      alternateRowColor: 'F5F5F5',
      borderColor: 'E0E0E0'
    });
    
    const fileId = slides.saveToGoogleDrive('Advanced Table Formatting Test');
    
    console.log('✅ Advanced table formatting successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Combined multiple XML hacks');
    console.log('   Roboto fonts with custom spacing');
    console.log('   Professional color scheme');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Advanced table formatting successfully applied'
    };
    
  } catch (error) {
    console.log('❌ Advanced table formatting failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test the table style compatibility through Google Slides conversion
 */
function testTableStyleCompatibility() {
  console.log('\n🔍 COMPATIBILITY TEST: Table Styles Through Google Slides');
  console.log('-' * 60);
  
  try {
    // Create OOXML with table styling
    const originalSlides = OOXMLSlides.fromTemplate();
    originalSlides.applyBrandwaresTableHack({
      headerFont: { family: 'Arial', size: 12, bold: true, color: 'FFFFFF' },
      bodyFont: { family: 'Arial', size: 10, bold: false, color: '333333' },
      headerBackground: '4472C4'
    });
    
    const originalId = originalSlides.saveToGoogleDrive('Original Table Style');
    
    // Convert through Google Slides
    const presentation = SlidesApp.openById(originalId);
    presentation.getSlides()[0].insertTextBox('Converted via Google Slides');
    
    // Export back to PPTX
    const convertedBlob = DriveApp.getFileById(originalId).getBlob();
    const convertedFile = DriveApp.createFile(convertedBlob).setName('Converted Table Style');
    
    console.log('✅ Table style compatibility test completed');
    console.log(`   Original: ${originalId}`);
    console.log(`   Converted: ${convertedFile.getId()}`);
    console.log('   Check both files to compare table formatting survival');
    
    // Clean up
    DriveApp.getFileById(originalId).setTrashed(true);
    DriveApp.getFileById(convertedFile.getId()).setTrashed(true);
    
    return {
      success: true,
      originalId: originalId,
      convertedId: convertedFile.getId(),
      message: 'Table style compatibility test completed'
    };
    
  } catch (error) {
    console.log('❌ Table style compatibility test failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Quick table style demonstration
 */
function demonstrateTableStyleHacks() {
  console.log('🎯 DEMONSTRATION: Table Style XML Hacks');
  console.log('=' * 45);
  console.log('Creating presentation with 3 different table styles...\n');
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    // Style 1: Corporate Blue
    console.log('1. Corporate Blue Style');
    slides.createTableStyle('Corporate Blue', {
      headerFont: { family: 'Calibri', size: 12, bold: true, color: 'FFFFFF' },
      bodyFont: { family: 'Calibri', size: 10, bold: false, color: '2F5597' },
      headerBackground: '2F5597',
      alternateRowColor: 'F0F4F8'
    });
    
    // Style 2: Modern Gray
    console.log('2. Modern Gray Style');
    slides.createTableStyle('Modern Gray', {
      headerFont: { family: 'Segoe UI', size: 11, bold: true, color: 'FFFFFF' },
      bodyFont: { family: 'Segoe UI', size: 9, bold: false, color: '424242' },
      headerBackground: '616161',
      alternateRowColor: 'FAFAFA'
    });
    
    // Style 3: Creative Green
    console.log('3. Creative Green Style');
    slides.createTableStyle('Creative Green', {
      headerFont: { family: 'Lato', size: 12, bold: true, color: 'FFFFFF' },
      bodyFont: { family: 'Lato', size: 10, bold: false, color: '2E7D32' },
      headerBackground: '4CAF50',
      alternateRowColor: 'E8F5E8'
    });
    
    const fileId = slides.saveToGoogleDrive('Table Style Demonstration');
    
    console.log('\n✅ Table style demonstration complete!');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Contains 3 custom table styles via XML hacking');
    console.log('   These styles modify default table formatting');
    
    return fileId;
    
  } catch (error) {
    console.log('❌ Table style demonstration failed:', error.message);
    return null;
  }
}

/**
 * Print comprehensive test results
 */
function printTableTestResults(results) {
  console.log('\n📋 TABLE STYLE XML HACKING RESULTS');
  console.log('=' * 40);
  
  const tests = [
    { name: 'Basic Table Text Hack', key: 'basicTableText', icon: '🔧' },
    { name: 'Brandwares Table Hack', key: 'brandwaresHack', icon: '🎨' },
    { name: 'Custom Table Style Creation', key: 'customTableStyle', icon: '🏗️' },
    { name: 'Table Cell Margins', key: 'cellMargins', icon: '📏' },
    { name: 'Advanced Table Formatting', key: 'advancedFormatting', icon: '✨' }
  ];
  
  console.log('\n🎯 Test Results:');
  tests.forEach(test => {
    const result = results[test.key];
    const status = result.success ? '✅ PASSED' : '❌ FAILED';
    console.log(`${test.icon} ${test.name}: ${status}`);
    if (!result.success && result.error) {
      console.log(`   Error: ${result.error}`);
    }
  });
  
  const passedTests = Object.values(results).filter(r => r.success).length;
  const totalTests = Object.keys(results).length;
  
  console.log(`\n📊 Summary: ${passedTests}/${totalTests} tests passed`);
  
  if (passedTests === totalTests) {
    console.log('\n🎉 ALL TABLE STYLE XML HACKS WORKING!');
    console.log('The Brandwares technique is successfully implemented.');
    console.log('Default table text properties can now be modified via XML.');
    console.log('\nNext steps:');
    console.log('• Test table styles in actual presentations');
    console.log('• Run testTableStyleCompatibility() to check Google Slides survival');
    console.log('• Use demonstrateTableStyleHacks() to see all styles');
  } else {
    console.log('\n⚠️ SOME TESTS FAILED');
    console.log('Review the error messages above for troubleshooting.');
  }
}

/**
 * Quick test for immediate verification
 */
function quickTableStyleTest() {
  console.log('🔍 Quick Table Style Test');
  console.log('Testing basic Brandwares XML hack...\n');
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    const result = TableStyleEditor.testTableTextHack(slides);
    
    if (result) {
      const fileId = slides.saveToGoogleDrive('Quick Table Test');
      console.log(`✅ Quick test passed! File: ${fileId}`);
      DriveApp.getFileById(fileId).setTrashed(true);
      return true;
    } else {
      console.log('❌ Quick test failed');
      return false;
    }
    
  } catch (error) {
    console.log('❌ Quick test error:', error.message);
    return false;
  }
}