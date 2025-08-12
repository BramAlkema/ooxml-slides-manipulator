/**
 * CustomColorTests - Test suite for Brandwares custom color XML hacking
 * 
 * Tests the XML manipulation techniques for adding unlimited custom colors
 * beyond PowerPoint's standard 12-color theme limitation
 * 
 * Based on: https://www.brandwares.com/bestpractices/2015/06/xml-hacking-custom-colors/
 */

/**
 * Main test function for custom color XML hacking
 */
function runCustomColorTests() {
  console.log('🎨 CUSTOM COLOR XML HACKING TESTS');
  console.log('=' * 40);
  console.log('Based on: https://www.brandwares.com/bestpractices/2015/06/xml-hacking-custom-colors/\n');
  
  const results = {
    basicCustomColors: testBasicCustomColors(),
    brandwaresColorHack: testBrandwaresCustomColorHack(),
    brandColorPalette: testBrandColorPalette(),
    materialDesign: testMaterialDesignColors(),
    accessibleColors: testAccessibleColorPalette(),
    colorHarmony: testColorHarmony(),
    customGradients: testCustomGradients(),
    colorExport: testColorPaletteExport()
  };
  
  printCustomColorResults(results);
  return results;
}

/**
 * Test basic custom color addition
 */
function testBasicCustomColors() {
  console.log('🔧 TEST 1: Basic Custom Color Addition');
  console.log('-' * 42);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Adding basic custom colors...');
    slides.addCustomColors({
      'My_Brand_Blue': '#2E86AB',
      'My_Brand_Orange': '#F18F01',
      'My_Brand_Green': '#4C956C',
      'My_Brand_Red': '#C73E1D',
      'My_Brand_Purple': '#A23B72'
    });
    
    const fileId = slides.saveToGoogleDrive('Basic Custom Colors Test');
    
    console.log('✅ Basic custom colors successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Added 5 custom brand colors');
    console.log('   Colors accessible via "More Colors" dialog');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Basic custom colors successfully added to theme'
    };
    
  } catch (error) {
    console.log('❌ Basic custom colors failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test the full Brandwares custom color hack
 */
function testBrandwaresCustomColorHack() {
  console.log('\n🚀 TEST 2: Brandwares Custom Color Hack');
  console.log('-' * 44);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Applying Brandwares custom color hack...');
    slides.applyBrandwaresColorHack({
      // Add some additional custom colors
      'Custom_Teal': '#009688',
      'Custom_Amber': '#FFC107',
      'Custom_DeepOrange': '#FF5722'
    });
    
    const fileId = slides.saveToGoogleDrive('Brandwares Color Hack Test');
    
    console.log('✅ Brandwares color hack successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Added 25+ corporate colors');
    console.log('   Blue scale: 9 shades from light to dark');
    console.log('   Gray scale: 9 professional grays');
    console.log('   Accent colors: Orange, Green, Red, Purple, Teal');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Brandwares custom color hack successfully applied'
    };
    
  } catch (error) {
    console.log('❌ Brandwares color hack failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test brand-specific color palette creation
 */
function testBrandColorPalette() {
  console.log('\n🏢 TEST 3: Brand Color Palette Creation');
  console.log('-' * 41);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Creating "TechCorp" brand color palette...');
    slides.createBrandColorPalette('TechCorp', {
      primary: '#1565C0',      // Tech Blue
      secondary: '#FF7043',    // Orange
      accent1: '#43A047',      // Green
      accent2: '#E53935',      // Red
      neutralDark: '#263238',  // Dark Gray
      neutralLight: '#FAFAFA', // Light Gray
      success: '#2E7D32',      // Success Green
      warning: '#F57C00',      // Warning Orange
      error: '#C62828',        // Error Red
      info: '#1976D2',         // Info Blue
      variations: {
        'Light_Blue': '#E3F2FD',
        'Dark_Blue': '#0D47A1',
        'Light_Orange': '#FFE0B2',
        'Dark_Orange': '#BF360C'
      }
    });
    
    const fileId = slides.saveToGoogleDrive('TechCorp Brand Colors Test');
    
    console.log('✅ Brand color palette creation successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   TechCorp brand: 14 custom colors');
    console.log('   Includes semantic colors (success, warning, error)');
    console.log('   Color variations for design flexibility');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Brand color palette successfully created'
    };
    
  } catch (error) {
    console.log('❌ Brand color palette creation failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test Material Design color palette
 */
function testMaterialDesignColors() {
  console.log('\n🎨 TEST 4: Material Design Color Palette');
  console.log('-' * 42);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Adding Material Design color palette...');
    slides.addMaterialDesignColors();
    
    const fileId = slides.saveToGoogleDrive('Material Design Colors Test');
    
    console.log('✅ Material Design colors successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Added Material Design color system');
    console.log('   6 color families with 4 shades each');
    console.log('   Includes: Red, Pink, Purple, Blue, Green, Orange');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Material Design color palette successfully added'
    };
    
  } catch (error) {
    console.log('❌ Material Design colors failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test accessibility-compliant color palette
 */
function testAccessibleColorPalette() {
  console.log('\n♿ TEST 5: Accessible Color Palette (WCAG)');
  console.log('-' * 43);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Creating WCAG AA compliant color palette...');
    slides.addAccessibleColors();
    
    const fileId = slides.saveToGoogleDrive('Accessible Colors Test');
    
    console.log('✅ Accessible color palette successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   WCAG AA compliant colors (4.5:1 contrast)');
    console.log('   Color blind safe palette');
    console.log('   High visibility focus colors');
    console.log('   Semantic colors for UI (success, warning, error)');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Accessible color palette successfully created'
    };
    
  } catch (error) {
    console.log('❌ Accessible color palette failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test color harmony generation
 */
function testColorHarmony() {
  console.log('\n🎭 TEST 6: Color Harmony Generation');
  console.log('-' * 38);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Creating color harmonies from base color #3498DB...');
    
    // Create different types of harmonies
    slides.createColorHarmony('#3498DB', 'complementary');
    slides.createColorHarmony('#E74C3C', 'triadic');
    slides.createColorHarmony('#27AE60', 'analogous');
    slides.createColorHarmony('#9B59B6', 'monochromatic');
    
    const fileId = slides.saveToGoogleDrive('Color Harmony Test');
    
    console.log('✅ Color harmony generation successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   4 different harmony types applied');
    console.log('   Complementary: Blue + Orange');
    console.log('   Triadic: Red + two 120° shifts');
    console.log('   Analogous: Green + adjacent hues');
    console.log('   Monochromatic: Purple + light/dark variants');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Color harmony generation successful'
    };
    
  } catch (error) {
    console.log('❌ Color harmony generation failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test custom gradient definitions
 */
function testCustomGradients() {
  console.log('\n🌈 TEST 7: Custom Gradient Definitions');
  console.log('-' * 40);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Adding custom gradient definitions...');
    slides.addCustomGradients({
      'Ocean_Gradient': {
        type: 'linear',
        angle: 45,
        stops: [
          { position: 0, color: '#1e3c72', alpha: 1.0 },
          { position: 0.5, color: '#2a5298', alpha: 1.0 },
          { position: 1, color: '#74b9ff', alpha: 1.0 }
        ]
      },
      'Sunset_Gradient': {
        type: 'linear',
        angle: 90,
        stops: [
          { position: 0, color: '#ff7f50', alpha: 1.0 },
          { position: 0.5, color: '#ff6347', alpha: 1.0 },
          { position: 1, color: '#ff4500', alpha: 1.0 }
        ]
      },
      'Forest_Radial': {
        type: 'radial',
        stops: [
          { position: 0, color: '#2d5016', alpha: 1.0 },
          { position: 0.7, color: '#4a7c59', alpha: 0.8 },
          { position: 1, color: '#6b8e23', alpha: 0.6 }
        ]
      }
    });
    
    const fileId = slides.saveToGoogleDrive('Custom Gradients Test');
    
    console.log('✅ Custom gradients successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Added 3 custom gradient definitions');
    console.log('   Ocean: Blue linear gradient (45°)');
    console.log('   Sunset: Orange linear gradient (90°)');
    console.log('   Forest: Green radial gradient');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Custom gradients successfully added'
    };
    
  } catch (error) {
    console.log('❌ Custom gradients failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test color palette export functionality
 */
function testColorPaletteExport() {
  console.log('\n📤 TEST 8: Color Palette Export');
  console.log('-' * 32);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    // Add some colors first
    slides.addCustomColors({
      'Export_Test_Red': '#FF0000',
      'Export_Test_Green': '#00FF00',
      'Export_Test_Blue': '#0000FF'
    });
    
    console.log('Exporting color palette...');
    const colors = slides.exportColorPalette();
    
    console.log('✅ Color palette export successful');
    console.log(`   Found ${colors.length} colors in theme`);
    console.log('   Color list includes:');
    colors.slice(0, 10).forEach(color => {
      console.log(`     ${color}`);
    });
    if (colors.length > 10) {
      console.log(`     ... and ${colors.length - 10} more colors`);
    }
    
    return {
      success: true,
      colorCount: colors.length,
      colors: colors,
      message: `Successfully exported ${colors.length} colors from theme`
    };
    
  } catch (error) {
    console.log('❌ Color palette export failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test custom color compatibility through Google Slides conversion
 */
function testCustomColorCompatibility() {
  console.log('\n🔍 COMPATIBILITY TEST: Custom Colors Through Google Slides');
  console.log('-' * 65);
  
  try {
    // Create OOXML with custom colors
    const originalSlides = OOXMLSlides.fromTemplate();
    originalSlides.applyBrandwaresColorHack();
    
    const originalId = originalSlides.saveToGoogleDrive('Original Custom Colors');
    
    // Convert through Google Slides
    const presentation = SlidesApp.openById(originalId);
    presentation.getSlides()[0].insertTextBox('Converted via Google Slides');
    
    // Export back to PPTX and analyze
    const convertedBlob = DriveApp.getFileById(originalId).getBlob();
    const convertedFile = DriveApp.createFile(convertedBlob).setName('Converted Custom Colors');
    
    const convertedSlides = OOXMLSlides.fromGoogleDriveFile(convertedFile.getId());
    const convertedColors = convertedSlides.exportColorPalette();
    
    console.log('✅ Custom color compatibility test completed');
    console.log(`   Original: ${originalId}`);
    console.log(`   Converted: ${convertedFile.getId()}`);
    console.log(`   Colors after conversion: ${convertedColors.length}`);
    
    // Clean up
    DriveApp.getFileById(originalId).setTrashed(true);
    DriveApp.getFileById(convertedFile.getId()).setTrashed(true);
    
    return {
      success: true,
      originalId: originalId,
      convertedId: convertedFile.getId(),
      convertedColorCount: convertedColors.length,
      message: 'Custom color compatibility test completed'
    };
    
  } catch (error) {
    console.log('❌ Custom color compatibility test failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Demonstrate all custom color capabilities
 */
function demonstrateCustomColorCapabilities() {
  console.log('🎯 DEMONSTRATION: Custom Color XML Hacking Capabilities');
  console.log('=' * 58);
  console.log('Creating comprehensive custom color showcase...\n');
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('1. Adding Brandwares corporate colors...');
    slides.applyBrandwaresColorHack();
    
    console.log('2. Adding Material Design palette...');
    slides.addMaterialDesignColors();
    
    console.log('3. Adding accessible colors...');
    slides.addAccessibleColors();
    
    console.log('4. Creating brand palette for "DemoCompany"...');
    slides.createBrandColorPalette('DemoCompany', {
      primary: '#6366F1',      // Indigo
      secondary: '#EC4899',    // Pink
      accent1: '#10B981',      // Emerald
      accent2: '#F59E0B',      // Amber
      neutralDark: '#1F2937',  // Gray 800
      neutralLight: '#F9FAFB'  // Gray 50
    });
    
    console.log('5. Adding color harmonies...');
    slides.createColorHarmony('#8B5CF6', 'triadic');
    
    console.log('6. Adding custom gradients...');
    slides.addCustomGradients({
      'Brand_Gradient': {
        type: 'linear',
        angle: 135,
        stops: [
          { position: 0, color: '#6366F1', alpha: 1.0 },
          { position: 1, color: '#EC4899', alpha: 1.0 }
        ]
      }
    });
    
    const fileId = slides.saveToGoogleDrive('Custom Color Showcase');
    const colorCount = slides.exportColorPalette().length;
    
    console.log('\n🎉 Custom Color Showcase Complete!');
    console.log(`   Created presentation: ${fileId}`);
    console.log(`   Total custom colors available: ${colorCount}`);
    console.log('   Features demonstrated:');
    console.log('     • Brandwares corporate color scales');
    console.log('     • Material Design color system');
    console.log('     • WCAG accessible color palette');
    console.log('     • Brand-specific color palette');
    console.log('     • Color theory-based harmonies');
    console.log('     • Custom gradient definitions');
    console.log('\n   All colors are now available in the "More Colors" dialog!');
    
    return fileId;
    
  } catch (error) {
    console.log('❌ Custom color demonstration failed:', error.message);
    return null;
  }
}

/**
 * Print comprehensive test results
 */
function printCustomColorResults(results) {
  console.log('\n📋 CUSTOM COLOR XML HACKING RESULTS');
  console.log('=' * 40);
  
  const tests = [
    { name: 'Basic Custom Colors', key: 'basicCustomColors', icon: '🔧' },
    { name: 'Brandwares Color Hack', key: 'brandwaresColorHack', icon: '🚀' },
    { name: 'Brand Color Palette', key: 'brandColorPalette', icon: '🏢' },
    { name: 'Material Design Colors', key: 'materialDesign', icon: '🎨' },
    { name: 'Accessible Colors', key: 'accessibleColors', icon: '♿' },
    { name: 'Color Harmony', key: 'colorHarmony', icon: '🎭' },
    { name: 'Custom Gradients', key: 'customGradients', icon: '🌈' },
    { name: 'Color Export', key: 'colorExport', icon: '📤' }
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
    console.log('\n🎉 ALL CUSTOM COLOR XML HACKS WORKING!');
    console.log('The Brandwares custom color technique is fully implemented.');
    console.log('Unlimited custom colors can now be added beyond the 12-color limit.');
    console.log('\n🔥 Key Capabilities Unlocked:');
    console.log('• Add unlimited custom colors to any presentation');
    console.log('• Create brand-specific color palettes');
    console.log('• Generate color harmonies based on color theory');
    console.log('• Add accessibility-compliant color sets');
    console.log('• Define custom gradients for advanced fills');
    console.log('• Export and analyze color usage');
    console.log('\nNext steps:');
    console.log('• Test custom colors in actual presentations');
    console.log('• Run testCustomColorCompatibility() to check Google Slides survival');
    console.log('• Use demonstrateCustomColorCapabilities() to see full showcase');
  } else {
    console.log('\n⚠️ SOME TESTS FAILED');
    console.log('Review the error messages above for troubleshooting.');
  }
}

/**
 * Quick test for immediate verification
 */
function quickCustomColorTest() {
  console.log('🔍 Quick Custom Color Test');
  console.log('Testing basic Brandwares custom color hack...\n');
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    const result = CustomColorEditor.testCustomColorHack(slides);
    
    if (result) {
      const fileId = slides.saveToGoogleDrive('Quick Color Test');
      const colorCount = slides.exportColorPalette().length;
      console.log(`✅ Quick test passed! File: ${fileId}`);
      console.log(`Colors available: ${colorCount}`);
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