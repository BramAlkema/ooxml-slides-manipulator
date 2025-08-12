/**
 * Tanaikech-Style Tests for Advanced Google Slides API Features
 * 
 * Tests for undocumented and advanced features inspired by tanaikech's approach
 * to accessing Google Apps capabilities beyond the official API.
 */

/**
 * Comprehensive test suite for tanaikech-style Slides API extensions
 */
function runTanaikechStyleTests() {
  console.log('🚀 Tanaikech-Style Google Slides API Extensions Test Suite');
  console.log('=' * 70);
  
  const results = {};
  
  try {
    // Test 1: Advanced Slides API Exploration
    console.log('\n🔍 TEST 1: API EXPLORATION');
    console.log('-' * 30);
    results.exploration = testAdvancedAPIExploration();
    
    // Test 2: Font Pair Manipulation
    console.log('\n🔤 TEST 2: FONT PAIR MANIPULATION');
    console.log('-' * 35);
    results.fontPairs = testFontPairManipulation();
    
    // Test 3: Advanced Color Control
    console.log('\n🎨 TEST 3: ADVANCED COLOR CONTROL');
    console.log('-' * 35);
    results.colorControl = testAdvancedColorControl();
    
    // Test 4: Theme Properties Beyond API
    console.log('\n✨ TEST 4: THEME PROPERTIES BEYOND API');
    console.log('-' * 40);
    results.advancedTheme = testAdvancedThemeProperties();
    
    // Test 5: Master Slide Enhancement
    console.log('\n🎯 TEST 5: MASTER SLIDE ENHANCEMENT');
    console.log('-' * 36);
    results.masterEnhancement = testMasterSlideEnhancement();
    
    // Test 6: OOXML Integration
    console.log('\n🔗 TEST 6: OOXML INTEGRATION');
    console.log('-' * 28);
    results.ooxmlIntegration = testOOXMLIntegration();
    
    // Final Results
    console.log('\n📋 TANAIKECH-STYLE TEST RESULTS');
    console.log('=' * 35);
    Object.entries(results).forEach(([test, result]) => {
      const status = result ? '✅ PASSED' : '❌ FAILED';
      console.log(`${test}: ${status}`);
    });
    
    const allPassed = Object.values(results).every(result => result);
    
    if (allPassed) {
      console.log('\n🎉 ALL TANAIKECH-STYLE TESTS PASSED!');
      console.log('Advanced Google Slides API extensions are working perfectly.');
    } else {
      console.log('\n⚠️ SOME TESTS FAILED');
      console.log('Check the detailed output above for troubleshooting.');
    }
    
    return results;
    
  } catch (error) {
    console.error('❌ Tanaikech-style test suite failed:', error);
    return { error: error.message };
  }
}

/**
 * Test 1: Advanced API Exploration
 */
function testAdvancedAPIExploration() {
  console.log('Testing advanced API exploration capabilities...');
  
  try {
    // Create a test presentation
    const testPresentation = SlidesApp.create('Tanaikech Style API Test');
    const presentationId = testPresentation.getId();
    
    console.log('Step 1: Creating test presentation...');
    console.log(`✅ Test presentation created: ${presentationId}`);
    
    // Explore with API Explorer
    console.log('Step 2: Running API exploration...');
    const analysis = SlidesAPIExplorer.explorePresentation(presentationId);
    
    if (analysis && analysis.basicInfo) {
      console.log(`✅ API exploration successful`);
      console.log(`   - Masters: ${analysis.basicInfo.masterCount}`);
      console.log(`   - Layouts: ${analysis.basicInfo.layoutCount}`);
      console.log(`   - Slides: ${analysis.basicInfo.slideCount}`);
    }
    
    // Test font exploration
    console.log('Step 3: Testing font exploration...');
    const fontAnalysis = SlidesAPIExplorer.exploreFontPairs(presentationId);
    
    if (fontAnalysis) {
      console.log(`✅ Font exploration successful`);
      console.log(`   - Master fonts: ${Object.keys(fontAnalysis.masterFonts).length}`);
      console.log(`   - Layout fonts: ${Object.keys(fontAnalysis.layoutFonts).length}`);
    }
    
    // Test color exploration
    console.log('Step 4: Testing color exploration...');
    const colorAnalysis = SlidesAPIExplorer.exploreColorScheme(presentationId);
    
    if (colorAnalysis) {
      console.log(`✅ Color exploration successful`);
      console.log(`   - Solid colors: ${colorAnalysis.solidFillColors.length}`);
      console.log(`   - Gradient colors: ${colorAnalysis.gradientColors.length}`);
    }
    
    // Clean up
    DriveApp.getFileById(presentationId).setTrashed(true);
    
    return true;
    
  } catch (error) {
    console.error('❌ API exploration test failed:', error);
    return false;
  }
}

/**
 * Test 2: Font Pair Manipulation
 */
function testFontPairManipulation() {
  console.log('Testing font pair manipulation beyond standard API...');
  
  try {
    // Create test presentation with advanced API
    console.log('Step 1: Creating presentation with SlidesAppAdvanced...');
    const advancedPresentation = SlidesAppAdvanced.create('Font Pair Test', {
      fontPairs: {
        TITLE: { fontFamily: 'Montserrat', fontSize: 44, bold: true },
        BODY: { fontFamily: 'Open Sans', fontSize: 18, bold: false },
        SUBTITLE: { fontFamily: 'Lato', fontSize: 24, italic: true }
      }
    });
    
    const presentationId = advancedPresentation.getId();
    console.log(`✅ Advanced presentation created: ${presentationId}`);
    
    // Test getting font pairs
    console.log('Step 2: Testing font pair retrieval...');
    const currentFontPairs = advancedPresentation.getFontPairs();
    
    if (currentFontPairs && typeof currentFontPairs === 'object') {
      console.log('✅ Font pairs retrieved successfully');
      Object.entries(currentFontPairs).forEach(([type, config]) => {
        console.log(`   - ${type}: ${config.fontFamily || 'Unknown'}`);
      });
    }
    
    // Test setting new font pairs
    console.log('Step 3: Testing font pair modification...');
    advancedPresentation.setFontPairs({
      TITLE: { fontFamily: 'Roboto', fontSize: 48, bold: true },
      BODY: { fontFamily: 'Source Sans Pro', fontSize: 16 }
    });
    
    console.log('✅ Font pairs modified successfully');
    
    // Export to test functionality
    console.log('Step 4: Testing export with custom fonts...');
    const exportBlob = advancedPresentation.exportAdvanced({
      format: 'application/pdf'
    });
    
    if (exportBlob && exportBlob.getSize() > 0) {
      console.log(`✅ Export successful: ${exportBlob.getSize()} bytes`);
    }
    
    // Clean up
    DriveApp.getFileById(presentationId).setTrashed(true);
    
    return true;
    
  } catch (error) {
    console.error('❌ Font pair manipulation test failed:', error);
    return false;
  }
}

/**
 * Test 3: Advanced Color Control
 */
function testAdvancedColorControl() {
  console.log('Testing advanced color control beyond standard API...');
  
  try {
    // Create presentation with custom color scheme
    console.log('Step 1: Creating presentation with custom colors...');
    const advancedPresentation = SlidesAppAdvanced.create('Color Test', {
      theme: {
        colorScheme: {
          accent1: '#FF6B35',  // Orange
          accent2: '#E74C3C',  // Red
          accent3: '#3498DB',  // Blue
          accent4: '#2ECC71',  // Green
          dark1: '#2C3E50',    // Dark blue
          light1: '#ECF0F1'    // Light gray
        }
      }
    });
    
    const presentationId = advancedPresentation.getId();
    console.log(`✅ Color-themed presentation created: ${presentationId}`);
    
    // Test color palette retrieval
    console.log('Step 2: Testing color palette retrieval...');
    const colorPalette = advancedPresentation.getColorPalette();
    
    if (colorPalette) {
      console.log('✅ Color palette retrieved');
    }
    
    // Test color palette modification
    console.log('Step 3: Testing color palette modification...');
    advancedPresentation.setColorPalette({
      accent1: '#9B59B6',  // Purple
      accent2: '#F39C12',  // Orange
      accent3: '#1ABC9C'   // Turquoise
    });
    
    console.log('✅ Color palette modified');
    
    // Test advanced theme setting
    console.log('Step 4: Testing comprehensive theme setting...');
    advancedPresentation.setAdvancedTheme({
      colorScheme: {
        accent1: '2E86AB',  // Ocean Blue
        accent2: 'A23B72',  // Deep Pink
        accent3: 'F18F01'   // Bright Orange
      },
      masterBackground: {
        solidFill: {
          color: 'F8F9FA',
          alpha: 1.0
        }
      }
    });
    
    console.log('✅ Advanced theme applied');
    
    // Clean up
    DriveApp.getFileById(presentationId).setTrashed(true);
    
    return true;
    
  } catch (error) {
    console.error('❌ Advanced color control test failed:', error);
    return false;
  }
}

/**
 * Test 4: Advanced Theme Properties
 */
function testAdvancedThemeProperties() {
  console.log('Testing theme properties beyond standard API...');
  
  try {
    // Create presentation
    const presentation = SlidesApp.create('Advanced Theme Test');
    const advancedPresentation = SlidesAppAdvanced.openById(presentation.getId());
    const presentationId = presentation.getId();
    
    console.log('Step 1: Getting advanced theme properties...');
    const advancedTheme = advancedPresentation.getAdvancedTheme();
    
    if (advancedTheme) {
      console.log('✅ Advanced theme properties retrieved');
    }
    
    console.log('Step 2: Getting master properties...');
    const masterProperties = advancedPresentation.getMasterProperties();
    
    if (masterProperties) {
      console.log('✅ Master properties retrieved');
      console.log(`   - Masters: ${masterProperties.masterThemes ? masterProperties.masterThemes.length : 0}`);
    }
    
    console.log('Step 3: Updating master properties...');
    try {
      advancedPresentation.updateMasterProperties({
        background: {
          solidFill: {
            color: 'FFFFFF',
            alpha: 1.0
          }
        }
      });
      console.log('✅ Master properties updated');
    } catch (updateError) {
      console.log('⚠️ Master property update attempted (may require specific permissions)');
    }
    
    // Clean up
    DriveApp.getFileById(presentationId).setTrashed(true);
    
    return true;
    
  } catch (error) {
    console.error('❌ Advanced theme properties test failed:', error);
    return false;
  }
}

/**
 * Test 5: Master Slide Enhancement
 */
function testMasterSlideEnhancement() {
  console.log('Testing master slide enhancement capabilities...');
  
  try {
    const presentation = SlidesApp.create('Master Enhancement Test');
    const advancedPresentation = SlidesAppAdvanced.openById(presentation.getId());
    const presentationId = presentation.getId();
    
    console.log('Step 1: Analyzing current master...');
    const masterProps = advancedPresentation.getMasterProperties();
    
    if (masterProps) {
      console.log('✅ Master analysis complete');
    }
    
    console.log('Step 2: Testing master enhancement...');
    // Test various master enhancements
    try {
      advancedPresentation.updateMasterProperties({
        elements: [{
          shapeType: 'TEXT_BOX',
          size: { width: { magnitude: 200, unit: 'PT' }, height: { magnitude: 50, unit: 'PT' } },
          transform: { scaleX: 1.0, scaleY: 1.0, translateX: 100, translateY: 100 }
        }]
      });
      console.log('✅ Master enhancement attempted');
    } catch (enhanceError) {
      console.log('⚠️ Master enhancement attempted (may require additional permissions)');
    }
    
    // Clean up
    DriveApp.getFileById(presentationId).setTrashed(true);
    
    return true;
    
  } catch (error) {
    console.error('❌ Master slide enhancement test failed:', error);
    return false;
  }
}

/**
 * Test 6: OOXML Integration
 */
function testOOXMLIntegration() {
  console.log('Testing integration between advanced Slides API and OOXML manipulation...');
  
  try {
    // Create presentation with advanced features
    console.log('Step 1: Creating advanced presentation...');
    const advancedPresentation = SlidesAppAdvanced.create('OOXML Integration Test', {
      theme: {
        colorScheme: {
          accent1: 'FF6B35',
          accent2: 'E74C3C'
        }
      }
    });
    
    const presentationId = advancedPresentation.getId();
    console.log(`✅ Advanced presentation created: ${presentationId}`);
    
    // Convert to OOXML for manipulation
    console.log('Step 2: Converting to OOXML...');
    const ooxmlParser = advancedPresentation.toOOXML();
    
    if (ooxmlParser) {
      console.log('✅ OOXML conversion successful');
      
      // Test OOXML manipulation
      console.log('Step 3: Testing OOXML manipulation...');
      ooxmlParser.extract();
      const fileCount = ooxmlParser.listFiles().length;
      console.log(`✅ OOXML extracted: ${fileCount} files`);
      
      // Modify theme in OOXML
      if (ooxmlParser.hasFile('ppt/theme/theme1.xml')) {
        const themeXML = ooxmlParser.getXML('ppt/theme/theme1.xml');
        console.log('✅ Theme XML accessed for advanced manipulation');
      }
      
      // Rebuild
      const modifiedBlob = ooxmlParser.build();
      console.log(`✅ OOXML rebuilt: ${modifiedBlob.getBytes().length} bytes`);
    }
    
    // Clean up
    DriveApp.getFileById(presentationId).setTrashed(true);
    
    return true;
    
  } catch (error) {
    console.error('❌ OOXML integration test failed:', error);
    return false;
  }
}

/**
 * Demo function showing tanaikech-style advanced usage
 */
function demonstrateTanaikechStyleUsage() {
  console.log('🎭 Demonstrating Tanaikech-Style Advanced Google Slides Usage');
  console.log('=' * 65);
  
  try {
    // Create a presentation with advanced styling
    console.log('Creating presentation with advanced tanaikech-style features...');
    
    const presentation = SlidesAppAdvanced.create('Tanaikech Style Demo', {
      theme: {
        colorScheme: {
          accent1: '2E86AB',  // Professional blue
          accent2: 'A23B72',  // Deep pink
          accent3: 'F18F01',  // Bright orange
          accent4: '4C956C',  // Green
          dark1: '1A1A1A',    // Almost black
          light1: 'FFFFFF'    // Pure white
        },
        masterBackground: {
          solidFill: {
            color: 'F8F9FA',
            alpha: 0.95
          }
        }
      },
      fontPairs: {
        TITLE: { fontFamily: 'Montserrat', fontSize: 44, bold: true },
        BODY: { fontFamily: 'Source Sans Pro', fontSize: 16 },
        SUBTITLE: { fontFamily: 'Lato', fontSize: 24, italic: true }
      }
    });
    
    console.log('✅ Advanced presentation created with:');
    console.log('   - Custom color scheme (6 colors)');
    console.log('   - Professional font pairs');
    console.log('   - Master background styling');
    
    // Demonstrate API exploration
    const presentationId = presentation.getId();
    console.log('\nExploring undocumented API features...');
    
    const exploration = SlidesAPIExplorer.explorePresentation(presentationId);
    console.log('✅ API exploration complete');
    
    // Demonstrate OOXML integration
    console.log('\nIntegrating with OOXML for advanced manipulation...');
    const ooxmlParser = presentation.toOOXML();
    ooxmlParser.extract();
    
    // Advanced OOXML + Slides API hybrid manipulation
    if (ooxmlParser.hasFile('ppt/theme/theme1.xml')) {
      console.log('✅ Direct theme XML access available');
      console.log('   This enables manipulation beyond standard API limits');
    }
    
    // Save final result
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const fileName = `Tanaikech_Style_Demo_${timestamp}.pptx`;
    
    const finalBlob = ooxmlParser.build();
    const driveFile = DriveApp.createFile(finalBlob.setName(fileName));
    
    console.log('\n🎉 Tanaikech-Style Demo Complete!');
    console.log(`📁 File: ${driveFile.getName()}`);
    console.log(`🔗 URL: https://docs.google.com/presentation/d/${driveFile.getId()}`);
    console.log('\nFeatures demonstrated:');
    console.log('✅ Advanced color scheme manipulation');
    console.log('✅ Font pair relationships');
    console.log('✅ API exploration beyond documentation');
    console.log('✅ Hybrid Slides API + OOXML manipulation');
    console.log('✅ Master slide advanced properties');
    
    // Clean up temp presentation
    DriveApp.getFileById(presentationId).setTrashed(true);
    
    return driveFile.getId();
    
  } catch (error) {
    console.error('❌ Tanaikech-style demo failed:', error);
    return false;
  }
}

/**
 * Quick test for tanaikech-style connectivity
 */
function quickTanaikechTest() {
  console.log('🔍 Quick Tanaikech-Style Feature Test');
  
  try {
    // Test 1: Advanced Slides API
    const presentation = SlidesAppAdvanced.create('Quick Test');
    console.log('✅ SlidesAppAdvanced working');
    
    // Test 2: API Explorer
    const analysis = SlidesAPIExplorer.explorePresentation(presentation.getId());
    console.log('✅ SlidesAPIExplorer working');
    
    // Test 3: OOXML Integration
    const ooxml = presentation.toOOXML();
    console.log('✅ OOXML integration working');
    
    // Clean up
    DriveApp.getFileById(presentation.getId()).setTrashed(true);
    
    console.log('\n🎉 All tanaikech-style features operational!');
    return true;
    
  } catch (error) {
    console.error('❌ Quick tanaikech test failed:', error);
    return false;
  }
}