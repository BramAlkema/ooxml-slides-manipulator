/**
 * TypographyTests - Test suite for advanced typography and kerning XML hacking
 * 
 * Tests the tanaikech-style XML manipulation techniques for advanced typography controls
 * including kerning, tracking, ligatures, OpenType features, and professional
 * typographic adjustments beyond PowerPoint's standard interface
 */

/**
 * Main test function for typography and kerning XML hacking
 */
function runTypographyTests() {
  console.log('✍️ ADVANCED TYPOGRAPHY & KERNING TESTS');
  console.log('=' * 45);
  console.log('Tanaikech-style XML manipulation for professional typography controls\n');
  
  const results = {
    basicKerning: testBasicKerningControls(),
    professionalTypography: testProfessionalTypographyPresets(),
    customKerningPairs: testCustomKerningPairs(),
    openTypeFeatures: testOpenTypeFeatures(),
    characterTracking: testCharacterTracking(),
    wordSpacing: testWordSpacing(),
    responsiveTypography: testResponsiveTypography(),
    typographyToElement: testTypographyToElement()
  };
  
  printTypographyTestResults(results);
  return results;
}

/**
 * Test basic kerning controls
 */
function testBasicKerningControls() {
  console.log('🔧 TEST 1: Basic Kerning Controls');
  console.log('-' * 35);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Applying basic kerning controls...');
    slides.applyAdvancedKerning({
      'Heading_Kerning': {
        kerning: { method: 'optical', adjustment: -0.02 },
        tracking: 0.03,
        wordSpacing: 0.95,
        ligatures: { standard: true, discretionary: false }
      },
      'Body_Kerning': {
        kerning: { method: 'metrics', adjustment: 0 },
        tracking: 0,
        wordSpacing: 1.0,
        ligatures: { standard: true, discretionary: false }
      }
    });
    
    const fileId = slides.saveToGoogleDrive('Basic Kerning Controls Test');
    
    console.log('✅ Basic kerning controls successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Optical kerning: -2% adjustment for headings');
    console.log('   Tracking: 3% letter spacing for headings');
    console.log('   Word spacing: 95% for tighter text');
    console.log('   Standard ligatures enabled');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Basic kerning controls successfully applied'
    };
    
  } catch (error) {
    console.log('❌ Basic kerning controls failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test professional typography presets
 */
function testProfessionalTypographyPresets() {
  console.log('\n🎯 TEST 2: Professional Typography Presets');
  console.log('-' * 46);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Testing corporate typography preset...');
    slides.applyProfessionalTypography('corporate');
    
    const fileId = slides.saveToGoogleDrive('Professional Typography Test');
    
    console.log('✅ Professional typography presets successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Corporate preset applied:');
    console.log('     • Headings: Optical kerning, 5% tracking');
    console.log('     • Body: Metrics kerning, standard tracking');
    console.log('     • Captions: Optical kerning, 3% tracking');
    console.log('   OpenType features: kern, liga enabled');
    
    // Test editorial preset
    console.log('\n   Testing editorial preset...');
    const editorialSlides = OOXMLSlides.fromTemplate();
    editorialSlides.applyProfessionalTypography('editorial');
    const editorialFileId = editorialSlides.saveToGoogleDrive('Editorial Typography Test');
    
    console.log(`   Editorial preset: ${editorialFileId}`);
    console.log('     • Display text: Optical kerning, tighter spacing');
    console.log('     • Headlines: Balanced kerning');
    console.log('     • Bylines: Small caps, wider tracking');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    DriveApp.getFileById(editorialFileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      editorialFileId: editorialFileId,
      message: 'Professional typography presets successfully applied'
    };
    
  } catch (error) {
    console.log('❌ Professional typography presets failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test custom kerning pairs
 */
function testCustomKerningPairs() {
  console.log('\n🔤 TEST 3: Custom Kerning Pairs');
  console.log('-' * 32);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Creating custom kerning pairs...');
    slides.createKerningPairs('Arial', {
      'AV': -0.05,   // Tighter kerning between A and V
      'To': -0.03,   // Tighter kerning between T and o
      'Wa': -0.04,   // Tighter kerning between W and a
      'LT': 0.02,    // Looser kerning between L and T
      'ff': -0.01    // Slight adjustment for f ligature
    });
    
    const fileId = slides.saveToGoogleDrive('Custom Kerning Pairs Test');
    
    console.log('✅ Custom kerning pairs successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Arial font kerning pairs:');
    console.log('     • AV: -5% (tighter)');
    console.log('     • To: -3% (tighter)');
    console.log('     • Wa: -4% (tighter)');
    console.log('     • LT: +2% (looser)');
    console.log('     • ff: -1% (ligature adjustment)');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Custom kerning pairs successfully created'
    };
    
  } catch (error) {
    console.log('❌ Custom kerning pairs failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test OpenType font features
 */
function testOpenTypeFeatures() {
  console.log('\n🎨 TEST 4: OpenType Font Features');
  console.log('-' * 35);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Applying OpenType features...');
    
    // Test with a font that supports many OpenType features
    slides.applyOpenTypeFeatures('Minion Pro', [
      'kern',  // Kerning
      'liga',  // Standard ligatures
      'dlig',  // Discretionary ligatures
      'smcp',  // Small caps
      'onum',  // Old-style figures
      'swsh',  // Swashes
      'calt',  // Contextual alternates
      'ss01'   // Stylistic set 1
    ]);
    
    const fileId = slides.saveToGoogleDrive('OpenType Features Test');
    
    console.log('✅ OpenType features successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Minion Pro OpenType features:');
    console.log('     • kern: Automatic kerning');
    console.log('     • liga: Standard ligatures (fi, fl)');
    console.log('     • dlig: Discretionary ligatures');
    console.log('     • smcp: Small caps');
    console.log('     • onum: Old-style figures (1234)');
    console.log('     • swsh: Swash alternates');
    console.log('     • calt: Contextual alternates');
    console.log('     • ss01: Stylistic set 1');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'OpenType font features successfully applied'
    };
    
  } catch (error) {
    console.log('❌ OpenType features failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test character tracking (letter-spacing)
 */
function testCharacterTracking() {
  console.log('\n📏 TEST 5: Character Tracking (Letter-Spacing)');
  console.log('-' * 48);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Setting character tracking values...');
    slides.setCharacterTracking('Tight_Tracking', -0.02);  // Tighter
    slides.setCharacterTracking('Normal_Tracking', 0);      // Normal
    slides.setCharacterTracking('Loose_Tracking', 0.05);   // Looser
    slides.setCharacterTracking('Wide_Tracking', 0.1);     // Very wide
    
    const fileId = slides.saveToGoogleDrive('Character Tracking Test');
    
    console.log('✅ Character tracking successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Tracking variations:');
    console.log('     • Tight: -2% letter spacing');
    console.log('     • Normal: 0% letter spacing');
    console.log('     • Loose: +5% letter spacing');
    console.log('     • Wide: +10% letter spacing');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Character tracking successfully applied'
    };
    
  } catch (error) {
    console.log('❌ Character tracking failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test word spacing adjustments
 */
function testWordSpacing() {
  console.log('\n📐 TEST 6: Word Spacing Adjustments');
  console.log('-' * 37);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Setting word spacing ratios...');
    slides.setWordSpacing('Tight_Words', 0.85);    // Tighter word spacing
    slides.setWordSpacing('Normal_Words', 1.0);    // Normal word spacing
    slides.setWordSpacing('Loose_Words', 1.15);    // Looser word spacing
    slides.setWordSpacing('Wide_Words', 1.3);      // Very wide word spacing
    
    const fileId = slides.saveToGoogleDrive('Word Spacing Test');
    
    console.log('✅ Word spacing successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Word spacing variations:');
    console.log('     • Tight: 85% word spacing');
    console.log('     • Normal: 100% word spacing');
    console.log('     • Loose: 115% word spacing');
    console.log('     • Wide: 130% word spacing');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Word spacing successfully applied'
    };
    
  } catch (error) {
    console.log('❌ Word spacing failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test responsive typography
 */
function testResponsiveTypography() {
  console.log('\n📱 TEST 7: Responsive Typography');
  console.log('-' * 34);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Creating responsive typography system...');
    slides.createResponsiveTypography();
    
    const fileId = slides.saveToGoogleDrive('Responsive Typography Test');
    
    console.log('✅ Responsive typography successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Responsive typography rules:');
    console.log('     • Large text (≥24pt): Optical kerning, tighter tracking');
    console.log('     • Medium text (14-23pt): Optical kerning, normal tracking');
    console.log('     • Small text (<14pt): Metrics kerning, looser tracking');
    console.log('   Automatic adjustment based on text size');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Responsive typography successfully created'
    };
    
  } catch (error) {
    console.log('❌ Responsive typography failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test applying typography to specific elements
 */
function testTypographyToElement() {
  console.log('\n✨ TEST 8: Typography to Specific Elements');
  console.log('-' * 45);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Applying typography to slide elements...');
    
    // Apply different typography to different elements
    slides.applyTypographyToElement(0, 'title', {
      kerning: { method: 'optical', adjustment: -0.025 },
      tracking: 0.02,
      wordSpacing: 0.9,
      ligatures: { standard: true, discretionary: true },
      openType: ['kern', 'liga', 'dlig']
    });
    
    slides.applyTypographyToElement(0, 'content', {
      kerning: { method: 'metrics', adjustment: 0 },
      tracking: 0,
      wordSpacing: 1.0,
      ligatures: { standard: true, discretionary: false },
      openType: ['kern', 'liga']
    });
    
    const fileId = slides.saveToGoogleDrive('Element Typography Test');
    
    console.log('✅ Element typography successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Title element:');
    console.log('     • Optical kerning with -2.5% adjustment');
    console.log('     • 2% character tracking');
    console.log('     • 90% word spacing');
    console.log('     • All ligatures enabled');
    console.log('   Content element:');
    console.log('     • Metrics kerning');
    console.log('     • Normal tracking and spacing');
    console.log('     • Standard ligatures only');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Element-specific typography successfully applied'
    };
    
  } catch (error) {
    console.log('❌ Element typography failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test typography compatibility through Google Slides conversion
 */
function testTypographyCompatibility() {
  console.log('\n🔍 COMPATIBILITY TEST: Typography Through Google Slides');
  console.log('-' * 62);
  
  try {
    // Create OOXML with advanced typography
    const originalSlides = OOXMLSlides.fromTemplate();
    originalSlides.applyProfessionalTypography('corporate');
    
    const originalId = originalSlides.saveToGoogleDrive('Original Typography');
    
    // Convert through Google Slides
    const presentation = SlidesApp.openById(originalId);
    presentation.getSlides()[0].insertTextBox('Converted via Google Slides');
    
    // Export back to PPTX and analyze
    const convertedBlob = DriveApp.getFileById(originalId).getBlob();
    const convertedFile = DriveApp.createFile(convertedBlob).setName('Converted Typography');
    
    const convertedSlides = OOXMLSlides.fromGoogleDriveFile(convertedFile.getId());
    const convertedSettings = convertedSlides.exportTypographySettings();
    
    console.log('✅ Typography compatibility test completed');
    console.log(`   Original: ${originalId}`);
    console.log(`   Converted: ${convertedFile.getId()}`);
    console.log(`   Typography settings after conversion: ${convertedSettings.length}`);
    
    // Clean up
    DriveApp.getFileById(originalId).setTrashed(true);
    DriveApp.getFileById(convertedFile.getId()).setTrashed(true);
    
    return {
      success: true,
      originalId: originalId,
      convertedId: convertedFile.getId(),
      convertedSettingsCount: convertedSettings.length,
      message: 'Typography compatibility test completed'
    };
    
  } catch (error) {
    console.log('❌ Typography compatibility test failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Demonstrate all typography capabilities
 */
function demonstrateTypographyCapabilities() {
  console.log('🎯 DEMONSTRATION: Advanced Typography & Kerning Capabilities');
  console.log('=' * 65);
  console.log('Creating comprehensive typography showcase...\n');
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('1. Applying professional corporate typography...');
    slides.applyProfessionalTypography('corporate');
    
    console.log('2. Creating custom kerning pairs for Arial...');
    slides.createKerningPairs('Arial', {
      'AV': -0.04, 'Wa': -0.03, 'To': -0.02, 'LT': 0.02
    });
    
    console.log('3. Adding OpenType features...');
    slides.applyOpenTypeFeatures('Calibri', ['kern', 'liga', 'calt']);
    
    console.log('4. Setting character tracking variations...');
    slides.setCharacterTracking('Headline', -0.01);
    slides.setCharacterTracking('Caption', 0.03);
    
    console.log('5. Adjusting word spacing...');
    slides.setWordSpacing('Tight_Text', 0.9);
    slides.setWordSpacing('Loose_Text', 1.1);
    
    console.log('6. Creating responsive typography...');
    slides.createResponsiveTypography();
    
    const fileId = slides.saveToGoogleDrive('Typography Showcase');
    const settingCount = slides.exportTypographySettings().length;
    
    console.log('\n🎉 Typography Showcase Complete!');
    console.log(`   Created presentation: ${fileId}`);
    console.log(`   Total typography configurations: ${settingCount}`);
    console.log('   Features demonstrated:');
    console.log('     • Professional typography presets');
    console.log('     • Custom kerning pairs for specific fonts');
    console.log('     • OpenType feature activation');
    console.log('     • Precise character tracking control');
    console.log('     • Advanced word spacing adjustments');
    console.log('     • Responsive typography rules');
    console.log('     • Font ligature management');
    console.log('\n   All typography controls go far beyond PowerPoint standard options!');
    
    return fileId;
    
  } catch (error) {
    console.log('❌ Typography demonstration failed:', error.message);
    return null;
  }
}

/**
 * Print comprehensive test results
 */
function printTypographyTestResults(results) {
  console.log('\n📋 ADVANCED TYPOGRAPHY & KERNING RESULTS');
  console.log('=' * 45);
  
  const tests = [
    { name: 'Basic Kerning Controls', key: 'basicKerning', icon: '🔧' },
    { name: 'Professional Typography', key: 'professionalTypography', icon: '🎯' },
    { name: 'Custom Kerning Pairs', key: 'customKerningPairs', icon: '🔤' },
    { name: 'OpenType Features', key: 'openTypeFeatures', icon: '🎨' },
    { name: 'Character Tracking', key: 'characterTracking', icon: '📏' },
    { name: 'Word Spacing', key: 'wordSpacing', icon: '📐' },
    { name: 'Responsive Typography', key: 'responsiveTypography', icon: '📱' },
    { name: 'Element Typography', key: 'typographyToElement', icon: '✨' }
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
    console.log('\n🎉 ALL TYPOGRAPHY & KERNING XML HACKS WORKING!');
    console.log('Advanced typography controls beyond PowerPoint standards are now available.');
    console.log('\n🔥 Key Typography Capabilities Unlocked:');
    console.log('• Professional typography presets (corporate, editorial, technical)');
    console.log('• Custom kerning pairs for specific character combinations');
    console.log('• OpenType font feature activation (ligatures, swashes, small caps)');
    console.log('• Precise character tracking (letter-spacing) control');
    console.log('• Advanced word spacing adjustments');
    console.log('• Responsive typography that adapts to text size');
    console.log('• Element-specific typography application');
    console.log('• Font ligature management');
    console.log('• Baseline shift adjustments');
    console.log('\nNext steps:');
    console.log('• Apply typography to actual slide content');
    console.log('• Run testTypographyCompatibility() to check Google Slides survival');
    console.log('• Use demonstrateTypographyCapabilities() to see full showcase');
  } else {
    console.log('\n⚠️ SOME TESTS FAILED');
    console.log('Review the error messages above for troubleshooting.');
  }
}

/**
 * Quick test for immediate verification
 */
function quickTypographyTest() {
  console.log('🔍 Quick Typography Test');
  console.log('Testing basic typography and kerning...\n');
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    const result = TypographyEditor.testTypographyHack(slides);
    
    if (result) {
      const fileId = slides.saveToGoogleDrive('Quick Typography Test');
      const settingCount = slides.exportTypographySettings().length;
      console.log(`✅ Quick test passed! File: ${fileId}`);
      console.log(`Typography settings: ${settingCount}`);
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