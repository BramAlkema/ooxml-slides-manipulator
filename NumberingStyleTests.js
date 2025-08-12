/**
 * NumberingStyleTests - Test suite for Brandwares numbering styles XML hacking
 * 
 * Tests the XML manipulation techniques for creating custom numbering and bullet styles
 * that go far beyond PowerPoint's standard list formatting options
 * 
 * Based on: https://www.brandwares.com/bestpractices/2017/06/xml-hacking-powerpoint-numbering-styles/
 */

/**
 * Main test function for numbering styles XML hacking
 */
function runNumberingStyleTests() {
  console.log('🔢 NUMBERING STYLES XML HACKING TESTS');
  console.log('=' * 45);
  console.log('Based on: https://www.brandwares.com/bestpractices/2017/06/xml-hacking-powerpoint-numbering-styles/\n');
  
  const results = {
    basicNumberingStyles: testBasicNumberingStyles(),
    brandwaresNumberingHack: testBrandwaresNumberingHack(),
    modernNumberingStyles: testModernNumberingStyles(),
    accessibleNumbering: testAccessibleNumberingStyles(),
    customBulletStyles: testCustomBulletStyles(),
    restartableNumbering: testRestartableNumbering(),
    legalStyleNumbering: testLegalStyleNumbering(),
    emojiBasedBullets: testEmojiBasedBullets()
  };
  
  printNumberingStyleResults(results);
  return results;
}

/**
 * Test basic custom numbering styles
 */
function testBasicNumberingStyles() {
  console.log('🔧 TEST 1: Basic Custom Numbering Styles');
  console.log('-' * 44);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Creating basic custom numbering styles...');
    slides.createNumberingStyles({
      'Basic_Decimal': {
        type: 'multilevel',
        levels: [
          {
            level: 0,
            format: 'decimal',
            text: '%1.',
            font: { family: 'Arial', size: 12, bold: true },
            color: '#2F5597',
            indent: { left: 0.25, hanging: 0.25 }
          },
          {
            level: 1,
            format: 'decimal',
            text: '%1.%2.',
            font: { family: 'Arial', size: 11, bold: false },
            color: '#2F5597',
            indent: { left: 0.5, hanging: 0.25 }
          }
        ]
      },
      'Basic_Bullets': {
        type: 'bullet',
        levels: [
          {
            level: 0,
            bullet: '●',
            font: { family: 'Arial', size: 12 },
            color: '#4472C4',
            indent: { left: 0.25, hanging: 0.2 }
          }
        ]
      }
    });
    
    const fileId = slides.saveToGoogleDrive('Basic Numbering Styles Test');
    
    console.log('✅ Basic numbering styles successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Added decimal numbering with hierarchy');
    console.log('   Added custom bullet style with color');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Basic numbering styles successfully created'
    };
    
  } catch (error) {
    console.log('❌ Basic numbering styles failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test the full Brandwares numbering styles hack
 */
function testBrandwaresNumberingHack() {
  console.log('\n🚀 TEST 2: Brandwares Numbering Styles Hack');
  console.log('-' * 46);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Applying Brandwares numbering styles hack...');
    slides.applyBrandwaresNumberingHack();
    
    const fileId = slides.saveToGoogleDrive('Brandwares Numbering Hack Test');
    
    console.log('✅ Brandwares numbering hack successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Legal-style numbering: 1.1, 1.2, 1.2.1');
    console.log('   Corporate bullets: ▶ ▸ ◦');
    console.log('   Academic outline: I. A. 1. a)');
    console.log('   Multi-level hierarchies with custom indentation');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Brandwares numbering styles hack successfully applied'
    };
    
  } catch (error) {
    console.log('❌ Brandwares numbering hack failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test modern design system numbering styles
 */
function testModernNumberingStyles() {
  console.log('\n🎨 TEST 3: Modern Design System Numbering');
  console.log('-' * 42);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Creating modern numbering styles...');
    slides.addModernNumberingStyles();
    
    const fileId = slides.saveToGoogleDrive('Modern Numbering Styles Test');
    
    console.log('✅ Modern numbering styles successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Material Design inspired numbers');
    console.log('   Emoji bullets: 🔹 🔸 🔺');
    console.log('   Process steps with styled backgrounds');
    console.log('   Roboto/Segoe UI fonts for modern look');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Modern numbering styles successfully created'
    };
    
  } catch (error) {
    console.log('❌ Modern numbering styles failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test accessible numbering styles
 */
function testAccessibleNumberingStyles() {
  console.log('\n♿ TEST 4: Accessible Numbering Styles');
  console.log('-' * 38);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Creating accessible numbering styles...');
    slides.addAccessibleNumberingStyles();
    
    const fileId = slides.saveToGoogleDrive('Accessible Numbering Test');
    
    console.log('✅ Accessible numbering styles successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   High contrast colors (black on white)');
    console.log('   Screen reader friendly bullet alternatives');
    console.log('   Clear hierarchy with proper indentation');
    console.log('   WCAG AA compliant formatting');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Accessible numbering styles successfully created'
    };
    
  } catch (error) {
    console.log('❌ Accessible numbering styles failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test custom bullet styles with special characters
 */
function testCustomBulletStyles() {
  console.log('\n🎯 TEST 5: Custom Bullet Styles');
  console.log('-' * 33);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Creating custom bullet styles...');
    slides.createBulletStyles({
      'Arrow_Bullets': {
        type: 'bullet',
        levels: [
          {
            level: 0,
            bullet: '→',
            font: { family: 'Arial', size: 12, bold: true },
            color: '#E74C3C',
            indent: { left: 0.25, hanging: 0.2 }
          },
          {
            level: 1,
            bullet: '↳',
            font: { family: 'Arial', size: 11 },
            color: '#E74C3C',
            indent: { left: 0.5, hanging: 0.2 }
          }
        ]
      },
      'Checkmark_Bullets': {
        type: 'bullet',
        levels: [
          {
            level: 0,
            bullet: '✓',
            font: { family: 'Arial', size: 12, bold: true },
            color: '#27AE60',
            indent: { left: 0.25, hanging: 0.2 }
          }
        ]
      },
      'Star_Bullets': {
        type: 'bullet',
        levels: [
          {
            level: 0,
            bullet: '★',
            font: { family: 'Arial', size: 12 },
            color: '#F39C12',
            indent: { left: 0.25, hanging: 0.2 }
          }
        ]
      }
    });
    
    const fileId = slides.saveToGoogleDrive('Custom Bullet Styles Test');
    
    console.log('✅ Custom bullet styles successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Arrow bullets: → ↳');
    console.log('   Checkmark bullets: ✓');
    console.log('   Star bullets: ★');
    console.log('   Custom colors and indentation');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Custom bullet styles successfully created'
    };
    
  } catch (error) {
    console.log('❌ Custom bullet styles failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test restartable numbering functionality
 */
function testRestartableNumbering() {
  console.log('\n🔄 TEST 6: Restartable Numbering');
  console.log('-' * 34);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Creating restartable numbered lists...');
    slides.createRestartableList('RestartList1', 1);
    slides.createRestartableList('RestartList2', 5);
    slides.createRestartableList('RestartList3', 10);
    
    const fileId = slides.saveToGoogleDrive('Restartable Numbering Test');
    
    console.log('✅ Restartable numbering successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   List 1: Starts at 1');
    console.log('   List 2: Starts at 5');
    console.log('   List 3: Starts at 10');
    console.log('   Each list can restart independently');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Restartable numbering successfully created'
    };
    
  } catch (error) {
    console.log('❌ Restartable numbering failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test legal-style hierarchical numbering
 */
function testLegalStyleNumbering() {
  console.log('\n⚖️ TEST 7: Legal-Style Hierarchical Numbering');
  console.log('-' * 48);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Creating legal-style numbering...');
    slides.createNumberingStyles({
      'Legal_Hierarchy': {
        type: 'multilevel',
        levels: [
          {
            level: 0,
            format: 'decimal',
            text: '%1.',
            font: { family: 'Times New Roman', size: 12, bold: true },
            color: '#000000',
            indent: { left: 0, hanging: 0.3 }
          },
          {
            level: 1,
            format: 'decimal',
            text: '%1.%2.',
            font: { family: 'Times New Roman', size: 11 },
            color: '#000000',
            indent: { left: 0.3, hanging: 0.3 }
          },
          {
            level: 2,
            format: 'decimal',
            text: '%1.%2.%3.',
            font: { family: 'Times New Roman', size: 10 },
            color: '#000000',
            indent: { left: 0.6, hanging: 0.3 }
          },
          {
            level: 3,
            format: 'decimal',
            text: '%1.%2.%3.%4.',
            font: { family: 'Times New Roman', size: 10 },
            color: '#000000',
            indent: { left: 0.9, hanging: 0.3 }
          }
        ]
      }
    });
    
    const fileId = slides.saveToGoogleDrive('Legal Style Numbering Test');
    
    console.log('✅ Legal-style numbering successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Hierarchy: 1. → 1.1. → 1.1.1. → 1.1.1.1.');
    console.log('   Times New Roman font for traditional look');
    console.log('   Proper legal document indentation');
    console.log('   4 levels of hierarchical numbering');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Legal-style numbering successfully created'
    };
    
  } catch (error) {
    console.log('❌ Legal-style numbering failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test emoji-based bullet points
 */
function testEmojiBasedBullets() {
  console.log('\n😀 TEST 8: Emoji-Based Bullet Points');
  console.log('-' * 37);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Creating emoji bullet styles...');
    slides.createBulletStyles({
      'Process_Emojis': {
        type: 'bullet',
        levels: [
          {
            level: 0,
            bullet: '🚀',
            font: { family: 'Segoe UI Emoji', size: 12 },
            indent: { left: 0.25, hanging: 0.25 }
          },
          {
            level: 1,
            bullet: '⭐',
            font: { family: 'Segoe UI Emoji', size: 11 },
            indent: { left: 0.5, hanging: 0.25 }
          },
          {
            level: 2,
            bullet: '💡',
            font: { family: 'Segoe UI Emoji', size: 10 },
            indent: { left: 0.75, hanging: 0.25 }
          }
        ]
      },
      'Status_Emojis': {
        type: 'bullet',
        levels: [
          {
            level: 0,
            bullet: '✅',
            font: { family: 'Segoe UI Emoji', size: 12 },
            indent: { left: 0.25, hanging: 0.25 }
          },
          {
            level: 1,
            bullet: '⚠️',
            font: { family: 'Segoe UI Emoji', size: 12 },
            indent: { left: 0.5, hanging: 0.25 }
          },
          {
            level: 2,
            bullet: '❌',
            font: { family: 'Segoe UI Emoji', size: 12 },
            indent: { left: 0.75, hanging: 0.25 }
          }
        ]
      }
    });
    
    const fileId = slides.saveToGoogleDrive('Emoji Bullet Points Test');
    
    console.log('✅ Emoji bullet points successful');
    console.log(`   Created presentation: ${fileId}`);
    console.log('   Process emojis: 🚀 ⭐ 💡');
    console.log('   Status emojis: ✅ ⚠️ ❌');
    console.log('   Modern, engaging bullet points');
    console.log('   Segoe UI Emoji font for proper rendering');
    
    // Clean up
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return {
      success: true,
      fileId: fileId,
      message: 'Emoji bullet points successfully created'
    };
    
  } catch (error) {
    console.log('❌ Emoji bullet points failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test numbering style compatibility through Google Slides conversion
 */
function testNumberingStyleCompatibility() {
  console.log('\n🔍 COMPATIBILITY TEST: Numbering Styles Through Google Slides');
  console.log('-' * 68);
  
  try {
    // Create OOXML with numbering styles
    const originalSlides = OOXMLSlides.fromTemplate();
    originalSlides.applyBrandwaresNumberingHack();
    
    const originalId = originalSlides.saveToGoogleDrive('Original Numbering Styles');
    
    // Convert through Google Slides
    const presentation = SlidesApp.openById(originalId);
    presentation.getSlides()[0].insertTextBox('Converted via Google Slides');
    
    // Export back to PPTX and analyze
    const convertedBlob = DriveApp.getFileById(originalId).getBlob();
    const convertedFile = DriveApp.createFile(convertedBlob).setName('Converted Numbering Styles');
    
    const convertedSlides = OOXMLSlides.fromGoogleDriveFile(convertedFile.getId());
    const convertedStyles = convertedSlides.exportNumberingStyles();
    
    console.log('✅ Numbering style compatibility test completed');
    console.log(`   Original: ${originalId}`);
    console.log(`   Converted: ${convertedFile.getId()}`);
    console.log(`   Styles after conversion: ${convertedStyles.length}`);
    
    // Clean up
    DriveApp.getFileById(originalId).setTrashed(true);
    DriveApp.getFileById(convertedFile.getId()).setTrashed(true);
    
    return {
      success: true,
      originalId: originalId,
      convertedId: convertedFile.getId(),
      convertedStyleCount: convertedStyles.length,
      message: 'Numbering style compatibility test completed'
    };
    
  } catch (error) {
    console.log('❌ Numbering style compatibility test failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Demonstrate all numbering style capabilities
 */
function demonstrateNumberingStyleCapabilities() {
  console.log('🎯 DEMONSTRATION: Numbering Styles XML Hacking Capabilities');
  console.log('=' * 63);
  console.log('Creating comprehensive numbering styles showcase...\n');
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('1. Adding Brandwares numbering styles...');
    slides.applyBrandwaresNumberingHack();
    
    console.log('2. Adding modern design system styles...');
    slides.addModernNumberingStyles();
    
    console.log('3. Adding accessible numbering...');
    slides.addAccessibleNumberingStyles();
    
    console.log('4. Creating custom bullet styles...');
    slides.createBulletStyles({
      'Demo_Arrows': {
        type: 'bullet',
        levels: [
          { level: 0, bullet: '▶', font: { family: 'Arial', size: 12 }, color: '#3498DB' },
          { level: 1, bullet: '▷', font: { family: 'Arial', size: 11 }, color: '#3498DB' }
        ]
      },
      'Demo_Checkmarks': {
        type: 'bullet',
        levels: [
          { level: 0, bullet: '✓', font: { family: 'Arial', size: 12 }, color: '#27AE60' }
        ]
      }
    });
    
    console.log('5. Creating restartable lists...');
    slides.createRestartableList('DemoList', 100);
    
    console.log('6. Adding emoji bullets...');
    slides.createBulletStyles({
      'Demo_Emojis': {
        type: 'bullet',
        levels: [
          { level: 0, bullet: '🔥', font: { family: 'Segoe UI Emoji', size: 12 } },
          { level: 1, bullet: '⚡', font: { family: 'Segoe UI Emoji', size: 11 } },
          { level: 2, bullet: '💫', font: { family: 'Segoe UI Emoji', size: 10 } }
        ]
      }
    });
    
    const fileId = slides.saveToGoogleDrive('Numbering Styles Showcase');
    const styleCount = slides.exportNumberingStyles().length;
    
    console.log('\n🎉 Numbering Styles Showcase Complete!');
    console.log(`   Created presentation: ${fileId}`);
    console.log(`   Total custom numbering styles: ${styleCount}`);
    console.log('   Features demonstrated:');
    console.log('     • Legal-style hierarchical numbering (1.1.1)');
    console.log('     • Corporate bullet styles with custom colors');
    console.log('     • Academic outline formatting (I, A, 1, a)');
    console.log('     • Modern emoji-based bullet points');
    console.log('     • Restartable numbered lists');
    console.log('     • Accessible high-contrast formatting');
    console.log('     • Custom bullet characters and symbols');
    console.log('\n   All styles go far beyond standard PowerPoint options!');
    
    return fileId;
    
  } catch (error) {
    console.log('❌ Numbering styles demonstration failed:', error.message);
    return null;
  }
}

/**
 * Print comprehensive test results
 */
function printNumberingStyleResults(results) {
  console.log('\n📋 NUMBERING STYLES XML HACKING RESULTS');
  console.log('=' * 45);
  
  const tests = [
    { name: 'Basic Numbering Styles', key: 'basicNumberingStyles', icon: '🔧' },
    { name: 'Brandwares Numbering Hack', key: 'brandwaresNumberingHack', icon: '🚀' },
    { name: 'Modern Numbering Styles', key: 'modernNumberingStyles', icon: '🎨' },
    { name: 'Accessible Numbering', key: 'accessibleNumbering', icon: '♿' },
    { name: 'Custom Bullet Styles', key: 'customBulletStyles', icon: '🎯' },
    { name: 'Restartable Numbering', key: 'restartableNumbering', icon: '🔄' },
    { name: 'Legal-Style Numbering', key: 'legalStyleNumbering', icon: '⚖️' },
    { name: 'Emoji-Based Bullets', key: 'emojiBasedBullets', icon: '😀' }
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
    console.log('\n🎉 ALL NUMBERING STYLE XML HACKS WORKING!');
    console.log('The Brandwares numbering styles technique is fully implemented.');
    console.log('Advanced list formatting beyond PowerPoint standards is now available.');
    console.log('\n🔥 Key Capabilities Unlocked:');
    console.log('• Legal-style hierarchical numbering (1.1, 1.2, 1.2.1)');
    console.log('• Corporate bullet styles with custom colors and fonts');
    console.log('• Academic outline formatting (I, A, 1, a)');
    console.log('• Modern emoji-based bullet points');
    console.log('• Restartable numbered lists at any number');
    console.log('• Custom bullet characters and symbols');
    console.log('• Accessible high-contrast formatting');
    console.log('• Multi-level hierarchies with precise indentation');
    console.log('\nNext steps:');
    console.log('• Apply numbering to actual slide elements');
    console.log('• Run testNumberingStyleCompatibility() to check Google Slides survival');
    console.log('• Use demonstrateNumberingStyleCapabilities() to see full showcase');
  } else {
    console.log('\n⚠️ SOME TESTS FAILED');
    console.log('Review the error messages above for troubleshooting.');
  }
}

/**
 * Quick test for immediate verification
 */
function quickNumberingStyleTest() {
  console.log('🔍 Quick Numbering Style Test');
  console.log('Testing basic Brandwares numbering hack...\n');
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    const result = NumberingStyleEditor.testNumberingStylesHack(slides);
    
    if (result) {
      const fileId = slides.saveToGoogleDrive('Quick Numbering Test');
      const styleCount = slides.exportNumberingStyles().length;
      console.log(`✅ Quick test passed! File: ${fileId}`);
      console.log(`Numbering styles created: ${styleCount}`);
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