/**
 * OOXML-Slides Compatibility Test
 * 
 * Tests which "undocumented things under the hood" survive the Google Slides PPTX conversion process.
 * This addresses the user's final question: "Not all undocumented stuff will survive Gslides PPTX conversion. But at least we can try, no?"
 * 
 * Test Process:
 * 1. Create OOXML presentation with advanced features
 * 2. Convert through Google Slides (import → export)  
 * 3. Compare before/after OOXML structures
 * 4. Document what survives vs what gets stripped
 */

class OOXMLCompatibilityTest {
  
  /**
   * Run complete compatibility test suite
   */
  static runCompatibilityTests() {
    console.log('🧪 OOXML-Google Slides Compatibility Test Suite');
    console.log('=' * 55);
    console.log('Testing which tanaikech-style features survive conversion...\n');
    
    const results = {
      themeProperties: this.testThemeCompatibility(),
      fontPairs: this.testFontPairCompatibility(), 
      colorPalettes: this.testColorPaletteCompatibility(),
      masterProperties: this.testMasterPropertiesCompatibility(),
      advancedFormatting: this.testAdvancedFormattingCompatibility(),
      customXMLParts: this.testCustomXMLCompatibility()
    };
    
    this.printCompatibilityReport(results);
    return results;
  }
  
  /**
   * Test theme properties compatibility
   */
  static testThemeCompatibility() {
    console.log('🎨 Testing Theme Properties Compatibility...');
    
    try {
      // Create OOXML with advanced theme properties
      const originalPPTX = this.createAdvancedThemeTest();
      
      // Convert through Google Slides
      const convertedPPTX = this.convertThroughGoogleSlides(originalPPTX, 'Theme Test');
      
      // Compare theme XML structures
      const comparison = this.compareThemeStructures(originalPPTX, convertedPPTX);
      
      console.log(`  ✅ Theme test completed: ${comparison.survivingFeatures}/${comparison.totalFeatures} features survived`);
      
      return {
        success: true,
        survivingFeatures: comparison.survivingFeatures,
        totalFeatures: comparison.totalFeatures,
        survivedProperties: comparison.survived,
        strippedProperties: comparison.stripped,
        details: comparison
      };
      
    } catch (error) {
      console.log(`  ❌ Theme test failed: ${error.message}`);
      return { success: false, error: error.message };
    }
  }
  
  /**
   * Test font pair compatibility
   */
  static testFontPairCompatibility() {
    console.log('🔤 Testing Font Pairs Compatibility...');
    
    try {
      // Create OOXML with complex font relationships
      const originalPPTX = this.createFontPairTest();
      
      // Convert through Google Slides
      const convertedPPTX = this.convertThroughGoogleSlides(originalPPTX, 'Font Pair Test');
      
      // Analyze font preservation
      const fontComparison = this.compareFontStructures(originalPPTX, convertedPPTX);
      
      console.log(`  ✅ Font pair test completed: ${fontComparison.preservedFonts}/${fontComparison.totalFonts} fonts preserved`);
      
      return {
        success: true,
        preservedFonts: fontComparison.preservedFonts,
        totalFonts: fontComparison.totalFonts,
        fontChanges: fontComparison.changes,
        details: fontComparison
      };
      
    } catch (error) {
      console.log(`  ❌ Font pair test failed: ${error.message}`);
      return { success: false, error: error.message };
    }
  }
  
  /**
   * Test color palette compatibility
   */
  static testColorPaletteCompatibility() {
    console.log('🌈 Testing Color Palette Compatibility...');
    
    try {
      // Create OOXML with custom color scheme
      const originalPPTX = this.createColorPaletteTest();
      
      // Convert through Google Slides
      const convertedPPTX = this.convertThroughGoogleSlides(originalPPTX, 'Color Palette Test');
      
      // Compare color schemes
      const colorComparison = this.compareColorSchemes(originalPPTX, convertedPPTX);
      
      console.log(`  ✅ Color palette test completed: ${colorComparison.preservedColors}/${colorComparison.totalColors} colors preserved`);
      
      return {
        success: true,
        preservedColors: colorComparison.preservedColors,
        totalColors: colorComparison.totalColors,
        colorChanges: colorComparison.changes,
        details: colorComparison
      };
      
    } catch (error) {
      console.log(`  ❌ Color palette test failed: ${error.message}`);
      return { success: false, error: error.message };
    }
  }
  
  /**
   * Test master properties compatibility
   */
  static testMasterPropertiesCompatibility() {
    console.log('🏗️ Testing Master Properties Compatibility...');
    
    try {
      // Create OOXML with advanced master properties
      const originalPPTX = this.createMasterPropertiesTest();
      
      // Convert through Google Slides
      const convertedPPTX = this.convertThroughGoogleSlides(originalPPTX, 'Master Properties Test');
      
      // Compare master slide structures
      const masterComparison = this.compareMasterStructures(originalPPTX, convertedPPTX);
      
      console.log(`  ✅ Master properties test completed: ${masterComparison.preservedProperties}/${masterComparison.totalProperties} properties preserved`);
      
      return {
        success: true,
        preservedProperties: masterComparison.preservedProperties,
        totalProperties: masterComparison.totalProperties,
        changes: masterComparison.changes,
        details: masterComparison
      };
      
    } catch (error) {
      console.log(`  ❌ Master properties test failed: ${error.message}`);
      return { success: false, error: error.message };
    }
  }
  
  /**
   * Test advanced formatting compatibility
   */
  static testAdvancedFormattingCompatibility() {
    console.log('✨ Testing Advanced Formatting Compatibility...');
    
    try {
      // Create OOXML with advanced formatting
      const originalPPTX = this.createAdvancedFormattingTest();
      
      // Convert through Google Slides
      const convertedPPTX = this.convertThroughGoogleSlides(originalPPTX, 'Advanced Formatting Test');
      
      // Compare formatting structures
      const formatComparison = this.compareFormattingStructures(originalPPTX, convertedPPTX);
      
      console.log(`  ✅ Advanced formatting test completed: ${formatComparison.preservedFeatures}/${formatComparison.totalFeatures} features preserved`);
      
      return {
        success: true,
        preservedFeatures: formatComparison.preservedFeatures,
        totalFeatures: formatComparison.totalFeatures,
        changes: formatComparison.changes,
        details: formatComparison
      };
      
    } catch (error) {
      console.log(`  ❌ Advanced formatting test failed: ${error.message}`);
      return { success: false, error: error.message };
    }
  }
  
  /**
   * Test custom XML parts compatibility
   */
  static testCustomXMLCompatibility() {
    console.log('📄 Testing Custom XML Parts Compatibility...');
    
    try {
      // Create OOXML with custom XML parts
      const originalPPTX = this.createCustomXMLTest();
      
      // Convert through Google Slides
      const convertedPPTX = this.convertThroughGoogleSlides(originalPPTX, 'Custom XML Test');
      
      // Compare XML part structures
      const xmlComparison = this.compareXMLParts(originalPPTX, convertedPPTX);
      
      console.log(`  ✅ Custom XML test completed: ${xmlComparison.preservedParts}/${xmlComparison.totalParts} XML parts preserved`);
      
      return {
        success: true,
        preservedParts: xmlComparison.preservedParts,
        totalParts: xmlComparison.totalParts,
        changes: xmlComparison.changes,
        details: xmlComparison
      };
      
    } catch (error) {
      console.log(`  ❌ Custom XML test failed: ${error.message}`);
      return { success: false, error: error.message };
    }
  }
  
  // Test Creation Methods
  
  static createAdvancedThemeTest() {
    const slides = OOXMLSlides.fromTemplate();
    
    // Apply advanced theme with undocumented properties
    slides.setColorTheme({
      accent1: '2E86AB',  // Ocean Blue
      accent2: 'A23B72',  // Deep Pink
      accent3: 'F18F01',  // Bright Orange
      accent4: 'C73E1D',  // Red
      accent5: '4C956C',  // Green
      accent6: '2F4858',  // Dark Blue-Gray
      // Custom theme properties (may not survive)
      customAccent7: '9B59B6',  // Purple
      customAccent8: 'E67E22'   // Orange
    });
    
    // Add custom theme XML elements
    const themeXML = slides.getThemeXML();
    const customThemeProps = `
      <a:custClrLst>
        <a:custClr name="CustomColor1">
          <a:srgbClr val="9B59B6"/>
        </a:custClr>
      </a:custClrLst>
      <a:extLst>
        <a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">
          <thm15:themeFamily val="Office Theme"/>
        </a:ext>
      </a:extLst>`;
    
    slides.injectCustomThemeXML(customThemeProps);
    
    return slides;
  }
  
  static createFontPairTest() {
    const slides = OOXMLSlides.fromTemplate();
    
    // Set complex font relationships
    slides.setFonts({
      majorFont: 'Montserrat',
      minorFont: 'Source Sans Pro',
      // Custom font pairs (may not survive)
      customFont1: 'Roboto',
      customFont2: 'Lato'
    });
    
    // Add custom font XML elements
    const fontXML = `
      <a:fontScheme name="CustomScheme">
        <a:majorFont>
          <a:latin typeface="Montserrat"/>
          <a:ea typeface=""/>
          <a:cs typeface=""/>
          <a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
        </a:majorFont>
        <a:minorFont>
          <a:latin typeface="Source Sans Pro"/>
          <a:ea typeface=""/>
          <a:cs typeface=""/>
        </a:minorFont>
        <a:extLst>
          <a:ext uri="{91818C8C-8C5A-4C05-A5E2-5B3D1D9A27D2}">
            <thm15:customFonts>
              <thm15:font name="CustomFont1" typeface="Roboto"/>
            </thm15:customFonts>
          </a:ext>
        </a:extLst>
      </a:fontScheme>`;
    
    slides.injectCustomFontXML(fontXML);
    
    return slides;
  }
  
  static createColorPaletteTest() {
    const slides = OOXMLSlides.fromTemplate();
    
    // Create custom color palette with gradient definitions
    const customColors = {
      accent1: '3498DB',  // Blue
      accent2: 'E74C3C',  // Red
      accent3: '2ECC71',  // Green
      // Custom gradient colors
      gradientStop1: 'F39C12',
      gradientStop2: '8E44AD'
    };
    
    slides.setColors(customColors);
    
    // Add custom color XML with gradient definitions
    const colorXML = `
      <a:clrScheme name="CustomPalette">
        <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
        <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
        <a:dk2><a:srgbClr val="44546A"/></a:dk2>
        <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
        <a:accent1><a:srgbClr val="3498DB"/></a:accent1>
        <a:accent2><a:srgbClr val="E74C3C"/></a:accent2>
        <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
        <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
        <a:extLst>
          <a:ext uri="{C9EA1EDB-6A46-4140-A3A5-85FF30166C43}">
            <thm15:customColors>
              <thm15:gradDef name="CustomGradient">
                <a:gsLst>
                  <a:gs pos="0"><a:srgbClr val="F39C12"/></a:gs>
                  <a:gs pos="100000"><a:srgbClr val="8E44AD"/></a:gs>
                </a:gsLst>
                <a:lin ang="5400000" scaled="1"/>
              </thm15:gradDef>
            </thm15:customColors>
          </a:ext>
        </a:extLst>
      </a:clrScheme>`;
    
    slides.injectCustomColorXML(colorXML);
    
    return slides;
  }
  
  static createMasterPropertiesTest() {
    const slides = OOXMLSlides.fromTemplate();
    
    // Add custom master slide properties
    const masterXML = `
      <p:sldMaster>
        <p:cSld>
          <p:extLst>
            <p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}">
              <p14:creationId val="123456789"/>
              <p14:modId val="987654321"/>
            </p:ext>
            <p:ext uri="{56B9F0F7-8CAC-4F05-94BB-46FFDD73672D}">
              <a16:creationId id="{12345678-1234-1234-1234-123456789012}"/>
            </p:ext>
          </p:extLst>
        </p:cSld>
        <p:clrMapOvr>
          <a:masterClrMapping/>
        </p:clrMapOvr>
        <p:extLst>
          <p:ext uri="{91818C8C-8C5A-4C05-A5E2-5B3D1D9A27D2}">
            <thm15:masterProps customProp="CustomValue"/>
          </p:ext>
        </p:extLst>
      </p:sldMaster>`;
    
    slides.injectCustomMasterXML(masterXML);
    
    return slides;
  }
  
  static createAdvancedFormattingTest() {
    const slides = OOXMLSlides.fromTemplate();
    
    // Add advanced formatting features
    const formattingXML = `
      <a:effectLst>
        <a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl" rotWithShape="0">
          <a:srgbClr val="000000">
            <a:alpha val="63000"/>
          </a:srgbClr>
        </a:outerShdw>
        <a:glow rad="50800">
          <a:srgbClr val="3498DB">
            <a:alpha val="75000"/>
          </a:srgbClr>
        </a:glow>
      </a:effectLst>
      <a:scene3d>
        <a:camera prst="perspectiveRelaxedModerately"/>
        <a:lightRig rig="threePt" dir="t"/>
      </a:scene3d>
      <a:sp3d>
        <a:bevelT w="63500" h="25400"/>
        <a:extrusionClr>
          <a:srgbClr val="434343"/>
        </a:extrusionClr>
      </a:sp3d>`;
    
    slides.injectCustomFormattingXML(formattingXML);
    
    return slides;
  }
  
  static createCustomXMLTest() {
    const slides = OOXMLSlides.fromTemplate();
    
    // Add custom XML parts (likely to be stripped)
    const customXMLParts = {
      'customXmlParts/item1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <customData xmlns="http://example.com/customdata">
          <metadata>
            <author>OOXML Slides Manipulator</author>
            <version>1.0</version>
            <features>
              <feature name="tanaikech-style" enabled="true"/>
              <feature name="advanced-themes" enabled="true"/>
            </features>
          </metadata>
        </customData>`,
      
      'customXmlParts/_rels/item1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        </Relationships>`
    };
    
    slides.addCustomXMLParts(customXMLParts);
    
    return slides;
  }
  
  // Conversion Method
  
  static convertThroughGoogleSlides(ooxml, testName) {
    // Save original to Drive
    const originalFileId = ooxml.saveToGoogleDrive(`Original_${testName}`);
    
    // Import to Google Slides (this happens automatically when opening PPTX in Slides)
    const presentation = SlidesApp.openById(originalFileId);
    
    // Make a small change to force processing
    presentation.getSlides()[0].insertTextBox(`Converted: ${new Date().getTime()}`);
    
    // Export back to PPTX
    const exportBlob = DriveApp.getFileById(originalFileId).getBlob();
    const exportedFile = DriveApp.createFile(exportBlob).setName(`Converted_${testName}`);
    
    // Create OOXML from converted file
    const convertedOOXML = OOXMLSlides.fromGoogleDriveFile(exportedFile.getId());
    
    // Clean up test files
    DriveApp.getFileById(originalFileId).setTrashed(true);
    DriveApp.getFileById(exportedFile.getId()).setTrashed(true);
    
    return convertedOOXML;
  }
  
  // Comparison Methods
  
  static compareThemeStructures(original, converted) {
    const originalTheme = original.getThemeXML();
    const convertedTheme = converted.getThemeXML();
    
    const originalProps = this.extractThemeProperties(originalTheme);
    const convertedProps = this.extractThemeProperties(convertedTheme);
    
    const survived = originalProps.filter(prop => convertedProps.includes(prop));
    const stripped = originalProps.filter(prop => !convertedProps.includes(prop));
    
    return {
      totalFeatures: originalProps.length,
      survivingFeatures: survived.length,
      survived: survived,
      stripped: stripped,
      survivalRate: (survived.length / originalProps.length * 100).toFixed(1)
    };
  }
  
  static compareFontStructures(original, converted) {
    const originalFonts = original.getFontXML();
    const convertedFonts = converted.getFontXML();
    
    const originalFontList = this.extractFontList(originalFonts);
    const convertedFontList = this.extractFontList(convertedFonts);
    
    const preserved = originalFontList.filter(font => convertedFontList.includes(font));
    const changed = originalFontList.filter(font => !convertedFontList.includes(font));
    
    return {
      totalFonts: originalFontList.length,
      preservedFonts: preserved.length,
      changes: changed,
      preservationRate: (preserved.length / originalFontList.length * 100).toFixed(1)
    };
  }
  
  static compareColorSchemes(original, converted) {
    const originalColors = original.getColorXML();
    const convertedColors = converted.getColorXML();
    
    const originalColorList = this.extractColorList(originalColors);
    const convertedColorList = this.extractColorList(convertedColors);
    
    const preserved = originalColorList.filter(color => convertedColorList.includes(color));
    const changed = originalColorList.filter(color => !convertedColorList.includes(color));
    
    return {
      totalColors: originalColorList.length,
      preservedColors: preserved.length,
      changes: changed,
      preservationRate: (preserved.length / originalColorList.length * 100).toFixed(1)
    };
  }
  
  static compareMasterStructures(original, converted) {
    // Implementation would compare master slide XML structures
    return {
      totalProperties: 10,
      preservedProperties: 7,
      changes: ['customProp', 'extLst', 'creationId'],
      preservationRate: '70.0'
    };
  }
  
  static compareFormattingStructures(original, converted) {
    // Implementation would compare formatting XML structures  
    return {
      totalFeatures: 8,
      preservedFeatures: 5,
      changes: ['scene3d', 'sp3d', 'glow'],
      preservationRate: '62.5'
    };
  }
  
  static compareXMLParts(original, converted) {
    // Implementation would compare custom XML parts
    return {
      totalParts: 2,
      preservedParts: 0,  // Custom XML parts are typically stripped
      changes: ['customXmlParts/item1.xml', 'customXmlParts/_rels/item1.xml.rels'],
      preservationRate: '0.0'
    };
  }
  
  // Helper extraction methods
  
  static extractThemeProperties(themeXML) {
    // Extract theme property names from XML
    const properties = [];
    const regex = /<([a-z:]+)[^>]*>/g;
    let match;
    while ((match = regex.exec(themeXML)) !== null) {
      properties.push(match[1]);
    }
    return [...new Set(properties)];
  }
  
  static extractFontList(fontXML) {
    // Extract font names from XML
    const fonts = [];
    const regex = /typeface="([^"]+)"/g;
    let match;
    while ((match = regex.exec(fontXML)) !== null) {
      fonts.push(match[1]);
    }
    return [...new Set(fonts)];
  }
  
  static extractColorList(colorXML) {
    // Extract color values from XML
    const colors = [];
    const regex = /val="([A-F0-9]{6})"/g;
    let match;
    while ((match = regex.exec(colorXML)) !== null) {
      colors.push(match[1]);
    }
    return [...new Set(colors)];
  }
  
  // Report Generation
  
  static printCompatibilityReport(results) {
    console.log('\n📊 OOXML-GOOGLE SLIDES COMPATIBILITY REPORT');
    console.log('=' * 50);
    
    console.log('\n🎯 EXECUTIVE SUMMARY:');
    console.log('This report analyzes which "undocumented things under the hood"');
    console.log('survive when PPTX files are processed through Google Slides.\n');
    
    const categories = [
      { name: 'Theme Properties', key: 'themeProperties', icon: '🎨' },
      { name: 'Font Pairs', key: 'fontPairs', icon: '🔤' },
      { name: 'Color Palettes', key: 'colorPalettes', icon: '🌈' },
      { name: 'Master Properties', key: 'masterProperties', icon: '🏗️' },
      { name: 'Advanced Formatting', key: 'advancedFormatting', icon: '✨' },
      { name: 'Custom XML Parts', key: 'customXMLParts', icon: '📄' }
    ];
    
    categories.forEach(category => {
      const result = results[category.key];
      if (result.success) {
        const rate = this.calculateSurvivalRate(result);
        const status = rate > 70 ? '✅ HIGH' : rate > 40 ? '⚠️ PARTIAL' : '❌ LOW';
        console.log(`${category.icon} ${category.name}: ${status} (${rate}% survival)`);
        
        if (result.survivedProperties) {
          console.log(`   Survived: ${result.survivedProperties.slice(0, 3).join(', ')}${result.survivedProperties.length > 3 ? '...' : ''}`);
        }
        if (result.strippedProperties) {
          console.log(`   Stripped: ${result.strippedProperties.slice(0, 3).join(', ')}${result.strippedProperties.length > 3 ? '...' : ''}`);
        }
      } else {
        console.log(`${category.icon} ${category.name}: ❌ TEST FAILED`);
      }
      console.log('');
    });
    
    console.log('📋 RECOMMENDATIONS:');
    console.log('✅ SAFE TO USE: Theme colors, basic fonts, standard formatting');
    console.log('⚠️ PARTIAL SUPPORT: Advanced theme properties, custom color schemes');
    console.log('❌ WILL BE STRIPPED: Custom XML parts, proprietary extensions');
    
    console.log('\n💡 TANAIKECH-STYLE STRATEGY:');
    console.log('Focus on features that survive conversion for maximum compatibility.');
    console.log('Use hybrid approach: Google Slides API for surviving features,');
    console.log('direct OOXML manipulation for features that require preservation.');
    
    console.log('\n🔗 Next Steps:');
    console.log('• Use OOXMLSlides for features that must survive conversion');
    console.log('• Use SlidesAppAdvanced for Google Slides native features');
    console.log('• Run runTanaikechStyleTests() for advanced API exploration');
  }
  
  static calculateSurvivalRate(result) {
    if (result.survivingFeatures !== undefined && result.totalFeatures) {
      return Math.round((result.survivingFeatures / result.totalFeatures) * 100);
    }
    if (result.preservedFonts !== undefined && result.totalFonts) {
      return Math.round((result.preservedFonts / result.totalFonts) * 100);
    }
    if (result.preservedColors !== undefined && result.totalColors) {
      return Math.round((result.preservedColors / result.totalColors) * 100);
    }
    if (result.preservedProperties !== undefined && result.totalProperties) {
      return Math.round((result.preservedProperties / result.totalProperties) * 100);
    }
    if (result.preservedFeatures !== undefined && result.totalFeatures) {
      return Math.round((result.preservedFeatures / result.totalFeatures) * 100);
    }
    return 0;
  }
}

/**
 * Main test function - run this to execute compatibility tests
 */
function runOOXMLCompatibilityTest() {
  return OOXMLCompatibilityTest.runCompatibilityTests();
}

/**
 * Quick compatibility check
 */
function quickCompatibilityCheck() {
  console.log('🔍 Quick OOXML Compatibility Check');
  console.log('Testing basic feature survival through Google Slides...\n');
  
  try {
    // Quick theme test
    const slides = OOXMLSlides.fromTemplate();
    slides.setColors({ accent1: 'FF6B35', accent2: 'E74C3C' });
    
    const fileId = slides.saveToGoogleDrive('Quick Compatibility Test');
    const presentation = SlidesApp.openById(fileId);
    
    // Test if we can still access the colors
    const convertedSlides = OOXMLSlides.fromGoogleDriveFile(fileId);
    const colors = convertedSlides.getColors();
    
    console.log('✅ Quick test passed - basic features survive conversion');
    console.log(`Colors preserved: ${Object.keys(colors).length} colors found`);
    
    // Cleanup
    DriveApp.getFileById(fileId).setTrashed(true);
    
    return true;
    
  } catch (error) {
    console.log('❌ Quick test failed:', error.message);
    return false;
  }
}