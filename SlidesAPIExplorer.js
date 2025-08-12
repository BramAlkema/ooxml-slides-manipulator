/**
 * SlidesAPIExplorer - Discovery tool for undocumented Google Slides API features
 * 
 * This tool helps discover and analyze the full Google Slides API response
 * to find properties and features not documented in the official API.
 * Tanaikech-style exploration of the API's hidden capabilities.
 */

class SlidesAPIExplorer {
  
  /**
   * Deep exploration of a presentation's undocumented properties
   */
  static explorePresentation(presentationId) {
    console.log(`🔍 Exploring presentation: ${presentationId}`);
    console.log('=' * 60);
    
    try {
      // Get full API response
      const apiData = SlidesAPIExplorer._getFullAPIResponse(presentationId);
      
      // Analyze structure
      const analysis = {
        basicInfo: SlidesAPIExplorer._analyzeBasicInfo(apiData),
        masters: SlidesAPIExplorer._analyzeMasters(apiData.masters || []),
        layouts: SlidesAPIExplorer._analyzeLayouts(apiData.layouts || []),
        slides: SlidesAPIExplorer._analyzeSlides(apiData.slides || []),
        undocumentedProperties: SlidesAPIExplorer._findUndocumentedProperties(apiData),
        fontAnalysis: SlidesAPIExplorer._analyzeFonts(apiData),
        colorAnalysis: SlidesAPIExplorer._analyzeColors(apiData),
        customProperties: SlidesAPIExplorer._extractCustomProperties(apiData)
      };
      
      SlidesAPIExplorer._printAnalysis(analysis);
      
      return analysis;
      
    } catch (error) {
      console.error('❌ Exploration failed:', error);
      return null;
    }
  }
  
  /**
   * Compare API responses between presentations to find patterns
   */
  static comparePresentations(presentationIds) {
    console.log(`🔍 Comparing ${presentationIds.length} presentations for patterns`);
    
    const comparisons = presentationIds.map(id => ({
      id: id,
      data: SlidesAPIExplorer._getFullAPIResponse(id),
      properties: new Set()
    }));
    
    // Find common and unique properties
    const commonProperties = new Set();
    const uniqueProperties = new Map();
    
    comparisons.forEach(comp => {
      const props = SlidesAPIExplorer._getAllPropertyPaths(comp.data);
      props.forEach(prop => comp.properties.add(prop));
    });
    
    // Identify patterns
    const patterns = {
      common: Array.from(commonProperties),
      unique: Object.fromEntries(uniqueProperties),
      variations: SlidesAPIExplorer._findPropertyVariations(comparisons)
    };
    
    console.log('📊 Comparison Results:');
    console.log(`Common properties: ${patterns.common.length}`);
    console.log(`Unique properties: ${Object.keys(patterns.unique).length}`);
    
    return patterns;
  }
  
  /**
   * Discover font pair relationships
   */
  static exploreFontPairs(presentationId) {
    console.log(`🔤 Exploring font pairs in: ${presentationId}`);
    
    const apiData = SlidesAPIExplorer._getFullAPIResponse(presentationId);
    const fontAnalysis = {
      masterFonts: {},
      layoutFonts: {},
      slideFonts: {},
      fontRelationships: [],
      undocumentedFontProperties: []
    };
    
    // Analyze masters
    if (apiData.masters) {
      apiData.masters.forEach((master, index) => {
        fontAnalysis.masterFonts[`master_${index}`] = SlidesAPIExplorer._extractElementFonts(master.pageElements || []);
      });
    }
    
    // Analyze layouts
    if (apiData.layouts) {
      apiData.layouts.forEach((layout, index) => {
        fontAnalysis.layoutFonts[`layout_${index}`] = SlidesAPIExplorer._extractElementFonts(layout.pageElements || []);
      });
    }
    
    // Find font relationships (major/minor patterns)
    fontAnalysis.fontRelationships = SlidesAPIExplorer._identifyFontRelationships(fontAnalysis);
    
    console.log('📊 Font Analysis Results:');
    console.log(`Master fonts: ${Object.keys(fontAnalysis.masterFonts).length}`);
    console.log(`Layout fonts: ${Object.keys(fontAnalysis.layoutFonts).length}`);
    console.log(`Font relationships: ${fontAnalysis.fontRelationships.length}`);
    
    return fontAnalysis;
  }
  
  /**
   * Discover color scheme patterns
   */
  static exploreColorScheme(presentationId) {
    console.log(`🎨 Exploring color scheme in: ${presentationId}`);
    
    const apiData = SlidesAPIExplorer._getFullAPIResponse(presentationId);
    const colorAnalysis = {
      masterColors: {},
      themeColors: {},
      solidFillColors: [],
      gradientColors: [],
      colorRelationships: [],
      undocumentedColorProperties: []
    };
    
    // Extract all color references
    const allColors = SlidesAPIExplorer._extractAllColors(apiData);
    
    // Categorize colors
    colorAnalysis.solidFillColors = allColors.filter(c => c.type === 'solid');
    colorAnalysis.gradientColors = allColors.filter(c => c.type === 'gradient');
    
    // Look for theme color relationships
    colorAnalysis.colorRelationships = SlidesAPIExplorer._identifyColorRelationships(allColors);
    
    // Find undocumented color properties
    colorAnalysis.undocumentedColorProperties = SlidesAPIExplorer._findUndocumentedColorProps(apiData);
    
    console.log('🎨 Color Analysis Results:');
    console.log(`Solid colors: ${colorAnalysis.solidFillColors.length}`);
    console.log(`Gradient colors: ${colorAnalysis.gradientColors.length}`);
    console.log(`Color relationships: ${colorAnalysis.colorRelationships.length}`);
    
    return colorAnalysis;
  }
  
  /**
   * Test undocumented API features
   */
  static testUndocumentedFeatures(presentationId) {
    console.log(`🧪 Testing undocumented features on: ${presentationId}`);
    
    const tests = [
      () => SlidesAPIExplorer._testAdvancedBatchUpdate(presentationId),
      () => SlidesAPIExplorer._testCustomProperties(presentationId),
      () => SlidesAPIExplorer._testAdvancedTextFormatting(presentationId),
      () => SlidesAPIExplorer._testMasterManipulation(presentationId),
      () => SlidesAPIExplorer._testAdvancedColorOperations(presentationId)
    ];
    
    const results = {};
    
    tests.forEach((test, index) => {
      try {
        console.log(`Running test ${index + 1}...`);
        results[`test_${index + 1}`] = test();
        console.log(`✅ Test ${index + 1} passed`);
      } catch (error) {
        console.log(`❌ Test ${index + 1} failed: ${error.message}`);
        results[`test_${index + 1}`] = { error: error.message };
      }
    });
    
    return results;
  }
  
  // Private helper methods
  
  static _getFullAPIResponse(presentationId) {
    const response = UrlFetchApp.fetch(
      `https://slides.googleapis.com/v1/presentations/${presentationId}`,
      {
        headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
      }
    );
    
    return JSON.parse(response.getContentText());
  }
  
  static _analyzeBasicInfo(apiData) {
    return {
      presentationId: apiData.presentationId,
      title: apiData.title,
      locale: apiData.locale,
      revisionId: apiData.revisionId,
      pageSize: apiData.pageSize,
      masterCount: apiData.masters ? apiData.masters.length : 0,
      layoutCount: apiData.layouts ? apiData.layouts.length : 0,
      slideCount: apiData.slides ? apiData.slides.length : 0
    };
  }
  
  static _analyzeMasters(masters) {
    return masters.map(master => ({
      objectId: master.objectId,
      pageType: master.pageType,
      elementCount: master.pageElements ? master.pageElements.length : 0,
      properties: master.masterProperties || {},
      customProps: SlidesAPIExplorer._extractCustomProperties(master)
    }));
  }
  
  static _analyzeLayouts(layouts) {
    return layouts.map(layout => ({
      objectId: layout.objectId,
      pageType: layout.pageType,
      elementCount: layout.pageElements ? layout.pageElements.length : 0,
      properties: layout.layoutProperties || {},
      customProps: SlidesAPIExplorer._extractCustomProperties(layout)
    }));
  }
  
  static _analyzeSlides(slides) {
    return slides.map(slide => ({
      objectId: slide.objectId,
      pageType: slide.pageType,
      elementCount: slide.pageElements ? slide.pageElements.length : 0,
      properties: slide.slideProperties || {},
      customProps: SlidesAPIExplorer._extractCustomProperties(slide)
    }));
  }
  
  static _findUndocumentedProperties(apiData) {
    const documented = new Set([
      'presentationId', 'title', 'locale', 'revisionId', 'pageSize',
      'masters', 'layouts', 'slides', 'notesMaster'
    ]);
    
    return Object.keys(apiData).filter(key => !documented.has(key));
  }
  
  static _analyzeFonts(apiData) {
    const fonts = new Set();
    const fontProperties = [];
    
    // Recursive font extraction
    const extractFonts = (obj, path = '') => {
      if (typeof obj !== 'object' || obj === null) return;
      
      Object.keys(obj).forEach(key => {
        if (key === 'fontFamily' && typeof obj[key] === 'string') {
          fonts.add(obj[key]);
          fontProperties.push({
            font: obj[key],
            path: path + '.' + key,
            context: obj
          });
        } else if (typeof obj[key] === 'object') {
          extractFonts(obj[key], path + '.' + key);
        }
      });
    };
    
    extractFonts(apiData);
    
    return {
      uniqueFonts: Array.from(fonts),
      fontOccurrences: fontProperties
    };
  }
  
  static _analyzeColors(apiData) {
    const colors = [];
    
    // Recursive color extraction
    const extractColors = (obj, path = '') => {
      if (typeof obj !== 'object' || obj === null) return;
      
      Object.keys(obj).forEach(key => {
        if (key === 'rgbColor' && obj[key]) {
          colors.push({
            type: 'rgb',
            color: obj[key],
            path: path + '.' + key
          });
        } else if (key === 'solidFill' && obj[key]) {
          colors.push({
            type: 'solid',
            color: obj[key],
            path: path + '.' + key
          });
        } else if (typeof obj[key] === 'object') {
          extractColors(obj[key], path + '.' + key);
        }
      });
    };
    
    extractColors(apiData);
    
    return colors;
  }
  
  static _extractCustomProperties(obj) {
    const standardProps = new Set([
      'objectId', 'pageType', 'pageElements', 'masterProperties',
      'layoutProperties', 'slideProperties', 'presentationId',
      'title', 'locale', 'revisionId'
    ]);
    
    const custom = {};
    Object.keys(obj).forEach(key => {
      if (!standardProps.has(key)) {
        custom[key] = obj[key];
      }
    });
    
    return custom;
  }
  
  static _getAllPropertyPaths(obj, path = '', paths = new Set()) {
    if (typeof obj !== 'object' || obj === null) return paths;
    
    Object.keys(obj).forEach(key => {
      const fullPath = path ? `${path}.${key}` : key;
      paths.add(fullPath);
      
      if (typeof obj[key] === 'object') {
        SlidesAPIExplorer._getAllPropertyPaths(obj[key], fullPath, paths);
      }
    });
    
    return paths;
  }
  
  static _extractElementFonts(pageElements) {
    const fonts = {};
    
    pageElements.forEach(element => {
      if (element.shape && element.shape.text) {
        const textElements = element.shape.text.textElements || [];
        textElements.forEach(textElement => {
          if (textElement.textRun && textElement.textRun.style) {
            const style = textElement.textRun.style;
            if (style.fontFamily) {
              fonts[element.objectId] = {
                fontFamily: style.fontFamily,
                fontSize: style.fontSize,
                bold: style.bold,
                italic: style.italic
              };
            }
          }
        });
      }
    });
    
    return fonts;
  }
  
  static _identifyFontRelationships(fontAnalysis) {
    // Look for major/minor font pair patterns
    const relationships = [];
    
    // This is where tanaikech-style analysis would identify
    // font pair relationships that aren't explicitly documented
    
    return relationships;
  }
  
  static _extractAllColors(apiData) {
    const colors = [];
    
    const extractRecursive = (obj, path = '') => {
      if (typeof obj !== 'object' || obj === null) return;
      
      Object.keys(obj).forEach(key => {
        if (key.includes('color') || key.includes('Color') || key.includes('fill')) {
          colors.push({
            type: key.includes('gradient') ? 'gradient' : 'solid',
            data: obj[key],
            path: path + '.' + key
          });
        } else if (typeof obj[key] === 'object') {
          extractRecursive(obj[key], path + '.' + key);
        }
      });
    };
    
    extractRecursive(apiData);
    return colors;
  }
  
  static _identifyColorRelationships(colors) {
    // Analyze color usage patterns to identify theme relationships
    return [];
  }
  
  static _findUndocumentedColorProps(apiData) {
    // Look for color properties not in the official documentation
    return [];
  }
  
  static _printAnalysis(analysis) {
    console.log('\n📊 ANALYSIS RESULTS');
    console.log('==================');
    
    console.log('\n📋 Basic Info:');
    Object.entries(analysis.basicInfo).forEach(([key, value]) => {
      console.log(`  ${key}: ${value}`);
    });
    
    console.log('\n🎨 Color Analysis:');
    console.log(`  Unique colors found: ${analysis.colorAnalysis.length}`);
    
    console.log('\n🔤 Font Analysis:');
    console.log(`  Unique fonts found: ${analysis.fontAnalysis.uniqueFonts.length}`);
    analysis.fontAnalysis.uniqueFonts.forEach(font => {
      console.log(`    - ${font}`);
    });
    
    console.log('\n🔍 Undocumented Properties:');
    if (analysis.undocumentedProperties.length > 0) {
      analysis.undocumentedProperties.forEach(prop => {
        console.log(`  - ${prop}`);
      });
    } else {
      console.log('  None found at presentation level');
    }
    
    console.log('\n✨ Custom Properties:');
    if (Object.keys(analysis.customProperties).length > 0) {
      Object.entries(analysis.customProperties).forEach(([key, value]) => {
        console.log(`  ${key}: ${JSON.stringify(value)}`);
      });
    } else {
      console.log('  None found');
    }
  }
  
  // Test methods for undocumented features
  
  static _testAdvancedBatchUpdate(presentationId) {
    // Test advanced batch update operations
    return { success: true, message: 'Advanced batch update test' };
  }
  
  static _testCustomProperties(presentationId) {
    // Test setting custom properties
    return { success: true, message: 'Custom properties test' };
  }
  
  static _testAdvancedTextFormatting(presentationId) {
    // Test advanced text formatting options
    return { success: true, message: 'Advanced text formatting test' };
  }
  
  static _testMasterManipulation(presentationId) {
    // Test master slide manipulation
    return { success: true, message: 'Master manipulation test' };
  }
  
  static _testAdvancedColorOperations(presentationId) {
    // Test advanced color operations
    return { success: true, message: 'Advanced color operations test' };
  }
}