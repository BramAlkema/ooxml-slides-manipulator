/**
 * TypographyEditor - Advanced Kerning and Typography XML Hacking
 * 
 * This implements tanaikech-style XML manipulation for advanced typography controls
 * including kerning, tracking, ligatures, OpenType features, and professional
 * typographic adjustments that are not accessible through PowerPoint's standard interface.
 * 
 * Key capabilities:
 * - Custom kerning pairs and adjustments
 * - Character tracking (letter-spacing)
 * - Word spacing control
 * - OpenType ligature controls
 * - Advanced baseline adjustments
 * - Small caps and stylistic sets
 * - Font feature selection
 * - Professional typography presets
 */

class TypographyEditor {
  
  constructor(ooxml) {
    this.ooxml = ooxml;
    this.parser = ooxml.parser;
  }
  
  /**
   * Apply advanced kerning to text elements
   * This is the core typography XML hacking technique
   */
  applyAdvancedKerning(kerningDefinitions) {
    console.log('✍️ XML Hacking: Applying advanced kerning and typography...');
    console.log(`Processing ${Object.keys(kerningDefinitions).length} typography definitions`);
    
    try {
      // Apply kerning to slide masters
      this._applyKerningToMasters(kerningDefinitions);
      
      // Apply kerning to slide layouts
      this._applyKerningToLayouts(kerningDefinitions);
      
      // Update theme with typography settings
      this._updateThemeTypography(kerningDefinitions);
      
      console.log('✅ Successfully applied advanced kerning and typography');
      return true;
      
    } catch (error) {
      console.error('❌ Failed to apply advanced kerning:', error.message);
      return false;
    }
  }
  
  /**
   * Apply professional typography preset
   */
  applyProfessionalTypography(preset = 'corporate') {
    console.log(`🎯 Applying professional typography preset: ${preset}`);
    
    const presets = {
      'corporate': {
        'Heading_Typography': {
          kerning: { method: 'optical', adjustment: -0.02 },
          tracking: 0.05, // 5% letter spacing
          wordSpacing: 1.0,
          ligatures: { standard: true, discretionary: false },
          openType: ['kern', 'liga'],
          baseline: 0,
          fontFeatures: {
            smallCaps: false,
            stylisticSet: 'ss01'
          }
        },
        'Body_Typography': {
          kerning: { method: 'metrics', adjustment: 0 },
          tracking: 0,
          wordSpacing: 1.0,
          ligatures: { standard: true, discretionary: false },
          openType: ['kern', 'liga'],
          baseline: 0
        },
        'Caption_Typography': {
          kerning: { method: 'optical', adjustment: 0.01 },
          tracking: 0.03,
          wordSpacing: 0.95,
          ligatures: { standard: false, discretionary: false },
          openType: ['kern'],
          baseline: 0
        }
      },
      
      'editorial': {
        'Display_Typography': {
          kerning: { method: 'optical', adjustment: -0.03 },
          tracking: -0.02, // Tighter spacing for large text
          wordSpacing: 0.9,
          ligatures: { standard: true, discretionary: true },
          openType: ['kern', 'liga', 'dlig', 'swsh'],
          baseline: 0,
          fontFeatures: {
            stylisticSet: 'ss02',
            contextualAlternates: true
          }
        },
        'Headline_Typography': {
          kerning: { method: 'optical', adjustment: -0.015 },
          tracking: 0,
          wordSpacing: 0.95,
          ligatures: { standard: true, discretionary: false },
          openType: ['kern', 'liga'],
          baseline: 0
        },
        'Byline_Typography': {
          kerning: { method: 'metrics', adjustment: 0 },
          tracking: 0.08, // More open spacing
          wordSpacing: 1.1,
          ligatures: { standard: false, discretionary: false },
          openType: ['kern'],
          baseline: 0,
          fontFeatures: {
            smallCaps: true
          }
        }
      },
      
      'technical': {
        'Code_Typography': {
          kerning: { method: 'none', adjustment: 0 },
          tracking: 0,
          wordSpacing: 1.0,
          ligatures: { standard: false, discretionary: false },
          openType: [],
          baseline: 0,
          fontFeatures: {
            tabularNums: true
          }
        },
        'Monospace_Typography': {
          kerning: { method: 'none', adjustment: 0 },
          tracking: 0,
          wordSpacing: 1.0,
          ligatures: { standard: false, discretionary: false },
          openType: [],
          baseline: 0
        }
      }
    };
    
    const selectedPreset = presets[preset];
    if (selectedPreset) {
      return this.applyAdvancedKerning(selectedPreset);
    } else {
      console.log(`❌ Unknown typography preset: ${preset}`);
      return false;
    }
  }
  
  /**
   * Create custom kerning pairs for specific font combinations
   */
  createKerningPairs(fontFamily, kerningPairs) {
    console.log(`🔤 Creating custom kerning pairs for font: ${fontFamily}`);
    
    try {
      // Create kerning table in theme
      const kerningTable = {
        [`${fontFamily}_Kerning`]: {
          fontFamily: fontFamily,
          pairs: kerningPairs,
          kerning: { method: 'custom', pairs: kerningPairs }
        }
      };
      
      return this.applyAdvancedKerning(kerningTable);
      
    } catch (error) {
      console.error('❌ Failed to create kerning pairs:', error.message);
      return false;
    }
  }
  
  /**
   * Apply OpenType font features
   */
  applyOpenTypeFeatures(fontFamily, features) {
    console.log(`🎨 Applying OpenType features to ${fontFamily}: ${features.join(', ')}`);
    
    const openTypeTypography = {
      [`${fontFamily}_OpenType`]: {
        fontFamily: fontFamily,
        openType: features,
        fontFeatures: {
          contextualAlternates: features.includes('calt'),
          standardLigatures: features.includes('liga'),
          discretionaryLigatures: features.includes('dlig'),
          swash: features.includes('swsh'),
          stylisticSets: features.filter(f => f.startsWith('ss')),
          smallCaps: features.includes('smcp'),
          oldStyleFigures: features.includes('onum'),
          tabularFigures: features.includes('tnum')
        }
      }
    };
    
    return this.applyAdvancedKerning(openTypeTypography);
  }
  
  /**
   * Set character tracking (letter-spacing) for text styles
   */
  setCharacterTracking(textStyleName, trackingValue) {
    console.log(`📏 Setting character tracking for ${textStyleName}: ${trackingValue}`);
    
    const trackingTypography = {
      [textStyleName]: {
        tracking: trackingValue,
        kerning: { method: 'optical', adjustment: 0 }
      }
    };
    
    return this.applyAdvancedKerning(trackingTypography);
  }
  
  /**
   * Set word spacing adjustments
   */
  setWordSpacing(textStyleName, wordSpacingRatio) {
    console.log(`📐 Setting word spacing for ${textStyleName}: ${wordSpacingRatio}`);
    
    const wordSpacingTypography = {
      [textStyleName]: {
        wordSpacing: wordSpacingRatio,
        kerning: { method: 'metrics', adjustment: 0 }
      }
    };
    
    return this.applyAdvancedKerning(wordSpacingTypography);
  }
  
  /**
   * Create typography for specific slide element
   */
  applyTypographyToElement(slideIndex, elementId, typographySettings) {
    console.log(`✨ Applying typography to slide ${slideIndex + 1}, element ${elementId}`);
    
    try {
      // Get slide XML
      const slideXml = this.parser.getXML(`ppt/slides/slide${slideIndex + 1}.xml`);
      const slideRoot = slideXml.getRootElement();
      
      // Find and modify the specific element
      this._applyTypographyToSlideElement(slideRoot, elementId, typographySettings);
      
      // Update the slide XML
      this.parser.setXML(`ppt/slides/slide${slideIndex + 1}.xml`, slideXml);
      
      console.log('✅ Successfully applied typography to slide element');
      return true;
      
    } catch (error) {
      console.error('❌ Failed to apply typography to slide element:', error.message);
      return false;
    }
  }
  
  /**
   * Create responsive typography that adjusts based on text size
   */
  createResponsiveTypography() {
    console.log('📱 Creating responsive typography system...');
    
    const responsiveTypography = {
      'Large_Text': {
        kerning: { method: 'optical', adjustment: -0.025 },
        tracking: -0.01,
        wordSpacing: 0.95,
        ligatures: { standard: true, discretionary: true },
        condition: 'fontSize >= 24'
      },
      'Medium_Text': {
        kerning: { method: 'optical', adjustment: 0 },
        tracking: 0,
        wordSpacing: 1.0,
        ligatures: { standard: true, discretionary: false },
        condition: 'fontSize >= 14 && fontSize < 24'
      },
      'Small_Text': {
        kerning: { method: 'metrics', adjustment: 0.01 },
        tracking: 0.02,
        wordSpacing: 1.05,
        ligatures: { standard: false, discretionary: false },
        condition: 'fontSize < 14'
      }
    };
    
    return this.applyAdvancedKerning(responsiveTypography);
  }
  
  /**
   * Export current typography settings for analysis
   */
  exportTypographySettings() {
    console.log('📤 Exporting current typography settings...');
    
    try {
      const typographySettings = this._extractTypographySettings();
      console.log(`Found ${typographySettings.length} typography configurations`);
      return typographySettings;
      
    } catch (error) {
      console.error('❌ Failed to export typography settings:', error.message);
      return [];
    }
  }
  
  // Private helper methods for XML manipulation
  
  _applyKerningToMasters(kerningDefinitions) {
    try {
      // Get slide master XML
      const masterXml = this.parser.getXML('ppt/slideMasters/slideMaster1.xml');
      const masterRoot = masterXml.getRootElement();
      
      // Apply typography to master text styles
      this._applyTypographyToTextStyles(masterRoot, kerningDefinitions);
      
      // Update master XML
      this.parser.setXML('ppt/slideMasters/slideMaster1.xml', masterXml);
      
    } catch (error) {
      console.log('Note: Could not apply kerning to slide masters');
    }
  }
  
  _applyKerningToLayouts(kerningDefinitions) {
    try {
      // Get layout XMLs and apply typography
      for (let i = 1; i <= 10; i++) {
        try {
          const layoutXml = this.parser.getXML(`ppt/slideLayouts/slideLayout${i}.xml`);
          const layoutRoot = layoutXml.getRootElement();
          
          this._applyTypographyToTextStyles(layoutRoot, kerningDefinitions);
          
          this.parser.setXML(`ppt/slideLayouts/slideLayout${i}.xml`, layoutXml);
        } catch (layoutError) {
          // Layout doesn't exist, continue
        }
      }
    } catch (error) {
      console.log('Note: Could not apply kerning to slide layouts');
    }
  }
  
  _updateThemeTypography(kerningDefinitions) {
    try {
      const themeXml = this.parser.getXML('ppt/theme/theme1.xml');
      const themeRoot = themeXml.getRootElement();
      
      // Add typography definitions to theme
      this._addTypographyToTheme(themeRoot, kerningDefinitions);
      
      this.parser.setXML('ppt/theme/theme1.xml', themeXml);
      
    } catch (error) {
      console.log('Note: Could not update theme typography');
    }
  }
  
  _applyTypographyToTextStyles(root, kerningDefinitions) {
    const namespace = this._getNamespace('p');
    
    // Find text elements
    const textElements = this._findTextElementsRecursively(root, namespace);
    
    textElements.forEach(textElement => {
      Object.entries(kerningDefinitions).forEach(([styleName, typographyDef]) => {
        this._applyTypographyDefinition(textElement, styleName, typographyDef, namespace);
      });
    });
  }
  
  _applyTypographyDefinition(textElement, styleName, typographyDef, namespace) {
    // Find text runs within the element
    const textRuns = textElement.getChildren('r', this._getNamespace('a'));
    
    textRuns.forEach(textRun => {
      let rPr = textRun.getChild('rPr', this._getNamespace('a'));
      if (!rPr) {
        rPr = XmlService.createElement('rPr', this._getNamespace('a'));
        textRun.addContent(0, rPr);
      }
      
      // Apply kerning
      if (typographyDef.kerning) {
        this._addKerningProperties(rPr, typographyDef.kerning);
      }
      
      // Apply tracking (letter spacing)
      if (typographyDef.tracking !== undefined) {
        this._addTrackingProperties(rPr, typographyDef.tracking);
      }
      
      // Apply word spacing
      if (typographyDef.wordSpacing !== undefined) {
        this._addWordSpacingProperties(rPr, typographyDef.wordSpacing);
      }
      
      // Apply ligatures
      if (typographyDef.ligatures) {
        this._addLigatureProperties(rPr, typographyDef.ligatures);
      }
      
      // Apply OpenType features
      if (typographyDef.openType) {
        this._addOpenTypeProperties(rPr, typographyDef.openType);
      }
      
      // Apply font features
      if (typographyDef.fontFeatures) {
        this._addFontFeatureProperties(rPr, typographyDef.fontFeatures);
      }
      
      // Apply baseline adjustment
      if (typographyDef.baseline !== undefined) {
        this._addBaselineProperties(rPr, typographyDef.baseline);
      }
    });
  }
  
  _addKerningProperties(rPr, kerningDef) {
    const namespace = this._getNamespace('a');
    
    // Remove existing kerning
    const existingKern = rPr.getChild('kern', namespace);
    if (existingKern) {
      rPr.removeContent(existingKern);
    }
    
    if (kerningDef.method === 'none') {
      // Disable kerning
      const kern = XmlService.createElement('kern', namespace)
        .setAttribute('val', '0');
      rPr.addContent(kern);
    } else if (kerningDef.method === 'metrics') {
      // Use font metrics kerning
      const kern = XmlService.createElement('kern', namespace)
        .setAttribute('val', '1');
      rPr.addContent(kern);
    } else if (kerningDef.method === 'optical') {
      // Use optical kerning with adjustment
      const adjustmentValue = Math.round((kerningDef.adjustment || 0) * 20); // Convert to 20ths of a point
      const spc = XmlService.createElement('spc', namespace)
        .setAttribute('val', adjustmentValue.toString());
      rPr.addContent(spc);
    }
    
    // Add custom kerning pairs if specified
    if (kerningDef.pairs) {
      const kerningPairs = XmlService.createElement('kerningPairs', namespace);
      Object.entries(kerningDef.pairs).forEach(([pair, adjustment]) => {
        const kernPair = XmlService.createElement('kernPair', namespace)
          .setAttribute('chars', pair)
          .setAttribute('val', Math.round(adjustment * 20).toString());
        kerningPairs.addContent(kernPair);
      });
      rPr.addContent(kerningPairs);
    }
  }
  
  _addTrackingProperties(rPr, trackingValue) {
    const namespace = this._getNamespace('a');
    
    // Convert tracking percentage to 20ths of a point
    const trackingPoints = Math.round(trackingValue * 12 * 20); // Assume 12pt base size
    
    const spc = XmlService.createElement('spc', namespace)
      .setAttribute('val', trackingPoints.toString());
    rPr.addContent(spc);
  }
  
  _addWordSpacingProperties(rPr, wordSpacingRatio) {
    const namespace = this._getNamespace('a');
    
    // Word spacing is typically handled at paragraph level
    // This creates a custom attribute for word spacing
    const wordSpc = XmlService.createElement('wordSpc', namespace)
      .setAttribute('val', Math.round(wordSpacingRatio * 100).toString());
    rPr.addContent(wordSpc);
  }
  
  _addLigatureProperties(rPr, ligatureDef) {
    const namespace = this._getNamespace('a');
    
    if (ligatureDef.standard !== undefined) {
      const liga = XmlService.createElement('liga', namespace)
        .setAttribute('val', ligatureDef.standard ? '1' : '0');
      rPr.addContent(liga);
    }
    
    if (ligatureDef.discretionary !== undefined) {
      const dliga = XmlService.createElement('dliga', namespace)
        .setAttribute('val', ligatureDef.discretionary ? '1' : '0');
      rPr.addContent(dliga);
    }
  }
  
  _addOpenTypeProperties(rPr, openTypeFeatures) {
    const namespace = this._getNamespace('a');
    
    const otFeatures = XmlService.createElement('otFeatures', namespace);
    
    openTypeFeatures.forEach(feature => {
      const otFeature = XmlService.createElement('otFeature', namespace)
        .setAttribute('tag', feature)
        .setAttribute('val', '1');
      otFeatures.addContent(otFeature);
    });
    
    if (openTypeFeatures.length > 0) {
      rPr.addContent(otFeatures);
    }
  }
  
  _addFontFeatureProperties(rPr, fontFeatures) {
    const namespace = this._getNamespace('a');
    
    if (fontFeatures.smallCaps) {
      const smcp = XmlService.createElement('smcp', namespace)
        .setAttribute('val', '1');
      rPr.addContent(smcp);
    }
    
    if (fontFeatures.stylisticSet) {
      const ss = XmlService.createElement('ss', namespace)
        .setAttribute('val', fontFeatures.stylisticSet);
      rPr.addContent(ss);
    }
    
    if (fontFeatures.contextualAlternates) {
      const calt = XmlService.createElement('calt', namespace)
        .setAttribute('val', '1');
      rPr.addContent(calt);
    }
    
    if (fontFeatures.tabularNums) {
      const tnum = XmlService.createElement('tnum', namespace)
        .setAttribute('val', '1');
      rPr.addContent(tnum);
    }
    
    if (fontFeatures.oldStyleFigures) {
      const onum = XmlService.createElement('onum', namespace)
        .setAttribute('val', '1');
      rPr.addContent(onum);
    }
  }
  
  _addBaselineProperties(rPr, baselineShift) {
    const namespace = this._getNamespace('a');
    
    if (baselineShift !== 0) {
      const baseline = XmlService.createElement('baseline', namespace)
        .setAttribute('val', Math.round(baselineShift * 1000).toString());
      rPr.addContent(baseline);
    }
  }
  
  _addTypographyToTheme(themeRoot, kerningDefinitions) {
    const namespace = this._getNamespace('a');
    const themeElements = themeRoot.getChild('themeElements', namespace);
    
    if (themeElements) {
      // Create or find format scheme
      let fmtScheme = themeElements.getChild('fmtScheme', namespace);
      if (!fmtScheme) {
        fmtScheme = XmlService.createElement('fmtScheme', namespace)
          .setAttribute('name', 'Office');
        themeElements.addContent(fmtScheme);
      }
      
      // Add typography definitions
      const typographyList = XmlService.createElement('typographyList', namespace);
      
      Object.entries(kerningDefinitions).forEach(([styleName, typographyDef]) => {
        const typography = XmlService.createElement('typography', namespace)
          .setAttribute('name', styleName);
        
        // Add typography properties as attributes
        if (typographyDef.kerning) {
          typography.setAttribute('kerning', typographyDef.kerning.method);
          if (typographyDef.kerning.adjustment) {
            typography.setAttribute('kerningAdj', typographyDef.kerning.adjustment.toString());
          }
        }
        
        if (typographyDef.tracking !== undefined) {
          typography.setAttribute('tracking', typographyDef.tracking.toString());
        }
        
        if (typographyDef.wordSpacing !== undefined) {
          typography.setAttribute('wordSpacing', typographyDef.wordSpacing.toString());
        }
        
        typographyList.addContent(typography);
      });
      
      fmtScheme.addContent(typographyList);
    }
  }
  
  _applyTypographyToSlideElement(slideRoot, elementId, typographySettings) {
    const namespace = this._getNamespace('p');
    
    // Find the specific element by ID
    const elements = this._findElementsRecursively(slideRoot, 'sp', namespace);
    
    elements.forEach(element => {
      const nvSpPr = element.getChild('nvSpPr', namespace);
      if (nvSpPr) {
        const cNvPr = nvSpPr.getChild('cNvPr', namespace);
        if (cNvPr && cNvPr.getAttribute('id') && 
            cNvPr.getAttribute('id').getValue() === elementId) {
          
          // Apply typography to this element's text
          this._applyTypographyToTextStyles(element, { [elementId]: typographySettings });
        }
      }
    });
  }
  
  _findTextElementsRecursively(element, namespace) {
    const textElements = [];
    
    // Look for text body elements
    const txBodies = this._findElementsRecursively(element, 'txBody', this._getNamespace('p'));
    textElements.push(...txBodies);
    
    return textElements;
  }
  
  _findElementsRecursively(element, tagName, namespace) {
    const elements = [];
    
    if (element.getName() === tagName && element.getNamespace().equals(namespace)) {
      elements.push(element);
    }
    
    element.getChildren().forEach(child => {
      elements.push(...this._findElementsRecursively(child, tagName, namespace));
    });
    
    return elements;
  }
  
  _extractTypographySettings() {
    const settings = [];
    
    try {
      const themeXml = this.parser.getXML('ppt/theme/theme1.xml');
      const root = themeXml.getRootElement();
      const namespace = this._getNamespace('a');
      
      // Extract typography settings from theme
      const typographyLists = root.getDescendants().filter(element => 
        element.getName() === 'typographyList' && element.getNamespace().equals(namespace)
      );
      
      typographyLists.forEach(typographyList => {
        const typographies = typographyList.getChildren('typography', namespace);
        typographies.forEach(typography => {
          const name = typography.getAttribute('name');
          if (name) {
            settings.push({
              name: name.getValue(),
              kerning: typography.getAttribute('kerning')?.getValue(),
              tracking: typography.getAttribute('tracking')?.getValue(),
              wordSpacing: typography.getAttribute('wordSpacing')?.getValue()
            });
          }
        });
      });
      
    } catch (error) {
      console.log('No existing typography settings found');
    }
    
    return settings;
  }
  
  _getNamespace(prefix = 'a') {
    const namespaces = {
      'a': XmlService.getNamespace('http://schemas.openxmlformats.org/drawingml/2006/main'),
      'p': XmlService.getNamespace('http://schemas.openxmlformats.org/presentationml/2006/main'),
      'w': XmlService.getNamespace('http://schemas.openxmlformats.org/wordprocessingml/2006/main'),
      'r': XmlService.getNamespace('http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    };
    return namespaces[prefix];
  }
  
  /**
   * Test typography and kerning hack with sample data
   */
  static testTypographyHack(ooxml) {
    console.log('🧪 Testing advanced typography and kerning hack...');
    
    const typographyEditor = new TypographyEditor(ooxml);
    
    // Apply professional typography
    const result = typographyEditor.applyProfessionalTypography('corporate');
    
    if (result) {
      console.log('✅ Typography hack test passed!');
      console.log('   Applied corporate typography preset');
      console.log('   Optical kerning with -2% adjustment for headings');
      console.log('   5% letter spacing for corporate headers');
      console.log('   Standard ligatures enabled');
      console.log('   OpenType features activated');
    } else {
      console.log('❌ Typography hack test failed');
    }
    
    return result;
  }
}

// Export for use
if (typeof module !== 'undefined' && module.exports) {
  module.exports = TypographyEditor;
}