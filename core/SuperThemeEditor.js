/**
 * SuperThemeEditor - Advanced SuperTheme XML Manipulation
 * 
 * Implements Microsoft PowerPoint SuperTheme functionality with tanaikech-style XML hacking.
 * SuperThemes are Microsoft's advanced theme format that support:
 * - Multiple design variants (2-8 designs) in a single theme file
 * - Multiple slide size variants (16:9, 4:3, 16:10, etc.)
 * - Appear directly in PowerPoint's Design tab > Variants Gallery
 * - Prevent graphic distortion when changing slide sizes
 * 
 * Based on Brandwares research: https://www.brandwares.com/bestpractices/2018/02/brandwares-custom-superthemes/
 * 
 * Key capabilities:
 * - Analyze existing SuperTheme structures
 * - Create custom SuperThemes with multiple design/size variants
 * - Extract individual themes from SuperThemes
 * - Generate SuperTheme XML structures programmatically
 * - Support for theme variant management
 */

class SuperThemeEditor {
  
  constructor(ooxml) {
    this.ooxml = ooxml;
    this.parser = ooxml.parser;
    this.variants = new Map();
    this.themeVariantManager = null;
  }
  
  /**
   * Analyze an existing SuperTheme file (.thmx)
   * @param {Blob} superThemeBlob - SuperTheme .thmx file
   * @returns {Object} Analysis of SuperTheme structure
   */
  analyzeSuperTheme(superThemeBlob) {
    console.log('🔍 Analyzing SuperTheme structure...');
    
    try {
      // Extract SuperTheme files using CloudPPTXService
      const extractedFiles = CloudPPTXService.unzipPPTX(superThemeBlob);
      
      // Find theme variant manager
      const variantManagerXml = extractedFiles['themeVariants/themeVariantManager.xml'];
      if (!variantManagerXml) {
        throw new Error('No themeVariantManager.xml found - not a valid SuperTheme');
      }
      
      // Parse variant manager
      const variantManager = this._parseThemeVariantManager(variantManagerXml);
      
      // Analyze each variant
      const variantAnalysis = [];
      variantManager.variants.forEach((variant, index) => {
        const variantPath = `themeVariants/variant${index + 1}`;
        const analysis = this._analyzeVariant(extractedFiles, variantPath, variant);
        variantAnalysis.push(analysis);
      });
      
      const analysis = {
        totalVariants: variantManager.variants.length,
        designVariants: this._countUniqueDesigns(variantManager.variants),
        sizeVariants: this._countUniqueSizes(variantManager.variants),
        variants: variantAnalysis,
        structure: this._analyzeStructure(extractedFiles)
      };
      
      console.log(`✅ SuperTheme analysis complete:`);
      console.log(`   Design variants: ${analysis.designVariants}`);
      console.log(`   Size variants: ${analysis.sizeVariants}`);
      console.log(`   Total combinations: ${analysis.totalVariants}`);
      
      return analysis;
      
    } catch (error) {
      console.error('❌ Failed to analyze SuperTheme:', error.message);
      throw error;
    }
  }
  
  /**
   * Create a custom SuperTheme with multiple design and size variants
   * @param {Object} superThemeDefinition - SuperTheme configuration
   * @returns {Blob} Generated SuperTheme .thmx file
   */
  createSuperTheme(superThemeDefinition) {
    console.log('🎨 Creating custom SuperTheme...');
    console.log(`Creating ${superThemeDefinition.designs.length} design variants`);
    console.log(`With ${superThemeDefinition.sizes.length} size variants each`);
    
    try {
      const files = {};
      
      // Create theme variant manager
      const variantManager = this._createThemeVariantManager(superThemeDefinition);
      files['themeVariants/themeVariantManager.xml'] = variantManager;
      
      // Create each variant
      let variantIndex = 1;
      superThemeDefinition.designs.forEach((design, designIndex) => {
        superThemeDefinition.sizes.forEach((size, sizeIndex) => {
          const variantFiles = this._createVariant(design, size, variantIndex, designIndex, sizeIndex);
          Object.entries(variantFiles).forEach(([path, content]) => {
            files[`themeVariants/variant${variantIndex}/${path}`] = content;
          });
          variantIndex++;
        });
      });
      
      // Create base theme structure
      const baseThemeFiles = this._createBaseTheme(superThemeDefinition);
      Object.entries(baseThemeFiles).forEach(([path, content]) => {
        files[path] = content;
      });
      
      // Create content types and relationships
      files['[Content_Types].xml'] = this._createSuperThemeContentTypes();
      files['_rels/.rels'] = this._createSuperThemeRelationships();
      
      // Package into .thmx file
      const superThemeBlob = CloudPPTXService.zipPPTX(files);
      
      console.log(`✅ SuperTheme created with ${variantIndex - 1} variants`);
      return superThemeBlob;
      
    } catch (error) {
      console.error('❌ Failed to create SuperTheme:', error.message);
      throw error;
    }
  }
  
  /**
   * Extract individual themes from a SuperTheme
   * @param {Blob} superThemeBlob - SuperTheme .thmx file
   * @param {number} variantIndex - Which variant to extract (1-based)
   * @returns {Object} Extracted theme data
   */
  extractThemeVariant(superThemeBlob, variantIndex) {
    console.log(`📤 Extracting theme variant ${variantIndex} from SuperTheme...`);
    
    try {
      const extractedFiles = CloudPPTXService.unzipPPTX(superThemeBlob);
      const variantPath = `themeVariants/variant${variantIndex}`;
      
      // Extract theme files for this variant
      const themeFiles = {};
      Object.keys(extractedFiles).forEach(path => {
        if (path.startsWith(variantPath + '/theme/')) {
          const relativePath = path.replace(variantPath + '/theme/', '');
          themeFiles[relativePath] = extractedFiles[path];
        }
      });
      
      if (Object.keys(themeFiles).length === 0) {
        throw new Error(`Variant ${variantIndex} not found in SuperTheme`);
      }
      
      console.log(`✅ Extracted ${Object.keys(themeFiles).length} theme files`);
      return themeFiles;
      
    } catch (error) {
      console.error('❌ Failed to extract theme variant:', error.message);
      throw error;
    }
  }
  
  /**
   * Convert regular PPTX themes to SuperTheme format
   * @param {Array<Object>} themeDefinitions - Array of theme configurations
   * @returns {Blob} Generated SuperTheme
   */
  convertToSuperTheme(themeDefinitions) {
    console.log('🔄 Converting themes to SuperTheme format...');
    
    const superThemeDefinition = {
      name: 'Custom SuperTheme',
      designs: [],
      sizes: [
        { name: '16:9', width: 12192000, height: 6858000 },
        { name: '4:3', width: 9144000, height: 6858000 },
        { name: '16:10', width: 10972800, height: 6858000 }
      ]
    };
    
    // Convert each theme definition to design variant
    themeDefinitions.forEach((themeDef, index) => {
      superThemeDefinition.designs.push({
        name: themeDef.name || `Design ${index + 1}`,
        vid: this._generateGUID(),
        colorScheme: themeDef.colorScheme,
        fontScheme: themeDef.fontScheme,
        effectScheme: themeDef.effectScheme || this._getDefaultEffectScheme()
      });
    });
    
    return this.createSuperTheme(superThemeDefinition);
  }
  
  /**
   * Create responsive SuperTheme that adapts to different screen sizes
   * @param {Object} responsiveConfig - Responsive configuration
   * @returns {Blob} Responsive SuperTheme
   */
  createResponsiveSuperTheme(responsiveConfig) {
    console.log('📱 Creating responsive SuperTheme...');
    
    const superThemeDefinition = {
      name: 'Responsive SuperTheme',
      designs: responsiveConfig.designs,
      sizes: [
        { name: 'Mobile 16:9', width: 12192000, height: 6858000 },
        { name: 'Tablet 4:3', width: 9144000, height: 6858000 },
        { name: 'Desktop 16:10', width: 10972800, height: 6858000 },
        { name: 'Ultrawide 21:9', width: 16192000, height: 6858000 }
      ]
    };
    
    // Add responsive adjustments for each size
    superThemeDefinition.designs.forEach(design => {
      design.responsiveAdjustments = {
        mobile: { fontSize: 0.8, spacing: 0.9 },
        tablet: { fontSize: 1.0, spacing: 1.0 },
        desktop: { fontSize: 1.1, spacing: 1.1 },
        ultrawide: { fontSize: 1.2, spacing: 1.2 }
      };
    });
    
    return this.createSuperTheme(superThemeDefinition);
  }
  
  // Private helper methods
  
  _parseThemeVariantManager(xml) {
    const doc = XmlService.parse(xml);
    const root = doc.getRootElement();
    const namespace = XmlService.getNamespace('http://schemas.microsoft.com/office/thememl/2012/main');
    
    const variantList = root.getChild('themeVariantLst', namespace);
    const variants = [];
    
    if (variantList) {
      const variantElements = variantList.getChildren('themeVariant', namespace);
      variantElements.forEach(element => {
        variants.push({
          name: element.getAttribute('name').getValue(),
          vid: element.getAttribute('vid').getValue(),
          width: parseInt(element.getAttribute('cx').getValue()),
          height: parseInt(element.getAttribute('cy').getValue()),
          relationshipId: element.getAttribute('id', 
            XmlService.getNamespace('http://schemas.openxmlformats.org/officeDocument/2006/relationships')).getValue()
        });
      });
    }
    
    return { variants };
  }
  
  _analyzeVariant(extractedFiles, variantPath, variant) {
    const themeXmlPath = `${variantPath}/theme/theme/theme1.xml`;
    const themeXml = extractedFiles[themeXmlPath];
    
    if (!themeXml) {
      return { error: 'Theme XML not found' };
    }
    
    // Parse theme XML to extract color scheme, fonts, etc.
    const doc = XmlService.parse(themeXml);
    const root = doc.getRootElement();
    const namespace = XmlService.getNamespace('http://schemas.openxmlformats.org/drawingml/2006/main');
    
    const analysis = {
      name: variant.name,
      size: `${variant.width}x${variant.height}`,
      vid: variant.vid,
      colorScheme: this._extractColorScheme(root, namespace),
      fontScheme: this._extractFontScheme(root, namespace),
      files: Object.keys(extractedFiles).filter(path => path.startsWith(variantPath)).length
    };
    
    return analysis;
  }
  
  _extractColorScheme(themeRoot, namespace) {
    const themeElements = themeRoot.getChild('themeElements', namespace);
    if (!themeElements) return {};
    
    const colorScheme = themeElements.getChild('clrScheme', namespace);
    if (!colorScheme) return {};
    
    const colors = {};
    ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].forEach(colorName => {
      const colorElement = colorScheme.getChild(colorName, namespace);
      if (colorElement) {
        const srgbElement = colorElement.getChild('srgbClr', namespace);
        if (srgbElement) {
          colors[colorName] = srgbElement.getAttribute('val').getValue();
        }
      }
    });
    
    return colors;
  }
  
  _extractFontScheme(themeRoot, namespace) {
    const themeElements = themeRoot.getChild('themeElements', namespace);
    if (!themeElements) return {};
    
    const fontScheme = themeElements.getChild('fontScheme', namespace);
    if (!fontScheme) return {};
    
    const fonts = {};
    const majorFont = fontScheme.getChild('majorFont', namespace);
    const minorFont = fontScheme.getChild('minorFont', namespace);
    
    if (majorFont) {
      const latin = majorFont.getChild('latin', namespace);
      if (latin) fonts.majorFont = latin.getAttribute('typeface').getValue();
    }
    
    if (minorFont) {
      const latin = minorFont.getChild('latin', namespace);
      if (latin) fonts.minorFont = latin.getAttribute('typeface').getValue();
    }
    
    return fonts;
  }
  
  _countUniqueDesigns(variants) {
    const uniqueVids = new Set(variants.map(v => v.vid));
    return uniqueVids.size;
  }
  
  _countUniqueSizes(variants) {
    const uniqueSizes = new Set(variants.map(v => `${v.width}x${v.height}`));
    return uniqueSizes.size;
  }
  
  _analyzeStructure(extractedFiles) {
    const structure = {
      hasThemeVariants: false,
      hasBaseTheme: false,
      totalFiles: Object.keys(extractedFiles).length,
      fileTypes: {}
    };
    
    Object.keys(extractedFiles).forEach(path => {
      if (path.startsWith('themeVariants/')) {
        structure.hasThemeVariants = true;
      }
      if (path.startsWith('theme/')) {
        structure.hasBaseTheme = true;
      }
      
      const extension = path.split('.').pop();
      structure.fileTypes[extension] = (structure.fileTypes[extension] || 0) + 1;
    });
    
    return structure;
  }
  
  _createThemeVariantManager(superThemeDefinition) {
    let xml = '<?xml version="1.0" encoding="utf-8" standalone="yes"?>\n';
    xml += '<t:themeVariantManager xmlns:t="http://schemas.microsoft.com/office/thememl/2012/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n';
    xml += '  <t:themeVariantLst>\n';
    
    let relationshipId = 1;
    superThemeDefinition.designs.forEach((design, designIndex) => {
      superThemeDefinition.sizes.forEach((size, sizeIndex) => {
        xml += `    <t:themeVariant name="${design.name} ${size.name}" vid="${design.vid}" cx="${size.width}" cy="${size.height}" r:id="rId${relationshipId}" />\n`;
        relationshipId++;
      });
    });
    
    xml += '  </t:themeVariantLst>\n';
    xml += '</t:themeVariantManager>';
    
    return xml;
  }
  
  _createVariant(design, size, variantIndex, designIndex, sizeIndex) {
    const files = {};
    
    // Create theme XML for this variant
    files['theme/theme/theme1.xml'] = this._createThemeXml(design, size);
    files['theme/presentation.xml'] = this._createPresentationXml(size);
    
    // Create slide master and layouts
    files['theme/slideMasters/slideMaster1.xml'] = this._createSlideMaster(design, size);
    files['theme/slideMasters/_rels/slideMaster1.xml.rels'] = this._createSlideMasterRels();
    
    // Create standard slide layouts
    for (let i = 1; i <= 12; i++) {
      files[`theme/slideLayouts/slideLayout${i}.xml`] = this._createSlideLayout(i, design, size);
      files[`theme/slideLayouts/_rels/slideLayout${i}.xml.rels`] = this._createSlideLayoutRels();
    }
    
    // Create relationships and properties
    files['theme/_rels/presentation.xml.rels'] = this._createThemePresentationRels();
    files['_rels/.rels'] = this._createVariantRels();
    files['docProps/thumbnail.jpeg'] = this._createThumbnail(design);
    
    return files;
  }
  
  _createThemeXml(design, size) {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${design.name} ${size.name}">\n`;
    xml += '  <a:themeElements>\n';
    
    // Color scheme
    xml += `    <a:clrScheme name="${design.name} Colors">\n`;
    Object.entries(design.colorScheme).forEach(([colorName, colorValue]) => {
      xml += `      <a:${colorName}>\n`;
      xml += `        <a:srgbClr val="${colorValue}"/>\n`;
      xml += `      </a:${colorName}>\n`;
    });
    xml += '    </a:clrScheme>\n';
    
    // Font scheme
    xml += `    <a:fontScheme name="${design.name} Fonts">\n`;
    xml += '      <a:majorFont>\n';
    xml += `        <a:latin typeface="${design.fontScheme.majorFont}"/>\n`;
    xml += '      </a:majorFont>\n';
    xml += '      <a:minorFont>\n';
    xml += `        <a:latin typeface="${design.fontScheme.minorFont}"/>\n`;
    xml += '      </a:minorFont>\n';
    xml += '    </a:fontScheme>\n';
    
    // Effect scheme
    xml += '    <a:fmtScheme name="Office">\n';
    xml += '      <!-- Standard effect scheme -->\n';
    xml += '    </a:fmtScheme>\n';
    
    xml += '  </a:themeElements>\n';
    xml += '</a:theme>';
    
    return xml;
  }
  
  _createPresentationXml(size) {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">\n';
    xml += `  <p:sldSz cx="${size.width}" cy="${size.height}"/>\n`;
    xml += '</p:presentation>';
    return xml;
  }
  
  _createSuperThemeContentTypes() {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n';
    xml += '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n';
    xml += '  <Default Extension="xml" ContentType="application/xml"/>\n';
    xml += '  <Default Extension="jpeg" ContentType="image/jpeg"/>\n';
    xml += '  <Default Extension="emf" ContentType="image/x-emf"/>\n';
    xml += '  <Override PartName="/themeVariants/themeVariantManager.xml" ContentType="application/vnd.ms-office.themevariantmanager+xml"/>\n';
    xml += '</Types>';
    return xml;
  }
  
  _createSuperThemeRelationships() {
    let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n';
    xml += '  <Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/themevariantmanager" Target="themeVariants/themeVariantManager.xml"/>\n';
    xml += '</Relationships>';
    return xml;
  }
  
  _generateGUID() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
      const r = Math.random() * 16 | 0;
      const v = c === 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16).toUpperCase();
    });
  }
  
  _getDefaultEffectScheme() {
    return {
      name: 'Office',
      fillStyleLst: [],
      lnStyleLst: [],
      effectStyleLst: []
    };
  }
  
  /**
   * Test SuperTheme creation with sample data
   */
  static testSuperThemeCreation(ooxml) {
    console.log('🧪 Testing SuperTheme creation...');
    
    const editor = new SuperThemeEditor(ooxml);
    
    // Create sample SuperTheme with multiple designs and sizes
    const superThemeDefinition = {
      name: 'Test SuperTheme',
      designs: [
        {
          name: 'Corporate',
          vid: '{0E01D92C-1466-42F6-BFE8-5BD5C0EECBBA}',
          colorScheme: {
            dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6',
            accent1: '4472C4', accent2: 'E70013', accent3: 'A5A5A5',
            accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47'
          },
          fontScheme: { majorFont: 'Calibri', minorFont: 'Calibri' }
        },
        {
          name: 'Creative',
          vid: '{E9A03591-8CE2-4DA8-9D87-195F6010481A}',
          colorScheme: {
            dk1: '000000', lt1: 'FFFFFF', dk2: '2F4F4F', lt2: 'F5F5F5',
            accent1: 'FF6B35', accent2: 'F7931E', accent3: 'FFD23F',
            accent4: '06FFA5', accent5: '118AB2', accent6: '073B4C'
          },
          fontScheme: { majorFont: 'Montserrat', minorFont: 'Open Sans' }
        }
      ],
      sizes: [
        { name: '16:9', width: 12192000, height: 6858000 },
        { name: '4:3', width: 9144000, height: 6858000 },
        { name: '16:10', width: 10972800, height: 6858000 }
      ]
    };
    
    try {
      const superThemeBlob = editor.createSuperTheme(superThemeDefinition);
      
      console.log('✅ SuperTheme creation test passed!');
      console.log('   Created 2 design variants × 3 size variants = 6 total variants');
      console.log('   Corporate and Creative designs');
      console.log('   16:9, 4:3, and 16:10 aspect ratios');
      
      return superThemeBlob;
      
    } catch (error) {
      console.log('❌ SuperTheme creation test failed:', error.message);
      return null;
    }
  }
}

// Export for use
if (typeof module !== 'undefined' && module.exports) {
  module.exports = SuperThemeEditor;
}