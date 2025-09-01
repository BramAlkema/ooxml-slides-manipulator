# OOXML Hidden Features API Extension - Technical Specification

## Architecture Overview

The OOXML Hidden Features API Extension builds upon the existing OOXMLCore foundation to provide deep access to PowerPoint's advanced features through direct XML manipulation. The architecture follows a layered approach:

```
┌─────────────────────────────────────────────────────────────┐
│                    Hidden Features API                     │
├─────────────────────────────────────────────────────────────┤
│  DrawMLEngine │ ThemeEngine │ TableEngine │ AnimationEngine │
├─────────────────────────────────────────────────────────────┤
│                  OOXML Extension Adapter                   │
├─────────────────────────────────────────────────────────────┤
│              OOXMLCore (Existing Foundation)               │
├─────────────────────────────────────────────────────────────┤
│                    FFlatePPTXService                       │
└─────────────────────────────────────────────────────────────┘
```

## Core Components Design

### 1. DrawML Shape Manipulation Engine

#### 1.1 Advanced 3D Effects System

**Class**: `DrawML3DEngine`

```javascript
class DrawML3DEngine {
  constructor(ooxml) {
    this.ooxml = ooxml;
    this.namespace = {
      a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
      p: 'http://schemas.openxmlformats.org/presentationml/2006/main'
    };
  }

  /**
   * Apply 3D bevel effects to shape
   * @param {string} shapeId - Target shape identifier
   * @param {Object} bevelConfig - 3D bevel configuration
   */
  apply3DBevel(shapeId, bevelConfig) {
    const shapeXml = this._getShapeXml(shapeId);
    const spPr = this._getOrCreateShapeProperties(shapeXml);
    
    // Create 3D properties element
    const sp3d = XmlService.createElement('sp3d', this.namespace.a);
    
    // Top bevel
    if (bevelConfig.topBevel) {
      const bevelT = XmlService.createElement('bevelT', this.namespace.a)
        .setAttribute('w', this._pointsToEMU(bevelConfig.topBevel.width))
        .setAttribute('h', this._pointsToEMU(bevelConfig.topBevel.height))
        .setAttribute('prst', bevelConfig.topBevel.preset || 'relaxedInset');
      sp3d.addContent(bevelT);
    }
    
    // Bottom bevel
    if (bevelConfig.bottomBevel) {
      const bevelB = XmlService.createElement('bevelB', this.namespace.a)
        .setAttribute('w', this._pointsToEMU(bevelConfig.bottomBevel.width))
        .setAttribute('h', this._pointsToEMU(bevelConfig.bottomBevel.height))
        .setAttribute('prst', bevelConfig.bottomBevel.preset || 'relaxedInset');
      sp3d.addContent(bevelB);
    }
    
    // Extrusion
    if (bevelConfig.extrusion) {
      const extrusionClr = XmlService.createElement('extrusionClr', this.namespace.a);
      const srgbClr = XmlService.createElement('srgbClr', this.namespace.a)
        .setAttribute('val', bevelConfig.extrusion.color || '000000');
      extrusionClr.addContent(srgbClr);
      sp3d.addContent(extrusionClr);
      
      sp3d.setAttribute('extrusionH', this._pointsToEMU(bevelConfig.extrusion.height || 0));
    }
    
    // Lighting and material
    if (bevelConfig.lighting) {
      sp3d.setAttribute('rig', bevelConfig.lighting.rig || 'threePt')
           .setAttribute('dir', bevelConfig.lighting.direction || 't');
    }
    
    if (bevelConfig.material) {
      sp3d.setAttribute('prstMaterial', bevelConfig.material || 'matte');
    }
    
    spPr.addContent(sp3d);
    this._updateShapeXml(shapeId, shapeXml);
  }
}
```

#### 1.2 Custom Geometry Engine

**Class**: `CustomGeometryEngine`

```javascript
class CustomGeometryEngine {
  constructor(ooxml) {
    this.ooxml = ooxml;
    this.pathBuilder = new GeometryPathBuilder();
  }

  /**
   * Create custom geometry shape with path commands
   * @param {string} shapeId - Target shape identifier
   * @param {Array} pathCommands - Array of path commands
   */
  createCustomGeometry(shapeId, pathCommands) {
    const shapeXml = this._getShapeXml(shapeId);
    const spPr = this._getOrCreateShapeProperties(shapeXml);
    
    // Remove existing geometry
    const existingGeom = spPr.getChild('custGeom', this.namespace.a) || 
                        spPr.getChild('prstGeom', this.namespace.a);
    if (existingGeom) {
      spPr.removeContent(existingGeom);
    }
    
    // Create custom geometry
    const custGeom = XmlService.createElement('custGeom', this.namespace.a);
    
    // Add adjustment values if specified
    if (pathCommands.adjustments) {
      const avLst = XmlService.createElement('avLst', this.namespace.a);
      pathCommands.adjustments.forEach((adj, index) => {
        const gd = XmlService.createElement('gd', this.namespace.a)
          .setAttribute('name', `adj${index + 1}`)
          .setAttribute('fmla', `val ${adj.value}`);
        avLst.addContent(gd);
      });
      custGeom.addContent(avLst);
    }
    
    // Create rectangle (bounding box)
    const rect = XmlService.createElement('rect', this.namespace.a)
      .setAttribute('l', '0')
      .setAttribute('t', '0')
      .setAttribute('r', 'w')
      .setAttribute('b', 'h');
    custGeom.addContent(rect);
    
    // Create path list
    const pathLst = XmlService.createElement('pathLst', this.namespace.a);
    
    pathCommands.paths.forEach(pathData => {
      const path = XmlService.createElement('path', this.namespace.a)
        .setAttribute('w', pathData.width || '100000')
        .setAttribute('h', pathData.height || '100000');
      
      if (pathData.fill) {
        path.setAttribute('fill', pathData.fill);
      }
      if (pathData.stroke) {
        path.setAttribute('stroke', pathData.stroke);
      }
      
      // Add path commands
      pathData.commands.forEach(cmd => {
        const element = this._createPathCommand(cmd);
        path.addContent(element);
      });
      
      pathLst.addContent(path);
    });
    
    custGeom.addContent(pathLst);
    spPr.addContent(custGeom);
    this._updateShapeXml(shapeId, shapeXml);
  }
  
  _createPathCommand(command) {
    const { type, points } = command;
    let element;
    
    switch (type) {
      case 'moveTo':
        element = XmlService.createElement('moveTo', this.namespace.a);
        const movePt = XmlService.createElement('pt', this.namespace.a)
          .setAttribute('x', points[0].toString())
          .setAttribute('y', points[1].toString());
        element.addContent(movePt);
        break;
        
      case 'lineTo':
        element = XmlService.createElement('lnTo', this.namespace.a);
        const linePt = XmlService.createElement('pt', this.namespace.a)
          .setAttribute('x', points[0].toString())
          .setAttribute('y', points[1].toString());
        element.addContent(linePt);
        break;
        
      case 'cubicBezTo':
        element = XmlService.createElement('cubicBezTo', this.namespace.a);
        for (let i = 0; i < 3; i++) {
          const pt = XmlService.createElement('pt', this.namespace.a)
            .setAttribute('x', points[i * 2].toString())
            .setAttribute('y', points[i * 2 + 1].toString());
          element.addContent(pt);
        }
        break;
        
      case 'close':
        element = XmlService.createElement('close', this.namespace.a);
        break;
    }
    
    return element;
  }
}
```

### 2. Advanced Theme Manipulation System

#### 2.1 SuperTheme Engine Architecture

**Class**: `SuperThemeEngine`

```javascript
class SuperThemeEngine extends BrandwaresExtension {
  constructor(ooxml) {
    super(ooxml);
    this.variantManager = new ThemeVariantManager();
  }

  /**
   * Create Microsoft SuperTheme with multiple design variants
   * @param {Object} superThemeConfig - SuperTheme configuration
   */
  async createSuperTheme(superThemeConfig) {
    console.log(`🎨 Creating SuperTheme: ${superThemeConfig.name}`);
    
    const files = {};
    let relationshipId = 1;
    
    // Create theme variant manager
    const variantManager = this._buildVariantManagerXML(superThemeConfig);
    files['themeVariants/themeVariantManager.xml'] = variantManager;
    
    // Create each design×size variant combination
    for (const design of superThemeConfig.designs) {
      for (const size of superThemeConfig.sizes) {
        const variantFiles = await this._createThemeVariant(
          design, size, relationshipId++
        );
        
        Object.entries(variantFiles).forEach(([path, content]) => {
          files[`themeVariants/variant${relationshipId - 1}/${path}`] = content;
        });
      }
    }
    
    // Add base theme structure
    files['theme/theme1.xml'] = this._createBaseThemeXML(superThemeConfig);
    files['_rels/.rels'] = this._createSuperThemeRelationships(superThemeConfig);
    files['[Content_Types].xml'] = this._createSuperThemeContentTypes(superThemeConfig);
    
    // Package as .thmx file
    const superThemeBlob = await this.ooxml.core.buildFromFiles(files);
    superThemeBlob.setName(`${superThemeConfig.name}.thmx`);
    
    return superThemeBlob;
  }
  
  _buildVariantManagerXML(superThemeConfig) {
    const doc = XmlService.createDocument();
    const root = XmlService.createElement('themeVariantManager', 
      XmlService.getNamespace('http://schemas.microsoft.com/office/thememl/2012/main'));
    
    const variantList = XmlService.createElement('themeVariantLst', 
      XmlService.getNamespace('http://schemas.microsoft.com/office/thememl/2012/main'));
    
    let relationshipId = 1;
    superThemeConfig.designs.forEach(design => {
      superThemeConfig.sizes.forEach(size => {
        const variant = XmlService.createElement('themeVariant',
          XmlService.getNamespace('http://schemas.microsoft.com/office/thememl/2012/main'))
          .setAttribute('name', `${design.name} ${size.name}`)
          .setAttribute('vid', design.vid || this._generateGUID())
          .setAttribute('cx', size.width.toString())
          .setAttribute('cy', size.height.toString())
          .setAttribute('id', `rId${relationshipId++}`,
            XmlService.getNamespace('http://schemas.openxmlformats.org/officeDocument/2006/relationships'));
            
        variantList.addContent(variant);
      });
    });
    
    root.addContent(variantList);
    doc.setRootElement(root);
    return XmlService.getRawFormat().format(doc);
  }
}
```

#### 2.2 Custom Color Scheme Engine

**Class**: `CustomColorSchemeEngine`

```javascript
class CustomColorSchemeEngine extends BrandwaresExtension {
  constructor(ooxml) {
    super(ooxml);
    this.colorHarmonyGenerator = new ColorHarmonyGenerator();
  }

  /**
   * Implement Brandwares custom color technique
   * @param {Object} customColors - Named color definitions
   */
  addBrandwaresCustomColors(customColors) {
    console.log('🎨 Implementing Brandwares custom color XML hack...');
    
    try {
      const themeXml = this.parser.getXML('ppt/theme/theme1.xml');
      const root = themeXml.getRootElement();
      const namespace = XmlService.getNamespace(this.namespace.a);
      
      // Navigate to theme elements
      const themeElements = root.getChild('themeElements', namespace);
      if (!themeElements) {
        throw new Error('Theme elements not found');
      }
      
      // Create or get custom color list
      let custClrLst = themeElements.getChild('custClrLst', namespace);
      if (!custClrLst) {
        custClrLst = XmlService.createElement('custClrLst', namespace);
        
        // Insert after color scheme but before font scheme
        const clrScheme = themeElements.getChild('clrScheme', namespace);
        const fontScheme = themeElements.getChild('fontScheme', namespace);
        
        if (fontScheme) {
          themeElements.addContent(
            themeElements.getContent().indexOf(fontScheme), 
            custClrLst
          );
        } else {
          themeElements.addContent(custClrLst);
        }
      }
      
      // Add each custom color
      Object.entries(customColors).forEach(([name, colorValue]) => {
        this._addCustomColorDefinition(custClrLst, name, colorValue, namespace);
      });
      
      // Update the theme XML
      this.parser.setXML('ppt/theme/theme1.xml', themeXml);
      
      console.log(`✅ Added ${Object.keys(customColors).length} custom colors`);
      return true;
      
    } catch (error) {
      console.error('❌ Brandwares custom color hack failed:', error.message);
      throw error;
    }
  }
  
  _addCustomColorDefinition(custClrLst, name, colorValue, namespace) {
    const custClr = XmlService.createElement('custClr', namespace)
      .setAttribute('name', name);
    
    if (typeof colorValue === 'string') {
      // Simple hex color
      const rgb = colorValue.replace('#', '').toUpperCase();
      const srgbClr = XmlService.createElement('srgbClr', namespace)
        .setAttribute('val', rgb);
      custClr.addContent(srgbClr);
      
    } else if (typeof colorValue === 'object') {
      // Complex color with modifiers
      const rgb = colorValue.rgb.replace('#', '').toUpperCase();
      const srgbClr = XmlService.createElement('srgbClr', namespace)
        .setAttribute('val', rgb);
      
      // Add color modifiers
      if (colorValue.tint !== undefined) {
        const tint = XmlService.createElement('tint', namespace)
          .setAttribute('val', Math.round(colorValue.tint * 100000).toString());
        srgbClr.addContent(tint);
      }
      
      if (colorValue.shade !== undefined) {
        const shade = XmlService.createElement('shade', namespace)
          .setAttribute('val', Math.round(colorValue.shade * 100000).toString());
        srgbClr.addContent(shade);
      }
      
      if (colorValue.alpha !== undefined) {
        const alpha = XmlService.createElement('alpha', namespace)
          .setAttribute('val', Math.round(colorValue.alpha * 100000).toString());
        srgbClr.addContent(alpha);
      }
      
      custClr.addContent(srgbClr);
    }
    
    custClrLst.addContent(custClr);
  }
}
```

### 3. Advanced Table Style System

#### 3.1 Complete Brandwares Table Implementation

Building on the existing `BrandwaresTableStyles` class with enhanced capabilities:

```javascript
class AdvancedTableStyleEngine extends BrandwaresTableStyles {
  constructor(ooxml) {
    super(ooxml);
    this.conditionalFormattingEngine = new ConditionalFormattingEngine();
  }

  /**
   * Create enterprise table style with conditional formatting
   * @param {Object} styleConfig - Enhanced table style configuration
   */
  createEnterpriseTableStyle(styleConfig) {
    console.log(`🏢 Creating enterprise table style: ${styleConfig.name}`);
    
    const enhancedConfig = {
      ...styleConfig,
      parts: this._enhanceTableParts(styleConfig.parts),
      conditionalFormatting: styleConfig.conditionalFormatting || {}
    };
    
    // Create base table style
    const result = super.createCustomTableStyle(enhancedConfig);
    
    // Add conditional formatting rules
    if (styleConfig.conditionalFormatting) {
      this._addConditionalFormatting(styleConfig.name, styleConfig.conditionalFormatting);
    }
    
    // Add data validation styling
    if (styleConfig.dataValidation) {
      this._addDataValidationStyling(styleConfig.name, styleConfig.dataValidation);
    }
    
    return result;
  }
  
  _enhanceTableParts(parts) {
    const enhancedParts = { ...parts };
    
    // Add advanced styling for each table part
    Object.keys(enhancedParts).forEach(partName => {
      const part = enhancedParts[partName];
      
      // Enhanced border styling
      if (part.border && part.border.style) {
        part.border = this._createAdvancedBorder(part.border);
      }
      
      // Enhanced fill with gradients/patterns
      if (part.fill && part.fill.type === 'gradient') {
        part.fill = this._createGradientFill(part.fill);
      } else if (part.fill && part.fill.type === 'pattern') {
        part.fill = this._createPatternFill(part.fill);
      }
      
      // Advanced typography
      if (part.text && part.text.advanced) {
        part.text = this._enhanceTypography(part.text);
      }
    });
    
    return enhancedParts;
  }
  
  _createAdvancedBorder(borderConfig) {
    return {
      ...borderConfig,
      compound: borderConfig.compound || 'sng',
      cap: borderConfig.cap || 'flat',
      join: borderConfig.join || 'round',
      dashStyle: borderConfig.dashStyle || 'solid',
      lineStyle: borderConfig.lineStyle || 'sng'
    };
  }
  
  _createGradientFill(gradientConfig) {
    return {
      type: 'gradient',
      direction: gradientConfig.direction || 'linear',
      angle: gradientConfig.angle || 0,
      stops: gradientConfig.stops.map(stop => ({
        position: stop.position,
        color: stop.color,
        alpha: stop.alpha || 100000
      }))
    };
  }
  
  _addConditionalFormatting(styleName, rules) {
    console.log(`🔧 Adding conditional formatting to: ${styleName}`);
    
    try {
      const tableStylesXml = this._getTableStyles();
      const styleElement = this._findStyleInCollection(tableStylesXml, styleName);
      
      if (styleElement) {
        const condFmtLst = XmlService.createElement('condFmtLst', this.namespace.a);
        
        rules.forEach((rule, index) => {
          const condFmt = this._createConditionalFormatRule(rule, index);
          condFmtLst.addContent(condFmt);
        });
        
        styleElement.addContent(condFmtLst);
        this.parser.setXML('ppt/tableStyles/tableStyle1.xml', tableStylesXml);
      }
      
    } catch (error) {
      console.warn('Could not add conditional formatting:', error.message);
    }
  }
}
```

### 4. Animation & Transition System

#### 4.1 Advanced Animation Engine

**Class**: `AnimationEngine`

```javascript
class AnimationEngine {
  constructor(ooxml) {
    this.ooxml = ooxml;
    this.namespace = {
      p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
      a: 'http://schemas.openxmlformats.org/drawingml/2006/main'
    };
  }

  /**
   * Add advanced animation sequence to slide
   * @param {string} slideId - Target slide identifier
   * @param {Array} animationSequence - Animation commands
   */
  addAnimationSequence(slideId, animationSequence) {
    console.log(`🎬 Adding animation sequence to slide: ${slideId}`);
    
    const slideXml = this.parser.getXML(`ppt/slides/slide${slideId}.xml`);
    const root = slideXml.getRootElement();
    
    // Create timing root
    const timing = XmlService.createElement('timing', this.namespace.p);
    const tnLst = XmlService.createElement('tnLst', this.namespace.p);
    const par = XmlService.createElement('par', this.namespace.p);
    
    // Main sequence
    const cTn = XmlService.createElement('cTn', this.namespace.p)
      .setAttribute('id', '1')
      .setAttribute('dur', 'indefinite')
      .setAttribute('restart', 'never')
      .setAttribute('nodeType', 'tmRoot');
    
    const childTnLst = XmlService.createElement('childTnLst', this.namespace.p);
    const seq = XmlService.createElement('seq', this.namespace.p)
      .setAttribute('concurrent', '1')
      .setAttribute('nextAc', 'seek');
    
    const seqCTn = XmlService.createElement('cTn', this.namespace.p)
      .setAttribute('id', '2')
      .setAttribute('dur', 'indefinite')
      .setAttribute('nodeType', 'mainSeq');
    
    const seqChildTnLst = XmlService.createElement('childTnLst', this.namespace.p);
    
    // Add each animation in sequence
    let animationId = 3;
    animationSequence.forEach((animation, index) => {
      const animElement = this._createAnimationElement(animation, animationId++);
      seqChildTnLst.addContent(animElement);
    });
    
    seqCTn.addContent(seqChildTnLst);
    seq.addContent(seqCTn);
    childTnLst.addContent(seq);
    cTn.addContent(childTnLst);
    par.addContent(cTn);
    tnLst.addContent(par);
    timing.addContent(tnLst);
    
    // Add to slide
    const cSld = root.getChild('cSld', this.namespace.p);
    cSld.addContent(timing);
    
    this.parser.setXML(`ppt/slides/slide${slideId}.xml`, slideXml);
  }
  
  _createAnimationElement(animation, id) {
    const { type, target, properties } = animation;
    
    switch (type) {
      case 'entrance':
        return this._createEntranceAnimation(target, properties, id);
      case 'emphasis':
        return this._createEmphasisAnimation(target, properties, id);
      case 'exit':
        return this._createExitAnimation(target, properties, id);
      case 'motionPath':
        return this._createMotionPathAnimation(target, properties, id);
      default:
        throw new Error(`Unknown animation type: ${type}`);
    }
  }
  
  _createEntranceAnimation(target, properties, id) {
    const par = XmlService.createElement('par', this.namespace.p);
    const cTn = XmlService.createElement('cTn', this.namespace.p)
      .setAttribute('id', id.toString())
      .setAttribute('fill', 'hold');
    
    // Timing properties
    if (properties.duration) {
      cTn.setAttribute('dur', properties.duration.toString());
    }
    
    if (properties.delay) {
      cTn.setAttribute('delay', properties.delay.toString());
    }
    
    // Animation effects
    const childTnLst = XmlService.createElement('childTnLst', this.namespace.p);
    
    if (properties.effect === 'fadeIn') {
      const animEffect = this._createFadeInEffect(target, properties);
      childTnLst.addContent(animEffect);
    } else if (properties.effect === 'slideInFromLeft') {
      const animEffect = this._createSlideInEffect(target, properties, 'left');
      childTnLst.addContent(animEffect);
    }
    // Add more animation effects as needed
    
    cTn.addContent(childTnLst);
    par.addContent(cTn);
    
    return par;
  }
}
```

### 5. Implementation Details

#### 5.1 OOXML Namespace Management

```javascript
const OOXML_NAMESPACES = {
  // Core OOXML namespaces
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  
  // Extended namespaces for advanced features
  p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',
  p15: 'http://schemas.microsoft.com/office/powerpoint/2012/main',
  p16: 'http://schemas.microsoft.com/office/powerpoint/2016/main',
  
  // Theme variant namespaces
  thm15: 'http://schemas.microsoft.com/office/thememl/2012/main',
  
  // Custom metadata namespaces
  cp: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  dc: 'http://purl.org/dc/elements/1.1/',
  dcterms: 'http://purl.org/dc/terms/',
  dcmitype: 'http://purl.org/dc/dcmitype/',
  xsi: 'http://www.w3.org/2001/XMLSchema-instance'
};
```

#### 5.2 Error Handling Strategy

```javascript
class OOXMLHiddenFeaturesError extends Error {
  constructor(code, message, context = {}) {
    super(message);
    this.name = 'OOXMLHiddenFeaturesError';
    this.code = code;
    this.context = context;
    this.timestamp = new Date().toISOString();
  }
}

const ERROR_CODES = {
  DRAWML_001: 'Invalid 3D bevel configuration',
  DRAWML_002: 'Custom geometry path command failed',
  THEME_001: 'SuperTheme variant creation failed',
  THEME_002: 'Custom color definition invalid',
  TABLE_001: 'Table style part not recognized',
  ANIMATION_001: 'Animation sequence validation failed',
  XML_001: 'XML namespace resolution failed',
  XML_002: 'XML element creation failed'
};
```

#### 5.3 Performance Optimization

```javascript
class OOXMLPerformanceOptimizer {
  constructor() {
    this.xmlCache = new Map();
    this.namespaceCache = new Map();
    this.performanceMetrics = {
      xmlParseTime: 0,
      xmlSerializeTime: 0,
      memoryUsage: 0
    };
  }

  /**
   * Cache frequently accessed XML elements
   */
  cacheXMLElement(path, element) {
    this.xmlCache.set(path, {
      element: element,
      timestamp: Date.now(),
      accessCount: (this.xmlCache.get(path)?.accessCount || 0) + 1
    });
  }

  /**
   * Batch XML updates for better performance
   */
  batchXMLUpdates(updates) {
    console.log(`⚡ Batching ${updates.length} XML updates...`);
    
    const startTime = performance.now();
    
    // Group updates by file
    const updatesByFile = new Map();
    updates.forEach(update => {
      if (!updatesByFile.has(update.file)) {
        updatesByFile.set(update.file, []);
      }
      updatesByFile.get(update.file).push(update);
    });
    
    // Apply all updates to each file in one pass
    updatesByFile.forEach((fileUpdates, filePath) => {
      const xml = this.parser.getXML(filePath);
      fileUpdates.forEach(update => {
        this._applyXMLUpdate(xml, update);
      });
      this.parser.setXML(filePath, xml);
    });
    
    const endTime = performance.now();
    this.performanceMetrics.xmlSerializeTime += (endTime - startTime);
    
    console.log(`✅ Batch update completed in ${endTime - startTime}ms`);
  }
}
```

## Integration Architecture

### API Surface Design

The Hidden Features API will extend the existing extension system:

```javascript
// Main API entry point
class OOXMLHiddenFeaturesAPI extends BaseExtension {
  constructor(ooxml) {
    super(ooxml);
    
    this.drawml = new DrawMLEngine(ooxml);
    this.themes = new AdvancedThemeEngine(ooxml);
    this.tables = new AdvancedTableStyleEngine(ooxml);
    this.animations = new AnimationEngine(ooxml);
    this.metadata = new CustomMetadataEngine(ooxml);
  }

  /**
   * Unified method to apply hidden features
   */
  async applyHiddenFeatures(featuresConfig) {
    const results = {};
    
    if (featuresConfig.drawml) {
      results.drawml = await this.drawml.applyFeatures(featuresConfig.drawml);
    }
    
    if (featuresConfig.themes) {
      results.themes = await this.themes.applyFeatures(featuresConfig.themes);
    }
    
    if (featuresConfig.tables) {
      results.tables = await this.tables.applyFeatures(featuresConfig.tables);
    }
    
    if (featuresConfig.animations) {
      results.animations = await this.animations.applyFeatures(featuresConfig.animations);
    }
    
    if (featuresConfig.metadata) {
      results.metadata = await this.metadata.applyFeatures(featuresConfig.metadata);
    }
    
    return results;
  }
}
```

## Testing Strategy

### Unit Testing Framework

```javascript
class HiddenFeaturesTestSuite {
  constructor() {
    this.testResults = [];
  }

  async runAllTests() {
    console.log('🧪 Running OOXML Hidden Features Test Suite...');
    
    await this.testDrawMLFeatures();
    await this.testThemeFeatures();
    await this.testTableFeatures();
    await this.testAnimationFeatures();
    
    this.reportResults();
  }

  async testDrawMLFeatures() {
    const testCases = [
      { name: '3D Bevel Application', test: this.test3DBevel },
      { name: 'Custom Geometry Creation', test: this.testCustomGeometry },
      { name: 'Advanced Fill Effects', test: this.testAdvancedFills }
    ];
    
    for (const testCase of testCases) {
      try {
        const result = await testCase.test.call(this);
        this.testResults.push({ 
          category: 'DrawML', 
          name: testCase.name, 
          result: 'PASS',
          details: result 
        });
      } catch (error) {
        this.testResults.push({ 
          category: 'DrawML', 
          name: testCase.name, 
          result: 'FAIL',
          error: error.message 
        });
      }
    }
  }
}
```

This technical specification provides the detailed implementation approach for exposing PowerPoint's hidden OOXML capabilities through direct XML manipulation, building upon the existing Brandwares techniques in the codebase while extending them to cover the full spectrum of advanced PowerPoint features.