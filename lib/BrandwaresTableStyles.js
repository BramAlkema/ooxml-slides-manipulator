/**
 * BrandwaresTableStyles - Advanced XML Hacking for Custom Table Styles
 * 
 * Based on: https://www.brandwares.com/bestpractices/2015/07/xml-hacking-custom-table-styles/
 * 
 * PURPOSE:
 * Implements the complete Brandwares XML hacking technique for creating custom
 * table styles that go beyond PowerPoint's built-in capabilities. This includes:
 * - Custom table style definitions in tableStyles.xml
 * - Advanced color schemes and gradients
 * - Conditional formatting for different table parts
 * - Brand-compliant table templates
 * - Accessibility-aware styling
 * 
 * ARCHITECTURE:
 * - Direct XML manipulation of tableStyles directory
 * - Theme integration for consistent branding
 * - Support for complex table style hierarchies
 * - Validation and error handling
 * 
 * AI CONTEXT:
 * This implements the advanced table styling techniques from Brandwares that
 * require direct OOXML manipulation. Use this for enterprise-grade table
 * styling that can't be achieved through standard PowerPoint APIs.
 */

/**
 * Table style part definitions (from OOXML spec)
 */
const BRANDWARES_TABLE_STYLE_PARTS = {
  WHOLE_TABLE: 'wholeTbl',
  FIRST_ROW: 'firstRow',
  LAST_ROW: 'lastRow', 
  FIRST_COLUMN: 'firstCol',
  LAST_COLUMN: 'lastCol',
  BAND_1_HORIZONTAL: 'band1H',
  BAND_2_HORIZONTAL: 'band2H',
  BAND_1_VERTICAL: 'band1V',
  BAND_2_VERTICAL: 'band2V',
  NE_CELL: 'neCell',
  NW_CELL: 'nwCell', 
  SE_CELL: 'seCell',
  SW_CELL: 'swCell'
};

/**
 * Advanced color schemes for enterprise tables
 */
const ENTERPRISE_COLOR_SCHEMES = {
  corporate: {
    name: 'Corporate Blue',
    primary: '2E5E8C',
    secondary: 'F0F4F8',
    accent: '5B9BD5',
    text: '2C3E50',
    textLight: 'FFFFFF'
  },
  finance: {
    name: 'Financial Green',
    primary: '1E7E34', 
    secondary: 'F1F8E9',
    accent: '4CAF50',
    text: '1B5E20',
    textLight: 'FFFFFF'
  },
  technology: {
    name: 'Tech Purple',
    primary: '6A1B9A',
    secondary: 'F3E5F5', 
    accent: '9C27B0',
    text: '4A148C',
    textLight: 'FFFFFF'
  },
  medical: {
    name: 'Medical Red',
    primary: 'C62828',
    secondary: 'FFEBEE',
    accent: 'F44336', 
    text: 'B71C1C',
    textLight: 'FFFFFF'
  }
};

/**
 * Brandwares Custom Table Styles Engine
 */
class BrandwaresTableStyles {
  
  constructor(ooxml) {
    this.ooxml = ooxml;
    this.parser = ooxml.parser;
    this.namespace = {
      a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
      p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
      r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    };
  }
  
  /**
   * Create a complete custom table style using Brandwares XML hacking
   */
  createCustomTableStyle(styleConfig) {
    console.log(`🎨 Creating custom table style: ${styleConfig.name}`);
    
    try {
      // Validate the style configuration
      this._validateStyleConfig(styleConfig);
      
      // Create the table style XML structure
      const tableStyleXml = this._buildTableStyleXML(styleConfig);
      
      // Add to table styles collection
      const tableStylesXml = this._getOrCreateTableStyles();
      this._addStyleToCollection(tableStylesXml, tableStyleXml);
      
      // Update content types and relationships
      this._updateContentTypes(styleConfig.name);
      this._updateRelationships(styleConfig.name);
      
      // Apply to theme if requested
      if (styleConfig.setAsDefault) {
        this._setAsDefaultTableStyle(styleConfig.name);
      }
      
      console.log(`✅ Custom table style '${styleConfig.name}' created successfully`);
      return {
        success: true,
        styleName: styleConfig.name,
        styleId: this._generateStyleId(styleConfig.name)
      };
      
    } catch (error) {
      const ooxmlError = OOXMLErrorFactory.createAppError(
        OOXMLErrorCodes.A027_SHAPE_ERROR,
        `Failed to create custom table style: ${error.message}`,
        { styleName: styleConfig.name, originalError: error.message }
      );
      
      OOXMLErrorLogger.logError(ooxmlError, { operation: 'create_custom_table_style' });
      throw ooxmlError;
    }
  }
  
  /**
   * Apply enterprise color scheme to table style
   */
  applyEnterpriseColorScheme(schemeName, customizations = {}) {
    console.log(`🎨 Applying enterprise color scheme: ${schemeName}`);
    
    const scheme = ENTERPRISE_COLOR_SCHEMES[schemeName];
    if (!scheme) {
      throw OOXMLErrorFactory.createAppError(
        OOXMLErrorCodes.A026_TEMPLATE_ERROR,
        `Unknown color scheme: ${schemeName}`,
        { availableSchemes: Object.keys(ENTERPRISE_COLOR_SCHEMES) }
      );
    }
    
    const styleConfig = {
      name: `Enterprise_${scheme.name.replace(/\s+/g, '_')}`,
      description: `Enterprise table style with ${scheme.name} color scheme`,
      colorScheme: scheme,
      ...customizations,
      
      // Define all table parts with enterprise styling
      parts: {
        [BRANDWARES_TABLE_STYLE_PARTS.WHOLE_TABLE]: {
          fill: { type: 'solid', color: scheme.secondary },
          border: { color: scheme.primary, width: 1 },
          text: { 
            font: { family: 'Segoe UI', size: 11 },
            color: scheme.text 
          }
        },
        
        [BRANDWARES_TABLE_STYLE_PARTS.FIRST_ROW]: {
          fill: { type: 'solid', color: scheme.primary },
          border: { color: scheme.primary, width: 2 },
          text: { 
            font: { family: 'Segoe UI', size: 12, bold: true },
            color: scheme.textLight
          }
        },
        
        [BRANDWARES_TABLE_STYLE_PARTS.BAND_1_HORIZONTAL]: {
          fill: { type: 'solid', color: scheme.secondary },
          text: { 
            font: { family: 'Segoe UI', size: 11 },
            color: scheme.text 
          }
        },
        
        [BRANDWARES_TABLE_STYLE_PARTS.BAND_2_HORIZONTAL]: {
          fill: { type: 'solid', color: 'FFFFFF' },
          text: { 
            font: { family: 'Segoe UI', size: 11 },
            color: scheme.text 
          }
        }
      }
    };
    
    return this.createCustomTableStyle(styleConfig);
  }
  
  /**
   * Create accessibility-compliant table style
   */
  createAccessibleTableStyle(options = {}) {
    console.log('♿ Creating accessibility-compliant table style...');
    
    const accessibleConfig = {
      name: 'AccessibleTable',
      description: 'WCAG 2.1 AA compliant table style',
      setAsDefault: options.setAsDefault || false,
      
      parts: {
        [BRANDWARES_TABLE_STYLE_PARTS.WHOLE_TABLE]: {
          fill: { type: 'solid', color: 'FFFFFF' },
          border: { color: '000000', width: 2 },
          text: { 
            font: { family: 'Arial', size: 12 },
            color: '000000' // High contrast
          }
        },
        
        [BRANDWARES_TABLE_STYLE_PARTS.FIRST_ROW]: {
          fill: { type: 'solid', color: '000000' },
          border: { color: '000000', width: 3 },
          text: { 
            font: { family: 'Arial', size: 14, bold: true },
            color: 'FFFFFF' // High contrast
          }
        },
        
        [BRANDWARES_TABLE_STYLE_PARTS.BAND_1_HORIZONTAL]: {
          fill: { type: 'solid', color: 'F5F5F5' }, // Light gray for alternating
          text: { 
            font: { family: 'Arial', size: 12 },
            color: '000000'
          }
        }
      },
      
      // Accessibility enhancements
      accessibility: {
        highContrast: true,
        minimumFontSize: 12,
        boldHeaders: true,
        clearBorders: true
      }
    };
    
    return this.createCustomTableStyle(accessibleConfig);
  }
  
  /**
   * Create modern material design table style
   */
  createMaterialDesignTableStyle(materialTheme = 'blue') {
    console.log(`🎨 Creating Material Design table style: ${materialTheme}`);
    
    const materialColors = {
      blue: { primary: '1976D2', secondary: 'E3F2FD', accent: '42A5F5' },
      green: { primary: '388E3C', secondary: 'E8F5E8', accent: '66BB6A' },
      purple: { primary: '7B1FA2', secondary: 'F3E5F5', accent: 'BA68C8' },
      orange: { primary: 'F57C00', secondary: 'FFF3E0', accent: 'FFB74D' }
    };
    
    const colors = materialColors[materialTheme] || materialColors.blue;
    
    const materialConfig = {
      name: `Material_${materialTheme}`,
      description: `Material Design table style in ${materialTheme}`,
      
      parts: {
        [BRANDWARES_TABLE_STYLE_PARTS.WHOLE_TABLE]: {
          fill: { type: 'solid', color: 'FFFFFF' },
          border: { color: 'E0E0E0', width: 1 },
          text: { 
            font: { family: 'Roboto', size: 14 },
            color: '212121'
          },
          padding: { top: 12, bottom: 12, left: 16, right: 16 }
        },
        
        [BRANDWARES_TABLE_STYLE_PARTS.FIRST_ROW]: {
          fill: { type: 'solid', color: colors.primary },
          border: { color: colors.primary, width: 0 },
          text: { 
            font: { family: 'Roboto', size: 14, bold: true },
            color: 'FFFFFF'
          },
          padding: { top: 16, bottom: 16, left: 16, right: 16 }
        },
        
        [BRANDWARES_TABLE_STYLE_PARTS.BAND_1_HORIZONTAL]: {
          fill: { type: 'solid', color: colors.secondary },
          text: { 
            font: { family: 'Roboto', size: 14 },
            color: '424242'
          }
        }
      },
      
      // Material Design enhancements
      materialDesign: {
        elevation: 2,
        borderRadius: 4,
        rippleEffect: true
      }
    };
    
    return this.createCustomTableStyle(materialConfig);
  }
  
  /**
   * Export table style for reuse
   */
  exportTableStyle(styleName) {
    console.log(`📤 Exporting table style: ${styleName}`);
    
    try {
      const tableStylesXml = this._getTableStyles();
      const styleElement = this._findStyleInCollection(tableStylesXml, styleName);
      
      if (!styleElement) {
        throw new Error(`Table style '${styleName}' not found`);
      }
      
      const exportData = {
        name: styleName,
        xml: XmlService.getRawFormat().format(styleElement),
        timestamp: new Date().toISOString(),
        version: '1.0.0'
      };
      
      console.log(`✅ Table style '${styleName}' exported successfully`);
      return exportData;
      
    } catch (error) {
      throw OOXMLErrorFactory.createAppError(
        OOXMLErrorCodes.A027_SHAPE_ERROR,
        `Failed to export table style: ${error.message}`,
        { styleName }
      );
    }
  }
  
  /**
   * Import table style from export data
   */
  importTableStyle(exportData) {
    console.log(`📥 Importing table style: ${exportData.name}`);
    
    try {
      // Parse the XML
      const styleElement = XmlService.parse(exportData.xml).getRootElement();
      
      // Add to collection
      const tableStylesXml = this._getOrCreateTableStyles();
      this._addStyleToCollection(tableStylesXml, styleElement);
      
      // Update content types
      this._updateContentTypes(exportData.name);
      
      console.log(`✅ Table style '${exportData.name}' imported successfully`);
      return {
        success: true,
        styleName: exportData.name
      };
      
    } catch (error) {
      throw OOXMLErrorFactory.createAppError(
        OOXMLErrorCodes.A027_SHAPE_ERROR,
        `Failed to import table style: ${error.message}`,
        { styleName: exportData.name }
      );
    }
  }
  
  // PRIVATE METHODS - XML MANIPULATION
  
  /**
   * Build complete table style XML using Brandwares techniques
   * @private
   */
  _buildTableStyleXML(styleConfig) {
    const doc = XmlService.createDocument();
    const root = XmlService.createElement('tblStyle', XmlService.getNamespace(this.namespace.a))
      .setAttribute('styleId', this._generateStyleId(styleConfig.name))
      .setAttribute('styleName', styleConfig.name);
    
    if (styleConfig.description) {
      root.setAttribute('description', styleConfig.description);
    }
    
    // Add all table parts
    for (const [partName, partConfig] of Object.entries(styleConfig.parts)) {
      const partElement = this._createTableStylePart(partName, partConfig);
      root.addContent(partElement);
    }
    
    doc.setRootElement(root);
    return doc;
  }
  
  /**
   * Create individual table style part XML
   * @private
   */
  _createTableStylePart(partName, partConfig) {
    const namespace = XmlService.getNamespace(this.namespace.a);
    const partElement = XmlService.createElement('tblStylePr', namespace)
      .setAttribute('type', partName);
    
    // Add fill properties
    if (partConfig.fill) {
      const fillElement = this._createFillXML(partConfig.fill);
      partElement.addContent(fillElement);
    }
    
    // Add border properties  
    if (partConfig.border) {
      const borderElement = this._createBorderXML(partConfig.border);
      partElement.addContent(borderElement);
    }
    
    // Add text properties
    if (partConfig.text) {
      const textElement = this._createTextPropertiesXML(partConfig.text);
      partElement.addContent(textElement);
    }
    
    // Add cell properties (margins/padding)
    if (partConfig.padding) {
      const cellElement = this._createCellPropertiesXML(partConfig.padding);
      partElement.addContent(cellElement);
    }
    
    return partElement;
  }
  
  /**
   * Create fill XML (solid, gradient, pattern)
   * @private
   */
  _createFillXML(fillConfig) {
    const namespace = XmlService.getNamespace(this.namespace.a);
    
    if (fillConfig.type === 'solid') {
      const solidFill = XmlService.createElement('solidFill', namespace);
      const srgbClr = XmlService.createElement('srgbClr', namespace)
        .setAttribute('val', fillConfig.color);
      solidFill.addContent(srgbClr);
      return solidFill;
    }
    
    if (fillConfig.type === 'gradient') {
      const gradFill = XmlService.createElement('gradFill', namespace);
      const gsLst = XmlService.createElement('gsLst', namespace);
      
      fillConfig.stops.forEach((stop, index) => {
        const gs = XmlService.createElement('gs', namespace)
          .setAttribute('pos', (stop.position * 1000).toString());
        const srgbClr = XmlService.createElement('srgbClr', namespace)
          .setAttribute('val', stop.color);
        gs.addContent(srgbClr);
        gsLst.addContent(gs);
      });
      
      gradFill.addContent(gsLst);
      return gradFill;
    }
    
    // Default to no fill
    return XmlService.createElement('noFill', namespace);
  }
  
  /**
   * Create border XML
   * @private
   */
  _createBorderXML(borderConfig) {
    const namespace = XmlService.getNamespace(this.namespace.a);
    const tcBdr = XmlService.createElement('tcBdr', namespace);
    
    ['left', 'right', 'top', 'bottom'].forEach(side => {
      const borderSide = XmlService.createElement(side, namespace)
        .setAttribute('w', ((borderConfig.width || 1) * 12700).toString())
        .setAttribute('cap', 'flat')
        .setAttribute('cmpd', 'sng')
        .setAttribute('algn', 'ctr');
      
      const solidFill = XmlService.createElement('solidFill', namespace);
      const srgbClr = XmlService.createElement('srgbClr', namespace)
        .setAttribute('val', borderConfig.color || '000000');
      solidFill.addContent(srgbClr);
      borderSide.addContent(solidFill);
      
      tcBdr.addContent(borderSide);
    });
    
    return tcBdr;
  }
  
  /**
   * Create text properties XML
   * @private
   */
  _createTextPropertiesXML(textConfig) {
    const namespace = XmlService.getNamespace(this.namespace.a);
    const tcTxStyle = XmlService.createElement('tcTxStyle', namespace);
    
    if (textConfig.font) {
      const fontRef = XmlService.createElement('fontRef', namespace)
        .setAttribute('idx', 'minor');
      
      if (textConfig.font.family) {
        const latin = XmlService.createElement('latin', namespace)
          .setAttribute('typeface', textConfig.font.family);
        fontRef.addContent(latin);
      }
      
      tcTxStyle.addContent(fontRef);
      
      // Text properties
      const defRPr = XmlService.createElement('defRPr', namespace);
      
      if (textConfig.font.size) {
        defRPr.setAttribute('sz', (textConfig.font.size * 100).toString());
      }
      
      if (textConfig.font.bold) {
        defRPr.setAttribute('b', '1');
      }
      
      if (textConfig.font.italic) {
        defRPr.setAttribute('i', '1');
      }
      
      if (textConfig.color) {
        const solidFill = XmlService.createElement('solidFill', namespace);
        const srgbClr = XmlService.createElement('srgbClr', namespace)
          .setAttribute('val', textConfig.color);
        solidFill.addContent(srgbClr);
        defRPr.addContent(solidFill);
      }
      
      tcTxStyle.addContent(defRPr);
    }
    
    return tcTxStyle;
  }
  
  /**
   * Create cell properties XML (margins/padding)
   * @private
   */
  _createCellPropertiesXML(paddingConfig) {
    const namespace = XmlService.getNamespace(this.namespace.a);
    const tcPr = XmlService.createElement('tcPr', namespace);
    
    // Convert points to EMUs
    const toEMU = (points) => Math.round(points * 12700);
    
    const marL = XmlService.createElement('marL', namespace)
      .setAttribute('val', toEMU(paddingConfig.left || 5.4).toString());
    const marR = XmlService.createElement('marR', namespace)
      .setAttribute('val', toEMU(paddingConfig.right || 5.4).toString());
    const marT = XmlService.createElement('marT', namespace)
      .setAttribute('val', toEMU(paddingConfig.top || 2.4).toString());
    const marB = XmlService.createElement('marB', namespace)
      .setAttribute('val', toEMU(paddingConfig.bottom || 2.4).toString());
    
    tcPr.addContent(marL);
    tcPr.addContent(marR);
    tcPr.addContent(marT);
    tcPr.addContent(marB);
    
    return tcPr;
  }
  
  /**
   * Get or create table styles XML collection
   * @private
   */
  _getOrCreateTableStyles() {
    try {
      return this.parser.getXML('ppt/tableStyles/tableStyle1.xml');
    } catch (error) {
      // Create new table styles collection
      const doc = XmlService.createDocument();
      const root = XmlService.createElement('tblStyleLst', XmlService.getNamespace(this.namespace.a))
        .setAttribute('def', '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}');
      
      doc.setRootElement(root);
      return doc;
    }
  }
  
  /**
   * Add style to table styles collection
   * @private
   */
  _addStyleToCollection(tableStylesXml, styleElement) {
    const root = tableStylesXml.getRootElement();
    root.addContent(styleElement.getRootElement());
    this.parser.setXML('ppt/tableStyles/tableStyle1.xml', tableStylesXml);
  }
  
  /**
   * Generate unique style ID
   * @private
   */
  _generateStyleId(styleName) {
    const hash = styleName.split('').reduce((acc, char) => {
      return ((acc << 5) - acc + char.charCodeAt(0)) & 0xffffffff;
    }, 0);
    
    return `{${Math.abs(hash).toString(16).toUpperCase().padStart(8, '0')}-ABCD-EFGH-IJKL-123456789ABC}`;
  }
  
  /**
   * Validate style configuration
   * @private
   */
  _validateStyleConfig(styleConfig) {
    if (!styleConfig.name) {
      throw new Error('Style name is required');
    }
    
    if (!styleConfig.parts || Object.keys(styleConfig.parts).length === 0) {
      throw new Error('At least one table part must be defined');
    }
    
    // Validate part names
    for (const partName of Object.keys(styleConfig.parts)) {
      if (!Object.values(BRANDWARES_TABLE_STYLE_PARTS).includes(partName)) {
        throw new Error(`Invalid table part: ${partName}`);
      }
    }
  }
  
  /**
   * Update content types for new table style
   * @private
   */
  _updateContentTypes(styleName) {
    try {
      const contentTypesXml = this.parser.getXML('[Content_Types].xml');
      const root = contentTypesXml.getRootElement();
      
      // Add override for table style
      const override = XmlService.createElement('Override')
        .setAttribute('PartName', `/ppt/tableStyles/${styleName}.xml`)
        .setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml');
      
      root.addContent(override);
      this.parser.setXML('[Content_Types].xml', contentTypesXml);
      
    } catch (error) {
      console.warn('Could not update content types:', error.message);
    }
  }
  
  /**
   * Update relationships for new table style
   * @private
   */
  _updateRelationships(styleName) {
    try {
      const relsXml = this.parser.getXML('ppt/_rels/presentation.xml.rels');
      const root = relsXml.getRootElement();
      
      // Generate new relationship ID
      const relationships = root.getChildren('Relationship', XmlService.getNamespace('http://schemas.openxmlformats.org/package/2006/relationships'));
      const newId = `rId${relationships.length + 1}`;
      
      // Add relationship for table style
      const relationship = XmlService.createElement('Relationship', XmlService.getNamespace('http://schemas.openxmlformats.org/package/2006/relationships'))
        .setAttribute('Id', newId)
        .setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles')
        .setAttribute('Target', `tableStyles/${styleName}.xml`);
      
      root.addContent(relationship);
      this.parser.setXML('ppt/_rels/presentation.xml.rels', relsXml);
      
    } catch (error) {
      console.warn('Could not update relationships:', error.message);
    }
  }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    BrandwaresTableStyles,
    TABLE_STYLE_PARTS: BRANDWARES_TABLE_STYLE_PARTS,
    ENTERPRISE_COLOR_SCHEMES
  };
}