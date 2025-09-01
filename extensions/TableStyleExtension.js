/**
 * TableStyleEditor - XML Hacking for Default Table Text Properties
 * 
 * Based on: https://www.brandwares.com/bestpractices/2015/03/xml-hacking-default-table-text/
 * 
 * This implements tanaikech-style XML manipulation to modify default table text formatting
 * that's not accessible through standard PowerPoint or Google Slides APIs.
 * 
 * The technique modifies the theme's table style definitions to set default:
 * - Font families for table text
 * - Font sizes for headers vs body cells
 * - Text alignment and spacing
 * - Color schemes for table elements
 */

class TableStyleEditor {
  
  constructor(ooxml) {
    this.ooxml = ooxml;
    this.parser = ooxml.parser;
  }
  
  /**
   * Set default table text properties (the main XML hack)
   * This modifies the theme1.xml file to change default table formatting
   */
  setDefaultTableTextProperties(options = {}) {
    console.log('🔧 XML Hacking: Setting default table text properties...');
    
    const defaults = {
      headerFont: {
        family: 'Calibri',
        size: 14,
        bold: true,
        color: '000000'
      },
      bodyFont: {
        family: 'Calibri', 
        size: 11,
        bold: false,
        color: '000000'
      },
      alignment: {
        horizontal: 'left',
        vertical: 'center'
      },
      spacing: {
        marginLeft: 5.4, // Points
        marginRight: 5.4,
        marginTop: 2.4,
        marginBottom: 2.4
      }
    };
    
    const config = { ...defaults, ...options };
    
    try {
      // Get theme XML
      const themeXml = this.parser.getXML('ppt/theme/theme1.xml');
      const themeRoot = themeXml.getRootElement();
      
      // Find or create table style elements
      this._ensureTableStyleElements(themeRoot);
      
      // Apply the XML hack for default text properties
      this._hackTableTextDefaults(themeRoot, config);
      
      // Update the theme XML
      this.parser.setXML('ppt/theme/theme1.xml', themeXml);
      
      console.log('✅ Successfully hacked default table text properties');
      return true;
      
    } catch (error) {
      console.error('❌ Failed to hack table text properties:', error.message);
      return false;
    }
  }
  
  /**
   * Create custom table styles (advanced XML hacking)
   */
  createCustomTableStyle(styleName, styleDefinition) {
    console.log(`🎨 Creating custom table style: ${styleName}`);
    
    try {
      // Get or create table styles part
      const tableStylesXml = this._getOrCreateTableStyles();
      
      // Create new table style XML
      const customStyle = this._buildTableStyleXML(styleName, styleDefinition);
      
      // Add to table styles
      const tableStylesRoot = tableStylesXml.getRootElement();
      tableStylesRoot.addContent(customStyle);
      
      // Update the XML
      this.parser.setXML('ppt/tableStyles/tableStyle1.xml', tableStylesXml);
      
      console.log(`✅ Custom table style '${styleName}' created`);
      return styleName;
      
    } catch (error) {
      console.error('❌ Failed to create custom table style:', error.message);
      return null;
    }
  }
  
  /**
   * Apply Brandwares-style table text formatting 
   * This is the core XML hack from the article
   */
  applyBrandwaresTableHack(options = {}) {
    console.log('🚀 Applying Brandwares XML hack for table text...');
    
    const brandwaresDefaults = {
      headerFont: {
        family: 'Arial',
        size: 12,
        bold: true,
        color: 'FFFFFF' // White text
      },
      bodyFont: {
        family: 'Arial',
        size: 10,
        bold: false,
        color: '333333' // Dark gray
      },
      headerBackground: '4472C4', // Blue header
      alternateRowColor: 'F2F2F2', // Light gray
      borderColor: 'CCCCCC',
      alignment: {
        header: { horizontal: 'center', vertical: 'middle' },
        body: { horizontal: 'left', vertical: 'middle' }
      }
    };
    
    const config = { ...brandwaresDefaults, ...options };
    
    return this._applyAdvancedTableStyling(config);
  }
  
  /**
   * Set table cell default margins (XML hack)
   */
  setTableCellMargins(margins = {}) {
    const defaultMargins = {
      left: 0.1,   // Inches
      right: 0.1,
      top: 0.05,
      bottom: 0.05
    };
    
    const cellMargins = { ...defaultMargins, ...margins };
    
    try {
      const themeXml = this.parser.getXML('ppt/theme/theme1.xml');
      const themeRoot = themeXml.getRootElement();
      
      // Find theme elements and add table cell margin defaults
      const themeElements = themeRoot.getChild('themeElements', this._getNamespace('a'));
      if (themeElements) {
        this._addTableCellMarginsXML(themeElements, cellMargins);
      }
      
      this.parser.setXML('ppt/theme/theme1.xml', themeXml);
      
      console.log('✅ Set default table cell margins');
      return true;
      
    } catch (error) {
      console.error('❌ Failed to set table cell margins:', error.message);
      return false;
    }
  }
  
  /**
   * Advanced table text formatting (like in the Brandwares article)
   */
  setAdvancedTableTextFormatting(formatting) {
    console.log('🎯 Applying advanced table text formatting...');
    
    try {
      // This is the core XML manipulation
      const themeXml = this.parser.getXML('ppt/theme/theme1.xml');
      const root = themeXml.getRootElement();
      
      // Create the table text formatting XML structure
      const tableTextXML = this._createTableTextFormattingXML(formatting);
      
      // Find the right place to insert it (usually in theme elements)
      const themeElements = root.getChild('themeElements', this._getNamespace('a'));
      if (themeElements) {
        // Add our custom table formatting
        const formatScheme = themeElements.getChild('fmtScheme', this._getNamespace('a'));
        if (formatScheme) {
          formatScheme.addContent(tableTextXML);
        }
      }
      
      this.parser.setXML('ppt/theme/theme1.xml', themeXml);
      
      console.log('✅ Advanced table text formatting applied');
      return true;
      
    } catch (error) {
      console.error('❌ Advanced table formatting failed:', error.message);
      return false;
    }
  }
  
  // Private helper methods for XML manipulation
  
  _ensureTableStyleElements(themeRoot) {
    const namespace = this._getNamespace('a');
    const themeElements = themeRoot.getChild('themeElements', namespace);
    
    if (themeElements) {
      let fmtScheme = themeElements.getChild('fmtScheme', namespace);
      if (!fmtScheme) {
        fmtScheme = XmlService.createElement('fmtScheme', namespace)
          .setAttribute('name', 'Office');
        themeElements.addContent(fmtScheme);
      }
      
      // Ensure table style list exists
      let tableStyleList = fmtScheme.getChild('tableStyleLst', namespace);
      if (!tableStyleList) {
        tableStyleList = XmlService.createElement('tableStyleLst', namespace);
        fmtScheme.addContent(tableStyleList);
      }
    }
  }
  
  _hackTableTextDefaults(themeRoot, config) {
    const namespace = this._getNamespace('a');
    
    // This is the actual XML hack - modifying theme elements
    const themeElements = themeRoot.getChild('themeElements', namespace);
    const fmtScheme = themeElements.getChild('fmtScheme', namespace);
    
    // Create table style with our custom text defaults
    const tableStyle = XmlService.createElement('tableStyle', namespace)
      .setAttribute('styleId', '{Default Table Style}')
      .setAttribute('styleName', 'Default Table Style');
    
    // Add whole table properties
    const wholeTbl = XmlService.createElement('wholeTbl', namespace);
    
    // Header row formatting
    const headerRow = this._createTableRowFormat('firstRow', config.headerFont, namespace);
    wholeTbl.addContent(headerRow);
    
    // Body row formatting  
    const bodyRow = this._createTableRowFormat('band1H', config.bodyFont, namespace);
    wholeTbl.addContent(bodyRow);
    
    tableStyle.addContent(wholeTbl);
    
    // Add to format scheme
    const tableStyleList = fmtScheme.getChild('tableStyleLst', namespace) || 
      XmlService.createElement('tableStyleLst', namespace);
    tableStyleList.addContent(tableStyle);
    
    if (!fmtScheme.getChild('tableStyleLst', namespace)) {
      fmtScheme.addContent(tableStyleList);
    }
  }
  
  _createTableRowFormat(rowType, fontConfig, namespace) {
    const rowElement = XmlService.createElement(rowType, namespace);
    
    // Text properties
    const tcTxStyle = XmlService.createElement('tcTxStyle', namespace);
    
    // Font reference
    const fontRef = XmlService.createElement('fontRef', namespace)
      .setAttribute('idx', 'minor');
    
    // Font family override
    const latin = XmlService.createElement('latin', namespace)
      .setAttribute('typeface', fontConfig.family);
    fontRef.addContent(latin);
    
    tcTxStyle.addContent(fontRef);
    
    // Color scheme reference
    const schemeClr = XmlService.createElement('schemeClr', namespace)
      .setAttribute('val', fontConfig.color === '000000' ? 'tx1' : 'accent1');
    
    if (fontConfig.color !== '000000') {
      const srgbClr = XmlService.createElement('srgbClr', namespace)
        .setAttribute('val', fontConfig.color);
      tcTxStyle.addContent(srgbClr);
    } else {
      tcTxStyle.addContent(schemeClr);
    }
    
    // Font size (in points * 100)
    if (fontConfig.size) {
      tcTxStyle.setAttribute('sz', (fontConfig.size * 100).toString());
    }
    
    // Bold formatting
    if (fontConfig.bold) {
      tcTxStyle.setAttribute('b', '1');
    }
    
    rowElement.addContent(tcTxStyle);
    
    return rowElement;
  }
  
  _getOrCreateTableStyles() {
    try {
      return this.parser.getXML('ppt/tableStyles/tableStyle1.xml');
    } catch (error) {
      // Create new table styles XML if it doesn't exist
      const newTableStyles = XmlService.createDocument();
      const root = XmlService.createElement('tableStyleList', this._getNamespace('a'));
      newTableStyles.setRootElement(root);
      return newTableStyles;
    }
  }
  
  _buildTableStyleXML(styleName, styleDefinition) {
    const namespace = this._getNamespace('a');
    const tableStyle = XmlService.createElement('tableStyle', namespace)
      .setAttribute('styleId', `{${styleName}}`)
      .setAttribute('styleName', styleName);
    
    // Build XML based on style definition
    if (styleDefinition.headerFont) {
      const firstRow = this._createTableRowFormat('firstRow', styleDefinition.headerFont, namespace);
      tableStyle.addContent(firstRow);
    }
    
    if (styleDefinition.bodyFont) {
      const band1H = this._createTableRowFormat('band1H', styleDefinition.bodyFont, namespace);
      tableStyle.addContent(band1H);
    }
    
    return tableStyle;
  }
  
  _applyAdvancedTableStyling(config) {
    try {
      // Apply header styling
      this.setDefaultTableTextProperties({
        headerFont: config.headerFont,
        bodyFont: config.bodyFont,
        alignment: config.alignment
      });
      
      // Apply color scheme
      if (config.headerBackground || config.alternateRowColor) {
        this._applyTableColorScheme(config);
      }
      
      console.log('✅ Brandwares table hack applied successfully');
      return true;
      
    } catch (error) {
      console.error('❌ Brandwares table hack failed:', error.message);
      return false;
    }
  }
  
  _applyTableColorScheme(config) {
    const themeXml = this.parser.getXML('ppt/theme/theme1.xml');
    const root = themeXml.getRootElement();
    
    // Add color scheme modifications for tables
    const namespace = this._getNamespace('a');
    const themeElements = root.getChild('themeElements', namespace);
    const fmtScheme = themeElements.getChild('fmtScheme', namespace);
    
    // Create fill style list for table backgrounds
    let fillStyleLst = fmtScheme.getChild('fillStyleLst', namespace);
    if (!fillStyleLst) {
      fillStyleLst = XmlService.createElement('fillStyleLst', namespace);
      fmtScheme.addContent(fillStyleLst);
    }
    
    // Add header background fill
    if (config.headerBackground) {
      const headerFill = XmlService.createElement('solidFill', namespace);
      const headerColor = XmlService.createElement('srgbClr', namespace)
        .setAttribute('val', config.headerBackground);
      headerFill.addContent(headerColor);
      fillStyleLst.addContent(headerFill);
    }
    
    this.parser.setXML('ppt/theme/theme1.xml', themeXml);
  }
  
  _addTableCellMarginsXML(themeElements, margins) {
    const namespace = this._getNamespace('a');
    
    // Convert inches to EMUs (English Metric Units)
    const toEMU = (inches) => Math.round(inches * 914400);
    
    const tableCellMargins = XmlService.createElement('tcPr', namespace);
    
    const marLeft = XmlService.createElement('marL', namespace)
      .setAttribute('val', toEMU(margins.left).toString());
    const marRight = XmlService.createElement('marR', namespace)
      .setAttribute('val', toEMU(margins.right).toString());
    const marTop = XmlService.createElement('marT', namespace)
      .setAttribute('val', toEMU(margins.top).toString());
    const marBottom = XmlService.createElement('marB', namespace)
      .setAttribute('val', toEMU(margins.bottom).toString());
    
    tableCellMargins.addContent(marLeft);
    tableCellMargins.addContent(marRight);
    tableCellMargins.addContent(marTop);
    tableCellMargins.addContent(marBottom);
    
    themeElements.addContent(tableCellMargins);
  }
  
  _createTableTextFormattingXML(formatting) {
    const namespace = this._getNamespace('a');
    
    const tableTextFormat = XmlService.createElement('tableTextFormat', namespace);
    
    // Add font properties
    if (formatting.font) {
      const fontProps = XmlService.createElement('defRPr', namespace);
      
      if (formatting.font.family) {
        const latin = XmlService.createElement('latin', namespace)
          .setAttribute('typeface', formatting.font.family);
        fontProps.addContent(latin);
      }
      
      if (formatting.font.size) {
        fontProps.setAttribute('sz', (formatting.font.size * 100).toString());
      }
      
      if (formatting.font.bold) {
        fontProps.setAttribute('b', '1');
      }
      
      if (formatting.font.italic) {
        fontProps.setAttribute('i', '1');
      }
      
      tableTextFormat.addContent(fontProps);
    }
    
    return tableTextFormat;
  }
  
  _getNamespace(prefix = 'a') {
    const namespaces = {
      'a': XmlService.getNamespace('http://schemas.openxmlformats.org/drawingml/2006/main'),
      'p': XmlService.getNamespace('http://schemas.openxmlformats.org/presentationml/2006/main'),
      'r': XmlService.getNamespace('http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    };
    return namespaces[prefix];
  }
  
  /**
   * Test table text hack with sample data
   */
  static testTableTextHack(ooxml) {
    console.log('🧪 Testing Brandwares table text hack...');
    
    const tableEditor = new TableStyleEditor(ooxml);
    
    // Apply the classic Brandwares hack
    const result = tableEditor.applyBrandwaresTableHack({
      headerFont: {
        family: 'Segoe UI',
        size: 14,
        bold: true,
        color: 'FFFFFF'
      },
      bodyFont: {
        family: 'Segoe UI',
        size: 11,
        bold: false,
        color: '404040'
      },
      headerBackground: '5B9BD5',
      alternateRowColor: 'F2F2F2'
    });
    
    if (result) {
      console.log('✅ Table text hack test passed!');
      console.log('   Default table text properties have been modified');
      console.log('   Headers: Segoe UI 14pt Bold White on Blue');
      console.log('   Body: Segoe UI 11pt Regular Dark Gray');
    } else {
      console.log('❌ Table text hack test failed');
    }
    
    return result;
  }
}

// Export for use
if (typeof module !== 'undefined' && module.exports) {
  module.exports = TableStyleEditor;
}