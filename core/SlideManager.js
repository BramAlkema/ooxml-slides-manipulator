/**
 * SlideManager - Slide master and layout manipulation
 * Handles slide masters, layouts, and presentation structure
 */
class SlideManager {
  constructor(ooxmlParser) {
    this.parser = ooxmlParser;
    this.namespaces = {
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    };
  }

  /**
   * Get all slide masters
   * @returns {Array} Array of slide master information
   */
  getSlideMasters() {
    const masters = [];
    const files = this.parser.listFiles();
    
    files.forEach(file => {
      if (file.startsWith('ppt/slideMasters/') && file.endsWith('.xml')) {
        const masterDoc = this.parser.getXML(file);
        const masterId = file.match(/slideMaster(\d+)\.xml/)[1];
        masters.push({
          id: masterId,
          path: file,
          document: masterDoc
        });
      }
    });
    
    return masters;
  }

  /**
   * Get all slide layouts
   * @returns {Array} Array of slide layout information
   */
  getSlideLayouts() {
    const layouts = [];
    const files = this.parser.listFiles();
    
    files.forEach(file => {
      if (file.startsWith('ppt/slideLayouts/') && file.endsWith('.xml')) {
        const layoutDoc = this.parser.getXML(file);
        const layoutId = file.match(/slideLayout(\d+)\.xml/)[1];
        layouts.push({
          id: layoutId,
          path: file,
          document: layoutDoc
        });
      }
    });
    
    return layouts;
  }

  /**
   * Modify slide master properties
   * @param {string} masterId - Master ID to modify
   * @param {Object} properties - Properties to update
   */
  modifySlideMaster(masterId, properties) {
    const masterPath = `ppt/slideMasters/slideMaster${masterId}.xml`;
    
    if (!this.parser.hasFile(masterPath)) {
      throw new Error(`Slide master not found: ${masterId}`);
    }
    
    const masterDoc = this.parser.getXML(masterPath);
    const root = masterDoc.getRootElement();
    
    // Update properties
    if (properties.background) {
      this._setBackground(root, properties.background);
    }
    
    if (properties.textStyles) {
      this._updateTextStyles(root, properties.textStyles);
    }
    
    this.parser.setXML(masterPath, masterDoc);
    return this;
  }

  /**
   * Modify slide layout
   * @param {string} layoutId - Layout ID to modify  
   * @param {Object} properties - Properties to update
   */
  modifySlideLayout(layoutId, properties) {
    const layoutPath = `ppt/slideLayouts/slideLayout${layoutId}.xml`;
    
    if (!this.parser.hasFile(layoutPath)) {
      throw new Error(`Slide layout not found: ${layoutId}`);
    }
    
    const layoutDoc = this.parser.getXML(layoutPath);
    const root = layoutDoc.getRootElement();
    
    // Update layout properties
    if (properties.name) {
      root.setAttribute('type', properties.name);
    }
    
    if (properties.placeholders) {
      this._updatePlaceholders(root, properties.placeholders);
    }
    
    this.parser.setXML(layoutPath, layoutDoc);
    return this;
  }

  /**
   * Set presentation slide size
   * @param {Object} size - {width: number, height: number} in pixels
   */
  setSlideSize(size) {
    const presDoc = this.parser.getPresentationXML();
    const root = presDoc.getRootElement();
    const sldSz = root.getChild('sldSz', XmlService.getNamespace('p', this.namespaces.p));
    
    if (sldSz) {
      // Convert pixels to EMU (English Metric Units)
      // 1 pixel = 12700 EMU (at 96 DPI)
      const widthEMU = Math.round(size.width * 12700);
      const heightEMU = Math.round(size.height * 12700);
      
      sldSz.setAttribute('cx', widthEMU.toString());
      sldSz.setAttribute('cy', heightEMU.toString());
    }
    
    this.parser.setPresentationXML(presDoc);
    return this;
  }

  /**
   * Get current slide size
   * @returns {Object} {width: number, height: number} in pixels
   */
  getSlideSize() {
    const presDoc = this.parser.getPresentationXML();
    const root = presDoc.getRootElement();
    const sldSz = root.getChild('sldSz', XmlService.getNamespace('p', this.namespaces.p));
    
    if (sldSz) {
      const widthEMU = parseInt(sldSz.getAttribute('cx').getValue());
      const heightEMU = parseInt(sldSz.getAttribute('cy').getValue());
      
      // Convert EMU back to pixels
      return {
        width: Math.round(widthEMU / 12700),
        height: Math.round(heightEMU / 12700)
      };
    }
    
    return { width: 960, height: 720 }; // Default PowerPoint size
  }

  /**
   * Add a new slide layout based on existing layout
   * @param {string} baseLayoutId - ID of layout to copy from
   * @param {Object} modifications - Modifications to apply
   * @returns {string} New layout ID
   */
  createSlideLayout(baseLayoutId, modifications = {}) {
    const baseLayoutPath = `ppt/slideLayouts/slideLayout${baseLayoutId}.xml`;
    
    if (!this.parser.hasFile(baseLayoutPath)) {
      throw new Error(`Base slide layout not found: ${baseLayoutId}`);
    }
    
    // Find next available layout ID
    const layouts = this.getSlideLayouts();
    const maxId = Math.max(...layouts.map(l => parseInt(l.id)));
    const newId = maxId + 1;
    
    // Copy base layout
    const baseDoc = this.parser.getXML(baseLayoutPath);
    const newLayoutPath = `ppt/slideLayouts/slideLayout${newId}.xml`;
    
    // Apply modifications
    const root = baseDoc.getRootElement();
    if (modifications.name) {
      root.setAttribute('type', modifications.name);
    }
    
    // Save new layout
    this.parser.setXML(newLayoutPath, baseDoc);
    
    // TODO: Update relationships and content types
    this._updateLayoutRelationships(newId);
    
    return newId.toString();
  }

  // Private helper methods
  _setBackground(root, background) {
    // Find or create background element
    const cSld = root.getChild('cSld', XmlService.getNamespace('p', this.namespaces.p));
    let bg = cSld.getChild('bg', XmlService.getNamespace('p', this.namespaces.p));
    
    if (!bg) {
      bg = XmlService.createElement('bg', XmlService.getNamespace('p', this.namespaces.p));
      cSld.addContent(0, bg);
    }
    
    // Clear existing background
    bg.removeContent();
    
    if (background.color) {
      // Solid color background
      const bgPr = XmlService.createElement('bgPr', XmlService.getNamespace('p', this.namespaces.p));
      const solidFill = XmlService.createElement('solidFill', XmlService.getNamespace('a', this.namespaces.a));
      const srgbClr = XmlService.createElement('srgbClr', XmlService.getNamespace('a', this.namespaces.a));
      srgbClr.setAttribute('val', background.color.replace('#', ''));
      
      solidFill.addContent(srgbClr);
      bgPr.addContent(solidFill);
      bg.addContent(bgPr);
    }
  }

  _updateTextStyles(root, textStyles) {
    const txStyles = root.getChild('txStyles', XmlService.getNamespace('p', this.namespaces.p));
    if (!txStyles) return;
    
    // Update title style
    if (textStyles.title) {
      const titleStyle = txStyles.getChild('titleStyle', XmlService.getNamespace('p', this.namespaces.p));
      if (titleStyle) {
        this._applyTextStyle(titleStyle, textStyles.title);
      }
    }
    
    // Update body style
    if (textStyles.body) {
      const bodyStyle = txStyles.getChild('bodyStyle', XmlService.getNamespace('p', this.namespaces.p));
      if (bodyStyle) {
        this._applyTextStyle(bodyStyle, textStyles.body);
      }
    }
  }

  _applyTextStyle(styleElement, style) {
    // Find first level paragraph properties
    const lvl1pPr = styleElement.getChild('lvl1pPr', XmlService.getNamespace('a', this.namespaces.a));
    if (!lvl1pPr) return;
    
    if (style.fontSize) {
      const defRPr = lvl1pPr.getChild('defRPr', XmlService.getNamespace('a', this.namespaces.a));
      if (defRPr) {
        defRPr.setAttribute('sz', (style.fontSize * 100).toString());
      }
    }
  }

  _updatePlaceholders(root, placeholders) {
    const cSld = root.getChild('cSld', XmlService.getNamespace('p', this.namespaces.p));
    const spTree = cSld.getChild('spTree', XmlService.getNamespace('p', this.namespaces.p));
    
    placeholders.forEach(placeholder => {
      // Find placeholder shape
      const shapes = spTree.getChildren('sp', XmlService.getNamespace('p', this.namespaces.p));
      shapes.forEach(shape => {
        const nvSpPr = shape.getChild('nvSpPr', XmlService.getNamespace('p', this.namespaces.p));
        const nvPr = nvSpPr.getChild('nvPr', XmlService.getNamespace('p', this.namespaces.p));
        const ph = nvPr.getChild('ph', XmlService.getNamespace('p', this.namespaces.p));
        
        if (ph && ph.getAttribute('type') && ph.getAttribute('type').getValue() === placeholder.type) {
          // Update placeholder position/size
          if (placeholder.position || placeholder.size) {
            this._updateShapeTransform(shape, placeholder.position, placeholder.size);
          }
        }
      });
    });
  }

  _updateShapeTransform(shape, position, size) {
    const spPr = shape.getChild('spPr', XmlService.getNamespace('p', this.namespaces.p));
    let xfrm = spPr.getChild('xfrm', XmlService.getNamespace('a', this.namespaces.a));
    
    if (!xfrm) {
      xfrm = XmlService.createElement('xfrm', XmlService.getNamespace('a', this.namespaces.a));
      spPr.addContent(0, xfrm);
    }
    
    if (position) {
      let off = xfrm.getChild('off', XmlService.getNamespace('a', this.namespaces.a));
      if (!off) {
        off = XmlService.createElement('off', XmlService.getNamespace('a', this.namespaces.a));
        xfrm.addContent(off);
      }
      off.setAttribute('x', (position.x * 12700).toString());
      off.setAttribute('y', (position.y * 12700).toString());
    }
    
    if (size) {
      let ext = xfrm.getChild('ext', XmlService.getNamespace('a', this.namespaces.a));
      if (!ext) {
        ext = XmlService.createElement('ext', XmlService.getNamespace('a', this.namespaces.a));
        xfrm.addContent(ext);
      }
      ext.setAttribute('cx', (size.width * 12700).toString());
      ext.setAttribute('cy', (size.height * 12700).toString());
    }
  }

  _updateLayoutRelationships(layoutId) {
    // Update content types
    const contentTypesPath = '[Content_Types].xml';
    if (this.parser.hasFile(contentTypesPath)) {
      const contentDoc = this.parser.getXML(contentTypesPath);
      // Add new layout to content types
      // Implementation would add Override element for new layout
    }
    
    // Update master relationships
    // Implementation would update slide master relationships file
  }
}