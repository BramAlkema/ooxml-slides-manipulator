/**
 * OOXMLSlides - Main API class for advanced PowerPoint/Google Slides manipulation
 * Provides a clean interface for OOXML editing operations
 */
class OOXMLSlides {
  constructor(fileId, options = {}) {
    this.fileId = fileId;
    this.options = {
      createBackup: options.createBackup !== false,
      autoSave: options.autoSave !== false,
      ...options
    };
    
    this.fileHandler = new FileHandler();
    this.parser = null;
    this.theme = null;
    this.slides = null;
    this.isLoaded = false;
  }

  /**
   * Load and parse the PPTX file
   * @returns {OOXMLSlides} Fluent interface
   */
  load() {
    try {
      // Validate file access
      if (!this.fileHandler.validateAccess(this.fileId)) {
        throw new Error(`Cannot access file: ${this.fileId}`);
      }
      
      // Create backup if requested
      if (this.options.createBackup) {
        this.backupId = this.fileHandler.createBackup(this.fileId);
        console.log(`Backup created: ${this.backupId}`);
      }
      
      // Load file content
      const fileBlob = this.fileHandler.loadFile(this.fileId);
      
      // Initialize parser and managers
      this.parser = new OOXMLParser(fileBlob);
      this.theme = new ThemeEditor(this.parser);
      this.slides = new SlideManager(this.parser);
      
      this.isLoaded = true;
      return this;
    } catch (error) {
      throw new Error(`Failed to load OOXML file: ${error.message}`);
    }
  }

  /**
   * Save modifications
   * @param {Object} options - Save options
   * @returns {string} New file ID or updated file ID
   */
  save(options = {}) {
    this._ensureLoaded();
    
    try {
      const modifiedBlob = this.parser.build();
      
      if (options.asNewFile !== false) {
        // Save as new file
        const newFileId = this.fileHandler.saveFile(modifiedBlob, {
          name: options.name,
          folderId: options.folderId
        });
        
        if (options.importToSlides) {
          return this.fileHandler.importToSlides(modifiedBlob, {
            name: options.name,
            folderId: options.folderId
          });
        }
        
        return newFileId;
      } else {
        // Update existing file
        return this.fileHandler.updateFile(this.fileId, modifiedBlob);
      }
    } catch (error) {
      throw new Error(`Failed to save OOXML file: ${error.message}`);
    }
  }

  // Theme operations (fluent interface)
  
  /**
   * Set color palette
   * @param {Object|Array} colors - Color palette object or array
   * @returns {OOXMLSlides} Fluent interface
   */
  setColors(colors) {
    this._ensureLoaded();
    
    if (Array.isArray(colors)) {
      this.theme.setColorPaletteFromArray(colors);
    } else {
      this.theme.setColorPalette(colors);
    }
    
    return this;
  }

  /**
   * Set predefined color theme
   * @param {string} themeName - Theme name ('office', 'colorful', 'dark')
   * @returns {OOXMLSlides} Fluent interface
   */
  setColorTheme(themeName) {
    this._ensureLoaded();
    this.theme.setColorTheme(themeName);
    return this;
  }

  /**
   * Set font pair
   * @param {string} majorFont - Heading font
   * @param {string} minorFont - Body font
   * @returns {OOXMLSlides} Fluent interface
   */
  setFonts(majorFont, minorFont) {
    this._ensureLoaded();
    this.theme.setFontPair(majorFont, minorFont);
    return this;
  }

  /**
   * Get current theme information
   * @returns {Object} Theme information
   */
  getTheme() {
    this._ensureLoaded();
    return {
      colors: this.theme.getColorScheme(),
      fonts: this.theme.getFontScheme()
    };
  }

  // Slide operations
  
  /**
   * Set slide dimensions
   * @param {number} width - Width in pixels
   * @param {number} height - Height in pixels
   * @returns {OOXMLSlides} Fluent interface
   */
  setSize(width, height) {
    this._ensureLoaded();
    this.slides.setSlideSize({ width, height });
    return this;
  }

  /**
   * Get slide dimensions
   * @returns {Object} {width, height} in pixels
   */
  getSize() {
    this._ensureLoaded();
    return this.slides.getSlideSize();
  }

  /**
   * Modify slide master
   * @param {string} masterId - Master ID
   * @param {Object} properties - Properties to modify
   * @returns {OOXMLSlides} Fluent interface
   */
  modifyMaster(masterId, properties) {
    this._ensureLoaded();
    this.slides.modifySlideMaster(masterId, properties);
    return this;
  }

  /**
   * Modify slide layout
   * @param {string} layoutId - Layout ID
   * @param {Object} properties - Properties to modify
   * @returns {OOXMLSlides} Fluent interface
   */
  modifyLayout(layoutId, properties) {
    this._ensureLoaded();
    this.slides.modifySlideLayout(layoutId, properties);
    return this;
  }

  // Utility methods
  
  /**
   * Get file information
   * @returns {Object} File metadata
   */
  getFileInfo() {
    return this.fileHandler.getFileInfo(this.fileId);
  }

  /**
   * List all files in the OOXML package
   * @returns {Array<string>} File paths
   */
  listFiles() {
    this._ensureLoaded();
    return this.parser.listFiles();
  }

  /**
   * Export raw XML content for inspection
   * @param {string} path - XML file path
   * @returns {string} XML content
   */
  exportXML(path) {
    this._ensureLoaded();
    const xmlDoc = this.parser.getXML(path);
    return XmlService.getPrettyFormat().format(xmlDoc);
  }

  /**
   * Chain multiple operations
   * @param {Function} callback - Function to call with this instance
   * @returns {OOXMLSlides} Fluent interface
   */
  batch(callback) {
    callback(this);
    return this;
  }

  // Private methods
  
  _ensureLoaded() {
    if (!this.isLoaded) {
      throw new Error('File not loaded. Call load() first.');
    }
  }

  // Static factory methods
  
  /**
   * Create instance and load file
   * @param {string} fileId - Google Drive file ID
   * @param {Object} options - Options
   * @returns {OOXMLSlides} Loaded instance
   */
  static fromFile(fileId, options = {}) {
    return new OOXMLSlides(fileId, options).load();
  }

  /**
   * Create instance from blob
   * @param {Blob} blob - PPTX blob
   * @param {Object} options - Options
   * @returns {OOXMLSlides} Loaded instance
   */
  static fromBlob(blob, options = {}) {
    const instance = new OOXMLSlides(null, options);
    instance.parser = new OOXMLParser(blob);
    instance.theme = new ThemeEditor(instance.parser);
    instance.slides = new SlideManager(instance.parser);
    instance.isLoaded = true;
    return instance;
  }
}

// Example usage and presets
OOXMLSlides.Presets = {
  /**
   * Apply modern color theme
   */
  modernTheme: (slides) => {
    return slides
      .setColors(['#2C3E50', '#FFFFFF', '#34495E', '#ECF0F1', '#3498DB', '#E74C3C'])
      .setFonts('Segoe UI', 'Segoe UI');
  },

  /**
   * Apply corporate color theme  
   */
  corporateTheme: (slides) => {
    return slides
      .setColors(['#1F2937', '#FFFFFF', '#374151', '#F3F4F6', '#3B82F6', '#10B981'])
      .setFonts('Arial', 'Arial');
  },

  /**
   * Apply creative color theme
   */
  creativeTheme: (slides) => {
    return slides
      .setColors(['#000000', '#FFFFFF', '#444444', '#F0F0F0', '#FF6B6B', '#4ECDC4'])
      .setFonts('Montserrat', 'Open Sans');
  }
};

// Export for Google Apps Script
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { OOXMLSlides, OOXMLParser, ThemeEditor, SlideManager, FileHandler };
}