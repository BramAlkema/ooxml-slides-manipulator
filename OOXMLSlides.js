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
      this.tables = new TableStyleEditor(this);
      this.colors = new CustomColorEditor(this);
      this.numbering = new NumberingStyleEditor(this);
      this.typography = new TypographyEditor(this);
      this.superTheme = new SuperThemeEditor(this);
      this.xmlSearchReplace = new XMLSearchReplaceEditor(this);
      this.languageStandardization = new LanguageStandardizationEditor(this);
      
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

  /**
   * Save to Google Drive with specified name
   * @param {string} fileName - Name for the new file
   * @returns {string} File ID of the saved file
   */
  saveToGoogleDrive(fileName) {
    return this.save({ name: fileName });
  }

  /**
   * Get theme XML for compatibility analysis
   * @returns {string} Theme XML content
   */
  getThemeXML() {
    this._ensureLoaded();
    return this.theme.getThemeXML();
  }

  /**
   * Get font XML for compatibility analysis
   * @returns {string} Font XML content
   */
  getFontXML() {
    this._ensureLoaded();
    return this.theme.getFontXML();
  }

  /**
   * Get color XML for compatibility analysis
   * @returns {string} Color XML content
   */
  getColorXML() {
    this._ensureLoaded();
    return this.theme.getColorXML();
  }

  /**
   * Get current colors for analysis
   * @returns {Object} Color mapping
   */
  getColors() {
    this._ensureLoaded();
    return this.theme.getColors();
  }

  /**
   * Inject custom theme XML for testing
   * @param {string} customXML - Custom XML to inject
   */
  injectCustomThemeXML(customXML) {
    this._ensureLoaded();
    this.theme.injectCustomXML(customXML);
  }

  /**
   * Inject custom font XML for testing
   * @param {string} customXML - Custom font XML to inject
   */
  injectCustomFontXML(customXML) {
    this._ensureLoaded();
    this.theme.injectCustomFontXML(customXML);
  }

  /**
   * Inject custom color XML for testing
   * @param {string} customXML - Custom color XML to inject
   */
  injectCustomColorXML(customXML) {
    this._ensureLoaded();
    this.theme.injectCustomColorXML(customXML);
  }

  /**
   * Inject custom master XML for testing
   * @param {string} customXML - Custom master XML to inject
   */
  injectCustomMasterXML(customXML) {
    this._ensureLoaded();
    this.slides.injectCustomMasterXML(customXML);
  }

  /**
   * Inject custom formatting XML for testing
   * @param {string} customXML - Custom formatting XML to inject
   */
  injectCustomFormattingXML(customXML) {
    this._ensureLoaded();
    this.parser.injectCustomFormattingXML(customXML);
  }

  /**
   * Add custom XML parts for testing
   * @param {Object} xmlParts - Custom XML parts to add
   */
  addCustomXMLParts(xmlParts) {
    this._ensureLoaded();
    this.parser.addCustomXMLParts(xmlParts);
  }

  // Table Styling Operations (Brandwares XML Hack Implementation)

  /**
   * Set default table text properties using XML hacking technique
   * Based on: https://www.brandwares.com/bestpractices/2015/03/xml-hacking-default-table-text/
   * @param {Object} options - Table text formatting options
   * @returns {OOXMLSlides} Fluent interface
   */
  setDefaultTableText(options = {}) {
    this._ensureLoaded();
    this.tables.setDefaultTableTextProperties(options);
    return this;
  }

  /**
   * Apply the classic Brandwares table text hack
   * @param {Object} options - Brandwares-style formatting options
   * @returns {OOXMLSlides} Fluent interface
   */
  applyBrandwaresTableHack(options = {}) {
    this._ensureLoaded();
    this.tables.applyBrandwaresTableHack(options);
    return this;
  }

  /**
   * Create custom table style with advanced formatting
   * @param {string} styleName - Name for the custom style
   * @param {Object} styleDefinition - Style configuration
   * @returns {OOXMLSlides} Fluent interface
   */
  createTableStyle(styleName, styleDefinition) {
    this._ensureLoaded();
    this.tables.createCustomTableStyle(styleName, styleDefinition);
    return this;
  }

  /**
   * Set table cell margins (XML hack for default margins)
   * @param {Object} margins - Margin settings in inches
   * @returns {OOXMLSlides} Fluent interface
   */
  setTableCellMargins(margins = {}) {
    this._ensureLoaded();
    this.tables.setTableCellMargins(margins);
    return this;
  }

  // Custom Color Operations (Brandwares XML Hack Implementation)

  /**
   * Add custom colors beyond the standard 12-color theme
   * Based on: https://www.brandwares.com/bestpractices/2015/06/xml-hacking-custom-colors/
   * @param {Object} customColors - Object mapping color names to hex values
   * @returns {OOXMLSlides} Fluent interface
   */
  addCustomColors(customColors) {
    this._ensureLoaded();
    this.colors.addCustomColors(customColors);
    return this;
  }

  /**
   * Apply the classic Brandwares custom color hack
   * @param {Object} options - Additional color options to merge
   * @returns {OOXMLSlides} Fluent interface
   */
  applyBrandwaresColorHack(options = {}) {
    this._ensureLoaded();
    this.colors.applyBrandwaresCustomColorHack(options);
    return this;
  }

  /**
   * Create a brand-specific custom color palette
   * @param {string} brandName - Name of the brand
   * @param {Object} brandColors - Brand color definitions
   * @returns {OOXMLSlides} Fluent interface
   */
  createBrandColorPalette(brandName, brandColors) {
    this._ensureLoaded();
    this.colors.createBrandColorPalette(brandName, brandColors);
    return this;
  }

  /**
   * Create Material Design color palette
   * @returns {OOXMLSlides} Fluent interface
   */
  addMaterialDesignColors() {
    this._ensureLoaded();
    this.colors.createMaterialDesignPalette();
    return this;
  }

  /**
   * Create accessibility-compliant WCAG color palette
   * @returns {OOXMLSlides} Fluent interface
   */
  addAccessibleColors() {
    this._ensureLoaded();
    this.colors.createAccessibleColorPalette();
    return this;
  }

  /**
   * Add custom gradient definitions
   * @param {Object} gradientDefinitions - Gradient definitions
   * @returns {OOXMLSlides} Fluent interface
   */
  addCustomGradients(gradientDefinitions) {
    this._ensureLoaded();
    this.colors.addCustomGradients(gradientDefinitions);
    return this;
  }

  /**
   * Create color harmony based on color theory
   * @param {string} baseColor - Base color in hex format
   * @param {string} harmonyType - Type of harmony (complementary, triadic, analogous, monochromatic)
   * @returns {OOXMLSlides} Fluent interface
   */
  createColorHarmony(baseColor, harmonyType = 'complementary') {
    this._ensureLoaded();
    this.colors.createColorHarmony(baseColor, harmonyType);
    return this;
  }

  /**
   * Export current color palette for analysis
   * @returns {Array} Array of color hex values
   */
  exportColorPalette() {
    this._ensureLoaded();
    return this.colors.exportColorPalette();
  }

  // Numbering Style Operations (Brandwares XML Hack Implementation)

  /**
   * Create custom numbering and bullet styles
   * Based on: https://www.brandwares.com/bestpractices/2017/06/xml-hacking-powerpoint-numbering-styles/
   * @param {Object} numberingDefinitions - Numbering style definitions
   * @returns {OOXMLSlides} Fluent interface
   */
  createNumberingStyles(numberingDefinitions) {
    this._ensureLoaded();
    this.numbering.createCustomNumberingStyles(numberingDefinitions);
    return this;
  }

  /**
   * Apply the classic Brandwares numbering styles hack
   * @returns {OOXMLSlides} Fluent interface
   */
  applyBrandwaresNumberingHack() {
    this._ensureLoaded();
    this.numbering.applyBrandwaresNumberingHack();
    return this;
  }

  /**
   * Create modern design system numbering styles
   * @returns {OOXMLSlides} Fluent interface
   */
  addModernNumberingStyles() {
    this._ensureLoaded();
    this.numbering.createModernNumberingStyles();
    return this;
  }

  /**
   * Create accessible numbering styles (WCAG compliant)
   * @returns {OOXMLSlides} Fluent interface
   */
  addAccessibleNumberingStyles() {
    this._ensureLoaded();
    this.numbering.createAccessibleNumberingStyles();
    return this;
  }

  /**
   * Create custom bullet styles with symbols and graphics
   * @param {Object} bulletDefinitions - Bullet style definitions
   * @returns {OOXMLSlides} Fluent interface
   */
  createBulletStyles(bulletDefinitions) {
    this._ensureLoaded();
    this.numbering.createCustomBulletStyles(bulletDefinitions);
    return this;
  }

  /**
   * Create numbered list with restart capability
   * @param {string} listId - Unique identifier for the list
   * @param {number} restartAt - Number to restart counting at
   * @returns {OOXMLSlides} Fluent interface
   */
  createRestartableList(listId, restartAt = 1) {
    this._ensureLoaded();
    this.numbering.createRestartableNumberedList(listId, restartAt);
    return this;
  }

  /**
   * Apply numbering style to specific slide element
   * @param {number} slideIndex - Zero-based slide index
   * @param {string} elementId - Element ID to apply numbering to
   * @param {string} numberingStyleId - Numbering style identifier
   * @returns {OOXMLSlides} Fluent interface
   */
  applyNumberingToElement(slideIndex, elementId, numberingStyleId) {
    this._ensureLoaded();
    this.numbering.applyNumberingToSlideElement(slideIndex, elementId, numberingStyleId);
    return this;
  }

  /**
   * Export current numbering styles for analysis
   * @returns {Array} Array of numbering style definitions
   */
  exportNumberingStyles() {
    this._ensureLoaded();
    return this.numbering.exportNumberingStyles();
  }

  // Typography and Kerning Operations (Tanaikech-Style XML Manipulation)

  /**
   * Apply advanced kerning and typography controls
   * @param {Object} kerningDefinitions - Typography and kerning definitions
   * @returns {OOXMLSlides} Fluent interface
   */
  applyAdvancedKerning(kerningDefinitions) {
    this._ensureLoaded();
    this.typography.applyAdvancedKerning(kerningDefinitions);
    return this;
  }

  /**
   * Apply professional typography presets
   * @param {string} preset - Typography preset ('corporate', 'editorial', 'technical')
   * @returns {OOXMLSlides} Fluent interface
   */
  applyProfessionalTypography(preset = 'corporate') {
    this._ensureLoaded();
    this.typography.applyProfessionalTypography(preset);
    return this;
  }

  /**
   * Create custom kerning pairs for specific font combinations
   * @param {string} fontFamily - Font family to apply kerning to
   * @param {Object} kerningPairs - Object mapping character pairs to adjustments
   * @returns {OOXMLSlides} Fluent interface
   */
  createKerningPairs(fontFamily, kerningPairs) {
    this._ensureLoaded();
    this.typography.createKerningPairs(fontFamily, kerningPairs);
    return this;
  }

  /**
   * Apply OpenType font features
   * @param {string} fontFamily - Font family to apply features to
   * @param {Array} features - Array of OpenType feature tags
   * @returns {OOXMLSlides} Fluent interface
   */
  applyOpenTypeFeatures(fontFamily, features) {
    this._ensureLoaded();
    this.typography.applyOpenTypeFeatures(fontFamily, features);
    return this;
  }

  /**
   * Set character tracking (letter-spacing) for text styles
   * @param {string} textStyleName - Name of the text style
   * @param {number} trackingValue - Tracking value as percentage (0.05 = 5%)
   * @returns {OOXMLSlides} Fluent interface
   */
  setCharacterTracking(textStyleName, trackingValue) {
    this._ensureLoaded();
    this.typography.setCharacterTracking(textStyleName, trackingValue);
    return this;
  }

  /**
   * Set word spacing adjustments
   * @param {string} textStyleName - Name of the text style
   * @param {number} wordSpacingRatio - Word spacing ratio (1.0 = normal, 1.2 = 20% wider)
   * @returns {OOXMLSlides} Fluent interface
   */
  setWordSpacing(textStyleName, wordSpacingRatio) {
    this._ensureLoaded();
    this.typography.setWordSpacing(textStyleName, wordSpacingRatio);
    return this;
  }

  /**
   * Apply typography to specific slide element
   * @param {number} slideIndex - Zero-based slide index
   * @param {string} elementId - Element ID to apply typography to
   * @param {Object} typographySettings - Typography settings object
   * @returns {OOXMLSlides} Fluent interface
   */
  applyTypographyToElement(slideIndex, elementId, typographySettings) {
    this._ensureLoaded();
    this.typography.applyTypographyToElement(slideIndex, elementId, typographySettings);
    return this;
  }

  /**
   * Create responsive typography that adjusts based on text size
   * @returns {OOXMLSlides} Fluent interface
   */
  createResponsiveTypography() {
    this._ensureLoaded();
    this.typography.createResponsiveTypography();
    return this;
  }

  /**
   * Export current typography settings for analysis
   * @returns {Array} Array of typography configuration objects
   */
  exportTypographySettings() {
    this._ensureLoaded();
    return this.typography.exportTypographySettings();
  }

  // SuperTheme Operations (Microsoft PowerPoint SuperTheme XML Manipulation)

  /**
   * Analyze an existing SuperTheme file structure
   * @param {Blob} superThemeBlob - SuperTheme .thmx file
   * @returns {Object} Analysis of SuperTheme structure
   */
  analyzeSuperTheme(superThemeBlob) {
    return SuperThemeEditor.prototype.analyzeSuperTheme.call({ parser: this.parser }, superThemeBlob);
  }

  /**
   * Create a custom SuperTheme with multiple design and size variants
   * @param {Object} superThemeDefinition - SuperTheme configuration
   * @returns {Blob} Generated SuperTheme .thmx file
   */
  createSuperTheme(superThemeDefinition) {
    return SuperThemeEditor.prototype.createSuperTheme.call({ parser: this.parser }, superThemeDefinition);
  }

  /**
   * Extract individual theme variant from SuperTheme
   * @param {Blob} superThemeBlob - SuperTheme .thmx file
   * @param {number} variantIndex - Which variant to extract (1-based)
   * @returns {Object} Extracted theme data
   */
  extractThemeVariant(superThemeBlob, variantIndex) {
    return SuperThemeEditor.prototype.extractThemeVariant.call({ parser: this.parser }, superThemeBlob, variantIndex);
  }

  /**
   * Convert regular PPTX themes to SuperTheme format
   * @param {Array<Object>} themeDefinitions - Array of theme configurations
   * @returns {Blob} Generated SuperTheme
   */
  convertToSuperTheme(themeDefinitions) {
    return SuperThemeEditor.prototype.convertToSuperTheme.call({ parser: this.parser }, themeDefinitions);
  }

  /**
   * Create responsive SuperTheme that adapts to different screen sizes
   * @param {Object} responsiveConfig - Responsive configuration
   * @returns {Blob} Responsive SuperTheme
   */
  createResponsiveSuperTheme(responsiveConfig) {
    return SuperThemeEditor.prototype.createResponsiveSuperTheme.call({ parser: this.parser }, responsiveConfig);
  }

  // XML Search & Replace Operations (Multi-file XML Manipulation)

  /**
   * Perform multi-file XML search and replace operation
   * @param {Object} searchReplaceConfig - Search and replace configuration
   * @returns {Object} Results of the operation
   */
  searchReplaceXML(searchReplaceConfig) {
    this._ensureLoaded();
    return this.xmlSearchReplace.searchReplaceXML(searchReplaceConfig);
  }

  /**
   * Search and replace with regular expressions
   * @param {string} searchPattern - Regular expression pattern
   * @param {string} replacePattern - Replacement pattern
   * @param {Object} options - Additional options
   * @returns {Object} Results of the operation
   */
  searchReplaceRegex(searchPattern, replacePattern, options = {}) {
    this._ensureLoaded();
    return this.xmlSearchReplace.searchReplaceRegex(searchPattern, replacePattern, options);
  }

  /**
   * Search and replace XML attributes specifically
   * @param {string} attributeName - Name of the attribute to search
   * @param {string} searchValue - Value to search for
   * @param {string} replaceValue - Value to replace with
   * @param {Object} options - Additional options
   * @returns {Object} Results of the operation
   */
  searchReplaceAttribute(attributeName, searchValue, replaceValue, options = {}) {
    this._ensureLoaded();
    return this.xmlSearchReplace.searchReplaceAttribute(attributeName, searchValue, replaceValue, options);
  }

  /**
   * Find all occurrences of a pattern without replacing
   * @param {string} searchPattern - Pattern to search for
   * @param {Object} options - Search options
   * @returns {Array<Object>} All matches found
   */
  findInXML(searchPattern, options = {}) {
    this._ensureLoaded();
    return this.xmlSearchReplace.findAll(searchPattern, options);
  }

  // Language Standardization Operations (Brandwares Production Technique)

  /**
   * Standardize all language tags to a target language
   * @param {string} targetLanguage - Target language code (e.g., 'en-US', 'fr-FR')
   * @param {Object} options - Standardization options
   * @returns {Object} Results of the standardization
   */
  standardizeLanguage(targetLanguage, options = {}) {
    this._ensureLoaded();
    return this.languageStandardization.standardizeLanguage(targetLanguage, options);
  }

  /**
   * Analyze current language tag distribution
   * @returns {Object} Language analysis results
   */
  analyzeLanguages() {
    this._ensureLoaded();
    return this.languageStandardization.analyzeLanguages();
  }

  /**
   * Standardize to US English (most common business requirement)
   * @returns {Object} Results of standardization
   */
  standardizeToUSEnglish() {
    this._ensureLoaded();
    return this.languageStandardization.standardizeToUSEnglish();
  }

  /**
   * Fix mixed language paragraphs (common PowerPoint issue)
   * @param {string} targetLanguage - Target language for consistency
   * @returns {Object} Results of the fix
   */
  fixMixedLanguageParagraphs(targetLanguage) {
    this._ensureLoaded();
    return this.languageStandardization.fixMixedLanguageParagraphs(targetLanguage);
  }

  /**
   * Generate comprehensive language standardization report
   * @returns {Object} Language analysis report
   */
  generateLanguageReport() {
    this._ensureLoaded();
    return this.languageStandardization.generateLanguageReport();
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
    instance.tables = new TableStyleEditor(instance);
    instance.colors = new CustomColorEditor(instance);
    instance.numbering = new NumberingStyleEditor(instance);
    instance.typography = new TypographyEditor(instance);
    instance.superTheme = new SuperThemeEditor(instance);
    instance.xmlSearchReplace = new XMLSearchReplaceEditor(instance);
    instance.languageStandardization = new LanguageStandardizationEditor(instance);
    instance.isLoaded = true;
    return instance;
  }

  /**
   * Create new presentation from template (like tanaikech's approach)
   * @param {Object} templateOptions - Template creation options
   * @param {Object} instanceOptions - Instance options
   * @returns {OOXMLSlides} New presentation instance
   */
  static fromTemplate(templateOptions = {}, instanceOptions = {}) {
    const instance = new OOXMLSlides(null, instanceOptions);
    instance.parser = OOXMLParser.fromTemplate(templateOptions);
    instance.theme = new ThemeEditor(instance.parser);
    instance.slides = new SlideManager(instance.parser);
    instance.tables = new TableStyleEditor(instance);
    instance.colors = new CustomColorEditor(instance);
    instance.numbering = new NumberingStyleEditor(instance);
    instance.typography = new TypographyEditor(instance);
    instance.superTheme = new SuperThemeEditor(instance);
    instance.xmlSearchReplace = new XMLSearchReplaceEditor(instance);
    instance.languageStandardization = new LanguageStandardizationEditor(instance);
    instance.isLoaded = true;
    return instance;
  }

  /**
   * Create instance from Google Drive file
   * @param {string} fileId - Google Drive file ID
   * @param {Object} options - Options
   * @returns {OOXMLSlides} Loaded instance
   */
  static fromGoogleDriveFile(fileId, options = {}) {
    return new OOXMLSlides(fileId, options).load();
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