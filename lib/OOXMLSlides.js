/**
 * OOXMLSlides - PowerPoint-Specific Layer with OOXML JSON Service Integration
 * 
 * PURPOSE:
 * PowerPoint-specific wrapper that provides high-level PowerPoint operations
 * using the OOXML JSON Service backend and integrates with the Extension Framework
 * for easy brandbook compliance and customization.
 * 
 * ARCHITECTURE:
 * - Built on OOXMLJsonService for JSON manifest manipulation
 * - Integrates with ExtensionFramework for modular functionality
 * - Provides PowerPoint-specific operations (slides, themes, layouts)
 * - Supports server-side operations for complex transformations
 * - Supports session-based large file handling
 * - Includes built-in validation and error handling
 * 
 * AI CONTEXT:
 * This is the next-generation PowerPoint manipulation class. It uses JSON manifests
 * for easy XML editing, server-side operations for complex transformations, and
 * seamlessly integrates with the extension framework.
 * 
 * USAGE:
 * const slides = new OOXMLSlides(fileId);
 * await slides.load();
 * await slides.applyBrandColors({primary: '#FF0000'});
 * await slides.validateCompliance(brandGuidelines);
 * const result = await slides.save();
 */

class OOXMLSlides {
  
  /**
   * Initialize PowerPoint manipulation with extension support
   * 
   * @param {string} fileId - Google Drive file ID
   * @param {Object} options - Configuration options
   * @param {boolean} options.createBackup - Create backup before modifications
   * @param {boolean} options.autoSave - Automatically save changes
   * @param {boolean} options.enableExtensions - Enable extension framework
   * @param {Array} options.preloadExtensions - Extensions to load automatically
   * 
   * AI_USAGE: Create instance for PowerPoint manipulation with extensions
   * Example: const slides = new OOXMLSlides(fileId, {enableExtensions: true})
   */
  constructor(fileId, options = {}) {
    this.fileId = fileId;
    this.options = {
      createBackup: true,
      autoSave: true,
      enableExtensions: true,
      preloadExtensions: [],
      trackMetrics: true,
      enableLogging: true,
      ...options
    };
    
    // Core components  
    this.fileHandler = new FileHandler();
    this.manifest = null; // OOXML JSON manifest
    this.session = null; // Session for large files
    this.isLoaded = false;
    this.backupId = null;
    
    // Extension management
    this.extensions = new Map();
    this.extensionResults = new Map();
    
    // Performance tracking
    this.metrics = {
      loadTime: null,
      operations: [],
      extensionCalls: []
    };
    
    this._log('OOXMLSlides initialized with OOXML JSON Service integration');
  }
  
  /**
   * Load and parse the PowerPoint file
   * 
   * @returns {Promise<OOXMLSlides>} This instance for method chaining
   * @throws {Error} If file loading fails
   * 
   * AI_USAGE: Always call this first before other operations
   * Example: await slides.load()
   */
  async load() {
    const startTime = Date.now();
    
    try {
      this._log('Loading PowerPoint file via OOXML JSON Service...');
      
      // Validate file access
      if (!this.fileHandler.validateAccess(this.fileId)) {
        throw new Error(`Cannot access file: ${this.fileId}`);
      }
      
      // Create backup if requested
      if (this.options.createBackup) {
        this.backupId = this.fileHandler.createBackup(this.fileId);
        this._log(`Backup created: ${this.backupId}`);
      }
      
      // Check file size and decide on approach
      const fileBlob = this.fileHandler.loadFile(this.fileId);
      const sizeBytes = fileBlob.getBytes().length;
      const sizeMB = sizeBytes / (1024 * 1024);
      
      // Use session-based approach for large files
      const useSession = sizeMB > 25;
      
      if (useSession) {
        this._log(`Large file detected (${sizeMB.toFixed(1)}MB), using session approach`);
        this.session = await OOXMLJsonService.createSession();
        await OOXMLJsonService.uploadToSession(this.session.uploadUrl, fileBlob);
        this.manifest = await OOXMLJsonService.unwrap(fileBlob, { useSession: true });
      } else {
        this._log(`Small file (${sizeMB.toFixed(1)}MB), using direct approach`);
        this.manifest = await OOXMLJsonService.unwrap(fileBlob);
      }
      
      this._log(`Unwrapped to ${this.manifest.entries.length} entries (${this.manifest.kind} format)`);
      
      // Load extensions if enabled
      if (this.options.enableExtensions) {
        await this._loadExtensions();
      }
      
      this.isLoaded = true;
      this.metrics.loadTime = Date.now() - startTime;
      
      this._log(`PowerPoint file loaded successfully (${this.metrics.loadTime}ms)`);
      return this;
      
    } catch (error) {
      this._error(`Failed to load PowerPoint file: ${error.message}`);
      throw error;
    }
  }
  
  /**
   * Save the modified PowerPoint file
   * 
   * @param {Object} options - Save options
   * @param {string} options.filename - Optional new filename
   * @param {boolean} options.overwrite - Overwrite original file
   * @returns {Promise<Object>} Save result with file information
   * 
   * AI_USAGE: Call this after making modifications
   * Example: const result = await slides.save()
   */
  async save(options = {}) {
    this._ensureLoaded();
    
    try {
      this._log('Saving PowerPoint file via OOXML JSON Service...');
      
      // Prepare save options
      const saveOptions = {
        filename: options.filename || 'presentation.pptx',
        overwrite: options.overwrite !== false,
        ...options
      };
      
      let savedFileId;
      
      // Use session-based save for large files
      if (this.session) {
        this._log('Using session-based save for large file');
        const result = await OOXMLJsonService.rewrap(this.manifest, {
          gcsIn: this.session.gcsIn,
          filename: saveOptions.filename
        });
        
        // Download from GCS and save to Drive
        if (result.gcsOut) {
          const downloadResponse = await UrlFetchApp.fetch(this.session.downloadUrl);
          const modifiedBlob = downloadResponse.getBlob();
          modifiedBlob.setName(saveOptions.filename);
          
          const savedFile = this.fileHandler.saveFile(
            this.fileId,
            modifiedBlob,
            saveOptions
          );
          savedFileId = savedFile.getId();
        }
      } else {
        // Direct save for small files
        this._log('Using direct save for small file');
        savedFileId = await OOXMLJsonService.rewrap(this.manifest, saveOptions);
      }
      
      // Get file information
      const savedFile = DriveApp.getFileById(savedFileId);
      
      const result = {
        success: true,
        fileId: savedFileId,
        filename: savedFile.getName(),
        size: savedFile.getSize(),
        url: savedFile.getUrl(),
        timestamp: new Date().toISOString(),
        metrics: this.getMetrics()
      };
      
      this._log(`PowerPoint file saved successfully: ${result.filename}`);
      return result;
      
    } catch (error) {
      this._error(`Failed to save PowerPoint file: ${error.message}`);
      throw error;
    }
  }
  
  // ===== BUILT-IN POWERPOINT OPERATIONS =====
  
  /**
   * Apply brand colors using the BrandColors extension
   * 
   * @param {Object} brandColors - Brand color configuration
   * @param {string} brandColors.primary - Primary brand color
   * @param {string} brandColors.secondary - Secondary brand color
   * @param {string} brandColors.accent - Accent color
   * @returns {Promise<Object>} Application result
   * 
   * AI_USAGE: Apply consistent brand colors across presentation
   * Example: await slides.applyBrandColors({primary: '#FF0000', secondary: '#0000FF'})
   */
  async applyBrandColors(brandColors, options = {}) {
    this._ensureLoaded();
    
    const extension = await this._getOrCreateExtension('BrandColors', {
      brandColors,
      ...options
    });
    
    const result = await extension.execute();
    this._trackExtensionCall('applyBrandColors', result);
    
    return result;
  }
  
  /**
   * Validate brand compliance using the BrandCompliance extension
   * 
   * @param {Object} brandGuidelines - Brand guidelines configuration
   * @param {Array} brandGuidelines.allowedFonts - Allowed font families
   * @param {Array} brandGuidelines.allowedColors - Allowed color palette
   * @param {Object} brandGuidelines.logoRequirements - Logo requirements
   * @returns {Promise<Object>} Compliance report
   * 
   * AI_USAGE: Check presentation against brand guidelines
   * Example: const report = await slides.validateCompliance(guidelines)
   */
  async validateCompliance(brandGuidelines, options = {}) {
    this._ensureLoaded();
    
    const extension = await this._getOrCreateExtension('BrandCompliance', {
      brandGuidelines,
      ...options
    });
    
    const result = await extension.execute();
    this._trackExtensionCall('validateCompliance', result);
    
    return result;
  }
  
  /**
   * Apply advanced typography and kerning
   * 
   * @param {Object} typographyConfig - Typography configuration
   * @returns {Promise<Object>} Typography application result
   * 
   * AI_USAGE: Apply professional typography settings
   * Example: await slides.applyTypography({kerning: 'tight', tracking: 0.02})
   */
  async applyTypography(typographyConfig, options = {}) {
    this._ensureLoaded();
    
    // Use existing TypographyEditor or create custom extension
    const editor = new TypographyEditor(this);
    const result = editor.applyAdvancedKerning(typographyConfig);
    
    this._trackOperation('applyTypography', result);
    return result;
  }
  
  /**
   * Apply custom table styles
   * 
   * @param {Object} tableStyles - Table style configuration
   * @returns {Promise<Object>} Table style application result
   * 
   * AI_USAGE: Apply consistent table formatting
   * Example: await slides.applyTableStyles({headerStyle: 'bold', stripes: true})
   */
  async applyTableStyles(tableStyles, options = {}) {
    this._ensureLoaded();
    
    // Use existing TableStyleEditor
    const editor = new TableStyleEditor(this);
    const result = editor.applyTableStyles(tableStyles);
    
    this._trackOperation('applyTableStyles', result);
    return result;
  }
  
  // ===== EXTENSION FRAMEWORK INTEGRATION =====
  
  /**
   * Register a custom extension
   * 
   * @param {string} name - Extension name
   * @param {Function} extensionClass - Extension class
   * @param {Object} metadata - Extension metadata
   * @returns {boolean} Registration success
   * 
   * AI_USAGE: Register custom extensions for specific needs
   * Example: slides.registerExtension('CustomBrand', MyExtension, {type: 'THEME'})
   */
  registerExtension(name, extensionClass, metadata = {}) {
    return ExtensionFramework.register(name, extensionClass, metadata);
  }
  
  /**
   * Use a specific extension
   * 
   * @param {string} extensionName - Name of registered extension
   * @param {Object} options - Extension options
   * @returns {Promise<Object>} Extension execution result
   * 
   * AI_USAGE: Execute custom extensions
   * Example: await slides.useExtension('CustomValidation', {rules: myRules})
   */
  async useExtension(extensionName, options = {}) {
    this._ensureLoaded();
    
    const extension = await this._getOrCreateExtension(extensionName, options);
    const result = await extension.execute();
    
    this._trackExtensionCall(extensionName, result);
    return result;
  }
  
  /**
   * Get list of available extensions
   * 
   * @param {string} filterType - Optional filter by extension type
   * @returns {Array} Available extensions
   * 
   * AI_USAGE: Discover available extensions
   * Example: const themeExtensions = slides.listExtensions('THEME')
   */
  listExtensions(filterType = null) {
    return ExtensionFramework.list(filterType);
  }
  
  /**
   * Get extension template for custom development
   * 
   * @param {string} extensionType - Extension type (THEME, VALIDATION, etc.)
   * @returns {Object} Template structure and examples
   * 
   * AI_USAGE: Get boilerplate for creating custom extensions
   * Example: const template = slides.getExtensionTemplate('VALIDATION')
   */
  getExtensionTemplate(extensionType) {
    return ExtensionFramework.getTemplate(extensionType);
  }
  
  // ===== CORE OOXML OPERATIONS =====
  
  /**
   * Get file content from the OOXML archive
   * 
   * @param {string} filename - File path within OOXML
   * @returns {string|null} File content
   * 
   * AI_USAGE: Access any file within the PowerPoint archive
   * Example: const slideXml = slides.getFile('ppt/slides/slide1.xml')
   */
  getFile(filename) {
    this._ensureLoaded();
    return this.core.getFile(filename);
  }
  
  /**
   * Set file content in the OOXML archive
   * 
   * @param {string} filename - File path within OOXML
   * @param {string} content - New file content
   * @returns {OOXMLSlides} This instance for method chaining
   * 
   * AI_USAGE: Modify files within the PowerPoint archive
   * Example: slides.setFile('ppt/slides/slide1.xml', modifiedXml)
   */
  setFile(filename, content) {
    this._ensureLoaded();
    this.core.setFile(filename, content);
    return this;
  }
  
  /**
   * List all files in the OOXML archive
   * 
   * @returns {Array<string>} List of filenames
   * 
   * AI_USAGE: Explore the PowerPoint file structure
   * Example: const files = slides.listFiles()
   */
  listFiles() {
    this._ensureLoaded();
    return this.core.listFiles();
  }
  
  /**
   * Get PowerPoint metadata and information
   * 
   * @returns {Object} Presentation metadata
   * 
   * AI_USAGE: Get information about the presentation
   * Example: const info = slides.getInfo()
   */
  getInfo() {
    this._ensureLoaded();
    
    const coreMetadata = this.core.getMetadata();
    const files = this.listFiles();
    
    // Count slides
    const slideCount = files.filter(f => 
      f.includes('ppt/slides/slide') && f.endsWith('.xml')
    ).length;
    
    // Count masters and layouts
    const masterCount = files.filter(f => 
      f.includes('ppt/slideMasters/') && f.endsWith('.xml')
    ).length;
    
    const layoutCount = files.filter(f => 
      f.includes('ppt/slideLayouts/') && f.endsWith('.xml')
    ).length;
    
    return {
      fileId: this.fileId,
      format: this.core.getFormatInfo(),
      slides: {
        count: slideCount,
        masters: masterCount,
        layouts: layoutCount
      },
      files: {
        total: coreMetadata.fileCount,
        modified: this.core.hasModifications()
      },
      extensions: {
        loaded: this.extensions.size,
        available: ExtensionFramework.list().length
      },
      metadata: coreMetadata
    };
  }
  
  /**
   * Get performance metrics
   * 
   * @returns {Object} Performance and usage metrics
   * 
   * AI_USAGE: Monitor performance and usage
   * Example: const metrics = slides.getMetrics()
   */
  getMetrics() {
    return {
      loadTime: this.metrics.loadTime,
      operations: this.metrics.operations.length,
      extensionCalls: this.metrics.extensionCalls.length,
      coreMetrics: this.core?.getMetadata(),
      extensionMetrics: Array.from(this.extensions.values()).map(ext => ({
        name: ext._framework?.name,
        metrics: ext.getMetrics?.()
      }))
    };
  }
  
  // ===== PRIVATE METHODS =====
  
  /**
   * Ensure the presentation is loaded
   * @private
   */
  _ensureLoaded() {
    if (!this.isLoaded) {
      throw new Error('Presentation not loaded. Call load() first.');
    }
  }
  
  /**
   * Load configured extensions
   * @private
   */
  async _loadExtensions() {
    this._log('Loading extensions...');
    
    for (const extensionName of this.options.preloadExtensions) {
      try {
        const extension = ExtensionFramework.create(extensionName, this);
        this.extensions.set(extensionName, extension);
        this._log(`Loaded extension: ${extensionName}`);
      } catch (error) {
        this._warn(`Failed to load extension ${extensionName}: ${error.message}`);
      }
    }
  }
  
  /**
   * Get or create extension instance
   * @private
   */
  async _getOrCreateExtension(name, options = {}) {
    if (this.extensions.has(name)) {
      return this.extensions.get(name);
    }
    
    try {
      const extension = ExtensionFramework.create(name, this, options);
      this.extensions.set(name, extension);
      return extension;
    } catch (error) {
      this._error(`Failed to create extension ${name}: ${error.message}`);
      throw error;
    }
  }
  
  /**
   * Track extension calls for metrics
   * @private
   */
  _trackExtensionCall(extensionName, result) {
    if (this.options.trackMetrics) {
      this.metrics.extensionCalls.push({
        extension: extensionName,
        timestamp: new Date().toISOString(),
        success: !result.error,
        result: result
      });
    }
  }
  
  /**
   * Track operations for metrics
   * @private
   */
  _trackOperation(operationName, result) {
    if (this.options.trackMetrics) {
      this.metrics.operations.push({
        operation: operationName,
        timestamp: new Date().toISOString(),
        success: !result.error,
        result: result
      });
    }
  }
  
  /**
   * Logging utilities
   * @private
   */
  _log(message) {
    if (this.options.enableLogging) {
      console.log(`[OOXMLSlidesV2] ${message}`);
    }
  }
  
  _warn(message) {
    console.warn(`[OOXMLSlidesV2] ${message}`);
  }
  
  _error(message) {
    console.error(`[OOXMLSlidesV2] ${message}`);
  }
  
  /**
   * Cleanup resources
   * 
   * AI_USAGE: Call when done with the instance
   * Example: slides.cleanup()
   */
  cleanup() {
    this._log('Cleaning up resources...');
    
    // Cleanup extensions
    for (const extension of this.extensions.values()) {
      if (extension.cleanup) {
        extension.cleanup();
      }
    }
    this.extensions.clear();
    
    // Clear results and metrics
    this.extensionResults.clear();
    this.metrics.operations = [];
    this.metrics.extensionCalls = [];
    
    this._log('Cleanup completed');
  }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = OOXMLSlides;
}