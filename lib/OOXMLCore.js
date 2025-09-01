/**
 * OOXMLCore v2 - Universal OOXML Manipulation Engine
 * 
 * PURPOSE:
 * Universal engine for manipulating Office Open XML formats (.pptx, .docx, .xlsx)
 * Provides format-agnostic ZIP/XML handling with pluggable format-specific layers
 * 
 * ARCHITECTURE:
 * - Core: Format-agnostic ZIP extraction, XML manipulation, recompression
 * - Formats: Pluggable handlers for specific OOXML formats
 * - Dependencies: FFlatePPTXService for reliable ZIP operations
 * 
 * AI CONTEXT:
 * This is the foundation class for all OOXML operations. It handles the universal
 * aspects (ZIP archives, XML files, relationships) while delegating format-specific
 * logic to specialized handlers. Always use this for any OOXML work.
 * 
 * TESTING STRATEGY:
 * Each public method has corresponding unit tests. Tests verify both success paths
 * and error conditions. Mock FFlatePPTXService for isolated testing.
 */

class OOXMLCore {
  
  /**
   * Create a new OOXML manipulation instance
   * 
   * @param {Blob} fileBlob - The OOXML file to manipulate
   * @param {string} formatType - Format type: 'pptx', 'docx', 'xlsx'
   * @param {Object} options - Configuration options
   * @param {boolean} options.validateFormat - Validate file format on creation
   * @param {boolean} options.enableLogging - Enable detailed logging
   * 
   * AI_USAGE: Use this constructor to initialize OOXML processing for any format
   * Example: const core = new OOXMLCore(pptxBlob, 'pptx', {validateFormat: true})
   */
  constructor(fileBlob, formatType = 'pptx', options = {}) {
    // Input validation
    if (!fileBlob || typeof fileBlob.getBytes !== 'function') {
      throw new Error('OOXML_CORE_001: Invalid file blob provided');
    }
    
    if (!this._isValidFormatType(formatType)) {
      throw new Error(`OOXML_CORE_002: Unsupported format type: ${formatType}`);
    }
    
    // Core properties
    this.originalBlob = fileBlob;
    this.formatType = formatType;
    this.options = {
      validateFormat: false,
      enableLogging: true,
      ...options
    };
    
    // State management
    this.files = new Map(); // filename -> content mapping
    this.isExtracted = false;
    this.isModified = false;
    this.metadata = {
      originalSize: fileBlob.getBytes().length,
      fileCount: 0,
      extractedAt: null,
      modifiedAt: null
    };
    
    // Format configuration
    this.formatConfig = this._getFormatConfig(formatType);
    
    // Core OOXML components (universal across formats)
    this.contentTypes = null;
    this.relationships = null;
    this.appProperties = null;
    this.coreProperties = null;
    
    this._log(`OOXMLCore initialized: ${formatType} format, ${this.metadata.originalSize} bytes`);
  }
  
  /**
   * Extract the OOXML archive into individual files
   * 
   * @returns {Promise<OOXMLCore>} This instance for method chaining
   * @throws {Error} If extraction fails or file is invalid
   * 
   * AI_USAGE: Always call this before accessing files. It's safe to call multiple times.
   * Example: await core.extract().then(core => core.getFile('ppt/slides/slide1.xml'))
   */
  async extract() {
    if (this.isExtracted) {
      this._log('Already extracted, skipping');
      return this;
    }
    
    try {
      this._log(`Extracting ${this.formatConfig.name} archive...`);
      
      // Use FFlatePPTXService for reliable extraction
      const extractedFiles = await this._extractViaCloudFunction();
      
      // Convert to internal Map structure
      this.files.clear();
      Object.entries(extractedFiles).forEach(([filename, content]) => {
        this.files.set(filename, content);
      });
      
      this.metadata.fileCount = this.files.size;
      this.metadata.extractedAt = new Date().toISOString();
      this.isExtracted = true;
      
      // Load core OOXML components
      await this._loadCoreComponents();
      
      this._log(`✅ Extracted ${this.files.size} files successfully`);
      return this;
      
    } catch (error) {
      const wrappedError = new Error(`OOXML_CORE_003: Extraction failed: ${error.message}`);
      wrappedError.originalError = error;
      throw wrappedError;
    }
  }
  
  /**
   * Get a file from the extracted archive
   * 
   * @param {string} filename - Path within the OOXML archive
   * @returns {string|null} File content as string, or null if not found
   * 
   * AI_USAGE: Use this to access any file within the OOXML archive
   * Example: const slideXml = core.getFile('ppt/slides/slide1.xml')
   */
  getFile(filename) {
    this._requireExtracted();
    
    const content = this.files.get(filename);
    this._log(`Get file: ${filename} -> ${content ? 'found' : 'not found'}`);
    return content || null;
  }
  
  /**
   * Set/modify a file within the archive
   * 
   * @param {string} filename - Path within the OOXML archive
   * @param {string} content - New file content
   * @returns {OOXMLCore} This instance for method chaining
   * 
   * AI_USAGE: Use this to modify or add files to the OOXML archive
   * Example: core.setFile('ppt/slides/slide1.xml', modifiedXmlContent)
   */
  setFile(filename, content) {
    this._requireExtracted();
    
    if (typeof content !== 'string') {
      throw new Error(`OOXML_CORE_004: File content must be string, got: ${typeof content}`);
    }
    
    const existed = this.files.has(filename);
    this.files.set(filename, content);
    
    if (!existed) {
      this.metadata.fileCount++;
    }
    
    this.isModified = true;
    this.metadata.modifiedAt = new Date().toISOString();
    
    this._log(`Set file: ${filename} (${existed ? 'modified' : 'added'})`);
    return this;
  }
  
  /**
   * List all files in the archive
   * 
   * @returns {Array<string>} Array of filenames
   * 
   * AI_USAGE: Use this to explore the structure of an OOXML file
   * Example: const files = core.listFiles().filter(f => f.includes('slide'))
   */
  listFiles() {
    this._requireExtracted();
    return Array.from(this.files.keys()).sort();
  }
  
  /**
   * Compress the modified files back into an OOXML blob
   * 
   * @returns {Promise<Blob>} New OOXML file blob
   * @throws {Error} If compression fails
   * 
   * AI_USAGE: Call this after making modifications to get the final file
   * Example: const newPptx = await core.compress()
   */
  async compress() {
    this._requireExtracted();
    
    try {
      this._log(`Compressing ${this.files.size} files...`);
      
      // Convert Map to object for cloud function
      const filesObject = {};
      this.files.forEach((content, filename) => {
        filesObject[filename] = content;
      });
      
      // Use FFlatePPTXService for reliable compression
      const compressedBlob = await this._compressViaCloudFunction(filesObject);
      
      // Set proper MIME type
      const finalBlob = compressedBlob.setContentType(this.formatConfig.contentType);
      
      this._log(`✅ Compressed to ${finalBlob.getBytes().length} bytes`);
      return finalBlob;
      
    } catch (error) {
      const wrappedError = new Error(`OOXML_CORE_005: Compression failed: ${error.message}`);
      wrappedError.originalError = error;
      throw wrappedError;
    }
  }
  
  /**
   * Get format information
   * 
   * @returns {Object} Format configuration object
   * 
   * AI_USAGE: Use this to understand what format you're working with
   * Example: if (core.getFormatInfo().type === 'pptx') { ... }
   */
  getFormatInfo() {
    return { ...this.formatConfig };
  }
  
  /**
   * Get metadata about the OOXML file
   * 
   * @returns {Object} Metadata object with size, file count, timestamps
   * 
   * AI_USAGE: Use this for debugging and monitoring
   * Example: console.log(`Processing ${core.getMetadata().fileCount} files`)
   */
  getMetadata() {
    return { ...this.metadata };
  }
  
  /**
   * Check if the archive has been modified
   * 
   * @returns {boolean} True if any files have been modified/added
   * 
   * AI_USAGE: Use this to determine if compression is needed
   * Example: if (core.isModified()) { await core.compress() }
   */
  hasModifications() {
    return this.isModified;
  }
  
  // PRIVATE METHODS - Internal implementation details
  
  /**
   * Validate format type
   * @private
   */
  _isValidFormatType(formatType) {
    const validFormats = ['pptx', 'docx', 'xlsx'];
    return validFormats.includes(formatType);
  }
  
  /**
   * Get format-specific configuration
   * @private
   */
  _getFormatConfig(formatType) {
    const configs = {
      pptx: {
        type: 'pptx',
        name: 'PowerPoint',
        extension: '.pptx',
        contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        mainDocument: 'ppt/presentation.xml',
        documentPath: 'ppt/',
        slidesPath: 'ppt/slides/',
        mastersPath: 'ppt/slideMasters/',
        layoutsPath: 'ppt/slideLayouts/',
        themePath: 'ppt/theme/',
        mediaPath: 'ppt/media/'
      },
      docx: {
        type: 'docx',
        name: 'Word',
        extension: '.docx',
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        mainDocument: 'word/document.xml',
        documentPath: 'word/',
        stylesPath: 'word/styles.xml',
        settingsPath: 'word/settings.xml',
        mediaPath: 'word/media/'
      },
      xlsx: {
        type: 'xlsx',
        name: 'Excel',
        extension: '.xlsx',
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        mainDocument: 'xl/workbook.xml',
        documentPath: 'xl/',
        worksheetsPath: 'xl/worksheets/',
        sharedStringsPath: 'xl/sharedStrings.xml',
        stylesPath: 'xl/styles.xml',
        mediaPath: 'xl/media/'
      }
    };
    
    return configs[formatType] || configs.pptx;
  }
  
  /**
   * Extract files via FFlatePPTXService
   * @private
   */
  async _extractViaCloudFunction() {
    if (typeof FFlatePPTXService === 'undefined') {
      throw new Error('FFlatePPTXService not available');
    }
    
    return FFlatePPTXService.extractFiles(this.originalBlob);
  }
  
  /**
   * Compress files via FFlatePPTXService
   * @private
   */
  async _compressViaCloudFunction(filesObject) {
    if (typeof FFlatePPTXService === 'undefined') {
      throw new Error('FFlatePPTXService not available');
    }
    
    return FFlatePPTXService.compressFiles(filesObject);
  }
  
  /**
   * Load core OOXML components
   * @private
   */
  async _loadCoreComponents() {
    try {
      // Content Types - defines what types of files are in the archive
      this.contentTypes = this.getFile('[Content_Types].xml');
      
      // Relationships - defines how files relate to each other
      this.relationships = this.getFile('_rels/.rels');
      
      // App Properties - application metadata
      this.appProperties = this.getFile('docProps/app.xml');
      
      // Core Properties - document metadata
      this.coreProperties = this.getFile('docProps/core.xml');
      
      this._log('Core OOXML components loaded');
      
    } catch (error) {
      this._log(`Warning: Could not load all core components: ${error.message}`);
    }
  }
  
  /**
   * Require that the archive has been extracted
   * @private
   */
  _requireExtracted() {
    if (!this.isExtracted) {
      throw new Error('OOXML_CORE_006: Must call extract() before accessing files');
    }
  }
  
  /**
   * Log message if logging is enabled
   * @private
   */
  _log(message) {
    if (this.options.enableLogging) {
      console.log(`[OOXMLCore] ${message}`);
    }
  }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = OOXMLCore;
}