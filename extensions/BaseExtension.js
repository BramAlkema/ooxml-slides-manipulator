/**
 * BaseExtension - Foundation class for all Slides API extensions
 * 
 * PURPOSE:
 * Provides common functionality, patterns, and utilities that all extensions
 * can inherit from. Reduces boilerplate and ensures consistent behavior.
 * 
 * FEATURES:
 * - Standard lifecycle methods (init, validate, execute, cleanup)
 * - Error handling and logging utilities
 * - Configuration management
 * - XML manipulation helpers
 * - Color and typography utilities
 * - File and blob handling
 * - Caching and performance helpers
 * 
 * AI CONTEXT:
 * Extend this class when creating new Slides API extensions. It provides
 * all the common patterns needed for PowerPoint manipulation, XML processing,
 * and integration with the OOXML framework.
 * 
 * USAGE PATTERNS:
 * - Theme Extensions: Inherit and implement theme manipulation
 * - Content Extensions: Inherit and implement content processing
 * - Validation Extensions: Inherit and implement compliance checking
 */

class BaseExtension {
  
  /**
   * Initialize extension with context and options
   * 
   * @param {Object} context - PowerPoint context (OOXMLSlides instance)
   * @param {Object} options - Extension-specific configuration
   * 
   * AI_USAGE: Always call super(context, options) in your constructor
   */
  constructor(context, options = {}) {
    // Core dependencies
    this.context = context;
    this.options = this._mergeDefaults(options);
    
    // Framework integration
    this._framework = null; // Will be set by ExtensionFramework
    this._initialized = false;
    this._cache = new Map();
    
    // Performance tracking
    this._metrics = {
      startTime: null,
      operations: [],
      errors: []
    };
    
    // XML namespaces (commonly used)
    this.namespaces = {
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    };
  }
  
  /**
   * Get default options for this extension type
   * Override in subclasses to provide extension-specific defaults
   * 
   * @returns {Object} Default options
   */
  getDefaults() {
    return {
      enableLogging: true,
      enableCaching: true,
      validateInput: true,
      trackMetrics: true
    };
  }
  
  /**
   * Initialize the extension
   * Override in subclasses for custom initialization
   * 
   * @returns {Promise<void>}
   * 
   * AI_USAGE: Call this before using any extension functionality
   */
  async init() {
    if (this._initialized) {
      return;
    }
    
    try {
      this._startMetrics();
      this.log('Initializing extension...');
      
      // Validate context and dependencies
      this._validateContext();
      
      // Initialize caching if enabled
      if (this.options.enableCaching) {
        this._initializeCache();
      }
      
      // Custom initialization logic (override in subclasses)
      await this._customInit();
      
      this._initialized = true;
      this.log('Extension initialized successfully');
      
    } catch (error) {
      this.error(`Initialization failed: ${error.message}`);
      throw error;
    }
  }
  
  /**
   * Custom initialization logic (override in subclasses)
   * @protected
   */
  async _customInit() {
    // Override in subclasses
  }
  
  /**
   * Validate extension input
   * Override in subclasses for custom validation
   * 
   * @param {any} input - Input to validate
   * @returns {boolean} Validation result
   * 
   * AI_USAGE: Call this to validate inputs before processing
   */
  validate(input) {
    if (!this.options.validateInput) {
      return true;
    }
    
    try {
      // Basic validation (override for custom logic)
      if (input === null || input === undefined) {
        throw new Error('Input cannot be null or undefined');
      }
      
      return this._customValidate(input);
      
    } catch (error) {
      this.error(`Validation failed: ${error.message}`);
      return false;
    }
  }
  
  /**
   * Custom validation logic (override in subclasses)
   * @protected
   */
  _customValidate(input) {
    return true; // Override in subclasses
  }
  
  /**
   * Execute the main extension functionality
   * Override in subclasses to implement core logic
   * 
   * @param {any} input - Input data
   * @returns {Promise<any>} Execution result
   * 
   * AI_USAGE: This is the main entry point for extension functionality
   */
  async execute(input) {
    await this.init();
    
    if (!this.validate(input)) {
      throw new Error('Input validation failed');
    }
    
    try {
      this.log('Executing extension...');
      const result = await this._customExecute(input);
      this.log('Extension executed successfully');
      return result;
      
    } catch (error) {
      this.error(`Execution failed: ${error.message}`);
      this._trackError(error);
      throw error;
    }
  }
  
  /**
   * Custom execution logic (override in subclasses)
   * @protected
   */
  async _customExecute(input) {
    throw new Error('_customExecute must be implemented in subclasses');
  }
  
  /**
   * Clean up resources
   * Override in subclasses for custom cleanup
   * 
   * AI_USAGE: Call this when done with the extension
   */
  cleanup() {
    this.log('Cleaning up extension...');
    
    // Clear cache
    this._cache.clear();
    
    // Custom cleanup logic (override in subclasses)
    this._customCleanup();
    
    this.log('Extension cleaned up');
  }
  
  /**
   * Custom cleanup logic (override in subclasses)
   * @protected
   */
  _customCleanup() {
    // Override in subclasses
  }
  
  // ===== UTILITY METHODS =====
  
  /**
   * Logging utilities with extension context
   */
  log(message) {
    if (this.options.enableLogging) {
      const prefix = this._framework?.name || this.constructor.name;
      console.log(`[${prefix}] ${message}`);
    }
  }
  
  warn(message) {
    const prefix = this._framework?.name || this.constructor.name;
    console.warn(`[${prefix}] ${message}`);
  }
  
  error(message) {
    const prefix = this._framework?.name || this.constructor.name;
    console.error(`[${prefix}] ${message}`);
  }
  
  /**
   * Cache utilities
   */
  getCached(key) {
    return this._cache.get(key);
  }
  
  setCached(key, value, ttl = null) {
    if (ttl) {
      // Simple TTL implementation
      setTimeout(() => this._cache.delete(key), ttl);
    }
    return this._cache.set(key, value);
  }
  
  clearCache() {
    this._cache.clear();
  }
  
  /**
   * XML manipulation utilities
   */
  
  /**
   * Get XML element with namespace support
   * 
   * @param {Element} parent - Parent XML element
   * @param {string} tagName - Element tag name
   * @param {string} namespace - Namespace prefix
   * @returns {Element|null} Found element
   */
  getXmlElement(parent, tagName, namespace = 'a') {
    if (!parent || !tagName) return null;
    
    try {
      const ns = XmlService.getNamespace(namespace, this.namespaces[namespace]);
      return parent.getChild(tagName, ns);
    } catch (error) {
      this.warn(`Failed to get XML element ${tagName}: ${error.message}`);
      return null;
    }
  }
  
  /**
   * Get multiple XML elements with namespace support
   * 
   * @param {Element} parent - Parent XML element
   * @param {string} tagName - Element tag name
   * @param {string} namespace - Namespace prefix
   * @returns {Array<Element>} Found elements
   */
  getXmlElements(parent, tagName, namespace = 'a') {
    if (!parent || !tagName) return [];
    
    try {
      const ns = XmlService.getNamespace(namespace, this.namespaces[namespace]);
      return parent.getChildren(tagName, ns);
    } catch (error) {
      this.warn(`Failed to get XML elements ${tagName}: ${error.message}`);
      return [];
    }
  }
  
  /**
   * Create XML element with namespace support
   * 
   * @param {string} tagName - Element tag name
   * @param {string} namespace - Namespace prefix
   * @returns {Element} Created element
   */
  createXmlElement(tagName, namespace = 'a') {
    try {
      const ns = XmlService.getNamespace(namespace, this.namespaces[namespace]);
      return XmlService.createElement(tagName, ns);
    } catch (error) {
      this.error(`Failed to create XML element ${tagName}: ${error.message}`);
      throw error;
    }
  }
  
  /**
   * Color utilities
   */
  
  /**
   * Convert hex color to RGB object
   * 
   * @param {string} hex - Hex color (e.g., '#FF0000')
   * @returns {Object} RGB object {r, g, b}
   */
  hexToRgb(hex) {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
      r: parseInt(result[1], 16),
      g: parseInt(result[2], 16),
      b: parseInt(result[3], 16)
    } : null;
  }
  
  /**
   * Convert RGB object to hex color
   * 
   * @param {Object} rgb - RGB object {r, g, b}
   * @returns {string} Hex color
   */
  rgbToHex(rgb) {
    const componentToHex = (c) => {
      const hex = c.toString(16);
      return hex.length === 1 ? "0" + hex : hex;
    };
    
    return `#${componentToHex(rgb.r)}${componentToHex(rgb.g)}${componentToHex(rgb.b)}`;
  }
  
  /**
   * Validate color format
   * 
   * @param {string} color - Color value
   * @returns {boolean} Valid color
   */
  isValidColor(color) {
    if (!color || typeof color !== 'string') return false;
    
    // Check hex format
    if (/^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/.test(color)) return true;
    
    // Check RGB format
    if (/^rgb\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*\)$/.test(color)) return true;
    
    return false;
  }
  
  /**
   * File and blob utilities
   */
  
  /**
   * Get file content from context
   * 
   * @param {string} filename - File path within OOXML
   * @returns {string|null} File content
   */
  getFile(filename) {
    if (!this.context?.core) {
      this.warn('No OOXML core available');
      return null;
    }
    
    return this.context.core.getFile(filename);
  }
  
  /**
   * Set file content in context
   * 
   * @param {string} filename - File path within OOXML
   * @param {string} content - File content
   * @returns {boolean} Success status
   */
  setFile(filename, content) {
    if (!this.context?.core) {
      this.warn('No OOXML core available');
      return false;
    }
    
    try {
      this.context.core.setFile(filename, content);
      return true;
    } catch (error) {
      this.error(`Failed to set file ${filename}: ${error.message}`);
      return false;
    }
  }
  
  // ===== PRIVATE METHODS =====
  
  /**
   * Merge options with defaults
   * @private
   */
  _mergeDefaults(options) {
    return {
      ...this.getDefaults(),
      ...options
    };
  }
  
  /**
   * Validate context and dependencies
   * @private
   */
  _validateContext() {
    if (!this.context) {
      throw new Error('Extension context is required');
    }
    
    // Validate required context methods/properties
    const required = ['core'];
    for (const prop of required) {
      if (!this.context[prop]) {
        throw new Error(`Context missing required property: ${prop}`);
      }
    }
  }
  
  /**
   * Initialize cache
   * @private
   */
  _initializeCache() {
    this._cache = new Map();
    this.log('Cache initialized');
  }
  
  /**
   * Start performance metrics
   * @private
   */
  _startMetrics() {
    if (this.options.trackMetrics) {
      this._metrics.startTime = Date.now();
    }
  }
  
  /**
   * Track error for metrics
   * @private
   */
  _trackError(error) {
    if (this.options.trackMetrics) {
      this._metrics.errors.push({
        message: error.message,
        timestamp: new Date().toISOString()
      });
    }
  }
  
  /**
   * Get performance metrics
   * 
   * @returns {Object} Performance metrics
   */
  getMetrics() {
    if (!this.options.trackMetrics) {
      return { tracking: false };
    }
    
    return {
      ...this._metrics,
      duration: this._metrics.startTime ? Date.now() - this._metrics.startTime : 0,
      errorCount: this._metrics.errors.length
    };
  }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = BaseExtension;
}