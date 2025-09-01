/**
 * ExtensionAutoLoader - Automatic Discovery and Loading of OOXMLSlides Extensions
 * 
 * PURPOSE:
 * Automatically discovers, validates, and registers extensions from the extensions/ folder,
 * then dynamically adds their methods to OOXMLSlides instances.
 * 
 * USAGE:
 * ExtensionAutoLoader.discoverAndLoad(); // Auto-discover all extensions
 * const slides = new OOXMLSlides(fileId); 
 * await slides.createFromPrompt("startup deck"); // Method added by extension
 */

class ExtensionAutoLoader {
  
  /**
   * Get discovered extensions cache (lazy initialization)
   * @private
   */
  static _getExtensions() {
    if (!this.__extensions) {
      this.__extensions = new Map();
    }
    return this.__extensions;
  }
  
  /**
   * Extension discovery patterns
   * @private
   */
  static get _DISCOVERY_PATTERNS() {
    return {
      // File naming patterns for auto-discovery
      extensions: /^([A-Z][a-zA-Z]+)Extension\.js$/,
      methods: /^create([A-Z][a-zA-Z]+)$/,
      templates: /^([a-z_]+)Template\.js$/
    };
  }
  
  /**
   * Discover and load all extensions from the extensions folder
   * 
   * @param {Object} options - Loading options
   * @param {string} options.path - Path to extensions folder (default: 'extensions/')
   * @param {Array} options.include - Specific extensions to include
   * @param {Array} options.exclude - Extensions to exclude
   * @param {boolean} options.autoRegister - Auto-register with OOXMLSlides (default: true)
   * @returns {Promise<Object>} Discovery results
   */
  static async discoverAndLoad(options = {}) {
    ConsoleFormatter.header('🔍 Extension Auto-Discovery');
    
    const config = {
      path: 'extensions/',
      include: [],
      exclude: [],
      autoRegister: true,
      ...options
    };
    
    try {
      // Step 1: Discover extension files
      const discovered = await this._discoverExtensionFiles(config);
      ConsoleFormatter.status('PASS', 'File Discovery', `Found ${discovered.length} extension files`);
      
      // Step 2: Load and validate extensions
      const loaded = await this._loadExtensions(discovered);
      ConsoleFormatter.status('PASS', 'Extension Loading', `Loaded ${loaded.valid.length} valid extensions`);
      
      // Step 3: Auto-register with OOXMLSlides
      if (config.autoRegister) {
        const registered = await this._registerWithOOXMLSlides(loaded.valid);
        ConsoleFormatter.status('PASS', 'Auto-Registration', `Registered ${registered} methods`);
      }
      
      // Step 4: Generate summary
      const summary = this._generateLoadingSummary(loaded);
      ConsoleFormatter.summary('Extension Loading Results', summary);
      
      return {
        success: true,
        discovered: discovered.length,
        loaded: loaded.valid.length,
        failed: loaded.invalid.length,
        registered: config.autoRegister ? loaded.valid.length : 0,
        extensions: loaded.valid,
        errors: loaded.errors
      };
      
    } catch (error) {
      ConsoleFormatter.error('Extension auto-loading failed', error);
      throw error;
    }
  }
  
  /**
   * Discover extension files based on naming patterns
   * @private
   */
  static async _discoverExtensionFiles(config) {
    const files = [];
    
    // In Google Apps Script, we can't scan folders dynamically
    // So we'll use a registry approach with known extensions
    const knownExtensions = [
      'CreateFromPromptExtension',
      'BrandColorsExtension', 
      'BrandComplianceExtension',
      'CustomColorExtension',
      'LanguageStandardizationExtension',
      'NumberingStyleExtension',
      'SuperThemeExtension',
      'TableStyleExtension',
      'ThemeExtension',
      'TypographyExtension',
      'XMLSearchReplaceExtension'
    ];
    
    for (const extName of knownExtensions) {
      // Check if the extension class exists in global scope
      if (typeof globalThis[extName] !== 'undefined') {
        files.push({
          name: extName,
          className: extName,
          type: this._detectExtensionType(extName),
          path: `extensions/${extName}.js`
        });
      }
    }
    
    // Filter by include/exclude lists
    return files.filter(file => {
      if (config.include.length > 0 && !config.include.includes(file.name)) {
        return false;
      }
      if (config.exclude.includes(file.name)) {
        return false;
      }
      return true;
    });
  }
  
  /**
   * Load and validate discovered extensions
   * @private
   */
  static async _loadExtensions(discovered) {
    const results = {
      valid: [],
      invalid: [],
      errors: []
    };
    
    for (const fileInfo of discovered) {
      try {
        ConsoleFormatter.info(`Loading ${fileInfo.name}...`, 'EXTENSION');
        
        // Get the extension class
        const ExtensionClass = globalThis[fileInfo.className];
        
        if (!ExtensionClass) {
          throw new Error(`Extension class ${fileInfo.className} not found`);
        }
        
        // Validate extension structure
        const validation = this._validateExtension(ExtensionClass, fileInfo);
        
        if (validation.valid) {
          // Create extension instance
          const extension = new ExtensionClass();
          
          // Extract methods to add to OOXMLSlides
          const methods = this._extractExtensionMethods(extension);
          
          results.valid.push({
            ...fileInfo,
            class: ExtensionClass,
            instance: extension,
            methods: methods,
            validation: validation
          });
          
          ConsoleFormatter.status('PASS', fileInfo.name, `${methods.length} methods discovered`);
          
        } else {
          results.invalid.push({
            ...fileInfo,
            validation: validation
          });
          
          ConsoleFormatter.status('FAIL', fileInfo.name, validation.errors.join(', '));
        }
        
      } catch (error) {
        results.errors.push({
          extension: fileInfo.name,
          error: error.message
        });
        
        ConsoleFormatter.status('ERROR', fileInfo.name, error.message);
      }
    }
    
    return results;
  }
  
  /**
   * Register extension methods with OOXMLSlides prototype
   * @private
   */
  static async _registerWithOOXMLSlides(validExtensions) {
    let registeredCount = 0;
    
    for (const ext of validExtensions) {
      for (const method of ext.methods) {
        // Add method to OOXMLSlides prototype
        OOXMLSlides.prototype[method.name] = method.implementation;
        
        // Add static methods to OOXMLSlides class
        if (method.static) {
          OOXMLSlides[method.name] = method.implementation;
        }
        
        registeredCount++;
        
        ConsoleFormatter.indent(1, `+ ${method.static ? 'static ' : ''}${method.name}()`);
      }
    }
    
    return registeredCount;
  }
  
  /**
   * Validate extension structure and requirements
   * @private
   */
  static _validateExtension(ExtensionClass, fileInfo) {
    const errors = [];
    
    try {
      // Check if it's a proper class
      if (typeof ExtensionClass !== 'function') {
        errors.push('Not a valid class constructor');
      }
      
      // Check for required methods based on extension type
      const typeInfo = ExtensionFramework._EXTENSION_TYPES[fileInfo.type];
      
      if (typeInfo && typeInfo.requiredMethods) {
        const prototype = ExtensionClass.prototype;
        
        for (const requiredMethod of typeInfo.requiredMethods) {
          if (typeof prototype[requiredMethod] !== 'function') {
            errors.push(`Missing required method: ${requiredMethod}`);
          }
        }
      }
      
      // Check for proper extension metadata
      if (!ExtensionClass.prototype.getExtensionInfo) {
        errors.push('Missing getExtensionInfo() method');
      }
      
    } catch (error) {
      errors.push(`Validation error: ${error.message}`);
    }
    
    return {
      valid: errors.length === 0,
      errors: errors
    };
  }
  
  /**
   * Extract methods from extension that should be added to OOXMLSlides
   * @private
   */
  static _extractExtensionMethods(extensionInstance) {
    const methods = [];
    const proto = Object.getPrototypeOf(extensionInstance);
    const methodNames = Object.getOwnPropertyNames(proto);
    
    for (const methodName of methodNames) {
      if (methodName === 'constructor' || methodName === 'getExtensionInfo') {
        continue;
      }
      
      const method = proto[methodName];
      if (typeof method === 'function') {
        // Check if this method should be exposed to OOXMLSlides
        const shouldExpose = this._shouldExposeMethod(methodName, method);
        
        if (shouldExpose) {
          methods.push({
            name: methodName,
            implementation: method.bind(extensionInstance),
            static: methodName.startsWith('create') || methodName.startsWith('from'),
            description: this._extractMethodDescription(method)
          });
        }
      }
    }
    
    return methods;
  }
  
  /**
   * Determine if a method should be exposed to OOXMLSlides
   * @private
   */
  static _shouldExposeMethod(methodName, method) {
    // Expose methods that start with common patterns
    const exposePatterns = [
      /^create/, /^from/, /^apply/, /^generate/, /^build/,
      /^add/, /^set/, /^update/, /^transform/, /^convert/
    ];
    
    return exposePatterns.some(pattern => pattern.test(methodName));
  }
  
  /**
   * Extract method description from comments or metadata
   * @private
   */
  static _extractMethodDescription(method) {
    const source = method.toString();
    const commentMatch = source.match(/\/\*\*(.*?)\*\//s);
    
    if (commentMatch) {
      const comment = commentMatch[1];
      const descMatch = comment.match(/\*\s*(.+?)(?:\n|\*\/)/);
      if (descMatch) {
        return descMatch[1].trim();
      }
    }
    
    return 'Extension method (no description available)';
  }
  
  /**
   * Detect extension type from name
   * @private
   */
  static _detectExtensionType(name) {
    if (name.includes('Theme') || name.includes('Color') || name.includes('Font')) {
      return 'THEME';
    } else if (name.includes('Template') || name.includes('Generator')) {
      return 'TEMPLATE';
    } else if (name.includes('Validation') || name.includes('Compliance')) {
      return 'VALIDATION';
    } else if (name.includes('Content') || name.includes('Text')) {
      return 'CONTENT';
    } else if (name.includes('Export') || name.includes('Convert')) {
      return 'EXPORT';
    }
    
    return 'CONTENT'; // Default type
  }
  
  /**
   * Generate loading summary
   * @private
   */
  static _generateLoadingSummary(results) {
    return {
      'Extensions Loaded': results.valid.length,
      'Extensions Failed': results.invalid.length,
      'Total Methods': results.valid.reduce((sum, ext) => sum + ext.methods.length, 0),
      'Static Methods': results.valid.reduce((sum, ext) => 
        sum + ext.methods.filter(m => m.static).length, 0
      ),
      'Instance Methods': results.valid.reduce((sum, ext) => 
        sum + ext.methods.filter(m => !m.static).length, 0
      )
    };
  }
  
  /**
   * List all loaded extensions and their methods
   * 
   * @returns {Array} Extension information
   */
  static getLoadedExtensions() {
    return Array.from(this._getExtensions().values()).map(ext => ({
      name: ext.name,
      type: ext.type,
      methods: ext.methods.map(m => ({
        name: m.name,
        static: m.static,
        description: m.description
      }))
    }));
  }
  
  /**
   * Reload all extensions (useful during development)
   * 
   * @returns {Promise<Object>} Reload results
   */
  static async reload() {
    ConsoleFormatter.info('Reloading extensions...', 'RELOAD');
    
    // Clear existing extensions
    this._getExtensions().clear();
    
    // Re-discover and load
    return await this.discoverAndLoad();
  }
}