/**
 * OOXML Slides Extension Framework v2
 * 
 * PURPOSE:
 * Provides a systematic, extensible architecture for creating brandbook-compliant
 * PowerPoint extensions. Makes it easy for users to add custom functionality
 * without modifying core code.
 * 
 * ARCHITECTURE:
 * - Base Extension Class: Provides common patterns and utilities
 * - Extension Registry: Manages extension lifecycle and dependencies
 * - Validation System: Ensures extensions follow best practices
 * - Template System: Provides boilerplate for common extension types
 * - Hook System: Allows extensions to integrate with core operations
 * 
 * AI CONTEXT:
 * This is the foundation for all Slides API extensions. When users want to add
 * brandbook compliance features, typography controls, color management, or any
 * PowerPoint-specific functionality, they extend this framework. It provides
 * validation, error handling, and integration patterns.
 * 
 * EXTENSION TYPES:
 * - ThemeExtension: Colors, fonts, styles
 * - ContentExtension: Text, shapes, layout manipulation
 * - ValidationExtension: Brand compliance checking
 * - TemplateExtension: Slide templates and layouts
 * - ExportExtension: Custom export formats and workflows
 */

class ExtensionFramework {
  
  /**
   * Get extension registry (lazy initialization)
   * @private
   */
  static _getRegistry() {
    if (!this.__registry) {
      this.__registry = new Map();
    }
    return this.__registry;
  }
  
  /**
   * Get hooks registry (lazy initialization)
   * @private
   */
  static _getHooks() {
    if (!this.__hooks) {
      this.__hooks = new Map();
    }
    return this.__hooks;
  }
  
  /**
   * Get validators registry (lazy initialization)
   * @private
   */
  static _getValidators() {
    if (!this.__validators) {
      this.__validators = new Map();
    }
    return this.__validators;
  }
  
  /**
   * Extension types and their requirements
   * @private
   */
  static get _EXTENSION_TYPES() {
    return {
      THEME: {
        name: 'Theme Extension',
        requiredMethods: ['applyTheme', 'validateTheme'],
        optionalMethods: ['getThemeInfo', 'createPreview'],
        dependencies: ['OOXMLCoreV2'],
        namespace: 'theme'
      },
      CONTENT: {
        name: 'Content Extension',
        requiredMethods: ['processContent', 'validateContent'],
        optionalMethods: ['getContentInfo', 'transformContent'],
        dependencies: ['OOXMLCoreV2'],
        namespace: 'content'
      },
      VALIDATION: {
        name: 'Validation Extension',
        requiredMethods: ['validate', 'getViolations'],
        optionalMethods: ['autoFix', 'generateReport'],
        dependencies: ['OOXMLCoreV2'],
        namespace: 'validation'
      },
      TEMPLATE: {
        name: 'Template Extension',
        requiredMethods: ['createTemplate', 'applyTemplate'],
        optionalMethods: ['getTemplateInfo', 'customizeTemplate'],
        dependencies: ['OOXMLCoreV2'],
        namespace: 'template'
      },
      EXPORT: {
        name: 'Export Extension',
        requiredMethods: ['export', 'getSupportedFormats'],
        optionalMethods: ['getExportOptions', 'validateExport'],
        dependencies: ['OOXMLCoreV2', 'FFlatePPTXServiceV2'],
        namespace: 'export'
      }
    };
  }
  
  /**
   * Register a new extension
   * 
   * @param {string} name - Extension name (unique identifier)
   * @param {Object} extensionClass - Extension class or factory
   * @param {Object} metadata - Extension metadata and configuration
   * @returns {boolean} Success status
   * 
   * AI_USAGE: Use this to register custom extensions with the framework
   * Example: ExtensionFramework.register('BrandColors', BrandColorExtension, {type: 'THEME'})
   */
  static register(name, extensionClass, metadata = {}) {
    try {
      this._validateExtensionRegistration(name, extensionClass, metadata);
      
      const extensionInfo = {
        name,
        class: extensionClass,
        type: metadata.type || 'CONTENT',
        version: metadata.version || '1.0.0',
        description: metadata.description || '',
        author: metadata.author || 'Unknown',
        dependencies: metadata.dependencies || [],
        configuration: metadata.configuration || {},
        hooks: metadata.hooks || [],
        registeredAt: new Date().toISOString()
      };
      
      this._getRegistry().set(name, extensionInfo);
      this._registerHooks(name, metadata.hooks || []);
      
      console.log(`✅ Extension registered: ${name} (${extensionInfo.type})`);
      return true;
      
    } catch (error) {
      console.error(`❌ Extension registration failed: ${name} - ${error.message}`);
      return false;
    }
  }
  
  /**
   * Create an instance of a registered extension
   * 
   * @param {string} name - Extension name
   * @param {Object} context - PowerPoint context (OOXMLSlides instance)
   * @param {Object} options - Extension-specific options
   * @returns {Object} Extension instance
   * 
   * AI_USAGE: Use this to instantiate extensions for use
   * Example: const brandColors = ExtensionFramework.create('BrandColors', slidesInstance)
   */
  static create(name, context, options = {}) {
    const extensionInfo = this._getRegistry().get(name);
    
    if (!extensionInfo) {
      throw new Error(`Extension not found: ${name}`);
    }
    
    try {
      // Validate dependencies
      this._validateDependencies(extensionInfo.dependencies);
      
      // Create extension instance
      const ExtensionClass = extensionInfo.class;
      const instance = new ExtensionClass(context, options);
      
      // Enhance instance with framework methods
      this._enhanceInstance(instance, extensionInfo);
      
      console.log(`🔧 Extension created: ${name}`);
      return instance;
      
    } catch (error) {
      console.error(`❌ Extension creation failed: ${name} - ${error.message}`);
      throw error;
    }
  }
  
  /**
   * Get information about registered extensions
   * 
   * @param {string} filterType - Optional filter by extension type
   * @returns {Array} List of extension information
   * 
   * AI_USAGE: Use this to discover available extensions
   * Example: const themeExtensions = ExtensionFramework.list('THEME')
   */
  static list(filterType = null) {
    const extensions = Array.from(this._getRegistry().values());
    
    if (filterType) {
      return extensions.filter(ext => ext.type === filterType);
    }
    
    return extensions;
  }
  
  /**
   * Execute hooks for a given event
   * 
   * @param {string} hookName - Hook event name
   * @param {Object} context - Event context
   * @returns {Array} Hook results
   * 
   * AI_USAGE: Used internally by the framework to execute extension hooks
   * Example: ExtensionFramework.executeHooks('beforeApplyTheme', {theme: themeData})
   */
  static executeHooks(hookName, context = {}) {
    const hooks = this._getHooks().get(hookName) || [];
    const results = [];
    
    for (const hook of hooks) {
      try {
        const result = hook.callback(context);
        results.push({
          extension: hook.extension,
          success: true,
          result: result
        });
      } catch (error) {
        console.error(`❌ Hook execution failed: ${hookName} (${hook.extension}) - ${error.message}`);
        results.push({
          extension: hook.extension,
          success: false,
          error: error.message
        });
      }
    }
    
    return results;
  }
  
  /**
   * Get extension templates for common patterns
   * 
   * @param {string} type - Extension type
   * @returns {Object} Template structure and examples
   * 
   * AI_USAGE: Use this to get boilerplate code for new extensions
   * Example: const template = ExtensionFramework.getTemplate('THEME')
   */
  static getTemplate(type) {
    const typeInfo = this._EXTENSION_TYPES[type];
    
    if (!typeInfo) {
      throw new Error(`Unknown extension type: ${type}`);
    }
    
    return {
      type: type,
      info: typeInfo,
      template: this._generateTemplate(typeInfo),
      example: this._generateExample(typeInfo)
    };
  }
  
  // PRIVATE METHODS - Framework implementation details
  
  /**
   * Validate extension registration
   * @private
   */
  static _validateExtensionRegistration(name, extensionClass, metadata) {
    if (!name || typeof name !== 'string') {
      throw new Error('Extension name must be a non-empty string');
    }
    
    if (this._getRegistry().has(name)) {
      throw new Error(`Extension already registered: ${name}`);
    }
    
    if (!extensionClass || typeof extensionClass !== 'function') {
      throw new Error('Extension must be a class or constructor function');
    }
    
    const type = metadata.type || 'CONTENT';
    const typeInfo = this._EXTENSION_TYPES[type];
    
    if (!typeInfo) {
      throw new Error(`Invalid extension type: ${type}`);
    }
    
    // Validate required methods exist on prototype
    const prototype = extensionClass.prototype;
    for (const method of typeInfo.requiredMethods) {
      if (typeof prototype[method] !== 'function') {
        throw new Error(`Extension missing required method: ${method}`);
      }
    }
  }
  
  /**
   * Validate dependencies are available
   * @private
   */
  static _validateDependencies(dependencies) {
    for (const dep of dependencies) {
      if (typeof globalThis[dep] === 'undefined') {
        throw new Error(`Missing dependency: ${dep}`);
      }
    }
  }
  
  /**
   * Register hooks for an extension
   * @private
   */
  static _registerHooks(extensionName, hooks) {
    for (const hook of hooks) {
      if (!this._getHooks().has(hook.name)) {
        this._getHooks().set(hook.name, []);
      }
      
      this._getHooks().get(hook.name).push({
        extension: extensionName,
        callback: hook.callback,
        priority: hook.priority || 0
      });
      
      // Sort hooks by priority
      this._getHooks().get(hook.name).sort((a, b) => b.priority - a.priority);
    }
  }
  
  /**
   * Enhance extension instance with framework methods
   * @private
   */
  static _enhanceInstance(instance, extensionInfo) {
    // Add framework utilities
    instance._framework = {
      name: extensionInfo.name,
      type: extensionInfo.type,
      version: extensionInfo.version,
      
      // Logging utilities
      log: (message) => console.log(`[${extensionInfo.name}] ${message}`),
      warn: (message) => console.warn(`[${extensionInfo.name}] ${message}`),
      error: (message) => console.error(`[${extensionInfo.name}] ${message}`),
      
      // Hook utilities
      executeHooks: (hookName, context) => ExtensionFramework.executeHooks(hookName, context),
      
      // Validation utilities
      validate: (data, schema) => this._validateData(data, schema),
      
      // Configuration access
      config: extensionInfo.configuration
    };
    
    // Add error handling wrapper
    const originalMethods = {};
    const typeInfo = this._EXTENSION_TYPES[extensionInfo.type];
    
    for (const method of [...typeInfo.requiredMethods, ...typeInfo.optionalMethods]) {
      if (typeof instance[method] === 'function') {
        originalMethods[method] = instance[method];
        instance[method] = this._wrapMethod(instance, method, originalMethods[method]);
      }
    }
  }
  
  /**
   * Wrap extension methods with error handling and hooks
   * @private
   */
  static _wrapMethod(instance, methodName, originalMethod) {
    return function(...args) {
      try {
        // Execute before hooks
        instance._framework.executeHooks(`before${methodName}`, { instance, args });
        
        // Execute original method
        const result = originalMethod.apply(instance, args);
        
        // Execute after hooks
        instance._framework.executeHooks(`after${methodName}`, { instance, args, result });
        
        return result;
        
      } catch (error) {
        instance._framework.error(`Method ${methodName} failed: ${error.message}`);
        
        // Execute error hooks
        instance._framework.executeHooks(`error${methodName}`, { instance, args, error });
        
        throw error;
      }
    };
  }
  
  /**
   * Generate template code for extension type
   * @private
   */
  static _generateTemplate(typeInfo) {
    return `
/**
 * ${typeInfo.name} - Custom Extension
 * 
 * AI_USAGE: This extension provides ${typeInfo.name.toLowerCase()} functionality
 */
class Custom${typeInfo.name.replace(/\s/g, '')} {
  
  constructor(context, options = {}) {
    this.context = context;
    this.options = options;
  }
  
  ${typeInfo.requiredMethods.map(method => `
  /**
   * Required method: ${method}
   * TODO: Implement your ${method} logic here
   */
  ${method}() {
    // Implementation goes here
    throw new Error('Method ${method} not implemented');
  }`).join('')}
  
  ${typeInfo.optionalMethods.map(method => `
  /**
   * Optional method: ${method}
   * TODO: Implement your ${method} logic here if needed
   */
  ${method}() {
    // Optional implementation
  }`).join('')}
}

// Register the extension
ExtensionFramework.register('Custom${typeInfo.name.replace(/\s/g, '')}', Custom${typeInfo.name.replace(/\s/g, '')}, {
  type: '${Object.keys(this._EXTENSION_TYPES).find(key => this._EXTENSION_TYPES[key] === typeInfo)}',
  version: '1.0.0',
  description: 'Custom ${typeInfo.name.toLowerCase()} extension',
  author: 'Your Name'
});
`;
  }
  
  /**
   * Generate example implementation
   * @private
   */
  static _generateExample(typeInfo) {
    if (typeInfo.name === 'Theme Extension') {
      return this._generateThemeExample();
    } else if (typeInfo.name === 'Content Extension') {
      return this._generateContentExample();
    } else if (typeInfo.name === 'Validation Extension') {
      return this._generateValidationExample();
    }
    
    return '// Example implementation coming soon';
  }
  
  /**
   * Generate theme extension example
   * @private
   */
  static _generateThemeExample() {
    return `
// Example: Brand Color Extension
class BrandColorExtension {
  constructor(context, options = {}) {
    this.context = context;
    this.brandColors = options.brandColors || {};
  }
  
  applyTheme() {
    const colors = {
      primary: this.brandColors.primary || '#FF0000',
      secondary: this.brandColors.secondary || '#0000FF',
      accent: this.brandColors.accent || '#00FF00'
    };
    
    this._framework.log('Applying brand colors...');
    return this.context.applyCustomColors(colors);
  }
  
  validateTheme() {
    const required = ['primary', 'secondary'];
    const missing = required.filter(color => !this.brandColors[color]);
    
    if (missing.length > 0) {
      throw new Error(\`Missing required brand colors: \${missing.join(', ')}\`);
    }
    
    return true;
  }
}
`;
  }
  
  /**
   * Generate content extension example
   * @private
   */
  static _generateContentExample() {
    return `
// Example: Logo Insertion Extension
class LogoInsertionExtension {
  constructor(context, options = {}) {
    this.context = context;
    this.logoConfig = options.logoConfig || {};
  }
  
  processContent() {
    this._framework.log('Inserting company logo...');
    
    // Add logo to master slides
    return this.context.slides.addLogoToMasters({
      imageId: this.logoConfig.imageId,
      position: this.logoConfig.position || 'top-right',
      size: this.logoConfig.size || 'small'
    });
  }
  
  validateContent() {
    if (!this.logoConfig.imageId) {
      throw new Error('Logo image ID is required');
    }
    
    return true;
  }
}
`;
  }
  
  /**
   * Generate validation extension example
   * @private
   */
  static _generateValidationExample() {
    return `
// Example: Brand Compliance Validator
class BrandComplianceValidator {
  constructor(context, options = {}) {
    this.context = context;
    this.brandGuidelines = options.brandGuidelines || {};
  }
  
  validate() {
    const violations = this.getViolations();
    
    this._framework.log(\`Found \${violations.length} brand violations\`);
    
    return {
      passed: violations.length === 0,
      violations: violations,
      score: this._calculateScore(violations)
    };
  }
  
  getViolations() {
    const violations = [];
    
    // Check font compliance
    const fonts = this.context.theme.getFonts();
    const allowedFonts = this.brandGuidelines.allowedFonts || [];
    
    if (allowedFonts.length > 0) {
      for (const font of fonts) {
        if (!allowedFonts.includes(font)) {
          violations.push({
            type: 'font',
            message: \`Unauthorized font: \${font}\`,
            severity: 'medium'
          });
        }
      }
    }
    
    return violations;
  }
  
  _calculateScore(violations) {
    const totalPoints = 100;
    const deductions = violations.reduce((sum, v) => {
      return sum + (v.severity === 'high' ? 20 : v.severity === 'medium' ? 10 : 5);
    }, 0);
    
    return Math.max(0, totalPoints - deductions);
  }
}
`;
  }
  
  /**
   * Validate data against schema
   * @private
   */
  static _validateData(data, schema) {
    // Simple validation implementation
    // In a full implementation, this would use a proper schema validator
    return true;
  }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = ExtensionFramework;
}