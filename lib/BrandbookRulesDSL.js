/**
 * BrandbookRulesDSL - Declarative Brand Compliance Rules Engine
 * 
 * PURPOSE:
 * Provides a declarative JSON-based rules engine for brand compliance validation
 * and auto-fixing. Uses XPath selectors for precise targeting and supports
 * complex validation logic with automatic remediation.
 * 
 * ARCHITECTURE:
 * - JSON Schema validation for rule definitions
 * - XPath-based element selection with namespace handling
 * - Pluggable expectation matchers (hex, equals, regex, range)
 * - Auto-fix capabilities with before/after tracking
 * - Violation reporting with actionable recommendations
 * 
 * AI CONTEXT:
 * This replaces hardcoded brand validation logic with declarative JSON rules.
 * Rules can be authored by brand managers without code changes. The engine
 * provides automatic fixing where possible and detailed violation reports.
 */

/**
 * JSON Schema for Brandbook Rules validation
 */
const BRANDBOOK_RULES_SCHEMA = {
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "$id": "https://example.com/brand-rules.schema.json",
  "type": "object",
  "properties": {
    "profile": { "type": "string" },
    "version": { "type": "string", "default": "1.0.0" },
    "metadata": {
      "type": "object",
      "properties": {
        "name": { "type": "string" },
        "description": { "type": "string" },
        "author": { "type": "string" },
        "created": { "type": "string", "format": "date-time" },
        "updated": { "type": "string", "format": "date-time" }
      }
    },
    "rules": { 
      "type": "array", 
      "items": { "$ref": "#/definitions/rule" },
      "minItems": 1
    }
  },
  "required": ["rules"],
  "definitions": {
    "rule": {
      "type": "object",
      "properties": {
        "id": { 
          "type": "string",
          "pattern": "^[a-zA-Z0-9._-]+$",
          "description": "Unique rule identifier"
        },
        "desc": { 
          "type": "string",
          "description": "Human-readable rule description"
        },
        "category": {
          "type": "string",
          "enum": ["color", "font", "layout", "content", "accessibility"],
          "description": "Rule category for organization"
        },
        "where": { 
          "type": "string",
          "description": "OOXML file path or pattern"
        },
        "xpath": { 
          "type": "string",
          "description": "XPath selector for target elements"
        },
        "expect": { 
          "type": "object",
          "description": "Expectation definition"
        },
        "autofix": { 
          "type": "boolean",
          "default": false,
          "description": "Enable automatic fixing"
        },
        "weight": { 
          "type": "number", 
          "minimum": 0,
          "maximum": 10,
          "default": 1,
          "description": "Rule importance weight"
        },
        "enabled": {
          "type": "boolean",
          "default": true,
          "description": "Rule is active"
        },
        "tags": {
          "type": "array",
          "items": { "type": "string" },
          "description": "Rule tags for filtering"
        }
      },
      "required": ["id", "where", "xpath", "expect"]
    }
  }
};

/**
 * Violation record structure
 */
class BrandViolation {
  constructor(ruleId, where, message, options = {}) {
    this.ruleId = ruleId;
    this.where = where;
    this.message = message;
    this.xpath = options.xpath || '';
    this.expected = options.expected || null;
    this.actual = options.actual || null;
    this.before = options.before || null;
    this.after = options.after || null;
    this.autoFixed = options.autoFixed || false;
    this.weight = options.weight || 1;
    this.category = options.category || 'general';
    this.timestamp = new Date().toISOString();
  }
  
  toJSON() {
    return {
      ruleId: this.ruleId,
      where: this.where,
      message: this.message,
      xpath: this.xpath,
      expected: this.expected,
      actual: this.actual,
      before: this.before,
      after: this.after,
      autoFixed: this.autoFixed,
      weight: this.weight,
      category: this.category,
      timestamp: this.timestamp
    };
  }
}

/**
 * Expectation matchers for different validation types
 */
class ExpectationMatchers {
  
  /**
   * Match hex color values
   */
  static hex(actual, expected) {
    const normalizeHex = (color) => {
      if (!color) return null;
      const hex = color.toString().replace('#', '').toUpperCase();
      return hex.length === 3 ? 
        hex.split('').map(c => c + c).join('') : hex;
    };
    
    const actualNorm = normalizeHex(actual);
    const expectedNorm = normalizeHex(expected);
    
    return {
      matches: actualNorm === expectedNorm,
      actual: actualNorm ? `#${actualNorm}` : actual,
      expected: expectedNorm ? `#${expectedNorm}` : expected
    };
  }
  
  /**
   * Match exact string values
   */
  static equals(actual, expected) {
    return {
      matches: actual === expected,
      actual,
      expected
    };
  }
  
  /**
   * Match using regular expressions
   */
  static regex(actual, pattern, flags = 'i') {
    const regex = new RegExp(pattern, flags);
    return {
      matches: regex.test(actual),
      actual,
      expected: pattern
    };
  }
  
  /**
   * Match numeric ranges
   */
  static range(actual, min, max) {
    const num = parseFloat(actual);
    const matches = !isNaN(num) && num >= min && num <= max;
    return {
      matches,
      actual: num,
      expected: `${min}-${max}`
    };
  }
  
  /**
   * Match against list of allowed values
   */
  static oneOf(actual, allowed) {
    return {
      matches: allowed.includes(actual),
      actual,
      expected: allowed.join(' | ')
    };
  }
  
  /**
   * Match font family names with normalization
   */
  static font(actual, expected) {
    const normalize = (font) => {
      if (!font) return '';
      return font.toLowerCase()
        .replace(/['"]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
    };
    
    const actualNorm = normalize(actual);
    const expectedNorm = normalize(expected);
    
    return {
      matches: actualNorm === expectedNorm,
      actual: actualNorm,
      expected: expectedNorm
    };
  }
}

/**
 * XPath evaluator with namespace handling
 */
class XPathEvaluator {
  
  constructor() {
    this.namespaces = {
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
      'dc': 'http://purl.org/dc/elements/1.1/',
      'dcterms': 'http://purl.org/dc/terms/',
      'dcmitype': 'http://purl.org/dc/dcmitype/',
      'xsi': 'http://www.w3.org/2001/XMLSchema-instance'
    };
  }
  
  /**
   * Evaluate XPath expression against XML content
   */
  evaluate(xmlContent, xpathExpr) {
    try {
      // For Google Apps Script environment, use simple XML parsing
      if (typeof XmlService !== 'undefined') {
        return this._evaluateWithXmlService(xmlContent, xpathExpr);
      }
      
      // For Node.js environment, would use xpath library
      if (typeof require !== 'undefined') {
        return this._evaluateWithXPath(xmlContent, xpathExpr);
      }
      
      // Fallback: simple regex-based matching for basic selectors
      return this._evaluateWithRegex(xmlContent, xpathExpr);
      
    } catch (error) {
      throw new OOXMLError('V042', `XPath evaluation failed: ${error.message}`, {
        xpath: xpathExpr,
        xmlLength: xmlContent.length
      });
    }
  }
  
  /**
   * Extract value from XPath result
   */
  extractValue(result, xpath) {
    if (!result || result.length === 0) {
      return null;
    }
    
    // Handle different result types
    if (Array.isArray(result)) {
      return result.length > 0 ? this._extractSingleValue(result[0]) : null;
    }
    
    return this._extractSingleValue(result);
  }
  
  // PRIVATE METHODS
  
  /**
   * Evaluate using Google Apps Script XmlService
   * @private
   */
  _evaluateWithXmlService(xmlContent, xpathExpr) {
    try {
      const document = XmlService.parse(xmlContent);
      const root = document.getRootElement();
      
      // Set up namespaces
      Object.entries(this.namespaces).forEach(([prefix, uri]) => {
        const namespace = XmlService.getNamespace(prefix, uri);
        root.addNamespace(namespace);
      });
      
      // Simple XPath-like selectors using XmlService
      return this._traverseWithXmlService(root, xpathExpr);
      
    } catch (error) {
      throw new Error(`XmlService evaluation failed: ${error.message}`);
    }
  }
  
  /**
   * Evaluate using Node.js xpath library
   * @private
   */
  _evaluateWithXPath(xmlContent, xpathExpr) {
    // This would require xpath and xmldom packages in Node.js
    // Implementation placeholder for server-side evaluation
    throw new Error('Node.js XPath evaluation not implemented');
  }
  
  /**
   * Fallback regex-based evaluation
   * @private
   */
  _evaluateWithRegex(xmlContent, xpathExpr) {
    // Convert simple XPath to regex patterns
    const patterns = this._xpathToRegex(xpathExpr);
    const results = [];
    
    for (const pattern of patterns) {
      const matches = xmlContent.match(pattern);
      if (matches) {
        results.push(...matches);
      }
    }
    
    return results;
  }
  
  /**
   * Traverse XML using XmlService
   * @private
   */
  _traverseWithXmlService(element, selector) {
    // Simplified XPath-like traversal
    // This is a basic implementation - full XPath would need more complex parsing
    
    if (selector.includes('@')) {
      // Attribute selector
      const [elementPath, attrName] = selector.split('@');
      const targetElements = this._findElements(element, elementPath);
      
      return targetElements.map(el => {
        const attr = el.getAttribute(attrName);
        return attr ? attr.getValue() : null;
      }).filter(v => v !== null);
    } else {
      // Element selector
      const targetElements = this._findElements(element, selector);
      return targetElements.map(el => el.getText()).filter(t => t);
    }
  }
  
  /**
   * Find elements by simplified path
   * @private
   */
  _findElements(root, path) {
    // Very basic element finding - would need full XPath parser for production
    const parts = path.split('/').filter(p => p && p !== '*');
    let current = [root];
    
    for (const part of parts) {
      const next = [];
      for (const element of current) {
        const children = element.getChildren(part, element.getNamespace());
        next.push(...children);
      }
      current = next;
    }
    
    return current;
  }
  
  /**
   * Convert simple XPath to regex patterns
   * @private
   */
  _xpathToRegex(xpath) {
    const patterns = [];
    
    // Handle attribute selectors like @val
    if (xpath.includes('@')) {
      const attrMatch = xpath.match(/@(\w+)/);
      if (attrMatch) {
        const attrName = attrMatch[1];
        patterns.push(new RegExp(`${attrName}="([^"]*)"`, 'g'));
        patterns.push(new RegExp(`${attrName}='([^']*)'`, 'g'));
      }
    }
    
    // Handle element text content
    const elementMatch = xpath.match(/(\w+)$/);
    if (elementMatch && !xpath.includes('@')) {
      const elementName = elementMatch[1];
      patterns.push(new RegExp(`<${elementName}[^>]*>([^<]*)</${elementName}>`, 'g'));
    }
    
    return patterns;
  }
  
  /**
   * Extract single value from XPath result
   * @private
   */
  _extractSingleValue(result) {
    if (typeof result === 'string') {
      return result;
    }
    
    if (result && typeof result.getText === 'function') {
      return result.getText();
    }
    
    if (result && typeof result.getValue === 'function') {
      return result.getValue();
    }
    
    return result;
  }
}

/**
 * Auto-fix engine for brand violations
 */
class BrandAutoFixer {
  
  constructor() {
    this.xpathEvaluator = new XPathEvaluator();
  }
  
  /**
   * Apply auto-fix for a rule violation
   */
  applyFix(xmlContent, rule, violation) {
    try {
      const fixedXml = this._applyRuleFix(xmlContent, rule, violation);
      
      return {
        success: true,
        before: violation.before,
        after: this._extractAfterValue(fixedXml, rule),
        xmlContent: fixedXml
      };
      
    } catch (error) {
      throw new OOXMLError('V044', `Auto-fix failed: ${error.message}`, {
        ruleId: rule.id,
        xpath: rule.xpath
      });
    }
  }
  
  // PRIVATE METHODS
  
  /**
   * Apply specific fix based on rule type
   * @private
   */
  _applyRuleFix(xmlContent, rule, violation) {
    const expectation = rule.expect;
    
    if (expectation.hex) {
      return this._fixHexColor(xmlContent, rule, expectation.hex);
    }
    
    if (expectation.equals) {
      return this._fixTextValue(xmlContent, rule, expectation.equals);
    }
    
    if (expectation.font) {
      return this._fixFontValue(xmlContent, rule, expectation.font);
    }
    
    throw new Error(`No auto-fix handler for expectation type: ${Object.keys(expectation)}`);
  }
  
  /**
   * Fix hex color values
   * @private
   */
  _fixHexColor(xmlContent, rule, expectedHex) {
    const normalizedHex = expectedHex.replace('#', '').toUpperCase();
    
    // Replace color values in various OOXML color elements
    let fixedXml = xmlContent;
    
    // Fix srgbClr elements
    fixedXml = fixedXml.replace(
      /<a:srgbClr\s+val="[^"]*"/g,
      `<a:srgbClr val="${normalizedHex}"`
    );
    
    // Fix other color representations as needed
    fixedXml = fixedXml.replace(
      /rgb\([^)]*\)/g,
      this._hexToRgb(normalizedHex)
    );
    
    return fixedXml;
  }
  
  /**
   * Fix text values
   * @private
   */
  _fixTextValue(xmlContent, rule, expectedValue) {
    // This would need more sophisticated XML manipulation
    // For now, simple regex replacement
    const xpath = rule.xpath;
    
    if (xpath.includes('@')) {
      // Attribute replacement
      const attrMatch = xpath.match(/@(\w+)/);
      if (attrMatch) {
        const attrName = attrMatch[1];
        return xmlContent.replace(
          new RegExp(`${attrName}="[^"]*"`, 'g'),
          `${attrName}="${expectedValue}"`
        );
      }
    } else {
      // Element text replacement
      const elementMatch = xpath.match(/(\w+)$/);
      if (elementMatch) {
        const elementName = elementMatch[1];
        return xmlContent.replace(
          new RegExp(`(<${elementName}[^>]*>)[^<]*(</${elementName}>)`, 'g'),
          `$1${expectedValue}$2`
        );
      }
    }
    
    return xmlContent;
  }
  
  /**
   * Fix font values
   * @private
   */
  _fixFontValue(xmlContent, rule, expectedFont) {
    return xmlContent.replace(
      /typeface="[^"]*"/g,
      `typeface="${expectedFont}"`
    );
  }
  
  /**
   * Convert hex to RGB
   * @private
   */
  _hexToRgb(hex) {
    const r = parseInt(hex.substring(0, 2), 16);
    const g = parseInt(hex.substring(2, 4), 16);
    const b = parseInt(hex.substring(4, 6), 16);
    return `rgb(${r}, ${g}, ${b})`;
  }
  
  /**
   * Extract value after fix
   * @private
   */
  _extractAfterValue(xmlContent, rule) {
    try {
      const results = this.xpathEvaluator.evaluate(xmlContent, rule.xpath);
      return this.xpathEvaluator.extractValue(results, rule.xpath);
    } catch (error) {
      return null;
    }
  }
}

/**
 * Main Brandbook Rules Engine
 */
class BrandbookRulesEngine {
  
  constructor() {
    this.xpathEvaluator = new XPathEvaluator();
    this.autoFixer = new BrandAutoFixer();
    this.matchers = ExpectationMatchers;
  }
  
  /**
   * Validate brandbook rules configuration
   */
  validateRulesConfig(rulesConfig) {
    try {
      // Basic schema validation (would use ajv in Node.js)
      this._validateSchema(rulesConfig, BRANDBOOK_RULES_SCHEMA);
      
      // Additional semantic validation
      this._validateRuleLogic(rulesConfig);
      
      return { valid: true, errors: [] };
      
    } catch (error) {
      throw new OOXMLError('V045', `Rules schema validation failed: ${error.message}`, {
        rulesCount: rulesConfig.rules?.length || 0
      });
    }
  }
  
  /**
   * Execute brand compliance validation
   */
  async executeValidation(manifest, rulesConfig, options = {}) {
    const startTime = Date.now();
    const opts = {
      enableAutoFix: true,
      failFast: false,
      maxViolations: 100,
      ...options
    };
    
    try {
      this.validateRulesConfig(rulesConfig);
      
      const violations = [];
      const fixedFiles = new Map();
      let totalWeight = 0;
      let violationWeight = 0;
      
      for (const rule of rulesConfig.rules) {
        if (!rule.enabled) continue;
        
        totalWeight += rule.weight;
        
        try {
          const ruleViolations = await this._executeRule(
            manifest, rule, fixedFiles, opts.enableAutoFix
          );
          
          violations.push(...ruleViolations);
          violationWeight += ruleViolations.reduce((sum, v) => sum + v.weight, 0);
          
          if (opts.failFast && ruleViolations.length > 0) {
            break;
          }
          
          if (violations.length >= opts.maxViolations) {
            break;
          }
          
        } catch (error) {
          const violation = new BrandViolation(
            rule.id,
            rule.where,
            `Rule execution failed: ${error.message}`,
            { weight: rule.weight, category: rule.category }
          );
          violations.push(violation);
          violationWeight += rule.weight;
        }
      }
      
      // Apply fixes to manifest
      if (opts.enableAutoFix && fixedFiles.size > 0) {
        this._applyFixesToManifest(manifest, fixedFiles);
      }
      
      const score = totalWeight > 0 ? 
        Math.round(((totalWeight - violationWeight) / totalWeight) * 100) : 100;
      
      const duration = Date.now() - startTime;
      
      return {
        version: 'ooxml-json-1',
        profile: rulesConfig.profile || 'default',
        score,
        totalRules: rulesConfig.rules.filter(r => r.enabled).length,
        violations: violations.map(v => v.toJSON()),
        autoFixed: violations.filter(v => v.autoFixed).length,
        duration,
        timestamp: new Date().toISOString()
      };
      
    } catch (error) {
      throw new OOXMLError('V041', `Validation execution failed: ${error.message}`, {
        profile: rulesConfig.profile,
        rulesCount: rulesConfig.rules?.length || 0
      });
    }
  }
  
  /**
   * Get available rule templates
   */
  getRuleTemplates() {
    return {
      colorRule: {
        id: "brand.colors.accent1",
        desc: "Theme accent color must match brand color",
        category: "color",
        where: "ppt/theme/theme1.xml",
        xpath: "//a:accent1//a:srgbClr/@val",
        expect: { hex: "#005BBB" },
        autofix: true,
        weight: 5
      },
      fontRule: {
        id: "brand.fonts.major",
        desc: "Major font must be brand font",
        category: "font",
        where: "ppt/theme/theme1.xml",
        xpath: "//a:majorFont/a:latin/@typeface",
        expect: { font: "Inter" },
        autofix: true,
        weight: 3
      },
      contentRule: {
        id: "brand.content.disclaimer",
        desc: "Slides must contain disclaimer text",
        category: "content",
        where: "ppt/slides/slide*.xml",
        xpath: "//a:t[contains(text(), 'disclaimer')]",
        expect: { regex: ".*confidential.*" },
        autofix: false,
        weight: 2
      }
    };
  }
  
  // PRIVATE METHODS
  
  /**
   * Execute a single rule against the manifest
   * @private
   */
  async _executeRule(manifest, rule, fixedFiles, enableAutoFix) {
    const violations = [];
    const matchingEntries = this._findMatchingEntries(manifest, rule.where);
    
    for (const entry of matchingEntries) {
      if (entry.type !== 'xml') continue;
      
      try {
        let xmlContent = fixedFiles.get(entry.path) || entry.text;
        const results = this.xpathEvaluator.evaluate(xmlContent, rule.xpath);
        const actualValue = this.xpathEvaluator.extractValue(results, rule.xpath);
        
        const matchResult = this._evaluateExpectation(actualValue, rule.expect);
        
        if (!matchResult.matches) {
          const violation = new BrandViolation(
            rule.id,
            entry.path,
            `Expected ${matchResult.expected}, got ${matchResult.actual}`,
            {
              xpath: rule.xpath,
              expected: matchResult.expected,
              actual: matchResult.actual,
              before: matchResult.actual,
              weight: rule.weight,
              category: rule.category
            }
          );
          
          // Apply auto-fix if enabled
          if (enableAutoFix && rule.autofix) {
            try {
              const fixResult = this.autoFixer.applyFix(xmlContent, rule, violation);
              if (fixResult.success) {
                fixedFiles.set(entry.path, fixResult.xmlContent);
                violation.after = fixResult.after;
                violation.autoFixed = true;
              }
            } catch (fixError) {
              // Auto-fix failed, but rule violation still recorded
              violation.message += ` (auto-fix failed: ${fixError.message})`;
            }
          }
          
          violations.push(violation);
        }
        
      } catch (error) {
        const violation = new BrandViolation(
          rule.id,
          entry.path,
          `Rule evaluation failed: ${error.message}`,
          { weight: rule.weight, category: rule.category }
        );
        violations.push(violation);
      }
    }
    
    return violations;
  }
  
  /**
   * Find manifest entries matching file pattern
   * @private
   */
  _findMatchingEntries(manifest, pattern) {
    if (pattern.includes('*')) {
      const regex = new RegExp(pattern.replace(/\*/g, '.*'));
      return manifest.entries.filter(entry => regex.test(entry.path));
    } else {
      return manifest.entries.filter(entry => entry.path === pattern);
    }
  }
  
  /**
   * Evaluate expectation against actual value
   * @private
   */
  _evaluateExpectation(actual, expectation) {
    if (expectation.hex) {
      return this.matchers.hex(actual, expectation.hex);
    }
    
    if (expectation.equals) {
      return this.matchers.equals(actual, expectation.equals);
    }
    
    if (expectation.regex) {
      return this.matchers.regex(actual, expectation.regex, expectation.flags);
    }
    
    if (expectation.range) {
      return this.matchers.range(actual, expectation.range.min, expectation.range.max);
    }
    
    if (expectation.oneOf) {
      return this.matchers.oneOf(actual, expectation.oneOf);
    }
    
    if (expectation.font) {
      return this.matchers.font(actual, expectation.font);
    }
    
    throw new Error(`Unknown expectation type: ${Object.keys(expectation)}`);
  }
  
  /**
   * Apply fixes to the original manifest
   * @private
   */
  _applyFixesToManifest(manifest, fixedFiles) {
    for (const [path, fixedContent] of fixedFiles) {
      const entry = manifest.entries.find(e => e.path === path);
      if (entry && entry.type === 'xml') {
        entry.text = fixedContent;
      }
    }
  }
  
  /**
   * Basic schema validation
   * @private
   */
  _validateSchema(data, schema) {
    // Basic validation - would use ajv or similar in production
    if (!data.rules || !Array.isArray(data.rules)) {
      throw new Error('Rules must be an array');
    }
    
    for (const rule of data.rules) {
      if (!rule.id || !rule.where || !rule.xpath || !rule.expect) {
        throw new Error(`Rule missing required fields: ${JSON.stringify(rule)}`);
      }
    }
  }
  
  /**
   * Validate rule logic
   * @private
   */
  _validateRuleLogic(rulesConfig) {
    const ruleIds = new Set();
    
    for (const rule of rulesConfig.rules) {
      if (ruleIds.has(rule.id)) {
        throw new Error(`Duplicate rule ID: ${rule.id}`);
      }
      ruleIds.add(rule.id);
      
      // Validate XPath syntax (basic check)
      if (!rule.xpath || rule.xpath.trim() === '') {
        throw new Error(`Invalid XPath for rule ${rule.id}`);
      }
      
      // Validate expectation has at least one matcher
      if (!rule.expect || Object.keys(rule.expect).length === 0) {
        throw new Error(`Rule ${rule.id} has no expectations`);
      }
    }
  }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    BrandbookRulesEngine,
    BrandViolation,
    ExpectationMatchers,
    XPathEvaluator,
    BrandAutoFixer,
    BRANDBOOK_RULES_SCHEMA
  };
}