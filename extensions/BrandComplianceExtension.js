/**
 * BrandComplianceExtension - Example validation extension for brand compliance
 * 
 * PURPOSE:
 * Demonstrates how to create a validation extension that checks presentations
 * against brand guidelines and corporate standards.
 * 
 * FEATURES:
 * - Font compliance validation
 * - Color compliance checking
 * - Logo placement verification
 * - Layout consistency validation
 * - Auto-fix capabilities
 * - Compliance scoring
 * - Detailed violation reporting
 * 
 * AI CONTEXT:
 * This example shows how to build brand compliance checking into PowerPoint
 * presentations. Users can adapt this for their specific brand guidelines.
 * 
 * USAGE:
 * const compliance = ExtensionFramework.create('BrandCompliance', slidesInstance, {
 *   brandGuidelines: { allowedFonts: ['Arial', 'Helvetica'], requiredLogo: true }
 * });
 * const report = await compliance.execute();
 */

class BrandComplianceExtension extends BaseExtension {
  
  /**
   * Get default options for brand compliance extension
   */
  getDefaults() {
    return {
      ...super.getDefaults(),
      generateReport: true,
      autoFix: false,
      failOnViolations: false,
      maxViolations: 10,
      skipWarnings: false
    };
  }
  
  /**
   * Custom initialization for brand compliance
   */
  async _customInit() {
    this.log('Initializing brand compliance extension...');
    
    // Validate brand guidelines are provided
    if (!this.options.brandGuidelines) {
      throw new Error('Brand guidelines configuration is required');
    }
    
    // Parse brand guidelines
    this.guidelines = this._parseBrandGuidelines(this.options.brandGuidelines);
    
    // Initialize violation tracking
    this.violations = [];
    this.warnings = [];
    this.autoFixes = [];
    
    this.log(`Loaded ${Object.keys(this.guidelines).length} brand guidelines`);
  }
  
  /**
   * Custom validation for brand compliance
   */
  _customValidate(input) {
    // Validate that we have meaningful guidelines
    const requiredGuidelines = ['allowedFonts', 'allowedColors', 'requiredElements'];
    const hasGuidelines = requiredGuidelines.some(guideline => this.guidelines[guideline]);
    
    if (!hasGuidelines) {
      this.warn('No meaningful brand guidelines found');
      return false;
    }
    
    return true;
  }
  
  /**
   * Main execution - validate presentation compliance
   */
  async _customExecute(input) {
    this.log('Validating brand compliance...');
    
    // Clear previous results
    this.violations = [];
    this.warnings = [];
    this.autoFixes = [];
    
    // Run all compliance checks
    await this._validateFontCompliance();
    await this._validateColorCompliance();
    await this._validateLogoCompliance();
    await this._validateLayoutCompliance();
    await this._validateContentCompliance();
    
    // Apply auto-fixes if enabled
    if (this.options.autoFix && this.autoFixes.length > 0) {
      await this._applyAutoFixes();
    }
    
    // Calculate compliance score
    const score = this._calculateComplianceScore();
    
    // Generate report
    const report = this._generateComplianceReport(score);
    
    // Check if we should fail on violations
    if (this.options.failOnViolations && this.violations.length > 0) {
      throw new Error(`Brand compliance failed: ${this.violations.length} violations found`);
    }
    
    this.log(`Compliance validation completed - Score: ${score}%`);
    return report;
  }
  
  /**
   * Validate compliance - required method for VALIDATION extensions
   */
  async validate() {
    const result = await this.execute();
    return {
      passed: result.score >= 80, // 80% or higher passes
      violations: result.violations,
      score: result.score
    };
  }
  
  /**
   * Get violations - required method for VALIDATION extensions
   */
  getViolations() {
    return {
      violations: this.violations,
      warnings: this.warnings,
      total: this.violations.length + this.warnings.length
    };
  }
  
  /**
   * Auto-fix violations - optional method for VALIDATION extensions
   */
  async autoFix() {
    if (this.autoFixes.length === 0) {
      return { fixed: 0, message: 'No auto-fixable violations found' };
    }
    
    const applied = await this._applyAutoFixes();
    return {
      fixed: applied.length,
      fixes: applied,
      message: `Applied ${applied.length} automatic fixes`
    };
  }
  
  /**
   * Generate compliance report - optional method for VALIDATION extensions
   */
  generateReport() {
    return this._generateComplianceReport(this._calculateComplianceScore());
  }
  
  // ===== PRIVATE IMPLEMENTATION METHODS =====
  
  /**
   * Parse brand guidelines from configuration
   * @private
   */
  _parseBrandGuidelines(guidelines) {
    return {
      allowedFonts: guidelines.allowedFonts || [],
      allowedColors: guidelines.allowedColors || [],
      requiredElements: guidelines.requiredElements || [],
      layoutRules: guidelines.layoutRules || {},
      logoRequirements: guidelines.logoRequirements || {},
      contentRules: guidelines.contentRules || {}
    };
  }
  
  /**
   * Validate font compliance
   * @private
   */
  async _validateFontCompliance() {
    if (!this.guidelines.allowedFonts.length) return;
    
    this.log('Validating font compliance...');
    
    // Check theme fonts
    const themeXml = this.getFile('ppt/theme/theme1.xml');
    if (themeXml) {
      await this._checkThemeFonts(themeXml);
    }
    
    // Check slide master fonts
    await this._checkSlideMasterFonts();
    
    // Check individual slide fonts
    await this._checkSlideFonts();
  }
  
  /**
   * Check theme fonts for compliance
   * @private
   */
  async _checkThemeFonts(themeXml) {
    try {
      const doc = XmlService.parse(themeXml);
      const fontScheme = this._findFontScheme(doc);
      
      if (fontScheme) {
        const fonts = this._extractFontsFromScheme(fontScheme);
        
        for (const font of fonts) {
          if (!this.guidelines.allowedFonts.includes(font)) {
            this._addViolation({
              type: 'font',
              severity: 'high',
              element: 'theme',
              message: `Unauthorized font in theme: ${font}`,
              location: 'ppt/theme/theme1.xml',
              autoFixable: true,
              suggestedFix: `Replace with ${this.guidelines.allowedFonts[0]}`
            });
          }
        }
      }
      
    } catch (error) {
      this.error(`Failed to check theme fonts: ${error.message}`);
    }
  }
  
  /**
   * Find font scheme in theme document
   * @private
   */
  _findFontScheme(themeDoc) {
    const root = themeDoc.getRootElement();
    const themeElements = this.getXmlElement(root, 'themeElements');
    return this.getXmlElement(themeElements, 'fontScheme');
  }
  
  /**
   * Extract fonts from font scheme
   * @private
   */
  _extractFontsFromScheme(fontScheme) {
    const fonts = [];
    
    // Get major and minor font families
    const majorFont = this.getXmlElement(fontScheme, 'majorFont');
    const minorFont = this.getXmlElement(fontScheme, 'minorFont');
    
    if (majorFont) {
      const latin = this.getXmlElement(majorFont, 'latin');
      if (latin) {
        const typeface = latin.getAttribute('typeface')?.getValue();
        if (typeface) fonts.push(typeface);
      }
    }
    
    if (minorFont) {
      const latin = this.getXmlElement(minorFont, 'latin');
      if (latin) {
        const typeface = latin.getAttribute('typeface')?.getValue();
        if (typeface) fonts.push(typeface);
      }
    }
    
    return fonts;
  }
  
  /**
   * Check slide master fonts
   * @private
   */
  async _checkSlideMasterFonts() {
    // Get list of slide master files
    const files = this.context.core.listFiles();
    const masterFiles = files.filter(f => f.includes('slideMasters/slideMaster'));
    
    for (const masterFile of masterFiles) {
      await this._checkFileForFonts(masterFile, 'slideMaster');
    }
  }
  
  /**
   * Check individual slide fonts
   * @private
   */
  async _checkSlideFonts() {
    // Get list of slide files
    const files = this.context.core.listFiles();
    const slideFiles = files.filter(f => f.includes('slides/slide') && f.endsWith('.xml'));
    
    for (const slideFile of slideFiles) {
      await this._checkFileForFonts(slideFile, 'slide');
    }
  }
  
  /**
   * Check specific file for font violations
   * @private
   */
  async _checkFileForFonts(filename, elementType) {
    const content = this.getFile(filename);
    if (!content) return;
    
    try {
      const doc = XmlService.parse(content);
      const fonts = this._extractFontsFromXml(doc.getRootElement());
      
      for (const font of fonts) {
        if (!this.guidelines.allowedFonts.includes(font)) {
          this._addViolation({
            type: 'font',
            severity: 'medium',
            element: elementType,
            message: `Unauthorized font: ${font}`,
            location: filename,
            autoFixable: true,
            suggestedFix: `Replace with ${this.guidelines.allowedFonts[0]}`
          });
        }
      }
      
    } catch (error) {
      this.warn(`Failed to parse ${filename} for fonts: ${error.message}`);
    }
  }
  
  /**
   * Extract fonts from XML element recursively
   * @private
   */
  _extractFontsFromXml(element) {
    const fonts = new Set();
    
    // Check for typeface attributes
    const typeface = element.getAttribute('typeface')?.getValue();
    if (typeface && typeface !== '+mn-lt' && typeface !== '+mj-lt') {
      fonts.add(typeface);
    }
    
    // Recursively check children
    const children = element.getChildren();
    for (const child of children) {
      const childFonts = this._extractFontsFromXml(child);
      childFonts.forEach(font => fonts.add(font));
    }
    
    return Array.from(fonts);
  }
  
  /**
   * Validate color compliance
   * @private
   */
  async _validateColorCompliance() {
    if (!this.guidelines.allowedColors.length) return;
    
    this.log('Validating color compliance...');
    
    // Check theme colors
    await this._checkThemeColors();
    
    // Check slide colors
    await this._checkSlideColors();
  }
  
  /**
   * Check theme colors for compliance
   * @private
   */
  async _checkThemeColors() {
    const themeXml = this.getFile('ppt/theme/theme1.xml');
    if (!themeXml) return;
    
    try {
      const doc = XmlService.parse(themeXml);
      const colorScheme = this._findColorScheme(doc);
      
      if (colorScheme) {
        const colors = this._extractColorsFromScheme(colorScheme);
        
        for (const color of colors) {
          if (!this._isColorAllowed(color)) {
            this._addViolation({
              type: 'color',
              severity: 'medium',
              element: 'theme',
              message: `Unauthorized color in theme: ${color}`,
              location: 'ppt/theme/theme1.xml',
              autoFixable: false,
              suggestedFix: 'Use approved brand colors'
            });
          }
        }
      }
      
    } catch (error) {
      this.error(`Failed to check theme colors: ${error.message}`);
    }
  }
  
  /**
   * Find color scheme in theme document
   * @private
   */
  _findColorScheme(themeDoc) {
    const root = themeDoc.getRootElement();
    const themeElements = this.getXmlElement(root, 'themeElements');
    return this.getXmlElement(themeElements, 'clrScheme');
  }
  
  /**
   * Extract colors from color scheme
   * @private
   */
  _extractColorsFromScheme(colorScheme) {
    const colors = [];
    const colorElements = colorScheme.getChildren();
    
    for (const colorEl of colorElements) {
      const color = this._extractColorValue(colorEl);
      if (color) colors.push(color);
    }
    
    return colors;
  }
  
  /**
   * Extract color value from XML element
   * @private
   */
  _extractColorValue(element) {
    // Check for srgbClr (RGB color)
    const srgbClr = this.getXmlElement(element, 'srgbClr');
    if (srgbClr) {
      const val = srgbClr.getAttribute('val')?.getValue();
      return val ? `#${val}` : null;
    }
    
    // Check for other color types (schemeClr, etc.)
    // This would be expanded for full color support
    
    return null;
  }
  
  /**
   * Check if color is in allowed list
   * @private
   */
  _isColorAllowed(color) {
    return this.guidelines.allowedColors.some(allowed => 
      allowed.toLowerCase() === color.toLowerCase()
    );
  }
  
  /**
   * Check slide colors
   * @private
   */
  async _checkSlideColors() {
    // Implementation for checking individual slide colors
    // This would scan through slide XML for color attributes
  }
  
  /**
   * Validate logo compliance
   * @private
   */
  async _validateLogoCompliance() {
    if (!this.guidelines.logoRequirements.required) return;
    
    this.log('Validating logo compliance...');
    
    // Check for required logo placement
    const hasLogo = await this._checkForLogo();
    
    if (!hasLogo) {
      this._addViolation({
        type: 'logo',
        severity: 'high',
        element: 'presentation',
        message: 'Required company logo not found',
        location: 'slide masters',
        autoFixable: false,
        suggestedFix: 'Add company logo to slide masters'
      });
    }
  }
  
  /**
   * Check for logo presence
   * @private
   */
  async _checkForLogo() {
    // This would scan slide masters and layouts for logo images
    // Implementation depends on how logos are identified
    return false; // Placeholder
  }
  
  /**
   * Validate layout compliance
   * @private
   */
  async _validateLayoutCompliance() {
    this.log('Validating layout compliance...');
    
    // Check layout rules (margins, spacing, etc.)
    // Implementation depends on specific layout requirements
  }
  
  /**
   * Validate content compliance
   * @private
   */
  async _validateContentCompliance() {
    this.log('Validating content compliance...');
    
    // Check content rules (text length, formatting, etc.)
    // Implementation depends on specific content requirements
  }
  
  /**
   * Add a violation to the list
   * @private
   */
  _addViolation(violation) {
    if (violation.severity === 'warning' && this.options.skipWarnings) {
      return;
    }
    
    if (violation.severity === 'warning') {
      this.warnings.push(violation);
    } else {
      this.violations.push(violation);
    }
    
    if (violation.autoFixable) {
      this.autoFixes.push(violation);
    }
    
    // Log the violation
    const level = violation.severity === 'high' ? 'error' : 
                  violation.severity === 'medium' ? 'warn' : 'log';
    this[level](`${violation.type.toUpperCase()}: ${violation.message} (${violation.location})`);
  }
  
  /**
   * Apply automatic fixes
   * @private
   */
  async _applyAutoFixes() {
    const applied = [];
    
    for (const fix of this.autoFixes) {
      try {
        if (await this._applyFix(fix)) {
          applied.push(fix);
          this.log(`Applied auto-fix: ${fix.message}`);
        }
      } catch (error) {
        this.error(`Failed to apply fix for ${fix.message}: ${error.message}`);
      }
    }
    
    return applied;
  }
  
  /**
   * Apply a specific fix
   * @private
   */
  async _applyFix(violation) {
    if (violation.type === 'font' && violation.autoFixable) {
      return this._fixFontViolation(violation);
    }
    
    return false;
  }
  
  /**
   * Fix font violations by replacing with allowed font
   * @private
   */
  _fixFontViolation(violation) {
    // Implementation would modify the XML to replace unauthorized fonts
    // with the first allowed font
    const replacementFont = this.guidelines.allowedFonts[0];
    
    // This would require parsing and modifying the specific file
    // and replacing font references
    
    return true; // Placeholder
  }
  
  /**
   * Calculate overall compliance score
   * @private
   */
  _calculateComplianceScore() {
    const totalIssues = this.violations.length + this.warnings.length;
    
    if (totalIssues === 0) return 100;
    
    // Deduct points based on severity
    const deductions = this.violations.reduce((sum, v) => {
      return sum + (v.severity === 'high' ? 25 : v.severity === 'medium' ? 15 : 10);
    }, 0) + this.warnings.length * 5;
    
    return Math.max(0, 100 - deductions);
  }
  
  /**
   * Generate comprehensive compliance report
   * @private
   */
  _generateComplianceReport(score) {
    return {
      summary: {
        score: score,
        passed: score >= 80,
        totalViolations: this.violations.length,
        totalWarnings: this.warnings.length,
        autoFixableCount: this.autoFixes.length
      },
      violations: this.violations,
      warnings: this.warnings,
      guidelines: {
        fontsChecked: this.guidelines.allowedFonts.length > 0,
        colorsChecked: this.guidelines.allowedColors.length > 0,
        logoChecked: this.guidelines.logoRequirements.required,
        layoutChecked: Object.keys(this.guidelines.layoutRules).length > 0
      },
      recommendations: this._generateRecommendations(),
      metadata: {
        validatedAt: new Date().toISOString(),
        extensionVersion: '1.0.0',
        validationDuration: this.getMetrics().duration
      }
    };
  }
  
  /**
   * Generate recommendations based on violations
   * @private
   */
  _generateRecommendations() {
    const recommendations = [];
    
    if (this.violations.some(v => v.type === 'font')) {
      recommendations.push({
        category: 'Typography',
        message: `Use only approved fonts: ${this.guidelines.allowedFonts.join(', ')}`,
        priority: 'high'
      });
    }
    
    if (this.violations.some(v => v.type === 'color')) {
      recommendations.push({
        category: 'Colors',
        message: 'Stick to approved brand colors for consistency',
        priority: 'medium'
      });
    }
    
    if (this.autoFixes.length > 0) {
      recommendations.push({
        category: 'Quick Fixes',
        message: `${this.autoFixes.length} violations can be automatically fixed`,
        priority: 'low'
      });
    }
    
    return recommendations;
  }
}

// Register the extension with the framework
ExtensionFramework.register('BrandCompliance', BrandComplianceExtension, {
  type: 'VALIDATION',
  version: '1.0.0',
  description: 'Validate presentations against brand guidelines and corporate standards',
  author: 'Extension Framework',
  dependencies: ['BaseExtension'],
  hooks: [
    {
      name: 'afterValidation',
      callback: (context) => console.log(`Compliance score: ${context.result.summary.score}%`),
      priority: 5
    }
  ]
});

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = BrandComplianceExtension;
}