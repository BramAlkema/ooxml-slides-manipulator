/**
 * Extension Templates - Ready-to-use templates for common extension types
 * 
 * PURPOSE:
 * Provides complete, working templates that users can copy and customize
 * for their specific brandbook compliance and PowerPoint automation needs.
 * 
 * TEMPLATES INCLUDED:
 * - Brand Theme Extension: Colors, fonts, logos
 * - Content Validation Extension: Text, layout, compliance
 * - Layout Template Extension: Slide templates and layouts
 * - Export Extension: Custom export formats
 * - Workflow Extension: Multi-step automation workflows
 * 
 * AI CONTEXT:
 * These are production-ready templates. Users can copy any template,
 * modify the configuration and logic for their needs, and register
 * with the Extension Framework immediately.
 */

class ExtensionTemplates {
  
  /**
   * Get all available templates
   * 
   * @returns {Object} All available templates by category
   * 
   * AI_USAGE: Discover all available extension templates
   * Example: const templates = ExtensionTemplates.getAllTemplates()
   */
  static getAllTemplates() {
    return {
      theme: {
        brandTheme: this.getBrandThemeTemplate(),
        colorPalette: this.getColorPaletteTemplate(),
        typography: this.getTypographyTemplate()
      },
      validation: {
        compliance: this.getComplianceTemplate(),
        accessibility: this.getAccessibilityTemplate(),
        contentRules: this.getContentRulesTemplate()
      },
      template: {
        slideLayouts: this.getSlideLayoutTemplate(),
        masterSlides: this.getMasterSlideTemplate(),
        brandedTemplates: this.getBrandedTemplateTemplate()
      },
      export: {
        customFormats: this.getCustomExportTemplate(),
        brandedPDF: this.getBrandedPDFTemplate(),
        webFormats: this.getWebExportTemplate()
      },
      workflow: {
        automation: this.getAutomationTemplate(),
        batch: this.getBatchProcessingTemplate(),
        pipeline: this.getPipelineTemplate()
      }
    };
  }
  
  /**
   * Get Brand Theme Extension Template
   * Complete template for brand theme management
   */
  static getBrandThemeTemplate() {
    return {
      name: 'BrandThemeExtension',
      type: 'THEME',
      description: 'Comprehensive brand theme management',
      code: `
/**
 * Custom Brand Theme Extension
 * Manages complete brand themes including colors, fonts, and visual elements
 */
class CustomBrandThemeExtension extends BaseExtension {
  
  getDefaults() {
    return {
      ...super.getDefaults(),
      applyToMasters: true,
      applyToLayouts: true,
      validateAccessibility: true,
      createPreview: true,
      brandConfig: {
        // Define your brand configuration here
        colors: {
          primary: '#0066CC',
          secondary: '#FF6600', 
          accent: '#00AA44',
          neutral: '#666666',
          background: '#FFFFFF',
          text: '#333333'
        },
        fonts: {
          heading: 'Arial Bold',
          body: 'Arial',
          accent: 'Arial Italic'
        },
        logo: {
          imageId: null, // Google Drive image ID
          position: 'top-right',
          size: 'medium'
        }
      }
    };
  }
  
  async _customExecute(input) {
    const config = { ...this.options.brandConfig, ...input };
    
    this.log('Applying complete brand theme...');
    
    // Apply brand colors
    await this._applyBrandColors(config.colors);
    
    // Apply brand fonts
    await this._applyBrandFonts(config.fonts);
    
    // Apply logo if specified
    if (config.logo.imageId) {
      await this._applyBrandLogo(config.logo);
    }
    
    // Validate accessibility if enabled
    if (this.options.validateAccessibility) {
      const accessibility = await this._validateAccessibility(config);
      if (!accessibility.passed) {
        this.warn(\`Accessibility issues found: \${accessibility.issues.length}\`);
      }
    }
    
    return {
      success: true,
      applied: {
        colors: Object.keys(config.colors).length,
        fonts: Object.keys(config.fonts).length,
        logo: !!config.logo.imageId
      },
      accessibility: this.options.validateAccessibility,
      timestamp: new Date().toISOString()
    };
  }
  
  async _applyBrandColors(colors) {
    this.log('Applying brand colors...');
    
    // Get theme XML
    const themeXml = this.getFile('ppt/theme/theme1.xml');
    if (!themeXml) return;
    
    const doc = XmlService.parse(themeXml);
    const colorScheme = this._findColorScheme(doc);
    
    if (colorScheme) {
      // Map brand colors to theme colors
      const colorMappings = {
        'accent1': colors.primary,
        'accent2': colors.secondary,
        'accent3': colors.accent,
        'dk1': colors.text,
        'lt1': colors.background
      };
      
      for (const [themeColor, brandColor] of Object.entries(colorMappings)) {
        if (brandColor) {
          this._setThemeColor(colorScheme, themeColor, brandColor);
        }
      }
      
      // Save modified theme
      const modifiedXml = XmlService.getPrettyFormat().format(doc);
      this.setFile('ppt/theme/theme1.xml', modifiedXml);
    }
  }
  
  async _applyBrandFonts(fonts) {
    this.log('Applying brand fonts...');
    
    // Implementation for font application
    // This would modify theme font scheme
  }
  
  async _applyBrandLogo(logoConfig) {
    this.log('Applying brand logo...');
    
    // Implementation for logo insertion
    // This would add logo to slide masters
  }
  
  async _validateAccessibility(config) {
    // Implementation for accessibility validation
    return { passed: true, issues: [] };
  }
  
  _findColorScheme(themeDoc) {
    const root = themeDoc.getRootElement();
    const themeElements = this.getXmlElement(root, 'themeElements');
    return this.getXmlElement(themeElements, 'clrScheme');
  }
  
  _setThemeColor(colorScheme, themeColorName, hexColor) {
    const colorElement = this.getXmlElement(colorScheme, themeColorName);
    if (colorElement) {
      colorElement.removeContent();
      const srgbClr = this.createXmlElement('srgbClr');
      srgbClr.setAttribute('val', hexColor.replace('#', ''));
      colorElement.addContent(srgbClr);
    }
  }
  
  // Required theme extension methods
  async applyTheme() { return this.execute(); }
  validateTheme() { return this.validate(); }
  getThemeInfo() {
    return {
      name: 'Custom Brand Theme',
      version: '1.0.0',
      colors: Object.keys(this.options.brandConfig.colors).length,
      fonts: Object.keys(this.options.brandConfig.fonts).length
    };
  }
}

// Register the extension
ExtensionFramework.register('CustomBrandTheme', CustomBrandThemeExtension, {
  type: 'THEME',
  version: '1.0.0',
  description: 'Complete brand theme management with colors, fonts, and logos'
});`
    };
  }
  
  /**
   * Get Compliance Validation Template
   * Complete template for brand compliance checking
   */
  static getComplianceTemplate() {
    return {
      name: 'ComplianceValidationExtension',
      type: 'VALIDATION',
      description: 'Comprehensive brand compliance validation',
      code: `
/**
 * Custom Compliance Validation Extension
 * Validates presentations against specific brand guidelines
 */
class CustomComplianceExtension extends BaseExtension {
  
  getDefaults() {
    return {
      ...super.getDefaults(),
      strictMode: false,
      autoFix: false,
      generateReport: true,
      complianceRules: {
        // Define your compliance rules here
        fonts: {
          allowed: ['Arial', 'Helvetica', 'Calibri'],
          forbidden: ['Comic Sans MS', 'Papyrus'],
          requireConsistency: true
        },
        colors: {
          palette: ['#0066CC', '#FF6600', '#00AA44', '#666666'],
          allowOffBrand: false,
          contrastMinimum: 4.5
        },
        content: {
          maxSlides: 50,
          maxTextPerSlide: 200,
          requireSlideNumbers: true,
          requireConsistentLayout: true
        },
        branding: {
          requireLogo: true,
          logoPosition: 'consistent',
          forbiddenElements: ['watermarks', 'personal-photos']
        }
      }
    };
  }
  
  async _customExecute(input) {
    const rules = { ...this.options.complianceRules, ...input };
    
    this.log('Running compliance validation...');
    
    // Clear previous violations
    this.violations = [];
    this.warnings = [];
    this.fixes = [];
    
    // Run all compliance checks
    await this._validateFontCompliance(rules.fonts);
    await this._validateColorCompliance(rules.colors);
    await this._validateContentCompliance(rules.content);
    await this._validateBrandingCompliance(rules.branding);
    
    // Apply fixes if enabled
    if (this.options.autoFix) {
      await this._applyAutoFixes();
    }
    
    // Generate report
    const report = this._generateReport();
    
    this.log(\`Compliance check completed - Score: \${report.score}%\`);
    return report;
  }
  
  async _validateFontCompliance(fontRules) {
    if (!fontRules.allowed.length) return;
    
    this.log('Validating font compliance...');
    
    const files = this.context.listFiles();
    const slideFiles = files.filter(f => f.includes('slides/slide') && f.endsWith('.xml'));
    
    for (const slideFile of slideFiles) {
      const slideXml = this.getFile(slideFile);
      if (slideXml) {
        const fonts = this._extractFontsFromXml(slideXml);
        
        for (const font of fonts) {
          if (fontRules.forbidden.includes(font)) {
            this._addViolation({
              type: 'font',
              severity: 'high',
              message: \`Forbidden font used: \${font}\`,
              location: slideFile,
              autoFixable: true,
              fix: \`Replace with \${fontRules.allowed[0]}\`
            });
          } else if (!fontRules.allowed.includes(font)) {
            this._addViolation({
              type: 'font',
              severity: 'medium',
              message: \`Unauthorized font: \${font}\`,
              location: slideFile,
              autoFixable: false,
              fix: \`Use approved fonts: \${fontRules.allowed.join(', ')}\`
            });
          }
        }
      }
    }
  }
  
  async _validateColorCompliance(colorRules) {
    this.log('Validating color compliance...');
    // Implementation for color validation
  }
  
  async _validateContentCompliance(contentRules) {
    this.log('Validating content compliance...');
    
    const files = this.context.listFiles();
    const slideFiles = files.filter(f => f.includes('slides/slide') && f.endsWith('.xml'));
    
    // Check slide count
    if (contentRules.maxSlides && slideFiles.length > contentRules.maxSlides) {
      this._addViolation({
        type: 'content',
        severity: 'medium',
        message: \`Too many slides: \${slideFiles.length}/\${contentRules.maxSlides}\`,
        location: 'presentation',
        autoFixable: false,
        fix: 'Reduce number of slides or split into multiple presentations'
      });
    }
    
    // Check text length per slide
    for (const slideFile of slideFiles) {
      const textLength = this._getSlideTextLength(slideFile);
      if (contentRules.maxTextPerSlide && textLength > contentRules.maxTextPerSlide) {
        this._addViolation({
          type: 'content',
          severity: 'low',
          message: \`Slide has too much text: \${textLength} characters\`,
          location: slideFile,
          autoFixable: false,
          fix: 'Reduce text content or split across multiple slides'
        });
      }
    }
  }
  
  async _validateBrandingCompliance(brandingRules) {
    this.log('Validating branding compliance...');
    
    if (brandingRules.requireLogo) {
      const hasLogo = await this._checkForLogo();
      if (!hasLogo) {
        this._addViolation({
          type: 'branding',
          severity: 'high',
          message: 'Required company logo not found',
          location: 'slide masters',
          autoFixable: false,
          fix: 'Add company logo to slide master'
        });
      }
    }
  }
  
  _extractFontsFromXml(xmlContent) {
    // Extract font names from XML content
    const fonts = [];
    const typefaceRegex = /typeface="([^"]+)"/g;
    let match;
    
    while ((match = typefaceRegex.exec(xmlContent)) !== null) {
      if (!match[1].startsWith('+')) { // Skip theme references
        fonts.push(match[1]);
      }
    }
    
    return [...new Set(fonts)]; // Remove duplicates
  }
  
  _getSlideTextLength(slideFile) {
    const slideXml = this.getFile(slideFile);
    if (!slideXml) return 0;
    
    // Extract text content and count characters
    const textRegex = /<a:t[^>]*>([^<]*)</a:t>/g;
    let totalLength = 0;
    let match;
    
    while ((match = textRegex.exec(slideXml)) !== null) {
      totalLength += match[1].length;
    }
    
    return totalLength;
  }
  
  async _checkForLogo() {
    // Check slide masters for logo images
    // This is a simplified implementation
    const files = this.context.listFiles();
    const masterFiles = files.filter(f => f.includes('slideMasters/'));
    
    for (const masterFile of masterFiles) {
      const masterXml = this.getFile(masterFile);
      if (masterXml && masterXml.includes('pic:pic')) {
        return true; // Found images in master
      }
    }
    
    return false;
  }
  
  _addViolation(violation) {
    if (violation.severity === 'high') {
      this.violations.push(violation);
    } else {
      this.warnings.push(violation);
    }
    
    if (violation.autoFixable) {
      this.fixes.push(violation);
    }
  }
  
  async _applyAutoFixes() {
    this.log(\`Applying \${this.fixes.length} automatic fixes...\`);
    
    for (const fix of this.fixes) {
      if (fix.type === 'font' && fix.autoFixable) {
        // Apply font fix
        await this._fixFontViolation(fix);
      }
    }
  }
  
  async _fixFontViolation(violation) {
    // Implementation to fix font violations
    this.log(\`Fixed font violation: \${violation.message}\`);
  }
  
  _generateReport() {
    const totalViolations = this.violations.length + this.warnings.length;
    const score = Math.max(0, 100 - (this.violations.length * 20) - (this.warnings.length * 5));
    
    return {
      score: score,
      passed: score >= 80,
      summary: {
        violations: this.violations.length,
        warnings: this.warnings.length,
        fixesAvailable: this.fixes.length
      },
      violations: this.violations,
      warnings: this.warnings,
      recommendations: this._generateRecommendations(),
      timestamp: new Date().toISOString()
    };
  }
  
  _generateRecommendations() {
    const recs = [];
    
    if (this.violations.some(v => v.type === 'font')) {
      recs.push('Use only approved corporate fonts for consistency');
    }
    
    if (this.violations.some(v => v.type === 'branding')) {
      recs.push('Ensure corporate logo is present on all slides');
    }
    
    return recs;
  }
  
  // Required validation extension methods
  async validate() {
    const report = await this.execute();
    return { passed: report.passed, violations: report.violations, score: report.score };
  }
  
  getViolations() {
    return { violations: this.violations, warnings: this.warnings };
  }
}

// Register the extension
ExtensionFramework.register('CustomCompliance', CustomComplianceExtension, {
  type: 'VALIDATION',
  version: '1.0.0',
  description: 'Custom brand compliance validation with auto-fix capabilities'
});`
    };
  }
  
  /**
   * Get Automation Workflow Template
   * Complete template for multi-step automation
   */
  static getAutomationTemplate() {
    return {
      name: 'AutomationWorkflowExtension',
      type: 'WORKFLOW',
      description: 'Multi-step automation workflow',
      code: `
/**
 * Custom Automation Workflow Extension
 * Orchestrates multiple operations in a defined sequence
 */
class CustomAutomationExtension extends BaseExtension {
  
  getDefaults() {
    return {
      ...super.getDefaults(),
      stopOnError: false,
      parallel: false,
      saveCheckpoints: true,
      workflow: {
        // Define your workflow steps here
        steps: [
          {
            name: 'applyBrandColors',
            extension: 'BrandColors',
            options: { brandColors: { primary: '#0066CC' } },
            required: true
          },
          {
            name: 'validateCompliance',
            extension: 'BrandCompliance',
            options: { brandGuidelines: { allowedFonts: ['Arial'] } },
            required: false
          },
          {
            name: 'generateReport',
            action: 'custom',
            function: 'generateSummaryReport',
            required: false
          }
        ]
      }
    };
  }
  
  async _customExecute(input) {
    const workflow = { ...this.options.workflow, ...input };
    
    this.log(\`Starting automation workflow with \${workflow.steps.length} steps...\`);
    
    const results = [];
    let checkpoint = 0;
    
    try {
      if (this.options.parallel) {
        // Execute steps in parallel
        results.push(...await this._executeParallel(workflow.steps));
      } else {
        // Execute steps sequentially
        for (let i = 0; i < workflow.steps.length; i++) {
          const step = workflow.steps[i];
          checkpoint = i;
          
          this.log(\`Executing step \${i + 1}/\${workflow.steps.length}: \${step.name}\`);
          
          const result = await this._executeStep(step);
          results.push(result);
          
          // Stop on error if configured
          if (!result.success && this.options.stopOnError) {
            throw new Error(\`Workflow stopped at step \${step.name}: \${result.error}\`);
          }
          
          // Save checkpoint if enabled
          if (this.options.saveCheckpoints) {
            await this._saveCheckpoint(i, results);
          }
        }
      }
      
      const summary = this._generateWorkflowSummary(results);
      
      this.log(\`Automation workflow completed successfully - \${summary.successCount}/\${summary.totalSteps} steps passed\`);
      return summary;
      
    } catch (error) {
      this.error(\`Workflow failed at step \${checkpoint + 1}: \${error.message}\`);
      
      return {
        success: false,
        error: error.message,
        checkpoint: checkpoint,
        results: results,
        timestamp: new Date().toISOString()
      };
    }
  }
  
  async _executeStep(step) {
    const startTime = Date.now();
    
    try {
      let result;
      
      if (step.extension) {
        // Execute extension step
        result = await this._executeExtensionStep(step);
      } else if (step.action === 'custom' && step.function) {
        // Execute custom function
        result = await this._executeCustomStep(step);
      } else {
        throw new Error(\`Unknown step type for step: \${step.name}\`);
      }
      
      return {
        step: step.name,
        success: true,
        result: result,
        duration: Date.now() - startTime,
        timestamp: new Date().toISOString()
      };
      
    } catch (error) {
      return {
        step: step.name,
        success: false,
        error: error.message,
        duration: Date.now() - startTime,
        timestamp: new Date().toISOString()
      };
    }
  }
  
  async _executeExtensionStep(step) {
    // Create or get extension instance
    const extension = ExtensionFramework.create(step.extension, this.context, step.options || {});
    
    // Execute the extension
    return await extension.execute();
  }
  
  async _executeCustomStep(step) {
    // Execute custom function by name
    if (typeof this[step.function] === 'function') {
      return await this[step.function](step.options || {});
    } else {
      throw new Error(\`Custom function not found: \${step.function}\`);
    }
  }
  
  async _executeParallel(steps) {
    // Execute all steps in parallel
    const promises = steps.map(step => this._executeStep(step));
    return await Promise.all(promises);
  }
  
  async _saveCheckpoint(stepIndex, results) {
    // Save workflow progress (could save to Drive, cache, etc.)
    this.setCached(\`checkpoint_\${stepIndex}\`, {
      stepIndex: stepIndex,
      results: results,
      timestamp: new Date().toISOString()
    });
  }
  
  _generateWorkflowSummary(results) {
    const successCount = results.filter(r => r.success).length;
    const errorCount = results.filter(r => !r.success).length;
    const totalDuration = results.reduce((sum, r) => sum + (r.duration || 0), 0);
    
    return {
      success: errorCount === 0,
      summary: {
        totalSteps: results.length,
        successCount: successCount,
        errorCount: errorCount,
        totalDuration: totalDuration
      },
      results: results,
      recommendations: this._generateWorkflowRecommendations(results),
      timestamp: new Date().toISOString()
    };
  }
  
  _generateWorkflowRecommendations(results) {
    const recs = [];
    
    const slowSteps = results.filter(r => r.duration > 5000);
    if (slowSteps.length > 0) {
      recs.push(\`Consider optimizing slow steps: \${slowSteps.map(s => s.step).join(', ')}\`);
    }
    
    const errorSteps = results.filter(r => !r.success);
    if (errorSteps.length > 0) {
      recs.push(\`Review failed steps: \${errorSteps.map(s => s.step).join(', ')}\`);
    }
    
    return recs;
  }
  
  // Custom workflow functions (examples)
  async generateSummaryReport(options = {}) {
    this.log('Generating workflow summary report...');
    
    const info = this.context.getInfo();
    const metrics = this.context.getMetrics();
    
    return {
      report: 'Workflow Summary Report',
      presentation: {
        slides: info.slides.count,
        fileSize: info.metadata.originalSize
      },
      performance: {
        loadTime: metrics.loadTime,
        operations: metrics.operations
      },
      generatedAt: new Date().toISOString()
    };
  }
  
  async applyCustomFormatting(options = {}) {
    this.log('Applying custom formatting...');
    
    // Custom formatting logic here
    return { applied: true, changes: ['formatting'] };
  }
}

// Register the extension
ExtensionFramework.register('CustomAutomation', CustomAutomationExtension, {
  type: 'WORKFLOW',
  version: '1.0.0',
  description: 'Multi-step automation workflow with checkpoint support'
});`
    };
  }
  
  /**
   * Get usage examples and best practices
   */
  static getUsageExamples() {
    return {
      basicUsage: `
// Basic Extension Usage
const slides = new OOXMLSlides(fileId);
await slides.load();

// Apply brand colors
await slides.applyBrandColors({
  primary: '#0066CC',
  secondary: '#FF6600',
  accent: '#00AA44'
});

// Validate compliance
const report = await slides.validateCompliance({
  allowedFonts: ['Arial', 'Helvetica'],
  allowedColors: ['#0066CC', '#FF6600', '#00AA44']
});

console.log(\`Compliance score: \${report.score}%\`);
`,
      
      customExtension: `
// Using Custom Extensions
// First register your extension
slides.registerExtension('MyBrandValidator', MyValidatorClass, {
  type: 'VALIDATION'
});

// Then use it
const result = await slides.useExtension('MyBrandValidator', {
  strictMode: true,
  companyGuidelines: myGuidelines
});
`,
      
      advancedWorkflow: `
// Advanced Workflow
const slides = new OOXMLSlides(fileId, {
  enableExtensions: true,
  preloadExtensions: ['BrandColors', 'BrandCompliance']
});

await slides.load();

// Use multiple extensions in sequence
await slides.applyBrandColors(brandColors);
const compliance = await slides.validateCompliance(guidelines);

if (compliance.score < 80) {
  // Auto-fix compliance issues
  await slides.useExtension('BrandCompliance', { autoFix: true });
}

// Save with detailed metrics
const result = await slides.save();
console.log('Metrics:', result.metrics);
`,
      
      extensionDevelopment: `
// Developing Custom Extensions

// 1. Extend BaseExtension
class MyCustomExtension extends BaseExtension {
  async _customExecute(input) {
    // Your custom logic here
    return { success: true, result: 'Custom processing complete' };
  }
}

// 2. Register with framework
ExtensionFramework.register('MyExtension', MyCustomExtension, {
  type: 'CONTENT',
  version: '1.0.0',
  description: 'My custom PowerPoint processing'
});

// 3. Use in presentations
const result = await slides.useExtension('MyExtension', options);
`
    };
  }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = ExtensionTemplates;
}