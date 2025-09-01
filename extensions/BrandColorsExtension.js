/**
 * BrandColorsExtension - Example theme extension for brand color management
 * 
 * PURPOSE:
 * Demonstrates how to create a brandbook-compliant extension for managing
 * corporate color palettes in PowerPoint presentations.
 * 
 * FEATURES:
 * - Apply brand color schemes
 * - Validate color compliance
 * - Support multiple brand variations
 * - Color accessibility checking
 * - Automatic color harmony generation
 * 
 * AI CONTEXT:
 * This is an example extension showing best practices for theme manipulation.
 * Users can copy and modify this pattern for their own brand requirements.
 * 
 * USAGE:
 * const brandColors = ExtensionFramework.create('BrandColors', slidesInstance, {
 *   brandColors: { primary: '#FF0000', secondary: '#0000FF' }
 * });
 * await brandColors.execute();
 */

class BrandColorsExtension extends BaseExtension {
  
  /**
   * Get default options for brand colors extension
   */
  getDefaults() {
    return {
      ...super.getDefaults(),
      autoGenerateHarmony: true,
      validateAccessibility: true,
      preserveUserColors: false,
      generatePreview: true
    };
  }
  
  /**
   * Custom initialization for brand colors
   */
  async _customInit() {
    this.log('Initializing brand colors extension...');
    
    // Validate brand colors are provided
    if (!this.options.brandColors) {
      throw new Error('Brand colors configuration is required');
    }
    
    // Parse and validate brand colors
    this.brandPalette = this._parseBrandColors(this.options.brandColors);
    
    // Generate color harmony if enabled
    if (this.options.autoGenerateHarmony) {
      this.brandPalette = this._generateColorHarmony(this.brandPalette);
    }
    
    this.log(`Loaded ${Object.keys(this.brandPalette).length} brand colors`);
  }
  
  /**
   * Custom validation for brand colors
   */
  _customValidate(input) {
    // Validate that we have minimum required colors
    const required = ['primary'];
    const missing = required.filter(color => !this.brandPalette[color]);
    
    if (missing.length > 0) {
      this.error(`Missing required brand colors: ${missing.join(', ')}`);
      return false;
    }
    
    // Validate color accessibility if enabled
    if (this.options.validateAccessibility) {
      return this._validateAccessibility();
    }
    
    return true;
  }
  
  /**
   * Main execution - apply brand colors to presentation
   */
  async _customExecute(input) {
    this.log('Applying brand colors to presentation...');
    
    // Apply colors to theme
    await this._applyColorsToTheme();
    
    // Apply colors to slide masters
    await this._applyColorsToMasters();
    
    // Apply colors to slide layouts
    await this._applyColorsToLayouts();
    
    // Generate preview if enabled
    let preview = null;
    if (this.options.generatePreview) {
      preview = this._generateColorPreview();
    }
    
    const result = {
      success: true,
      colorsApplied: Object.keys(this.brandPalette).length,
      palette: this.brandPalette,
      preview: preview,
      timestamp: new Date().toISOString()
    };
    
    this.log(`Successfully applied ${result.colorsApplied} brand colors`);
    return result;
  }
  
  /**
   * Apply theme colors - required method for THEME extensions
   */
  async applyTheme() {
    return this.execute();
  }
  
  /**
   * Validate theme - required method for THEME extensions
   */
  validateTheme() {
    return this.validate();
  }
  
  /**
   * Get theme information - optional method for THEME extensions
   */
  getThemeInfo() {
    return {
      type: 'Brand Colors',
      version: '1.0.0',
      palette: this.brandPalette,
      features: [
        'Brand color application',
        'Color harmony generation',
        'Accessibility validation',
        'Preview generation'
      ]
    };
  }
  
  /**
   * Create theme preview - optional method for THEME extensions
   */
  createPreview() {
    return this._generateColorPreview();
  }
  
  // ===== PRIVATE IMPLEMENTATION METHODS =====
  
  /**
   * Parse brand colors from configuration
   * @private
   */
  _parseBrandColors(colors) {
    const parsed = {};
    
    for (const [name, color] of Object.entries(colors)) {
      if (this.isValidColor(color)) {
        parsed[name] = color.toUpperCase();
      } else {
        this.warn(`Invalid color format for ${name}: ${color}`);
      }
    }
    
    return parsed;
  }
  
  /**
   * Generate color harmony from primary colors
   * @private
   */
  _generateColorHarmony(palette) {
    const enhanced = { ...palette };
    
    // Generate secondary colors if not provided
    if (palette.primary && !palette.secondary) {
      enhanced.secondary = this._generateComplementary(palette.primary);
      this.log(`Generated secondary color: ${enhanced.secondary}`);
    }
    
    // Generate accent colors
    if (palette.primary && !palette.accent) {
      enhanced.accent = this._generateTriadic(palette.primary);
      this.log(`Generated accent color: ${enhanced.accent}`);
    }
    
    // Generate light and dark variants
    if (palette.primary) {
      enhanced.light = this._lightenColor(palette.primary, 0.3);
      enhanced.dark = this._darkenColor(palette.primary, 0.3);
    }
    
    return enhanced;
  }
  
  /**
   * Generate complementary color
   * @private
   */
  _generateComplementary(hexColor) {
    const rgb = this.hexToRgb(hexColor);
    if (!rgb) return hexColor;
    
    // Simple complementary: invert RGB values
    const comp = {
      r: 255 - rgb.r,
      g: 255 - rgb.g,
      b: 255 - rgb.b
    };
    
    return this.rgbToHex(comp);
  }
  
  /**
   * Generate triadic color
   * @private
   */
  _generateTriadic(hexColor) {
    const rgb = this.hexToRgb(hexColor);
    if (!rgb) return hexColor;
    
    // Simple triadic: rotate hue by 120 degrees
    // This is a simplified implementation
    const triadic = {
      r: rgb.g,
      g: rgb.b,
      b: rgb.r
    };
    
    return this.rgbToHex(triadic);
  }
  
  /**
   * Lighten a color
   * @private
   */
  _lightenColor(hexColor, factor) {
    const rgb = this.hexToRgb(hexColor);
    if (!rgb) return hexColor;
    
    const lighter = {
      r: Math.min(255, Math.round(rgb.r + (255 - rgb.r) * factor)),
      g: Math.min(255, Math.round(rgb.g + (255 - rgb.g) * factor)),
      b: Math.min(255, Math.round(rgb.b + (255 - rgb.b) * factor))
    };
    
    return this.rgbToHex(lighter);
  }
  
  /**
   * Darken a color
   * @private
   */
  _darkenColor(hexColor, factor) {
    const rgb = this.hexToRgb(hexColor);
    if (!rgb) return hexColor;
    
    const darker = {
      r: Math.max(0, Math.round(rgb.r * (1 - factor))),
      g: Math.max(0, Math.round(rgb.g * (1 - factor))),
      b: Math.max(0, Math.round(rgb.b * (1 - factor)))
    };
    
    return this.rgbToHex(darker);
  }
  
  /**
   * Validate color accessibility
   * @private
   */
  _validateAccessibility() {
    this.log('Validating color accessibility...');
    
    // Simple contrast check between primary and secondary
    if (this.brandPalette.primary && this.brandPalette.secondary) {
      const contrast = this._calculateContrast(
        this.brandPalette.primary,
        this.brandPalette.secondary
      );
      
      if (contrast < 4.5) {
        this.warn(`Low contrast between primary and secondary colors: ${contrast.toFixed(2)}`);
        return false;
      }
    }
    
    return true;
  }
  
  /**
   * Calculate contrast ratio between two colors
   * @private
   */
  _calculateContrast(color1, color2) {
    const getLuminance = (hex) => {
      const rgb = this.hexToRgb(hex);
      if (!rgb) return 0;
      
      const [r, g, b] = [rgb.r, rgb.g, rgb.b].map(c => {
        c = c / 255;
        return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
      });
      
      return 0.2126 * r + 0.7152 * g + 0.0722 * b;
    };
    
    const l1 = getLuminance(color1);
    const l2 = getLuminance(color2);
    const lighter = Math.max(l1, l2);
    const darker = Math.min(l1, l2);
    
    return (lighter + 0.05) / (darker + 0.05);
  }
  
  /**
   * Apply colors to presentation theme
   * @private
   */
  async _applyColorsToTheme() {
    this.log('Applying colors to theme...');
    
    // Get theme file
    const themeXml = this.getFile('ppt/theme/theme1.xml');
    if (!themeXml) {
      this.warn('Theme file not found');
      return;
    }
    
    // Parse and modify theme XML
    try {
      const doc = XmlService.parse(themeXml);
      const root = doc.getRootElement();
      
      // Find color scheme element
      const themeElements = this.getXmlElement(root, 'themeElements');
      const colorScheme = this.getXmlElement(themeElements, 'clrScheme');
      
      if (colorScheme) {
        // Apply brand colors to theme
        this._updateThemeColors(colorScheme);
        
        // Save modified theme
        const modifiedXml = XmlService.getPrettyFormat().format(doc);
        this.setFile('ppt/theme/theme1.xml', modifiedXml);
        
        this.log('Theme colors updated successfully');
      }
      
    } catch (error) {
      this.error(`Failed to update theme colors: ${error.message}`);
    }
  }
  
  /**
   * Update theme color elements
   * @private
   */
  _updateThemeColors(colorScheme) {
    const colorMappings = {
      'accent1': this.brandPalette.primary,
      'accent2': this.brandPalette.secondary,
      'accent3': this.brandPalette.accent
    };
    
    for (const [themeName, brandColor] of Object.entries(colorMappings)) {
      if (brandColor) {
        const colorElement = this.getXmlElement(colorScheme, themeName);
        if (colorElement) {
          this._setColorValue(colorElement, brandColor);
        }
      }
    }
  }
  
  /**
   * Set color value in XML element
   * @private
   */
  _setColorValue(element, hexColor) {
    // Remove existing color elements
    element.removeContent();
    
    // Add new srgbClr element
    const srgbClr = this.createXmlElement('srgbClr');
    srgbClr.setAttribute('val', hexColor.replace('#', ''));
    element.addContent(srgbClr);
  }
  
  /**
   * Apply colors to slide masters
   * @private
   */
  async _applyColorsToMasters() {
    this.log('Applying colors to slide masters...');
    
    // This would contain logic to update slide masters
    // Implementation depends on specific requirements
  }
  
  /**
   * Apply colors to slide layouts
   * @private
   */
  async _applyColorsToLayouts() {
    this.log('Applying colors to slide layouts...');
    
    // This would contain logic to update slide layouts
    // Implementation depends on specific requirements
  }
  
  /**
   * Generate color preview
   * @private
   */
  _generateColorPreview() {
    const preview = {
      palette: this.brandPalette,
      swatches: [],
      metadata: {
        totalColors: Object.keys(this.brandPalette).length,
        hasAccessibilityValidation: this.options.validateAccessibility,
        generatedAt: new Date().toISOString()
      }
    };
    
    // Generate color swatches
    for (const [name, color] of Object.entries(this.brandPalette)) {
      const rgb = this.hexToRgb(color);
      preview.swatches.push({
        name,
        hex: color,
        rgb: rgb,
        isLight: this._isLightColor(color)
      });
    }
    
    return preview;
  }
  
  /**
   * Check if color is light or dark
   * @private
   */
  _isLightColor(hexColor) {
    const rgb = this.hexToRgb(hexColor);
    if (!rgb) return false;
    
    // Calculate perceived lightness
    const brightness = (rgb.r * 299 + rgb.g * 587 + rgb.b * 114) / 1000;
    return brightness > 128;
  }
}

// Register the extension with the framework
ExtensionFramework.register('BrandColors', BrandColorsExtension, {
  type: 'THEME',
  version: '1.0.0',
  description: 'Apply and manage brand color palettes in PowerPoint presentations',
  author: 'Extension Framework',
  dependencies: ['BaseExtension'],
  hooks: [
    {
      name: 'beforeApplyTheme',
      callback: (context) => console.log('Applying brand colors...'),
      priority: 10
    }
  ]
});

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = BrandColorsExtension;
}