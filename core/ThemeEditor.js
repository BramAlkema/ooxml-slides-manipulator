/**
 * ThemeEditor - Advanced theme manipulation for PPTX files
 * Handles font pairs, color palettes, and theme properties
 */
class ThemeEditor {
  constructor(ooxmlParser) {
    this.parser = ooxmlParser;
    this.namespaces = {
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
    };
  }

  /**
   * Get current color scheme from theme
   * @returns {Object} Color scheme object
   */
  getColorScheme() {
    const themeDoc = this.parser.getThemeXML();
    const root = themeDoc.getRootElement();
    const themeElements = root.getChild('themeElements', XmlService.getNamespace('a', this.namespaces.a));
    const clrScheme = themeElements.getChild('clrScheme', XmlService.getNamespace('a', this.namespaces.a));
    
    const colors = {};
    const colorElements = clrScheme.getChildren();
    
    colorElements.forEach(colorEl => {
      const name = colorEl.getName();
      const colorValue = this._extractColorValue(colorEl);
      colors[name] = colorValue;
    });
    
    return colors;
  }

  /**
   * Set color palette for the theme
   * @param {Object} colorPalette - Object with color definitions
   * @example
   * setColorPalette({
   *   accent1: '#FF0000',
   *   accent2: '#00FF00', 
   *   accent3: '#0000FF',
   *   dk1: '#000000',
   *   lt1: '#FFFFFF'
   * })
   */
  setColorPalette(colorPalette) {
    const themeDoc = this.parser.getThemeXML();
    const root = themeDoc.getRootElement();
    const themeElements = root.getChild('themeElements', XmlService.getNamespace('a', this.namespaces.a));
    const clrScheme = themeElements.getChild('clrScheme', XmlService.getNamespace('a', this.namespaces.a));
    
    // Update each color in the palette
    Object.keys(colorPalette).forEach(colorName => {
      const colorEl = clrScheme.getChild(colorName, XmlService.getNamespace('a', this.namespaces.a));
      if (colorEl) {
        this._setColorValue(colorEl, colorPalette[colorName]);
      }
    });
    
    this.parser.setThemeXML(themeDoc);
    return this;
  }

  /**
   * Get current font scheme
   * @returns {Object} Font scheme with major and minor fonts
   */
  getFontScheme() {
    const themeDoc = this.parser.getThemeXML();
    const root = themeDoc.getRootElement();
    const themeElements = root.getChild('themeElements', XmlService.getNamespace('a', this.namespaces.a));
    const fontScheme = themeElements.getChild('fontScheme', XmlService.getNamespace('a', this.namespaces.a));
    
    const majorFont = this._extractFontInfo(fontScheme.getChild('majorFont', XmlService.getNamespace('a', this.namespaces.a)));
    const minorFont = this._extractFontInfo(fontScheme.getChild('minorFont', XmlService.getNamespace('a', this.namespaces.a)));
    
    return {
      major: majorFont,
      minor: minorFont
    };
  }

  /**
   * Set font pair for the theme
   * @param {string} majorFont - Font for headings (e.g., 'Calibri')
   * @param {string} minorFont - Font for body text (e.g., 'Arial')
   */
  setFontPair(majorFont, minorFont) {
    const themeDoc = this.parser.getThemeXML();
    const root = themeDoc.getRootElement();
    const themeElements = root.getChild('themeElements', XmlService.getNamespace('a', this.namespaces.a));
    const fontScheme = themeElements.getChild('fontScheme', XmlService.getNamespace('a', this.namespaces.a));
    
    // Update major font (headings)
    if (majorFont) {
      const majorFontEl = fontScheme.getChild('majorFont', XmlService.getNamespace('a', this.namespaces.a));
      this._setFontFamily(majorFontEl, majorFont);
    }
    
    // Update minor font (body)
    if (minorFont) {
      const minorFontEl = fontScheme.getChild('minorFont', XmlService.getNamespace('a', this.namespaces.a));
      this._setFontFamily(minorFontEl, minorFont);
    }
    
    this.parser.setThemeXML(themeDoc);
    return this;
  }

  /**
   * Set complete color scheme with predefined themes
   * @param {string} themeName - 'office', 'colorful', 'dark', 'minimal', etc.
   */
  setColorTheme(themeName) {
    const themes = this._getPredefinedThemes();
    const theme = themes[themeName];
    
    if (!theme) {
      throw new Error(`Unknown theme: ${themeName}. Available: ${Object.keys(themes).join(', ')}`);
    }
    
    return this.setColorPalette(theme);
  }

  /**
   * Create custom color scheme from array of colors
   * @param {Array<string>} colors - Array of hex colors
   */
  setColorPaletteFromArray(colors) {
    if (colors.length < 6) {
      throw new Error('Need at least 6 colors for a complete palette');
    }
    
    const colorPalette = {
      dk1: colors[0] || '#000000',  // Dark 1
      lt1: colors[1] || '#FFFFFF',  // Light 1
      dk2: colors[2] || '#44546A',  // Dark 2
      lt2: colors[3] || '#E7E6E6',  // Light 2
      accent1: colors[4] || '#5B9BD5', // Accent 1
      accent2: colors[5] || '#70AD47', // Accent 2
      accent3: colors[6] || '#A5A5A5', // Accent 3
      accent4: colors[7] || '#FFC000', // Accent 4
      accent5: colors[8] || '#4472C4', // Accent 5
      accent6: colors[9] || '#70AD47'  // Accent 6
    };
    
    return this.setColorPalette(colorPalette);
  }

  // Private helper methods
  _extractColorValue(colorElement) {
    const children = colorElement.getChildren();
    for (let child of children) {
      if (child.getName() === 'srgbClr') {
        return '#' + child.getAttribute('val').getValue();
      } else if (child.getName() === 'sysClr') {
        return child.getAttribute('val').getValue();
      }
    }
    return null;
  }

  _setColorValue(colorElement, hexColor) {
    try {
      // Remove existing color children
      const children = colorElement.getChildren();
      // Use removeContent instead of removeChild for XmlService
      children.forEach(child => {
        colorElement.removeContent(child);
      });
      
      // Add new srgbClr element
      const srgbClr = XmlService.createElement('srgbClr', XmlService.getNamespace('a', this.namespaces.a));
      srgbClr.setAttribute('val', hexColor.replace('#', ''));
      colorElement.addContent(srgbClr);
    } catch (error) {
      // Fallback: create new element structure
      console.log('XML manipulation fallback for color:', hexColor);
      const srgbClr = XmlService.createElement('srgbClr', XmlService.getNamespace('a', this.namespaces.a));
      srgbClr.setAttribute('val', hexColor.replace('#', ''));
      colorElement.addContent(srgbClr);
    }
  }

  _extractFontInfo(fontElement) {
    const latin = fontElement.getChild('latin', XmlService.getNamespace('a', this.namespaces.a));
    return {
      typeface: latin ? latin.getAttribute('typeface').getValue() : null
    };
  }

  _setFontFamily(fontElement, fontName) {
    const latin = fontElement.getChild('latin', XmlService.getNamespace('a', this.namespaces.a));
    if (latin) {
      latin.getAttribute('typeface').setValue(fontName);
    }
  }

  _getPredefinedThemes() {
    return {
      office: {
        dk1: '#000000',
        lt1: '#FFFFFF', 
        dk2: '#44546A',
        lt2: '#E7E6E6',
        accent1: '#5B9BD5',
        accent2: '#70AD47',
        accent3: '#A5A5A5',
        accent4: '#FFC000',
        accent5: '#4472C4',
        accent6: '#70AD47'
      },
      colorful: {
        dk1: '#000000',
        lt1: '#FFFFFF',
        dk2: '#44546A', 
        lt2: '#E7E6E6',
        accent1: '#E74C3C',
        accent2: '#9B59B6',
        accent3: '#3498DB',
        accent4: '#1ABC9C',
        accent5: '#F39C12',
        accent6: '#2ECC71'
      },
      dark: {
        dk1: '#FFFFFF',
        lt1: '#000000',
        dk2: '#EEEEEE',
        lt2: '#1C1C1C', 
        accent1: '#FF6B6B',
        accent2: '#4ECDC4',
        accent3: '#45B7D1',
        accent4: '#96CEB4',
        accent5: '#FECA57',
        accent6: '#FF9FF3'
      }
    };
  }
}