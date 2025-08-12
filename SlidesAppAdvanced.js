/**
 * SlidesAppAdvanced - Extended Google Slides API (Tanaikech Style)
 * 
 * Provides access to "undocumented things under the hood" including:
 * - Font pair manipulation
 * - Color palette control 
 * - Theme properties beyond standard API
 * - Advanced slide master manipulation
 * - Custom layout properties
 * 
 * Inspired by tanaikech's approach to accessing Google Apps features
 * not exposed through official APIs.
 */

class SlidesAppAdvanced {
  
  /**
   * Enhanced presentation wrapper with advanced features
   */
  static openById(presentationId) {
    const presentation = SlidesApp.openById(presentationId);
    return new AdvancedPresentation(presentation, presentationId);
  }
  
  /**
   * Create presentation with advanced theme options
   */
  static create(title, options = {}) {
    const presentation = SlidesApp.create(title);
    const advanced = new AdvancedPresentation(presentation, presentation.getId());
    
    if (options.theme) {
      advanced.setAdvancedTheme(options.theme);
    }
    
    if (options.fontPairs) {
      advanced.setFontPairs(options.fontPairs);
    }
    
    return advanced;
  }
  
  /**
   * Access presentation via Drive API for additional properties
   */
  static getAdvancedProperties(presentationId) {
    try {
      // Use Drive API v3 to get additional metadata
      const response = UrlFetchApp.fetch(
        `https://www.googleapis.com/drive/v3/files/${presentationId}?fields=*`,
        {
          headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
        }
      );
      
      const metadata = JSON.parse(response.getContentText());
      
      // Use Slides API v1 for presentation-specific data
      const slidesResponse = UrlFetchApp.fetch(
        `https://slides.googleapis.com/v1/presentations/${presentationId}`,
        {
          headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
        }
      );
      
      const slidesData = JSON.parse(slidesResponse.getContentText());
      
      return {
        drive: metadata,
        slides: slidesData,
        // Extract undocumented properties
        advancedProperties: SlidesAppAdvanced._extractAdvancedProperties(slidesData)
      };
      
    } catch (error) {
      console.error('Failed to get advanced properties:', error);
      return null;
    }
  }
  
  /**
   * Extract undocumented properties from Slides API response
   */
  static _extractAdvancedProperties(slidesData) {
    const advanced = {};
    
    // Extract master theme information
    if (slidesData.masters && slidesData.masters.length > 0) {
      advanced.masterThemes = slidesData.masters.map(master => ({
        objectId: master.objectId,
        colorScheme: master.masterProperties?.colorScheme,
        fontPairs: SlidesAppAdvanced._extractFontPairs(master)
      }));
    }
    
    // Extract layout-specific properties  
    if (slidesData.layouts) {
      advanced.layoutProperties = slidesData.layouts.map(layout => ({
        objectId: layout.objectId,
        layoutProperties: layout.layoutProperties,
        // Look for undocumented properties
        customProperties: SlidesAppAdvanced._extractCustomProperties(layout)
      }));
    }
    
    // Extract presentation-level theme data
    if (slidesData.presentationStyle) {
      advanced.presentationStyle = slidesData.presentationStyle;
    }
    
    return advanced;
  }
  
  /**
   * Extract font pair information (tanaikech style)
   */
  static _extractFontPairs(masterOrLayout) {
    const fontPairs = {};
    
    // Look in placeholder elements
    if (masterOrLayout.pageElements) {
      masterOrLayout.pageElements.forEach(element => {
        if (element.shape && element.shape.placeholder) {
          const placeholder = element.shape.placeholder;
          const textStyle = element.shape.text?.textElements?.[0]?.textRun?.style;
          
          if (textStyle) {
            fontPairs[placeholder.type] = {
              fontFamily: textStyle.fontFamily,
              fontSize: textStyle.fontSize,
              bold: textStyle.bold,
              italic: textStyle.italic
            };
          }
        }
      });
    }
    
    return fontPairs;
  }
  
  /**
   * Extract custom properties not in standard API
   */
  static _extractCustomProperties(object) {
    const custom = {};
    
    // Look for properties that aren't in the documented API
    Object.keys(object).forEach(key => {
      if (!['objectId', 'pageType', 'pageElements', 'layoutProperties'].includes(key)) {
        custom[key] = object[key];
      }
    });
    
    return custom;
  }
}

/**
 * Advanced Presentation wrapper class
 */
class AdvancedPresentation {
  
  constructor(presentation, presentationId) {
    this.presentation = presentation;
    this.presentationId = presentationId;
    this._cache = {};
  }
  
  // Proxy standard methods
  getSlides() { return this.presentation.getSlides(); }
  getId() { return this.presentation.getId(); }
  getName() { return this.presentation.getName(); }
  getUrl() { return this.presentation.getUrl(); }
  
  /**
   * Get advanced theme properties (tanaikech style)
   */
  getAdvancedTheme() {
    if (this._cache.advancedTheme) {
      return this._cache.advancedTheme;
    }
    
    const properties = SlidesAppAdvanced.getAdvancedProperties(this.presentationId);
    if (properties && properties.advancedProperties.masterThemes) {
      this._cache.advancedTheme = properties.advancedProperties.masterThemes[0];
      return this._cache.advancedTheme;
    }
    
    return null;
  }
  
  /**
   * Set advanced theme properties beyond standard API
   */
  setAdvancedTheme(themeOptions) {
    try {
      const requests = [];
      
      // Color scheme manipulation
      if (themeOptions.colorScheme) {
        requests.push(...this._buildColorSchemeRequests(themeOptions.colorScheme));
      }
      
      // Font manipulation
      if (themeOptions.fonts) {
        requests.push(...this._buildFontRequests(themeOptions.fonts));
      }
      
      // Master slide background
      if (themeOptions.masterBackground) {
        requests.push(...this._buildMasterBackgroundRequests(themeOptions.masterBackground));
      }
      
      if (requests.length > 0) {
        this._executeBatchUpdate(requests);
        this._cache = {}; // Clear cache
      }
      
      return this;
      
    } catch (error) {
      console.error('Failed to set advanced theme:', error);
      throw error;
    }
  }
  
  /**
   * Get font pairs (major/minor font relationships)
   */
  getFontPairs() {
    const theme = this.getAdvancedTheme();
    return theme ? theme.fontPairs : {};
  }
  
  /**
   * Set font pairs using advanced API techniques
   */
  setFontPairs(fontPairs) {
    try {
      const requests = [];
      
      Object.entries(fontPairs).forEach(([placeholderType, fontConfig]) => {
        requests.push({
          updateTextStyle: {
            objectId: this._getMasterObjectId(),
            style: {
              fontFamily: fontConfig.fontFamily,
              fontSize: fontConfig.fontSize ? { magnitude: fontConfig.fontSize, unit: 'PT' } : undefined,
              bold: fontConfig.bold,
              italic: fontConfig.italic
            },
            textRange: { type: 'ALL' },
            fields: 'fontFamily,fontSize,bold,italic'
          }
        });
      });
      
      if (requests.length > 0) {
        this._executeBatchUpdate(requests);
      }
      
      return this;
      
    } catch (error) {
      console.error('Failed to set font pairs:', error);
      throw error;
    }
  }
  
  /**
   * Access color palette beyond standard API
   */
  getColorPalette() {
    const theme = this.getAdvancedTheme();
    return theme ? theme.colorScheme : null;
  }
  
  /**
   * Set color palette using advanced techniques
   */
  setColorPalette(colors) {
    const requests = this._buildColorSchemeRequests(colors);
    if (requests.length > 0) {
      this._executeBatchUpdate(requests);
      this._cache = {};
    }
    return this;
  }
  
  /**
   * Get slide master properties not available in standard API
   */
  getMasterProperties() {
    const properties = SlidesAppAdvanced.getAdvancedProperties(this.presentationId);
    return properties ? properties.advancedProperties : null;
  }
  
  /**
   * Manipulate slide master beyond standard API
   */
  updateMasterProperties(masterOptions) {
    try {
      const requests = [];
      const masterId = this._getMasterObjectId();
      
      if (masterOptions.background) {
        requests.push({
          updatePageProperties: {
            objectId: masterId,
            pageProperties: {
              pageBackgroundFill: this._buildBackgroundFill(masterOptions.background)
            },
            fields: 'pageBackgroundFill'
          }
        });
      }
      
      if (masterOptions.elements) {
        masterOptions.elements.forEach(element => {
          requests.push(this._buildElementRequest(masterId, element));
        });
      }
      
      if (requests.length > 0) {
        this._executeBatchUpdate(requests);
      }
      
      return this;
      
    } catch (error) {
      console.error('Failed to update master properties:', error);
      throw error;
    }
  }
  
  /**
   * Export presentation with advanced options
   */
  exportAdvanced(options = {}) {
    try {
      let exportUrl = `https://www.googleapis.com/drive/v3/files/${this.presentationId}/export`;
      
      const params = [];
      if (options.format) params.push(`mimeType=${encodeURIComponent(options.format)}`);
      if (options.size) params.push(`size=${options.size}`);
      if (options.scale) params.push(`scale=${options.scale}`);
      
      if (params.length > 0) {
        exportUrl += '?' + params.join('&');
      }
      
      const response = UrlFetchApp.fetch(exportUrl, {
        headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
      });
      
      return response.getBlob();
      
    } catch (error) {
      console.error('Failed to export with advanced options:', error);
      throw error;
    }
  }
  
  /**
   * Get presentation as OOXML for advanced manipulation
   */
  toOOXML() {
    const pptxBlob = this.exportAdvanced({ 
      format: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' 
    });
    return new OOXMLParser(pptxBlob);
  }
  
  // Private helper methods
  
  _getMasterObjectId() {
    if (!this._cache.masterId) {
      const properties = SlidesAppAdvanced.getAdvancedProperties(this.presentationId);
      if (properties && properties.slides.masters && properties.slides.masters.length > 0) {
        this._cache.masterId = properties.slides.masters[0].objectId;
      } else {
        throw new Error('No slide master found');
      }
    }
    return this._cache.masterId;
  }
  
  _buildColorSchemeRequests(colorScheme) {
    const requests = [];
    const masterId = this._getMasterObjectId();
    
    // This is where tanaikech-style magic happens - manipulating colors
    // beyond what the standard API exposes
    Object.entries(colorScheme).forEach(([colorType, colorValue]) => {
      requests.push({
        updateSlideProperties: {
          objectId: masterId,
          slideProperties: {
            // Advanced color manipulation
            masterColorScheme: {
              [colorType]: {
                rgbColor: this._hexToRgb(colorValue)
              }
            }
          },
          fields: 'masterColorScheme'
        }
      });
    });
    
    return requests;
  }
  
  _buildFontRequests(fonts) {
    const requests = [];
    
    Object.entries(fonts).forEach(([fontType, fontFamily]) => {
      // Advanced font manipulation using undocumented properties
      requests.push({
        updatePresentationStyle: {
          style: {
            fontFamilyMapping: {
              [fontType]: fontFamily
            }
          },
          fields: 'fontFamilyMapping'
        }
      });
    });
    
    return requests;
  }
  
  _buildMasterBackgroundRequests(backgroundConfig) {
    const requests = [];
    const masterId = this._getMasterObjectId();
    
    requests.push({
      updatePageProperties: {
        objectId: masterId,
        pageProperties: {
          pageBackgroundFill: this._buildBackgroundFill(backgroundConfig)
        },
        fields: 'pageBackgroundFill'
      }
    });
    
    return requests;
  }
  
  _buildBackgroundFill(config) {
    if (config.solidFill) {
      return {
        solidFill: {
          color: {
            rgbColor: this._hexToRgb(config.solidFill.color)
          },
          alpha: config.solidFill.alpha || 1.0
        }
      };
    }
    
    if (config.gradientFill) {
      return {
        gradientFill: config.gradientFill
      };
    }
    
    return null;
  }
  
  _buildElementRequest(pageId, element) {
    return {
      createShape: {
        objectId: element.objectId || Utilities.getUuid(),
        shapeType: element.shapeType || 'TEXT_BOX',
        elementProperties: {
          pageObjectId: pageId,
          size: element.size,
          transform: element.transform
        }
      }
    };
  }
  
  _executeBatchUpdate(requests) {
    try {
      const body = JSON.stringify({ requests: requests });
      
      const response = UrlFetchApp.fetch(
        `https://slides.googleapis.com/v1/presentations/${this.presentationId}:batchUpdate`,
        {
          method: 'POST',
          headers: {
            'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
            'Content-Type': 'application/json'
          },
          payload: body
        }
      );
      
      const result = JSON.parse(response.getContentText());
      
      if (result.error) {
        throw new Error(`Batch update failed: ${result.error.message}`);
      }
      
      return result;
      
    } catch (error) {
      console.error('Batch update error:', error);
      throw error;
    }
  }
  
  _hexToRgb(hex) {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
      red: parseInt(result[1], 16) / 255,
      green: parseInt(result[2], 16) / 255,
      blue: parseInt(result[3], 16) / 255
    } : null;
  }
}