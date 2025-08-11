/**
 * Validators - Input validation and error handling utilities
 */
class Validators {
  /**
   * Validate hex color format
   * @param {string} color - Hex color to validate
   * @returns {boolean}
   */
  static isValidHexColor(color) {
    if (typeof color !== 'string') return false;
    
    // Remove # if present
    const cleanColor = color.replace('#', '');
    
    // Check if 3 or 6 character hex
    return /^[0-9A-Fa-f]{3}$|^[0-9A-Fa-f]{6}$/.test(cleanColor);
  }

  /**
   * Normalize hex color (ensure 6 digits with #)
   * @param {string} color - Hex color
   * @returns {string} Normalized hex color
   */
  static normalizeHexColor(color) {
    if (!this.isValidHexColor(color)) {
      throw new Error(`Invalid hex color: ${color}`);
    }
    
    let cleanColor = color.replace('#', '').toUpperCase();
    
    // Expand 3-digit to 6-digit
    if (cleanColor.length === 3) {
      cleanColor = cleanColor.split('').map(char => char + char).join('');
    }
    
    return '#' + cleanColor;
  }

  /**
   * Validate color palette object
   * @param {Object} palette - Color palette
   * @returns {boolean}
   */
  static isValidColorPalette(palette) {
    if (!palette || typeof palette !== 'object') return false;
    
    const requiredColors = ['dk1', 'lt1', 'accent1'];
    const validColors = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
    
    // Check required colors exist
    for (let color of requiredColors) {
      if (!palette[color]) return false;
    }
    
    // Check all provided colors are valid
    for (let [colorName, colorValue] of Object.entries(palette)) {
      if (!validColors.includes(colorName)) return false;
      if (!this.isValidHexColor(colorValue)) return false;
    }
    
    return true;
  }

  /**
   * Validate font name
   * @param {string} fontName - Font name
   * @returns {boolean}
   */
  static isValidFontName(fontName) {
    if (typeof fontName !== 'string') return false;
    if (fontName.length === 0 || fontName.length > 100) return false;
    
    // Basic validation - no control characters
    return !/[\x00-\x1F\x7F]/.test(fontName);
  }

  /**
   * Validate slide dimensions
   * @param {Object} size - {width, height} object
   * @returns {boolean}
   */
  static isValidSlideSize(size) {
    if (!size || typeof size !== 'object') return false;
    if (typeof size.width !== 'number' || typeof size.height !== 'number') return false;
    if (size.width <= 0 || size.height <= 0) return false;
    if (size.width > 10000 || size.height > 10000) return false; // Reasonable limits
    
    return true;
  }

  /**
   * Validate Google Drive file ID
   * @param {string} fileId - File ID
   * @returns {boolean}
   */
  static isValidFileId(fileId) {
    if (typeof fileId !== 'string') return false;
    
    // Google Drive file IDs are typically 44 characters
    // Can contain letters, numbers, underscores, and hyphens
    return /^[a-zA-Z0-9_-]{10,50}$/.test(fileId);
  }

  /**
   * Validate OOXML file path
   * @param {string} path - File path in OOXML package
   * @returns {boolean}
   */
  static isValidOOXMLPath(path) {
    if (typeof path !== 'string') return false;
    if (path.length === 0) return false;
    
    // Basic path validation
    return !/^\/|\/\/|[<>:"|?*]/.test(path);
  }

  /**
   * Sanitize file name
   * @param {string} fileName - Original file name
   * @returns {string} Sanitized file name
   */
  static sanitizeFileName(fileName) {
    if (typeof fileName !== 'string') return 'untitled';
    
    // Remove invalid characters
    let sanitized = fileName.replace(/[<>:"/\\|?*\x00-\x1f]/g, '_');
    
    // Limit length
    if (sanitized.length > 100) {
      const ext = sanitized.substring(sanitized.lastIndexOf('.'));
      sanitized = sanitized.substring(0, 97 - ext.length) + ext;
    }
    
    return sanitized || 'untitled';
  }

  /**
   * Validate theme name
   * @param {string} themeName - Theme name
   * @returns {boolean}
   */
  static isValidThemeName(themeName) {
    const validThemes = ['office', 'colorful', 'dark', 'minimal'];
    return validThemes.includes(themeName);
  }

  /**
   * Validate placeholder type
   * @param {string} type - Placeholder type
   * @returns {boolean}
   */
  static isValidPlaceholderType(type) {
    const validTypes = [
      'title', 'body', 'ctrTitle', 'subTitle', 'dt', 'sldNum', 'ftr', 'hdr',
      'obj', 'chart', 'tbl', 'clipArt', 'dgm', 'media', 'sldImg', 'pic'
    ];
    return validTypes.includes(type);
  }
}

/**
 * ErrorHandler - Centralized error handling
 */
class ErrorHandler {
  /**
   * Handle OOXML parsing errors
   * @param {Error} error - Original error
   * @param {string} context - Context where error occurred
   * @throws {Error} Enhanced error
   */
  static handleOOXMLError(error, context) {
    let message = `OOXML Error in ${context}: ${error.message}`;
    
    // Enhance common errors
    if (error.message.includes('zip')) {
      message += '\nTip: File may be corrupted or not a valid PPTX file.';
    } else if (error.message.includes('XML')) {
      message += '\nTip: XML structure may be invalid or unsupported.';
    } else if (error.message.includes('namespace')) {
      message += '\nTip: XML namespace issue - file format may not be supported.';
    }
    
    throw new Error(message);
  }

  /**
   * Handle Drive API errors
   * @param {Error} error - Original error
   * @param {string} fileId - File ID if applicable
   * @throws {Error} Enhanced error
   */
  static handleDriveError(error, fileId = null) {
    let message = `Drive API Error: ${error.message}`;
    
    if (fileId) {
      message += `\nFile ID: ${fileId}`;
    }
    
    // Enhance common errors
    if (error.message.includes('not found')) {
      message += '\nTip: Check if file exists and you have access permissions.';
    } else if (error.message.includes('permission')) {
      message += '\nTip: Insufficient permissions. Ensure you can edit this file.';
    } else if (error.message.includes('quota')) {
      message += '\nTip: Storage quota exceeded. Free up space or upgrade storage.';
    }
    
    throw new Error(message);
  }

  /**
   * Validate operation parameters
   * @param {Object} params - Parameters to validate
   * @param {Object} rules - Validation rules
   * @throws {Error} Validation error
   */
  static validateParams(params, rules) {
    for (let [param, rule] of Object.entries(rules)) {
      const value = params[param];
      
      if (rule.required && (value === undefined || value === null)) {
        throw new Error(`Required parameter missing: ${param}`);
      }
      
      if (value !== undefined && rule.type && typeof value !== rule.type) {
        throw new Error(`Parameter ${param} must be of type ${rule.type}`);
      }
      
      if (rule.validator && !rule.validator(value)) {
        throw new Error(`Invalid parameter ${param}: ${value}`);
      }
    }
  }

  /**
   * Create user-friendly error messages
   * @param {Error} error - Original error
   * @returns {string} User-friendly message
   */
  static getUserMessage(error) {
    const message = error.message.toLowerCase();
    
    if (message.includes('file not found')) {
      return 'The file could not be found. Please check the file ID and your permissions.';
    } else if (message.includes('invalid hex color')) {
      return 'Please provide valid hex colors (e.g., #FF0000 or #F00).';
    } else if (message.includes('invalid font')) {
      return 'Please provide a valid font name.';
    } else if (message.includes('zip') || message.includes('corrupt')) {
      return 'The file appears to be corrupted or is not a valid PowerPoint file.';
    } else {
      return 'An unexpected error occurred. Please try again or contact support.';
    }
  }
}