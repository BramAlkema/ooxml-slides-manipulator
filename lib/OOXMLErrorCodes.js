/**
 * OOXMLErrorCodes - Systematic Error Taxonomy for Production Debugging
 * 
 * PURPOSE:
 * Provides a systematic, greppable error taxonomy for the OOXML JSON system.
 * Each error code is unique, contextual, and designed for fast debugging in
 * production environments.
 * 
 * ARCHITECTURE:
 * - C00x: Core ZIP/XML operations
 * - S01x: Service/Network operations  
 * - A02x: Application/Slides semantics
 * - E03x: Extension framework
 * - V04x: Validation and compliance
 * 
 * AI CONTEXT:
 * Use these specific error codes throughout the system instead of generic errors.
 * This enables systematic debugging, alerting, and log analysis in production.
 * Each error includes context for detailed troubleshooting.
 */

class OOXMLUniqueError extends Error {
  /**
   * Create a systematic OOXML error with code and context
   * 
   * @param {string} code - Error code (e.g., 'C001', 'S010')
   * @param {string} message - Human-readable error message
   * @param {Object} context - Additional context for debugging
   * @param {string} correlationId - Request correlation ID
   */
  constructor(code, message, context = {}, correlationId = null) {
    super(message);
    this.name = 'OOXMLUniqueError';
    this.code = code;
    this.context = context;
    this.correlationId = correlationId;
    this.timestamp = new Date().toISOString();
    
    // Ensure error is properly captured in stack traces
    if (Error.captureStackTrace) {
      Error.captureStackTrace(this, OOXMLUniqueError);
    }
  }
  
  /**
   * Generate greppable log format
   * @returns {string} Formatted log line
   */
  toLogFormat() {
    const ctx = {
      ...this.context,
      corr: this.correlationId,
      timestamp: this.timestamp
    };
    return `ERR[${this.code}] ${this.message} ctx=${JSON.stringify(ctx)}`;
  }
  
  /**
   * Convert to JSON for API responses
   * @returns {Object} Structured error object
   */
  toJSON() {
    return {
      error: true,
      code: this.code,
      message: this.message,
      context: this.context,
      correlation: this.correlationId,
      timestamp: this.timestamp
    };
  }
}

/**
 * Error Code Registry with systematic taxonomy
 * Compatible with Google Apps Script V8 runtime
 */
class OOXMLErrorCodes {
  
  /**
   * Get error code metadata
   * @param {string} code - Error code
   * @returns {Object} Error metadata
   */
  static getMetadata(code) {
    const category = this._getCategory(code);
    const description = this._getDescription(code);
    const severity = this._getSeverity(code);
    
    return {
      code,
      category,
      description,
      severity,
      isRetryable: this._isRetryable(code),
      alerting: this._getAlertingLevel(code)
    };
  }
  
  /**
   * Create standardized error with proper context
   * @param {string} code - Error code
   * @param {string} message - Error message
   * @param {Object} context - Additional context
   * @param {string} correlationId - Request correlation ID
   * @returns {OOXMLError} Structured error
   */
  static create(code, message, context = {}, correlationId = null) {
    return new OOXMLUniqueError(code, message, context, correlationId);
  }
  
  /**
   * Check if error code is valid
   * @param {string} code - Error code to validate
   * @returns {boolean} True if valid error code
   */
  static isValid(code) {
    return Object.values(this.CODES).includes(code);
  }
  
  /**
   * Get all error codes by category
   * @param {string} category - Category prefix (C, S, A, E, V)
   * @returns {Array} Error codes in category
   */
  static getByCategory(category) {
    const prefix = category.toUpperCase();
    return Object.values(this.CODES).filter(code => 
      typeof code === 'string' && code.startsWith(prefix)
    );
  }
  
  // PRIVATE METHODS
  
  static _getCategory(code) {
    const categories = {
      'C': 'Core ZIP/XML',
      'S': 'Service/Network', 
      'A': 'Application/Slides',
      'E': 'Extension Framework',
      'V': 'Validation/Compliance'
    };
    return categories[code.charAt(0)] || 'Unknown';
  }
  
  static _getDescription(code) {
    const descriptions = {
      'C001': 'Bad ZIP file format',
      'C002': 'Missing required OOXML part',
      'C003': 'XML parsing failed',
      'C004': 'ZIP file corrupt or damaged',
      'C005': 'Invalid OOXML structure',
      'S010': 'HTTP 4xx client error',
      'S011': 'HTTP 5xx server error', 
      'S012': 'Request timeout',
      'A020': 'Theme file not found',
      'A021': 'Layout mapping missing',
      'E030': 'Extension not found',
      'E031': 'Extension execution failed',
      'V040': 'Rule parsing error',
      'V041': 'Rule execution error'
    };
    return descriptions[code] || 'Unknown error';
  }
  
  static _getSeverity(code) {
    const firstChar = code.charAt(0);
    const severityMap = {
      'C': 'HIGH',    // Core errors are always high severity
      'S': 'MEDIUM',  // Service errors vary
      'A': 'HIGH',    // Application errors are important  
      'E': 'LOW',     // Extension errors are usually recoverable
      'V': 'MEDIUM'   // Validation errors are warnings
    };
    return severityMap[firstChar] || 'LOW';
  }
  
  static _isRetryable(code) {
    const retryableCodes = ['S012', 'S014', 'S015', 'S017', 'E034'];
    return retryableCodes.includes(code);
  }
  
  static _getAlertingLevel(code) {
    const severity = this._getSeverity(code);
    const alertMap = {
      'HIGH': 'IMMEDIATE',
      'MEDIUM': 'WARNING', 
      'LOW': 'INFO'
    };
    return alertMap[severity] || 'INFO';
  }
}

// Error codes defined as static property for GAS compatibility
OOXMLErrorCodes.CODES = {
  // Core ZIP/XML Operations (C00x)
  C001_BAD_ZIP: 'C001',
  C002_MISSING_PART: 'C002',
  C003_XML_PARSE: 'C003',
  C004_ZIP_CORRUPT: 'C004',
  C005_INVALID_OOXML: 'C005',
  C006_COMPRESSION_FAILED: 'C006',
  C007_EXTRACTION_FAILED: 'C007',
  C008_CONTENT_TYPES_MISSING: 'C008',
  C009_RELATIONSHIP_MISSING: 'C009',
  
  // Service/Network Operations (S01x)
  S010_HTTP_4XX: 'S010',
  S011_HTTP_5XX: 'S011',
  S012_TIMEOUT: 'S012',
  S013_RETRY_EXHAUSTED: 'S013',
  S014_NETWORK_ERROR: 'S014',
  S015_SERVICE_UNAVAILABLE: 'S015',
  S016_AUTHENTICATION_FAILED: 'S016',
  S017_RATE_LIMITED: 'S017',
  S018_REQUEST_TOO_LARGE: 'S018',
  S019_SESSION_EXPIRED: 'S019',
  
  // Application/Slides Semantics (A02x)
  A020_THEME_NOT_FOUND: 'A020',
  A021_LAYOUT_MAP_MISSING: 'A021',
  A022_SLIDE_NOT_FOUND: 'A022',
  A023_MASTER_MISSING: 'A023',
  A024_INVALID_SLIDE_ID: 'A024',
  A025_PRESENTATION_CORRUPT: 'A025',
  A026_UNSUPPORTED_FORMAT: 'A026',
  A027_BRAND_CONFIG_INVALID: 'A027',
  A028_COLOR_PARSE_FAILED: 'A028',
  A029_FONT_NOT_AVAILABLE: 'A029',
  
  // Extension Framework (E03x)
  E030_EXTENSION_MISSING: 'E030',
  E031_EXTENSION_FAILED: 'E031',
  E032_BAD_CONFIG: 'E032',
  E033_DEPENDENCY_MISSING: 'E033',
  E034_EXTENSION_TIMEOUT: 'E034',
  E035_INVALID_EXTENSION_TYPE: 'E035',
  E036_EXTENSION_CONFLICT: 'E036',
  E037_REGISTRATION_FAILED: 'E037',
  E038_LIFECYCLE_ERROR: 'E038',
  E039_EXTENSION_DISABLED: 'E039',
  
  // Validation and Compliance (V04x)
  V040_RULE_PARSE: 'V040',
  V041_RULE_EXEC: 'V041',
  V042_XPATH_INVALID: 'V042',
  V043_VALIDATION_FAILED: 'V043',
  V044_AUTOFIX_FAILED: 'V044',
  V045_COMPLIANCE_SCHEMA_INVALID: 'V045',
  V046_RULE_CONFLICT: 'V046',
  V047_PROFILE_NOT_FOUND: 'V047',
  V048_NAMESPACE_ERROR: 'V048',
  V049_ASSERTION_FAILED: 'V049'
};

// Add convenience accessors for direct access
Object.keys(OOXMLErrorCodes.CODES).forEach(key => {
  OOXMLErrorCodes[key] = OOXMLErrorCodes.CODES[key];
});

/**
 * Utility functions for common error scenarios
 */
const OOXML_ERRORS = {
  /**
   * Create a ZIP processing error
   */
  badZip: (message, context, correlationId) => 
    OOXMLErrorCodes.create(OOXMLErrorCodes.CODES.C001_BAD_ZIP, message, context, correlationId),
  
  /**
   * Create a missing part error
   */
  missingPart: (partName, context, correlationId) => 
    OOXMLErrorCodes.create(OOXMLErrorCodes.CODES.C002_MISSING_PART, `Missing OOXML part: ${partName}`, context, correlationId),
  
  /**
   * Create an XML parsing error
   */
  xmlParse: (xmlPath, parseError, context, correlationId) => 
    OOXMLErrorCodes.create(OOXMLErrorCodes.CODES.C003_XML_PARSE, `XML parse failed: ${xmlPath} - ${parseError}`, context, correlationId),
  
  /**
   * Create a network timeout error
   */
  timeout: (operation, timeoutMs, context, correlationId) =>
    OOXMLErrorCodes.create(OOXMLErrorCodes.CODES.S012_TIMEOUT, `Operation timeout: ${operation} (${timeoutMs}ms)`, context, correlationId),
  
  /**
   * Create a theme not found error
   */
  themeNotFound: (themeId, context, correlationId) =>
    OOXMLErrorCodes.create(OOXMLErrorCodes.CODES.A020_THEME_NOT_FOUND, `Theme not found: ${themeId}`, context, correlationId),
  
  /**
   * Create an extension failure error
   */
  extensionFailed: (extensionName, error, context, correlationId) =>
    OOXMLErrorCodes.create(OOXMLErrorCodes.CODES.E031_EXTENSION_FAILED, `Extension failed: ${extensionName} - ${error}`, context, correlationId)
};