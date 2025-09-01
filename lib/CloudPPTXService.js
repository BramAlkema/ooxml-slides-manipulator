/**
 * CloudPPTXService v2 - Google Apps Script ↔ Cloud Function Bridge
 * 
 * PURPOSE:
 * Provides reliable ZIP/unzip capabilities for OOXML files by bridging Google Apps Script
 * with a Google Cloud Function. Solves the limitation that GAS Utilities.unzip() cannot
 * handle most real-world OOXML files due to compression differences.
 * 
 * ARCHITECTURE:
 * - Primary: Uses Google Cloud Function with JSZip for reliable extraction/compression
 * - Fallback: Native GAS Utilities.unzip/zip (limited compatibility)
 * - Error Handling: Comprehensive error codes and retry logic
 * - Monitoring: Detailed logging and performance metrics
 * 
 * AI CONTEXT:
 * This is the reliability layer for all OOXML operations. When OOXML files need to be
 * extracted or compressed, use this service. It handles network errors, Cloud Function
 * failures, and provides fallback mechanisms. Always use this instead of direct
 * Utilities.unzip for OOXML files.
 * 
 * TESTING STRATEGY:
 * Mock UrlFetchApp and Utilities for unit testing. Test network failures, malformed
 * responses, and fallback scenarios. Verify error codes and retry behavior.
 */

class CloudPPTXService {
  
  /**
   * Configuration and constants
   * @private
   */
  static get _CONFIG() {
    return {
      // Cloud Function endpoint
      cloudFunctionUrl: 'https://pptx-router-kmoetjdbla-uc.a.run.app',
      
      // Request settings
      requestTimeout: 30000, // 30 seconds
      retryAttempts: 3,
      retryDelayMs: 1000,
      
      // Size limits
      maxFileSizeMB: 50,
      maxFileCount: 1000,
      
      // Endpoints
      endpoints: {
        unzip: '/unzip',
        zip: '/zip',
        health: '/health'
      }
    };
  }
  
  /**
   * Error codes for systematic error handling
   * @private
   */
  static get _ERROR_CODES() {
    return {
      CLOUD_FUNCTION_001: 'Cloud Function not available',
      CLOUD_FUNCTION_002: 'Invalid response from Cloud Function',
      CLOUD_FUNCTION_003: 'File size exceeds limit',
      CLOUD_FUNCTION_004: 'Network request failed',
      CLOUD_FUNCTION_005: 'Invalid input parameters',
      CLOUD_FUNCTION_006: 'Fallback operation failed',
      CLOUD_FUNCTION_007: 'Compression/decompression failed'
    };
  }
  
  /**
   * Extract OOXML file into individual files
   * 
   * @param {Blob} ooxmlBlob - The OOXML file to extract (pptx, docx, xlsx)
   * @param {Object} options - Extraction options
   * @param {boolean} options.useFallback - Try fallback if Cloud Function fails
   * @param {boolean} options.validateFiles - Validate extracted files
   * @returns {Promise<Object>} Object mapping filenames to content
   * @throws {Error} With specific error codes for different failure types
   * 
   * AI_USAGE: Use this to extract any OOXML file. Handles all error scenarios.
   * Example: const files = await CloudPPTXService.extractFiles(pptxBlob)
   */
  static async extractFiles(ooxmlBlob, options = {}) {
    const startTime = Date.now();
    const opts = {
      useFallback: true,
      validateFiles: true,
      ...options
    };
    
    // Input validation
    this._validateExtractInput(ooxmlBlob);
    
    try {
      this._log(`🔄 Extracting OOXML file: ${ooxmlBlob.getBytes().length} bytes`);
      
      // Try Cloud Function first
      const result = await this._extractViaCloudFunction(ooxmlBlob);
      
      // Validate results if requested
      if (opts.validateFiles) {
        this._validateExtractedFiles(result);
      }
      
      const duration = Date.now() - startTime;
      this._log(`✅ Extracted ${Object.keys(result).length} files in ${duration}ms`);
      
      return result;
      
    } catch (error) {
      this._log(`❌ Cloud Function extraction failed: ${error.message}`);
      
      // Try fallback if enabled
      if (opts.useFallback) {
        this._log('🔄 Attempting fallback extraction...');
        return this._extractViaFallback(ooxmlBlob);
      }
      
      // Re-throw with context
      const wrappedError = new Error(`${this._ERROR_CODES.CLOUD_FUNCTION_007}: ${error.message}`);
      wrappedError.originalError = error;
      wrappedError.duration = Date.now() - startTime;
      throw wrappedError;
    }
  }
  
  /**
   * Compress files into an OOXML blob
   * 
   * @param {Object} files - Object mapping filenames to content strings
   * @param {Object} options - Compression options
   * @param {string} options.mimeType - MIME type for the output blob
   * @param {string} options.filename - Filename for the output blob
   * @param {boolean} options.useFallback - Try fallback if Cloud Function fails
   * @returns {Promise<Blob>} Compressed OOXML blob
   * @throws {Error} With specific error codes for different failure types
   * 
   * AI_USAGE: Use this to create OOXML files from extracted content.
   * Example: const pptxBlob = await CloudPPTXService.compressFiles(files)
   */
  static async compressFiles(files, options = {}) {
    const startTime = Date.now();
    const opts = {
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      filename: 'document.pptx',
      useFallback: true,
      ...options
    };
    
    // Input validation
    this._validateCompressInput(files);
    
    try {
      this._log(`🔄 Compressing ${Object.keys(files).length} files to OOXML`);
      
      // Try Cloud Function first
      const result = await this._compressViaCloudFunction(files, opts);
      
      const duration = Date.now() - startTime;
      this._log(`✅ Compressed to ${result.getBytes().length} bytes in ${duration}ms`);
      
      return result;
      
    } catch (error) {
      this._log(`❌ Cloud Function compression failed: ${error.message}`);
      
      // Try fallback if enabled
      if (opts.useFallback) {
        this._log('🔄 Attempting fallback compression...');
        return this._compressViaFallback(files, opts);
      }
      
      // Re-throw with context
      const wrappedError = new Error(`${this._ERROR_CODES.CLOUD_FUNCTION_007}: ${error.message}`);
      wrappedError.originalError = error;
      wrappedError.duration = Date.now() - startTime;
      throw wrappedError;
    }
  }
  
  /**
   * Check if Cloud Function is available and healthy
   * 
   * @returns {Promise<Object>} Health check results
   * 
   * AI_USAGE: Use this to verify Cloud Function availability before operations.
   * Example: const health = await CloudPPTXService.healthCheck()
   */
  static async healthCheck() {
    const startTime = Date.now();
    
    try {
      this._log('🔍 Performing Cloud Function health check...');
      
      const response = await this._makeRequest(this._CONFIG.endpoints.health, {
        method: 'GET',
        timeout: 5000
      });
      
      const duration = Date.now() - startTime;
      const isHealthy = response.getResponseCode() < 400;
      
      const result = {
        available: isHealthy,
        responseTime: duration,
        statusCode: response.getResponseCode(),
        timestamp: new Date().toISOString(),
        url: this._CONFIG.cloudFunctionUrl
      };
      
      if (isHealthy) {
        this._log(`✅ Cloud Function healthy (${duration}ms)`);
      } else {
        this._log(`❌ Cloud Function unhealthy: ${response.getResponseCode()}`);
      }
      
      return result;
      
    } catch (error) {
      const duration = Date.now() - startTime;
      this._log(`❌ Health check failed: ${error.message}`);
      
      return {
        available: false,
        responseTime: duration,
        error: error.message,
        timestamp: new Date().toISOString(),
        url: this._CONFIG.cloudFunctionUrl
      };
    }
  }
  
  /**
   * Get service statistics and configuration
   * 
   * @returns {Object} Service information
   * 
   * AI_USAGE: Use this for debugging and monitoring.
   * Example: const info = CloudPPTXService.getServiceInfo()
   */
  static getServiceInfo() {
    return {
      version: '2.0.0',
      cloudFunctionUrl: this._CONFIG.cloudFunctionUrl,
      maxFileSizeMB: this._CONFIG.maxFileSizeMB,
      maxFileCount: this._CONFIG.maxFileCount,
      requestTimeout: this._CONFIG.requestTimeout,
      retryAttempts: this._CONFIG.retryAttempts,
      capabilities: {
        extraction: true,
        compression: true,
        fallback: true,
        healthCheck: true,
        errorRecovery: true
      },
      supportedFormats: ['pptx', 'docx', 'xlsx', 'zip']
    };
  }
  
  // PRIVATE METHODS - Internal implementation details
  
  /**
   * Extract via Cloud Function
   * @private
   */
  static async _extractViaCloudFunction(ooxmlBlob) {
    const base64Data = Utilities.base64Encode(ooxmlBlob.getBytes());
    const payload = { data: base64Data };
    
    const response = await this._makeRequest(this._CONFIG.endpoints.unzip, {
      method: 'POST',
      payload: JSON.stringify(payload),
      headers: { 'Content-Type': 'application/json' }
    });
    
    const result = JSON.parse(response.getContentText());
    
    if (!result.success) {
      throw new Error(result.error || 'Unknown Cloud Function error');
    }
    
    if (!result.files || typeof result.files !== 'object') {
      throw new Error(this._ERROR_CODES.CLOUD_FUNCTION_002);
    }
    
    return result.files;
  }
  
  /**
   * Compress via Cloud Function
   * @private
   */
  static async _compressViaCloudFunction(files, options) {
    const payload = { files };
    
    const response = await this._makeRequest(this._CONFIG.endpoints.zip, {
      method: 'POST',
      payload: JSON.stringify(payload),
      headers: { 'Content-Type': 'application/json' }
    });
    
    const result = JSON.parse(response.getContentText());
    
    if (!result.success) {
      throw new Error(result.error || 'Unknown Cloud Function error');
    }
    
    if (!result.pptxData) {
      throw new Error(this._ERROR_CODES.CLOUD_FUNCTION_002);
    }
    
    // Convert base64 back to blob
    const binaryData = Utilities.base64Decode(result.pptxData);
    return Utilities.newBlob(binaryData, options.mimeType, options.filename);
  }
  
  /**
   * Fallback extraction using GAS native methods
   * @private
   */
  static _extractViaFallback(ooxmlBlob) {
    try {
      this._log('🔄 Using GAS native unzip fallback...');
      
      const files = Utilities.unzip(ooxmlBlob);
      const fileMap = {};
      
      files.forEach(file => {
        if (!file.isGoogleType()) {
          fileMap[file.getName()] = file.getDataAsString();
        }
      });
      
      if (Object.keys(fileMap).length === 0) {
        throw new Error('No files extracted via fallback method');
      }
      
      this._log(`✅ Fallback extracted ${Object.keys(fileMap).length} files`);
      return fileMap;
      
    } catch (error) {
      throw new Error(`${this._ERROR_CODES.CLOUD_FUNCTION_006}: ${error.message}`);
    }
  }
  
  /**
   * Fallback compression using GAS native methods
   * @private
   */
  static _compressViaFallback(files, options) {
    try {
      this._log('🔄 Using GAS native zip fallback...');
      
      const blobs = [];
      Object.entries(files).forEach(([filename, content]) => {
        blobs.push(Utilities.newBlob(content, 'text/plain', filename));
      });
      
      const zipBlob = Utilities.zip(blobs, options.filename);
      return zipBlob.setContentType(options.mimeType);
      
    } catch (error) {
      throw new Error(`${this._ERROR_CODES.CLOUD_FUNCTION_006}: ${error.message}`);
    }
  }
  
  /**
   * Make HTTP request with retry logic
   * @private
   */
  static async _makeRequest(endpoint, options = {}) {
    const url = `${this._CONFIG.cloudFunctionUrl}${endpoint}`;
    const requestOptions = this._buildRequestOptions(options);
    
    let lastError;
    
    for (let attempt = 1; attempt <= this._CONFIG.retryAttempts; attempt++) {
      try {
        const response = await this._executeRequest(url, requestOptions, attempt);
        const result = this._handleResponse(response, attempt);
        
        if (result.success) {
          return result.response;
        }
        
        lastError = result.error;
        
      } catch (error) {
        lastError = error;
        if (!this._shouldRetry(error, attempt)) {
          break;
        }
        this._delayBeforeRetry(error.message);
      }
    }
    
    throw new Error(`${this._ERROR_CODES.CLOUD_FUNCTION_004}: ${lastError.message}`);
  }
  
  /**
   * Build request options with defaults
   * @private
   */
  static _buildRequestOptions(options) {
    return {
      method: 'GET',
      muteHttpExceptions: true,
      timeout: this._CONFIG.requestTimeout,
      ...options
    };
  }
  
  /**
   * Execute HTTP request
   * @private
   */
  static async _executeRequest(url, requestOptions, attempt) {
    this._log(`📡 HTTP ${requestOptions.method} ${url} (attempt ${attempt})`);
    return UrlFetchApp.fetch(url, requestOptions);
  }
  
  /**
   * Handle HTTP response and determine retry strategy
   * @private
   */
  static _handleResponse(response, attempt) {
    const statusCode = response.getResponseCode();
    
    if (statusCode >= 200 && statusCode < 400) {
      return { success: true, response };
    }
    
    const errorMessage = `HTTP ${statusCode}: ${response.getContentText()}`;
    
    // Server error - worth retrying
    if (statusCode >= 500 && attempt < this._CONFIG.retryAttempts) {
      this._delayBeforeRetry(`Server error ${statusCode}`);
      return { success: false, error: new Error(errorMessage) };
    }
    
    // Client error or final attempt - don't retry
    throw new Error(errorMessage);
  }
  
  /**
   * Check if error should trigger retry
   * @private
   */
  static _shouldRetry(error, attempt) {
    return attempt < this._CONFIG.retryAttempts && this._isRetryableError(error);
  }
  
  /**
   * Add delay before retry
   * @private
   */
  static _delayBeforeRetry(reason) {
    this._log(`⏳ Retrying in ${this._CONFIG.retryDelayMs}ms (${reason})`);
    Utilities.sleep(this._CONFIG.retryDelayMs);
  }
  
  /**
   * Determine if an error is worth retrying
   * @private
   */
  static _isRetryableError(error) {
    const retryableMessages = [
      'timeout',
      'network',
      'connection',
      'temporary',
      'unavailable'
    ];
    
    const errorMessage = error.message.toLowerCase();
    return retryableMessages.some(msg => errorMessage.includes(msg));
  }
  
  /**
   * Validate extraction input
   * @private
   */
  static _validateExtractInput(ooxmlBlob) {
    if (!ooxmlBlob) {
      throw new Error(`${this._ERROR_CODES.CLOUD_FUNCTION_005}: Missing blob parameter`);
    }
    
    if (typeof ooxmlBlob.getBytes !== 'function') {
      throw new Error(`${this._ERROR_CODES.CLOUD_FUNCTION_005}: Invalid blob object`);
    }
    
    const sizeBytes = ooxmlBlob.getBytes().length;
    const sizeMB = sizeBytes / (1024 * 1024);
    
    if (sizeMB > this._CONFIG.maxFileSizeMB) {
      throw new Error(
        `${this._ERROR_CODES.CLOUD_FUNCTION_003}: File size ${sizeMB.toFixed(2)}MB exceeds limit of ${this._CONFIG.maxFileSizeMB}MB`
      );
    }
  }
  
  /**
   * Validate compression input
   * @private
   */
  static _validateCompressInput(files) {
    if (!files || typeof files !== 'object') {
      throw new Error(`${this._ERROR_CODES.CLOUD_FUNCTION_005}: Files must be an object`);
    }
    
    const fileCount = Object.keys(files).length;
    if (fileCount === 0) {
      throw new Error(`${this._ERROR_CODES.CLOUD_FUNCTION_005}: No files provided`);
    }
    
    if (fileCount > this._CONFIG.maxFileCount) {
      throw new Error(
        `${this._ERROR_CODES.CLOUD_FUNCTION_003}: File count ${fileCount} exceeds limit of ${this._CONFIG.maxFileCount}`
      );
    }
    
    // Validate file contents
    for (const [filename, content] of Object.entries(files)) {
      if (typeof content !== 'string' && typeof content !== 'object') {
        throw new Error(`${this._ERROR_CODES.CLOUD_FUNCTION_005}: Invalid content for file ${filename}`);
      }
    }
  }
  
  /**
   * Validate extracted files
   * @private
   */
  static _validateExtractedFiles(files) {
    if (!files || typeof files !== 'object') {
      throw new Error(`${this._ERROR_CODES.CLOUD_FUNCTION_002}: Invalid files object`);
    }
    
    const fileCount = Object.keys(files).length;
    if (fileCount === 0) {
      throw new Error(`${this._ERROR_CODES.CLOUD_FUNCTION_002}: No files extracted`);
    }
    
    // Check for essential OOXML files
    const essentialFiles = ['[Content_Types].xml', '_rels/.rels'];
    const hasEssentialFiles = essentialFiles.some(filename => files.hasOwnProperty(filename));
    
    if (!hasEssentialFiles) {
      this._log('⚠️ Warning: No essential OOXML files found in extraction');
    }
  }
  
  /**
   * Log message with timestamp
   * @private
   */
  static _log(message) {
    const timestamp = new Date().toISOString().substr(11, 8);
    console.log(`[${timestamp}] [CloudPPTXService] ${message}`);
  }
}

// Backward compatibility aliases
CloudPPTXService.unzipPPTX = CloudPPTXService.extractFiles;
CloudPPTXService.zipPPTX = CloudPPTXService.compressFiles;

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = CloudPPTXService;
}