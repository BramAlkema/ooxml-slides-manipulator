/**
 * FFlatePPTXService - Client-side PPTX ZIP operations with fflate
 * 
 * PURPOSE:
 * Provides fast, reliable ZIP/unzip capabilities for OOXML files using the fflate
 * library directly in Google Apps Script. Eliminates network dependencies and
 * provides better performance than CloudPPTXService.
 * 
 * ARCHITECTURE:
 * - Client-side: Uses fflate library for ZIP operations
 * - Synchronous: No network calls, immediate processing
 * - Memory efficient: Streams data processing
 * - Type handling: Automatic Uint8Array ↔ String conversion
 * 
 * DEPENDENCIES:
 * - fflate library (must be included in GAS project)
 * 
 * AI CONTEXT:
 * This replaces CloudPPTXService. Use for all OOXML ZIP operations.
 * Handles both text (XML) and binary files automatically.
 * 
 * TESTING STRATEGY:
 * Mock fflate functions for unit testing. Test binary/text file handling,
 * error scenarios, and conversion utilities.
 */

class FFlatePPTXService {
  
  /**
   * Configuration and constants
   * @private
   */
  static get _CONFIG() {
    return {
      // Size limits
      maxFileSizeMB: 100,
      maxFileCount: 2000,
      
      // Text file extensions (will be converted from Uint8Array to string)
      textExtensions: new Set([
        '.xml', '.rels', '.txt', '.json'
      ]),
      
      // Known text files by path
      textFiles: new Set([
        '[Content_Types].xml',
        '_rels/.rels'
      ]),
      
      // Encoding
      encoding: 'utf-8'
    };
  }

  /**
   * Extract files from PPTX/OOXML blob
   * @param {Blob} ooxmlBlob - The OOXML file to extract
   * @param {Object} options - Extraction options
   * @param {boolean} options.autoConvertText - Convert XML files to strings (default: true)
   * @param {boolean} options.validateFiles - Validate extracted files (default: true)
   * @returns {Object} Object mapping filenames to content (string for XML, Uint8Array for binary)
   * @throws {Error} With specific error codes for different failure types
   * 
   * AI_USAGE: Use this to extract any OOXML file. Returns mixed content types.
   * Example: const files = FFlatePPTXService.extractFiles(pptxBlob)
   */
  static extractFiles(ooxmlBlob, options = {}) {
    const startTime = Date.now();
    const opts = {
      autoConvertText: true,
      validateFiles: true,
      ...options
    };
    
    try {
      // Input validation
      this._validateExtractInput(ooxmlBlob);
      
      this._log(`🔄 Extracting OOXML file: ${ooxmlBlob.getBytes().length} bytes`);
      
      // Convert blob to Uint8Array
      const ooxmlData = new Uint8Array(ooxmlBlob.getBytes());
      
      // Extract with fflate
      const extractedFiles = this._extractWithFflate(ooxmlData);
      
      // Convert text files if requested
      const processedFiles = opts.autoConvertText 
        ? this._processExtractedFiles(extractedFiles)
        : extractedFiles;
      
      // Validate results if requested
      if (opts.validateFiles) {
        this._validateExtractedFiles(processedFiles);
      }
      
      const duration = Date.now() - startTime;
      this._log(`✅ Extracted ${Object.keys(processedFiles).length} files in ${duration}ms`);
      
      return processedFiles;
      
    } catch (error) {
      const duration = Date.now() - startTime;
      this._logError(`❌ Extract failed after ${duration}ms: ${error.message}`);
      throw new Error(`FFLATE_EXTRACT_ERROR: ${error.message}`);
    }
  }

  /**
   * Compress files into PPTX/OOXML blob
   * @param {Object} files - Object mapping filenames to content
   * @param {Object} options - Compression options
   * @param {number} options.compressionLevel - 0-9, default: 6
   * @param {boolean} options.autoConvertText - Convert strings to Uint8Array (default: true)
   * @returns {Blob} The compressed OOXML file
   * @throws {Error} With specific error codes for different failure types
   * 
   * AI_USAGE: Use this to create OOXML files from modified content.
   * Example: const pptxBlob = FFlatePPTXService.compressFiles(files)
   */
  static compressFiles(files, options = {}) {
    const startTime = Date.now();
    const opts = {
      compressionLevel: 6,
      autoConvertText: true,
      ...options
    };
    
    try {
      // Input validation
      this._validateCompressInput(files);
      
      this._log(`🔄 Compressing ${Object.keys(files).length} files`);
      
      // Prepare files for compression
      const preparedFiles = opts.autoConvertText 
        ? this._prepareFilesForCompression(files)
        : files;
      
      // Compress with fflate
      const compressedData = this._compressWithFflate(preparedFiles, opts.compressionLevel);
      
      // Create blob
      const resultBlob = Utilities.newBlob(compressedData, 'application/vnd.openxmlformats-officedocument.presentationml.presentation', 'presentation.pptx');
      
      const duration = Date.now() - startTime;
      this._log(`✅ Compressed to ${compressedData.length} bytes in ${duration}ms`);
      
      return resultBlob;
      
    } catch (error) {
      const duration = Date.now() - startTime;
      this._logError(`❌ Compress failed after ${duration}ms: ${error.message}`);
      throw new Error(`FFLATE_COMPRESS_ERROR: ${error.message}`);
    }
  }

  /**
   * Convert string to Uint8Array
   * @param {string} str - String to convert
   * @returns {Uint8Array} Encoded bytes
   */
  static stringToUint8Array(str) {
    const encoder = new TextEncoder();
    return encoder.encode(str);
  }

  /**
   * Convert Uint8Array to string
   * @param {Uint8Array} uint8Array - Bytes to convert
   * @returns {string} Decoded string
   */
  static uint8ArrayToString(uint8Array) {
    const decoder = new TextDecoder(this._CONFIG.encoding);
    return decoder.decode(uint8Array);
  }

  /**
   * Check if filename should be treated as text
   * @param {string} filename - File path to check
   * @returns {boolean} True if file should be text
   * @private
   */
  static _isTextFile(filename) {
    const config = this._CONFIG;
    
    // Check exact filename matches
    if (config.textFiles.has(filename)) {
      return true;
    }
    
    // Check extension matches
    const ext = filename.substring(filename.lastIndexOf('.'));
    return config.textExtensions.has(ext);
  }

  /**
   * Extract files using fflate
   * @param {Uint8Array} ooxmlData - OOXML file bytes
   * @returns {Object} Extracted files as Uint8Arrays
   * @private
   */
  static _extractWithFflate(ooxmlData) {
    try {
      // Note: This assumes fflate is available globally
      // In actual GAS implementation, you'd import/include fflate
      
      // Placeholder for fflate.unzipSync - replace with actual fflate call
      if (typeof fflate === 'undefined') {
        throw new Error('fflate library not available');
      }
      
      return fflate.unzipSync(ooxmlData);
      
    } catch (error) {
      throw new Error(`Fflate extraction failed: ${error.message}`);
    }
  }

  /**
   * Process extracted files (convert text files to strings)
   * @param {Object} extractedFiles - Raw extracted files
   * @returns {Object} Processed files with text conversion
   * @private
   */
  static _processExtractedFiles(extractedFiles) {
    const processedFiles = {};
    
    for (const [filename, content] of Object.entries(extractedFiles)) {
      if (this._isTextFile(filename)) {
        // Convert Uint8Array to string for text files
        processedFiles[filename] = this.uint8ArrayToString(content);
      } else {
        // Keep as Uint8Array for binary files
        processedFiles[filename] = content;
      }
    }
    
    return processedFiles;
  }

  /**
   * Prepare files for compression (convert strings to Uint8Array)
   * @param {Object} files - Files to prepare
   * @returns {Object} Files prepared for compression
   * @private
   */
  static _prepareFilesForCompression(files) {
    const preparedFiles = {};
    
    for (const [filename, content] of Object.entries(files)) {
      if (typeof content === 'string') {
        // Convert string to Uint8Array
        preparedFiles[filename] = this.stringToUint8Array(content);
      } else if (content instanceof Uint8Array) {
        // Already Uint8Array
        preparedFiles[filename] = content;
      } else {
        throw new Error(`Invalid content type for ${filename}: ${typeof content}`);
      }
    }
    
    return preparedFiles;
  }

  /**
   * Compress files using fflate
   * @param {Object} files - Files to compress (all Uint8Arrays)
   * @param {number} compressionLevel - Compression level
   * @returns {Uint8Array} Compressed data
   * @private
   */
  static _compressWithFflate(files, compressionLevel) {
    try {
      // Note: This assumes fflate is available globally
      if (typeof fflate === 'undefined') {
        throw new Error('fflate library not available');
      }
      
      return fflate.zipSync(files, { level: compressionLevel });
      
    } catch (error) {
      throw new Error(`Fflate compression failed: ${error.message}`);
    }
  }

  /**
   * Validate extraction input
   * @param {Blob} ooxmlBlob - Blob to validate
   * @private
   */
  static _validateExtractInput(ooxmlBlob) {
    if (!ooxmlBlob) {
      throw new Error('OOXML blob is required');
    }
    
    if (typeof ooxmlBlob.getBytes !== 'function') {
      throw new Error('Invalid blob object');
    }
    
    const sizeBytes = ooxmlBlob.getBytes().length;
    const sizeMB = sizeBytes / (1024 * 1024);
    
    if (sizeMB > this._CONFIG.maxFileSizeMB) {
      throw new Error(`File too large: ${sizeMB.toFixed(1)}MB (max: ${this._CONFIG.maxFileSizeMB}MB)`);
    }
    
    if (sizeBytes === 0) {
      throw new Error('Empty OOXML file');
    }
  }

  /**
   * Validate compression input
   * @param {Object} files - Files to validate
   * @private
   */
  static _validateCompressInput(files) {
    if (!files || typeof files !== 'object') {
      throw new Error('Files object is required');
    }
    
    const fileCount = Object.keys(files).length;
    if (fileCount === 0) {
      throw new Error('No files to compress');
    }
    
    if (fileCount > this._CONFIG.maxFileCount) {
      throw new Error(`Too many files: ${fileCount} (max: ${this._CONFIG.maxFileCount})`);
    }
    
    // Check for required OOXML files
    if (!files['[Content_Types].xml']) {
      throw new Error('Missing required file: [Content_Types].xml');
    }
  }

  /**
   * Validate extracted files
   * @param {Object} files - Extracted files to validate
   * @private
   */
  static _validateExtractedFiles(files) {
    const fileCount = Object.keys(files).length;
    
    if (fileCount === 0) {
      throw new Error('No files extracted');
    }
    
    // Check for required OOXML structure
    if (!files['[Content_Types].xml']) {
      throw new Error('Invalid OOXML: Missing [Content_Types].xml');
    }
    
    if (!files['_rels/.rels']) {
      throw new Error('Invalid OOXML: Missing _rels/.rels');
    }
  }

  /**
   * Log message with timestamp
   * @param {string} message - Message to log
   * @private
   */
  static _log(message) {
    const timestamp = new Date().toISOString();
    console.log(`[${timestamp}] FFlatePPTXService: ${message}`);
  }

  /**
   * Log error message
   * @param {string} message - Error message to log
   * @private
   */
  static _logError(message) {
    const timestamp = new Date().toISOString();
    console.error(`[${timestamp}] FFlatePPTXService: ${message}`);
  }

  /**
   * Get service information
   * @returns {Object} Service info and configuration
   */
  static getServiceInfo() {
    return {
      service: 'FFlatePPTXService',
      version: '1.0.0',
      type: 'client-side',
      library: 'fflate',
      config: this._CONFIG,
      capabilities: [
        'Fast ZIP extraction',
        'Synchronous processing',
        'Automatic text/binary handling',
        'Memory efficient streaming',
        'No network dependencies'
      ]
    };
  }
}

// Legacy compatibility aliases
FFlatePPTXService.unzipPPTX = FFlatePPTXService.extractFiles;
FFlatePPTXService.zipPPTX = FFlatePPTXService.compressFiles;