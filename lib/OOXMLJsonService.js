/**
 * OOXMLJsonService - One-Click GAS + Cloud Run OOXML JSON Helper
 * 
 * PURPOSE:
 * Provides a complete OOXML manipulation platform with one-click Cloud Run deployment,
 * JSON manifests for XML editing, server-side operations, and large file support.
 * 
 * ARCHITECTURE:
 * - GAS-based Cloud Run deployment with preflight checks
 * - JSON manifest system (XML inline, binaries referenced)
 * - Server-side operations: replaceText, upsertPart, removePart, renamePart
 * - Session-based large file handling with GCS signed URLs
 * - Integrated billing and budget management
 * 
 * AI CONTEXT:
 * This is the next-generation OOXML service that replaces CloudPPTXService.
 * Use this for all OOXML operations. It provides JSON manifests for easy XML editing,
 * server-side operations for complex transformations, and handles files up to 100MB+.
 * 
 * TESTING STRATEGY:
 * Test deployment flow, manifest generation, server-side operations, and large file sessions.
 * Verify billing checks, budget setup, and error handling.
 */

class OOXMLJsonService {
  
  /**
   * Configuration for GCP deployment and operations
   * @private
   */
  static get _CONFIG() {
    return {
      PROJECT_ID: PropertiesService.getScriptProperties().getProperty('GCP_PROJECT_ID') || '',
      REGION: 'europe-west4',
      SERVICE: 'ooxml-json',
      BUCKET: null, // null => auto: "<PROJECT_ID>-ooxml-src"
      PUBLIC: true, // false => protected; GAS fetches ID token
      RUN_SA: null, // optional service account email for Cloud Run
      
      // Budget defaults
      BUDGET_NAME: 'OOXML Helper Budget',
      BUDGET_CURRENCY: 'EUR',
      BUDGET_AMOUNT_UNITS: '5',
      BUDGET_THRESHOLDS: [0.10, 0.50, 0.90]
    };
  }
  
  /**
   * Required APIs for Cloud Run deployment
   * @private
   */
  static get _REQUIRED_APIS() {
    return [
      'run.googleapis.com',
      'cloudbuild.googleapis.com',
      'artifactregistry.googleapis.com',
      'iam.googleapis.com',
      'serviceusage.googleapis.com',
      'cloudbilling.googleapis.com',
      'cloudresourcemanager.googleapis.com',
      'billingbudgets.googleapis.com',
      'storage.googleapis.com'
    ];
  }
  
  /**
   * Error codes for systematic error handling
   * @private
   */
  static get _ERROR_CODES() {
    return {
      OOXML_JSON_001: 'Cloud Run service not deployed',
      OOXML_JSON_002: 'Invalid manifest format',
      OOXML_JSON_003: 'File size exceeds session limit',
      OOXML_JSON_004: 'Server-side operation failed',
      OOXML_JSON_005: 'Session creation failed',
      OOXML_JSON_006: 'GCP billing not enabled',
      OOXML_JSON_007: 'Required APIs not enabled'
    };
  }
  
  /**
   * Get service base URL from Script Properties
   * @returns {string} Cloud Run service base URL
   * @throws {Error} If CF_BASE not configured
   * 
   * AI_USAGE: Use this to get the deployed service URL.
   * Example: const baseUrl = OOXMLJsonService.getServiceUrl()
   */
  static getServiceUrl() {
    const url = PropertiesService.getScriptProperties().getProperty('CF_BASE');
    if (!url) {
      throw new Error(`${this._ERROR_CODES.OOXML_JSON_001}: Run initAndDeploy() first`);
    }
    return url;
  }
  
  /**
   * Unwrap OOXML file into JSON manifest
   * 
   * @param {string|Blob} fileIdOrBlob - Drive file ID, filename, or blob
   * @param {Object} options - Options for unwrapping
   * @param {boolean} options.useSession - Use session for large files
   * @returns {Promise<Object>} JSON manifest with entries array
   * @throws {Error} With specific error codes for different failure types
   * 
   * AI_USAGE: Use this to convert OOXML to editable JSON manifest.
   * Example: const manifest = await OOXMLJsonService.unwrap('presentation.pptx')
   */
  static async unwrap(fileIdOrBlob, options = {}) {
    const startTime = Date.now();
    const opts = { useSession: false, ...options };
    
    try {
      this._log(`🔄 Unwrapping OOXML to JSON manifest`);
      
      // Handle different input types
      let blob;
      if (typeof fileIdOrBlob === 'string') {
        blob = this._getFileByIdOrName(fileIdOrBlob).getBlob();
      } else {
        blob = fileIdOrBlob;
      }
      
      // Check if we should use session for large files
      const sizeBytes = blob.getBytes().length;
      const sizeMB = sizeBytes / (1024 * 1024);
      
      if (sizeMB > 25 || opts.useSession) {
        this._log(`📁 Using session flow for ${sizeMB.toFixed(1)}MB file`);
        return this._unwrapViaSession(blob);
      }
      
      // Small file - direct unwrap
      const manifest = await this._unwrapDirect(blob);
      
      const duration = Date.now() - startTime;
      this._log(`✅ Unwrapped to ${manifest.entries.length} entries in ${duration}ms`);
      
      return manifest;
      
    } catch (error) {
      const duration = Date.now() - startTime;
      this._log(`❌ Unwrap failed: ${error.message}`);
      
      const wrappedError = new Error(`${this._ERROR_CODES.OOXML_JSON_002}: ${error.message}`);
      wrappedError.originalError = error;
      wrappedError.duration = duration;
      throw wrappedError;
    }
  }
  
  /**
   * Rewrap JSON manifest back to OOXML file
   * 
   * @param {Object} manifest - JSON manifest with entries
   * @param {Object} options - Options for rewrapping
   * @param {string} options.filename - Output filename
   * @param {string} options.gcsIn - GCS path for original file (session mode)
   * @param {boolean} options.saveToGCS - Save result to GCS instead of Drive
   * @returns {Promise<string|Object>} Drive file ID or GCS result
   * @throws {Error} With specific error codes for different failure types
   * 
   * AI_USAGE: Use this to convert edited JSON manifest back to OOXML.
   * Example: const fileId = await OOXMLJsonService.rewrap(manifest, {filename: 'updated.pptx'})
   */
  static async rewrap(manifest, options = {}) {
    const startTime = Date.now();
    const opts = {
      filename: 'output.pptx',
      saveToGCS: false,
      ...options
    };
    
    try {
      this._log(`🔄 Rewrapping JSON manifest to OOXML`);
      
      // Validate manifest
      this._validateManifest(manifest);
      
      // Use session mode if gcsIn provided
      if (opts.gcsIn) {
        this._log(`📁 Using session rewrap with GCS`);
        return this._rewrapViaSession(manifest, opts);
      }
      
      // Direct rewrap
      const result = await this._rewrapDirect(manifest, opts);
      
      const duration = Date.now() - startTime;
      this._log(`✅ Rewrapped manifest in ${duration}ms`);
      
      return result;
      
    } catch (error) {
      const duration = Date.now() - startTime;
      this._log(`❌ Rewrap failed: ${error.message}`);
      
      const wrappedError = new Error(`${this._ERROR_CODES.OOXML_JSON_002}: ${error.message}`);
      wrappedError.originalError = error;
      wrappedError.duration = duration;
      throw wrappedError;
    }
  }
  
  /**
   * Process OOXML file with server-side operations
   * 
   * @param {string|Blob} fileIdOrBlob - Input file
   * @param {Array} operations - Array of operations to perform
   * @param {Object} options - Processing options
   * @param {string} options.filename - Output filename
   * @param {boolean} options.saveToGCS - Save result to GCS
   * @returns {Promise<string|Object>} Result with file ID or GCS path and report
   * @throws {Error} With specific error codes for different failure types
   * 
   * AI_USAGE: Use this for complex server-side OOXML transformations.
   * Example: const result = await OOXMLJsonService.process(fileId, [
   *   {type: 'replaceText', find: 'ACME', replace: 'DeltaQuad'},
   *   {type: 'upsertPart', path: 'ppt/customXml/item1.xml', text: '<metadata/>'}
   * ])
   */
  static async process(fileIdOrBlob, operations, options = {}) {
    const startTime = Date.now();
    const opts = {
      filename: 'processed.pptx',
      saveToGCS: false,
      ...options
    };
    
    try {
      this._log(`🔄 Processing OOXML with ${operations.length} operations`);
      
      // Handle different input types
      let blob;
      if (typeof fileIdOrBlob === 'string') {
        blob = this._getFileByIdOrName(fileIdOrBlob).getBlob();
      } else {
        blob = fileIdOrBlob;
      }
      
      // Validate operations
      this._validateOperations(operations);
      
      // Execute server-side processing
      const result = await this._processServerSide(blob, operations, opts);
      
      const duration = Date.now() - startTime;
      this._log(`✅ Processed with ${result.report?.replaced || 0} replacements in ${duration}ms`);
      
      return result;
      
    } catch (error) {
      const duration = Date.now() - startTime;
      this._log(`❌ Processing failed: ${error.message}`);
      
      const wrappedError = new Error(`${this._ERROR_CODES.OOXML_JSON_004}: ${error.message}`);
      wrappedError.originalError = error;
      wrappedError.duration = duration;
      throw wrappedError;
    }
  }
  
  /**
   * Create new session for large file operations
   * 
   * @returns {Promise<Object>} Session object with signed URLs and GCS paths
   * @throws {Error} If session creation fails
   * 
   * AI_USAGE: Use this for large file workflows.
   * Example: const session = await OOXMLJsonService.createSession()
   */
  static async createSession() {
    try {
      this._log('🔄 Creating new session for large file operations');
      
      const baseUrl = this.getServiceUrl();
      const headers = this._getAuthHeaders(baseUrl);
      
      const response = await UrlFetchApp.fetch(`${baseUrl}/session/new`, {
        method: 'POST',
        headers: { ...headers, 'Content-Type': 'application/json' },
        payload: JSON.stringify({}),
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() >= 300) {
        throw new Error(response.getContentText());
      }
      
      const session = JSON.parse(response.getContentText());
      this._log(`✅ Created session: ${session.sessionId}`);
      
      return session;
      
    } catch (error) {
      this._log(`❌ Session creation failed: ${error.message}`);
      throw new Error(`${this._ERROR_CODES.OOXML_JSON_005}: ${error.message}`);
    }
  }
  
  /**
   * Upload file to session signed URL
   * 
   * @param {string} uploadUrl - Signed upload URL from session
   * @param {string|Blob} fileIdOrBlob - File to upload
   * @returns {Promise<void>}
   * 
   * AI_USAGE: Use this to upload large files to session.
   * Example: await OOXMLJsonService.uploadToSession(session.uploadUrl, 'large.pptx')
   */
  static async uploadToSession(uploadUrl, fileIdOrBlob) {
    try {
      let blob;
      if (typeof fileIdOrBlob === 'string') {
        blob = this._getFileByIdOrName(fileIdOrBlob).getBlob();
      } else {
        blob = fileIdOrBlob;
      }
      
      this._log(`📤 Uploading ${blob.getBytes().length} bytes to session`);
      
      const response = await UrlFetchApp.fetch(uploadUrl, {
        method: 'PUT',
        contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        payload: blob.getBytes(),
        followRedirects: true,
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() >= 300) {
        throw new Error(`${response.getResponseCode()}: ${response.getContentText()}`);
      }
      
      this._log('✅ Upload to session completed');
      
    } catch (error) {
      this._log(`❌ Session upload failed: ${error.message}`);
      throw error;
    }
  }
  
  /**
   * Check service health and availability
   * 
   * @returns {Promise<Object>} Health status
   * 
   * AI_USAGE: Use this to verify service is running.
   * Example: const health = await OOXMLJsonService.healthCheck()
   */
  static async healthCheck() {
    try {
      const baseUrl = this.getServiceUrl();
      
      const response = await UrlFetchApp.fetch(`${baseUrl}/ping`, {
        method: 'GET',
        muteHttpExceptions: true
      });
      
      const isHealthy = response.getResponseCode() < 400;
      const result = JSON.parse(response.getContentText());
      
      return {
        available: isHealthy,
        statusCode: response.getResponseCode(),
        version: result.version,
        timestamp: new Date().toISOString()
      };
      
    } catch (error) {
      return {
        available: false,
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }
  }
  
  /**
   * Get service configuration and info
   * 
   * @returns {Object} Service information
   */
  static getServiceInfo() {
    return {
      version: '1.0.0',
      config: this._CONFIG,
      capabilities: {
        unwrap: true,
        rewrap: true,
        process: true,
        sessions: true,
        largeFiles: true,
        serverSideOps: true
      },
      supportedFormats: ['pptx', 'docx', 'xlsx', 'thmx'],
      operations: [
        'replaceText',
        'upsertPart', 
        'removePart',
        'renamePart'
      ]
    };
  }
  
  // PRIVATE METHODS - Internal implementation
  
  /**
   * Direct unwrap for small files
   * @private
   */
  static async _unwrapDirect(blob) {
    const baseUrl = this.getServiceUrl();
    const headers = this._getAuthHeaders(baseUrl);
    const zipB64 = Utilities.base64Encode(blob.getBytes());
    
    const response = await UrlFetchApp.fetch(`${baseUrl}/unwrap`, {
      method: 'POST',
      headers: { ...headers, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ zipB64 }),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() >= 300) {
      throw new Error(response.getContentText());
    }
    
    return JSON.parse(response.getContentText());
  }
  
  /**
   * Session-based unwrap for large files
   * @private
   */
  static async _unwrapViaSession(blob) {
    const session = await this.createSession();
    await this.uploadToSession(session.uploadUrl, blob);
    
    const baseUrl = this.getServiceUrl();
    const headers = this._getAuthHeaders(baseUrl);
    
    const response = await UrlFetchApp.fetch(`${baseUrl}/unwrap`, {
      method: 'POST',
      headers: { ...headers, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ gcsIn: session.gcsIn }),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() >= 300) {
      throw new Error(response.getContentText());
    }
    
    const manifest = JSON.parse(response.getContentText());
    manifest._session = session; // Attach session for rewrap
    return manifest;
  }
  
  /**
   * Direct rewrap for small files
   * @private
   */
  static async _rewrapDirect(manifest, options) {
    const baseUrl = this.getServiceUrl();
    const headers = this._getAuthHeaders(baseUrl);
    
    const response = await UrlFetchApp.fetch(`${baseUrl}/rewrap`, {
      method: 'POST',
      headers: { ...headers, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ manifest }),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() >= 300) {
      throw new Error(response.getContentText());
    }
    
    const result = JSON.parse(response.getContentText());
    
    if (options.saveToGCS) {
      return result;
    }
    
    // Create Drive file
    const bytes = Utilities.base64Decode(result.zipB64);
    const mime = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
    const file = DriveApp.createFile(Utilities.newBlob(bytes, mime, options.filename));
    
    return file.getId();
  }
  
  /**
   * Session-based rewrap for large files
   * @private
   */
  static async _rewrapViaSession(manifest, options) {
    const session = manifest._session || options;
    const baseUrl = this.getServiceUrl();
    const headers = this._getAuthHeaders(baseUrl);
    
    const payload = {
      manifest,
      gcsIn: session.gcsIn,
      gcsOut: session.gcsOut
    };
    
    const response = await UrlFetchApp.fetch(`${baseUrl}/rewrap`, {
      method: 'POST',
      headers: { ...headers, 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() >= 300) {
      throw new Error(response.getContentText());
    }
    
    return JSON.parse(response.getContentText());
  }
  
  /**
   * Server-side processing
   * @private
   */
  static async _processServerSide(blob, operations, options) {
    const baseUrl = this.getServiceUrl();
    const headers = this._getAuthHeaders(baseUrl);
    const zipB64 = Utilities.base64Encode(blob.getBytes());
    
    const response = await UrlFetchApp.fetch(`${baseUrl}/process`, {
      method: 'POST',
      headers: { ...headers, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ 
        zipB64,
        ops: operations 
      }),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() >= 300) {
      throw new Error(response.getContentText());
    }
    
    const result = JSON.parse(response.getContentText());
    
    if (options.saveToGCS) {
      return result;
    }
    
    // Create Drive file if zipB64 returned
    if (result.zipB64) {
      const bytes = Utilities.base64Decode(result.zipB64);
      const mime = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
      const file = DriveApp.createFile(Utilities.newBlob(bytes, mime, options.filename));
      
      return {
        fileId: file.getId(),
        report: result.report
      };
    }
    
    return result;
  }
  
  /**
   * Get file by ID or name
   * @private
   */
  static _getFileByIdOrName(idOrName) {
    if (/^[a-zA-Z0-9\-_]{20,}$/.test(idOrName)) {
      return DriveApp.getFileById(idOrName);
    }
    
    const files = DriveApp.getFilesByName(idOrName);
    if (!files.hasNext()) {
      throw new Error(`File not found: ${idOrName}`);
    }
    
    return files.next();
  }
  
  /**
   * Get authentication headers
   * @private
   */
  static _getAuthHeaders(baseUrl) {
    if (this._CONFIG.PUBLIC) {
      return {};
    }
    
    // For private services, would need ID token
    throw new Error('Private service authentication not implemented. Set PUBLIC: true');
  }
  
  /**
   * Validate manifest structure
   * @private
   */
  static _validateManifest(manifest) {
    if (!manifest || typeof manifest !== 'object') {
      throw new Error('Invalid manifest: must be object');
    }
    
    if (!Array.isArray(manifest.entries)) {
      throw new Error('Invalid manifest: entries must be array');
    }
    
    for (const entry of manifest.entries) {
      if (!entry.path || !entry.type) {
        throw new Error('Invalid manifest entry: path and type required');
      }
      
      if (entry.type === 'xml' && typeof entry.text !== 'string') {
        throw new Error(`Invalid XML entry ${entry.path}: text must be string`);
      }
    }
  }
  
  /**
   * Validate operations array
   * @private
   */
  static _validateOperations(operations) {
    if (!Array.isArray(operations)) {
      throw new Error('Operations must be array');
    }
    
    const validTypes = ['replaceText', 'upsertPart', 'removePart', 'renamePart'];
    
    for (const op of operations) {
      if (!op.type || !validTypes.includes(op.type)) {
        throw new Error(`Invalid operation type: ${op.type}. Must be one of: ${validTypes.join(', ')}`);
      }
    }
  }
  
  /**
   * Log message with timestamp
   * @private
   */
  static _log(message) {
    const timestamp = new Date().toISOString().substr(11, 8);
    console.log(`[${timestamp}] [OOXMLJsonService] ${message}`);
  }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = OOXMLJsonService;
}