/**
 * OOXMLExtensionAdapter - Bridge between Extensions and OOXML JSON Service
 * 
 * PURPOSE:
 * Adapts the existing extension framework to work with the new OOXML JSON Service
 * backend. Provides server-side operations capabilities while maintaining the
 * existing extension API compatibility.
 * 
 * ARCHITECTURE:
 * - Translates extension operations to server-side operations
 * - Provides JSON manifest manipulation helpers
 * - Maintains compatibility with existing BaseExtension pattern
 * - Adds server-side operation batching for performance
 * 
 * AI CONTEXT:
 * This adapter allows existing extensions to leverage the powerful server-side
 * operations of the OOXML JSON Service while maintaining API compatibility.
 * It automatically batches operations for optimal performance.
 */

class OOXMLExtensionAdapter {
  
  /**
   * Initialize adapter with manifest and optional session
   * 
   * @param {Object} manifest - OOXML JSON manifest
   * @param {Object} session - Optional session for large files
   * @param {Object} options - Adapter options
   */
  constructor(manifest, session = null, options = {}) {
    this.manifest = manifest;
    this.session = session;
    this.options = {
      enableBatching: true,
      batchSize: 50,
      enableServerSideOps: true,
      ...options
    };
    
    // Operation batching
    this.pendingOperations = [];
    this.batchTimer = null;
    
    // Performance tracking
    this.metrics = {
      operationsProcessed: 0,
      batchesExecuted: 0,
      serverSideOpsUsed: 0,
      fallbackOpsUsed: 0
    };
    
    this._log('OOXMLExtensionAdapter initialized');
  }
  
  /**
   * Get file content by path (XML or binary)
   * 
   * @param {string} path - File path in OOXML
   * @returns {string|null} File content or null if not found
   * 
   * AI_USAGE: Use this to read OOXML file contents in extensions.
   * Example: const slideXml = adapter.getFile('ppt/slides/slide1.xml')
   */
  getFile(path) {
    const entry = this.manifest.entries.find(e => e.path === path);
    
    if (!entry) {
      return null;
    }
    
    if (entry.type === 'xml') {
      return entry.text;
    } else if (entry.type === 'bin' && entry.dataB64) {
      return entry.dataB64;
    }
    
    return null;
  }
  
  /**
   * Set file content by path
   * 
   * @param {string} path - File path in OOXML
   * @param {string} content - File content (XML text or base64 for binary)
   * @param {string} contentType - Optional content type for new files
   * 
   * AI_USAGE: Use this to modify OOXML file contents in extensions.
   * Example: adapter.setFile('ppt/slides/slide1.xml', updatedXml)
   */
  setFile(path, content, contentType = null) {
    let entry = this.manifest.entries.find(e => e.path === path);
    
    if (entry) {
      // Update existing entry
      if (entry.type === 'xml') {
        entry.text = content;
      } else if (entry.type === 'bin') {
        entry.dataB64 = content;
      }
    } else {
      // Create new entry
      const isXml = /\.xml$/i.test(path);
      
      entry = {
        path,
        type: isXml ? 'xml' : 'bin'
      };
      
      if (isXml) {
        entry.text = content;
      } else {
        entry.dataB64 = content;
      }
      
      this.manifest.entries.push(entry);
      
      // Add to content types if specified
      if (contentType) {
        this._addContentType(path, contentType);
      }
    }
    
    this._trackOperation('setFile', { path, size: content.length });
  }
  
  /**
   * Remove file by path
   * 
   * @param {string} path - File path to remove
   * 
   * AI_USAGE: Use this to remove OOXML files in extensions.
   * Example: adapter.removeFile('ppt/slides/slide10.xml')
   */
  removeFile(path) {
    const index = this.manifest.entries.findIndex(e => e.path === path);
    
    if (index >= 0) {
      this.manifest.entries.splice(index, 1);
      this._removeContentType(path);
      this._trackOperation('removeFile', { path });
      return true;
    }
    
    return false;
  }
  
  /**
   * Replace text across multiple files using server-side operations
   * 
   * @param {string|RegExp} find - Text to find
   * @param {string} replace - Replacement text
   * @param {Object} options - Replace options
   * @param {string} options.scope - File path prefix to limit scope
   * @param {boolean} options.regex - Treat find as regex
   * @param {string} options.flags - Regex flags
   * @returns {Promise<Object>} Replace operation result
   * 
   * AI_USAGE: Use this for efficient text replacement across presentation.
   * Example: await adapter.replaceText('ACME Corp', 'DeltaQuad Inc', {scope: 'ppt/slides/'})
   */
  async replaceText(find, replace, options = {}) {
    if (this.options.enableServerSideOps) {
      // Use server-side operation for efficiency
      const operation = {
        type: 'replaceText',
        find: find.toString(),
        replace: replace,
        scope: options.scope || '',
        regex: find instanceof RegExp || options.regex,
        flags: options.flags || 'g'
      };
      
      return this._addServerSideOperation(operation);
    } else {
      // Fallback to client-side replacement
      return this._replaceTextClientSide(find, replace, options);
    }
  }
  
  /**
   * Upsert (insert or update) a file part
   * 
   * @param {string} path - File path
   * @param {string} content - File content
   * @param {string} contentType - Content type for new files
   * @returns {Promise<Object>} Upsert operation result
   * 
   * AI_USAGE: Use this to add or update OOXML parts.
   * Example: await adapter.upsertPart('ppt/customXml/item1.xml', xmlContent, 'application/xml')
   */
  async upsertPart(path, content, contentType = null) {
    if (this.options.enableServerSideOps) {
      const operation = {
        type: 'upsertPart',
        path,
        text: /\.xml$/i.test(path) ? content : undefined,
        dataB64: !/\.xml$/i.test(path) ? content : undefined,
        contentType
      };
      
      return this._addServerSideOperation(operation);
    } else {
      // Fallback to direct manifest manipulation
      this.setFile(path, content, contentType);
      return { success: true, operation: 'upsertPart', path };
    }
  }
  
  /**
   * Remove a file part
   * 
   * @param {string} path - File path to remove
   * @returns {Promise<Object>} Remove operation result
   * 
   * AI_USAGE: Use this to remove OOXML parts.
   * Example: await adapter.removePart('ppt/slides/slide10.xml')
   */
  async removePart(path) {
    if (this.options.enableServerSideOps) {
      const operation = {
        type: 'removePart',
        path
      };
      
      return this._addServerSideOperation(operation);
    } else {
      // Fallback to direct manifest manipulation
      const removed = this.removeFile(path);
      return { success: removed, operation: 'removePart', path };
    }
  }
  
  /**
   * Rename a file part
   * 
   * @param {string} fromPath - Current file path
   * @param {string} toPath - New file path
   * @param {string} contentType - Optional content type for renamed file
   * @returns {Promise<Object>} Rename operation result
   * 
   * AI_USAGE: Use this to rename OOXML parts.
   * Example: await adapter.renamePart('ppt/slides/slide1.xml', 'ppt/slides/intro.xml')
   */
  async renamePart(fromPath, toPath, contentType = null) {
    if (this.options.enableServerSideOps) {
      const operation = {
        type: 'renamePart',
        from: fromPath,
        to: toPath,
        contentType
      };
      
      return this._addServerSideOperation(operation);
    } else {
      // Fallback to direct manifest manipulation
      const content = this.getFile(fromPath);
      if (content) {
        this.setFile(toPath, content, contentType);
        this.removeFile(fromPath);
        return { success: true, operation: 'renamePart', from: fromPath, to: toPath };
      }
      return { success: false, operation: 'renamePart', error: 'Source file not found' };
    }
  }
  
  /**
   * Execute pending server-side operations
   * 
   * @returns {Promise<Object>} Execution result with report
   * 
   * AI_USAGE: Use this to flush pending operations (usually automatic).
   * Example: const result = await adapter.executeOperations()
   */
  async executeOperations() {
    if (this.pendingOperations.length === 0) {
      return { success: true, operationsExecuted: 0 };
    }
    
    try {
      this._log(`Executing ${this.pendingOperations.length} server-side operations`);
      
      // Execute operations via OOXML JSON Service
      const result = await OOXMLJsonService.process(
        this.manifest,
        this.pendingOperations,
        {
          filename: 'temp.pptx',
          saveToGCS: !!this.session
        }
      );
      
      // Update manifest with processed result
      if (result.manifest) {
        this.manifest = result.manifest;
      }
      
      // Clear pending operations
      const executedCount = this.pendingOperations.length;
      this.pendingOperations = [];
      
      // Update metrics
      this.metrics.batchesExecuted++;
      this.metrics.operationsProcessed += executedCount;
      this.metrics.serverSideOpsUsed += executedCount;
      
      this._log(`✅ Executed ${executedCount} operations successfully`);
      
      return {
        success: true,
        operationsExecuted: executedCount,
        report: result.report || {},
        metrics: this.metrics
      };
      
    } catch (error) {
      this._log(`❌ Failed to execute operations: ${error.message}`);
      
      // Mark operations as failed but don't clear them
      // They can be retried or executed individually
      throw error;
    }
  }
  
  /**
   * Get list of files by pattern
   * 
   * @param {string|RegExp} pattern - File path pattern
   * @returns {Array} Matching file entries
   * 
   * AI_USAGE: Use this to find files matching a pattern.
   * Example: const slides = adapter.getFilesByPattern(/ppt\/slides\/slide\d+\.xml/)
   */
  getFilesByPattern(pattern) {
    const regex = pattern instanceof RegExp ? pattern : new RegExp(pattern);
    return this.manifest.entries.filter(entry => regex.test(entry.path));
  }
  
  /**
   * Get adapter metrics and performance data
   * 
   * @returns {Object} Performance metrics
   */
  getMetrics() {
    return {
      ...this.metrics,
      pendingOperations: this.pendingOperations.length,
      manifestEntries: this.manifest.entries.length,
      manifestKind: this.manifest.kind,
      sessionActive: !!this.session
    };
  }
  
  // PRIVATE METHODS
  
  /**
   * Add server-side operation to batch
   * @private
   */
  async _addServerSideOperation(operation) {
    this.pendingOperations.push(operation);
    
    // Auto-execute if batch size reached
    if (this.options.enableBatching && 
        this.pendingOperations.length >= this.options.batchSize) {
      return this.executeOperations();
    }
    
    // Set timer for delayed execution
    if (this.options.enableBatching && !this.batchTimer) {
      this.batchTimer = setTimeout(() => {
        this.batchTimer = null;
        this.executeOperations().catch(error => {
          this._log(`Batch execution failed: ${error.message}`);
        });
      }, 100); // 100ms delay
    }
    
    return { success: true, operation: operation.type, queued: true };
  }
  
  /**
   * Client-side text replacement fallback
   * @private
   */
  _replaceTextClientSide(find, replace, options) {
    const regex = find instanceof RegExp ? find : 
      new RegExp(find.toString().replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), options.flags || 'g');
    
    let replacementCount = 0;
    
    for (const entry of this.manifest.entries) {
      if (entry.type === 'xml' && entry.text) {
        // Apply scope filter if specified
        if (options.scope && !entry.path.startsWith(options.scope)) {
          continue;
        }
        
        const originalText = entry.text;
        entry.text = originalText.replace(regex, replace);
        
        if (entry.text !== originalText) {
          replacementCount++;
        }
      }
    }
    
    this.metrics.fallbackOpsUsed++;
    
    return {
      success: true,
      operation: 'replaceText',
      replacements: replacementCount,
      method: 'client-side'
    };
  }
  
  /**
   * Add content type override
   * @private
   */
  _addContentType(partName, contentType) {
    const contentTypesEntry = this.manifest.entries.find(
      e => e.path === '[Content_Types].xml' && e.type === 'xml'
    );
    
    if (contentTypesEntry && contentTypesEntry.text) {
      const pn = partName.startsWith('/') ? partName : '/' + partName;
      const overrideXml = `<Override PartName="${pn}" ContentType="${contentType}"/>`;
      
      // Insert before closing </Types>
      contentTypesEntry.text = contentTypesEntry.text.replace(
        /<\/Types>\s*$/i,
        `${overrideXml}</Types>`
      );
    }
  }
  
  /**
   * Remove content type override
   * @private
   */
  _removeContentType(partName) {
    const contentTypesEntry = this.manifest.entries.find(
      e => e.path === '[Content_Types].xml' && e.type === 'xml'
    );
    
    if (contentTypesEntry && contentTypesEntry.text) {
      const pn = partName.startsWith('/') ? partName : '/' + partName;
      const escapedPath = pn.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const regex = new RegExp(
        `<Override\\b[^>]*PartName="${escapedPath}"[^>]*/>\\s*`,
        'i'
      );
      
      contentTypesEntry.text = contentTypesEntry.text.replace(regex, '');
    }
  }
  
  /**
   * Track operation for metrics
   * @private
   */
  _trackOperation(operation, details) {
    this.metrics.operationsProcessed++;
    this._log(`Operation: ${operation} - ${JSON.stringify(details)}`);
  }
  
  /**
   * Log message with timestamp
   * @private
   */
  _log(message) {
    const timestamp = new Date().toISOString().substr(11, 8);
    console.log(`[${timestamp}] [OOXMLExtensionAdapter] ${message}`);
  }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = OOXMLExtensionAdapter;
}