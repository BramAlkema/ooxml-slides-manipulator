/**
 * CorrelationTracker - Request Correlation ID Management
 * 
 * PURPOSE:
 * Provides comprehensive correlation ID tracking throughout the OOXML JSON
 * system. Enables request tracing across service boundaries, detailed
 * observability, and systematic debugging in production environments.
 * 
 * ARCHITECTURE:
 * - Automatic correlation ID generation and propagation
 * - Thread-local storage simulation for Google Apps Script
 * - Integration with error handling and logging systems
 * - Cross-service correlation tracking
 * 
 * AI CONTEXT:
 * Use this to track requests end-to-end across the entire system. Every
 * operation should have a correlation ID that can be traced through logs,
 * errors, and performance metrics for production debugging.
 */

/**
 * Correlation context for tracking request flow
 */
class CorrelationContext {
  constructor(correlationId, metadata = {}) {
    this.correlationId = correlationId;
    this.startTime = Date.now();
    this.metadata = {
      userAgent: this._getUserAgent(),
      sessionId: this._getSessionId(),
      timestamp: new Date().toISOString(),
      ...metadata
    };
    this.operations = [];
    this.errors = [];
    this.metrics = {};
  }
  
  /**
   * Add operation to context
   */
  addOperation(operation, duration = null, success = true) {
    this.operations.push({
      operation,
      duration,
      success,
      timestamp: new Date().toISOString()
    });
  }
  
  /**
   * Add error to context
   */
  addError(error, operation = null) {
    this.errors.push({
      error: error instanceof Error ? error.message : String(error),
      code: error.code || 'UNKNOWN',
      operation,
      timestamp: new Date().toISOString()
    });
  }
  
  /**
   * Set metric value
   */
  setMetric(key, value) {
    this.metrics[key] = value;
  }
  
  /**
   * Get total request duration
   */
  getDuration() {
    return Date.now() - this.startTime;
  }
  
  /**
   * Export context for logging
   */
  toJSON() {
    return {
      correlationId: this.correlationId,
      duration: this.getDuration(),
      metadata: this.metadata,
      operations: this.operations,
      errors: this.errors,
      metrics: this.metrics
    };
  }
  
  // PRIVATE METHODS
  
  _getUserAgent() {
    if (typeof navigator !== 'undefined') {
      return navigator.userAgent;
    }
    return 'Google Apps Script';
  }
  
  _getSessionId() {
    try {
      return Session.getActiveUser().getEmail().substring(0, 8);
    } catch (e) {
      return 'anonymous';
    }
  }
}

/**
 * Main correlation tracker with context management
 */
class CorrelationTracker {
  
  constructor() {
    // Simulate thread-local storage for GAS
    this.contexts = new Map();
    this.currentCorrelationId = null;
  }
  
  /**
   * Generate a new correlation ID
   */
  generateCorrelationId() {
    // Generate UUID v4-like identifier
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
      const r = Math.random() * 16 | 0;
      const v = c === 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });
  }
  
  /**
   * Start a new correlation context
   */
  startContext(correlationId = null, metadata = {}) {
    const id = correlationId || this.generateCorrelationId();
    const context = new CorrelationContext(id, metadata);
    
    this.contexts.set(id, context);
    this.currentCorrelationId = id;
    
    this._log('info', 'Correlation context started', { correlationId: id });
    
    return id;
  }
  
  /**
   * End correlation context and cleanup
   */
  endContext(correlationId = null) {
    const id = correlationId || this.currentCorrelationId;
    if (!id) return null;
    
    const context = this.contexts.get(id);
    if (!context) return null;
    
    const summary = context.toJSON();
    this._log('info', 'Correlation context ended', summary);
    
    // Cleanup
    this.contexts.delete(id);
    if (this.currentCorrelationId === id) {
      this.currentCorrelationId = null;
    }
    
    return summary;
  }
  
  /**
   * Get current correlation ID
   */
  getCurrentCorrelationId() {
    return this.currentCorrelationId;
  }
  
  /**
   * Get correlation context
   */
  getContext(correlationId = null) {
    const id = correlationId || this.currentCorrelationId;
    return id ? this.contexts.get(id) : null;
  }
  
  /**
   * Track operation with correlation
   */
  trackOperation(operation, fn, correlationId = null) {
    const id = correlationId || this.currentCorrelationId;
    const context = this.getContext(id);
    
    const startTime = Date.now();
    let success = true;
    let error = null;
    
    try {
      const result = fn();
      
      // Handle async functions
      if (result && typeof result.then === 'function') {
        return result
          .then(asyncResult => {
            const duration = Date.now() - startTime;
            if (context) {
              context.addOperation(operation, duration, true);
            }
            this._logOperation(operation, duration, true, id);
            return asyncResult;
          })
          .catch(asyncError => {
            const duration = Date.now() - startTime;
            if (context) {
              context.addOperation(operation, duration, false);
              context.addError(asyncError, operation);
            }
            this._logOperation(operation, duration, false, id, asyncError);
            throw asyncError;
          });
      }
      
      // Synchronous result
      const duration = Date.now() - startTime;
      if (context) {
        context.addOperation(operation, duration, true);
      }
      this._logOperation(operation, duration, true, id);
      
      return result;
      
    } catch (syncError) {
      success = false;
      error = syncError;
      
      const duration = Date.now() - startTime;
      if (context) {
        context.addOperation(operation, duration, false);
        context.addError(syncError, operation);
      }
      this._logOperation(operation, duration, false, id, syncError);
      
      throw syncError;
    }
  }
  
  /**
   * Add custom metric to current context
   */
  addMetric(key, value, correlationId = null) {
    const context = this.getContext(correlationId);
    if (context) {
      context.setMetric(key, value);
    }
  }
  
  /**
   * Log error with correlation context
   */
  logError(error, operation = null, correlationId = null) {
    const id = correlationId || this.currentCorrelationId;
    const context = this.getContext(id);
    
    if (context) {
      context.addError(error, operation);
    }
    
    this._log('error', error.message || String(error), {
      correlationId: id,
      operation,
      code: error.code,
      stack: error.stack
    });
  }
  
  /**
   * Get correlation headers for HTTP requests
   */
  getHeaders(correlationId = null) {
    const id = correlationId || this.currentCorrelationId;
    return id ? { 'x-correlation-id': id } : {};
  }
  
  /**
   * Extract correlation ID from HTTP headers
   */
  extractFromHeaders(headers) {
    if (!headers) return null;
    
    // Handle different header formats
    return headers['x-correlation-id'] || 
           headers['X-Correlation-Id'] || 
           headers['correlation-id'] ||
           null;
  }
  
  /**
   * Wrap function with automatic correlation tracking
   */
  wrap(operation, fn) {
    return (...args) => {
      return this.trackOperation(operation, () => fn.apply(this, args));
    };
  }
  
  /**
   * Create child correlation for sub-operations
   */
  createChild(parentCorrelationId = null, operation = 'child') {
    const parentId = parentCorrelationId || this.currentCorrelationId;
    const childId = this.generateCorrelationId();
    
    const metadata = {
      parentCorrelationId: parentId,
      operation
    };
    
    this.startContext(childId, metadata);
    return childId;
  }
  
  /**
   * Get all active contexts (for debugging)
   */
  getActiveContexts() {
    return Array.from(this.contexts.entries()).map(([id, context]) => ({
      correlationId: id,
      duration: context.getDuration(),
      operations: context.operations.length,
      errors: context.errors.length
    }));
  }
  
  /**
   * Cleanup stale contexts
   */
  cleanup(maxAge = 300000) { // 5 minutes default
    const cutoff = Date.now() - maxAge;
    const staleIds = [];
    
    for (const [id, context] of this.contexts) {
      if (context.startTime < cutoff) {
        staleIds.push(id);
      }
    }
    
    staleIds.forEach(id => {
      this._log('warn', 'Cleaning up stale correlation context', { correlationId: id });
      this.contexts.delete(id);
    });
    
    return staleIds.length;
  }
  
  // PRIVATE METHODS
  
  /**
   * Log operation metrics
   * @private
   */
  _logOperation(operation, duration, success, correlationId, error = null) {
    const logData = {
      operation,
      duration_ms: duration,
      success,
      correlationId
    };
    
    if (error) {
      logData.error = error.message;
      logData.code = error.code;
    }
    
    this._log('metrics', `Operation: ${operation}`, logData);
  }
  
  /**
   * Internal logging method
   * @private
   */
  _log(level, message, context = {}) {
    const timestamp = new Date().toISOString();
    const logEntry = {
      timestamp,
      level,
      message,
      ...context
    };
    
    // Use console.log with structured format
    console.log(JSON.stringify(logEntry));
  }
}

/**
 * Correlation middleware for service integration
 */
class CorrelationMiddleware {
  
  constructor(tracker) {
    this.tracker = tracker;
  }
  
  /**
   * Express-style middleware for Cloud Run service
   */
  expressMiddleware() {
    return (req, res, next) => {
      // Extract or generate correlation ID
      let correlationId = req.headers['x-correlation-id'];
      if (!correlationId) {
        correlationId = this.tracker.generateCorrelationId();
      }
      
      // Start correlation context
      this.tracker.startContext(correlationId, {
        method: req.method,
        path: req.path,
        userAgent: req.headers['user-agent'],
        ip: req.ip
      });
      
      // Set response header
      res.setHeader('x-correlation-id', correlationId);
      
      // Add to request object
      req.correlationId = correlationId;
      req.tracker = this.tracker;
      
      // End context on response finish
      res.on('finish', () => {
        this.tracker.addMetric('statusCode', res.statusCode);
        this.tracker.endContext(correlationId);
      });
      
      next();
    };
  }
  
  /**
   * Google Apps Script UrlFetchApp wrapper
   */
  wrapUrlFetch(url, options = {}) {
    const correlationId = this.tracker.getCurrentCorrelationId();
    
    // Add correlation header
    const headers = {
      ...options.headers,
      ...this.tracker.getHeaders(correlationId)
    };
    
    const wrappedOptions = {
      ...options,
      headers
    };
    
    return this.tracker.trackOperation('http_request', () => {
      return UrlFetchApp.fetch(url, wrappedOptions);
    });
  }
  
  /**
   * OOXMLJsonService wrapper with correlation
   */
  wrapOOXMLService(service) {
    const wrappedService = {};
    
    ['unwrap', 'rewrap', 'process', 'createSession', 'healthCheck'].forEach(method => {
      if (typeof service[method] === 'function') {
        wrappedService[method] = this.tracker.wrap(`ooxml_${method}`, service[method].bind(service));
      }
    });
    
    return wrappedService;
  }
}

/**
 * Correlation-aware error handler
 */
class CorrelatedErrorHandler {
  
  constructor(tracker) {
    this.tracker = tracker;
  }
  
  /**
   * Handle error with correlation context
   */
  handleError(error, operation = null) {
    const correlationId = this.tracker.getCurrentCorrelationId();
    
    // Enhance error with correlation
    if (error instanceof Error && !error.correlationId) {
      error.correlationId = correlationId;
    }
    
    // Log to correlation context
    this.tracker.logError(error, operation);
    
    // Return enhanced error info
    return {
      error: error.message,
      code: error.code || 'UNKNOWN',
      correlationId,
      timestamp: new Date().toISOString(),
      operation
    };
  }
  
  /**
   * Create correlation-aware error
   */
  createError(code, message, context = {}) {
    const correlationId = this.tracker.getCurrentCorrelationId();
    
    const error = new OOXMLError(code, message, {
      ...context,
      correlationId
    }, correlationId);
    
    this.tracker.logError(error);
    
    return error;
  }
}

// Global tracker instance
const globalTracker = new CorrelationTracker();

/**
 * Convenience functions for easy integration
 */
const Correlation = {
  // Start new correlation
  start: (correlationId, metadata) => globalTracker.startContext(correlationId, metadata),
  
  // End correlation
  end: (correlationId) => globalTracker.endContext(correlationId),
  
  // Get current ID
  getId: () => globalTracker.getCurrentCorrelationId(),
  
  // Track operation
  track: (operation, fn) => globalTracker.trackOperation(operation, fn),
  
  // Wrap function
  wrap: (operation, fn) => globalTracker.wrap(operation, fn),
  
  // Add metric
  metric: (key, value) => globalTracker.addMetric(key, value),
  
  // Log error
  error: (error, operation) => globalTracker.logError(error, operation),
  
  // Get headers
  headers: () => globalTracker.getHeaders(),
  
  // Create child
  child: (operation) => globalTracker.createChild(null, operation),
  
  // Cleanup
  cleanup: () => globalTracker.cleanup(),
  
  // Get tracker instance
  getTracker: () => globalTracker
};

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    CorrelationTracker,
    CorrelationContext,
    CorrelationMiddleware,
    CorrelatedErrorHandler,
    Correlation
  };
}