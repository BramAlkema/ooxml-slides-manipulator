/**
 * StructuredObservability - Production Logging and Monitoring
 * 
 * PURPOSE:
 * Provides comprehensive structured logging, metrics collection, and observability
 * for the OOXML JSON system. Enables production monitoring, debugging, and
 * performance analysis with greppable logs and actionable metrics.
 * 
 * ARCHITECTURE:
 * - Structured JSON logging with correlation tracking
 * - Performance metrics collection and aggregation
 * - Error tracking with context and categorization
 * - Health monitoring and alerting
 * - Integration with external monitoring systems
 * 
 * AI CONTEXT:
 * Use this throughout the system for production observability. All logs are
 * structured JSON for easy parsing, searching, and alerting. Metrics enable
 * performance monitoring and capacity planning.
 */

/**
 * Log levels with numeric priorities
 */
const LOG_LEVELS = {
  DEBUG: { level: 'debug', priority: 10 },
  INFO: { level: 'info', priority: 20 },
  WARN: { level: 'warn', priority: 30 },
  ERROR: { level: 'error', priority: 40 },
  FATAL: { level: 'fatal', priority: 50 }
};

/**
 * Metric types for different data collection patterns
 */
const METRIC_TYPES = {
  COUNTER: 'counter',      // Incrementing values (requests, errors)
  GAUGE: 'gauge',          // Point-in-time values (memory, connections)
  HISTOGRAM: 'histogram',  // Distribution of values (latency, size)
  TIMER: 'timer'          // Duration measurements
};

/**
 * Structured logger with correlation and context
 */
class StructuredLogger {
  
  constructor(component, config = {}) {
    this.component = component;
    this.config = {
      minLevel: LOG_LEVELS.INFO.priority,
      enableConsole: true,
      enableMetrics: true,
      maxContextSize: 1000,
      ...config
    };
    
    this.metrics = new MetricsCollector(component);
    this.contextStack = [];
  }
  
  /**
   * Add persistent context to all logs
   */
  addContext(context) {
    this.contextStack.push(context);
    return () => {
      const index = this.contextStack.indexOf(context);
      if (index > -1) {
        this.contextStack.splice(index, 1);
      }
    };
  }
  
  /**
   * Debug level logging
   */
  debug(message, context = {}) {
    this._log(LOG_LEVELS.DEBUG, message, context);
  }
  
  /**
   * Info level logging
   */
  info(message, context = {}) {
    this._log(LOG_LEVELS.INFO, message, context);
  }
  
  /**
   * Warning level logging
   */
  warn(message, context = {}) {
    this._log(LOG_LEVELS.WARN, message, context);
    this.metrics.increment('warnings', { component: this.component });
  }
  
  /**
   * Error level logging
   */
  error(message, context = {}) {
    this._log(LOG_LEVELS.ERROR, message, context);
    this.metrics.increment('errors', { component: this.component, level: 'error' });
  }
  
  /**
   * Fatal level logging
   */
  fatal(message, context = {}) {
    this._log(LOG_LEVELS.FATAL, message, context);
    this.metrics.increment('errors', { component: this.component, level: 'fatal' });
  }
  
  /**
   * Log operation with timing
   */
  operation(operationName, fn, context = {}) {
    const startTime = Date.now();
    const timer = this.metrics.startTimer(`operation.${operationName}`);
    
    this.debug(`Starting operation: ${operationName}`, context);
    
    try {
      const result = fn();
      
      // Handle async operations
      if (result && typeof result.then === 'function') {
        return result
          .then(asyncResult => {
            const duration = Date.now() - startTime;
            timer.end({ success: true });
            this.info(`Operation completed: ${operationName}`, { 
              ...context, 
              duration,
              success: true 
            });
            return asyncResult;
          })
          .catch(asyncError => {
            const duration = Date.now() - startTime;
            timer.end({ success: false, error: asyncError.code });
            this.error(`Operation failed: ${operationName}`, { 
              ...context, 
              duration,
              error: asyncError.message,
              errorCode: asyncError.code 
            });
            throw asyncError;
          });
      }
      
      // Synchronous result
      const duration = Date.now() - startTime;
      timer.end({ success: true });
      this.info(`Operation completed: ${operationName}`, { 
        ...context, 
        duration,
        success: true 
      });
      
      return result;
      
    } catch (syncError) {
      const duration = Date.now() - startTime;
      timer.end({ success: false, error: syncError.code });
      this.error(`Operation failed: ${operationName}`, { 
        ...context, 
        duration,
        error: syncError.message,
        errorCode: syncError.code 
      });
      throw syncError;
    }
  }
  
  /**
   * Log structured event
   */
  event(eventName, eventData = {}) {
    this._log(LOG_LEVELS.INFO, `Event: ${eventName}`, {
      eventType: 'structured_event',
      eventName,
      eventData
    });
    
    this.metrics.increment('events', { 
      component: this.component, 
      eventName 
    });
  }
  
  /**
   * Log API request/response
   */
  apiCall(method, url, statusCode, duration, context = {}) {
    const success = statusCode < 400;
    const level = success ? LOG_LEVELS.INFO : LOG_LEVELS.ERROR;
    
    this._log(level, `API ${method} ${url}`, {
      apiCall: true,
      method,
      url,
      statusCode,
      duration,
      success,
      ...context
    });
    
    this.metrics.increment('api_calls', { 
      component: this.component, 
      method, 
      status: this._getStatusCategory(statusCode) 
    });
    
    this.metrics.recordValue('api_duration', duration, { 
      method, 
      endpoint: this._extractEndpoint(url) 
    });
  }
  
  /**
   * Get metrics collector
   */
  getMetrics() {
    return this.metrics;
  }
  
  /**
   * Get logging statistics
   */
  getStats() {
    return this.metrics.getSnapshot();
  }
  
  // PRIVATE METHODS
  
  /**
   * Core logging method
   * @private
   */
  _log(levelObj, message, context = {}) {
    if (levelObj.priority < this.config.minLevel) {
      return;
    }
    
    const timestamp = new Date().toISOString();
    const correlationId = Correlation?.getId() || null;
    
    // Build context from stack and provided context
    const fullContext = this._buildContext(context);
    
    const logEntry = {
      timestamp,
      level: levelObj.level,
      component: this.component,
      message,
      correlation: correlationId,
      ...fullContext
    };
    
    // Truncate large contexts
    const serialized = JSON.stringify(logEntry);
    if (serialized.length > this.config.maxContextSize) {
      logEntry._truncated = true;
      logEntry._originalSize = serialized.length;
    }
    
    if (this.config.enableConsole) {
      console.log(JSON.stringify(logEntry));
    }
    
    // Forward to external systems if configured
    this._forwardLog(logEntry);
  }
  
  /**
   * Build context from stack and parameters
   * @private
   */
  _buildContext(context) {
    const fullContext = {};
    
    // Add stacked contexts
    for (const stackContext of this.contextStack) {
      Object.assign(fullContext, stackContext);
    }
    
    // Add provided context
    Object.assign(fullContext, context);
    
    return fullContext;
  }
  
  /**
   * Forward log to external systems
   * @private
   */
  _forwardLog(logEntry) {
    // Could integrate with external logging services
    // For now, store in PropertiesService for debugging
    try {
      if (logEntry.level === 'error' || logEntry.level === 'fatal') {
        const errorLogs = PropertiesService.getScriptProperties().getProperty('ERROR_LOGS') || '[]';
        const errors = JSON.parse(errorLogs);
        errors.push(logEntry);
        
        // Keep only last 100 errors
        if (errors.length > 100) {
          errors.splice(0, errors.length - 100);
        }
        
        PropertiesService.getScriptProperties().setProperty('ERROR_LOGS', JSON.stringify(errors));
      }
    } catch (e) {
      // Fail silently to avoid logging recursion
    }
  }
  
  /**
   * Get HTTP status category
   * @private
   */
  _getStatusCategory(statusCode) {
    if (statusCode < 300) return '2xx';
    if (statusCode < 400) return '3xx';
    if (statusCode < 500) return '4xx';
    return '5xx';
  }
  
  /**
   * Extract endpoint from URL
   * @private
   */
  _extractEndpoint(url) {
    try {
      const urlObj = new URL(url);
      return urlObj.pathname;
    } catch (e) {
      return url;
    }
  }
}

/**
 * Metrics collection and aggregation
 */
class MetricsCollector {
  
  constructor(component) {
    this.component = component;
    this.metrics = new Map();
    this.timers = new Map();
    this.startTime = Date.now();
  }
  
  /**
   * Increment a counter metric
   */
  increment(name, tags = {}, value = 1) {
    const key = this._getMetricKey(name, tags);
    const existing = this.metrics.get(key) || { 
      name, 
      type: METRIC_TYPES.COUNTER, 
      value: 0, 
      tags,
      updated: Date.now() 
    };
    
    existing.value += value;
    existing.updated = Date.now();
    this.metrics.set(key, existing);
  }
  
  /**
   * Set a gauge metric value
   */
  setGauge(name, value, tags = {}) {
    const key = this._getMetricKey(name, tags);
    this.metrics.set(key, {
      name,
      type: METRIC_TYPES.GAUGE,
      value,
      tags,
      updated: Date.now()
    });
  }
  
  /**
   * Record a value in a histogram
   */
  recordValue(name, value, tags = {}) {
    const key = this._getMetricKey(name, tags);
    const existing = this.metrics.get(key) || {
      name,
      type: METRIC_TYPES.HISTOGRAM,
      values: [],
      tags,
      updated: Date.now()
    };
    
    existing.values.push({ value, timestamp: Date.now() });
    
    // Keep only last 1000 values
    if (existing.values.length > 1000) {
      existing.values.splice(0, existing.values.length - 1000);
    }
    
    existing.updated = Date.now();
    this.metrics.set(key, existing);
  }
  
  /**
   * Start a timer
   */
  startTimer(name, tags = {}) {
    const timerId = `${name}_${Date.now()}_${Math.random()}`;
    const timer = {
      name,
      tags,
      startTime: Date.now(),
      end: (endTags = {}) => {
        const duration = Date.now() - timer.startTime;
        this.recordValue(name, duration, { ...tags, ...endTags });
        this.timers.delete(timerId);
        return duration;
      }
    };
    
    this.timers.set(timerId, timer);
    return timer;
  }
  
  /**
   * Get current metric snapshot
   */
  getSnapshot() {
    const snapshot = {
      component: this.component,
      timestamp: new Date().toISOString(),
      uptime: Date.now() - this.startTime,
      activeTimers: this.timers.size,
      metrics: {}
    };
    
    // Process metrics
    for (const [key, metric] of this.metrics) {
      if (metric.type === METRIC_TYPES.HISTOGRAM) {
        snapshot.metrics[key] = {
          ...metric,
          statistics: this._calculateHistogramStats(metric.values)
        };
        delete snapshot.metrics[key].values; // Remove raw values from snapshot
      } else {
        snapshot.metrics[key] = { ...metric };
      }
    }
    
    return snapshot;
  }
  
  /**
   * Get metrics for specific name pattern
   */
  getMetrics(namePattern = null) {
    const result = [];
    
    for (const [key, metric] of this.metrics) {
      if (!namePattern || metric.name.includes(namePattern)) {
        result.push({ key, ...metric });
      }
    }
    
    return result;
  }
  
  /**
   * Reset all metrics
   */
  reset() {
    this.metrics.clear();
    this.timers.clear();
    this.startTime = Date.now();
  }
  
  // PRIVATE METHODS
  
  /**
   * Generate metric key from name and tags
   * @private
   */
  _getMetricKey(name, tags) {
    const tagStr = Object.keys(tags)
      .sort()
      .map(key => `${key}=${tags[key]}`)
      .join(',');
    
    return `${name}[${tagStr}]`;
  }
  
  /**
   * Calculate histogram statistics
   * @private
   */
  _calculateHistogramStats(values) {
    if (values.length === 0) {
      return { count: 0, min: 0, max: 0, avg: 0, p50: 0, p95: 0, p99: 0 };
    }
    
    const sorted = values.map(v => v.value).sort((a, b) => a - b);
    const count = sorted.length;
    const sum = sorted.reduce((acc, val) => acc + val, 0);
    
    return {
      count,
      min: sorted[0],
      max: sorted[count - 1],
      avg: sum / count,
      p50: this._percentile(sorted, 0.5),
      p95: this._percentile(sorted, 0.95),
      p99: this._percentile(sorted, 0.99)
    };
  }
  
  /**
   * Calculate percentile value
   * @private
   */
  _percentile(sortedValues, percentile) {
    const index = Math.ceil(sortedValues.length * percentile) - 1;
    return sortedValues[Math.max(0, index)];
  }
}

/**
 * Health monitor for system components
 */
class HealthMonitor {
  
  constructor() {
    this.checks = new Map();
    this.logger = new StructuredLogger('HealthMonitor');
  }
  
  /**
   * Register a health check
   */
  register(name, checkFn, config = {}) {
    this.checks.set(name, {
      name,
      checkFn,
      config: {
        interval: 60000,      // 1 minute default
        timeout: 5000,        // 5 second timeout
        retries: 3,
        ...config
      },
      lastCheck: null,
      lastResult: null,
      failureCount: 0
    });
  }
  
  /**
   * Run all health checks
   */
  async runAll() {
    const results = new Map();
    
    for (const [name, check] of this.checks) {
      try {
        const result = await this._runCheck(check);
        results.set(name, result);
      } catch (error) {
        this.logger.error(`Health check failed: ${name}`, { error: error.message });
        results.set(name, {
          name,
          healthy: false,
          error: error.message,
          timestamp: new Date().toISOString()
        });
      }
    }
    
    return this._summarizeHealth(results);
  }
  
  /**
   * Run specific health check
   */
  async runCheck(name) {
    const check = this.checks.get(name);
    if (!check) {
      throw new Error(`Health check not found: ${name}`);
    }
    
    return this._runCheck(check);
  }
  
  /**
   * Get health status summary
   */
  getStatus() {
    const status = {
      timestamp: new Date().toISOString(),
      healthy: true,
      checks: []
    };
    
    for (const [name, check] of this.checks) {
      const checkStatus = {
        name,
        healthy: check.lastResult?.healthy || false,
        lastCheck: check.lastCheck,
        failureCount: check.failureCount
      };
      
      if (check.lastResult?.error) {
        checkStatus.error = check.lastResult.error;
      }
      
      status.checks.push(checkStatus);
      
      if (!checkStatus.healthy) {
        status.healthy = false;
      }
    }
    
    return status;
  }
  
  // PRIVATE METHODS
  
  /**
   * Execute single health check
   * @private
   */
  async _runCheck(check) {
    const startTime = Date.now();
    
    try {
      // Run check with timeout
      const result = await this._withTimeout(
        check.checkFn(),
        check.config.timeout
      );
      
      const duration = Date.now() - startTime;
      
      check.lastCheck = new Date().toISOString();
      check.lastResult = {
        name: check.name,
        healthy: true,
        duration,
        timestamp: check.lastCheck,
        data: result
      };
      check.failureCount = 0;
      
      this.logger.debug(`Health check passed: ${check.name}`, { duration });
      
      return check.lastResult;
      
    } catch (error) {
      const duration = Date.now() - startTime;
      
      check.lastCheck = new Date().toISOString();
      check.lastResult = {
        name: check.name,
        healthy: false,
        duration,
        timestamp: check.lastCheck,
        error: error.message
      };
      check.failureCount++;
      
      this.logger.warn(`Health check failed: ${check.name}`, { 
        error: error.message,
        failureCount: check.failureCount,
        duration 
      });
      
      return check.lastResult;
    }
  }
  
  /**
   * Run function with timeout
   * @private
   */
  async _withTimeout(promise, timeoutMs) {
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        reject(new Error(`Health check timeout after ${timeoutMs}ms`));
      }, timeoutMs);
      
      Promise.resolve(promise)
        .then(result => {
          clearTimeout(timer);
          resolve(result);
        })
        .catch(error => {
          clearTimeout(timer);
          reject(error);
        });
    });
  }
  
  /**
   * Summarize overall health
   * @private
   */
  _summarizeHealth(results) {
    const summary = {
      timestamp: new Date().toISOString(),
      healthy: true,
      totalChecks: results.size,
      passedChecks: 0,
      failedChecks: 0,
      checks: Array.from(results.values())
    };
    
    for (const result of results.values()) {
      if (result.healthy) {
        summary.passedChecks++;
      } else {
        summary.failedChecks++;
        summary.healthy = false;
      }
    }
    
    return summary;
  }
}

/**
 * Performance profiler for detailed analysis
 */
class PerformanceProfiler {
  
  constructor() {
    this.profiles = new Map();
    this.logger = new StructuredLogger('PerformanceProfiler');
  }
  
  /**
   * Start profiling session
   */
  startProfiling(sessionName) {
    const session = {
      name: sessionName,
      startTime: Date.now(),
      operations: [],
      memorySnapshots: [],
      markers: []
    };
    
    this.profiles.set(sessionName, session);
    this.logger.debug(`Started profiling session: ${sessionName}`);
    
    return {
      addMarker: (name, data = {}) => this._addMarker(session, name, data),
      recordOperation: (name, duration, data = {}) => this._recordOperation(session, name, duration, data),
      takeMemorySnapshot: () => this._takeMemorySnapshot(session),
      end: () => this._endProfiling(sessionName)
    };
  }
  
  /**
   * Get profiling results
   */
  getProfile(sessionName) {
    return this.profiles.get(sessionName);
  }
  
  /**
   * Get all active profiles
   */
  getActiveProfiles() {
    return Array.from(this.profiles.keys());
  }
  
  // PRIVATE METHODS
  
  /**
   * Add marker to profiling session
   * @private
   */
  _addMarker(session, name, data) {
    session.markers.push({
      name,
      timestamp: Date.now(),
      offset: Date.now() - session.startTime,
      data
    });
  }
  
  /**
   * Record operation in profiling session
   * @private
   */
  _recordOperation(session, name, duration, data) {
    session.operations.push({
      name,
      duration,
      timestamp: Date.now(),
      offset: Date.now() - session.startTime,
      data
    });
  }
  
  /**
   * Take memory snapshot
   * @private
   */
  _takeMemorySnapshot(session) {
    try {
      // Approximate memory usage for GAS
      const snapshot = {
        timestamp: Date.now(),
        offset: Date.now() - session.startTime,
        memoryUsed: this._getApproximateMemoryUsage()
      };
      
      session.memorySnapshots.push(snapshot);
      return snapshot;
    } catch (e) {
      this.logger.warn('Failed to take memory snapshot', { error: e.message });
      return null;
    }
  }
  
  /**
   * End profiling session
   * @private
   */
  _endProfiling(sessionName) {
    const session = this.profiles.get(sessionName);
    if (!session) return null;
    
    session.endTime = Date.now();
    session.totalDuration = session.endTime - session.startTime;
    
    const summary = {
      sessionName,
      totalDuration: session.totalDuration,
      totalOperations: session.operations.length,
      totalMarkers: session.markers.length,
      memorySnapshots: session.memorySnapshots.length,
      topOperations: this._getTopOperations(session),
      timeline: this._buildTimeline(session)
    };
    
    this.logger.info(`Profiling session completed: ${sessionName}`, summary);
    
    // Remove from active profiles
    this.profiles.delete(sessionName);
    
    return { session, summary };
  }
  
  /**
   * Get top operations by duration
   * @private
   */
  _getTopOperations(session) {
    return session.operations
      .sort((a, b) => b.duration - a.duration)
      .slice(0, 10)
      .map(op => ({ name: op.name, duration: op.duration }));
  }
  
  /**
   * Build timeline of events
   * @private
   */
  _buildTimeline(session) {
    const events = [
      ...session.markers.map(m => ({ ...m, type: 'marker' })),
      ...session.operations.map(o => ({ ...o, type: 'operation' }))
    ];
    
    return events.sort((a, b) => a.offset - b.offset);
  }
  
  /**
   * Approximate memory usage
   * @private
   */
  _getApproximateMemoryUsage() {
    // Very rough approximation for GAS
    return {
      estimated: true,
      driveQuotaUsed: DriveApp.getStorageUsed(),
      timestamp: Date.now()
    };
  }
}

// Global observability instances
const globalLogger = new StructuredLogger('Global');
const globalHealthMonitor = new HealthMonitor();
const globalProfiler = new PerformanceProfiler();

/**
 * Convenience observability facade
 */
const Observability = {
  // Logging
  logger: (component) => new StructuredLogger(component),
  log: globalLogger,
  
  // Metrics
  metrics: globalLogger.getMetrics(),
  
  // Health monitoring
  health: globalHealthMonitor,
  
  // Performance profiling
  profiler: globalProfiler,
  
  // Setup common health checks
  setupDefaultHealthChecks: () => {
    globalHealthMonitor.register('ooxml_service', async () => {
      const health = await OOXMLJsonService.healthCheck();
      if (!health.available) throw new Error('OOXML service unhealthy');
      return health;
    });
    
    globalHealthMonitor.register('deployment', () => {
      const status = OOXMLDeployment.getDeploymentStatus();
      if (!status.deployed) throw new Error('Service not deployed');
      return status;
    });
  },
  
  // Get system overview
  getSystemOverview: async () => {
    const health = globalHealthMonitor.getStatus();
    const metrics = globalLogger.getMetrics().getSnapshot();
    
    return {
      timestamp: new Date().toISOString(),
      health,
      metrics,
      activeProfiles: globalProfiler.getActiveProfiles()
    };
  }
};

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    StructuredLogger,
    MetricsCollector,
    HealthMonitor,
    PerformanceProfiler,
    Observability,
    LOG_LEVELS,
    METRIC_TYPES
  };
}