/**
 * ConsoleFormatter - Utility for consistent console output formatting
 * 
 * Provides standardized formatting functions for console output across
 * the OOXML system, including headers, separators, status indicators, etc.
 */

class ConsoleFormatter {
  
  /**
   * Create a header with title and separator line
   * @param {string} title - Header title
   * @param {string} char - Character for separator (default: '=')
   * @param {number} width - Total width (default: auto-size to title)
   */
  static header(title, char = '=', width = null) {
    const actualWidth = width || Math.max(title.length, 40);
    console.log(title);
    console.log(char.repeat(actualWidth));
  }
  
  /**
   * Create a section header with title and separator
   * @param {string} title - Section title
   * @param {string} char - Character for separator (default: '-')
   */
  static section(title, char = '-') {
    console.log('\n' + title);
    console.log(char.repeat(title.length));
  }
  
  /**
   * Create a separator line
   * @param {string} char - Character for separator (default: '-')
   * @param {number} width - Width of separator (default: 40)
   */
  static separator(char = '-', width = 40) {
    console.log(char.repeat(width));
  }
  
  /**
   * Log with status indicator
   * @param {string} status - Status ('PASS', 'FAIL', 'WARN', 'INFO', 'ERROR', 'ACTION')
   * @param {string} message - Message to log
   * @param {string} details - Optional details
   */
  static status(status, message, details = null) {
    const icons = {
      'PASS': '✅',
      'FAIL': '❌', 
      'WARN': '⚠️',
      'INFO': 'ℹ️',
      'ERROR': '🚨',
      'ACTION': '🔧',
      'SKIP': '⏭️'
    };
    
    const icon = icons[status] || '•';
    const fullMessage = details ? `${message}: ${details}` : message;
    console.log(`  ${icon} ${fullMessage}`);
  }
  
  /**
   * Log a test result
   * @param {string} testName - Name of the test
   * @param {boolean} passed - Whether test passed
   * @param {string} details - Optional details
   * @param {number} duration - Optional duration in ms
   */
  static testResult(testName, passed, details = null, duration = null) {
    const status = passed ? 'PASS' : 'FAIL';
    const durationStr = duration ? ` (${duration}ms)` : '';
    const detailsStr = details ? ` - ${details}` : '';
    
    this.status(status, `${testName}${detailsStr}${durationStr}`);
  }
  
  /**
   * Log a progress indicator
   * @param {number} current - Current progress
   * @param {number} total - Total items
   * @param {string} label - Progress label
   */
  static progress(current, total, label = 'Progress') {
    const percentage = Math.round((current / total) * 100);
    const bar = '█'.repeat(Math.floor(percentage / 5)) + '░'.repeat(20 - Math.floor(percentage / 5));
    console.log(`${label}: [${bar}] ${current}/${total} (${percentage}%)`);
  }
  
  /**
   * Log a summary box
   * @param {string} title - Summary title
   * @param {Object} stats - Statistics object
   */
  static summary(title, stats) {
    console.log('\n' + '='.repeat(50));
    console.log(`📋 ${title.toUpperCase()}`);
    console.log('='.repeat(50));
    
    Object.entries(stats).forEach(([key, value]) => {
      const displayKey = key.charAt(0).toUpperCase() + key.slice(1).replace(/([A-Z])/g, ' $1');
      console.log(`${displayKey}: ${value}`);
    });
  }
  
  /**
   * Log with indentation
   * @param {number} level - Indentation level
   * @param {string} message - Message to log
   * @param {string} prefix - Optional prefix character
   */
  static indent(level, message, prefix = ' ') {
    const indent = prefix.repeat(level * 2);
    console.log(`${indent}${message}`);
  }
  
  /**
   * Log a step in a process
   * @param {number} stepNumber - Step number
   * @param {string} title - Step title
   * @param {string} description - Optional description
   */
  static step(stepNumber, title, description = null) {
    const stepTitle = `STEP ${stepNumber}: ${title.toUpperCase()}`;
    console.log(`\n📋 ${stepTitle}`);
    console.log('-'.repeat(stepTitle.length));
    
    if (description) {
      console.log(description);
    }
  }
  
  /**
   * Log a list of items with consistent formatting
   * @param {Array} items - Array of items to list
   * @param {string} bullet - Bullet character (default: '•')
   * @param {number} indent - Indentation level (default: 1)
   */
  static list(items, bullet = '•', indent = 1) {
    const indentStr = ' '.repeat(indent * 2);
    items.forEach(item => {
      console.log(`${indentStr}${bullet} ${item}`);
    });
  }
  
  /**
   * Log an error with stack trace formatting
   * @param {string} message - Error message
   * @param {Error} error - Error object (optional)
   * @param {string} context - Context information (optional)
   */
  static error(message, error = null, context = null) {
    console.error(`🚨 ERROR: ${message}`);
    
    if (context) {
      console.error(`   Context: ${context}`);
    }
    
    if (error) {
      console.error(`   Details: ${error.message}`);
      if (error.stack) {
        console.error(`   Stack: ${error.stack.split('\n')[1]?.trim()}`);
      }
    }
  }
  
  /**
   * Log a warning with context
   * @param {string} message - Warning message
   * @param {string} recommendation - Optional recommendation
   */
  static warning(message, recommendation = null) {
    console.warn(`⚠️  WARNING: ${message}`);
    if (recommendation) {
      console.warn(`   Recommendation: ${recommendation}`);
    }
  }
  
  /**
   * Log an informational message
   * @param {string} message - Info message
   * @param {string} category - Optional category
   */
  static info(message, category = null) {
    const prefix = category ? `ℹ️  [${category.toUpperCase()}]` : 'ℹ️ ';
    console.log(`${prefix} ${message}`);
  }
  
  /**
   * Log success message
   * @param {string} message - Success message
   */
  static success(message) {
    console.log(`✅ ${message}`);
  }
  
  /**
   * Create a table-like output
   * @param {Array} headers - Table headers
   * @param {Array} rows - Array of row arrays
   * @param {number} columnWidth - Width of each column (default: 20)
   */
  static table(headers, rows, columnWidth = 20) {
    // Header
    const headerRow = headers.map(h => h.padEnd(columnWidth)).join(' | ');
    console.log(headerRow);
    console.log('-'.repeat(headerRow.length));
    
    // Rows
    rows.forEach(row => {
      const formattedRow = row.map(cell => 
        String(cell).padEnd(columnWidth)
      ).join(' | ');
      console.log(formattedRow);
    });
  }
  
  /**
   * Clear console (if supported)
   */
  static clear() {
    // GAS doesn't support console.clear(), so we'll simulate it
    console.log('\n'.repeat(50));
    console.log('Console cleared');
    console.log('='.repeat(20));
  }
}

// Convenience aliases for common operations
const fmt = ConsoleFormatter;

// Export both the class and convenience alias
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { ConsoleFormatter, fmt };
}