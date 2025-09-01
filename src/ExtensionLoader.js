/**
 * Extension Loader - Initialize and auto-load all extensions
 * 
 * Call this once at startup to automatically discover and register
 * all extensions with OOXMLSlides.
 */

/**
 * Initialize the extension system and auto-load all extensions
 * Call this once when the script starts
 */
async function initializeExtensions() {
  try {
    ConsoleFormatter.header('🚀 Initializing Extension System');
    
    // Auto-discover and load all extensions
    const results = await ExtensionAutoLoader.discoverAndLoad({
      autoRegister: true,
      exclude: [] // Exclude any problematic extensions
    });
    
    if (results.success) {
      ConsoleFormatter.success(`Extension system initialized successfully`);
      ConsoleFormatter.info(`${results.loaded} extensions loaded, ${results.registered} methods registered`);
      
      // Log available methods
      const extensions = ExtensionAutoLoader.getLoadedExtensions();
      extensions.forEach(ext => {
        ConsoleFormatter.indent(1, `📦 ${ext.name}: ${ext.methods.length} methods`);
        ext.methods.forEach(method => {
          const prefix = method.static ? 'OOXMLSlides.' : 'slides.';
          ConsoleFormatter.indent(2, `• ${prefix}${method.name}()`);
        });
      });
      
      return results;
    } else {
      throw new Error('Extension initialization failed');
    }
    
  } catch (error) {
    ConsoleFormatter.error('Failed to initialize extensions', error);
    throw error;
  }
}

/**
 * Quick extension status check
 */
function getExtensionStatus() {
  const loaded = ExtensionAutoLoader.getLoadedExtensions();
  
  return {
    loaded: loaded.length,
    extensions: loaded.map(ext => ext.name),
    totalMethods: loaded.reduce((sum, ext) => sum + ext.methods.length, 0)
  };
}

/**
 * Test that extensions are working
 */
async function testExtensions() {
  ConsoleFormatter.header('🧪 Testing Extension System');
  
  try {
    // Test that createFromPrompt was added to OOXMLSlides
    if (typeof OOXMLSlides.createFromPrompt === 'function') {
      ConsoleFormatter.status('PASS', 'createFromPrompt', 'Method registered successfully');
    } else {
      ConsoleFormatter.status('FAIL', 'createFromPrompt', 'Method not found');
      return false;
    }
    
    // Test other registered methods
    const status = getExtensionStatus();
    ConsoleFormatter.summary('Extension System Status', {
      'Extensions Loaded': status.loaded,
      'Total Methods': status.totalMethods,
      'Extensions': status.extensions.join(', ')
    });
    
    ConsoleFormatter.success('Extension system test completed');
    return true;
    
  } catch (error) {
    ConsoleFormatter.error('Extension system test failed', error);
    return false;
  }
}