/**
 * WEB APP REQUEST HANDLER
 * Legacy handler - functionality moved to Main.js
 * This file exists for backwards compatibility only
 */

// Note: doPost and doGet are now handled by Main.js
// This file is kept for reference but handlers are disabled to avoid conflicts

function doPost_Legacy(e) {
  try {
    var body = e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    var fn = body.fn;
    var args = body.args || [];
    
    ConsoleFormatter.info(`Web app request: ${fn}`, 'REQUEST');
    
    if (!fn) {
      return ContentService.createTextOutput(JSON.stringify({ 
        error: 'Missing function name (fn parameter)' 
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Check if function exists in global scope
    if (typeof this[fn] !== 'function') {
      return ContentService.createTextOutput(JSON.stringify({ 
        error: `Function '${fn}' not found or not accessible` 
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Execute the function
    var result = this[fn](...args);
    
    ConsoleFormatter.success(`Function '${fn}' executed successfully`);
    
    return ContentService.createTextOutput(JSON.stringify({ 
      result: result 
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    ConsoleFormatter.error('Web app error', error);
    return ContentService.createTextOutput(JSON.stringify({ 
      error: error.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet_Legacy(e) {
  // Handle GET requests for testing
  var fn = e.parameter.fn || 'ping';
  var name = e.parameter.name || 'GET-Request';
  
  try {
    // Allow specific test functions via GET for automation
    var allowedGETFunctions = ['ping', 'testOOXMLCore', 'testFFlatePPTXService', 'testExtensionSystem', 'runAcidTest'];
    
    if (allowedGETFunctions.includes(fn)) {
      let result;
      if (fn === 'ping') {
        result = ping(name);
      } else if (typeof this[fn] === 'function') {
        result = this[fn]();
      } else {
        throw new Error(`Function '${fn}' not found`);
      }
      
      return ContentService.createTextOutput(JSON.stringify({ 
        result: result 
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ 
      error: `GET method only supports: ${allowedGETFunctions.join(', ')}, requested: ${fn}` 
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ 
      error: error.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}