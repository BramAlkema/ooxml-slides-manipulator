/**
 * CloudPPTXService - Bridge between Google Apps Script and Cloud Function
 * Provides ZIP/unzip capabilities for PPTX manipulation using Google Cloud Function
 * Solves the limitation that Google Apps Script's Utilities.unzip() cannot handle PPTX files
 */
class CloudPPTXService {
  
  /**
   * Cloud Function URL for PPTX operations
   * TODO: Replace with actual deployed function URL
   */
  static get CLOUD_FUNCTION_URL() {
    return 'https://pptx-router-149429464972.us-central1.run.app'; // Production
    // return 'http://localhost:8080'; // For local testing
  }
  
  /**
   * Unzip a PPTX blob using Cloud Function
   * @param {Blob} pptxBlob - The PPTX file to unzip
   * @returns {Object} Object with file contents keyed by filename
   */
  static unzipPPTX(pptxBlob) {
    try {
      console.log(`Unzipping PPTX via Cloud Function: ${pptxBlob.getBytes().length} bytes`);
      
      // Convert blob to base64 for transmission
      const pptxBase64 = Utilities.base64Encode(pptxBlob.getBytes());
      
      // Call Cloud Function unzip endpoint
      const response = UrlFetchApp.fetch(`${this.CLOUD_FUNCTION_URL}/unzip`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
          data: pptxBase64
        })
      });
      
      const result = JSON.parse(response.getContentText());
      
      if (!result.success) {
        throw new Error(`Cloud unzip failed: ${result.error || 'Unknown error'}`);
      }
      
      console.log(`✅ Successfully unzipped ${result.fileCount} files from PPTX`);
      console.log(`   Files: ${Object.keys(result.files).join(', ')}`);
      
      return result.files;
      
    } catch (error) {
      console.error('Failed to unzip PPTX via Cloud Function:', error);
      throw new Error(`Cloud unzip error: ${error.message}`);
    }
  }
  
  /**
   * Zip files into a PPTX blob using Cloud Function
   * @param {Object} files - Object with filename -> content mappings
   * @returns {Blob} PPTX blob
   */
  static zipPPTX(files) {
    try {
      console.log(`Zipping ${Object.keys(files).length} files via Cloud Function`);
      
      // Call Cloud Function zip endpoint
      const response = UrlFetchApp.fetch(`${this.CLOUD_FUNCTION_URL}/zip`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
          files: files
        })
      });
      
      const result = JSON.parse(response.getContentText());
      
      if (!result.success) {
        throw new Error(`Cloud zip failed: ${result.error || 'Unknown error'}`);
      }
      
      // Convert base64 result back to blob
      const pptxData = Utilities.base64Decode(result.pptxData);
      const pptxBlob = Utilities.newBlob(pptxData, MimeType.MICROSOFT_POWERPOINT, 'presentation.pptx');
      
      console.log(`✅ Successfully zipped ${result.fileCount} files into PPTX: ${result.size} bytes`);
      
      return pptxBlob;
      
    } catch (error) {
      console.error('Failed to zip PPTX via Cloud Function:', error);
      throw new Error(`Cloud zip error: ${error.message}`);
    }
  }
  
  /**
   * Test the Cloud Function connectivity
   * @returns {boolean} True if Cloud Function is accessible
   */
  static testCloudFunction() {
    try {
      console.log('Testing Cloud Function connectivity...');
      
      // Test with actual PPTX data using the unzip endpoint
      const templateBlob = PPTXTemplate.createMinimalTemplate();
      const result = this.unzipPPTX(templateBlob);
      
      // Check if we got valid PPTX structure back
      const hasExpectedFiles = result && 
        typeof result === 'object' &&
        Object.keys(result).length > 0;
      
      if (hasExpectedFiles) {
        console.log('✅ Cloud Function connectivity test passed!');
        console.log(`   Extracted ${Object.keys(result).length} files`);
        return true;
      } else {
        console.log('❌ Cloud Function connectivity test failed: invalid response structure');
        return false;
      }
      
    } catch (error) {
      console.error('❌ Cloud Function connectivity test failed:', error.message);
      return false;
    }
  }
  
  /**
   * Check if we can use Cloud Function (fallback detection)
   * @returns {boolean} True if Cloud Function is available
   */
  static isCloudFunctionAvailable() {
    try {
      // Try a simple GET request to test connectivity
      const response = UrlFetchApp.fetch(`${this.CLOUD_FUNCTION_URL}/test`, {
        method: 'GET',
        muteHttpExceptions: true
      });
      
      // Accept any response that's not a network error (even 404/405 means service is up)
      const responseCode = response.getResponseCode();
      return responseCode >= 200 && responseCode < 500;
      
    } catch (error) {
      console.log('Cloud Function not available:', error.message);
      return false;
    }
  }
  
  /**
   * Enhanced PPTX manipulation that uses Cloud Function when available
   * Falls back to Google Apps Script native methods when Cloud Function is unavailable
   * @param {Blob} pptxBlob - Original PPTX blob
   * @param {Function} manipulator - Function that takes files object and returns modified files
   * @returns {Blob} Modified PPTX blob
   */
  static manipulatePPTX(pptxBlob, manipulator) {
    try {
      if (this.isCloudFunctionAvailable()) {
        console.log('Using Cloud Function for PPTX manipulation...');
        
        // Extract files using Cloud Function
        const files = this.unzipPPTX(pptxBlob);
        
        // Apply manipulations
        const modifiedFiles = manipulator(files);
        
        // Repackage using Cloud Function
        return this.zipPPTX(modifiedFiles);
        
      } else {
        console.log('Cloud Function unavailable, falling back to Google Apps Script...');
        
        // Try Google Apps Script native unzip (will fail for most PPTX files)
        try {
          const files = Utilities.unzip(pptxBlob);
          const fileMap = {};
          
          files.forEach(file => {
            fileMap[file.getName()] = file.getDataAsString();
          });
          
          const modifiedFiles = manipulator(fileMap);
          
          // Convert back to blobs
          const blobs = [];
          Object.entries(modifiedFiles).forEach(([filename, content]) => {
            blobs.push(Utilities.newBlob(content, 'text/plain', filename));
          });
          
          return Utilities.zip(blobs);
          
        } catch (nativeError) {
          throw new Error(
            `PPTX manipulation failed: Cloud Function unavailable and Google Apps Script cannot unzip this PPTX. ` +
            `Deploy the Cloud Function or use a different PPTX file. Native error: ${nativeError.message}`
          );
        }
      }
      
    } catch (error) {
      throw new Error(`PPTX manipulation failed: ${error.message}`);
    }
  }
}