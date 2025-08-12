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
    return 'http://localhost:8080'; // For local testing
    // return 'https://your-region-your-project.cloudfunctions.net/pptx-router'; // For production
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
      const pptxBlob = Utilities.newBlob(pptxData, MimeType.MICROSOFT_POWERPOINT, 'presentation.pptx');\n      \n      console.log(`✅ Successfully zipped ${result.fileCount} files into PPTX: ${result.size} bytes`);\n      \n      return pptxBlob;\n      \n    } catch (error) {\n      console.error('Failed to zip PPTX via Cloud Function:', error);\n      throw new Error(`Cloud zip error: ${error.message}`);\n    }\n  }\n  \n  /**\n   * Test the Cloud Function connectivity\n   * @returns {boolean} True if Cloud Function is accessible\n   */\n  static testCloudFunction() {\n    try {\n      console.log('Testing Cloud Function connectivity...');\n      \n      // Simple round-trip test\n      const testFiles = {\n        'test.txt': 'Hello from Google Apps Script!'\n      };\n      \n      const pptxBlob = this.zipPPTX(testFiles);\n      const extractedFiles = this.unzipPPTX(pptxBlob);\n      \n      const success = extractedFiles['test.txt'] === testFiles['test.txt'];\n      \n      if (success) {\n        console.log('✅ Cloud Function connectivity test passed!');\n      } else {\n        console.log('❌ Cloud Function connectivity test failed: content mismatch');\n      }\n      \n      return success;\n      \n    } catch (error) {\n      console.error('❌ Cloud Function connectivity test failed:', error);\n      return false;\n    }\n  }\n  \n  /**\n   * Check if we can use Cloud Function (fallback detection)\n   * @returns {boolean} True if Cloud Function is available\n   */\n  static isCloudFunctionAvailable() {\n    try {\n      // Try a simple request with short timeout\n      const response = UrlFetchApp.fetch(`${this.CLOUD_FUNCTION_URL}/zip`, {\n        method: 'POST',\n        headers: { 'Content-Type': 'application/json' },\n        payload: JSON.stringify({ files: { 'test.txt': 'test' } }),\n        muteHttpExceptions: true\n      });\n      \n      return response.getResponseCode() === 200;\n      \n    } catch (error) {\n      console.log('Cloud Function not available:', error.message);\n      return false;\n    }\n  }\n  \n  /**\n   * Enhanced PPTX manipulation that uses Cloud Function when available\n   * Falls back to Google Apps Script native methods when Cloud Function is unavailable\n   * @param {Blob} pptxBlob - Original PPTX blob\n   * @param {Function} manipulator - Function that takes files object and returns modified files\n   * @returns {Blob} Modified PPTX blob\n   */\n  static manipulatePPTX(pptxBlob, manipulator) {\n    try {\n      if (this.isCloudFunctionAvailable()) {\n        console.log('Using Cloud Function for PPTX manipulation...');\n        \n        // Extract files using Cloud Function\n        const files = this.unzipPPTX(pptxBlob);\n        \n        // Apply manipulations\n        const modifiedFiles = manipulator(files);\n        \n        // Repackage using Cloud Function\n        return this.zipPPTX(modifiedFiles);\n        \n      } else {\n        console.log('Cloud Function unavailable, falling back to Google Apps Script...');\n        \n        // Try Google Apps Script native unzip (will fail for most PPTX files)\n        try {\n          const files = Utilities.unzip(pptxBlob);\n          const fileMap = {};\n          \n          files.forEach(file => {\n            fileMap[file.getName()] = file.getDataAsString();\n          });\n          \n          const modifiedFiles = manipulator(fileMap);\n          \n          // Convert back to blobs\n          const blobs = [];\n          Object.entries(modifiedFiles).forEach(([filename, content]) => {\n            blobs.push(Utilities.newBlob(content, 'text/plain', filename));\n          });\n          \n          return Utilities.zip(blobs);\n          \n        } catch (nativeError) {\n          throw new Error(\n            `PPTX manipulation failed: Cloud Function unavailable and Google Apps Script cannot unzip this PPTX. ` +\n            `Deploy the Cloud Function or use a different PPTX file. Native error: ${nativeError.message}`\n          );\n        }\n      }\n      \n    } catch (error) {\n      throw new Error(`PPTX manipulation failed: ${error.message}`);\n    }\n  }\n}