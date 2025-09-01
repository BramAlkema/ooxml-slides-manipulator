/**
 * OOXMLParser - Core class for OOXML ZIP/XML manipulation
 * Handles the low-level unzipping, XML parsing, and rezipping of PPTX files
 * Enhanced with tanaikech's template approach
 */
class OOXMLParser {
  constructor(fileBlob) {
    this.originalBlob = fileBlob;
    this.files = new Map();
    this.isExtracted = false;
  }

  /**
   * Create new parser from PPTX template
   * @param {Object} options - Template options
   * @returns {OOXMLParser} New parser instance
   */
  static fromTemplate(options = {}) {
    const templateBlob = PPTXTemplate.createFromTemplate(options);
    return new OOXMLParser(templateBlob);
  }

  /**
   * Extract PPTX file (ZIP archive) into individual files
   * Uses FFlatePPTXService for reliable PPTX extraction
   */
  extract() {
    if (this.isExtracted) return this;
    
    try {
      // Use FFlatePPTXService for extraction (handles PPTX files properly)
      const extractedFiles = FFlatePPTXService.extractFiles(this.originalBlob);
      
      // Convert to internal format
      Object.entries(extractedFiles).forEach(([fileName, content]) => {
        const isXML = fileName.endsWith('.xml') || fileName.endsWith('.rels');
        let processedContent = content;
        
        // Cloud Function returns base64-encoded content, decode for XML files
        if (isXML && typeof content === 'string') {
          // Check if content looks like XML already (starts with < or <?xml)
          if (content.trim().startsWith('<')) {
            // Already text content, use as is
            processedContent = content;
          } else {
            // Try to decode base64 content
            try {
              // Try standard base64 decode first
              const decoded = Utilities.base64Decode(content);
              processedContent = Utilities.newBlob(decoded).getDataAsString('UTF-8');
            } catch (decodeError) {
              try {
                // Try base64WebSafe decode as fallback
                const decoded = Utilities.base64DecodeWebSafe(content);
                processedContent = Utilities.newBlob(decoded).getDataAsString('UTF-8');
              } catch (decodeError2) {
                // If both decode methods fail, assume it's already text content
                processedContent = content;
              }
            }
          }
        }
        
        this.files.set(fileName, {
          content: processedContent,
          isXML: isXML,
          blob: null // We have text content, can recreate blob if needed
        });
      });
      
      this.isExtracted = true;
      console.log(`✅ Extracted ${this.files.size} files from PPTX`);
      return this;
      
    } catch (error) {
      // Fallback to native Google Apps Script method if Cloud Function fails
      console.log('Cloud extraction failed, trying native method:', error.message);
      
      try {
        const zipBlob = Utilities.unzip(this.originalBlob);
        
        zipBlob.forEach(file => {
          const fileName = file.getName();
          const content = file.getDataAsString();
          this.files.set(fileName, {
            content: content,
            isXML: fileName.endsWith('.xml') || fileName.endsWith('.rels'),
            blob: file
          });
        });
        
        this.isExtracted = true;
        console.log(`✅ Extracted ${this.files.size} files using native method`);
        return this;
        
      } catch (nativeError) {
        throw new Error(
          `Failed to extract PPTX: Cloud Function error: ${error.message}. ` +
          `Native method error: ${nativeError.message}. ` +
          `Consider deploying the Cloud Function for better PPTX support.`
        );
      }
    }
  }

  /**
   * Get XML content as parsed document
   * @param {string} path - Path to XML file (e.g., 'ppt/presentation.xml')
   * @returns {XmlService.Document}
   */
  getXML(path) {
    if (!this.isExtracted) this.extract();
    
    const file = this.files.get(path);
    if (!file) {
      throw new Error(`File not found: ${path}`);
    }
    
    if (!file.isXML) {
      throw new Error(`File is not XML: ${path}`);
    }
    
    try {
      return XmlService.parse(file.content);
    } catch (error) {
      throw new Error(`Failed to parse XML ${path}: ${error.message}`);
    }
  }

  /**
   * Update XML file with new content
   * @param {string} path - Path to XML file
   * @param {XmlService.Document} xmlDoc - Updated XML document
   */
  setXML(path, xmlDoc) {
    if (!this.isExtracted) this.extract();
    
    const file = this.files.get(path);
    if (!file) {
      throw new Error(`File not found: ${path}`);
    }
    
    try {
      // Use getRawFormat() like tanaikech - preserves OOXML structure better
      const xmlString = XmlService.getRawFormat().format(xmlDoc);
      file.content = xmlString;
      file.blob = Utilities.newBlob(xmlString, 'application/xml', path);
    } catch (error) {
      throw new Error(`Failed to update XML ${path}: ${error.message}`);
    }
  }

  /**
   * Get raw file content (for non-XML files like images)
   * @param {string} path - Path to file
   * @returns {string|Blob}
   */
  getFile(path) {
    if (!this.isExtracted) this.extract();
    
    const file = this.files.get(path);
    if (!file) {
      throw new Error(`File not found: ${path}`);
    }
    
    return file.isXML ? file.content : file.blob;
  }

  /**
   * Add or update a file in the archive
   * @param {string} path - Path to file
   * @param {string|Blob} content - File content
   */
  setFile(path, content) {
    if (!this.isExtracted) this.extract();
    
    const isXML = path.endsWith('.xml') || path.endsWith('.rels');
    let blob;
    
    if (typeof content === 'string') {
      blob = Utilities.newBlob(content, isXML ? 'application/xml' : 'application/octet-stream', path);
    } else {
      blob = content;
    }
    
    this.files.set(path, {
      content: typeof content === 'string' ? content : null,
      isXML: isXML,
      blob: blob
    });
  }

  /**
   * List all files in the archive
   * @returns {Array<string>} Array of file paths
   */
  listFiles() {
    if (!this.isExtracted) this.extract();
    return Array.from(this.files.keys());
  }

  /**
   * Check if file exists in archive
   * @param {string} path - Path to check
   * @returns {boolean}
   */
  hasFile(path) {
    if (!this.isExtracted) this.extract();
    return this.files.has(path);
  }

  /**
   * Create new PPTX blob from current files
   * Uses FFlatePPTXService for reliable PPTX creation
   * @returns {Blob}
   */
  build() {
    if (!this.isExtracted) {
      throw new Error('Cannot build - archive not extracted');
    }
    
    try {
      // Prepare files for FFlatePPTXService (content strings)
      const fileContents = {};
      
      this.files.forEach((file, fileName) => {
        if (file.content !== null) {
          // We have string content
          fileContents[fileName] = file.content;
        } else if (file.blob) {
          // Convert blob to string
          fileContents[fileName] = file.blob.getDataAsString();
        } else {
          throw new Error(`No content available for file: ${fileName}`);
        }
      });
      
      // Use FFlatePPTXService for reliable zipping
      try {
        console.log(`Building PPTX with ${Object.keys(fileContents).length} files via Cloud Function...`);
        return FFlatePPTXService.compressFiles(fileContents);
        
      } catch (cloudError) {
        console.log('Cloud build failed, trying native method:', cloudError.message);
        
        // Fallback to native method
        const blobs = [];
        this.files.forEach((file, fileName) => {
          if (file.blob) {
            blobs.push(file.blob);
          } else {
            // Create blob from content
            const blob = Utilities.newBlob(
              file.content,
              file.isXML ? 'application/xml' : 'application/octet-stream',
              fileName
            );
            blobs.push(blob);
          }
        });
        
        const result = Utilities.zip(blobs);
        console.log(`✅ Built PPTX using native method: ${result.getBytes().length} bytes`);
        return result;
      }
      
    } catch (error) {
      throw new Error(`Failed to build PPTX: ${error.message}`);
    }
  }

  /**
   * Get theme XML specifically (common operation)
   * @returns {XmlService.Document}
   */
  getThemeXML() {
    return this.getXML('ppt/theme/theme1.xml');
  }

  /**
   * Set theme XML specifically (common operation)
   * @param {XmlService.Document} themeDoc
   */
  setThemeXML(themeDoc) {
    this.setXML('ppt/theme/theme1.xml', themeDoc);
  }

  /**
   * Get presentation XML
   * @returns {XmlService.Document}
   */
  getPresentationXML() {
    return this.getXML('ppt/presentation.xml');
  }

  /**
   * Set presentation XML
   * @param {XmlService.Document} presDoc
   */
  setPresentationXML(presDoc) {
    this.setXML('ppt/presentation.xml', presDoc);
  }
}