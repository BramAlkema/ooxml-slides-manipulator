/**
 * OOXMLParser - Core class for OOXML ZIP/XML manipulation
 * Handles the low-level unzipping, XML parsing, and rezipping of PPTX files
 */
class OOXMLParser {
  constructor(fileBlob) {
    this.originalBlob = fileBlob;
    this.files = new Map();
    this.isExtracted = false;
  }

  /**
   * Extract PPTX file (ZIP archive) into individual files
   */
  extract() {
    if (this.isExtracted) return this;
    
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
      return this;
    } catch (error) {
      throw new Error(`Failed to extract PPTX: ${error.message}`);
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
      const xmlString = XmlService.getPrettyFormat().format(xmlDoc);
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
   * @returns {Blob}
   */
  build() {
    if (!this.isExtracted) {
      throw new Error('Cannot build - archive not extracted');
    }
    
    try {
      const blobs = Array.from(this.files.values()).map(file => file.blob);
      return Utilities.zip(blobs);
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