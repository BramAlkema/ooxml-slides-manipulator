/**
 * FileHandler - Google Drive integration utilities
 * Handles file loading, saving, and Drive API operations
 */
class FileHandler {
  constructor() {
    this.mimeTypes = {
      PPTX: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      POTX: 'application/vnd.openxmlformats-officedocument.presentationml.template',
      PPSX: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow',
      PPAM: 'application/vnd.ms-powerpoint.addin.macroEnabled.12',
      PPTM: 'application/vnd.ms-powerpoint.presentation.macroEnabled.12',
      POTM: 'application/vnd.ms-powerpoint.template.macroEnabled.12',
      PPSM: 'application/vnd.ms-powerpoint.slideshow.macroEnabled.12',
      GOOGLE_SLIDES: 'application/vnd.google-apps.presentation'
    };
    
    // All supported PowerPoint MIME types
    this.supportedPPTXTypes = [
      this.mimeTypes.PPTX,
      this.mimeTypes.POTX,
      this.mimeTypes.PPSX,
      this.mimeTypes.PPAM,
      this.mimeTypes.PPTM,
      this.mimeTypes.POTM,
      this.mimeTypes.PPSM
    ];
  }

  /**
   * Load PPTX file from Google Drive
   * @param {string} fileId - Google Drive file ID
   * @returns {Blob} File blob
   */
  loadFile(fileId) {
    try {
      const file = DriveApp.getFileById(fileId);
      const mimeType = file.getBlob().getContentType();
      
      if (mimeType === this.mimeTypes.GOOGLE_SLIDES) {
        // Convert Google Slides to PPTX
        return this._convertSlidesToPPTX(fileId);
      } else if (this.supportedPPTXTypes.includes(mimeType) || mimeType === 'application/zip') {
        return file.getBlob();
      } else {
        throw new Error(`Unsupported file type: ${mimeType}. Supported: ${this.supportedPPTXTypes.join(', ')}`);
      }
    } catch (error) {
      throw new Error(`Failed to load file ${fileId}: ${error.message}`);
    }
  }

  /**
   * Save modified PPTX to Google Drive
   * @param {Blob} pptxBlob - Modified PPTX blob
   * @param {Object} options - Save options
   * @returns {string} New file ID
   */
  saveFile(pptxBlob, options = {}) {
    try {
      const fileName = options.name || `Modified_Presentation_${new Date().getTime()}.pptx`;
      const folderId = options.folderId;
      
      let file;
      if (folderId) {
        const folder = DriveApp.getFolderById(folderId);
        file = folder.createFile(pptxBlob.setName(fileName));
      } else {
        file = DriveApp.createFile(pptxBlob.setName(fileName));
      }
      
      return file.getId();
    } catch (error) {
      throw new Error(`Failed to save file: ${error.message}`);
    }
  }

  /**
   * Update existing file with modified content
   * @param {string} fileId - Existing file ID
   * @param {Blob} pptxBlob - Modified PPTX blob
   */
  updateFile(fileId, pptxBlob) {
    try {
      const file = DriveApp.getFileById(fileId);
      file.setContent(pptxBlob);
      return fileId;
    } catch (error) {
      throw new Error(`Failed to update file ${fileId}: ${error.message}`);
    }
  }

  /**
   * Import PPTX back to Google Slides
   * @param {Blob} pptxBlob - PPTX blob to import
   * @param {Object} options - Import options
   * @returns {string} Google Slides file ID
   */
  importToSlides(pptxBlob, options = {}) {
    try {
      // First save as PPTX
      const tempFileId = this.saveFile(pptxBlob, { name: 'temp_import.pptx' });
      
      // Convert to Google Slides
      const resource = {
        title: options.name || 'Imported Presentation',
        mimeType: this.mimeTypes.GOOGLE_SLIDES,
        parents: options.folderId ? [{ id: options.folderId }] : undefined
      };
      
      // Use Drive API to convert
      const slidesFile = Drive.Files.insert(resource, DriveApp.getFileById(tempFileId).getBlob(), {
        convert: true
      });
      
      // Clean up temp file
      DriveApp.getFileById(tempFileId).setTrashed(true);
      
      return slidesFile.id;
    } catch (error) {
      throw new Error(`Failed to import to Slides: ${error.message}`);
    }
  }

  /**
   * Get file metadata
   * @param {string} fileId - File ID
   * @returns {Object} File metadata
   */
  getFileInfo(fileId) {
    try {
      const file = DriveApp.getFileById(fileId);
      return {
        id: fileId,
        name: file.getName(),
        mimeType: file.getBlob().getContentType(),
        size: file.getSize(),
        lastModified: file.getLastUpdated(),
        url: file.getUrl()
      };
    } catch (error) {
      throw new Error(`Failed to get file info ${fileId}: ${error.message}`);
    }
  }

  /**
   * Create a copy of the file before modification
   * @param {string} fileId - Original file ID
   * @param {string} suffix - Suffix for backup name
   * @returns {string} Backup file ID
   */
  createBackup(fileId, suffix = '_backup') {
    try {
      const originalFile = DriveApp.getFileById(fileId);
      const backupName = originalFile.getName().replace(/\.[^/.]+$/, '') + suffix + '.pptx';
      const backup = originalFile.makeCopy(backupName);
      return backup.getId();
    } catch (error) {
      throw new Error(`Failed to create backup of ${fileId}: ${error.message}`);
    }
  }

  /**
   * List PPTX files in a folder
   * @param {string} folderId - Folder ID (optional, defaults to root)
   * @returns {Array} Array of file objects
   */
  listPPTXFiles(folderId) {
    try {
      const folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
      const files = folder.getFilesByType(this.mimeTypes.PPTX);
      const fileList = [];
      
      while (files.hasNext()) {
        const file = files.next();
        fileList.push({
          id: file.getId(),
          name: file.getName(),
          size: file.getSize(),
          lastModified: file.getLastUpdated()
        });
      }
      
      return fileList;
    } catch (error) {
      throw new Error(`Failed to list files: ${error.message}`);
    }
  }

  // Private methods
  _convertSlidesToPPTX(slidesId) {
    try {
      // Export Google Slides as PPTX
      const url = `https://docs.google.com/presentation/d/${slidesId}/export/pptx`;
      const response = UrlFetchApp.fetch(url, {
        headers: {
          'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
        }
      });
      
      if (response.getResponseCode() === 200) {
        return response.getBlob();
      } else {
        throw new Error(`Export failed with status: ${response.getResponseCode()}`);
      }
    } catch (error) {
      throw new Error(`Failed to convert Slides to PPTX: ${error.message}`);
    }
  }

  /**
   * Validate file access and permissions
   * @param {string} fileId - File ID to validate
   * @returns {boolean} True if accessible
   */
  validateAccess(fileId) {
    try {
      const file = DriveApp.getFileById(fileId);
      // Try to read basic properties
      file.getName();
      file.getSize();
      return true;
    } catch (error) {
      return false;
    }
  }
}