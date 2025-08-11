/**
 * OOXMLSlides Usage Examples
 * Demonstrates various ways to use the library
 */

// Example 1: Basic theme modification
function basicThemeExample() {
  const fileId = 'your-google-drive-file-id';
  
  try {
    const slides = OOXMLSlides.fromFile(fileId);
    
    // Modify colors and fonts
    slides
      .setColors(['#2C3E50', '#FFFFFF', '#34495E', '#ECF0F1', '#3498DB', '#E74C3C'])
      .setFonts('Segoe UI', 'Segoe UI');
    
    // Save as new file
    const newFileId = slides.save({ 
      name: 'Modified Presentation',
      importToSlides: true 
    });
    
    console.log(`Modified presentation saved: ${newFileId}`);
    return newFileId;
    
  } catch (error) {
    console.error('Error:', ErrorHandler.getUserMessage(error));
    throw error;
  }
}

// Example 2: Using predefined themes
function predefinedThemeExample() {
  const fileId = 'your-file-id';
  
  const slides = OOXMLSlides.fromFile(fileId);
  
  // Apply corporate theme
  OOXMLSlides.Presets.corporateTheme(slides);
  
  // Or use method chaining
  slides.setColorTheme('colorful');
  
  return slides.save({ name: 'Corporate Theme Presentation' });
}

// Example 3: Advanced color palette customization
function advancedColorExample() {
  const fileId = 'your-file-id';
  
  const slides = OOXMLSlides.fromFile(fileId);
  
  // Custom color palette
  const brandColors = {
    dk1: '#1A1A1A',      // Dark text
    lt1: '#FFFFFF',      // Light text  
    dk2: '#666666',      // Dark accent
    lt2: '#F5F5F5',      // Light accent
    accent1: '#FF6B35',  // Brand primary
    accent2: '#004E7C',  // Brand secondary
    accent3: '#8B9DC3',  // Supporting color
    accent4: '#F7931E',  // Highlight color
    accent5: '#C5C6C7',  // Neutral
    accent6: '#0F3460'   // Dark blue
  };
  
  slides.setColors(brandColors);
  
  return slides.save({ name: 'Brand Colors Applied' });
}

// Example 4: Slide size and master modification
function slideMasterExample() {
  const fileId = 'your-file-id';
  
  const slides = OOXMLSlides.fromFile(fileId);
  
  // Change slide size to widescreen
  slides.setSize(1920, 1080);
  
  // Modify slide master
  slides.modifyMaster('1', {
    background: { color: '#F8F9FA' },
    textStyles: {
      title: { fontSize: 44 },
      body: { fontSize: 24 }
    }
  });
  
  return slides.save({ name: 'Widescreen with Custom Master' });
}

// Example 5: Batch operations
function batchOperationsExample() {
  const fileId = 'your-file-id';
  
  return OOXMLSlides
    .fromFile(fileId)
    .batch(slides => {
      // Get current theme info
      const currentTheme = slides.getTheme();
      console.log('Current theme:', currentTheme);
      
      // Apply multiple modifications
      slides
        .setColorTheme('dark')
        .setFonts('Arial', 'Arial')
        .setSize(1920, 1080);
      
      // Modify layouts
      slides.modifyLayout('1', {
        placeholders: [{
          type: 'title',
          position: { x: 100, y: 50 },
          size: { width: 800, height: 100 }
        }]
      });
    })
    .save({ 
      name: 'Batch Modified Presentation',
      importToSlides: true 
    });
}

// Example 6: Error handling and validation
function robustExample() {
  const fileId = 'your-file-id';
  
  try {
    // Validate file ID format
    if (!Validators.isValidFileId(fileId)) {
      throw new Error('Invalid file ID format');
    }
    
    const slides = OOXMLSlides.fromFile(fileId, {
      createBackup: true  // Create backup before modification
    });
    
    // Validate colors before applying
    const colors = ['#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF'];
    colors.forEach(color => {
      if (!Validators.isValidHexColor(color)) {
        throw new Error(`Invalid color: ${color}`);
      }
    });
    
    slides.setColors(colors);
    
    return slides.save({ name: 'Validated Modification' });
    
  } catch (error) {
    console.error('Robust example error:', ErrorHandler.getUserMessage(error));
    return null;
  }
}

// Example 7: File inspection and debugging
function inspectionExample() {
  const fileId = 'your-file-id';
  
  const slides = OOXMLSlides.fromFile(fileId);
  
  // Get file information
  const info = slides.getFileInfo();
  console.log('File info:', info);
  
  // List all files in OOXML package
  const files = slides.listFiles();
  console.log('OOXML files:', files);
  
  // Export theme XML for inspection
  const themeXML = slides.exportXML('ppt/theme/theme1.xml');
  console.log('Theme XML:', themeXML);
  
  // Get current theme data
  const theme = slides.getTheme();
  console.log('Current theme:', theme);
  
  return { info, files, theme };
}

// Example 8: Working with multiple files
function multipleFilesExample() {
  const fileIds = ['file-id-1', 'file-id-2', 'file-id-3'];
  const results = [];
  
  fileIds.forEach(fileId => {
    try {
      const slides = OOXMLSlides.fromFile(fileId);
      
      // Apply consistent branding
      OOXMLSlides.Presets.corporateTheme(slides);
      
      const newFileId = slides.save({ 
        name: `Branded_${slides.getFileInfo().name}`,
        importToSlides: true 
      });
      
      results.push({ original: fileId, modified: newFileId, status: 'success' });
      
    } catch (error) {
      results.push({ original: fileId, error: error.message, status: 'error' });
    }
  });
  
  return results;
}

// Example 9: Custom theme creation
function createCustomTheme() {
  // Define a custom theme preset
  OOXMLSlides.Presets.techTheme = (slides) => {
    return slides
      .setColors({
        dk1: '#0D1B2A',      // Deep navy
        lt1: '#FFFFFF',      // White
        dk2: '#415A77',      // Steel blue
        lt2: '#E0E1DD',      // Light gray
        accent1: '#778DA9',  // Blue gray
        accent2: '#1B263B',  // Dark blue
        accent3: '#F25C54',  // Coral red
        accent4: '#F2CC8F',  // Golden
        accent5: '#81B29A',  // Sage green
        accent6: '#3D5A80'   // Medium blue
      })
      .setFonts('Roboto', 'Open Sans')
      .setSize(1920, 1080);
  };
  
  // Use the custom theme
  const fileId = 'your-file-id';
  const slides = OOXMLSlides.fromFile(fileId);
  
  OOXMLSlides.Presets.techTheme(slides);
  
  return slides.save({ name: 'Tech Theme Applied' });
}

// Example 10: Google Apps Script integration
function gasIntegrationExample() {
  /**
   * This function can be called from Google Apps Script
   * Set up a trigger or call manually
   */
  function modifyPresentationFromTrigger() {
    const fileId = PropertiesService.getScriptProperties().getProperty('PRESENTATION_FILE_ID');
    
    if (!fileId) {
      throw new Error('No file ID configured in script properties');
    }
    
    try {
      const slides = OOXMLSlides.fromFile(fileId, {
        createBackup: true,
        autoSave: true
      });
      
      // Apply modifications based on current date/context
      const now = new Date();
      const isWorkWeek = now.getDay() >= 1 && now.getDay() <= 5;
      
      if (isWorkWeek) {
        OOXMLSlides.Presets.corporateTheme(slides);
      } else {
        OOXMLSlides.Presets.creativeTheme(slides);
      }
      
      const resultId = slides.save({
        name: `Auto-Modified_${Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd')}`,
        importToSlides: true
      });
      
      // Send notification email
      MailApp.sendEmail({
        to: Session.getActiveUser().getEmail(),
        subject: 'Presentation Modified',
        body: `Your presentation has been automatically modified. New file: ${resultId}`
      });
      
      return resultId;
      
    } catch (error) {
      console.error('Auto-modification failed:', error);
      
      // Send error notification
      MailApp.sendEmail({
        to: Session.getActiveUser().getEmail(),
        subject: 'Presentation Modification Failed',
        body: `Auto-modification failed: ${error.message}`
      });
      
      throw error;
    }
  }
}