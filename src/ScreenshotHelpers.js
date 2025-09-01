/**
 * ScreenshotHelpers - Authenticated screenshot and comparison functions
 * Separated from Main.js to keep code organized
 */

/**
 * Test font parsing functionality without creating presentations
 */
function testFontParsingSync(prompt) {
  try {
    console.log('🧪 Testing font parsing functionality:', prompt);
    
    // Test the parsing functions directly
    var fontPair = parseFontPairFromPrompt(prompt);
    var colorPalette = parseColorPaletteFromPrompt(prompt);
    
    return {
      success: true,
      message: 'Font parsing test completed',
      fontPair: fontPair,
      colorPalette: colorPalette,
      prompt: prompt,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('❌ Font parsing test failed:', error);
    return {
      success: false,
      error: error.message,
      stack: error.stack,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * Get slide as image using authenticated Apps Script environment
 */
function getSlideAsImageSync(presentationId, slideIndex) {
  try {
    console.log('📸 Getting slide as image:', presentationId, 'slide:', slideIndex || 0);
    
    // Open the presentation using Apps Script (authenticated)
    var presentation = SlidesApp.openById(presentationId);
    var slides = presentation.getSlides();
    
    if (slides.length === 0) {
      return {
        success: false,
        error: 'No slides found in presentation',
        timestamp: new Date().toISOString()
      };
    }
    
    var slideIndex = slideIndex || 0;
    if (slideIndex >= slides.length) {
      slideIndex = 0;
    }
    
    var slide = slides[slideIndex];
    
    // Get slide thumbnail using Drive API (authenticated through Apps Script)
    try {
      // Use Drive API to get thumbnail
      var file = DriveApp.getFileById(presentationId);
      var thumbnail = Drive.Files.get(presentationId, {fields: 'thumbnailLink'});
      
      return {
        success: true,
        presentationId: presentationId,
        slideIndex: slideIndex,
        totalSlides: slides.length,
        thumbnailLink: thumbnail.thumbnailLink,
        presentationTitle: presentation.getName(),
        slideId: slide.getObjectId(),
        timestamp: new Date().toISOString()
      };
      
    } catch (driveError) {
      return {
        success: false,
        error: 'Could not get slide thumbnail: ' + driveError.message,
        presentationId: presentationId,
        timestamp: new Date().toISOString()
      };
    }
    
  } catch (error) {
    console.error('❌ Get slide as image failed:', error);
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * Create font comparison presentations and return both IDs with slide info
 */
function createFontComparisonSync() {
  try {
    console.log('🎨 Creating font comparison presentations...');
    
    // Create BEFORE presentation (default fonts)
    var beforeTitle = 'BEFORE: Default Fonts - ' + new Date().toISOString();
    var beforePresentation = SlidesApp.create(beforeTitle);
    var beforeId = beforePresentation.getId();
    
    // Create AFTER presentation (Merriweather/Inter fonts)
    var afterPrompt = 'Create presentation with https://coolors.co/edd3c4-c8adc0-7765e3-3b60e4-080708 and Merriweather/Inter fonts';
    var afterFontPair = parseFontPairFromPrompt(afterPrompt);
    var afterColorPalette = parseColorPaletteFromPrompt(afterPrompt);
    
    var afterTitle = 'AFTER: Merriweather/Inter Fonts - ' + new Date().toISOString();
    var afterPresentation = SlidesApp.create(afterTitle);
    var afterId = afterPresentation.getId();
    
    // Add content to AFTER presentation with font styling
    var afterSlides = afterPresentation.getSlides();
    if (afterSlides.length > 0) {
      var firstSlide = afterSlides[0];
      
      // Clear default content
      var shapes = firstSlide.getShapes();
      shapes.forEach(function(shape) {
        shape.remove();
      });
      
      // Add font comparison content
      var titleShape = firstSlide.insertTextBox('Font Pair Demonstration', 50, 50, 600, 80);
      var titleRange = titleShape.getText();
      titleRange.getTextStyle()
        .setFontFamily(afterFontPair.heading)
        .setFontSize(36)
        .setBold(true);
      
      if (afterColorPalette.primary) {
        titleRange.getTextStyle().setForegroundColor(afterColorPalette.primary);
      }
      
      // Add comparison text
      var comparisonText = `AFTER: Using ${afterFontPair.heading} for headings and ${afterFontPair.body} for body text.\n\nThis demonstrates the font pair change system working - parsing font names from prompts and applying them to presentation content.`;
      var bodyShape = firstSlide.insertTextBox(comparisonText, 50, 150, 600, 200);
      var bodyRange = bodyShape.getText();
      bodyRange.getTextStyle()
        .setFontFamily(afterFontPair.body)
        .setFontSize(16);
      
      if (afterColorPalette.accent) {
        bodyRange.getTextStyle().setForegroundColor(afterColorPalette.accent);
      }
    }
    
    // Get slide thumbnails for both
    var beforeThumbnail = null;
    var afterThumbnail = null;
    
    try {
      var beforeFile = Drive.Files.get(beforeId, {fields: 'thumbnailLink'});
      beforeThumbnail = beforeFile.thumbnailLink;
    } catch (thumbError) {
      console.log('⚠️ Could not get BEFORE thumbnail:', thumbError.message);
    }
    
    try {
      var afterFile = Drive.Files.get(afterId, {fields: 'thumbnailLink'});
      afterThumbnail = afterFile.thumbnailLink;
    } catch (thumbError) {
      console.log('⚠️ Could not get AFTER thumbnail:', thumbError.message);
    }
    
    return {
      success: true,
      message: 'Font comparison presentations created',
      before: {
        presentationId: beforeId,
        title: beforeTitle,
        editUrl: `https://docs.google.com/presentation/d/${beforeId}/edit`,
        thumbnailLink: beforeThumbnail,
        fontInfo: 'Default Google Slides fonts'
      },
      after: {
        presentationId: afterId,
        title: afterTitle,
        editUrl: `https://docs.google.com/presentation/d/${afterId}/edit`,
        thumbnailLink: afterThumbnail,
        fontPair: afterFontPair,
        colorPalette: afterColorPalette,
        fontInfo: `${afterFontPair.heading} (headings) + ${afterFontPair.body} (body)`
      },
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('❌ Font comparison creation failed:', error);
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}