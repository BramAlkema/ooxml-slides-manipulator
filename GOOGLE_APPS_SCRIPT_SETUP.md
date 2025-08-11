# Google Apps Script Setup Guide

## ✅ Completed Setup

1. **Clasp Authentication** - Successfully logged in as info@bramalkema.nl
2. **Project Created** - New Google Apps Script project created
3. **Files Uploaded** - All 10 library files pushed to Google Apps Script
4. **OAuth Scopes Configured** - Required permissions set in appsscript.json

## 🔗 Project Information

- **Project URL**: https://script.google.com/d/1feN12V9_9EgBR6lHIh1FcCRXJlT-w-uFEe3NYdN_AuaUiiWk0Ov8jICB/edit
- **Project Title**: OOXML Slides Manipulator
- **Type**: Standalone Google Apps Script

## 📋 Next Steps in Google Apps Script Web IDE

### 1. Enable Drive API

1. Open the project in Google Apps Script web IDE
2. Click **Resources** > **Advanced Google Services**
3. Enable **Google Drive API v3**
4. Also enable it in **Google Cloud Console** (link provided in the dialog)

### 2. Authorize OAuth Scopes

The following scopes are configured and will be requested on first run:
- `https://www.googleapis.com/auth/drive`
- `https://www.googleapis.com/auth/drive.file`
- `https://www.googleapis.com/auth/presentations`
- `https://www.googleapis.com/auth/script.external_request`

### 3. Run Quick Tests

1. In the Google Apps Script editor, select **QuickTest.js**
2. Choose the function: **`runQuickTests`**
3. Click **Run** (you'll need to authorize permissions first)
4. Check the **Execution transcript** for results

### 4. Test with Real Files

1. Upload a PowerPoint (.pptx) file to Google Drive
2. Get the file ID from the URL or sharing settings
3. Update `TEST_CONFIG` in **basic-test.js** with your file ID
4. Run **`runAllTests()`** for comprehensive testing

## 🎯 Quick Test Commands

```javascript
// Test library loading
runQuickTests()

// Check setup requirements
setupCheck()

// Test with a real file (replace with your file ID)
const slides = OOXMLSlides.fromFile('your-file-id-here');
const theme = slides.getTheme();
console.log(theme);
```

## 🧪 Example Usage

```javascript
// Basic theme modification
function modifyPresentationTheme() {
  const fileId = 'your-google-drive-file-id';
  
  const slides = OOXMLSlides.fromFile(fileId);
  
  // Apply corporate theme
  slides
    .setColors(['#2C3E50', '#FFFFFF', '#34495E', '#ECF0F1', '#3498DB', '#E74C3C'])
    .setFonts('Arial', 'Calibri')
    .setSize(1920, 1080);
  
  // Save as new Google Slides file
  const newFileId = slides.save({
    name: 'Modified Presentation',
    importToSlides: true
  });
  
  console.log('Modified presentation saved:', newFileId);
  return newFileId;
}
```

## 🔧 Troubleshooting

### Common Issues

1. **"DriveApp is not defined"**
   - Enable Drive API in Advanced Google Services
   - Also enable in Google Cloud Console

2. **"Insufficient permissions"**
   - Re-run the function to trigger OAuth authorization
   - Check that all required scopes are approved

3. **"File not found"**
   - Verify file ID is correct
   - Ensure you have access to the file
   - File must be a .pptx format

4. **"Invalid OOXML structure"**
   - File may be corrupted
   - Try with a different PowerPoint file
   - Ensure file is actually a .pptx (not .ppt)

### Debug Commands

```javascript
// Check if all classes are loaded
testLibraryLoading()

// Validate file ID format
Validators.isValidFileId('your-file-id')

// Test file access
const fileHandler = new FileHandler();
fileHandler.validateAccess('your-file-id')
```

## 📚 Available Functions

### Core Classes
- `OOXMLSlides` - Main API class
- `OOXMLParser` - ZIP/XML manipulation
- `ThemeEditor` - Color and font editing
- `SlideManager` - Master and layout control
- `FileHandler` - Google Drive integration
- `Validators` - Input validation
- `ErrorHandler` - Error management

### Test Functions
- `runQuickTests()` - Basic functionality tests
- `runAllTests()` - Comprehensive test suite
- `setupCheck()` - Setup verification

### Example Functions
- See `usage-examples.js` for 10 complete examples
- See `basic-test.js` for testing patterns

## 🎉 Success Indicators

When everything is working correctly, you should see:
- ✅ All library classes loaded successfully
- ✅ Drive API access confirmed
- ✅ File operations working
- ✅ Theme modifications applied
- ✅ OOXML parsing successful

## 🚀 Ready to Use!

Once the setup is complete, you can:
1. Edit PowerPoint color palettes not exposed in APIs
2. Change font pairs for entire presentations
3. Modify slide masters and layouts
4. Automate complex presentation transformations
5. Create custom theme presets

The library unlocks **undocumented PowerPoint features** through direct OOXML manipulation!