# OOXML Slides Manipulator

[![MIT License](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285f4?logo=google&logoColor=white)](https://script.google.com/)
[![PowerPoint](https://img.shields.io/badge/PowerPoint-B7472A?logo=microsoftpowerpoint&logoColor=white)](https://www.microsoft.com/en-us/microsoft-365/powerpoint)

A modern Google Apps Script library for advanced PowerPoint/Google Slides manipulation through direct OOXML editing.

> 🎯 **Unlock the power of undocumented PowerPoint features** - Edit font pairs, color palettes, and theme properties that aren't exposed through standard APIs.

## Features

- **Theme Management**: Edit font pairs, color palettes, and theme properties
- **Slide Master Control**: Modify layouts and master slides  
- **Advanced Shapes**: Manipulate shapes beyond standard API limits
- **Template System**: Create and modify presentation templates
- **Google Drive Integration**: Seamless file handling in Drive

## Architecture

```
OOXMLSlides/
├── core/
│   ├── OOXMLParser.js      # ZIP/XML manipulation
│   ├── ThemeEditor.js      # Theme and color manipulation  
│   ├── SlideManager.js     # Slide and layout management
│   └── ShapeManipulator.js # Advanced shape operations
├── utils/
│   ├── FileHandler.js      # Drive API integration
│   ├── Validators.js       # Input validation
│   └── Constants.js        # OOXML constants and mappings
└── OOXMLSlides.js          # Main API class
```

## Quick Start

```javascript
// Load and modify a PowerPoint file
const slides = OOXMLSlides.fromFile('your-google-drive-file-id');

// Apply theme changes
slides
  .setColors(['#2C3E50', '#FFFFFF', '#34495E', '#ECF0F1', '#3498DB', '#E74C3C'])
  .setFonts('Segoe UI', 'Segoe UI')
  .setSize(1920, 1080);

// Save as new Google Slides file
const newFileId = slides.save({
  name: 'Modified Presentation',
  importToSlides: true
});
```

## Advanced Usage

```javascript
// Use predefined themes
OOXMLSlides.Presets.corporateTheme(slides);

// Custom color palette with validation
const colors = {
  dk1: '#1A1A1A',
  lt1: '#FFFFFF', 
  accent1: '#FF6B35',
  accent2: '#004E7C'
};
slides.setColors(colors);

// Batch operations
slides.batch(s => {
  s.setColorTheme('dark')
   .modifyMaster('1', { background: { color: '#F8F9FA' } })
   .modifyLayout('1', { name: 'customLayout' });
});
```

## Installation

### Google Apps Script

1. Open [Google Apps Script](https://script.google.com/)
2. Create a new project
3. Copy the library files into your project:
   - `OOXMLSlides.js` (main library)
   - Files from `core/` directory
   - Files from `utils/` directory
4. Enable the following Google Services:
   - Drive API
   - Advanced Google services > Drive API

### Using Clasp

```bash
npm install -g @google/clasp
git clone https://github.com/username/ooxml-slides-manipulator.git
cd ooxml-slides-manipulator
clasp create --type standalone
clasp push
```

## API Reference

### OOXMLSlides Class

#### Static Methods
```javascript
// Load and initialize
OOXMLSlides.fromFile(fileId, options)
OOXMLSlides.fromBlob(blob, options)
```

#### Theme Methods
```javascript
.setColors(colors)          // Set color palette
.setColorTheme(themeName)   // Apply predefined theme
.setFonts(major, minor)     // Set font pair
.getTheme()                 // Get current theme info
```

#### Slide Methods
```javascript
.setSize(width, height)     // Set slide dimensions
.getSize()                  // Get current dimensions
.modifyMaster(id, props)    // Modify slide master
.modifyLayout(id, props)    // Modify slide layout
```

#### File Operations
```javascript
.load()                     // Load and parse file
.save(options)              // Save modifications
.getFileInfo()              // Get file metadata
.exportXML(path)            // Export raw XML
```

## Examples

See the `examples/` directory for comprehensive usage examples:
- `usage-examples.js` - Complete examples for all features

## Testing

Run the test suite in Google Apps Script:
```javascript
// Update TEST_CONFIG in test/basic-test.js with your file IDs
runAllTests();
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

## Acknowledgments

- Inspired by [tanaikech's](https://github.com/tanaikech) Google Apps Script explorations
- Built upon [BramAlkema's](https://github.com/BramAlkema) OOXML manipulation techniques
- Thanks to the OOXML specification contributors

## License

MIT License - see [LICENSE](LICENSE) file for details.