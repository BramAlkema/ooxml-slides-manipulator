# OOXML Slides Manipulator - Project Index

## Project Architecture

### Core Universal OOXML Engine
- **OOXMLCore.js** - Universal OOXML manipulation engine
  - `class OOXMLCore` - Format-agnostic ZIP/XML handling
  - Supports: PowerPoint (.pptx), Word (.docx), Excel (.xlsx)
  - Key methods: `extract()`, `compress()`, `getFile()`, `setFile()`

### PowerPoint-Specific Layers
- **OOXMLSlides.js** - Original PowerPoint manipulator
  - `class OOXMLSlides` - Legacy implementation with tanaikech-style features
- **OOXMLSlides_Clean.js** - Clean PowerPoint layer
  - `class OOXMLSlidesV2` - Built on OOXMLCore, clean architecture

### Core Editors & Processors
- **ThemeEditor.js** - `class ThemeEditor` - Theme manipulation
- **CustomColorEditor.js** - `class CustomColorEditor` - Color scheme management
- **TableStyleEditor.js** - `class TableStyleEditor` - Table formatting
- **TypographyEditor.js** - `class TypographyEditor` - Font and text styling
- **NumberingStyleEditor.js** - `class NumberingStyleEditor` - List formatting
- **SuperThemeEditor.js** - `class SuperThemeEditor` - Advanced theme management
- **XMLSearchReplaceEditor.js** - `class XMLSearchReplaceEditor` - Bulk text operations
- **SlideManager.js** - `class SlideManager` - Slide operations
- **LanguageStandardizationEditor.js** - `class LanguageStandardizationEditor` - Language normalization

### Utilities & Infrastructure
- **FileHandler.js** - `class FileHandler` - Google Drive integration
- **Validators.js** - `class Validators`, `class ErrorHandler` - Validation & error handling
- **CloudPPTXService.js** - `class CloudPPTXService` - Cloud Function integration
- **OOXMLParser.js** - `class OOXMLParser` - ZIP/XML parsing utilities

### Templates & Resources
- **PPTXTemplate.js** - `class PPTXTemplate` - Template management
- **MinimalPPTXTemplate.js** - `class MinimalPPTXTemplate` - Minimal PPTX creation

### Testing & Validation
- **TestOOXMLCore.js** - `function testOOXMLCore()` - Universal core testing
- **AcidTestFramework.js** - `function runAcidTest()` - Comprehensive validation
- **CompleteRoundtripDemo.js** - `function completeOOXMLRoundtripDemo()` - Full workflow test
- **QuickTest.js** - `function quickOOXMLTest()` - Bypass test for MIME issues

### Web Integration
- **WebAppHandler.js** - HTTP request handling
  - `function doPost()`, `function doGet()` - API endpoints
- **Main.js** - Main entry points
  - `function doPost()`, `function doGet()`, `function ping()`

### Advanced Features
- **SlidesAppAdvanced.js** - `class SlidesAppAdvanced` - Advanced Google Slides API
- **SlidesAPIExplorer.js** - `class SlidesAPIExplorer` - API analysis
- **AutoSetup.js** - `function autoSetup()` - Environment validation
- **PublishPresentations.js** - `function publishAcidTestPresentations()` - Publishing

## Working Solutions to Known Problems

### Problem 1: MIME Type Recognition
**Issue**: Google Apps Script FileHandler doesn't recognize PPTX MIME types correctly
**Error**: `Unsupported file type: application/vnd.openxmlformats-officedocument.presentationml.presentation`

**Solution**: Fixed in FileHandler.js:
```javascript
// Wrong MIME type (original)
PPTX: 'application/vnd.openxmlformats-presentationml.presentation'

// Correct MIME type
PPTX: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'

// Added support for all PowerPoint variants
supportedPPTXTypes: [
  'application/vnd.openxmlformats-officedocument.presentationml.presentation', // .pptx
  'application/vnd.openxmlformats-officedocument.presentationml.template',     // .potx
  'application/vnd.openxmlformats-officedocument.presentationml.slideshow',    // .ppsx
  'application/vnd.ms-powerpoint.presentation.macroEnabled.12',               // .pptm
  // ... etc
]
```

### Problem 2: Domain Restrictions (bramalkema.nl workspace)
**Issue**: Google Workspace domain restrictions block external access and clasp operations
**Error**: 403 Forbidden, Chinese/Dutch error pages

**Solutions**:
1. **Workspace Admin Setting**: Set "Sharing outside of Bram Alkema" to "ON"
2. **Manual Code Updates**: When clasp push/deploy fails, manually copy code in GAS editor
3. **Direct Function Testing**: Use GAS editor's Run button instead of external API calls
4. **Web App Alternative**: Deploy as web app with "ANYONE" access when domain allows

### Problem 3: Naming Conflicts
**Issue**: Multiple class definitions with same name
**Error**: `SyntaxError: Identifier 'OOXMLSlides' has already been declared`

**Solution**: Renamed clean implementation:
- `OOXMLSlides_Clean.js` → class `OOXMLSlidesV2`
- Updated all references in test files

### Problem 4: Cloud Function Integration
**Issue**: Reliable ZIP extraction and recompression for OOXML files

**Solution**: CloudPPTXService.js provides:
- `extractPPTX(blob)` - Extract ZIP to files object
- `compressPPTX(files)` - Recompress files to ZIP blob
- Deployed Cloud Function at: `https://us-central1-sys-06010664847399265559325060.cloudfunctions.net/pptxRouter`

### Problem 5: Testing Access Issues
**Issue**: Cannot run automated tests due to domain restrictions

**Solutions**:
1. **Manual GAS Testing**: Open script editor, run functions directly
2. **API Executable**: Use clasp run (when working)
3. **Bypass Tests**: Create QuickTest.js that works with raw blobs
4. **Web App Endpoints**: HTTP POST to deployed web app

## Deployment Patterns

### Successful Deployment Methods
1. **clasp deploy** - Works even when clasp push fails
2. **Manual Copy-Paste** - Reliable fallback for blocked operations
3. **GAS Editor Direct** - Always works for testing

### File Push Order (optimized)
```
appsscript.json (config first)
Main.js (entry points)
core/OOXMLCore.js (universal engine)
core/* (all core modules)
utils/* (utilities)
OOXMLSlides*.js (PowerPoint layers)
*Test*.js (testing files)
```

## Testing Strategy

### Test Hierarchy
1. **Unit Tests**: Individual class methods
2. **Integration Tests**: OOXMLCore + specific editors
3. **End-to-End Tests**: Full import → modify → export workflow
4. **Acid Tests**: Comprehensive validation framework

### Current Test Status
- ✅ Universal OOXML Core architecture complete
- ⚠️ MIME type fix needed (manual deployment)
- 🚧 Domain restrictions limit automated testing
- ✅ Manual testing via GAS editor works

## Architecture Decisions

### Universal Core Design
- **Separation of Concerns**: Format-agnostic core + format-specific layers
- **Future-Ready**: Easy extension to Word (.docx) and Excel (.xlsx)
- **Clean Dependencies**: Core has no PowerPoint-specific code

### Error Handling Strategy
- **Graceful Degradation**: Fallback methods when Cloud Function unavailable  
- **Detailed Logging**: Console output for debugging in GAS environment
- **Recovery Patterns**: Backup creation before modifications

## Known Working Configurations

### Apps Script Configuration
```json
{
  "executionApi": { "access": "MYSELF" },
  "webapp": { "access": "ANYONE", "executeAs": "USER_DEPLOYING" },
  "oauthScopes": [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/script.external_request"
  ]
}
```

### Cloud Function Setup
- **Platform**: Google Cloud Functions
- **Runtime**: Node.js
- **Endpoint**: HTTP trigger with CORS enabled
- **Libraries**: JSZip, xml2js for OOXML processing

## Next Steps & Maintenance

### Immediate Actions Needed
1. Deploy MIME type fix to FileHandler.js
2. Test universal OOXML core with fixed MIME types
3. Document successful end-to-end workflow

### Future Enhancements
1. Extend OOXMLCore to support Word (.docx)
2. Add Excel (.xlsx) support
3. Implement batch processing capabilities
4. Add more sophisticated theme management

### Monitoring Points
- Domain restriction changes
- Google Apps Script API updates  
- Cloud Function availability
- MIME type handling changes