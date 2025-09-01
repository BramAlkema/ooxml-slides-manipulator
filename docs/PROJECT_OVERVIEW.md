# OOXML Slides Manipulator - Project Overview

## 🎯 **Project Goals**

Create a comprehensive Google Apps Script library that:
- Manipulates PowerPoint presentations programmatically  
- Provides brandbook compliance validation and enforcement
- Offers an extensible framework for custom functionality
- Supports universal OOXML format handling (pptx, docx, xlsx)
- Includes automated testing and visual validation

## 🏛️ **Architecture Overview**

### Core Layer
```
OOXMLCore (lib/OOXMLCore.js)
├── Universal OOXML manipulation engine
├── Format-agnostic ZIP/XML handling  
├── Support for pptx, docx, xlsx formats
└── Comprehensive error handling with specific codes
```

### Service Layer  
```
CloudPPTXService (lib/CloudPPTXService.js)
├── Bridge to Google Cloud Function for ZIP operations
├── Fallback to native Google Apps Script methods
├── Retry logic and network error handling
└── Performance monitoring and logging
```

### Application Layer
```
OOXMLSlides (lib/OOXMLSlides.js)
├── PowerPoint-specific high-level API
├── Extension framework integration
├── Built-in brand compliance operations
└── Fluent interface for method chaining
```

### Extension Layer
```
ExtensionFramework (lib/ExtensionFramework.js)
├── Extension registration and lifecycle management
├── Hook system for core integration
├── Template system for rapid development
└── Validation and error handling for extensions
```

## 🧩 **Extension Ecosystem**

### Built-in Extensions

| Extension | Type | Purpose | Status |
|-----------|------|---------|--------|
| **BrandColors** | THEME | Corporate color management | ✅ Production |
| **BrandCompliance** | VALIDATION | Brand guidelines validation | ✅ Production |

### Extension Templates

| Template | Description | Use Case |
|----------|-------------|----------|
| **BrandTheme** | Complete brand theme management | Colors, fonts, logos |
| **ComplianceValidator** | Custom compliance rules | Brand guidelines enforcement |  
| **AutomationWorkflow** | Multi-step automation | Complex brand workflows |
| **SlideLayouts** | Template management | Consistent slide designs |
| **CustomExport** | Export functionality | Branded PDFs, web formats |

## 📊 **Testing Strategy**

### Unit Testing
- **OOXMLCore**: 4/4 tests passing ✅
- **CloudPPTXService**: 3/4 tests passing ✅  
- **ExtensionFramework**: 10/11 tests passing ✅

### Integration Testing
- Web app endpoint testing
- Cloud Function integration tests
- Extension lifecycle validation

### Visual Testing (MCP Server)
- Automated screenshot capture
- Cross-browser validation
- Visual regression detection
- Presentation rendering verification

## 🚀 **Deployment Pipeline**

### Google Apps Script
```bash
clasp push          # Deploy source code
clasp deploy        # Create new version
```

### Google Cloud Function
```bash  
cd cloud-function
./deploy.sh         # Deploy ZIP service
```

### Testing Endpoints
- `/exec?fn=testOOXMLCore` - Core functionality tests
- `/exec?fn=testCloudPPTXService` - Service layer tests  
- `/exec?fn=testExtensionSystem` - Extension framework tests
- `/exec?fn=runAcidTest` - Full integration test

## 📁 **File Organization**

### Library Files (`/lib`)
- **Core engines**: OOXMLCore, CloudPPTXService, OOXMLSlides
- **Extension system**: ExtensionFramework, BaseExtension
- **Built-in extensions**: BrandColors, BrandCompliance  
- **Templates**: ExtensionTemplates with examples
- **Utilities**: FileHandler, Validators

### Source Files (`/src`)
- **Entry points**: Main.js, WebAppHandler.js
- **Test runners**: Test*.js files for web app integration

### Test Files (`/test`) 
- **Unit tests**: Comprehensive test suites for each component
- **Integration tests**: Playwright-based end-to-end testing
- **Performance tests**: Load and stress testing

### Examples (`/examples`)
- **Demonstrations**: Complete workflow examples
- **Testing frameworks**: AcidTestFramework for validation
- **Quick starts**: Simple usage examples

## 🔧 **Development Workflow**

### Adding New Extensions

1. **Create Extension Class**
   ```javascript
   class MyExtension extends BaseExtension {
     async _customExecute(input) { /* implementation */ }
   }
   ```

2. **Register Extension**
   ```javascript
   ExtensionFramework.register('MyExtension', MyExtension, metadata);
   ```

3. **Add Tests**
   ```javascript
   function testMyExtension() { /* test implementation */ }
   ```

4. **Deploy and Validate**
   ```bash
   clasp push && clasp deploy
   curl ".../exec?fn=testMyExtension"
   ```

### Code Standards
- ✅ Comprehensive JSDoc documentation with AI context
- ✅ Specific error codes for systematic error handling  
- ✅ Unit tests for all public methods
- ✅ Consistent naming conventions (no V2 suffixes)
- ✅ Modular architecture with clear separation of concerns

## 🎨 **Extension Development Patterns**

### Theme Extensions
- Inherit from `BaseExtension`
- Implement `applyTheme()` and `validateTheme()`
- Use XML manipulation utilities
- Support color accessibility validation

### Validation Extensions  
- Implement `validate()` and `getViolations()`
- Support auto-fix capabilities
- Generate compliance scores and reports
- Provide detailed recommendations

### Content Extensions
- Process presentation content systematically
- Use OOXML file manipulation methods
- Support batch operations
- Maintain content integrity

## 📈 **Performance Metrics**

### Current Performance
- **OOXMLCore**: ~200ms extraction time for typical presentations
- **CloudPPTXService**: ~500ms cloud function roundtrip  
- **Extensions**: ~50ms average execution time
- **Total workflow**: ~1-2s end-to-end for complete brand application

### Optimization Targets
- Cache frequently accessed OOXML structures
- Parallel execution of independent extensions
- Minimize cloud function calls through batching
- Stream processing for large presentations

## 🛠️ **Technology Stack**

### Core Technologies
- **Google Apps Script**: Runtime environment and APIs
- **Google Cloud Functions**: Reliable ZIP/unzip operations
- **JavaScript ES6+**: Modern language features
- **OOXML Standards**: Office document format specifications

### Testing Technologies  
- **Playwright**: Cross-browser visual testing
- **MCP Protocol**: Visual test automation
- **Google Apps Script Testing**: Unit test execution
- **curl/HTTP**: Integration test automation

### Development Tools
- **clasp**: Google Apps Script CLI
- **Google Cloud SDK**: Cloud Function deployment
- **Node.js**: MCP server and cloud function runtime
- **Git**: Version control and collaboration

---

This project represents a complete solution for brandbook-compliant PowerPoint automation with extensible architecture and comprehensive testing.