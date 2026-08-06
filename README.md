# OOXML Slides PPTX Processing Platform

A powerful Google Apps Script platform for manipulating PowerPoint presentations through direct OOXML editing. Process PPTX files in the cloud without leaving Google Workspace - unzip, modify, and zip PPTX files with one-click deployment to Google Cloud free tier.

## 🎯 **Core Capabilities**

- **📦 Unzip PPTX Files**: Extract OOXML structure from PowerPoint files using fflate
- **✏️ Modify OOXML Content**: Edit themes, colors, fonts, slides, and metadata as JSON/XML
- **🔨 Zip Back to PPTX**: Reconstruct valid PowerPoint files after modifications
- **☁️ Deploy from GAS**: One-click deployment to US free tier directly from Google Apps Script
- **🚀 Cloud Processing**: Process PPTX files on Google Cloud Run (512MB free tier)
- **🔄 Server-Side Operations**: Batch operations like text replacement without downloading
- **📋 JSON Manifest System**: Work with OOXML as editable JSON structures
- **🎨 Theme Management**: Apply brand colors, fonts, and compliance rules
- **💰 Free Tier Optimized**: Runs on Google Cloud free tier (us-central1)
- **🧪 Layered Testing**: deterministic local OOXML tests in CI plus opt-in Playwright deployment checks

## 🏗️ **Project Structure**

```
├── lib/              # Core libraries and extensions
│   ├── OOXMLJsonService.js    # OOXML JSON Service integration
│   ├── OOXMLExtensionAdapter.js # Extension-to-JSON service bridge
│   ├── OOXMLSlides.js         # PowerPoint-specific API with JSON backend
│   ├── ExtensionFramework.js  # Extension registration and lifecycle
│   ├── BaseExtension.js       # Base class for all extensions
│   ├── BrandColorsExtension.js    # Brand color management
│   ├── BrandComplianceExtension.js # Compliance validation
│   ├── ExtensionTemplates.js  # Templates for creating extensions
│   ├── OOXMLCore.js           # Legacy OOXML engine (fallback)
│   └── CloudPPTXService.js    # Legacy Cloud Function service (fallback)
│
├── src/              # Source files and handlers  
│   ├── Main.js                # Main entry point
│   ├── WebAppHandler.js       # HTTP request handler
│   ├── OOXMLDeployment.js     # One-click Cloud Run deployment
│   └── Test*.js              # Test runners
│
├── test/             # Unit tests
│   ├── TestOOXMLCore.js           # OOXMLCore test suite
│   ├── TestCloudPPTXService.js    # CloudPPTXService tests
│   └── *.spec.js             # Playwright tests
│
├── examples/         # Usage examples and demos
│   ├── OOXMLJsonQuickStart.js # Complete OOXML JSON Service examples
│   ├── AcidTestFramework.js   # Comprehensive testing framework
│   ├── CompleteRoundtripDemo.js # Full demo workflow
│   └── QuickTest.js          # Quick validation tests
│
├── docs/             # Documentation
├── mcp-server/       # Visual testing server
└── cloud-function/   # Google Cloud Function for ZIP operations
```

## 🚀 **Quick Start**

### 1. Deploy from Google Apps Script (No CLI Required!)

```javascript
// Copy DeployFromGAS.js to your Google Apps Script project, then:

// Step 1: Set your GCP Project ID
setupProjectId();  // Enter your project ID when prompted

// Step 2: Run preflight checks (opens sidebar)
showPreflightChecks();  // Complete billing & API setup in sidebar

// Step 3: Deploy to US free tier (takes 3-5 minutes)
await deployToUSFreeTier();  // Deploys to us-central1 free tier

// Step 4: Test the deployment
await testDeployedService();  // Verifies everything works
```

### 2. Basic PPTX Unzip/Modify/Zip Workflow

```javascript
// STEP 1: UNZIP - Extract PPTX to editable JSON/XML
const manifest = await OOXMLJsonService.unwrap('presentation.pptx');
console.log(`Extracted ${manifest.entries.length} files from PPTX`);

// STEP 2: MODIFY - Edit the OOXML content
manifest.entries.forEach(entry => {
  if (entry.type === 'xml' && entry.path.includes('slide')) {
    // Modify slide content
    entry.text = entry.text.replace('Old Company', 'New Company');
    entry.text = entry.text.replace('#FF0000', '#0066CC'); // Change colors
  }
});

// STEP 3: ZIP - Reconstruct valid PPTX file
const outputFileId = await OOXMLJsonService.rewrap(manifest, {
  filename: 'modified-presentation.pptx'
});
console.log(`Created modified PPTX: ${outputFileId}`);
```

### 3. Server-Side Operations (Advanced)

```javascript
// Powerful server-side operations for efficiency
const operations = [
  { type: 'replaceText', find: 'ACME Corp', replace: 'DeltaQuad Inc' },
  { type: 'upsertPart', path: 'ppt/customXml/metadata.xml', text: '<metadata/>' }
];

const result = await OOXMLJsonService.process('presentation.pptx', operations);
console.log(`Processed with ${result.report.replaced} replacements`);
```

### 4. Large File Sessions (100MB+)

```javascript  
// Handle large presentations efficiently
const session = await OOXMLJsonService.createSession();
await OOXMLJsonService.uploadToSession(session.uploadUrl, 'large-presentation.pptx');

const manifest = await OOXMLJsonService.unwrap(null, { gcsIn: session.gcsIn });
// Edit manifest...
const result = await OOXMLJsonService.rewrap(manifest, { 
  gcsIn: session.gcsIn, 
  gcsOut: session.gcsOut 
});
```

### Extension Development

```javascript
// Create custom extension
class MyBrandExtension extends BaseExtension {
  async _customExecute(input) {
    // Your custom brand logic here
    return { success: true, applied: input.changes };
  }
}

// Register extension
ExtensionFramework.register('MyBrand', MyBrandExtension, {
  type: 'THEME',
  version: '1.0.0'
});

// Use extension
const result = await slides.useExtension('MyBrand', options);
```

## 🧩 **Service Architecture** 

| Component | Purpose | Technology |
|-----------|---------|------------|
| **OOXML JSON Service** | Cloud Run service for OOXML ↔ JSON conversion | Node.js, Express, fflate |
| **Server-Side Operations** | Efficient text/part manipulation on server | Cloud Run with GCS integration |
| **Session Management** | Large file handling via signed URLs | Google Cloud Storage |
| **Extension Framework** | Modular brandbook compliance system | Google Apps Script |
| **One-Click Deployment** | Automated Cloud Run deployment | GAS → Cloud Build → Cloud Run |

## 🔧 **Extension Types**

| Type | Purpose | Examples |
|------|---------|----------|
| **THEME** | Colors, fonts, logos | Brand colors, typography, logo placement |
| **VALIDATION** | Compliance checking | Brand guidelines, accessibility, content rules |
| **CONTENT** | Content manipulation | Text processing, layout changes, media handling |
| **TEMPLATE** | Slide templates | Master slides, layouts, branded templates |
| **EXPORT** | Custom exports | PDF generation, web formats, custom workflows |

## 📋 **Built-in Extensions**

### BrandColors Extension
- Apply corporate color schemes
- Generate color harmony
- Accessibility validation
- Color preview generation

### BrandCompliance Extension  
- Font compliance checking
- Color validation
- Content rule enforcement
- Auto-fix capabilities
- Detailed compliance scoring

## 🧪 **Testing**

The project includes comprehensive testing at multiple levels:

### Local CI Tests

```bash
npm test
```

This exercises PPTX compression and extraction entirely in-process, including
text and binary OOXML parts. It does not depend on a deployed Apps Script URL,
Google authentication, or the permissions of a historical test presentation.

### Live Apps Script Tests

```bash
GAS_PROJECT_URL="https://script.google.com/macros/s/.../exec" npm run test:live
```

Live deployment checks are explicit because they create presentations and rely
on the current deployment and its Google account permissions.

### Deployment Health Check
```javascript
// Check if OOXML JSON Service is deployed and healthy
const health = await OOXMLJsonService.healthCheck();
const status = OOXMLDeployment.getDeploymentStatus();
```

### Unit Tests
```bash
# Test OOXML JSON Service integration
curl "https://script.google.com/.../exec?fn=testOOXMLJsonService"

# Test server-side operations
curl "https://script.google.com/.../exec?fn=testServerSideOperations"

# Test extension framework with JSON backend
curl "https://script.google.com/.../exec?fn=testExtensionFramework"
```

### Visual Tests (MCP Server)
```bash
cd mcp-server
npm test
```

### Integration Tests
```bash
# Run complete workflow test with JSON service
curl "https://script.google.com/.../exec?fn=runCompleteWorkflowTest"
```

## 🔧 **Architecture**

### Core Components

1. **OOXMLJsonService** - Next-generation OOXML manipulation
   - JSON manifest system for XML editing
   - Server-side operations for efficiency
   - Session-based large file handling
   - One-click Cloud Run deployment

2. **OOXMLExtensionAdapter** - Extension-to-JSON bridge
   - Adapts extensions to JSON service backend
   - Server-side operation batching
   - Maintains API compatibility

3. **OOXMLDeployment** - Cloud deployment automation
   - GCP preflight checks (billing, APIs, budgets)
   - Automated Cloud Run deployment
   - Service health monitoring and cleanup

4. **ExtensionFramework** - Modular extensions
   - Extension registration and lifecycle
   - Hook system for integration
   - Template system for development

5. **OOXMLSlides** - PowerPoint-specific API
   - High-level PowerPoint operations with JSON backend
   - Extension integration with server-side ops
   - Built-in brand compliance tools

### Extension Development Pattern

```javascript
// 1. Extend BaseExtension
class CustomExtension extends BaseExtension {
  getDefaults() { return { /* defaults */ }; }
  async _customExecute(input) { /* logic */ }
}

// 2. Register with framework  
ExtensionFramework.register('Custom', CustomExtension, metadata);

// 3. Use in presentations
await slides.useExtension('Custom', options);
```

## 📚 **API Reference**

### OOXMLJsonService
- `unwrap(fileId, options)` - Convert OOXML to JSON manifest
- `rewrap(manifest, options)` - Convert JSON manifest to OOXML
- `process(fileId, operations, options)` - Server-side operations
- `createSession()` - Create session for large files
- `uploadToSession(url, file)` - Upload to session
- `healthCheck()` - Check service health

### OOXMLDeployment
- `showGcpPreflight()` - Display setup interface
- `initAndDeploy(options)` - Deploy Cloud Run service
- `getDeploymentStatus()` - Check deployment status
- `cleanup(options)` - Remove deployment resources

### OOXMLSlides (Enhanced)
- `load()` - Initialize with JSON service backend
- `applyBrandColors(colors)` - Apply brand colors
- `validateCompliance(rules)` - Check compliance
- `useExtension(name, options)` - Use custom extension
- `save()` - Save with JSON service backend

### Server-Side Operations
- `replaceText` - Text replacement across files
- `upsertPart` - Insert or update OOXML part
- `removePart` - Remove OOXML part
- `renamePart` - Rename OOXML part

## 🎨 **Extension Templates**

Get started quickly with built-in templates:

```javascript
// Get template for extension type
const template = ExtensionFramework.getTemplate('THEME');
console.log(template.template); // Full template code
console.log(template.example);  // Example implementation

// Available templates
const templates = ExtensionTemplates.getAllTemplates();
// Returns: theme, validation, template, export, workflow templates
```

## 🔒 **Security & Compliance**

- All extensions run in sandboxed environment
- Input validation and sanitization  
- Comprehensive error codes and logging
- OAuth-based authentication for Cloud Functions
- Brand compliance validation and scoring

## 🚦 **Status**

- ✅ OOXML JSON Service: Production ready with Cloud Run deployment
- ✅ Server-Side Operations: Production ready with batching and error handling
- ✅ Large File Sessions: Production ready with GCS integration
- ✅ One-Click Deployment: Production ready with preflight checks
- ✅ Extension Framework: Enhanced with JSON service backend
- ✅ Built-in Extensions: Adapted for server-side operations
- ✅ Automated Testing: Fully integrated with JSON service workflows
- ✅ Visual Testing: MCP server operational
- ✅ Cost Management: Built-in billing checks and budget controls

## 📄 **License**

MIT License - see LICENSE file for details

## 🤝 **Contributing**

1. Create extensions using the template system
2. Add comprehensive tests for all functionality
3. Follow the established patterns and conventions
4. Update documentation for new features

---

**Next-generation OOXML manipulation with Cloud Run deployment, JSON manifests, and server-side operations - made with ❤️ for enterprise PowerPoint automation**
