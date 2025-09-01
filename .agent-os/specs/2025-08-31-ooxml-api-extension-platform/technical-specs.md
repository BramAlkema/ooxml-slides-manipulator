# OOXML API Extension Platform - Technical Specifications

## Architecture Overview

The OOXML API Extension Platform builds upon the existing OOXML Slides project to create a comprehensive **Google Slides API++** platform. The architecture leverages proven components while introducing new capabilities for universal OOXML operations.

```
┌─────────────────────────────────────────────────────────────────┐
│                    OOXML API Extension Platform                │
├─────────────────────────────────────────────────────────────────┤
│  🌐 API Gateway Layer                                          │
│  ├─ REST API Endpoints        ├─ GraphQL Interface            │
│  ├─ Rate Limiting            ├─ Authentication/Authorization   │
│  └─ Request Validation       └─ Response Caching              │
├─────────────────────────────────────────────────────────────────┤
│  🔧 Universal Operations Engine                                │
│  ├─ Global Search/Replace     ├─ Template Management          │
│  ├─ Batch Processing         ├─ Theme Control                 │
│  └─ Custom OOXML Operations   └─ Workflow Automation          │
├─────────────────────────────────────────────────────────────────┤
│  🎯 Extension Framework                                        │
│  ├─ Plugin Registry          ├─ Extension Loader              │
│  ├─ API Extension Interface  ├─ Custom Operation Builder      │
│  └─ Version Management       └─ Marketplace Integration       │
├─────────────────────────────────────────────────────────────────┤
│  📊 Enhanced OOXML Core (Existing + Extensions)               │
│  ├─ OOXMLJsonService         ├─ Universal OOXML Engine        │
│  ├─ Session Management       ├─ Large File Support            │
│  └─ Cloud Run Integration    └─ GCS Storage                   │
├─────────────────────────────────────────────────────────────────┤
│  🏗️ Infrastructure Layer                                       │
│  ├─ Auto-scaling Cloud Run   ├─ Redis Caching                │
│  ├─ Cloud Storage            ├─ Monitoring & Logging          │
│  └─ Security & Compliance    └─ Backup & Recovery             │
└─────────────────────────────────────────────────────────────────┘
```

## Core Components

### 1. Universal Operations Engine

#### Global Search and Replace Service

**Purpose**: Enable text replacement operations across all OOXML elements that the Google Slides API cannot access.

**Technical Implementation**:
```javascript
class UniversalSearchReplace {
  constructor(ooxmlService) {
    this.ooxmlService = ooxmlService;
    this.searchPatterns = new Map();
    this.replaceOperations = new Map();
  }

  async performGlobalReplace(fileId, operations, options = {}) {
    // 1. Extract OOXML to JSON manifest
    const manifest = await this.ooxmlService.unwrap(fileId);
    
    // 2. Process all XML parts systematically
    const processedParts = await this._processAllParts(manifest, operations);
    
    // 3. Apply changes to binary parts if needed
    const processedBinaries = await this._processBinaryParts(manifest, operations);
    
    // 4. Rebuild and validate OOXML
    return await this.ooxmlService.rewrap({
      ...manifest,
      parts: processedParts,
      binaries: processedBinaries
    }, options);
  }

  async _processAllParts(manifest, operations) {
    const processedParts = {};
    
    for (const [partPath, partContent] of Object.entries(manifest.parts)) {
      let processedContent = partContent;
      
      // Apply each operation in sequence
      for (const operation of operations) {
        processedContent = await this._applyOperation(
          processedContent, 
          operation, 
          partPath
        );
      }
      
      processedParts[partPath] = processedContent;
    }
    
    return processedParts;
  }

  async _applyOperation(content, operation, partPath) {
    switch (operation.type) {
      case 'textReplace':
        return this._performTextReplace(content, operation);
      case 'attributeReplace':
        return this._performAttributeReplace(content, operation);
      case 'elementReplace':
        return this._performElementReplace(content, operation);
      case 'conditional':
        return this._performConditionalReplace(content, operation, partPath);
      default:
        throw new Error(`Unsupported operation type: ${operation.type}`);
    }
  }
}
```

#### Template Management System

**Purpose**: Provide programmatic control over theme elements, colors, fonts, and layouts.

**Technical Implementation**:
```javascript
class TemplateManager {
  constructor(ooxmlService) {
    this.ooxmlService = ooxmlService;
    this.themeCache = new Map();
    this.fontCache = new Map();
  }

  async applyColorScheme(fileId, colorScheme, options = {}) {
    const operations = [
      // Theme colors in theme XML
      {
        type: 'themeColorReplace',
        targetPath: 'ppt/theme/theme1.xml',
        mappings: this._buildColorMappings(colorScheme)
      },
      // Master slide color references
      {
        type: 'masterSlideColorUpdate',
        targetPattern: 'ppt/slideMasters/slideMaster*.xml',
        mappings: this._buildMasterColorMappings(colorScheme)
      },
      // Layout color inheritance
      {
        type: 'layoutColorUpdate',
        targetPattern: 'ppt/slideLayouts/slideLayout*.xml',
        mappings: this._buildLayoutColorMappings(colorScheme)
      }
    ];

    return await this.ooxmlService.process(fileId, operations, {
      validateTheme: true,
      preserveCustomColors: options.preserveCustom || false,
      generatePreview: options.generatePreview || true
    });
  }

  async applyFontScheme(fileId, fontScheme, options = {}) {
    const operations = [
      // Theme fonts
      {
        type: 'themeFontReplace',
        targetPath: 'ppt/theme/theme1.xml',
        majorFont: fontScheme.major,
        minorFont: fontScheme.minor
      },
      // Direct font references in slides
      {
        type: 'directFontReplace',
        targetPattern: 'ppt/slides/slide*.xml',
        mappings: this._buildFontMappings(fontScheme)
      },
      // Master and layout font updates
      {
        type: 'templateFontUpdate',
        targetPattern: 'ppt/slide{Masters,Layouts}/*.xml',
        mappings: this._buildTemplateFontMappings(fontScheme)
      }
    ];

    return await this.ooxmlService.process(fileId, operations, {
      validateFonts: true,
      embedFonts: options.embedFonts || false,
      generateFallbacks: options.generateFallbacks || true
    });
  }
}
```

### 2. Enhanced Extension Framework

#### Plugin Architecture for Custom Operations

**Purpose**: Allow developers to create custom OOXML operations that extend the platform's capabilities.

**Technical Implementation**:
```javascript
class ExtensionRegistry {
  constructor() {
    this.extensions = new Map();
    this.hooks = new Map();
    this.validators = new Map();
  }

  registerExtension(name, extension, metadata = {}) {
    // Validate extension interface
    this._validateExtensionInterface(extension);
    
    // Register with metadata
    this.extensions.set(name, {
      instance: extension,
      metadata: {
        version: metadata.version || '1.0.0',
        author: metadata.author || 'Unknown',
        description: metadata.description || '',
        capabilities: metadata.capabilities || [],
        dependencies: metadata.dependencies || [],
        hooks: metadata.hooks || []
      },
      registeredAt: new Date()
    });

    // Register hooks if provided
    if (metadata.hooks) {
      this._registerHooks(name, metadata.hooks);
    }

    return this;
  }

  async executeExtension(name, input, context = {}) {
    const extension = this.extensions.get(name);
    if (!extension) {
      throw new Error(`Extension '${name}' not found`);
    }

    // Pre-execution hooks
    await this._executeHooks('beforeExtension', { name, input, context });

    try {
      // Execute main extension logic
      const result = await extension.instance.execute(input, context);

      // Post-execution hooks
      await this._executeHooks('afterExtension', { name, input, result, context });

      return result;
    } catch (error) {
      // Error hooks
      await this._executeHooks('extensionError', { name, input, error, context });
      throw error;
    }
  }
}

// Base extension interface
class BaseAPIExtension {
  constructor(config = {}) {
    this.config = config;
    this.name = this.constructor.name;
  }

  // Required methods
  async execute(input, context) {
    throw new Error(`Extension ${this.name} must implement execute method`);
  }

  getCapabilities() {
    return [];
  }

  validate(input) {
    return { valid: true, errors: [] };
  }

  // Optional lifecycle methods
  async init(context) {}
  async cleanup(context) {}
}

// Example custom extension
class GlobalTextTransformExtension extends BaseAPIExtension {
  getCapabilities() {
    return ['textTransform', 'batchOperation', 'undoableOperation'];
  }

  async execute(input, context) {
    const { fileId, transformations, options = {} } = input;
    
    // Build operations for each transformation
    const operations = transformations.map(transform => ({
      type: 'globalTextTransform',
      ...transform,
      preserveFormatting: options.preserveFormatting || true,
      scope: options.scope || 'all'
    }));

    // Execute via OOXML service
    return await context.ooxmlService.process(fileId, operations, {
      validateResult: true,
      createBackup: options.createBackup || true,
      generateReport: true
    });
  }
}
```

### 3. API Gateway and Interface Design

#### REST API Endpoints

**Base URL**: `https://api.ooxml-extension.com/v1`

**Core Endpoints**:

```yaml
# Universal Operations
POST /presentations/{id}/operations/search-replace
POST /presentations/{id}/operations/batch-transform
POST /presentations/{id}/operations/custom

# Template Management
POST /presentations/{id}/templates/apply-colors
POST /presentations/{id}/templates/apply-fonts
POST /presentations/{id}/templates/apply-theme
GET  /templates/color-schemes
GET  /templates/font-schemes

# Batch Processing
POST /batch/operations
GET  /batch/{batchId}/status
GET  /batch/{batchId}/results

# Extensions
GET  /extensions
POST /extensions/{name}/execute
GET  /extensions/{name}/capabilities

# Sessions (Large Files)
POST /sessions
POST /sessions/{id}/upload
GET  /sessions/{id}/status
```

**Example API Usage**:
```javascript
// Global search and replace
const response = await fetch('/v1/presentations/abc123/operations/search-replace', {
  method: 'POST',
  headers: {
    'Authorization': 'Bearer ' + token,
    'Content-Type': 'application/json'
  },
  body: JSON.stringify({
    operations: [
      {
        type: 'textReplace',
        find: 'ACME Corporation',
        replace: 'DeltaQuad Industries',
        scope: 'all',
        matchCase: false
      },
      {
        type: 'colorReplace',
        find: '#FF0000',
        replace: '#0066CC',
        scope: 'theme'
      }
    ],
    options: {
      preserveFormatting: true,
      createBackup: true,
      validateResult: true
    }
  })
});

// Apply color scheme
const colorResult = await fetch('/v1/presentations/abc123/templates/apply-colors', {
  method: 'POST',
  headers: {
    'Authorization': 'Bearer ' + token,
    'Content-Type': 'application/json'
  },
  body: JSON.stringify({
    colorScheme: {
      primary: '#0066CC',
      secondary: '#FF6600',
      accent: '#00AA44',
      neutral: '#666666'
    },
    options: {
      preserveCustomColors: false,
      validateAccessibility: true
    }
  })
});
```

#### GraphQL Interface

**Schema Definition**:
```graphql
type Presentation {
  id: ID!
  title: String!
  slideCount: Int!
  theme: Theme
  lastModified: DateTime!
}

type Theme {
  colors: ColorScheme!
  fonts: FontScheme!
  effects: EffectScheme!
}

type ColorScheme {
  primary: Color!
  secondary: Color!
  accent: [Color!]!
  neutral: [Color!]!
}

type OperationResult {
  success: Boolean!
  presentationId: ID!
  operationsApplied: Int!
  errors: [OperationError!]!
  report: OperationReport
}

type Query {
  presentation(id: ID!): Presentation
  extensions: [Extension!]!
  templates: TemplateLibrary!
  batchOperation(id: ID!): BatchOperation
}

type Mutation {
  globalSearchReplace(
    presentationId: ID!,
    operations: [SearchReplaceInput!]!,
    options: OperationOptions
  ): OperationResult!
  
  applyColorScheme(
    presentationId: ID!,
    scheme: ColorSchemeInput!,
    options: TemplateOptions
  ): OperationResult!
  
  executeExtension(
    name: String!,
    input: JSON!,
    context: JSON
  ): ExtensionResult!
  
  createBatchOperation(
    presentationIds: [ID!]!,
    operations: [OperationInput!]!,
    options: BatchOptions
  ): BatchOperation!
}
```

### 4. Performance and Scalability Architecture

#### Caching Strategy

```javascript
class OOXMLCacheManager {
  constructor(redisClient, gcsStorage) {
    this.redis = redisClient;
    this.gcs = gcsStorage;
    this.memoryCache = new Map();
  }

  async getCachedManifest(fileId, version) {
    // L1: Memory cache
    const memoryKey = `manifest:${fileId}:${version}`;
    if (this.memoryCache.has(memoryKey)) {
      return this.memoryCache.get(memoryKey);
    }

    // L2: Redis cache
    const redisKey = `ooxml:manifest:${fileId}:${version}`;
    const cached = await this.redis.get(redisKey);
    if (cached) {
      const manifest = JSON.parse(cached);
      this.memoryCache.set(memoryKey, manifest);
      return manifest;
    }

    return null;
  }

  async setCachedManifest(fileId, version, manifest, ttl = 3600) {
    const memoryKey = `manifest:${fileId}:${version}`;
    const redisKey = `ooxml:manifest:${fileId}:${version}`;
    
    // Store in memory
    this.memoryCache.set(memoryKey, manifest);
    
    // Store in Redis with TTL
    await this.redis.setex(redisKey, ttl, JSON.stringify(manifest));
  }
}
```

#### Auto-scaling Configuration

```yaml
apiVersion: serving.knative.dev/v1
kind: Service
metadata:
  name: ooxml-api-extension
  annotations:
    run.googleapis.com/cpu-throttling: "false"
    autoscaling.knative.dev/minScale: "1"
    autoscaling.knative.dev/maxScale: "100"
    autoscaling.knative.dev/target: "80"
spec:
  template:
    metadata:
      annotations:
        run.googleapis.com/memory: "4Gi"
        run.googleapis.com/cpu: "2"
    spec:
      containers:
      - image: gcr.io/project/ooxml-api-extension
        ports:
        - containerPort: 8080
        env:
        - name: NODE_ENV
          value: production
        - name: REDIS_URL
          valueFrom:
            secretKeyRef:
              name: redis-config
              key: url
        resources:
          limits:
            cpu: "2"
            memory: "4Gi"
          requests:
            cpu: "1"
            memory: "2Gi"
```

### 5. Security and Compliance

#### Authentication and Authorization

```javascript
class SecurityManager {
  constructor(config) {
    this.config = config;
    this.tokenValidator = new TokenValidator(config.oauth);
    this.permissionEngine = new PermissionEngine(config.rbac);
  }

  async validateRequest(req, res, next) {
    try {
      // Extract and validate token
      const token = this.extractToken(req);
      const claims = await this.tokenValidator.validate(token);
      
      // Check permissions for requested operation
      const permission = this.determineRequiredPermission(req);
      const hasPermission = await this.permissionEngine.check(
        claims.sub, 
        permission, 
        this.extractResourceContext(req)
      );
      
      if (!hasPermission) {
        return res.status(403).json({
          error: 'INSUFFICIENT_PERMISSIONS',
          message: `Permission '${permission}' required for this operation`
        });
      }

      // Add user context to request
      req.user = claims;
      req.permissions = await this.permissionEngine.getUserPermissions(claims.sub);
      
      next();
    } catch (error) {
      res.status(401).json({
        error: 'AUTHENTICATION_FAILED',
        message: error.message
      });
    }
  }

  determineRequiredPermission(req) {
    const { method, path } = req;
    
    const permissionMap = {
      'POST /v1/presentations/*/operations/search-replace': 'presentations:modify',
      'POST /v1/presentations/*/templates/apply-colors': 'templates:apply',
      'POST /v1/batch/operations': 'batch:create',
      'POST /v1/extensions/*/execute': 'extensions:execute'
    };

    const pattern = `${method} ${path}`.replace(/\/[^\/]+\/presentations\/[^\/]+/, '/presentations/*');
    return permissionMap[pattern] || 'presentations:read';
  }
}
```

## Data Models and Storage

### OOXML Manifest Structure

```javascript
const OOXMLManifestSchema = {
  id: 'string',
  version: 'string',
  format: 'pptx|docx|xlsx',
  metadata: {
    originalSize: 'number',
    fileCount: 'number',
    extractedAt: 'datetime',
    checksum: 'string'
  },
  parts: {
    // XML parts as editable JSON/text
    '[content_types].xml': 'string',
    'ppt/presentation.xml': 'string',
    'ppt/slides/slide1.xml': 'string',
    'ppt/theme/theme1.xml': 'string'
    // ... other XML parts
  },
  binaries: {
    // Binary parts as GCS references
    'ppt/media/image1.png': {
      gcsPath: 'string',
      size: 'number',
      mimeType: 'string',
      checksum: 'string'
    }
  },
  relationships: {
    // OOXML relationships
    '_rels/.rels': 'object',
    'ppt/_rels/presentation.xml.rels': 'object'
  }
};
```

### Extension Metadata Schema

```javascript
const ExtensionSchema = {
  name: 'string',
  version: 'string',
  author: 'string',
  description: 'string',
  homepage: 'string',
  repository: 'string',
  capabilities: ['string'],
  dependencies: ['string'],
  hooks: {
    beforeOperation: 'boolean',
    afterOperation: 'boolean',
    onError: 'boolean'
  },
  parameters: {
    // JSON Schema for extension parameters
    type: 'object',
    properties: 'object',
    required: ['string']
  },
  examples: [{
    name: 'string',
    description: 'string',
    input: 'object',
    expectedOutput: 'object'
  }]
};
```

## Testing Strategy

### Unit Testing Framework

```javascript
describe('UniversalSearchReplace', () => {
  let service;
  let mockOOXMLService;

  beforeEach(() => {
    mockOOXMLService = new MockOOXMLService();
    service = new UniversalSearchReplace(mockOOXMLService);
  });

  it('should perform global text replacement across all parts', async () => {
    // Setup mock manifest
    const manifest = createMockManifest({
      'ppt/slides/slide1.xml': '<p:sp><p:txBody><a:p><a:r><a:t>ACME Corp</a:t></a:r></a:p></p:txBody></p:sp>',
      'ppt/slides/slide2.xml': '<p:sp><p:txBody><a:p><a:r><a:t>ACME Corp Rules</a:t></a:r></a:p></p:txBody></p:sp>'
    });
    
    mockOOXMLService.setMockManifest('test-file', manifest);

    // Execute replacement
    const result = await service.performGlobalReplace('test-file', [{
      type: 'textReplace',
      find: 'ACME Corp',
      replace: 'DeltaQuad Inc'
    }]);

    // Verify results
    expect(result.success).toBe(true);
    expect(result.report.replacements).toBe(2);
    expect(mockOOXMLService.getProcessedManifest()['ppt/slides/slide1.xml'])
      .toContain('DeltaQuad Inc');
  });

  it('should handle complex conditional replacements', async () => {
    // Test conditional replacement logic
  });

  it('should maintain OOXML integrity after operations', async () => {
    // Test OOXML validation
  });
});

// Integration tests
describe('API Integration', () => {
  it('should handle complete workflow from API to result', async () => {
    const response = await request(app)
      .post('/v1/presentations/test-file/operations/search-replace')
      .set('Authorization', 'Bearer ' + testToken)
      .send({
        operations: [
          { type: 'textReplace', find: 'test', replace: 'production' }
        ]
      });

    expect(response.status).toBe(200);
    expect(response.body.success).toBe(true);
  });
});
```

### Visual Validation Testing

```javascript
class VisualValidationFramework {
  constructor(playwrightConfig) {
    this.browser = null;
    this.config = playwrightConfig;
  }

  async validatePresentationRender(presentationId, expectedScreenshot = null) {
    // Open presentation in Google Slides
    const page = await this.browser.newPage();
    await page.goto(`https://docs.google.com/presentation/d/${presentationId}/edit`);
    
    // Wait for full render
    await page.waitForSelector('[data-test-id="presentation-canvas"]', { timeout: 30000 });
    await page.waitForTimeout(2000); // Allow for animations
    
    // Capture screenshot
    const screenshot = await page.screenshot({ fullPage: true });
    
    // Compare with baseline if provided
    if (expectedScreenshot) {
      const similarity = await this.compareImages(screenshot, expectedScreenshot);
      return {
        similarity,
        passed: similarity > 0.95, // 95% similarity threshold
        screenshot: screenshot.toString('base64')
      };
    }
    
    return { screenshot: screenshot.toString('base64') };
  }

  async compareImages(image1, image2) {
    // Use image comparison library (e.g., pixelmatch)
    return 0.98; // Mock similarity score
  }
}
```

This technical specification provides the foundation for building a comprehensive OOXML API Extension Platform that truly extends Google Slides API capabilities beyond their current limitations, making advanced OOXML operations accessible through simple, powerful interfaces.