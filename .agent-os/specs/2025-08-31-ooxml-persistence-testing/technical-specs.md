# OOXML Persistence Validation Testing - Technical Specifications

## Architecture Overview

The OOXML Persistence Validation Testing system implements a comprehensive roundtrip testing pipeline that validates OOXML feature preservation through Google Slides operations.

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│ Test PPTX       │────│ Google Slides    │────│ Downloaded PPTX │
│ Generation      │    │ Upload/Process   │    │ Analysis        │
└─────────────────┘    └──────────────────┘    └─────────────────┘
         │                       │                       │
         ▼                       ▼                       ▼
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│ OOXML Feature   │    │ Playwright       │    │ OOXML JSON      │
│ Catalog         │    │ Automation       │    │ Service         │
└─────────────────┘    └──────────────────┘    └─────────────────┘
         │                       │                       │
         ▼                       ▼                       ▼
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│ Visual          │    │ Screenshot       │    │ Compatibility   │
│ Comparison      │    │ Capture          │    │ Reports         │
└─────────────────┘    └──────────────────┘    └─────────────────┘
```

## Core Components

### 1. Test Presentation Generator

**File**: `test/ooxml-persistence/generators/TestPresentationGenerator.js`

```javascript
class TestPresentationGenerator {
  /**
   * Generate test presentations with specific OOXML features
   * @param {Object} featureSet - OOXML features to include
   * @returns {Promise<Object>} Generated presentation data
   */
  async generateTestPresentation(featureSet) {
    const features = {
      tableStyles: featureSet.includeTableStyles,
      customThemes: featureSet.includeCustomThemes,
      drawMLElements: featureSet.includeDrawML,
      typography: featureSet.includeTypography,
      masterSlides: featureSet.includeMasterSlides
    };
    
    // Use existing extension system
    const presentation = await this.createBasePresentation();
    
    if (features.tableStyles) {
      await this.addBrandwaresTableStyles(presentation);
    }
    
    if (features.customThemes) {
      await this.addCustomThemeElements(presentation);
    }
    
    // Continue for other features...
    
    return {
      presentation,
      features,
      baseline: await this.generateOOXMLBaseline(presentation)
    };
  }
}
```

### 2. Google Slides Automation Engine

**File**: `test/ooxml-persistence/automation/GoogleSlidesAutomation.js`

```javascript
class GoogleSlidesAutomation {
  constructor(page, config) {
    this.page = page;
    this.config = config;
    this.downloadPath = config.downloadPath;
  }
  
  /**
   * Complete roundtrip test workflow
   * @param {Buffer} pptxBuffer - Original PPTX file
   * @returns {Promise<Object>} Roundtrip results
   */
  async executeRoundtripTest(pptxBuffer, testOperations = []) {
    const results = {
      uploadResult: null,
      operationResults: [],
      downloadResult: null,
      timing: {}
    };
    
    // Step 1: Upload PPTX
    results.timing.uploadStart = Date.now();
    results.uploadResult = await this.uploadPPTX(pptxBuffer);
    results.timing.uploadEnd = Date.now();
    
    // Step 2: Perform test operations
    for (const operation of testOperations) {
      const operationResult = await this.executeOperation(operation);
      results.operationResults.push(operationResult);
    }
    
    // Step 3: Download processed PPTX
    results.timing.downloadStart = Date.now();
    results.downloadResult = await this.downloadPPTX();
    results.timing.downloadEnd = Date.now();
    
    return results;
  }
  
  /**
   * Upload PPTX to Google Slides
   */
  async uploadPPTX(pptxBuffer) {
    // Navigate to Google Slides
    await this.page.goto('https://docs.google.com/presentation/u/0/');
    
    // Wait for upload interface
    await this.page.waitForSelector('[aria-label="File"]');
    await this.page.click('[aria-label="File"]');
    
    // Handle file upload
    const fileInput = await this.page.waitForSelector('input[type="file"]');
    
    // Create temporary file for upload
    const tempFile = await this.createTempFile(pptxBuffer);
    await fileInput.setInputFiles(tempFile);
    
    // Wait for processing
    await this.page.waitForSelector('.punch-viewer-content', { timeout: 60000 });
    
    return {
      success: true,
      url: this.page.url(),
      timestamp: Date.now()
    };
  }
  
  /**
   * Download processed PPTX
   */
  async downloadPPTX() {
    await this.page.click('[aria-label="File"]');
    await this.page.waitForSelector('[role="menu"]');
    
    // Click Download
    await this.page.click('text="Download"');
    await this.page.waitForSelector('[role="menu"]');
    
    // Click PPTX option
    const downloadPromise = this.page.waitForDownload();
    await this.page.click('text="Microsoft PowerPoint (.pptx)"');
    
    const download = await downloadPromise;
    const buffer = await this.readDownloadAsBuffer(download);
    
    return {
      success: true,
      buffer,
      filename: download.suggestedFilename(),
      size: buffer.length
    };
  }
}
```

### 3. OOXML Analysis Engine

**File**: `test/ooxml-persistence/analysis/OOXMLAnalysisEngine.js`

```javascript
class OOXMLAnalysisEngine {
  constructor() {
    this.ooxmlService = new OOXMLJsonService();
  }
  
  /**
   * Compare OOXML structures before and after roundtrip
   * @param {Buffer} originalBuffer - Original PPTX
   * @param {Buffer} processedBuffer - Google Slides processed PPTX
   * @returns {Promise<Object>} Analysis results
   */
  async compareOOXMLStructures(originalBuffer, processedBuffer) {
    const originalManifest = await this.ooxmlService.generateManifest(originalBuffer);
    const processedManifest = await this.ooxmlService.generateManifest(processedBuffer);
    
    const analysis = {
      structuralChanges: this.analyzeStructuralChanges(originalManifest, processedManifest),
      featurePreservation: this.analyzeFeaturePreservation(originalManifest, processedManifest),
      contentChanges: this.analyzeContentChanges(originalManifest, processedManifest),
      qualityScore: 0
    };
    
    analysis.qualityScore = this.calculateQualityScore(analysis);
    
    return analysis;
  }
  
  /**
   * Analyze specific OOXML feature preservation
   */
  analyzeFeaturePreservation(original, processed) {
    const features = {
      tableStyles: this.checkTableStylePreservation(original, processed),
      themes: this.checkThemePreservation(original, processed),
      drawML: this.checkDrawMLPreservation(original, processed),
      typography: this.checkTypographyPreservation(original, processed),
      masterSlides: this.checkMasterSlidePreservation(original, processed)
    };
    
    return features;
  }
  
  /**
   * Check table style preservation
   */
  checkTableStylePreservation(original, processed) {
    const originalTableStyles = this.extractTableStyles(original);
    const processedTableStyles = this.extractTableStyles(processed);
    
    return {
      preserved: originalTableStyles.filter(style => 
        this.findMatchingStyle(style, processedTableStyles)
      ),
      lost: originalTableStyles.filter(style => 
        !this.findMatchingStyle(style, processedTableStyles)
      ),
      modified: this.findModifiedStyles(originalTableStyles, processedTableStyles),
      score: this.calculatePreservationScore(originalTableStyles, processedTableStyles)
    };
  }
}
```

### 4. Visual Comparison Engine

**File**: `test/ooxml-persistence/visual/VisualComparisonEngine.js`

```javascript
class VisualComparisonEngine {
  constructor(page, config) {
    this.page = page;
    this.config = config;
    this.screenshotOptions = {
      fullPage: true,
      type: 'png',
      quality: 100
    };
  }
  
  /**
   * Perform visual comparison of presentations
   * @param {string} originalPptxPath - Path to original PPTX
   * @param {string} processedPptxPath - Path to processed PPTX
   * @returns {Promise<Object>} Visual comparison results
   */
  async compareVisually(originalPptxPath, processedPptxPath) {
    const originalScreenshots = await this.captureSlideScreenshots(originalPptxPath);
    const processedScreenshots = await this.captureSlideScreenshots(processedPptxPath);
    
    const comparisons = [];
    
    for (let i = 0; i < Math.min(originalScreenshots.length, processedScreenshots.length); i++) {
      const comparison = await this.compareSlideImages(
        originalScreenshots[i], 
        processedScreenshots[i],
        i + 1
      );
      comparisons.push(comparison);
    }
    
    return {
      totalSlides: Math.max(originalScreenshots.length, processedScreenshots.length),
      comparedSlides: comparisons.length,
      overallSimilarity: this.calculateOverallSimilarity(comparisons),
      slideComparisons: comparisons,
      visualArtifacts: this.detectVisualArtifacts(comparisons)
    };
  }
  
  /**
   * Capture screenshots of all slides in a presentation
   */
  async captureSlideScreenshots(pptxPath) {
    // Upload to Google Slides for rendering
    await this.uploadForScreenshots(pptxPath);
    
    const screenshots = [];
    const slideCount = await this.getSlideCount();
    
    for (let i = 1; i <= slideCount; i++) {
      await this.navigateToSlide(i);
      await this.page.waitForTimeout(2000); // Allow rendering
      
      const screenshot = await this.page.screenshot({
        ...this.screenshotOptions,
        path: `temp/slide-${i}-${Date.now()}.png`
      });
      
      screenshots.push({
        slideNumber: i,
        buffer: screenshot,
        timestamp: Date.now()
      });
    }
    
    return screenshots;
  }
  
  /**
   * Compare two slide images using pixel-perfect analysis
   */
  async compareSlideImages(original, processed, slideNumber) {
    const pixelmatch = require('pixelmatch');
    const { PNG } = require('pngjs');
    
    const img1 = PNG.sync.read(original.buffer);
    const img2 = PNG.sync.read(processed.buffer);
    
    const { width, height } = img1;
    const diff = new PNG({ width, height });
    
    const pixelDifferences = pixelmatch(
      img1.data, img2.data, diff.data, width, height,
      { threshold: 0.1, alpha: 0.1 }
    );
    
    const totalPixels = width * height;
    const similarity = 1 - (pixelDifferences / totalPixels);
    
    return {
      slideNumber,
      similarity,
      pixelDifferences,
      totalPixels,
      diffImage: PNG.sync.write(diff),
      width,
      height,
      threshold: similarity > 0.95 ? 'PASS' : similarity > 0.85 ? 'WARN' : 'FAIL'
    };
  }
}
```

### 5. Agent OS Test Orchestration Framework

**File**: `test/ooxml-persistence/AgentOSPersistenceValidation.spec.js`
**Integration**: Extends existing Agent OS test patterns and MCP infrastructure

```javascript
const { test, expect } = require('@playwright/test');
const ExtensionTestGenerator = require('./generators/ExtensionTestGenerator');
const MCPGoogleSlidesAutomation = require('./automation/MCPGoogleSlidesAutomation');
const AgentOSManifestAnalyzer = require('./analysis/AgentOSManifestAnalyzer');
const MCPVisualValidator = require('./visual/MCPVisualValidator');
const AgentOSCompatibilityReporter = require('./reporting/AgentOSCompatibilityReporter');
// Import existing Agent OS components
const { ExtensionFramework } = require('../../lib/ExtensionFramework');
const { OOXMLJsonService } = require('../../lib/OOXMLJsonService');

test.describe('Agent OS Extension Framework Persistence Validation', () => {
  let extensionGenerator, mcpAutomation, manifestAnalyzer, visualValidator, reporter;
  let agentOSContext;
  
  test.beforeAll(async () => {
    // Initialize Agent OS context
    agentOSContext = await ExtensionFramework.initializeContext();
    extensionGenerator = new ExtensionTestGenerator(agentOSContext);
    manifestAnalyzer = new AgentOSManifestAnalyzer(agentOSContext);
    reporter = new AgentOSCompatibilityReporter();
  });
  
  test.beforeEach(async ({ page }) => {
    // Use existing Agent OS MCP infrastructure
    mcpAutomation = new MCPGoogleSlidesAutomation(page, {
      downloadPath: './test-downloads/',
      timeout: 60000,
      mcpServerEnabled: true,
      agentOSContext: agentOSContext
    });
    visualValidator = new MCPVisualValidator(page, {
      screenshotPath: './test-screenshots/',
      extensionAware: true
    });
  });
  
  /**
   * Test Agent OS TableStyleExtension preservation
   */
  test('Agent OS TableStyleExtension survives Google Slides roundtrip', async () => {
    // Generate test presentation using Agent OS Extension Framework
    const testPresentation = await extensionGenerator.generateTestPresentation({
      includeTableStyles: true,
      tableStyle: {
        type: 'TableStyle',
        config: {
          styleTypes: ['corporate', 'modern', 'minimalist'],
          includeConditionalFormatting: true,
          brandCompliant: true
        }
      }
    });
    
    // Execute roundtrip test using MCP infrastructure
    const roundtripResult = await mcpAutomation.executeExtensionRoundtripTest(
      testPresentation.jsonManifest, 
      [{
        type: 'extension-validation',
        extension: 'TableStyle',
        operations: ['apply', 'validate', 'preserve']
      }]
    );
    
    expect(roundtripResult.uploadResult.success).toBe(true);
    expect(roundtripResult.downloadResult.manifest).toBeDefined();
    
    // Analyze extension preservation using Agent OS patterns
    const analysis = await manifestAnalyzer.compareAgentOSManifests(
      testPresentation.jsonManifest,
      roundtripResult.downloadResult.manifest
    );
    
    // Validate Agent OS TableStyleExtension preservation
    expect(analysis.extensionPreservation.tableStyles.score).toBeGreaterThan(0.8);
    expect(analysis.extensionPreservation.tableStyles.aiAgentCompatibility.rating).toBeGreaterThan(0.85);
    expect(analysis.aiAgentRecommendations.length).toBeLessThan(3);
    
    // Visual comparison with extension awareness
    const visualComparison = await visualValidator.compareVisualsWithExtensionAwareness(
      testPresentation.jsonManifest,
      roundtripResult.downloadResult.manifest
    );
    
    expect(visualComparison.overallSimilarity).toBeGreaterThan(0.90);
    expect(visualComparison.aiAgentVisualRecommendations.length).toBeLessThan(2);
    
    // Generate Agent OS compatibility report
    await reporter.generateExtensionReport('TableStyle', analysis, visualComparison);
  });
  
  /**
   * Test Agent OS BrandColorsExtension and ThemeExtension preservation
   */
  test('Agent OS BrandColors and Theme extensions survive roundtrip', async () => {
    const testPresentation = await extensionGenerator.generateTestPresentation({
      includeBrandColors: true,
      includeThemes: true,
      brandColors: {
        type: 'BrandColors',
        config: {
          primaryColors: ['#0066CC', '#FF6600', '#00AA44'],
          secondaryColors: ['#666666', '#CCCCCC'],
          customPalette: true
        }
      },
      theme: {
        type: 'Theme',
        config: {
          themeTypes: ['corporate', 'modern'],
          includeColorSchemes: true,
          superThemeEnabled: true
        }
      }
    });
    
    const roundtripResult = await mcpAutomation.executeExtensionRoundtripTest(
      testPresentation.jsonManifest, 
      [{
        type: 'extension-validation',
        extensions: ['BrandColors', 'Theme'],
        operations: ['apply', 'validate', 'preserve']
      }]
    );
    
    const analysis = await manifestAnalyzer.compareAgentOSManifests(
      testPresentation.jsonManifest,
      roundtripResult.downloadResult.manifest
    );
    
    // Validate Agent OS extension preservation
    expect(analysis.extensionPreservation.brandColors.score).toBeGreaterThan(0.8);
    expect(analysis.extensionPreservation.themes.score).toBeGreaterThan(0.75);
    expect(analysis.agentOSQualityScore).toBeGreaterThan(0.80);
    
    await reporter.generateExtensionReport('BrandColors-Theme', analysis);
  });
  
  /**
   * Test Agent OS XMLSearchReplaceExtension preservation
   */
  test('Agent OS XMLSearchReplaceExtension operations survive roundtrip', async () => {
    const testPresentation = await extensionGenerator.generateTestPresentation({
      includeXMLOperations: true,
      xmlOperations: {
        type: 'XMLSearchReplace',
        config: {
          searchPatterns: [{
            pattern: '<a:solidFill>',
            replacement: '<a:solidFill><a:srgbClr val="FF6600"/>',
            scope: 'shapes'
          }],
          preserveStructure: true,
          validateXML: true
        }
      }
    });
    
    const roundtripResult = await mcpAutomation.executeExtensionRoundtripTest(
      testPresentation.jsonManifest,
      [{
        type: 'extension-validation',
        extension: 'XMLSearchReplace',
        operations: ['apply', 'validate', 'preserve']
      }]
    );
    
    const analysis = await manifestAnalyzer.compareAgentOSManifests(
      testPresentation.jsonManifest,
      roundtripResult.downloadResult.manifest
    );
    
    expect(analysis.extensionPreservation.xmlOperations.score).toBeGreaterThan(0.70);
    expect(analysis.extensionPreservation.xmlOperations.structuralIntegrity).toBe(true);
    
    await reporter.generateExtensionReport('XMLSearchReplace', analysis);
  });
  
  /**
   * Comprehensive Agent OS Extension Framework compatibility matrix test
   */
  test('Generate complete Agent OS Extension Framework compatibility matrix', async () => {
    const extensionSets = [
      { 
        name: 'BrandColors Only', 
        includeBrandColors: true,
        brandColors: { type: 'BrandColors', config: { primaryColors: ['#0066CC'] } }
      },
      { 
        name: 'Theme Only', 
        includeThemes: true,
        theme: { type: 'Theme', config: { themeTypes: ['corporate'] } }
      },
      { 
        name: 'TableStyle Only', 
        includeTableStyles: true,
        tableStyle: { type: 'TableStyle', config: { styleTypes: ['modern'] } }
      },
      { 
        name: 'Typography Only', 
        includeTypography: true,
        typography: { type: 'Typography', config: { fontPairs: ['arial-bold'] } }
      },
      { 
        name: 'All Agent OS Extensions', 
        includeBrandColors: true, 
        includeThemes: true, 
        includeTableStyles: true, 
        includeTypography: true,
        includeXMLOperations: true
      }
    ];
    
    const compatibilityMatrix = [];
    
    for (const extensionSet of extensionSets) {
      const testPresentation = await extensionGenerator.generateTestPresentation(extensionSet);
      const roundtripResult = await mcpAutomation.executeExtensionRoundtripTest(testPresentation.jsonManifest);
      const analysis = await manifestAnalyzer.compareAgentOSManifests(
        testPresentation.jsonManifest,
        roundtripResult.downloadResult.manifest
      );
      
      compatibilityMatrix.push({
        extensionSet: extensionSet.name,
        analysis,
        aiAgentRecommendations: analysis.aiAgentRecommendations,
        timestamp: Date.now()
      });
    }
    
    await reporter.generateAgentOSCompatibilityMatrix(compatibilityMatrix);
    
    // Validate overall Agent OS compatibility for AI agents
    const averageScore = compatibilityMatrix.reduce((sum, result) => 
      sum + result.analysis.agentOSQualityScore, 0) / compatibilityMatrix.length;
    
    const aiAgentReadiness = compatibilityMatrix.every(result => 
      result.analysis.aiAgentRecommendations.length < 5
    );
    
    expect(averageScore).toBeGreaterThan(0.75); // Higher threshold for Agent OS
    expect(aiAgentReadiness).toBe(true);
  });
});
```

## Agent OS Data Models

### Agent OS Extension Test Specification
```typescript
interface AgentOSExtensionTestSpec {
  extensions: {
    brandColors: boolean;
    themes: boolean;
    tableStyles: boolean;
    typography: boolean;
    xmlOperations: boolean;
  };
  extensionConfigs: {
    [extensionName: string]: ExtensionConfig;
  };
  complexity: 'simple' | 'moderate' | 'complex';
  slideCount: number;
  contentTypes: string[];
  aiAgentOptimized: boolean;
}

interface ExtensionConfig {
  type: string;
  config: Record<string, any>;
  agentOSCompatible: boolean;
}
```

### Agent OS Analysis Results
```typescript
interface AgentOSAnalysisResult {
  structuralChanges: {
    manifestFilesAdded: string[];
    manifestFilesRemoved: string[];
    manifestFilesModified: string[];
    jsonManifestIntegrity: boolean;
  };
  extensionPreservation: {
    brandColors: ExtensionPreservationResult;
    themes: ExtensionPreservationResult;
    tableStyles: ExtensionPreservationResult;
    typography: ExtensionPreservationResult;
    xmlOperations: ExtensionPreservationResult;
  };
  agentOSQualityScore: number; // 0-1, Agent OS specific scoring
  aiAgentRecommendations: AIAgentRecommendation[];
  ooxmlJsonServiceCompatibility: boolean;
}

interface ExtensionPreservationResult {
  score: number;
  preserved: any[];
  lost: any[];
  modified: any[];
  aiAgentCompatibility: {
    rating: number;
    issues: string[];
    recommendations: string[];
  };
  extensionFrameworkCompliance: boolean;
}

interface AIAgentRecommendation {
  type: 'extension-optimization' | 'structure-improvement' | 'ai-compatibility';
  extension?: string;
  recommendation: string;
  confidence: number;
  priority: 'low' | 'medium' | 'high';
}
```

### Agent OS Visual Comparison Results
```typescript
interface AgentOSVisualComparisonResult {
  overallSimilarity: number; // 0-1
  slideComparisons: ExtensionAwareSlideComparison[];
  extensionSpecificArtifacts: ExtensionVisualArtifact[];
  aiAgentVisualRecommendations: AIAgentVisualRecommendation[];
  mcpServerMetrics: {
    screenshotCaptureTime: number;
    comparisonProcessingTime: number;
    extensionAwarenessAccuracy: number;
  };
}

interface ExtensionAwareSlideComparison {
  slideNumber: number;
  similarity: number;
  extensionContext: {
    appliedExtensions: string[];
    extensionArtifacts: any[];
  };
  aiAgentRelevantChanges: boolean;
}

interface AIAgentVisualRecommendation {
  type: 'visual-quality' | 'extension-rendering' | 'ai-compatibility';
  recommendation: string;
  affectedExtensions: string[];
  confidence: number;
}
```

## Implementation Strategy

### Phase 1: Core Infrastructure
1. Set up test presentation generator
2. Implement Google Slides automation basics
3. Create OOXML analysis framework
4. Basic roundtrip testing

### Phase 2: Feature-Specific Testing
1. Table styles preservation testing
2. Theme and color scheme testing
3. DrawML elements testing
4. Typography testing

### Phase 3: Visual Validation
1. Screenshot capture automation
2. Pixel-perfect comparison engine
3. Visual regression detection
4. Render quality metrics

### Phase 4: Reporting and Analytics
1. Compatibility matrix generation
2. Executive dashboard
3. Trend analysis
4. Performance metrics

### Phase 5: Integration and Optimization
1. CI/CD pipeline integration
2. Performance optimization
3. Parallel test execution
4. Cloud-based testing

## Agent OS Configuration Management

**File**: `test/ooxml-persistence/config/agent-os-persistence-config.js`
**Integration**: Follows Agent OS JSON-based configuration patterns

```javascript
const { ExtensionFramework } = require('../../../lib/ExtensionFramework');
const { OOXMLJsonService } = require('../../../lib/OOXMLJsonService');

module.exports = {
  agentOSIntegration: {
    extensionFrameworkEnabled: true,
    ooxmlJsonServiceEnabled: true,
    mcpServerEnabled: true,
    cloudRunIntegration: true
  },
  
  testSuites: {
    quick: {
      extensionSets: ['BrandColors', 'Theme'],
      operations: ['extension-apply', 'save'],
      visualComparison: false,
      aiAgentOptimized: true
    },
    comprehensive: {
      extensionSets: ['BrandColors', 'Theme', 'TableStyle', 'Typography', 'XMLSearchReplace'],
      operations: ['extension-apply', 'save', 'copy', 'edit', 'share'],
      visualComparison: true,
      aiAgentOptimized: true,
      generateRecommendations: true
    },
    regression: {
      baselineDir: './test-baselines/',
      compareAgainstBaseline: true,
      generateNewBaseline: false,
      extensionFrameworkAware: true
    }
  },
  
  agentOSThresholds: {
    extensionPreservation: 0.8,
    visualSimilarity: 0.9,
    aiAgentCompatibility: 0.85,
    performanceTimeout: 120000,
    ooxmlJsonServiceTimeout: 60000
  },
  
  reporting: {
    outputDir: './test-reports/agent-os-persistence/',
    formats: ['json', 'html', 'csv'], // JSON first for AI agent consumption
    includeScreenshots: true,
    aiAgentReports: true,
    extensionFrameworkMetrics: true,
    jsonSchemaCompliant: true
  },
  
  // Agent OS specific configuration
  extensionFramework: {
    loadExistingExtensions: true,
    validateExtensionCompatibility: true,
    trackExtensionPerformance: true
  },
  
  ooxmlJsonService: {
    useEstablishedService: true,
    sessionBasedTesting: true,
    manifestComparison: true,
    serverSideOperations: ['replaceText', 'upsertPart', 'removePart']
  },
  
  mcpServer: {
    extendExistingServer: true,
    reuseAuthentication: true,
    leveragePlaywrightSetup: true,
    enableVisualTesting: true
  }
};
```

## Agent OS Testing Strategy

### Extension Framework Unit Tests
- Individual Agent OS extension testing using established patterns
- Mock Google Slides responses with Extension Framework context
- OOXMLJsonService JSON manifest analysis accuracy
- BaseExtension inheritance and lifecycle testing

### Agent OS Integration Tests
- End-to-end Extension Framework roundtrip testing
- Cross-browser compatibility using existing Playwright infrastructure
- Performance impact on OOXMLJsonService operations
- MCP server integration and tool availability

### Agent OS Regression Tests
- Extension Framework baseline comparison
- Agent OS extension preservation trend analysis
- Visual regression detection with extension awareness
- AI agent recommendation consistency validation

## Monitoring and Alerting

### Key Metrics
- Feature preservation rates
- Test execution time
- Visual similarity scores
- Error rates and types

### Alerts
- Preservation score drops below threshold
- New visual regressions detected
- Test execution failures
- Performance degradation

## Security and Privacy

### Data Handling
- Temporary file cleanup
- No sensitive data in test presentations
- Secure authentication management
- Download directory isolation

### Access Control
- Test-only Google accounts
- Limited API permissions
- Encrypted credential storage
- Audit logging

This Agent OS-integrated technical specification provides the foundation for implementing comprehensive Extension Framework persistence validation that directly supports our mission to "democratize enterprise-grade PowerPoint automation for AI agents." By building on established Agent OS architecture patterns, this system ensures AI agents can confidently deploy sophisticated OOXML manipulations in production scenarios while maintaining compatibility with the existing Cloud-Native Architecture and Extension-First Framework.