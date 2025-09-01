/**
 * OOXML JSON Quick Start Examples
 * 
 * PURPOSE:
 * Demonstrates the complete OOXML JSON Helper workflow from deployment
 * to advanced operations. Shows both basic usage and advanced server-side
 * operations for efficient PowerPoint manipulation.
 * 
 * USAGE:
 * Run these examples in Google Apps Script after setting up the OOXML JSON Service.
 * Make sure to set your GCP_PROJECT_ID in Script Properties first.
 */

// ===== SETUP AND DEPLOYMENT =====

/**
 * Example 1: One-Click Deployment Setup
 * 
 * This shows the complete setup flow from preflight checks to deployment.
 * Run this first to deploy the OOXML JSON Service to Cloud Run.
 */
function example1_DeploymentSetup() {
  try {
    Logger.log('=== OOXML JSON Service Deployment ===');
    
    // Step 1: Show preflight interface for user setup
    Logger.log('Step 1: Run preflight checks...');
    OOXMLDeployment.showGcpPreflight();
    
    // Note: User will interact with the sidebar to:
    // - Verify billing is enabled
    // - Enable required APIs
    // - Set up budget alerts
    
    Logger.log('Complete the preflight checks in the sidebar, then run example1b_Deploy()');
    
  } catch (error) {
    Logger.log(`Deployment setup failed: ${error.message}`);
  }
}

/**
 * Example 1b: Execute Deployment
 * 
 * Run this after completing preflight checks to deploy the service.
 */
async function example1b_Deploy() {
  try {
    Logger.log('=== Deploying OOXML JSON Service ===');
    
    // Deploy the service (takes 3-5 minutes)
    const serviceUrl = await OOXMLDeployment.initAndDeploy();
    
    Logger.log(`✅ Service deployed successfully!`);
    Logger.log(`Service URL: ${serviceUrl}`);
    
    // Verify deployment
    const status = OOXMLDeployment.getDeploymentStatus();
    Logger.log(`Deployment status: ${JSON.stringify(status, null, 2)}`);
    
  } catch (error) {
    Logger.log(`Deployment failed: ${error.message}`);
  }
}

// ===== BASIC OPERATIONS =====

/**
 * Example 2: Basic File Unwrapping
 * 
 * Shows how to unwrap a PowerPoint file into a JSON manifest
 * for easy XML manipulation.
 */
async function example2_BasicUnwrap() {
  try {
    Logger.log('=== Basic OOXML Unwrapping ===');
    
    // Unwrap a PowerPoint file
    const manifest = await OOXMLJsonService.unwrap('sample.pptx');
    
    Logger.log(`File type: ${manifest.kind}`);
    Logger.log(`Total entries: ${manifest.entries.length}`);
    
    // Show XML entries
    const xmlEntries = manifest.entries.filter(e => e.type === 'xml');
    Logger.log(`XML entries: ${xmlEntries.length}`);
    
    xmlEntries.slice(0, 5).forEach(entry => {
      Logger.log(`- ${entry.path} (${entry.text.length} chars)`);
    });
    
    // Show binary entries
    const binEntries = manifest.entries.filter(e => e.type === 'bin');
    Logger.log(`Binary entries: ${binEntries.length}`);
    
    return manifest;
    
  } catch (error) {
    Logger.log(`Unwrap failed: ${error.message}`);
  }
}

/**
 * Example 3: Simple Text Replacement
 * 
 * Shows basic text replacement across an entire presentation
 * using the JSON manifest.
 */
async function example3_SimpleTextReplacement() {
  try {
    Logger.log('=== Simple Text Replacement ===');
    
    // Load and unwrap presentation
    const manifest = await OOXMLJsonService.unwrap('company-presentation.pptx');
    
    // Replace company name in all XML content
    let replacementCount = 0;
    
    manifest.entries.forEach(entry => {
      if (entry.type === 'xml' && entry.text.includes('ACME Corp')) {
        entry.text = entry.text.replace(/ACME Corp/g, 'DeltaQuad Inc');
        replacementCount++;
      }
    });
    
    Logger.log(`Made replacements in ${replacementCount} files`);
    
    // Rewrap and save
    const outputFileId = await OOXMLJsonService.rewrap(manifest, {
      filename: 'company-presentation-updated.pptx'
    });
    
    Logger.log(`✅ Updated presentation saved: ${outputFileId}`);
    
    return outputFileId;
    
  } catch (error) {
    Logger.log(`Text replacement failed: ${error.message}`);
  }
}

// ===== SERVER-SIDE OPERATIONS =====

/**
 * Example 4: Server-Side Text Replacement
 * 
 * Uses the powerful server-side /process endpoint for efficient
 * text replacement without downloading the entire manifest.
 */
async function example4_ServerSideOperations() {
  try {
    Logger.log('=== Server-Side Operations ===');
    
    // Define operations to perform server-side
    const operations = [
      {
        type: 'replaceText',
        scope: 'ppt/slides/',  // Only in slide content
        find: 'ACME Corp',
        replace: 'DeltaQuad Inc'
      },
      {
        type: 'replaceText',
        scope: 'ppt/slides/',
        find: 'Old Tagline',
        replace: 'Innovation Through Technology'
      },
      {
        type: 'upsertPart',
        path: 'ppt/customXml/item1.xml',
        text: '<?xml version="1.0"?><metadata><updated>' + new Date().toISOString() + '</updated></metadata>',
        contentType: 'application/xml'
      }
    ];
    
    // Execute all operations server-side in one call
    const result = await OOXMLJsonService.process(
      'company-presentation.pptx',
      operations,
      { filename: 'company-presentation-processed.pptx' }
    );
    
    Logger.log(`✅ Server-side processing complete!`);
    Logger.log(`File ID: ${result.fileId}`);
    Logger.log(`Report: ${JSON.stringify(result.report, null, 2)}`);
    
    return result;
    
  } catch (error) {
    Logger.log(`Server-side operations failed: ${error.message}`);
  }
}

/**
 * Example 5: Large File Session Workflow
 * 
 * Demonstrates session-based handling for large presentations
 * using GCS signed URLs for efficient processing.
 */
async function example5_LargeFileSession() {
  try {
    Logger.log('=== Large File Session Workflow ===');
    
    // Create session for large file operations
    const session = await OOXMLJsonService.createSession();
    Logger.log(`Session created: ${session.sessionId}`);
    
    // Upload large file to session
    await OOXMLJsonService.uploadToSession(session.uploadUrl, 'large-presentation.pptx');
    Logger.log('Large file uploaded to session');
    
    // Unwrap from GCS
    const manifest = await OOXMLJsonService.unwrap(null, { 
      gcsIn: session.gcsIn 
    });
    
    Logger.log(`Unwrapped ${manifest.entries.length} entries from large file`);
    
    // Make modifications to manifest
    manifest.entries.forEach(entry => {
      if (entry.type === 'xml' && entry.path.includes('slide')) {
        entry.text = entry.text.replace(/Large Corp/g, 'Mega Enterprise');
      }
    });
    
    // Rewrap to GCS
    const result = await OOXMLJsonService.rewrap(manifest, {
      gcsIn: session.gcsIn,
      gcsOut: session.gcsOut
    });
    
    Logger.log(`✅ Large file processing complete: ${result.gcsOut}`);
    Logger.log(`Download URL: ${session.downloadUrl}`);
    
    return result;
    
  } catch (error) {
    Logger.log(`Large file session failed: ${error.message}`);
  }
}

// ===== INTEGRATION WITH EXTENSION FRAMEWORK =====

/**
 * Example 6: Extension Framework Integration
 * 
 * Shows how to use the upgraded OOXMLSlides class with
 * OOXML JSON Service backend and extension framework.
 */
async function example6_ExtensionIntegration() {
  try {
    Logger.log('=== Extension Framework Integration ===');
    
    // Initialize with OOXML JSON Service backend
    const slides = new OOXMLSlides('brand-template.pptx', {
      enableExtensions: true,
      createBackup: true
    });
    
    // Load using JSON service
    await slides.load();
    Logger.log('Presentation loaded with JSON service backend');
    
    // Apply brand colors using extension
    const colorResult = await slides.applyBrandColors({
      primary: '#0066CC',
      secondary: '#FF6600',
      accent: '#00AA44'
    });
    
    Logger.log(`Brand colors applied: ${JSON.stringify(colorResult)}`);
    
    // Validate compliance
    const complianceResult = await slides.validateCompliance({
      allowedFonts: ['Arial', 'Helvetica', 'Calibri'],
      allowedColors: ['#0066CC', '#FF6600', '#00AA44'],
      requireLogo: true
    });
    
    Logger.log(`Compliance score: ${complianceResult.score}%`);
    
    // Use custom extension for advanced operations
    const customResult = await slides.useExtension('BrandCompliance', {
      autoFix: true,
      strictMode: false
    });
    
    Logger.log(`Custom extension result: ${JSON.stringify(customResult)}`);
    
    // Save with JSON service backend
    const saveResult = await slides.save({
      filename: 'branded-presentation.pptx'
    });
    
    Logger.log(`✅ Presentation saved: ${saveResult.fileId}`);
    
    return saveResult;
    
  } catch (error) {
    Logger.log(`Extension integration failed: ${error.message}`);
  }
}

// ===== ADVANCED WORKFLOWS =====

/**
 * Example 7: Batch Processing Multiple Presentations
 * 
 * Shows how to efficiently process multiple presentations
 * using server-side operations and batching.
 */
async function example7_BatchProcessing() {
  try {
    Logger.log('=== Batch Processing Multiple Presentations ===');
    
    const presentations = [
      'quarterly-report-q1.pptx',
      'quarterly-report-q2.pptx', 
      'quarterly-report-q3.pptx',
      'quarterly-report-q4.pptx'
    ];
    
    const results = [];
    
    // Define common operations for all presentations
    const standardOperations = [
      {
        type: 'replaceText',
        find: '2023',
        replace: '2024'
      },
      {
        type: 'replaceText', 
        scope: 'ppt/slides/',
        find: 'Confidential',
        replace: 'Internal Use Only'
      },
      {
        type: 'upsertPart',
        path: 'ppt/customXml/metadata.xml',
        text: `<?xml version="1.0"?><metadata>
          <updated>${new Date().toISOString()}</updated>
          <version>2024-Q1</version>
          <status>Updated</status>
        </metadata>`,
        contentType: 'application/xml'
      }
    ];
    
    // Process each presentation
    for (const presentation of presentations) {
      Logger.log(`Processing: ${presentation}`);
      
      const result = await OOXMLJsonService.process(
        presentation,
        standardOperations,
        { filename: presentation.replace('.pptx', '-updated.pptx') }
      );
      
      results.push({
        original: presentation,
        processed: result.fileId,
        report: result.report
      });
      
      Logger.log(`✅ Processed: ${presentation} -> ${result.fileId}`);
    }
    
    Logger.log(`🎉 Batch processing complete! Processed ${results.length} presentations`);
    
    // Summary report
    const totalReplacements = results.reduce((sum, r) => sum + (r.report.replaced || 0), 0);
    Logger.log(`Total text replacements: ${totalReplacements}`);
    
    return results;
    
  } catch (error) {
    Logger.log(`Batch processing failed: ${error.message}`);
  }
}

/**
 * Example 8: Health Check and Monitoring
 * 
 * Shows how to monitor the OOXML JSON Service health
 * and get performance metrics.
 */
async function example8_HealthAndMonitoring() {
  try {
    Logger.log('=== Health Check and Monitoring ===');
    
    // Check service health
    const health = await OOXMLJsonService.healthCheck();
    Logger.log(`Service health: ${JSON.stringify(health, null, 2)}`);
    
    // Get service information
    const info = OOXMLJsonService.getServiceInfo();
    Logger.log(`Service info: ${JSON.stringify(info, null, 2)}`);
    
    // Check deployment status
    const status = OOXMLDeployment.getDeploymentStatus();
    Logger.log(`Deployment status: ${JSON.stringify(status, null, 2)}`);
    
    // Performance test
    const startTime = Date.now();
    
    const testManifest = await OOXMLJsonService.unwrap('test-presentation.pptx');
    const unwrapTime = Date.now() - startTime;
    
    const rewrapStart = Date.now();
    const testFileId = await OOXMLJsonService.rewrap(testManifest, {
      filename: 'test-output.pptx'
    });
    const rewrapTime = Date.now() - rewrapStart;
    
    Logger.log(`Performance metrics:`);
    Logger.log(`- Unwrap time: ${unwrapTime}ms`);
    Logger.log(`- Rewrap time: ${rewrapTime}ms`);
    Logger.log(`- Total round-trip: ${unwrapTime + rewrapTime}ms`);
    
    return {
      health,
      info,
      status,
      performance: {
        unwrapTime,
        rewrapTime,
        totalTime: unwrapTime + rewrapTime
      }
    };
    
  } catch (error) {
    Logger.log(`Health check failed: ${error.message}`);
  }
}

// ===== UTILITY FUNCTIONS =====

/**
 * Utility: Clean up test files and sessions
 */
function utility_Cleanup() {
  try {
    Logger.log('=== Cleanup Test Resources ===');
    
    // Clean up test files (optional - be careful!)
    const testFiles = [
      'company-presentation-updated.pptx',
      'company-presentation-processed.pptx',
      'branded-presentation.pptx',
      'test-output.pptx'
    ];
    
    testFiles.forEach(filename => {
      try {
        const files = DriveApp.getFilesByName(filename);
        while (files.hasNext()) {
          const file = files.next();
          DriveApp.removeFile(file);
          Logger.log(`Deleted: ${filename}`);
        }
      } catch (e) {
        // File not found, skip
      }
    });
    
    Logger.log('✅ Cleanup complete');
    
  } catch (error) {
    Logger.log(`Cleanup failed: ${error.message}`);
  }
}

/**
 * Utility: Complete service removal (use with caution!)
 */
function utility_RemoveService() {
  try {
    Logger.log('=== Complete Service Removal ===');
    
    const result = OOXMLDeployment.cleanup({
      removeService: true,      // Delete Cloud Run service
      removeBucket: false,      // Keep bucket (may have other files)
      clearProperties: true     // Clear Script Properties
    });
    
    Logger.log(`Cleanup result: ${JSON.stringify(result, null, 2)}`);
    
    if (result.errors.length > 0) {
      Logger.log('⚠️ Some cleanup operations failed:');
      result.errors.forEach(error => Logger.log(`- ${error}`));
    } else {
      Logger.log('✅ Service removal complete');
    }
    
  } catch (error) {
    Logger.log(`Service removal failed: ${error.message}`);
  }
}