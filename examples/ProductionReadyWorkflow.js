/**
 * ProductionReadyWorkflow - Complete Enterprise Example
 * 
 * PURPOSE:
 * Demonstrates the full production-ready capabilities of the enhanced OOXML
 * system including systematic error handling, correlation tracking, brand
 * compliance DSL, and comprehensive observability.
 * 
 * This example shows how to build enterprise-grade PowerPoint automation
 * with proper logging, monitoring, and error handling.
 */

/**
 * Example 1: Production Deployment with Monitoring
 * 
 * Complete deployment workflow with health checks and monitoring setup.
 */
async function example1_ProductionDeployment() {
  const logger = Observability.logger('DeploymentExample');
  const correlationId = Correlation.start('production-deployment');
  
  try {
    logger.info('Starting production deployment workflow', { correlationId });
    
    // Setup observability
    Observability.setupDefaultHealthChecks();
    
    // Verify prerequisites
    logger.operation('verify_prerequisites', () => {
      const projectId = PropertiesService.getScriptProperties().getProperty('GCP_PROJECT_ID');
      if (!projectId) {
        throw OOXMLErrorCodes.create(
          'S010', 
          'GCP_PROJECT_ID not configured', 
          { step: 'prerequisites' },
          correlationId
        );
      }
      return { projectId };
    });
    
    // Run preflight checks
    const preflightResult = await logger.operation('preflight_checks', async () => {
      // This would open the sidebar for user interaction
      OOXMLDeployment.showGcpPreflight();
      
      // Simulate preflight completion
      return { billing: true, apis: true, budget: true };
    });
    
    // Deploy service
    const serviceUrl = await logger.operation('deploy_service', async () => {
      return await OOXMLDeployment.initAndDeploy({
        skipBillingCheck: false
      });
    });
    
    // Verify deployment
    const healthCheck = await logger.operation('verify_deployment', async () => {
      const health = await Observability.health.runAll();
      if (!health.healthy) {
        throw OOXMLErrorCodes.create(
          'S015',
          'Health checks failed after deployment',
          { health },
          correlationId
        );
      }
      return health;
    });
    
    logger.info('Production deployment completed successfully', {
      serviceUrl,
      healthStatus: healthCheck,
      correlationId
    });
    
    return {
      success: true,
      serviceUrl,
      healthStatus: healthCheck,
      correlationId
    };
    
  } catch (error) {
    OOXMLErrorLogger.logError(error, { operation: 'production_deployment' });
    throw error;
  } finally {
    Correlation.end(correlationId);
  }
}

/**
 * Example 2: Brand Compliance with DSL Rules
 * 
 * Advanced brand compliance using the declarative rules DSL with auto-fixing.
 */
async function example2_BrandComplianceDSL() {
  const logger = Observability.logger('BrandComplianceExample');
  const correlationId = Correlation.start('brand-compliance-dsl');
  
  try {
    logger.info('Starting brand compliance validation with DSL rules', { correlationId });
    
    // Define comprehensive brand rules
    const brandRules = {
      profile: 'enterprise-brand-2024',
      version: '2.0.0',
      metadata: {
        name: 'Enterprise Brand Guidelines 2024',
        description: 'Complete brand compliance rules with auto-fixing',
        author: 'Brand Team',
        updated: new Date().toISOString()
      },
      rules: [
        {
          id: 'brand.colors.primary',
          desc: 'Primary brand color must be corporate blue',
          category: 'color',
          where: 'ppt/theme/theme1.xml',
          xpath: '//a:accent1//a:srgbClr/@val',
          expect: { hex: '#0066CC' },
          autofix: true,
          weight: 10
        },
        {
          id: 'brand.colors.secondary',
          desc: 'Secondary brand color must be corporate orange',
          category: 'color',
          where: 'ppt/theme/theme1.xml',
          xpath: '//a:accent2//a:srgbClr/@val',
          expect: { hex: '#FF6600' },
          autofix: true,
          weight: 8
        },
        {
          id: 'brand.fonts.headers',
          desc: 'Header font must be brand font family',
          category: 'font',
          where: 'ppt/theme/theme1.xml',
          xpath: '//a:majorFont/a:latin/@typeface',
          expect: { font: 'Helvetica Neue' },
          autofix: true,
          weight: 7
        },
        {
          id: 'brand.fonts.body',
          desc: 'Body font must be brand font family',
          category: 'font',
          where: 'ppt/theme/theme1.xml',
          xpath: '//a:minorFont/a:latin/@typeface',
          expect: { font: 'Helvetica Neue' },
          autofix: true,
          weight: 6
        },
        {
          id: 'brand.content.confidentiality',
          desc: 'All slides must contain confidentiality notice',
          category: 'content',
          where: 'ppt/slides/slide*.xml',
          xpath: '//a:t[contains(translate(text(), "ABCDEFGHIJKLMNOPQRSTUVWXYZ", "abcdefghijklmnopqrstuvwxyz"), "confidential")]',
          expect: { regex: '.*[Cc]onfidential.*' },
          autofix: false,
          weight: 5
        },
        {
          id: 'brand.layout.logo_placement',
          desc: 'Logo must be present in slide masters',
          category: 'layout',
          where: 'ppt/slideMasters/slideMaster*.xml',
          xpath: '//a:blip/@r:embed',
          expect: { regex: '.*logo.*' },
          autofix: false,
          weight: 4
        }
      ]
    };
    
    // Load presentation
    const manifest = await logger.operation('load_presentation', async () => {
      return await OOXMLJsonService.unwrap('enterprise-presentation.pptx');
    });
    
    // Validate rules configuration
    const rulesEngine = new BrandbookRulesEngine();
    const rulesValidation = logger.operation('validate_rules', () => {
      return rulesEngine.validateRulesConfig(brandRules);
    });
    
    if (!rulesValidation.valid) {
      throw OOXMLErrorCodes.create(
        'V045',
        'Brand rules configuration invalid',
        { errors: rulesValidation.errors },
        correlationId
      );
    }
    
    // Execute comprehensive validation
    const validationResult = await logger.operation('execute_validation', async () => {
      return await rulesEngine.executeValidation(manifest, brandRules, {
        enableAutoFix: true,
        failFast: false,
        maxViolations: 50
      });
    });
    
    // Analyze results
    const analysis = logger.operation('analyze_results', () => {
      const criticalViolations = validationResult.violations.filter(v => v.weight >= 8);
      const autoFixedViolations = validationResult.violations.filter(v => v.autoFixed);
      const manualViolations = validationResult.violations.filter(v => !v.autoFixed);
      
      return {
        score: validationResult.score,
        totalViolations: validationResult.violations.length,
        criticalViolations: criticalViolations.length,
        autoFixed: autoFixedViolations.length,
        requiresManualFix: manualViolations.length,
        passed: validationResult.score >= 85, // 85% threshold
        categories: this._groupViolationsByCategory(validationResult.violations)
      };
    });
    
    // Save compliant version if auto-fixes were applied
    let outputFileId = null;
    if (validationResult.autoFixed > 0) {
      outputFileId = await logger.operation('save_compliant_version', async () => {
        return await OOXMLJsonService.rewrap(manifest, {
          filename: 'enterprise-presentation-compliant.pptx'
        });
      });
    }
    
    logger.info('Brand compliance validation completed', {
      analysis,
      outputFileId,
      correlationId
    });
    
    return {
      success: true,
      validationResult,
      analysis,
      outputFileId,
      correlationId
    };
    
  } catch (error) {
    OOXMLErrorLogger.logError(error, { operation: 'brand_compliance_dsl' });
    throw error;
  } finally {
    Correlation.end(correlationId);
  }
}

/**
 * Example 3: High-Performance Batch Processing
 * 
 * Enterprise-scale batch processing with monitoring and error recovery.
 */
async function example3_BatchProcessingWithMonitoring() {
  const logger = Observability.logger('BatchProcessingExample');
  const correlationId = Correlation.start('batch-processing');
  
  try {
    logger.info('Starting high-performance batch processing', { correlationId });
    
    // Setup performance monitoring
    const profiler = Observability.profiler.startProfiling('batch-processing');
    
    // Define batch operations
    const batchOperations = [
      {
        type: 'replaceText',
        scope: 'ppt/slides/',
        find: '2023',
        replace: '2024'
      },
      {
        type: 'replaceText',
        scope: 'ppt/slides/',
        find: 'Q4 2023',
        replace: 'Q1 2024'
      },
      {
        type: 'upsertPart',
        path: 'ppt/customXml/batchInfo.xml',
        text: `<?xml version="1.0"?>
          <batchInfo>
            <processedAt>${new Date().toISOString()}</processedAt>
            <batchId>${correlationId}</batchId>
            <version>2024.1</version>
          </batchInfo>`,
        contentType: 'application/xml'
      }
    ];
    
    // Get list of presentations to process
    const presentations = [
      'quarterly-report-q1.pptx',
      'quarterly-report-q2.pptx',
      'quarterly-report-q3.pptx',
      'quarterly-report-q4.pptx',
      'annual-summary.pptx',
      'board-presentation.pptx'
    ];
    
    profiler.addMarker('batch_start', { presentationCount: presentations.length });
    
    const results = [];
    const errors = [];
    
    // Process each presentation with error recovery
    for (let i = 0; i < presentations.length; i++) {
      const presentation = presentations[i];
      const childCorrelationId = Correlation.child(`batch-item-${i}`);
      
      try {
        profiler.addMarker(`processing_start_${i}`, { presentation });
        
        const result = await logger.operation(`process_${presentation}`, async () => {
          
          // Check if file exists
          const files = DriveApp.getFilesByName(presentation);
          if (!files.hasNext()) {
            throw OOXMLErrorCodes.create(
              'A022',
              `Presentation not found: ${presentation}`,
              { presentation, batchIndex: i },
              childCorrelationId
            );
          }
          
          // Process with server-side operations
          const processResult = await OOXMLJsonService.process(
            presentation,
            batchOperations,
            { filename: `${presentation.replace('.pptx', '-updated.pptx')}` }
          );
          
          // Verify processing results
          if (processResult.report.errors.length > 0) {
            logger.warn(`Processing warnings for ${presentation}`, {
              errors: processResult.report.errors,
              correlationId: childCorrelationId
            });
          }
          
          return {
            presentation,
            outputFileId: processResult.fileId,
            report: processResult.report,
            correlationId: childCorrelationId
          };
        });
        
        results.push(result);
        profiler.addMarker(`processing_end_${i}`, { success: true });
        
        // Add small delay to avoid rate limiting
        if (i < presentations.length - 1) {
          await new Promise(resolve => setTimeout(resolve, 100));
        }
        
      } catch (error) {
        const errorInfo = {
          presentation,
          error: error.message,
          errorCode: error.code,
          correlationId: childCorrelationId
        };
        
        errors.push(errorInfo);
        
        OOXMLErrorLogger.logError(error, { 
          operation: 'batch_processing_item',
          presentation,
          batchIndex: i 
        });
        
        profiler.addMarker(`processing_end_${i}`, { success: false, error: error.message });
        
        // Continue processing other files
        logger.warn(`Failed to process ${presentation}, continuing with batch`, errorInfo);
      } finally {
        Correlation.end(childCorrelationId);
      }
    }
    
    profiler.addMarker('batch_end', { 
      processed: results.length, 
      failed: errors.length 
    });
    
    // Generate batch summary
    const summary = {
      totalPresentations: presentations.length,
      successfullyProcessed: results.length,
      failed: errors.length,
      successRate: (results.length / presentations.length) * 100,
      totalReplacements: results.reduce((sum, r) => sum + (r.report?.replaced || 0), 0),
      totalUpserted: results.reduce((sum, r) => sum + (r.report?.upserted || 0), 0),
      errors: errors
    };
    
    const profilingResults = profiler.end();
    
    logger.info('Batch processing completed', {
      summary,
      performance: profilingResults.summary,
      correlationId
    });
    
    return {
      success: true,
      summary,
      results,
      errors,
      performance: profilingResults,
      correlationId
    };
    
  } catch (error) {
    OOXMLErrorLogger.logError(error, { operation: 'batch_processing' });
    throw error;
  } finally {
    Correlation.end(correlationId);
  }
}

/**
 * Example 4: Comprehensive System Health Check
 * 
 * Production-grade system monitoring and health verification.
 */
async function example4_SystemHealthMonitoring() {
  const logger = Observability.logger('SystemHealthExample');
  const correlationId = Correlation.start('system-health-check');
  
  try {
    logger.info('Starting comprehensive system health check', { correlationId });
    
    // Setup custom health checks
    const healthMonitor = Observability.health;
    
    // Register additional health checks
    healthMonitor.register('file_access', async () => {
      // Test file access
      const testFiles = DriveApp.getFilesByName('test-file.pptx');
      return { canAccessFiles: true, testFileFound: testFiles.hasNext() };
    }, { timeout: 3000 });
    
    healthMonitor.register('processing_capability', async () => {
      // Test basic processing capability
      const testOps = [{ type: 'replaceText', find: 'test', replace: 'test' }];
      
      // This would need a small test file
      const testResult = { canProcess: true, testOperations: testOps.length };
      return testResult;
    }, { timeout: 10000 });
    
    healthMonitor.register('correlation_tracking', () => {
      // Test correlation system
      const testCorrelationId = Correlation.getId();
      return { 
        correlationActive: !!testCorrelationId,
        currentCorrelationId: testCorrelationId 
      };
    });
    
    // Run all health checks
    const healthResults = await logger.operation('run_health_checks', async () => {
      return await healthMonitor.runAll();
    });
    
    // Get system metrics
    const systemMetrics = logger.operation('collect_metrics', () => {
      return Observability.getSystemOverview();
    });
    
    // Check deployment status
    const deploymentStatus = logger.operation('check_deployment', () => {
      return OOXMLDeployment.getDeploymentStatus();
    });
    
    // Generate health report
    const healthReport = {
      timestamp: new Date().toISOString(),
      correlationId,
      overall: {
        healthy: healthResults.healthy && deploymentStatus.deployed,
        score: this._calculateHealthScore(healthResults, deploymentStatus),
        issues: this._identifyHealthIssues(healthResults, deploymentStatus)
      },
      components: {
        deployment: deploymentStatus,
        health: healthResults,
        metrics: systemMetrics
      },
      recommendations: this._generateHealthRecommendations(healthResults, deploymentStatus)
    };
    
    // Log health status
    if (healthReport.overall.healthy) {
      logger.info('System health check passed', { healthReport });
    } else {
      logger.error('System health check failed', { 
        healthReport,
        issues: healthReport.overall.issues 
      });
    }
    
    return healthReport;
    
  } catch (error) {
    OOXMLErrorLogger.logError(error, { operation: 'system_health_monitoring' });
    throw error;
  } finally {
    Correlation.end(correlationId);
  }
}

/**
 * Example 5: Complete Production Workflow
 * 
 * End-to-end example combining all production features.
 */
async function example5_CompleteProductionWorkflow() {
  const logger = Observability.logger('ProductionWorkflowExample');
  const correlationId = Correlation.start('complete-production-workflow');
  
  try {
    logger.info('Starting complete production workflow', { correlationId });
    
    // Phase 1: System Verification
    const healthCheck = await logger.operation('system_verification', async () => {
      const health = await example4_SystemHealthMonitoring();
      if (!health.overall.healthy) {
        throw OOXMLErrorCodes.create(
          'S015',
          'System health check failed',
          { healthScore: health.overall.score },
          correlationId
        );
      }
      return health;
    });
    
    // Phase 2: Brand Compliance Validation
    const complianceResult = await logger.operation('brand_compliance', async () => {
      return await example2_BrandComplianceDSL();
    });
    
    // Phase 3: Batch Processing (if compliance passed)
    let batchResult = null;
    if (complianceResult.analysis.passed) {
      batchResult = await logger.operation('batch_processing', async () => {
        return await example3_BatchProcessingWithMonitoring();
      });
    } else {
      logger.warn('Skipping batch processing due to compliance failures', {
        complianceScore: complianceResult.analysis.score
      });
    }
    
    // Phase 4: Generate comprehensive report
    const finalReport = logger.operation('generate_report', () => {
      return {
        workflow: 'complete-production',
        timestamp: new Date().toISOString(),
        correlationId,
        phases: {
          systemHealth: {
            passed: healthCheck.overall.healthy,
            score: healthCheck.overall.score
          },
          brandCompliance: {
            passed: complianceResult.analysis.passed,
            score: complianceResult.analysis.score,
            violations: complianceResult.analysis.totalViolations,
            autoFixed: complianceResult.analysis.autoFixed
          },
          batchProcessing: batchResult ? {
            passed: batchResult.success,
            processed: batchResult.summary.successfullyProcessed,
            failed: batchResult.summary.failed,
            successRate: batchResult.summary.successRate
          } : { skipped: true }
        },
        overall: {
          success: healthCheck.overall.healthy && 
                   complianceResult.analysis.passed && 
                   (batchResult?.success !== false),
          recommendations: this._generateWorkflowRecommendations({
            healthCheck,
            complianceResult,
            batchResult
          })
        }
      };
    });
    
    logger.info('Complete production workflow finished', { finalReport });
    
    return finalReport;
    
  } catch (error) {
    OOXMLErrorLogger.logError(error, { operation: 'complete_production_workflow' });
    throw error;
  } finally {
    Correlation.end(correlationId);
  }
}

// UTILITY HELPER METHODS

function _groupViolationsByCategory(violations) {
  const categories = {};
  for (const violation of violations) {
    const category = violation.category || 'general';
    if (!categories[category]) {
      categories[category] = [];
    }
    categories[category].push(violation);
  }
  return categories;
}

function _calculateHealthScore(healthResults, deploymentStatus) {
  let score = 0;
  let maxScore = 0;
  
  // Deployment status (30 points)
  maxScore += 30;
  if (deploymentStatus.deployed && deploymentStatus.health?.available) {
    score += 30;
  } else if (deploymentStatus.deployed) {
    score += 15;
  }
  
  // Health checks (70 points)
  maxScore += 70;
  if (healthResults.healthy) {
    score += 70;
  } else {
    score += Math.round((healthResults.passedChecks / healthResults.totalChecks) * 70);
  }
  
  return Math.round((score / maxScore) * 100);
}

function _identifyHealthIssues(healthResults, deploymentStatus) {
  const issues = [];
  
  if (!deploymentStatus.deployed) {
    issues.push('Service not deployed');
  }
  
  if (!deploymentStatus.health?.available) {
    issues.push('Service health check failed');
  }
  
  for (const check of healthResults.checks) {
    if (!check.healthy) {
      issues.push(`Health check failed: ${check.name}`);
    }
  }
  
  return issues;
}

function _generateHealthRecommendations(healthResults, deploymentStatus) {
  const recommendations = [];
  
  if (!deploymentStatus.deployed) {
    recommendations.push('Deploy the OOXML JSON service using OOXMLDeployment.initAndDeploy()');
  }
  
  if (deploymentStatus.deployed && !deploymentStatus.health?.available) {
    recommendations.push('Check Cloud Run logs and service configuration');
  }
  
  const failedChecks = healthResults.checks.filter(c => !c.healthy);
  if (failedChecks.length > 0) {
    recommendations.push(`Review failed health checks: ${failedChecks.map(c => c.name).join(', ')}`);
  }
  
  return recommendations;
}

function _generateWorkflowRecommendations(results) {
  const recommendations = [];
  
  if (!results.healthCheck.overall.healthy) {
    recommendations.push('Address system health issues before running production workflows');
  }
  
  if (!results.complianceResult.analysis.passed) {
    recommendations.push(`Improve brand compliance score (current: ${results.complianceResult.analysis.score}%)`);
  }
  
  if (results.batchResult && results.batchResult.summary.failed > 0) {
    recommendations.push(`Review and fix ${results.batchResult.summary.failed} failed batch processing items`);
  }
  
  if (recommendations.length === 0) {
    recommendations.push('System is operating optimally');
  }
  
  return recommendations;
}

// Export functions for testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    example1_ProductionDeployment,
    example2_BrandComplianceDSL,
    example3_BatchProcessingWithMonitoring,
    example4_SystemHealthMonitoring,
    example5_CompleteProductionWorkflow
  };
}