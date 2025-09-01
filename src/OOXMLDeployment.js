/**
 * OOXMLDeployment - One-Click GAS + Cloud Run Deployment
 * 
 * PURPOSE:
 * Provides complete Google Apps Script-based deployment of the OOXML JSON Cloud Run service.
 * Includes preflight checks (billing, APIs), budget setup, one-click deployment, and monitoring.
 * 
 * ARCHITECTURE:
 * - Preflight UI: Sidebar for billing/API verification and budget setup
 * - One-click deploy: Packages source, uploads to GCS, triggers Cloud Build
 * - Service management: Health checks, configuration, monitoring
 * - Integration: Works seamlessly with OOXMLJsonService
 * 
 * AI CONTEXT:
 * This handles all deployment aspects of the OOXML JSON system. Use showGcpPreflight()
 * to verify setup, then initAndDeploy() for deployment. The CF_BASE property is automatically
 * set for OOXMLJsonService to use.
 */

class OOXMLDeployment {
  
  /**
   * Configuration for deployment
   * @private
   */
  static get _CONFIG() {
    return {
      PROJECT_ID: PropertiesService.getScriptProperties().getProperty('GCP_PROJECT_ID') || '',
      REGION: 'europe-west4',
      SERVICE: 'ooxml-json',
      BUCKET: null, // null => auto: "<PROJECT_ID>-ooxml-src"
      PUBLIC: true,
      RUN_SA: null,
      
      // Budget defaults
      BUDGET_NAME: 'OOXML Helper Budget',
      BUDGET_CURRENCY: 'EUR',
      BUDGET_AMOUNT_UNITS: '5',
      BUDGET_THRESHOLDS: [0.10, 0.50, 0.90]
    };
  }
  
  /**
   * Required APIs for full deployment
   * @private
   */
  static get _REQUIRED_APIS() {
    return [
      'run.googleapis.com',
      'cloudbuild.googleapis.com',
      'artifactregistry.googleapis.com',
      'iam.googleapis.com',
      'serviceusage.googleapis.com',
      'cloudbilling.googleapis.com',
      'cloudresourcemanager.googleapis.com',
      'billingbudgets.googleapis.com',
      'storage.googleapis.com'
    ];
  }
  
  /**
   * Show GCP preflight sidebar for setup verification
   * 
   * AI_USAGE: Use this to guide users through GCP setup before deployment.
   * Example: OOXMLDeployment.showGcpPreflight()
   */
  static showGcpPreflight() {
    var html = HtmlService.createHtmlOutput(this._generatePreflightHtml())
      .setTitle('GCP Preflight - OOXML JSON Setup')
      .setWidth(480);
    
    // Use appropriate UI container
    try {
      SpreadsheetApp.getUi().showSidebar(html);
    } catch (e) {
      try {
        DocumentApp.getUi().showSidebar(html);
      } catch (e) {
        // Fallback for standalone script
        var modal = HtmlService.createHtmlOutput(this._generatePreflightHtml())
          .setTitle('GCP Preflight')
          .setWidth(500)
          .setHeight(600);
        Logger.log('Open this URL in browser for preflight checks: ' + 
                  'https://script.google.com/your-script-id/exec');
      }
    }
  }
  
  /**
   * Initialize and deploy the OOXML JSON Cloud Run service
   * 
   * @param {Object} options - Deployment options
   * @param {string} options.projectId - Override PROJECT_ID
   * @param {string} options.region - Override REGION
   * @param {boolean} options.skipBillingCheck - Skip billing verification
   * @returns {Promise<string>} Deployed service URL
   * @throws {Error} If deployment fails
   * 
   * AI_USAGE: Use this for one-click deployment after preflight checks.
   * Example: var url = await OOXMLDeployment.initAndDeploy()
   */
  static async initAndDeploy(options = {}) {
    var config = { ...this._CONFIG, ...options };
    var { PROJECT_ID: projectId, REGION: region, SERVICE: service } = config;
    
    if (!projectId) {
      throw new Error('Set GCP_PROJECT_ID in Script Properties or pass projectId option');
    }
    
    try {
      this._log(`🚀 Starting deployment for project: ${projectId}`);
      
      // Billing verification
      if (!options.skipBillingCheck) {
        var billing = this._checkBilling(projectId);
        if (!billing.enabled) {
          throw new Error('Billing not enabled. Use showGcpPreflight() first.');
        }
        this._log(`✅ Billing verified: ${billing.accountId}`);
      }
      
      // Enable required APIs
      await this._enableRequiredApis(projectId);
      this._log('✅ APIs enabled');
      
      // Ensure source bucket exists
      var bucket = config.BUCKET || `${projectId}-ooxml-src`;
      await this._ensureBucket(projectId, bucket, region);
      this._log(`✅ Bucket ready: ${bucket}`);
      
      // Package and upload source
      var sourceZip = this._makeSourceZip();
      var objectPath = `src/${Date.now()}/${sourceZip.name}`;
      var uploadResult = await this._uploadToGCS(projectId, bucket, objectPath, sourceZip.bytes);
      this._log(`✅ Source uploaded: gs://${bucket}/${objectPath}`);
      
      // Start Cloud Build deployment
      var buildId = await this._startCloudBuildDeploy(
        projectId, region, bucket, objectPath, service, config.PUBLIC, config.RUN_SA
      );
      this._log(`🔨 Cloud Build started: ${buildId}`);
      
      // Wait for build completion
      await this._waitForBuild(projectId, buildId, 300);
      this._log('✅ Build completed successfully');
      
      // Get deployed service URL
      var serviceUrl = await this._getRunUrl(projectId, region, service);
      
      // Store service URL for OOXMLJsonService
      PropertiesService.getScriptProperties().setProperties({
        'CF_BASE': serviceUrl,
        'OOXML_SERVICE_DEPLOYED': new Date().toISOString(),
        'OOXML_SERVICE_PROJECT': projectId,
        'OOXML_SERVICE_REGION': region
      });
      
      this._log(`🎉 Deployment complete: ${serviceUrl}`);
      
      // Verify service health
      var health = await this._verifyDeployment(serviceUrl);
      if (!health.available) {
        throw new Error(`Service deployed but health check failed: ${health.error}`);
      }
      
      this._log('✅ Service health verified');
      return serviceUrl;
      
    } catch (error) {
      this._log(`❌ Deployment failed: ${error.message}`);
      throw error;
    }
  }
  
  /**
   * Check deployment status and service health
   * 
   * @returns {Object} Deployment status information
   * 
   * AI_USAGE: Use this to check if service is properly deployed and running.
   * Example: var status = OOXMLDeployment.getDeploymentStatus()
   */
  static getDeploymentStatus() {
    var properties = PropertiesService.getScriptProperties().getProperties();
    
    var status = {
      deployed: !!properties.CF_BASE,
      serviceUrl: properties.CF_BASE || null,
      deployedAt: properties.OOXML_SERVICE_DEPLOYED || null,
      project: properties.OOXML_SERVICE_PROJECT || null,
      region: properties.OOXML_SERVICE_REGION || null,
      health: null
    };
    
    // Check service health if deployed
    if (status.deployed) {
      try {
        var response = UrlFetchApp.fetch(`${status.serviceUrl}/ping`, {
          method: 'GET',
          muteHttpExceptions: true,
          timeout: 5000
        });
        
        var isHealthy = response.getResponseCode() < 400;
        status.health = {
          available: isHealthy,
          statusCode: response.getResponseCode(),
          response: isHealthy ? JSON.parse(response.getContentText()) : null,
          checkedAt: new Date().toISOString()
        };
      } catch (error) {
        status.health = {
          available: false,
          error: error.message,
          checkedAt: new Date().toISOString()
        };
      }
    }
    
    return status;
  }
  
  /**
   * Clean up deployment resources
   * 
   * @param {Object} options - Cleanup options
   * @param {boolean} options.removeService - Delete Cloud Run service
   * @param {boolean} options.removeBucket - Delete source bucket
   * @param {boolean} options.clearProperties - Clear Script Properties
   * @returns {Object} Cleanup results
   * 
   * AI_USAGE: Use this to clean up deployment resources.
   * Example: var result = OOXMLDeployment.cleanup({removeService: true})
   */
  static cleanup(options = {}) {
    var opts = {
      removeService: false,
      removeBucket: false,
      clearProperties: true,
      ...options
    };
    
    var results = {
      serviceRemoved: false,
      bucketRemoved: false,
      propertiesCleared: false,
      errors: []
    };
    
    try {
      var properties = PropertiesService.getScriptProperties().getProperties();
      var projectId = properties.OOXML_SERVICE_PROJECT;
      var region = properties.OOXML_SERVICE_REGION;
      var bucket = this._CONFIG.BUCKET || `${projectId}-ooxml-src`;
      
      // Remove Cloud Run service
      if (opts.removeService && projectId && region) {
        try {
          this._deleteCloudRunService(projectId, region, this._CONFIG.SERVICE);
          results.serviceRemoved = true;
          this._log('✅ Cloud Run service removed');
        } catch (error) {
          results.errors.push(`Service removal failed: ${error.message}`);
        }
      }
      
      // Remove source bucket
      if (opts.removeBucket && projectId) {
        try {
          this._deleteBucket(projectId, bucket);
          results.bucketRemoved = true;
          this._log('✅ Source bucket removed');
        } catch (error) {
          results.errors.push(`Bucket removal failed: ${error.message}`);
        }
      }
      
      // Clear Script Properties
      if (opts.clearProperties) {
        PropertiesService.getScriptProperties().deleteProperty('CF_BASE');
        PropertiesService.getScriptProperties().deleteProperty('OOXML_SERVICE_DEPLOYED');
        PropertiesService.getScriptProperties().deleteProperty('OOXML_SERVICE_PROJECT');
        PropertiesService.getScriptProperties().deleteProperty('OOXML_SERVICE_REGION');
        results.propertiesCleared = true;
        this._log('✅ Script Properties cleared');
      }
      
    } catch (error) {
      results.errors.push(`Cleanup failed: ${error.message}`);
    }
    
    return results;
  }
  
  // PREFLIGHT CHECK METHODS
  
  /**
   * Check billing status for project
   * @private
   */
  static _checkBilling(projectId = this._CONFIG.PROJECT_ID) {
    try {
      var response = this._gapi(
        `https://cloudbilling.googleapis.com/v1/projects/${projectId}/billingInfo`,
        'GET'
      );
      
      return {
        enabled: !!response.billingEnabled,
        accountName: response.billingAccountName || '',
        accountId: (response.billingAccountName || '').replace('billingAccounts/', '') || null
      };
    } catch (error) {
      return { enabled: false, error: error.message };
    }
  }
  
  /**
   * Check if specific API is enabled
   * @private
   */
  static _checkApi(api, projectId = this._CONFIG.PROJECT_ID) {
    try {
      var response = this._gapi(
        `https://serviceusage.googleapis.com/v1/projects/${projectId}/services/${api}`,
        'GET',
        null,
        true
      );
      
      return {
        api,
        enabled: response && response.state === 'ENABLED'
      };
    } catch (error) {
      return { api, enabled: false, error: error.message };
    }
  }
  
  /**
   * Enable all required APIs
   * @private
   */
  static async _enableRequiredApis(projectId = this._CONFIG.PROJECT_ID) {
    var payload = { serviceIds: this._REQUIRED_APIS };
    
    return this._gapi(
      `https://serviceusage.googleapis.com/v1/projects/${projectId}/services:batchEnable`,
      'POST',
      payload
    );
  }
  
  /**
   * Ensure budget exists
   * @private
   */
  static _ensureBudget(displayName, currencyCode, amountUnits, thresholds) {
    var billing = this._checkBilling();
    if (!billing.enabled) {
      throw new Error('Billing not enabled on project');
    }
    
    var billingAccountName = billing.accountName;
    var projectNumber = this._getProjectNumber();
    
    var budget = {
      displayName: displayName || this._CONFIG.BUDGET_NAME,
      budgetFilter: { projects: [`projects/${projectNumber}`] },
      amount: {
        specifiedAmount: {
          currencyCode: currencyCode || this._CONFIG.BUDGET_CURRENCY,
          units: String(amountUnits || this._CONFIG.BUDGET_AMOUNT_UNITS)
        }
      },
      thresholdRules: (thresholds || this._CONFIG.BUDGET_THRESHOLDS).map(p => ({
        thresholdPercent: p
      })),
      allUpdatesRule: { disableDefaultIamRecipients: false }
    };
    
    // Check if budget exists
    var existing = this._listBudgets(billingAccountName).find(
      b => b.displayName === budget.displayName
    );
    
    if (existing) {
      // Update existing budget
      var updateMask = 'displayName,budgetFilter,amount,thresholdRules,allUpdatesRule';
      var url = `https://billingbudgets.googleapis.com/v1/${existing.name}?updateMask=${encodeURIComponent(updateMask)}`;
      return this._gapi(url, 'PATCH', budget);
    } else {
      // Create new budget
      var url = `https://billingbudgets.googleapis.com/v1/${billingAccountName}/budgets`;
      return this._gapi(url, 'POST', budget);
    }
  }
  
  // CLOUD DEPLOYMENT METHODS
  
  /**
   * Ensure GCS bucket exists
   * @private
   */
  static async _ensureBucket(projectId, bucketName, location) {
    // Check if bucket exists
    var exists = this._gapi(
      `https://storage.googleapis.com/storage/v1/b/${encodeURIComponent(bucketName)}`,
      'GET',
      null,
      true
    );
    
    if (exists) {
      this._log(`📦 Bucket exists: ${bucketName}`);
      return;
    }
    
    // Create bucket
    var bucketConfig = {
      name: bucketName,
      location: location,
      iamConfiguration: {
        uniformBucketLevelAccess: { enabled: true }
      }
    };
    
    this._gapi(
      `https://storage.googleapis.com/storage/v1/b?project=${encodeURIComponent(projectId)}`,
      'POST',
      bucketConfig
    );
    
    this._log(`📦 Bucket created: ${bucketName}`);
  }
  
  /**
   * Upload source zip to GCS
   * @private
   */
  static async _uploadToGCS(projectId, bucket, objectName, bytes) {
    var url = `https://storage.googleapis.com/upload/storage/v1/b/${encodeURIComponent(bucket)}/o?uploadType=media&name=${encodeURIComponent(objectName)}`;
    
    var response = this._gapiRaw(url, 'POST', bytes, {
      'Content-Type': 'application/zip'
    });
    
    var metadata = JSON.parse(response);
    return { bucket: metadata.bucket, name: metadata.name };
  }
  
  /**
   * Start Cloud Build deployment
   * @private
   */
  static async _startCloudBuildDeploy(projectId, region, bucket, objectName, service, allowUnauth, runSaEmail) {
    var args = [
      'run', 'deploy', service,
      '--source', '.',
      '--region', region,
      '--project', projectId,
      '--quiet'
    ];
    
    if (allowUnauth) {
      args.push('--allow-unauthenticated');
    }
    
    if (runSaEmail) {
      args.push('--service-account', runSaEmail);
    }
    
    // Set environment variables for the service
    var envVars = [
      `BUCKET=${this._CONFIG.BUCKET || (projectId + '-ooxml-src')}`,
      `REGION=${region}`
    ];
    args.push('--set-env-vars', envVars.join(','));
    
    var buildConfig = {
      source: {
        storageSource: { bucket, object: objectName }
      },
      steps: [{
        name: 'gcr.io/cloud-builders/gcloud',
        args
      }],
      options: {
        logging: 'CLOUD_LOGGING_ONLY'
      }
    };
    
    var response = this._gapi(
      `https://cloudbuild.googleapis.com/v1/projects/${projectId}/builds`,
      'POST',
      buildConfig
    );
    
    return response.id;
  }
  
  /**
   * Wait for Cloud Build to complete
   * @private
   */
  static async _waitForBuild(projectId, buildId, timeoutSec = 300) {
    var url = `https://cloudbuild.googleapis.com/v1/projects/${projectId}/builds/${buildId}`;
    var startTime = Date.now();
    
    while (true) {
      var build = this._gapi(url, 'GET');
      
      if (build.logUrl) {
        this._log(`📋 Build logs: ${build.logUrl}`);
      }
      
      var finalStates = ['SUCCESS', 'FAILURE', 'CANCELLED', 'TIMEOUT'];
      if (finalStates.includes(build.status)) {
        if (build.status !== 'SUCCESS') {
          throw new Error(`Build ${build.status}: ${build.statusDetail || 'Unknown error'}`);
        }
        return;
      }
      
      var elapsed = (Date.now() - startTime) / 1000;
      if (elapsed > timeoutSec) {
        this._log(`⏰ Build timeout after ${elapsed}s. Re-run _waitForBuild later with buildId: ${buildId}`);
        return;
      }
      
      this._log(`⏳ Build ${build.status}... (${elapsed.toFixed(0)}s)`);
      Utilities.sleep(5000);
    }
  }
  
  /**
   * Get Cloud Run service URL
   * @private
   */
  static async _getRunUrl(projectId, region, service) {
    var url = `https://run.googleapis.com/v2/projects/${projectId}/locations/${region}/services/${service}`;
    var response = this._gapi(url, 'GET');
    return response.uri;
  }
  
  /**
   * Verify deployment by checking service health
   * @private
   */
  static async _verifyDeployment(serviceUrl) {
    try {
      var response = UrlFetchApp.fetch(`${serviceUrl}/ping`, {
        method: 'GET',
        muteHttpExceptions: true,
        timeout: 10000
      });
      
      var isHealthy = response.getResponseCode() < 400;
      var result = isHealthy ? JSON.parse(response.getContentText()) : null;
      
      return {
        available: isHealthy,
        statusCode: response.getResponseCode(),
        version: result?.version,
        timestamp: new Date().toISOString()
      };
    } catch (error) {
      return {
        available: false,
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }
  }
  
  // UTILITY METHODS
  
  /**
   * Create source package for Cloud Run
   * @private
   */
  static _makeSourceZip() {
    var files = [
      Utilities.newBlob(this._getServiceIndexMjs(), 'text/plain', 'index.mjs'),
      Utilities.newBlob(this._getServicePackageJson(), 'application/json', 'package.json'),
      Utilities.newBlob('node_modules\n.git\n*.log\n', 'text/plain', '.gcloudignore'),
      Utilities.newBlob('# OOXML JSON Service\nDeployed from Google Apps Script\n', 'text/markdown', 'README.md')
    ];
    
    var zip = Utilities.zip(files, 'ooxml-json-source.zip');
    return { name: zip.getName(), bytes: zip.getBytes() };
  }
  
  /**
   * Generate package.json for service
   * @private
   */
  static _getServicePackageJson() {
    return JSON.stringify({
      type: 'module',
      dependencies: {
        '@google-cloud/storage': '^7.12.1',
        express: '^4.19.2',
        fflate: '^0.8.2'
      },
      engines: { node: '>=18' },
      scripts: { start: 'node index.mjs' }
    }, null, 2);
  }
  
  /**
   * Generate index.mjs for service (from the guide)
   * @private
   */
  static _getServiceIndexMjs() {
    return `import express from "express";
import { unzipSync, zipSync, base64Decode, base64Encode } from "fflate";
import { Storage } from "@google-cloud/storage";
import crypto from "node:crypto";

var app = express();
app.use(express.json({ limit: "100mb" }));
var storage = new Storage();
var BUCKET = process.env.BUCKET;
var REGION = process.env.REGION || "europe-west4";

var toU8 = (b64) => base64Decode(b64.replace(/-/g, "+").replace(/_/g, "/"));
var toB64 = (u8) => base64Encode(u8);
var isXml = (p) => /\\.xml$/i.test(p);

// Sessions
app.post("/session/new", async (req, res) => {
  try {
    if (!BUCKET) return res.status(500).json({ error: "BUCKET not set" });
    var sid = crypto.randomUUID();
    var inKey = \`\${sid}/in.pptx\`;
    var outKey = \`\${sid}/out.pptx\`;

    var input = storage.bucket(BUCKET).file(inKey);
    var output = storage.bucket(BUCKET).file(outKey);

    var [uploadUrl] = await input.getSignedUrl({ 
      version: "v4", action: "write", expires: Date.now() + 15*60*1000,
      contentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" 
    });
    var [downloadUrl] = await output.getSignedUrl({ 
      version: "v4", action: "read", expires: Date.now() + 60*60*1000 
    });

    res.json({ 
      sessionId: sid, 
      gcsIn: \`gs://\${BUCKET}/\${inKey}\`, 
      gcsOut: \`gs://\${BUCKET}/\${outKey}\`, 
      uploadUrl, 
      downloadUrl 
    });
  } catch (e) { 
    res.status(500).json({ error: String(e?.message || e) }); 
  }
});

// Unwrap
app.post("/unwrap", async (req, res) => {
  try {
    var { zipB64, gcsIn } = req.body || {}; 
    let buf;
    if (zipB64) buf = toU8(zipB64);
    else if (gcsIn) { 
      var [, , bucket, ...rest] = gcsIn.split("/"); 
      var key = rest.join("/"); 
      var [data] = await storage.bucket(bucket).file(key).download(); 
      buf = data; 
    }
    else return res.status(400).json({ error: "zipB64 or gcsIn required" });

    var bag = unzipSync(buf, { filter: () => true });
    var paths = Object.keys(bag).sort();
    var entries = paths.map((path) => 
      isXml(path) 
        ? { path, type: "xml", text: new TextDecoder().decode(bag[path]) } 
        : { path, type: "bin" }
    );
    res.json({ version: "ooxml-json-1", kind: detectKind(entries), entries });
  } catch (e) { 
    res.status(500).json({ error: String(e?.message || e) }); 
  }
});

// Rewrap
app.post("/rewrap", async (req, res) => {
  try {
    var { manifest, gcsIn, gcsOut } = req.body || {};
    if (!manifest || !Array.isArray(manifest.entries)) {
      return res.status(400).json({ error: "manifest.entries missing" });
    }
    if (!gcsIn && !manifest.zipB64) {
      return res.status(400).json({ error: "gcsIn or manifest.zipB64 required" });
    }

    let raw;
    if (gcsIn) { 
      var [, , bucket, ...rest] = gcsIn.split("/"); 
      var key = rest.join("/"); 
      var [data] = await storage.bucket(bucket).file(key).download(); 
      raw = data; 
    } else { 
      raw = toU8(manifest.zipB64); 
    }

    var bag = unzipSync(raw, { filter: () => true });
    var enc = new TextEncoder();
    for (var e of manifest.entries) {
      if (!e || !e.path) continue;
      if (e.type === "xml" && typeof e.text === "string") {
        bag[e.path] = enc.encode(e.text);
      } else if (e.type === "bin" && e.dataB64) {
        bag[e.path] = toU8(e.dataB64);
      }
    }
    var ordered = Object.fromEntries(Object.keys(bag).sort().map(p => [p, bag[p]]));
    var outZip = zipSync(ordered, { level: 6, zip64: true });

    if (gcsOut) {
      var [, , outBucket, ...outRest] = gcsOut.split("/"); 
      var outKey = outRest.join("/");
      await storage.bucket(outBucket).file(outKey).save(Buffer.from(outZip), { 
        contentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" 
      });
      return res.json({ gcsOut });
    }
    res.json({ zipB64: base64Encode(outZip) });
  } catch (e) { 
    res.status(500).json({ error: String(e?.message || e) }); 
  }
});

// Process
app.post("/process", async (req, res) => {
  try {
    var { zipB64, gcsIn, gcsOut, ops } = req.body || {};
    if (!Array.isArray(ops)) return res.status(400).json({ error: "ops[] required" });
    if (!zipB64 && !gcsIn) return res.status(400).json({ error: "zipB64 or gcsIn required" });

    let buf;
    if (zipB64) buf = toU8(zipB64); 
    else { 
      var [, , bucket, ...rest] = gcsIn.split("/"); 
      var key = rest.join("/"); 
      var [data] = await storage.bucket(bucket).file(key).download(); 
      buf = data; 
    }

    var bag = unzipSync(buf, { filter: () => true });
    var enc = new TextEncoder(), dec = new TextDecoder();
    var getText = (p) => dec.decode(bag[p]);
    var setText = (p, s) => { bag[p] = enc.encode(s); };
    var setBin = (p, b64) => { bag[p] = toU8(b64); };
    var del = (p) => { delete bag[p]; };
    var exists = (p) => Object.prototype.hasOwnProperty.call(bag, p);

    var ctPath = "[Content_Types].xml";
    let ctXml = exists(ctPath) ? getText(ctPath) : 
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>';
    var report = { replaced: 0, upserted: 0, removed: 0, renamed: 0, errors: [] };

    for (var op of ops) {
      try {
        switch (op.type) {
          case "replaceText": {
            var scope = op.scope || "";
            var find = String(op.find);
            var replace = String(op.replace || "");
            var re = op.regex ? new RegExp(op.find, op.flags || "g") : 
              new RegExp(find.replace(/[\\\.\*\+\?\^\$\{\}\(\)\|\[\]]/g, '\\$&'), "g");
            for (var p of Object.keys(bag)) {
              if (!/\\.xml$/i.test(p)) continue; 
              if (scope && !p.startsWith(scope)) continue;
              var before = getText(p); 
              var after = before.replace(re, replace);
              if (after !== before) { setText(p, after); report.replaced++; }
            }
            break;
          }
          case "upsertPart": {
            if (!op.path) throw new Error("upsertPart.path missing");
            if (/\\.xml$/i.test(op.path)) { 
              setText(op.path, String(op.text ?? "")); 
              if (op.contentType) ctXml = ensureOverride(ctXml, op.path, op.contentType); 
            } else { 
              if (!op.dataB64) throw new Error("upsertPart.dataB64 missing for binary"); 
              setBin(op.path, op.dataB64); 
              if (op.contentType) ctXml = ensureOverride(ctXml, op.path, op.contentType); 
            }
            report.upserted++; 
            break;
          }
          case "removePart": { 
            if (!op.path) throw new Error("removePart.path missing"); 
            if (exists(op.path)) { del(op.path); report.removed++; } 
            ctXml = removeOverride(ctXml, op.path); 
            break; 
          }
          case "renamePart": { 
            if (!op.from || !op.to) throw new Error("renamePart.from/to missing"); 
            if (exists(op.from)) { 
              bag[op.to] = bag[op.from]; 
              delete bag[op.from]; 
              if (op.contentType) { 
                ctXml = removeOverride(ctXml, op.from); 
                ctXml = ensureOverride(ctXml, op.to, op.contentType); 
              } 
              report.renamed++; 
            } 
            break; 
          }
          default: 
            throw new Error("unknown op.type: " + op.type);
        }
      } catch (e) { 
        report.errors.push({ op, message: String(e?.message || e) }); 
      }
    }

    bag[ctPath] = enc.encode(ctXml);
    var ordered = Object.fromEntries(Object.keys(bag).sort().map(p => [p, bag[p]]));
    var outZip = zipSync(ordered, { level: 6, zip64: true });

    if (gcsOut) { 
      var [, , b, ...rest] = gcsOut.split("/"); 
      var k = rest.join("/"); 
      await storage.bucket(b).file(k).save(Buffer.from(outZip), { 
        contentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation" 
      }); 
      return res.json({ gcsOut, report }); 
    }
    res.json({ zipB64: toB64(outZip), report });
  } catch (e) { 
    res.status(500).json({ error: String(e?.message || e) }); 
  }
});

app.get("/ping", (_req, res) => res.json({ ok: true, version: "ooxml-json-1" }));

function detectKind(entries) {
  var ct = entries.find(e => e.path === "[Content_Types].xml" && e.type === "xml"); 
  if (!ct) return "unknown";
  var has = (s) => ct.text.includes(\`ContentType="\${s}"\`);
  if (has("application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml")) return "pptx";
  if (has("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")) return "docx";
  if (has("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")) return "xlsx";
  if (has("application/vnd.openxmlformats-officedocument.theme+xml")) return "thmx";
  return "unknown";
}

function ensureOverride(ctXml, partName, contentType) {
  var pn = partName.startsWith('/') ? partName : '/' + partName;
  var escapedPn = pn.replace(/[\\\.\*\+\?\^\$\{\}\(\)\|\[\]]/g,'\\\\$&');
  var rx = new RegExp('<Override\\\\b[^>]*PartName="' + escapedPn + '"[^>]*/>', 'i');
  if (rx.test(ctXml)) return ctXml;
  return ctXml.replace(/<\\/Types>\\s*$/i, '<Override PartName="' + pn + '" ContentType="' + contentType + '"/><\\/Types>');
}

function removeOverride(ctXml, partName) {
  var pn = partName.startsWith('/') ? partName : '/' + partName;
  var escapedPn = pn.replace(/[\\\.\*\+\?\^\$\{\}\(\)\|\[\]]/g,'\\\\$&');
  var rx = new RegExp('<Override\\\\b[^>]*PartName="' + escapedPn + '"[^>]*/>\\s*', 'i');
  return ctXml.replace(rx, '');
}

var port = process.env.PORT || 8080;
app.listen(port, () => console.log("OOXML JSON Service listening on", port));
`;
  }
  
  /**
   * Generate preflight HTML interface
   * @private
   */
  static _generatePreflightHtml() {
    var apis = this._REQUIRED_APIS.map(api => 
      `<li><code>${api}</code> — <span data-api="${api}">checking…</span></li>`
    ).join('');
    
    return `
    <style>
      body { font: 13px/1.4 system-ui, Segoe UI, Roboto, Arial; margin: 12px; }
      h3 { margin: 8px 0 6px; }
      button { padding: 8px 10px; border-radius: 8px; border: 1px solid #dadce0; background: #fff; cursor: pointer; margin: 4px 2px; }
      button.primary { background: #0b57d0; color: #fff; border-color: #0b57d0; }
      .ok { color: #137333; }
      .bad { color: #c5221f; }
      .warn { color: #b06000; }
      input[type=text] { width: 100%; padding: 6px; border: 1px solid #dadce0; border-radius: 6px; }
      .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin: 8px 0; }
      .status-row { margin: 8px 0; }
      .section { margin: 16px 0; }
    </style>
    
    <h3>🚀 OOXML JSON Helper - GCP Setup</h3>
    
    <div class="section">
      <div>Project: <code id="projectId">${this._CONFIG.PROJECT_ID || 'Not set'}</code></div>
      <div class="status-row">Billing: <span id="billingStatus">checking…</span></div>
    </div>
    
    <div class="section">
      <h4>📋 Required APIs</h4>
      <ul>${apis}</ul>
      <button id="enableApis">Enable Missing APIs</button>
    </div>
    
    <div class="section">
      <h4>💰 Budget Setup</h4>
      <div class="grid">
        <div><label>Name</label><input id="budgetName" type="text" value="${this._CONFIG.BUDGET_NAME}"></div>
        <div><label>Currency</label><input id="currency" type="text" value="${this._CONFIG.BUDGET_CURRENCY}"></div>
        <div><label>Amount</label><input id="amount" type="text" value="${this._CONFIG.BUDGET_AMOUNT_UNITS}"></div>
        <div><label>Thresholds (%)</label><input id="thresholds" type="text" value="10,50,90"></div>
      </div>
      <button id="createBudget">Create/Update Budget</button>
    </div>
    
    <div class="section">
      <button class="primary" id="recheck">🔄 Re-check Status</button>
      <button class="primary" id="deploy" disabled>🚀 Deploy Service</button>
    </div>
    
    <div><small id="msg"></small></div>
    
    <script>
      function setMsg(text, className) {
        var el = document.getElementById('msg');
        el.textContent = text || '';
        el.className = className || '';
      }
      
      function setBilling(status, className) {
        var el = document.getElementById('billingStatus');
        el.textContent = status;
        el.className = className || '';
      }
      
      function setApi(api, status, className) {
        var el = document.querySelector('[data-api="' + api + '"]');
        if (el) {
          el.textContent = status;
          el.className = className || '';
        }
      }
      
      function updateDeployButton() {
        var billingOk = document.getElementById('billingStatus').className === 'ok';
        var deployBtn = document.getElementById('deploy');
        deployBtn.disabled = !billingOk;
      }
      
      function checkStatus() {
        setMsg('');
        setBilling('checking…', '');
        
        // Check billing
        google.script.run
          .withSuccessHandler(function(result) {
            var status = result.enabled ? 
              \`ENABLED (\${result.accountId})\` : 'NOT ENABLED';
            setBilling(status, result.enabled ? 'ok' : 'bad');
            updateDeployButton();
          })
          ._checkBilling();
        
        // Check APIs
        ${JSON.stringify(this._REQUIRED_APIS)}.forEach(function(api) {
          setApi(api, 'checking…', '');
          google.script.run
            .withSuccessHandler(function(result) {
              setApi(result.api, result.enabled ? 'ENABLED' : 'DISABLED', 
                    result.enabled ? 'ok' : 'bad');
            })
            ._checkApi(api);
        });
      }
      
      // Event handlers
      document.getElementById('enableApis').onclick = function() {
        setMsg('Enabling APIs… (this may take 1-2 minutes)');
        google.script.run
          .withSuccessHandler(function() {
            setMsg('APIs enabled successfully!', 'ok');
            setTimeout(checkStatus, 2000);
          })
          .withFailureHandler(function(error) {
            setMsg('API enable failed: ' + error.message, 'bad');
          })
          ._enableRequiredApis();
      };
      
      document.getElementById('createBudget').onclick = function() {
        var name = document.getElementById('budgetName').value.trim();
        var currency = document.getElementById('currency').value.trim();
        var amount = document.getElementById('amount').value.trim();
        var thresholds = document.getElementById('thresholds').value
          .split(',').map(s => parseFloat(s)/100).filter(n => !isNaN(n));
        
        setMsg('Creating/updating budget…');
        google.script.run
          .withSuccessHandler(function(info) {
            setMsg('Budget ready: ' + (info.displayName || name), 'ok');
          })
          .withFailureHandler(function(error) {
            setMsg('Budget failed: ' + (error.message || error), 'bad');
          })
          ._ensureBudget(name, currency, amount, thresholds);
      };
      
      document.getElementById('recheck').onclick = checkStatus;
      
      document.getElementById('deploy').onclick = function() {
        if (confirm('Deploy OOXML JSON service to Cloud Run?')) {
          setMsg('Starting deployment… (this will take several minutes)');
          google.script.run
            .withSuccessHandler(function(url) {
              setMsg('🎉 Deployment successful! Service URL: ' + url, 'ok');
            })
            .withFailureHandler(function(error) {
              setMsg('❌ Deployment failed: ' + (error.message || error), 'bad');
            })
            .initAndDeploy();
        }
      };
      
      // Initial check
      checkStatus();
    </script>`;
  }
  
  // GCP API HELPER METHODS
  
  /**
   * Get project number from project ID
   * @private
   */
  static _getProjectNumber() {
    var response = this._gapi(
      `https://cloudresourcemanager.googleapis.com/v1/projects/${this._CONFIG.PROJECT_ID}`,
      'GET'
    );
    return response.projectNumber;
  }
  
  /**
   * List budgets for billing account
   * @private
   */
  static _listBudgets(billingAccountName) {
    var budgets = [];
    let pageToken = '';
    
    do {
      var url = `https://billingbudgets.googleapis.com/v1/${billingAccountName}/budgets` +
                  (pageToken ? `?pageToken=${encodeURIComponent(pageToken)}` : '');
      
      var response = this._gapi(url, 'GET', null, true) || {};
      (response.budgets || []).forEach(budget => budgets.push(budget));
      pageToken = response.nextPageToken || '';
    } while (pageToken);
    
    return budgets;
  }
  
  /**
   * Delete Cloud Run service
   * @private
   */
  static _deleteCloudRunService(projectId, region, service) {
    var url = `https://run.googleapis.com/v2/projects/${projectId}/locations/${region}/services/${service}`;
    return this._gapi(url, 'DELETE');
  }
  
  /**
   * Delete GCS bucket
   * @private
   */
  static _deleteBucket(projectId, bucketName) {
    // Note: This only works for empty buckets
    var url = `https://storage.googleapis.com/storage/v1/b/${encodeURIComponent(bucketName)}`;
    return this._gapi(url, 'DELETE');
  }
  
  /**
   * Make authenticated GCP API call
   * @private
   */
  static _gapi(url, method = 'GET', body = null, returnNullOn404 = false) {
    var options = {
      method: method.toUpperCase(),
      headers: {
        'Authorization': `Bearer ${ScriptApp.getOAuthToken()}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    };
    
    if (body) {
      options.payload = JSON.stringify(body);
    }
    
    var response = UrlFetchApp.fetch(url, options);
    var statusCode = response.getResponseCode();
    
    if (returnNullOn404 && statusCode === 404) {
      return null;
    }
    
    if (statusCode < 200 || statusCode >= 300) {
      var errorText = response.getContentText();
      throw new Error(`${method} ${url} => ${statusCode}: ${errorText}`);
    }
    
    var responseText = response.getContentText();
    return responseText ? JSON.parse(responseText) : {};
  }
  
  /**
   * Make raw authenticated GCP API call for binary data
   * @private
   */
  static _gapiRaw(url, method = 'POST', bytes = null, extraHeaders = {}) {
    var headers = {
      'Authorization': `Bearer ${ScriptApp.getOAuthToken()}`,
      ...extraHeaders
    };
    
    var options = {
      method: method.toUpperCase(),
      headers,
      muteHttpExceptions: true
    };
    
    if (bytes) {
      options.payload = bytes;
    }
    
    var response = UrlFetchApp.fetch(url, options);
    var statusCode = response.getResponseCode();
    
    if (statusCode < 200 || statusCode >= 300) {
      var errorText = response.getContentText();
      throw new Error(`RAW ${method} ${url} => ${statusCode}: ${errorText}`);
    }
    
    return response.getContentText();
  }
  
  /**
   * Log message with timestamp
   * @private
   */
  static _log(message) {
    var timestamp = new Date().toISOString().substr(11, 8);
    console.log(`[${timestamp}] [OOXMLDeployment] ${message}`);
  }
}

// Export for use in other modules
if (typeof module !== 'undefined' && module.exports) {
  module.exports = OOXMLDeployment;
}