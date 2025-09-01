# PPTX Processing Service Deployment Guide

Complete guide for deploying the OOXML PPTX processing service to Google Cloud free tier from Google Apps Script.

## 📋 Table of Contents

1. [Overview](#overview)
2. [Prerequisites](#prerequisites)
3. [Deployment from Google Apps Script](#deployment-from-google-apps-script)
4. [Alternative: Command Line Deployment](#alternative-command-line-deployment)
5. [Testing & Validation](#testing--validation)
6. [Troubleshooting](#troubleshooting)
7. [Cost & Limits](#cost--limits)

## Overview

This service processes PPTX files by:
- **Unzipping** them to access OOXML structure
- **Modifying** XML/JSON content programmatically
- **Zipping** them back into valid PPTX files

All processing happens on Google Cloud Run (US free tier) and is controlled from Google Apps Script.

## Prerequisites

### Required
- Google Cloud Project with billing enabled (free tier is sufficient)
- Google Apps Script project
- Google Workspace account

### Optional (for local development)
- Node.js 18+ 
- Google Cloud CLI (`gcloud`)
- Playwright (for testing)

## Deployment from Google Apps Script

### Step 1: Set Up Google Apps Script Project

1. **Create or open a Google Apps Script project:**
   ```
   https://script.google.com
   ```

2. **Copy the deployment files to your project:**
   - Copy all files from `src/` folder
   - Copy all files from `lib/` folder
   - Copy `DeployFromGAS.js`

3. **Enable Advanced Google Services:**
   - In Apps Script Editor: Resources → Advanced Google Services
   - Enable: Drive API, Slides API (if using presentations)

### Step 2: Configure Project ID

```javascript
// In Google Apps Script, run:
function setupProjectId() {
  const YOUR_PROJECT_ID = 'your-actual-project-id'; // Change this!
  PropertiesService.getScriptProperties()
    .setProperty('GCP_PROJECT_ID', YOUR_PROJECT_ID);
  console.log('✅ Project ID configured');
}
```

### Step 3: Run Preflight Checks

```javascript
// This opens a sidebar for setup verification
function showPreflightChecks() {
  OOXMLDeployment.showGcpPreflight();
}
```

In the sidebar:
1. ✅ Verify billing is enabled
2. ✅ Enable required APIs
3. ✅ Set up budget alerts (optional)
4. ✅ Click "Deploy" when ready

### Step 4: Deploy to Free Tier

```javascript
// Deploy the service (takes 3-5 minutes)
async function deployToUSFreeTier() {
  const serviceUrl = await OOXMLDeployment.initAndDeploy({
    region: 'us-central1'  // US free tier
  });
  
  console.log('✅ Deployed to:', serviceUrl);
  
  // Save URL for later use
  PropertiesService.getScriptProperties()
    .setProperty('CLOUD_SERVICE_URL', serviceUrl);
}
```

### Step 5: Verify Deployment

```javascript
// Test the deployed service
async function testDeployedService() {
  // Health check
  const health = await OOXMLJsonService.healthCheck();
  console.log('Health:', health);
  
  // Process a test file
  const manifest = await OOXMLJsonService.unwrap('test.pptx');
  console.log('Files extracted:', manifest.entries.length);
}
```

## Alternative: Command Line Deployment

If you prefer using the command line:

### Option 1: Simple Deployment Script

```bash
# Make sure gcloud is installed and authenticated
gcloud auth login
gcloud config set project YOUR_PROJECT_ID

# Run the deployment script
chmod +x deploy-simple-us.sh
./deploy-simple-us.sh
```

### Option 2: Manual Deployment

```bash
# Enable APIs
gcloud services enable run.googleapis.com
gcloud services enable cloudbuild.googleapis.com

# Deploy from source
gcloud run deploy pptx-processor \
  --source . \
  --region us-central1 \
  --allow-unauthenticated \
  --memory 512Mi \
  --timeout 60s \
  --max-instances 3
```

## Testing & Validation

### Local Testing (Optional)

```bash
# Install dependencies
npm install

# Test PPTX processing locally
node test-pptx-mvp.js

# Test cloud service simulation  
node test-cloud-pptx-mvp.js

# Test with Brave browser
npx playwright test --project=Brave
```

### GAS Testing

```javascript
// Complete workflow test
async function testCompleteWorkflow() {
  // 1. Create test presentation
  const testFile = createTestPresentation();
  
  // 2. Unzip (extract)
  const manifest = await OOXMLJsonService.unwrap(testFile.getId());
  
  // 3. Modify
  manifest.entries.forEach(entry => {
    if (entry.type === 'xml') {
      entry.text = entry.text.replace('Test', 'Modified');
    }
  });
  
  // 4. Zip (rebuild)
  const outputId = await OOXMLJsonService.rewrap(manifest);
  
  console.log('✅ Workflow complete:', outputId);
}
```

## Service Endpoints

Once deployed, your service provides:

### Health Check
```bash
curl YOUR_SERVICE_URL/health
# Returns: {"status":"healthy","region":"us-central1"}
```

### Extract (Unzip)
```javascript
POST YOUR_SERVICE_URL/extract
{
  "zipB64": "base64_encoded_pptx"
}
// Returns: JSON manifest with all files
```

### Rebuild (Zip)
```javascript
POST YOUR_SERVICE_URL/rebuild
{
  "manifest": {
    "zipB64": "original_base64",
    "entries": [/* modified entries */]
  }
}
// Returns: {"zipB64": "modified_pptx_base64"}
```

### Process (Server-side operations)
```javascript
POST YOUR_SERVICE_URL/process
{
  "zipB64": "base64_pptx",
  "ops": [
    {"type": "replaceText", "find": "old", "replace": "new"}
  ]
}
// Returns: Modified PPTX and operation report
```

## Troubleshooting

### Common Issues

#### "Billing not enabled"
- Enable billing in Google Cloud Console
- Even free tier requires billing to be enabled
- You won't be charged within free tier limits

#### "APIs not enabled"
Run in Cloud Shell or terminal:
```bash
gcloud services enable run.googleapis.com
gcloud services enable cloudbuild.googleapis.com
gcloud services enable storage.googleapis.com
```

#### "Deployment failed"
Check:
1. Project ID is correct
2. You have owner/editor permissions
3. APIs are enabled
4. Billing is active

#### "Service not responding"
- Check deployment status: `OOXMLDeployment.getDeploymentStatus()`
- Verify URL is saved in Script Properties
- Check Cloud Run logs in Console

### View Logs

```bash
# Command line
gcloud run services logs read pptx-processor --region=us-central1

# Or in Google Cloud Console
https://console.cloud.google.com/run
```

## Cost & Limits

### Free Tier Limits (us-central1)
- **Memory**: 512MB per instance
- **CPU**: 1 vCPU
- **Requests**: 2 million/month free
- **Compute time**: 180,000 vCPU-seconds/month free
- **Memory time**: 360,000 GB-seconds/month free
- **Network**: 1GB/month free egress

### File Size Limits
- **Max file size**: 10MB (configurable)
- **Timeout**: 60 seconds
- **Max instances**: 3 concurrent

### Estimated Monthly Cost
- **Within free tier**: $0.00
- **Typical usage** (10K requests): $0.00
- **Heavy usage** (100K requests): ~$5-10

## Architecture

```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────────┐
│  Google Apps    │────▶│  Cloud Run       │────▶│  PPTX Files     │
│  Script (GAS)   │◀────│  Service (US)    │◀────│  (Drive)        │
└─────────────────┘     └──────────────────┘     └─────────────────┘
        │                       │                          │
        │                       │                          │
    Base64 PPTX            fflate lib               Google Drive
    JSON Control          Unzip/Zip                 File Storage
```

## Security

### Authentication
- Service allows unauthenticated access by default
- Can be restricted to specific users/service accounts
- OAuth2 available for user-specific access

### Data Flow
1. PPTX files are base64 encoded in GAS
2. Sent to Cloud Run for processing
3. Processed in memory (not stored)
4. Returned as base64 to GAS
5. Saved to Google Drive

### Best Practices
- Don't include sensitive data in PPTX files
- Use service accounts for production
- Enable audit logging if needed
- Set up budget alerts

## Advanced Configuration

### Custom Memory/CPU
```javascript
// In OOXMLDeployment.js
static get _CONFIG() {
  return {
    MEMORY: '1Gi',      // Increase memory
    CPU: 2,             // Increase CPU
    MAX_INSTANCES: 10,  // More concurrent instances
    TIMEOUT: '300s'     // Longer timeout
  };
}
```

### Private Service
```javascript
// Deploy with authentication required
const serviceUrl = await OOXMLDeployment.initAndDeploy({
  region: 'us-central1',
  allowUnauthenticated: false  // Require auth
});
```

### Custom Domain
1. Verify domain in Cloud Console
2. Map service to domain
3. Update SSL certificates

## Monitoring

### Health Dashboard
```javascript
// Create monitoring dashboard
function createMonitoringDashboard() {
  const status = OOXMLDeployment.getDeploymentStatus();
  const health = OOXMLJsonService.healthCheck();
  
  // Log to Stackdriver
  console.log({
    service: 'pptx-processor',
    status: status,
    health: health,
    timestamp: new Date()
  });
}
```

### Alerts
Set up in Google Cloud Console:
1. Cloud Run → Service → Metrics
2. Create alert policy
3. Set thresholds (errors, latency, etc.)

## Cleanup

### Remove Service (Use with Caution!)
```javascript
// This will delete your deployed service
function removeService() {
  const result = OOXMLDeployment.cleanup({
    removeService: true,
    removeBucket: false,
    clearProperties: true
  });
  console.log('Cleanup result:', result);
}
```

## Support

- **Issues**: Create an issue in the repository
- **Documentation**: See README.md and examples/
- **Cloud Console**: https://console.cloud.google.com
- **GAS Support**: https://developers.google.com/apps-script

---

**Remember**: The entire deployment can be done from Google Apps Script - no command line required! 🚀