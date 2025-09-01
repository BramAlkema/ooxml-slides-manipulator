#!/bin/bash

# Simple Free Tier US Deployment - No Dockerfile
# Uses Cloud Run source deployment without custom Dockerfile

set -e  # Exit on any error

echo "🇺🇸 Simple Free Tier Deployment to US"
echo "====================================="

# Check authentication and project
PROJECT=$(gcloud config get-value project 2>/dev/null)
if [ -z "$PROJECT" ]; then
    echo "❌ No project set"
    echo "Please run: gcloud config set project YOUR_PROJECT_ID"
    exit 1
fi

echo "📋 Project: $PROJECT"
echo "📋 Region: us-central1 (Free Tier)"

# Set region
gcloud config set run/region us-central1

# Enable APIs
echo "🔧 Enabling required APIs..."
gcloud services enable run.googleapis.com --quiet
gcloud services enable cloudbuild.googleapis.com --quiet

# Create simple deployment directory
DEPLOY_DIR="/tmp/simple-pptx-deploy"
rm -rf $DEPLOY_DIR
mkdir -p $DEPLOY_DIR
cd $DEPLOY_DIR

# Create simple package.json
echo "📦 Creating simple package.json..."
cat > package.json << 'EOF'
{
  "name": "pptx-processor",
  "version": "1.0.0",
  "main": "index.js",
  "scripts": {
    "start": "node index.js"
  },
  "dependencies": {
    "express": "^4.19.2",
    "fflate": "^0.8.2"
  },
  "engines": {
    "node": "18"
  }
}
EOF

# Create simple index.js
echo "📝 Creating simple service..."
cat > index.js << 'EOF'
const express = require('express');
const { unzipSync, zipSync } = require('fflate');

const app = express();
app.use(express.json({ limit: '10mb' }));

// Root endpoint
app.get('/', (req, res) => {
  res.json({
    service: 'PPTX Processor US Free Tier',
    version: '1.0.0',
    region: 'us-central1',
    endpoints: ['/ping', '/extract', '/rebuild']
  });
});

// Ping endpoint
app.get('/ping', (req, res) => {
  res.json({ 
    message: 'pong', 
    timestamp: new Date().toISOString(),
    region: 'us-central1'
  });
});

// Extract PPTX
app.post('/extract', (req, res) => {
  try {
    const { zipB64 } = req.body;
    if (!zipB64) {
      return res.status(400).json({ error: 'zipB64 required' });
    }
    
    const buffer = Buffer.from(zipB64, 'base64');
    const extracted = unzipSync(new Uint8Array(buffer));
    
    const entries = [];
    const decoder = new TextDecoder();
    
    for (const [path, data] of Object.entries(extracted)) {
      const isXml = path.endsWith('.xml') || path.endsWith('.rels');
      entries.push({
        path,
        type: isXml ? 'xml' : 'bin',
        text: isXml ? decoder.decode(data) : undefined,
        size: data.length
      });
    }
    
    res.json({ success: true, entries, totalFiles: entries.length });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Rebuild PPTX
app.post('/rebuild', (req, res) => {
  try {
    const { manifest } = req.body;
    if (!manifest?.zipB64) {
      return res.status(400).json({ error: 'manifest.zipB64 required' });
    }
    
    const originalBuffer = Buffer.from(manifest.zipB64, 'base64');
    const extracted = unzipSync(new Uint8Array(originalBuffer));
    
    // Apply modifications
    const encoder = new TextEncoder();
    if (manifest.entries) {
      for (const entry of manifest.entries) {
        if (entry.type === 'xml' && entry.text) {
          extracted[entry.path] = encoder.encode(entry.text);
        }
      }
    }
    
    const rebuilt = zipSync(extracted, { level: 6 });
    const rebuiltBase64 = Buffer.from(rebuilt).toString('base64');
    
    res.json({ success: true, zipB64: rebuiltBase64 });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`🚀 PPTX Processor running on port ${PORT} (us-central1)`);
});
EOF

# Deploy using source (no Dockerfile)
echo "🚀 Deploying to Cloud Run..."
gcloud run deploy pptx-processor \
  --source . \
  --region us-central1 \
  --allow-unauthenticated \
  --memory 512Mi \
  --cpu 1 \
  --timeout 60s \
  --max-instances 3 \
  --set-env-vars "NODE_ENV=production" \
  --quiet

# Get service URL
SERVICE_URL=$(gcloud run services describe pptx-processor --region=us-central1 --format="value(status.url)")

echo ""
echo "✅ Deployment Complete!"
echo "======================="
echo "🔗 Service URL: $SERVICE_URL"
echo ""
echo "🧪 Test commands:"
echo "curl $SERVICE_URL/ping"
echo "curl $SERVICE_URL/"
echo ""
echo "📋 Update your GAS with:"
echo "const CLOUD_SERVICE_URL = '$SERVICE_URL';"

# Cleanup
cd /
rm -rf $DEPLOY_DIR

echo ""
echo "🎉 Your PPTX service is live on US free tier!"