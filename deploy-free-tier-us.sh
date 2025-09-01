#!/bin/bash

# Deploy PPTX Processing Service to US Free Tier
# Optimized for Google Cloud Free Tier limits

set -e  # Exit on any error

echo "🇺🇸 Deploying PPTX Service to US Free Tier"
echo "==========================================="

# Check if gcloud is authenticated
if ! gcloud auth list --filter=status:ACTIVE --format="value(account)" | grep -q "@"; then
    echo "❌ Not authenticated with Google Cloud"
    echo "Please run: gcloud auth login"
    exit 1
fi

# Get current project
PROJECT=$(gcloud config get-value project 2>/dev/null)
if [ -z "$PROJECT" ]; then
    echo "❌ No project set"
    echo "Please run: gcloud config set project YOUR_PROJECT_ID"
    exit 1
fi

echo "📋 Project: $PROJECT"
echo "📋 Region: us-central1 (Free Tier)"
echo "📋 Service: ooxml-pptx-processor"

# Set default region for free tier
gcloud config set run/region us-central1

# Enable required APIs (free on all projects)
echo "🔧 Enabling required APIs..."
gcloud services enable cloudbuild.googleapis.com --quiet
gcloud services enable run.googleapis.com --quiet
gcloud services enable storage.googleapis.com --quiet

# Create optimized package.json for free tier
echo "📦 Creating optimized free-tier package.json..."
cat > /tmp/free-tier-package.json << 'EOF'
{
  "name": "pptx-processor-free-tier",
  "version": "1.0.0",
  "description": "Free-tier PPTX processing service",
  "main": "index.js",
  "scripts": {
    "start": "node index.js"
  },
  "dependencies": {
    "express": "^4.19.2",
    "fflate": "^0.8.2"
  },
  "engines": {
    "node": ">=18"
  }
}
EOF

# Create optimized index.js for free tier
echo "📝 Creating optimized free-tier service..."
cat > /tmp/free-tier-index.js << 'EOF'
const express = require('express');
const { unzipSync, zipSync } = require('fflate');

const app = express();

// Free tier optimizations
app.use(express.json({ limit: '10mb' })); // Reduced for free tier
app.use(express.raw({ type: 'application/octet-stream', limit: '10mb' }));

// Health check endpoint (required for Cloud Run)
app.get('/health', (req, res) => {
  res.json({ 
    status: 'healthy',
    service: 'pptx-processor-free-tier',
    region: 'us-central1',
    timestamp: new Date().toISOString()
  });
});

// Root endpoint
app.get('/', (req, res) => {
  res.json({
    service: 'PPTX Processor Free Tier',
    version: '1.0.0',
    endpoints: ['/health', '/extract', '/rebuild', '/ping'],
    region: 'us-central1',
    limits: {
      maxFileSize: '10MB',
      timeout: '60s',
      memory: '512MB'
    }
  });
});

// Ping endpoint
app.get('/ping', (req, res) => {
  res.json({ 
    message: 'pong',
    timestamp: new Date().toISOString(),
    freeTier: true
  });
});

// Extract PPTX contents
app.post('/extract', async (req, res) => {
  try {
    const { zipB64 } = req.body;
    
    if (!zipB64) {
      return res.status(400).json({ error: 'zipB64 required' });
    }
    
    // Convert base64 to Uint8Array
    const buffer = Buffer.from(zipB64, 'base64');
    const uint8Array = new Uint8Array(buffer);
    
    // Unzip using fflate
    const extracted = unzipSync(uint8Array);
    
    // Process files
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
    
    res.json({
      success: true,
      entries,
      totalFiles: entries.length,
      extractedAt: new Date().toISOString()
    });
    
  } catch (error) {
    console.error('Extract error:', error);
    res.status(500).json({ 
      error: 'Extraction failed',
      message: error.message 
    });
  }
});

// Rebuild PPTX from modified contents
app.post('/rebuild', async (req, res) => {
  try {
    const { manifest } = req.body;
    
    if (!manifest || !manifest.zipB64) {
      return res.status(400).json({ error: 'manifest.zipB64 required' });
    }
    
    // Start with original ZIP data
    const originalBuffer = Buffer.from(manifest.zipB64, 'base64');
    const originalUint8 = new Uint8Array(originalBuffer);
    const extracted = unzipSync(originalUint8);
    
    // Apply modifications from manifest.entries
    const encoder = new TextEncoder();
    
    if (manifest.entries) {
      for (const entry of manifest.entries) {
        if (entry.type === 'xml' && entry.text) {
          extracted[entry.path] = encoder.encode(entry.text);
        }
      }
    }
    
    // Rebuild ZIP
    const rebuilt = zipSync(extracted, { level: 6 });
    const rebuiltBase64 = Buffer.from(rebuilt).toString('base64');
    
    res.json({
      success: true,
      zipB64: rebuiltBase64,
      originalSize: originalBuffer.length,
      rebuiltSize: rebuilt.length,
      rebuiltAt: new Date().toISOString()
    });
    
  } catch (error) {
    console.error('Rebuild error:', error);
    res.status(500).json({ 
      error: 'Rebuild failed',
      message: error.message 
    });
  }
});

// Error handling middleware
app.use((error, req, res, next) => {
  console.error('Unhandled error:', error);
  res.status(500).json({ 
    error: 'Internal server error',
    freeTier: true
  });
});

// Start server
const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`🚀 PPTX Processor Free Tier running on port ${PORT}`);
  console.log(`🇺🇸 Region: us-central1`);
  console.log(`📦 Memory: 512MB`);
  console.log(`⏱️  Timeout: 60s`);
});
EOF

# Create Dockerfile optimized for free tier
echo "🐳 Creating free-tier optimized Dockerfile..."
cat > /tmp/free-tier-Dockerfile << 'EOF'
# Use official Node.js runtime (free tier compatible)
FROM node:18-slim

# Install curl for healthcheck
RUN apt-get update && apt-get install -y curl && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy package files
COPY package.json ./

# Install dependencies (only production)
RUN npm install --only=production

# Copy application code
COPY index.js ./

# Expose port
EXPOSE 8080

# Start application (no healthcheck to avoid issues)
CMD ["npm", "start"]
EOF

# Create temporary deployment directory
DEPLOY_DIR="/tmp/pptx-free-tier-deploy"
rm -rf $DEPLOY_DIR
mkdir -p $DEPLOY_DIR

# Copy files to deployment directory
cp /tmp/free-tier-package.json $DEPLOY_DIR/package.json
cp /tmp/free-tier-index.js $DEPLOY_DIR/index.js
cp /tmp/free-tier-Dockerfile $DEPLOY_DIR/Dockerfile

echo "📂 Deployment files ready in: $DEPLOY_DIR"

# Deploy to Cloud Run (free tier optimized)
echo "🚀 Deploying to Cloud Run (Free Tier)..."
cd $DEPLOY_DIR

gcloud run deploy ooxml-pptx-processor \
  --source . \
  --region us-central1 \
  --platform managed \
  --allow-unauthenticated \
  --memory 512Mi \
  --cpu 1 \
  --timeout 60s \
  --concurrency 10 \
  --min-instances 0 \
  --max-instances 3 \
  --port 8080 \
  --set-env-vars "NODE_ENV=production,REGION=us-central1" \
  --quiet

# Get the deployed URL
echo "🔗 Getting service URL..."
SERVICE_URL=$(gcloud run services describe ooxml-pptx-processor --region=us-central1 --format="value(status.url)")

echo ""
echo "✅ Free Tier Deployment Complete!"
echo "=================================="
echo "🇺🇸 Region: us-central1 (Free Tier)"
echo "🔗 Service URL: $SERVICE_URL"
echo "💰 Cost: FREE (within limits)"
echo "📊 Limits:"
echo "   - Memory: 512MB"
echo "   - CPU: 1 vCPU"
echo "   - Timeout: 60s"
echo "   - Max instances: 3"
echo "   - File size: 10MB"
echo ""
echo "🧪 Test your deployment:"
echo "curl $SERVICE_URL/ping"
echo "curl $SERVICE_URL/health"
echo ""
echo "📋 Update your GAS script with:"
echo "const CLOUD_FUNCTION_URL = '$SERVICE_URL';"

# Cleanup
rm -rf $DEPLOY_DIR
rm -f /tmp/free-tier-*

echo ""
echo "🎉 Your PPTX processing service is now running on US free tier!"