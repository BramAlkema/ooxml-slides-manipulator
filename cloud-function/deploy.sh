#!/bin/bash

# Deploy JSZip Cloud Function to Google Cloud Platform
# Run this script after authenticating with: gcloud auth login

set -e  # Exit on any error

echo "🚀 Deploying JSZip Cloud Function to Google Cloud Platform"
echo "============================================================"

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

echo "📋 Current project: $PROJECT"
echo "📋 Current directory: $(pwd)"

# Enable required APIs
echo "🔧 Enabling required APIs..."
gcloud services enable cloudfunctions.googleapis.com
gcloud services enable cloudbuild.googleapis.com

# Deploy the function
echo "📦 Deploying Cloud Function..."
gcloud functions deploy pptx-router \
    --runtime nodejs18 \
    --trigger http \
    --entry-point pptxRouter \
    --allow-unauthenticated \
    --memory 512MB \
    --timeout 120s \
    --region us-central1

# Get the deployed URL
echo "🔗 Getting deployed function URL..."
FUNCTION_URL=$(gcloud functions describe pptx-router --region=us-central1 --format="value(httpsTrigger.url)")

echo ""
echo "✅ Deployment Complete!"
echo "========================"
echo "Function URL: $FUNCTION_URL"
echo ""
echo "📝 Next steps:"
echo "1. Update CloudPPTXService.js with this URL:"
echo "   static get CLOUD_FUNCTION_URL() {"
echo "     return '$FUNCTION_URL';"
echo "   }"
echo ""
echo "2. Push to Google Apps Script:"
echo "   clasp push"
echo ""
echo "3. Test the integration:"
echo "   runAllCloudFunctionTests();"