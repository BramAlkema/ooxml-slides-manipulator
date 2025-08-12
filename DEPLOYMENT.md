# Google Cloud Platform Deployment Guide

## Prerequisites

1. **Google Cloud Account**: Ensure you have access to Google Cloud Platform
2. **gcloud CLI**: Already installed in this project
3. **Project**: You'll need a GCP project with billing enabled

## Authentication & Setup

1. **Authenticate with Google Cloud:**
   ```bash
   export PATH="$HOME/google-cloud-sdk/bin:$PATH"
   gcloud auth login
   ```

2. **Set your project:**
   ```bash
   gcloud config set project YOUR_PROJECT_ID
   ```

3. **Enable required APIs:**
   ```bash
   gcloud services enable cloudfunctions.googleapis.com
   gcloud services enable cloudbuild.googleapis.com
   ```

## Deploy the Cloud Function

1. **Navigate to the cloud function directory:**
   ```bash
   cd cloud-function
   ```

2. **Deploy the function:**
   ```bash
   npm run deploy
   ```
   
   Or manually:
   ```bash
   gcloud functions deploy pptx-router \
     --runtime nodejs18 \
     --trigger http \
     --entry-point pptxRouter \
     --allow-unauthenticated \
     --memory 512MB \
     --timeout 120s
   ```

3. **Get the deployed URL:**
   ```bash
   gcloud functions describe pptx-router --format="value(httpsTrigger.url)"
   ```

## Update Google Apps Script

1. **Update CloudPPTXService.js:**
   Replace the localhost URL with your deployed function URL:
   ```javascript
   static get CLOUD_FUNCTION_URL() {
     return 'https://YOUR_REGION-YOUR_PROJECT.cloudfunctions.net/pptx-router';
   }
   ```

2. **Push to Google Apps Script:**
   ```bash
   clasp push
   ```

## Test the Deployment

Run the test suite in Google Apps Script:
```javascript
runAllCloudFunctionTests();
```

## Expected Deployed Function URL Format

Your function will be available at:
```
https://[REGION]-[PROJECT_ID].cloudfunctions.net/pptx-router
```

Where:
- `REGION` is typically `us-central1` (default)
- `PROJECT_ID` is your Google Cloud project ID

## Troubleshooting

### Function Timeout
If processing large PPTX files, increase timeout:
```bash
gcloud functions deploy pptx-router --timeout 300s
```

### Memory Issues
Increase memory allocation:
```bash
gcloud functions deploy pptx-router --memory 1024MB
```

### CORS Issues
The function includes CORS headers, but if you encounter issues, verify the headers in `index.js`.

## Cost Optimization

- **Memory**: Start with 512MB (sufficient for most PPTX files)
- **Timeout**: 120s should handle most files, adjust as needed
- **Invocations**: Function is only called during PPTX manipulation

## Security Note

The current deployment uses `--allow-unauthenticated` for simplicity. For production:

1. Remove `--allow-unauthenticated`
2. Use service account authentication
3. Implement API key authentication if needed