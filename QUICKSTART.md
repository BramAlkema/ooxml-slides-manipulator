# 🚀 5-Minute Quick Start Guide

Get your PPTX processing service running in 5 minutes! No command line required - everything runs from Google Apps Script.

## What You'll Build

A cloud service that can:
- 📦 **Unzip** PowerPoint files to access internal structure
- ✏️ **Modify** themes, colors, text, and layouts
- 🔨 **Zip** everything back into valid PPTX files

## Step 1: Create Google Apps Script Project (1 min)

1. Go to [script.google.com](https://script.google.com)
2. Click **New Project**
3. Name it: "PPTX Processor"

## Step 2: Add the Code (2 min)

1. **Delete** the default `Code.gs` content
2. **Copy & paste** this deployment code:

```javascript
// ===== QUICK DEPLOYMENT CODE =====

function quickSetup() {
  // Set your Google Cloud Project ID here
  const PROJECT_ID = 'your-project-id'; // ← CHANGE THIS!
  
  PropertiesService.getScriptProperties()
    .setProperty('GCP_PROJECT_ID', PROJECT_ID);
  
  console.log('✅ Project configured:', PROJECT_ID);
  console.log('Next: Run quickDeploy()');
}

async function quickDeploy() {
  try {
    console.log('🚀 Starting deployment...');
    
    // This deploys to US free tier automatically
    const url = await OOXMLDeployment.initAndDeploy({
      region: 'us-central1',
      skipBillingCheck: true // Remove if unsure about billing
    });
    
    console.log('✅ SUCCESS! Service deployed to:', url);
    
    // Save for later use
    PropertiesService.getScriptProperties()
      .setProperty('CLOUD_SERVICE_URL', url);
    
    return url;
  } catch (error) {
    console.error('❌ Deployment failed:', error.message);
    console.log('See DEPLOYMENT.md for troubleshooting');
  }
}

async function testPPTX() {
  console.log('🧪 Testing PPTX processing...');
  
  // Create a test presentation
  const pres = SlidesApp.create('Test PPTX');
  const slide = pres.getSlides()[0];
  slide.getShapes()[0].getText().setText('Original Text');
  
  const fileId = pres.getId();
  
  // UNZIP: Extract PPTX structure
  const manifest = await OOXMLJsonService.unwrap(fileId);
  console.log(`📦 Extracted ${manifest.entries.length} files`);
  
  // MODIFY: Change content
  manifest.entries.forEach(entry => {
    if (entry.type === 'xml' && entry.text) {
      entry.text = entry.text.replace('Original Text', 'Modified Text');
    }
  });
  
  // ZIP: Create new PPTX
  const outputId = await OOXMLJsonService.rewrap(manifest, {
    filename: 'Modified Test.pptx'
  });
  
  console.log('✅ Success! Modified file:', outputId);
  console.log('Check your Google Drive for the file');
}
```

3. **Add required libraries** (copy from this repo):
   - Copy contents of `src/OOXMLDeployment.js`
   - Copy contents of `lib/OOXMLJsonService.js`
   - Add as new files in your project

## Step 3: Configure Your Project (1 min)

1. Get your Google Cloud Project ID:
   - Go to [console.cloud.google.com](https://console.cloud.google.com)
   - Select or create a project
   - Copy the Project ID (looks like: `my-project-123456`)

2. In Apps Script, run:
   ```javascript
   quickSetup()  // Enter your project ID in the code first!
   ```

3. Check the console output for confirmation

## Step 4: Deploy to Free Tier (1 min)

1. Run in Apps Script:
   ```javascript
   quickDeploy()
   ```

2. Wait for "SUCCESS!" message (takes 2-3 minutes)
3. Copy the service URL from the output

## Step 5: Test It Works! (1 min)

Run the test function:
```javascript
testPPTX()
```

Check your Google Drive - you should see "Modified Test.pptx"!

## 🎉 That's It! You're Done!

Your PPTX processing service is now running on Google Cloud free tier.

## What's Next?

### Try the Full Workflow

```javascript
async function processMyPresentation() {
  // 1. UNZIP your PPTX
  const manifest = await OOXMLJsonService.unwrap('my-presentation.pptx');
  
  // 2. MODIFY what you need
  manifest.entries.forEach(entry => {
    if (entry.type === 'xml') {
      // Replace company name
      entry.text = entry.text.replace('OldCompany', 'NewCompany');
      
      // Change colors
      entry.text = entry.text.replace('#FF0000', '#0066CC');
      
      // Update dates
      entry.text = entry.text.replace('2023', '2024');
    }
  });
  
  // 3. ZIP back to PPTX
  const outputId = await OOXMLJsonService.rewrap(manifest, {
    filename: 'updated-presentation.pptx'
  });
  
  console.log('Done! File saved:', outputId);
}
```

### Common Operations

**Change all fonts:**
```javascript
entry.text = entry.text.replace(/font="Arial"/g, 'font="Helvetica"');
```

**Update theme colors:**
```javascript
if (entry.path.includes('theme')) {
  entry.text = entry.text.replace('accent1="FF0000"', 'accent1="0066CC"');
}
```

**Replace slide content:**
```javascript
if (entry.path.includes('slide1.xml')) {
  entry.text = entry.text.replace('<a:t>Old Title</a:t>', '<a:t>New Title</a:t>');
}
```

## Troubleshooting

### "Project ID not found"
- Make sure you changed `your-project-id` to your actual project ID
- Run `quickSetup()` again with the correct ID

### "Billing not enabled"
- Go to [console.cloud.google.com/billing](https://console.cloud.google.com/billing)
- Enable billing (you won't be charged in free tier)
- Try deployment again

### "APIs not enabled"
The deployment should enable APIs automatically, but if not:
1. Go to [console.cloud.google.com/apis](https://console.cloud.google.com/apis)
2. Enable: Cloud Run API, Cloud Build API

### "Deployment failed"
Check that you have:
- ✅ Valid Google Cloud project
- ✅ Billing enabled (even for free tier)
- ✅ Editor/Owner permissions on the project

## Free Tier Limits

Your service runs FREE with these limits:
- 2 million requests/month
- 360,000 GB-seconds of memory/month
- 180,000 vCPU-seconds/month
- 1 GB network egress/month

**Typical usage cost: $0.00** 🎉

## Advanced Features

Once you're comfortable with the basics:

### Server-Side Operations (Faster!)
```javascript
// Process without downloading the whole file
const result = await OOXMLJsonService.process(
  'presentation.pptx',
  [
    { type: 'replaceText', find: 'Q3 2023', replace: 'Q4 2024' },
    { type: 'replaceText', find: 'Draft', replace: 'Final' }
  ]
);
```

### Brand Compliance
```javascript
const slides = new OOXMLSlides('template.pptx');
await slides.applyBrandColors({
  primary: '#0066CC',
  secondary: '#FF6600'
});
```

### Large Files (100MB+)
```javascript
const session = await OOXMLJsonService.createSession();
// Process large files via Google Cloud Storage
```

## Get Help

- 📖 [Full Documentation](README.md)
- 🔧 [Deployment Guide](DEPLOYMENT.md)
- 💡 [Examples](examples/)
- 🐛 [Troubleshooting](DEPLOYMENT.md#troubleshooting)

## One-Line Deploy (For Experts)

If you already have everything set up:
```javascript
OOXMLDeployment.initAndDeploy({region:'us-central1'}).then(url => console.log('Deployed:', url))
```

---

**Remember:** Everything runs from Google Apps Script - no terminal needed! 🚀

**Total time:** Under 5 minutes from zero to working PPTX processor! ⏱️