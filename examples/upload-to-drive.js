/**
 * Upload the brandware PPTX to Google Drive and make it publicly accessible
 */

const fs = require('fs');
const { google } = require('googleapis');

async function uploadToDrive() {
  console.log('📤 Uploading brandware PPTX to Google Drive...\n');
  
  try {
    // Read the service account key
    const keyFile = JSON.parse(fs.readFileSync('./oauth-key.json', 'utf8'));
    
    // Create JWT auth
    const auth = new google.auth.JWT(
      keyFile.client_email,
      null,
      keyFile.private_key,
      ['https://www.googleapis.com/auth/drive']
    );
    
    // Initialize Drive API
    const drive = google.drive({ version: 'v3', auth });
    
    // Read the PPTX file
    const fileBuffer = fs.readFileSync('./brandware-typography-tables-demo.pptx');
    
    // Upload to Google Drive
    const fileMetadata = {
      name: 'Brandware Typography & Tables Demo.pptx',
      description: 'Advanced PowerPoint with custom typography, kerning, and table styles created via OOXML manipulation'
    };
    
    const media = {
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      body: require('stream').Readable.from(fileBuffer)
    };
    
    console.log('  Uploading file to Google Drive...');
    const uploadResponse = await drive.files.create({
      resource: fileMetadata,
      media: media,
      fields: 'id,name,webViewLink,webContentLink'
    });
    
    const fileId = uploadResponse.data.id;
    console.log('  ✅ File uploaded:', uploadResponse.data.name);
    console.log('  📁 File ID:', fileId);
    
    // Make the file publicly viewable
    console.log('  Making file publicly accessible...');
    await drive.permissions.create({
      fileId: fileId,
      resource: {
        role: 'reader',
        type: 'anyone'
      }
    });
    
    console.log('  ✅ File made publicly accessible');
    
    // Get the updated file info
    const fileInfo = await drive.files.get({
      fileId: fileId,
      fields: 'id,name,webViewLink,webContentLink,size,createdTime'
    });
    
    console.log('\n🎉 Upload completed successfully!');
    console.log('📊 File Details:');
    console.log('  Name:', fileInfo.data.name);
    console.log('  ID:', fileInfo.data.id);
    console.log('  Size:', Math.round(fileInfo.data.size / 1024), 'KB');
    console.log('  Created:', fileInfo.data.createdTime);
    console.log('  View Link:', fileInfo.data.webViewLink);
    console.log('  Download Link:', fileInfo.data.webContentLink);
    
    return {
      fileId: fileInfo.data.id,
      viewLink: fileInfo.data.webViewLink,
      downloadLink: fileInfo.data.webContentLink,
      name: fileInfo.data.name
    };
    
  } catch (error) {
    console.error('❌ Upload failed:', error.message);
    throw error;
  }
}

// Execute if called directly
if (require.main === module) {
  uploadToDrive().catch(console.error);
}

module.exports = { uploadToDrive };