const functions = require('@google-cloud/functions-framework');
const JSZip = require('jszip');

/**
 * Main router function for PPTX operations
 */
functions.http('pptxRouter', async (req, res) => {
  // Set CORS headers for Apps Script access
  res.set('Access-Control-Allow-Origin', '*');
  res.set('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.set('Access-Control-Allow-Headers', 'Content-Type');

  // Handle preflight requests
  if (req.method === 'OPTIONS') {
    res.status(204).send('');
    return;
  }

  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Only POST method allowed' });
    return;
  }

  // Route based on URL path
  const path = req.path || req.url;
  
  if (path.endsWith('/unzip') || path.includes('unzip')) {
    await handlePptxUnzip(req, res);
  } else if (path.endsWith('/zip') || path.includes('zip')) {
    await handlePptxZip(req, res);
  } else {
    // Default to unzip for backward compatibility
    await handlePptxUnzip(req, res);
  }
});

/**
 * Handle PPTX unzip operations
 */
async function handlePptxUnzip(req, res) {

  try {
    console.log('Processing PPTX unzip request...');
    
    // Get the PPTX data from request body
    let pptxData = req.body;
    
    // Handle different content types
    if (req.get('content-type') === 'application/json') {
      pptxData = Buffer.from(pptxData.data || pptxData, 'base64');
    } else if (typeof pptxData === 'string') {
      pptxData = Buffer.from(pptxData, 'base64');
    } else if (!Buffer.isBuffer(pptxData)) {
      pptxData = Buffer.from(pptxData);
    }
    
    if (!pptxData || pptxData.length === 0) {
      res.status(400).json({ error: 'No PPTX data provided' });
      return;
    }

    console.log(`Received PPTX data: ${pptxData.length} bytes`);

    // Load the PPTX file with JSZip
    const zip = new JSZip();
    const zipContent = await zip.loadAsync(pptxData);
    
    console.log(`Successfully loaded ZIP with ${Object.keys(zipContent.files).length} files`);

    // Extract all files as text content
    const files = {};
    const promises = [];

    Object.keys(zipContent.files).forEach(filename => {
      if (!zipContent.files[filename].dir) {
        promises.push(
          zipContent.files[filename].async('text').then(content => {
            files[filename] = content;
            console.log(`Extracted: ${filename} (${content.length} chars)`);
          }).catch(err => {
            // Try as binary if text fails
            return zipContent.files[filename].async('base64').then(content => {
              files[filename] = {
                type: 'binary',
                content: content
              };
              console.log(`Extracted as binary: ${filename} (${content.length} chars base64)`);
            });
          })
        );
      }
    });

    await Promise.all(promises);

    console.log(`Successfully extracted ${Object.keys(files).length} files`);

    // Return the extracted files
    res.status(200).json({
      success: true,
      files: files,
      fileCount: Object.keys(files).length,
      originalSize: pptxData.length
    });

  } catch (error) {
    console.error('Error processing PPTX:', error);
    res.status(500).json({
      error: 'Failed to process PPTX file',
      message: error.message,
      stack: error.stack
    });
  }
}

/**
 * Handle PPTX zip operations
 */
async function handlePptxZip(req, res) {

  try {
    console.log('Processing PPTX zip request...');
    
    const { files } = req.body;
    if (!files || typeof files !== 'object') {
      res.status(400).json({ error: 'No files object provided' });
      return;
    }

    console.log(`Creating PPTX from ${Object.keys(files).length} files`);

    // Create new ZIP file
    const zip = new JSZip();

    // Add each file to the ZIP
    Object.entries(files).forEach(([filename, content]) => {
      if (typeof content === 'object' && content.type === 'binary') {
        // Handle binary files
        zip.file(filename, content.content, { base64: true });
        console.log(`Added binary file: ${filename}`);
      } else {
        // Handle text files
        zip.file(filename, content);
        console.log(`Added text file: ${filename} (${content.length} chars)`);
      }
    });

    // Generate the PPTX file
    const pptxData = await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });

    console.log(`Generated PPTX: ${pptxData.length} bytes`);

    // Return the PPTX file as base64
    const base64Data = pptxData.toString('base64');
    
    res.status(200).json({
      success: true,
      pptxData: base64Data,
      size: pptxData.length,
      fileCount: Object.keys(files).length
    });

  } catch (error) {
    console.error('Error creating PPTX:', error);
    res.status(500).json({
      error: 'Failed to create PPTX file',
      message: error.message,
      stack: error.stack
    });
  }
}