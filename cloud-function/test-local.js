/**
 * Test script for local Google Cloud Function
 */
const fs = require('fs');
const path = require('path');

// Test the pptxUnzip function
async function testPptxUnzip() {
  console.log('Testing pptxUnzip function...');
  
  try {
    // Use tanaikech's proven base64 template (this is a real PPTX file)
    const tanaikechTemplate = "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";
    
    const response = await fetch('http://localhost:8080/unzip', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ data: tanaikechTemplate })
    });
    
    const result = await response.json();
    console.log('pptxUnzip response success:', result.success);
    console.log('pptxUnzip response file count:', result.fileCount);
    
    if (result.files) {
      console.log('Files extracted:', Object.keys(result.files));
    }
    
    return result;
    
  } catch (error) {
    console.error('pptxUnzip test failed:', error);
    return null;
  }
}

// Test the pptxZip function  
async function testPptxZip() {
  console.log('Testing pptxZip function...');
  
  try {
    // Create test files object
    const testFiles = {
      '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8"?><Types></Types>',
      '_rels/.rels': '<?xml version="1.0" encoding="UTF-8"?><Relationships></Relationships>',
      'ppt/presentation.xml': '<?xml version="1.0" encoding="UTF-8"?><presentation></presentation>'
    };
    
    const response = await fetch('http://localhost:8080/zip', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ files: testFiles })
    });
    
    const result = await response.json();
    console.log('pptxZip response success:', result.success);
    console.log('pptxZip response size:', result.size);
    
    return result;
    
  } catch (error) {
    console.error('pptxZip test failed:', error);
    return null;
  }
}

// Run tests
async function runTests() {
  console.log('Starting Cloud Function tests...\n');
  
  // Test unzip first
  const unzipResult = await testPptxUnzip();
  
  console.log('\n' + '='.repeat(50) + '\n');
  
  // Test zip
  const zipResult = await testPptxZip();
  
  console.log('\n' + '='.repeat(50));
  console.log('Tests completed!');
}

// Only run if this file is executed directly
if (require.main === module) {
  runTests();
}

module.exports = { testPptxUnzip, testPptxZip };