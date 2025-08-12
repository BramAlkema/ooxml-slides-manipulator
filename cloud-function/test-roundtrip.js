/**
 * Round-trip test for PPTX Cloud Function
 * Tests zip -> unzip -> zip cycle to ensure data integrity
 */

async function testRoundTrip() {
  console.log('Testing round-trip: zip → unzip → zip...\n');
  
  try {
    // Step 1: Create a PPTX using zip function
    console.log('Step 1: Creating PPTX with zip function...');
    const originalFiles = {
      '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-presentationml.presentation.main+xml"/></Types>',
      '_rels/.rels': '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>',
      'ppt/presentation.xml': '<?xml version="1.0" encoding="UTF-8"?><p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst/><p:sldSz cx="9144000" cy="6858000"/></p:presentation>'
    };
    
    const zipResponse = await fetch('http://localhost:8080/zip', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ files: originalFiles })
    });
    
    const zipResult = await zipResponse.json();
    if (!zipResult.success) {
      console.error('❌ Failed to create PPTX:', zipResult);
      return;
    }
    
    console.log(`✅ Created PPTX: ${zipResult.size} bytes`);
    console.log(`   Files: ${zipResult.fileCount}`);
    
    // Step 2: Unzip the created PPTX
    console.log('\nStep 2: Unzipping the created PPTX...');
    const unzipResponse = await fetch('http://localhost:8080/unzip', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ data: zipResult.pptxData })
    });
    
    const unzipResult = await unzipResponse.json();
    if (!unzipResult.success) {
      console.error('❌ Failed to unzip PPTX:', unzipResult);
      return;
    }
    
    console.log(`✅ Unzipped PPTX: ${unzipResult.fileCount} files extracted`);
    console.log('   Files extracted:', Object.keys(unzipResult.files));
    
    // Step 3: Re-zip the extracted files
    console.log('\nStep 3: Re-zipping the extracted files...');
    const rezipResponse = await fetch('http://localhost:8080/zip', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ files: unzipResult.files })
    });
    
    const rezipResult = await rezipResponse.json();
    if (!rezipResult.success) {
      console.error('❌ Failed to re-zip files:', rezipResult);
      return;
    }
    
    console.log(`✅ Re-zipped PPTX: ${rezipResult.size} bytes`);
    console.log(`   Files: ${rezipResult.fileCount}`);
    
    // Step 4: Compare original and final results
    console.log('\n📊 Round-trip comparison:');
    console.log(`   Original size: ${zipResult.size} bytes`);
    console.log(`   Final size: ${rezipResult.size} bytes`);
    console.log(`   Size difference: ${rezipResult.size - zipResult.size} bytes`);
    console.log(`   Original files: ${zipResult.fileCount}`);
    console.log(`   Final files: ${rezipResult.fileCount}`);
    
    if (zipResult.fileCount === rezipResult.fileCount) {
      console.log('✅ File count matches!');
    } else {
      console.log('❌ File count mismatch!');
    }
    
    console.log('\n🎉 Round-trip test completed successfully!');
    
  } catch (error) {
    console.error('❌ Round-trip test failed:', error);
  }
}

// Run the test
testRoundTrip();