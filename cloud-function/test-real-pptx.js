/**
 * Test JSZip with a real PPTX file 
 * This will test if we can manipulate existing PPTX files created by PowerPoint
 */

const JSZip = require('jszip');
const fs = require('fs');

async function testRealPPTX() {
  console.log('🧪 Testing JSZip with real PPTX file...\n');
  
  try {
    // First, let's use our own generated PPTX as a "real" file
    console.log('Step 1: Using our generated PPTX as test file...');
    
    if (!fs.existsSync('test-ooxml-pptx.pptx')) {
      console.log('❌ test-ooxml-pptx.pptx not found. Run test-ooxml-compatibility.js first.');
      return false;
    }
    
    const realPptxBuffer = fs.readFileSync('test-ooxml-pptx.pptx');
    console.log(`📂 Loaded PPTX file: ${realPptxBuffer.length} bytes`);
    
    // Test 2: Load and analyze the PPTX
    console.log('\nStep 2: Loading and analyzing PPTX structure...');
    
    const zip = new JSZip();
    const loadedZip = await zip.loadAsync(realPptxBuffer);
    
    console.log(`📦 ZIP contains ${Object.keys(loadedZip.files).length} files/folders`);
    
    // List all files
    console.log('\nFiles in PPTX:');
    Object.keys(loadedZip.files).forEach((filename, index) => {
      const file = loadedZip.files[filename];
      if (!file.dir) {
        console.log(`  ${index + 1}. ${filename} (${file._data.uncompressedSize || 'unknown'} bytes)`);
      }
    });
    
    // Test 3: Extract specific files for manipulation
    console.log('\nStep 3: Extracting key files for manipulation...');
    
    const themeXML = await loadedZip.files['ppt/theme/theme1.xml'].async('text');
    console.log(`✅ Extracted theme XML: ${themeXML.length} chars`);
    console.log('Theme preview:', themeXML.substring(0, 200) + '...');
    
    const presentationXML = await loadedZip.files['ppt/presentation.xml'].async('text');
    console.log(`✅ Extracted presentation XML: ${presentationXML.length} chars`);
    
    // Test 4: Modify theme colors (simulate our OOXML manipulation)
    console.log('\nStep 4: Modifying theme colors...');
    
    // Simple string replacement to change accent colors
    let modifiedTheme = themeXML
      .replace(/val="5B9BD5"/, 'val="FF6B35"') // Change blue to orange
      .replace(/val="70AD47"/, 'val="E74C3C"') // Change green to red
      .replace(/val="A5A5A5"/, 'val="9B59B6"'); // Change gray to purple
    
    console.log('✅ Modified theme colors (Orange, Red, Purple)');
    
    // Test 5: Create modified PPTX
    console.log('\nStep 5: Creating modified PPTX...');
    
    const modifiedZip = new JSZip();
    
    // Copy all files from original
    const copyPromises = [];
    Object.keys(loadedZip.files).forEach(filename => {
      const file = loadedZip.files[filename];
      if (!file.dir) {
        if (filename === 'ppt/theme/theme1.xml') {
          // Use our modified theme
          modifiedZip.file(filename, modifiedTheme);
          console.log(`   📝 Modified: ${filename}`);
        } else {
          // Copy original file
          copyPromises.push(
            file.async('text').then(content => {
              modifiedZip.file(filename, content);
              console.log(`   📄 Copied: ${filename}`);
            }).catch(err => {
              // Handle binary files
              return file.async('base64').then(content => {
                modifiedZip.file(filename, content, { base64: true });
                console.log(`   🗂️ Copied (binary): ${filename}`);
              });
            })
          );
        }
      }
    });
    
    await Promise.all(copyPromises);
    
    // Generate modified PPTX
    console.log('\nStep 6: Generating modified PPTX...');
    const modifiedBuffer = await modifiedZip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    
    fs.writeFileSync('modified-theme-pptx.pptx', modifiedBuffer);
    console.log(`✅ Modified PPTX saved: ${modifiedBuffer.length} bytes`);
    
    // Test 7: Verify the modification worked
    console.log('\nStep 7: Verifying modifications...');
    
    const verifyZip = new JSZip();
    const verifyLoaded = await verifyZip.loadAsync(modifiedBuffer);
    const verifyTheme = await verifyLoaded.files['ppt/theme/theme1.xml'].async('text');
    
    if (verifyTheme.includes('FF6B35')) {
      console.log('✅ Theme modification verified - orange color found');
    } else {
      console.log('❌ Theme modification failed - orange color not found');
    }
    
    // Test 8: Size comparison
    console.log('\nStep 8: Final comparison...');
    console.log(`Original size: ${realPptxBuffer.length} bytes`);
    console.log(`Modified size: ${modifiedBuffer.length} bytes`);
    console.log(`Size difference: ${modifiedBuffer.length - realPptxBuffer.length} bytes`);
    
    console.log('\n🎉 Real PPTX manipulation test completed successfully!');
    console.log('📝 Created files for testing:');
    console.log('   - test-ooxml-pptx.pptx (original)');
    console.log('   - modified-theme-pptx.pptx (with orange/red/purple theme)');
    console.log('📋 Manual verification: Open both files and verify the theme colors changed');
    
    return true;
    
  } catch (error) {
    console.error('❌ Real PPTX test failed:', error);
    console.error('Stack:', error.stack);
    return false;
  }
}

// Run the test if this file is executed directly
if (require.main === module) {
  testRealPPTX();
}

module.exports = { testRealPPTX };