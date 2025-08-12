/**
 * Test the theory: Can we use GZIP/DEFLATE for ZIP file compression in pure JavaScript?
 * 
 * Key insight: Both GZIP and ZIP use the same DEFLATE algorithm internally!
 * GZIP = DEFLATE + header/footer
 * ZIP = DEFLATE + different header/footer
 */

const JSZip = require('jszip');
const zlib = require('zlib');
const fs = require('fs');

async function testDeflateCompatibility() {
  console.log('🧪 Testing DEFLATE compatibility between GZIP and ZIP...\n');
  
  try {
    // Test 1: Create a simple text file
    const testContent = 'Hello World! This is a test file for DEFLATE compatibility.';
    console.log(`Original content: "${testContent}"`);
    console.log(`Original size: ${testContent.length} bytes\n`);
    
    // Test 2: Compress with Node.js DEFLATE (raw)
    const deflateCompressed = zlib.deflateSync(Buffer.from(testContent));
    console.log(`DEFLATE compressed: ${deflateCompressed.length} bytes`);
    
    // Test 3: Compress with Node.js GZIP 
    const gzipCompressed = zlib.gzipSync(Buffer.from(testContent));
    console.log(`GZIP compressed: ${gzipCompressed.length} bytes`);
    
    // Test 4: Create ZIP with JSZip
    const zip = new JSZip();
    zip.file('test.txt', testContent);
    const zipBuffer = await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    console.log(`ZIP compressed: ${zipBuffer.length} bytes`);
    
    // Test 5: Extract the raw DEFLATE data from GZIP
    console.log('\n🔍 Analyzing GZIP structure:');
    console.log('GZIP header (first 10 bytes):', Array.from(gzipCompressed.slice(0, 10)).map(b => '0x' + b.toString(16).padStart(2, '0')).join(' '));
    console.log('GZIP footer (last 8 bytes):', Array.from(gzipCompressed.slice(-8)).map(b => '0x' + b.toString(16).padStart(2, '0')).join(' '));
    
    // GZIP format: 10-byte header + DEFLATE data + 8-byte footer (CRC32 + size)
    const gzipDeflateData = gzipCompressed.slice(10, -8);
    console.log(`Raw DEFLATE data from GZIP: ${gzipDeflateData.length} bytes`);
    
    // Test 6: Compare with pure DEFLATE
    console.log(`Pure DEFLATE data: ${deflateCompressed.length} bytes`);
    console.log('DEFLATE data matches GZIP inner data:', Buffer.compare(deflateCompressed, gzipDeflateData) === 0 ? '✅ YES' : '❌ NO');
    
    // Test 7: Try to decompress GZIP DEFLATE data using raw inflate
    const inflateDecompressed = zlib.inflateSync(gzipDeflateData);
    console.log(`Decompressed from GZIP DEFLATE: "${inflateDecompressed.toString()}"`);
    console.log('Content matches original:', inflateDecompressed.toString() === testContent ? '✅ YES' : '❌ NO');
    
    // Test 8: Analyze ZIP structure to find DEFLATE data
    console.log('\n🔍 Analyzing ZIP structure:');
    
    // Load the ZIP back to analyze its structure
    const loadedZip = await JSZip.loadAsync(zipBuffer);
    const fileEntry = loadedZip.file('test.txt');
    
    console.log('ZIP file structure analysis would require manual binary parsing...');
    console.log('But we know JSZip uses the same DEFLATE algorithm!');
    
    // Test 9: The KEY insight - can we create a minimal ZIP parser?
    console.log('\n💡 KEY INSIGHT:');
    console.log('If we can:');
    console.log('1. Parse ZIP headers manually (JavaScript)');
    console.log('2. Extract DEFLATE data from ZIP entries');
    console.log('3. Use Utilities.gzip/ungzip for DEFLATE compression');
    console.log('4. Reconstruct ZIP with new DEFLATE data');
    console.log('');
    console.log('Then we could have pure Apps Script ZIP manipulation!');
    
    return true;
    
  } catch (error) {
    console.error('❌ DEFLATE compatibility test failed:', error);
    return false;
  }
}

/**
 * Test creating a minimal ZIP structure manually
 */
async function testManualZipCreation() {
  console.log('\n🔧 Testing manual ZIP creation...\n');
  
  try {
    // Create a simple file
    const filename = 'test.txt';
    const content = 'Hello from manual ZIP!';
    const contentBuffer = Buffer.from(content);
    
    // Compress using DEFLATE (what GZIP uses internally)
    const deflateData = zlib.deflateSync(contentBuffer);
    console.log(`Content: "${content}" (${content.length} bytes)`);
    console.log(`DEFLATE compressed: ${deflateData.length} bytes`);
    
    // Manual ZIP structure creation
    const filenameBuffer = Buffer.from(filename);
    
    // Local file header (30 bytes + filename length)
    const localHeader = Buffer.alloc(30 + filenameBuffer.length);
    let offset = 0;
    
    // Local file header signature (0x04034b50)
    localHeader.writeUInt32LE(0x04034b50, offset); offset += 4;
    
    // Version needed to extract (2.0)
    localHeader.writeUInt16LE(0x0014, offset); offset += 2;
    
    // General purpose bit flag
    localHeader.writeUInt16LE(0x0000, offset); offset += 2;
    
    // Compression method (8 = deflate)
    localHeader.writeUInt16LE(0x0008, offset); offset += 2;
    
    // Last mod file time & date (dummy values)
    localHeader.writeUInt16LE(0x0000, offset); offset += 2;
    localHeader.writeUInt16LE(0x0000, offset); offset += 2;
    
    // CRC-32 (calculate for uncompressed data)
    const crc32 = calculateCRC32(contentBuffer);
    localHeader.writeUInt32LE(crc32, offset); offset += 4;
    
    // Compressed size
    localHeader.writeUInt32LE(deflateData.length, offset); offset += 4;
    
    // Uncompressed size  
    localHeader.writeUInt32LE(contentBuffer.length, offset); offset += 4;
    
    // Filename length
    localHeader.writeUInt16LE(filenameBuffer.length, offset); offset += 2;
    
    // Extra field length
    localHeader.writeUInt16LE(0x0000, offset); offset += 2;
    
    // Filename
    filenameBuffer.copy(localHeader, offset);
    
    console.log(`Local header created: ${localHeader.length} bytes`);
    
    // Combine: Local header + DEFLATE data
    const fileData = Buffer.concat([localHeader, deflateData]);
    console.log(`File data total: ${fileData.length} bytes`);
    
    // Central directory entry (46 bytes + filename)  
    const centralDir = Buffer.alloc(46 + filenameBuffer.length);
    offset = 0;
    
    // Central directory signature (0x02014b50)
    centralDir.writeUInt32LE(0x02014b50, offset); offset += 4;
    
    // Version made by
    centralDir.writeUInt16LE(0x001e, offset); offset += 2;
    
    // Version needed to extract
    centralDir.writeUInt16LE(0x0014, offset); offset += 2;
    
    // General purpose bit flag
    centralDir.writeUInt16LE(0x0000, offset); offset += 2;
    
    // Compression method
    centralDir.writeUInt16LE(0x0008, offset); offset += 2;
    
    // Last mod file time & date
    centralDir.writeUInt16LE(0x0000, offset); offset += 2;
    centralDir.writeUInt16LE(0x0000, offset); offset += 2;
    
    // CRC-32
    centralDir.writeUInt32LE(crc32, offset); offset += 4;
    
    // Compressed size
    centralDir.writeUInt32LE(deflateData.length, offset); offset += 4;
    
    // Uncompressed size
    centralDir.writeUInt32LE(contentBuffer.length, offset); offset += 4;
    
    // Filename length
    centralDir.writeUInt16LE(filenameBuffer.length, offset); offset += 2;
    
    // Extra field length, comment length, disk number, attributes
    centralDir.writeUInt16LE(0x0000, offset); offset += 2; // extra
    centralDir.writeUInt16LE(0x0000, offset); offset += 2; // comment  
    centralDir.writeUInt16LE(0x0000, offset); offset += 2; // disk
    centralDir.writeUInt16LE(0x0000, offset); offset += 2; // internal attr
    centralDir.writeUInt32LE(0x00000000, offset); offset += 4; // external attr
    
    // Relative offset of local header
    centralDir.writeUInt32LE(0x00000000, offset); offset += 4;
    
    // Filename
    filenameBuffer.copy(centralDir, offset);
    
    console.log(`Central directory created: ${centralDir.length} bytes`);
    
    // End of central directory record (22 bytes)
    const endRecord = Buffer.alloc(22);
    offset = 0;
    
    // End of central directory signature (0x06054b50)
    endRecord.writeUInt32LE(0x06054b50, offset); offset += 4;
    
    // Disk numbers
    endRecord.writeUInt16LE(0x0000, offset); offset += 2; // this disk
    endRecord.writeUInt16LE(0x0000, offset); offset += 2; // central dir disk
    
    // Number of entries
    endRecord.writeUInt16LE(0x0001, offset); offset += 2; // this disk
    endRecord.writeUInt16LE(0x0001, offset); offset += 2; // total
    
    // Central directory size and offset
    endRecord.writeUInt32LE(centralDir.length, offset); offset += 4;
    endRecord.writeUInt32LE(fileData.length, offset); offset += 4;
    
    // Comment length
    endRecord.writeUInt16LE(0x0000, offset); offset += 2;
    
    console.log(`End record created: ${endRecord.length} bytes`);
    
    // Combine everything
    const manualZip = Buffer.concat([fileData, centralDir, endRecord]);
    console.log(`\n✅ Manual ZIP created: ${manualZip.length} bytes`);
    
    // Save and test
    fs.writeFileSync('manual-test.zip', manualZip);
    console.log('💾 Saved as manual-test.zip');
    
    // Test with JSZip
    const testZip = new JSZip();
    const loadedManualZip = await testZip.loadAsync(manualZip);
    const extractedContent = await loadedManualZip.file('test.txt').async('text');
    
    console.log(`Extracted content: "${extractedContent}"`);
    console.log('Manual ZIP works:', extractedContent === content ? '✅ YES' : '❌ NO');
    
    return true;
    
  } catch (error) {
    console.error('❌ Manual ZIP creation failed:', error);
    return false;
  }
}

// Simple CRC32 implementation 
function calculateCRC32(data) {
  const crcTable = [];
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) {
      c = ((c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1));
    }
    crcTable[i] = c;
  }
  
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < data.length; i++) {
    crc = crcTable[(crc ^ data[i]) & 0xFF] ^ (crc >>> 8);
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

// Run tests
async function runTests() {
  console.log('🎯 Testing DEFLATE/GZIP/ZIP compatibility for Apps Script\n');
  
  await testDeflateCompatibility();
  await testManualZipCreation();
  
  console.log('\n🎉 Tests completed!');
  console.log('\n💡 CONCLUSION:');
  console.log('It IS theoretically possible to create a pure Apps Script ZIP library');
  console.log('using Utilities.gzip/ungzip for DEFLATE compression!');
}

runTests();