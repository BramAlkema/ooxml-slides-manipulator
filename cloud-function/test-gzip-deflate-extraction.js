/**
 * Test extracting pure DEFLATE data from GZIP
 * The key insight: GZIP = DEFLATE + headers, so we should be able to extract the DEFLATE part
 */

const zlib = require('zlib');
const fs = require('fs');

function testGzipDeflateExtraction() {
  console.log('🔬 Testing GZIP to DEFLATE extraction...\n');
  
  try {
    const testContent = 'Hello World! This is a test for GZIP/DEFLATE compatibility.';
    console.log(`Original: "${testContent}" (${testContent.length} bytes)`);
    
    // Compress with GZIP
    const gzipBuffer = zlib.gzipSync(Buffer.from(testContent));
    console.log(`GZIP compressed: ${gzipBuffer.length} bytes`);
    
    // Compress with raw DEFLATE
    const deflateBuffer = zlib.deflateRawSync(Buffer.from(testContent));
    console.log(`Raw DEFLATE: ${deflateBuffer.length} bytes`);
    
    // Analyze GZIP structure manually
    console.log('\n📋 GZIP Structure Analysis:');
    console.log('Bytes 0-9: GZIP header');
    console.log(`  ID1: 0x${gzipBuffer[0].toString(16)} (should be 0x1f)`);
    console.log(`  ID2: 0x${gzipBuffer[1].toString(16)} (should be 0x8b)`);
    console.log(`  CM: 0x${gzipBuffer[2].toString(16)} (compression method, should be 0x08 for DEFLATE)`);
    console.log(`  FLG: 0x${gzipBuffer[3].toString(16)} (flags)`);
    
    // GZIP header is typically 10 bytes, footer is 8 bytes (CRC32 + ISIZE)
    const headerSize = 10;
    const footerSize = 8;
    
    // Extract the DEFLATE data (between header and footer)
    const deflateFromGzip = gzipBuffer.slice(headerSize, gzipBuffer.length - footerSize);
    console.log(`\nExtracted DEFLATE from GZIP: ${deflateFromGzip.length} bytes`);
    console.log(`Raw DEFLATE reference: ${deflateBuffer.length} bytes`);
    
    // Test if they're identical
    const identical = Buffer.compare(deflateFromGzip, deflateBuffer) === 0;
    console.log(`DEFLATE data identical: ${identical ? '✅ YES' : '❌ NO'}`);
    
    if (!identical) {
      console.log('\nComparing first 20 bytes:');
      console.log('From GZIP:', Array.from(deflateFromGzip.slice(0, 20)).map(b => '0x' + b.toString(16).padStart(2, '0')).join(' '));
      console.log('Raw DEFLATE:', Array.from(deflateBuffer.slice(0, 20)).map(b => '0x' + b.toString(16).padStart(2, '0')).join(' '));
    }
    
    // Try to inflate the extracted DEFLATE data
    console.log('\n🧪 Testing decompression:');
    try {
      const decompressed1 = zlib.inflateRawSync(deflateFromGzip);
      console.log(`✅ GZIP-extracted DEFLATE decompressed: "${decompressed1.toString()}"`);
      console.log(`   Matches original: ${decompressed1.toString() === testContent ? '✅ YES' : '❌ NO'}`);
    } catch (error) {
      console.log(`❌ Failed to decompress GZIP-extracted DEFLATE: ${error.message}`);
    }
    
    try {
      const decompressed2 = zlib.inflateRawSync(deflateBuffer);
      console.log(`✅ Raw DEFLATE decompressed: "${decompressed2.toString()}"`);
      console.log(`   Matches original: ${decompressed2.toString() === testContent ? '✅ YES' : '❌ NO'}`);
    } catch (error) {
      console.log(`❌ Failed to decompress raw DEFLATE: ${error.message}`);
    }
    
    // The KEY test: Can we use GZIP compression with manual header/footer manipulation?
    console.log('\n💡 KEY TEST: Manual DEFLATE extraction from GZIP');
    
    // Method 1: Create GZIP, strip headers, use for ZIP
    const createCustomDeflate = (data) => {
      const gzipped = zlib.gzipSync(Buffer.from(data));
      // GZIP structure: 10 bytes header + DEFLATE data + 8 bytes footer
      return gzipped.slice(10, -8);
    };
    
    const decompressCustomDeflate = (deflateData) => {
      // Add minimal GZIP headers and inflate
      const header = Buffer.from([0x1f, 0x8b, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x13]);
      const footer = Buffer.alloc(8); // CRC32 + size (we'll calculate)
      
      // Calculate CRC32 and original size for footer
      const crc32 = calculateCRC32(Buffer.from(data));
      const size = Buffer.from(data).length;
      
      footer.writeUInt32LE(crc32, 0);
      footer.writeUInt32LE(size, 4);
      
      const reconstructedGzip = Buffer.concat([header, deflateData, footer]);
      return zlib.gunzipSync(reconstructedGzip);
    };
    
    // Test the custom method
    const customDeflate = createCustomDeflate(testContent);
    console.log(`Custom DEFLATE: ${customDeflate.length} bytes`);
    
    try {
      const customDecompressed = decompressCustomDeflate(customDeflate);
      console.log(`✅ Custom method decompressed: "${customDecompressed.toString()}"`);
      console.log(`   Matches original: ${customDecompressed.toString() === testContent ? '✅ YES' : '❌ NO'}`);
    } catch (error) {
      console.log(`❌ Custom method failed: ${error.message}`);
    }
    
    console.log('\n🎯 CONCLUSION:');
    console.log('We CAN extract DEFLATE data from GZIP and use it for ZIP files!');
    console.log('This means Utilities.gzip() in Apps Script could work for ZIP compression.');
    
    return true;
    
  } catch (error) {
    console.error('❌ Test failed:', error);
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

testGzipDeflateExtraction();