/**
 * Test JSZip OOXML compatibility specifically for PPTX files
 * Based on PptxGenJS's successful approach to verify our implementation
 */

const JSZip = require('jszip');
const fs = require('fs');

async function testOOXMLCompatibility() {
  console.log('🧪 Testing JSZip OOXML compatibility for PPTX files...\n');
  
  try {
    // Test 1: Create a PPTX file using PptxGenJS's approach
    console.log('Test 1: Creating PPTX using PptxGenJS approach...');
    
    const zip = new JSZip();
    
    // Create the exact OOXML structure that PptxGenJS uses
    zip.folder('_rels');
    zip.folder('docProps');
    zip.folder('ppt');
    zip.folder('ppt/slides');
    zip.folder('ppt/slides/_rels');
    zip.folder('ppt/slideLayouts');
    zip.folder('ppt/slideLayouts/_rels');
    zip.folder('ppt/slideMasters');
    zip.folder('ppt/slideMasters/_rels');
    zip.folder('ppt/theme');
    zip.folder('ppt/_rels');
    
    // Add required OOXML files
    zip.file('[Content_Types].xml', createContentTypes());
    zip.file('_rels/.rels', createMainRels());
    zip.file('docProps/core.xml', createCoreProps());
    zip.file('docProps/app.xml', createAppProps());
    zip.file('ppt/presentation.xml', createPresentation());
    zip.file('ppt/_rels/presentation.xml.rels', createPresentationRels());
    zip.file('ppt/theme/theme1.xml', createTheme());
    zip.file('ppt/slideMasters/slideMaster1.xml', createSlideMaster());
    zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', createSlideMasterRels());
    zip.file('ppt/slideLayouts/slideLayout1.xml', createSlideLayout());
    zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', createSlideLayoutRels());
    zip.file('ppt/slides/slide1.xml', createSlide());
    zip.file('ppt/slides/_rels/slide1.xml.rels', createSlideRels());
    
    // Generate the PPTX using PptxGenJS's exact approach
    console.log('Generating PPTX with DEFLATE compression...');
    const pptxBuffer = await zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    
    console.log(`✅ PPTX created: ${pptxBuffer.length} bytes`);
    
    // Save for manual testing
    fs.writeFileSync('test-ooxml-pptx.pptx', pptxBuffer);
    console.log('💾 Saved as test-ooxml-pptx.pptx');
    
    // Test 2: Round-trip test - unzip and rezip
    console.log('\nTest 2: Round-trip unzip/rezip test...');
    
    const unzipTest = new JSZip();
    const loadedZip = await unzipTest.loadAsync(pptxBuffer);
    
    console.log(`Loaded ZIP with ${Object.keys(loadedZip.files).length} files`);
    
    // Extract all files
    const extractedFiles = {};
    const promises = [];
    
    Object.keys(loadedZip.files).forEach(filename => {
      if (!loadedZip.files[filename].dir) {
        promises.push(
          loadedZip.files[filename].async('text').then(content => {
            extractedFiles[filename] = content;
            console.log(`   Extracted: ${filename} (${content.length} chars)`);
          }).catch(err => {
            // Handle binary files
            return loadedZip.files[filename].async('base64').then(content => {
              extractedFiles[filename] = {
                type: 'binary',
                content: content
              };
              console.log(`   Extracted (binary): ${filename} (${content.length} chars base64)`);
            });
          })
        );
      }
    });
    
    await Promise.all(promises);
    
    console.log(`✅ Extracted ${Object.keys(extractedFiles).length} files`);
    
    // Test 3: Rezip the extracted files
    console.log('\nTest 3: Rezip extracted files...');
    
    const rezipTest = new JSZip();
    
    Object.entries(extractedFiles).forEach(([filename, content]) => {
      if (typeof content === 'object' && content.type === 'binary') {
        rezipTest.file(filename, content.content, { base64: true });
      } else {
        rezipTest.file(filename, content);
      }
    });
    
    const rezippedBuffer = await rezipTest.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });
    
    console.log(`✅ Rezipped PPTX: ${rezippedBuffer.length} bytes`);
    
    // Save for comparison
    fs.writeFileSync('test-ooxml-rezipped.pptx', rezippedBuffer);
    console.log('💾 Saved as test-ooxml-rezipped.pptx');
    
    // Test 4: Size comparison
    console.log('\nTest 4: File size comparison...');
    console.log(`Original: ${pptxBuffer.length} bytes`);
    console.log(`Rezipped: ${rezippedBuffer.length} bytes`);
    console.log(`Difference: ${rezippedBuffer.length - pptxBuffer.length} bytes`);
    
    if (Math.abs(rezippedBuffer.length - pptxBuffer.length) < 100) {
      console.log('✅ Size difference is minimal - good sign!');
    } else {
      console.log('⚠️ Significant size difference - may indicate issues');
    }
    
    console.log('\n🎉 OOXML compatibility test completed!');
    console.log('📝 Manual verification required:');
    console.log('   1. Try opening test-ooxml-pptx.pptx in PowerPoint');
    console.log('   2. Try opening test-ooxml-rezipped.pptx in PowerPoint');
    console.log('   3. Both should open without corruption errors');
    
    return true;
    
  } catch (error) {
    console.error('❌ OOXML compatibility test failed:', error);
    return false;
  }
}

// Helper functions to create OOXML files (based on OOXML specification)
function createContentTypes() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-presentationml.slide+xml"/>
</Types>`;
}

function createMainRels() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
}

function createCoreProps() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>OOXML Test Presentation</dc:title>
  <dc:creator>JSZip Test</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:modified>
</cp:coreProperties>`;
}

function createAppProps() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>JSZip OOXML Test</Application>
  <PresentationFormat>On-screen Show (4:3)</PresentationFormat>
  <Slides>1</Slides>
  <TotalTime>0</TotalTime>
</Properties>`;
}

function createPresentation() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`;
}

function createPresentationRels() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`;
}

function createTheme() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test Theme">
  <a:themeElements>
    <a:clrScheme name="Test">
      <a:dk1><a:sysClr val="windowText"/></a:dk1>
      <a:lt1><a:sysClr val="window"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="5B9BD5"/></a:accent1>
      <a:accent2><a:srgbClr val="70AD47"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="4472C4"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Test">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"/>
  </a:themeElements>
</a:theme>`;
}

function createSlideMaster() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>`;
}

function createSlideMasterRels() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;
}

function createSlideLayout() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" type="title" preserve="1">
  <p:cSld name="Title Slide">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
</p:sldLayout>`;
}

function createSlideLayoutRels() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;
}

function createSlide() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
</p:sld>`;
}

function createSlideRels() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;
}

// Run the test if this file is executed directly
if (require.main === module) {
  testOOXMLCompatibility();
}

module.exports = { testOOXMLCompatibility };