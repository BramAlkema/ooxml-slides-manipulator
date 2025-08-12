/**
 * DebugTest - Simple debugging functions to isolate issues
 */

/**
 * Test the ZIP creation step by step
 */
function debugZipCreation() {
  console.log('🔍 Debug: ZIP Creation Test');
  console.log('=' * 30);
  
  try {
    // Step 1: Test basic blob creation
    console.log('Step 1: Testing basic blob creation...');
    const testBlob = Utilities.newBlob('Hello World', 'text/plain', 'test.txt');
    console.log('✅ Basic blob created:', testBlob.getName(), 'Size:', testBlob.getBytes().length);
    
    // Step 2: Test XML blob creation
    console.log('Step 2: Testing XML blob creation...');
    const xmlContent = '<?xml version="1.0"?><root><test>Hello</test></root>';
    const xmlBlob = Utilities.newBlob(xmlContent, 'application/xml', 'test.xml');
    console.log('✅ XML blob created:', xmlBlob.getName(), 'Size:', xmlBlob.getBytes().length);
    
    // Step 3: Test array of blobs
    console.log('Step 3: Testing blob array...');
    const blobArray = [testBlob, xmlBlob];
    console.log('✅ Blob array created with', blobArray.length, 'items');
    
    // Step 4: Test ZIP creation
    console.log('Step 4: Testing ZIP creation...');
    const zipBlob = Utilities.zip(blobArray);
    console.log('✅ ZIP created:', zipBlob.getName(), 'Size:', zipBlob.getBytes().length);
    
    // Step 5: Test ZIP extraction
    console.log('Step 5: Testing ZIP extraction...');
    const extractedFiles = Utilities.unzip(zipBlob);
    console.log('✅ ZIP extracted:', extractedFiles.length, 'files');
    
    console.log('🎉 All ZIP operations working correctly!');
    return true;
    
  } catch (error) {
    console.error('❌ ZIP creation failed:', error);
    return false;
  }
}

/**
 * Test minimal XML creation
 */
function debugMinimalXML() {
  console.log('🔍 Debug: Minimal XML Creation');
  console.log('=' * 32);
  
  try {
    // Test each XML file creation individually
    const xmlFiles = [
      { name: 'Content Types', content: PPTXTemplate._getContentTypesXML() },
      { name: 'Main Rels', content: PPTXTemplate._getMainRelsXML() },
      { name: 'Presentation', content: PPTXTemplate._getPresentationXML() },
      { name: 'Theme', content: PPTXTemplate._getThemeXML() }
    ];
    
    xmlFiles.forEach(file => {
      try {
        console.log(`Testing ${file.name} XML...`);
        const blob = Utilities.newBlob(file.content, MimeType.PLAIN_TEXT, `${file.name}.xml`);
        console.log(`✅ ${file.name}: ${blob.getBytes().length} bytes`);
      } catch (error) {
        console.log(`❌ ${file.name}: ${error.message}`);
      }
    });
    
    return true;
    
  } catch (error) {
    console.error('❌ XML creation failed:', error);
    return false;
  }
}

/**
 * Test step-by-step template creation
 */
function debugStepByStepTemplate() {
  console.log('🔍 Debug: Step-by-Step Template Creation');
  console.log('=' * 42);
  
  try {
    console.log('Step 1: Creating individual XML blobs...');
    const files = [];
    
    // Create each file individually with error checking
    const filesToCreate = [
      { name: '[Content_Types].xml', content: PPTXTemplate._getContentTypesXML() },
      { name: '_rels/.rels', content: PPTXTemplate._getMainRelsXML() },
      { name: 'ppt/presentation.xml', content: PPTXTemplate._getPresentationXML() },
      { name: 'ppt/theme/theme1.xml', content: PPTXTemplate._getThemeXML() }
    ];
    
    filesToCreate.forEach((fileInfo, index) => {
      try {
        console.log(`  Creating file ${index + 1}: ${fileInfo.name}`);
        const blob = Utilities.newBlob(fileInfo.content, MimeType.PLAIN_TEXT, fileInfo.name);
        files.push(blob);
        console.log(`  ✅ Created: ${blob.getBytes().length} bytes`);
      } catch (error) {
        console.log(`  ❌ Failed to create ${fileInfo.name}: ${error.message}`);
        throw error;
      }
    });
    
    console.log(`Step 2: All ${files.length} files created successfully`);
    
    console.log('Step 3: Creating ZIP using tanaikech approach...');
    const zipBlob = Utilities.zip(files, 'debug-template.pptx');
    console.log(`✅ ZIP created: ${zipBlob.getBytes().length} bytes`);
    
    console.log('Step 4: Setting MIME type...');
    const finalBlob = zipBlob.setContentType(MimeType.MICROSOFT_POWERPOINT);
    console.log(`✅ Final blob: ${finalBlob.getName()}, ${finalBlob.getBytes().length} bytes`);
    
    console.log('Step 5: Testing extraction...');
    const extracted = Utilities.unzip(finalBlob);
    console.log(`✅ Extracted ${extracted.length} files from template`);
    
    console.log('🎉 Step-by-step template creation successful!');
    return finalBlob;
    
  } catch (error) {
    console.error('❌ Step-by-step template failed:', error);
    return null;
  }
}

/**
 * Simple working PPTX test
 */
function createSimplestPPTX() {
  console.log('🎯 Creating Simplest Possible PPTX');
  console.log('=' * 35);
  
  try {
    // Create the absolute minimum files needed for a PPTX
    const files = [];
    
    // Minimal Content Types
    const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-presentationml.presentation.main+xml"/>
</Types>`;
    
    // Minimal Main Relationships
    const mainRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;
    
    // Minimal Presentation
    const presentation = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>`;
    
    // Create blobs using tanaikech's approach
    files.push(Utilities.newBlob(contentTypes, MimeType.PLAIN_TEXT, '[Content_Types].xml'));
    files.push(Utilities.newBlob(mainRels, MimeType.PLAIN_TEXT, '_rels/.rels'));
    files.push(Utilities.newBlob(presentation, MimeType.PLAIN_TEXT, 'ppt/presentation.xml'));
    
    console.log(`Created ${files.length} minimal files`);
    
    // Create ZIP using tanaikech's approach
    const zipBlob = Utilities.zip(files, 'minimal.pptx');
    const pptxBlob = zipBlob.setContentType(MimeType.MICROSOFT_POWERPOINT);
    
    console.log(`✅ Minimal PPTX created: ${pptxBlob.getBytes().length} bytes`);
    
    // Save it to test
    const file = DriveApp.createFile(pptxBlob);
    console.log(`💾 Saved to Drive: ${file.getId()}`);
    console.log(`🔗 URL: ${file.getUrl()}`);
    
    return file.getId();
    
  } catch (error) {
    console.error('❌ Simplest PPTX creation failed:', error);
    return null;
  }
}

/**
 * Run all debug tests
 */
function runDebugTests() {
  console.log('🐛 Running Debug Tests');
  console.log('=' * 23);
  
  const results = [];
  
  results.push({ name: 'ZIP Creation', result: debugZipCreation() });
  console.log('');
  
  results.push({ name: 'Minimal XML', result: debugMinimalXML() });
  console.log('');
  
  results.push({ name: 'Step-by-Step Template', result: debugStepByStepTemplate() });
  console.log('');
  
  results.push({ name: 'Simplest PPTX', result: createSimplestPPTX() });
  
  console.log('\n📊 Debug Results:');
  results.forEach(test => {
    const status = test.result ? '✅' : '❌';
    console.log(`${status} ${test.name}`);
  });
  
  return results;
}