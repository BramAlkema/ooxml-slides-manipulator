/**
 * Quick test that bypasses FileHandler MIME type issues
 */
function quickOOXMLTest() {
  try {
    console.log('🧪 Quick OOXML Core test...');
    
    // Create a test presentation
    const testPresentation = SlidesApp.create('Quick OOXML Test');
    const slide = testPresentation.getSlides()[0];
    slide.insertTextBox('Testing Universal OOXML Core');
    console.log(`✅ Test presentation: ${testPresentation.getId()}`);
    
    // Export as PPTX directly
    const pptxUrl = `https://docs.google.com/presentation/d/${testPresentation.getId()}/export/pptx`;
    const response = UrlFetchApp.fetch(pptxUrl, {
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`Export failed: ${response.getResponseCode()}`);
    }
    
    const pptxBlob = response.getBlob();
    console.log(`✅ Got PPTX blob: ${pptxBlob.getSize()} bytes`);
    
    // Test OOXMLCore directly with the blob
    console.log('🔧 Testing OOXMLCore with blob...');
    const core = new OOXMLCore(pptxBlob, 'pptx');
    
    // Extract files
    core.extract();
    console.log(`✅ Extracted ${core.files.size} files`);
    
    // Test getting a file
    const presentationXml = core.getFile('ppt/presentation.xml');
    if (presentationXml) {
      console.log('✅ Found presentation.xml');
      console.log(`📄 XML length: ${presentationXml.length} chars`);
    }
    
    // Test file modification
    const modifiedXml = presentationXml.replace(/Testing Universal OOXML Core/g, 'MODIFIED BY UNIVERSAL CORE');
    core.setFile('ppt/presentation.xml', modifiedXml);
    console.log('✅ Modified presentation.xml');
    
    // Test recompression
    const modifiedBlob = core.compress();
    console.log(`✅ Recompressed to ${modifiedBlob.getSize()} bytes`);
    
    // Save the modified file
    const modifiedFile = DriveApp.createFile(modifiedBlob.setName('QuickTest_Modified.pptx'));
    console.log(`✅ Saved modified file: ${modifiedFile.getId()}`);
    
    console.log('🎉 Universal OOXML Core test PASSED!');
    return {
      success: true,
      originalPresentation: testPresentation.getId(),
      modifiedFile: modifiedFile.getId(),
      extractedFiles: core.files.size
    };
    
  } catch (error) {
    console.error('💥 Quick test failed:', error);
    return { success: false, error: error.toString() };
  }
}