/**
 * Basic tests for OOXMLSlides library
 * Run these in Google Apps Script to test functionality
 */

// Test configuration - update with your actual file IDs
const TEST_CONFIG = {
  // Replace with actual Google Drive PPTX file ID for testing
  SAMPLE_PPTX_ID: 'your-test-file-id-here',
  TEST_FOLDER_ID: 'your-test-folder-id-here' // Optional
};

/**
 * Test 1: Basic library loading and file access
 */
function test1_BasicLoading() {
  console.log('=== Test 1: Basic Loading ===');
  
  try {
    // Test file ID validation
    const isValid = Validators.isValidFileId(TEST_CONFIG.SAMPLE_PPTX_ID);
    console.log('File ID validation:', isValid);
    
    // Test file handler
    const fileHandler = new FileHandler();
    const hasAccess = fileHandler.validateAccess(TEST_CONFIG.SAMPLE_PPTX_ID);
    console.log('File access:', hasAccess);
    
    if (hasAccess) {
      const fileInfo = fileHandler.getFileInfo(TEST_CONFIG.SAMPLE_PPTX_ID);
      console.log('File info:', fileInfo);
    }
    
    console.log('✅ Test 1 passed');
    return true;
    
  } catch (error) {
    console.error('❌ Test 1 failed:', error.message);
    return false;
  }
}

/**
 * Test 2: OOXML parsing
 */
function test2_OOXMLParsing() {
  console.log('=== Test 2: OOXML Parsing ===');
  
  try {
    const fileHandler = new FileHandler();
    const blob = fileHandler.loadFile(TEST_CONFIG.SAMPLE_PPTX_ID);
    console.log('File loaded, size:', blob.getSize());
    
    const parser = new OOXMLParser(blob);
    parser.extract();
    
    const files = parser.listFiles();
    console.log('OOXML files found:', files.length);
    console.log('Sample files:', files.slice(0, 5));
    
    // Test theme XML access
    if (parser.hasFile('ppt/theme/theme1.xml')) {
      const themeXML = parser.getThemeXML();
      console.log('Theme XML loaded successfully');
    }
    
    // Test presentation XML access
    if (parser.hasFile('ppt/presentation.xml')) {
      const presXML = parser.getPresentationXML();
      console.log('Presentation XML loaded successfully');
    }
    
    console.log('✅ Test 2 passed');
    return true;
    
  } catch (error) {
    console.error('❌ Test 2 failed:', error.message);
    return false;
  }
}

/**
 * Test 3: Theme editing
 */
function test3_ThemeEditing() {
  console.log('=== Test 3: Theme Editing ===');
  
  try {
    const slides = OOXMLSlides.fromFile(TEST_CONFIG.SAMPLE_PPTX_ID);
    
    // Get current theme
    const currentTheme = slides.getTheme();
    console.log('Current colors:', Object.keys(currentTheme.colors));
    console.log('Current fonts:', currentTheme.fonts);
    
    // Test color validation
    const testColors = ['#FF0000', '#00FF00', '#0000FF', '#FFFF00'];
    testColors.forEach(color => {
      const isValid = Validators.isValidHexColor(color);
      console.log(`Color ${color} valid:`, isValid);
    });
    
    // Test font validation
    const testFonts = ['Arial', 'Calibri', 'Times New Roman'];
    testFonts.forEach(font => {
      const isValid = Validators.isValidFontName(font);
      console.log(`Font ${font} valid:`, isValid);
    });
    
    console.log('✅ Test 3 passed');
    return true;
    
  } catch (error) {
    console.error('❌ Test 3 failed:', error.message);
    return false;
  }
}

/**
 * Test 4: Full modification workflow (READ-ONLY - doesn't save)
 */
function test4_ModificationWorkflow() {
  console.log('=== Test 4: Modification Workflow ===');
  
  try {
    const slides = OOXMLSlides.fromFile(TEST_CONFIG.SAMPLE_PPTX_ID);
    
    // Test size operations
    const currentSize = slides.getSize();
    console.log('Current slide size:', currentSize);
    
    // Test theme operations (in memory only)
    slides.setColors(['#2C3E50', '#FFFFFF', '#34495E', '#ECF0F1', '#3498DB', '#E74C3C']);
    slides.setFonts('Arial', 'Calibri');
    
    // Verify changes in memory
    const modifiedTheme = slides.getTheme();
    console.log('Modified theme colors:', Object.keys(modifiedTheme.colors));
    
    // Test build process (creates blob but doesn't save)
    const modifiedBlob = slides.parser.build();
    console.log('Modified PPTX blob size:', modifiedBlob.getSize());
    
    console.log('✅ Test 4 passed (read-only)');
    return true;
    
  } catch (error) {
    console.error('❌ Test 4 failed:', error.message);
    return false;
  }
}

/**
 * Test 5: Actual save operation (creates new file)
 * Only run this if you want to create a test file
 */
function test5_SaveOperation() {
  console.log('=== Test 5: Save Operation ===');
  
  const CREATE_TEST_FILE = false; // Set to true to actually save
  
  if (!CREATE_TEST_FILE) {
    console.log('⏭️ Test 5 skipped (set CREATE_TEST_FILE to true to run)');
    return true;
  }
  
  try {
    const slides = OOXMLSlides.fromFile(TEST_CONFIG.SAMPLE_PPTX_ID, {
      createBackup: false // Don't create backup for test
    });
    
    // Apply simple modifications
    slides.setColors(['#FF0000', '#FFFFFF', '#000000', '#CCCCCC', '#0000FF', '#00FF00']);
    slides.setFonts('Arial', 'Arial');
    
    // Save as new file
    const newFileId = slides.save({
      name: `OOXMLSlides_Test_${new Date().getTime()}`,
      folderId: TEST_CONFIG.TEST_FOLDER_ID,
      importToSlides: false // Keep as PPTX for testing
    });
    
    console.log('✅ Test file created:', newFileId);
    
    // Verify the saved file
    const fileHandler = new FileHandler();
    const newFileInfo = fileHandler.getFileInfo(newFileId);
    console.log('New file info:', newFileInfo);
    
    console.log('✅ Test 5 passed');
    return true;
    
  } catch (error) {
    console.error('❌ Test 5 failed:', error.message);
    return false;
  }
}

/**
 * Run all tests
 */
function runAllTests() {
  console.log('🚀 Starting OOXMLSlides Tests');
  console.log('Make sure to update TEST_CONFIG with valid file IDs');
  
  const results = [];
  
  results.push({ test: 'Basic Loading', passed: test1_BasicLoading() });
  results.push({ test: 'OOXML Parsing', passed: test2_OOXMLParsing() });
  results.push({ test: 'Theme Editing', passed: test3_ThemeEditing() });
  results.push({ test: 'Modification Workflow', passed: test4_ModificationWorkflow() });
  results.push({ test: 'Save Operation', passed: test5_SaveOperation() });
  
  console.log('📊 Test Results:');
  results.forEach(result => {
    console.log(`${result.passed ? '✅' : '❌'} ${result.test}`);
  });
  
  const passedCount = results.filter(r => r.passed).length;
  console.log(`\n${passedCount}/${results.length} tests passed`);
  
  return results;
}

/**
 * Quick validation test that doesn't require file access
 */
function quickValidationTest() {
  console.log('=== Quick Validation Test ===');
  
  // Test validators without file access
  const colorTests = [
    '#FF0000', '#f00', 'FF0000', 'invalid', '#GGGGGG'
  ];
  
  colorTests.forEach(color => {
    const isValid = Validators.isValidHexColor(color);
    const normalized = isValid ? Validators.normalizeHexColor(color) : 'invalid';
    console.log(`${color} -> valid: ${isValid}, normalized: ${normalized}`);
  });
  
  // Test file ID validation
  const fileIdTests = [
    'valid_file_id_123456789',
    'invalid file id with spaces',
    '1234567890123456789012345678901234567890123456789', // too long
    '' // empty
  ];
  
  fileIdTests.forEach(fileId => {
    const isValid = Validators.isValidFileId(fileId);
    console.log(`"${fileId}" -> valid: ${isValid}`);
  });
  
  console.log('✅ Quick validation test completed');
}

// Helper function to get library info
function getLibraryInfo() {
  return {
    version: '1.0.0',
    description: 'OOXML Slides Manipulator for Google Apps Script',
    classes: ['OOXMLSlides', 'OOXMLParser', 'ThemeEditor', 'SlideManager', 'FileHandler', 'Validators', 'ErrorHandler'],
    features: [
      'Theme and color palette editing',
      'Font pair modification', 
      'Slide master manipulation',
      'Slide size adjustment',
      'Google Drive integration',
      'Error handling and validation'
    ]
  };
}