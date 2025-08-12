/**
 * Test XML Search & Replace and Language Standardization on Real Analysis Files
 * Tests the complete integration using the actual PPTX files from /analysis folder
 */

function testAnalysisFiles() {
  try {
    console.log('=== Testing XML Search & Replace and Language Standardization ===');
    
    // Initialize OOXML system
    const ooxml = new OOXMLSlides();
    
    // Test 1: Create a new presentation and apply XML operations
    console.log('\n1. Creating test presentation with XML search & replace...');
    
    const testSlide = ooxml.createBlankPresentation()
      .addSlide()
      .addTitle('XML Search & Replace Test')
      .addContent('This is a test paragraph with mixed language tags.')
      .addContent('Another paragraph for testing language standardization.');
    
    // Test XML search and replace functionality
    console.log('\n2. Testing XML search & replace operations...');
    
    const searchReplaceConfig = {
      operations: [
        {
          searchPattern: 'test',
          replaceWith: 'verified',
          options: { caseSensitive: false, wholeWord: false }
        },
        {
          searchPattern: 'paragraph',
          replaceWith: 'section',
          options: { caseSensitive: false, wholeWord: true }
        }
      ],
      scope: 'slides'
    };
    
    const xmlSearchResult = testSlide.searchReplaceXML(searchReplaceConfig);
    console.log('XML Search & Replace Result:', xmlSearchResult);
    
    // Test 3: Language standardization
    console.log('\n3. Testing language standardization...');
    
    const languageResult = testSlide.standardizeLanguage('en-US', {
      includeSlides: true,
      includeMasters: true,
      includeLayouts: true,
      preserveSpecialElements: true
    });
    console.log('Language Standardization Result:', languageResult);
    
    // Test 4: Export and analyze
    console.log('\n4. Exporting test presentation...');
    
    const exportedBlob = testSlide.exportToPPTX();
    const exportedFile = DriveApp.createFile(exportedBlob)
      .setName(`XMLSearchReplace_Test_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss')}.pptx`);
    
    console.log('Test presentation created:', exportedFile.getUrl());
    
    // Test 5: Advanced XML analysis on exported file
    console.log('\n5. Analyzing exported file XML structure...');
    
    const reloadedOOXML = new OOXMLSlides();
    const analysisResult = reloadedOOXML.loadFromBlob(exportedBlob);
    
    // Perform advanced XML search to verify changes
    const verificationConfig = {
      operations: [
        {
          searchPattern: 'verified',
          replaceWith: 'verified', // No change, just search
          options: { searchOnly: true }
        }
      ],
      scope: 'all'
    };
    
    const verificationResult = reloadedOOXML.searchReplaceXML(verificationConfig);
    console.log('Verification Search Result:', verificationResult);
    
    // Test 6: Typography integration with XML operations
    console.log('\n6. Testing typography + XML search & replace integration...');
    
    const combinedTest = reloadedOOXML
      .applyProfessionalTypography('editorial')
      .searchReplaceXML({
        operations: [
          {
            searchPattern: '<a:latin typeface="[^"]*"',
            replaceWith: '<a:latin typeface="Calibri"',
            options: { useRegex: true }
          }
        ],
        scope: 'theme'
      });
    
    console.log('Combined Typography + XML Result:', combinedTest);
    
    return {
      success: true,
      testFile: exportedFile.getUrl(),
      results: {
        xmlSearchReplace: xmlSearchResult,
        languageStandardization: languageResult,
        verification: verificationResult,
        combinedTest: combinedTest
      }
    };
    
  } catch (error) {
    console.error('Test failed:', error);
    return {
      success: false,
      error: error.toString(),
      stack: error.stack
    };
  }
}

/**
 * Test language standardization on a complex multi-language scenario
 */
function testComplexLanguageStandardization() {
  try {
    console.log('=== Testing Complex Language Standardization ===');
    
    const ooxml = new OOXMLSlides();
    
    // Create a presentation with intentionally mixed language tags
    const mixedLanguageSlide = ooxml.createBlankPresentation()
      .addSlide()
      .addTitle('Multi-Language Test Presentation')
      .addContent('English text with embedded français words and 中文 characters.')
      .addContent('Another mixed paragraph with Español and Deutsch elements.');
    
    // Analyze language tags before standardization
    console.log('\n1. Analyzing original language distribution...');
    
    const originalAnalysis = mixedLanguageSlide.searchReplaceXML({
      operations: [
        {
          searchPattern: 'lang="[^"]*"',
          replaceWith: 'lang="$1"', // No change, just analyze
          options: { useRegex: true, searchOnly: true }
        }
      ],
      scope: 'all'
    });
    
    console.log('Original Language Tags:', originalAnalysis);
    
    // Standardize to English (US)
    console.log('\n2. Standardizing all language tags to en-US...');
    
    const standardizationResult = mixedLanguageSlide.standardizeLanguage('en-US', {
      includeSlides: true,
      includeMasters: true,
      includeLayouts: true,
      includeTheme: true,
      preserveSpecialElements: false, // Force standardization
      reportChanges: true
    });
    
    console.log('Standardization Result:', standardizationResult);
    
    // Verify standardization
    console.log('\n3. Verifying language standardization...');
    
    const verificationAnalysis = mixedLanguageSlide.searchReplaceXML({
      operations: [
        {
          searchPattern: 'lang="en-US"',
          replaceWith: 'lang="en-US"', // No change, just count
          options: { searchOnly: true }
        },
        {
          searchPattern: 'lang="(?!en-US)[^"]*"',
          replaceWith: 'lang="$1"', // Find non-English tags
          options: { useRegex: true, searchOnly: true }
        }
      ],
      scope: 'all'
    });
    
    console.log('Verification Analysis:', verificationAnalysis);
    
    // Export result
    const exportedBlob = mixedLanguageSlide.exportToPPTX();
    const exportedFile = DriveApp.createFile(exportedBlob)
      .setName(`LanguageStandardization_Test_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss')}.pptx`);
    
    console.log('Language standardization test file:', exportedFile.getUrl());
    
    return {
      success: true,
      testFile: exportedFile.getUrl(),
      originalAnalysis: originalAnalysis,
      standardizationResult: standardizationResult,
      verificationAnalysis: verificationAnalysis
    };
    
  } catch (error) {
    console.error('Complex language test failed:', error);
    return {
      success: false,
      error: error.toString(),
      stack: error.stack
    };
  }
}

/**
 * Run complete test suite for XML operations
 */
function runCompleteXMLTestSuite() {
  console.log('=== Running Complete XML Test Suite ===');
  
  const results = {
    timestamp: new Date().toISOString(),
    tests: {}
  };
  
  // Test 1: Basic XML Search & Replace
  console.log('\n--- Test 1: Basic XML Operations ---');
  results.tests.basicXMLOperations = testAnalysisFiles();
  
  // Test 2: Complex Language Standardization
  console.log('\n--- Test 2: Complex Language Standardization ---');
  results.tests.complexLanguageStandardization = testComplexLanguageStandardization();
  
  // Test 3: Integration with existing systems
  console.log('\n--- Test 3: Integration Test ---');
  try {
    const ooxml = new OOXMLSlides();
    const integrationTest = ooxml.createBlankPresentation()
      .addSlide()
      .addTitle('Integration Test')
      .applyCustomColors({ primaryColor: '#FF6B35', secondaryColor: '#004E89' })
      .applyProfessionalTypography('corporate')
      .searchReplaceXML({
        operations: [
          {
            searchPattern: 'Integration',
            replaceWith: 'Complete Integration',
            options: { caseSensitive: false }
          }
        ],
        scope: 'slides'
      })
      .standardizeLanguage('en-US');
    
    const integrationBlob = integrationTest.exportToPPTX();
    const integrationFile = DriveApp.createFile(integrationBlob)
      .setName(`Integration_Test_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss')}.pptx`);
    
    results.tests.integration = {
      success: true,
      testFile: integrationFile.getUrl()
    };
    
    console.log('Integration test file:', integrationFile.getUrl());
    
  } catch (error) {
    results.tests.integration = {
      success: false,
      error: error.toString()
    };
  }
  
  console.log('\n=== Test Suite Complete ===');
  console.log('Results Summary:', JSON.stringify(results, null, 2));
  
  return results;
}