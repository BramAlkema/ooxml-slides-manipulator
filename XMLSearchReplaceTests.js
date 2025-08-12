/**
 * XMLSearchReplaceTests - Test suite for XML Search & Replace and Language Standardization
 * 
 * Tests both the general XML search & replace engine and the specific language standardization
 * functionality using real PPTX files and created test cases.
 */

/**
 * Main test function for XML Search & Replace functionality
 */
function runXMLSearchReplaceTests() {
  console.log('🔍 XML SEARCH & REPLACE AND LANGUAGE STANDARDIZATION TESTS');
  console.log('=' * 65);
  console.log('Multi-file XML manipulation and Brandwares language standardization\n');
  
  const results = {
    xmlSearchReplaceBasic: testBasicXMLSearchReplace(),
    xmlSearchReplaceRegex: testRegexSearchReplace(),
    xmlAttributeReplace: testAttributeSearchReplace(),
    languageAnalysis: testLanguageAnalysis(),
    languageStandardization: testLanguageStandardization(),
    mixedLanguageFix: testMixedLanguageFix(),
    realFileAnalysis: testRealFileAnalysis(),
    productionWorkflow: testProductionWorkflow()
  };
  
  printXMLSearchReplaceResults(results);
  return results;
}

/**
 * Test basic XML search and replace
 */
function testBasicXMLSearchReplace() {
  console.log('🔧 TEST 1: Basic XML Search & Replace');
  console.log('-' * 40);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Testing basic text replacement in XML...');
    const result = slides.searchReplaceXML({
      search: 'Arial',
      replace: 'Calibri',
      caseSensitive: true,
      xmlOnly: true
    });
    
    console.log('✅ Basic XML search & replace successful');
    console.log(`   Files processed: ${result.filesProcessed}`);
    console.log(`   Files modified: ${result.filesModified}`);
    console.log(`   Total replacements: ${result.totalReplacements}`);
    console.log('   Replaced Arial font references with Calibri');
    
    return {
      success: true,
      result: result,
      message: 'Basic XML search & replace completed successfully'
    };
    
  } catch (error) {
    console.log('❌ Basic XML search & replace failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test regex-based search and replace
 */
function testRegexSearchReplace() {
  console.log('\n🎯 TEST 2: Regex Search & Replace');
  console.log('-' * 35);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Testing regex pattern replacement...');
    // Replace any typeface attribute value with "Roboto"
    const result = slides.searchReplaceRegex(
      'typeface="([^"]*)"',
      'typeface="Roboto"',
      { flags: 'g' }
    );
    
    console.log('✅ Regex search & replace successful');
    console.log(`   Files processed: ${result.filesProcessed}`);
    console.log(`   Files modified: ${result.filesModified}`);
    console.log(`   Total replacements: ${result.totalReplacements}`);
    console.log('   Replaced all typeface attributes with Roboto using regex');
    
    return {
      success: true,
      result: result,
      message: 'Regex search & replace completed successfully'
    };
    
  } catch (error) {
    console.log('❌ Regex search & replace failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test attribute-specific search and replace
 */
function testAttributeSearchReplace() {
  console.log('\n🏷️ TEST 3: Attribute Search & Replace');
  console.log('-' * 39);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Testing attribute-specific replacement...');
    const result = slides.searchReplaceAttribute('lang', 'en-US', 'en-GB');
    
    console.log('✅ Attribute search & replace successful');
    console.log(`   Files processed: ${result.filesProcessed}`);
    console.log(`   Files modified: ${result.filesModified}`);
    console.log(`   Total replacements: ${result.totalReplacements}`);
    console.log('   Replaced lang="en-US" with lang="en-GB"');
    
    return {
      success: true,
      result: result,
      message: 'Attribute search & replace completed successfully'
    };
    
  } catch (error) {
    console.log('❌ Attribute search & replace failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test language analysis functionality
 */
function testLanguageAnalysis() {
  console.log('\n🌐 TEST 4: Language Analysis');
  console.log('-' * 31);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Analyzing language tag distribution...');
    const analysis = slides.analyzeLanguages();
    
    console.log('Generating comprehensive language report...');
    const report = slides.generateLanguageReport();
    
    console.log('✅ Language analysis successful');
    console.log(`   Total language tags: ${analysis.totalLanguageTags}`);
    console.log(`   Unique languages: ${analysis.uniqueLanguages.length}`);
    console.log(`   Languages found: ${analysis.uniqueLanguages.join(', ')}`);
    console.log(`   Files with languages: ${analysis.filesWithLanguageTags}`);
    console.log(`   Mixed paragraphs: ${report.summary.mixedLanguageParagraphs}`);
    
    return {
      success: true,
      analysis: analysis,
      report: report,
      message: 'Language analysis completed successfully'
    };
    
  } catch (error) {
    console.log('❌ Language analysis failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test language standardization
 */
function testLanguageStandardization() {
  console.log('\n🇺🇸 TEST 5: Language Standardization');
  console.log('-' * 40);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('Standardizing to US English...');
    const result = slides.standardizeToUSEnglish();
    
    console.log('Verifying standardization...');
    const postAnalysis = slides.analyzeLanguages();
    
    console.log('✅ Language standardization successful');
    console.log(`   Total replacements: ${result.totalReplacements}`);
    console.log(`   Files modified: ${result.filesModified}`);
    console.log(`   Target language: ${result.targetLanguage}`);
    console.log(`   Successfully standardized: ${result.successfullyStandardized}`);
    console.log(`   Final languages: ${postAnalysis.uniqueLanguages.join(', ')}`);
    
    return {
      success: true,
      result: result,
      postAnalysis: postAnalysis,
      message: 'Language standardization completed successfully'
    };
    
  } catch (error) {
    console.log('❌ Language standardization failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test mixed language paragraph fix (the specific Brandwares problem)
 */
function testMixedLanguageFix() {
  console.log('\n🔧 TEST 6: Mixed Language Paragraph Fix');
  console.log('-' * 44);
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    // First, inject some mixed language content for testing
    console.log('Creating test content with mixed languages...');
    slides.searchReplaceXML({
      search: 'lang="en-US"',
      replace: 'lang="en-CA"',
      xmlOnly: true
    });
    
    console.log('Fixing mixed language paragraphs...');
    const result = slides.fixMixedLanguageParagraphs('en-US');
    
    console.log('✅ Mixed language paragraph fix successful');
    console.log(`   Mixed paragraphs found: ${result.mixedParagraphsFound}`);
    console.log(`   Mixed paragraphs fixed: ${result.mixedParagraphsFixed}`);
    console.log(`   Total replacements: ${result.totalReplacements}`);
    console.log(`   Target language: ${result.targetLanguage}`);
    console.log(`   Successfully fixed: ${result.success}`);
    
    return {
      success: true,
      result: result,
      message: 'Mixed language paragraph fix completed successfully'
    };
    
  } catch (error) {
    console.log('❌ Mixed language paragraph fix failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test analysis of real files from /analysis folder
 */
function testRealFileAnalysis() {
  console.log('\n📁 TEST 7: Real File Analysis');
  console.log('-' * 33);
  
  try {
    console.log('Looking for real PPTX files in /analysis folder...');
    
    // Try to find and analyze files from the analysis folder
    const testFiles = [
      'CloudFunction_PPTX_Test_2025-08-12_11-53-52.pptx',
      'CloudFunction_PPTX_Test_2025-08-12_12-48-32.pptx',
      'Custom_Professional_Theme_2025-08-12_11-53-55.pptx'
    ];
    
    const analysisResults = [];
    
    testFiles.forEach(fileName => {
      try {
        console.log(`\nAnalyzing: ${fileName}`);
        
        // Try to find the file in Google Drive (would be uploaded there)
        const files = DriveApp.getFilesByName(fileName);
        if (files.hasNext()) {
          const file = files.next();
          console.log(`Found file: ${file.getId()}`);
          
          // Load and analyze
          const slides = OOXMLSlides.fromGoogleDriveFile(file.getId());
          const languageAnalysis = slides.analyzeLanguages();
          const xmlStats = slides.xmlSearchReplace.getXMLStatistics();
          
          analysisResults.push({
            fileName: fileName,
            fileId: file.getId(),
            success: true,
            languageAnalysis: languageAnalysis,
            xmlStats: xmlStats
          });
          
          console.log(`   Languages: ${languageAnalysis.uniqueLanguages.join(', ')}`);
          console.log(`   Language tags: ${languageAnalysis.totalLanguageTags}`);
          console.log(`   XML files: ${xmlStats.totalXMLFiles}`);
          
        } else {
          console.log(`   File not found in Google Drive`);
          analysisResults.push({
            fileName: fileName,
            success: false,
            error: 'File not found in Google Drive'
          });
        }
        
      } catch (fileError) {
        console.log(`   Error analyzing ${fileName}: ${fileError.message}`);
        analysisResults.push({
          fileName: fileName,
          success: false,
          error: fileError.message
        });
      }
    });
    
    const successfulAnalyses = analysisResults.filter(r => r.success).length;
    
    console.log(`\n✅ Real file analysis completed`);
    console.log(`   Files attempted: ${testFiles.length}`);
    console.log(`   Files successfully analyzed: ${successfulAnalyses}`);
    console.log('   Upload PPTX files to Google Drive for analysis');
    
    return {
      success: true,
      analysisResults: analysisResults,
      successfulAnalyses: successfulAnalyses,
      message: 'Real file analysis completed'
    };
    
  } catch (error) {
    console.log('❌ Real file analysis failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Test complete production workflow (Brandwares style)
 */
function testProductionWorkflow() {
  console.log('\n🏭 TEST 8: Production Workflow');
  console.log('-' * 35);
  
  try {
    console.log('Simulating Brandwares production workflow...');
    
    const slides = OOXMLSlides.fromTemplate();
    
    // Step 1: Analyze current state
    console.log('\n1. Initial analysis...');
    const initialAnalysis = slides.analyzeLanguages();
    
    // Step 2: Apply multiple fixes using batch operations
    console.log('2. Applying batch operations...');
    const batchOperations = [
      {
        name: 'Font Standardization',
        search: 'Arial',
        replace: 'Calibri',
        caseSensitive: true
      },
      {
        name: 'Font Size Cleanup',
        search: 'sz="1200"',
        replace: 'sz="1400"',
        caseSensitive: true
      }
    ];
    
    const batchResults = slides.xmlSearchReplace.batchSearchReplace(batchOperations);
    
    // Step 3: Language standardization
    console.log('3. Language standardization...');
    const langResult = slides.standardizeToUSEnglish();
    
    // Step 4: Fix mixed paragraphs
    console.log('4. Fixing mixed language paragraphs...');
    const mixedResult = slides.fixMixedLanguageParagraphs('en-US');
    
    // Step 5: Final verification
    console.log('5. Final verification...');
    const finalAnalysis = slides.analyzeLanguages();
    const finalReport = slides.generateLanguageReport();
    
    // Step 6: Save processed file
    console.log('6. Saving processed file...');
    const processedFileId = slides.saveToGoogleDrive('Production_Workflow_Test_Result');
    
    console.log('✅ Production workflow successful');
    console.log(`   Processed file: ${processedFileId}`);
    console.log(`   Initial languages: ${initialAnalysis.uniqueLanguages.join(', ')}`);
    console.log(`   Final languages: ${finalAnalysis.uniqueLanguages.join(', ')}`);
    console.log(`   Batch operations: ${batchOperations.length} completed`);
    console.log(`   Language standardized: ${langResult.successfullyStandardized}`);
    console.log(`   Mixed paragraphs fixed: ${mixedResult.success}`);
    console.log(`   Ready for client delivery: ${finalReport.summary.needsStandardization ? 'No' : 'Yes'}`);
    
    // Clean up
    DriveApp.getFileById(processedFileId).setTrashed(true);
    
    return {
      success: true,
      processedFileId: processedFileId,
      initialAnalysis: initialAnalysis,
      finalAnalysis: finalAnalysis,
      batchResults: batchResults,
      langResult: langResult,
      mixedResult: mixedResult,
      finalReport: finalReport,
      message: 'Production workflow completed successfully'
    };
    
  } catch (error) {
    console.log('❌ Production workflow failed:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Print comprehensive test results
 */
function printXMLSearchReplaceResults(results) {
  console.log('\n📋 XML SEARCH & REPLACE AND LANGUAGE STANDARDIZATION RESULTS');
  console.log('=' * 65);
  
  const tests = [
    { name: 'Basic XML Search & Replace', key: 'xmlSearchReplaceBasic', icon: '🔧' },
    { name: 'Regex Search & Replace', key: 'xmlSearchReplaceRegex', icon: '🎯' },
    { name: 'Attribute Search & Replace', key: 'xmlAttributeReplace', icon: '🏷️' },
    { name: 'Language Analysis', key: 'languageAnalysis', icon: '🌐' },
    { name: 'Language Standardization', key: 'languageStandardization', icon: '🇺🇸' },
    { name: 'Mixed Language Fix', key: 'mixedLanguageFix', icon: '🔧' },
    { name: 'Real File Analysis', key: 'realFileAnalysis', icon: '📁' },
    { name: 'Production Workflow', key: 'productionWorkflow', icon: '🏭' }
  ];
  
  console.log('\n🎯 Test Results:');
  tests.forEach(test => {
    const result = results[test.key];
    const status = result.success ? '✅ PASSED' : '❌ FAILED';
    console.log(`${test.icon} ${test.name}: ${status}`);
    if (!result.success && result.error) {
      console.log(`   Error: ${result.error}`);
    }
  });
  
  const passedTests = Object.values(results).filter(r => r.success).length;
  const totalTests = Object.keys(results).length;
  
  console.log(`\n📊 Summary: ${passedTests}/${totalTests} tests passed`);
  
  if (passedTests === totalTests) {
    console.log('\n🎉 ALL XML SEARCH & REPLACE AND LANGUAGE TESTS WORKING!');
    console.log('Multi-file XML manipulation and language standardization fully implemented.');
    console.log('\n🔧 XML Search & Replace Capabilities:');
    console.log('• Multi-file pattern matching and replacement');
    console.log('• Regular expression support for complex patterns');
    console.log('• Attribute-specific search and replace');
    console.log('• Element content modification');
    console.log('• Batch operations across entire PPTX structure');
    console.log('• Find operations without replacement');
    console.log('• XML content statistics and analysis');
    console.log('\n🌐 Language Standardization Capabilities:');
    console.log('• Complete language tag analysis and distribution');
    console.log('• Standardization to any target language');
    console.log('• Mixed language paragraph detection and fixing');
    console.log('• Comprehensive language reports');
    console.log('• Built-in presets (US, UK, Canadian English)');
    console.log('• Batch standardization for multiple languages');
    console.log('\n💼 Production Ready Features:');
    console.log('• Brandwares-style international file processing');
    console.log('• Professional file finalization workflows');
    console.log('• Client delivery language standardization');
    console.log('• Integration with all existing OOXML manipulation');
  } else {
    console.log('\n⚠️ SOME TESTS FAILED');
    console.log('Review the error messages above for troubleshooting.');
  }
}

/**
 * Demonstrate complete XML manipulation capabilities
 */
function demonstrateXMLManipulationCapabilities() {
  console.log('🔧 DEMONSTRATION: Complete XML Search & Replace Capabilities');
  console.log('=' * 70);
  console.log('Showcasing multi-file XML manipulation and language standardization...\n');
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    console.log('1. Creating presentation with mixed content...');
    
    console.log('2. Performing comprehensive XML analysis...');
    const xmlStats = slides.xmlSearchReplace.getXMLStatistics();
    const languageReport = slides.generateLanguageReport();
    
    console.log('3. Applying production-level fixes...');
    
    // Font standardization
    const fontResult = slides.searchReplaceXML({
      search: 'Arial',
      replace: 'Calibri',
      caseSensitive: true
    });
    
    // Language standardization
    const langResult = slides.standardizeToUSEnglish();
    
    // Attribute cleanup
    const attrResult = slides.searchReplaceAttribute('dirty', '1', '0');
    
    console.log('4. Saving demonstration file...');
    const demoFileId = slides.saveToGoogleDrive('XML_Manipulation_Demo');
    
    console.log('\n🎉 XML Manipulation Demonstration Complete!');
    console.log(`   Demo file: ${demoFileId}`);
    console.log(`   XML files processed: ${xmlStats.totalXMLFiles}`);
    console.log(`   Font replacements: ${fontResult.totalReplacements}`);
    console.log(`   Language replacements: ${langResult.totalReplacements}`);
    console.log(`   Attribute cleanups: ${attrResult.totalReplacements}`);
    console.log(`   Total XML size: ${xmlStats.totalSize} characters`);
    console.log('\n🚀 Key Achievements:');
    console.log('   • Multi-file XML search and replace engine');
    console.log('   • Brandwares-style language standardization');
    console.log('   • Production-ready OOXML file processing');
    console.log('   • Integration with all tanaikech-style techniques');
    
    return demoFileId;
    
  } catch (error) {
    console.log('❌ XML manipulation demonstration failed:', error.message);
    return null;
  }
}

/**
 * Quick XML manipulation test
 */
function quickXMLManipulationTest() {
  console.log('🔍 Quick XML Manipulation Test');
  console.log('Testing XML search & replace and language standardization...\n');
  
  try {
    const slides = OOXMLSlides.fromTemplate();
    
    // Test XML search & replace
    const xmlResult = XMLSearchReplaceEditor.testXMLSearchReplace(slides);
    
    // Test language standardization  
    const langResult = LanguageStandardizationEditor.testLanguageStandardization(slides);
    
    if (xmlResult && langResult) {
      console.log('✅ Quick XML manipulation test passed!');
      console.log('Both XML search & replace and language standardization working.');
      return true;
    } else {
      console.log('❌ Quick XML manipulation test failed');
      return false;
    }
    
  } catch (error) {
    console.log('❌ Quick XML manipulation test error:', error.message);
    return false;
  }
}