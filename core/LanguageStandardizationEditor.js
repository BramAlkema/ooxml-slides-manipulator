/**
 * LanguageStandardizationEditor - PowerPoint Language Tag Standardization
 * 
 * Implements the Brandwares technique for standardizing language tags across OOXML files.
 * PowerPoint puts language tags everywhere, often with mixed languages in a single paragraph.
 * This creates inconsistent spell-checking, formatting, and localization behavior.
 * 
 * Example problem:
 * <a:p>
 *   <a:r>
 *     <a:rPr lang="en-US" dirty="0">...</a:rPr>
 *     <a:t>Test</a:t>
 *   </a:r>
 *   <a:endParaRPr lang="en-CA" dirty="0">...</a:endParaRPr>
 * </a:p>
 * 
 * Solution: Replace all language tags with target language (e.g., lang="en-US")
 * 
 * Based on Brandwares production workflow for international PowerPoint file processing.
 * Essential for file finalization and client delivery in specific target languages.
 */

class LanguageStandardizationEditor {
  
  constructor(ooxml) {
    this.ooxml = ooxml;
    this.parser = ooxml.parser;
    this.xmlSearchReplace = new XMLSearchReplaceEditor(ooxml);
  }
  
  /**
   * Standardize all language tags to a target language
   * @param {string} targetLanguage - Target language code (e.g., 'en-US', 'fr-FR', 'de-DE')
   * @param {Object} options - Standardization options
   * @returns {Object} Results of the standardization
   */
  standardizeLanguage(targetLanguage, options = {}) {
    console.log(`🌐 Standardizing language tags to: ${targetLanguage}`);
    
    try {
      // First, analyze current language distribution
      const languageAnalysis = this.analyzeLanguages();
      console.log(`Found ${languageAnalysis.totalLanguageTags} language tags in ${languageAnalysis.uniqueLanguages.length} languages`);
      
      if (languageAnalysis.uniqueLanguages.length > 0) {
        console.log(`   Languages found: ${languageAnalysis.uniqueLanguages.join(', ')}`);
      }
      
      // Perform standardization using regex to catch all language tag patterns
      const searchPattern = 'lang="([^"]*)"';
      const replacePattern = `lang="${targetLanguage}"`;
      
      const result = this.xmlSearchReplace.searchReplaceRegex(
        searchPattern,
        replacePattern,
        {
          flags: 'g',
          caseSensitive: true,
          xmlOnly: true,
          filePattern: options.filePattern,
          excludeFiles: options.excludeFiles || []
        }
      );
      
      // Verify standardization
      const postAnalysis = this.analyzeLanguages();
      
      console.log(`✅ Language standardization complete:`);
      console.log(`   Total replacements: ${result.totalReplacements}`);
      console.log(`   Files modified: ${result.filesModified}`);
      console.log(`   Languages after: ${postAnalysis.uniqueLanguages.join(', ')}`);
      
      return {
        ...result,
        targetLanguage: targetLanguage,
        beforeAnalysis: languageAnalysis,
        afterAnalysis: postAnalysis,
        successfullyStandardized: postAnalysis.uniqueLanguages.length === 1 && postAnalysis.uniqueLanguages[0] === targetLanguage
      };
      
    } catch (error) {
      console.error('❌ Language standardization failed:', error.message);
      throw error;
    }
  }
  
  /**
   * Analyze current language tag distribution
   * @returns {Object} Language analysis results
   */
  analyzeLanguages() {
    console.log('🔍 Analyzing language tag distribution...');
    
    try {
      const languageMatches = this.xmlSearchReplace.findAll('lang="([^"]*)"', {
        useRegex: true,
        regexFlags: 'g'
      });
      
      const languageStats = {};
      const uniqueLanguages = new Set();
      
      languageMatches.forEach(match => {
        // Extract language from the regex capture group
        const fullMatch = match.match;
        const langMatch = fullMatch.match(/lang="([^"]*)"/);
        if (langMatch) {
          const language = langMatch[1];
          uniqueLanguages.add(language);
          
          languageStats[language] = languageStats[language] || {
            count: 0,
            files: new Set()
          };
          languageStats[language].count++;
          languageStats[language].files.add(match.fileName);
        }
      });
      
      // Convert sets to arrays for serialization
      const processedStats = {};
      Object.entries(languageStats).forEach(([lang, stats]) => {
        processedStats[lang] = {
          count: stats.count,
          files: Array.from(stats.files),
          fileCount: stats.files.size
        };
      });
      
      const analysis = {
        totalLanguageTags: languageMatches.length,
        uniqueLanguages: Array.from(uniqueLanguages).sort(),
        languageStats: processedStats,
        filesWithLanguageTags: new Set(languageMatches.map(m => m.fileName)).size
      };
      
      console.log(`Found ${analysis.totalLanguageTags} language tags`);
      if (analysis.uniqueLanguages.length > 0) {
        console.log(`Languages: ${analysis.uniqueLanguages.join(', ')}`);
      }
      
      return analysis;
      
    } catch (error) {
      console.error('❌ Language analysis failed:', error.message);
      throw error;
    }
  }
  
  /**
   * Standardize to US English (most common business requirement)
   * @returns {Object} Results of standardization
   */
  standardizeToUSEnglish() {
    console.log('🇺🇸 Standardizing to US English (en-US)...');
    return this.standardizeLanguage('en-US');
  }
  
  /**
   * Standardize to UK English
   * @returns {Object} Results of standardization
   */
  standardizeToUKEnglish() {
    console.log('🇬🇧 Standardizing to UK English (en-GB)...');
    return this.standardizeLanguage('en-GB');
  }
  
  /**
   * Standardize to Canadian English
   * @returns {Object} Results of standardization
   */
  standardizeToCanadianEnglish() {
    console.log('🇨🇦 Standardizing to Canadian English (en-CA)...');
    return this.standardizeLanguage('en-CA');
  }
  
  /**
   * Batch standardize multiple presentations to different target languages
   * @param {Array<Object>} languageTargets - Array of {targetLanguage, description}
   * @returns {Array<Object>} Results for each target language
   */
  batchLanguageStandardization(languageTargets) {
    console.log(`🌍 Batch language standardization: ${languageTargets.length} target languages`);
    
    const results = [];
    
    languageTargets.forEach((target, index) => {
      console.log(`\n${index + 1}. ${target.description || target.targetLanguage}:`);
      try {
        const result = this.standardizeLanguage(target.targetLanguage);
        result.targetDescription = target.description;
        results.push(result);
      } catch (error) {
        results.push({
          targetLanguage: target.targetLanguage,
          targetDescription: target.description,
          error: error.message,
          successfullyStandardized: false
        });
      }
    });
    
    console.log('\n📊 Batch Standardization Summary:');
    const successful = results.filter(r => r.successfullyStandardized).length;
    console.log(`   Successful: ${successful}/${languageTargets.length}`);
    
    return results;
  }
  
  /**
   * Fix mixed language paragraphs (common PowerPoint issue)
   * This addresses the specific problem shown in the example where one paragraph
   * contains multiple language tags
   * @param {string} targetLanguage - Target language for consistency
   * @returns {Object} Results of the fix
   */
  fixMixedLanguageParagraphs(targetLanguage) {
    console.log(`🔧 Fixing mixed language paragraphs with target: ${targetLanguage}`);
    
    try {
      // Find paragraphs with multiple language tags
      const mixedParagraphs = this._findMixedLanguageParagraphs();
      
      console.log(`Found ${mixedParagraphs.length} paragraphs with mixed languages`);
      
      if (mixedParagraphs.length === 0) {
        return {
          mixedParagraphsFound: 0,
          mixedParagraphsFixed: 0,
          message: 'No mixed language paragraphs found'
        };
      }
      
      // Standardize all language tags
      const standardizationResult = this.standardizeLanguage(targetLanguage);
      
      // Verify fix
      const postFixMixedParagraphs = this._findMixedLanguageParagraphs();
      
      console.log(`✅ Mixed language paragraph fix complete:`);
      console.log(`   Paragraphs with mixed languages before: ${mixedParagraphs.length}`);
      console.log(`   Paragraphs with mixed languages after: ${postFixMixedParagraphs.length}`);
      console.log(`   Total language tag replacements: ${standardizationResult.totalReplacements}`);
      
      return {
        mixedParagraphsFound: mixedParagraphs.length,
        mixedParagraphsFixed: mixedParagraphs.length - postFixMixedParagraphs.length,
        totalReplacements: standardizationResult.totalReplacements,
        targetLanguage: targetLanguage,
        success: postFixMixedParagraphs.length === 0
      };
      
    } catch (error) {
      console.error('❌ Failed to fix mixed language paragraphs:', error.message);
      throw error;
    }
  }
  
  /**
   * Generate language standardization report
   * @returns {Object} Comprehensive language report
   */
  generateLanguageReport() {
    console.log('📋 Generating comprehensive language report...');
    
    try {
      const analysis = this.analyzeLanguages();
      const mixedParagraphs = this._findMixedLanguageParagraphs();
      const fileStats = this._analyzeLanguagesByFileType();
      
      const report = {
        summary: {
          totalLanguageTags: analysis.totalLanguageTags,
          uniqueLanguages: analysis.uniqueLanguages.length,
          filesAffected: analysis.filesWithLanguageTags,
          mixedLanguageParagraphs: mixedParagraphs.length,
          needsStandardization: analysis.uniqueLanguages.length > 1
        },
        languageDistribution: analysis.languageStats,
        fileTypeAnalysis: fileStats,
        mixedParagraphDetails: mixedParagraphs,
        recommendations: this._generateRecommendations(analysis, mixedParagraphs)
      };
      
      console.log('✅ Language report generated');
      console.log(`   Languages found: ${analysis.uniqueLanguages.join(', ')}`);
      console.log(`   Mixed paragraphs: ${mixedParagraphs.length}`);
      console.log(`   Standardization needed: ${report.summary.needsStandardization ? 'Yes' : 'No'}`);
      
      return report;
      
    } catch (error) {
      console.error('❌ Failed to generate language report:', error.message);
      throw error;
    }
  }
  
  // Private helper methods
  
  _findMixedLanguageParagraphs() {
    // Find paragraphs (<a:p>...</a:p>) that contain multiple different lang attributes
    const paragraphMatches = this.xmlSearchReplace.findAll(
      '<a:p>(.*?)</a:p>',
      {
        useRegex: true,
        regexFlags: 'gs' // dot matches newlines
      }
    );
    
    const mixedParagraphs = [];
    
    paragraphMatches.forEach(match => {
      const paragraphContent = match.groups[0] || '';
      const languageMatches = paragraphContent.match(/lang="([^"]*)"/g);
      
      if (languageMatches && languageMatches.length > 1) {
        const languages = [...new Set(languageMatches)];
        if (languages.length > 1) {
          mixedParagraphs.push({
            fileName: match.fileName,
            lineNumber: match.lineNumber,
            languages: languages,
            context: match.context
          });
        }
      }
    });
    
    return mixedParagraphs;
  }
  
  _analyzeLanguagesByFileType() {
    const fileStats = {};
    
    this.parser.files.forEach((fileData, fileName) => {
      if (fileData.isXML) {
        const fileType = this._getFileType(fileName);
        
        if (!fileStats[fileType]) {
          fileStats[fileType] = {
            totalFiles: 0,
            filesWithLanguages: 0,
            totalLanguageTags: 0,
            languages: new Set()
          };
        }
        
        fileStats[fileType].totalFiles++;
        
        const languageMatches = fileData.content.match(/lang="([^"]*)"/g);
        if (languageMatches) {
          fileStats[fileType].filesWithLanguages++;
          fileStats[fileType].totalLanguageTags += languageMatches.length;
          
          languageMatches.forEach(match => {
            const lang = match.match(/lang="([^"]*)"/)[1];
            fileStats[fileType].languages.add(lang);
          });
        }
      }
    });
    
    // Convert sets to arrays
    Object.values(fileStats).forEach(stats => {
      stats.languages = Array.from(stats.languages);
    });
    
    return fileStats;
  }
  
  _getFileType(fileName) {
    if (fileName.includes('/slides/')) return 'slides';
    if (fileName.includes('/slideLayouts/')) return 'slideLayouts';
    if (fileName.includes('/slideMasters/')) return 'slideMasters';
    if (fileName.includes('/theme/')) return 'theme';
    if (fileName.includes('presentation.xml')) return 'presentation';
    return 'other';
  }
  
  _generateRecommendations(analysis, mixedParagraphs) {
    const recommendations = [];
    
    if (analysis.uniqueLanguages.length > 1) {
      recommendations.push({
        type: 'standardization',
        priority: 'high',
        message: `Standardize ${analysis.uniqueLanguages.length} different languages to single target language`,
        action: 'Run standardizeLanguage() with target language'
      });
    }
    
    if (mixedParagraphs.length > 0) {
      recommendations.push({
        type: 'mixed_paragraphs',
        priority: 'medium',
        message: `Fix ${mixedParagraphs.length} paragraphs with mixed language tags`,
        action: 'Run fixMixedLanguageParagraphs() to ensure consistency'
      });
    }
    
    if (analysis.uniqueLanguages.includes('en-US') && analysis.uniqueLanguages.length > 1) {
      recommendations.push({
        type: 'us_english',
        priority: 'low',
        message: 'Consider standardizing to en-US for business presentations',
        action: 'Run standardizeToUSEnglish() for common business standard'
      });
    }
    
    if (recommendations.length === 0) {
      recommendations.push({
        type: 'good',
        priority: 'info',
        message: 'Language tags are already consistent',
        action: 'No action needed'
      });
    }
    
    return recommendations;
  }
  
  /**
   * Test language standardization functionality
   */
  static testLanguageStandardization(ooxml) {
    console.log('🧪 Testing language standardization...');
    
    const editor = new LanguageStandardizationEditor(ooxml);
    
    try {
      // Test 1: Analyze current languages
      console.log('\nTest 1: Language analysis');
      const analysis = editor.analyzeLanguages();
      
      // Test 2: Generate comprehensive report
      console.log('\nTest 2: Language report generation');
      const report = editor.generateLanguageReport();
      
      // Test 3: Standardize to US English
      console.log('\nTest 3: Standardization to US English');
      const standardizationResult = editor.standardizeToUSEnglish();
      
      // Test 4: Verify standardization
      console.log('\nTest 4: Post-standardization verification');
      const postAnalysis = editor.analyzeLanguages();
      
      console.log('\n✅ Language standardization tests completed!');
      console.log(`   Original languages: ${analysis.uniqueLanguages.length}`);
      console.log(`   Replacements made: ${standardizationResult.totalReplacements}`);
      console.log(`   Final languages: ${postAnalysis.uniqueLanguages.length}`);
      console.log(`   Successfully standardized: ${standardizationResult.successfullyStandardized}`);
      
      return standardizationResult.successfullyStandardized;
      
    } catch (error) {
      console.log('❌ Language standardization test failed:', error.message);
      return false;
    }
  }
}

// Export for use
if (typeof module !== 'undefined' && module.exports) {
  module.exports = LanguageStandardizationEditor;
}