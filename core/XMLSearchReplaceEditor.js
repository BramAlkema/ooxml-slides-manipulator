/**
 * XMLSearchReplaceEditor - Multi-file XML Search & Replace Engine
 * 
 * Implements comprehensive search and replace functionality across all OOXML files.
 * This is the core engine for advanced XML manipulation including:
 * - Multi-file pattern matching and replacement
 * - Regular expression support for complex patterns
 * - Attribute-specific search and replace
 * - Element content modification
 * - Namespace-aware XML processing
 * - Batch operations across entire PPTX structure
 * 
 * Based on Brandwares production techniques for OOXML file repair and standardization.
 * Essential for professional PowerPoint file processing and cleanup.
 */

class XMLSearchReplaceEditor {
  
  constructor(ooxml) {
    this.ooxml = ooxml;
    this.parser = ooxml.parser;
  }
  
  /**
   * Perform multi-file XML search and replace operation
   * @param {Object} searchReplaceConfig - Search and replace configuration
   * @returns {Object} Results of the operation
   */
  searchReplaceXML(searchReplaceConfig) {
    console.log('🔍 Multi-file XML Search & Replace Operation');
    console.log(`Search pattern: "${searchReplaceConfig.search}"`);
    console.log(`Replace pattern: "${searchReplaceConfig.replace}"`);
    
    try {
      this._ensureExtracted();
      
      const results = {
        filesProcessed: 0,
        filesModified: 0,
        totalReplacements: 0,
        modifiedFiles: [],
        errors: []
      };
      
      // Process all XML files in the PPTX
      this.parser.files.forEach((fileData, fileName) => {
        if (fileData.isXML) {
          try {
            const replacementCount = this._searchReplaceInFile(fileName, fileData, searchReplaceConfig);
            results.filesProcessed++;
            
            if (replacementCount > 0) {
              results.filesModified++;
              results.totalReplacements += replacementCount;
              results.modifiedFiles.push({
                fileName: fileName,
                replacements: replacementCount
              });
            }
          } catch (fileError) {
            results.errors.push({
              fileName: fileName,
              error: fileError.message
            });
          }
        }
      });
      
      console.log(`✅ Search & Replace complete:`);
      console.log(`   Files processed: ${results.filesProcessed}`);
      console.log(`   Files modified: ${results.filesModified}`);
      console.log(`   Total replacements: ${results.totalReplacements}`);
      
      if (results.modifiedFiles.length > 0) {
        console.log('   Modified files:');
        results.modifiedFiles.forEach(file => {
          console.log(`     • ${file.fileName}: ${file.replacements} replacements`);
        });
      }
      
      return results;
      
    } catch (error) {
      console.error('❌ Search & Replace failed:', error.message);
      throw error;
    }
  }
  
  /**
   * Search and replace with regular expressions
   * @param {string} searchPattern - Regular expression pattern
   * @param {string} replacePattern - Replacement pattern (can include $1, $2, etc.)
   * @param {Object} options - Additional options
   * @returns {Object} Results of the operation
   */
  searchReplaceRegex(searchPattern, replacePattern, options = {}) {
    console.log('🔍 Multi-file XML Regex Search & Replace');
    console.log(`Regex pattern: /${searchPattern}/${options.flags || 'g'}`);
    
    const searchReplaceConfig = {
      search: searchPattern,
      replace: replacePattern,
      useRegex: true,
      regexFlags: options.flags || 'g',
      caseSensitive: options.caseSensitive !== false,
      xmlOnly: options.xmlOnly !== false,
      filePattern: options.filePattern,
      excludeFiles: options.excludeFiles || []
    };
    
    return this.searchReplaceXML(searchReplaceConfig);
  }
  
  /**
   * Search and replace XML attributes specifically
   * @param {string} attributeName - Name of the attribute to search
   * @param {string} searchValue - Value to search for
   * @param {string} replaceValue - Value to replace with
   * @param {Object} options - Additional options
   * @returns {Object} Results of the operation
   */
  searchReplaceAttribute(attributeName, searchValue, replaceValue, options = {}) {
    console.log('🔍 XML Attribute Search & Replace');
    console.log(`Attribute: ${attributeName}`);
    console.log(`Search: "${searchValue}" → Replace: "${replaceValue}"`);
    
    const searchPattern = `${attributeName}="${searchValue}"`;
    const replacePattern = `${attributeName}="${replaceValue}"`;
    
    const searchReplaceConfig = {
      search: searchPattern,
      replace: replacePattern,
      useRegex: false,
      caseSensitive: options.caseSensitive !== false,
      xmlOnly: true,
      filePattern: options.filePattern,
      excludeFiles: options.excludeFiles || []
    };
    
    return this.searchReplaceXML(searchReplaceConfig);
  }
  
  /**
   * Search and replace XML element content
   * @param {string} elementName - Name of the XML element
   * @param {string} searchContent - Content to search for
   * @param {string} replaceContent - Content to replace with
   * @param {Object} options - Additional options
   * @returns {Object} Results of the operation
   */
  searchReplaceElementContent(elementName, searchContent, replaceContent, options = {}) {
    console.log('🔍 XML Element Content Search & Replace');
    console.log(`Element: <${elementName}>`);
    console.log(`Content: "${searchContent}" → "${replaceContent}"`);
    
    // Create regex pattern to match element content
    const searchPattern = `(<${elementName}[^>]*>)${this._escapeRegex(searchContent)}(<\\/${elementName}>)`;
    const replacePattern = `$1${replaceContent}$2`;
    
    const searchReplaceConfig = {
      search: searchPattern,
      replace: replacePattern,
      useRegex: true,
      regexFlags: 'g',
      caseSensitive: options.caseSensitive !== false,
      xmlOnly: true,
      filePattern: options.filePattern,
      excludeFiles: options.excludeFiles || []
    };
    
    return this.searchReplaceXML(searchReplaceConfig);
  }
  
  /**
   * Batch search and replace operations
   * @param {Array<Object>} operations - Array of search/replace operations
   * @returns {Array<Object>} Results for each operation
   */
  batchSearchReplace(operations) {
    console.log(`🔍 Batch Search & Replace: ${operations.length} operations`);
    
    const allResults = [];
    
    operations.forEach((operation, index) => {
      console.log(`\nOperation ${index + 1}:`);
      try {
        const result = this.searchReplaceXML(operation);
        result.operationIndex = index;
        result.operationName = operation.name || `Operation ${index + 1}`;
        allResults.push(result);
      } catch (error) {
        allResults.push({
          operationIndex: index,
          operationName: operation.name || `Operation ${index + 1}`,
          error: error.message,
          filesProcessed: 0,
          filesModified: 0,
          totalReplacements: 0
        });
      }
    });
    
    // Summary
    const totalReplacements = allResults.reduce((sum, result) => sum + (result.totalReplacements || 0), 0);
    const totalModifiedFiles = allResults.reduce((sum, result) => sum + (result.filesModified || 0), 0);
    
    console.log('\n📊 Batch Operation Summary:');
    console.log(`   Total operations: ${operations.length}`);
    console.log(`   Total replacements: ${totalReplacements}`);
    console.log(`   Total files modified: ${totalModifiedFiles}`);
    
    return allResults;
  }
  
  /**
   * Find all occurrences of a pattern without replacing
   * @param {string} searchPattern - Pattern to search for
   * @param {Object} options - Search options
   * @returns {Array<Object>} All matches found
   */
  findAll(searchPattern, options = {}) {
    console.log(`🔍 Find All: "${searchPattern}"`);
    
    try {
      this._ensureExtracted();
      
      const matches = [];
      
      this.parser.files.forEach((fileData, fileName) => {
        if (fileData.isXML) {
          const fileMatches = this._findInFile(fileName, fileData, searchPattern, options);
          matches.push(...fileMatches);
        }
      });
      
      console.log(`Found ${matches.length} matches across ${new Set(matches.map(m => m.fileName)).size} files`);
      
      return matches;
      
    } catch (error) {
      console.error('❌ Find operation failed:', error.message);
      throw error;
    }
  }
  
  /**
   * Get statistics about XML content
   * @returns {Object} XML content statistics
   */
  getXMLStatistics() {
    console.log('📊 Analyzing XML content statistics...');
    
    try {
      this._ensureExtracted();
      
      const stats = {
        totalXMLFiles: 0,
        totalSize: 0,
        fileTypes: {},
        commonAttributes: {},
        commonElements: {},
        namespaces: new Set()
      };
      
      this.parser.files.forEach((fileData, fileName) => {
        if (fileData.isXML) {
          stats.totalXMLFiles++;
          stats.totalSize += fileData.content.length;
          
          // File type analysis
          const fileType = this._getFileType(fileName);
          stats.fileTypes[fileType] = (stats.fileTypes[fileType] || 0) + 1;
          
          // Attribute and element analysis
          this._analyzeXMLContent(fileData.content, stats);
        }
      });
      
      stats.namespaces = Array.from(stats.namespaces);
      
      console.log(`✅ XML Statistics:`);
      console.log(`   Total XML files: ${stats.totalXMLFiles}`);
      console.log(`   Total size: ${stats.totalSize} characters`);
      console.log(`   File types: ${Object.keys(stats.fileTypes).join(', ')}`);
      console.log(`   Namespaces: ${stats.namespaces.length}`);
      
      return stats;
      
    } catch (error) {
      console.error('❌ Statistics analysis failed:', error.message);
      throw error;
    }
  }
  
  // Private helper methods
  
  _ensureExtracted() {
    if (!this.parser.isExtracted) {
      this.parser.extract();
    }
  }
  
  _searchReplaceInFile(fileName, fileData, config) {
    let content = fileData.content;
    let replacementCount = 0;
    
    // Apply file pattern filter if specified
    if (config.filePattern && !fileName.match(config.filePattern)) {
      return 0;
    }
    
    // Skip excluded files
    if (config.excludeFiles.includes(fileName)) {
      return 0;
    }
    
    if (config.useRegex) {
      const regex = new RegExp(config.search, config.regexFlags);
      const matches = content.match(regex);
      if (matches) {
        replacementCount = matches.length;
        content = content.replace(regex, config.replace);
      }
    } else {
      const searchFlags = config.caseSensitive ? 'g' : 'gi';
      const searchRegex = new RegExp(this._escapeRegex(config.search), searchFlags);
      const matches = content.match(searchRegex);
      if (matches) {
        replacementCount = matches.length;
        content = content.replace(searchRegex, config.replace);
      }
    }
    
    if (replacementCount > 0) {
      // Update the file content
      fileData.content = content;
      
      // Update any cached XML document if it exists
      if (fileData.xmlDoc) {
        delete fileData.xmlDoc;
      }
    }
    
    return replacementCount;
  }
  
  _findInFile(fileName, fileData, searchPattern, options) {
    const content = fileData.content;
    const matches = [];
    
    let regex;
    if (options.useRegex) {
      regex = new RegExp(searchPattern, options.regexFlags || 'g');
    } else {
      const flags = options.caseSensitive === false ? 'gi' : 'g';
      regex = new RegExp(this._escapeRegex(searchPattern), flags);
    }
    
    let match;
    while ((match = regex.exec(content)) !== null) {
      // Calculate line number
      const beforeMatch = content.substring(0, match.index);
      const lineNumber = (beforeMatch.match(/\n/g) || []).length + 1;
      
      // Get context around the match
      const contextStart = Math.max(0, match.index - 50);
      const contextEnd = Math.min(content.length, match.index + match[0].length + 50);
      const context = content.substring(contextStart, contextEnd);
      
      matches.push({
        fileName: fileName,
        lineNumber: lineNumber,
        index: match.index,
        match: match[0],
        context: context.trim(),
        groups: match.slice(1) // Capture groups if any
      });
    }
    
    return matches;
  }
  
  _escapeRegex(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  
  _getFileType(fileName) {
    if (fileName.includes('/slides/')) return 'slide';
    if (fileName.includes('/slideLayouts/')) return 'slideLayout';
    if (fileName.includes('/slideMasters/')) return 'slideMaster';
    if (fileName.includes('/theme/')) return 'theme';
    if (fileName.includes('presentation.xml')) return 'presentation';
    if (fileName.includes('[Content_Types]')) return 'contentTypes';
    if (fileName.includes('.rels')) return 'relationships';
    return 'other';
  }
  
  _analyzeXMLContent(content, stats) {
    // Find all attributes
    const attributeRegex = /(\w+)="([^"]*)"/g;
    let match;
    while ((match = attributeRegex.exec(content)) !== null) {
      const attrName = match[1];
      stats.commonAttributes[attrName] = (stats.commonAttributes[attrName] || 0) + 1;
    }
    
    // Find all elements
    const elementRegex = /<(\w+)[\s>]/g;
    while ((match = elementRegex.exec(content)) !== null) {
      const elementName = match[1];
      stats.commonElements[elementName] = (stats.commonElements[elementName] || 0) + 1;
    }
    
    // Find namespaces
    const namespaceRegex = /xmlns:(\w+)="([^"]*)"/g;
    while ((match = namespaceRegex.exec(content)) !== null) {
      stats.namespaces.add(`${match[1]}: ${match[2]}`);
    }
  }
  
  /**
   * Test multi-file XML search and replace functionality
   */
  static testXMLSearchReplace(ooxml) {
    console.log('🧪 Testing multi-file XML search & replace...');
    
    const editor = new XMLSearchReplaceEditor(ooxml);
    
    try {
      // Test 1: Simple text replacement
      console.log('\nTest 1: Simple text replacement');
      const result1 = editor.searchReplaceXML({
        search: 'Arial',
        replace: 'Calibri',
        caseSensitive: true,
        xmlOnly: true
      });
      
      // Test 2: Attribute replacement
      console.log('\nTest 2: Attribute replacement');
      const result2 = editor.searchReplaceAttribute('lang', 'en-US', 'en-GB');
      
      // Test 3: Regex replacement
      console.log('\nTest 3: Regex replacement');
      const result3 = editor.searchReplaceRegex(
        'typeface="([^"]*)"',
        'typeface="Roboto"',
        { flags: 'g' }
      );
      
      // Test 4: Find all occurrences
      console.log('\nTest 4: Find all font references');
      const matches = editor.findAll('typeface=');
      
      console.log('\n✅ XML Search & Replace tests completed!');
      console.log(`   Test 1 replacements: ${result1.totalReplacements}`);
      console.log(`   Test 2 replacements: ${result2.totalReplacements}`);
      console.log(`   Test 3 replacements: ${result3.totalReplacements}`);
      console.log(`   Find operation matches: ${matches.length}`);
      
      return true;
      
    } catch (error) {
      console.log('❌ XML Search & Replace test failed:', error.message);
      return false;
    }
  }
}

// Export for use
if (typeof module !== 'undefined' && module.exports) {
  module.exports = XMLSearchReplaceEditor;
}