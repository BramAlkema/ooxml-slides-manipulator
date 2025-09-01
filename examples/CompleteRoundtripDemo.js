/**
 * COMPLETE OOXML ROUNDTRIP DEMONSTRATION
 * Google Slides → PPTX Download → OOXML Manipulation → Enhanced PPTX
 * Shows before/after with real table formatting and typography changes
 */

function completeOOXMLRoundtripDemo() {
  console.log('🔄 === COMPLETE OOXML ROUNDTRIP DEMO === 🔄');
  console.log('Google Slides → PPTX → OOXML Manipulation → Enhanced PPTX');
  
  try {
    const results = {
      timestamp: new Date().toISOString(),
      workflow: {},
      beforeAfter: {},
      files: {}
    };
    
    // STEP 1: Create baseline presentation with basic content
    console.log('\n📝 STEP 1: Creating baseline presentation...');
    results.workflow.step1 = createBaselinePresentation();
    
    // STEP 2: Simulate PPTX download from Google Slides
    console.log('\n📥 STEP 2: Creating PPTX (simulating Google Slides export)...');
    results.workflow.step2 = createBasicPPTX();
    
    // STEP 3: Apply OOXML enhancements
    console.log('\n🎨 STEP 3: Applying OOXML enhancements...');
    results.workflow.step3 = applyOOXMLEnhancements(results.workflow.step2.presentation);
    
    // STEP 4: Document before/after changes
    console.log('\n📊 STEP 4: Documenting before/after changes...');
    results.beforeAfter = documentChanges(results.workflow.step2, results.workflow.step3);
    
    // STEP 5: Export both versions for comparison
    console.log('\n💾 STEP 5: Exporting comparison files...');
    results.files = exportComparisonFiles(results.workflow.step2, results.workflow.step3);
    
    console.log('\n✅ === ROUNDTRIP DEMO COMPLETE === ✅');
    console.log('Results:', JSON.stringify(results, null, 2));
    
    return results;
    
  } catch (error) {
    console.error('💥 Roundtrip demo failed:', error);
    return {
      success: false,
      error: error.toString(),
      stack: error.stack
    };
  }
}

function createBaselinePresentation() {
  console.log('Creating baseline presentation with basic table and text...');
  
  const ooxml = new OOXMLSlides();
  const baseline = ooxml.createBlankPresentation()
    .addSlide()
    .addTitle('Quarterly Financial Report')
    .addContent('Basic financial data for Q4 2024 performance review');
  
  // Add table slide with basic formatting
  baseline.addSlide()
    .addTitle('Revenue Analysis')
    .addContent('Department | Q3 Revenue | Q4 Revenue | Growth Rate\nSales | $125,000 | $142,000 | 13.6%\nMarketing | $89,000 | $95,000 | 6.7%\nProduct | $156,000 | $178,000 | 14.1%\nSupport | $67,000 | $73,000 | 9.0%');
  
  // Add typography slide
  baseline.addSlide()
    .addTitle('Typography Sample')
    .addContent('This slide demonstrates basic typography.\nStandard fonts: Arial, Calibri\nBasic formatting: Bold, Italic, Underline\nNo advanced kerning or custom spacing');
  
  return {
    success: true,
    slides: 3,
    content: 'Basic table with standard formatting, basic typography',
    styling: 'Default Google Slides theme, standard fonts'
  };
}

function createBasicPPTX() {
  console.log('Creating basic PPTX (simulating Google Slides download)...');
  
  const ooxml = new OOXMLSlides();
  const basic = ooxml.createBlankPresentation()
    .addSlide()
    .addTitle('Quarterly Financial Report')
    .addContent('Basic financial data for Q4 2024 performance review');
  
  // Create table with basic styling
  basic.addSlide()
    .addTitle('Revenue Analysis')
    .addContent('Department | Q3 Revenue | Q4 Revenue | Growth Rate\nSales | $125,000 | $142,000 | 13.6%\nMarketing | $89,000 | $95,000 | 6.7%\nProduct | $156,000 | $178,000 | 14.1%\nSupport | $67,000 | $73,000 | 9.0%');
  
  // Apply basic table styling (before enhancement)
  basic.applyTableStyles({
    headerStyle: {
      backgroundColor: '#d9d9d9',
      textColor: '#000000',
      fontWeight: 'normal'
    },
    dataStyle: {
      alternatingRows: false,
      primaryColor: '#ffffff',
      secondaryColor: '#ffffff',
      borderColor: '#000000'
    }
  });
  
  // Add typography slide with basic fonts
  basic.addSlide()
    .addTitle('Typography Sample')
    .addContent('This slide demonstrates basic typography.\nStandard fonts: Arial, Calibri\nBasic formatting: Bold, Italic, Underline\nNo advanced kerning or custom spacing');
  
  // Export baseline PPTX
  const basicBlob = basic.exportToPPTX();
  const basicFile = DriveApp.createFile(basicBlob)
    .setName(`BEFORE_Basic_PPTX_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss')}.pptx`);
  
  return {
    success: true,
    presentation: basic,
    file: basicFile.getUrl(),
    fileId: basicFile.getId(),
    styling: {
      tableHeaders: 'Light gray background, black text',
      tableData: 'White background, no alternating rows',
      typography: 'Standard Arial/Calibri, no kerning',
      colors: 'Default theme colors'
    }
  };
}

function applyOOXMLEnhancements(basicPresentation) {
  console.log('Applying advanced OOXML enhancements...');
  
  // Create enhanced version from the basic presentation
  const enhanced = basicPresentation;
  
  // ENHANCEMENT 1: Professional table styling
  console.log('  → Applying professional table styling...');
  const tableResult = enhanced.applyTableStyles({
    headerStyle: {
      backgroundColor: '#1f4e79',    // Professional dark blue
      textColor: '#ffffff',          // White text
      fontWeight: 'bold',
      fontSize: '12pt',
      fontFamily: 'Calibri'
    },
    dataStyle: {
      alternatingRows: true,
      primaryColor: '#f8f9fa',       // Light gray
      secondaryColor: '#ffffff',     // White
      borderColor: '#1f4e79',        // Matching blue borders
      borderWidth: '1pt'
    },
    numberFormatting: {
      currency: true,
      percentage: true,
      thousandsSeparator: true
    }
  });
  
  // ENHANCEMENT 2: Professional color scheme
  console.log('  → Applying professional color scheme...');
  const colorResult = enhanced.applyCustomColors({
    primaryColor: '#1f4e79',     // Corporate blue
    secondaryColor: '#70ad47',   // Success green  
    accentColor: '#ff6b35'       // Alert orange
  });
  
  // ENHANCEMENT 3: Advanced typography with kerning
  console.log('  → Applying advanced typography and kerning...');
  const typographyResult = enhanced.applyProfessionalTypography('corporate', {
    customKerningPairs: [
      { pair: 'AV', adjustment: -0.1 },
      { pair: 'WA', adjustment: -0.05 },
      { pair: 'To', adjustment: -0.08 },
      { pair: 'Yo', adjustment: -0.06 }
    ],
    openTypeFeatures: ['liga', 'kern'],
    characterTracking: 'tight',
    fonts: {
      headingFont: 'Calibri Bold',
      bodyFont: 'Calibri',
      tableFont: 'Calibri'
    }
  });
  
  // ENHANCEMENT 4: XML-based content improvements
  console.log('  → Applying XML-based content enhancements...');
  const xmlResult = enhanced.searchReplaceXML({
    operations: [
      // Format currency properly
      {
        searchPattern: '\\$([0-9]+),([0-9]+)',
        replaceWith: '$$$1,$2.00',
        options: { useRegex: true }
      },
      // Format percentages
      {
        searchPattern: '([0-9]+\\.?[0-9]*)%',
        replaceWith: '$1%',
        options: { useRegex: true }
      },
      // Enhance department names
      {
        searchPattern: 'Sales',
        replaceWith: 'Sales Department',
        options: { caseSensitive: true }
      },
      {
        searchPattern: 'Marketing',
        replaceWith: 'Marketing Division',
        options: { caseSensitive: true }
      },
      {
        searchPattern: 'Product',
        replaceWith: 'Product Development',
        options: { caseSensitive: true }
      },
      {
        searchPattern: 'Support',
        replaceWith: 'Customer Support',
        options: { caseSensitive: true }
      }
    ],
    scope: 'slides'
  });
  
  // ENHANCEMENT 5: Language standardization
  console.log('  → Standardizing language tags...');
  const languageResult = enhanced.standardizeLanguage('en-US', {
    includeSlides: true,
    includeMasters: true,
    includeLayouts: true,
    preserveSpecialElements: true
  });
  
  // Export enhanced PPTX
  const enhancedBlob = enhanced.exportToPPTX();
  const enhancedFile = DriveApp.createFile(enhancedBlob)
    .setName(`AFTER_Enhanced_PPTX_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss')}.pptx`);
  
  return {
    success: true,
    presentation: enhanced,
    file: enhancedFile.getUrl(),
    fileId: enhancedFile.getId(),
    enhancements: {
      tableStyles: tableResult,
      customColors: colorResult,
      typography: typographyResult,
      xmlOperations: xmlResult,
      languageStandardization: languageResult
    },
    styling: {
      tableHeaders: 'Corporate blue background (#1f4e79), white bold text',
      tableData: 'Alternating gray/white rows with blue borders',
      typography: 'Corporate Calibri with custom kerning pairs (AV, WA, To, Yo)',
      colors: 'Professional 3-color scheme (blue, green, orange)',
      xmlEnhancements: '6 content improvements applied',
      languageStandardization: 'All tags normalized to en-US'
    }
  };
}

function documentChanges(beforeData, afterData) {
  console.log('Documenting before/after changes...');
  
  const changes = {
    tableFormatting: {
      before: beforeData.styling,
      after: afterData.styling,
      improvements: [
        'Header background: Light gray → Professional corporate blue (#1f4e79)',
        'Header text: Black → White bold with improved contrast',
        'Data rows: Plain white → Alternating gray/white for readability',
        'Borders: Basic black → Professional blue matching theme',
        'Department names: Enhanced with descriptive titles'
      ]
    },
    typography: {
      before: 'Standard Arial/Calibri, no kerning adjustments',
      after: 'Corporate Calibri with custom kerning pairs: AV(-0.1), WA(-0.05), To(-0.08), Yo(-0.06)',
      improvements: [
        'Applied professional kerning pairs for improved readability',
        'Enabled OpenType ligatures and kerning features',
        'Tight character tracking for executive presentation style',
        'Consistent font hierarchy: Calibri Bold (headings), Calibri (body)'
      ]
    },
    colorScheme: {
      before: 'Default theme colors',
      after: 'Professional 3-color corporate scheme',
      improvements: [
        'Primary: #1f4e79 (Corporate blue for headers and accents)',
        'Secondary: #70ad47 (Success green for positive metrics)',
        'Accent: #ff6b35 (Alert orange for attention-drawing elements)',
        'Applied consistently across all presentation elements'
      ]
    },
    contentEnhancements: {
      xmlOperations: 6,
      improvements: [
        'Currency formatting: $125,000 → $125,000.00 (consistent decimal places)',
        'Department names: Sales → Sales Department (more descriptive)',
        'Enhanced all department titles for professional clarity',
        'Standardized percentage formatting',
        'Language tags normalized to en-US throughout document'
      ]
    }
  };
  
  return changes;
}

function exportComparisonFiles(beforeData, afterData) {
  console.log('Exporting comparison files...');
  
  // Create summary comparison document
  const ooxml = new OOXMLSlides();
  const comparison = ooxml.createBlankPresentation()
    .addSlide()
    .addTitle('🔄 OOXML Enhancement Comparison')
    .addContent('Before and After demonstration of advanced OOXML manipulation');
  
  // Before slide
  comparison.addSlide()
    .addTitle('📋 BEFORE: Basic Google Slides Export')
    .addContent('• Light gray table headers with black text\n' +
               '• Plain white data rows, no alternating colors\n' +
               '• Standard Arial/Calibri fonts, no kerning\n' +
               '• Default theme colors\n' +
               '• Basic department names (Sales, Marketing, Product)\n' +
               '• Standard currency formatting');
  
  // After slide
  comparison.addSlide()
    .addTitle('✨ AFTER: Enhanced OOXML Processing')
    .addContent('• Corporate blue headers (#1f4e79) with white bold text\n' +
               '• Alternating gray/white rows with blue borders\n' +
               '• Corporate Calibri with custom kerning (AV, WA, To, Yo)\n' +
               '• Professional 3-color scheme (blue, green, orange)\n' +
               '• Enhanced department names (Sales Department, Marketing Division)\n' +
               '• Improved currency formatting ($125,000.00)');
  
  // Technical details slide
  comparison.addSlide()
    .addTitle('🔧 Technical Enhancements Applied')
    .addContent('• 6 XML search & replace operations\n' +
               '• 4 custom kerning pairs for typography\n' +
               '• 3-color professional theme application\n' +
               '• Language tag standardization (en-US)\n' +
               '• OpenType ligature and kerning features\n' +
               '• Advanced table styling with alternating rows');
  
  // Apply the enhanced styling to the comparison document itself
  comparison
    .applyCustomColors({
      primaryColor: '#1f4e79',
      secondaryColor: '#70ad47',
      accentColor: '#ff6b35'
    })
    .applyProfessionalTypography('corporate');
  
  const comparisonBlob = comparison.exportToPPTX();
  const comparisonFile = DriveApp.createFile(comparisonBlob)
    .setName(`COMPARISON_BeforeAfter_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss')}.pptx`);
  
  return {
    beforeFile: beforeData.file,
    afterFile: afterData.file,
    comparisonFile: comparisonFile.getUrl(),
    summary: 'Three files created: BEFORE (basic), AFTER (enhanced), COMPARISON (side-by-side analysis)'
  };
}

// Main execution function
function runCompleteRoundtripDemo() {
  return completeOOXMLRoundtripDemo();
}