/**
 * ACID TEST FRAMEWORK
 * Color schemes × Font pairs × Table styles = Template quality test
 * Start with 2×2×2 = 8 combinations
 */

function runAcidTest() {
  ConsoleFormatter.header('🧪 ACID TEST FRAMEWORK');
  ConsoleFormatter.info('Testing template across color schemes, font pairs, and table styles');
  
  try {
    // Define test parameters (start small)
    const colorSchemes = [
      {
        name: 'Corporate Blue',
        primary: '#1f4e79',
        secondary: '#70ad47', 
        accent: '#ff6b35'
      },
      {
        name: 'Startup Orange',
        primary: '#ff6b35',
        secondary: '#1f4e79',
        accent: '#70ad47'
      }
    ];
    
    const fontPairs = [
      {
        name: 'Corporate',
        heading: 'Calibri',
        body: 'Calibri',
        headingSize: 24,
        bodySize: 14,
        headingBold: true
      },
      {
        name: 'Modern',
        heading: 'Arial',
        body: 'Arial',
        headingSize: 26,
        bodySize: 13,
        headingBold: true
      }
    ];
    
    const tableStyles = [
      {
        name: 'Executive',
        headerBg: 'primary', // Use primary color
        headerText: '#ffffff',
        alternatingRows: true,
        borderStyle: 'subtle'
      },
      {
        name: 'Financial',
        headerBg: '#f8f9fa',
        headerText: 'primary', // Use primary color for text
        alternatingRows: false,
        borderStyle: 'grid'
      }
    ];
    
    ConsoleFormatter.info(`Testing ${colorSchemes.length} × ${fontPairs.length} × ${tableStyles.length} = ${colorSchemes.length * fontPairs.length * tableStyles.length} combinations`, 'ACID TEST');
    
    const results = [];
    let combinationIndex = 1;
    
    // Test all combinations
    for (const colorScheme of colorSchemes) {
      for (const fontPair of fontPairs) {
        for (const tableStyle of tableStyles) {
          ConsoleFormatter.step(combinationIndex, `${colorScheme.name} + ${fontPair.name} + ${tableStyle.name}`);
          
          const testResult = createTestPresentation(colorScheme, fontPair, tableStyle, combinationIndex);
          results.push({
            combination: combinationIndex,
            colorScheme: colorScheme.name,
            fontPair: fontPair.name,
            tableStyle: tableStyle.name,
            ...testResult
          });
          
          combinationIndex++;
        }
      }
    }
    
    ConsoleFormatter.header('ACID TEST COMPLETE');
    ConsoleFormatter.section('📊 Results Summary');
    results.forEach(result => {
      const status = result.success ? 'PASS' : 'FAIL';
      ConsoleFormatter.status(status, `Combination ${result.combination}`, `${result.colorScheme}/${result.fontPair}/${result.tableStyle}`);
    });
    
    return {
      success: true,
      totalCombinations: results.length,
      passedTests: results.filter(r => r.success).length,
      failedTests: results.filter(r => !r.success).length,
      results: results,
      message: 'Acid test framework executed successfully'
    };
    
  } catch (error) {
    ConsoleFormatter.error('Acid test failed', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

function createTestPresentation(colorScheme, fontPair, tableStyle, combinationIndex) {
  ConsoleFormatter.indent(1, `→ Creating presentation with combination ${combinationIndex}...`);
  
  try {
    // Create test presentation
    const presentationTitle = `AcidTest_${combinationIndex}_${colorScheme.name}_${fontPair.name}_${tableStyle.name}`;
    const presentation = SlidesApp.create(presentationTitle);
    
    // Get first slide
    const slides = presentation.getSlides();
    const titleSlide = slides[0];
    
    // SLIDE 1: Title slide (tests font pairs + color scheme)
    const titlePlaceholder = titleSlide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
    if (titlePlaceholder) {
      const titleText = titlePlaceholder.asShape().getText();
      titleText.setText('Brand Template Acid Test');
      
      // Apply font and color
      titleText.getTextStyle()
        .setFontFamily(fontPair.heading)
        .setFontSize(fontPair.headingSize)
        .setBold(fontPair.headingBold)
        .setForegroundColor(colorScheme.primary);
    }
    
    const subtitlePlaceholder = titleSlide.getPlaceholder(SlidesApp.PlaceholderType.SUBTITLE);
    if (subtitlePlaceholder) {
      const subtitleText = subtitlePlaceholder.asShape().getText();
      subtitleText.setText(`${colorScheme.name} • ${fontPair.name} • ${tableStyle.name}`);
      
      subtitleText.getTextStyle()
        .setFontFamily(fontPair.body)
        .setFontSize(fontPair.bodySize)
        .setBold(false)
        .setForegroundColor(colorScheme.secondary);
    }
    
    // SLIDE 2: Table slide (tests table styles + color integration)
    const tableSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.TITLE_AND_BODY);
    
    const tableTitle = tableSlide.getPlaceholder(SlidesApp.PlaceholderType.TITLE);
    if (tableTitle) {
      const tableTitleText = tableTitle.asShape().getText();
      tableTitleText.setText('Revenue Analysis');
      
      tableTitleText.getTextStyle()
        .setFontFamily(fontPair.heading)
        .setFontSize(fontPair.headingSize)
        .setBold(fontPair.headingBold)
        .setForegroundColor(colorScheme.primary);
    }
    
    // Create table content
    const tableBody = tableSlide.getPlaceholder(SlidesApp.PlaceholderType.BODY);
    if (tableBody) {
      const tableText = tableBody.asShape().getText();
      tableText.setText(
        'Department | Q3 Revenue | Q4 Revenue | Growth\n' +
        'Sales | $125,000 | $142,000 | +13.6%\n' +
        'Marketing | $89,000 | $95,000 | +6.7%\n' +
        'Product | $156,000 | $178,000 | +14.1%'
      );
      
      tableText.getTextStyle()
        .setFontFamily(fontPair.body)
        .setFontSize(fontPair.bodySize);
    }
    
    // Apply advanced styling using our OOXML extensions
    if (typeof OOXMLSlides !== 'undefined') {
      try {
        const enhanced = OOXMLSlides.fromGoogleSlides(presentation.getId());
        
        // Apply color scheme
        enhanced.applyCustomColors({
          primaryColor: colorScheme.primary,
          secondaryColor: colorScheme.secondary,
          accentColor: colorScheme.accent
        });
        
        // Apply table styling
        const headerBgColor = tableStyle.headerBg === 'primary' ? colorScheme.primary : tableStyle.headerBg;
        const headerTextColor = tableStyle.headerText === 'primary' ? colorScheme.primary : tableStyle.headerText;
        
        enhanced.applyTableStyles({
          headerStyle: {
            backgroundColor: headerBgColor,
            textColor: headerTextColor,
            fontFamily: fontPair.heading,
            fontWeight: 'bold'
          },
          dataStyle: {
            alternatingRows: tableStyle.alternatingRows,
            primaryColor: '#ffffff',
            secondaryColor: '#f8f9fa',
            borderColor: colorScheme.primary,
            fontFamily: fontPair.body
          }
        });
        
        ConsoleFormatter.success('Enhanced with OOXML styling');
      } catch (ooxmlError) {
        ConsoleFormatter.warning('OOXML enhancement failed', ooxmlError.message);
      }
    }
    
    ConsoleFormatter.success(`Presentation created: ${presentation.getUrl()}`);
    
    return {
      success: true,
      presentationId: presentation.getId(),
      presentationUrl: presentation.getUrl(),
      presentationName: presentation.getName()
    };
    
  } catch (error) {
    ConsoleFormatter.error(`Failed to create combination ${combinationIndex}`, error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// Export function for easy execution
function runTemplateAcidTest() {
  return runAcidTest();
}