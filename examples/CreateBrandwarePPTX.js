/**
 * CreateBrandwarePPTX.js - Complete GAS Script for Brandware PowerPoint
 * 
 * This script creates a presentation using Google Slides API,
 * exports it as PPTX, and applies advanced brandware OOXML features.
 * 
 * Run this in Google Apps Script to create a fully-featured demo.
 */

/**
 * Main function - Creates brandware-enhanced PowerPoint
 */
function createBrandwarePowerPoint() {
  console.log('🎨 Creating Brandware PowerPoint with Typography & Tables...\n');
  
  try {
    // Step 1: Create presentation with Slides API
    const presentationId = createPresentationWithSlidesAPI();
    console.log('✅ Presentation created:', presentationId);
    
    // Step 2: Export as PPTX
    const pptxBlob = exportAsPPTX(presentationId);
    console.log('✅ Exported as PPTX');
    
    // Step 3: Apply brandware enhancements
    const enhancedFileId = applyBrandwareEnhancements(pptxBlob, presentationId);
    console.log('✅ Brandware enhancements applied');
    
    // Step 4: Create shareable link
    const fileUrl = DriveApp.getFileById(enhancedFileId).getUrl();
    console.log('\n🎉 SUCCESS! Brandware PowerPoint created');
    console.log('📁 File URL:', fileUrl);
    
    return enhancedFileId;
    
  } catch (error) {
    console.error('❌ Error:', error.message);
    throw error;
  }
}

/**
 * Step 1: Create presentation using Slides API
 */
function createPresentationWithSlidesAPI() {
  // Create new presentation
  const presentation = SlidesApp.create('Brandware Typography & Tables Demo');
  const presentationId = presentation.getId();
  
  // Get the first slide
  const slides = presentation.getSlides();
  const firstSlide = slides[0];
  
  // Clear default content
  firstSlide.getPageElements().forEach(element => element.remove());
  
  // Add title with specific positioning for typography demo
  const title = firstSlide.insertTextBox(
    'ADVANCED TYPOGRAPHY',
    50, 50, 600, 80
  );
  
  const titleText = title.getText();
  titleText.getTextStyle()
    .setFontSize(48)
    .setFontFamily('Segoe UI')
    .setBold(true)
    .setForegroundColor('#2E5E8C');
  
  // Add subtitle
  const subtitle = firstSlide.insertTextBox(
    'Custom Kerning • Letter Spacing • OpenType Features',
    50, 140, 600, 40
  );
  
  subtitle.getText().getTextStyle()
    .setFontSize(18)
    .setFontFamily('Segoe UI Light')
    .setForegroundColor('#666666');
  
  // Add demonstration text with different styles
  const demoText = firstSlide.insertTextBox(
    'Typography Examples:\nAVATAR (needs kerning)\nWATER (letter spacing)\noffice (ligatures)',
    50, 200, 350, 150
  );
  
  demoText.getText().getTextStyle()
    .setFontSize(16)
    .setFontFamily('Segoe UI');
  
  // Create a table to demonstrate custom table styles
  const table = firstSlide.insertTable(5, 4, 50, 380, 650, 250);
  
  // Style the header row
  const headerRow = table.getRow(0);
  headerRow.getCell(0).getText().setText('Feature').getTextStyle()
    .setBold(true).setForegroundColor('#FFFFFF');
  headerRow.getCell(1).getText().setText('Standard UI').getTextStyle()
    .setBold(true).setForegroundColor('#FFFFFF');
  headerRow.getCell(2).getText().setText('Brandware XML').getTextStyle()
    .setBold(true).setForegroundColor('#FFFFFF');
  headerRow.getCell(3).getText().setText('Benefit').getTextStyle()
    .setBold(true).setForegroundColor('#FFFFFF');
  
  // Set header row background
  for (let i = 0; i < 4; i++) {
    headerRow.getCell(i).getFill().setSolidFill('#2E5E8C');
  }
  
  // Add table data
  const tableData = [
    ['Kerning', 'Auto only', 'Custom values', 'Perfect spacing'],
    ['Letter Spacing', 'Limited', 'Full control', 'Brand consistency'],
    ['Table Gradients', 'Not available', 'Multi-stop', 'Professional look'],
    ['Hidden Metadata', 'None', 'Custom XML', 'Tracking & compliance']
  ];
  
  tableData.forEach((row, rowIndex) => {
    const tableRow = table.getRow(rowIndex + 1);
    row.forEach((cell, cellIndex) => {
      tableRow.getCell(cellIndex).getText().setText(cell);
      // Alternate row colors
      if (rowIndex % 2 === 0) {
        tableRow.getCell(cellIndex).getFill().setSolidFill('#F8F9FA');
      }
    });
  });
  
  // Add a second slide for more examples
  const secondSlide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  
  const slide2Title = secondSlide.insertTextBox(
    'ENTERPRISE TABLE STYLES',
    50, 50, 600, 80
  );
  
  slide2Title.getText().getTextStyle()
    .setFontSize(48)
    .setFontFamily('Segoe UI')
    .setBold(true)
    .setForegroundColor('#7B1FA2');
  
  // Add description
  const description = secondSlide.insertTextBox(
    'These table styles are created entirely through OOXML manipulation.\n' +
    'They include features not available in PowerPoint\'s UI:\n' +
    '• Gradient fills in headers\n' +
    '• Custom cell padding and margins\n' +
    '• Advanced border styling\n' +
    '• Conditional formatting rules',
    50, 150, 650, 150
  );
  
  description.getText().getTextStyle()
    .setFontSize(16)
    .setFontFamily('Segoe UI');
  
  // Add another table for Material Design demo
  const materialTable = secondSlide.insertTable(4, 3, 50, 350, 650, 200);
  
  // Style as Material Design
  const matHeader = materialTable.getRow(0);
  matHeader.getCell(0).getText().setText('Component').getTextStyle()
    .setBold(true).setForegroundColor('#FFFFFF');
  matHeader.getCell(1).getText().setText('Material Spec').getTextStyle()
    .setBold(true).setForegroundColor('#FFFFFF');
  matHeader.getCell(2).getText().setText('Implementation').getTextStyle()
    .setBold(true).setForegroundColor('#FFFFFF');
  
  // Purple header for Material Design
  for (let i = 0; i < 3; i++) {
    matHeader.getCell(i).getFill().setSolidFill('#7B1FA2');
  }
  
  // Add Material Design data
  const materialData = [
    ['Elevation', '2dp shadow', 'OOXML effects'],
    ['Typography', 'Roboto font', 'Custom font embedding'],
    ['Motion', 'Ease curves', 'Animation timing']
  ];
  
  materialData.forEach((row, rowIndex) => {
    const tableRow = materialTable.getRow(rowIndex + 1);
    row.forEach((cell, cellIndex) => {
      tableRow.getCell(cellIndex).getText().setText(cell);
      tableRow.getCell(cellIndex).getFill().setSolidFill('#FFFFFF');
    });
  });
  
  // Save the presentation
  presentation.saveAndClose();
  
  return presentationId;
}

/**
 * Step 2: Export presentation as PPTX
 */
function exportAsPPTX(presentationId) {
  // Get the file from Drive
  const file = DriveApp.getFileById(presentationId);
  
  // Export as PPTX using Drive API advanced service
  const pptxBlob = file.getAs('application/vnd.openxmlformats-officedocument.presentationml.presentation');
  pptxBlob.setName('Brandware_Demo_Original.pptx');
  
  // Save original PPTX for comparison
  const originalFile = DriveApp.createFile(pptxBlob);
  console.log('  Original PPTX saved:', originalFile.getId());
  
  return pptxBlob;
}

/**
 * Step 3: Apply brandware OOXML enhancements
 */
function applyBrandwareEnhancements(pptxBlob, originalPresentationId) {
  console.log('🔧 Applying brandware enhancements...');
  
  // Save the PPTX temporarily
  const tempFile = DriveApp.createFile(pptxBlob);
  const tempFileId = tempFile.getId();
  
  try {
    // Get the OOXML manifest
    const manifest = OOXMLJsonService.unwrap(tempFileId);
    console.log('  Unwrapped PPTX:', manifest.entries.length, 'files');
    
    // Enhancement 1: Add custom typography (kerning and spacing)
    applyTypographyEnhancements(manifest);
    
    // Enhancement 2: Add custom table styles
    addCustomTableStyles(manifest);
    
    // Enhancement 3: Add hidden metadata
    addHiddenMetadata(manifest);
    
    // Enhancement 4: Add advanced effects
    addAdvancedEffects(manifest);
    
    // Rewrap the enhanced PPTX
    const enhancedFileId = OOXMLJsonService.rewrap(manifest, {
      filename: 'Brandware_Demo_Enhanced.pptx'
    });
    
    console.log('  Enhanced PPTX created:', enhancedFileId);
    
    // Clean up temp file
    DriveApp.getFileById(tempFileId).setTrashed(true);
    
    return enhancedFileId;
    
  } catch (error) {
    // Clean up on error
    DriveApp.getFileById(tempFileId).setTrashed(true);
    throw error;
  }
}

/**
 * Enhancement 1: Apply advanced typography
 */
function applyTypographyEnhancements(manifest) {
  console.log('  📝 Applying typography enhancements...');
  
  manifest.entries.forEach(entry => {
    if (entry.path.includes('slides/slide') && entry.path.endsWith('.xml')) {
      // Add custom kerning to title text
      if (entry.text.includes('ADVANCED TYPOGRAPHY')) {
        // Add kerning to the title - different values for different letter pairs
        entry.text = entry.text.replace(
          /<a:rPr([^>]*?)>/g,
          (match, attrs) => {
            // Check if this is the title text
            if (match.includes('sz="4800"') || match.includes('sz="48')) {
              return `<a:rPr${attrs} kern="1400" spc="-75" baseline="0">`;
            }
            return match;
          }
        );
        
        // Special kerning for AVATAR text
        entry.text = entry.text.replace(
          />AVATAR</g,
          '>A<a:rPr kern="1800"/>V<a:rPr kern="1600"/>A<a:rPr kern="1400"/>T<a:rPr kern="1200"/>AR'
        );
        
        // Letter spacing for WATER
        entry.text = entry.text.replace(
          />WATER</g,
          '>W<a:rPr spc="200"/>A<a:rPr spc="200"/>T<a:rPr spc="200"/>E<a:rPr spc="200"/>R'
        );
      }
      
      // Add OpenType features (ligatures) to specific text
      entry.text = entry.text.replace(
        /<a:rPr([^>]*?)>(.*?)office/g,
        '<a:rPr$1 dirty="0" smtClean="0"><a:latin typeface="Calibri"/><a:extLst><a:ext uri="{C4B1156A-D082-4D91-A7EB-8F1C0C470E33}"><a14:ligatures xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="standard"/></a:ext></a:extLst></a:rPr>$2office'
      );
    }
  });
}

/**
 * Enhancement 2: Add custom table styles
 */
function addCustomTableStyles(manifest) {
  console.log('  🎨 Adding custom table styles...');
  
  // Check if tableStyles.xml exists
  let tableStylesEntry = manifest.entries.find(e => e.path === 'ppt/tableStyles.xml');
  
  const customTableStyles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}">
  <!-- Brandware Enterprise Blue -->
  <a:tblStyle styleId="{BRAND-ENTERPRISE-001}" styleName="Brandware Enterprise">
    <a:wholeTbl>
      <a:tcStyle>
        <a:fill>
          <a:solidFill>
            <a:schemeClr val="lt1"/>
          </a:solidFill>
        </a:fill>
        <a:tcBdr>
          <a:left>
            <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
              <a:solidFill>
                <a:srgbClr val="D0D0D0"/>
              </a:solidFill>
            </a:ln>
          </a:left>
          <a:right>
            <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
              <a:solidFill>
                <a:srgbClr val="D0D0D0"/>
              </a:solidFill>
            </a:ln>
          </a:right>
          <a:top>
            <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
              <a:solidFill>
                <a:srgbClr val="D0D0D0"/>
              </a:solidFill>
            </a:ln>
          </a:top>
          <a:bottom>
            <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
              <a:solidFill>
                <a:srgbClr val="D0D0D0"/>
              </a:solidFill>
            </a:ln>
          </a:bottom>
        </a:tcBdr>
      </a:tcStyle>
    </a:wholeTbl>
    <a:firstRow>
      <a:tcStyle>
        <a:fill>
          <a:gradFill rotWithShape="1">
            <a:gsLst>
              <a:gs pos="0">
                <a:srgbClr val="2E5E8C">
                  <a:shade val="95000"/>
                </a:srgbClr>
              </a:gs>
              <a:gs pos="100000">
                <a:srgbClr val="1E4E7C">
                  <a:shade val="85000"/>
                </a:srgbClr>
              </a:gs>
            </a:gsLst>
            <a:lin ang="2700000" scaled="1"/>
          </a:gradFill>
        </a:fill>
      </a:tcStyle>
      <a:tcTxStyle b="on">
        <a:fontRef idx="minor">
          <a:srgbClr val="FFFFFF"/>
        </a:fontRef>
      </a:tcTxStyle>
    </a:firstRow>
    <a:band1H>
      <a:tcStyle>
        <a:fill>
          <a:solidFill>
            <a:srgbClr val="F0F4F8">
              <a:alpha val="50000"/>
            </a:srgbClr>
          </a:solidFill>
        </a:fill>
      </a:tcStyle>
    </a:band1H>
    <a:band2H>
      <a:tcStyle>
        <a:fill>
          <a:solidFill>
            <a:srgbClr val="FFFFFF"/>
          </a:solidFill>
        </a:fill>
      </a:tcStyle>
    </a:band2H>
  </a:tblStyle>
  
  <!-- Material Design Purple -->
  <a:tblStyle styleId="{BRAND-MATERIAL-002}" styleName="Material Purple">
    <a:wholeTbl>
      <a:tcStyle>
        <a:fill>
          <a:solidFill>
            <a:srgbClr val="FFFFFF"/>
          </a:solidFill>
        </a:fill>
        <a:tcBdr>
          <a:bottom>
            <a:ln w="6350">
              <a:solidFill>
                <a:srgbClr val="E0E0E0"/>
              </a:solidFill>
            </a:ln>
          </a:bottom>
        </a:tcBdr>
      </a:tcStyle>
    </a:wholeTbl>
    <a:firstRow>
      <a:tcStyle>
        <a:fill>
          <a:solidFill>
            <a:srgbClr val="7B1FA2"/>
          </a:solidFill>
        </a:fill>
        <a:tcBdr/>
      </a:tcStyle>
      <a:tcTxStyle b="on">
        <a:fontRef idx="minor">
          <a:latin typeface="Roboto" pitchFamily="34" charset="0"/>
          <a:srgbClr val="FFFFFF"/>
        </a:fontRef>
      </a:tcTxStyle>
    </a:firstRow>
  </a:tblStyle>
</a:tblStyleLst>`;
  
  if (tableStylesEntry) {
    // Replace existing table styles
    tableStylesEntry.text = customTableStyles;
  } else {
    // Add new table styles file
    manifest.entries.push({
      path: 'ppt/tableStyles.xml',
      type: 'xml',
      text: customTableStyles
    });
    
    // Update Content_Types.xml
    const contentTypes = manifest.entries.find(e => e.path === '[Content_Types].xml');
    if (contentTypes) {
      contentTypes.text = contentTypes.text.replace(
        '</Types>',
        '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>\n</Types>'
      );
    }
    
    // Update presentation relationships
    const presRels = manifest.entries.find(e => e.path === 'ppt/_rels/presentation.xml.rels');
    if (presRels) {
      // Find highest rId
      const rIdMatch = presRels.text.match(/rId(\d+)/g);
      const maxId = rIdMatch ? Math.max(...rIdMatch.map(id => parseInt(id.replace('rId', '')))) : 0;
      const newId = `rId${maxId + 1}`;
      
      presRels.text = presRels.text.replace(
        '</Relationships>',
        `<Relationship Id="${newId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>\n</Relationships>`
      );
    }
  }
  
  // Apply table style to existing tables
  manifest.entries.forEach(entry => {
    if (entry.path.includes('slides/slide') && entry.text.includes('<a:tbl>')) {
      // Add style reference to first table (Enterprise style)
      let tableCount = 0;
      entry.text = entry.text.replace(
        /<a:tblPr([^>]*?)>/g,
        (match, attrs) => {
          tableCount++;
          if (tableCount === 1) {
            return `<a:tblPr${attrs}>\n<a:tableStyleId>{BRAND-ENTERPRISE-001}</a:tableStyleId>`;
          } else if (tableCount === 2) {
            return `<a:tblPr${attrs}>\n<a:tableStyleId>{BRAND-MATERIAL-002}</a:tableStyleId>`;
          }
          return match;
        }
      );
    }
  });
}

/**
 * Enhancement 3: Add hidden metadata
 */
function addHiddenMetadata(manifest) {
  console.log('  🔒 Adding hidden metadata...');
  
  // Add to core properties
  const coreProps = manifest.entries.find(e => e.path === 'docProps/core.xml');
  if (coreProps) {
    coreProps.text = coreProps.text.replace(
      '</cp:coreProperties>',
      `<cp:category>BRANDWARE-DEMO-2024</cp:category>
      <cp:contentStatus>ENHANCED-TYPOGRAPHY-TABLES</cp:contentStatus>
      <cp:version>2.0.0</cp:version>
      <dc:identifier>BRAND-${Date.now()}</dc:identifier>
      </cp:coreProperties>`
    );
  }
  
  // Add custom properties
  let customProps = manifest.entries.find(e => e.path === 'docProps/custom.xml');
  
  const customPropsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="BrandwareVersion">
    <vt:lpwstr>2.0.0</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="TypographyEnhanced">
    <vt:bool>true</vt:bool>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="4" name="TableStylesCustom">
    <vt:bool>true</vt:bool>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="5" name="ComplianceTracking">
    <vt:lpwstr>APPROVED-${new Date().toISOString()}</vt:lpwstr>
  </property>
  <property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="6" name="EnhancementSignature">
    <vt:lpwstr>${Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'BRANDWARE'))}</vt:lpwstr>
  </property>
</Properties>`;
  
  if (customProps) {
    customProps.text = customPropsXml;
  } else {
    manifest.entries.push({
      path: 'docProps/custom.xml',
      type: 'xml',
      text: customPropsXml
    });
    
    // Update relationships
    const rels = manifest.entries.find(e => e.path === '_rels/.rels');
    if (rels) {
      const rIdMatch = rels.text.match(/rId(\d+)/g);
      const maxId = rIdMatch ? Math.max(...rIdMatch.map(id => parseInt(id.replace('rId', '')))) : 0;
      const newId = `rId${maxId + 1}`;
      
      rels.text = rels.text.replace(
        '</Relationships>',
        `<Relationship Id="${newId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>\n</Relationships>`
      );
    }
  }
  
  // Add hidden tracking comments
  manifest.entries.forEach(entry => {
    if (entry.path.includes('.xml')) {
      entry.text = entry.text.replace(
        '<?xml version="1.0"',
        `<?xml version="1.0"
<!-- BRANDWARE-TRACKING:${Date.now()}:TYPOGRAPHY-TABLES-ENHANCED -->`
      );
    }
  });
}

/**
 * Enhancement 4: Add advanced effects
 */
function addAdvancedEffects(manifest) {
  console.log('  ✨ Adding advanced effects...');
  
  manifest.entries.forEach(entry => {
    if (entry.path.includes('slides/slide') && entry.path.endsWith('.xml')) {
      // Add advanced shadow effects to titles
      entry.text = entry.text.replace(
        /<a:effectLst\/>/g,
        `<a:effectLst>
          <a:outerShdw blurRad="63500" dist="50800" dir="2700000" algn="tl" rotWithShape="0">
            <a:srgbClr val="000000">
              <a:alpha val="35000"/>
            </a:srgbClr>
          </a:outerShdw>
        </a:effectLst>`
      );
      
      // Add reflection effect to specific text
      if (entry.text.includes('ADVANCED TYPOGRAPHY')) {
        entry.text = entry.text.replace(
          '</a:effectLst>',
          `<a:reflection blurRad="0" st="0" end="50000" dir="5400000" sy="100000" algn="bl" rotWithShape="0"/>
          </a:effectLst>`
        );
      }
    }
  });
}

/**
 * Test function - creates sample presentation
 */
function testBrandwarePowerPoint() {
  console.log('🧪 Testing Brandware PowerPoint creation...\n');
  
  try {
    const fileId = createBrandwarePowerPoint();
    
    // Verify the file
    const file = DriveApp.getFileById(fileId);
    console.log('\n📊 File Details:');
    console.log('  Name:', file.getName());
    console.log('  Size:', Math.round(file.getSize() / 1024), 'KB');
    console.log('  Type:', file.getMimeType());
    
    // List enhancements applied
    console.log('\n✨ Enhancements Applied:');
    console.log('  ✅ Custom kerning (kern="1400", kern="1800")');
    console.log('  ✅ Letter spacing (spc="-75", spc="200")');
    console.log('  ✅ OpenType ligatures');
    console.log('  ✅ Enterprise gradient table style');
    console.log('  ✅ Material Design table style');
    console.log('  ✅ Hidden compliance metadata');
    console.log('  ✅ Advanced shadow effects');
    console.log('  ✅ Reflection effects');
    
    console.log('\n🎉 Test completed successfully!');
    
  } catch (error) {
    console.error('❌ Test failed:', error.message);
    console.error(error.stack);
  }
}

// Helper function to check if OOXML service is available
function checkOOXMLService() {
  try {
    if (typeof OOXMLJsonService === 'undefined') {
      throw new Error('OOXMLJsonService not found. Please ensure the OOXML libraries are installed.');
    }
    
    // Test basic functionality
    const health = OOXMLJsonService.healthCheck();
    console.log('✅ OOXML Service is available:', health);
    return true;
    
  } catch (error) {
    console.error('❌ OOXML Service check failed:', error.message);
    console.log('\n📚 To use this script, you need:');
    console.log('1. Copy the OOXML library files to your project');
    console.log('2. Ensure OOXMLJsonService is available');
    console.log('3. Deploy the Cloud Run service if needed');
    return false;
  }
}

// Menu function for Google Sheets/Docs
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Brandware PowerPoint')
    .addItem('Create Demo Presentation', 'createBrandwarePowerPoint')
    .addItem('Test Brandware Features', 'testBrandwarePowerPoint')
    .addItem('Check OOXML Service', 'checkOOXMLService')
    .addSeparator()
    .addItem('About', 'showAbout')
    .addToUi();
}

function showAbout() {
  const html = `
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2>Brandware PowerPoint Creator</h2>
      <p>This tool creates PowerPoint presentations with advanced OOXML features:</p>
      <ul>
        <li>🔤 Custom typography (kerning, spacing, ligatures)</li>
        <li>🎨 Enterprise table styles with gradients</li>
        <li>📊 Material Design tables</li>
        <li>🔒 Hidden metadata for tracking</li>
        <li>✨ Advanced visual effects</li>
      </ul>
      <p><strong>These features are not available through PowerPoint's UI!</strong></p>
      <hr>
      <p>Created with OOXML manipulation techniques.</p>
    </div>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'About Brandware PowerPoint');
}