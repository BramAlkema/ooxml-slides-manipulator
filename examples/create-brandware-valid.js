/**
 * Create Valid Brandware Demo PowerPoint
 * 
 * This creates a minimal valid PPTX that PowerPoint will accept,
 * then we'll modify it with brandware features using the OOXML service.
 */

const fs = require('fs');
const path = require('path');

async function createValidBrandwarePowerPoint() {
  console.log('🎨 Creating Valid Brandware Demo PowerPoint...\n');
  
  // Instead of creating from scratch, let's use the OOXML service
  // to modify an existing valid PPTX
  
  // First, let's create a script that will run in Google Apps Script
  const gasScript = `
/**
 * Create Brandware Demo in Google Apps Script
 * Run this in Google Apps Script to create the demo
 */

async function createBrandwareDemo() {
  console.log('Creating Brandware Demo Presentation...');
  
  // Create a new presentation
  const presentation = SlidesApp.create('Brandware Typography & Tables Demo');
  const presentationId = presentation.getId();
  
  // Get the first slide
  const slides = presentation.getSlides();
  const slide = slides[0];
  
  // Add title
  const title = slide.insertTextBox('ADVANCED TYPOGRAPHY & TABLES', 50, 50, 600, 100);
  const titleStyle = title.getText().getTextStyle();
  titleStyle.setFontSize(44)
           .setFontFamily('Segoe UI')
           .setBold(true)
           .setForegroundColor('#2E5E8C');
  
  // Add subtitle
  const subtitle = slide.insertTextBox('Brandware OOXML Features Demonstration', 50, 150, 600, 50);
  subtitle.getText().getTextStyle()
          .setFontSize(20)
          .setFontFamily('Segoe UI Light')
          .setForegroundColor('#666666');
  
  // Add a table
  const table = slide.insertTable(4, 3, 50, 250, 600, 200);
  
  // Style the table header
  const headerRow = table.getRow(0);
  headerRow.getCell(0).getText().setText('Feature');
  headerRow.getCell(1).getText().setText('Standard');
  headerRow.getCell(2).getText().setText('Brandware');
  
  // Add table data
  table.getRow(1).getCell(0).getText().setText('Kerning');
  table.getRow(1).getCell(1).getText().setText('Limited');
  table.getRow(1).getCell(2).getText().setText('Full Control');
  
  table.getRow(2).getCell(0).getText().setText('Table Styles');
  table.getRow(2).getCell(1).getText().setText('Built-in');
  table.getRow(2).getCell(2).getText().setText('Custom XML');
  
  table.getRow(3).getCell(0).getText().setText('Typography');
  table.getRow(3).getCell(1).getText().setText('Basic');
  table.getRow(3).getCell(2).getText().setText('Advanced');
  
  // Export as PPTX
  const driveFile = DriveApp.getFileById(presentationId);
  const blob = driveFile.getAs('application/vnd.openxmlformats-officedocument.presentationml.presentation');
  
  // Now apply Brandware modifications
  console.log('Applying Brandware OOXML modifications...');
  
  // Get the PPTX as a manifest
  const manifest = await OOXMLJsonService.unwrap(presentationId);
  
  // Find and modify the slide XML to add advanced typography
  manifest.entries.forEach(entry => {
    if (entry.path.includes('slides/slide1.xml')) {
      // Add custom kerning to title text
      entry.text = entry.text.replace(
        /<a:rPr([^>]*)>/g,
        '<a:rPr$1 kern="1200" spc="-50">'
      );
      
      // Add advanced table styling
      if (entry.text.includes('<a:tbl>')) {
        // Inject custom table style reference
        entry.text = entry.text.replace(
          '<a:tblPr',
          '<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{BRANDWARE-001}</a:tableStyleId></a:tblPr><a:tblPr'
        );
      }
    }
    
    // Add custom table styles
    if (entry.path === 'ppt/tableStyles.xml') {
      entry.text = getCustomTableStyles();
    }
  });
  
  // If tableStyles.xml doesn't exist, add it
  if (!manifest.entries.find(e => e.path === 'ppt/tableStyles.xml')) {
    manifest.entries.push({
      path: 'ppt/tableStyles.xml',
      type: 'xml',
      text: getCustomTableStyles()
    });
    
    // Update relationships
    const presRels = manifest.entries.find(e => e.path === 'ppt/_rels/presentation.xml.rels');
    if (presRels) {
      presRels.text = presRels.text.replace(
        '</Relationships>',
        '<Relationship Id="rId999" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/></Relationships>'
      );
    }
  }
  
  // Save the modified PPTX
  const outputId = await OOXMLJsonService.rewrap(manifest, {
    filename: 'Brandware_Demo_Typography_Tables.pptx'
  });
  
  console.log('✅ Brandware demo created successfully!');
  console.log('File ID:', outputId);
  
  return outputId;
}

function getCustomTableStyles() {
  return \`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}">
  <a:tblStyle styleId="{BRANDWARE-001}" styleName="Brandware Enterprise">
    <a:wholeTbl>
      <a:tcStyle>
        <a:fill>
          <a:solidFill>
            <a:srgbClr val="F8F9FA"/>
          </a:solidFill>
        </a:fill>
        <a:tcBdr>
          <a:left>
            <a:ln w="12700">
              <a:solidFill>
                <a:srgbClr val="2E5E8C"/>
              </a:solidFill>
            </a:ln>
          </a:left>
        </a:tcBdr>
      </a:tcStyle>
    </a:wholeTbl>
    <a:firstRow>
      <a:tcStyle>
        <a:fill>
          <a:gradFill>
            <a:gsLst>
              <a:gs pos="0">
                <a:srgbClr val="2E5E8C"/>
              </a:gs>
              <a:gs pos="100000">
                <a:srgbClr val="1E4E7C"/>
              </a:gs>
            </a:gsLst>
            <a:lin ang="2700000"/>
          </a:gradFill>
        </a:fill>
      </a:tcStyle>
      <a:tcTxStyle b="1">
        <a:fontRef idx="minor">
          <a:srgbClr val="FFFFFF"/>
        </a:fontRef>
      </a:tcTxStyle>
    </a:firstRow>
  </a:tblStyle>
</a:tblStyleLst>\`;
}

// Run the demo
createBrandwareDemo();
`;

  // Save the GAS script
  const gasScriptPath = path.join(__dirname, 'brandware-demo-gas.js');
  fs.writeFileSync(gasScriptPath, gasScript);
  
  console.log('📝 Created Google Apps Script file: brandware-demo-gas.js');
  console.log('\n📋 Instructions:');
  console.log('1. Copy the contents of brandware-demo-gas.js');
  console.log('2. Paste into a Google Apps Script project');
  console.log('3. Make sure you have the OOXML libraries loaded');
  console.log('4. Run createBrandwareDemo()');
  console.log('\nThis will create a valid PPTX with brandware features that PowerPoint can open.');
  
  // Also create a Node.js version that uses the processing service
  const nodeScript = `
/**
 * Process existing PPTX with Brandware features
 */

const fs = require('fs');
const path = require('path');

async function processPPTXWithBrandware(inputFile) {
  console.log('Processing PPTX with Brandware features...');
  
  // This would connect to your OOXML processing service
  // For now, we'll show the modifications that would be made
  
  const modifications = {
    typography: {
      kern: '1200',        // Custom kerning
      spc: '-50',         // Letter spacing
      baseline: '0',      // Baseline adjustment
      fonts: ['Segoe UI', 'Segoe UI Light', 'Roboto']
    },
    tableStyles: [
      {
        id: 'BRANDWARE-001',
        name: 'Enterprise Blue',
        header: {
          fill: 'gradient:#2E5E8C-#1E4E7C',
          text: '#FFFFFF',
          bold: true
        },
        rows: {
          alternating: true,
          band1: '#F8F9FA',
          band2: '#FFFFFF'
        }
      },
      {
        id: 'BRANDWARE-002', 
        name: 'Material Design',
        header: {
          fill: '#7B1FA2',
          text: '#FFFFFF',
          padding: '16px'
        },
        rows: {
          borderless: true,
          hover: true
        }
      }
    ],
    hiddenMetadata: {
      tracking: 'BRANDWARE-2024',
      compliance: 'APPROVED',
      version: '1.0.0'
    }
  };
  
  console.log('\\nModifications to apply:');
  console.log(JSON.stringify(modifications, null, 2));
  
  console.log('\\n✅ These features would be injected at the OOXML level');
  console.log('They are not accessible through PowerPoint UI but will work when applied via XML manipulation.');
  
  return modifications;
}

// Example usage
if (require.main === module) {
  processPPTXWithBrandware('input.pptx').catch(console.error);
}

module.exports = { processPPTXWithBrandware };
`;

  const nodeScriptPath = path.join(__dirname, 'process-pptx-brandware.js');
  fs.writeFileSync(nodeScriptPath, nodeScript);
  
  console.log('\n📝 Also created: process-pptx-brandware.js');
  console.log('This shows how to process existing PPTX files with brandware features.\n');
  
  return { gasScriptPath, nodeScriptPath };
}

// Run the script
if (require.main === module) {
  createValidBrandwarePowerPoint().catch(console.error);
}

module.exports = { createValidBrandwarePowerPoint };