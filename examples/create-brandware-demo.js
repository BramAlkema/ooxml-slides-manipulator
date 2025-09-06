/**
 * Create Brandware Demo PowerPoint
 * 
 * This script creates a PowerPoint presentation demonstrating advanced
 * typography and table styles using the Brandware XML hacking techniques.
 * 
 * The generated PPTX will showcase:
 * - Custom table styles not available in PowerPoint UI
 * - Advanced typography with custom kerning
 * - Enterprise color schemes
 * - Material Design tables
 * - Accessibility-compliant formatting
 */

const fs = require('fs');
const path = require('path');

// Import required libraries
let JSZip, OOXMLJsonService;

try {
  JSZip = require('jszip');
} catch (e) {
  console.log('JSZip not found, using mock for demo');
}

/**
 * Create a minimal PPTX structure with advanced features
 */
async function createBrandwarePowerPoint() {
  console.log('🎨 Creating Brandware Demo PowerPoint...\n');
  
  const zip = new JSZip();
  
  // Create the basic PPTX structure
  const pptxStructure = {
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`,

    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`,

    'docProps/core.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Brandware Typography &amp; Table Styles Demo</dc:title>
  <dc:creator>OOXML Brandware Engine</dc:creator>
  <cp:category>DEMO-BRANDWARE-2024</cp:category>
  <cp:contentStatus>SHOWCASE-ADVANCED-FEATURES</cp:contentStatus>
  <dcterms:created xsi:type="dcterms:W3CDTF">2024-01-01T00:00:00Z</dcterms:created>
</cp:coreProperties>`,

    'docProps/app.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>OOXML Brandware Demo</Application>
  <PresentationFormat>Custom</PresentationFormat>
  <TotalTime>0</TotalTime>
  <Slides>1</Slides>
</Properties>`,

    'ppt/_rels/presentation.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>
</Relationships>`,

    'ppt/presentation.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <!-- BRANDWARE-TRACKING:DEMO-2024 -->
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId3"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle>
    <!-- Advanced Typography Settings -->
    <a:defPPr>
      <a:defRPr lang="en-US" sz="1800" kern="1200" spc="0" baseline="0">
        <!-- Custom kerning and spacing -->
        <a:latin typeface="Segoe UI" pitchFamily="34" charset="0"/>
        <a:ea typeface="Arial Unicode MS"/>
        <a:cs typeface="Arial Unicode MS"/>
      </a:defRPr>
    </a:defPPr>
  </p:defaultTextStyle>
</p:presentation>`,

    // CUSTOM TABLE STYLES - The Brandware Magic!
    'ppt/tableStyles.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}">
  <!-- Enterprise Blue Table Style -->
  <a:tblStyle styleId="{E8034E78-1234-5678-9ABC-ENTERPRISE001}" styleName="Enterprise Blue">
    <!-- Whole Table -->
    <a:wholeTbl>
      <a:tcTxStyle>
        <a:fontRef idx="minor">
          <a:scrgbClr r="10" g="10" b="10"/>
        </a:fontRef>
        <a:schemeClr val="tx1"/>
      </a:tcTxStyle>
      <a:tcStyle>
        <a:fill>
          <a:solidFill>
            <a:srgbClr val="F0F4F8"/>
          </a:solidFill>
        </a:fill>
        <a:tcBdr>
          <a:left>
            <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
              <a:solidFill>
                <a:srgbClr val="2E5E8C"/>
              </a:solidFill>
            </a:ln>
          </a:left>
          <a:right>
            <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
              <a:solidFill>
                <a:srgbClr val="2E5E8C"/>
              </a:solidFill>
            </a:ln>
          </a:right>
          <a:top>
            <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
              <a:solidFill>
                <a:srgbClr val="2E5E8C"/>
              </a:solidFill>
            </a:ln>
          </a:top>
          <a:bottom>
            <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
              <a:solidFill>
                <a:srgbClr val="2E5E8C"/>
              </a:solidFill>
            </a:ln>
          </a:bottom>
        </a:tcBdr>
      </a:tcStyle>
    </a:wholeTbl>
    <!-- First Row (Header) -->
    <a:firstRow>
      <a:tcTxStyle>
        <a:fontRef idx="minor">
          <a:scrgbClr r="100" g="100" b="100"/>
        </a:fontRef>
        <a:b/>
      </a:tcTxStyle>
      <a:tcStyle>
        <a:fill>
          <a:gradFill rotWithShape="1">
            <a:gsLst>
              <a:gs pos="0">
                <a:srgbClr val="2E5E8C"/>
              </a:gs>
              <a:gs pos="100000">
                <a:srgbClr val="1E4E7C"/>
              </a:gs>
            </a:gsLst>
            <a:lin ang="2700000" scaled="1"/>
          </a:gradFill>
        </a:fill>
      </a:tcStyle>
    </a:firstRow>
    <!-- Banded Rows -->
    <a:band1H>
      <a:tcStyle>
        <a:fill>
          <a:solidFill>
            <a:srgbClr val="E8F2FC" alpha="50000"/>
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
  
  <!-- Material Design Table Style -->
  <a:tblStyle styleId="{MAT00001-2345-6789-ABCD-MATERIAL002}" styleName="Material Purple">
    <a:wholeTbl>
      <a:tcTxStyle>
        <a:fontRef idx="minor">
          <a:latin typeface="Roboto"/>
        </a:fontRef>
      </a:tcTxStyle>
      <a:tcStyle>
        <a:fill>
          <a:solidFill>
            <a:srgbClr val="FFFFFF"/>
          </a:solidFill>
        </a:fill>
        <a:tcBdr>
          <a:bottom>
            <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">
              <a:solidFill>
                <a:srgbClr val="E0E0E0"/>
              </a:solidFill>
            </a:ln>
          </a:bottom>
        </a:tcBdr>
      </a:tcStyle>
    </a:wholeTbl>
    <a:firstRow>
      <a:tcTxStyle>
        <a:b/>
        <a:fontRef idx="minor">
          <a:solidFill>
            <a:srgbClr val="FFFFFF"/>
          </a:solidFill>
        </a:fontRef>
      </a:tcTxStyle>
      <a:tcStyle>
        <a:fill>
          <a:solidFill>
            <a:srgbClr val="7B1FA2"/>
          </a:solidFill>
        </a:fill>
      </a:tcStyle>
    </a:firstRow>
  </a:tblStyle>
</a:tblStyleLst>`,

    // Slide with advanced typography and table
    'ppt/slides/slide1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="9144000" cy="6858000"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="9144000" cy="6858000"/>
        </a:xfrm>
      </p:grpSpPr>
      
      <!-- Title with Advanced Typography -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="274320"/>
            <a:ext cx="8229600" cy="1143000"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr">
              <a:spcBef>
                <a:spcPts val="0"/>
              </a:spcBef>
            </a:pPr>
            <!-- Advanced Typography: Custom Kerning -->
            <a:r>
              <a:rPr lang="en-US" sz="4400" b="1" kern="1400" spc="-50" baseline="0">
                <a:solidFill>
                  <a:srgbClr val="2E5E8C"/>
                </a:solidFill>
                <a:latin typeface="Segoe UI Light" pitchFamily="34" charset="0"/>
                <a:effectLst>
                  <a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl">
                    <a:srgbClr val="000000" alpha="25000"/>
                  </a:outerShdw>
                </a:effectLst>
              </a:rPr>
              <a:t>TYPOGRAPHY</a:t>
            </a:r>
            <a:r>
              <a:rPr lang="en-US" sz="4400" b="1" kern="800" spc="100">
                <a:solidFill>
                  <a:srgbClr val="7B1FA2"/>
                </a:solidFill>
                <a:latin typeface="Segoe UI" pitchFamily="34" charset="0"/>
              </a:rPr>
              <a:t> &amp; </a:t>
            </a:r>
            <a:r>
              <a:rPr lang="en-US" sz="4400" b="1" kern="1200" spc="-25">
                <a:solidFill>
                  <a:srgbClr val="2E5E8C"/>
                </a:solidFill>
                <a:latin typeface="Segoe UI Light" pitchFamily="34" charset="0"/>
              </a:rPr>
              <a:t>TABLES</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      
      <!-- Subtitle with Typography Features -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Subtitle"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="1600200"/>
            <a:ext cx="8229600" cy="571500"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr"/>
            <a:r>
              <a:rPr lang="en-US" sz="2000" kern="200" spc="50">
                <a:solidFill>
                  <a:srgbClr val="666666"/>
                </a:solidFill>
                <a:latin typeface="Segoe UI Semilight"/>
              </a:rPr>
              <a:t>Advanced OOXML Features • Custom Kerning • Enterprise Tables</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      
      <!-- Custom Table with Enterprise Style -->
      <p:graphicFrame>
        <p:nvGraphicFramePr>
          <p:cNvPr id="4" name="Table"/>
          <p:cNvGraphicFramePr>
            <a:graphicFrameLocks noGrp="1"/>
          </p:cNvGraphicFramePr>
          <p:nvPr/>
        </p:nvGraphicFramePr>
        <p:xfrm>
          <a:off x="457200" y="2514600"/>
          <a:ext cx="8229600" cy="3657600"/>
        </p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl>
              <a:tblPr firstRow="1" bandRow="1">
                <a:tableStyleId>{E8034E78-1234-5678-9ABC-ENTERPRISE001}</a:tableStyleId>
              </a:tblPr>
              <a:tblGrid>
                <a:gridCol w="2743200"/>
                <a:gridCol w="2743200"/>
                <a:gridCol w="2743200"/>
              </a:tblGrid>
              <!-- Header Row -->
              <a:tr h="635000">
                <a:tc>
                  <a:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                      <a:r>
                        <a:rPr lang="en-US" sz="1800" b="1">
                          <a:solidFill>
                            <a:srgbClr val="FFFFFF"/>
                          </a:solidFill>
                        </a:rPr>
                        <a:t>Feature</a:t>
                      </a:r>
                    </a:p>
                  </a:txBody>
                  <a:tcPr/>
                </a:tc>
                <a:tc>
                  <a:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                      <a:r>
                        <a:rPr lang="en-US" sz="1800" b="1">
                          <a:solidFill>
                            <a:srgbClr val="FFFFFF"/>
                          </a:solidFill>
                        </a:rPr>
                        <a:t>Standard PowerPoint</a:t>
                      </a:r>
                    </a:p>
                  </a:txBody>
                  <a:tcPr/>
                </a:tc>
                <a:tc>
                  <a:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                      <a:r>
                        <a:rPr lang="en-US" sz="1800" b="1">
                          <a:solidFill>
                            <a:srgbClr val="FFFFFF"/>
                          </a:solidFill>
                        </a:rPr>
                        <a:t>Brandware OOXML</a:t>
                      </a:r>
                    </a:p>
                  </a:txBody>
                  <a:tcPr/>
                </a:tc>
              </a:tr>
              <!-- Data Rows -->
              <a:tr h="508000">
                <a:tc>
                  <a:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                      <a:r>
                        <a:rPr lang="en-US" sz="1600" kern="100">
                          <a:solidFill>
                            <a:srgbClr val="2C3E50"/>
                          </a:solidFill>
                        </a:rPr>
                        <a:t>Kerning Control</a:t>
                      </a:r>
                    </a:p>
                  </a:txBody>
                  <a:tcPr/>
                </a:tc>
                <a:tc>
                  <a:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                      <a:r>
                        <a:rPr lang="en-US" sz="1600">
                          <a:solidFill>
                            <a:srgbClr val="666666"/>
                          </a:solidFill>
                        </a:rPr>
                        <a:t>Limited</a:t>
                      </a:r>
                    </a:p>
                  </a:txBody>
                  <a:tcPr/>
                </a:tc>
                <a:tc>
                  <a:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                      <a:r>
                        <a:rPr lang="en-US" sz="1600" b="1">
                          <a:solidFill>
                            <a:srgbClr val="27AE60"/>
                          </a:solidFill>
                        </a:rPr>
                        <a:t>Full Control</a:t>
                      </a:r>
                    </a:p>
                  </a:txBody>
                  <a:tcPr/>
                </a:tc>
              </a:tr>
              <a:tr h="508000">
                <a:tc>
                  <a:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                      <a:r>
                        <a:rPr lang="en-US" sz="1600" kern="100">
                          <a:solidFill>
                            <a:srgbClr val="2C3E50"/>
                          </a:solidFill>
                        </a:rPr>
                        <a:t>Table Styles</a:t>
                      </a:r>
                    </a:p>
                  </a:txBody>
                  <a:tcPr/>
                </a:tc>
                <a:tc>
                  <a:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                      <a:r>
                        <a:rPr lang="en-US" sz="1600">
                          <a:solidFill>
                            <a:srgbClr val="666666"/>
                          </a:solidFill>
                        </a:rPr>
                        <a:t>Built-in Only</a:t>
                      </a:r>
                    </a:p>
                  </a:txBody>
                  <a:tcPr/>
                </a:tc>
                <a:tc>
                  <a:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                      <a:r>
                        <a:rPr lang="en-US" sz="1600" b="1">
                          <a:solidFill>
                            <a:srgbClr val="27AE60"/>
                          </a:solidFill>
                        </a:rPr>
                        <a:t>Unlimited Custom</a:t>
                      </a:r>
                    </a:p>
                  </a:txBody>
                  <a:tcPr/>
                </a:tc>
              </a:tr>
              <a:tr h="508000">
                <a:tc>
                  <a:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                      <a:r>
                        <a:rPr lang="en-US" sz="1600" kern="100">
                          <a:solidFill>
                            <a:srgbClr val="2C3E50"/>
                          </a:solidFill>
                        </a:rPr>
                        <a:t>Gradient Fills</a:t>
                      </a:r>
                    </a:p>
                  </a:txBody>
                  <a:tcPr/>
                </a:tc>
                <a:tc>
                  <a:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                      <a:r>
                        <a:rPr lang="en-US" sz="1600">
                          <a:solidFill>
                            <a:srgbClr val="666666"/>
                          </a:solidFill>
                        </a:rPr>
                        <a:t>Basic</a:t>
                      </a:r>
                    </a:p>
                  </a:txBody>
                  <a:tcPr/>
                </a:tc>
                <a:tc>
                  <a:txBody>
                    <a:bodyPr/>
                    <a:lstStyle/>
                    <a:p>
                      <a:r>
                        <a:rPr lang="en-US" sz="1600" b="1">
                          <a:solidFill>
                            <a:srgbClr val="27AE60"/>
                          </a:solidFill>
                        </a:rPr>
                        <a:t>Multi-stop Advanced</a:t>
                      </a:r>
                    </a:p>
                  </a:txBody>
                  <a:tcPr/>
                </a:tc>
              </a:tr>
            </a:tbl>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
      
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`,

    'ppt/slides/_rels/slide1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`,

    // Minimal required files for valid PPTX
    'ppt/slideLayouts/slideLayout1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sldLayout>`,

    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`,

    'ppt/slideMasters/slideMaster1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgRef idx="1001">
        <a:schemeClr val="bg1"/>
      </p:bgRef>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>`,

    'ppt/slideMasters/_rels/slideMaster1.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`,

    'ppt/theme/theme1.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Brandware">
  <a:themeElements>
    <a:clrScheme name="Brandware">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1E4E7C"/></a:dk2>
      <a:lt2><a:srgbClr val="E8F2FC"/></a:lt2>
      <a:accent1><a:srgbClr val="2E5E8C"/></a:accent1>
      <a:accent2><a:srgbClr val="7B1FA2"/></a:accent2>
      <a:accent3><a:srgbClr val="27AE60"/></a:accent3>
      <a:accent4><a:srgbClr val="F39C12"/></a:accent4>
      <a:accent5><a:srgbClr val="E74C3C"/></a:accent5>
      <a:accent6><a:srgbClr val="3498DB"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Brandware">
      <a:majorFont>
        <a:latin typeface="Segoe UI Light" panose="020B0503020204020204"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Segoe UI" panose="020B0503020204020204"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Brandware">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
            <a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="1"/>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
            <a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="0"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
            <a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>`
  };

  // Add all files to the ZIP
  for (const [path, content] of Object.entries(pptxStructure)) {
    zip.file(path, content);
  }

  // Generate the PPTX file
  const pptxBuffer = await zip.generateAsync({ 
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 9 }
  });

  // Save the file
  const outputPath = path.join(__dirname, 'brandware-typography-tables-demo.pptx');
  fs.writeFileSync(outputPath, pptxBuffer);

  console.log('✅ PowerPoint created successfully!');
  console.log(`📁 File saved to: ${outputPath}`);
  console.log('\n📊 Features demonstrated:');
  console.log('  • Custom kerning values (kern="1400", kern="800", kern="1200")');
  console.log('  • Letter spacing control (spc="-50", spc="100")');
  console.log('  • Custom table style "Enterprise Blue" with gradient header');
  console.log('  • Material Design table style "Material Purple"');
  console.log('  • Banded rows with transparency');
  console.log('  • Advanced typography in table cells');
  console.log('  • Multiple font weights and styles');
  console.log('\n🎨 This PPTX contains OOXML features not accessible through PowerPoint UI!');
  
  return outputPath;
}

// Run the demo
if (require.main === module) {
  createBrandwarePowerPoint().catch(console.error);
}

module.exports = { createBrandwarePowerPoint };