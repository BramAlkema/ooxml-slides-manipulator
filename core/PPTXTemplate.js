/**
 * PPTXTemplate - Base PPTX template system
 * Provides minimal PPTX template for creating new presentations from scratch
 * Based on tanaikech's approach but enhanced for theme manipulation
 */
class PPTXTemplate {
  
  /**
   * Get base PPTX template as base64 string
   * This is a minimal valid PPTX file with basic structure
   */
  static getBase64Template() {
    // Note: This is a placeholder base64 string
    // In production, this should be a proper base64-encoded minimal PPTX file
    // For now, we'll use the createMinimalTemplate() method instead
    return null;
  }

  /**
   * Create a new PPTX file from template
   * @param {Object} options - Creation options
   * @returns {Blob} PPTX blob
   */
  static createFromTemplate(options = {}) {
    try {
      // For now, use the minimal template approach
      // TODO: Replace with proper base64 template in production
      const templateBlob = this.createMinimalTemplate();
      
      // Extract files from template
      const files = Utilities.unzip(templateBlob);
      const fileMap = new Map();
      
      files.forEach(file => {
        fileMap.set(file.getName(), file);
      });
      
      // Apply any template modifications
      if (options.slideSize) {
        this._modifySlideSize(fileMap, options.slideSize);
      }
      
      if (options.title) {
        this._modifyTitle(fileMap, options.title);
      }
      
      // Repack into PPTX
      const modifiedFiles = Array.from(fileMap.values());
      return Utilities.zip(modifiedFiles);
      
    } catch (error) {
      throw new Error(`Failed to create PPTX from template: ${error.message}`);
    }
  }

  /**
   * Create minimal template structure without base64
   * For testing when base64 template is not available
   */
  static createMinimalTemplate() {
    const files = [];
    
    // [Content_Types].xml
    files.push(Utilities.newBlob(
      this._getContentTypesXML(),
      'application/xml',
      '[Content_Types].xml'
    ));
    
    // _rels/.rels
    files.push(Utilities.newBlob(
      this._getMainRelsXML(),
      'application/xml',
      '_rels/.rels'
    ));
    
    // ppt/presentation.xml
    files.push(Utilities.newBlob(
      this._getPresentationXML(),
      'application/xml',
      'ppt/presentation.xml'
    ));
    
    // ppt/theme/theme1.xml
    files.push(Utilities.newBlob(
      this._getThemeXML(),
      'application/xml',
      'ppt/theme/theme1.xml'
    ));
    
    // ppt/slideMasters/slideMaster1.xml
    files.push(Utilities.newBlob(
      this._getSlideMasterXML(),
      'application/xml',
      'ppt/slideMasters/slideMaster1.xml'
    ));
    
    // ppt/slideLayouts/slideLayout1.xml
    files.push(Utilities.newBlob(
      this._getSlideLayoutXML(),
      'application/xml',
      'ppt/slideLayouts/slideLayout1.xml'
    ));
    
    // ppt/slides/slide1.xml
    files.push(Utilities.newBlob(
      this._getSlideXML(),
      'application/xml',
      'ppt/slides/slide1.xml'
    ));
    
    return Utilities.zip(files);
  }

  // Private helper methods for template modification
  static _modifySlideSize(fileMap, size) {
    const presFile = fileMap.get('ppt/presentation.xml');
    if (!presFile) return;
    
    const presXML = presFile.getDataAsString();
    const presDoc = XmlService.parse(presXML);
    const root = presDoc.getRootElement();
    
    // Find sldSz element and update dimensions
    const sldSz = root.getChild('sldSz', XmlService.getNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main'));
    
    if (sldSz) {
      // Convert pixels to EMU (English Metric Units)
      const widthEMU = Math.round(size.width * 12700);
      const heightEMU = Math.round(size.height * 12700);
      
      sldSz.setAttribute('cx', widthEMU.toString());
      sldSz.setAttribute('cy', heightEMU.toString());
    }
    
    // Update the file in the map
    const modifiedXML = XmlService.getRawFormat().format(presDoc);
    fileMap.set('ppt/presentation.xml', Utilities.newBlob(modifiedXML, 'application/xml', 'ppt/presentation.xml'));
  }

  static _modifyTitle(fileMap, title) {
    // This would modify the title in slide1.xml
    // Implementation depends on slide structure
  }

  // XML template generators
  static _getContentTypesXML() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-presentationml.slide+xml"/>
</Types>`;
  }

  static _getMainRelsXML() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;
  }

  static _getPresentationXML() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle>
    <a:defPPr>
      <a:defRPr lang="en-US"/>
    </a:defPPr>
  </p:defaultTextStyle>
</p:presentation>`;
  }

  static _getThemeXML() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText"/></a:dk1>
      <a:lt1><a:sysClr val="window"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="5B9BD5"/></a:accent1>
      <a:accent2><a:srgbClr val="70AD47"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="4472C4"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office"/>
  </a:themeElements>
</a:theme>`;
  }

  static _getSlideMasterXML() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
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
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>`;
  }

  static _getSlideLayoutXML() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" type="title" preserve="1">
  <p:cSld name="Title Slide">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
</p:sldLayout>`;
  }

  static _getSlideXML() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
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
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
</p:sld>`;
  }
}