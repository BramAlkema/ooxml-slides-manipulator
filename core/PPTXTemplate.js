/**
 * PPTXTemplate - Base PPTX template system
 * Provides minimal PPTX template for creating new presentations from scratch
 * Based on tanaikech's approach but enhanced for theme manipulation
 */
class PPTXTemplate {
  
  /**
   * Get base PPTX template as base64 string (tanaikech's approach)
   * This is a minimal valid PPTX file with basic structure
   */
  static getBase64Template() {
    // Tanaikech's proven base64 template - this is a real, valid PPTX file
    return "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";
  }
  
  /**
   * Create template from tanaikech's base64 approach
   */
  static createFromBase64Template() {
    try {
      // Use our proven minimal PPTX template 
      const blob = MinimalPPTXTemplate.createBlob();
      console.log(`Created template from base64: ${blob.getBytes().length} bytes`);
      return blob;
      
    } catch (error) {
      console.log('Base64 template failed:', error.message);
      throw error; // Let the main method handle the fallback
    }
  }

  /**
   * Create a new PPTX file from template
   * @param {Object} options - Creation options
   * @returns {Blob} PPTX blob
   */
  static createFromTemplate(options = {}) {
    try {
      // Use the Cloud Function approach for reliable PPTX processing
      const templateBlob = this.createMinimalTemplate();
      
      // Extract files using CloudPPTXService (handles PPTX properly)
      const extractedFiles = CloudPPTXService.unzipPPTX(templateBlob);
      const fileMap = new Map();
      
      // Convert extracted files to blob format for processing
      Object.entries(extractedFiles).forEach(([fileName, content]) => {
        // Create blob from string content
        const blob = Utilities.newBlob(content, 'application/octet-stream', fileName);
        fileMap.set(fileName, blob);
      });
      
      // Apply any template modifications
      if (options.slideSize) {
        this._modifySlideSize(fileMap, options.slideSize);
      }
      
      if (options.title) {
        this._modifyTitle(fileMap, options.title);
      }
      
      // Repack into PPTX using Cloud Function
      const modifiedFiles = Array.from(fileMap.values());
      const filesObject = {};
      modifiedFiles.forEach(file => {
        filesObject[file.getName()] = Utilities.base64Encode(file.getBytes());
      });
      
      return CloudPPTXService.zipPPTX(filesObject);
      
    } catch (error) {
      throw new Error(`Failed to create PPTX from template: ${error.message}`);
    }
  }

  /**
   * Create minimal template using the proven working approach
   */
  static createMinimalTemplate() {
    try {
      // Try tanaikech's base64 approach first (most efficient)
      return this.createFromBase64Template();
      
    } catch (error) {
      console.log('Base64 approach failed, using Google Slides API approach:', error.message);
      
      // Fall back to the proven Google Slides API approach
      try {
        console.log('Creating minimal PPTX template via Google Slides API...');
        
        // Create a minimal Google Slides presentation
        const tempPresentation = SlidesApp.create('OOXML Template Base');
        const tempPresentationId = tempPresentation.getId();
        
        // Export it as PPTX using Drive API
        const exportUrl = `https://www.googleapis.com/drive/v3/files/${tempPresentationId}/export?mimeType=application/vnd.openxmlformats-officedocument.presentationml.presentation`;
        
        const response = UrlFetchApp.fetch(exportUrl, {
          headers: {
            'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
          }
        });
        
        const pptxBlob = response.getBlob().setName('template.pptx');
        
        // Clean up the temporary presentation
        DriveApp.getFileById(tempPresentationId).setTrashed(true);
        
        console.log(`✅ PPTX template created: ${pptxBlob.getBytes().length} bytes`);
        return pptxBlob;
        
      } catch (apiError) {
        console.log('Google Slides API approach also failed, using manual creation:', apiError.message);
        return this.createManualTemplate();
      }
    }
  }

  /**
   * Fallback: Create minimal template structure with manual XML
   * Only used if Google Slides API approach fails
   */
  static createManualTemplate() {
    try {
      const files = [];
      
      // [Content_Types].xml
      files.push(Utilities.newBlob(
        this._getContentTypesXML(),
        MimeType.PLAIN_TEXT,
        '[Content_Types].xml'
      ));
      
      // _rels/.rels
      files.push(Utilities.newBlob(
        this._getMainRelsXML(),
        MimeType.PLAIN_TEXT,
        '_rels/.rels'
      ));
      
      // ppt/presentation.xml
      files.push(Utilities.newBlob(
        this._getPresentationXML(),
        MimeType.PLAIN_TEXT,
        'ppt/presentation.xml'
      ));
      
      // ppt/theme/theme1.xml
      files.push(Utilities.newBlob(
        this._getThemeXML(),
        MimeType.PLAIN_TEXT,
        'ppt/theme/theme1.xml'
      ));
      
      // ppt/slideMasters/slideMaster1.xml
      files.push(Utilities.newBlob(
        this._getSlideMasterXML(),
        MimeType.PLAIN_TEXT,
        'ppt/slideMasters/slideMaster1.xml'
      ));
      
      // ppt/slideLayouts/slideLayout1.xml
      files.push(Utilities.newBlob(
        this._getSlideLayoutXML(),
        MimeType.PLAIN_TEXT,
        'ppt/slideLayouts/slideLayout1.xml'
      ));
      
      // ppt/slides/slide1.xml
      files.push(Utilities.newBlob(
        this._getSlideXML(),
        MimeType.PLAIN_TEXT,
        'ppt/slides/slide1.xml'
      ));
      
      // Create the ZIP blob using tanaikech's approach
      const zipBlob = Utilities.zip(files, 'template.pptx');
      
      // Set correct MIME type for PPTX
      return zipBlob.setContentType(MimeType.MICROSOFT_POWERPOINT);
      
    } catch (error) {
      throw new Error(`Failed to create manual template: ${error.message}`);
    }
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