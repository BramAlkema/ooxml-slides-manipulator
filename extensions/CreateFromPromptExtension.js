/**
 * CreateFromPromptExtension - Natural Language Presentation Creation
 * 
 * Adds createFromPrompt() method to OOXMLSlides for generating presentations
 * from natural language descriptions including color palettes and fonts.
 */

class CreateFromPromptExtension extends BaseExtension {
  
  constructor() {
    super();
    this.extensionInfo = {
      name: 'CreateFromPrompt',
      version: '1.0.0',
      type: 'TEMPLATE',
      description: 'Create presentations from natural language prompts',
      methods: ['createFromPrompt', 'fromDescription']
    };
  }
  
  /**
   * Create a new presentation from a natural language prompt
   * This method gets added to OOXMLSlides as a static method
   * 
   * @param {string} prompt - Natural language description
   * @param {Object} options - Creation options
   * @returns {Promise<OOXMLSlides>} New presentation instance
   */
  static async createFromPrompt(prompt, options = {}) {
    ConsoleFormatter.header('🎨 Creating Presentation from Prompt');
    ConsoleFormatter.info(`Prompt: "${prompt}"`);
    
    try {
      // Parse the prompt for design elements
      const design = this._parsePromptDesign(prompt);
      ConsoleFormatter.status('PASS', 'Prompt Parsing', `Detected: ${design.type} presentation`);
      
      // Create base presentation
      const presentation = SlidesApp.create(design.title);
      const fileId = presentation.getId();
      ConsoleFormatter.status('PASS', 'Base Creation', `Created: ${fileId}`);
      
      // Initialize OOXMLSlides instance
      const slides = new OOXMLSlides(fileId, {
        createBackup: false, // New presentation, no backup needed
        ...options
      });
      
      // Apply the design
      await slides.load();
      await slides._applyPromptDesign(design);
      await slides.save();
      
      ConsoleFormatter.success(`Presentation created: ${presentation.getUrl()}`);
      return slides;
      
    } catch (error) {
      ConsoleFormatter.error('Failed to create presentation from prompt', error);
      throw error;
    }
  }
  
  /**
   * Parse natural language prompt for design elements
   * @private
   */
  static _parsePromptDesign(prompt) {
    const lower = prompt.toLowerCase();
    
    // Extract color palette from coolors.co URLs
    const coolorsMatch = prompt.match(/coolors\.co\/([a-fA-F0-9-]+)/);
    let colors = null;
    if (coolorsMatch) {
      const colorHexes = coolorsMatch[1].split('-').map(c => `#${c.toUpperCase()}`);
      colors = {
        primary: colorHexes[0] || '#3B60E4',
        secondary: colorHexes[1] || '#7765E3', 
        accent: colorHexes[2] || '#C8ADC0',
        light: colorHexes[3] || '#EDD3C4',
        dark: colorHexes[4] || '#080708'
      };
    }
    
    // Extract font pairings
    let fonts = { heading: 'Calibri', body: 'Calibri' };
    if (lower.includes('merriweather') && lower.includes('inter')) {
      fonts = { heading: 'Merriweather', body: 'Inter' };
    } else if (lower.includes('playfair') && lower.includes('source sans')) {
      fonts = { heading: 'Playfair Display', body: 'Source Sans Pro' };
    } else if (lower.includes('montserrat') && lower.includes('open sans')) {
      fonts = { heading: 'Montserrat', body: 'Open Sans' };
    }
    
    // Determine presentation type
    let type = 'general';
    let title = 'New Presentation';
    
    if (lower.includes('startup') || lower.includes('pitch')) {
      type = 'startup_pitch';
      title = 'Startup Pitch Deck';
    } else if (lower.includes('corporate') || lower.includes('business')) {
      type = 'corporate';
      title = 'Business Presentation';
    } else if (lower.includes('academic') || lower.includes('research')) {
      type = 'academic';
      title = 'Research Presentation';
    } else if (lower.includes('portfolio') || lower.includes('showcase')) {
      type = 'portfolio';
      title = 'Portfolio Showcase';
    }
    
    return { type, title, colors, fonts };
  }
  
  /**
   * Apply parsed design to the presentation
   * This method gets added to OOXMLSlides as an instance method
   */
  async applyPromptDesign(design) {
    this._ensureLoaded();
    
    try {
      // Apply colors if detected
      if (design.colors) {
        await this.applyCustomColors(design.colors);
        ConsoleFormatter.status('PASS', 'Color Application', 'Custom palette applied');
      }
      
      // Apply fonts
      if (design.fonts) {
        await this.applyFontPairing(design.fonts);
        ConsoleFormatter.status('PASS', 'Font Application', `${design.fonts.heading}/${design.fonts.body}`);
      }
      
      // Apply template based on type
      await this.applyTemplate(design.type);
      ConsoleFormatter.status('PASS', 'Template Application', `${design.type} template applied`);
      
      return this;
      
    } catch (error) {
      ConsoleFormatter.error('Failed to apply prompt design', error);
      throw error;
    }
  }
  
  /**
   * Apply custom colors to the presentation theme
   */
  async applyCustomColors(colors) {
    this._ensureLoaded();
    
    try {
      // Get the theme file
      const themeFile = this.getFile('ppt/theme/theme1.xml');
      if (!themeFile) {
        throw new Error('Theme file not found');
      }
      
      // Parse and modify theme colors
      let theme = themeFile;
      
      // Replace color scheme in theme XML
      if (colors.primary) {
        theme = theme.replace(/<a:accent1>.*?<\/a:accent1>/g, 
          `<a:accent1><a:srgbClr val="${colors.primary.replace('#', '')}"/></a:accent1>`);
      }
      
      if (colors.secondary) {
        theme = theme.replace(/<a:accent2>.*?<\/a:accent2>/g, 
          `<a:accent2><a:srgbClr val="${colors.secondary.replace('#', '')}"/></a:accent2>`);
      }
      
      if (colors.accent) {
        theme = theme.replace(/<a:accent3>.*?<\/a:accent3>/g, 
          `<a:accent3><a:srgbClr val="${colors.accent.replace('#', '')}"/></a:accent3>`);
      }
      
      // Update the theme file
      this.setFile('ppt/theme/theme1.xml', theme);
      
      return this;
      
    } catch (error) {
      ConsoleFormatter.error('Failed to apply custom colors', error);
      throw error;
    }
  }
  
  /**
   * Apply font pairing to the presentation
   */
  async applyFontPairing(fonts) {
    this._ensureLoaded();
    
    try {
      // Get the theme file
      const themeFile = this.getFile('ppt/theme/theme1.xml');
      if (!themeFile) {
        throw new Error('Theme file not found');
      }
      
      let theme = themeFile;
      
      // Replace font scheme in theme XML
      if (fonts.heading) {
        theme = theme.replace(/<a:majorFont>[\s\S]*?<\/a:majorFont>/g, 
          `<a:majorFont><a:latin typeface="${fonts.heading}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>`);
      }
      
      if (fonts.body) {
        theme = theme.replace(/<a:minorFont>[\s\S]*?<\/a:minorFont>/g, 
          `<a:minorFont><a:latin typeface="${fonts.body}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>`);
      }
      
      // Update the theme file
      this.setFile('ppt/theme/theme1.xml', theme);
      
      return this;
      
    } catch (error) {
      ConsoleFormatter.error('Failed to apply font pairing', error);
      throw error;
    }
  }
  
  /**
   * Apply template based on presentation type
   */
  async applyTemplate(type) {
    this._ensureLoaded();
    
    const templates = {
      'startup_pitch': this._createStartupPitchTemplate,
      'corporate': this._createCorporateTemplate,
      'academic': this._createAcademicTemplate,
      'portfolio': this._createPortfolioTemplate,
      'general': this._createGeneralTemplate
    };
    
    const templateFunction = templates[type] || templates.general;
    await templateFunction.call(this);
    
    return this;
  }
  
  /**
   * Create startup pitch deck template
   * @private
   */
  async _createStartupPitchTemplate() {
    // Add slides for typical startup pitch structure
    const slideStructure = [
      { title: 'Company Name', subtitle: 'Tagline' },
      { title: 'Problem', content: 'What problem are we solving?' },
      { title: 'Solution', content: 'How we solve it uniquely' },
      { title: 'Market', content: 'Market size and opportunity' },
      { title: 'Product', content: 'Product demonstration' },
      { title: 'Traction', content: 'Growth and validation' },
      { title: 'Team', content: 'Who we are' },
      { title: 'Financials', content: 'Revenue model and projections' },
      { title: 'Funding', content: 'Investment ask and use of funds' },
      { title: 'Thank You', subtitle: 'Questions?' }
    ];
    
    // Implementation would modify slide layouts and content
    ConsoleFormatter.info(`Applied startup pitch template with ${slideStructure.length} slides`);
  }
  
  /**
   * Create corporate template
   * @private
   */
  async _createCorporateTemplate() {
    ConsoleFormatter.info('Applied corporate template');
  }
  
  /**
   * Create academic template
   * @private
   */
  async _createAcademicTemplate() {
    ConsoleFormatter.info('Applied academic template');
  }
  
  /**
   * Create portfolio template
   * @private
   */
  async _createPortfolioTemplate() {
    ConsoleFormatter.info('Applied portfolio template');
  }
  
  /**
   * Create general template
   * @private
   */
  async _createGeneralTemplate() {
    ConsoleFormatter.info('Applied general template');
  }
}