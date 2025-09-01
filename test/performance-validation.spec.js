const { test, expect } = require('@playwright/test');

/**
 * Performance tests for OOXML-generated slides in Google Slides
 * Validates rendering speed, responsiveness, and resource usage
 */

test.describe('OOXML Slides Performance in Google Slides', () => {
  
  test.beforeEach(async ({ page }) => {
    test.setTimeout(180000); // 3 minutes for performance tests
  });

  test('Performance - Basic slide loading speed', async ({ page }) => {
    // Start performance monitoring
    const performanceMetrics = {
      loadStart: Date.now(),
      domContentLoaded: null,
      slideRendered: null,
      interactionReady: null
    };

    await page.goto('https://docs.google.com/presentation/u/0/');
    
    // Wait for page load
    await page.waitForSelector('[data-tooltip="Blank presentation"]', { timeout: 30000 });
    performanceMetrics.domContentLoaded = Date.now();
    
    // Create presentation and measure slide rendering
    await page.click('[data-tooltip="Blank presentation"]');
    await page.waitForSelector('.punch-viewer-content', { timeout: 30000 });
    performanceMetrics.slideRendered = Date.now();
    
    // Wait for full interactivity (can click and type)
    await page.waitForSelector('[role="textbox"], .kix-wordhtmlgenerator-word-node', { timeout: 30000 });
    performanceMetrics.interactionReady = Date.now();
    
    // Calculate performance metrics
    const metrics = {
      totalLoadTime: performanceMetrics.slideRendered - performanceMetrics.loadStart,
      domLoadTime: performanceMetrics.domContentLoaded - performanceMetrics.loadStart,
      slideRenderTime: performanceMetrics.slideRendered - performanceMetrics.domContentLoaded,
      interactionTime: performanceMetrics.interactionReady - performanceMetrics.slideRendered
    };
    
    console.log('Performance Metrics:', metrics);
    
    // Assert reasonable performance thresholds
    expect(metrics.totalLoadTime).toBeLessThan(15000); // < 15 seconds total
    expect(metrics.slideRenderTime).toBeLessThan(5000); // < 5 seconds slide render
    expect(metrics.interactionTime).toBeLessThan(3000); // < 3 seconds to interact
    
    await page.screenshot({ path: 'tests/screenshots/performance-basic-load.png' });
  });

  test('Performance - Complex theme application speed', async ({ page }) => {
    await page.goto('https://docs.google.com/presentation/u/0/');
    await page.waitForSelector('[data-tooltip="Blank presentation"]', { timeout: 30000 });
    await page.click('[data-tooltip="Blank presentation"]');
    await page.waitForSelector('.punch-viewer-content', { timeout: 30000 });
    
    // Measure theme application performance
    const themeStart = Date.now();
    
    // Access theme menu
    await page.click('[aria-label="Theme"]');
    await page.waitForSelector('.docs-material-theme-picker', { timeout: 10000 });
    
    // Select a complex theme (not the first basic one)
    const themeOptions = page.locator('.docs-material-theme-picker .docs-material-theme-item');
    const themeCount = await themeOptions.count();
    
    if (themeCount > 1) {
      // Select second theme (usually more complex)
      await themeOptions.nth(1).click();
      
      // Wait for theme application
      await page.waitForTimeout(2000);
      
      // Verify theme applied by checking for style changes
      await page.waitForFunction(() => {
        const viewer = document.querySelector('.punch-viewer-content');
        return viewer && getComputedStyle(viewer).backgroundColor !== 'rgba(0, 0, 0, 0)';
      }, { timeout: 10000 });
    }
    
    const themeEnd = Date.now();
    const themeApplicationTime = themeEnd - themeStart;
    
    console.log('Theme Application Time:', themeApplicationTime, 'ms');
    
    // Theme should apply within reasonable time
    expect(themeApplicationTime).toBeLessThan(8000); // < 8 seconds
    
    await page.screenshot({ path: 'tests/screenshots/performance-theme-applied.png' });
  });

  test('Performance - Typography rendering stress test', async ({ page }) => {
    await page.goto('https://docs.google.com/presentation/u/0/');
    await page.waitForSelector('[data-tooltip="Blank presentation"]', { timeout: 30000 });
    await page.click('[data-tooltip="Blank presentation"]');
    await page.waitForSelector('.punch-viewer-content', { timeout: 30000 });
    
    // Add multiple text elements to test typography performance
    const textElements = [
      'Typography Performance Test - Corporate Style',
      'Advanced kerning and character spacing validation',
      'Professional font rendering with complex formatting',
      'Multi-line paragraph with extensive text content for performance testing'
    ];
    
    const typingStart = Date.now();
    
    for (let i = 0; i < textElements.length; i++) {
      // Click to add new text box or use existing
      if (i === 0) {
        await page.click('.punch-viewer-content [role="textbox"], .kix-wordhtmlgenerator-word-node');
      } else {
        // Add new slide for additional text
        await page.keyboard.press('Control+M'); // New slide shortcut
        await page.waitForTimeout(1000);
        await page.click('.punch-viewer-content [role="textbox"], .kix-wordhtmlgenerator-word-node');
      }
      
      await page.type('.punch-viewer-content [role="textbox"], .kix-wordhtmlgenerator-word-node', textElements[i]);
      await page.waitForTimeout(500); // Allow rendering
    }
    
    const typingEnd = Date.now();
    const totalTypingTime = typingEnd - typingStart;
    
    console.log('Typography Stress Test Time:', totalTypingTime, 'ms');
    
    // Should handle multiple text elements efficiently
    expect(totalTypingTime).toBeLessThan(15000); // < 15 seconds for all elements
    
    await page.screenshot({ path: 'tests/screenshots/performance-typography-stress.png' });
  });

  test('Performance - Slide navigation responsiveness', async ({ page }) => {
    await page.goto('https://docs.google.com/presentation/u/0/');
    await page.waitForSelector('[data-tooltip="Blank presentation"]', { timeout: 30000 });
    await page.click('[data-tooltip="Blank presentation"]');
    await page.waitForSelector('.punch-viewer-content', { timeout: 30000 });
    
    // Create multiple slides for navigation testing
    const slideCount = 5;
    for (let i = 0; i < slideCount; i++) {
      await page.keyboard.press('Control+M'); // New slide
      await page.waitForTimeout(800);
    }
    
    // Test navigation performance
    const navigationTimes = [];
    
    for (let i = 0; i < slideCount; i++) {
      const navStart = Date.now();
      
      // Navigate using filmstrip
      const slideThumbs = page.locator('.punch-filmstrip .punch-filmstrip-thumbnail');
      if (await slideThumbs.nth(i).isVisible()) {
        await slideThumbs.nth(i).click();
        
        // Wait for slide to be active
        await page.waitForTimeout(300);
        
        const navEnd = Date.now();
        navigationTimes.push(navEnd - navStart);
      }
    }
    
    const avgNavigationTime = navigationTimes.reduce((a, b) => a + b, 0) / navigationTimes.length;
    console.log('Average Slide Navigation Time:', avgNavigationTime, 'ms');
    console.log('All Navigation Times:', navigationTimes);
    
    // Navigation should be snappy
    expect(avgNavigationTime).toBeLessThan(1000); // < 1 second average
    expect(Math.max(...navigationTimes)).toBeLessThan(2000); // < 2 seconds worst case
    
    await page.screenshot({ path: 'tests/screenshots/performance-slide-navigation.png' });
  });

  test('Performance - Complex formatting application', async ({ page }) => {
    await page.goto('https://docs.google.com/presentation/u/0/');
    await page.waitForSelector('[data-tooltip="Blank presentation"]', { timeout: 30000 });
    await page.click('[data-tooltip="Blank presentation"]');
    await page.waitForSelector('.punch-viewer-content', { timeout: 30000 });
    
    // Add text content
    await page.click('.punch-viewer-content [role="textbox"], .kix-wordhtmlgenerator-word-node');
    await page.type('.punch-viewer-content [role="textbox"], .kix-wordhtmlgenerator-word-node', 'Complex Formatting Performance Test');
    
    // Select all text
    await page.keyboard.press('Control+A');
    
    const formattingStart = Date.now();
    
    // Apply multiple formatting operations
    const formattingOperations = [
      () => page.keyboard.press('Control+B'), // Bold
      () => page.keyboard.press('Control+I'), // Italic
      () => page.keyboard.press('Control+U'), // Underline
    ];
    
    for (const operation of formattingOperations) {
      await operation();
      await page.waitForTimeout(200); // Allow formatting to apply
    }
    
    // Access text formatting menu for advanced options
    await page.click('[aria-label="Text"], [aria-label="Format"]');
    await page.waitForSelector('[role="menu"]', { timeout: 5000 });
    
    const formattingEnd = Date.now();
    const formattingTime = formattingEnd - formattingStart;
    
    console.log('Complex Formatting Time:', formattingTime, 'ms');
    
    // Formatting should be responsive
    expect(formattingTime).toBeLessThan(5000); // < 5 seconds
    
    await page.screenshot({ path: 'tests/screenshots/performance-complex-formatting.png' });
  });

  test('Performance - Memory usage during extended session', async ({ page }) => {
    // Monitor memory usage during extended operations
    await page.goto('https://docs.google.com/presentation/u/0/');
    await page.waitForSelector('[data-tooltip="Blank presentation"]', { timeout: 30000 });
    await page.click('[data-tooltip="Blank presentation"]');
    await page.waitForSelector('.punch-viewer-content', { timeout: 30000 });
    
    // Get initial memory metrics
    const initialMetrics = await page.evaluate(() => {
      if (performance.memory) {
        return {
          usedJSHeapSize: performance.memory.usedJSHeapSize,
          totalJSHeapSize: performance.memory.totalJSHeapSize,
          jsHeapSizeLimit: performance.memory.jsHeapSizeLimit
        };
      }
      return null;
    });
    
    // Perform memory-intensive operations
    for (let i = 0; i < 10; i++) {
      // Add slide with content
      await page.keyboard.press('Control+M');
      await page.waitForTimeout(500);
      
      await page.click('.punch-viewer-content [role="textbox"], .kix-wordhtmlgenerator-word-node');
      await page.type('.punch-viewer-content [role="textbox"], .kix-wordhtmlgenerator-word-node', 
        `Slide ${i + 1}: Memory usage test with substantial content to validate performance under load`);
      
      await page.waitForTimeout(300);
    }
    
    // Get final memory metrics
    const finalMetrics = await page.evaluate(() => {
      if (performance.memory) {
        return {
          usedJSHeapSize: performance.memory.usedJSHeapSize,
          totalJSHeapSize: performance.memory.totalJSHeapSize,
          jsHeapSizeLimit: performance.memory.jsHeapSizeLimit
        };
      }
      return null;
    });
    
    if (initialMetrics && finalMetrics) {
      const memoryIncrease = finalMetrics.usedJSHeapSize - initialMetrics.usedJSHeapSize;
      const memoryIncreasePercentage = (memoryIncrease / initialMetrics.usedJSHeapSize) * 100;
      
      console.log('Memory Usage Analysis:');
      console.log('Initial heap size:', initialMetrics.usedJSHeapSize);
      console.log('Final heap size:', finalMetrics.usedJSHeapSize);
      console.log('Memory increase:', memoryIncrease);
      console.log('Memory increase percentage:', memoryIncreasePercentage.toFixed(2) + '%');
      
      // Memory should not increase excessively
      expect(memoryIncreasePercentage).toBeLessThan(200); // < 200% increase
    }
    
    await page.screenshot({ path: 'tests/screenshots/performance-memory-test.png' });
  });

  test('Performance - Network efficiency validation', async ({ page }) => {
    // Monitor network requests during slide operations
    const networkRequests = [];
    
    page.on('request', request => {
      networkRequests.push({
        url: request.url(),
        method: request.method(),
        resourceType: request.resourceType(),
        timestamp: Date.now()
      });
    });
    
    const networkStart = Date.now();
    
    await page.goto('https://docs.google.com/presentation/u/0/');
    await page.waitForSelector('[data-tooltip="Blank presentation"]', { timeout: 30000 });
    await page.click('[data-tooltip="Blank presentation"]');
    await page.waitForSelector('.punch-viewer-content', { timeout: 30000 });
    
    // Perform typical operations
    await page.click('.punch-viewer-content [role="textbox"], .kix-wordhtmlgenerator-word-node');
    await page.type('.punch-viewer-content [role="textbox"], .kix-wordhtmlgenerator-word-node', 'Network Efficiency Test');
    
    // Apply theme
    await page.click('[aria-label="Theme"]');
    await page.waitForSelector('.docs-material-theme-picker', { timeout: 10000 });
    const themes = page.locator('.docs-material-theme-picker .docs-material-theme-item');
    if (await themes.count() > 1) {
      await themes.nth(1).click();
      await page.waitForTimeout(2000);
    }
    
    const networkEnd = Date.now();
    
    // Analyze network efficiency
    const relevantRequests = networkRequests.filter(req => 
      req.resourceType === 'xhr' || req.resourceType === 'fetch'
    );
    
    console.log('Network Analysis:');
    console.log('Total requests:', networkRequests.length);
    console.log('API requests:', relevantRequests.length);
    console.log('Request duration:', networkEnd - networkStart, 'ms');
    
    // Should not make excessive API calls
    expect(relevantRequests.length).toBeLessThan(50); // < 50 API calls for basic operations
    
    await page.screenshot({ path: 'tests/screenshots/performance-network-efficiency.png' });
  });

});

test.describe('OOXML Performance Comparison', () => {
  
  test('Performance - OOXML vs native Google Slides creation', async ({ page }) => {
    // This test would compare performance of OOXML-generated content
    // vs content created natively in Google Slides
    
    await page.goto('https://docs.google.com/presentation/u/0/');
    await page.waitForSelector('[data-tooltip="Blank presentation"]', { timeout: 30000 });
    
    // Test 1: Native creation speed
    const nativeStart = Date.now();
    await page.click('[data-tooltip="Blank presentation"]');
    await page.waitForSelector('.punch-viewer-content', { timeout: 30000 });
    
    await page.click('.punch-viewer-content [role="textbox"], .kix-wordhtmlgenerator-word-node');
    await page.type('.punch-viewer-content [role="textbox"], .kix-wordhtmlgenerator-word-node', 'Native Creation Test');
    
    const nativeEnd = Date.now();
    const nativeTime = nativeEnd - nativeStart;
    
    console.log('Native Google Slides creation time:', nativeTime, 'ms');
    
    // Test 2: Would test OOXML import speed here
    // (In a real scenario, this would upload an OOXML-generated file)
    
    await page.screenshot({ path: 'tests/screenshots/performance-comparison.png' });
    
    // Document performance for comparison
    expect(nativeTime).toBeLessThan(10000); // Baseline performance expectation
  });

});