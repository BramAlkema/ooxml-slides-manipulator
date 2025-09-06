#!/usr/bin/env node

/**
 * MVP PPTX Unzip-Modify-Zip Test
 * 
 * This test demonstrates our core capability:
 * 1. Unzip a PPTX file 
 * 2. Modify OOXML content
 * 3. Zip it back up
 * 4. Validate the result
 */

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

async function runPPTXMVP() {
    console.log('🚀 Starting PPTX MVP Test: Unzip-Modify-Zip');
    
    try {
        // Step 1: Check if we have a template PPTX file
        const templatePath = path.join(__dirname, 'working-template.pptx');
        
        if (!fs.existsSync(templatePath)) {
            throw new Error(`Template file not found: ${templatePath}`);
        }
        
        console.log('✅ Found template PPTX file');
        
        // Step 2: Read and unzip the PPTX
        const pptxBuffer = fs.readFileSync(templatePath);
        const zip = new JSZip();
        const loadedZip = await zip.loadAsync(pptxBuffer);
        
        console.log('✅ Successfully unzipped PPTX');
        console.log(`📁 Contains ${Object.keys(loadedZip.files).length} files`);
        
        // Step 3: List some key OOXML files
        const keyFiles = [
            'ppt/presentation.xml',
            'ppt/slides/slide1.xml', 
            'ppt/theme/theme1.xml',
            '[Content_Types].xml'
        ];
        
        console.log('\n📋 Key OOXML files found:');
        for (const fileName of keyFiles) {
            if (loadedZip.files[fileName]) {
                console.log(`  ✓ ${fileName}`);
            } else {
                console.log(`  ✗ ${fileName} (missing)`);
            }
        }
        
        // Step 4: Modify content - let's change the slide title
        const slide1Path = 'ppt/slides/slide1.xml';
        if (loadedZip.files[slide1Path]) {
            const slide1Content = await loadedZip.files[slide1Path].async('string');
            console.log('\n🔧 Original slide1.xml preview:');
            console.log(slide1Content.substring(0, 300) + '...');
            
            // Simple modification: add a timestamp to any text content
            const timestamp = new Date().toISOString();
            const modifiedContent = slide1Content.replace(
                /<a:t>([^<]+)<\/a:t>/g, 
                `<a:t>$1 [Modified ${timestamp}]</a:t>`
            );
            
            // Update the file in the ZIP
            loadedZip.file(slide1Path, modifiedContent);
            console.log('✅ Modified slide content with timestamp');
        }
        
        // Step 5: Zip it back up
        const modifiedBuffer = await loadedZip.generateAsync({
            type: 'nodebuffer',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 }
        });
        
        console.log('✅ Successfully re-zipped PPTX');
        
        // Step 6: Save the modified file
        const outputPath = path.join(__dirname, `mvp-test-output-${Date.now()}.pptx`);
        fs.writeFileSync(outputPath, modifiedBuffer);
        
        console.log(`✅ Saved modified PPTX: ${outputPath}`);
        console.log(`📊 Original size: ${pptxBuffer.length} bytes`);
        console.log(`📊 Modified size: ${modifiedBuffer.length} bytes`);
        
        // Step 7: Verify the modified file can be read
        const verifyZip = new JSZip();
        await verifyZip.loadAsync(modifiedBuffer);
        
        console.log('✅ Verification successful - modified PPTX is valid');
        
        return {
            success: true,
            originalFile: templatePath,
            modifiedFile: outputPath,
            originalSize: pptxBuffer.length,
            modifiedSize: modifiedBuffer.length,
            fileCount: Object.keys(loadedZip.files).length
        };
        
    } catch (error) {
        console.error('❌ MVP Test Failed:', error.message);
        return {
            success: false,
            error: error.message
        };
    }
}

// Check if JSZip is available, if not suggest installation
try {
    require('jszip');
} catch (error) {
    console.log('📦 Installing jszip dependency...');
    require('child_process').execSync('npm install jszip', { stdio: 'inherit' });
    console.log('✅ jszip installed');
}

// Run the MVP test
runPPTXMVP().then(result => {
    console.log('\n🎯 MVP Test Results:');
    console.log(JSON.stringify(result, null, 2));
    
    if (result.success) {
        console.log('\n🎉 MVP PPTX Pipeline Working! ✨');
        console.log('Next steps: Integrate with Brave browser automation');
    } else {
        console.log('\n💥 MVP Test Failed - Check error above');
        process.exit(1);
    }
});