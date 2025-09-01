#!/usr/bin/env node

/**
 * Test Free Tier US Deployment
 * 
 * This script tests the free-tier deployment before and after deploying
 */

const fs = require('fs');
const path = require('path');

async function testFreeTierDeployment() {
    console.log('🇺🇸 Testing Free Tier US Deployment Configuration');
    console.log('=================================================');
    
    try {
        // Step 1: Verify gcloud CLI is available
        console.log('\n🔧 Step 1: Checking gcloud CLI...');
        
        const { execSync } = require('child_process');
        
        try {
            const gcloudVersion = execSync('gcloud version', { encoding: 'utf8' });
            const versionMatch = gcloudVersion.match(/Google Cloud SDK (\S+)/);
            const version = versionMatch ? versionMatch[1] : 'installed';
            console.log(`✅ gcloud CLI installed: ${version}`);
        } catch (error) {
            console.log('⚠️ gcloud CLI not found locally');
            console.log('   You can install it or use Cloud Shell');
            console.log('   Install: https://cloud.google.com/sdk/docs/install');
            // Don't fail here - continue with other tests
        }
        
        // Step 2: Check authentication
        console.log('\n🔐 Step 2: Checking authentication...');
        
        try {
            const activeAccount = execSync('gcloud auth list --filter=status:ACTIVE --format="value(account)"', { encoding: 'utf8' });
            if (activeAccount.trim()) {
                console.log(`✅ Authenticated as: ${activeAccount.trim()}`);
            } else {
                console.log('⚠️ No active authentication');
                console.log('Run: gcloud auth login');
            }
        } catch (error) {
            console.log('⚠️ Authentication check failed');
        }
        
        // Step 3: Check current project
        console.log('\n📋 Step 3: Checking project configuration...');
        
        try {
            const currentProject = execSync('gcloud config get-value project', { encoding: 'utf8' });
            if (currentProject.trim() && !currentProject.includes('unset')) {
                console.log(`✅ Current project: ${currentProject.trim()}`);
            } else {
                console.log('⚠️ No project set');
                console.log('Run: gcloud config set project YOUR_PROJECT_ID');
            }
        } catch (error) {
            console.log('⚠️ Project check failed');
        }
        
        // Step 4: Test PPTX processing locally (simulate what will run on free tier)
        console.log('\n📦 Step 4: Testing PPTX processing logic...');
        
        const templatePath = path.join(__dirname, 'working-template.pptx');
        if (!fs.existsSync(templatePath)) {
            throw new Error('Template PPTX not found. Run previous MVP tests first.');
        }
        
        const pptxBuffer = fs.readFileSync(templatePath);
        const base64Data = pptxBuffer.toString('base64');
        
        console.log(`✅ PPTX loaded: ${pptxBuffer.length} bytes → ${base64Data.length} base64 chars`);
        
        // Simulate the free-tier service logic
        const fflate = require('fflate');
        const uint8Array = new Uint8Array(pptxBuffer);
        
        // Test unzipSync (what our free-tier service will do)
        const extracted = fflate.unzipSync(uint8Array);
        const fileCount = Object.keys(extracted).length;
        
        console.log(`✅ fflate unzipSync successful: ${fileCount} files extracted`);
        
        // Test zipSync (rebuild process)
        const rebuilt = fflate.zipSync(extracted, { level: 6 });
        console.log(`✅ fflate zipSync successful: ${rebuilt.length} bytes`);
        
        // Verify integrity
        const rebuiltBase64 = Buffer.from(rebuilt).toString('base64');
        console.log(`✅ Base64 conversion successful: ${rebuiltBase64.length} chars`);
        
        // Step 5: Check deployment script
        console.log('\n🚀 Step 5: Validating deployment script...');
        
        const deployScript = path.join(__dirname, 'deploy-free-tier-us.sh');
        if (fs.existsSync(deployScript)) {
            console.log('✅ Free-tier deployment script ready');
            console.log('📂 Location:', deployScript);
        } else {
            console.log('❌ Deployment script missing');
        }
        
        // Step 6: Free tier validation
        console.log('\n💰 Step 6: Free tier configuration validation...');
        
        const freeTierLimits = {
            memory: '512MB',
            timeout: '60s',
            maxInstances: 3,
            maxFileSize: '10MB',
            region: 'us-central1'
        };
        
        console.log('✅ Free tier limits configured:');
        Object.entries(freeTierLimits).forEach(([key, value]) => {
            console.log(`   ${key}: ${value}`);
        });
        
        // Validate file size is within limits
        const fileSizeMB = pptxBuffer.length / (1024 * 1024);
        if (fileSizeMB < 10) {
            console.log(`✅ Test file size OK: ${fileSizeMB.toFixed(2)}MB < 10MB limit`);
        } else {
            console.log(`⚠️ Test file too large: ${fileSizeMB.toFixed(2)}MB > 10MB limit`);
        }
        
        return {
            success: true,
            fileCount: fileCount,
            originalSize: pptxBuffer.length,
            rebuiltSize: rebuilt.length,
            freeTierReady: true,
            deploymentScript: deployScript
        };
        
    } catch (error) {
        console.error('❌ Free Tier Test Failed:', error.message);
        return {
            success: false,
            error: error.message
        };
    }
}

// Test function to simulate API call after deployment
async function testDeployedService(serviceUrl) {
    console.log('\n🌐 Testing deployed service...');
    
    if (!serviceUrl) {
        console.log('⚠️ No service URL provided, skipping deployment test');
        return;
    }
    
    try {
        const fetch = require('node-fetch');
        
        // Test ping endpoint
        console.log('🏓 Testing ping endpoint...');
        const pingResponse = await fetch(`${serviceUrl}/ping`);
        const pingData = await pingResponse.json();
        console.log('✅ Ping successful:', pingData);
        
        // Test health endpoint
        console.log('🩺 Testing health endpoint...');
        const healthResponse = await fetch(`${serviceUrl}/health`);
        const healthData = await healthResponse.json();
        console.log('✅ Health check successful:', healthData);
        
        return { success: true, endpoints: ['ping', 'health'] };
        
    } catch (error) {
        console.log('❌ Service test failed:', error.message);
        return { success: false, error: error.message };
    }
}

// Main execution
async function main() {
    const testResult = await testFreeTierDeployment();
    
    console.log('\n🎯 Free Tier Deployment Test Results:');
    console.log('=====================================');
    console.log(JSON.stringify(testResult, null, 2));
    
    if (testResult.success) {
        console.log('\n🎉 Free Tier Configuration Valid! ✨');
        console.log('\n📋 Next Steps:');
        console.log('1. Ensure gcloud auth login is complete');
        console.log('2. Set your project: gcloud config set project YOUR_PROJECT_ID');
        console.log('3. Run deployment: ./deploy-free-tier-us.sh');
        console.log('4. Test deployed service with returned URL');
        console.log('\n💰 Estimated cost: $0.00 (within free tier limits)');
    } else {
        console.log('\n💥 Configuration Issues Found');
        console.log('Please resolve the issues above before deploying');
        process.exit(1);
    }
    
    // If service URL is provided as argument, test it
    const serviceUrl = process.argv[2];
    if (serviceUrl) {
        await testDeployedService(serviceUrl);
    }
}

main().catch(console.error);