# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an OOXML PPTX Processing Platform that manipulates PowerPoint files through direct OOXML editing. The core workflow is:
1. **Unzip** PPTX files to extract OOXML structure
2. **Modify** themes, colors, fonts, slides as JSON/XML
3. **Zip** back to valid PPTX files

The platform is deployed entirely from Google Apps Script to Google Cloud Run (free tier), without requiring CLI tools.

## Key Architecture Concepts

### Deployment Architecture
- **Primary deployment method**: Google Apps Script → Google Cloud Run (no CLI required)
- **Region**: us-central1 (US free tier, 512MB memory)
- **Service**: Cloud Run with fflate for ZIP operations
- **Data transfer**: Base64 encoding for binary data between GAS and Cloud Run

### Core Services
1. **OOXMLJsonService** (`lib/OOXMLJsonService.js`): Cloud Run service for OOXML↔JSON conversion
2. **OOXMLDeployment** (`src/OOXMLDeployment.js`): Automated Cloud Run deployment from GAS
3. **FFlatePPTXService** (`lib/FFlatePPTXService.js`): ZIP/unzip operations using fflate library
4. **ExtensionFramework** (`lib/ExtensionFramework.js`): Modular extension system for brand compliance

### Critical Implementation Details
- All PPTX processing happens server-side on Cloud Run, not in GAS
- Use `DeployFromGAS.js` for deployment - users copy this to their GAS project
- The workflow test files demonstrate the unzip→modify→zip pattern
- Brave browser is configured for Playwright testing via `executablePath`

## Common Development Commands

### Testing
```bash
# Run Playwright tests with Brave browser
npm test

# Test core PPTX processing MVP
npm run test:mvp

# Test cloud deployment integration
npm run test:cloud

# Test deployment configuration
npm run test:deployment

# Run specific integration tests
npm run test:integration
npm run test:brave
```

### Google Apps Script (Clasp)
```bash
# Login to Google Apps Script
npm run clasp:login

# Push changes to GAS
npm run clasp:push

# Pull changes from GAS
npm run clasp:pull

# Open GAS project in browser
npm run clasp:open

# Deploy to GAS
npm run clasp:deploy
```

### Installation
```bash
# Install dependencies and Playwright browsers
npm install
```

## Testing Approach

1. **MVP Tests** (`test-pptx-mvp.js`): Validates core unzip/modify/zip workflow locally
2. **Cloud Tests** (`test-cloud-pptx-mvp.js`): Tests Cloud Run deployment
3. **Playwright Tests**: Browser automation tests using Brave browser exclusively
4. **Integration Tests**: End-to-end GAS→Cloud Run workflow validation

## Deployment Workflow

When users ask about deployment:
1. Direct them to use `DeployFromGAS.js` - copy to GAS project
2. Run `setupProjectId()` to configure GCP project
3. Run `showPreflightChecks()` for setup validation
4. Run `deployToUSFreeTier()` for automated deployment
5. Run `testDeployedService()` to verify

Never suggest CLI deployment unless specifically requested - the platform is designed for GAS-based deployment.

## File Structure Notes

- `src/`: GAS source files and handlers
- `lib/`: Core libraries including OOXML services
- `test/`: Playwright and integration tests
- `examples/`: Usage examples and demos
- `DeployFromGAS.js`: User-facing deployment script

## Important Patterns

1. **Error Handling**: Check for deployment status before operations
2. **Free Tier Limits**: 512MB memory, 60s timeout in us-central1
3. **Binary Data**: Always use Base64 encoding for PPTX data transfer
4. **Testing**: Brave browser is the default for all Playwright tests

## GitHub Workflow Configuration

The GitHub Actions workflow (`.github/workflows/test.yml`) is configured to:
- Run only on push/PR (no daily schedule)
- Test with Brave browser exclusively
- Validate file structure for current architecture
- Check documentation completeness

## When Making Changes

1. Maintain the GAS-first deployment approach
2. Test unzip→modify→zip workflow for any PPTX changes
3. Update `DeployFromGAS.js` if deployment process changes
4. Ensure Brave browser compatibility for browser automation
5. Keep us-central1 as default region for free tier optimization