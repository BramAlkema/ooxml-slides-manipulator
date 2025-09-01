# PPTX Processor Project-Specific Instructions

## Project Context

This is a Google Apps Script-based platform for processing PowerPoint files through OOXML manipulation. The core workflow is: unzip PPTX → modify XML/JSON → zip back to PPTX.

## Key Principles

1. **GAS-First Development**: Everything deploys and runs from Google Apps Script
2. **Free Tier Focus**: Optimize for Google Cloud free tier (us-central1)
3. **No CLI Required**: Users should never need command line access
4. **5-Minute Setup**: From zero to deployed in under 5 minutes

## Development Guidelines

### When Adding Features

1. **Check GAS Compatibility**: Ensure new code works in V8 runtime
2. **Maintain Deployment Flow**: Don't break the one-click deployment
3. **Document in Code**: Add clear comments for GAS users
4. **Test with Brave**: Use Playwright with Brave browser configuration

### When Modifying Core Systems

1. **Preserve Unzip/Zip Flow**: Core pipeline must remain intact
2. **Keep Base64 Encoding**: GAS requires base64 for binary data
3. **Respect Free Tier Limits**: 512MB memory, 60s timeout
4. **Update Agent OS Specs**: Keep .agent-os/ files current

## File Organization

```
/src/               # Google Apps Script source files
  Main.js          # Entry point for Web App
  OOXMLDeployment.js # Cloud deployment from GAS

/lib/               # Core libraries
  OOXMLJsonService.js # Main API for PPTX processing
  FFlatePPTXService.js # Local PPTX operations

/examples/          # Usage examples
  OOXMLJsonQuickStart.js # Complete examples

/test/              # Playwright tests
  *-integration.spec.js # Brave browser tests

/.agent-os/         # Agent OS configuration
  /product/        # Product definition
  /specs/          # Technical specifications
  /instructions/   # Development guidelines
```

## Key Commands

```bash
# Core testing
npm test:mvp        # Test basic PPTX processing
npm test:cloud      # Test cloud simulation
npm test:deployment # Validate deployment readiness

# Integration testing
npm test:brave      # Brave browser tests
npm test:integration # GAS integration

# Development
clasp push          # Deploy to Google Apps Script
clasp open          # Open in GAS editor
```

## Core Workflow

### Standard Development Flow
1. Code in local IDE
2. Test with `npm test:mvp`
3. Push to GAS with `clasp push`
4. Test deployment with GAS functions
5. Validate with Brave browser tests

### Deployment Flow (GAS)
```javascript
setupProjectId()       // Configure project
showPreflightChecks()  // Verify setup
deployToUSFreeTier()   // Deploy service
testDeployedService()  // Validate
```

## Critical Files

- `src/OOXMLDeployment.js` - Cloud deployment system
- `lib/OOXMLJsonService.js` - Core PPTX processing
- `DeployFromGAS.js` - User deployment script
- `QUICKSTART.md` - User onboarding