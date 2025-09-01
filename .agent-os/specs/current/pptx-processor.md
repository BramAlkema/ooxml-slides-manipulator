# PPTX Processor Specification

## Overview
Complete specification for the OOXML PPTX processing platform with unzip/modify/zip capabilities.

## Core Requirements

### Functional Requirements

#### FR1: PPTX Processing
- **FR1.1**: Unzip PPTX files to extract OOXML structure
- **FR1.2**: Parse XML files into editable JSON format
- **FR1.3**: Modify themes, colors, fonts, and content
- **FR1.4**: Reconstruct valid PPTX files from modified content
- **FR1.5**: Validate PPTX integrity after modifications

#### FR2: Cloud Deployment
- **FR2.1**: Deploy from Google Apps Script without CLI
- **FR2.2**: Run on Google Cloud Run free tier
- **FR2.3**: Use us-central1 region for free tier eligibility
- **FR2.4**: Auto-enable required Google Cloud APIs
- **FR2.5**: Handle billing verification and setup

#### FR3: API Operations
- **FR3.1**: Provide unwrap() method for PPTX extraction
- **FR3.2**: Provide rewrap() method for PPTX reconstruction
- **FR3.3**: Support server-side text replacement
- **FR3.4**: Enable batch processing operations
- **FR3.5**: Handle large files via session management

### Non-Functional Requirements

#### NFR1: Performance
- **NFR1.1**: Process standard PPTX (< 10MB) in < 10 seconds
- **NFR1.2**: Support files up to 100MB via sessions
- **NFR1.3**: Handle 3 concurrent requests (free tier limit)
- **NFR1.4**: 60-second timeout per request
- **NFR1.5**: 512MB memory per instance

#### NFR2: Reliability
- **NFR2.1**: 99.9% uptime within free tier limits
- **NFR2.2**: Graceful error handling with clear messages
- **NFR2.3**: Automatic retry for transient failures
- **NFR2.4**: Data integrity validation after processing
- **NFR2.5**: Rollback capability for failed operations

#### NFR3: Security
- **NFR3.1**: OAuth2 authentication for GAS integration
- **NFR3.2**: Optional authentication for Cloud Run service
- **NFR3.3**: No permanent storage of user data
- **NFR3.4**: Secure base64 encoding for data transfer
- **NFR3.5**: HTTPS-only communication

#### NFR4: Usability
- **NFR4.1**: 5-minute quickstart from zero to deployed
- **NFR4.2**: No command line knowledge required
- **NFR4.3**: Clear error messages and troubleshooting
- **NFR4.4**: Comprehensive examples and documentation
- **NFR4.5**: Visual feedback in GAS sidebar during setup

## Technical Architecture

### System Components

```
┌─────────────────────────────────────────────────────────┐
│                   Google Apps Script                     │
│  ┌─────────────────────────────────────────────────┐   │
│  │          OOXMLDeployment Service                 │   │
│  │  - showGcpPreflight()                           │   │
│  │  - initAndDeploy()                              │   │
│  │  - getDeploymentStatus()                        │   │
│  └─────────────────────────────────────────────────┘   │
│  ┌─────────────────────────────────────────────────┐   │
│  │           OOXMLJsonService                       │   │
│  │  - unwrap(fileId)                               │   │
│  │  - rewrap(manifest)                             │   │
│  │  - process(fileId, operations)                  │   │
│  └─────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────┘
                           │
                    Base64 PPTX Data
                           │
┌─────────────────────────────────────────────────────────┐
│                 Google Cloud Run Service                 │
│  ┌─────────────────────────────────────────────────┐   │
│  │              Express.js Server                   │   │
│  │  POST /extract - Unzip PPTX                     │   │
│  │  POST /rebuild - Zip PPTX                       │   │
│  │  POST /process - Server-side operations         │   │
│  │  GET /health - Health check                     │   │
│  └─────────────────────────────────────────────────┘   │
│  ┌─────────────────────────────────────────────────┐   │
│  │               fflate Library                     │   │
│  │  - unzipSync() - Extract files                  │   │
│  │  - zipSync() - Compress files                   │   │
│  └─────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────┘
```

### Data Flow

1. **Upload**: User selects PPTX from Google Drive
2. **Encode**: GAS converts to base64
3. **Send**: POST to Cloud Run service
4. **Process**: Unzip → Modify → Zip
5. **Return**: Base64 response to GAS
6. **Save**: Write to Google Drive

### API Endpoints

#### /extract
```javascript
POST /extract
{
  "zipB64": "base64_encoded_pptx"
}

Response:
{
  "entries": [
    {
      "path": "ppt/slides/slide1.xml",
      "type": "xml",
      "text": "<?xml version=\"1.0\"?>..."
    }
  ]
}
```

#### /rebuild
```javascript
POST /rebuild
{
  "manifest": {
    "zipB64": "original_base64",
    "entries": [/* modified entries */]
  }
}

Response:
{
  "zipB64": "modified_pptx_base64"
}
```

#### /process
```javascript
POST /process
{
  "zipB64": "base64_pptx",
  "ops": [
    {"type": "replaceText", "find": "old", "replace": "new"}
  ]
}

Response:
{
  "zipB64": "modified_base64",
  "report": {
    "replaced": 5,
    "errors": []
  }
}
```

## Implementation Plan

### Phase 1: Core Infrastructure ✅
- [x] GAS deployment system
- [x] Cloud Run service template
- [x] fflate integration
- [x] Basic unzip/zip functionality

### Phase 2: API Development ✅
- [x] OOXMLJsonService implementation
- [x] Extract/rebuild endpoints
- [x] Server-side operations
- [x] Error handling

### Phase 3: Testing & Documentation ✅
- [x] Playwright test suite
- [x] Brave browser integration
- [x] README and QUICKSTART
- [x] DEPLOYMENT guide

### Phase 4: Optimization (Current)
- [ ] Performance tuning
- [ ] Caching strategy
- [ ] Monitoring setup
- [ ] Usage analytics

### Phase 5: Advanced Features (Future)
- [ ] Template library
- [ ] AI-powered modifications
- [ ] Batch processing UI
- [ ] Version control integration

## Testing Strategy

### Unit Tests
- fflate operations
- XML parsing
- JSON manipulation
- Base64 encoding/decoding

### Integration Tests
- GAS to Cloud Run communication
- File upload/download
- Error scenarios
- Timeout handling

### End-to-End Tests
- Complete workflow validation
- Browser automation with Brave
- Performance benchmarks
- Free tier limit validation

### Test Commands
```bash
npm test:mvp        # Basic PPTX processing
npm test:cloud      # Cloud service simulation
npm test:deployment # Deployment validation
npm test:integration # Full integration tests
```

## Deployment Checklist

### Pre-Deployment
- [ ] Google Cloud project created
- [ ] Billing enabled (free tier)
- [ ] Project ID configured in GAS
- [ ] APIs enabled (Cloud Run, Build)

### Deployment Steps
1. Run `setupProjectId()`
2. Run `showPreflightChecks()`
3. Complete sidebar verification
4. Run `deployToUSFreeTier()`
5. Run `testDeployedService()`

### Post-Deployment
- [ ] Service URL saved
- [ ] Health check passing
- [ ] Test PPTX processed
- [ ] Monitoring configured

## Success Criteria

1. **Deployment Time**: < 5 minutes from start to finish
2. **Processing Speed**: < 10 seconds for 10MB PPTX
3. **Cost**: $0.00 within free tier limits
4. **Reliability**: 99.9% success rate
5. **User Satisfaction**: Clear docs, easy setup

## Risk Mitigation

### Technical Risks
- **Free tier exhaustion**: Monitor usage, implement quotas
- **Large file handling**: Use session management for > 10MB
- **API changes**: Version lock dependencies

### Operational Risks
- **Service downtime**: Health checks, auto-restart
- **Data loss**: Validation before/after processing
- **Security breach**: Authentication, encryption

## Future Enhancements

1. **AI Integration**: Smart content suggestions
2. **Template Marketplace**: Share/sell templates
3. **Collaboration**: Multi-user editing
4. **Version Control**: Track presentation changes
5. **Analytics**: Usage insights and optimization