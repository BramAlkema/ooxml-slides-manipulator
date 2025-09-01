# OAuth Client Project Issue - Root Cause Analysis

## Problem Summary
Google Apps Script API calls fail with project mismatch despite correct configuration:
- **Configured Project**: sys-06010664847399265559325060 (OOXML Slides Manipulator)
- **API Consumer Project**: 32555940559 (causing SERVICE_DISABLED errors)

## Root Cause Discovery

### Investigation Steps
```bash
# Check token details
curl -s "https://www.googleapis.com/oauth2/v1/tokeninfo?access_token=$(gcloud auth print-access-token)" 
```

### Key Finding
```json
{
  "issued_to": "32555940559.apps.googleusercontent.com",
  "audience": "32555940559.apps.googleusercontent.com",
  "scope": "https://www.googleapis.com/auth/cloud-platform ..."
}
```

## Root Cause Identified ✅

**The gcloud authentication is using an OAuth client from project 32555940559**, which is likely Google's internal/default project for gcloud CLI.

### Why This Causes Issues:
1. **API Billing**: Google routes API calls to the OAuth client's project (32555940559)
2. **Service Enablement**: APIs must be enabled on the OAuth client's project, not our target project
3. **Project Override Ignored**: Even with quota-project and other settings, the OAuth client project takes precedence

### Configuration vs Reality:
- ✅ **gcloud project**: sys-06010664847399265559325060  
- ✅ **Apps Script GCP association**: 149429464972 (correct project number)
- ✅ **Quota project**: sys-06010664847399265559325060
- ❌ **OAuth Client**: 32555940559.apps.googleusercontent.com (WRONG PROJECT)

## Current Working Status

### ✅ Working Components:
- **clasp deployment**: `clasp push`, `clasp deploy` work perfectly
- **Web app access**: External calls work with `"access": "ANYONE_ANONYMOUS"`
- **Manual execution**: GAS editor execution works perfectly
- **Core functionality**: Universal OOXML manipulator fully functional

### ❌ Blocked Components:
- **Apps Script Execution API**: `https://script.googleapis.com/v1/scripts/{id}:run`
- **clasp run**: Same OAuth client issue
- **Programmatic function calls**: All blocked by wrong OAuth client

## Solution Path

### Option 1: Custom OAuth Client (Recommended)
1. **Create OAuth credentials** in sys-06010664847399265559325060
2. **Download client_secret.json**
3. **Re-authenticate gcloud** with custom client:
   ```bash
   gcloud auth login --cred-file=client_secret.json
   ```

### Option 2: Service Account with Domain Delegation
1. **Create service account** in sys-06010664847399265559325060
2. **Enable domain-wide delegation**
3. **Use service account** for API calls

### Option 3: clasp Custom OAuth
1. **Create OAuth client** for Apps Script access
2. **Use clasp with custom credentials**:
   ```bash
   clasp logout
   clasp login --creds ./client_secret.json
   ```

## Technical Details

### Error Pattern:
```json
{
  "error": {
    "code": 403,
    "status": "PERMISSION_DENIED",
    "details": [{
      "reason": "SERVICE_DISABLED",
      "metadata": {
        "consumer": "projects/32555940559",  // Wrong project!
        "service": "script.googleapis.com"
      }
    }]
  }
}
```

### API Call Flow:
1. **gcloud auth print-access-token** → Token from OAuth client 32555940559
2. **curl with Bearer token** → Google sees client 32555940559
3. **script.googleapis.com** → Routes to project 32555940559
4. **SERVICE_DISABLED** → APIs not enabled on project 32555940559

## Impact Assessment

### High Priority (Working) ✅:
- Universal OOXML core development and testing
- Automated deployment pipeline
- Web app external integration
- Manual validation workflows

### Medium Priority (Blocked) ❌:
- Automated testing via API
- CI/CD integration
- Programmatic function execution
- External automation scripts

## Next Actions Required

1. **Create OAuth Client**:
   - Go to Google Cloud Console → APIs & Services → Credentials
   - Create OAuth 2.0 Client ID (Desktop application)
   - Download JSON credentials

2. **Re-authenticate**:
   ```bash
   gcloud auth login --cred-file=./oauth_client.json
   ```

3. **Test API Access**:
   ```bash
   curl -X POST \
     -H "Authorization: Bearer $(gcloud auth print-access-token)" \
     -H "Content-Type: application/json" \
     -d '{"function":"ping","parameters":["CustomOAuth"],"devMode":false}' \
     "https://script.googleapis.com/v1/scripts/{SCRIPT_ID}:run"
   ```

## Lessons Learned

### Key Insight:
**OAuth client project determines API billing/routing**, not the target project or quota project settings.

### Configuration Hierarchy:
1. **OAuth Client Project** (highest priority - determines API routing)
2. **Quota Project** (billing override - but doesn't work if APIs disabled on OAuth project)
3. **Target Resource Project** (lowest priority - where resources live)

### Authentication vs Authorization:
- **Authentication**: Which OAuth client (determines project context)
- **Authorization**: What permissions (determined by scopes and IAM)

The issue was **authentication-level project mismatch**, not authorization or configuration.

---

**Date**: 2025-08-13  
**Status**: Root cause identified, solution path defined  
**Priority**: Medium (doesn't block core functionality)