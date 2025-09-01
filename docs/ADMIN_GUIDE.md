# Google Workspace Admin Guide - Enable Apps Script API Access

## Overview
Your workspace (bramalkema.nl) is blocking ALL Apps Script API operations via organizational policies. Here's how to fix it as an admin.

## Step 1: Access Google Admin Console
1. Go to **https://admin.google.com**
2. Sign in with your admin account (info@bramalkema.nl)
3. Navigate to the main dashboard

## Step 2: Enable Apps Script in Admin Console
### Path: Apps → Google Workspace → Apps Script

1. **Main Menu** → **Apps**
2. **Google Workspace** 
3. **Apps Script**
4. **Settings for Apps Script**

### Configure Apps Script Settings:
- ✅ **Allow users to create Apps Script projects**: ON
- ✅ **Allow users to run Apps Script projects**: ON  
- ✅ **Allow users to share Apps Script projects**: ON
- ✅ **Allow external apps to access Apps Script API**: ON ← **KEY SETTING**

## Step 3: Enable Drive API Access
### Path: Apps → Google Workspace → Drive and Docs

1. **Main Menu** → **Apps**
2. **Google Workspace**
3. **Drive and Docs**
4. **Features and Applications**

### Configure Drive Settings:
- ✅ **Allow users to install Google Drive apps**: ON
- ✅ **Allow Drive SDK apps**: ON ← **Important for clasp**

## Step 4: API Access Controls
### Path: Security → API Controls

1. **Main Menu** → **Security**
2. **API Controls**
3. **App access control**

### Configure API Access:
1. **Domain-owned apps** → Click **Configure**
2. **Trusted apps** → Add these OAuth Client IDs:
   ```
   1072944905499-vm2v2i5dvn0a0d2o4ca36i1vge8cvbn0.apps.googleusercontent.com
   ```
   (This is clasp's default OAuth client)

3. **Scopes** → Ensure these are allowed:
   ```
   https://www.googleapis.com/auth/script.projects
   https://www.googleapis.com/auth/script.deployments  
   https://www.googleapis.com/auth/script.webapp.deploy
   https://www.googleapis.com/auth/drive.file
   https://www.googleapis.com/auth/drive.metadata.readonly
   ```

## Step 5: Enable Google Cloud Platform APIs
### Path: Security → API Controls → Manage Google Services

1. **Security** → **API Controls**
2. **Manage Google Services**
3. Find and configure:
   - **Google Apps Script API**: Allow
   - **Google Drive API**: Allow
   - **Google Cloud Platform**: Allow

## Step 6: Service Account Settings
### Path: Security → Access and data control → API controls

1. **Security** → **Access and data control**
2. **API controls** 
3. **Domain-wide delegation**
4. **Allow users to authorize trusted OAuth2 apps**: ✅ **ON**

## Step 7: Third-Party Apps Policy
### Path: Security → Access and data control → API controls → App access control

1. Click **Configure** next to "Unconfigured"
2. Select: **Allow users to access any app (not recommended for most organizations)**
   
   OR (More Secure):
   
3. Select: **Allow users to install and run only trusted apps**
4. Add clasp OAuth client to trusted apps list

## Step 8: Verify Settings Propagation
⚠️ **Important**: Changes can take **up to 24 hours** to propagate

### Immediate Check:
1. **Admin Console** → **Reports** → **Audit and investigation** → **Apps Script**
2. Check for any blocked API calls

## Step 9: Test Configuration
After changes propagate, test with:

```bash
# Re-authenticate clasp
clasp logout
clasp login

# Test operations
clasp list-deployments
clasp push
clasp deploy -d "test-v1"
```

## Alternative Solution: Custom OAuth Client

If organizational policies still block clasp's default client:

### Create Custom OAuth Client:
1. **https://console.cloud.google.com**
2. **APIs & Services** → **Credentials**
3. **Create Credentials** → **OAuth 2.0 Client IDs**
4. **Application type**: Desktop application
5. **Name**: "Clasp CLI for bramalkema.nl"
6. Download the `client_secret.json` file

### Use Custom Client:
```bash
clasp logout
clasp login --creds ./client_secret.json
```

### Whitelist Custom Client:
1. **Admin Console** → **Security** → **API Controls**
2. **App access control** → **Trusted apps**
3. Add your new OAuth Client ID

## Troubleshooting Common Issues

### Issue: Still getting 403 errors
**Solution**: 
- Verify all settings saved correctly
- Wait 24 hours for propagation
- Check audit logs for specific blocks

### Issue: "GCP project not found"
**Solution**:
1. **Apps Script Editor** → **Resources** → **Cloud Platform Project**
2. **Change Project** → Create new or link existing
3. Enable required APIs on that project

### Issue: "Insufficient permissions"
**Solution**:
- Ensure user has **Editor** role on Apps Script project
- Check organizational units (OUs) - settings might be overridden at OU level

## Verification Checklist

After admin changes:

- [ ] clasp login works without errors
- [ ] clasp list-deployments returns data (not 403)
- [ ] clasp push succeeds
- [ ] clasp deploy creates new deployment
- [ ] Apps Script API shows as enabled in GCP console

## Security Notes

⚠️ **Before enabling broad API access**, consider:
- Create specific organizational units for developers
- Apply these settings only to developer OU
- Monitor API usage via audit logs
- Regular review of authorized applications

## Contact Support

If issues persist after following this guide:
1. **Google Workspace Support**: Mention "Apps Script API domain restrictions"
2. **Include error logs**: Specific 403 error messages from clasp
3. **Reference**: clasp GitHub issues for workspace-specific problems