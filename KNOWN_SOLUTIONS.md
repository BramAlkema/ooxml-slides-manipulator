# Known Solutions - Google Slides Integration

## Screenshot/Export Issues

### Problem: Cannot screenshot or export presentations
**Symptoms:**
- `/pub` URLs return "This document is not published" 
- `/export?format=png` returns HTTP 401 Unauthorized
- Playwright/Selenium tests fail with authentication screens

**Root Cause:** 
Google Slides requires proper sharing permissions for public access and export functionality.

**Solutions:**

1. **Set proper Drive API permissions and publish revisions (CORRECT METHOD):**
```javascript
// Step 1: Create proper Drive API permission for "anyoneWithLink"
var permission = Drive.Permissions.create({
  "role": "reader",
  "type": "anyone", 
  "allowFileDiscovery": false
}, presentationId);

// Step 2: Set the head revision to published using Drive API  
var revisions = Drive.Revisions.list(presentationId);
var items = revisions.items;
var headRevisionId = items[items.length - 1].id;

var publishedRevision = Drive.Revisions.update({
  "published": true,
  "publishAuto": true, 
  "publishedOutsideDomain": true
}, presentationId, headRevisionId);
```

**API References:**
- [Drive Permissions API](https://developers.google.com/drive/api/v3/reference/permissions/create)  
- [Drive Revisions API](https://developers.google.com/drive/api/v3/reference/revisions)

2. **Use export URLs with proper sharing:**
```bash
# These work ONLY if sharing is set to "Anyone with link can view"
https://docs.google.com/presentation/d/{ID}/export?format=png
https://docs.google.com/presentation/d/{ID}/export?format=svg
https://docs.google.com/presentation/d/{ID}/export?format=pdf
```

3. **Alternative screenshot approaches:**
```python
# Method 1: Direct export (requires proper sharing)
png_url = f"https://docs.google.com/presentation/d/{id}/export?format=png"

# Method 2: Published presentation (requires setSharing)
pub_url = f"https://docs.google.com/presentation/d/{id}/pub"

# Method 3: Embedded view (sometimes works without auth)
embed_url = f"https://docs.google.com/presentation/d/{id}/embed"
```

**Implementation Notes:**
- Always call `file.setSharing()` immediately after `SlidesApp.create()`
- Check sharing permissions in Drive if exports still fail
- Export URLs return raw image data, not HTML pages
- SVG exports maintain vector quality and text selectability

**Known Issues:**
- Export URLs may still return HTTP 401 even with proper sharing set via API
- Sharing permissions may take time to propagate (try waiting 30+ seconds)
- Some Google accounts/domains may have additional restrictions on programmatic exports
- Alternative: Use manual screenshot tools or browser automation for reliable capture

**Working Alternatives:**
1. Manual screenshot with `screencapture` (macOS) or browser tools
2. Browser automation targeting `/pub` URLs (if sharing is properly set)
3. Use Google Apps Script's built-in Drive API to create images within the script

## Font Application Issues

### Problem: Enhanced font functions not returning parsed data
**Symptoms:**
- API returns basic response without `fontPair` or `colorPalette` data
- Font parsing functions exist but aren't being called

**Root Cause:** 
JavaScript errors in enhanced function cause fallback to basic implementation.

**Solutions:**
1. Add error handling around font parsing
2. Use debug functions to isolate parsing vs. application issues
3. Check Google Apps Script execution logs for runtime errors

## Authentication Issues

### Problem: Playwright/Selenium tests hit sign-in screens
**Root Cause:** 
Accessing edit URLs or improperly shared presentations requires authentication.

**Solutions:**
1. Use public/published URLs only: `/pub` not `/edit`
2. Ensure presentations have proper sharing permissions
3. Use export URLs for automated screenshot capture
4. Test with incognito/private browsing to verify public access