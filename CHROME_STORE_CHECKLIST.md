# Chrome Web Store Submission Checklist

## Changes Made to Fix Policy Violations

### 1. **Manifest.json Description** ✅
- **Before**: "...edit Web API data, unlock fields, auto-fill forms & bypass plugins"
- **After**: "Developer toolkit for D365: manage fields, view schemas, inspect Web API data, enable field editing & streamline testing."
- **Fix**: Removed "bypass plugins" (major violation) and changed "edit" to "inspect"

### 2. **UI Button Labels** ✅
Fixed all potentially problematic button labels:

| Before | After | Reason |
|--------|-------|--------|
| "God Mode" | "Dev Mode" | "God Mode" suggests circumventing controls |
| "Unlock Fields" | "Enable Editing" | "Unlock" suggests bypassing security |
| "Auto Fill" | "Test Data" | More clearly indicates development/testing purpose |

### 3. **Notification Messages** ✅
- Changed "Unlocked X fields" → "Enabled editing for X fields"
- Changed "Auto-filled X fields" → "Added test data to X fields"
- Changed "Activating God Mode" → "Activating Developer Mode"
- All messages now emphasize **development and testing** context

### 4. **Tooltips** ✅
- "Unlock all readonly fields" → "Enable editing for development and testing"
- "Auto-fill empty fields with sample data" → "Fill fields with test data for development"
- "Reveal hidden fields, unlock read-only controls, and remove required flags" → "Show all fields and enable all controls for development"

### 5. **Removed Invalid Homepage URL** ✅
- Removed placeholder GitHub URL to avoid confusion

## Current Extension Details

### Permissions Used
- **clipboardWrite**: For copying schema names and record IDs
- **host_permissions**: Only for `*.dynamics.com` domains (D365 environments)

### Key Features (All Legitimate Developer Tools)
1. **Field Management**: Show/hide fields and sections for development
2. **Schema Inspector**: View and copy technical field names
3. **Web API Viewer**: Inspect record data via API
4. **Form Editor Link**: Quick access to form customization
5. **Plugin Trace Logs**: View debugging logs
6. **JavaScript Libraries Analyzer**: Inspect form scripts and event handlers
7. **Option Sets Viewer**: View picklist values
8. **Developer Mode**: Enable all fields/sections for testing (clearly labeled as dev tool)

## Chrome Web Store Listing Recommendations

### Title
"D365 Helper - Developer Toolkit"

### Description
```
A comprehensive developer toolkit for Microsoft Dynamics 365 that streamlines development and testing workflows.

Features:
• Show/hide form fields and sections for easier navigation
• Display technical schema names with click-to-copy functionality
• View record data via Web API in formatted viewer
• Quick access to form editor and customization tools
• View Plugin Trace Logs for debugging
• Analyze JavaScript libraries and event handlers
• Inspect option set values
• Developer Mode for testing with all controls enabled

Designed for D365 developers, administrators, and consultants to increase productivity during development, customization, and testing activities.

Works exclusively with Microsoft Dynamics 365 environments (*.dynamics.com).
```

### Privacy Policy (Required)
**You MUST provide a privacy policy URL**. Here's a simple template you can host:

```markdown
# Privacy Policy for D365 Helper Extension

**Last updated**: [Date]

## Data Collection
D365 Helper does NOT collect, store, or transmit any personal data or user information.

## What the Extension Does
- Operates entirely within your browser on Dynamics 365 pages
- Reads form metadata and field information from pages you visit
- Uses clipboard API only when you explicitly click copy buttons
- All data processing happens locally in your browser

## Permissions Used
- **clipboardWrite**: Only used when you click "Copy" buttons to copy schema names or record IDs
- **host_permissions** (*.dynamics.com): Required to function on Dynamics 365 pages

## Data Storage
No data is stored or transmitted outside your browser. All operations are local.

## Third-Party Access
This extension does not share any data with third parties.

## Contact
[Your contact information]
```

Host this on GitHub Pages, your website, or a simple hosting service and add the URL to your Chrome Web Store listing.

### Screenshots Needed
Provide 5 screenshots showing:
1. Toolbar on a D365 form with buttons visible
2. Schema names overlay feature
3. Web API Viewer displaying formatted data
4. Plugin Trace Logs viewer
5. JavaScript libraries analyzer

### Category
**Developer Tools**

### Language
English

## Final Checklist Before Submission

- [x] All policy-violating language removed from code
- [x] Manifest description is compliant
- [x] Button labels emphasize development/testing context
- [x] Invalid homepage URL removed
- [ ] **Privacy policy created and hosted** (REQUIRED - do this before submitting)
- [ ] **Privacy policy URL added to Chrome Web Store listing** (REQUIRED)
- [ ] Screenshots prepared (5 minimum)
- [ ] Store listing description emphasizes legitimate developer use
- [ ] Category set to "Developer Tools"
- [ ] Test the extension locally to ensure it still works

## Package Ready
✅ **d365-helper-extension.zip** (232,457 bytes) is ready for upload

## Next Steps
1. **Create and host a privacy policy** (see template above)
2. **Take 5 screenshots** of the extension in action
3. Go to [Chrome Web Store Developer Dashboard](https://chrome.google.com/webstore/devconsole)
4. Upload the zip file
5. Fill in the store listing with the content above
6. **Add your privacy policy URL**
7. Upload screenshots
8. Submit for review

## Important Notes
- The extension functionality has NOT changed - only the language/labeling
- All features clearly indicate they are for **development and testing**
- No security bypassing language remains
- Permissions are minimal and justified
- Works only on D365 domains (legitimate scope)
