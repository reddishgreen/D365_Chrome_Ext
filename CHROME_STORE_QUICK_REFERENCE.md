# Chrome Web Store Submission - Quick Reference

Use this as a quick copy-paste reference while filling out the Chrome Web Store Developer Dashboard.

---

## Account Tab

**Contact Email**: [Enter your email]

⚠️ **IMPORTANT**: Verify your email before continuing (check inbox)

---

## Store Listing Tab

### Basic Info

**Name**: D365 Helper - Developer Toolkit

**Summary** (132 char max):
```
Developer toolkit for D365: manage fields, view schemas, inspect Web API data, enable field editing & streamline testing.
```

**Category**: Developer Tools

**Language**: English (United States)

---

## Privacy Practices Tab

### 1. Single Purpose Description

```
This extension provides developer productivity tools exclusively for Microsoft Dynamics 365, including field management, schema name display, Web API data inspection, and developer testing features.
```

### 2. Permissions Justifications

**activeTab:**
```
The activeTab permission is required to inject the developer toolbar into active Dynamics 365 form pages. When users click the extension icon or use toolbar features, we need access to the current tab to:
- Display the floating toolbar interface
- Read form metadata and field schema names
- Show/hide form fields and sections
- Access the Dynamics 365 Client API (Xrm) to interact with form controls
This permission is only used when the user actively engages with the extension on a D365 page.
```

**clipboardWrite:**
```
The clipboardWrite permission enables users to quickly copy developer information to their clipboard with one click. This is used when:
- Copying individual field schema names
- Copying all schema names at once
- Copying record IDs
- Copying Web API data values
All clipboard operations are triggered explicitly by the user clicking a "Copy" button. No data is copied automatically or in the background.
```

**storage:**
```
The storage permission saves user preferences locally in the browser to improve user experience:
- Remembering toolbar state (minimized/expanded)
- Saving user's preferred toolbar position
- Storing display preferences
All data is stored locally on the user's device only. No data is sent to external servers or shared with third parties.
```

**Host Permissions (*.dynamics.com, *.crm*.dynamics.com):**
```
Host permissions for Dynamics 365 domains are essential for the extension to function. These permissions allow the extension to:
- Run exclusively on Microsoft Dynamics 365 instances (dynamics.com and regional CRM domains)
- Access form metadata and field information using the D365 Client API
- Interact with Web API endpoints for data viewing and editing
- Inject the developer toolbar interface onto D365 pages
The extension is specifically designed for D365 developers and only activates on official Microsoft Dynamics 365 URLs. It does not run on any other websites.
```

### 3. Remote Code

**Does your extension execute remote code?**: No

**Explanation:**
```
This extension does not fetch, execute, or load any code from remote servers. All JavaScript code is bundled within the extension package at build time using Webpack. The extension only interacts with the user's own Dynamics 365 environment's Web API for data viewing and editing purposes.
```

### 4. Data Usage

**Handles personal or sensitive data?**: No

**Certification:**
```
This extension does not collect, store, transmit, or process any personal or sensitive user data. All operations are performed locally in the user's browser. While the extension can view and edit data within the user's Dynamics 365 environment, this data never leaves the user's browser, is not sent to any external servers, and is not stored by the extension (except user preferences).
```

**Check all certifications**:
- [x] Complies with Chrome Web Store Developer Program Policies
- [x] Provided accurate information about data collection and usage
- [x] All permissions are necessary for core functionality
- [x] Will not sell user data to third parties
- [x] Will not use or transfer data for unrelated purposes
- [x] Will not use data for creditworthiness or lending

### 5. Privacy Policy

**Option 1 - Paste directly:**
Copy and paste contents from `PRIVACY_POLICY.md`

**Option 2 - Provide URL:**
```
https://github.com/[yourusername]/d365-helper/blob/main/PRIVACY_POLICY.md
```

---

## Distribution Tab

**Visibility**: Public

**Countries**: All countries

---

## Upload Package

**File**: d365-helper-v1.0.0.zip

---

## Checklist Before Submit

- [ ] Contact email entered and verified
- [ ] Single purpose description provided
- [ ] All 4 permission justifications entered (activeTab, clipboardWrite, storage, host permissions)
- [ ] Remote code answered (No)
- [ ] Data usage certification completed
- [ ] All certifications checked
- [ ] Privacy policy provided
- [ ] Package uploaded (d365-helper-v1.0.0.zip)
- [ ] At least 1 screenshot uploaded (1280x800 or 640x400)
- [ ] Small promo tile uploaded (440x280)
- [ ] Distribution settings configured

---

## After Submission

- Extension review typically takes 1-3 business days
- You'll receive email notification
- Monitor your Chrome Web Store developer dashboard
- Check spam folder for emails from Chrome Web Store

---

## Need Help?

See full documentation in:
- `CHROME_WEB_STORE_SUBMISSION.md` - Complete submission guide
- `PRIVACY_PRACTICES_RESPONSES.md` - Detailed privacy responses
- `PRIVACY_POLICY.md` - Privacy policy content
- `PUBLICATION_CHECKLIST.md` - Full checklist
