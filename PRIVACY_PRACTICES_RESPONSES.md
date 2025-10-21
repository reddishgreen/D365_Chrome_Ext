# Chrome Web Store Privacy Practices Responses

Use these responses when filling out the Privacy practices tab in the Chrome Web Store Developer Dashboard.

---

## 1. Single Purpose Description

**What to enter:**

```
This extension provides developer productivity tools exclusively for Microsoft Dynamics 365, including field management, schema name display, Web API data viewing and editing capabilities.
```

---

## 2. Permission Justifications

### activeTab Permission

**Justification:**

```
The activeTab permission is required to inject the developer toolbar into active Dynamics 365 form pages. When users click the extension icon or use toolbar features, we need access to the current tab to:
- Display the floating toolbar interface
- Read form metadata and field schema names
- Show/hide form fields and sections
- Access the Dynamics 365 Client API (Xrm) to interact with form controls
This permission is only used when the user actively engages with the extension on a D365 page.
```

### clipboardWrite Permission

**Justification:**

```
The clipboardWrite permission enables users to quickly copy developer information to their clipboard with one click. This is used when:
- Copying individual field schema names
- Copying all schema names at once
- Copying record IDs
- Copying Web API data values
All clipboard operations are triggered explicitly by the user clicking a "Copy" button. No data is copied automatically or in the background.
```

### storage Permission

**Justification:**

```
The storage permission saves user preferences locally in the browser to improve user experience:
- Remembering toolbar state (minimized/expanded)
- Saving user's preferred toolbar position
- Storing display preferences
All data is stored locally on the user's device only. No data is sent to external servers or shared with third parties.
```

### Host Permissions (*.dynamics.com, *.crm*.dynamics.com)

**Justification:**

```
Host permissions for Dynamics 365 domains are essential for the extension to function. These permissions allow the extension to:
- Run exclusively on Microsoft Dynamics 365 instances (dynamics.com and regional CRM domains)
- Access form metadata and field information using the D365 Client API
- Interact with Web API endpoints for data viewing and editing
- Inject the developer toolbar interface onto D365 pages
The extension is specifically designed for D365 developers and only activates on official Microsoft Dynamics 365 URLs. It does not run on any other websites.
```

---

## 3. Remote Code Declaration

**Does your extension execute remote code?**

```
No
```

**Explanation:**

```
This extension does not fetch, execute, or load any code from remote servers. All JavaScript code is bundled within the extension package at build time using Webpack. The extension only interacts with the user's own Dynamics 365 environment's Web API for data viewing and editing purposes.
```

---

## 4. Data Usage Certification

### Does your extension handle personal or sensitive user data?

**Answer:** No

**Explanation:**

```
This extension does not collect, store, transmit, or process any personal or sensitive user data. All operations are performed locally in the user's browser. While the extension can view and edit data within the user's Dynamics 365 environment, this data:
- Never leaves the user's browser
- Is not sent to any external servers
- Is not stored by the extension (except user preferences via storage permission)
- Is only accessed when the user explicitly uses extension features
```

### Data Handling

**Data collected:** None

**Data usage:** None

**Data transmission:** None

---

## 5. Certification Statement

**I certify that:**

- [x] My extension complies with Chrome Web Store Developer Program Policies
- [x] I have provided accurate information about data collection and usage
- [x] All permissions requested are necessary for the extension's core functionality
- [x] I will not sell user data to third parties
- [x] I will not use or transfer user data for purposes unrelated to the extension's core functionality
- [x] I will not use or transfer user data to determine creditworthiness or for lending purposes

---

## 6. Contact Information Required

Before you can publish, ensure you have:

1. **Contact Email**: Enter a valid email address on the Account tab
2. **Email Verification**: Verify your contact email (check your inbox for verification link)

**Recommended Email Format:**
```
support@yourdomain.com
OR
your.name@gmail.com
```

---

## Additional Store Listing Information

### Category
```
Developer Tools
```

### Target Audience
```
Developers and administrators working with Microsoft Dynamics 365
```

### Maturity Rating
```
Everyone
```

### Language
```
English (United States)
```

---

## Privacy Policy URL

If hosting your privacy policy online, use this URL format:
```
https://github.com/yourusername/d365-helper/blob/main/PRIVACY_POLICY.md
```

Otherwise, you can paste the contents of PRIVACY_POLICY.md directly into the text field.

---

## Support Information

### Support URL (optional but recommended)
```
https://github.com/yourusername/d365-helper/issues
```

### Support Email
```
your.email@example.com
```

---

## Disclaimer Text (Include in Store Description)

```
Disclaimer: This extension is an independent tool created for developer productivity.
It is not affiliated with, endorsed by, or sponsored by Microsoft Corporation.
Microsoft, Dynamics 365, and related trademarks are property of Microsoft Corporation.
```

---

## Summary Checklist

Before submitting, verify you have completed:

- [ ] Single purpose description entered
- [ ] All 4 permission justifications provided (activeTab, clipboardWrite, storage, host permissions)
- [ ] Remote code declaration answered (No)
- [ ] Data usage certification completed
- [ ] Contact email entered and verified on Account tab
- [ ] Privacy policy provided (URL or text)
- [ ] All certifications checked
- [ ] Store listing complete with images
- [ ] Disclaimer about Microsoft trademarks included

---

**Once all items are complete, you can submit your extension for review!**
