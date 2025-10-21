# Chrome Web Store Submission Guide

## Prerequisites

1. **Google Account** with Chrome Web Store Developer access
2. **One-time $5 registration fee** (if not already registered)
3. **Promotional Images** (see requirements below)

## Step 1: Build for Production

```bash
npm run build
```

This creates the `dist` folder with all production files.

## Step 2: Create a ZIP File

Zip the contents of the `dist` folder (NOT the dist folder itself):

1. Navigate to the `dist` folder
2. Select ALL files inside (not the folder)
3. Right-click and create ZIP archive
4. Name it: `d365-helper-v1.0.0.zip`

**Important**: The ZIP should contain files directly, not a folder containing files.

## Step 3: Prepare Promotional Images

### Required Images

#### 1. Small Promo Tile (440x280 px) - REQUIRED
- Showcase: "D365 Helper" logo/text
- Background: Blue gradient (#0078d4)
- Text: "Developer Toolkit"

#### 2. Screenshots (1280x800 or 640x400 px) - At least 1 REQUIRED
Create 3-5 screenshots showing:
- Screenshot 1: Toolbar with all features visible
- Screenshot 2: Schema names overlay on a D365 form
- Screenshot 3: Web API Viewer in action
- Screenshot 4: Edit mode with lookup field editing
- Screenshot 5: Auto-fill and unlock features

Tips for screenshots:
- Use a clean D365 environment
- Annotate with arrows/highlights
- Show the extension in action
- Use consistent branding

#### 3. Large Promo Tile (920x680 px) - OPTIONAL
- More detailed version of small tile
- Can include feature list

#### 4. Marquee Promo Tile (1400x560 px) - OPTIONAL
- Banner-style promotional image
- Highlight key features

### Tools for Creating Images
- **Canva** (easiest, has templates)
- **Figma** (professional)
- **Photoshop/GIMP**
- **PowerPoint** (export as PNG)

## Step 4: Fill Out Chrome Web Store Developer Dashboard

### Go to: https://chrome.google.com/webstore/devconsole

### Item Details

**Display Name**: D365 Helper - Developer Toolkit

**Summary** (132 char max):
```
Essential developer toolkit for Microsoft Dynamics 365. Manage fields, view schemas, edit Web API data, unlock fields & more!
```

**Description**:
Use content from `STORE_LISTING.md`

**Category**: Developer Tools

**Language**: English (United States)

### Privacy Practices Tab

**IMPORTANT**: You must complete the Privacy practices tab before you can publish.

**See `PRIVACY_PRACTICES_RESPONSES.md` for detailed responses to all required fields.**

**Quick Summary**:

1. **Single Purpose Description**: Developer productivity tools for Microsoft Dynamics 365
2. **Permission Justifications**: Detailed justifications provided in PRIVACY_PRACTICES_RESPONSES.md for:
   - activeTab
   - clipboardWrite
   - storage
   - host permissions
3. **Remote Code**: No
4. **Data Handling**: No data collected, stored, or transmitted
5. **Certifications**: Complete all required certifications in the dashboard

**Privacy Policy**:
Upload `PRIVACY_POLICY.md` or host it on GitHub and provide the URL

**Contact Email**:
- Enter your contact email on the Account tab
- Verify your email before publishing (check inbox for verification link)

### Store Listing

**Upload**:
1. Small Promo Tile (440x280)
2. At least 1 Screenshot
3. Icons are from dist/icons folder

**Website**: Your GitHub repository URL

**Support Email**: Your email address

### Distribution

**Visibility**: Public

**Regions**: All regions

## Step 5: Submit for Review

1. Click "Submit for Review"
2. Review can take 1-3 days (sometimes up to a week)
3. You'll receive email notification

## Step 6: After Approval

Once approved:
1. Extension will be live on Chrome Web Store
2. Update the manifest.json `homepage_url` with the actual store URL
3. Share the link with D365 community!

## Common Rejection Reasons (and How to Avoid)

❌ **Misleading Description**
✅ Be accurate about what the extension does

❌ **Missing Privacy Policy**
✅ Include PRIVACY_POLICY.md and link in manifest

❌ **Unclear Permission Usage**
✅ Clearly justify each permission (done above)

❌ **Poor Quality Images**
✅ Use high-quality, professional screenshots

❌ **Trademark Issues**
✅ Include disclaimer about Microsoft trademarks

## Post-Launch Checklist

- [ ] Monitor reviews and respond to user feedback
- [ ] Set up GitHub for issue tracking
- [ ] Create release notes for future updates
- [ ] Consider creating a landing page
- [ ] Share in D365 communities (Reddit, Forums, LinkedIn)

## Updating the Extension

When releasing updates:
1. Update `version` in manifest.json (e.g., 1.0.1, 1.1.0, 2.0.0)
2. Build: `npm run build`
3. Create new ZIP
4. Upload to Chrome Web Store
5. Provide update description
6. Submit for review

### Version Numbering
- **Major (1.x.x)**: Breaking changes
- **Minor (x.1.x)**: New features
- **Patch (x.x.1)**: Bug fixes

## Support & Maintenance

- Respond to user reviews within 48 hours
- Fix critical bugs within 1 week
- Release updates every 2-3 months with improvements
- Monitor Chrome Web Store dashboard for issues

---

**Good luck with your submission! 🚀**
