# Chrome Web Store Publication Checklist

## ✅ Pre-Submission Checklist

### Code & Build
- [x] Production build completed (`npm run build`)
- [x] All features tested and working
- [x] No console errors
- [x] Extension works on multiple D365 instances
- [x] Icons created (16px, 48px, 128px)

### Documentation
- [x] README.md updated
- [x] PRIVACY_POLICY.md created
- [x] STORE_LISTING.md created
- [x] CHROME_WEB_STORE_SUBMISSION.md guide created

### Manifest.json
- [x] Name: "D365 Helper - Developer Toolkit"
- [x] Version: 1.0.0
- [x] Description optimized (under 132 chars)
- [x] All permissions justified
- [x] Homepage URL (update after GitHub repo creation)
- [x] Author field added

## 📦 Files to Submit

### In the ZIP file (from `dist` folder):
- [ ] manifest.json
- [ ] content.bundle.js
- [ ] content.css
- [ ] injected.bundle.js
- [ ] webapi-viewer.bundle.js
- [ ] webapi-viewer.html
- [ ] popup.bundle.js
- [ ] popup.html
- [ ] icons/ folder (icon16.png, icon48.png, icon128.png)

## 🎨 Promotional Materials Needed

### Required
- [ ] Small Promo Tile (440x280 px)
- [ ] At least 1 Screenshot (1280x800 px)

### Recommended
- [ ] 3-5 Screenshots showing different features
- [ ] Large Promo Tile (920x680 px)
- [ ] Marquee Promo Tile (1400x560 px)

### Screenshot Ideas
1. Main toolbar with all buttons visible on a D365 form
2. Schema names overlay feature in action
3. Web API Viewer displaying record data
4. Edit mode with lookup field editing
5. Before/After: Auto-fill feature demonstration

## 📝 Store Listing Information

### Required Fields
- [ ] Display name
- [ ] Summary (132 char max)
- [ ] Detailed description (from STORE_LISTING.md)
- [ ] Category: Developer Tools
- [ ] Language: English
- [ ] Privacy policy URL or inline content
- [ ] Single purpose description
- [ ] Permission justifications
- [ ] Support email
- [ ] Website URL (GitHub)

### Privacy Questions
- [ ] Confirm: Does NOT collect user data
- [ ] Confirm: Does NOT use remote code
- [ ] Confirm: All processing is local

## 🚀 Submission Steps

1. **Create ZIP File**
   ```bash
   # Navigate to dist folder
   cd dist
   # Select all files (NOT the dist folder itself)
   # Create ZIP: d365-helper-v1.0.0.zip
   ```

2. **Chrome Web Store Developer Dashboard**
   - Go to: https://chrome.google.com/webstore/devconsole
   - Click "New Item"
   - Upload ZIP file
   - Fill out all required fields
   - Upload promotional images
   - Submit for review

3. **Review Timeline**
   - Initial review: 1-3 business days
   - May take up to 1 week
   - Check email for updates

## 🔍 Pre-Launch Testing

### Test on Clean Installation
- [ ] Install from dist folder (Load unpacked)
- [ ] Test on D365 form page
- [ ] All toolbar buttons work
- [ ] Show/Hide fields works
- [ ] Schema names overlay works
- [ ] Copy to clipboard works
- [ ] Unlock fields works
- [ ] Auto-fill works
- [ ] Web API viewer opens
- [ ] Edit mode works
- [ ] Bypass plugins option works
- [ ] Form editor link works
- [ ] Minimize/expand works
- [ ] No console errors

### Test on Different Environments
- [ ] D365 Sales
- [ ] D365 Customer Service
- [ ] Power Apps Model-Driven App
- [ ] Different browsers (Chrome, Edge)

## 📧 Communication

### Support Email Setup
- [ ] Create dedicated support email
- [ ] Set up auto-responder
- [ ] Prepare FAQ document

### Community
- [ ] GitHub repository created
- [ ] Issue tracking enabled
- [ ] Contributing guidelines
- [ ] Code of conduct

## 🎯 Post-Submission

### Once Approved
- [ ] Update manifest.json homepage_url with store link
- [ ] Update README badges with store link
- [ ] Share on LinkedIn
- [ ] Share on D365 Community forums
- [ ] Share on Reddit (r/Dynamics365, r/PowerPlatform)
- [ ] Create blog post/announcement
- [ ] Monitor initial reviews

### Monitoring
- [ ] Set up Google Alerts for extension name
- [ ] Check Chrome Web Store reviews daily (first week)
- [ ] Respond to reviews within 48 hours
- [ ] Track usage statistics

## 📊 Success Metrics

### Week 1 Targets
- [ ] 50+ installs
- [ ] 4+ star average rating
- [ ] Positive user feedback

### Month 1 Targets
- [ ] 200+ installs
- [ ] 10+ reviews
- [ ] Feature requests documented

## 🛠️ Future Updates

### Version 1.1.0 Ideas
- Export form data to Excel
- Bulk update fields
- Form validation viewer
- Business rules analyzer
- Record comparison tool
- Dark mode support

---

## 🎉 Ready to Submit?

Double-check this entire list, then:

1. Build: `npm run build`
2. Create ZIP from `dist` contents
3. Prepare promotional images
4. Go to Chrome Web Store Developer Console
5. Submit!

**Good luck! 🚀**
