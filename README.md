# D365 Helper Chrome Extension

A powerful developer toolkit for Microsoft Dynamics 365, built with React and TypeScript.

## Features

### Toolbar (Bookmark Bar Style)
A sleek toolbar that appears at the top of all D365 form pages with quick access to developer tools.

### Field Management
- **Show All Fields** - Make all fields on the form visible
- **Hide All Fields** - Hide all fields on the form
- **Show All Sections** - Expand all form sections
- **Hide All Sections** - Collapse all form sections

### Schema Name Tools
- **Show Schema Names** - Display logical/schema names as overlays on all fields
- **Copy Schema Names** - Copy all field schema names to clipboard
- **Click to Copy** - Click any schema name overlay to copy individual names

### Developer Tools
- **Copy Record ID** - Quickly copy the current record's GUID to clipboard
- **Web API Viewer** - Open current record's Web API data in a beautifully formatted viewer
- **Form Editor** - Quick link to open the current form in the form editor

### Web API Viewer
A dedicated viewer that displays Web API data with:
- JSON syntax highlighting
- Expandable/collapsible sections
- Search functionality
- Copy to clipboard
- Formatted dates
- Type indicators (strings, numbers, booleans, nulls)

## Installation

### Prerequisites
- Node.js (v16 or higher)
- npm or yarn

### Build Steps

1. **Install Dependencies**
   ```bash
   npm install
   ```

2. **Build the Extension**

   For development (with watch mode):
   ```bash
   npm run dev
   ```

   For production:
   ```bash
   npm run build
   ```

3. **Add Icons** (Optional but recommended)

   Add your icon files to the `icons/` directory:
   - `icon16.png` (16x16)
   - `icon48.png` (48x48)
   - `icon128.png` (128x128)

4. **Load in Chrome**

   - Open Chrome and go to `chrome://extensions/`
   - Enable "Developer mode" (toggle in top right)
   - Click "Load unpacked"
   - Select the `dist` folder from your project

## Usage

1. Navigate to any Dynamics 365 form (e.g., Contact, Account, Custom Entity)
2. The D365 Helper toolbar will automatically appear at the top of the page
3. Use the buttons to:
   - Toggle field and section visibility
   - Show/hide schema names
   - Copy schema names or record IDs
   - Open Web API viewer or form editor

### Keyboard Shortcuts
The toolbar can be minimized by clicking the "−" button and restored by clicking "D365 Helper".

## Project Structure

```
D365_Chrome_Ext/
├── src/
│   ├── content/              # Content script (injected into D365 pages)
│   │   ├── components/
│   │   │   └── D365Toolbar.tsx
│   │   ├── utils/
│   │   │   └── D365Helper.ts
│   │   ├── styles.css
│   │   └── index.tsx
│   ├── webapi-viewer/        # Web API viewer page
│   │   ├── components/
│   │   │   ├── WebAPIViewer.tsx
│   │   │   └── WebAPIViewer.css
│   │   ├── index.html
│   │   └── index.tsx
│   └── popup/                # Extension popup
│       ├── components/
│       │   ├── Popup.tsx
│       │   └── Popup.css
│       ├── index.html
│       └── index.tsx
├── icons/                    # Extension icons
├── dist/                     # Build output (generated)
├── manifest.json             # Chrome extension manifest
├── webpack.config.js         # Webpack configuration
├── tsconfig.json            # TypeScript configuration
└── package.json             # Dependencies
```

## Development

### Watch Mode
Run the extension in development mode with auto-rebuild:
```bash
npm run dev
```

After making changes, go to `chrome://extensions/` and click the refresh icon on the D365 Helper extension.

### Clean Build
Remove the dist folder and rebuild:
```bash
npm run clean
npm run build
```

## Technologies Used

- **React 18** - UI framework
- **TypeScript** - Type-safe development
- **Webpack** - Bundling and build process
- **Chrome Extension API** - Browser integration
- **Dynamics 365 Client API (Xrm)** - D365 form interactions

## Supported D365 URLs

The extension automatically activates on:
- `*.dynamics.com`
- `*.crm.dynamics.com`
- `*.crm*.dynamics.com`

## Troubleshooting

### Toolbar not appearing
- Ensure you're on a Dynamics 365 form page (not a view or dashboard)
- Check the browser console for errors
- Refresh the page
- Reload the extension in `chrome://extensions/`

### Schema names not showing
- Some fields may not have accessible schema names
- Ensure you're on a form with visible fields
- Check browser console for errors

### Web API viewer not loading data
- Ensure you're authenticated to D365
- Check CORS and security settings
- Verify the record has data

## Future Enhancements

Potential features for future versions:
- Export form data to various formats (JSON, CSV, Excel)
- Form validation rule viewer
- Business rule analyzer
- Power Automate flow trigger
- Environment switcher
- Custom field inspector
- Record duplication detector

## Contributing

Feel free to submit issues, fork the repository, and create pull requests for any improvements.

## License

MIT License - feel free to use and modify as needed.

## Support

For issues and questions, please check:
- Browser console for error messages
- Ensure latest Chrome version
- Verify D365 permissions

---

Made with ❤️ for D365 Developers
