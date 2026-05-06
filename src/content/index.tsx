import React from 'react';
import ReactDOM from 'react-dom/client';
import D365Toolbar from './components/D365Toolbar';
import { setShellContainerOffset } from './utils/shellLayout';

// Inject the script that has access to window.Xrm
const injectPageScript = () => {
  const script = document.createElement('script');
  // Add cache-busting parameter using extension version
  const manifest = chrome.runtime.getManifest();
  const scriptUrl = chrome.runtime.getURL('injected.bundle.js') + '?v=' + manifest.version + '.' + Date.now();
  script.src = scriptUrl;
  script.onload = () => {
    script.remove();
  };
  script.onerror = (err) => {
    console.error('[D365 Helper] Failed to load injected script:', err);
  };
  (document.head || document.documentElement).appendChild(script);
};

// Check if we're on a D365 form page
const isFormPage = (): boolean => {
  // Check if Xrm and Xrm.Page are available (indicates a form page)
  const xrm = (window as any).Xrm;
  if (xrm && xrm.Page && xrm.Page.data && xrm.Page.data.entity) {
    return true;
  }

  // Check URL patterns for form pages
  const url = window.location.href;
  if (url.includes('pagetype=entityrecord') ||
      url.includes('etn=') ||
      url.includes('extraqs=') ||
      url.includes('formid=')) {
    return true;
  }

  return false;
};

// Wait for DOM to be ready
const initializeToolbar = () => {
  // Only initialize on form pages
  if (!isFormPage()) {
    return;
  }

  // Always mount the toolbar component on form pages, regardless of `showTool`.
  // The component itself decides whether to render the visible bar — but its keyboard
  // listeners (Ctrl+K, user-bound chords) stay attached either way.
  chrome.storage.sync.get(['toolbarPosition'], (result) => {
    const toolbarPosition = result.toolbarPosition !== undefined ? result.toolbarPosition : 'bottom';

    // Inject the page script first
    injectPageScript();

    // Create container for toolbar
    const toolbarContainer = document.createElement('div');
    toolbarContainer.id = 'd365-helper-toolbar-root';

    // Apply position class to the root container
    if (toolbarPosition === 'bottom') {
      toolbarContainer.classList.add('toolbar-bottom');
    }

    // Insert toolbar at the very top of body
    document.body.insertBefore(toolbarContainer, document.body.firstChild);

    // Reserve vertical space for the toolbar (toolbar may reclaim it if hidden)
    setShellContainerOffset(70, toolbarPosition);

    // Render React toolbar
    const root = ReactDOM.createRoot(toolbarContainer);
    root.render(<D365Toolbar />);
  });
};

// Listen for setting changes that need a full page reload to take effect.
// `showTool` no longer requires a reload — the toolbar reacts to it in-place
// so its keyboard listeners survive a hide/show cycle.
chrome.storage.onChanged.addListener((changes, namespace) => {
  if (namespace === 'sync' && (changes.toolbarPosition || changes.schemaOverlayColor)) {
    window.location.reload();
  }
});

// Initialize when DOM is ready
const init = () => {
  // Wait a bit for Xrm to load
  setTimeout(() => {
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', initializeToolbar);
    } else {
      initializeToolbar();
    }
  }, 1000);
};

init();
