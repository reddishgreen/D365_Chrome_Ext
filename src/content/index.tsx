import React from 'react';
import ReactDOM from 'react-dom/client';
import D365Toolbar from './components/D365Toolbar';
import { setShellContainerOffset } from './utils/shellLayout';

// Inject the script that has access to window.Xrm
const injectPageScript = () => {
  const script = document.createElement('script');
  script.src = chrome.runtime.getURL('injected.bundle.js');
  script.onload = () => {
    script.remove();
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
    console.log('D365 Helper: Not a form page, skipping toolbar initialization');
    return;
  }

  console.log('D365 Helper: Initializing toolbar on form page');

  // Inject the page script first
  injectPageScript();

  // Create container for toolbar
  const toolbarContainer = document.createElement('div');
  toolbarContainer.id = 'd365-helper-toolbar-root';

  // Insert toolbar at the very top of body
  document.body.insertBefore(toolbarContainer, document.body.firstChild);

  // Reserve vertical space for the toolbar
  setShellContainerOffset(70);

  // Render React toolbar
  const root = ReactDOM.createRoot(toolbarContainer);
  root.render(<D365Toolbar />);
};

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
