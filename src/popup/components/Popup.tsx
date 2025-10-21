import React from 'react';
import './Popup.css';

const Popup: React.FC = () => {
  return (
    <div className="popup">
      <header className="popup-header">
        <h1>⚡ D365 Helper</h1>
        <p className="version">Version 1.0.0</p>
      </header>

      <div className="popup-content">
        <section className="info-section">
          <h2>Features</h2>
          <ul className="feature-list">
            <li>
              <strong>Show/Hide Fields</strong> - Toggle visibility of all form fields
            </li>
            <li>
              <strong>Show/Hide Sections</strong> - Toggle visibility of all form sections
            </li>
            <li>
              <strong>Schema Names</strong> - Display logical names as overlays on fields
            </li>
            <li>
              <strong>Copy Schema Names</strong> - Copy all field schema names to clipboard
            </li>
            <li>
              <strong>Copy Record ID</strong> - Copy current record ID to clipboard
            </li>
            <li>
              <strong>Web API Viewer</strong> - View record data via Web API in a new tab
            </li>
            <li>
              <strong>Form Editor</strong> - Quick access to form editor
            </li>
            <li>
              <strong>Plugin Trace Logs</strong> - Open Plugin Trace Logs list in a new tab
            </li>
          </ul>
        </section>

        <section className="info-section">
          <h2>Usage</h2>
          <p>
            Navigate to any Dynamics 365 form page and the toolbar will appear at the top of the page.
          </p>
          <p>
            Click the buttons in the toolbar to use the different features.
          </p>
        </section>

        <section className="info-section">
          <h2>About</h2>
          <p>
            D365 Helper is a developer toolkit designed to make working with Dynamics 365 faster and more efficient.
          </p>
        </section>
      </div>

      <footer className="popup-footer">
        <p>Made with ❤️ for D365 Developers</p>
      </footer>
    </div>
  );
};

export default Popup;
