import React, { useState, useEffect } from 'react';
import './Popup.css';

type TabType = 'features' | 'settings';

const Popup: React.FC = () => {
  const [activeTab, setActiveTab] = useState<TabType>('features');
  const [showTool, setShowTool] = useState(true);
  const [notificationDuration, setNotificationDuration] = useState(3);
  const [toolbarPosition, setToolbarPosition] = useState<'top' | 'bottom'>('top');
  const [schemaOverlayColor, setSchemaOverlayColor] = useState('#0078d4');
  const [traceLogLimit, setTraceLogLimit] = useState(20);
  const [skipPluginsByDefault, setSkipPluginsByDefault] = useState(false);

  useEffect(() => {
    chrome.storage.sync.get([
      'showTool',
      'notificationDuration',
      'toolbarPosition',
      'schemaOverlayColor',
      'traceLogLimit',
      'skipPluginsByDefault'
    ], (result) => {
      if (result.showTool !== undefined) {
        setShowTool(result.showTool);
      }
      if (result.notificationDuration !== undefined) {
        setNotificationDuration(result.notificationDuration);
      }
      if (result.toolbarPosition !== undefined) {
        setToolbarPosition(result.toolbarPosition);
      }
      if (result.schemaOverlayColor !== undefined) {
        setSchemaOverlayColor(result.schemaOverlayColor);
      }
      if (result.traceLogLimit !== undefined) {
        setTraceLogLimit(result.traceLogLimit);
      }
      if (result.skipPluginsByDefault !== undefined) {
        setSkipPluginsByDefault(result.skipPluginsByDefault);
      }
    });
  }, []);

  const handleToggleShowTool = (value: boolean) => {
    setShowTool(value);
    chrome.storage.sync.set({ showTool: value });
  };

  const handleNotificationDurationChange = (value: number) => {
    setNotificationDuration(value);
    chrome.storage.sync.set({ notificationDuration: value });
  };

  const handleToolbarPositionChange = (value: 'top' | 'bottom') => {
    setToolbarPosition(value);
    chrome.storage.sync.set({ toolbarPosition: value });
  };

  const handleSchemaOverlayColorChange = (value: string) => {
    setSchemaOverlayColor(value);
    chrome.storage.sync.set({ schemaOverlayColor: value });
  };

  const handleTraceLogLimitChange = (value: number) => {
    setTraceLogLimit(value);
    chrome.storage.sync.set({ traceLogLimit: value });
  };

  const handleSkipPluginsByDefaultChange = (value: boolean) => {
    setSkipPluginsByDefault(value);
    chrome.storage.sync.set({ skipPluginsByDefault: value });
  };

  return (
    <div className="popup">
      <header className="popup-header">
        <h1>
          <a
            href="https://www.reddishgreen.co.uk"
            target="_blank"
            rel="noopener noreferrer"
            style={{ display: 'inline-block', lineHeight: 0 }}
          >
            <img
              className="popup-logo"
              src={chrome.runtime.getURL('icons/RG%20Logo_White_Stacked.svg')}
              alt="RG Logo"
              onError={(e) => {
                e.currentTarget.style.display = 'none';
              }}
            />
          </a>
          D365 Helper
        </h1>
        <p className="version">Version {chrome.runtime.getManifest().version}</p>
      </header>

      <div className="popup-tabs">
        <button
          className={`popup-tab ${activeTab === 'features' ? 'popup-tab-active' : ''}`}
          onClick={() => setActiveTab('features')}
        >
          Features
        </button>
        <button
          className={`popup-tab ${activeTab === 'settings' ? 'popup-tab-active' : ''}`}
          onClick={() => setActiveTab('settings')}
        >
          Settings
        </button>
      </div>

      <div className="popup-content">
        {activeTab === 'features' && (
          <>
            <section className="info-section">
              <h2>
                <span className="info-section-icon" aria-hidden="true">&#x2139;</span>
                About D365 Helper
              </h2>
              <p>
                A professional developer toolkit designed to streamline Dynamics 365 development workflows.
                Built to enhance productivity with powerful debugging and customization tools.
              </p>
            </section>

            <section className="info-section">
              <h2>
                <span className="info-section-icon" aria-hidden="true">&#x1F4CB;</span>
                Field Controls
              </h2>
              <p>Toggle visibility of all form fields and restore to original state</p>
            </section>

            <section className="info-section">
              <h2>
                <span className="info-section-icon" aria-hidden="true">&#x1F4C1;</span>
                Section Management
              </h2>
              <p>Show or hide all form sections with one click</p>
            </section>

            <section className="info-section">
              <h2>
                <span className="info-section-icon" aria-hidden="true">&#x1F4C4;</span>
                Schema Names
              </h2>
              <p>Display and copy logical field names for development</p>
            </section>

            <section className="info-section">
              <h2>
                <span className="info-section-icon" aria-hidden="true">&#x1F4DD;</span>
                Record Tools
              </h2>
              <p>Copy record IDs and open Web API viewer for debugging</p>
            </section>

            <section className="info-section">
              <h2>
                <span className="info-section-icon" aria-hidden="true">&#x1F527;</span>
                Developer Mode
              </h2>
              <p>Enable all fields, unlock controls, and add test data</p>
            </section>

            <section className="info-section">
              <h2>
                <span className="info-section-icon" aria-hidden="true">&#x1F4CA;</span>
                Data Analysis
              </h2>
              <p>View plugin trace logs, option sets, and form libraries</p>
            </section>

            <section className="info-section warning-section">
              <h2>
                <span className="info-section-icon" aria-hidden="true">&#9888;</span>
                Important Notice
              </h2>
              <p>
                This extension is intended for authorized developers only. Use only on systems where you have proper permissions.
                All data modifications require appropriate D365 security roles and organizational approval.
              </p>
            </section>
          </>
        )}

        {activeTab === 'settings' && (
          <>
            <h3 className="settings-section-title">General</h3>

            <section className="info-section settings-row">
              <div className="settings-label">
                <p className="settings-description">
                  <span className="settings-icon" aria-hidden="true">&#9881;</span>
                  Display the D365 Helper toolbar on form pages
                </p>
              </div>
              <div className="settings-toggle">
                <button
                  className={`toggle-btn ${showTool ? 'toggle-yes' : ''}`}
                  onClick={() => handleToggleShowTool(true)}
                >
                  On
                </button>
                <button
                  className={`toggle-btn ${!showTool ? 'toggle-no' : ''}`}
                  onClick={() => handleToggleShowTool(false)}
                >
                  Off
                </button>
              </div>
            </section>

            <section className="info-section settings-row">
              <div className="settings-label">
                <p className="settings-description">
                  <span className="settings-icon" aria-hidden="true">&#128205;</span>
                  Toolbar position on form pages
                </p>
              </div>
              <div className="settings-toggle">
                <button
                  className={`toggle-btn ${toolbarPosition === 'top' ? 'toggle-yes' : ''}`}
                  onClick={() => handleToolbarPositionChange('top')}
                >
                  Top
                </button>
                <button
                  className={`toggle-btn ${toolbarPosition === 'bottom' ? 'toggle-no' : ''}`}
                  onClick={() => handleToolbarPositionChange('bottom')}
                >
                  Bottom
                </button>
              </div>
            </section>

            <section className="info-section settings-row">
              <div className="settings-label">
                <p className="settings-description">
                  <span className="settings-icon" aria-hidden="true">&#9200;</span>
                  Notification display duration (seconds)
                </p>
              </div>
              <div className="settings-control">
                <input
                  type="range"
                  min="1"
                  max="10"
                  value={notificationDuration}
                  onChange={(e) => handleNotificationDurationChange(Number(e.target.value))}
                  className="slider"
                />
                <span className="slider-value">{notificationDuration}s</span>
              </div>
            </section>

            <h3 className="settings-section-title">Appearance</h3>

            <section className="info-section settings-row">
              <div className="settings-label">
                <p className="settings-description">
                  <span className="settings-icon" aria-hidden="true">&#127912;</span>
                  Schema overlay color
                </p>
              </div>
              <div className="settings-control">
                <input
                  type="color"
                  value={schemaOverlayColor}
                  onChange={(e) => handleSchemaOverlayColorChange(e.target.value)}
                  className="color-picker"
                />
                <span className="color-value">{schemaOverlayColor}</span>
              </div>
            </section>

            <h3 className="settings-section-title">Developer Tools</h3>

            <section className="info-section settings-row">
              <div className="settings-label">
                <p className="settings-description">
                  <span className="settings-icon" aria-hidden="true">&#128221;</span>
                  Plugin trace log default limit
                </p>
              </div>
              <div className="settings-control">
                <select
                  value={traceLogLimit}
                  onChange={(e) => handleTraceLogLimitChange(Number(e.target.value))}
                  className="dropdown"
                >
                  <option value={10}>10 records</option>
                  <option value={20}>20 records</option>
                  <option value={50}>50 records</option>
                  <option value={100}>100 records</option>
                  <option value={200}>200 records</option>
                </select>
              </div>
            </section>

            <section className="info-section settings-row">
              <div className="settings-label">
                <p className="settings-description">
                  <span className="settings-icon" aria-hidden="true">&#128295;</span>
                  Skip plugin execution by default (Web API Viewer)
                </p>
              </div>
              <div className="settings-toggle">
                <button
                  className={`toggle-btn ${skipPluginsByDefault ? 'toggle-yes' : ''}`}
                  onClick={() => handleSkipPluginsByDefaultChange(true)}
                >
                  On
                </button>
                <button
                  className={`toggle-btn ${!skipPluginsByDefault ? 'toggle-no' : ''}`}
                  onClick={() => handleSkipPluginsByDefaultChange(false)}
                >
                  Off
                </button>
              </div>
            </section>
          </>
        )}
      </div>

      <footer className="popup-footer">
        <p>Made by ReddishGreen</p>
      </footer>
    </div>
  );
};

export default Popup;
