import React, { useState, useEffect } from 'react';
import { D365Helper } from '../utils/D365Helper';

const D365Toolbar: React.FC = () => {
  const [isMinimized, setIsMinimized] = useState(false);
  const [showSchemaNames, setShowSchemaNames] = useState(false);
  const [notification, setNotification] = useState<string>('');

  const helper = new D365Helper();

  // Adjust shell-container margin when minimized state changes
  useEffect(() => {
    const shellContainer = document.getElementById('shell-container');
    if (shellContainer) {
      shellContainer.style.marginTop = isMinimized ? '35px' : '70px';
    }
  }, [isMinimized]);

  const showNotification = (message: string) => {
    setNotification(message);
    setTimeout(() => setNotification(''), 3000);
  };

  const handleToggleFields = async (show: boolean) => {
    try {
      await helper.toggleAllFields(show);
      showNotification(show ? 'All fields shown' : 'All fields hidden');
    } catch (error) {
      showNotification('Error toggling fields');
    }
  };

  const handleToggleSections = async (show: boolean) => {
    try {
      await helper.toggleAllSections(show);
      showNotification(show ? 'All sections shown' : 'All sections hidden');
    } catch (error) {
      showNotification('Error toggling sections');
    }
  };

  const handleCopyRecordId = async () => {
    try {
      const recordId = await helper.getRecordId();
      if (recordId) {
        await navigator.clipboard.writeText(recordId);
        showNotification('Record ID copied!');
      } else {
        showNotification('No record ID found');
      }
    } catch (error) {
      showNotification('Error copying record ID');
    }
  };

  const handleToggleSchemaNames = async () => {
    const newState = !showSchemaNames;
    setShowSchemaNames(newState);
    try {
      await helper.toggleSchemaOverlay(newState);
      showNotification(newState ? 'Schema names shown' : 'Schema names hidden');
    } catch (error) {
      showNotification('Error toggling schema names');
    }
  };

  const handleCopyAllSchemaNames = async () => {
    try {
      const schemaNames = await helper.getAllSchemaNames();
      await navigator.clipboard.writeText(schemaNames.join('\n'));
      showNotification(`Copied ${schemaNames.length} schema names!`);
    } catch (error) {
      showNotification('Error copying schema names');
    }
  };

  const handleOpenWebAPI = async () => {
    try {
      const url = await helper.getWebAPIUrl();
      if (url) {
        window.open(url, '_blank');
        showNotification('Opening Web API viewer...');
      } else {
        showNotification('Unable to generate Web API URL');
      }
    } catch (error) {
      showNotification('Error opening Web API viewer');
    }
  };

  const handleOpenFormEditor = async () => {
    try {
      const url = await helper.getFormEditorUrl();
      if (url) {
        window.open(url, '_blank');
        showNotification('Opening form editor...');
      } else {
        showNotification('Unable to open form editor');
      }
    } catch (error) {
      showNotification('Error opening form editor');
    }
  };

  const handleUnlockFields = async () => {
    try {
      const count = await helper.unlockFields();
      showNotification(`Unlocked ${count} fields!`);
    } catch (error) {
      showNotification('Error unlocking fields');
    }
  };

  const handleAutoFill = async () => {
    try {
      const count = await helper.autoFillForm();
      showNotification(`Auto-filled ${count} fields!`);
    } catch (error) {
      showNotification('Error auto-filling form');
    }
  };

  if (isMinimized) {
    return (
      <div className="d365-toolbar d365-toolbar-minimized">
        <button
          className="d365-toolbar-btn d365-toolbar-maximize"
          onClick={() => setIsMinimized(false)}
          title="Show D365 Helper"
        >
          ⚡ D365 Helper
        </button>
      </div>
    );
  }

  return (
    <div className="d365-toolbar">
      <div className="d365-toolbar-content">
        <div className="d365-toolbar-section">
          <span className="d365-toolbar-section-label">Fields:</span>
          <button
            className="d365-toolbar-btn"
            onClick={() => handleToggleFields(true)}
            title="Show all fields on the form"
          >
            Show All
          </button>
          <button
            className="d365-toolbar-btn"
            onClick={() => handleToggleFields(false)}
            title="Hide all fields on the form"
          >
            Hide All
          </button>
        </div>

        <div className="d365-toolbar-section">
          <span className="d365-toolbar-section-label">Sections:</span>
          <button
            className="d365-toolbar-btn"
            onClick={() => handleToggleSections(true)}
            title="Show all sections on the form"
          >
            Show All
          </button>
          <button
            className="d365-toolbar-btn"
            onClick={() => handleToggleSections(false)}
            title="Hide all sections on the form"
          >
            Hide All
          </button>
        </div>

        <div className="d365-toolbar-section">
          <span className="d365-toolbar-section-label">Schema:</span>
          <button
            className={`d365-toolbar-btn ${showSchemaNames ? 'd365-toolbar-btn-active' : ''}`}
            onClick={handleToggleSchemaNames}
            title="Toggle schema name overlays on fields"
          >
            {showSchemaNames ? 'Hide Names' : 'Show Names'}
          </button>
          <button
            className="d365-toolbar-btn"
            onClick={handleCopyAllSchemaNames}
            title="Copy all schema names to clipboard"
          >
            Copy All
          </button>
        </div>

        <div className="d365-toolbar-section">
          <span className="d365-toolbar-section-label">Form:</span>
          <button
            className="d365-toolbar-btn"
            onClick={handleUnlockFields}
            title="Unlock all readonly fields"
          >
            Unlock Fields
          </button>
          <button
            className="d365-toolbar-btn"
            onClick={handleAutoFill}
            title="Auto-fill empty fields with sample data"
          >
            Auto Fill
          </button>
        </div>

        <div className="d365-toolbar-section">
          <span className="d365-toolbar-section-label">Tools:</span>
          <button
            className="d365-toolbar-btn"
            onClick={handleCopyRecordId}
            title="Copy current record ID"
          >
            Copy ID
          </button>
          <button
            className="d365-toolbar-btn"
            onClick={handleOpenWebAPI}
            title="Open Web API data in new tab"
          >
            Web API
          </button>
          <button
            className="d365-toolbar-btn"
            onClick={handleOpenFormEditor}
            title="Open form in editor"
          >
            Form Editor
          </button>
        </div>
      </div>

      {notification && (
        <div className="d365-toolbar-notification">
          {notification}
        </div>
      )}
    </div>
  );
};

export default D365Toolbar;
