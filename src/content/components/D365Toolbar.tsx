import React, { useState, useEffect, useRef, useCallback } from 'react';
import { D365Helper } from '../utils/D365Helper';
import { restoreShellContainerLayout, setShellContainerOffset } from '../utils/shellLayout';
import FormLibrariesAnalyzer from './FormLibrariesAnalyzer';

const D365Toolbar: React.FC = () => {
  const [isMinimized, setIsMinimized] = useState(false);
  const [showSchemaNames, setShowSchemaNames] = useState(false);
  const [notification, setNotification] = useState<string>('');
  const [showLibraries, setShowLibraries] = useState(false);
  const [librariesData, setLibrariesData] = useState<any>(null);
  const toolbarRef = useRef<HTMLDivElement | null>(null);

  const helper = new D365Helper();

  const updateShellOffset = useCallback(() => {
    const toolbarHeight = toolbarRef.current?.getBoundingClientRect().height;
    const fallbackHeight = isMinimized ? 35 : 70;
    setShellContainerOffset(Math.round(toolbarHeight ?? fallbackHeight));
  }, [isMinimized]);

  useEffect(() => {
    updateShellOffset();
  }, [updateShellOffset]);

  useEffect(() => {
    const handleResize = () => updateShellOffset();
    window.addEventListener('resize', handleResize);
    handleResize();

    return () => {
      window.removeEventListener('resize', handleResize);
    };
  }, [updateShellOffset]);

  useEffect(() => {
    return () => {
      restoreShellContainerLayout();
    };
  }, []);

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

  const handleOpenPluginTraceLogs = () => {
    try {
      const url = helper.getPluginTraceLogsUrl();
      window.open(url, '_blank');
      showNotification('Opening Plugin Trace Logs...');
    } catch (error) {
      showNotification('Error opening Plugin Trace Logs');
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

  const handleShowLibraries = async () => {
    try {
      showNotification('Loading JavaScript libraries...');
      const data = await helper.getFormLibraries();
      setLibrariesData(data);
      setShowLibraries(true);
      showNotification('');
    } catch (error) {
      showNotification('Error loading libraries');
    }
  };

  const handleCloseLibraries = () => {
    setShowLibraries(false);
    setLibrariesData(null);
  };

  if (isMinimized) {
    return (
      <div ref={toolbarRef} className="d365-toolbar d365-toolbar-minimized">
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
    <div ref={toolbarRef} className="d365-toolbar">
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
          <button
            className="d365-toolbar-btn"
            onClick={handleOpenPluginTraceLogs}
            title="View Plugin Trace Logs"
          >
            Trace Logs
            onClick={handleShowLibraries}
            title="View JavaScript libraries and event handlers"
          >
            JS Libraries
          </button>
        </div>
      </div>

      {notification && (
        <div className="d365-toolbar-notification">
          {notification}
        </div>
      )}

      {showLibraries && (
        <FormLibrariesAnalyzer
          data={librariesData}
          onClose={handleCloseLibraries}
        />
      )}
    </div>
  );
};

export default D365Toolbar;
