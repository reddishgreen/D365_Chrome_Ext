import React, { useState, useEffect, useRef, useCallback } from 'react';
import { D365Helper } from '../utils/D365Helper';
import { restoreShellContainerLayout, setShellContainerOffset } from '../utils/shellLayout';
import FormLibrariesAnalyzer from './FormLibrariesAnalyzer';
import PluginTraceLogViewer, { PluginTraceLogData } from './PluginTraceLogViewer';
import OptionSetsViewer, { OptionSetsData } from './OptionSetsViewer';

const D365Toolbar: React.FC = () => {
  const [isMinimized, setIsMinimized] = useState(false);
  const [showSchemaNames, setShowSchemaNames] = useState(false);
  const [notification, setNotification] = useState<string>('');
  const [showLibraries, setShowLibraries] = useState(false);
  const [librariesData, setLibrariesData] = useState<any>(null);
  const [showTraceLogs, setShowTraceLogs] = useState(false);
  const [traceLogData, setTraceLogData] = useState<PluginTraceLogData | null>(null);
  const [showOptionSets, setShowOptionSets] = useState(false);
  const [optionSetData, setOptionSetData] = useState<OptionSetsData | null>(null);
  const toolbarRef = useRef<HTMLDivElement | null>(null);
  const helperRef = useRef<D365Helper | null>(null);

  // Create helper instance only once
  if (!helperRef.current) {
    helperRef.current = new D365Helper();
  }
  const helper = helperRef.current;

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

  const handleOpenSolutions = async () => {
    try {
      const url = await helper.getSolutionsUrl();
      if (url) {
        window.open(url, '_blank');
        showNotification('Opening solutions...');
      } else {
        showNotification('Unable to open solutions');
      }
    } catch (error) {
      showNotification('Error opening solutions');
    }
  };

  const handleOpenAdminCenter = () => {
    try {
      const url = helper.getAdminCenterUrl();
      window.open(url, '_blank');
      showNotification('Opening admin center...');
    } catch (error) {
      showNotification('Error opening admin center');
    }
  };

  const loadPluginTraceLogs = async (openModal: boolean = false) => {
    try {
      showNotification('Loading plugin trace logs...');
      const data = await helper.getPluginTraceLogs();
      setTraceLogData(data);
      if (openModal) {
        setShowTraceLogs(true);
      }
      showNotification('');
    } catch (error: any) {
      const message = error instanceof Error ? error.message : 'Error loading plugin trace logs';
      setTraceLogData({
        logs: [],
        error: message
      });
      setShowTraceLogs(true);
      showNotification('Error loading plugin trace logs');
    }
  };

  const handleShowTraceLogs = async () => {
    await loadPluginTraceLogs(true);
  };

  const handleRefreshTraceLogs = async () => {
    await loadPluginTraceLogs(false);
  };

  const handleCloseTraceLogs = () => {
    setShowTraceLogs(false);
    setTraceLogData(null);
  };

  const loadOptionSets = async (openModal: boolean = false) => {
    try {
      showNotification('Loading option sets...');
      const data = await helper.getOptionSets();
      setOptionSetData(data);
      if (openModal) {
        setShowOptionSets(true);
      }
      showNotification('');
    } catch (error: any) {
      const message = error instanceof Error ? error.message : 'Error loading option sets';
      setOptionSetData({
        attributes: [],
        error: message
      });
      setShowOptionSets(true);
      showNotification('Error loading option sets');
    }
  };

  const handleShowOptionSets = async () => {
    await loadOptionSets(true);
  };

  const handleRefreshOptionSets = async () => {
    await loadOptionSets(false);
  };

  const handleCloseOptionSets = () => {
    setShowOptionSets(false);
    setOptionSetData(null);
  };

  const handleUnlockFields = async () => {
    try {
      const count = await helper.unlockFields();
      showNotification(`Enabled editing for ${count} fields`);
    } catch (error) {
      showNotification('Error enabling field editing');
    }
  };

  const handleAutoFill = async () => {
    try {
      const count = await helper.autoFillForm();
      showNotification(`Added test data to ${count} fields`);
    } catch (error) {
      showNotification('Error adding test data');
    }
  };

  const handleEnableDevMode = async () => {
    try {
      showNotification('Activating Developer Mode...');
      await helper.toggleAllFields(true);
      await helper.toggleAllSections(true);
      const unlocked = await helper.unlockFields();
      const disabled = await helper.disableFieldRequirements();
      showNotification(`Dev Mode: enabled ${unlocked} fields and ${disabled} controls for testing.`);
    } catch (error) {
      showNotification('Error enabling Developer Mode');
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

  const handleCacheRefresh = () => {
    showNotification('Performing cache refresh...');
    // Trigger Ctrl+F5 (hard refresh)
    location.reload();
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
        <div className="d365-toolbar-logo-section">
          <img
            className="d365-toolbar-logo"
            src={chrome.runtime.getURL('icons/RG%20Logo_White_Stacked.svg')}
            alt="RG Logo"
            onError={(e) => {
              console.error('Logo failed to load:', chrome.runtime.getURL('icons/RG%20Logo_White_Stacked.svg'));
              e.currentTarget.style.display = 'none';
            }}
          />
        </div>
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
          <span className="d365-toolbar-section-label">Navigation:</span>
          <button
            className="d365-toolbar-btn"
            onClick={handleOpenSolutions}
            title="Open solutions page"
          >
            Solutions
          </button>
          <button
            className="d365-toolbar-btn"
            onClick={handleOpenAdminCenter}
            title="Open Power Platform admin center"
          >
            Admin Center
          </button>
        </div>

        <div className="d365-toolbar-section">
          <span className="d365-toolbar-section-label">Dev Tools:</span>
          <button
            className="d365-toolbar-btn"
            onClick={handleUnlockFields}
            title="Enable editing for development and testing"
          >
            Enable Editing
          </button>
          <button
            className="d365-toolbar-btn"
            onClick={handleAutoFill}
            title="Fill fields with test data for development"
          >
            Test Data
          </button>
          <button
            className="d365-toolbar-btn"
            onClick={handleEnableDevMode}
            title="Show all fields and enable all controls for development"
          >
            Dev Mode
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
            onClick={handleCacheRefresh}
            title="Perform hard refresh (Ctrl+F5) to clear cache"
          >
            Cache Refresh
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
            onClick={handleShowTraceLogs}
            title="View Plugin Trace Logs"
          >
            Trace Logs
          </button>
          <button
            className="d365-toolbar-btn"
            onClick={handleShowLibraries}
            title="View JavaScript libraries and event handlers"
          >
            JS Libraries
          </button>
          <button
            className="d365-toolbar-btn"
            onClick={handleShowOptionSets}
            title="View option set values used on this form"
          >
            Option Sets
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

      {showTraceLogs && (
        <PluginTraceLogViewer
          data={traceLogData}
          onClose={handleCloseTraceLogs}
          onRefresh={handleRefreshTraceLogs}
        />
      )}

      {showOptionSets && (
        <OptionSetsViewer
          data={optionSetData}
          onClose={handleCloseOptionSets}
          onRefresh={handleRefreshOptionSets}
        />
      )}
    </div>
  );
};

export default D365Toolbar;
