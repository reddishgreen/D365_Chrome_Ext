import React, { useState, useEffect, useRef, useCallback } from 'react';
import { D365Helper } from '../utils/D365Helper';
import { restoreShellContainerLayout, setShellContainerOffset } from '../utils/shellLayout';
import FormLibrariesAnalyzer from './FormLibrariesAnalyzer';
import PluginTraceLogViewer, { PluginTraceLogData } from './PluginTraceLogViewer';
import OptionSetsViewer, { OptionSetsData } from './OptionSetsViewer';
import ODataFieldsViewer, { ODataFieldsData } from './ODataFieldsViewer';

const D365Toolbar: React.FC = () => {
  const [showSchemaNames, setShowSchemaNames] = useState(false);
  const [allFieldsVisible, setAllFieldsVisible] = useState(false);
  const [allSectionsVisible, setAllSectionsVisible] = useState(false);
  const [notification, setNotification] = useState<string>('');
  const [showLibraries, setShowLibraries] = useState(false);
  const [librariesData, setLibrariesData] = useState<any>(null);
  const [showTraceLogs, setShowTraceLogs] = useState(false);
  const [traceLogData, setTraceLogData] = useState<PluginTraceLogData | null>(null);
  const [showOptionSets, setShowOptionSets] = useState(false);
  const [optionSetData, setOptionSetData] = useState<OptionSetsData | null>(null);
  const [showODataFields, setShowODataFields] = useState(false);
  const [odataFieldsData, setODataFieldsData] = useState<ODataFieldsData | null>(null);
  const [notificationDuration, setNotificationDuration] = useState(3);
  const [toolbarPosition, setToolbarPosition] = useState<'top' | 'bottom'>('top');
  const showSchemaNamesRef = useRef(showSchemaNames);
  const toolbarRef = useRef<HTMLDivElement | null>(null);
  const helperRef = useRef<D365Helper | null>(null);

  // Create helper instance only once
  if (!helperRef.current) {
    helperRef.current = new D365Helper();
  }
  const helper = helperRef.current;

  // Load settings
  useEffect(() => {
    chrome.storage.sync.get(['notificationDuration', 'toolbarPosition'], (result) => {
      if (result.notificationDuration !== undefined) {
        setNotificationDuration(result.notificationDuration);
      }
      if (result.toolbarPosition !== undefined) {
        setToolbarPosition(result.toolbarPosition);
      }
    });
  }, []);

  const updateShellOffset = useCallback(() => {
    const toolbarHeight = toolbarRef.current?.getBoundingClientRect().height;
    const fallbackHeight = 70;
    setShellContainerOffset(Math.round(toolbarHeight ?? fallbackHeight), toolbarPosition);
  }, [toolbarPosition]);

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

  useEffect(() => {
    showSchemaNamesRef.current = showSchemaNames;
  }, [showSchemaNames]);

  useEffect(() => {
    const { history } = window;
    let lastUrl = window.location.href;

    const resetSchemaOverlay = () => {
      if (!showSchemaNamesRef.current) {
        return;
      }

      showSchemaNamesRef.current = false;
      setShowSchemaNames(false);
      helper.toggleSchemaOverlay(false).catch((error) => {
        console.warn('D365 Helper: Failed to reset schema overlays after navigation', error);
      });
    };

    const handlePotentialNavigation = () => {
      const currentUrl = window.location.href;
      if (currentUrl !== lastUrl) {
        lastUrl = currentUrl;
        resetSchemaOverlay();
      }
    };

    const intervalId = window.setInterval(handlePotentialNavigation, 1000);

    window.addEventListener('popstate', handlePotentialNavigation);
    window.addEventListener('hashchange', handlePotentialNavigation);

    const originalPushState = history.pushState;
    const originalReplaceState = history.replaceState;

    const patchedPushState: History['pushState'] = function (
      this: History,
      data: any,
      unused: string,
      url?: string | URL | null
    ) {
      const unusedParam = typeof unused === 'string' ? unused : '';
      const result = originalPushState.call(this, data, unusedParam, url);
      handlePotentialNavigation();
      return result;
    };

    const patchedReplaceState: History['replaceState'] = function (
      this: History,
      data: any,
      unused: string,
      url?: string | URL | null
    ) {
      const unusedParam = typeof unused === 'string' ? unused : '';
      const result = originalReplaceState.call(this, data, unusedParam, url);
      handlePotentialNavigation();
      return result;
    };

    history.pushState = patchedPushState;
    history.replaceState = patchedReplaceState;

    return () => {
      window.clearInterval(intervalId);
      window.removeEventListener('popstate', handlePotentialNavigation);
      window.removeEventListener('hashchange', handlePotentialNavigation);
      history.pushState = originalPushState;
      history.replaceState = originalReplaceState;
    };
  }, [helper]);

  const showNotification = (message: string) => {
    setNotification(message);
    setTimeout(() => setNotification(''), notificationDuration * 1000);
  };

  const handleToggleFields = async () => {
    const newState = !allFieldsVisible;
    setAllFieldsVisible(newState);
    try {
      await helper.toggleAllFields(newState);
      showNotification(newState ? 'All fields shown' : 'Fields restored to original state');
    } catch (error: any) {
      setAllFieldsVisible(!newState); // Revert state on error
      const message = error?.message || 'Error toggling fields';
      showNotification(message);
    }
  };

  const handleToggleSections = async () => {
    const newState = !allSectionsVisible;
    setAllSectionsVisible(newState);
    try {
      await helper.toggleAllSections(newState);
      showNotification(newState ? 'All sections shown' : 'Sections restored to original state');
    } catch (error: any) {
      setAllSectionsVisible(!newState); // Revert state on error
      const message = error?.message || 'Error toggling sections';
      showNotification(message);
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

  const loadODataFields = async (openModal: boolean = false) => {
    try {
      showNotification('Loading OData fields...');
      const data = await helper.getODataFields();
      setODataFieldsData(data);
      if (openModal) {
        setShowODataFields(true);
      }
      showNotification('');
    } catch (error: any) {
      const message = error instanceof Error ? error.message : 'Error loading OData fields';
      setODataFieldsData({
        entityName: '',
        entitySetName: '',
        fields: [],
        error: message
      });
      setShowODataFields(true);
      showNotification('Error loading OData fields');
    }
  };

  const handleShowODataFields = async () => {
    await loadODataFields(true);
  };

  const handleRefreshODataFields = async () => {
    await loadODataFields(false);
  };

  const handleCloseODataFields = () => {
    setShowODataFields(false);
    setODataFieldsData(null);
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
      setAllFieldsVisible(true);
      setAllSectionsVisible(true);
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

  return (
    <div ref={toolbarRef} className={`d365-toolbar d365-toolbar-${toolbarPosition}`}>
      <div className="d365-toolbar-content">
        <div className="d365-toolbar-logo-section">
          <a
            href="https://www.reddishgreen.co.uk"
            target="_blank"
            rel="noopener noreferrer"
            style={{ display: 'inline-block', lineHeight: 0 }}
          >
            <img
              className="d365-toolbar-logo"
              src={chrome.runtime.getURL('icons/RG%20Logo_White_Stacked.svg')}
              alt="RG Logo"
              onError={(e) => {
                console.error('Logo failed to load:', chrome.runtime.getURL('icons/RG%20Logo_White_Stacked.svg'));
                e.currentTarget.style.display = 'none';
              }}
            />
          </a>
        </div>

        <div className="d365-toolbar-section">
          <span className="d365-toolbar-section-label">Fields:</span>
          <button
            className={`d365-toolbar-btn ${allFieldsVisible ? 'd365-toolbar-btn-active' : ''}`}
            onClick={handleToggleFields}
            title="Toggle visibility of all fields on the form"
          >
            {allFieldsVisible ? 'Hide All' : 'Show All'}
          </button>
        </div>

        <div className="d365-toolbar-section">
          <span className="d365-toolbar-section-label">Sections:</span>
          <button
            className={`d365-toolbar-btn ${allSectionsVisible ? 'd365-toolbar-btn-active' : ''}`}
            onClick={handleToggleSections}
            title="Toggle visibility of all sections on the form"
          >
            {allSectionsVisible ? 'Hide All' : 'Show All'}
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
          <button
            className="d365-toolbar-btn"
            onClick={handleShowODataFields}
            title="View OData field metadata for this entity"
          >
            OData Fields
          </button>
        </div>

        <div className="d365-toolbar-section d365-toolbar-actions">
          <span className="d365-toolbar-version-badge">v{chrome.runtime.getManifest().version}</span>
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

      {showODataFields && (
        <ODataFieldsViewer
          data={odataFieldsData}
          onClose={handleCloseODataFields}
          onRefresh={handleRefreshODataFields}
        />
      )}
    </div>
  );
};

export default D365Toolbar;

