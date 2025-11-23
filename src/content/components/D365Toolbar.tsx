import React, { useState, useEffect, useRef, useCallback } from 'react';
import { D365Helper } from '../utils/D365Helper';
import { restoreShellContainerLayout, setShellContainerOffset } from '../utils/shellLayout';
import FormLibrariesAnalyzer from './FormLibrariesAnalyzer';
import PluginTraceLogViewer, { PluginTraceLogData } from './PluginTraceLogViewer';
import OptionSetsViewer, { OptionSetsData } from './OptionSetsViewer';
import ODataFieldsViewer, { ODataFieldsData } from './ODataFieldsViewer';
import AuditHistoryViewer, { AuditHistoryData } from './AuditHistoryViewer';
import QueryBuilder from '../../query-builder/components/QueryBuilder';

// Toolbar configuration types
type SectionId = 'fields' | 'sections' | 'schema' | 'navigation' | 'devtools' | 'tools';

interface ToolbarConfig {
  sectionOrder: SectionId[];
  buttonVisibility: Record<string, boolean>;
}

const DEFAULT_TOOLBAR_CONFIG: ToolbarConfig = {
  sectionOrder: ['fields', 'sections', 'schema', 'navigation', 'devtools', 'tools'],
  buttonVisibility: {}
};

const D365Toolbar: React.FC = () => {
  const [showSchemaNames, setShowSchemaNames] = useState(false);
  const [allFieldsVisible, setAllFieldsVisible] = useState(false);
  const [allSectionsVisible, setAllSectionsVisible] = useState(false);
  const [fieldsBlurred, setFieldsBlurred] = useState(false);
  const [devModeActive, setDevModeActive] = useState(false);
  const [notification, setNotification] = useState<string>('');
  const [showLibraries, setShowLibraries] = useState(false);
  const [librariesData, setLibrariesData] = useState<any>(null);
  const [showTraceLogs, setShowTraceLogs] = useState(false);
  const [traceLogData, setTraceLogData] = useState<PluginTraceLogData | null>(null);
  const [showOptionSets, setShowOptionSets] = useState(false);
  const [optionSetData, setOptionSetData] = useState<OptionSetsData | null>(null);
  const [showODataFields, setShowODataFields] = useState(false);
  const [odataFieldsData, setODataFieldsData] = useState<ODataFieldsData | null>(null);
  const [showAuditHistory, setShowAuditHistory] = useState(false);
  const [auditHistoryData, setAuditHistoryData] = useState<AuditHistoryData | null>(null);
  const [showQueryBuilder, setShowQueryBuilder] = useState(false);
  const [notificationDuration, setNotificationDuration] = useState(3);
  const [toolbarPosition, setToolbarPosition] = useState<'top' | 'bottom'>('top');
  const [traceLogLimit, setTraceLogLimit] = useState(20);
  const [toolbarConfig, setToolbarConfig] = useState<ToolbarConfig>(DEFAULT_TOOLBAR_CONFIG);
  const showSchemaNamesRef = useRef(showSchemaNames);
  const allFieldsVisibleRef = useRef(allFieldsVisible);
  const allSectionsVisibleRef = useRef(allSectionsVisible);
  const fieldsBlurredRef = useRef(fieldsBlurred);
  const toolbarRef = useRef<HTMLDivElement | null>(null);
  const helperRef = useRef<D365Helper | null>(null);

  // Create helper instance only once
  if (!helperRef.current) {
    helperRef.current = new D365Helper();
  }
  const helper = helperRef.current;

  // Load settings
  useEffect(() => {
    chrome.storage.sync.get(['notificationDuration', 'toolbarPosition', 'traceLogLimit', 'toolbarConfig'], (result) => {
      if (result.notificationDuration !== undefined) {
        setNotificationDuration(result.notificationDuration);
      }
      if (result.toolbarPosition !== undefined) {
        setToolbarPosition(result.toolbarPosition);
      }
      if (result.traceLogLimit !== undefined) {
        setTraceLogLimit(result.traceLogLimit);
      }
      if (result.toolbarConfig !== undefined) {
        setToolbarConfig(result.toolbarConfig);
      }
    });

    // Listen for setting changes
    const handleStorageChange = (changes: { [key: string]: chrome.storage.StorageChange }) => {
      if (changes.notificationDuration) {
        setNotificationDuration(changes.notificationDuration.newValue);
      }
      if (changes.toolbarPosition) {
        setToolbarPosition(changes.toolbarPosition.newValue);
      }
      if (changes.traceLogLimit) {
        setTraceLogLimit(changes.traceLogLimit.newValue);
      }
      if (changes.toolbarConfig) {
        setToolbarConfig(changes.toolbarConfig.newValue);
      }
    };

    chrome.storage.onChanged.addListener(handleStorageChange);

    return () => {
      chrome.storage.onChanged.removeListener(handleStorageChange);
    };
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
    allFieldsVisibleRef.current = allFieldsVisible;
  }, [allFieldsVisible]);

  useEffect(() => {
    allSectionsVisibleRef.current = allSectionsVisible;
  }, [allSectionsVisible]);

  useEffect(() => {
    fieldsBlurredRef.current = fieldsBlurred;
  }, [fieldsBlurred]);

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

    const resetFieldsVisibility = () => {
      if (!allFieldsVisibleRef.current) {
        return;
      }

      allFieldsVisibleRef.current = false;
      setAllFieldsVisible(false);
      helper.toggleAllFields(false).catch((error: any) => {
        console.warn('D365 Helper: Failed to reset fields visibility after navigation', error);
      });
    };

    const resetSectionsVisibility = () => {
      if (!allSectionsVisibleRef.current) {
        return;
      }

      allSectionsVisibleRef.current = false;
      setAllSectionsVisible(false);
      helper.toggleAllSections(false).catch((error: any) => {
        console.warn('D365 Helper: Failed to reset sections visibility after navigation', error);
      });
    };

    const resetFieldsBlur = () => {
      if (!fieldsBlurredRef.current) {
        return;
      }

      fieldsBlurredRef.current = false;
      setFieldsBlurred(false);
      helper.toggleBlurFields(false).catch((error: any) => {
        console.warn('D365 Helper: Failed to reset fields blur after navigation', error);
      });
    };

    const handlePotentialNavigation = () => {
      const currentUrl = window.location.href;
      if (currentUrl !== lastUrl) {
        lastUrl = currentUrl;
        resetSchemaOverlay();
        resetFieldsVisibility();
        resetSectionsVisibility();
        resetFieldsBlur();
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

  const handleToggleBlurFields = async () => {
    const newState = !fieldsBlurred;
    setFieldsBlurred(newState);
    try {
      await helper.toggleBlurFields(newState);
      showNotification(newState ? 'Fields blurred for privacy' : 'Field blur removed');
    } catch (error: any) {
      setFieldsBlurred(!newState); // Revert state on error
      const message = error?.message || 'Error toggling field blur';
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

  const handleOpenQueryBuilder = () => {
    setShowQueryBuilder(true);
    showNotification('Opening Query Builder...');
  };

  const handleCloseQueryBuilder = () => {
    setShowQueryBuilder(false);
  };

  const loadPluginTraceLogs = async (openModal: boolean = false) => {
    try {
      showNotification('Loading plugin trace logs...');
      const data = await helper.getPluginTraceLogs(traceLogLimit);
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

  const loadAuditHistory = async (showLoading: boolean) => {
    try {
      if (showLoading) {
        showNotification('Loading audit history...');
      }
      const data = await helper.getAuditHistory();
      setAuditHistoryData(data);
      setShowAuditHistory(true);
      if (showLoading) {
        showNotification('Audit history loaded');
      }
    } catch (error: any) {
      const message = error?.message || 'Error loading audit history';
      showNotification(message);
      setAuditHistoryData({ records: [], error: message });
      setShowAuditHistory(true);
    }
  };

  const handleShowAuditHistory = async () => {
    await loadAuditHistory(true);
  };

  const handleRefreshAuditHistory = async () => {
    await loadAuditHistory(false);
  };

  const handleCloseAuditHistory = () => {
    setShowAuditHistory(false);
    setAuditHistoryData(null);
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
    const newState = !devModeActive;

    try {
      if (newState) {
        // Activate Dev Mode
        showNotification('Activating Developer Mode...');
        await helper.toggleAllFields(true);
        await helper.toggleAllSections(true);
        setAllFieldsVisible(true);
        setAllSectionsVisible(true);
        const unlocked = await helper.unlockFields();
        const disabled = await helper.disableFieldRequirements();
        setDevModeActive(true);
        showNotification(`Dev Mode: enabled ${unlocked} fields and ${disabled} controls for testing.`);
      } else {
        // Deactivate Dev Mode - reset everything
        showNotification('Deactivating Developer Mode...');

        // Reset fields visibility
        await helper.toggleAllFields(false);
        setAllFieldsVisible(false);

        // Reset sections visibility
        await helper.toggleAllSections(false);
        setAllSectionsVisible(false);

        // Reset schema overlay
        if (showSchemaNames) {
          await helper.toggleSchemaOverlay(false);
          setShowSchemaNames(false);
        }

        setDevModeActive(false);
        showNotification('Dev Mode deactivated - form reset to original state');
      }
    } catch (error) {
      showNotification(newState ? 'Error enabling Developer Mode' : 'Error disabling Developer Mode');
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

  // Check if button should be visible based on config
  const isButtonVisible = (buttonId: string): boolean => {
    return toolbarConfig.buttonVisibility[buttonId] !== false;
  };

  // Render section based on section ID
  const renderSection = (sectionId: SectionId) => {
    switch (sectionId) {
      case 'fields':
        if (!isButtonVisible('fields.showAll')) return null;
        return (
          <div key="fields" className="d365-toolbar-section">
            <span className="d365-toolbar-section-label">Fields:</span>
            <button
              className={`d365-toolbar-btn ${allFieldsVisible ? 'd365-toolbar-btn-active' : ''}`}
              onClick={handleToggleFields}
              title="Toggle visibility of all fields on the form"
            >
              {allFieldsVisible ? 'Hide All' : 'Show All'}
            </button>
          </div>
        );

      case 'sections':
        if (!isButtonVisible('sections.showAll')) return null;
        return (
          <div key="sections" className="d365-toolbar-section">
            <span className="d365-toolbar-section-label">Sections:</span>
            <button
              className={`d365-toolbar-btn ${allSectionsVisible ? 'd365-toolbar-btn-active' : ''}`}
              onClick={handleToggleSections}
              title="Toggle visibility of all sections on the form"
            >
              {allSectionsVisible ? 'Hide All' : 'Show All'}
            </button>
          </div>
        );

      case 'schema':
        const showNamesVisible = isButtonVisible('schema.showNames');
        const copyAllVisible = isButtonVisible('schema.copyAll');
        if (!showNamesVisible && !copyAllVisible) return null;
        return (
          <div key="schema" className="d365-toolbar-section">
            <span className="d365-toolbar-section-label">Schema:</span>
            {showNamesVisible && (
              <button
                className={`d365-toolbar-btn ${showSchemaNames ? 'd365-toolbar-btn-active' : ''}`}
                onClick={handleToggleSchemaNames}
                title="Toggle schema name overlays on fields"
              >
                {showSchemaNames ? 'Hide Names' : 'Show Names'}
              </button>
            )}
            {copyAllVisible && (
              <button
                className="d365-toolbar-btn"
                onClick={handleCopyAllSchemaNames}
                title="Copy all schema names to clipboard"
              >
                Copy All
              </button>
            )}
          </div>
        );

      case 'navigation':
        const solutionsVisible = isButtonVisible('navigation.solutions');
        const adminCenterVisible = isButtonVisible('navigation.adminCenter');
        if (!solutionsVisible && !adminCenterVisible) return null;
        return (
          <div key="navigation" className="d365-toolbar-section">
            <span className="d365-toolbar-section-label">Navigation:</span>
            {solutionsVisible && (
              <button
                className="d365-toolbar-btn"
                onClick={handleOpenSolutions}
                title="Open solutions page"
              >
                Solutions
              </button>
            )}
            {adminCenterVisible && (
              <button
                className="d365-toolbar-btn"
                onClick={handleOpenAdminCenter}
                title="Open Power Platform admin center"
              >
                Admin Center
              </button>
            )}
          </div>
        );

      case 'devtools':
        const enableEditingVisible = isButtonVisible('devtools.enableEditing');
        const testDataVisible = isButtonVisible('devtools.testData');
        const devModeVisible = isButtonVisible('devtools.devMode');
        const blurFieldsVisible = isButtonVisible('devtools.blurFields');
        if (!enableEditingVisible && !testDataVisible && !devModeVisible && !blurFieldsVisible) return null;
        return (
          <div key="devtools" className="d365-toolbar-section">
            <span className="d365-toolbar-section-label">Dev Tools:</span>
            {enableEditingVisible && (
              <button
                className="d365-toolbar-btn"
                onClick={handleUnlockFields}
                title="Enable editing for development and testing"
              >
                Enable Editing
              </button>
            )}
            {testDataVisible && (
              <button
                className="d365-toolbar-btn"
                onClick={handleAutoFill}
                title="Fill fields with test data for development"
              >
                Test Data
              </button>
            )}
            {devModeVisible && (
              <button
                className={`d365-toolbar-btn ${devModeActive ? 'd365-toolbar-btn-active' : ''}`}
                onClick={handleEnableDevMode}
                title="Toggle Developer Mode - shows all fields, sections, and enables editing"
              >
                {devModeActive ? 'Deactivate' : 'Dev Mode'}
              </button>
            )}
            {blurFieldsVisible && (
              <button
                className={`d365-toolbar-btn ${fieldsBlurred ? 'd365-toolbar-btn-active' : ''}`}
                onClick={handleToggleBlurFields}
                title="Blur field values for privacy when sharing screen or taking screenshots"
              >
                {fieldsBlurred ? 'Unblur' : 'Blur Fields'}
              </button>
            )}
          </div>
        );

      case 'tools':
        const copyIdVisible = isButtonVisible('tools.copyId');
        const cacheRefreshVisible = isButtonVisible('tools.cacheRefresh');
        const webApiVisible = isButtonVisible('tools.webApi');
        const traceLogsVisible = isButtonVisible('tools.traceLogs');
        const jsLibrariesVisible = isButtonVisible('tools.jsLibraries');
        const optionSetsVisible = isButtonVisible('tools.optionSets');
        const odataFieldsVisible = isButtonVisible('tools.odataFields');
        const auditHistoryVisible = isButtonVisible('tools.auditHistory');
        const formEditorVisible = isButtonVisible('tools.formEditor');
        const queryBuilderVisible = isButtonVisible('tools.queryBuilder');
        const anyVisible = copyIdVisible || cacheRefreshVisible || webApiVisible || traceLogsVisible || 
                          jsLibrariesVisible || optionSetsVisible || odataFieldsVisible || auditHistoryVisible || formEditorVisible || queryBuilderVisible;
        if (!anyVisible) return null;
        return (
          <div key="tools" className="d365-toolbar-section">
            <span className="d365-toolbar-section-label">Tools:</span>
            {copyIdVisible && (
              <button
                className="d365-toolbar-btn"
                onClick={handleCopyRecordId}
                title="Copy current record ID"
              >
                Copy ID
              </button>
            )}
            {cacheRefreshVisible && (
              <button
                className="d365-toolbar-btn"
                onClick={handleCacheRefresh}
                title="Perform hard refresh (Ctrl+F5) to clear cache"
              >
                Cache Refresh
              </button>
            )}
            {webApiVisible && (
              <button
                className="d365-toolbar-btn"
                onClick={handleOpenWebAPI}
                title="Open Web API data in new tab"
              >
                Web API
              </button>
            )}
            {queryBuilderVisible && (
              <button
                className="d365-toolbar-btn"
                onClick={handleOpenQueryBuilder}
                title="Open Advanced Find"
              >
                Advanced Find
              </button>
            )}
            {traceLogsVisible && (
              <button
                className="d365-toolbar-btn"
                onClick={handleShowTraceLogs}
                title="View Plugin Trace Logs"
              >
                Trace Logs
              </button>
            )}
            {jsLibrariesVisible && (
              <button
                className="d365-toolbar-btn"
                onClick={handleShowLibraries}
                title="View JavaScript libraries and event handlers"
              >
                JS Libraries
              </button>
            )}
            {optionSetsVisible && (
              <button
                className="d365-toolbar-btn"
                onClick={handleShowOptionSets}
                title="View option set values used on this form"
              >
                Option Sets
              </button>
            )}
            {odataFieldsVisible && (
              <button
                className="d365-toolbar-btn"
                onClick={handleShowODataFields}
                title="View OData field metadata for this entity"
              >
                OData Fields
              </button>
            )}
            {auditHistoryVisible && (
              <button
                className="d365-toolbar-btn"
                onClick={handleShowAuditHistory}
                title="View audit history for this record"
              >
                Audit History
              </button>
            )}
            {formEditorVisible && (
              <button
                className="d365-toolbar-btn"
                onClick={handleOpenFormEditor}
                title="Open form editor"
              >
                Form Editor
              </button>
            )}
          </div>
        );

      default:
        return null;
    }
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

        {toolbarConfig.sectionOrder.map(sectionId => renderSection(sectionId))}

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

      {showAuditHistory && auditHistoryData && (
        <AuditHistoryViewer
          data={auditHistoryData}
          onClose={handleCloseAuditHistory}
          onRefresh={handleRefreshAuditHistory}
        />
      )}
      
      {showQueryBuilder && (
        <QueryBuilder
          orgUrl={helper.getOrgUrl()}
          onClose={handleCloseQueryBuilder}
        />
      )}
    </div>
  );
};

export default D365Toolbar;
