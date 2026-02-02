import React, { useState, useEffect, useRef, useCallback } from 'react';
import { D365Helper } from '../utils/D365Helper';
import { restoreShellContainerLayout, setShellContainerOffset } from '../utils/shellLayout';
import FormLibrariesAnalyzer from './FormLibrariesAnalyzer';
import PluginTraceLogViewer, { PluginTraceLogData } from './PluginTraceLogViewer';
import OptionSetsViewer, { OptionSetsData } from './OptionSetsViewer';
import ODataFieldsViewer, { ODataFieldsData } from './ODataFieldsViewer';
import AuditHistoryViewer, { AuditHistoryData } from './AuditHistoryViewer';
import QueryBuilder from '../../query-builder/components/QueryBuilder';
import ImpersonationSelector, { ImpersonationData, SystemUser } from './ImpersonationSelector';
import PromptMakerViewer from './PromptMakerViewer';

// Toolbar configuration types
type SectionId = 'fields' | 'sections' | 'schema' | 'navigation' | 'devtools' | 'tools';

interface ToolbarConfig {
  sectionOrder: SectionId[];
  buttonVisibility: Record<string, boolean>;
  sectionLabels?: Partial<Record<SectionId, string>>;
  sectionButtons?: Partial<Record<SectionId, string[]>>;
}

const SECTION_IDS: SectionId[] = ['fields', 'sections', 'schema', 'navigation', 'devtools', 'tools'];

const DEFAULT_SECTION_LABELS: Record<SectionId, string> = {
  fields: 'Fields',
  sections: 'Sections',
  schema: 'Schema',
  navigation: 'Navigation',
  devtools: 'Dev Tools',
  tools: 'Tools'
};

const DEFAULT_SECTION_BUTTONS: Record<SectionId, string[]> = {
  fields: ['fields.showAll'],
  sections: ['sections.showAll'],
  schema: ['schema.showNames', 'schema.copyAll'],
  navigation: ['navigation.solutions', 'navigation.adminCenter'],
  devtools: ['devtools.enableEditing', 'devtools.testData', 'devtools.devMode', 'devtools.blurFields', 'devtools.impersonate'],
  tools: [
    'tools.copyId',
    'tools.cacheRefresh',
    'tools.webApi',
    'tools.queryBuilder',
    'tools.traceLogs',
    'tools.jsLibraries',
    'tools.optionSets',
    'tools.odataFields',
    'tools.auditHistory',
    'tools.formEditor',
    'tools.promptMaker'
  ]
};

const DEFAULT_TOOLBAR_CONFIG: ToolbarConfig = {
  sectionOrder: ['devtools', 'tools', 'sections', 'fields', 'schema', 'navigation'],
  buttonVisibility: {
    'fields.showAll': true,
    'sections.showAll': true,
    'schema.showNames': true,
    'schema.copyAll': true,
    'navigation.solutions': true,
    'navigation.adminCenter': true,
    'devtools.enableEditing': false,
    'devtools.testData': true,
    'devtools.devMode': true,
    'devtools.blurFields': false,
    'devtools.impersonate': true,
    'tools.copyId': true,
    'tools.cacheRefresh': true,
    'tools.webApi': true,
    'tools.queryBuilder': true,
    'tools.traceLogs': true,
    'tools.jsLibraries': true,
    'tools.optionSets': true,
    'tools.odataFields': true,
    'tools.auditHistory': true,
    'tools.formEditor': false,
    'tools.promptMaker': true
  },
  sectionLabels: {
    devtools: 'Form Controls',
    sections: 'Form',
    tools: 'Dev Tools'
  },
  sectionButtons: {
    devtools: [
      'devtools.enableEditing',
      'devtools.devMode',
      'devtools.testData',
      'schema.copyAll',
      'devtools.blurFields',
      'devtools.impersonate',
      'tools.copyId'
    ],
    tools: [
      'tools.webApi',
      'tools.traceLogs',
      'tools.queryBuilder',
      'tools.optionSets',
      'tools.cacheRefresh',
      'tools.jsLibraries',
      'tools.odataFields',
      'tools.auditHistory',
      'tools.formEditor'
    ],
    sections: ['sections.showAll'],
    fields: ['fields.showAll'],
    schema: ['schema.showNames'],
    navigation: ['navigation.solutions', 'navigation.adminCenter']
  }
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
  const [showPromptMaker, setShowPromptMaker] = useState(false);
  const [showOptionSets, setShowOptionSets] = useState(false);
  const [optionSetData, setOptionSetData] = useState<OptionSetsData | null>(null);
  const [showODataFields, setShowODataFields] = useState(false);
  const [odataFieldsData, setODataFieldsData] = useState<ODataFieldsData | null>(null);
  const [showAuditHistory, setShowAuditHistory] = useState(false);
  const [auditHistoryData, setAuditHistoryData] = useState<AuditHistoryData | null>(null);
  const [showQueryBuilder, setShowQueryBuilder] = useState(false);
  const [notificationDuration, setNotificationDuration] = useState(3);
  const [toolbarPosition, setToolbarPosition] = useState<'top' | 'bottom'>('bottom');
  const [traceLogLimit, setTraceLogLimit] = useState(20);
  const [toolbarConfig, setToolbarConfig] = useState<ToolbarConfig>(DEFAULT_TOOLBAR_CONFIG);
  const [showImpersonation, setShowImpersonation] = useState(false);
  const [impersonatedUser, setImpersonatedUser] = useState<SystemUser | null>(null);
  const [impersonationData, setImpersonationData] = useState<ImpersonationData | null>(null);
  const showSchemaNamesRef = useRef(showSchemaNames);
  const allFieldsVisibleRef = useRef(allFieldsVisible);
  const allSectionsVisibleRef = useRef(allSectionsVisible);
  const fieldsBlurredRef = useRef(fieldsBlurred);
  const devModeActiveRef = useRef(devModeActive);
  const toolbarRef = useRef<HTMLDivElement | null>(null);
  const helperRef = useRef<D365Helper | null>(null);
  const applyDevModeRef = useRef<((showNotification?: boolean) => Promise<void>) | null>(null);
  const removeDevModeRef = useRef<((showNotification?: boolean) => Promise<void>) | null>(null);

  // Create helper instance only once
  if (!helperRef.current) {
    helperRef.current = new D365Helper();
  }
  const helper = helperRef.current;

  // Load settings
  useEffect(() => {
    const normalizeToolbarConfig = (config?: Partial<ToolbarConfig> | null): ToolbarConfig => {
      const baseOrder = DEFAULT_TOOLBAR_CONFIG.sectionOrder;
      const rawOrder = Array.isArray(config?.sectionOrder) ? (config!.sectionOrder as SectionId[]) : baseOrder;

      const validOrder = rawOrder.filter((id) => SECTION_IDS.includes(id));
      const orderSet = new Set(validOrder);
      const sectionOrder: SectionId[] = [...validOrder, ...baseOrder.filter((id) => !orderSet.has(id))];

      const buttonVisibility = {
        ...DEFAULT_TOOLBAR_CONFIG.buttonVisibility,
        ...(config?.buttonVisibility || {})
      };

      const sectionLabels: Partial<Record<SectionId, string>> = {};
      const rawLabels: any = (config as any)?.sectionLabels;
      if (rawLabels && typeof rawLabels === 'object') {
        SECTION_IDS.forEach((id) => {
          const value = rawLabels[id];
          if (typeof value === 'string') {
            sectionLabels[id] = value.trim();
          }
        });
      }

      // Normalize section button layout (move/reorder buttons between sections)
      const normalizedSectionButtons: Record<SectionId, string[]> = SECTION_IDS.reduce((acc, id) => {
        acc[id] = [];
        return acc;
      }, {} as Record<SectionId, string[]>);

      const rawSectionButtons: any = (config as any)?.sectionButtons;
      const seen = new Set<string>();
      const allButtonIds = Object.values(DEFAULT_SECTION_BUTTONS).flat();

      if (rawSectionButtons && typeof rawSectionButtons === 'object') {
        SECTION_IDS.forEach((sectionId) => {
          const list = rawSectionButtons[sectionId];
          if (!Array.isArray(list)) return;

          list.forEach((buttonId: any) => {
            if (typeof buttonId !== 'string') return;
            if (!allButtonIds.includes(buttonId)) return;
            if (seen.has(buttonId)) return;
            normalizedSectionButtons[sectionId].push(buttonId);
            seen.add(buttonId);
          });
        });
      }

      // Add any missing buttons back to their default section
      allButtonIds.forEach((buttonId) => {
        if (seen.has(buttonId)) return;
        const defaultSectionId = (Object.entries(DEFAULT_SECTION_BUTTONS).find(([, ids]) => ids.includes(buttonId))?.[0] ||
          'tools') as SectionId;
        normalizedSectionButtons[defaultSectionId].push(buttonId);
        seen.add(buttonId);
      });

      return { sectionOrder, buttonVisibility, sectionLabels, sectionButtons: normalizedSectionButtons };
    };

    chrome.storage.sync.get(['notificationDuration', 'toolbarPosition', 'traceLogLimit', 'toolbarConfig', 'devModeActive', 'devModePersist'], (result) => {
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
        setToolbarConfig(normalizeToolbarConfig(result.toolbarConfig));
      }
      // Load Dev Mode state from storage and apply if enabled
      // Only restore Dev Mode if persistence is on (default true)
      const persistEnabled = result.devModePersist !== false;
      if (result.devModeActive === true && persistEnabled) {
        setDevModeActive(true);
        devModeActiveRef.current = true;
      } else if (result.devModeActive === true && !persistEnabled) {
        // Persist is off - clear the stored state so it doesn't linger
        chrome.storage.sync.set({ devModeActive: false });
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
        setToolbarConfig(normalizeToolbarConfig(changes.toolbarConfig.newValue));
      }
      if (changes.devModeActive !== undefined) {
        const newDevModeState = changes.devModeActive.newValue === true;
        if (newDevModeState !== devModeActive) {
          if (newDevModeState) {
            void applyDevMode(false);
          } else {
            void removeDevMode(false);
          }
        }
      }
    };

    chrome.storage.onChanged.addListener(handleStorageChange);

    return () => {
      chrome.storage.onChanged.removeListener(handleStorageChange);
    };
  }, []);

  const getSectionLabel = useCallback(
    (sectionId: SectionId, fallback: string) => {
      const override = toolbarConfig.sectionLabels?.[sectionId];
      const base = override !== undefined ? override.trim() : (fallback || '').trim();
      if (!base) return '';
      return base.endsWith(':') ? base : `${base}:`;
    },
    [toolbarConfig.sectionLabels]
  );

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
    devModeActiveRef.current = devModeActive;
  }, [devModeActive]);

  // Wait for form to be ready and apply Dev Mode if it was previously enabled
  useEffect(() => {
    let checkInterval: number | undefined;
    let maxAttempts = 30; // Try for up to 15 seconds (30 * 500ms)
    let attempts = 0;
    let hasApplied = false; // Track if we've already applied to avoid duplicates

    const checkAndApplyDevMode = async () => {
      try {
        // Check if form is ready by trying to get record ID
        try {
          await helper.getRecordId();
          // Form is ready, now check storage
          const result = await chrome.storage.sync.get(['devModeActive']);
          console.log('[D365 Helper] Dev Mode storage check:', result.devModeActive);
          
          if (result.devModeActive === true) {
            // Apply if storage says it should be active (user had it on when navigating between forms)
            if (!hasApplied && applyDevModeRef.current) {
              console.log('[D365 Helper] Reapplying Dev Mode from storage...');
              hasApplied = true;
              setDevModeActive(true);
              devModeActiveRef.current = true;
              // Wait a bit for D365 to finish initializing
              await new Promise(resolve => setTimeout(resolve, 1000));
              // Apply Dev Mode
              try {
                await applyDevModeRef.current(false);
                // Wait and reapply schema overlay in case D365 reset it
                await new Promise(resolve => setTimeout(resolve, 500));
                try {
                  await helper.toggleSchemaOverlay(true);
                  setShowSchemaNames(true);
                  showSchemaNamesRef.current = true;
                  console.log('[D365 Helper] Schema overlay reapplied after initial load');
                } catch (error) {
                  console.warn('[D365 Helper] Failed to reapply schema overlay', error);
                }
                console.log('[D365 Helper] Dev Mode applied successfully');
              } catch (error) {
                console.error('[D365 Helper] Failed to apply Dev Mode:', error);
                hasApplied = false; // Allow retry
              }
            }
            // Clear interval once we've applied
            if (hasApplied && checkInterval) {
              clearInterval(checkInterval);
              checkInterval = undefined;
            }
          } else {
            // Dev Mode is off in storage, make sure state matches
            if (devModeActive || devModeActiveRef.current) {
              console.log('[D365 Helper] Dev Mode is off in storage, resetting state');
              setDevModeActive(false);
              devModeActiveRef.current = false;
            }
            if (checkInterval) {
              clearInterval(checkInterval);
              checkInterval = undefined;
            }
          }
        } catch (error) {
          // Form not ready yet, continue checking
          attempts++;
          if (attempts >= maxAttempts) {
            console.warn('[D365 Helper] Form not ready after max attempts, giving up');
            // Give up after max attempts
            if (checkInterval) {
              clearInterval(checkInterval);
              checkInterval = undefined;
            }
          }
        }
      } catch (error) {
        console.warn('[D365 Helper] Failed to check Dev Mode state', error);
      }
    };

    // Start checking immediately, then every 500ms
    checkAndApplyDevMode();
    checkInterval = window.setInterval(checkAndApplyDevMode, 500);

    return () => {
      if (checkInterval) {
        clearInterval(checkInterval);
      }
    };
  }, []); // Only run once on mount

  useEffect(() => {
    const { history } = window;
    let lastUrl = window.location.href;
    let devModeReapplyTimeout: number | undefined;
    let devModeReapplyAttempts = 0;
    let devModeReapplyInProgress = false;
    const MAX_DEV_MODE_REAPPLY_ATTEMPTS = 3;

    const clearDevModeReapplyTimeout = () => {
      if (devModeReapplyTimeout !== undefined) {
        window.clearTimeout(devModeReapplyTimeout);
        devModeReapplyTimeout = undefined;
      }
    };

    const scheduleDevModeReapply = (delayMs: number = 1500) => {
      clearDevModeReapplyTimeout();
      devModeReapplyTimeout = window.setTimeout(() => {
        void reapplyDevModeAfterNavigation();
      }, delayMs);
    };

    // Helper function to check if current page is a form
    const isCurrentPageAForm = (): boolean => {
      const xrm = (window as any).Xrm;
      if (xrm && xrm.Page && xrm.Page.data && xrm.Page.data.entity) {
        return true;
      }
      const url = window.location.href;
      if (url.includes('pagetype=entityrecord') ||
          url.includes('etn=') ||
          url.includes('extraqs=') ||
          url.includes('formid=')) {
        return true;
      }
      return false;
    };

    const reapplyDevModeAfterNavigation = async () => {
      console.log('[D365 Helper] Reapplying Dev Mode after navigation...');
      // Check storage to see if Dev Mode should be active
      chrome.storage.sync.get(['devModeActive'], async (result) => {
        const shouldBeActive = result.devModeActive === true;
        console.log('[D365 Helper] Reapply check - Dev Mode should be active:', shouldBeActive);
        
        if (!shouldBeActive) {
          // Dev Mode is off in storage, make sure UI matches
          console.log('[D365 Helper] Dev Mode is off, resetting state');
          if (devModeActiveRef.current) {
            devModeActiveRef.current = false;
            setDevModeActive(false);
          }
          return;
        }
        
        // Dev Mode should be active, reapply it
        if (devModeReapplyInProgress) {
          console.log('[D365 Helper] Reapply already in progress, skipping');
          return;
        }
        devModeReapplyInProgress = true;

        // Wait for form to be ready before applying
        let formReady = false;
        let readyAttempts = 0;
        const maxReadyAttempts = 30;

        console.log('[D365 Helper] Waiting for form to be ready...');
        while (!formReady && readyAttempts < maxReadyAttempts) {
          try {
            await helper.getRecordId();
            formReady = true;
            console.log('[D365 Helper] Form is ready after', readyAttempts, 'attempts');
          } catch (error) {
            readyAttempts++;
            await new Promise(resolve => setTimeout(resolve, 300));
          }
        }

        if (!formReady) {
          console.warn('[D365 Helper] Form not ready after navigation, will retry');
          devModeReapplyInProgress = false;
          scheduleDevModeReapply(1500);
          return;
        }

        // Form is ready, wait a bit more for D365 to finish initializing
        await new Promise(resolve => setTimeout(resolve, 1000));
        
        // Form is ready, apply Dev Mode
        if (applyDevModeRef.current) {
          try {
            console.log('[D365 Helper] Applying Dev Mode after navigation...');
            await applyDevModeRef.current(false);
            // Wait a bit and reapply schema overlay in case D365 reset it
            await new Promise(resolve => setTimeout(resolve, 500));
            try {
              await helper.toggleSchemaOverlay(true);
              setShowSchemaNames(true);
              showSchemaNamesRef.current = true;
              console.log('[D365 Helper] Schema overlay reapplied after delay');
            } catch (error) {
              console.warn('[D365 Helper] Failed to reapply schema overlay', error);
            }
            console.log('[D365 Helper] Dev Mode reapplied successfully');
            devModeReapplyAttempts = 0;
            devModeReapplyInProgress = false;
          } catch (error) {
            console.error('[D365 Helper] Failed to reapply Dev Mode:', error);
            devModeReapplyAttempts += 1;
            if (devModeReapplyAttempts <= MAX_DEV_MODE_REAPPLY_ATTEMPTS) {
              // Form may still be loading; retry with a small backoff.
              const backoff = 600 + devModeReapplyAttempts * 500;
              console.log('[D365 Helper] Retrying Dev Mode reapply in', backoff, 'ms');
              scheduleDevModeReapply(backoff);
              devModeReapplyInProgress = false;
            } else {
              // Give up and reset Dev Mode so the UI never stays "stuck" active.
              console.warn('[D365 Helper] Dev Mode reapply failed after navigation; resetting Dev Mode.', error);
              if (removeDevModeRef.current) {
                await removeDevModeRef.current(false);
              }
              devModeReapplyInProgress = false;
            }
          }
        } else {
          console.warn('[D365 Helper] applyDevModeRef.current is null, cannot reapply');
          devModeReapplyInProgress = false;
        }
      });
    };

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
        console.log('[D365 Helper] Navigation detected to:', currentUrl);
        
        // Check if we're still on a form page
        const isFormPage = isCurrentPageAForm();
        console.log('[D365 Helper] Is current page a form?', isFormPage);
        
        if (!isFormPage) {
          // Navigated to a non-form page (grid, dashboard, etc.) - deactivate Dev Mode
          console.log('[D365 Helper] Navigated to non-form page, deactivating Dev Mode');
          if (devModeActiveRef.current && removeDevModeRef.current) {
            removeDevModeRef.current(false).catch((error) => {
              console.error('[D365 Helper] Failed to deactivate Dev Mode:', error);
            });
          } else {
            resetSchemaOverlay();
            resetFieldsVisibility();
            resetSectionsVisibility();
          }
          resetFieldsBlur();
        } else {
          // Still on a form page - check if Dev Mode should stay active
          chrome.storage.sync.get(['devModeActive', 'devModePersist'], (result) => {
            const shouldBeActive = result.devModeActive === true;
            const persistEnabled = result.devModePersist !== false; // default true for backwards compat
            if (shouldBeActive && persistEnabled) {
              // Keep Dev Mode on, but re-apply it for the newly loaded form.
              console.log('[D365 Helper] Staying on form page, reapplying Dev Mode');
              scheduleDevModeReapply(1500);
            } else {
              // Persistence is off or Dev Mode is inactive - deactivate
              if (shouldBeActive && !persistEnabled && devModeActiveRef.current && removeDevModeRef.current) {
                console.log('[D365 Helper] Dev Mode persist is off, deactivating on navigation');
                removeDevModeRef.current(true).catch((error) => {
                  console.error('[D365 Helper] Failed to deactivate Dev Mode:', error);
                });
              } else {
                resetSchemaOverlay();
                resetFieldsVisibility();
                resetSectionsVisibility();
              }
            }
          });
          resetFieldsBlur();
        }
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
      clearDevModeReapplyTimeout();
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

  const handleClearTraceLogs = async () => {
    // Non-destructive: just clear the UI list until user refreshes
    setTraceLogData({ logs: [], moreRecords: false });
    showNotification('Cleared trace log view (records not deleted). Click refresh to reload.');
  };

  const handleCloseTraceLogs = () => {
    setShowTraceLogs(false);
    setTraceLogData(null);
  };

  const handleOpenPromptMaker = () => {
    setShowPromptMaker(true);
    showNotification('Opening AI Prompt Maker...');
  };

  const handleClosePromptMaker = () => {
    setShowPromptMaker(false);
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

  // Helper function to apply Dev Mode
  const applyDevMode = useCallback(async (showNotificationMessage: boolean = true) => {
    if (showNotificationMessage) {
      showNotification('Activating Developer Mode...');
    }
    console.log('[D365 Helper] applyDevMode: Starting...');
    
    try {
      await helper.toggleAllFields(true);
      console.log('[D365 Helper] applyDevMode: Fields toggled');
    } catch (error) {
      console.error('[D365 Helper] applyDevMode: Failed to toggle fields', error);
    }
    
    try {
      await helper.toggleAllSections(true);
      console.log('[D365 Helper] applyDevMode: Sections toggled');
    } catch (error) {
      console.error('[D365 Helper] applyDevMode: Failed to toggle sections', error);
    }
    
    setAllFieldsVisible(true);
    setAllSectionsVisible(true);
    allFieldsVisibleRef.current = true;
    allSectionsVisibleRef.current = true;
    
    try {
      const unlocked = await helper.unlockFields();
      console.log('[D365 Helper] applyDevMode: Unlocked', unlocked, 'fields');
    } catch (error) {
      console.error('[D365 Helper] applyDevMode: Failed to unlock fields', error);
    }
    
    try {
      const disabled = await helper.disableFieldRequirements();
      console.log('[D365 Helper] applyDevMode: Disabled', disabled, 'required fields');
    } catch (error) {
      console.error('[D365 Helper] applyDevMode: Failed to disable requirements', error);
    }
    
    // Wait longer for DOM to update after showing fields/sections
    // D365 may need more time to fully render the form
    await new Promise(resolve => setTimeout(resolve, 800));
    
    // Activate schema names overlay after fields are visible
    // Always apply, don't check ref - D365 may have reset it
    try {
      await helper.toggleSchemaOverlay(true);
      setShowSchemaNames(true);
      showSchemaNamesRef.current = true;
      console.log('[D365 Helper] applyDevMode: Schema overlay applied');
    } catch (error) {
      console.error('[D365 Helper] applyDevMode: Failed to toggle schema overlay', error);
    }
    
    setDevModeActive(true);
    devModeActiveRef.current = true;
    // Save to storage
    chrome.storage.sync.set({ devModeActive: true });
    
    if (showNotificationMessage) {
      showNotification(`Dev Mode: enabled fields and controls for testing.`);
    }
    
    console.log('[D365 Helper] applyDevMode: Complete');
  }, [helper]);

  // Store refs to functions for use in effects
  useEffect(() => {
    applyDevModeRef.current = applyDevMode;
  }, [applyDevMode]);

  // Helper function to remove Dev Mode
  const removeDevMode = useCallback(async (showNotificationMessage: boolean = true) => {
    if (showNotificationMessage) {
      showNotification('Deactivating Developer Mode...');
    }

    // Reset fields visibility
    await helper.toggleAllFields(false);
    setAllFieldsVisible(false);
    allFieldsVisibleRef.current = false;

    // Reset sections visibility
    await helper.toggleAllSections(false);
    setAllSectionsVisible(false);
    allSectionsVisibleRef.current = false;

    // Reset schema overlay
    if (showSchemaNamesRef.current) {
      await helper.toggleSchemaOverlay(false);
      setShowSchemaNames(false);
      showSchemaNamesRef.current = false;
    }

    setDevModeActive(false);
    devModeActiveRef.current = false;
    // Save to storage
    chrome.storage.sync.set({ devModeActive: false });
    
    if (showNotificationMessage) {
      showNotification('Dev Mode deactivated - form reset to original state');
    }
  }, [helper]);

  // Store refs to functions for use in effects
  useEffect(() => {
    removeDevModeRef.current = removeDevMode;
  }, [removeDevMode]);

  const handleEnableDevMode = async () => {
    const newState = !devModeActive;

    try {
      if (newState) {
        await applyDevMode(true);
      } else {
        await removeDevMode(true);
      }
    } catch (error) {
      showNotification(newState ? 'Error enabling Developer Mode' : 'Error disabling Developer Mode');
    }
  };

  // ===== IMPERSONATION HANDLERS =====
  
  const loadImpersonationUsers = async () => {
    try {
      setImpersonationData(null); // Show loading state
      const data = await helper.getSystemUsers();
      setImpersonationData(data);
    } catch (error: any) {
      setImpersonationData({ users: [], error: error.message || 'Failed to load users' });
    }
  };

  const handleShowImpersonation = async () => {
    setShowImpersonation(true);
    await loadImpersonationUsers();
  };

  const handleCloseImpersonation = () => {
    setShowImpersonation(false);
    setImpersonationData(null);
  };

  const handleSelectImpersonation = async (user: SystemUser) => {
    try {
      showNotification(`Impersonating ${user.fullname}...`);
      await helper.setImpersonation(user.systemuserid, user.fullname, user.domainname);
      setImpersonatedUser(user);
      setShowImpersonation(false);
      setImpersonationData(null);
      showNotification(`Now impersonating: ${user.fullname}`);
    } catch (error) {
      showNotification('Error setting impersonation');
    }
  };

  const handleCancelImpersonation = async () => {
    try {
      await helper.clearImpersonation();
      const previousUser = impersonatedUser?.fullname;
      setImpersonatedUser(null);
      showNotification(`Stopped impersonating ${previousUser || 'user'}`);
    } catch (error) {
      showNotification('Error clearing impersonation');
    }
  };

  // Check for existing impersonation on mount (silent - may fail on non-form pages)
  useEffect(() => {
    const checkImpersonationStatus = async () => {
      const status = await helper.getImpersonationStatus();
      if (status.isImpersonating && status.user) {
        setImpersonatedUser(status.user);
      }
    };
    // Fire and forget - errors are handled silently in getImpersonationStatus
    checkImpersonationStatus();
  }, []);

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

  const getButtonsForSection = useCallback(
    (sectionId: SectionId): string[] => {
      const list = toolbarConfig.sectionButtons?.[sectionId];
      // IMPORTANT: empty arrays are valid (section intentionally has no buttons)
      if (Array.isArray(list)) return list;
      return DEFAULT_SECTION_BUTTONS[sectionId] || [];
    },
    [toolbarConfig.sectionButtons]
  );

  const renderButton = (buttonId: string): JSX.Element | null => {
    // Impersonation: always show indicator when active (even if user hid the button)
    if (buttonId === 'devtools.impersonate') {
      if (impersonatedUser) {
        return (
          <div key={buttonId} className="d365-impersonate-indicator">
            <span className="d365-impersonate-indicator-label">Acting as:</span>
            <span className="d365-impersonate-indicator-user">{impersonatedUser.fullname}</span>
            <button
              className="d365-impersonate-indicator-cancel"
              onClick={handleCancelImpersonation}
              title="Stop impersonating"
            >
              ✕
            </button>
          </div>
        );
      }

      if (!isButtonVisible(buttonId)) return null;
      return (
        <button
          key={buttonId}
          className="d365-toolbar-btn d365-toolbar-btn-impersonate"
          onClick={handleShowImpersonation}
          title="Impersonate another user for API calls"
        >
          Impersonate
        </button>
      );
    }

    if (!isButtonVisible(buttonId)) return null;

    switch (buttonId) {
      case 'fields.showAll':
        return (
          <button
            key={buttonId}
            className={`d365-toolbar-btn ${allFieldsVisible ? 'd365-toolbar-btn-active' : ''}`}
            onClick={handleToggleFields}
            title="Toggle visibility of all fields on the form"
          >
            {allFieldsVisible ? 'Hide All' : 'Show All'}
          </button>
        );

      case 'sections.showAll':
        return (
          <button
            key={buttonId}
            className={`d365-toolbar-btn ${allSectionsVisible ? 'd365-toolbar-btn-active' : ''}`}
            onClick={handleToggleSections}
            title="Toggle visibility of all sections on the form"
          >
            {allSectionsVisible ? 'Hide All' : 'Show All'}
          </button>
        );

      case 'schema.showNames':
        return (
          <button
            key={buttonId}
            className={`d365-toolbar-btn ${showSchemaNames ? 'd365-toolbar-btn-active' : ''}`}
            onClick={handleToggleSchemaNames}
            title="Toggle schema name overlays on fields"
          >
            {showSchemaNames ? 'Hide Names' : 'Show Names'}
          </button>
        );

      case 'schema.copyAll':
        return (
          <button
            key={buttonId}
            className="d365-toolbar-btn"
            onClick={handleCopyAllSchemaNames}
            title="Copy all schema names to clipboard"
          >
            Copy All
          </button>
        );

      case 'navigation.solutions':
        return (
          <button key={buttonId} className="d365-toolbar-btn" onClick={handleOpenSolutions} title="Open solutions page">
            Solutions
          </button>
        );

      case 'navigation.adminCenter':
        return (
          <button
            key={buttonId}
            className="d365-toolbar-btn"
            onClick={handleOpenAdminCenter}
            title="Open Power Platform admin center"
          >
            Admin Center
          </button>
        );

      case 'devtools.enableEditing':
        return (
          <button
            key={buttonId}
            className="d365-toolbar-btn"
            onClick={handleUnlockFields}
            title="Enable editing for development and testing"
          >
            Enable Editing
          </button>
        );

      case 'devtools.testData':
        return (
          <button
            key={buttonId}
            className="d365-toolbar-btn"
            onClick={handleAutoFill}
            title="Fill fields with test data for development"
          >
            Test Data
          </button>
        );

      case 'devtools.devMode':
        return (
          <button
            key={buttonId}
            className={`d365-toolbar-btn d365-toolbar-btn-dev-mode ${devModeActive ? 'd365-toolbar-btn-active' : ''}`}
            onClick={handleEnableDevMode}
            title="Toggle Developer Mode - shows all fields, sections, and enables editing"
          >
            {devModeActive ? 'Deactivate' : 'Dev Mode'}
          </button>
        );

      case 'devtools.blurFields':
        return (
          <button
            key={buttonId}
            className={`d365-toolbar-btn ${fieldsBlurred ? 'd365-toolbar-btn-active' : ''}`}
            onClick={handleToggleBlurFields}
            title="Blur field values for privacy when sharing screen or taking screenshots"
          >
            {fieldsBlurred ? 'Unblur' : 'Blur Fields'}
          </button>
        );

      case 'tools.copyId':
        return (
          <button key={buttonId} className="d365-toolbar-btn" onClick={handleCopyRecordId} title="Copy current record ID">
            Copy ID
          </button>
        );

      case 'tools.cacheRefresh':
        return (
          <button
            key={buttonId}
            className="d365-toolbar-btn d365-toolbar-btn-cache-refresh"
            onClick={handleCacheRefresh}
            title="Perform hard refresh (Ctrl+F5) to clear cache"
          >
            Cache Refresh
          </button>
        );

      case 'tools.webApi':
        return (
          <button key={buttonId} className="d365-toolbar-btn" onClick={handleOpenWebAPI} title="Open Web API data in new tab">
            Web API
          </button>
        );

      case 'tools.queryBuilder':
        return (
          <button key={buttonId} className="d365-toolbar-btn" onClick={handleOpenQueryBuilder} title="Open Advanced Find">
            Advanced Find
          </button>
        );

      case 'tools.traceLogs':
        return (
          <button key={buttonId} className="d365-toolbar-btn" onClick={handleShowTraceLogs} title="View Plugin Trace Logs">
            Plugin Trace Logs
          </button>
        );

      case 'tools.promptMaker':
        return (
          <button key={buttonId} className="d365-toolbar-btn" onClick={handleOpenPromptMaker} title="AI Prompt Maker - Generate context for AI assistants">
            AI Prompt Maker
          </button>
        );

      case 'tools.jsLibraries':
        return (
          <button
            key={buttonId}
            className="d365-toolbar-btn"
            onClick={handleShowLibraries}
            title="View JavaScript libraries and event handlers"
          >
            JS Libraries
          </button>
        );

      case 'tools.optionSets':
        return (
          <button key={buttonId} className="d365-toolbar-btn" onClick={handleShowOptionSets} title="View option set values used on this form">
            Option Sets
          </button>
        );

      case 'tools.odataFields':
        return (
          <button
            key={buttonId}
            className="d365-toolbar-btn"
            onClick={handleShowODataFields}
            title="View OData field metadata for this entity"
          >
            OData Fields
          </button>
        );

      case 'tools.auditHistory':
        return (
          <button key={buttonId} className="d365-toolbar-btn" onClick={handleShowAuditHistory} title="View audit history for this record">
            Audit History
          </button>
        );

      case 'tools.formEditor':
        return (
          <button key={buttonId} className="d365-toolbar-btn" onClick={handleOpenFormEditor} title="Open form editor">
            Form Editor
          </button>
        );

      default:
        return null;
    }
  };

  // Render section based on section layout config
  const renderSection = (sectionId: SectionId) => {
    const buttons = getButtonsForSection(sectionId);
    const renderedButtons = buttons
      .map((id) => renderButton(id))
      .filter((el): el is JSX.Element => Boolean(el));

    if (renderedButtons.length === 0) return null;

    const label = getSectionLabel(sectionId, DEFAULT_SECTION_LABELS[sectionId]);
    return (
      <div key={sectionId} className="d365-toolbar-section">
        {label && <span className="d365-toolbar-section-label">{label}</span>}
        {renderedButtons}
      </div>
    );
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
          onClear={handleClearTraceLogs}
        />
      )}

      {showPromptMaker && (
        <PromptMakerViewer
          onClose={handleClosePromptMaker}
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

      {showImpersonation && (
        <ImpersonationSelector
          data={impersonationData}
          onClose={handleCloseImpersonation}
          onSelect={handleSelectImpersonation}
          onRefresh={loadImpersonationUsers}
          currentImpersonation={impersonatedUser}
        />
      )}
    </div>
  );
};

export default D365Toolbar;
