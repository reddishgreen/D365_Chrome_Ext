import React, { useState, useEffect, useRef, useCallback } from 'react';
import { D365Helper } from '../utils/D365Helper';
import { restoreShellContainerLayout, setShellContainerOffset } from '../utils/shellLayout';
import FormLibrariesAnalyzer from './FormLibrariesAnalyzer';
import PluginTraceLogViewer, { PluginTraceLogData } from './PluginTraceLogViewer';
import OptionSetsViewer, { OptionSetsData } from './OptionSetsViewer';
import ODataFieldsViewer, { ODataFieldsData } from './ODataFieldsViewer';
import AuditHistoryViewer, { AuditHistoryData, AuditRecord } from './AuditHistoryViewer';
import QueryBuilder from '../../query-builder/components/QueryBuilder';
import ImpersonationSelector, { ImpersonationData, SystemUser } from './ImpersonationSelector';
import RecordNavigator from './RecordNavigator';
import { EntityInfo } from '../utils/D365Helper';
import PromptMakerViewer from './PromptMakerViewer';
import ActiveProcessesViewer, { ActiveProcessesData, ProcessRecord } from './ActiveProcessesViewer';
import PluginStepsViewer, { PluginStepsData, PluginStepRecord, PluginStepImage } from './PluginStepsViewer';
import PrivilegeDebugger, { PrivilegeDebugData } from './PrivilegeDebugger';
import CommandPalette, { PaletteCommand, CustomCommand, normaliseChord } from './CommandPalette';
import CloseIcon from './CloseIcon';

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
    'tools.promptMaker',
    'tools.navigateTo',
    'tools.commandPalette',
    'tools.activeProcesses',
    'tools.pluginSteps',
    'tools.privilegeDebug'
  ]
};

// Out-of-the-box keyboard chords. Applied on first install only — once the user
// edits any binding (or explicitly clears them) we don't re-add these.
// Shift+letter is intentionally chosen so chords don't conflict with Chrome shortcuts
// (which are mostly Ctrl-based) and don't fire while you're typing in form fields
// (the global listener bails when focus is in an editable element).
const DEFAULT_KEY_BINDINGS: Record<string, string> = {
  'cmd.devmode': 'shift+d',
  'cmd.auditHistory': 'shift+a',
  'cmd.webApi': 'shift+w',
  'cmd.odataFields': 'shift+o',
  'cmd.optionSets': 'shift+s',
  'cmd.impersonate': 'shift+i',
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
    'tools.promptMaker': true,
    'tools.navigateTo': true,
    'tools.commandPalette': true,
    'tools.activeProcesses': true,
    'tools.pluginSteps': true,
    'tools.privilegeDebug': true
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
      'tools.commandPalette',
      'tools.webApi',
      'tools.traceLogs',
      'tools.queryBuilder',
      'tools.optionSets',
      'tools.cacheRefresh',
      'tools.jsLibraries',
      'tools.odataFields',
      'tools.auditHistory',
      'tools.activeProcesses',
      'tools.pluginSteps',
      'tools.privilegeDebug',
      'tools.formEditor',
      'tools.navigateTo'
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
  const [auditHistoryLoading, setAuditHistoryLoading] = useState(false);
  const [auditHistoryNeedsFormReload, setAuditHistoryNeedsFormReload] = useState(false);
  const [showQueryBuilder, setShowQueryBuilder] = useState(false);
  const [notificationDuration, setNotificationDuration] = useState(3);
  const [toolbarPosition, setToolbarPosition] = useState<'top' | 'bottom'>('bottom');
  const [traceLogLimit, setTraceLogLimit] = useState(20);
  const [toolbarConfig, setToolbarConfig] = useState<ToolbarConfig>(DEFAULT_TOOLBAR_CONFIG);
  const [showImpersonation, setShowImpersonation] = useState(false);
  const [impersonatedUser, setImpersonatedUser] = useState<SystemUser | null>(null);
  const [impersonationData, setImpersonationData] = useState<ImpersonationData | null>(null);
  const [showRecordNavigator, setShowRecordNavigator] = useState(false);
  const [navigatorEntities, setNavigatorEntities] = useState<EntityInfo[] | null>(null);
  const [navigatorError, setNavigatorError] = useState<string | undefined>(undefined);
  const [showActiveProcesses, setShowActiveProcesses] = useState(false);
  const [activeProcessesData, setActiveProcessesData] = useState<ActiveProcessesData | null>(null);
  const [activeProcessesLoading, setActiveProcessesLoading] = useState(false);
  const [showPluginSteps, setShowPluginSteps] = useState(false);
  const [pluginStepsData, setPluginStepsData] = useState<PluginStepsData | null>(null);
  const [pluginStepsLoading, setPluginStepsLoading] = useState(false);
  const [showPrivilegeDebug, setShowPrivilegeDebug] = useState(false);
  const [privilegeDebugData, setPrivilegeDebugData] = useState<PrivilegeDebugData | null>(null);
  const [privilegeDebugLoading, setPrivilegeDebugLoading] = useState(false);
  const [showCommandPalette, setShowCommandPalette] = useState(false);
  const [paletteCompactMode, setPaletteCompactMode] = useState(false);
  // `showTool` (popup setting) hides the bar entirely but the component keeps
  // running so Ctrl+K and bound chords still work.
  const [showTool, setShowTool] = useState(true);
  const [pinnedCommandIds, setPinnedCommandIds] = useState<string[]>([]);
  const [customCommands, setCustomCommands] = useState<CustomCommand[]>([]);
  const [keyBindings, setKeyBindings] = useState<Record<string, string>>({});
  // Refs the global hotkey listener reads from so we don't re-attach on every render.
  const keyBindingsRef = useRef<Record<string, string>>({});
  const paletteCommandsRef = useRef<PaletteCommand[]>([]);
  const showCommandPaletteRef = useRef(false);
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

    chrome.storage.sync.get(['showTool', 'notificationDuration', 'toolbarPosition', 'traceLogLimit', 'toolbarConfig', 'devModeActive', 'devModePersist', 'paletteSettings'], (result) => {
      if (typeof result.showTool === 'boolean') {
        setShowTool(result.showTool);
      }
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
      // Apply default key bindings on first install (when paletteSettings.keyBindings
      // has never been written). After that we respect whatever the user has set —
      // including an explicitly empty object meaning "no defaults, please".
      const ps = (result.paletteSettings || {}) as any;
      if (ps && typeof ps === 'object') {
        if (typeof ps.compactMode === 'boolean') setPaletteCompactMode(ps.compactMode);
        if (Array.isArray(ps.pinnedCommandIds)) {
          setPinnedCommandIds(ps.pinnedCommandIds.filter((x: any) => typeof x === 'string'));
        }
        if (Array.isArray(ps.customCommands)) {
          setCustomCommands(
            ps.customCommands.filter(
              (c: any) =>
                c && typeof c.id === 'string' && typeof c.label === 'string' && typeof c.url === 'string'
            )
          );
        }
        const hasKeyBindingsField = Object.prototype.hasOwnProperty.call(ps, 'keyBindings');
        if (hasKeyBindingsField && ps.keyBindings && typeof ps.keyBindings === 'object') {
          const cleaned: Record<string, string> = {};
          Object.entries(ps.keyBindings).forEach(([k, v]) => {
            if (typeof k === 'string' && typeof v === 'string' && v) cleaned[k] = v;
          });
          setKeyBindings(cleaned);
          keyBindingsRef.current = cleaned;
        } else if (!hasKeyBindingsField) {
          // First-run install — seed defaults and persist so we don't reseed if user clears them.
          const seeded = { ...DEFAULT_KEY_BINDINGS };
          setKeyBindings(seeded);
          keyBindingsRef.current = seeded;
          chrome.storage.sync.set({
            paletteSettings: {
              compactMode: typeof ps.compactMode === 'boolean' ? ps.compactMode : false,
              pinnedCommandIds: Array.isArray(ps.pinnedCommandIds) ? ps.pinnedCommandIds : [],
              customCommands: Array.isArray(ps.customCommands) ? ps.customCommands : [],
              keyBindings: seeded,
            },
          });
        }
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
      if (changes.showTool && typeof changes.showTool.newValue === 'boolean') {
        setShowTool(changes.showTool.newValue);
      }
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
      if (changes.paletteSettings) {
        const ps = changes.paletteSettings.newValue || {};
        if (typeof ps.compactMode === 'boolean') setPaletteCompactMode(ps.compactMode);
        if (Array.isArray(ps.pinnedCommandIds)) {
          setPinnedCommandIds(ps.pinnedCommandIds.filter((x: any) => typeof x === 'string'));
        }
        if (Array.isArray(ps.customCommands)) {
          setCustomCommands(
            ps.customCommands.filter(
              (c: any) =>
                c && typeof c.id === 'string' && typeof c.label === 'string' && typeof c.url === 'string'
            )
          );
        }
        if (ps.keyBindings && typeof ps.keyBindings === 'object') {
          const cleaned: Record<string, string> = {};
          Object.entries(ps.keyBindings).forEach(([k, v]) => {
            if (typeof k === 'string' && typeof v === 'string' && v) cleaned[k] = v;
          });
          setKeyBindings(cleaned);
          keyBindingsRef.current = cleaned;
        } else if (ps && Object.prototype.hasOwnProperty.call(ps, 'keyBindings')) {
          setKeyBindings({});
          keyBindingsRef.current = {};
        }
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

  // Bar is hidden whenever the user has chosen compact mode OR has turned the toolbar
  // off via the popup. Either way the keyboard listeners remain mounted.
  const isBarHidden = paletteCompactMode || !showTool;

  const updateShellOffset = useCallback(() => {
    if (isBarHidden) {
      // Bar isn't visible; reclaim the space D365's shell was leaving for it.
      restoreShellContainerLayout();
      return;
    }
    const toolbarHeight = toolbarRef.current?.getBoundingClientRect().height;
    const fallbackHeight = 70;
    setShellContainerOffset(Math.round(toolbarHeight ?? fallbackHeight), toolbarPosition);
  }, [toolbarPosition, isBarHidden]);

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

  // Ctrl+K / Cmd+K opens the command palette; user-bound chords run their commands.
  // Reads from refs so we don't have to re-attach the listener on every render.
  useEffect(() => {
    const onKeyDown = (e: KeyboardEvent) => {
      const isMac = navigator.platform.toUpperCase().includes('MAC');
      const modifier = isMac ? e.metaKey : e.ctrlKey;
      if (modifier && (e.key === 'k' || e.key === 'K')) {
        e.preventDefault();
        e.stopPropagation();
        setShowCommandPalette((s) => !s);
        return;
      }

      // Don't fire user-bound chords while the palette itself is open
      // (palette has its own keyboard handling, including chord-capture mode).
      if (showCommandPaletteRef.current) return;

      // Skip when the user is typing in an editable element so we don't steal
      // chords from text inputs / D365 form fields.
      const target = e.target as HTMLElement | null;
      if (target) {
        const tag = target.tagName;
        const isEditable =
          tag === 'INPUT' ||
          tag === 'TEXTAREA' ||
          tag === 'SELECT' ||
          (target as any).isContentEditable === true;
        if (isEditable) return;
      }

      const chord = normaliseChord(e);
      if (!chord) return;

      const bindings = keyBindingsRef.current;
      const matchEntry = Object.entries(bindings).find(([, c]) => c === chord);
      if (!matchEntry) return;
      const [cmdId] = matchEntry;
      const cmd = paletteCommandsRef.current.find((c) => c.id === cmdId);
      if (!cmd) return;

      e.preventDefault();
      e.stopPropagation();
      try {
        cmd.run();
      } catch (err) {
        console.error('[D365 Helper] Bound command failed:', err);
      }
    };
    window.addEventListener('keydown', onKeyDown, true);
    return () => window.removeEventListener('keydown', onKeyDown, true);
  }, []);

  // Keep refs in sync so the keydown listener (which has [] deps) sees fresh values.
  useEffect(() => {
    showCommandPaletteRef.current = showCommandPalette;
  }, [showCommandPalette]);

  useEffect(() => {
    keyBindingsRef.current = keyBindings;
  }, [keyBindings]);

  // Listen for refresh requests from pop-out trace viewer window
  useEffect(() => {
    const listener = (message: any, _sender: any, sendResponse: (response?: any) => void) => {
      if (message.type === 'TRACE_LOG_REFRESH_REQUEST') {
        helper.getPluginTraceLogs(traceLogLimit).then((data) => {
          // Push updated data to the pop-out window via storage + broadcast
          chrome.storage.local.set({ d365_trace_log_popout_data: data }, () => {
            chrome.runtime.sendMessage({ type: 'TRACE_LOG_DATA_UPDATE', data });
          });
          sendResponse({ success: true });
        }).catch(() => {
          sendResponse({ success: false });
        });
        return true; // async response
      }
    };
    chrome.runtime.onMessage.addListener(listener);
    return () => chrome.runtime.onMessage.removeListener(listener);
  }, [traceLogLimit]);

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

          if (result.devModeActive === true) {
            // Apply if storage says it should be active (user had it on when navigating between forms)
            if (!hasApplied && applyDevModeRef.current) {
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
                } catch (error) {
                  console.warn('[D365 Helper] Failed to reapply schema overlay', error);
                }
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
      // Check storage to see if Dev Mode should be active
      chrome.storage.sync.get(['devModeActive'], async (result) => {
        const shouldBeActive = result.devModeActive === true;

        if (!shouldBeActive) {
          // Dev Mode is off in storage, make sure UI matches
          if (devModeActiveRef.current) {
            devModeActiveRef.current = false;
            setDevModeActive(false);
          }
          return;
        }
        
        // Dev Mode should be active, reapply it
        if (devModeReapplyInProgress) {
          return;
        }
        devModeReapplyInProgress = true;

        // Wait for form to be ready before applying
        let formReady = false;
        let readyAttempts = 0;
        const maxReadyAttempts = 30;

        while (!formReady && readyAttempts < maxReadyAttempts) {
          try {
            await helper.getRecordId();
            formReady = true;
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
            await applyDevModeRef.current(false);
            // Wait a bit and reapply schema overlay in case D365 reset it
            await new Promise(resolve => setTimeout(resolve, 500));
            try {
              await helper.toggleSchemaOverlay(true);
              setShowSchemaNames(true);
              showSchemaNamesRef.current = true;
            } catch (error) {
              console.warn('[D365 Helper] Failed to reapply schema overlay', error);
            }
            devModeReapplyAttempts = 0;
            devModeReapplyInProgress = false;
          } catch (error) {
            console.error('[D365 Helper] Failed to reapply Dev Mode:', error);
            devModeReapplyAttempts += 1;
            if (devModeReapplyAttempts <= MAX_DEV_MODE_REAPPLY_ATTEMPTS) {
              // Form may still be loading; retry with a small backoff.
              const backoff = 600 + devModeReapplyAttempts * 500;
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

        // Check if we're still on a form page
        const isFormPage = isCurrentPageAForm();
        
        if (!isFormPage) {
          // Navigated to a non-form page (grid, dashboard, etc.) - deactivate Dev Mode
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
              scheduleDevModeReapply(1500);
            } else {
              // Persistence is off or Dev Mode is inactive - deactivate
              if (shouldBeActive && !persistEnabled && devModeActiveRef.current && removeDevModeRef.current) {
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

  // Toggle handlers read state via refs so they always see the freshest value,
  // even when invoked through a stale closure (e.g. a pinned button or palette command
  // whose handler was captured during an earlier render).
  const handleToggleFields = async () => {
    const newState = !allFieldsVisibleRef.current;
    setAllFieldsVisible(newState);
    allFieldsVisibleRef.current = newState;
    try {
      await helper.toggleAllFields(newState);
      showNotification(newState ? 'All fields shown' : 'Fields restored to original state');
    } catch (error: any) {
      setAllFieldsVisible(!newState); // Revert state on error
      allFieldsVisibleRef.current = !newState;
      const message = error?.message || 'Error toggling fields';
      showNotification(message);
    }
  };

  const handleToggleSections = async () => {
    const newState = !allSectionsVisibleRef.current;
    setAllSectionsVisible(newState);
    allSectionsVisibleRef.current = newState;
    try {
      await helper.toggleAllSections(newState);
      showNotification(newState ? 'All sections shown' : 'Sections restored to original state');
    } catch (error: any) {
      setAllSectionsVisible(!newState); // Revert state on error
      allSectionsVisibleRef.current = !newState;
      const message = error?.message || 'Error toggling sections';
      showNotification(message);
    }
  };

  const handleToggleBlurFields = async () => {
    const newState = !fieldsBlurredRef.current;
    setFieldsBlurred(newState);
    fieldsBlurredRef.current = newState;
    try {
      await helper.toggleBlurFields(newState);
      showNotification(newState ? 'Fields blurred for privacy' : 'Field blur removed');
    } catch (error: any) {
      setFieldsBlurred(!newState); // Revert state on error
      fieldsBlurredRef.current = !newState;
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
    const newState = !showSchemaNamesRef.current;
    setShowSchemaNames(newState);
    showSchemaNamesRef.current = newState;
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

  const handlePopoutTraceLogs = () => {
    if (!traceLogData) return;
    // Store data and source tab for the pop-out window
    chrome.storage.local.set({
      d365_trace_log_popout_data: traceLogData,
      d365_trace_log_source_tab: null, // Will be set by background if needed
    }, () => {
      const url = chrome.runtime.getURL('trace-viewer.html');
      window.open(url, 'trace-viewer-window', 'width=1400,height=800');
      // Close the inline modal
      setShowTraceLogs(false);
      setTraceLogData(null);
    });
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

  const loadAuditHistory = async () => {
    setAuditHistoryLoading(true);
    try {
      const data = await helper.getAuditHistory();
      setAuditHistoryData(data);
    } catch (error: any) {
      const message = error?.message || 'Error loading audit history';
      showNotification(message);
      setAuditHistoryData({ records: [], error: message });
    } finally {
      setAuditHistoryLoading(false);
    }
  };

  const handleShowAuditHistory = async () => {
    setAuditHistoryNeedsFormReload(false);
    setShowAuditHistory(true);
    if (!auditHistoryData) {
      setAuditHistoryData({ records: [] });
    }
    await loadAuditHistory();
  };

  const handleRefreshAuditHistory = async () => {
    await loadAuditHistory();
  };

  const handleCloseAuditHistory = () => {
    const shouldReloadForm = auditHistoryNeedsFormReload;
    setShowAuditHistory(false);
    setAuditHistoryLoading(false);
    setAuditHistoryData(null);
    setAuditHistoryNeedsFormReload(false);

    if (shouldReloadForm) {
      showNotification('Refreshing form to reflect rollback changes...');
      setTimeout(() => location.reload(), 150);
    }
  };

  const handleAuditRollback = async (record: AuditRecord, skipPlugins?: boolean): Promise<boolean> => {
    try {
      await helper.rollbackFields(
        [{ fieldName: record.fieldName, oldValue: record.rollbackOldValue ?? record.oldValue }],
        skipPlugins
      );
      setAuditHistoryNeedsFormReload(true);
      showNotification(`Rolled back ${record.fieldName}. Refreshing audit history...`);
      loadAuditHistory().catch(() => undefined);
      return true;
    } catch (error: any) {
      const message = error?.message || 'Rollback failed';
      showNotification(message);
      return false;
    }
  };

  const handleAuditRollbackGroup = async (records: AuditRecord[], skipPlugins?: boolean): Promise<boolean> => {
    try {
      await helper.rollbackFields(
        records.map((record) => ({
          fieldName: record.fieldName,
          oldValue: record.rollbackOldValue ?? record.oldValue,
        })),
        skipPlugins
      );
      setAuditHistoryNeedsFormReload(true);
      showNotification(`Rolled back ${records.length} field${records.length !== 1 ? 's' : ''}. Refreshing audit history...`);
      loadAuditHistory().catch(() => undefined);
      return true;
    } catch (error: any) {
      const message = error?.message || 'Rollback failed';
      showNotification(message);
      return false;
    }
  };

  const handleShowRecordNavigator = async () => {
    setShowRecordNavigator(true);
    setNavigatorError(undefined);
    // Load entities if not already cached
    if (!navigatorEntities) {
      try {
        const entities = await helper.getAllEntities();
        setNavigatorEntities(entities);
      } catch (error: any) {
        setNavigatorError(error?.message || 'Failed to load entities');
      }
    }
  };

  const handleCloseRecordNavigator = () => {
    setShowRecordNavigator(false);
  };

  // ===== Active Processes =====
  const loadActiveProcesses = async () => {
    setActiveProcessesLoading(true);
    try {
      const data = await helper.getActiveProcesses();
      setActiveProcessesData(data);
    } catch (error: any) {
      const msg = error?.message || 'Failed to load processes';
      setActiveProcessesData({ entityName: '', processes: [], error: msg });
    } finally {
      setActiveProcessesLoading(false);
    }
  };

  const handleShowActiveProcesses = async () => {
    setShowActiveProcesses(true);
    if (!activeProcessesData) {
      setActiveProcessesData({ entityName: '', processes: [] });
    }
    await loadActiveProcesses();
  };

  const handleCloseActiveProcesses = () => {
    setShowActiveProcesses(false);
    setActiveProcessesData(null);
  };

  const handleRefreshActiveProcesses = async () => {
    await loadActiveProcesses();
  };

  const handleToggleProcess = async (process: ProcessRecord): Promise<boolean> => {
    try {
      const res = await helper.toggleProcess(process.id, !process.isActivated);
      if (!res?.success) throw new Error(res?.error || 'Toggle failed');
      showNotification(`${process.isActivated ? 'Deactivated' : 'Activated'} ${process.name}`);
      // Optimistic local update so the UI reacts even before refresh.
      setActiveProcessesData((prev) => {
        if (!prev) return prev;
        const next = {
          ...prev,
          processes: prev.processes.map((p) =>
            p.id === process.id
              ? { ...p, isActivated: !process.isActivated, statecode: !process.isActivated ? 1 : 0, statuscode: !process.isActivated ? 2 : 1 }
              : p
          ),
        };
        return next;
      });
      return true;
    } catch (error: any) {
      showNotification(error?.message || 'Toggle failed');
      return false;
    }
  };

  // ===== Plugin Steps =====
  const loadPluginSteps = async () => {
    setPluginStepsLoading(true);
    try {
      const data = await helper.getPluginSteps();
      setPluginStepsData(data);
    } catch (error: any) {
      const msg = error?.message || 'Failed to load plugin steps';
      setPluginStepsData({ entityName: '', steps: [], error: msg });
    } finally {
      setPluginStepsLoading(false);
    }
  };

  const handleShowPluginSteps = async () => {
    setShowPluginSteps(true);
    if (!pluginStepsData) {
      setPluginStepsData({ entityName: '', steps: [] });
    }
    await loadPluginSteps();
  };

  const handleClosePluginSteps = () => {
    setShowPluginSteps(false);
    setPluginStepsData(null);
  };

  const handleRefreshPluginSteps = async () => {
    await loadPluginSteps();
  };

  const handleUpdatePluginStepImage = async (
    stepId: string,
    image: PluginStepImage,
    patch: Partial<Pick<PluginStepImage, 'name' | 'entityAlias' | 'attributes' | 'imageType'>>
  ): Promise<boolean> => {
    try {
      const res = await helper.updatePluginStepImage({ id: image.id, ...patch });
      if (!res?.success) throw new Error(res?.error || 'Update failed');
      showNotification(`Updated image "${patch.name ?? image.name}"`);
      // Optimistic local update
      setPluginStepsData((prev) => {
        if (!prev) return prev;
        return {
          ...prev,
          steps: prev.steps.map((s) =>
            s.id !== stepId
              ? s
              : {
                  ...s,
                  images: (s.images || []).map((i) =>
                    i.id !== image.id
                      ? i
                      : {
                          ...i,
                          ...patch,
                          imageTypeLabel:
                            patch.imageType === undefined
                              ? i.imageTypeLabel
                              : patch.imageType === 0
                              ? 'Pre'
                              : patch.imageType === 1
                              ? 'Post'
                              : patch.imageType === 2
                              ? 'Both'
                              : i.imageTypeLabel,
                        }
                  ),
                }
          ),
        };
      });
      return true;
    } catch (error: any) {
      showNotification(error?.message || 'Update image failed');
      return false;
    }
  };

  const handleDeletePluginStepImage = async (stepId: string, imageId: string): Promise<boolean> => {
    try {
      const res = await helper.deletePluginStepImage(imageId);
      if (!res?.success) throw new Error(res?.error || 'Delete failed');
      showNotification('Image deleted');
      setPluginStepsData((prev) => {
        if (!prev) return prev;
        return {
          ...prev,
          steps: prev.steps.map((s) =>
            s.id !== stepId ? s : { ...s, images: (s.images || []).filter((i) => i.id !== imageId) }
          ),
        };
      });
      return true;
    } catch (error: any) {
      showNotification(error?.message || 'Delete image failed');
      return false;
    }
  };

  const handleCreatePluginStepImage = async (
    stepId: string,
    args: { name: string; entityAlias: string; attributes: string; imageType: number }
  ): Promise<boolean> => {
    try {
      const res = await helper.createPluginStepImage({ stepId, ...args });
      if (!res?.success) throw new Error(res?.error || 'Create failed');
      showNotification(`Added image "${args.name}"`);
      // Refresh to pick up the new image
      await loadPluginSteps();
      return true;
    } catch (error: any) {
      showNotification(error?.message || 'Create image failed');
      return false;
    }
  };

  const handleTogglePluginStep = async (step: PluginStepRecord): Promise<boolean> => {
    try {
      const res = await helper.togglePluginStep(step.id, !step.isEnabled);
      if (!res?.success) throw new Error(res?.error || 'Toggle failed');
      showNotification(`${step.isEnabled ? 'Disabled' : 'Enabled'} ${step.name}`);
      setPluginStepsData((prev) => {
        if (!prev) return prev;
        return {
          ...prev,
          steps: prev.steps.map((s) =>
            s.id === step.id ? { ...s, isEnabled: !step.isEnabled } : s
          ),
        };
      });
      return true;
    } catch (error: any) {
      showNotification(error?.message || 'Toggle failed');
      return false;
    }
  };

  // ===== Privilege Debugger =====
  const runPrivilegeDebug = async (entityName: string, recordId: string) => {
    setPrivilegeDebugLoading(true);
    try {
      const data = await helper.getPrivilegeDebug(entityName, recordId);
      setPrivilegeDebugData(data);
    } catch (error: any) {
      setPrivilegeDebugData({
        inputEntityName: entityName,
        inputRecordId: recordId,
        error: error?.message || 'Privilege check failed.',
      });
    } finally {
      setPrivilegeDebugLoading(false);
    }
  };

  const handleShowPrivilegeDebug = async () => {
    setShowPrivilegeDebug(true);
    // Auto-run if we're already on a record
    try {
      const entity = await helper.getEntityName();
      const id = await helper.getRecordId();
      if (entity && id) {
        await runPrivilegeDebug(entity, id);
      }
    } catch {
      // Ignore — user can enter manually.
    }
  };

  const handleClosePrivilegeDebug = () => {
    setShowPrivilegeDebug(false);
    setPrivilegeDebugData(null);
  };

  // ===== Command Palette =====
  const handleOpenCommandPalette = () => setShowCommandPalette(true);
  const handleCloseCommandPalette = () => setShowCommandPalette(false);

  const persistPaletteSettings = (
    next: {
      compactMode?: boolean;
      pinnedCommandIds?: string[];
      customCommands?: CustomCommand[];
      keyBindings?: Record<string, string>;
    }
  ) => {
    const merged = {
      compactMode: next.compactMode ?? paletteCompactMode,
      pinnedCommandIds: next.pinnedCommandIds ?? pinnedCommandIds,
      customCommands: next.customCommands ?? customCommands,
      keyBindings: next.keyBindings ?? keyBindings,
    };
    chrome.storage.sync.set({ paletteSettings: merged });
  };

  const handleSetBinding = (commandId: string, chord: string | null) => {
    setKeyBindings((prev) => {
      const next = { ...prev };
      if (chord) {
        next[commandId] = chord;
      } else {
        delete next[commandId];
      }
      keyBindingsRef.current = next;
      persistPaletteSettings({ keyBindings: next });
      return next;
    });
  };

  const handleTogglePin = (commandId: string) => {
    setPinnedCommandIds((prev) => {
      const next = prev.includes(commandId)
        ? prev.filter((id) => id !== commandId)
        : [...prev, commandId];
      persistPaletteSettings({ pinnedCommandIds: next });
      return next;
    });
  };

  const handleAddCustomCommand = (cmd: { label: string; url: string; description?: string }) => {
    const newCmd: CustomCommand = {
      id: `custom-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      label: cmd.label,
      url: cmd.url,
      description: cmd.description,
    };
    setCustomCommands((prev) => {
      const next = [...prev, newCmd];
      persistPaletteSettings({ customCommands: next });
      return next;
    });
    showNotification(`Added "${cmd.label}"`);
  };

  const handleDeleteCustomCommand = (id: string) => {
    setCustomCommands((prev) => {
      const next = prev.filter((c) => c.id !== id);
      persistPaletteSettings({ customCommands: next });
      return next;
    });
    setPinnedCommandIds((prev) => {
      if (!prev.includes(id)) return prev;
      const next = prev.filter((p) => p !== id);
      persistPaletteSettings({ pinnedCommandIds: next });
      return next;
    });
  };

  const handleToggleCompactMode = () => {
    setPaletteCompactMode((prev) => {
      const next = !prev;
      persistPaletteSettings({ compactMode: next });
      showNotification(next ? 'Compact toolbar enabled' : 'Full toolbar restored');
      return next;
    });
  };

  const handleToggleShowTool = () => {
    setShowTool((prev) => {
      const next = !prev;
      chrome.storage.sync.set({ showTool: next });
      showNotification(next ? 'Toolbar bar shown' : 'Toolbar bar hidden (Ctrl+K still works)');
      return next;
    });
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
    try {
      await helper.toggleAllFields(true);
    } catch (error) {
      console.error('[D365 Helper] applyDevMode: Failed to toggle fields', error);
    }
    
    try {
      await helper.toggleAllSections(true);
    } catch (error) {
      console.error('[D365 Helper] applyDevMode: Failed to toggle sections', error);
    }
    
    setAllFieldsVisible(true);
    setAllSectionsVisible(true);
    allFieldsVisibleRef.current = true;
    allSectionsVisibleRef.current = true;
    
    try {
      await helper.unlockFields();
    } catch (error) {
      console.error('[D365 Helper] applyDevMode: Failed to unlock fields', error);
    }
    
    try {
      await helper.disableFieldRequirements();
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
    // Use ref so repeated triggers (e.g. clicking a pinned palette command twice)
    // always read the freshest dev-mode state instead of a stale closure value.
    const newState = !devModeActiveRef.current;

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
              aria-label="Stop impersonating"
            >
              <CloseIcon size={12} />
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

      case 'tools.navigateTo':
        return (
          <button key={buttonId} className="d365-toolbar-btn" onClick={handleShowRecordNavigator} title="Navigate to a record by entity and ID">
            Navigate To
          </button>
        );

      case 'tools.commandPalette':
        return (
          <button key={buttonId} className="d365-toolbar-btn" onClick={handleOpenCommandPalette} title="Command Palette (Ctrl+K)">
            ⌘ Command (Ctrl+K)
          </button>
        );

      case 'tools.activeProcesses':
        return (
          <button key={buttonId} className="d365-toolbar-btn" onClick={handleShowActiveProcesses} title="View workflows, business rules, and flows on this entity">
            Active Processes
          </button>
        );

      case 'tools.pluginSteps':
        return (
          <button key={buttonId} className="d365-toolbar-btn" onClick={handleShowPluginSteps} title="View registered plugin steps for this entity">
            Plugin Steps
          </button>
        );

      case 'tools.privilegeDebug':
        return (
          <button key={buttonId} className="d365-toolbar-btn" onClick={handleShowPrivilegeDebug} title="Diagnose record access and security privileges">
            Privilege Debug
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

  // Commands available in the Ctrl+K palette
  const builtInPaletteCommands: PaletteCommand[] = [
    { id: 'cmd.toggleCompact', label: paletteCompactMode ? 'Switch to full toolbar' : 'Switch to compact toolbar', section: 'View', keywords: 'minimal hide bar reduce noise compact', run: handleToggleCompactMode },
    { id: 'cmd.toggleShowTool', label: showTool ? 'Hide toolbar bar' : 'Show toolbar bar', section: 'View', keywords: 'hide show bar visible toolbar', run: handleToggleShowTool },
    { id: 'cmd.fields.toggle', label: allFieldsVisible ? 'Hide all fields' : 'Show all fields', section: 'Form', keywords: 'show hide visible', run: handleToggleFields },
    { id: 'cmd.sections.toggle', label: allSectionsVisible ? 'Hide all sections' : 'Show all sections', section: 'Form', keywords: 'show hide visible', run: handleToggleSections },
    { id: 'cmd.schema.toggle', label: showSchemaNames ? 'Hide schema names' : 'Show schema names', section: 'Schema', keywords: 'logical name field', run: handleToggleSchemaNames },
    { id: 'cmd.schema.copy', label: 'Copy all schema names', section: 'Schema', keywords: 'clipboard fields', run: handleCopyAllSchemaNames },
    { id: 'cmd.devmode', label: devModeActive ? 'Deactivate Dev Mode' : 'Activate Dev Mode', section: 'Dev Tools', keywords: 'developer unlock', run: handleEnableDevMode },
    { id: 'cmd.unlock', label: 'Enable editing on locked fields', section: 'Dev Tools', keywords: 'unlock editable', run: handleUnlockFields },
    { id: 'cmd.testdata', label: 'Fill form with test data', section: 'Dev Tools', keywords: 'autofill mock', run: handleAutoFill },
    { id: 'cmd.blur', label: fieldsBlurred ? 'Unblur field values' : 'Blur field values', section: 'Dev Tools', keywords: 'privacy redact', run: handleToggleBlurFields },
    { id: 'cmd.impersonate', label: 'Impersonate another user', section: 'Dev Tools', keywords: 'security user act as', run: handleShowImpersonation },
    { id: 'cmd.copyId', label: 'Copy current record ID', section: 'Tools', keywords: 'guid clipboard', run: handleCopyRecordId },
    { id: 'cmd.cacheRefresh', label: 'Hard refresh (cache reset)', section: 'Tools', keywords: 'reload ctrl f5', run: handleCacheRefresh },
    { id: 'cmd.webApi', label: 'Open Web API viewer', section: 'Tools', keywords: 'odata json record', run: handleOpenWebAPI },
    { id: 'cmd.queryBuilder', label: 'Open Advanced Find / Query Builder', section: 'Tools', keywords: 'fetchxml odata', run: handleOpenQueryBuilder },
    { id: 'cmd.traceLogs', label: 'View Plugin Trace Logs', section: 'Tools', keywords: 'plugin debug', run: handleShowTraceLogs },
    { id: 'cmd.jsLibraries', label: 'View JS libraries & event handlers', section: 'Tools', keywords: 'script form', run: handleShowLibraries },
    { id: 'cmd.optionSets', label: 'View option sets on this form', section: 'Tools', keywords: 'picklist values', run: handleShowOptionSets },
    { id: 'cmd.odataFields', label: 'View OData field metadata', section: 'Tools', keywords: 'attribute metadata', run: handleShowODataFields },
    { id: 'cmd.auditHistory', label: 'View audit history', section: 'Tools', keywords: 'changes rollback', run: handleShowAuditHistory },
    { id: 'cmd.activeProcesses', label: 'View active processes (workflows, business rules, flows)', section: 'Tools', keywords: 'workflow business rule flow process', run: handleShowActiveProcesses },
    { id: 'cmd.pluginSteps', label: 'View registered plugin steps', section: 'Tools', keywords: 'sdk message processing', run: handleShowPluginSteps },
    { id: 'cmd.privilegeDebug', label: 'Diagnose record access (Privilege Debugger)', section: 'Tools', keywords: 'security role permission why cant i see', run: handleShowPrivilegeDebug },
    { id: 'cmd.formEditor', label: 'Open form editor', section: 'Tools', keywords: 'classic edit', run: handleOpenFormEditor },
    { id: 'cmd.navigateTo', label: 'Navigate to a record by entity & ID', section: 'Tools', keywords: 'open jump record', run: handleShowRecordNavigator },
    { id: 'cmd.promptMaker', label: 'AI Prompt Maker', section: 'Tools', keywords: 'context ai chatgpt claude', run: handleOpenPromptMaker },
    { id: 'cmd.solutions', label: 'Open Solutions page', section: 'Navigation', keywords: 'maker portal', run: handleOpenSolutions },
    { id: 'cmd.adminCenter', label: 'Open Power Platform Admin Center', section: 'Navigation', keywords: 'admin environment', run: handleOpenAdminCenter },
  ];

  // Append user-defined custom commands (open URL in new tab).
  const customPaletteCommands: PaletteCommand[] = customCommands.map((c) => ({
    id: c.id,
    label: c.label,
    description: c.description || c.url,
    section: 'Custom',
    keywords: c.url,
    isCustom: true,
    run: () => window.open(c.url, '_blank', 'noopener,noreferrer'),
  }));

  const paletteCommands: PaletteCommand[] = [...builtInPaletteCommands, ...customPaletteCommands];
  // Keep the ref in sync each render so the global hotkey listener can dispatch the latest handlers.
  paletteCommandsRef.current = paletteCommands;

  return (
    <div
      ref={toolbarRef}
      className={`d365-toolbar d365-toolbar-${toolbarPosition} ${
        isBarHidden ? 'd365-toolbar-hidden' : ''
      }`}
    >
      {!isBarHidden && (
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

          {toolbarConfig.sectionOrder.map((sectionId) => renderSection(sectionId))}

          <div className="d365-toolbar-section d365-toolbar-actions">
            <span className="d365-toolbar-version-badge">v{chrome.runtime.getManifest().version}</span>
          </div>
        </div>
      )}

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
          onPopout={handlePopoutTraceLogs}
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

      {showAuditHistory && (
        <AuditHistoryViewer
          data={auditHistoryData || { records: [] }}
          onClose={handleCloseAuditHistory}
          onRefresh={handleRefreshAuditHistory}
          onRollback={handleAuditRollback}
          onRollbackGroup={handleAuditRollbackGroup}
          isLoading={auditHistoryLoading}
        />
      )}
      
      {showRecordNavigator && (
        <RecordNavigator
          entities={navigatorEntities}
          orgUrl={helper.getOrgUrl()}
          onClose={handleCloseRecordNavigator}
          error={navigatorError}
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

      {showActiveProcesses && (
        <ActiveProcessesViewer
          data={activeProcessesData}
          onClose={handleCloseActiveProcesses}
          onRefresh={handleRefreshActiveProcesses}
          onToggle={handleToggleProcess}
          isLoading={activeProcessesLoading}
        />
      )}

      {showPluginSteps && (
        <PluginStepsViewer
          data={pluginStepsData}
          onClose={handleClosePluginSteps}
          onRefresh={handleRefreshPluginSteps}
          onToggle={handleTogglePluginStep}
          onUpdateImage={handleUpdatePluginStepImage}
          onDeleteImage={handleDeletePluginStepImage}
          onCreateImage={handleCreatePluginStepImage}
          isLoading={pluginStepsLoading}
        />
      )}

      {showPrivilegeDebug && (
        <PrivilegeDebugger
          data={privilegeDebugData}
          defaultEntityName={privilegeDebugData?.inputEntityName}
          defaultRecordId={privilegeDebugData?.inputRecordId}
          onClose={handleClosePrivilegeDebug}
          onRun={runPrivilegeDebug}
          isLoading={privilegeDebugLoading}
        />
      )}

      <CommandPalette
        open={showCommandPalette}
        commands={paletteCommands}
        pinnedIds={pinnedCommandIds}
        customCommands={customCommands}
        keyBindings={keyBindings}
        onClose={handleCloseCommandPalette}
        onTogglePin={handleTogglePin}
        onSetBinding={handleSetBinding}
        onAddCustomCommand={handleAddCustomCommand}
        onDeleteCustomCommand={handleDeleteCustomCommand}
      />
    </div>
  );
};

export default D365Toolbar;
