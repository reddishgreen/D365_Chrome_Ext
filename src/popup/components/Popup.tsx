import React, { useState, useEffect, useRef } from 'react';
import './Popup.css';

type TabType = 'features' | 'settings';

// Toolbar customization types
// NOTE: allow custom user-defined sections (string ids)
type SectionId = string;

interface ButtonConfig {
  id: string;
  label: string;
  description: string;
}

interface SectionConfig {
  id: SectionId;
  label: string;
  buttons: ButtonConfig[];
}

interface ToolbarConfig {
  sectionOrder: SectionId[];
  buttonVisibility: Record<string, boolean>;
  sectionLabels?: Partial<Record<SectionId, string>>;
  sectionButtons?: Partial<Record<SectionId, string[]>>;
}

// Default toolbar configuration
const DEFAULT_SECTIONS: SectionConfig[] = [
  {
    id: 'fields',
    label: 'Fields',
    buttons: [
      { id: 'fields.showAll', label: 'Show All / Hide All', description: 'Toggle visibility of all fields' }
    ]
  },
  {
    id: 'sections',
    label: 'Sections',
    buttons: [
      { id: 'sections.showAll', label: 'Show All / Hide All', description: 'Toggle visibility of all sections' }
    ]
  },
  {
    id: 'schema',
    label: 'Schema',
    buttons: [
      { id: 'schema.showNames', label: 'Show Names / Hide Names', description: 'Toggle schema name overlays' },
      { id: 'schema.copyAll', label: 'Copy All', description: 'Copy all schema names' }
    ]
  },
  {
    id: 'navigation',
    label: 'Navigation',
    buttons: [
      { id: 'navigation.solutions', label: 'Solutions', description: 'Open solutions page' },
      { id: 'navigation.adminCenter', label: 'Admin Center', description: 'Open Power Platform admin center' }
    ]
  },
  {
    id: 'devtools',
    label: 'Dev Tools',
    buttons: [
      { id: 'devtools.enableEditing', label: 'Enable Editing', description: 'Enable editing for development' },
      { id: 'devtools.testData', label: 'Test Data', description: 'Fill fields with test data' },
      { id: 'devtools.devMode', label: 'Dev Mode', description: 'Toggle Developer Mode' },
      { id: 'devtools.blurFields', label: 'Blur Fields', description: 'Blur fields for privacy' },
      { id: 'devtools.impersonate', label: 'Impersonate', description: 'Impersonate another user for API calls' }
    ]
  },
  {
    id: 'tools',
    label: 'Tools',
    buttons: [
      { id: 'tools.copyId', label: 'Copy ID', description: 'Copy current record ID' },
      { id: 'tools.cacheRefresh', label: 'Cache Refresh', description: 'Perform hard refresh' },
      { id: 'tools.webApi', label: 'Web API', description: 'Open Web API data' },
      { id: 'tools.queryBuilder', label: 'Advanced Find', description: 'Open Advanced Find / Query Builder' },
      { id: 'tools.traceLogs', label: 'Plugin Trace Logs', description: 'View Plugin Trace Logs' },
      { id: 'tools.jsLibraries', label: 'JS Libraries', description: 'View JavaScript libraries' },
      { id: 'tools.optionSets', label: 'Option Sets', description: 'View option set values' },
      { id: 'tools.odataFields', label: 'OData Fields', description: 'View OData field metadata' },
      { id: 'tools.auditHistory', label: 'Audit History', description: 'View audit history' },
      { id: 'tools.formEditor', label: 'Form Editor', description: 'Open form editor' }
    ]
  }
];

const DEFAULT_TOOLBAR_CONFIG: ToolbarConfig = {
  sectionOrder: ['fields', 'sections', 'schema', 'navigation', 'devtools', 'tools'],
  buttonVisibility: DEFAULT_SECTIONS.reduce((acc, section) => {
    section.buttons.forEach(button => {
      acc[button.id] = true;
    });
    return acc;
  }, {} as Record<string, boolean>),
  sectionLabels: {},
  sectionButtons: DEFAULT_SECTIONS.reduce((acc, section) => {
    acc[section.id] = section.buttons.map((b) => b.id);
    return acc;
  }, {} as Partial<Record<SectionId, string[]>>)
};

const Popup: React.FC = () => {
  const [activeTab, setActiveTab] = useState<TabType>('features');
  const [showTool, setShowTool] = useState(true);
  const [notificationDuration, setNotificationDuration] = useState(3);
  const [toolbarPosition, setToolbarPosition] = useState<'top' | 'bottom'>('top');
  const [schemaOverlayColor, setSchemaOverlayColor] = useState('#0078d4');
  const [traceLogLimit, setTraceLogLimit] = useState(20);
  const [skipPluginsByDefault, setSkipPluginsByDefault] = useState(false);
  const [toolbarConfig, setToolbarConfig] = useState<ToolbarConfig>(DEFAULT_TOOLBAR_CONFIG);
  const [draggedSectionIndex, setDraggedSectionIndex] = useState<number | null>(null);
  const [draggedButton, setDraggedButton] = useState<{ buttonId: string; fromSectionId: SectionId } | null>(null);
  const [settingsTransferMessage, setSettingsTransferMessage] = useState<string>('');
  const importFileInputRef = useRef<HTMLInputElement | null>(null);

  const DEFAULT_SECTION_IDS: SectionId[] = ['fields', 'sections', 'schema', 'navigation', 'devtools', 'tools'];
  const ALL_BUTTONS = DEFAULT_SECTIONS.flatMap((s) => s.buttons);
  const BUTTON_META = ALL_BUTTONS.reduce((acc, b) => {
    acc[b.id] = b;
    return acc;
  }, {} as Record<string, ButtonConfig>);
  const DEFAULT_SECTION_BUTTONS: Record<SectionId, string[]> = DEFAULT_SECTIONS.reduce((acc, section) => {
    acc[section.id] = section.buttons.map((b) => b.id);
    return acc;
  }, {} as Record<SectionId, string[]>);

  const normalizeToolbarConfig = (config?: Partial<ToolbarConfig> | null): ToolbarConfig => {
    const baseOrder = DEFAULT_TOOLBAR_CONFIG.sectionOrder;
    const rawOrder = Array.isArray(config?.sectionOrder) ? (config!.sectionOrder as SectionId[]) : baseOrder;

    // Allow custom sections: keep any string ids, de-dupe, and ensure all default sections exist
    const validOrder = rawOrder.filter((id) => typeof id === 'string' && id.trim().length > 0);
    const dedupedOrder: SectionId[] = [];
    const seenOrder = new Set<string>();
    validOrder.forEach((id) => {
      if (seenOrder.has(id)) return;
      seenOrder.add(id);
      dedupedOrder.push(id);
    });

    const orderSet = new Set(dedupedOrder);
    const sectionOrder: SectionId[] = [...dedupedOrder, ...baseOrder.filter((id) => !orderSet.has(id))];

    // Merge button visibility (new buttons default to true)
    const buttonVisibility = {
      ...DEFAULT_TOOLBAR_CONFIG.buttonVisibility,
      ...(config?.buttonVisibility || {})
    };

    // Normalize section button layout (move/reorder buttons between sections)
    const normalizedSectionButtons: Record<SectionId, string[]> = sectionOrder.reduce((acc, id) => {
      acc[id] = [];
      return acc;
    }, {} as Record<SectionId, string[]>);

    const rawSectionButtons: any = (config as any)?.sectionButtons;
    const seen = new Set<string>();

    if (rawSectionButtons && typeof rawSectionButtons === 'object') {
      sectionOrder.forEach((sectionId) => {
        const list = rawSectionButtons[sectionId];
        if (!Array.isArray(list)) return;

        list.forEach((buttonId: any) => {
          if (typeof buttonId !== 'string') return;
          if (!BUTTON_META[buttonId]) return; // only allow known buttons
          if (seen.has(buttonId)) return;
          normalizedSectionButtons[sectionId].push(buttonId);
          seen.add(buttonId);
        });
      });
    }

    // Add any missing buttons back to their default section
    ALL_BUTTONS.forEach((btn) => {
      if (seen.has(btn.id)) return;
      const defaultSectionId = (DEFAULT_SECTIONS.find((s) => s.buttons.some((b) => b.id === btn.id))?.id ||
        'tools') as SectionId;
      normalizedSectionButtons[defaultSectionId].push(btn.id);
      seen.add(btn.id);
    });

    // Normalize section label overrides (trim; preserve empty strings so users can hide labels)
    const sectionLabels: Partial<Record<SectionId, string>> = {};
    const rawLabels: any = (config as any)?.sectionLabels;
    if (rawLabels && typeof rawLabels === 'object') {
      Object.keys(rawLabels).forEach((key) => {
        const value = rawLabels[key];
        if (typeof value === 'string') {
          sectionLabels[key] = value.trim();
        }
      });
    }

    return { sectionOrder, buttonVisibility, sectionLabels, sectionButtons: normalizedSectionButtons };
  };

  useEffect(() => {
    chrome.storage.sync.get([
      'showTool',
      'notificationDuration',
      'toolbarPosition',
      'schemaOverlayColor',
      'traceLogLimit',
      'skipPluginsByDefault',
      'toolbarConfig'
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
      if (result.toolbarConfig !== undefined) {
        setToolbarConfig(normalizeToolbarConfig(result.toolbarConfig));
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

  const handleSectionOrderChange = (newOrder: SectionId[]) => {
    const newConfig = { ...toolbarConfig, sectionOrder: newOrder };
    setToolbarConfig(newConfig);
    chrome.storage.sync.set({ toolbarConfig: newConfig });
  };

  const handleButtonVisibilityChange = (buttonId: string, visible: boolean) => {
    const newConfig = {
      ...toolbarConfig,
      buttonVisibility: { ...toolbarConfig.buttonVisibility, [buttonId]: visible }
    };
    setToolbarConfig(newConfig);
    chrome.storage.sync.set({ toolbarConfig: newConfig });
  };

  const handleSectionLabelChange = (sectionId: SectionId, value: string, defaultLabel: string) => {
    const trimmed = value.trim();
    const current = { ...(toolbarConfig.sectionLabels || {}) };
    const isDefaultSection = DEFAULT_SECTION_IDS.includes(sectionId);

    // Allow blank labels. Only "same as default" removes the override.
    if (isDefaultSection && trimmed === defaultLabel) {
      delete current[sectionId];
    } else {
      current[sectionId] = trimmed;
    }

    const newConfig: ToolbarConfig = { ...toolbarConfig, sectionLabels: current };
    setToolbarConfig(newConfig);
    chrome.storage.sync.set({ toolbarConfig: newConfig });
  };

  const handleResetSectionLabel = (sectionId: SectionId) => {
    const current = { ...(toolbarConfig.sectionLabels || {}) };
    delete current[sectionId];
    const newConfig: ToolbarConfig = { ...toolbarConfig, sectionLabels: current };
    setToolbarConfig(newConfig);
    chrome.storage.sync.set({ toolbarConfig: newConfig });
  };

  const getButtonsForSection = (sectionId: SectionId): string[] => {
    const fromConfig = toolbarConfig.sectionButtons?.[sectionId];
    // IMPORTANT: empty arrays are valid (section intentionally has no buttons)
    if (Array.isArray(fromConfig)) return fromConfig;
    return DEFAULT_SECTION_BUTTONS[sectionId] || [];
  };

  const setToolbarConfigAndPersist = (next: ToolbarConfig) => {
    setToolbarConfig(next);
    chrome.storage.sync.set({ toolbarConfig: next });
  };

  const setSectionButtons = (sectionButtons: Partial<Record<SectionId, string[]>>) => {
    setToolbarConfigAndPersist({ ...toolbarConfig, sectionButtons });
  };

  const handleMoveButton = (buttonId: string, fromSectionId: SectionId, toSectionId: SectionId, toIndex?: number) => {
    const current: Record<SectionId, string[]> = {
      ...DEFAULT_SECTION_BUTTONS,
      ...(toolbarConfig.sectionButtons || {} as any)
    } as any;

    // Clone arrays
    const allSectionIds = toolbarConfig.sectionOrder;
    const next: Record<SectionId, string[]> = allSectionIds.reduce((acc, id) => {
      acc[id] = [...(current[id] || [])];
      return acc;
    }, {} as Record<SectionId, string[]>);

    const fromList = next[fromSectionId];
    const toList = next[toSectionId];

    const fromIndex = fromList.indexOf(buttonId);
    if (fromIndex >= 0) {
      fromList.splice(fromIndex, 1);
    }

    let insertIndex = typeof toIndex === 'number' ? toIndex : toList.length;
    if (fromSectionId === toSectionId && fromIndex >= 0 && typeof toIndex === 'number' && fromIndex < toIndex) {
      insertIndex = Math.max(0, insertIndex - 1);
    }

    // Prevent duplicates
    const existingIndex = toList.indexOf(buttonId);
    if (existingIndex >= 0) {
      toList.splice(existingIndex, 1);
      if (existingIndex < insertIndex) insertIndex = Math.max(0, insertIndex - 1);
    }

    toList.splice(insertIndex, 0, buttonId);

    setSectionButtons(next);
  };

  const handleButtonDragStart = (sectionId: SectionId, buttonId: string) => {
    setDraggedButton({ buttonId, fromSectionId: sectionId });
  };

  const handleButtonDragOver = (e: React.DragEvent) => {
    e.preventDefault();
  };

  const handleButtonDrop = (e: React.DragEvent, toSectionId: SectionId, toIndex?: number) => {
    e.preventDefault();
    e.stopPropagation();
    if (!draggedButton) return;
    handleMoveButton(draggedButton.buttonId, draggedButton.fromSectionId, toSectionId, toIndex);
    setDraggedButton(null);
  };

  const handleDragStart = (index: number) => {
    setDraggedSectionIndex(index);
  };

  const handleDragOver = (e: React.DragEvent, index: number) => {
    e.preventDefault();
  };

  const handleDrop = (e: React.DragEvent, dropIndex: number) => {
    e.preventDefault();

    // If a button is being dragged, dropping on a section card moves it to that section (append)
    if (draggedButton) {
      const toSectionId = toolbarConfig.sectionOrder[dropIndex];
      if (toSectionId) {
        handleMoveButton(draggedButton.buttonId, draggedButton.fromSectionId, toSectionId);
      }
      setDraggedButton(null);
      return;
    }

    if (draggedSectionIndex === null) return;

    const newOrder = [...toolbarConfig.sectionOrder];
    const [draggedSection] = newOrder.splice(draggedSectionIndex, 1);
    newOrder.splice(dropIndex, 0, draggedSection);

    handleSectionOrderChange(newOrder);
    setDraggedSectionIndex(null);
  };

  const handleAddSection = () => {
    // Create a stable, human-friendly id like custom-1, custom-2, ...
    const existing = new Set(toolbarConfig.sectionOrder);
    let n = 1;
    while (existing.has(`custom-${n}`)) n++;
    const id = `custom-${n}`;

    const sectionOrder = [...toolbarConfig.sectionOrder, id];
    const sectionButtons = { ...(toolbarConfig.sectionButtons || {}), [id]: [] as string[] };
    const sectionLabels = { ...(toolbarConfig.sectionLabels || {}), [id]: `Custom ${n}` };

    const next: ToolbarConfig = { ...toolbarConfig, sectionOrder, sectionButtons, sectionLabels };
    setToolbarConfig(next);
    chrome.storage.sync.set({ toolbarConfig: next });
  };

  const handleRemoveCustomSection = (sectionId: SectionId) => {
    if (DEFAULT_SECTION_IDS.includes(sectionId)) return;
    const confirmed = window.confirm(`Remove section \"${toolbarConfig.sectionLabels?.[sectionId] || sectionId}\"?`);
    if (!confirmed) return;

    // Move any buttons in this section back to Tools (append)
    const toolsId = 'tools';
    const currentButtons = toolbarConfig.sectionButtons?.[sectionId] || [];
    const nextButtons: any = { ...(toolbarConfig.sectionButtons || {}) };
    delete nextButtons[sectionId];
    nextButtons[toolsId] = [...(nextButtons[toolsId] || DEFAULT_SECTION_BUTTONS[toolsId] || []), ...currentButtons];

    const nextLabels: any = { ...(toolbarConfig.sectionLabels || {}) };
    delete nextLabels[sectionId];

    const sectionOrder = toolbarConfig.sectionOrder.filter((id) => id !== sectionId);

    const next: ToolbarConfig = { ...toolbarConfig, sectionOrder, sectionButtons: nextButtons, sectionLabels: nextLabels };
    setToolbarConfig(next);
    chrome.storage.sync.set({ toolbarConfig: next });
  };

  const handleResetToolbarCustomization = () => {
    const confirmed = window.confirm(
      'Reset toolbar customization to defaults?\n\nThis will restore original section names, section order, button layout, and visibility.'
    );
    if (!confirmed) return;

    const next = normalizeToolbarConfig(DEFAULT_TOOLBAR_CONFIG);
    setToolbarConfig(next);
    chrome.storage.sync.set({ toolbarConfig: next });
  };

  const handleExportSettings = async () => {
    setSettingsTransferMessage('');
    try {
      const keys = [
        'showTool',
        'notificationDuration',
        'toolbarPosition',
        'schemaOverlayColor',
        'traceLogLimit',
        'skipPluginsByDefault',
        'toolbarConfig'
      ];

      const result = await new Promise<any>((resolve) => chrome.storage.sync.get(keys, resolve));
      const payload = {
        exportedAt: new Date().toISOString(),
        extensionVersion: chrome.runtime.getManifest().version,
        data: {
          ...result,
          toolbarConfig: normalizeToolbarConfig(result.toolbarConfig)
        }
      };

      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `d365-helper-settings-${payload.extensionVersion}.json`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      setSettingsTransferMessage('Settings exported.');
    } catch (err: any) {
      setSettingsTransferMessage(err?.message ? `Export failed: ${err.message}` : 'Export failed.');
    }
  };

  const handleImportSettingsClick = () => {
    setSettingsTransferMessage('');
    importFileInputRef.current?.click();
  };

  const handleImportSettingsFile = async (file: File) => {
    setSettingsTransferMessage('');
    try {
      const text = await file.text();
      const parsed = JSON.parse(text);
      const imported = parsed?.data && typeof parsed.data === 'object' ? parsed.data : parsed;

      if (!imported || typeof imported !== 'object') {
        throw new Error('Invalid settings file format.');
      }

      const confirmed = window.confirm('Import settings from file? This will overwrite your current extension settings.');
      if (!confirmed) return;

      // Only apply known keys
      const next: any = {};
      if (typeof imported.showTool === 'boolean') next.showTool = imported.showTool;
      if (typeof imported.notificationDuration === 'number') next.notificationDuration = imported.notificationDuration;
      if (imported.toolbarPosition === 'top' || imported.toolbarPosition === 'bottom') next.toolbarPosition = imported.toolbarPosition;
      if (typeof imported.schemaOverlayColor === 'string') next.schemaOverlayColor = imported.schemaOverlayColor;
      if (typeof imported.traceLogLimit === 'number') next.traceLogLimit = imported.traceLogLimit;
      if (typeof imported.skipPluginsByDefault === 'boolean') next.skipPluginsByDefault = imported.skipPluginsByDefault;
      if (imported.toolbarConfig && typeof imported.toolbarConfig === 'object') {
        next.toolbarConfig = normalizeToolbarConfig(imported.toolbarConfig);
      }

      await new Promise<void>((resolve) => chrome.storage.sync.set(next, () => resolve()));

      // Refresh UI state from storage
      const refreshed = await new Promise<any>((resolve) =>
        chrome.storage.sync.get(
          [
            'showTool',
            'notificationDuration',
            'toolbarPosition',
            'schemaOverlayColor',
            'traceLogLimit',
            'skipPluginsByDefault',
            'toolbarConfig'
          ],
          resolve
        )
      );

      if (refreshed.showTool !== undefined) setShowTool(refreshed.showTool);
      if (refreshed.notificationDuration !== undefined) setNotificationDuration(refreshed.notificationDuration);
      if (refreshed.toolbarPosition !== undefined) setToolbarPosition(refreshed.toolbarPosition);
      if (refreshed.schemaOverlayColor !== undefined) setSchemaOverlayColor(refreshed.schemaOverlayColor);
      if (refreshed.traceLogLimit !== undefined) setTraceLogLimit(refreshed.traceLogLimit);
      if (refreshed.skipPluginsByDefault !== undefined) setSkipPluginsByDefault(refreshed.skipPluginsByDefault);
      if (refreshed.toolbarConfig !== undefined) setToolbarConfig(normalizeToolbarConfig(refreshed.toolbarConfig));

      setSettingsTransferMessage('Settings imported.');
    } catch (err: any) {
      setSettingsTransferMessage(err?.message ? `Import failed: ${err.message}` : 'Import failed.');
    }
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

            <h3 className="settings-section-title">Toolbar Customization</h3>
            <p className="settings-hint">
              Drag sections to reorder them. Drag buttons between sections to regroup them. Uncheck buttons to hide them from the toolbar.
            </p>

            <div className="toolbar-customization-actions">
              <button className="settings-transfer-btn" type="button" onClick={handleAddSection}>
                + Add section
              </button>
              <button className="settings-transfer-btn" type="button" onClick={handleResetToolbarCustomization}>
                Reset layout
              </button>
            </div>

            <div className="toolbar-customization">
              {toolbarConfig.sectionOrder.map((sectionId, index) => {
                const defaultSection = DEFAULT_SECTIONS.find(s => s.id === sectionId);
                const isDefaultSection = Boolean(defaultSection);

                const defaultLabel = defaultSection?.label || sectionId;
                const storedLabel = toolbarConfig.sectionLabels?.[sectionId];
                const hasCustomLabel = storedLabel !== undefined;
                const inputValue = hasCustomLabel ? storedLabel : defaultLabel;

                return (
                  <div
                    key={sectionId}
                    className={`toolbar-section-card ${draggedSectionIndex === index ? 'dragging' : ''}`}
                    onDragOver={(e) => handleDragOver(e, index)}
                    onDrop={(e) => handleDrop(e, index)}
                  >
                    <div className="toolbar-section-card-header">
                      <span
                        className="drag-handle"
                        aria-hidden="true"
                        draggable
                        onDragStart={() => handleDragStart(index)}
                      >
                        &#8942;&#8942;
                      </span>
                      <input
                        className="toolbar-section-card-title"
                        type="text"
                        value={inputValue}
                        onChange={(e) => handleSectionLabelChange(sectionId, e.target.value, defaultLabel)}
                        aria-label={`${defaultLabel} section name`}
                      />
                      {isDefaultSection ? (
                        <button
                          type="button"
                          className="toolbar-section-card-title-reset"
                          onClick={() => handleResetSectionLabel(sectionId)}
                          disabled={!hasCustomLabel}
                          title="Reset to default section name"
                        >
                          Reset
                        </button>
                      ) : (
                        <button
                          type="button"
                          className="toolbar-section-card-title-reset toolbar-section-card-title-remove"
                          onClick={() => handleRemoveCustomSection(sectionId)}
                          title="Remove this custom section"
                        >
                          Remove
                        </button>
                      )}
                    </div>
                    <div
                      className="toolbar-section-card-buttons"
                      onDragOver={handleButtonDragOver}
                      onDrop={(e) => handleButtonDrop(e, sectionId)}
                    >
                      {getButtonsForSection(sectionId).map((buttonId, buttonIndex) => {
                        const button = BUTTON_META[buttonId];
                        if (!button) return null;

                        return (
                          <div
                            key={button.id}
                            className="toolbar-button-row"
                            onDragOver={handleButtonDragOver}
                            onDrop={(e) => handleButtonDrop(e, sectionId, buttonIndex)}
                          >
                            <span
                              className="drag-handle toolbar-button-drag-handle"
                              aria-hidden="true"
                              draggable
                              onDragStart={() => handleButtonDragStart(sectionId, button.id)}
                            >
                              &#8942;&#8942;
                            </span>
                            <label className="toolbar-button-checkbox">
                              <input
                                type="checkbox"
                                checked={toolbarConfig.buttonVisibility[button.id] ?? true}
                                onChange={(e) => handleButtonVisibilityChange(button.id, e.target.checked)}
                              />
                              <span className="toolbar-button-label">{button.label}</span>
                              <span className="toolbar-button-description">{button.description}</span>
                            </label>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
            </div>

            <h3 className="settings-section-title">Settings Transfer</h3>
            <section className="info-section">
              <p className="settings-hint" style={{ marginTop: 0 }}>
                Export/import your extension settings to move them between Chrome profiles.
              </p>
              <div className="settings-transfer-actions">
                <button className="settings-transfer-btn" type="button" onClick={handleExportSettings}>
                  Export settings
                </button>
                <button className="settings-transfer-btn" type="button" onClick={handleImportSettingsClick}>
                  Import settings
                </button>
                <input
                  ref={importFileInputRef}
                  type="file"
                  accept="application/json"
                  className="settings-transfer-file"
                  onChange={async (e) => {
                    const file = e.target.files?.[0];
                    if (file) {
                      await handleImportSettingsFile(file);
                    }
                    // allow selecting the same file again
                    e.currentTarget.value = '';
                  }}
                />
              </div>
              {settingsTransferMessage && (
                <div className="settings-transfer-message">{settingsTransferMessage}</div>
              )}
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
