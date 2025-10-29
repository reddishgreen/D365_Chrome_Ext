import React, { useState, useEffect, useMemo, useCallback, useRef } from 'react';
import './WebAPIViewer.css';
import LookupFieldEditor from './LookupFieldEditor';
import { LookupEntityMetadata, LookupSelection } from './lookupTypes';

interface ApiData {
  [key: string]: any;
}

interface ApiPathInfo {
  baseUrl: string;
  entitySetName: string;
}

const deriveAttributeName = (lookupKey: string): string => {
  if (lookupKey.startsWith('_') && lookupKey.endsWith('_value')) {
    return lookupKey.slice(1, -6);
  }
  return lookupKey;
};

const escapeODataIdentifier = (value: string): string => value.replace(/'/g, "''");

const WebAPIViewer: React.FC = () => {
  const [apiUrl, setApiUrl] = useState<string>('');
  const [data, setData] = useState<ApiData | null>(null);
  const [editedData, setEditedData] = useState<ApiData | null>(null);
  const [loading, setLoading] = useState<boolean>(false);
  const [saving, setSaving] = useState<boolean>(false);
  const [error, setError] = useState<string>('');
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [expandedSections, setExpandedSections] = useState<Set<string>>(new Set(['root']));
  const [editMode, setEditMode] = useState<boolean>(false);
  const [skipPluginExecution, setSkipPluginExecution] = useState<boolean>(false);
  const [entityLogicalName, setEntityLogicalName] = useState<string | null>(null);

  const entityMetadataCache = useRef<Map<string, LookupEntityMetadata>>(new Map());
  const entitySetToLogicalCache = useRef<Map<string, string>>(new Map());
  const lookupTargetsCache = useRef<Map<string, string[]>>(new Map());

  const apiPathInfo = useMemo<ApiPathInfo>(() => {
    if (!apiUrl) {
      return { baseUrl: '', entitySetName: '' };
    }

    const baseMatch = apiUrl.match(/^(https?:\/\/[^/]+\/api\/data\/v[0-9.]+\/)/i);
    const baseUrl = baseMatch ? baseMatch[1] : '';
    let entitySetName = '';

    if (baseUrl && apiUrl.length > baseUrl.length) {
      const remainder = apiUrl.substring(baseUrl.length);
      const entityMatch = remainder.match(/^([^/?(]+)\(/);
      if (entityMatch) {
        entitySetName = entityMatch[1];
      }
    }

    return { baseUrl, entitySetName };
  }, [apiUrl]);

  const apiBaseUrl = apiPathInfo.baseUrl;
  const currentEntitySetName = apiPathInfo.entitySetName;

  useEffect(() => {
    entityMetadataCache.current.clear();
    lookupTargetsCache.current.clear();
    entitySetToLogicalCache.current.clear();
  }, [apiBaseUrl]);

  useEffect(() => {
    // Get URL from query parameter
    const params = new URLSearchParams(window.location.search);
    const url = params.get('url');

    if (url) {
      setApiUrl(decodeURIComponent(url));
      fetchData(decodeURIComponent(url));
    }
  }, []);

  const fetchData = async (url: string) => {
    setLoading(true);
    setError('');

    try {
      const response = await fetch(url, {
        headers: {
          'Accept': 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
        },
        credentials: 'include'
      });

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      const jsonData = await response.json();
      setData(jsonData);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to fetch data');
    } finally {
      setLoading(false);
    }
  };

  const getEntityMetadata = useCallback(
    async (logicalName: string): Promise<LookupEntityMetadata | null> => {
      if (!logicalName || !apiBaseUrl) {
        return null;
      }

      const cacheKey = logicalName.toLowerCase();
      const cached = entityMetadataCache.current.get(cacheKey);
      if (cached) {
        return cached;
      }

      const metadataUrl = `${apiBaseUrl}EntityDefinitions(LogicalName='${escapeODataIdentifier(
        logicalName
      )}')?$select=EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute`;

      const response = await fetch(metadataUrl, {
        headers: {
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
        },
        credentials: 'include',
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Metadata request failed (${response.status}): ${errorText}`);
      }

      const json = await response.json();
      const metadata: LookupEntityMetadata = {
        logicalName,
        entitySetName: json?.EntitySetName,
        primaryIdAttribute: json?.PrimaryIdAttribute,
        primaryNameAttribute: json?.PrimaryNameAttribute ?? null,
      };

      if (!metadata.entitySetName || !metadata.primaryIdAttribute) {
        throw new Error(`Metadata incomplete for entity ${logicalName}`);
      }

      entityMetadataCache.current.set(cacheKey, metadata);
      return metadata;
    },
    [apiBaseUrl]
  );

  const getEntityLogicalName = useCallback(
    async (entitySetName: string): Promise<string | null> => {
      if (!entitySetName || !apiBaseUrl) {
        return null;
      }

      const cacheKey = entitySetName.toLowerCase();
      const cached = entitySetToLogicalCache.current.get(cacheKey);
      if (cached) {
        return cached;
      }

      const logicalUrl = `${apiBaseUrl}EntityDefinitions?$select=LogicalName&$filter=EntitySetName eq '${escapeODataIdentifier(
        entitySetName
      )}'`;

      const response = await fetch(logicalUrl, {
        headers: {
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
        },
        credentials: 'include',
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to resolve logical name (${response.status}): ${errorText}`);
      }

      const json = await response.json();
      const logicalName =
        Array.isArray(json?.value) && json.value.length > 0
          ? json.value[0]?.LogicalName ?? null
          : null;

      if (logicalName) {
        entitySetToLogicalCache.current.set(cacheKey, logicalName);
      }

      return logicalName;
    },
    [apiBaseUrl]
  );

  const getLookupTargets = useCallback(
    async (entityLogical: string, attributeName: string): Promise<string[]> => {
      if (!entityLogical || !attributeName || !apiBaseUrl) {
        return [];
      }

      const cacheKey = `${entityLogical}.${attributeName}`.toLowerCase();
      const cached = lookupTargetsCache.current.get(cacheKey);
      if (cached) {
        return cached;
      }

      try {
        const attributeUrl = `${apiBaseUrl}EntityDefinitions(LogicalName='${escapeODataIdentifier(
          entityLogical
        )}')/Attributes(LogicalName='${escapeODataIdentifier(
          attributeName
        )}')/Microsoft.Dynamics.CRM.LookupAttributeMetadata?$select=Targets`;

        const response = await fetch(attributeUrl, {
          headers: {
            Accept: 'application/json',
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
          },
          credentials: 'include',
        });

        if (!response.ok) {
          // If attribute doesn't exist or isn't a lookup, return empty array
          // The lookup editor will use the current logical name as fallback
          console.warn(`Could not load lookup targets for ${entityLogical}.${attributeName}: ${response.status}`);
          const emptyTargets: string[] = [];
          lookupTargetsCache.current.set(cacheKey, emptyTargets);
          return emptyTargets;
        }

        const json = await response.json();
        let targets: string[] = [];

        if (Array.isArray(json?.Targets)) {
          targets = json.Targets.filter((item: unknown): item is string => typeof item === 'string');
        } else if (Array.isArray(json?.value)) {
          json.value.forEach((item: any) => {
            if (Array.isArray(item?.Targets)) {
              targets.push(
                ...item.Targets.filter((target: unknown): target is string => typeof target === 'string')
              );
            }
          });
        }

        lookupTargetsCache.current.set(cacheKey, targets);
        return targets;
      } catch (error) {
        // Network error or other fetch error - return empty array
        console.warn(`Error loading lookup targets for ${entityLogical}.${attributeName}:`, error);
        const emptyTargets: string[] = [];
        lookupTargetsCache.current.set(cacheKey, emptyTargets);
        return emptyTargets;
      }
    },
    [apiBaseUrl]
  );

  useEffect(() => {
    let cancelled = false;

    const resolveEntityName = async () => {
      if (!currentEntitySetName) {
        setEntityLogicalName(null);
        return;
      }

      try {
        const logical = await getEntityLogicalName(currentEntitySetName);
        if (!cancelled) {
          setEntityLogicalName(logical);
        }
      } catch (err) {
        console.error('Failed to resolve entity logical name', err);
        if (!cancelled) {
          setEntityLogicalName(null);
        }
      }
    };

    resolveEntityName();

    return () => {
      cancelled = true;
    };
  }, [currentEntitySetName, getEntityLogicalName]);

  const handleRefresh = () => {
    if (apiUrl) {
      fetchData(apiUrl);
    }
  };

  const handleCopyAll = async () => {
    if (data) {
      await navigator.clipboard.writeText(JSON.stringify(data, null, 2));
      alert('Data copied to clipboard!');
    }
  };

  const handleCopyUrl = async () => {
    await navigator.clipboard.writeText(apiUrl);
    alert('API URL copied to clipboard!');
  };

  const handleLookupSelection = useCallback(
    (lookupKey: string, selection: LookupSelection | null) => {
      setEditedData((previous) => {
        const base = previous ?? (data ? { ...data } : {});
        const updated: ApiData = { ...base };

        const attributeName = deriveAttributeName(lookupKey);
        const formattedKey = `${lookupKey}@OData.Community.Display.V1.FormattedValue`;
        const logicalKey = `${lookupKey}@Microsoft.Dynamics.CRM.lookuplogicalname`;
        const bindKey = `${attributeName}@odata.bind`;

        if (selection) {
          updated[lookupKey] = selection.recordId;
          updated[formattedKey] = selection.displayName;
          updated[logicalKey] = selection.logicalName;
          updated[bindKey] = `/${selection.entitySetName}(${selection.recordId})`;
        } else {
          updated[lookupKey] = null;
          updated[bindKey] = null;
          delete updated[formattedKey];
          delete updated[logicalKey];
        }

        return updated;
      });
    },
    [data]
  );

  const handleSaveChanges = async () => {
    if (!editedData || !apiUrl) return;

    setSaving(true);
    setError('');

    try {
      // Prepare the update payload - only include changed fields
      const updatePayload: any = {};

      for (const key in editedData) {
        if (key.endsWith('@odata.bind')) {
          if (editedData[key] !== data?.[key]) {
            updatePayload[key] = editedData[key];
          }
          continue;
        }

        if (
          !key.startsWith('@') &&
          !key.startsWith('_') &&
          key !== 'odata.etag' &&
          editedData[key] !== data?.[key]
        ) {
          updatePayload[key] = editedData[key];
        }
      }

      const headers: any = {
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
      };

      // Add headers to skip plugin execution if requested (for development/testing)
      if (skipPluginExecution) {
        headers['MSCRM.SuppressCallbackRegistrationExpanderJob'] = 'true';
        headers['MSCRM.BypassCustomPluginExecution'] = 'true';
      }

      const response = await fetch(apiUrl, {
        method: 'PATCH',
        headers: headers,
        credentials: 'include',
        body: JSON.stringify(updatePayload)
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`HTTP ${response.status}: ${errorText}`);
      }

      alert('Record updated successfully!');
      // Refresh the data
      await fetchData(apiUrl);
      setEditMode(false);
      setEditedData(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to save changes');
      alert('Error saving changes: ' + (err instanceof Error ? err.message : 'Unknown error'));
    } finally {
      setSaving(false);
    }
  };

  const handleEditModeToggle = () => {
    if (!editMode) {
      setEditedData(data ? JSON.parse(JSON.stringify(data)) : null);
    } else {
      setEditedData(null);
    }
    setEditMode((previous) => !previous);
  };

  const handleValueChange = (key: string, newValue: any) => {
    if (!editMode) {
      return;
    }

    setEditedData((previous) => {
      const base = previous ?? (data ? { ...data } : {});
      return {
        ...base,
        [key]: newValue,
      };
    });
  };

  const toggleSection = (key: string) => {
    const newExpanded = new Set(expandedSections);
    if (newExpanded.has(key)) {
      newExpanded.delete(key);
    } else {
      newExpanded.add(key);
    }
    setExpandedSections(newExpanded);
  };

  const handleCopyValue = async (value: any) => {
    const textToCopy = typeof value === 'string' ? value : JSON.stringify(value);
    await navigator.clipboard.writeText(textToCopy);
    // Show brief feedback
    const btn = event?.target as HTMLElement;
    if (btn) {
      const originalText = btn.textContent;
      btn.textContent = '✓';
      setTimeout(() => {
        btn.textContent = originalText || '';
      }, 1000);
    }
  };

  const renderValue = (value: any, key: string, level: number = 0, parentKey?: string): JSX.Element => {
    const isLookupField = key.startsWith('_') && key.endsWith('_value');
    const isEditable =
      editMode && level === 0 && !key.startsWith('@') && (isLookupField || !key.startsWith('_'));
    const currentValue = editMode && editedData ? editedData[key] : value;

    if (isLookupField) {
      const formattedKey = `${key}@OData.Community.Display.V1.FormattedValue`;
      const logicalKey = `${key}@Microsoft.Dynamics.CRM.lookuplogicalname`;
      const formattedValue =
        (editMode && editedData ? editedData[formattedKey] : undefined) ?? data?.[formattedKey];
      const lookupLogicalName =
        (editMode && editedData ? editedData[logicalKey] : undefined) ?? data?.[logicalKey];
      const attributeName = deriveAttributeName(key);
      const currentGuid =
        typeof currentValue === 'string' && currentValue ? currentValue : null;

      if (isEditable && apiBaseUrl) {
        const loadTargets =
          entityLogicalName && attributeName
            ? () => getLookupTargets(entityLogicalName, attributeName)
            : undefined;

        return (
          <>
            <LookupFieldEditor
              apiBaseUrl={apiBaseUrl}
              attributeName={attributeName}
              currentId={currentGuid}
              currentName={formattedValue ?? undefined}
              currentLogicalName={lookupLogicalName ?? undefined}
              loadTargets={loadTargets}
              getEntityMetadata={getEntityMetadata}
              onSelectionChange={(selection) => handleLookupSelection(key, selection)}
            />
            <button
              className="copy-btn"
              onClick={() => handleCopyValue(currentGuid)}
              title="Copy lookup GUID"
            >
              <img src={chrome.runtime.getURL('icons/rg_copy.svg')} alt="Copy" />
            </button>
          </>
        );
      }

      if (!formattedValue && !currentGuid) {
        return (
          <>
            <span className="value-null">null</span>
            <button className="copy-btn" onClick={() => handleCopyValue(null)} title="Copy value">
              <img src={chrome.runtime.getURL('icons/rg_copy.svg')} alt="Copy" />
            </button>
          </>
        );
      }

      const displayParts: string[] = [];
      if (formattedValue) {
        displayParts.push(String(formattedValue));
      }
      if (currentGuid) {
        displayParts.push(`(${currentGuid.toUpperCase()})`);
      }

      return (
        <>
          <span className="value-lookup">{displayParts.join(' ')}</span>
          <button
            className="copy-btn"
            onClick={() => handleCopyValue(currentGuid)}
            title="Copy lookup GUID"
          >
            <img src={chrome.runtime.getURL('icons/rg_copy.svg')} alt="Copy" />
          </button>
        </>
      );
    }

    if (value === null || typeof value === 'undefined') {
      return (
        <>
          {isEditable ? (
            <input
              type="text"
              className="edit-input"
              placeholder="null"
              value=""
              onChange={(e) => handleValueChange(key, e.target.value || null)}
            />
          ) : (
            <span className="value-null">null</span>
          )}
          <button className="copy-btn" onClick={() => handleCopyValue(null)} title="Copy value">
            <img src={chrome.runtime.getURL('icons/rg_copy.svg')} alt="Copy" />
          </button>
        </>
      );
    }

    if (typeof currentValue === 'boolean') {
      return (
        <>
          {isEditable ? (
            <select
              className="edit-select"
              value={currentValue.toString()}
              onChange={(e) => handleValueChange(key, e.target.value === 'true')}
            >
              <option value="true">true</option>
              <option value="false">false</option>
            </select>
          ) : (
            <span className="value-boolean">{currentValue.toString()}</span>
          )}
          <button className="copy-btn" onClick={() => handleCopyValue(currentValue)} title="Copy value">
            <img src={chrome.runtime.getURL('icons/rg_copy.svg')} alt="Copy" />
          </button>
        </>
      );
    }

    if (typeof currentValue === 'number') {
      return (
        <>
          {isEditable ? (
            <input
              type="number"
              className="edit-input"
              value={currentValue}
              onChange={(e) => handleValueChange(key, parseFloat(e.target.value))}
            />
          ) : (
            <span className="value-number">{currentValue}</span>
          )}
          <button className="copy-btn" onClick={() => handleCopyValue(currentValue)} title="Copy value">
            <img src={chrome.runtime.getURL('icons/rg_copy.svg')} alt="Copy" />
          </button>
        </>
      );
    }

    if (typeof currentValue === 'string') {
      if (currentValue.match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/)) {
        return (
          <>
            {isEditable ? (
              <input
                type="datetime-local"
                className="edit-input"
                value={currentValue.substring(0, 16)}
                onChange={(e) => handleValueChange(key, new Date(e.target.value).toISOString())}
              />
            ) : (
              <span className="value-string">
                "{currentValue}" <span className="value-hint">({new Date(currentValue).toLocaleString()})</span>
              </span>
            )}
            <button className="copy-btn" onClick={() => handleCopyValue(currentValue)} title="Copy value">
              <img src={chrome.runtime.getURL('icons/rg_copy.svg')} alt="Copy" />
            </button>
          </>
        );
      }

      return (
        <>
          {isEditable ? (
            <input
              type="text"
              className="edit-input"
              value={currentValue}
              onChange={(e) => handleValueChange(key, e.target.value)}
            />
          ) : (
            <span className="value-string">"{currentValue}"</span>
          )}
          <button className="copy-btn" onClick={() => handleCopyValue(currentValue)} title="Copy value">
            <img src={chrome.runtime.getURL('icons/rg_copy.svg')} alt="Copy" />
          </button>
        </>
      );
    }

    if (Array.isArray(value)) {
      const sectionKey = `${key}_${level}`;
      const isExpanded = expandedSections.has(sectionKey);

      return (
        <div className="value-array">
          <span className="toggle" onClick={() => toggleSection(sectionKey)}>
            {isExpanded ? '' : '?'} Array[{value.length}]
          </span>
          {isExpanded && (
            <div className="nested-content">
              {value.map((item, index) => (
                <div key={index} className="property">
                  <span className="property-key">[{index}]:</span>
                  {renderValue(item, `${key}_${index}`, level + 1)}
                </div>
              ))}
            </div>
          )}
        </div>
      );
    }

    if (typeof value === 'object') {
      const sectionKey = `${key}_${level}`;
      const isExpanded = expandedSections.has(sectionKey);
      const keys = Object.keys(value);

      return (
        <div className="value-object">
          <span className="toggle" onClick={() => toggleSection(sectionKey)}>
            {isExpanded ? '' : '?'} Object ({keys.length} properties)
          </span>
          {isExpanded && (
            <div className="nested-content">
              {keys.map((k) => (
                <div key={k} className="property">
                  <span className="property-key">{k}:</span>
                  {renderValue(value[k], `${key}_${k}`, level + 1)}
                </div>
              ))}
            </div>
          )}
        </div>
      );
    }

    return <span>{String(value)}</span>;
  };


const filterData = (obj: any, term: string): any => {
    if (!term) return obj;

    const lowerTerm = term.toLowerCase();

    if (typeof obj === 'object' && obj !== null) {
      const filtered: any = Array.isArray(obj) ? [] : {};

      for (const key in obj) {
        if (key.toLowerCase().includes(lowerTerm)) {
          filtered[key] = obj[key];
        } else if (typeof obj[key] === 'string' && obj[key].toLowerCase().includes(lowerTerm)) {
          filtered[key] = obj[key];
        } else if (typeof obj[key] === 'object') {
          const nested = filterData(obj[key], term);
          if (nested && Object.keys(nested).length > 0) {
            filtered[key] = nested;
          }
        }
      }

      return Object.keys(filtered).length > 0 ? filtered : null;
    }

    return obj;
  };

  const displayData = searchTerm ? filterData(data, searchTerm) : data;

  return (
    <div className="webapi-viewer">
      <header className="viewer-header">
        <h1>⚡ D365 Web API Viewer</h1>
        <div className="header-actions">
          <button onClick={handleRefresh} disabled={loading || saving} className="btn btn-primary">
            🔄 Refresh
          </button>
          {editMode ? (
            <>
              <label className="checkbox-label-inline">
                <input
                  type="checkbox"
                  checked={skipPluginExecution}
                  onChange={(e) => setSkipPluginExecution(e.target.checked)}
                />
                <span>Skip Plugin Execution</span>
              </label>
              <button onClick={handleSaveChanges} disabled={saving || !editedData} className="btn btn-success">
                {saving ? 'Saving...' : '💾 Save'}
              </button>
              <button onClick={handleEditModeToggle} disabled={saving} className="btn btn-secondary">
                ✕ Cancel
              </button>
              <div className="warning-text" style={{marginLeft: '10px', color: '#ff6b6b', fontSize: '12px'}}>
                ⚠️ For authorized development/testing only
              </div>
            </>
          ) : (
            <button onClick={handleEditModeToggle} disabled={!data} className="btn btn-primary">
              ✏️ Edit
            </button>
          )}
          <button onClick={handleCopyUrl} className="btn btn-secondary">
            📋 Copy URL
          </button>
          <button onClick={handleCopyAll} disabled={!data} className="btn btn-secondary">
            📋 Copy All
          </button>
        </div>
      </header>

      <div className="viewer-toolbar">
        <div className="url-display">
          <strong>API Endpoint:</strong>
          <code>{apiUrl}</code>
        </div>
        <div className="search-box">
          <input
            type="text"
            placeholder="Search in data..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="search-input"
          />
          {searchTerm && (
            <button onClick={() => setSearchTerm('')} className="clear-search">
              ✕
            </button>
          )}
        </div>
      </div>

      <main className="viewer-content">
        {loading && (
          <div className="loading">
            <div className="spinner"></div>
            <p>Loading data...</p>
          </div>
        )}

        {error && (
          <div className="error">
            <h2>Error</h2>
            <p>{error}</p>
            <button onClick={handleRefresh} className="btn btn-primary">
              Try Again
            </button>
          </div>
        )}

        {!loading && !error && displayData && (
          <div className="data-container">
            {Object.keys(displayData).map((key) => (
              <div key={key} className="property">
                <span className="property-key">{key}:</span>
                {renderValue(displayData[key], key)}
              </div>
            ))}
          </div>
        )}

        {!loading && !error && searchTerm && !displayData && (
          <div className="no-results">
            <p>No results found for "{searchTerm}"</p>
          </div>
        )}
      </main>
    </div>
  );
};

export default WebAPIViewer;
