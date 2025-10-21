import React, { useState, useEffect } from 'react';
import './WebAPIViewer.css';

interface ApiData {
  [key: string]: any;
}

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
  const [bypassPlugins, setBypassPlugins] = useState<boolean>(false);

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

  const handleSaveChanges = async () => {
    if (!editedData || !apiUrl) return;

    setSaving(true);
    setError('');

    try {
      // Prepare the update payload - only include changed fields
      const updatePayload: any = {};

      for (const key in editedData) {
        // Skip metadata fields and navigation properties
        if (!key.startsWith('@') &&
            !key.startsWith('_') &&
            key !== 'odata.etag' &&
            editedData[key] !== data?.[key]) {
          updatePayload[key] = editedData[key];
        }
      }

      const headers: any = {
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
      };

      // Add plugin bypass headers if requested
      if (bypassPlugins) {
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
      // Entering edit mode - copy current data
      setEditedData(JSON.parse(JSON.stringify(data)));
    } else {
      // Exiting edit mode - discard changes
      setEditedData(null);
    }
    setEditMode(!editMode);
  };

  const handleValueChange = (key: string, newValue: any) => {
    if (!editedData) return;

    setEditedData({
      ...editedData,
      [key]: newValue
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
    // Allow lookup fields (_fieldname_value) to be editable, but not other metadata fields
    const isLookupField = key.startsWith('_') && key.endsWith('_value');
    const isEditable = editMode && level === 0 && !key.startsWith('@') && (isLookupField || !key.startsWith('_'));
    const currentValue = editMode && editedData ? editedData[key] : value;

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
          <button className="copy-btn" onClick={() => handleCopyValue(null)} title="Copy value">📋</button>
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
          <button className="copy-btn" onClick={() => handleCopyValue(currentValue)} title="Copy value">📋</button>
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
          <button className="copy-btn" onClick={() => handleCopyValue(currentValue)} title="Copy value">📋</button>
        </>
      );
    }

    if (typeof currentValue === 'string') {
      // Check if it's a lookup field (navigation property)
      const isLookup = key.startsWith('_') && key.endsWith('_value');

      // Check if it's a date
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
            <button className="copy-btn" onClick={() => handleCopyValue(currentValue)} title="Copy value">📋</button>
          </>
        );
      }

      return (
        <>
          {isEditable ? (
            isLookup ? (
              <input
                type="text"
                className="edit-input edit-lookup"
                value={currentValue}
                onChange={(e) => handleValueChange(key, e.target.value)}
                placeholder="Lookup GUID"
                title="Enter lookup GUID value"
              />
            ) : (
              <input
                type="text"
                className="edit-input"
                value={currentValue}
                onChange={(e) => handleValueChange(key, e.target.value)}
              />
            )
          ) : (
            <span className="value-string">"{currentValue}"</span>
          )}
          <button className="copy-btn" onClick={() => handleCopyValue(currentValue)} title="Copy value">📋</button>
        </>
      );
    }

    if (Array.isArray(value)) {
      const sectionKey = `${key}_${level}`;
      const isExpanded = expandedSections.has(sectionKey);

      return (
        <div className="value-array">
          <span className="toggle" onClick={() => toggleSection(sectionKey)}>
            {isExpanded ? '▼' : '▶'} Array[{value.length}]
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
            {isExpanded ? '▼' : '▶'} Object ({keys.length} properties)
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
                  checked={bypassPlugins}
                  onChange={(e) => setBypassPlugins(e.target.checked)}
                />
                <span>Bypass Plugins</span>
              </label>
              <button onClick={handleSaveChanges} disabled={saving || !editedData} className="btn btn-success">
                {saving ? 'Saving...' : '💾 Save'}
              </button>
              <button onClick={handleEditModeToggle} disabled={saving} className="btn btn-secondary">
                ✕ Cancel
              </button>
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
