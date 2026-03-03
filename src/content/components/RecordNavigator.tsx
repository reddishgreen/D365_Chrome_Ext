import React, { useState, useEffect, useMemo, useRef } from 'react';
import './RecordNavigator.css';

export interface EntityInfo {
  LogicalName: string;
  DisplayName: string;
  EntitySetName: string;
  PrimaryIdAttribute: string;
  PrimaryNameAttribute: string | null;
}

interface RecordNavigatorProps {
  entities: EntityInfo[] | null;   // null = loading
  orgUrl: string;
  onClose: () => void;
  error?: string;
}

const GUID_REGEX = /^[0-9a-f]{8}-?[0-9a-f]{4}-?[0-9a-f]{4}-?[0-9a-f]{4}-?[0-9a-f]{12}$/i;

const cleanGuid = (raw: string): string => {
  return raw.replace(/[{}\s]/g, '').trim();
};

const isValidGuid = (value: string): boolean => {
  return GUID_REGEX.test(cleanGuid(value));
};

const formatGuid = (raw: string): string => {
  const clean = cleanGuid(raw).replace(/-/g, '').toLowerCase();
  if (clean.length !== 32) return raw;
  return `${clean.slice(0, 8)}-${clean.slice(8, 12)}-${clean.slice(12, 16)}-${clean.slice(16, 20)}-${clean.slice(20)}`;
};

const RecordNavigator: React.FC<RecordNavigatorProps> = ({
  entities,
  orgUrl,
  onClose,
  error,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedEntity, setSelectedEntity] = useState<EntityInfo | null>(null);
  const [showResults, setShowResults] = useState(false);
  const [activeIndex, setActiveIndex] = useState(-1);
  const [guidValue, setGuidValue] = useState('');
  const [guidTouched, setGuidTouched] = useState(false);

  const entityInputRef = useRef<HTMLInputElement>(null);
  const guidInputRef = useRef<HTMLInputElement>(null);
  const resultsRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (entityInputRef.current) {
      entityInputRef.current.focus();
    }
  }, []);

  const filteredEntities = useMemo(() => {
    if (!entities) return [];
    if (!searchTerm) return entities;
    const lower = searchTerm.toLowerCase();
    return entities.filter(e =>
      e.LogicalName.toLowerCase().includes(lower) ||
      e.DisplayName.toLowerCase().includes(lower)
    );
  }, [entities, searchTerm]);

  const handleSelectEntity = (entity: EntityInfo) => {
    setSelectedEntity(entity);
    setSearchTerm('');
    setShowResults(false);
    setActiveIndex(-1);
    // Focus the GUID input after selecting entity
    setTimeout(() => guidInputRef.current?.focus(), 50);
  };

  const handleClearEntity = () => {
    setSelectedEntity(null);
    setSearchTerm('');
    setTimeout(() => entityInputRef.current?.focus(), 50);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      if (!showResults) setShowResults(true);
      setActiveIndex(prev => Math.min(prev + 1, filteredEntities.length - 1));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setActiveIndex(prev => Math.max(prev - 1, 0));
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (activeIndex >= 0 && activeIndex < filteredEntities.length) {
        handleSelectEntity(filteredEntities[activeIndex]);
      }
    } else if (e.key === 'Escape') {
      if (showResults) {
        setShowResults(false);
      } else {
        onClose();
      }
    }
  };

  const handleGuidKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' && canNavigate) {
      e.preventDefault();
      handleOpenRecord();
    } else if (e.key === 'Escape') {
      onClose();
    }
  };

  const handleGuidPaste = (e: React.ClipboardEvent<HTMLInputElement>) => {
    e.preventDefault();
    const pasted = e.clipboardData.getData('text');
    const cleaned = cleanGuid(pasted);
    setGuidValue(cleaned);
    setGuidTouched(true);
  };

  const handleBlur = () => {
    setTimeout(() => setShowResults(false), 200);
  };

  const cleanedGuid = cleanGuid(guidValue);
  const guidIsValid = isValidGuid(cleanedGuid);
  const canNavigate = !!(selectedEntity && guidIsValid);

  const handleOpenRecord = () => {
    if (!selectedEntity || !guidIsValid) return;
    const guid = formatGuid(cleanedGuid);
    const url = `${orgUrl}/main.aspx?pagetype=entityrecord&etn=${selectedEntity.LogicalName}&id=${guid}`;
    window.open(url, '_blank');
  };

  const handleOpenWebApi = () => {
    if (!selectedEntity || !guidIsValid) return;
    const guid = formatGuid(cleanedGuid);
    const apiUrl = `${orgUrl}/api/data/v9.2/${selectedEntity.EntitySetName}(${guid})`;
    const viewerUrl = chrome.runtime.getURL(`webapi-viewer.html?url=${encodeURIComponent(apiUrl)}`);
    window.open(viewerUrl, '_blank');
  };

  const loading = entities === null && !error;

  const renderBody = () => {
    if (loading) {
      return (
        <div className="d365-navigate-content">
          <div className="d365-navigate-loading">
            <div className="d365-navigate-spinner"></div>
            <p>Loading entities...</p>
          </div>
        </div>
      );
    }

    if (error) {
      return (
        <div className="d365-navigate-content">
          <div className="d365-dialog-error">
            <div className="d365-dialog-error-icon">⚠</div>
            <div className="d365-dialog-error-message">{error}</div>
            <div className="d365-dialog-error-hint">
              Make sure you have permission to access entity metadata.
            </div>
          </div>
        </div>
      );
    }

    return (
      <div className="d365-navigate-content">
        <div className="d365-navigate-form">
          {/* Entity selector */}
          <div className="d365-navigate-field">
            <label className="d365-navigate-label">Table / Entity</label>

            {selectedEntity ? (
              <div className="d365-navigate-selected">
                <div className="d365-navigate-selected-info">
                  <span className="d365-navigate-selected-name">{selectedEntity.DisplayName}</span>
                  <span className="d365-navigate-selected-logical">{selectedEntity.LogicalName}</span>
                </div>
                <button
                  type="button"
                  className="d365-navigate-clear-btn"
                  onClick={handleClearEntity}
                  title="Clear selection"
                >
                  ✕
                </button>
              </div>
            ) : (
              <div className="d365-navigate-search">
                <input
                  ref={entityInputRef}
                  type="text"
                  className="d365-navigate-input"
                  placeholder={loading ? 'Loading entities...' : 'Search by display name or logical name...'}
                  value={searchTerm}
                  onChange={(e) => {
                    setSearchTerm(e.target.value);
                    setShowResults(true);
                    setActiveIndex(-1);
                  }}
                  onFocus={() => setShowResults(true)}
                  onBlur={handleBlur}
                  onKeyDown={handleKeyDown}
                  disabled={loading}
                />

                {showResults && filteredEntities.length > 0 && (
                  <div className="d365-navigate-results" ref={resultsRef}>
                    {filteredEntities.slice(0, 80).map((entity, index) => (
                      <button
                        key={entity.LogicalName}
                        type="button"
                        className={`d365-navigate-result ${index === activeIndex ? 'd365-navigate-result--active' : ''}`}
                        onClick={() => handleSelectEntity(entity)}
                        onMouseEnter={() => setActiveIndex(index)}
                      >
                        <span className="d365-navigate-result-name">{entity.DisplayName}</span>
                        <span className="d365-navigate-result-logical">{entity.LogicalName}</span>
                      </button>
                    ))}
                    {filteredEntities.length > 80 && (
                      <div className="d365-navigate-more">
                        +{filteredEntities.length - 80} more. Type to narrow results.
                      </div>
                    )}
                  </div>
                )}

                {showResults && searchTerm && filteredEntities.length === 0 && (
                  <div className="d365-navigate-results">
                    <div className="d365-navigate-no-results">No entities found</div>
                  </div>
                )}
              </div>
            )}
          </div>

          {/* GUID input */}
          <div className="d365-navigate-field">
            <label className="d365-navigate-label">Record ID (GUID)</label>
            <input
              ref={guidInputRef}
              type="text"
              className={`d365-navigate-guid-input${guidTouched && guidValue && !guidIsValid ? ' invalid' : ''}`}
              placeholder="e.g. a1b2c3d4-e5f6-7890-abcd-ef1234567890"
              value={guidValue}
              onChange={(e) => {
                setGuidValue(e.target.value);
                setGuidTouched(true);
              }}
              onPaste={handleGuidPaste}
              onKeyDown={handleGuidKeyDown}
            />
            {guidTouched && guidValue && !guidIsValid ? (
              <span className="d365-navigate-guid-error">Invalid GUID format</span>
            ) : (
              <span className="d365-navigate-guid-hint">Paste a GUID — curly braces and spaces are removed automatically</span>
            )}
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="d365-dialog-overlay">
      <div className="d365-dialog-modal d365-navigate-modal">
        <div className="d365-dialog-header d365-navigate-header">
          <h2>Navigate to Record</h2>
          <button className="d365-dialog-close" onClick={onClose} title="Close">
            ×
          </button>
        </div>

        {renderBody()}

        <div className="d365-navigate-footer">
          <button
            className="d365-navigate-btn d365-navigate-btn-cancel"
            onClick={onClose}
          >
            Cancel
          </button>
          <button
            className="d365-navigate-btn d365-navigate-btn-secondary"
            onClick={handleOpenWebApi}
            disabled={!canNavigate}
            title="Open record in Web API Viewer"
          >
            Web API
          </button>
          <button
            className="d365-navigate-btn d365-navigate-btn-primary"
            onClick={handleOpenRecord}
            disabled={!canNavigate}
          >
            Open Record
          </button>
        </div>
      </div>
    </div>
  );
};

export default RecordNavigator;
