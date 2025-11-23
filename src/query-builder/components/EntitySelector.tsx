import React, { useState, useEffect, useMemo, useRef } from 'react';
import { CrmApi } from '../utils/api';
import { EntityMetadata, SelectedEntity } from '../types';

interface EntitySelectorProps {
  api: CrmApi | null;
  onEntitySelected: (entity: EntityMetadata) => void;
  selectedEntity: SelectedEntity | null;
}

const EntitySelector: React.FC<EntitySelectorProps> = ({ api, onEntitySelected, selectedEntity }) => {
  const [entities, setEntities] = useState<EntityMetadata[]>([]);
  const [loading, setLoading] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [showResults, setShowResults] = useState(false);
  const [activeIndex, setActiveIndex] = useState(-1);
  const [error, setError] = useState<string | null>(null);
  
  const inputRef = useRef<HTMLInputElement>(null);
  const resultsRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (api) {
      setLoading(true);
      setError(null);
      api.getAllEntities()
        .then(data => {
          setEntities(data);
          setLoading(false);
        })
        .catch(err => {
          console.error('Failed to load entities', err);
          setError(err.message || 'Failed to load entities');
          setLoading(false);
        });
    }
  }, [api]);

  // Update search term if selected entity changes externally
  useEffect(() => {
    if (selectedEntity) {
      setSearchTerm(selectedEntity.displayName || selectedEntity.logicalName);
    }
  }, [selectedEntity]);

  const filteredEntities = useMemo(() => {
    if (!searchTerm) return entities;
    const lower = searchTerm.toLowerCase();
    return entities.filter(e => 
      e.LogicalName.toLowerCase().includes(lower) || 
      e.DisplayName.toLowerCase().includes(lower)
    );
  }, [entities, searchTerm]);

  const handleSelect = (entity: EntityMetadata) => {
    onEntitySelected(entity);
    setSearchTerm(entity.DisplayName || entity.LogicalName);
    setShowResults(false);
    setActiveIndex(-1);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      if (!showResults) setShowResults(true);
      setActiveIndex(prev => (prev < filteredEntities.length - 1 ? prev + 1 : prev));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setActiveIndex(prev => (prev > 0 ? prev - 1 : prev));
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (activeIndex >= 0 && activeIndex < filteredEntities.length) {
        handleSelect(filteredEntities[activeIndex]);
      }
    } else if (e.key === 'Escape') {
      setShowResults(false);
    }
  };

  const handleBlur = () => {
    // Delayed close to allow click to register
    setTimeout(() => setShowResults(false), 200);
  };

  return (
    <div className="lookup-editor">
      <label style={{ display: 'block', fontSize: '14px', marginBottom: '6px', fontWeight: 500 }}>
        Primary Entity
      </label>
      <div className="lookup-editor__search-row">
        <input
          ref={inputRef}
          type="text"
          className="lookup-editor__input"
          placeholder={loading ? "Loading entities..." : "Search entity by name..."}
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
        {loading && <span className="lookup-editor__loading">Loading...</span>}
      </div>

      {error && (
        <div className="lookup-editor__error">{error}</div>
      )}

      {showResults && !loading && (
        <div className="lookup-editor__results" ref={resultsRef}>
          {filteredEntities.slice(0, 100).map((entity, index) => (
            <button
              key={entity.LogicalName}
              type="button"
              className={`lookup-editor__result ${index === activeIndex ? 'lookup-editor__result--active' : ''}`}
              onClick={() => handleSelect(entity)}
              onMouseEnter={() => setActiveIndex(index)}
            >
              <span className="lookup-editor__result-name">{entity.DisplayName}</span>
              <span className="lookup-editor__result-meta">{entity.LogicalName}</span>
            </button>
          ))}
          {filteredEntities.length === 0 && (
            <div className="lookup-editor__no-results">No entities found</div>
          )}
        </div>
      )}
    </div>
  );
};

export default EntitySelector;
