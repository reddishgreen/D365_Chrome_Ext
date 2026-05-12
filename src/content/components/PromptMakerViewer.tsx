import React, { useState, useEffect, useMemo, useCallback, useRef } from 'react';
import CloseIcon from './CloseIcon';
import { CrmApi } from '../../query-builder/utils/api';
import { EntityMetadata, EntityMetadataComplete, RelationshipMetadata } from '../../query-builder/types';
import { generatePromptMarkdown, VerbosityLevel, PromptSelections, EntitySelection } from '../utils/promptGenerator';
import { D365Helper } from '../utils/D365Helper';
import './PromptMakerViewer.css';

interface PromptMakerViewerProps {
  onClose: () => void;
}

const PromptMakerViewer: React.FC<PromptMakerViewerProps> = ({ onClose }) => {
  const helper = useRef(new D365Helper()).current;
  const [orgUrl, setOrgUrl] = useState<string | null>(null);
  const [api, setApi] = useState<CrmApi | null>(null);
  const [allEntities, setAllEntities] = useState<EntityMetadata[]>([]);
  const [selectedEntityLogicalNames, setSelectedEntityLogicalNames] = useState<string[]>([]);
  const [entityMetadataMap, setEntityMetadataMap] = useState<Map<string, EntityMetadataComplete>>(new Map());
  const [loadingMetadata, setLoadingMetadata] = useState<Map<string, boolean>>(new Map());
  const [metadataErrors, setMetadataErrors] = useState<Map<string, string>>(new Map());
  const [selections, setSelections] = useState<PromptSelections>({
    entities: new Map(),
    selectedRelationships: []
  });
  const [verbosity, setVerbosity] = useState<VerbosityLevel>('full');
  const [copied, setCopied] = useState(false);
  const [activeEntityTab, setActiveEntityTab] = useState<string | null>(null);
  const [fieldSearchTerms, setFieldSearchTerms] = useState<Map<string, string>>(new Map());
  const [relationshipTypeFilter, setRelationshipTypeFilter] = useState<'all' | 'OneToMany' | 'ManyToOne' | 'ManyToMany'>('all');
  const [autoIncludeRelationships, setAutoIncludeRelationships] = useState(true);
  const [fieldsExpanded, setFieldsExpanded] = useState(true);
  const [relationshipsExpanded, setRelationshipsExpanded] = useState(true);
  const [showSaveDialog, setShowSaveDialog] = useState(false);
  const [promptName, setPromptName] = useState('');
  const [generatedAt, setGeneratedAt] = useState<string>(() => new Date().toISOString());
  const fileInputRef = useRef<HTMLInputElement>(null);
  // Tracks relationships the user has explicitly toggled off so auto-include
  // doesn't re-add them on subsequent state changes.
  const userDeselectedRelsRef = useRef<Set<string>>(new Set());
  // Guards the one-shot default-to-current-entity behavior so it doesn't
  // re-fire after the user removes the entity or loads a prompt.
  const defaultedToCurrentEntityRef = useRef(false);

  // Note: Prompts are now saved/loaded as JSON files to avoid Chrome storage quota limits

  // Initialize org URL and API
  useEffect(() => {
    const init = async () => {
      try {
        const url = helper.getOrgUrl();
        setOrgUrl(url);
        if (url) {
          const crmApi = new CrmApi(url);
          setApi(crmApi);
          // Load entities in background - don't block UI
          crmApi.getAllEntities().then(entities => {
            setAllEntities(entities);
          }).catch(error => {
            console.error('Failed to load entities', error);
          });
        }
      } catch (error) {
        console.error('Failed to initialize', error);
      }
    };
    init();
  }, [helper]);

  // Entity selector component (inline)
  const EntitySelector: React.FC<{ onAdd: (entity: EntityMetadata) => void }> = ({ onAdd }) => {
    const [searchTerm, setSearchTerm] = useState('');
    const [showResults, setShowResults] = useState(false);
    const [activeIndex, setActiveIndex] = useState(-1);

    const filteredEntities = useMemo(() => {
      if (!searchTerm) return allEntities.slice(0, 50);
      const lower = searchTerm.toLowerCase();
      return allEntities.filter(e =>
        e.LogicalName.toLowerCase().includes(lower) ||
        e.DisplayName.toLowerCase().includes(lower)
      ).slice(0, 100);
    }, [allEntities, searchTerm]);

    const handleSelect = (entity: EntityMetadata) => {
      onAdd(entity);
      setSearchTerm('');
      setShowResults(false);
    };

    return (
      <div className="pm-entity-selector">
        <input
          type="text"
          className="pm-input"
          placeholder="Search and select entities..."
          value={searchTerm}
          onChange={(e) => {
            setSearchTerm(e.target.value);
            setShowResults(true);
          }}
          onFocus={() => setShowResults(true)}
          onKeyDown={(e) => {
            if (e.key === 'ArrowDown') {
              e.preventDefault();
              setActiveIndex(prev => Math.min(prev + 1, filteredEntities.length - 1));
            } else if (e.key === 'ArrowUp') {
              e.preventDefault();
              setActiveIndex(prev => Math.max(prev - 1, -1));
            } else if (e.key === 'Enter' && activeIndex >= 0) {
              e.preventDefault();
              handleSelect(filteredEntities[activeIndex]);
            } else if (e.key === 'Escape') {
              setShowResults(false);
            }
          }}
        />
        {showResults && filteredEntities.length > 0 && (
          <div className="pm-dropdown">
            {filteredEntities.map((entity, index) => (
              <button
                key={entity.LogicalName}
                type="button"
                className={`pm-dropdown-item ${index === activeIndex ? 'active' : ''}`}
                onClick={() => handleSelect(entity)}
                onMouseEnter={() => setActiveIndex(index)}
              >
                <span className="pm-entity-name">{entity.DisplayName}</span>
                <span className="pm-entity-logical">{entity.LogicalName}</span>
              </button>
            ))}
          </div>
        )}
      </div>
    );
  };

  // Add entity handler
  const handleAddEntity = useCallback(async (entity: EntityMetadata) => {
    if (selectedEntityLogicalNames.includes(entity.LogicalName)) {
      return; // Already selected
    }

    setSelectedEntityLogicalNames(prev => [...prev, entity.LogicalName]);
    setActiveEntityTab(entity.LogicalName);

    // Fetch metadata if not already cached
    if (!entityMetadataMap.has(entity.LogicalName) && api) {
      setLoadingMetadata(prev => {
        const newMap = new Map(prev);
        newMap.set(entity.LogicalName, true);
        return newMap;
      });
      setMetadataErrors(prev => {
        const newMap = new Map(prev);
        newMap.delete(entity.LogicalName);
        return newMap;
      });
      
      try {
        const metadata = await api.getEntityMetadata(entity.LogicalName);
        
        setEntityMetadataMap(prev => {
          const newMap = new Map(prev);
          newMap.set(entity.LogicalName, metadata);
          return newMap;
        });

        // Initialize selections - select all fields by default
        setSelections(prev => {
          const newEntities = new Map(prev.entities);
          newEntities.set(entity.LogicalName, {
            entityLogicalName: entity.LogicalName,
            selectedFields: metadata.Attributes.map(a => a.LogicalName),
            selectedRelationships: []
          });
          return { ...prev, entities: newEntities };
        });
      } catch (error: any) {
        console.error('Failed to fetch entity metadata', error);
        setMetadataErrors(prev => {
          const newMap = new Map(prev);
          newMap.set(entity.LogicalName, error.message || 'Failed to load metadata');
          return newMap;
        });
      } finally {
        setLoadingMetadata(prev => {
          const newMap = new Map(prev);
          newMap.set(entity.LogicalName, false);
          return newMap;
        });
      }
    } else if (entityMetadataMap.has(entity.LogicalName)) {
      // Entity already loaded, just initialize selections if needed
      setSelections(prev => {
        if (prev.entities.has(entity.LogicalName)) return prev;
        const newEntities = new Map(prev.entities);
        const metadata = entityMetadataMap.get(entity.LogicalName)!;
        newEntities.set(entity.LogicalName, {
          entityLogicalName: entity.LogicalName,
          selectedFields: metadata.Attributes.map(a => a.LogicalName),
          selectedRelationships: []
        });
        return { ...prev, entities: newEntities };
      });
    }
  }, [selectedEntityLogicalNames, entityMetadataMap, api]);

  // Default the first entity to the form's current entity (if any). Runs once
  // after entities load. Skipped if the user already added an entity (e.g.
  // loaded a prompt from file) so we never clobber explicit intent.
  useEffect(() => {
    if (defaultedToCurrentEntityRef.current) return;
    if (allEntities.length === 0) return;
    if (selectedEntityLogicalNames.length > 0) {
      defaultedToCurrentEntityRef.current = true;
      return;
    }
    defaultedToCurrentEntityRef.current = true;
    (async () => {
      try {
        const currentEntity = await helper.getEntityName();
        if (!currentEntity) return;
        const match = allEntities.find(e => e.LogicalName === currentEntity);
        if (match) {
          handleAddEntity(match);
        }
      } catch {
        // Not on a form, or form not ready — leave the picker empty.
      }
    })();
  }, [allEntities, helper, handleAddEntity, selectedEntityLogicalNames]);

  // Remove entity handler
  const handleRemoveEntity = useCallback((entityLogicalName: string) => {
    setSelectedEntityLogicalNames(prev => prev.filter(name => name !== entityLogicalName));
    setSelections(prev => {
      const newEntities = new Map(prev.entities);
      newEntities.delete(entityLogicalName);
      return { ...prev, entities: newEntities };
    });
    if (activeEntityTab === entityLogicalName) {
      const remaining = selectedEntityLogicalNames.filter(name => name !== entityLogicalName);
      setActiveEntityTab(remaining.length > 0 ? remaining[0] : null);
    }
  }, [selectedEntityLogicalNames, activeEntityTab]);

  // Toggle field selection
  const handleToggleField = useCallback((entityLogicalName: string, fieldLogicalName: string) => {
    setSelections(prev => {
      const newEntities = new Map(prev.entities);
      const entitySelection = newEntities.get(entityLogicalName);
      if (!entitySelection) return prev;

      const newSelectedFields = entitySelection.selectedFields.includes(fieldLogicalName)
        ? entitySelection.selectedFields.filter(f => f !== fieldLogicalName)
        : [...entitySelection.selectedFields, fieldLogicalName];

      newEntities.set(entityLogicalName, {
        ...entitySelection,
        selectedFields: newSelectedFields
      });

      return { ...prev, entities: newEntities };
    });
  }, []);

  // Select all filtered fields
  const handleSelectAllFiltered = useCallback((entityLogicalName: string) => {
    const metadata = entityMetadataMap.get(entityLogicalName);
    if (!metadata) return;

    const searchTerm = fieldSearchTerms.get(entityLogicalName) || '';
    const filteredFields = metadata.Attributes.filter(attr => {
      if (!searchTerm) return true;
      const lower = searchTerm.toLowerCase();
      return attr.LogicalName.toLowerCase().includes(lower) ||
        attr.DisplayName.toLowerCase().includes(lower);
    }).map(attr => attr.LogicalName);

    setSelections(prev => {
      const newEntities = new Map(prev.entities);
      const entitySelection = newEntities.get(entityLogicalName);
      if (!entitySelection) return prev;

      newEntities.set(entityLogicalName, {
        ...entitySelection,
        selectedFields: Array.from(new Set([...entitySelection.selectedFields, ...filteredFields]))
      });

      return { ...prev, entities: newEntities };
    });
  }, [entityMetadataMap, fieldSearchTerms]);

  // Clear field selection
  const handleClearFields = useCallback((entityLogicalName: string) => {
    setSelections(prev => {
      const newEntities = new Map(prev.entities);
      const entitySelection = newEntities.get(entityLogicalName);
      if (!entitySelection) return prev;

      newEntities.set(entityLogicalName, {
        ...entitySelection,
        selectedFields: []
      });

      return { ...prev, entities: newEntities };
    });
  }, []);

  // Get available relationships between selected entities.
  // The same SchemaName can appear in multiple buckets (e.g. account.OneToMany
  // and contact.ManyToOne refer to the same 1:N metadata). Pick a canonical
  // form deterministically: OneToMany > ManyToMany > ManyToOne, so the labelled
  // type and direction don't depend on the order entities were added.
  const availableRelationships = useMemo(() => {
    const entitySet = new Set(selectedEntityLogicalNames);
    const rank = (t: RelationshipMetadata['RelationshipType']) =>
      t === 'OneToMany' ? 0 : t === 'ManyToMany' ? 1 : 2;
    const bySchema = new Map<string, RelationshipMetadata>();

    entityMetadataMap.forEach((metadata, entityLogicalName) => {
      if (!entitySet.has(entityLogicalName)) return;

      [...metadata.OneToManyRelationships, ...metadata.ManyToManyRelationships, ...metadata.ManyToOneRelationships]
        .forEach(rel => {
          const involvesSelected = entitySet.has(rel.ReferencingEntity) && entitySet.has(rel.ReferencedEntity);
          if (!involvesSelected) return;
          const existing = bySchema.get(rel.SchemaName);
          if (!existing || rank(rel.RelationshipType) < rank(existing.RelationshipType)) {
            bySchema.set(rel.SchemaName, rel);
          }
        });
    });

    return Array.from(bySchema.values()).filter(rel => {
      if (relationshipTypeFilter === 'all') return true;
      return rel.RelationshipType === relationshipTypeFilter;
    });
  }, [selectedEntityLogicalNames, entityMetadataMap, relationshipTypeFilter]);

  // Auto-include relationships — adds newly available rels but never re-adds
  // ones the user has explicitly toggled off (tracked in userDeselectedRelsRef).
  useEffect(() => {
    if (!autoIncludeRelationships || selectedEntityLogicalNames.length < 2) return;
    setSelections(prev => {
      const next = new Set(prev.selectedRelationships);
      let changed = false;
      availableRelationships.forEach(rel => {
        if (userDeselectedRelsRef.current.has(rel.SchemaName)) return;
        if (!next.has(rel.SchemaName)) {
          next.add(rel.SchemaName);
          changed = true;
        }
      });
      if (!changed) return prev;
      return { ...prev, selectedRelationships: Array.from(next) };
    });
  }, [autoIncludeRelationships, selectedEntityLogicalNames, availableRelationships]);

  // When the user re-enables auto-include, forget their previous deselections.
  useEffect(() => {
    if (autoIncludeRelationships) {
      userDeselectedRelsRef.current = new Set();
    }
  }, [autoIncludeRelationships]);

  // Toggle relationship selection. Tracks user intent so auto-include doesn't
  // override an explicit uncheck.
  const handleToggleRelationship = useCallback((schemaName: string) => {
    setSelections(prev => {
      const isSelected = prev.selectedRelationships.includes(schemaName);
      if (isSelected) {
        userDeselectedRelsRef.current.add(schemaName);
      } else {
        userDeselectedRelsRef.current.delete(schemaName);
      }
      const newSelected = isSelected
        ? prev.selectedRelationships.filter(r => r !== schemaName)
        : [...prev.selectedRelationships, schemaName];
      return { ...prev, selectedRelationships: newSelected };
    });
  }, []);

  // Generate markdown
  const markdown = useMemo(() => {
    if (selections.entities.size === 0) {
      return '# AI Prompt Maker\n\nSelect entities, fields, and relationships to generate context.';
    }

    return generatePromptMarkdown(entityMetadataMap, selections, {
      verbosity,
      includeRules: true,
      orgUrl: orgUrl || undefined,
      apiVersion: 'v9.2', // Match CrmApi default
      generatedAt
    });
  }, [entityMetadataMap, selections, verbosity, orgUrl, generatedAt]);

  // Surface entities the user added that haven't loaded yet so Copy/Save isn't silent.
  const pendingEntityCount = useMemo(() => {
    return selectedEntityLogicalNames.filter(name => !entityMetadataMap.has(name)).length;
  }, [selectedEntityLogicalNames, entityMetadataMap]);

  // Estimate token count (rough: 1 token ≈ 4 characters)
  const estimatedTokens = useMemo(() => {
    return Math.ceil(markdown.length / 4);
  }, [markdown]);

  // Copy to clipboard. Refresh the timestamp so the copied prompt reflects "now".
  const handleCopy = useCallback(async () => {
    const stamp = new Date().toISOString();
    setGeneratedAt(stamp);
    const fresh = selections.entities.size === 0
      ? markdown
      : generatePromptMarkdown(entityMetadataMap, selections, {
          verbosity,
          includeRules: true,
          orgUrl: orgUrl || undefined,
          apiVersion: 'v9.2',
          generatedAt: stamp
        });
    try {
      await navigator.clipboard.writeText(fresh);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (error) {
      console.error('Failed to copy', error);
    }
  }, [markdown, entityMetadataMap, selections, verbosity, orgUrl]);

  // Save prompt to file
  const handleSave = useCallback(() => {
    if (!promptName.trim()) {
      alert('Please enter a name for the prompt');
      return;
    }

    const stamp = new Date().toISOString();
    setGeneratedAt(stamp);
    const promptData = {
      id: `prompt_${Date.now()}`,
      name: promptName.trim(),
      savedAt: stamp,
      orgUrl,
      selections: {
        entities: Array.from(selections.entities.entries()).map(([key, value]) => [key, {
          entityLogicalName: value.entityLogicalName,
          selectedFields: value.selectedFields,
          selectedRelationships: value.selectedRelationships
        }]),
        selectedRelationships: selections.selectedRelationships
      },
      verbosity,
      entityMetadataMap: Array.from(entityMetadataMap.entries()),
      autoIncludeRelationships,
      relationshipTypeFilter
    };
    
    // Create a blob and download it
    const jsonString = JSON.stringify(promptData, null, 2);
    const blob = new Blob([jsonString], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `d365-prompt-${promptName.trim().replace(/[^a-z0-9]/gi, '_')}-${Date.now()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    setShowSaveDialog(false);
    setPromptName('');
    alert(`Prompt "${promptName.trim()}" saved to downloads!`);
  }, [promptName, orgUrl, selections, verbosity, entityMetadataMap, autoIncludeRelationships, relationshipTypeFilter]);

  // Load prompt from file
  const handleLoadFromFile = useCallback((event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const text = e.target?.result as string;
        const prompt = JSON.parse(text);
        
        // Restore entity metadata first
        if (prompt.entityMetadataMap && Array.isArray(prompt.entityMetadataMap)) {
          const restoredMetadata = new Map<string, EntityMetadataComplete>();
          prompt.entityMetadataMap.forEach(([key, value]: [string, EntityMetadataComplete]) => {
            restoredMetadata.set(key, value);
          });
          setEntityMetadataMap(restoredMetadata);
        }

        // Restore selections
        const restoredEntities = new Map<string, EntitySelection>();
        if (prompt.selections && prompt.selections.entities) {
          const entitiesArray = Array.isArray(prompt.selections.entities)
            ? prompt.selections.entities
            : [];

          entitiesArray.forEach(([key, value]: [string, any]) => {
            restoredEntities.set(key, {
              entityLogicalName: value.entityLogicalName as string,
              selectedFields: (value.selectedFields || []) as string[],
              selectedRelationships: (value.selectedRelationships || []) as string[]
            });
          });
        }

        // Restore selected entities
        const entityKeys = Array.from(restoredEntities.keys());
        setSelectedEntityLogicalNames(entityKeys);

        // Set active tab
        if (entityKeys.length > 0) {
          setActiveEntityTab(entityKeys[0]);
        }

        // Restore selections
        const restoredRelSelections = (prompt.selections?.selectedRelationships || []) as string[];
        setSelections({
          entities: restoredEntities,
          selectedRelationships: restoredRelSelections
        });

        // Restore verbosity
        setVerbosity((prompt.verbosity || 'full') as VerbosityLevel);

        // Restore other settings
        if (prompt.autoIncludeRelationships !== undefined) {
          setAutoIncludeRelationships(prompt.autoIncludeRelationships);
        }
        if (prompt.relationshipTypeFilter) {
          setRelationshipTypeFilter(prompt.relationshipTypeFilter);
        }

        // Reset deselection memory so auto-include matches the loaded selection.
        userDeselectedRelsRef.current = new Set();

        // Refresh the generated timestamp for the loaded prompt
        setGeneratedAt(new Date().toISOString());

        // Clear loading/error states
        const loadingMap = new Map<string, boolean>();
        const errorMap = new Map<string, string>();
        entityKeys.forEach(key => {
          loadingMap.set(key, false);
          errorMap.delete(key);
        });
        setLoadingMetadata(loadingMap);
        setMetadataErrors(errorMap);

        // Reset file input
        if (fileInputRef.current) {
          fileInputRef.current.value = '';
        }

        alert(`Prompt "${prompt.name || 'Untitled'}" loaded successfully! ${entityKeys.length} entities with their field selections restored.`);
      } catch (error) {
        console.error('[Prompt Maker] Error loading prompt from file:', error);
        alert('Error loading prompt file: ' + (error instanceof Error ? error.message : 'Invalid file format'));
      }
    };
    reader.onerror = () => {
      alert('Error reading file');
    };
    reader.readAsText(file);
  }, []);

  // Clear current prompt
  const handleClearPrompt = useCallback(() => {
    if (confirm('Are you sure you want to clear all selections? This cannot be undone.')) {
      setSelectedEntityLogicalNames([]);
      setSelections({
        entities: new Map(),
        selectedRelationships: []
      });
      setEntityMetadataMap(new Map());
      setActiveEntityTab(null);
      setFieldSearchTerms(new Map());
      setVerbosity('full');
      setAutoIncludeRelationships(true);
      setRelationshipTypeFilter('all');
      userDeselectedRelsRef.current = new Set();
      setGeneratedAt(new Date().toISOString());
    }
  }, []);

  // Open save dialog
  const handleOpenSave = useCallback(() => {
    setPromptName('');
    setShowSaveDialog(true);
  }, []);

  // Get filtered fields for an entity
  const getFilteredFields = useCallback((entityLogicalName: string) => {
    const metadata = entityMetadataMap.get(entityLogicalName);
    if (!metadata) {
      return [];
    }

    if (!metadata.Attributes) {
      console.warn('[Prompt Maker] Metadata has no Attributes property:', metadata);
      return [];
    }

    const searchTerm = fieldSearchTerms.get(entityLogicalName) || '';
    if (!searchTerm) {
      return metadata.Attributes;
    }

    const lower = searchTerm.toLowerCase();
    return metadata.Attributes.filter(attr =>
      attr.LogicalName.toLowerCase().includes(lower) ||
      attr.DisplayName.toLowerCase().includes(lower)
    );
  }, [entityMetadataMap, fieldSearchTerms]);

  return (
    <div className="d365-dialog-overlay" onClick={(e) => {
      if (e.target === e.currentTarget) onClose();
    }}>
      <div className="d365-dialog-modal pm-modal" onClick={(e) => e.stopPropagation()}>
        <div className="d365-dialog-header pm-header">
          <h2>AI Prompt Maker</h2>
          <button className="d365-dialog-close" onClick={onClose} title="Close" aria-label="Close"><CloseIcon /></button>
        </div>

        <div className="pm-content">
          {/* Left Panel - Selection */}
          <div className="pm-selection-panel">
            {/* Entity Selection */}
            <div className="pm-section">
              <h3>Entities</h3>
              <EntitySelector onAdd={handleAddEntity} />
              {selectedEntityLogicalNames.length > 0 && (
                <div className="pm-selected-entities">
                  {selectedEntityLogicalNames.map(logicalName => {
                    const entity = allEntities.find(e => e.LogicalName === logicalName);
                    return (
                      <div key={logicalName} className="pm-entity-tag">
                        <span>{entity?.DisplayName || logicalName}</span>
                        <button
                          type="button"
                          className="pm-remove-btn"
                          onClick={() => handleRemoveEntity(logicalName)}
                          title="Remove entity"
                          aria-label="Remove entity"
                        >
                          <CloseIcon size={12} />
                        </button>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            {/* Field Selection */}
            {selectedEntityLogicalNames.length > 0 && (
              <div className="pm-section">
                <h3 className="pm-section-header" onClick={() => setFieldsExpanded(!fieldsExpanded)}>
                  <span className="pm-chevron">{fieldsExpanded ? '▼' : '▶'}</span>
                  Fields
                </h3>
                {fieldsExpanded && (
                  <>
                <div className="pm-entity-tabs">
                  {selectedEntityLogicalNames.map(logicalName => {
                    const entity = allEntities.find(e => e.LogicalName === logicalName);
                    return (
                      <button
                        key={logicalName}
                        type="button"
                        className={`pm-tab ${activeEntityTab === logicalName ? 'active' : ''}`}
                        onClick={() => setActiveEntityTab(logicalName)}
                      >
                        {entity?.DisplayName || logicalName}
                      </button>
                    );
                  })}
                </div>

                {activeEntityTab && (
                  <div className="pm-field-selection">
                    <div className="pm-field-search">
                      <input
                        type="text"
                        className="pm-input"
                        placeholder="Search fields..."
                        value={fieldSearchTerms.get(activeEntityTab) || ''}
                        onChange={(e) => {
                          const newMap = new Map(fieldSearchTerms);
                          newMap.set(activeEntityTab, e.target.value);
                          setFieldSearchTerms(newMap);
                        }}
                      />
                      <div className="pm-field-actions">
                        <button
                          type="button"
                          className="pm-btn-small"
                          onClick={() => handleSelectAllFiltered(activeEntityTab)}
                        >
                          Select All Filtered
                        </button>
                        <button
                          type="button"
                          className="pm-btn-small"
                          onClick={() => handleClearFields(activeEntityTab)}
                        >
                          Clear
                        </button>
                      </div>
                    </div>

                    <div className="pm-field-list">
                      {loadingMetadata.get(activeEntityTab) ? (
                        <div className="pm-loading">Loading fields...</div>
                      ) : metadataErrors.has(activeEntityTab) ? (
                        <div className="pm-error">
                          Error: {metadataErrors.get(activeEntityTab)}
                        </div>
                      ) : getFilteredFields(activeEntityTab).length === 0 ? (
                        <div className="pm-empty">No fields found</div>
                      ) : (
                        getFilteredFields(activeEntityTab).map(attr => {
                          const entitySelection = selections.entities.get(activeEntityTab);
                          const isSelected = entitySelection?.selectedFields.includes(attr.LogicalName) || false;
                          return (
                            <label key={attr.LogicalName} className="pm-field-item">
                              <input
                                type="checkbox"
                                checked={isSelected}
                                onChange={() => handleToggleField(activeEntityTab, attr.LogicalName)}
                              />
                              <span className="pm-field-name">{attr.DisplayName}</span>
                              <span className="pm-field-logical">{attr.LogicalName}</span>
                              <span className="pm-field-type">{attr.AttributeType}</span>
                            </label>
                          );
                        })
                      )}
                    </div>
                  </div>
                )}
                  </>
                )}
              </div>
            )}

            {/* Relationship Selection */}
            {selectedEntityLogicalNames.length >= 2 && (
              <div className="pm-section">
                <h3 className="pm-section-header" onClick={() => setRelationshipsExpanded(!relationshipsExpanded)}>
                  <span className="pm-chevron">{relationshipsExpanded ? '▼' : '▶'}</span>
                  Relationships
                </h3>
                {relationshipsExpanded && (
                  <>
                <div className="pm-relationship-controls">
                  <label className="pm-checkbox-label">
                    <input
                      type="checkbox"
                      checked={autoIncludeRelationships}
                      onChange={(e) => setAutoIncludeRelationships(e.target.checked)}
                    />
                    Auto-include relationships between selected entities
                  </label>
                  <select
                    className="pm-select"
                    value={relationshipTypeFilter}
                    onChange={(e) => setRelationshipTypeFilter(e.target.value as any)}
                  >
                    <option value="all">All Types</option>
                    <option value="OneToMany">OneToMany</option>
                    <option value="ManyToOne">ManyToOne</option>
                    <option value="ManyToMany">ManyToMany</option>
                  </select>
                </div>

                <div className="pm-relationship-list">
                  {availableRelationships.map(rel => {
                    // ReferencingEntity = child (FK holder); ReferencedEntity = parent.
                    const childEntity = allEntities.find(e => e.LogicalName === rel.ReferencingEntity);
                    const parentEntity = allEntities.find(e => e.LogicalName === rel.ReferencedEntity);
                    const childName = childEntity?.DisplayName || rel.ReferencingEntity;
                    const parentName = parentEntity?.DisplayName || rel.ReferencedEntity;
                    const isSelected = selections.selectedRelationships.includes(rel.SchemaName);
                    const isSelfReferencing = rel.ReferencingEntity === rel.ReferencedEntity;

                    let pathLabel: string;
                    if (isSelfReferencing) {
                      pathLabel = `${parentName} (self-referencing)`;
                    } else if (rel.RelationshipType === 'OneToMany') {
                      pathLabel = `${parentName} → ${childName}`;
                    } else if (rel.RelationshipType === 'ManyToOne') {
                      pathLabel = `${childName} → ${parentName}`;
                    } else {
                      pathLabel = `${parentName} ↔ ${childName}`;
                    }

                    return (
                      <label key={rel.SchemaName} className="pm-relationship-item">
                        <input
                          type="checkbox"
                          checked={isSelected}
                          onChange={() => handleToggleRelationship(rel.SchemaName)}
                        />
                        <div className="pm-relationship-info">
                          <span className="pm-relationship-name">{rel.SchemaName}</span>
                          <span className="pm-relationship-type">{rel.RelationshipType}</span>
                          <span className="pm-relationship-path">{pathLabel}</span>
                        </div>
                      </label>
                    );
                  })}
                </div>
                  </>
                )}
              </div>
            )}
          </div>

          {/* Right Panel - Preview */}
          <div className="pm-preview-panel">
            <div className="pm-preview-header">
              <div className="pm-preview-controls">
                <label>
                  Verbosity:
                  <select
                    className="pm-select"
                    value={verbosity}
                    onChange={(e) => setVerbosity(e.target.value as VerbosityLevel)}
                  >
                    <option value="compact">Compact</option>
                    <option value="standard">Standard</option>
                    <option value="full">Full</option>
                  </select>
                </label>
                <button
                  type="button"
                  className="pm-btn-secondary"
                  onClick={handleClearPrompt}
                  title="Clear all selections"
                >
                  Clear
                </button>
                <button
                  type="button"
                  className="pm-btn-secondary"
                  onClick={() => fileInputRef.current?.click()}
                  title="Load prompt from file"
                >
                  Load
                </button>
                <input
                  ref={fileInputRef}
                  type="file"
                  accept=".json"
                  style={{ display: 'none' }}
                  onChange={handleLoadFromFile}
                />
                <button
                  type="button"
                  className="pm-btn-primary"
                  onClick={() => handleOpenSave()}
                  title="Save current prompt"
                >
                  Save
                </button>
                <button
                  type="button"
                  className={`pm-copy-btn ${copied ? 'copied' : ''}`}
                  onClick={handleCopy}
                >
                  {copied ? '✓ Copied!' : 'Copy'}
                </button>
              </div>
              <div className="pm-token-count">
                Estimated tokens: ~{estimatedTokens.toLocaleString()}
                {estimatedTokens > 100000 && (
                  <span className="pm-warning"> (Large prompt - may exceed AI context limits)</span>
                )}
                {pendingEntityCount > 0 && (
                  <span className="pm-warning"> ({pendingEntityCount} entit{pendingEntityCount === 1 ? 'y' : 'ies'} still loading)</span>
                )}
              </div>
            </div>
            <div className="pm-markdown-preview">
              <pre><code>{markdown}</code></pre>
            </div>
          </div>
        </div>
      </div>

      {/* Save Dialog */}
      {showSaveDialog && (
        <div className="pm-dialog-overlay" onClick={(e) => {
        if (e.target === e.currentTarget) {
          setShowSaveDialog(false);
          setPromptName('');
        }
        }}>
          <div className="pm-dialog" onClick={(e) => e.stopPropagation()}>
            <h3>Save Prompt to File</h3>
            <div className="pm-dialog-content">
              <label>
                Prompt Name:
                <input
                  type="text"
                  className="pm-input"
                  value={promptName}
                  onChange={(e) => setPromptName(e.target.value)}
                  placeholder="Enter prompt name..."
                  autoFocus
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      handleSave();
                    } else if (e.key === 'Escape') {
                      setShowSaveDialog(false);
                      setPromptName('');
                    }
                  }}
                />
              </label>
            </div>
            <div className="pm-dialog-actions">
              <button
                type="button"
                className="pm-btn-secondary"
                onClick={() => {
                  setShowSaveDialog(false);
                  setPromptName('');
                }}
              >
                Cancel
              </button>
              <button
                type="button"
                className="pm-btn-primary"
                onClick={handleSave}
              >
                Save to Downloads
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
};

export default PromptMakerViewer;
