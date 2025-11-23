import React, { useState, useEffect, useMemo } from 'react';
import { CrmApi } from '../utils/api';
import { SelectedEntity, JoinedEntity, RelationshipMetadata } from '../types';
import EnhancedSelect from './EnhancedSelect';

interface RelationshipManagerProps {
  api: CrmApi | null;
  entities: (SelectedEntity | JoinedEntity)[];
  onAddJoin: (entity: JoinedEntity) => void;
  onRemoveJoin: (alias: string) => void;
  isExpanded?: boolean;
  onToggleExpanded?: (expanded: boolean) => void;
}

const RelationshipManager: React.FC<RelationshipManagerProps> = ({ api, entities, onAddJoin, onRemoveJoin, isExpanded = false, onToggleExpanded }) => {
  const [isAdding, setIsAdding] = useState(false);
  const [fromAlias, setFromAlias] = useState<string>('');
  const [relationships, setRelationships] = useState<RelationshipMetadata[]>([]);
  const [selectedRelationship, setSelectedRelationship] = useState<RelationshipMetadata | null>(null);
  const [selectedRelationshipId, setSelectedRelationshipId] = useState<string>('');
  const [loadingRel, setLoadingRel] = useState(false);

  const fromEntity = useMemo(() => 
    entities.find(e => e.alias === fromAlias), 
    [entities, fromAlias]
  );

  useEffect(() => {
    if (entities.length > 0 && !fromAlias) {
      setFromAlias(entities[0].alias);
    }
  }, [entities]);

  useEffect(() => {
    if (api && fromEntity) {
      setLoadingRel(true);
      setRelationships([]);
      setSelectedRelationship(null);
      setSelectedRelationshipId('');

      // Fetch both ManyToOne and OneToMany
      Promise.all([
        api.getRelationships(fromEntity.logicalName, 'ManyToOne'),
        api.getRelationships(fromEntity.logicalName, 'OneToMany')
      ]).then(([mto, otm]) => {
        setRelationships([...mto, ...otm]);
        setLoadingRel(false);
      }).catch(err => {
        console.error(err);
        setLoadingRel(false);
      });
    }
  }, [api, fromEntity]);

  // Prepare options for relationship selector
  const relationshipOptions = useMemo(() => {
    const opts: Array<{ value: string; label: string; description?: string }> = [
      { value: '', label: 'Select a relationship...' }
    ];

    relationships.forEach(r => {
      const isMto = r.RelationshipType === 'ManyToOne';
      const target = isMto ? r.ReferencedEntity : r.ReferencingEntity;
      const label = `${r.SchemaName} → ${target} (${isMto ? 'Lookup' : 'OneToMany'})`;

      opts.push({
        value: r.SchemaName,
        label: label,
        description: `${isMto ? 'Many-to-One' : 'One-to-Many'} relationship`
      });
    });

    return opts;
  }, [relationships]);

  // Prepare options for entity selector
  const entityOptions = useMemo(() => {
    return entities.map(e => ({
      value: e.alias,
      label: `${e.displayName} (${e.alias === 'main' ? 'Main' : e.alias})`
    }));
  }, [entities]);

  const handleRelationshipChange = (schemaName: string) => {
    setSelectedRelationshipId(schemaName);
    const relationship = relationships.find(r => r.SchemaName === schemaName);
    setSelectedRelationship(relationship || null);
  };

  const handleAdd = async () => {
    if (!api || !selectedRelationship || !fromEntity) return;

    try {
      // Determine target entity logical name
      const isManyToOne = selectedRelationship.RelationshipType === 'ManyToOne';
      
      // If ManyToOne (Lookup), we are referencing another entity. Target is ReferencedEntity.
      // If OneToMany (Children), other entity is referencing us. Target is ReferencingEntity.
      const targetLogicalName = isManyToOne 
        ? selectedRelationship.ReferencedEntity 
        : selectedRelationship.ReferencingEntity;

      // Determine navigation property name (to use in $expand)
      // If ManyToOne, use ReferencingEntityNavigationPropertyName (on Us)
      // If OneToMany, use ReferencedEntityNavigationPropertyName (on Us - wait, actually for OData expand, we need the nav prop on the source entity)
      // Wait, let's verify standard metadata.
      // Account -> PrimaryContact (ManyToOne). Lookup is on Account. Nav prop "primarycontactid" on Account.
      // Metadata: ReferencingEntity: Account, ReferencedEntity: Contact. 
      // ReferencingEntityNavigationPropertyName: "primarycontactid" (On Account).
      
      // Account -> Contacts (OneToMany). Contact has lookup to Account.
      // Metadata: ReferencingEntity: Contact, ReferencedEntity: Account.
      // ReferencedEntityNavigationPropertyName: "contact_customer_accounts" (On Account).
      
      const navProp = isManyToOne
        ? selectedRelationship.ReferencingEntityNavigationPropertyName
        : selectedRelationship.ReferencedEntityNavigationPropertyName;

      if (!navProp) {
        alert('Navigation property not found for this relationship. OData expansion requires it.');
        return;
      }

      const targetMetadata = await api.getEntity(targetLogicalName);
      
      const alias = `${targetLogicalName}_${Math.floor(Math.random() * 1000)}`;
      
      const joined: JoinedEntity = {
        logicalName: targetMetadata.LogicalName,
        entitySetName: targetMetadata.EntitySetName,
        displayName: targetMetadata.DisplayName,
        primaryIdAttribute: targetMetadata.PrimaryIdAttribute,
        primaryNameAttribute: targetMetadata.PrimaryNameAttribute,
        alias: alias,
        relationshipName: selectedRelationship.SchemaName,
        parentAlias: fromEntity.alias,
        relationshipType: selectedRelationship.RelationshipType,
        navigationPropertyName: navProp
      };

      onAddJoin(joined);
      setIsAdding(false);
      setSelectedRelationshipId('');
      setSelectedRelationship(null);
    } catch (error) {
      console.error('Error adding joined entity', error);
      alert('Failed to add related entity');
    }
  };

  // Joined entities (exclude main)
  const joins = entities.filter(e => 'parentAlias' in e) as JoinedEntity[];

  // Auto-expand if there are no joins and isExpanded is true
  useEffect(() => {
    if (isExpanded && joins.length === 0 && !isAdding) {
      setIsAdding(true);
    }
  }, [isExpanded, joins.length]);

  return (
    <div className="relationship-manager">
      <div className="qb-row" style={{ alignItems: 'center', marginBottom: '10px' }}>
        <span style={{ fontWeight: 600, fontSize: '14px' }}>Related Entities ({joins.length})</span>
        {onToggleExpanded && (
          <button
            className="qb-btn qb-btn-secondary qb-btn--small"
            style={{ marginLeft: '10px' }}
            onClick={() => {
              setIsAdding(!isAdding);
              onToggleExpanded(!isAdding);
            }}
          >
            {isAdding ? (
              <>
                <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor" className="qb-icon--small">
                  <path d="M2 2l8 8M10 2l-8 8" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
                </svg>
                Cancel
              </>
            ) : (
              <>
                <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor" className="qb-icon--small">
                  <path d="M6 1v10M1 6h10" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
                </svg>
                Add Related
              </>
            )}
          </button>
        )}
      </div>

      {isAdding && (
        <div className="add-join-panel" style={{ background: '#f9f9f9', padding: '15px', border: '1px solid #eee', borderRadius: '4px', marginBottom: '15px' }}>
          <div className="qb-row">
            <div className="qb-col">
              <label style={{ display: 'block', fontSize: '14px', marginBottom: '6px', fontWeight: 500 }}>From Entity</label>
              <EnhancedSelect
                options={entityOptions}
                value={fromAlias}
                onChange={setFromAlias}
                searchable={false}
                size="medium"
              />
            </div>
            
            <div className="qb-col" style={{ flex: 2 }}>
              <label style={{ display: 'block', fontSize: '14px', marginBottom: '6px', fontWeight: 500 }}>Relationship</label>
              {loadingRel ? (
                <div style={{ fontSize: '12px', padding: '8px' }}>Loading relationships...</div>
              ) : (
                <EnhancedSelect
                  options={relationshipOptions}
                  value={selectedRelationshipId}
                  onChange={handleRelationshipChange}
                  searchable={true}
                  placeholder={relationships.length === 0 ? 'No relationships found' : 'Search and select a relationship...'}
                  size="medium"
                  disabled={relationships.length === 0}
                />
              )}
            </div>
          </div>
          <div style={{ marginTop: '10px', textAlign: 'right' }}>
            <button
              className="qb-btn qb-btn-primary qb-btn--medium"
              disabled={!selectedRelationship}
              onClick={handleAdd}
            >
              <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor" className="qb-icon">
                <path d="M7 2v10M2 7h10" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
              </svg>
              Add Relationship
            </button>
          </div>
        </div>
      )}

      {joins.length > 0 && (
        <div className="joins-list" style={{ display: 'flex', flexWrap: 'wrap', gap: '8px' }}>
          {joins.map(join => (
            <div key={join.alias} className="join-chip">
              <span className="join-chip__name">{join.displayName}</span>
              <span className="join-chip__relationship">via {join.relationshipName}</span>
              <button
                onClick={() => onRemoveJoin(join.alias)}
                className="join-chip__remove"
                title="Remove relationship"
              >
                <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor">
                  <path d="M2 2l8 8M10 2l-8 8" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
                </svg>
              </button>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default RelationshipManager;

