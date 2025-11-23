import React, { useState, useEffect, useMemo } from 'react';
import { CrmApi } from '../utils/api';
import { AttributeMetadata, SelectedEntity, JoinedEntity, QueryColumn } from '../types';

interface ColumnSelectorProps {
  api: CrmApi | null;
  entities: (SelectedEntity | JoinedEntity)[];
  selectedColumns: QueryColumn[];
  onToggleColumn: (column: QueryColumn) => void;
}

interface EntityAttributes {
  entityAlias: string;
  loading: boolean;
  attributes: AttributeMetadata[];
}

const ColumnSelector: React.FC<ColumnSelectorProps> = ({ api, entities, selectedColumns, onToggleColumn }) => {
  const [entityAttributes, setEntityAttributes] = useState<Record<string, EntityAttributes>>({});
  const [searchTerm, setSearchTerm] = useState('');
  const [expandedEntities, setExpandedEntities] = useState<Set<string>>(new Set(['main']));

  // Fetch attributes when entities change
  useEffect(() => {
    if (!api) return;

    entities.forEach(entity => {
      if (!entityAttributes[entity.alias]) {
        // Initialize loading state
        setEntityAttributes(prev => ({
          ...prev,
          [entity.alias]: { entityAlias: entity.alias, loading: true, attributes: [] }
        }));

        // Fetch
        api.getAttributes(entity.logicalName).then(attrs => {
          setEntityAttributes(prev => ({
            ...prev,
            [entity.alias]: { entityAlias: entity.alias, loading: false, attributes: attrs }
          }));
          
          // Auto-select Primary ID if not already selected
          const primaryId = attrs.find(a => a.IsPrimaryId);
          if (primaryId) {
             // Check if already selected
             // Actually parent handles selection state, but we can't check it easily here without creating a loop.
             // The parent should probably ensure primary ID is selected when entity is added.
             // But we can do a check here if we want to enforce UI.
             // Let's leave it to the user or parent init.
          }
        });
      }
    });
  }, [api, entities]);

  const toggleExpand = (alias: string) => {
    const newSet = new Set(expandedEntities);
    if (newSet.has(alias)) {
      newSet.delete(alias);
    } else {
      newSet.add(alias);
    }
    setExpandedEntities(newSet);
  };

  const getSelectAttributeName = (attr: AttributeMetadata) => {
    const type = attr.AttributeType?.toLowerCase();
    if (type === 'lookup' || type === 'customer' || type === 'owner') {
      return `_${attr.LogicalName}_value`;
    }
    return attr.LogicalName;
  };

  const handleToggle = (entityAlias: string, attr: AttributeMetadata) => {
    const attributeName = getSelectAttributeName(attr);
    const col: QueryColumn = {
      entityAlias,
      attribute: attributeName,
      displayName: attr.DisplayName,
      logicalName: attr.LogicalName,
      attributeType: attr.AttributeType
    };
    onToggleColumn(col);
  };

  const isSelected = (alias: string, attr: AttributeMetadata) => {
    const selectName = getSelectAttributeName(attr);
    const logicalLower = attr.LogicalName.toLowerCase();
    return selectedColumns.some(c =>
      c.entityAlias === alias &&
      (c.attribute === selectName || (c.logicalName && c.logicalName.toLowerCase() === logicalLower))
    );
  };

  // Filter attributes based on search
  const getFilteredAttributes = (alias: string) => {
    const data = entityAttributes[alias];
    if (!data || !data.attributes) return [];
    
    if (!searchTerm) return data.attributes;
    
    const lower = searchTerm.toLowerCase();
    return data.attributes.filter(a => 
      a.DisplayName.toLowerCase().includes(lower) || 
      a.LogicalName.toLowerCase().includes(lower)
    );
  };

  return (
    <div className="column-selector">
      <div className="qb-row" style={{ marginBottom: '10px' }}>
         <input 
           type="text" 
           className="qb-input" 
           placeholder="Search columns..." 
           value={searchTerm}
           onChange={e => setSearchTerm(e.target.value)}
         />
      </div>

      <div className="columns-container" style={{ 
        maxHeight: '400px', 
        overflowY: 'auto', 
        border: '1px solid #e1dfdd',
        background: 'white',
        borderRadius: '4px'
      }}>
        {entities.map(entity => {
          const data = entityAttributes[entity.alias];
          const attrs = getFilteredAttributes(entity.alias);
          const isExpanded = expandedEntities.has(entity.alias) || searchTerm.length > 0;
          
          if (!data) return null;

          return (
            <div key={entity.alias} className="entity-group">
              <div 
                className="entity-header" 
                style={{ 
                   padding: '8px 12px', 
                   background: '#f3f2f1', 
                   cursor: 'pointer',
                   fontWeight: 600,
                   borderBottom: '1px solid #e1dfdd',
                   display: 'flex',
                   alignItems: 'center',
                   gap: '8px'
                }}
                onClick={() => toggleExpand(entity.alias)}
              >
                 <span>{isExpanded ? '▼' : '▶'}</span>
                 <span>{entity.displayName} ({entity.alias === 'main' ? 'Main' : entity.alias})</span>
                 {data.loading && <span style={{fontSize: '11px', color: '#666'}}>(Loading...)</span>}
              </div>
              
              {isExpanded && (
                <div className="entity-attributes" style={{ padding: '5px 0' }}>
                   {attrs.length === 0 && !data.loading && (
                     <div style={{ padding: '10px', color: '#666' }}>No matching columns</div>
                   )}
                   
                   {attrs.map(attr => {
                     const checked = isSelected(entity.alias, attr);
                     return (
                       <div 
                         key={attr.LogicalName} 
                         className="attribute-row"
                         style={{ 
                           padding: '4px 12px 4px 32px',
                           display: 'flex',
                           alignItems: 'center',
                           gap: '8px'
                         }}
                         onClick={() => handleToggle(entity.alias, attr)}
                       >
                         <input 
                           type="checkbox" 
                           checked={checked} 
                           readOnly 
                           style={{ cursor: 'pointer' }}
                         />
                         <div style={{ flex: 1, cursor: 'pointer' }}>
                           <span style={{ color: '#333' }}>{attr.DisplayName}</span>
                           <span style={{ color: '#666', fontSize: '11px', marginLeft: '6px' }}>{attr.LogicalName}</span>
                           {attr.IsPrimaryId && <span style={{ color: '#0078d4', fontSize: '11px', marginLeft: '6px' }}>(Key)</span>}
                         </div>
                       </div>
                     );
                   })}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
};

export default ColumnSelector;

