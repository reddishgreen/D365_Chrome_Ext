import React, { useState, useEffect } from 'react';
import { CrmApi } from '../utils/api';
import { SelectedEntity, JoinedEntity, QueryFilter, AttributeMetadata } from '../types';
import EnhancedSelect from './EnhancedSelect';

interface FilterBuilderProps {
  api: CrmApi | null;
  entities: (SelectedEntity | JoinedEntity)[];
  filters: QueryFilter;
  onChange: (filters: QueryFilter) => void;
}

const OPERATORS = [
  // Basic Operators
  { value: 'eq', label: 'Equals', types: ['all'] },
  { value: 'ne', label: 'Does Not Equal', types: ['all'] },
  { value: 'null', label: 'Does Not Contain Data', types: ['all'] },
  { value: 'not null', label: 'Contains Data', types: ['all'] },

  // String Operators
  { value: 'contains', label: 'Contains', types: ['String', 'Memo'] },
  { value: 'not contains', label: 'Does Not Contain', types: ['String', 'Memo'] },
  { value: 'startswith', label: 'Begins With', types: ['String', 'Memo'] },
  { value: 'endswith', label: 'Ends With', types: ['String', 'Memo'] },

  // Numeric & Date Comparison Operators
  { value: 'gt', label: 'Greater Than', types: ['Integer', 'Money', 'Decimal', 'DateTime', 'Double'] },
  { value: 'ge', label: 'Greater Than or Equal To', types: ['Integer', 'Money', 'Decimal', 'DateTime', 'Double'] },
  { value: 'lt', label: 'Less Than', types: ['Integer', 'Money', 'Decimal', 'DateTime', 'Double'] },
  { value: 'le', label: 'Less Than or Equal To', types: ['Integer', 'Money', 'Decimal', 'DateTime', 'Double'] },

  // Date-Specific Operators
  { value: 'on', label: 'On', types: ['DateTime'] },
  { value: 'on-or-after', label: 'On or After', types: ['DateTime'] },
  { value: 'on-or-before', label: 'On or Before', types: ['DateTime'] },

  // Relative Date Operators - Daily
  { value: 'today', label: 'Today', types: ['DateTime'] },
  { value: 'yesterday', label: 'Yesterday', types: ['DateTime'] },
  { value: 'tomorrow', label: 'Tomorrow', types: ['DateTime'] },
  { value: 'last-x-days', label: 'Last X Days', types: ['DateTime'], needsValue: true },
  { value: 'next-x-days', label: 'Next X Days', types: ['DateTime'], needsValue: true },
  { value: 'older-than-x-days', label: 'Older Than X Days', types: ['DateTime'], needsValue: true },

  // Relative Date Operators - Weekly
  { value: 'this-week', label: 'This Week', types: ['DateTime'] },
  { value: 'last-week', label: 'Last Week', types: ['DateTime'] },
  { value: 'next-week', label: 'Next Week', types: ['DateTime'] },
  { value: 'last-x-weeks', label: 'Last X Weeks', types: ['DateTime'], needsValue: true },
  { value: 'next-x-weeks', label: 'Next X Weeks', types: ['DateTime'], needsValue: true },

  // Relative Date Operators - Monthly
  { value: 'this-month', label: 'This Month', types: ['DateTime'] },
  { value: 'last-month', label: 'Last Month', types: ['DateTime'] },
  { value: 'next-month', label: 'Next Month', types: ['DateTime'] },
  { value: 'last-x-months', label: 'Last X Months', types: ['DateTime'], needsValue: true },
  { value: 'next-x-months', label: 'Next X Months', types: ['DateTime'], needsValue: true },
  { value: 'older-than-x-months', label: 'Older Than X Months', types: ['DateTime'], needsValue: true },

  // Relative Date Operators - Yearly
  { value: 'this-year', label: 'This Year', types: ['DateTime'] },
  { value: 'last-year', label: 'Last Year', types: ['DateTime'] },
  { value: 'next-year', label: 'Next Year', types: ['DateTime'] },
  { value: 'last-x-years', label: 'Last X Years', types: ['DateTime'], needsValue: true },
  { value: 'next-x-years', label: 'Next X Years', types: ['DateTime'], needsValue: true },
  { value: 'older-than-x-years', label: 'Older Than X Years', types: ['DateTime'], needsValue: true },

  // Fiscal Period Operators
  { value: 'this-fiscal-year', label: 'This Fiscal Year', types: ['DateTime'] },
  { value: 'this-fiscal-period', label: 'This Fiscal Period', types: ['DateTime'] },
  { value: 'last-fiscal-year', label: 'Last Fiscal Year', types: ['DateTime'] },
  { value: 'last-fiscal-period', label: 'Last Fiscal Period', types: ['DateTime'] },
  { value: 'next-fiscal-year', label: 'Next Fiscal Year', types: ['DateTime'] },
  { value: 'next-fiscal-period', label: 'Next Fiscal Period', types: ['DateTime'] },
];

const FilterBuilder: React.FC<FilterBuilderProps> = ({ api, entities, filters, onChange }) => {
  const [attrCache, setAttrCache] = useState<Record<string, AttributeMetadata[]>>({});

  useEffect(() => {
    if (!api) return;
    // Pre-fetch attributes for all entities
    entities.forEach(e => {
      if (!attrCache[e.alias]) {
        api.getAttributes(e.logicalName).then(attrs => {
          setAttrCache(prev => ({ ...prev, [e.alias]: attrs }));
        });
      }
    });
  }, [api, entities]);

  const updateFilter = (id: string, changes: Partial<QueryFilter>) => {
    const updateRecursive = (node: QueryFilter): QueryFilter => {
      if (node.id === id) {
        return { ...node, ...changes };
      }
      if (node.children) {
        return { ...node, children: node.children.map(updateRecursive) };
      }
      return node;
    };
    onChange(updateRecursive(filters));
  };

  const addCondition = (parentId: string) => {
    const updateRecursive = (node: QueryFilter): QueryFilter => {
      if (node.id === parentId && node.children) {
        const newCondition: QueryFilter = {
          id: Math.random().toString(36).substr(2, 9),
          type: 'condition',
          entityAlias: 'main',
          attribute: '',
          operator: 'eq',
          value: ''
        };
        return { ...node, children: [...node.children, newCondition] };
      }
      if (node.children) {
        return { ...node, children: node.children.map(updateRecursive) };
      }
      return node;
    };
    onChange(updateRecursive(filters));
  };

  const addGroup = (parentId: string) => {
    const updateRecursive = (node: QueryFilter): QueryFilter => {
      if (node.id === parentId && node.children) {
        const newGroup: QueryFilter = {
          id: Math.random().toString(36).substr(2, 9),
          type: 'group',
          logicalOperator: 'and',
          children: [
            {
               id: Math.random().toString(36).substr(2, 9),
               type: 'condition',
               entityAlias: 'main',
               attribute: '',
               operator: 'eq',
               value: ''
            }
          ]
        };
        return { ...node, children: [...node.children, newGroup] };
      }
      if (node.children) {
        return { ...node, children: node.children.map(updateRecursive) };
      }
      return node;
    };
    onChange(updateRecursive(filters));
  };

  const removeNode = (id: string) => {
     // Cannot remove root
     if (id === filters.id) return;

     const removeRecursive = (node: QueryFilter): QueryFilter | null => {
       if (node.id === id) return null;
       if (node.children) {
         const newChildren = node.children.map(removeRecursive).filter(c => c !== null) as QueryFilter[];
         return { ...node, children: newChildren };
       }
       return node;
     };
     
     const result = removeRecursive(filters);
     if (result) onChange(result);
  };

  const renderNode = (node: QueryFilter, depth: number = 0) => {
    if (node.type === 'group') {
      const operatorOptions = [
        { value: 'and', label: 'AND' },
        { value: 'or', label: 'OR' }
      ];

      return (
        <div key={node.id} className={`filter-group ${depth > 0 ? 'filter-group--nested' : ''}`}>
           <div className="filter-group-header">
              <EnhancedSelect
                options={operatorOptions}
                value={node.logicalOperator || 'and'}
                onChange={(value) => updateFilter(node.id, { logicalOperator: value as 'and' | 'or' })}
                searchable={false}
                width="90px"
                size="small"
              />
              <button
                className="qb-btn qb-btn-secondary qb-btn--small"
                onClick={() => addCondition(node.id)}
              >
                <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor" className="qb-icon--small">
                  <path d="M6 1v10M1 6h10" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
                </svg>
                Condition
              </button>
              <button
                className="qb-btn qb-btn-secondary qb-btn--small"
                onClick={() => addGroup(node.id)}
              >
                <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor" className="qb-icon--small">
                  <path d="M6 1v10M1 6h10" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
                </svg>
                Group
              </button>
              {depth > 0 && (
                <button
                  className="qb-btn qb-btn-danger qb-btn--small"
                  onClick={() => removeNode(node.id)}
                  title="Delete Group"
                >
                  <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor" className="qb-icon--small">
                    <path d="M2 3h8M4 3V2a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1v1m1 0v7a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V3h6z" stroke="currentColor" strokeWidth="1.2" fill="none" strokeLinecap="round"/>
                  </svg>
                  Delete Group
                </button>
              )}
           </div>
           <div className="group-children">
             {node.children?.map(child => renderNode(child, depth + 1))}
           </div>
        </div>
      );
    } else {
      // Condition
      const entity = entities.find(e => e.alias === node.entityAlias) || entities[0];
      const attrs = attrCache[entity?.alias || ''] || [];
      const selectedAttr = attrs.find(a => a.LogicalName === node.attribute);

      // Filter operators based on type
      const allowedOperators = OPERATORS.filter(op => {
         if (op.types.includes('all')) return true;
         if (!selectedAttr) return false;
         // Map CRM types to our categories
         const type = selectedAttr.AttributeType;
         if (op.types.includes(type)) return true;
         if ((type === 'String' || type === 'Memo') && op.types.includes('String')) return true;
         if ((type === 'Integer' || type === 'Money' || type === 'Decimal' || type === 'Double') && op.types.includes('Integer')) return true;
         if ((type === 'DateTime') && op.types.includes('DateTime')) return true;
         return false;
      });

      const needsValue = ![
        'null', 'not null',
        'today', 'yesterday', 'tomorrow',
        'this-week', 'last-week', 'next-week',
        'this-month', 'last-month', 'next-month',
        'this-year', 'last-year', 'next-year',
        'this-fiscal-year', 'last-fiscal-year', 'next-fiscal-year',
        'this-fiscal-period', 'last-fiscal-period', 'next-fiscal-period'
      ].includes(node.operator || '');

      // Prepare options for selects
      const entityOptions = entities.map(e => ({
        value: e.alias,
        label: `${e.displayName} (${e.alias === 'main' ? 'Main' : e.alias})`
      }));

      const attributeOptions = [
        { value: '', label: 'Select Attribute...' },
        ...attrs.map(a => ({
          value: a.LogicalName,
          label: a.DisplayName,
          description: a.LogicalName
        }))
      ];

      const operatorOptions = allowedOperators.map(op => ({
        value: op.value,
        label: op.label
      }));

      return (
        <div key={node.id} className="filter-row">
           <div className="filter-row__field">
             <EnhancedSelect
               options={entityOptions}
               value={node.entityAlias || 'main'}
               onChange={(value) => updateFilter(node.id, { entityAlias: value, attribute: '', operator: 'eq', value: '' })}
               searchable={false}
               size="medium"
             />
           </div>

           <div className="filter-row__field">
             <EnhancedSelect
               options={attributeOptions}
               value={node.attribute || ''}
               onChange={(value) => updateFilter(node.id, { attribute: value })}
               searchable={true}
               placeholder="Select Attribute..."
               size="medium"
             />
           </div>

           <div className="filter-row__operator">
             <EnhancedSelect
               options={operatorOptions}
               value={node.operator || 'eq'}
               onChange={(value) => updateFilter(node.id, { operator: value })}
               searchable={false}
               size="medium"
             />
           </div>

           {needsValue && (
             <div className="filter-row__value">
               <input
                 type={selectedAttr?.AttributeType === 'DateTime' || node.operator?.includes('x-') ? (node.operator?.includes('x-') ? 'number' : 'date') : 'text'}
                 className="qb-input"
                 value={node.value}
                 placeholder="Value"
                 onChange={(e) => updateFilter(node.id, { value: e.target.value })}
               />
             </div>
           )}

           <div className="filter-row__actions">
             <button
               onClick={() => removeNode(node.id)}
               className="filter-remove-btn"
               title="Remove Condition"
             >
               <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor">
                 <path d="M3 4h8M5 4V3a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1v1m1 0v7a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V4h6z" stroke="currentColor" strokeWidth="1.2" fill="none" strokeLinecap="round"/>
               </svg>
             </button>
           </div>
        </div>
      );
    }
  };

  return (
    <div className="filter-builder">
      {renderNode(filters)}
    </div>
  );
};

export default FilterBuilder;

