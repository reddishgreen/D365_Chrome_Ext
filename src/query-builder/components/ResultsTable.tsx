import React, { useMemo, useState, useRef, useCallback, useEffect } from 'react';
import { QueryColumn, SelectedEntity, JoinedEntity } from '../types';

interface ResultsTableProps {
  data: any[];
  columns: QueryColumn[];
  entities: (SelectedEntity | JoinedEntity)[];
  loading: boolean;
  error: string | null;
  orgUrl?: string;
}

type SortDirection = 'asc' | 'desc';
interface SortConfig {
  key: string; // combination of alias and attribute
  direction: SortDirection;
}

const ResultsTable: React.FC<ResultsTableProps> = ({ data, columns, entities, loading, error, orgUrl }) => {
  const [sortConfig, setSortConfig] = useState<SortConfig | null>(null);
  const [columnWidths, setColumnWidths] = useState<Map<number, number>>(new Map());
  const [resizing, setResizing] = useState<{ columnIndex: number; startX: number; startWidth: number } | null>(null);
  const tableRef = useRef<HTMLTableElement>(null);

  const getFormattedValue = (row: any, col: QueryColumn, rawValue: boolean = false) => {
    // 1. Resolve the object holding the value (traverse alias/joins)
    let targetObj = row;
    
    if (col.entityAlias !== 'main') {
       const path: string[] = [];
       let currentAlias = col.entityAlias;
       
       // Reconstruct path from alias back to main
       while (currentAlias !== 'main') {
         const entity = entities.find(e => e.alias === currentAlias) as JoinedEntity;
         if (!entity) break;
         path.unshift(entity.navigationPropertyName);
         currentAlias = entity.parentAlias;
       }
       
       // Traverse
       for (const prop of path) {
         if (targetObj && targetObj[prop]) {
           targetObj = targetObj[prop];
           
           // Handle One-to-Many arrays during traversal
           if (Array.isArray(targetObj)) {
              // For now, we might be unable to dive deep into arrays without mapping
              // If it's an array, we might need to handle it in step 2
           }
         } else {
           targetObj = null;
           break;
         }
       }
    }

    if (!targetObj) return rawValue ? null : '';

    const resolveRawValue = (source: any) => {
      let val = source?.[col.attribute];
      if ((val === undefined || val === null) && col.logicalName) {
        const lookupKey = `_${col.logicalName}_value`;
        if (lookupKey !== col.attribute) {
          val = source?.[lookupKey];
        }
      }
      return val;
    };

    const resolveFormattedValue = (source: any) => {
      const formattedKey = `${col.attribute}@OData.Community.Display.V1.FormattedValue`;
      if (source && source[formattedKey] !== undefined) {
        return source[formattedKey];
      }
      if (col.logicalName) {
        const lookupFormattedKey = `_${col.logicalName}_value@OData.Community.Display.V1.FormattedValue`;
        if (lookupFormattedKey !== formattedKey && source && source[lookupFormattedKey] !== undefined) {
          return source[lookupFormattedKey];
        }
      }
      return undefined;
    };

    // 2. Get value
    // If targetObj is an array (from OneToMany), we map and join
    if (Array.isArray(targetObj)) {
       const values = targetObj.map(item => {
          let val = resolveRawValue(item);

          if (!rawValue) {
             const formatted = resolveFormattedValue(item);
             if (formatted !== undefined) {
               val = formatted;
             }
          }
          return val;
       }).filter(v => v !== null && v !== undefined);
       
       if (rawValue) return values.length > 0 ? values[0] : null;
       return values.join(', ');
    }

    // Single object value extraction
    let val = resolveRawValue(targetObj);
    
    // Try formatted value if not asking for raw value
    if (!rawValue) {
        const formatted = resolveFormattedValue(targetObj);
        if (formatted !== undefined) {
          val = formatted;
        }
    }

    if (rawValue) return val;

    if (val === null || val === undefined) return '';
    if (typeof val === 'object') return JSON.stringify(val);
    return String(val);
  };

  const handleSort = (col: QueryColumn) => {
    const key = `${col.entityAlias}.${col.attribute}`;
    setSortConfig(prev => {
        if (prev && prev.key === key) {
            if (prev.direction === 'asc') return { key, direction: 'desc' };
            return null; // Reset
        }
        return { key, direction: 'asc' };
    });
  };

  const sortedData = useMemo(() => {
     if (!sortConfig || !data) return data;

     const col = columns.find(c => `${c.entityAlias}.${c.attribute}` === sortConfig.key);
     if (!col) return data;

     return [...data].sort((a, b) => {
         const valA = getFormattedValue(a, col, true);
         const valB = getFormattedValue(b, col, true);

         if (valA === valB) return 0;
         if (valA === null || valA === undefined) return 1;
         if (valB === null || valB === undefined) return -1;

         if (valA < valB) return sortConfig.direction === 'asc' ? -1 : 1;
         if (valA > valB) return sortConfig.direction === 'asc' ? 1 : -1;
         return 0;
     });
  }, [data, sortConfig, columns, entities]);

  const getRecordUrl = (row: any) => {
    const mainEntity = entities.find(e => e.alias === 'main');
    if (!mainEntity) return null;
    
    const id = row[mainEntity.primaryIdAttribute];
    if (!id) return null;

    const baseUrl = orgUrl || window.location.origin;
    
    if (!baseUrl || baseUrl.startsWith('chrome-extension://')) return null;
    
    return `${baseUrl}/main.aspx?pagetype=entityrecord&etn=${mainEntity.logicalName}&id=${id}`;
  };

  const getColumnWidth = useCallback((columnIndex: number): number => {
    return columnWidths.get(columnIndex) || 150; // Default width
  }, [columnWidths]);

  const handleResizeStart = useCallback((columnIndex: number, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const startWidth = columnWidths.get(columnIndex) || 150;
    setResizing({ columnIndex, startX: e.clientX, startWidth });
  }, [columnWidths]);

  const handleResize = useCallback((e: MouseEvent) => {
    if (!resizing) return;
    
    const diff = e.clientX - resizing.startX;
    const newWidth = Math.max(50, resizing.startWidth + diff); // Minimum width of 50px
    
    setColumnWidths(prev => {
      const newMap = new Map(prev);
      newMap.set(resizing.columnIndex, newWidth);
      return newMap;
    });
  }, [resizing]);

  const handleResizeEnd = useCallback(() => {
    setResizing(null);
  }, []);

  useEffect(() => {
    if (resizing) {
      document.addEventListener('mousemove', handleResize);
      document.addEventListener('mouseup', handleResizeEnd);
      document.body.classList.add('resizing-column');
      
      return () => {
        document.removeEventListener('mousemove', handleResize);
        document.removeEventListener('mouseup', handleResizeEnd);
        document.body.classList.remove('resizing-column');
      };
    }
  }, [resizing, handleResize, handleResizeEnd]);

  if (loading) {
    return <div className="loading"><div className="spinner"></div>Loading results...</div>;
  }

  if (error) {
    return <div className="error">{error}</div>;
  }

  if (!data || data.length === 0) {
    return <div className="no-results">No results found.</div>;
  }

  return (
    <div className="results-table-container">
      <table className="results-table" ref={tableRef}>
        <thead>
          <tr>
            <th style={{ width: '40px' }}></th>
            {columns.map((col, idx) => {
                const sortKey = `${col.entityAlias}.${col.attribute}`;
                const isSorted = sortConfig?.key === sortKey;
                const width = getColumnWidth(idx);
                
                return (
                  <th 
                    key={idx} 
                    onClick={() => handleSort(col)} 
                    style={{ 
                      cursor: 'pointer', 
                      userSelect: 'none',
                      width: `${width}px`,
                      minWidth: `${width}px`,
                      maxWidth: `${width}px`,
                      position: 'relative'
                    }}
                    title="Click to sort"
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: '4px' }}>
                        {col.displayName}
                        {isSorted && (
                            <span style={{ fontSize: '12px' }}>
                                {sortConfig?.direction === 'asc' ? ' ↑' : ' ↓'}
                            </span>
                        )}
                    </div>
                    <div style={{ fontSize: '10px', color: '#666', fontWeight: 'normal' }}>
                      {col.entityAlias === 'main' ? '' : `${col.entityAlias}.`}{col.logicalName || col.attribute}
                    </div>
                    {idx < columns.length - 1 && (
                      <div
                        className="column-resize-handle"
                        onMouseDown={(e) => {
                          e.stopPropagation(); // Prevent sorting when starting resize
                          handleResizeStart(idx, e);
                        }}
                        onClick={(e) => e.stopPropagation()} // Prevent sorting when clicking handle
                        style={{
                          position: 'absolute',
                          right: '-2px',
                          top: 0,
                          bottom: 0,
                          width: '5px',
                          cursor: 'col-resize',
                          zIndex: 10
                        }}
                        title="Drag to resize column"
                      />
                    )}
                  </th>
                );
            })}
          </tr>
        </thead>
        <tbody>
          {sortedData.map((row, rIdx) => {
            const url = getRecordUrl(row);
            return (
              <tr key={rIdx}>
                <td>
                  {url && (
                    <a 
                      href={url} 
                      target="_blank" 
                      rel="noopener noreferrer"
                      className="link-icon"
                      title="Open Record"
                    >
                      ↗
                    </a>
                  )}
                </td>
                {columns.map((col, cIdx) => {
                  const width = getColumnWidth(cIdx);
                  return (
                    <td 
                      key={cIdx}
                      style={{
                        width: `${width}px`,
                        minWidth: `${width}px`,
                        maxWidth: `${width}px`
                      }}
                    >
                      {getFormattedValue(row, col)}
                    </td>
                  );
                })}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
};

export default ResultsTable;
