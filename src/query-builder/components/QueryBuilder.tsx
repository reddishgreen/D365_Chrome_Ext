import React, { useState, useEffect, useMemo } from 'react';
import './QueryBuilder.css';
import { CrmApi } from '../utils/api';
import { EntityMetadata, SelectedEntity, JoinedEntity, QueryColumn, QueryFilter, ViewMetadata, AttributeMetadata, RelationshipMetadata } from '../types';
import EntitySelector from './EntitySelector';
import RelationshipManager from './RelationshipManager';
import ColumnSelector from './ColumnSelector';
import FilterBuilder from './FilterBuilder';
import ResultsTable from './ResultsTable';
import ViewSelector from './ViewSelector';
import { buildODataQuery } from '../utils/odata';
import { parseViewFetchXml } from '../utils/viewParser';
import { exportToCSV, exportToExcel } from '../utils/export';

interface QueryBuilderProps {
  orgUrl?: string;
  onClose?: () => void;
}

const QueryBuilder: React.FC<QueryBuilderProps> = ({ orgUrl: propOrgUrl, onClose }) => {
  const [api, setApi] = useState<CrmApi | null>(null);
  const [selectedEntity, setSelectedEntity] = useState<SelectedEntity | null>(null);
  const [joinedEntities, setJoinedEntities] = useState<JoinedEntity[]>([]);
  const [selectedColumns, setSelectedColumns] = useState<QueryColumn[]>([]);
  const [filters, setFilters] = useState<QueryFilter>({
    id: 'root',
    type: 'group',
    logicalOperator: 'and',
    children: []
  });

  const [results, setResults] = useState<any[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [executedQuery, setExecutedQuery] = useState('');

  // Tab state
  const [activeTab, setActiveTab] = useState<'builder' | 'results'>('builder');
  const [showRelationshipManager, setShowRelationshipManager] = useState(false);

  // Export state
  const [showExportModal, setShowExportModal] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [exportProgress, setExportProgress] = useState<{ fetched: number; complete: boolean } | null>(null);
  const [exportError, setExportError] = useState<string | null>(null);

  useEffect(() => {
    if (propOrgUrl) {
      setApi(new CrmApi(propOrgUrl));
    } else {
      // Fallback to URL params if no prop provided
      const params = new URLSearchParams(window.location.search);
      const url = params.get('orgUrl');
      if (url) {
        setApi(new CrmApi(url));
      }
    }
  }, [propOrgUrl]);

  const handleEntitySelected = (meta: EntityMetadata) => {
    const main: SelectedEntity = {
      logicalName: meta.LogicalName,
      entitySetName: meta.EntitySetName,
      displayName: meta.DisplayName,
      primaryIdAttribute: meta.PrimaryIdAttribute,
      primaryNameAttribute: meta.PrimaryNameAttribute,
      alias: 'main'
    };
    setSelectedEntity(main);
    setJoinedEntities([]);
    setSelectedColumns([]);
    setFilters({
      id: 'root',
      type: 'group',
      logicalOperator: 'and',
      children: []
    });
    setResults([]);
    setExecutedQuery('');
  };

  const handleViewSelected = async (view: ViewMetadata) => {
    if (!api || !selectedEntity) return;

    try {
      const result = await parseViewFetchXml(view.fetchXml, api, {
        logicalName: selectedEntity.logicalName,
        primaryIdAttribute: selectedEntity.primaryIdAttribute,
        primaryNameAttribute: selectedEntity.primaryNameAttribute
      });

      setJoinedEntities(result.joinedEntities);
      setSelectedColumns(result.columns);
      setFilters(result.filters);

      setResults([]);
      setExecutedQuery('');
      setError(null);

    } catch (e) {
      console.error('Error parsing view', e);
      setError(`Error parsing view: ${e instanceof Error ? e.message : 'Unknown error'}`);
    }
  };

  const handleAddJoin = (joined: JoinedEntity) => {
    setJoinedEntities(prev => [...prev, joined]);
  };

  const handleRemoveJoin = (alias: string) => {
    setJoinedEntities(prev => prev.filter(j => j.alias !== alias));
    setSelectedColumns(prev => prev.filter(c => c.entityAlias !== alias));
  };

  const handleToggleColumn = (col: QueryColumn) => {
    setSelectedColumns(prev => {
      const exists = prev.some(c => c.entityAlias === col.entityAlias && c.attribute === col.attribute);
      if (exists) {
        return prev.filter(c => !(c.entityAlias === col.entityAlias && c.attribute === col.attribute));
      } else {
        return [...prev, col];
      }
    });
  };

  const allEntities = useMemo(() => {
    if (!selectedEntity) return [];
    return [selectedEntity, ...joinedEntities];
  }, [selectedEntity, joinedEntities]);

  const handleRunQuery = async () => {
    if (!api || !selectedEntity) return;

    setLoading(true);
    setError(null);
    setResults([]);

    // Switch to results tab immediately
    setActiveTab('results');

    try {
      const query = buildODataQuery(selectedEntity, joinedEntities, selectedColumns, filters);
      setExecutedQuery(query);

      // Debug logging
      console.log('Executing OData Query:', query);
      console.log('Selected Columns:', selectedColumns);
      console.log('Main Entity:', selectedEntity);

      const data = await api.executeQuery(query);
      setResults(data.value || []);
    } catch (err) {
      console.error('Query execution error:', err);
      console.error('Query was:', executedQuery);
      setError(err instanceof Error ? err.message : 'An error occurred');
    } finally {
      setLoading(false);
    }
  };

  const handleClose = () => {
    if (onClose) {
      onClose();
    } else {
      window.close();
    }
  };

  const handleExport = async (format: 'csv' | 'excel', exportAll: boolean) => {
    if (!api || !selectedEntity) return;

    setExporting(true);
    setExportError(null);
    setExportProgress(null);

    try {
      let dataToExport: any[];

      if (exportAll) {
        // Fetch ALL records (handles pagination)
        const query = buildODataQuery(selectedEntity, joinedEntities, selectedColumns, filters);
        dataToExport = await api.fetchAllPages(query, (fetched, complete) => {
          setExportProgress({ fetched, complete });
        });
      } else {
        // Export current results only
        dataToExport = results;
      }

      // Generate filename with entity name and timestamp
      const timestamp = new Date().toISOString().slice(0, 10);
      const filename = `${selectedEntity.displayName}_${timestamp}`;

      if (format === 'csv') {
        exportToCSV(dataToExport, selectedColumns, allEntities, filename);
      } else {
        exportToExcel(dataToExport, selectedColumns, allEntities, filename, propOrgUrl);
      }

      setShowExportModal(false);
    } catch (err) {
      console.error('Export error:', err);
      setExportError(err instanceof Error ? err.message : 'An error occurred during export');
    } finally {
      setExporting(false);
      setExportProgress(null);
    }
  };

  return (
    <div className="query-builder-overlay">
      <div className="query-builder-container">
        <header className="qb-header">
          <div className="qb-header-left">
            <img
              className="qb-logo"
              src={chrome.runtime.getURL('icons/RG%20Logo_White_Stacked.svg')}
              alt="RG Logo"
              onError={(e) => e.currentTarget.style.display = 'none'}
            />
            <h1>D365 Advanced Query Builder</h1>
          </div>
          <button className="qb-close-btn" onClick={handleClose} title="Close">×</button>
        </header>

        {/* Tab Navigation */}
        <div className="qb-tabs">
          <button
            className={`qb-tab ${activeTab === 'builder' ? 'active' : ''}`}
            onClick={() => setActiveTab('builder')}
          >
            Query Builder
          </button>
          <button
            className={`qb-tab ${activeTab === 'results' ? 'active' : ''}`}
            onClick={() => setActiveTab('results')}
            disabled={!selectedEntity}
          >
            Results {results.length > 0 ? `(${results.length})` : ''}
          </button>
        </div>

        <div className="qb-content-scroll">
          <div className="qb-container-inner">

            {activeTab === 'builder' && (
              <>
                {/* Entity Selection Section */}
                <div className="qb-section">
                  <div className="qb-section-title">
                    <span>Data Source</span>
                  </div>
                  <div className="qb-row">
                    <div className="qb-col">
                      <EntitySelector
                        api={api}
                        onEntitySelected={handleEntitySelected}
                        selectedEntity={selectedEntity}
                      />
                    </div>
                    <div className="qb-col">
                      {selectedEntity && (
                        <ViewSelector
                          api={api}
                          entityLogicalName={selectedEntity.logicalName}
                          onViewSelected={handleViewSelected}
                        />
                      )}
                    </div>
                  </div>

                  {selectedEntity && (
                    <>
                      <div style={{ marginTop: '15px' }}>
                        <button
                          className="qb-btn qb-btn-secondary qb-btn--medium"
                          onClick={() => setShowRelationshipManager(!showRelationshipManager)}
                        >
                          <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor" className="qb-icon">
                            <path d="M7 2v10M2 7h10" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
                          </svg>
                          Add Related Entity
                        </button>
                      </div>
                      {(showRelationshipManager || joinedEntities.length > 0) && (
                        <div style={{ marginTop: '15px', borderTop: '1px solid #eee', paddingTop: '15px' }}>
                          <RelationshipManager
                            api={api}
                            entities={allEntities}
                            onAddJoin={(joined) => {
                              handleAddJoin(joined);
                              setShowRelationshipManager(false);
                            }}
                            onRemoveJoin={handleRemoveJoin}
                            isExpanded={showRelationshipManager}
                            onToggleExpanded={setShowRelationshipManager}
                          />
                        </div>
                      )}
                    </>
                  )}
                </div>

                {selectedEntity && (
                  <>
                    {/* Columns Section */}
                    <div className="qb-section">
                      <div className="qb-section-title">
                        <span>Columns ({selectedColumns.length})</span>
                      </div>
                      <ColumnSelector
                        api={api}
                        entities={allEntities}
                        selectedColumns={selectedColumns}
                        onToggleColumn={handleToggleColumn}
                      />
                    </div>

                    {/* Filters Section */}
                    <div className="qb-section">
                      <div className="qb-section-title">
                        <span>Filters</span>
                      </div>
                      <FilterBuilder
                        api={api}
                        entities={allEntities}
                        filters={filters}
                        onChange={setFilters}
                      />
                    </div>

                    {/* Run Query Button at bottom of builder */}
                    <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: '10px' }}>
                      <button
                        className="qb-btn qb-btn-primary qb-btn--large"
                        onClick={handleRunQuery}
                        disabled={loading}
                      >
                        {loading ? (
                          <>
                            <svg className="qb-icon spinner" width="16" height="16" viewBox="0 0 16 16" fill="none">
                              <circle cx="8" cy="8" r="6" stroke="currentColor" strokeWidth="2" strokeDasharray="10 5" />
                            </svg>
                            Running...
                          </>
                        ) : (
                          <>
                            <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor" className="qb-icon">
                              <path d="M4 3l8 5-8 5V3z" fill="currentColor"/>
                            </svg>
                            Run Query
                          </>
                        )}
                      </button>
                    </div>
                  </>
                )}
              </>
            )}

            {activeTab === 'results' && (
              <div className="qb-section" style={{ flex: 1, overflow: 'hidden', display: 'flex', flexDirection: 'column', minHeight: '60vh' }}>
                <div className="qb-section-title">
                  <span>Results ({results.length})</span>
                  <div style={{ display: 'flex', gap: '10px' }}>
                    <button
                      className="qb-btn qb-btn-secondary qb-btn--medium"
                      onClick={() => setActiveTab('builder')}
                    >
                      <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor" className="qb-icon">
                        <path d="M8 3L4 7l4 4" stroke="currentColor" strokeWidth="1.5" fill="none" strokeLinecap="round" strokeLinejoin="round"/>
                      </svg>
                      Back to Builder
                    </button>
                    <button
                      className="qb-btn qb-btn-primary qb-btn--medium"
                      onClick={handleRunQuery}
                      disabled={loading}
                    >
                      {loading ? (
                        <>
                          <svg className="qb-icon spinner" width="14" height="14" viewBox="0 0 14 14" fill="none">
                            <circle cx="7" cy="7" r="5" stroke="currentColor" strokeWidth="2" strokeDasharray="8 4" />
                          </svg>
                          Running...
                        </>
                      ) : (
                        <>
                          <svg width="14" height="14" viewBox="0 0 14 14" fill="currentColor" className="qb-icon">
                            <path d="M11 5.5a4.5 4.5 0 1 1-9 0 4.5 4.5 0 0 1 9 0zM10 11l3 3" stroke="currentColor" strokeWidth="1.5" fill="none" strokeLinecap="round"/>
                          </svg>
                          Refresh
                        </>
                      )}
                    </button>
                    <button
                      className="qb-btn qb-btn-secondary qb-btn--medium"
                      onClick={() => setShowExportModal(true)}
                      disabled={loading || results.length === 0}
                      title="Export results to CSV or Excel"
                    >
                      <svg width="14" height="14" viewBox="0 0 14 14" fill="none" className="qb-icon">
                        <path d="M7 1v8M7 9L4 6M7 9l3-3M2 11h10v2H2v-2z" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                      </svg>
                      Export
                    </button>
                  </div>
                </div>

                {/* 
                  Debug Query Display Removed. 
                  If needed for debugging, uncomment below or move to console.
                */}
                {/* {executedQuery && (
                  <div style={{ fontSize: '11px', color: '#666', margin: '5px 0', fontFamily: 'monospace', wordBreak: 'break-all' }}>
                    Query: {executedQuery}
                  </div>
                )} */}

                <ResultsTable
                  data={results}
                  columns={selectedColumns}
                  entities={allEntities}
                  loading={loading}
                  error={error}
                  orgUrl={api?.['baseUrl']?.replace('/api/data/v9.2', '')}
                />
              </div>
            )}
          </div>
        </div>

        {/* Export Modal */}
        {showExportModal && (
          <div className="export-modal-overlay" onClick={() => !exporting && setShowExportModal(false)}>
            <div className="export-modal" onClick={(e) => e.stopPropagation()}>
              <div className="export-modal__header">
                <h2>Export Results</h2>
                <button
                  className="export-modal__close"
                  onClick={() => !exporting && setShowExportModal(false)}
                  disabled={exporting}
                >
                  ×
                </button>
              </div>

              <div className="export-modal__body">
                {exportError && (
                  <div className="export-modal__error">
                    {exportError}
                  </div>
                )}

                {exporting ? (
                  <div className="export-modal__progress">
                    <div className="spinner"></div>
                    <p>
                      {exportProgress
                        ? `Fetching records... ${exportProgress.fetched.toLocaleString()} records retrieved`
                        : 'Preparing export...'}
                    </p>
                    {exportProgress && !exportProgress.complete && (
                      <p className="export-modal__progress-hint">
                        Fetching all pages of data. This may take a while for large datasets.
                      </p>
                    )}
                  </div>
                ) : (
                  <>
                    <p className="export-modal__info">
                      Current results: <strong>{results.length.toLocaleString()}</strong> records
                    </p>
                    <p className="export-modal__description">
                      Choose to export the current page of results, or fetch and export ALL matching records (ignores the 5,000 record page limit).
                    </p>

                    <div className="export-modal__section">
                      <h3>Export Current Results ({results.length.toLocaleString()} records)</h3>
                      <div className="export-modal__buttons">
                        <button
                          className="qb-btn qb-btn-secondary qb-btn--medium"
                          onClick={() => handleExport('csv', false)}
                        >
                          <svg width="14" height="14" viewBox="0 0 14 14" fill="none" className="qb-icon">
                            <path d="M2 1h7l3 3v9H2V1z" stroke="currentColor" strokeWidth="1.2"/>
                            <text x="4" y="10" fontSize="5" fill="currentColor" fontWeight="bold">CSV</text>
                          </svg>
                          Download CSV
                        </button>
                        <button
                          className="qb-btn qb-btn-secondary qb-btn--medium"
                          onClick={() => handleExport('excel', false)}
                        >
                          <svg width="14" height="14" viewBox="0 0 14 14" fill="none" className="qb-icon">
                            <path d="M2 1h7l3 3v9H2V1z" stroke="#217346" strokeWidth="1.2"/>
                            <text x="3" y="10" fontSize="5" fill="#217346" fontWeight="bold">XLS</text>
                          </svg>
                          Download Excel
                        </button>
                      </div>
                    </div>

                    <div className="export-modal__section export-modal__section--highlight">
                      <h3>Export ALL Matching Records</h3>
                      <p className="export-modal__note">
                        This will fetch all pages of data matching your query and filters, bypassing the 5,000 record limit.
                      </p>
                      <div className="export-modal__buttons">
                        <button
                          className="qb-btn qb-btn-primary qb-btn--medium"
                          onClick={() => handleExport('csv', true)}
                        >
                          <svg width="14" height="14" viewBox="0 0 14 14" fill="none" className="qb-icon">
                            <path d="M2 1h7l3 3v9H2V1z" stroke="currentColor" strokeWidth="1.2"/>
                            <text x="4" y="10" fontSize="5" fill="currentColor" fontWeight="bold">CSV</text>
                          </svg>
                          Export All to CSV
                        </button>
                        <button
                          className="qb-btn qb-btn-primary qb-btn--medium"
                          onClick={() => handleExport('excel', true)}
                        >
                          <svg width="14" height="14" viewBox="0 0 14 14" fill="none" className="qb-icon">
                            <path d="M2 1h7l3 3v9H2V1z" stroke="currentColor" strokeWidth="1.2"/>
                            <text x="3" y="10" fontSize="5" fill="currentColor" fontWeight="bold">XLS</text>
                          </svg>
                          Export All to Excel
                        </button>
                      </div>
                    </div>
                  </>
                )}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default QueryBuilder;
