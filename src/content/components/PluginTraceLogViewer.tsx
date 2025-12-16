import React, { useMemo, useState } from 'react';

export interface PluginTraceLogRecord {
  id: string;
  createdOn?: string;
  messageName?: string;
  primaryEntity?: string;
  typeName?: string;
  mode?: number | string;
  depth?: number;
  operationCorrelationId?: string;
  performanceDurationMs?: number;
  executionStart?: string;
  requestId?: string;
  exceptionDetails?: string;
  messageBlock?: string;
}

export interface PluginTraceLogData {
  logs: PluginTraceLogRecord[];
  moreRecords?: boolean;
  error?: string;
}

interface PluginTraceLogViewerProps {
  data: PluginTraceLogData | null;
  onClose: () => void;
  onRefresh: () => void;
  onClear: () => void;
}

const formatDateTime = (value?: string): string => {
  if (!value) {
    return '—';
  }
  try {
    return new Date(value).toLocaleString();
  } catch {
    return value;
  }
};

const formatDuration = (value?: number): string => {
  if (value === null || value === undefined) {
    return '—';
  }
  if (Number.isFinite(value)) {
    return `${value} ms`;
  }
  return String(value);
};

const formatMode = (mode?: number | string): string => {
  if (typeof mode === 'number') {
    if (mode === 0) return 'Synchronous';
    if (mode === 1) return 'Asynchronous';
  }
  if (typeof mode === 'string') {
    return mode;
  }
  return 'Unknown';
};

const hasException = (log: PluginTraceLogRecord): boolean =>
  Boolean(log.exceptionDetails && log.exceptionDetails.trim().length > 0);

const normalize = (value?: string | number | null): string =>
  String(value ?? '').trim().toLowerCase();

const parseDateInput = (value: string): number | null => {
  if (!value) return null;
  const date = new Date(value);
  const time = date.getTime();
  return Number.isFinite(time) ? time : null;
};

const getModeKind = (mode?: number | string): 'sync' | 'async' | 'unknown' => {
  if (mode === 0) return 'sync';
  if (mode === 1) return 'async';
  if (typeof mode === 'string') {
    const m = mode.toLowerCase();
    if (m.includes('sync')) return 'sync';
    if (m.includes('async')) return 'async';
  }
  return 'unknown';
};

const PluginTraceLogViewer: React.FC<PluginTraceLogViewerProps> = ({
  data,
  onClose,
  onRefresh,
  onClear,
}) => {
  const [activeTab, setActiveTab] = useState<'recent' | 'exceptions'>('recent');
  const [selectedLog, setSelectedLog] = useState<PluginTraceLogRecord | null>(null);
  const [copiedField, setCopiedField] = useState<string>('');
  const [isClearing, setIsClearing] = useState(false);
  const [showFilters, setShowFilters] = useState(true);

  // Filter state (client-side, applied to currently loaded logs)
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');
  const [pluginFilter, setPluginFilter] = useState('');
  const [messageFilter, setMessageFilter] = useState('');
  const [entityFilter, setEntityFilter] = useState('');
  const [correlationIdFilter, setCorrelationIdFilter] = useState('');
  const [requestIdFilter, setRequestIdFilter] = useState('');
  const [textSearch, setTextSearch] = useState('');
  const [includeSync, setIncludeSync] = useState(true);
  const [includeAsync, setIncludeAsync] = useState(true);
  const [durationMin, setDurationMin] = useState('');
  const [durationMax, setDurationMax] = useState('');

  if (!data) {
    return null;
  }

  if (data.error) {
    return (
      <div className="d365-dialog-overlay">
        <div className="d365-dialog-modal d365-trace-modal">
          <div className="d365-dialog-header d365-trace-header">
            <h2>Plugin Trace Logs</h2>
            <button className="d365-dialog-close" onClick={onClose} title="Close">
              ×
            </button>
          </div>
          <div className="d365-dialog-content">
            <div className="d365-dialog-error">
              <div className="d365-dialog-error-icon">⚠️</div>
              <div className="d365-dialog-error-message">{data.error}</div>
              <div className="d365-dialog-error-hint">
                Ensure plugin tracing is enabled (Organization Settings &gt; Administration &gt; System Settings) and that
                trace records exist.
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  const baseLogs = useMemo(() => {
    return activeTab === 'exceptions' ? data.logs.filter(hasException) : data.logs;
  }, [activeTab, data.logs]);

  const filteredLogs = useMemo(() => {
    const fromTime = parseDateInput(dateFrom);
    const toTime = parseDateInput(dateTo);
    const plugin = normalize(pluginFilter);
    const message = normalize(messageFilter);
    const entity = normalize(entityFilter);
    const correlationId = normalize(correlationIdFilter);
    const requestId = normalize(requestIdFilter);
    const search = normalize(textSearch);

    const minDuration = durationMin.trim() ? Number(durationMin) : null;
    const maxDuration = durationMax.trim() ? Number(durationMax) : null;
    const hasMinDuration = Number.isFinite(minDuration as number);
    const hasMaxDuration = Number.isFinite(maxDuration as number);

    const restrictMode = !(includeSync && includeAsync);

    return baseLogs.filter((log) => {
      // Date range
      if (fromTime !== null || toTime !== null) {
        const createdTime = log.createdOn ? new Date(log.createdOn).getTime() : NaN;
        if (!Number.isFinite(createdTime)) return false;
        if (fromTime !== null && createdTime < fromTime) return false;
        if (toTime !== null && createdTime > toTime) return false;
      }

      // Mode
      if (!includeSync && !includeAsync) return false;
      if (restrictMode) {
        const kind = getModeKind(log.mode);
        if (kind === 'sync' && !includeSync) return false;
        if (kind === 'async' && !includeAsync) return false;
        // unknown -> include
      }

      // Duration range
      if (hasMinDuration || hasMaxDuration) {
        const value = log.performanceDurationMs;
        if (typeof value !== 'number' || !Number.isFinite(value)) return false;
        if (hasMinDuration && value < (minDuration as number)) return false;
        if (hasMaxDuration && value > (maxDuration as number)) return false;
      }

      // Field filters
      if (plugin && !normalize(log.typeName).includes(plugin)) return false;
      if (message && !normalize(log.messageName).includes(message)) return false;
      if (entity && !normalize(log.primaryEntity).includes(entity)) return false;
      if (correlationId && !normalize(log.operationCorrelationId).includes(correlationId)) return false;
      if (requestId && !normalize(log.requestId).includes(requestId)) return false;

      // Smart search across common fields
      if (search) {
        const haystack = [
          log.messageName,
          log.primaryEntity,
          log.typeName,
          log.operationCorrelationId,
          log.requestId,
          log.exceptionDetails,
          log.messageBlock,
        ]
          .map(normalize)
          .filter(Boolean)
          .join(' | ');

        if (!haystack.includes(search)) return false;
      }

      return true;
    });
  }, [
    baseLogs,
    dateFrom,
    dateTo,
    pluginFilter,
    messageFilter,
    entityFilter,
    correlationIdFilter,
    requestIdFilter,
    textSearch,
    includeSync,
    includeAsync,
    durationMin,
    durationMax,
  ]);

  // Auto-select first log if none selected
  useMemo(() => {
    if (filteredLogs.length > 0 && !selectedLog) {
      setSelectedLog(filteredLogs[0]);
    }
  }, [filteredLogs, selectedLog]);

  // If selected log is filtered out, clear selection so we can auto-select again
  useMemo(() => {
    if (selectedLog && !filteredLogs.some((l) => l.id === selectedLog.id)) {
      setSelectedLog(null);
    }
  }, [filteredLogs, selectedLog]);

  const clearFilters = () => {
    setDateFrom('');
    setDateTo('');
    setPluginFilter('');
    setMessageFilter('');
    setEntityFilter('');
    setCorrelationIdFilter('');
    setRequestIdFilter('');
    setTextSearch('');
    setIncludeSync(true);
    setIncludeAsync(true);
    setDurationMin('');
    setDurationMax('');
  };

  const setQuickRangeMinutes = (minutes: number) => {
    const now = Date.now();
    const from = new Date(now - minutes * 60 * 1000);
    const pad = (n: number) => String(n).padStart(2, '0');
    // datetime-local format: YYYY-MM-DDTHH:mm
    const localValue =
      `${from.getFullYear()}-${pad(from.getMonth() + 1)}-${pad(from.getDate())}` +
      `T${pad(from.getHours())}:${pad(from.getMinutes())}`;
    setDateFrom(localValue);
    setDateTo('');
  };

  const handleCopy = async (value?: string, key?: string) => {
    if (!value) {
      return;
    }
    try {
      await navigator.clipboard.writeText(value);
      if (key) {
        setCopiedField(key);
        setTimeout(() => setCopiedField(''), 2000);
      }
    } catch (error) {
      console.error('Failed to copy value', error);
    }
  };

  const handleClear = async () => {
    if (isClearing) return;
    try {
      setIsClearing(true);
      setSelectedLog(null);
      setCopiedField('');
      await onClear();
    } finally {
      setIsClearing(false);
    }
  };

  const renderLogListItem = (log: PluginTraceLogRecord) => {
    const isSelected = selectedLog?.id === log.id;

    return (
      <div
        key={log.id}
        className={`d365-trace-list-item ${isSelected ? 'd365-trace-list-item-selected' : ''}`}
        onClick={() => setSelectedLog(log)}
      >
        <div className="d365-trace-list-item-top">
          {hasException(log) && <span className="d365-trace-list-item-error">⚠</span>}
          <div className="d365-trace-list-item-message">{log.messageName || 'Unknown Message'}</div>
        </div>
        {log.primaryEntity && (
          <div className="d365-trace-list-item-entity">{log.primaryEntity}</div>
        )}
        <div className="d365-trace-list-item-time">{formatDateTime(log.createdOn)}</div>
      </div>
    );
  };

  const renderDetails = () => {
    if (!selectedLog) {
      return (
        <div className="d365-trace-details-empty">
          Select a trace log to view details
        </div>
      );
    }

    return (
      <div className="d365-trace-details-panel">
        <div className="d365-trace-details-header">
          <div className="d365-trace-details-title">
            <span className="d365-trace-message-name">{selectedLog.messageName || 'Unknown Message'}</span>
            {selectedLog.primaryEntity && (
              <span className="d365-trace-entity-badge">{selectedLog.primaryEntity}</span>
            )}
            {hasException(selectedLog) && <span className="d365-trace-error-badge">Exception</span>}
          </div>
          <div className="d365-trace-details-meta">
            <span>{formatMode(selectedLog.mode)}</span>
            <span>•</span>
            <span>{formatDuration(selectedLog.performanceDurationMs)}</span>
          </div>
        </div>

        <div className="d365-trace-details-content">
          <div className="d365-trace-detail-grid">
            <div className="d365-trace-detail">
              <span className="d365-trace-detail-label">Plugin Type</span>
              <span className="d365-trace-detail-value">{selectedLog.typeName || '—'}</span>
            </div>
            <div className="d365-trace-detail">
              <span className="d365-trace-detail-label">Created On</span>
              <span className="d365-trace-detail-value">{formatDateTime(selectedLog.createdOn)}</span>
            </div>
            <div className="d365-trace-detail">
              <span className="d365-trace-detail-label">Execution Start</span>
              <span className="d365-trace-detail-value">
                {formatDateTime(selectedLog.executionStart)}
              </span>
            </div>
            <div className="d365-trace-detail">
              <span className="d365-trace-detail-label">Depth</span>
              <span className="d365-trace-detail-value">{selectedLog.depth ?? '—'}</span>
            </div>
            <div className="d365-trace-detail">
              <span className="d365-trace-detail-label">Correlation ID</span>
              <div className="d365-trace-detail-value">
                {selectedLog.operationCorrelationId || '—'}
                {selectedLog.operationCorrelationId && (
                  <button
                    className="d365-trace-copy-btn"
                    onClick={() => handleCopy(selectedLog.operationCorrelationId, `correlation-${selectedLog.id}`)}
                    title="Copy correlation ID"
                  >
                    {copiedField === `correlation-${selectedLog.id}` ? (
                      <span className="d365-copy-success">✓</span>
                    ) : (
                      <img
                        src={chrome.runtime.getURL('icons/rg_copy.svg')}
                        alt="Copy"
                        className="d365-copy-icon"
                      />
                    )}
                  </button>
                )}
              </div>
            </div>
            <div className="d365-trace-detail">
              <span className="d365-trace-detail-label">Request ID</span>
              <div className="d365-trace-detail-value">
                {selectedLog.requestId || '—'}
                {selectedLog.requestId && (
                  <button
                    className="d365-trace-copy-btn"
                    onClick={() => handleCopy(selectedLog.requestId, `request-${selectedLog.id}`)}
                    title="Copy request ID"
                  >
                    {copiedField === `request-${selectedLog.id}` ? (
                      <span className="d365-copy-success">✓</span>
                    ) : (
                      <img
                        src={chrome.runtime.getURL('icons/rg_copy.svg')}
                        alt="Copy"
                        className="d365-copy-icon"
                      />
                    )}
                  </button>
                )}
              </div>
            </div>
          </div>

          {selectedLog.messageBlock && (
            <div className="d365-trace-section">
              <div className="d365-trace-section-header">
                <div className="d365-trace-section-title">Message Log</div>
                <button
                  className="d365-trace-copy-btn"
                  onClick={() => handleCopy(selectedLog.messageBlock, `message-${selectedLog.id}`)}
                  title="Copy message log"
                >
                  {copiedField === `message-${selectedLog.id}` ? (
                    <span className="d365-copy-success">✓ Copied</span>
                  ) : (
                    <>
                      <img
                        src={chrome.runtime.getURL('icons/rg_copy.svg')}
                        alt="Copy"
                        className="d365-copy-icon"
                      />
                      <span>Copy</span>
                    </>
                  )}
                </button>
              </div>
              <pre className="d365-trace-pre">{selectedLog.messageBlock}</pre>
            </div>
          )}

          {selectedLog.exceptionDetails && (
            <div className="d365-trace-section">
              <div className="d365-trace-section-header">
                <div className="d365-trace-section-title">Exception Details</div>
                <button
                  className="d365-trace-copy-btn"
                  onClick={() => handleCopy(selectedLog.exceptionDetails, `exception-${selectedLog.id}`)}
                  title="Copy exception details"
                >
                  {copiedField === `exception-${selectedLog.id}` ? (
                    <span className="d365-copy-success">✓ Copied</span>
                  ) : (
                    <>
                      <img
                        src={chrome.runtime.getURL('icons/rg_copy.svg')}
                        alt="Copy"
                        className="d365-copy-icon"
                      />
                      <span>Copy</span>
                    </>
                  )}
                </button>
              </div>
              <pre className="d365-trace-pre d365-trace-pre-error">{selectedLog.exceptionDetails}</pre>
            </div>
          )}
        </div>
      </div>
    );
  };

  return (
    <div className="d365-dialog-overlay">
      <div className="d365-dialog-modal d365-trace-modal">
        <div className="d365-dialog-header d365-trace-header">
          <div className="d365-trace-header-left">
            <h2>Plugin Trace Logs</h2>
          </div>
          <div className="d365-trace-header-actions">
            <button
              className="d365-trace-clear"
              onClick={handleClear}
              title="Clear the currently displayed logs (does not delete records)"
              disabled={isClearing || data.logs.length === 0}
            >
              🧹
            </button>
            <button
              className="d365-trace-refresh"
              onClick={onRefresh}
              title="Refresh trace logs"
              disabled={isClearing}
            >
              ↻
            </button>
            <button className="d365-dialog-close" onClick={onClose} title="Close">
              ×
            </button>
          </div>
        </div>

        <div className="d365-trace-filter-bar">
          <button
            className="d365-trace-filter-toggle"
            onClick={() => setShowFilters((prev) => !prev)}
            type="button"
            title={showFilters ? 'Hide filters' : 'Show filters'}
          >
            Filter {showFilters ? '▲' : '▼'}
          </button>
          <div className="d365-trace-filter-summary">
            Showing <strong>{filteredLogs.length}</strong> of <strong>{data.logs.length}</strong>
          </div>
          <button
            className="d365-trace-filter-clear-btn"
            onClick={clearFilters}
            type="button"
            title="Clear all filters"
            disabled={
              !dateFrom &&
              !dateTo &&
              !pluginFilter &&
              !messageFilter &&
              !entityFilter &&
              !correlationIdFilter &&
              !requestIdFilter &&
              !textSearch &&
              includeSync &&
              includeAsync &&
              !durationMin &&
              !durationMax
            }
          >
            Clear
          </button>
        </div>

        {showFilters && (
          <div className="d365-trace-filter-panel">
            <div className="d365-trace-filter-grid">
              <div className="d365-trace-filter-field">
                <label>Date From</label>
                <input
                  type="datetime-local"
                  value={dateFrom}
                  onChange={(e) => setDateFrom(e.target.value)}
                />
              </div>
              <div className="d365-trace-filter-field">
                <label>To</label>
                <input
                  type="datetime-local"
                  value={dateTo}
                  onChange={(e) => setDateTo(e.target.value)}
                />
              </div>
              <div className="d365-trace-filter-field d365-trace-filter-quick">
                <label>Quick</label>
                <div className="d365-trace-filter-quick-buttons">
                  <button type="button" onClick={() => setQuickRangeMinutes(5)}>Last 5m</button>
                  <button type="button" onClick={() => setQuickRangeMinutes(30)}>Last 30m</button>
                  <button type="button" onClick={() => setQuickRangeMinutes(60)}>Last 1h</button>
                </div>
              </div>

              <div className="d365-trace-filter-field">
                <label>Plugin</label>
                <input
                  type="text"
                  placeholder="Type name contains..."
                  value={pluginFilter}
                  onChange={(e) => setPluginFilter(e.target.value)}
                />
              </div>
              <div className="d365-trace-filter-field">
                <label>Message</label>
                <input
                  type="text"
                  placeholder="Message name contains..."
                  value={messageFilter}
                  onChange={(e) => setMessageFilter(e.target.value)}
                />
              </div>
              <div className="d365-trace-filter-field">
                <label>Entity</label>
                <input
                  type="text"
                  placeholder="Primary entity contains..."
                  value={entityFilter}
                  onChange={(e) => setEntityFilter(e.target.value)}
                />
              </div>

              <div className="d365-trace-filter-field">
                <label>Correlation Id</label>
                <input
                  type="text"
                  placeholder="Contains..."
                  value={correlationIdFilter}
                  onChange={(e) => setCorrelationIdFilter(e.target.value)}
                />
              </div>
              <div className="d365-trace-filter-field">
                <label>Request Id</label>
                <input
                  type="text"
                  placeholder="Contains..."
                  value={requestIdFilter}
                  onChange={(e) => setRequestIdFilter(e.target.value)}
                />
              </div>

              <div className="d365-trace-filter-field d365-trace-filter-wide">
                <label>Smart Search</label>
                <input
                  type="text"
                  placeholder="Search message, plugin, entity, ids, exception, message log..."
                  value={textSearch}
                  onChange={(e) => setTextSearch(e.target.value)}
                />
              </div>

              <div className="d365-trace-filter-field">
                <label>Mode</label>
                <div className="d365-trace-filter-checks">
                  <label>
                    <input
                      type="checkbox"
                      checked={includeSync}
                      onChange={(e) => setIncludeSync(e.target.checked)}
                    />
                    Synchronous
                  </label>
                  <label>
                    <input
                      type="checkbox"
                      checked={includeAsync}
                      onChange={(e) => setIncludeAsync(e.target.checked)}
                    />
                    Asynchronous
                  </label>
                </div>
              </div>

              <div className="d365-trace-filter-field">
                <label>Duration (ms)</label>
                <div className="d365-trace-filter-range">
                  <input
                    type="number"
                    min="0"
                    placeholder="Min"
                    value={durationMin}
                    onChange={(e) => setDurationMin(e.target.value)}
                  />
                  <span>–</span>
                  <input
                    type="number"
                    min="0"
                    placeholder="Max"
                    value={durationMax}
                    onChange={(e) => setDurationMax(e.target.value)}
                  />
                </div>
              </div>

              <div className="d365-trace-filter-field d365-trace-filter-check">
                <label>
                  <input
                    type="checkbox"
                    checked={activeTab === 'exceptions'}
                    onChange={(e) => {
                      setActiveTab(e.target.checked ? 'exceptions' : 'recent');
                      setSelectedLog(null);
                    }}
                  />
                  Exceptions only
                </label>
              </div>
            </div>
            <div className="d365-trace-filter-note">
              Filters apply to the currently loaded logs. Adjust “Plugin trace log default limit” in extension settings to fetch more.
            </div>
          </div>
        )}

        <div className="d365-trace-tabs">
          <button
            className={`d365-trace-tab ${activeTab === 'recent' ? 'd365-trace-tab-active' : ''}`}
            onClick={() => {
              setActiveTab('recent');
              setSelectedLog(null);
            }}
          >
            Recent ({data.logs.length})
          </button>
          <button
            className={`d365-trace-tab ${activeTab === 'exceptions' ? 'd365-trace-tab-active' : ''}`}
            onClick={() => {
              setActiveTab('exceptions');
              setSelectedLog(null);
            }}
          >
            Exceptions ({data.logs.filter(hasException).length})
          </button>
        </div>

        <div className="d365-trace-content-split">
          {filteredLogs.length === 0 ? (
            <div className="d365-trace-empty">
              {activeTab === 'exceptions'
                ? 'No exception entries were found in the selected trace logs.'
                : 'No plugin trace logs were returned. Try refreshing or increasing the trace depth.'}
            </div>
          ) : (
            <>
              <div className="d365-trace-list-panel">
                {filteredLogs.map(renderLogListItem)}
              </div>
              <div className="d365-trace-details-panel-wrapper">
                {renderDetails()}
                {data.moreRecords && selectedLog === null && (
                  <div className="d365-trace-more-records">
                    Showing the most recent {data.logs.length} entries. Use Dynamics views for full history.
                  </div>
                )}
              </div>
            </>
          )}
        </div>
      </div>
    </div>
  );
};

export default PluginTraceLogViewer;
