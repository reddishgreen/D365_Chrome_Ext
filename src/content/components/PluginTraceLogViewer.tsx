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
  createdByName?: string;
  createdById?: string;
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
    return '\u2014';
  }
  try {
    return new Date(value).toLocaleString();
  } catch {
    return value;
  }
};

const formatDuration = (value?: number): string => {
  if (value === null || value === undefined) {
    return '\u2014';
  }
  if (Number.isFinite(value)) {
    if (value >= 1000) {
      return `${(value / 1000).toFixed(1)}s`;
    }
    return `${value}ms`;
  }
  return String(value);
};

const formatDurationShort = (value?: number): string => {
  if (value === null || value === undefined) return '';
  if (!Number.isFinite(value)) return '';
  if (value >= 1000) return `${(value / 1000).toFixed(1)}s`;
  return `${value}ms`;
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

const formatModeShort = (mode?: number | string): string => {
  if (mode === 0) return 'Sync';
  if (mode === 1) return 'Async';
  if (typeof mode === 'string') {
    const m = mode.toLowerCase();
    if (m.includes('sync') && !m.includes('async')) return 'Sync';
    if (m.includes('async')) return 'Async';
  }
  return '';
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

const truncatePluginName = (name?: string): string => {
  if (!name) return '';
  // If it contains dots (namespace), show just the last segment
  const parts = name.split('.');
  if (parts.length > 2) {
    return `...${parts.slice(-2).join('.')}`;
  }
  return name;
};

const isSystemUserName = (name?: string): boolean =>
  Boolean(name && name.trim().toLowerCase() === 'system');

const getDisplayUserName = (name?: string): string | null => {
  if (!name) return null;
  return isSystemUserName(name) ? null : name.trim();
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
  const [messageExpanded, setMessageExpanded] = useState(true);
  const [exceptionExpanded, setExceptionExpanded] = useState(true);

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
              &times;
            </button>
          </div>
          <div className="d365-dialog-content">
            <div className="d365-dialog-error">
              <div className="d365-dialog-error-icon">&nbsp;&#9888;&#65039;</div>
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
          log.createdByName,
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

  const getDurationClass = (ms?: number): string => {
    if (ms === undefined || ms === null || !Number.isFinite(ms)) return '';
    if (ms >= 2000) return 'd365-trace-list-item-duration--slow';
    if (ms >= 500) return 'd365-trace-list-item-duration--medium';
    return '';
  };

  const renderLogListItem = (log: PluginTraceLogRecord) => {
    const isSelected = selectedLog?.id === log.id;
    const modeShort = formatModeShort(log.mode);
    const durationShort = formatDurationShort(log.performanceDurationMs);
    const createdByDisplay = getDisplayUserName(log.createdByName);

    return (
      <div
        key={log.id}
        className={`d365-trace-list-item ${isSelected ? 'd365-trace-list-item-selected' : ''} ${hasException(log) ? 'd365-trace-list-item--error' : ''}`}
        onClick={() => setSelectedLog(log)}
      >
        <div className="d365-trace-list-item-top">
          {hasException(log) && <span className="d365-trace-list-item-error">&nbsp;&#9888;</span>}
          <div className="d365-trace-list-item-message">{log.messageName || 'Unknown Message'}</div>
          <div className="d365-trace-list-item-badges">
            {modeShort && (
              <span className={`d365-trace-list-item-mode d365-trace-list-item-mode--${modeShort.toLowerCase()}`}>
                {modeShort}
              </span>
            )}
            {durationShort && (
              <span className={`d365-trace-list-item-duration ${getDurationClass(log.performanceDurationMs)}`}>
                {durationShort}
              </span>
            )}
          </div>
        </div>
        {log.typeName && (
          <div className="d365-trace-list-item-plugin" title={log.typeName}>
            {truncatePluginName(log.typeName)}
          </div>
        )}
        <div className="d365-trace-list-item-bottom">
          {log.primaryEntity && (
            <span className="d365-trace-list-item-entity">{log.primaryEntity}</span>
          )}
          <span className="d365-trace-list-item-time">{formatDateTime(log.createdOn)}</span>
        </div>
        {createdByDisplay && (
          <div className="d365-trace-list-item-user">
            <span className="d365-trace-list-item-user-icon">&nbsp;&#128100;</span>
            {createdByDisplay}
          </div>
        )}
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

    const createdByDisplay = getDisplayUserName(selectedLog.createdByName);

    const metaItems: React.ReactNode[] = [
      <span key="mode">{formatMode(selectedLog.mode)}</span>,
      <span key="duration">{formatDuration(selectedLog.performanceDurationMs)}</span>,
    ];

    if (createdByDisplay) {
      metaItems.push(
        <span key="user" className="d365-trace-details-meta-user">
          &nbsp;&#128100; {createdByDisplay}
        </span>
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
            {metaItems.map((item, index) => (
              <React.Fragment key={index}>
                {index > 0 && <span>&bull;</span>}
                {item}
              </React.Fragment>
            ))}
          </div>
        </div>

        <div className="d365-trace-details-content">
          <div className="d365-trace-detail-grid">
            <div className="d365-trace-detail">
              <span className="d365-trace-detail-label">Plugin Type</span>
              <span className="d365-trace-detail-value">{selectedLog.typeName || '\u2014'}</span>
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
              <span className="d365-trace-detail-value">{selectedLog.depth ?? '\u2014'}</span>
            </div>
            {createdByDisplay && (
              <div className="d365-trace-detail">
                <span className="d365-trace-detail-label">Created By</span>
                <span className="d365-trace-detail-value">
                  <span className="d365-trace-details-meta-user">&nbsp;&#128100; {createdByDisplay}</span>
                </span>
              </div>
            )}
            <div className="d365-trace-detail">
              <span className="d365-trace-detail-label">Correlation ID</span>
              <div className="d365-trace-detail-value">
                {selectedLog.operationCorrelationId || '\u2014'}
                {selectedLog.operationCorrelationId && (
                  <button
                    className="d365-trace-copy-btn"
                    onClick={() => handleCopy(selectedLog.operationCorrelationId, `correlation-${selectedLog.id}`)}
                    title="Copy correlation ID"
                  >
                    {copiedField === `correlation-${selectedLog.id}` ? (
                      <span className="d365-copy-success">&checkmark;</span>
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
                {selectedLog.requestId || '\u2014'}
                {selectedLog.requestId && (
                  <button
                    className="d365-trace-copy-btn"
                    onClick={() => handleCopy(selectedLog.requestId, `request-${selectedLog.id}`)}
                    title="Copy request ID"
                  >
                    {copiedField === `request-${selectedLog.id}` ? (
                      <span className="d365-copy-success">&checkmark;</span>
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
                <div className="d365-trace-section-actions">
                  <button
                    className="d365-trace-section-toggle"
                    onClick={() => setMessageExpanded(prev => !prev)}
                    title={messageExpanded ? 'Collapse' : 'Expand'}
                  >
                    {messageExpanded ? 'Collapse' : 'Expand'}
                  </button>
                  <button
                    className="d365-trace-copy-btn"
                    onClick={() => handleCopy(selectedLog.messageBlock, `message-${selectedLog.id}`)}
                    title="Copy message log"
                  >
                    {copiedField === `message-${selectedLog.id}` ? (
                      <span className="d365-copy-success">&checkmark; Copied</span>
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
              </div>
              <pre className={`d365-trace-pre ${messageExpanded ? 'd365-trace-pre--expanded' : 'd365-trace-pre--collapsed'}`}>
                {selectedLog.messageBlock}
              </pre>
            </div>
          )}

          {selectedLog.exceptionDetails && (
            <div className="d365-trace-section d365-trace-section--error">
              <div className="d365-trace-section-header">
                <div className="d365-trace-section-title d365-trace-section-title--error">Exception Details</div>
                <div className="d365-trace-section-actions">
                  <button
                    className="d365-trace-section-toggle"
                    onClick={() => setExceptionExpanded(prev => !prev)}
                    title={exceptionExpanded ? 'Collapse' : 'Expand'}
                  >
                    {exceptionExpanded ? 'Collapse' : 'Expand'}
                  </button>
                  <button
                    className="d365-trace-copy-btn"
                    onClick={() => handleCopy(selectedLog.exceptionDetails, `exception-${selectedLog.id}`)}
                    title="Copy exception details"
                  >
                    {copiedField === `exception-${selectedLog.id}` ? (
                      <span className="d365-copy-success">&checkmark; Copied</span>
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
              </div>
              <pre className={`d365-trace-pre d365-trace-pre-error ${exceptionExpanded ? 'd365-trace-pre--expanded' : 'd365-trace-pre--collapsed'}`}>
                {selectedLog.exceptionDetails}
              </pre>
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
              &#x1f9f9;
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
              &times;
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
            Filter {showFilters ? '\u25B2' : '\u25BC'}
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
                  <span>&ndash;</span>
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
              Filters apply to the currently loaded logs. Adjust &ldquo;Plugin trace log default limit&rdquo; in extension settings to fetch more.
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
