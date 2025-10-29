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

const PluginTraceLogViewer: React.FC<PluginTraceLogViewerProps> = ({
  data,
  onClose,
  onRefresh,
}) => {
  const [activeTab, setActiveTab] = useState<'recent' | 'exceptions'>('recent');
  const [selectedLog, setSelectedLog] = useState<PluginTraceLogRecord | null>(null);
  const [copiedField, setCopiedField] = useState<string>('');

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

  const filteredLogs = useMemo(() => {
    if (activeTab === 'exceptions') {
      return data.logs.filter(hasException);
    }
    return data.logs;
  }, [activeTab, data.logs]);

  // Auto-select first log if none selected
  useMemo(() => {
    if (filteredLogs.length > 0 && !selectedLog) {
      setSelectedLog(filteredLogs[0]);
    }
  }, [filteredLogs, selectedLog]);

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
            <button className="d365-trace-refresh" onClick={onRefresh} title="Refresh trace logs">
              ↻
            </button>
            <button className="d365-dialog-close" onClick={onClose} title="Close">
              ×
            </button>
          </div>
        </div>

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
