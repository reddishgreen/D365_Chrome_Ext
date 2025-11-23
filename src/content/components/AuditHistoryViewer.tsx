import React, { useState } from 'react';
import './AuditHistoryViewer.css';

export interface AuditRecord {
  auditId: string;
  action: string;
  fieldName: string;
  oldValue: string;
  newValue: string;
  changedBy: string;
  changedOn: string;
}

export interface AuditHistoryData {
  records: AuditRecord[];
  recordName?: string;
  entityName?: string;
  error?: string;
}

interface AuditHistoryViewerProps {
  data: AuditHistoryData;
  onClose: () => void;
  onRefresh: () => void;
}

const AuditHistoryViewer: React.FC<AuditHistoryViewerProps> = ({ data, onClose, onRefresh }) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [filterAction, setFilterAction] = useState<string>('all');

  const formatDate = (dateString: string): string => {
    try {
      const date = new Date(dateString);
      return date.toLocaleString();
    } catch {
      return dateString;
    }
  };

  const getActionBadgeClass = (action: string): string => {
    switch (action.toLowerCase()) {
      case 'create':
        return 'audit-action-badge audit-action-create';
      case 'update':
        return 'audit-action-badge audit-action-update';
      case 'delete':
        return 'audit-action-badge audit-action-delete';
      default:
        return 'audit-action-badge';
    }
  };

  const uniqueActions = Array.from(new Set(data.records.map(r => r.action)));

  const filteredRecords = data.records.filter(record => {
    const matchesSearch =
      record.fieldName.toLowerCase().includes(searchTerm.toLowerCase()) ||
      record.oldValue.toLowerCase().includes(searchTerm.toLowerCase()) ||
      record.newValue.toLowerCase().includes(searchTerm.toLowerCase()) ||
      record.changedBy.toLowerCase().includes(searchTerm.toLowerCase());

    const matchesAction = filterAction === 'all' || record.action === filterAction;

    return matchesSearch && matchesAction;
  });

  return (
    <div className="audit-history-overlay">
      <div className="audit-history-container">
        <div className="audit-history-header">
          <div className="audit-history-title-section">
            <h2>Audit History</h2>
            {data.recordName && (
              <div className="audit-history-subtitle">
                {data.recordName} ({data.entityName})
              </div>
            )}
          </div>
          <div className="audit-history-actions">
            <button onClick={onRefresh} className="audit-history-refresh-btn" title="Refresh audit history">
              ↻
            </button>
            <button onClick={onClose} className="audit-history-close-btn" title="Close">
              ✕
            </button>
          </div>
        </div>

        {data.error ? (
          <div className="audit-history-error">
            <p>{data.error}</p>
          </div>
        ) : (
          <>
            <div className="audit-history-filters">
              <input
                type="text"
                placeholder="Search fields, values, or users..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="audit-history-search"
              />
              <select
                value={filterAction}
                onChange={(e) => setFilterAction(e.target.value)}
                className="audit-history-filter"
              >
                <option value="all">All Actions</option>
                {uniqueActions.map(action => (
                  <option key={action} value={action}>{action}</option>
                ))}
              </select>
              <div className="audit-history-count">
                {filteredRecords.length} of {data.records.length} records
              </div>
            </div>

            {filteredRecords.length === 0 ? (
              <div className="audit-history-no-data">
                {data.records.length === 0 ? (
                  <p>No audit history found for this record.</p>
                ) : (
                  <p>No records match your filters.</p>
                )}
              </div>
            ) : (
              <div className="audit-history-table-container">
                <table className="audit-history-table">
                  <thead>
                    <tr>
                      <th>Action</th>
                      <th>Field</th>
                      <th>Old Value</th>
                      <th>New Value</th>
                      <th>Changed By</th>
                      <th>Changed On</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRecords.map((record) => (
                      <tr key={record.auditId}>
                        <td>
                          <span className={getActionBadgeClass(record.action)}>
                            {record.action}
                          </span>
                        </td>
                        <td className="audit-field-name">{record.fieldName}</td>
                        <td className="audit-value audit-old-value">{record.oldValue || '(empty)'}</td>
                        <td className="audit-value audit-new-value">{record.newValue || '(empty)'}</td>
                        <td className="audit-user">{record.changedBy}</td>
                        <td className="audit-date">{formatDate(record.changedOn)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
};

export default AuditHistoryViewer;
