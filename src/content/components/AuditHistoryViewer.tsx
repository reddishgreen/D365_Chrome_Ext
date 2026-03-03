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

interface AuditGroup {
  key: string;
  action: string;
  changedBy: string;
  changedOn: string;
  records: AuditRecord[];
  colorIndex: number;
}

// Accent colours cycling per save transaction
const GROUP_COLORS = ['#0078d4', '#107c10', '#8764b8', '#d83b01', '#008272'];

const AuditHistoryViewer: React.FC<AuditHistoryViewerProps> = ({ data, onClose, onRefresh }) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [filterAction, setFilterAction] = useState<string>('all');
  const [filterUser, setFilterUser] = useState<string>('all');
  const [collapsedGroups, setCollapsedGroups] = useState<Set<string>>(new Set());

  const formatDate = (dateString: string): string => {
    try {
      return new Date(dateString).toLocaleString();
    } catch {
      return dateString;
    }
  };

  const getActionBadgeClass = (action: string): string => {
    switch (action.toLowerCase()) {
      case 'create': return 'audit-action-badge audit-action-create';
      case 'update': return 'audit-action-badge audit-action-update';
      case 'delete': return 'audit-action-badge audit-action-delete';
      default:       return 'audit-action-badge';
    }
  };

  const uniqueActions = Array.from(new Set(data.records.map(r => r.action)));
  const uniqueUsers  = Array.from(new Set(data.records.map(r => r.changedBy))).sort();

  const filteredRecords = data.records.filter(record => {
    const q = searchTerm.toLowerCase();
    const matchesSearch =
      record.fieldName.toLowerCase().includes(q) ||
      record.oldValue.toLowerCase().includes(q) ||
      record.newValue.toLowerCase().includes(q) ||
      record.changedBy.toLowerCase().includes(q);

    const matchesAction = filterAction === 'all' || record.action === filterAction;
    const matchesUser   = filterUser   === 'all' || record.changedBy === filterUser;

    return matchesSearch && matchesAction && matchesUser;
  });

  // Group filtered records by save transaction (same timestamp + user + action)
  const buildGroups = (records: AuditRecord[]): AuditGroup[] => {
    const seen = new Map<string, AuditRecord[]>();
    records.forEach(record => {
      const key = `${record.changedOn}|${record.changedBy}|${record.action}`;
      if (!seen.has(key)) seen.set(key, []);
      seen.get(key)!.push(record);
    });
    let i = 0;
    return Array.from(seen.entries()).map(([key, recs]) => ({
      key,
      action:    recs[0].action,
      changedBy: recs[0].changedBy,
      changedOn: recs[0].changedOn,
      records:   recs,
      colorIndex: i++ % GROUP_COLORS.length,
    }));
  };

  const groups = buildGroups(filteredRecords);

  const toggleGroup = (key: string) => {
    setCollapsedGroups(prev => {
      const next = new Set(prev);
      next.has(key) ? next.delete(key) : next.add(key);
      return next;
    });
  };

  const collapseAll = () => setCollapsedGroups(new Set(groups.map(g => g.key)));
  const expandAll   = () => setCollapsedGroups(new Set());

  const totalFieldChanges = filteredRecords.length;

  return (
    <div className="audit-history-overlay">
      <div className="audit-history-container">

        {/* ── Header ── */}
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
            <button onClick={onRefresh} className="audit-history-refresh-btn" title="Refresh audit history">↻</button>
            <button onClick={onClose}   className="audit-history-close-btn"   title="Close">✕</button>
          </div>
        </div>

        {data.error ? (
          <div className="audit-history-error"><p>{data.error}</p></div>
        ) : (
          <>
            {/* ── Filters ── */}
            <div className="audit-history-filters">
              <input
                type="text"
                placeholder="Search fields, values, or users..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="audit-history-search"
              />
              <select value={filterAction} onChange={(e) => setFilterAction(e.target.value)} className="audit-history-filter">
                <option value="all">All Actions</option>
                {uniqueActions.map(a => <option key={a} value={a}>{a}</option>)}
              </select>
              <select value={filterUser} onChange={(e) => setFilterUser(e.target.value)} className="audit-history-filter">
                <option value="all">All Users</option>
                {uniqueUsers.map(u => <option key={u} value={u}>{u}</option>)}
              </select>
              <div className="audit-history-count">
                {groups.length} save{groups.length !== 1 ? 's' : ''} &middot; {totalFieldChanges} field change{totalFieldChanges !== 1 ? 's' : ''}
              </div>
              <div className="audit-history-collapse-btns">
                <button onClick={expandAll}   className="audit-collapse-btn" title="Expand all">Expand all</button>
                <button onClick={collapseAll} className="audit-collapse-btn" title="Collapse all">Collapse all</button>
              </div>
            </div>

            {groups.length === 0 ? (
              <div className="audit-history-no-data">
                {data.records.length === 0
                  ? <p>No audit history found for this record.</p>
                  : <p>No records match your filters.</p>}
              </div>
            ) : (
              <div className="audit-history-table-container">
                <table className="audit-history-table">
                  <thead>
                    <tr>
                      <th className="col-field">Field</th>
                      <th className="col-old">Old Value</th>
                      <th className="col-arrow"></th>
                      <th className="col-new">New Value</th>
                    </tr>
                  </thead>

                  {groups.map(group => {
                    const color     = GROUP_COLORS[group.colorIndex];
                    const collapsed = collapsedGroups.has(group.key);
                    const count     = group.records.length;

                    return (
                      <tbody key={group.key} className="audit-group-body">
                        {/* Sticky group header row */}
                        <tr
                          className="audit-group-header"
                          style={{ borderLeft: `4px solid ${color}` }}
                          onClick={() => toggleGroup(group.key)}
                        >
                          <td colSpan={4}>
                            <div className="audit-group-header-content">
                              <span className={getActionBadgeClass(group.action)}>{group.action}</span>
                              <span className="audit-group-user">{group.changedBy}</span>
                              <span className="audit-group-sep">·</span>
                              <span className="audit-group-date">{formatDate(group.changedOn)}</span>
                              <span className="audit-group-sep">·</span>
                              <span className="audit-group-count"
                                style={{ background: color + '1a', color }}>
                                {count} field{count !== 1 ? 's' : ''} changed
                              </span>
                              <span className="audit-group-chevron">{collapsed ? '▶' : '▼'}</span>
                            </div>
                          </td>
                        </tr>

                        {/* Field change rows */}
                        {!collapsed && group.records.map((record, idx) => (
                          <tr
                            key={`${record.auditId}_${idx}`}
                            className="audit-field-row"
                            style={{ borderLeft: `4px solid ${color}40` }}
                          >
                            <td className="audit-field-name">{record.fieldName}</td>
                            <td className="audit-old-value">
                              {record.oldValue
                                ? <span>{record.oldValue}</span>
                                : <em className="audit-empty">(empty)</em>}
                            </td>
                            <td className="audit-arrow-cell">→</td>
                            <td className="audit-new-value">
                              {record.newValue
                                ? <span>{record.newValue}</span>
                                : <em className="audit-empty">(empty)</em>}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    );
                  })}
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
