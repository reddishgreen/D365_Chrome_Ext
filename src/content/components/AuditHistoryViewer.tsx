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
  onRollback?: (fieldName: string, oldValue: string, skipPlugins?: boolean) => Promise<boolean>;
  onRollbackGroup?: (changes: { fieldName: string; oldValue: string }[], skipPlugins?: boolean) => Promise<boolean>;
}

interface AuditGroup {
  key: string;
  action: string;
  changedBy: string;
  changedOn: string;
  records: AuditRecord[];
  colorIndex: number;
}

type RollbackState = 'idle' | 'rolling-back' | 'success' | 'error';

interface ConfirmInfo {
  type: 'field' | 'group';
  fieldName?: string;
  oldValue?: string;
  changes?: { fieldName: string; oldValue: string }[];
  groupLabel?: string;
}

const GROUP_COLORS = ['#0078d4', '#107c10', '#8764b8', '#d83b01', '#008272'];

const AuditHistoryViewer: React.FC<AuditHistoryViewerProps> = ({
  data,
  onClose,
  onRefresh,
  onRollback,
  onRollbackGroup,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [filterAction, setFilterAction] = useState<string>('all');
  const [filterUser, setFilterUser] = useState<string>('all');
  const [collapsedGroups, setCollapsedGroups] = useState<Set<string>>(new Set());
  const [rollbackStates, setRollbackStates] = useState<Map<string, RollbackState>>(new Map());
  const [confirmInfo, setConfirmInfo] = useState<ConfirmInfo | null>(null);
  const [confirmBusy, setConfirmBusy] = useState(false);
  const [skipPlugins, setSkipPlugins] = useState(false);

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

  const truncateValue = (val: string, max: number = 60): string => {
    if (!val || val.length <= max) return val;
    return val.slice(0, max) + '...';
  };

  // -- Rollback handlers --

  const requestFieldRollback = (record: AuditRecord) => {
    setConfirmInfo({
      type: 'field',
      fieldName: record.fieldName,
      oldValue: record.oldValue,
    });
  };

  const requestGroupRollback = (group: AuditGroup) => {
    const changes = group.records.map(r => ({
      fieldName: r.fieldName,
      oldValue: r.oldValue,
    }));
    setConfirmInfo({
      type: 'group',
      changes,
      groupLabel: `${group.changedBy} - ${formatDate(group.changedOn)}`,
    });
  };

  const executeRollback = async () => {
    if (!confirmInfo) return;
    setConfirmBusy(true);

    try {
      if (confirmInfo.type === 'field' && onRollback && confirmInfo.fieldName !== undefined) {
        const key = confirmInfo.fieldName;
        setRollbackStates(prev => new Map(prev).set(key, 'rolling-back'));
        const success = await onRollback(confirmInfo.fieldName, confirmInfo.oldValue || '', skipPlugins);
        setRollbackStates(prev => new Map(prev).set(key, success ? 'success' : 'error'));
        if (success) {
          setTimeout(() => {
            setRollbackStates(prev => { const n = new Map(prev); n.delete(key); return n; });
          }, 2000);
        }
      } else if (confirmInfo.type === 'group' && onRollbackGroup && confirmInfo.changes) {
        const groupKey = confirmInfo.groupLabel || 'group';
        setRollbackStates(prev => new Map(prev).set(groupKey, 'rolling-back'));
        const success = await onRollbackGroup(confirmInfo.changes, skipPlugins);
        setRollbackStates(prev => new Map(prev).set(groupKey, success ? 'success' : 'error'));
        if (success) {
          setTimeout(() => {
            setRollbackStates(prev => { const n = new Map(prev); n.delete(groupKey); return n; });
          }, 2000);
        }
      }
    } finally {
      setConfirmBusy(false);
      setConfirmInfo(null);
    }
  };

  const getRollbackBtnClass = (key: string): string => {
    const state = rollbackStates.get(key);
    if (state === 'rolling-back') return 'audit-rollback-btn rolling-back';
    if (state === 'success') return 'audit-rollback-btn success';
    if (state === 'error') return 'audit-rollback-btn error';
    return 'audit-rollback-btn';
  };

  const getRollbackBtnLabel = (key: string): string => {
    const state = rollbackStates.get(key);
    if (state === 'rolling-back') return 'Rolling back...';
    if (state === 'success') return 'Done';
    if (state === 'error') return 'Failed';
    return 'Rollback';
  };

  const getGroupRollbackBtnClass = (key: string): string => {
    const state = rollbackStates.get(key);
    if (state === 'rolling-back') return 'audit-rollback-group-btn rolling-back';
    if (state === 'success') return 'audit-rollback-group-btn success';
    if (state === 'error') return 'audit-rollback-group-btn error';
    return 'audit-rollback-group-btn';
  };

  const getGroupRollbackBtnLabel = (key: string): string => {
    const state = rollbackStates.get(key);
    if (state === 'rolling-back') return 'Rolling back...';
    if (state === 'success') return 'Done';
    if (state === 'error') return 'Failed';
    return 'Rollback all';
  };

  const canRollback = !!(onRollback);
  const canRollbackGroup = !!(onRollbackGroup);

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
            <button onClick={onRefresh} className="audit-history-refresh-btn" title="Refresh">↻</button>
            <button onClick={onClose} className="audit-history-close-btn" title="Close">✕</button>
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
                <button onClick={expandAll} className="audit-collapse-btn">Expand all</button>
                <button onClick={collapseAll} className="audit-collapse-btn">Collapse all</button>
              </div>
            </div>

            {groups.length === 0 ? (
              <div className="audit-history-no-data">
                {data.records.length === 0
                  ? <p>No audit history found for this record.</p>
                  : <p>No records match your filters.</p>}
              </div>
            ) : (
              <div className="audit-history-content">
                {groups.map(group => {
                  const color = GROUP_COLORS[group.colorIndex];
                  const collapsed = collapsedGroups.has(group.key);
                  const count = group.records.length;
                  const groupRollbackKey = `${group.changedBy} - ${formatDate(group.changedOn)}`;

                  return (
                    <div key={group.key} className="audit-group-card">
                      {/* Group header */}
                      <div
                        className="audit-group-header"
                        style={{ borderLeftColor: color }}
                        onClick={() => toggleGroup(group.key)}
                      >
                        <div className="audit-group-header-left">
                          <span className={getActionBadgeClass(group.action)}>{group.action}</span>
                          <span className="audit-group-user">{group.changedBy}</span>
                          <span className="audit-group-sep">&middot;</span>
                          <span className="audit-group-date">{formatDate(group.changedOn)}</span>
                          <span className="audit-group-sep">&middot;</span>
                          <span
                            className="audit-group-count"
                            style={{ background: color + '15', color }}
                          >
                            {count} field{count !== 1 ? 's' : ''}
                          </span>
                        </div>
                        {canRollbackGroup && group.action.toLowerCase() === 'update' && !collapsed && count > 1 && (
                          <button
                            className={getGroupRollbackBtnClass(groupRollbackKey)}
                            onClick={(e) => { e.stopPropagation(); requestGroupRollback(group); }}
                            title="Rollback all fields in this save"
                          >
                            ↶ {getGroupRollbackBtnLabel(groupRollbackKey)}
                          </button>
                        )}
                        <span className={`audit-group-chevron${collapsed ? '' : ' expanded'}`}>▶</span>
                      </div>

                      {/* Field change rows */}
                      {!collapsed && (
                        <div className="audit-field-list">
                          {group.records.map((record, idx) => (
                            <div
                              key={`${record.auditId}_${idx}`}
                              className="audit-field-row"
                            >
                              <div className="audit-field-name-col">
                                <span className="audit-field-name">{record.fieldName}</span>
                              </div>

                              <div className="audit-values-section">
                                <div className="audit-value-block">
                                  <span className="audit-value-label was">Was</span>
                                  <span className="audit-value-content old">
                                    {record.oldValue
                                      ? record.oldValue
                                      : <span className="audit-empty-value">empty</span>}
                                  </span>
                                </div>

                                <div className="audit-value-arrow">→</div>

                                <div className="audit-value-block">
                                  <span className="audit-value-label now">Now</span>
                                  <span className="audit-value-content new">
                                    {record.newValue
                                      ? record.newValue
                                      : <span className="audit-empty-value">empty</span>}
                                  </span>
                                </div>
                              </div>

                              {canRollback && group.action.toLowerCase() === 'update' && (
                                <div className="audit-rollback-col">
                                  <button
                                    className={getRollbackBtnClass(record.fieldName)}
                                    onClick={() => requestFieldRollback(record)}
                                    title={`Revert ${record.fieldName} to previous value`}
                                  >
                                    ↶ {getRollbackBtnLabel(record.fieldName)}
                                  </button>
                                </div>
                              )}
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </>
        )}
      </div>

      {/* ── Confirmation dialog ── */}
      {confirmInfo && (
        <div className="audit-confirm-overlay" onClick={() => !confirmBusy && setConfirmInfo(null)}>
          <div className="audit-confirm-dialog" onClick={(e) => e.stopPropagation()}>
            <div className="audit-confirm-header">
              <span>⚠</span>
              <h3>Confirm Rollback</h3>
            </div>
            <div className="audit-confirm-body">
              {confirmInfo.type === 'field' ? (
                <>
                  <p>
                    Are you sure you want to revert <strong>{confirmInfo.fieldName}</strong> to its previous value?
                  </p>
                  <div className="audit-confirm-detail">
                    <div className="audit-confirm-field">
                      <span className="audit-confirm-field-label">Field</span>
                      <span className="audit-confirm-field-value">{confirmInfo.fieldName}</span>
                    </div>
                    <div className="audit-confirm-field">
                      <span className="audit-confirm-field-label">Revert to</span>
                      <span className="audit-confirm-field-value">
                        {confirmInfo.oldValue || '(empty)'}
                      </span>
                    </div>
                  </div>
                </>
              ) : (
                <>
                  <p>
                    Are you sure you want to rollback <strong>{confirmInfo.changes?.length} field{(confirmInfo.changes?.length || 0) > 1 ? 's' : ''}</strong> from this save?
                  </p>
                  <div className="audit-confirm-detail">
                    {confirmInfo.changes?.map((c, i) => (
                      <div key={i} className="audit-confirm-field">
                        <span className="audit-confirm-field-label">{c.fieldName}</span>
                        <span className="audit-confirm-field-value">
                          {truncateValue(c.oldValue) || '(empty)'}
                        </span>
                      </div>
                    ))}
                  </div>
                </>
              )}
              <label className="audit-confirm-skip-plugins">
                <input
                  type="checkbox"
                  checked={skipPlugins}
                  onChange={(e) => setSkipPlugins(e.target.checked)}
                />
                Skip plugin execution
              </label>
            </div>
            <div className="audit-confirm-actions">
              <button
                className="audit-confirm-cancel"
                onClick={() => setConfirmInfo(null)}
                disabled={confirmBusy}
              >
                Cancel
              </button>
              <button
                className="audit-confirm-proceed"
                onClick={executeRollback}
                disabled={confirmBusy}
              >
                {confirmBusy ? 'Rolling back...' : 'Rollback'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default AuditHistoryViewer;
