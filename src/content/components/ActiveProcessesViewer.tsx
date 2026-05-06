import React, { useMemo, useState } from 'react';
import './ActiveProcessesViewer.css';
import CloseIcon from './CloseIcon';

export interface ProcessRecord {
  id: string;
  name: string;
  category: number;
  categoryLabel: string;
  mode: number;
  modeLabel: string;
  statecode: number;
  statuscode: number;
  isActivated: boolean;
  triggers: string[];
  ownerName?: string;
  modifiedOn?: string;
  description?: string;
}

export interface ActiveProcessesData {
  entityName: string;
  processes: ProcessRecord[];
  error?: string;
}

interface ActiveProcessesViewerProps {
  data: ActiveProcessesData | null;
  onClose: () => void;
  onRefresh: () => void;
  onToggle: (process: ProcessRecord) => Promise<boolean>;
  isLoading?: boolean;
}

const CATEGORY_COLORS: Record<number, string> = {
  0: '#0078d4', // Workflow (Classic)
  2: '#107c10', // Business Rule
  3: '#8764b8', // Action
  4: '#d83b01', // Business Process Flow
  5: '#008272', // Modern Flow (Power Automate)
};

const ActiveProcessesViewer: React.FC<ActiveProcessesViewerProps> = ({
  data,
  onClose,
  onRefresh,
  onToggle,
  isLoading,
}) => {
  const [search, setSearch] = useState('');
  const [categoryFilter, setCategoryFilter] = useState<string>('all');
  const [stateFilter, setStateFilter] = useState<string>('all');
  const [busy, setBusy] = useState<Set<string>>(new Set());

  const processes = data?.processes || [];

  const categories = useMemo(() => {
    const map = new Map<number, string>();
    processes.forEach((p) => map.set(p.category, p.categoryLabel));
    return Array.from(map.entries()).sort((a, b) => a[0] - b[0]);
  }, [processes]);

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    return processes.filter((p) => {
      if (categoryFilter !== 'all' && String(p.category) !== categoryFilter) return false;
      if (stateFilter === 'active' && !p.isActivated) return false;
      if (stateFilter === 'inactive' && p.isActivated) return false;
      if (!q) return true;
      return (
        p.name.toLowerCase().includes(q) ||
        (p.description || '').toLowerCase().includes(q) ||
        (p.ownerName || '').toLowerCase().includes(q)
      );
    });
  }, [processes, search, categoryFilter, stateFilter]);

  const grouped = useMemo(() => {
    const map = new Map<number, ProcessRecord[]>();
    filtered.forEach((p) => {
      if (!map.has(p.category)) map.set(p.category, []);
      map.get(p.category)!.push(p);
    });
    return Array.from(map.entries()).sort((a, b) => a[0] - b[0]);
  }, [filtered]);

  const handleToggle = async (p: ProcessRecord) => {
    if (busy.has(p.id)) return;
    setBusy((prev) => new Set(prev).add(p.id));
    try {
      await onToggle(p);
    } finally {
      setBusy((prev) => {
        const next = new Set(prev);
        next.delete(p.id);
        return next;
      });
    }
  };

  const formatDate = (s?: string) => {
    if (!s) return '';
    try {
      return new Date(s).toLocaleString();
    } catch {
      return s;
    }
  };

  return (
    <div className="ap-overlay">
      <div className="ap-container">
        <div className="ap-header">
          <div>
            <h2>Active Processes</h2>
            {data?.entityName && <div className="ap-subtitle">{data.entityName}</div>}
          </div>
          <div className="ap-actions">
            {isLoading && <span className="ap-loading-pill">Loading...</span>}
            <button onClick={onRefresh} className="ap-icon-btn" title="Refresh" disabled={isLoading}>
              ↻
            </button>
            <button onClick={onClose} className="ap-icon-btn" title="Close" aria-label="Close">
              <CloseIcon />
            </button>
          </div>
        </div>

        {data?.error ? (
          <div className="ap-error">{data.error}</div>
        ) : (
          <>
            <div className="ap-filters">
              <input
                type="text"
                className="ap-search"
                placeholder="Search by name, description, or owner..."
                value={search}
                onChange={(e) => setSearch(e.target.value)}
              />
              <select
                className="ap-filter"
                value={categoryFilter}
                onChange={(e) => setCategoryFilter(e.target.value)}
              >
                <option value="all">All categories</option>
                {categories.map(([key, label]) => (
                  <option key={key} value={String(key)}>
                    {label}
                  </option>
                ))}
              </select>
              <select
                className="ap-filter"
                value={stateFilter}
                onChange={(e) => setStateFilter(e.target.value)}
              >
                <option value="all">All states</option>
                <option value="active">Activated only</option>
                <option value="inactive">Draft only</option>
              </select>
              <div className="ap-count">{filtered.length} process{filtered.length === 1 ? '' : 'es'}</div>
            </div>

            {processes.length === 0 && !isLoading ? (
              <div className="ap-empty">No processes found for this entity.</div>
            ) : grouped.length === 0 ? (
              <div className="ap-empty">No processes match your filters.</div>
            ) : (
              <div className="ap-content">
                {grouped.map(([cat, items]) => {
                  const color = CATEGORY_COLORS[cat] || '#5c6c7c';
                  return (
                    <div key={cat} className="ap-group">
                      <div className="ap-group-header" style={{ borderLeftColor: color }}>
                        <span className="ap-group-label" style={{ color }}>
                          {items[0].categoryLabel}
                        </span>
                        <span className="ap-group-count">{items.length}</span>
                      </div>
                      <div className="ap-list">
                        {items.map((p) => {
                          const isBusy = busy.has(p.id);
                          return (
                            <div key={p.id} className="ap-row">
                              <div className="ap-row-main">
                                <div className="ap-row-name">{p.name}</div>
                                <div className="ap-row-meta">
                                  <span className={`ap-pill ap-pill-${p.isActivated ? 'on' : 'off'}`}>
                                    {p.isActivated ? 'Activated' : 'Draft'}
                                  </span>
                                  <span className="ap-pill ap-pill-mode">{p.modeLabel}</span>
                                  {p.triggers.map((t) => (
                                    <span key={t} className="ap-pill ap-pill-trigger">
                                      {t}
                                    </span>
                                  ))}
                                  {p.ownerName && <span className="ap-meta-text">Owner: {p.ownerName}</span>}
                                  {p.modifiedOn && (
                                    <span className="ap-meta-text">Modified {formatDate(p.modifiedOn)}</span>
                                  )}
                                </div>
                                {p.description && <div className="ap-row-desc">{p.description}</div>}
                              </div>
                              <div className="ap-row-actions">
                                <button
                                  className={`ap-toggle-btn ${p.isActivated ? 'ap-toggle-deactivate' : 'ap-toggle-activate'}`}
                                  onClick={() => handleToggle(p)}
                                  disabled={isBusy}
                                  title={p.isActivated ? 'Deactivate process' : 'Activate process'}
                                >
                                  {isBusy ? '...' : p.isActivated ? 'Deactivate' : 'Activate'}
                                </button>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
};

export default ActiveProcessesViewer;
