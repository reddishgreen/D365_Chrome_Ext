import React, { useEffect, useMemo, useState } from 'react';
import './PluginStepsViewer.css';
import CloseIcon from './CloseIcon';

export interface PluginStepImage {
  id: string;
  name: string;
  entityAlias: string;
  imageType: number;        // 0=Pre, 1=Post, 2=Both
  imageTypeLabel: string;
  attributes: string;       // comma-separated logical names; empty = all
  messagePropertyName: string;
}

export interface PluginStepRecord {
  id: string;
  name: string;
  description?: string;
  message: string;
  primaryEntity?: string;
  stage: number;
  stageLabel: string;
  mode: number;
  modeLabel: string;
  rank: number;
  isEnabled: boolean;
  filteringAttributes?: string;
  pluginTypeName?: string;
  assemblyName?: string;
  isCustom?: boolean;
  images?: PluginStepImage[];
}

export interface PluginStepsData {
  entityName: string;
  steps: PluginStepRecord[];
  error?: string;
}

interface PluginStepsViewerProps {
  data: PluginStepsData | null;
  onClose: () => void;
  onRefresh: () => void;
  onToggle: (step: PluginStepRecord) => Promise<boolean>;
  onUpdateImage: (
    stepId: string,
    image: PluginStepImage,
    patch: Partial<Pick<PluginStepImage, 'name' | 'entityAlias' | 'attributes' | 'imageType'>>
  ) => Promise<boolean>;
  onDeleteImage: (stepId: string, imageId: string) => Promise<boolean>;
  onCreateImage: (
    stepId: string,
    args: { name: string; entityAlias: string; attributes: string; imageType: number }
  ) => Promise<boolean>;
  isLoading?: boolean;
}

const STAGE_COLORS: Record<number, string> = {
  10: '#8764b8',
  20: '#0078d4',
  40: '#107c10',
};

interface ImageEditState {
  name: string;
  entityAlias: string;
  attributes: string;
  imageType: number;
}

const PluginStepsViewer: React.FC<PluginStepsViewerProps> = ({
  data,
  onClose,
  onRefresh,
  onToggle,
  onUpdateImage,
  onDeleteImage,
  onCreateImage,
  isLoading,
}) => {
  const [search, setSearch] = useState('');
  const [messageFilter, setMessageFilter] = useState<string>('all');
  const [stateFilter, setStateFilter] = useState<string>('all');
  const [customOnly, setCustomOnly] = useState(true); // default on per spec
  const [busy, setBusy] = useState<Set<string>>(new Set());
  const [expandedSteps, setExpandedSteps] = useState<Set<string>>(new Set());
  const [editingImage, setEditingImage] = useState<string | null>(null);
  const [editState, setEditState] = useState<ImageEditState | null>(null);
  const [creatingFor, setCreatingFor] = useState<string | null>(null);
  const [createState, setCreateState] = useState<ImageEditState>({
    name: 'PreImage',
    entityAlias: 'PreImage',
    attributes: '',
    imageType: 0,
  });

  const steps = data?.steps || [];

  // If data refreshes, drop any open edit state that no longer references a known image.
  useEffect(() => {
    if (!editingImage) return;
    const stillExists = steps.some((s) => s.images?.some((i) => i.id === editingImage));
    if (!stillExists) {
      setEditingImage(null);
      setEditState(null);
    }
  }, [steps, editingImage]);

  const messages = useMemo(() => {
    return Array.from(new Set(steps.map((s) => s.message))).sort();
  }, [steps]);

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    return steps.filter((s) => {
      if (customOnly && !s.isCustom) return false;
      if (messageFilter !== 'all' && s.message !== messageFilter) return false;
      if (stateFilter === 'enabled' && !s.isEnabled) return false;
      if (stateFilter === 'disabled' && s.isEnabled) return false;
      if (!q) return true;
      return (
        s.name.toLowerCase().includes(q) ||
        (s.description || '').toLowerCase().includes(q) ||
        (s.pluginTypeName || '').toLowerCase().includes(q) ||
        (s.assemblyName || '').toLowerCase().includes(q) ||
        (s.filteringAttributes || '').toLowerCase().includes(q)
      );
    });
  }, [steps, search, messageFilter, stateFilter, customOnly]);

  const grouped = useMemo(() => {
    const map = new Map<string, PluginStepRecord[]>();
    filtered.forEach((s) => {
      const key = `${s.message}__${s.stage}`;
      if (!map.has(key)) map.set(key, []);
      map.get(key)!.push(s);
    });
    return Array.from(map.entries())
      .sort((a, b) => {
        const [aMsg, aStage] = a[0].split('__');
        const [bMsg, bStage] = b[0].split('__');
        if (Number(aStage) !== Number(bStage)) return Number(aStage) - Number(bStage);
        return aMsg.localeCompare(bMsg);
      })
      .map(([key, items]) => ({
        key,
        message: items[0].message,
        stage: items[0].stage,
        stageLabel: items[0].stageLabel,
        items: items.sort((x, y) => x.rank - y.rank),
      }));
  }, [filtered]);

  const toggleExpanded = (stepId: string) => {
    setExpandedSteps((prev) => {
      const next = new Set(prev);
      if (next.has(stepId)) next.delete(stepId);
      else next.add(stepId);
      return next;
    });
  };

  const handleToggle = async (s: PluginStepRecord) => {
    const key = `step:${s.id}`;
    if (busy.has(key)) return;
    setBusy((prev) => new Set(prev).add(key));
    try {
      await onToggle(s);
    } finally {
      setBusy((prev) => {
        const next = new Set(prev);
        next.delete(key);
        return next;
      });
    }
  };

  const beginEditImage = (img: PluginStepImage) => {
    setEditingImage(img.id);
    setEditState({
      name: img.name,
      entityAlias: img.entityAlias,
      attributes: img.attributes,
      imageType: img.imageType,
    });
  };

  const cancelEditImage = () => {
    setEditingImage(null);
    setEditState(null);
  };

  const saveImage = async (stepId: string, img: PluginStepImage) => {
    if (!editState) return;
    const key = `img:${img.id}`;
    setBusy((prev) => new Set(prev).add(key));
    try {
      const ok = await onUpdateImage(stepId, img, {
        name: editState.name,
        entityAlias: editState.entityAlias,
        attributes: editState.attributes,
        imageType: editState.imageType,
      });
      if (ok) {
        cancelEditImage();
      }
    } finally {
      setBusy((prev) => {
        const next = new Set(prev);
        next.delete(key);
        return next;
      });
    }
  };

  const deleteImage = async (stepId: string, img: PluginStepImage) => {
    if (!window.confirm(`Delete the "${img.name || img.entityAlias}" image?`)) return;
    const key = `img:${img.id}`;
    setBusy((prev) => new Set(prev).add(key));
    try {
      await onDeleteImage(stepId, img.id);
    } finally {
      setBusy((prev) => {
        const next = new Set(prev);
        next.delete(key);
        return next;
      });
    }
  };

  const beginCreateImage = (stepId: string) => {
    setCreatingFor(stepId);
    setCreateState({
      name: 'PreImage',
      entityAlias: 'PreImage',
      attributes: '',
      imageType: 0,
    });
  };

  const cancelCreateImage = () => {
    setCreatingFor(null);
  };

  const saveCreateImage = async (stepId: string) => {
    const key = `create:${stepId}`;
    setBusy((prev) => new Set(prev).add(key));
    try {
      const ok = await onCreateImage(stepId, { ...createState });
      if (ok) cancelCreateImage();
    } finally {
      setBusy((prev) => {
        const next = new Set(prev);
        next.delete(key);
        return next;
      });
    }
  };

  return (
    <div className="ps-overlay">
      <div className="ps-container">
        <div className="ps-header">
          <div>
            <h2>Plugin Steps</h2>
            {data?.entityName && <div className="ps-subtitle">{data.entityName}</div>}
          </div>
          <div className="ps-actions">
            {isLoading && <span className="ps-loading-pill">Loading...</span>}
            <button onClick={onRefresh} className="ps-icon-btn" title="Refresh" disabled={isLoading}>
              ↻
            </button>
            <button onClick={onClose} className="ps-icon-btn" title="Close" aria-label="Close">
              <CloseIcon />
            </button>
          </div>
        </div>

        {data?.error ? (
          <div className="ps-error">{data.error}</div>
        ) : (
          <>
            <div className="ps-filters">
              <input
                type="text"
                className="ps-search"
                placeholder="Search by name, type, assembly, filtering attrs..."
                value={search}
                onChange={(e) => setSearch(e.target.value)}
              />
              <select
                className="ps-filter"
                value={messageFilter}
                onChange={(e) => setMessageFilter(e.target.value)}
              >
                <option value="all">All messages</option>
                {messages.map((m) => (
                  <option key={m} value={m}>
                    {m}
                  </option>
                ))}
              </select>
              <select
                className="ps-filter"
                value={stateFilter}
                onChange={(e) => setStateFilter(e.target.value)}
              >
                <option value="all">All states</option>
                <option value="enabled">Enabled only</option>
                <option value="disabled">Disabled only</option>
              </select>
              <label className="ps-checkbox">
                <input
                  type="checkbox"
                  checked={customOnly}
                  onChange={(e) => setCustomOnly(e.target.checked)}
                />
                <span>Custom only</span>
              </label>
              <div className="ps-count">{filtered.length} step{filtered.length === 1 ? '' : 's'}</div>
            </div>

            {steps.length === 0 && !isLoading ? (
              <div className="ps-empty">No registered plugin steps for this entity.</div>
            ) : grouped.length === 0 ? (
              <div className="ps-empty">
                {customOnly && steps.length > 0
                  ? 'No custom plugin steps match your filters. Untick "Custom only" to see all.'
                  : 'No steps match your filters.'}
              </div>
            ) : (
              <div className="ps-content">
                {grouped.map((g) => {
                  const color = STAGE_COLORS[g.stage] || '#5c6c7c';
                  return (
                    <div key={g.key} className="ps-group">
                      <div className="ps-group-header" style={{ borderLeftColor: color }}>
                        <span className="ps-group-stage" style={{ color }}>
                          {g.stageLabel}
                        </span>
                        <span className="ps-group-msg">{g.message}</span>
                        <span className="ps-group-count">{g.items.length}</span>
                      </div>
                      <div className="ps-list">
                        {g.items.map((s) => {
                          const stepBusy = busy.has(`step:${s.id}`);
                          const isExpanded = expandedSteps.has(s.id);
                          const imageCount = s.images?.length || 0;
                          return (
                            <div key={s.id} className="ps-row">
                              <div className="ps-row-main">
                                <div className="ps-row-name">
                                  <span className="ps-rank">#{s.rank}</span>
                                  {s.name}
                                  {!s.isCustom && <span className="ps-pill ps-pill-system">System</span>}
                                </div>
                                <div className="ps-row-meta">
                                  <span className={`ps-pill ps-pill-${s.isEnabled ? 'on' : 'off'}`}>
                                    {s.isEnabled ? 'Enabled' : 'Disabled'}
                                  </span>
                                  <span className="ps-pill ps-pill-mode">{s.modeLabel}</span>
                                  {s.pluginTypeName && (
                                    <span className="ps-meta-text" title={s.pluginTypeName}>
                                      {s.pluginTypeName}
                                    </span>
                                  )}
                                  {s.assemblyName && (
                                    <span className="ps-meta-text ps-meta-dim">{s.assemblyName}</span>
                                  )}
                                </div>
                                {s.filteringAttributes && (
                                  <div className="ps-row-attrs">
                                    Filter attrs: <code>{s.filteringAttributes}</code>
                                  </div>
                                )}
                                {s.description && <div className="ps-row-desc">{s.description}</div>}

                                <div className="ps-images-bar">
                                  <button
                                    type="button"
                                    className="ps-images-toggle"
                                    onClick={() => toggleExpanded(s.id)}
                                  >
                                    {isExpanded ? '▾' : '▸'} {imageCount} image{imageCount === 1 ? '' : 's'}
                                  </button>
                                  {isExpanded && (
                                    <button
                                      type="button"
                                      className="ps-images-add"
                                      onClick={() => beginCreateImage(s.id)}
                                      disabled={creatingFor === s.id}
                                    >
                                      + Add image
                                    </button>
                                  )}
                                </div>

                                {isExpanded && (
                                  <div className="ps-images">
                                    {(s.images || []).length === 0 && creatingFor !== s.id && (
                                      <div className="ps-images-empty">
                                        No images registered. Pre/Post images give plugins a snapshot of the record.
                                      </div>
                                    )}
                                    {(s.images || []).map((img) => {
                                      const isEditing = editingImage === img.id;
                                      const imgBusy = busy.has(`img:${img.id}`);
                                      return (
                                        <div key={img.id} className={`ps-image ${isEditing ? 'ps-image-editing' : ''}`}>
                                          {!isEditing ? (
                                            <>
                                              <div className="ps-image-line">
                                                <span className={`ps-image-type-pill ps-image-type-${img.imageType}`}>
                                                  {img.imageTypeLabel}
                                                </span>
                                                <span className="ps-image-name">{img.name || '(no name)'}</span>
                                                <span className="ps-image-alias">
                                                  alias <code>{img.entityAlias || '-'}</code>
                                                </span>
                                              </div>
                                              <div className="ps-image-attrs">
                                                Attributes: {img.attributes ? <code>{img.attributes}</code> : <em>all attributes</em>}
                                              </div>
                                              <div className="ps-image-actions">
                                                <button
                                                  type="button"
                                                  className="ps-link-btn"
                                                  onClick={() => beginEditImage(img)}
                                                  disabled={imgBusy}
                                                >
                                                  Edit
                                                </button>
                                                <button
                                                  type="button"
                                                  className="ps-link-btn ps-link-danger"
                                                  onClick={() => deleteImage(s.id, img)}
                                                  disabled={imgBusy}
                                                >
                                                  Delete
                                                </button>
                                              </div>
                                            </>
                                          ) : (
                                            <ImageEditForm
                                              state={editState!}
                                              onChange={setEditState}
                                              onCancel={cancelEditImage}
                                              onSave={() => saveImage(s.id, img)}
                                              busy={imgBusy}
                                            />
                                          )}
                                        </div>
                                      );
                                    })}
                                    {creatingFor === s.id && (
                                      <div className="ps-image ps-image-editing">
                                        <ImageEditForm
                                          state={createState}
                                          onChange={setCreateState as any}
                                          onCancel={cancelCreateImage}
                                          onSave={() => saveCreateImage(s.id)}
                                          busy={busy.has(`create:${s.id}`)}
                                          createMode
                                        />
                                      </div>
                                    )}
                                  </div>
                                )}
                              </div>
                              <div className="ps-row-actions">
                                <button
                                  className={`ps-toggle-btn ${s.isEnabled ? 'ps-toggle-disable' : 'ps-toggle-enable'}`}
                                  onClick={() => handleToggle(s)}
                                  disabled={stepBusy}
                                  title={s.isEnabled ? 'Disable step' : 'Enable step'}
                                >
                                  {stepBusy ? '...' : s.isEnabled ? 'Disable' : 'Enable'}
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

interface ImageEditFormProps {
  state: ImageEditState;
  onChange: (next: ImageEditState) => void;
  onCancel: () => void;
  onSave: () => void;
  busy: boolean;
  createMode?: boolean;
}

const ImageEditForm: React.FC<ImageEditFormProps> = ({ state, onChange, onCancel, onSave, busy, createMode }) => {
  const set = (patch: Partial<ImageEditState>) => onChange({ ...state, ...patch });
  return (
    <div className="ps-image-form">
      <div className="ps-image-form-grid">
        <label>
          <span>Name</span>
          <input
            type="text"
            value={state.name}
            onChange={(e) => set({ name: e.target.value })}
            placeholder="PreImage"
          />
        </label>
        <label>
          <span>Entity alias</span>
          <input
            type="text"
            value={state.entityAlias}
            onChange={(e) => set({ entityAlias: e.target.value })}
            placeholder="PreImage"
          />
        </label>
        <label>
          <span>Image type</span>
          <select value={state.imageType} onChange={(e) => set({ imageType: Number(e.target.value) })}>
            <option value={0}>Pre</option>
            <option value={1}>Post</option>
            <option value={2}>Both</option>
          </select>
        </label>
        <label className="ps-image-form-attrs">
          <span>Attributes (comma-separated; empty = all)</span>
          <input
            type="text"
            value={state.attributes}
            onChange={(e) => set({ attributes: e.target.value })}
            placeholder="e.g. name,statecode,_primarycontactid_value"
          />
        </label>
      </div>
      <div className="ps-image-form-actions">
        <button type="button" className="ps-link-btn" onClick={onCancel} disabled={busy}>
          Cancel
        </button>
        <button type="button" className="ps-save-btn" onClick={onSave} disabled={busy}>
          {busy ? 'Saving…' : createMode ? 'Create' : 'Save'}
        </button>
      </div>
    </div>
  );
};

export default PluginStepsViewer;
