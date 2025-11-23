import React, { useState, useEffect, useMemo } from 'react';
import { CrmApi } from '../utils/api';
import { ViewMetadata } from '../types';
import EnhancedSelect from './EnhancedSelect';

interface ViewSelectorProps {
  api: CrmApi | null;
  entityLogicalName: string;
  onViewSelected: (view: ViewMetadata) => void;
}

const ViewSelector: React.FC<ViewSelectorProps> = ({ api, entityLogicalName, onViewSelected }) => {
  const [views, setViews] = useState<ViewMetadata[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [selectedViewId, setSelectedViewId] = useState<string>('');

  useEffect(() => {
    if (api && entityLogicalName) {
      setLoading(true);
      setError(null);
      api.getViews(entityLogicalName)
        .then(data => {
          setViews(data);
          setLoading(false);
        })
        .catch(err => {
          console.error('Failed to load views', err);
          setError(err.message || 'Failed to load views');
          setLoading(false);
        });
    } else {
        setViews([]);
    }
  }, [api, entityLogicalName]);

  const handleChange = (viewId: string) => {
    setSelectedViewId(viewId);
    const view = views.find(v => v.id === viewId);
    if (view) {
      onViewSelected(view);
    }
  };

  if (!api || !entityLogicalName) return null;

  const systemViews = views.filter(v => !v.isUserQuery);
  const userViews = views.filter(v => v.isUserQuery);

  // Prepare options with groups
  const options = useMemo(() => {
    const opts: Array<{ value: string; label: string; group?: string }> = [
      { value: '', label: 'Select a View to Populate Columns' }
    ];

    userViews.forEach(v => {
      opts.push({
        value: v.id,
        label: v.name,
        group: 'My Views'
      });
    });

    systemViews.forEach(v => {
      opts.push({
        value: v.id,
        label: v.name,
        group: 'System Views'
      });
    });

    return opts;
  }, [views, userViews, systemViews]);

  return (
    <div className="view-selector">
      <label style={{ display: 'block', fontSize: '14px', marginBottom: '6px', fontWeight: 500 }}>
        Select View (Optional)
      </label>
      <EnhancedSelect
        options={options}
        value={selectedViewId}
        onChange={handleChange}
        disabled={loading}
        searchable={true}
        placeholder={loading ? 'Loading views...' : 'Select a View to Populate Columns'}
        size="medium"
      />
      {error && (
        <div style={{ fontSize: '12px', color: '#a80000', marginTop: '4px', padding: '4px 8px', background: '#fdf3f4', borderRadius: '4px' }}>
          {error}
        </div>
      )}
    </div>
  );
};

export default ViewSelector;

