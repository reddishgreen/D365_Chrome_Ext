import { useState, useEffect, useRef } from 'react';
import { CrmApi } from '../../query-builder/utils/api';
import { EntityMetadataComplete } from '../../query-builder/types';

interface MetadataCache {
  [key: string]: EntityMetadataComplete;
}

// Global cache shared across all instances
const metadataCache: MetadataCache = {};

interface UseEntityMetadataResult {
  entity: EntityMetadataComplete | null;
  attributes: EntityMetadataComplete['Attributes'];
  relationships: {
    oneToMany: EntityMetadataComplete['OneToManyRelationships'];
    manyToOne: EntityMetadataComplete['ManyToOneRelationships'];
    manyToMany: EntityMetadataComplete['ManyToManyRelationships'];
  };
  loading: boolean;
  error: string | null;
  refresh: () => void;
}

export function useEntityMetadata(
  entityLogicalName: string | null,
  orgUrl: string | null
): UseEntityMetadataResult {
  const [entity, setEntity] = useState<EntityMetadataComplete | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [refreshTrigger, setRefreshTrigger] = useState(0);

  const cacheKey = entityLogicalName && orgUrl ? `${orgUrl}_${entityLogicalName}` : null;

  const fetchMetadata = async (forceRefresh: boolean = false) => {
    if (!entityLogicalName || !orgUrl) {
      setEntity(null);
      setLoading(false);
      return;
    }

    // Check cache first (unless forcing refresh)
    if (!forceRefresh && cacheKey && metadataCache[cacheKey]) {
      const cached = metadataCache[cacheKey];
      setEntity(cached);
      setLoading(false);
      setError(null);
      return;
    }

    setLoading(true);
    setError(null);

    try {
      const api = new CrmApi(orgUrl);
      const metadata = await api.getEntityMetadata(entityLogicalName);

      // Cache the result
      if (cacheKey) {
        metadataCache[cacheKey] = metadata;
      }

      setEntity(metadata);
      setLoading(false);
    } catch (err: any) {
      const errorMessage = err instanceof Error ? err.message : 'Failed to fetch entity metadata';
      setError(errorMessage);
      setLoading(false);
      setEntity(null);
    }
  };

  useEffect(() => {
    fetchMetadata();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [entityLogicalName, orgUrl, refreshTrigger]);

  const refresh = () => {
    if (cacheKey) {
      delete metadataCache[cacheKey];
    }
    setRefreshTrigger(prev => prev + 1);
  };

  return {
    entity,
    attributes: entity?.Attributes || [],
    relationships: {
      oneToMany: entity?.OneToManyRelationships || [],
      manyToOne: entity?.ManyToOneRelationships || [],
      manyToMany: entity?.ManyToManyRelationships || []
    },
    loading,
    error,
    refresh
  };
}
