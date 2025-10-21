import React, {
  useState,
  useEffect,
  useCallback,
  useRef,
  KeyboardEvent,
  ChangeEvent,
} from 'react';
import { LookupEntityMetadata, LookupSelection } from './lookupTypes';

const GUID_REGEX =
  /^[{(]?[0-9a-fA-F]{8}[-)]?[0-9a-fA-F]{4}[-)]?[0-9a-fA-F]{4}[-)]?[0-9a-fA-F]{4}[-)]?[0-9a-fA-F]{12}[)}]?$/;

const normalizeGuid = (value: string): string => {
  return value.replace(/[{}()]/g, '').toLowerCase();
};

const escapeODataString = (value: string): string => value.replace(/'/g, "''");

interface LookupFieldEditorProps {
  apiBaseUrl: string;
  attributeName: string;
  currentId?: string | null;
  currentName?: string | null;
  currentLogicalName?: string | null;
  loadTargets?: () => Promise<string[]>;
  getEntityMetadata: (logicalName: string) => Promise<LookupEntityMetadata | null>;
  onSelectionChange: (selection: LookupSelection | null) => void;
  disabled?: boolean;
}

interface LookupSearchResult extends LookupSelection {}

const LookupFieldEditor: React.FC<LookupFieldEditorProps> = ({
  apiBaseUrl,
  attributeName,
  currentId,
  currentName,
  currentLogicalName,
  loadTargets,
  getEntityMetadata,
  onSelectionChange,
  disabled = false,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [results, setResults] = useState<LookupSearchResult[]>([]);
  const [targets, setTargets] = useState<string[]>([]);
  const [selectedTarget, setSelectedTarget] = useState<string | undefined>(
    currentLogicalName || undefined
  );
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [showResults, setShowResults] = useState(false);
  const [selectedName, setSelectedName] = useState<string>(currentName ?? '');
  const [selectedId, setSelectedId] = useState<string | null>(currentId ?? null);
  const [activeIndex, setActiveIndex] = useState<number>(-1);
  const [isInputFocused, setIsInputFocused] = useState(false);
  const isMounted = useRef(true);
  const inputFocusRef = useRef(false);
  const resultsRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    return () => {
      isMounted.current = false;
    };
  }, []);

  useEffect(() => {
    setSelectedName(currentName ?? '');
  }, [currentName]);

  useEffect(() => {
    setSelectedId(currentId ?? null);
  }, [currentId]);

  useEffect(() => {
    setError(null);
  }, [searchTerm, selectedTarget]);

  useEffect(() => {
    if (results.length > 0) {
      setActiveIndex(0);
    } else {
      setActiveIndex(-1);
    }
  }, [results]);

  useEffect(() => {
    let cancelled = false;

    const loadLookupTargets = async () => {
      if (disabled) return;

      try {
        let availableTargets: string[] = [];

        if (loadTargets) {
          availableTargets = await loadTargets();
        }

        const combinedTargets = Array.from(
          new Set(
            [...availableTargets, currentLogicalName || ''].filter(
              (target): target is string => Boolean(target)
            )
          )
        );

        if (!cancelled && isMounted.current) {
          setTargets(combinedTargets);
          setSelectedTarget((previous) => {
            if (previous && combinedTargets.includes(previous)) {
              return previous;
            }
            if (currentLogicalName && combinedTargets.includes(currentLogicalName)) {
              return currentLogicalName;
            }
            return combinedTargets[0];
          });
        }
      } catch (err) {
        if (!cancelled && isMounted.current) {
          console.error('LookupFieldEditor: failed to load targets', err);
          setError(
            err instanceof Error ? err.message : 'Unable to load lookup metadata for search'
          );
        }
      }
    };

    loadLookupTargets();

    return () => {
      cancelled = true;
    };
  }, [loadTargets, currentLogicalName, disabled]);

  const mapResult = useCallback(
    (item: any, metadata: LookupEntityMetadata, logicalName: string): LookupSearchResult | null => {
      const rawId = item?.[metadata.primaryIdAttribute];
      if (!rawId || typeof rawId !== 'string') {
        return null;
      }

      const normalizedId = normalizeGuid(rawId);

      const rawName =
        (metadata.primaryNameAttribute && item?.[metadata.primaryNameAttribute]) || rawId;
      const displayName =
        typeof rawName === 'string' && rawName.trim().length > 0
          ? rawName
          : normalizedId.toUpperCase();

      return {
        recordId: normalizedId,
        displayName,
        logicalName,
        entitySetName: metadata.entitySetName,
      };
    },
    []
  );

  const fetchResults = useCallback(
    async (term: string, logicalName: string, exactId?: string) => {
      if (
        disabled ||
        !isMounted.current ||
        !apiBaseUrl ||
        !logicalName ||
        !getEntityMetadata ||
        (!term && !selectedTarget)
      ) {
        return;
      }

      setLoading(true);
      setError(null);

      try {
        const trimmedTerm = term.trim();
        const metadata = await getEntityMetadata(logicalName);
        if (!metadata || !isMounted.current) {
          throw new Error('Lookup metadata unavailable');
        }

        const selectFields = new Set<string>([metadata.primaryIdAttribute]);
        if (metadata.primaryNameAttribute) {
          selectFields.add(metadata.primaryNameAttribute);
        }
        const selectClause = Array.from(selectFields).join(',');

        const headers: Record<string, string> = {
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
        };

        const mapItems = (items: any[]): LookupSearchResult[] =>
          items
            .map((item) => mapResult(item, metadata, logicalName))
            .filter((item): item is LookupSearchResult => Boolean(item));

        let fetchedResults: LookupSearchResult[] = [];

        if (exactId) {
          const url = `${apiBaseUrl}${metadata.entitySetName}(${exactId})?$select=${selectClause}`;
          const response = await fetch(url, {
            headers,
            credentials: 'include',
          });

          if (response.ok) {
            const item = await response.json();
            const mapped = mapResult(item, metadata, logicalName);
            fetchedResults = mapped ? [mapped] : [];
          } else if (response.status === 404) {
            fetchedResults = [];
            setError('No record found with that ID');
          } else {
            const errText = await response.text();
            throw new Error(`Lookup fetch failed (${response.status}): ${errText}`);
          }
        } else {
          const params: string[] = [];
          if (selectClause) {
            params.push(`$select=${selectClause}`);
          }
          params.push('$top=20');

          const usesContains = Boolean(trimmedTerm && metadata.primaryNameAttribute);

          if (metadata.primaryNameAttribute) {
            if (trimmedTerm) {
              params.push(
                `$filter=contains(${metadata.primaryNameAttribute},'${escapeODataString(
                  trimmedTerm
                )}')`
              );
              params.push(`$orderby=${metadata.primaryNameAttribute}`);
              params.push('$count=true');
              headers['ConsistencyLevel'] = 'eventual';
            } else {
              params.push(`$orderby=${metadata.primaryNameAttribute}`);
            }
          }

          const queryUrl = `${apiBaseUrl}${metadata.entitySetName}?${params.join('&')}`;
          const response = await fetch(queryUrl, {
            headers,
            credentials: 'include',
          });

          if (!response.ok) {
            const errText = await response.text();
            throw new Error(`Lookup search failed (${response.status}): ${errText}`);
          }

          const json = await response.json();
          const items = Array.isArray(json?.value) ? json.value : [];
          fetchedResults = mapItems(items);

          if (trimmedTerm && usesContains && fetchedResults.length === 0) {
            setError('No records found for that search');
          }
        }

        if (!isMounted.current) {
          return;
        }

        setResults(fetchedResults);
        const shouldOpen =
          inputFocusRef.current && (fetchedResults.length > 0 || trimmedTerm.length > 0);
        setShowResults(shouldOpen);
      } catch (err) {
        if (!isMounted.current) {
          return;
        }
        console.error('LookupFieldEditor: search error', err);
        setResults([]);
        setShowResults(false);
        setError(err instanceof Error ? err.message : 'Lookup search failed');
      } finally {
        if (isMounted.current) {
          setLoading(false);
        }
      }
    },
    [apiBaseUrl, disabled, getEntityMetadata, mapResult, selectedTarget]
  );

  useEffect(() => {
    if (disabled || !selectedTarget) {
      return;
    }

    const trimmed = searchTerm.trim();
    const guidMatch = GUID_REGEX.test(trimmed) ? normalizeGuid(trimmed) : undefined;
    const shouldSearch = trimmed.length === 0 || trimmed.length >= 2 || Boolean(guidMatch);

    if (!shouldSearch) {
      setShowResults(false);
      return;
    }

    const handle = window.setTimeout(() => {
      fetchResults(trimmed, selectedTarget, guidMatch);
    }, 300);

    return () => {
      window.clearTimeout(handle);
    };
  }, [searchTerm, selectedTarget, fetchResults, disabled]);

  const handleResultClick = (result: LookupSearchResult) => {
    setSelectedName(result.displayName);
    setSelectedId(result.recordId);
    setSearchTerm('');
    setResults([]);
    setActiveIndex(-1);
    setShowResults(false);
    onSelectionChange(result);
  };

  const handleClearSelection = () => {
    setSelectedName('');
    setSelectedId(null);
    setResults([]);
    setSearchTerm('');
    setActiveIndex(-1);
    setShowResults(false);
    setError(null);
    onSelectionChange(null);
  };

  const handleTargetChange = (event: ChangeEvent<HTMLSelectElement>) => {
    const newTarget = event.target.value || undefined;
    setSelectedTarget(newTarget);
    setResults([]);
    setSearchTerm('');
    setActiveIndex(-1);
    setShowResults(false);
    setError(null);
    onSelectionChange(null);
  };

  const handleInputKeyDown = (event: KeyboardEvent<HTMLInputElement>) => {
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      if (results.length > 0) {
        setShowResults(true);
        setActiveIndex((prev) => {
          const nextIndex = prev + 1;
          return nextIndex >= results.length ? 0 : nextIndex;
        });
      }
      return;
    }

    if (event.key === 'ArrowUp') {
      event.preventDefault();
      if (results.length > 0) {
        setShowResults(true);
        setActiveIndex((prev) => {
          const nextIndex = prev - 1;
          return nextIndex < 0 ? results.length - 1 : nextIndex;
        });
      }
      return;
    }

    if (event.key === 'Enter') {
      event.preventDefault();
      if (activeIndex >= 0 && results[activeIndex]) {
        handleResultClick(results[activeIndex]);
      } else if (results.length === 1) {
        handleResultClick(results[0]);
      }
      return;
    }

    if (event.key === 'Escape') {
      setShowResults(false);
      return;
    }
  };

  const handleInputFocus = () => {
    setIsInputFocused(true);
    inputFocusRef.current = true;
    if (results.length > 0) {
      setShowResults(true);
    }
  };

  const handleInputBlur = () => {
    inputFocusRef.current = false;
    window.setTimeout(() => {
      if (isMounted.current) {
        setIsInputFocused(false);
        setShowResults(false);
      }
    }, 150);
  };

  if (disabled || (!targets.length && !currentLogicalName)) {
    return (
      <div className="lookup-editor lookup-editor--disabled">
        <span className="lookup-editor__label">Lookup:</span>
        <span className="lookup-editor__value">
          {currentName
            ? `${currentName} (${(currentId ?? '').toString().toUpperCase()})`
            : currentId
            ? currentId.toString().toUpperCase()
            : 'Not set'}
        </span>
      </div>
    );
  }

  return (
    <div className="lookup-editor">
      <div className="lookup-editor__selected">
        <span className="lookup-editor__label">Selected:</span>
        <span className="lookup-editor__value">
          {selectedName
            ? `${selectedName}${selectedId ? ` (${selectedId.toUpperCase()})` : ''}`
            : selectedId
            ? selectedId.toUpperCase()
            : 'None'}
        </span>
        <button
          type="button"
          className="lookup-editor__clear"
          onClick={handleClearSelection}
          disabled={!selectedId}
        >
          Clear
        </button>
      </div>

      <div className="lookup-editor__search-row">
        {targets.length > 1 && (
          <select
            className="lookup-editor__target"
            value={selectedTarget || ''}
            onChange={handleTargetChange}
          >
            {targets.map((target) => (
              <option key={target} value={target}>
                {target}
              </option>
            ))}
          </select>
        )}
        <input
          type="text"
          className="lookup-editor__input"
          placeholder={
            selectedTarget
              ? `Search ${selectedTarget} by name or paste ID`
              : `Search ${attributeName} lookup`
          }
          value={searchTerm}
          onChange={(e) => setSearchTerm(e.target.value)}
          onKeyDown={handleInputKeyDown}
          onFocus={handleInputFocus}
          onBlur={handleInputBlur}
        />
        {loading && <span className="lookup-editor__loading">Loading…</span>}
      </div>

      {error && <div className="lookup-editor__error">{error}</div>}

      {showResults && (
        <div className="lookup-editor__results" ref={resultsRef}>
          {results.map((result, index) => (
            <button
              type="button"
              key={`${result.recordId}_${result.logicalName}`}
              className={`lookup-editor__result${index === activeIndex ? ' lookup-editor__result--active' : ''}`}
              onClick={() => handleResultClick(result)}
              onMouseEnter={() => setActiveIndex(index)}
            >
              <span className="lookup-editor__result-name">{result.displayName}</span>
              <span className="lookup-editor__result-meta">
                {result.logicalName} - {result.recordId.toUpperCase()}
              </span>
            </button>
          ))}
          {!loading && results.length === 0 && (
            <div className="lookup-editor__no-results">No matches found</div>
          )}
        </div>
      )}
    </div>
  );
};

export default LookupFieldEditor;
