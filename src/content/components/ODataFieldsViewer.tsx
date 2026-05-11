import React, { useMemo, useState } from 'react';
import CloseIcon from './CloseIcon';

export interface ODataField {
  logicalName: string;
  schemaName: string;
  attributeType: string;
  odataBind?: string;
  optionSetValues?: string;
  relationshipName?: string;
  targetEntity?: string;
  /** Logical name of the parent attribute when this is a shadow attribute
   *  (e.g. lookup-name virtuals like `rg_username` extending `rg_user`). */
  attributeOf?: string;
}

export interface ODataFieldsData {
  entityName: string;
  entitySetName: string;
  fields: ODataField[];
  error?: string;
}

interface ODataFieldsViewerProps {
  data: ODataFieldsData | null;
  onClose: () => void;
  onRefresh: () => void;
}

const ODataFieldsViewer: React.FC<ODataFieldsViewerProps> = ({ data, onClose, onRefresh }) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [copiedKey, setCopiedKey] = useState<string>('');
  const [hideShadowAttrs, setHideShadowAttrs] = useState(true);

  if (!data) {
    return null;
  }

  const handleCopy = async (text: string, key: string) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopiedKey(key);
      setTimeout(() => setCopiedKey(''), 2000);
    } catch (error) {
      console.error('Failed to copy text', error);
    }
  };

  const fields = data.fields || [];

  const filteredFields = useMemo(() => {
    const term = searchTerm.trim().toLowerCase();
    return fields.filter(field => {
      if (hideShadowAttrs && field.attributeOf) return false;
      if (!term) return true;
      return (
        field.logicalName.toLowerCase().includes(term) ||
        field.schemaName.toLowerCase().includes(term) ||
        field.attributeType.toLowerCase().includes(term) ||
        field.targetEntity?.toLowerCase().includes(term) ||
        field.relationshipName?.toLowerCase().includes(term) ||
        field.optionSetValues?.toLowerCase().includes(term) ||
        field.attributeOf?.toLowerCase().includes(term)
      );
    });
  }, [fields, searchTerm, hideShadowAttrs]);

  const shadowCount = useMemo(
    () => fields.filter(f => f.attributeOf).length,
    [fields]
  );

  const handleCopyAll = () => {
    const output = filteredFields
      .map(field => {
        const parts = [
          field.logicalName,
          field.schemaName,
          field.attributeType,
          field.attributeOf || '',
          field.odataBind || '',
          field.optionSetValues || '',
          field.relationshipName || '',
          field.targetEntity || ''
        ];
        return parts.join('\t');
      })
      .join('\n');

    handleCopy(output, 'copy-all');
  };

  const renderBody = () => {
    if (data.error) {
      return (
        <div className="d365-dialog-content d365-options-content">
          <div className="d365-dialog-error">
            <div className="d365-dialog-error-icon">⚠</div>
            <div className="d365-dialog-error-message">{data.error}</div>
            <div className="d365-dialog-error-hint">
              Open an entity form to view OData field metadata.
            </div>
          </div>
        </div>
      );
    }

    return (
      <div className="d365-dialog-content d365-options-content">
        {filteredFields.length === 0 ? (
          <div className="d365-options-empty">
            {fields.length === 0
              ? 'No fields detected on this form.'
              : 'No fields match your search.'}
          </div>
        ) : (
          <div className="d365-odata-table-container">
            <table className="d365-odata-table">
              <thead>
                <tr>
                  <th>Logical Name</th>
                  <th>Schema Name</th>
                  <th>Type</th>
                  <th>OData Bind</th>
                  <th>Option Set Values</th>
                  <th>Relationship Name</th>
                  <th>Target Entity</th>
                  <th>Copy</th>
                </tr>
              </thead>
              <tbody>
                {filteredFields.map(field => {
                  const rowKey = `${field.logicalName}`;
                  const copyText = [
                    field.logicalName,
                    field.schemaName,
                    field.attributeType,
                    field.attributeOf ? `shadow of ${field.attributeOf}` : '',
                    field.odataBind || '',
                    field.optionSetValues || '',
                    field.relationshipName || '',
                    field.targetEntity || ''
                  ].filter(Boolean).join(' | ');

                  return (
                    <tr key={rowKey}>
                      <td className="d365-odata-logical">
                        <div className="d365-odata-cell">
                          <span className="d365-odata-cell-value">{field.logicalName}</span>
                          <button
                            className="d365-options-copy-btn d365-odata-copy-btn"
                            onClick={() => handleCopy(field.logicalName, `${rowKey}-logical`)}
                            title="Copy logical name"
                          >
                            {copiedKey === `${rowKey}-logical` ? (
                              <span className="d365-copy-success">V</span>
                            ) : (
                              <img
                                src={chrome.runtime.getURL('icons/rg_copy.svg')}
                                alt="Copy logical name"
                                className="d365-copy-icon"
                              />
                            )}
                          </button>
                        </div>
                      </td>
                      <td className="d365-odata-schema">
                        <div className="d365-odata-cell">
                          <span className="d365-odata-cell-value">{field.schemaName}</span>
                          <button
                            className="d365-options-copy-btn d365-odata-copy-btn"
                            onClick={() => handleCopy(field.schemaName, `${rowKey}-schema`)}
                            title="Copy schema name"
                          >
                            {copiedKey === `${rowKey}-schema` ? (
                              <span className="d365-copy-success">V</span>
                            ) : (
                              <img
                                src={chrome.runtime.getURL('icons/rg_copy.svg')}
                                alt="Copy schema name"
                                className="d365-copy-icon"
                              />
                            )}
                          </button>
                        </div>
                      </td>
                      <td className="d365-odata-type">
                        <div className="d365-odata-cell">
                          <span className="d365-odata-cell-value">
                            {field.attributeType}
                            {field.attributeOf && (
                              <span
                                className="d365-odata-shadow-tag"
                                title={`Shadow attribute extending ${field.attributeOf} (auto-generated by Dataverse, read-only)`}
                                style={{
                                  marginLeft: 6,
                                  padding: '1px 6px',
                                  borderRadius: 4,
                                  background: '#eef2ff',
                                  color: '#4338ca',
                                  fontSize: '0.75em',
                                  whiteSpace: 'nowrap'
                                }}
                              >
                                shadow of {field.attributeOf}
                              </span>
                            )}
                          </span>
                          <button
                            className="d365-options-copy-btn d365-odata-copy-btn"
                            onClick={() => handleCopy(field.attributeType, `${rowKey}-type`)}
                            title="Copy attribute type"
                          >
                            {copiedKey === `${rowKey}-type` ? (
                              <span className="d365-copy-success">V</span>
                            ) : (
                              <img
                                src={chrome.runtime.getURL('icons/rg_copy.svg')}
                                alt="Copy attribute type"
                                className="d365-copy-icon"
                              />
                            )}
                          </button>
                        </div>
                      </td>
                      <td className="d365-odata-bind">
                        <div className="d365-odata-cell">
                          <span className="d365-odata-cell-value">
                            {field.odataBind || '-'}
                          </span>
                          {field.odataBind && (
                            <button
                              className="d365-options-copy-btn d365-odata-copy-btn"
                              onClick={() => handleCopy(field.odataBind!, `${rowKey}-bind`)}
                              title="Copy OData bind path"
                            >
                              {copiedKey === `${rowKey}-bind` ? (
                                <span className="d365-copy-success">V</span>
                              ) : (
                                <img
                                  src={chrome.runtime.getURL('icons/rg_copy.svg')}
                                  alt="Copy OData bind"
                                  className="d365-copy-icon"
                                />
                              )}
                            </button>
                          )}
                        </div>
                      </td>
                      <td className="d365-odata-options">
                        <div className="d365-odata-cell">
                          {field.optionSetValues ? (
                            <span
                              className="d365-odata-truncated d365-odata-cell-value"
                              title={field.optionSetValues}
                            >
                              {field.optionSetValues}
                            </span>
                          ) : (
                            <span className="d365-odata-cell-value">-</span>
                          )}
                          {field.optionSetValues && (
                            <button
                              className="d365-options-copy-btn d365-odata-copy-btn"
                              onClick={() => handleCopy(field.optionSetValues!, `${rowKey}-options`)}
                              title="Copy option set values"
                            >
                              {copiedKey === `${rowKey}-options` ? (
                                <span className="d365-copy-success">V</span>
                              ) : (
                                <img
                                  src={chrome.runtime.getURL('icons/rg_copy.svg')}
                                  alt="Copy option set values"
                                  className="d365-copy-icon"
                                />
                              )}
                            </button>
                          )}
                        </div>
                      </td>
                      <td className="d365-odata-relationship">
                        <div className="d365-odata-cell">
                          {field.relationshipName ? (
                            <span
                              className="d365-odata-truncated d365-odata-cell-value"
                              title={field.relationshipName}
                            >
                              {field.relationshipName}
                            </span>
                          ) : (
                            <span className="d365-odata-cell-value">-</span>
                          )}
                          {field.relationshipName && (
                            <button
                              className="d365-options-copy-btn d365-odata-copy-btn"
                              onClick={() => handleCopy(field.relationshipName!, `${rowKey}-relationship`)}
                              title="Copy relationship name"
                            >
                              {copiedKey === `${rowKey}-relationship` ? (
                                <span className="d365-copy-success">V</span>
                              ) : (
                                <img
                                  src={chrome.runtime.getURL('icons/rg_copy.svg')}
                                  alt="Copy relationship name"
                                  className="d365-copy-icon"
                                />
                              )}
                            </button>
                          )}
                        </div>
                      </td>
                      <td className="d365-odata-target">
                        <div className="d365-odata-cell">
                          <span className="d365-odata-cell-value">{field.targetEntity || '-'}</span>
                          {field.targetEntity && (
                            <button
                              className="d365-options-copy-btn d365-odata-copy-btn"
                              onClick={() => handleCopy(field.targetEntity!, `${rowKey}-target`)}
                              title="Copy target entity"
                            >
                              {copiedKey === `${rowKey}-target` ? (
                                <span className="d365-copy-success">V</span>
                              ) : (
                                <img
                                  src={chrome.runtime.getURL('icons/rg_copy.svg')}
                                  alt="Copy target entity"
                                  className="d365-copy-icon"
                                />
                              )}
                            </button>
                          )}
                        </div>
                      </td>
                      <td>
                        <button
                          className="d365-options-copy-btn"
                          onClick={() => handleCopy(copyText, rowKey)}
                          title="Copy field info"
                        >
                          {copiedKey === rowKey ? (
                            <span className="d365-copy-success">✓</span>
                          ) : (
                            <img
                              src={chrome.runtime.getURL('icons/rg_copy.svg')}
                              alt="Copy"
                              className="d365-copy-icon"
                            />
                          )}
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    );
  };

  return (
    <div className="d365-dialog-overlay">
      <div className="d365-dialog-modal d365-odata-modal">
        <div className="d365-dialog-header d365-options-header">
          <div>
            <h2>OData Fields</h2>
            {data.entityName && (
              <div className="d365-odata-entity-info">
                Entity: {data.entityName} | EntitySet: {data.entitySetName}
              </div>
            )}
          </div>
          <div className="d365-options-header-actions">
            <button
              className="d365-trace-refresh"
              onClick={onRefresh}
              title="Refresh OData fields"
            >
              ↻
            </button>
            <button className="d365-dialog-close" onClick={onClose} title="Close" aria-label="Close">
              <CloseIcon />
            </button>
          </div>
        </div>

        {!data.error && (
          <div className="d365-options-toolbar">
            <div className="d365-options-count">
              Showing {filteredFields.length} of {fields.length} fields
              {hideShadowAttrs && shadowCount > 0 && (
                <span style={{ marginLeft: 8, color: '#6b7280' }}>
                  ({shadowCount} auto-generated hidden)
                </span>
              )}
            </div>
            <div className="d365-options-search" style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
              {shadowCount > 0 && (
                <label
                  style={{ display: 'flex', alignItems: 'center', gap: 4, fontSize: '0.85em', whiteSpace: 'nowrap' }}
                  title="Hide shadow attributes auto-generated by Dataverse for custom lookups (e.g. rg_username, rg_useryominame)"
                >
                  <input
                    type="checkbox"
                    checked={hideShadowAttrs}
                    onChange={(e) => setHideShadowAttrs(e.target.checked)}
                  />
                  Hide auto-generated
                </label>
              )}
              <input
                type="text"
                className="d365-options-search-input"
                placeholder="Search fields..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
              {searchTerm && (
                <button
                  className="d365-options-search-clear"
                  onClick={() => setSearchTerm('')}
                  title="Clear search"
                  aria-label="Clear search"
                >
                  <CloseIcon size={12} />
                </button>
              )}
              <button
                className="d365-libraries-copy-all-btn"
                onClick={handleCopyAll}
                title="Copy all visible fields to clipboard"
              >
                {copiedKey === 'copy-all' ? '✓ Copied!' : 'Copy All'}
              </button>
            </div>
          </div>
        )}

        {renderBody()}
      </div>
    </div>
  );
};

export default ODataFieldsViewer;
