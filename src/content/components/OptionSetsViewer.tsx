import React, { useMemo, useState } from 'react';

export interface OptionSetOption {
  value: number | string;
  label: string;
  color?: string | null;
  isDefault?: boolean;
}

export interface OptionSetAttribute {
  logicalName: string;
  displayLabel: string;
  attributeType: string;
  isMultiSelect: boolean;
  currentValue?: any;
  currentValueLabel?: string;
  options: OptionSetOption[];
  optionCount: number;
}

export interface OptionSetsData {
  attributes: OptionSetAttribute[];
  error?: string;
}

interface OptionSetsViewerProps {
  data: OptionSetsData | null;
  onClose: () => void;
  onRefresh: () => void;
}

const OptionSetsViewer: React.FC<OptionSetsViewerProps> = ({ data, onClose, onRefresh }) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [copiedKey, setCopiedKey] = useState<string>('');

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

  const attributes = data.attributes || [];

  const filteredAttributes = useMemo(() => {
    const term = searchTerm.trim().toLowerCase();
    if (!term) {
      return attributes;
    }

    return attributes.filter(attribute => {
      const matchAttribute =
        attribute.displayLabel.toLowerCase().includes(term) ||
        attribute.logicalName.toLowerCase().includes(term) ||
        attribute.attributeType.toLowerCase().includes(term);

      if (matchAttribute) {
        return true;
      }

      return attribute.options.some(option => {
        return (
          option.label.toLowerCase().includes(term) ||
          String(option.value).toLowerCase().includes(term)
        );
      });
    });
  }, [attributes, searchTerm]);

  const renderOptions = (attribute: OptionSetAttribute) => {
    if (!attribute.options || attribute.options.length === 0) {
      return <div className="d365-options-empty">No option values found.</div>;
    }

    return (
      <div className="d365-options-table">
        <div className="d365-options-table-head">
          <span>Value</span>
          <span>Label</span>
          <span></span>
        </div>
        {attribute.options.map((option, index) => {
          const key = `${attribute.logicalName}-${option.value}`;
          const hasColor = option.color && option.color !== 'transparent';

          return (
            <div
              key={key}
              className={`d365-options-table-row ${option.isDefault ? 'd365-options-table-row-default' : ''
                }`}
            >
              <span className="d365-options-option-value">{option.value}</span>
              <span className="d365-options-option-label">
                {hasColor && (
                  <span
                    className="d365-options-color-chip"
                    style={{ backgroundColor: option.color! }}
                    title={`Color: ${option.color}`}
                  />
                )}
                {option.label}
              </span>
              <button
                className="d365-options-copy-btn"
                onClick={() => handleCopy(`${option.value} - ${option.label}`, key)}
                title="Copy value and label"
              >
                {copiedKey === key ? 'V' : '??'}
              </button>
            </div>
          );
        })}
      </div>
    );
  };

  const renderBody = () => {
    if (data.error) {
      return (
        <div className="d365-dialog-content d365-options-content">
          <div className="d365-dialog-error">
            <div className="d365-dialog-error-icon">??</div>
            <div className="d365-dialog-error-message">{data.error}</div>
            <div className="d365-dialog-error-hint">
              Open an entity form and make sure it contains option set or picklist fields.
            </div>
          </div>
        </div>
      );
    }

    return (
      <div className="d365-dialog-content d365-options-content">
        {filteredAttributes.length === 0 ? (
          <div className="d365-options-empty">
            {attributes.length === 0
              ? 'No option set fields detected on this form.'
              : 'No option set fields match your search.'}
          </div>
        ) : (
          <div className="d365-options-list">
            {filteredAttributes.map(attribute => {
              const copyKey = `${attribute.logicalName}-all`;
              const serializedOptions = attribute.options
                .map(option => `${option.value}\t${option.label}`)
                .join('\n');

              return (
                <div key={attribute.logicalName} className="d365-options-attribute">
                  <div className="d365-options-attribute-header">
                    <div className="d365-options-attribute-title">
                      <span className="d365-options-attribute-name">
                        {attribute.displayLabel || attribute.logicalName}
                      </span>
                      <span className="d365-options-attribute-logical">
                        {attribute.logicalName}
                      </span>
                    </div>
                    <div className="d365-options-attribute-meta">
                      <span>
                        Type:{' '}
                        {attribute.attributeType === 'optionset'
                          ? 'Option Set'
                          : attribute.attributeType === 'multioptionset'
                            ? 'Multi Select Option Set'
                            : attribute.attributeType === 'boolean'
                              ? 'Two Options'
                              : attribute.attributeType}
                      </span>
                      <span>Options: {attribute.optionCount}</span>
                      {attribute.currentValueLabel && (
                        <span>Current: {attribute.currentValueLabel}</span>
                      )}
                    </div>
                    <div className="d365-options-attribute-actions">
                      <button
                        className="d365-options-copy-btn"
                        onClick={() => handleCopy(attribute.logicalName, `${attribute.logicalName}-logical`)}
                        title="Copy logical name"
                      >
                        {copiedKey === `${attribute.logicalName}-logical` ? 'V' : '??'}
                      </button>
                      {attribute.optionCount > 0 && (
                        <button
                          className="d365-options-copy-btn"
                          onClick={() => handleCopy(serializedOptions, copyKey)}
                          title="Copy all options (Tab separated)"
                        >
                          {copiedKey === copyKey ? 'V Copied' : 'Copy All'}
                        </button>
                      )}
                    </div>
                  </div>
                  {renderOptions(attribute)}
                </div>
              );
            })}
          </div>
        )}
      </div>
    );
  };

  return (
    <div className="d365-dialog-overlay">
      <div className="d365-dialog-modal d365-options-modal">
        <div className="d365-dialog-header d365-options-header">
          <h2>Option Sets</h2>
          <div className="d365-options-header-actions">
            <button
              className="d365-trace-refresh"
              onClick={onRefresh}
              title="Refresh option set values"
            >
              ??
            </button>
            <button className="d365-dialog-close" onClick={onClose} title="Close">
              ?
            </button>
          </div>
        </div>

        {!data.error && (
          <div className="d365-options-toolbar">
            <div className="d365-options-count">
              Showing {filteredAttributes.length} of {attributes.length}
            </div>
            <div className="d365-options-search">
              <input
                type="text"
                className="d365-options-search-input"
                placeholder="Search option sets..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
              {searchTerm && (
                <button
                  className="d365-options-search-clear"
                  onClick={() => setSearchTerm('')}
                  title="Clear search"
                >
                  ?
                </button>
              )}
            </div>
          </div>
        )}

        {renderBody()}
      </div>
    </div>
  );
};

export default OptionSetsViewer;
