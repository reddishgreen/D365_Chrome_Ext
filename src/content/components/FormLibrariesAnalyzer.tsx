import React, { useState } from 'react';

interface EventHandler {
  type: string;
  target: string;
  library: string;
  functionName: string;
  enabled: boolean;
}

interface Library {
  name: string;
  order: number;
}

interface FormLibrariesData {
  libraries: Library[];
  onLoad: EventHandler[];
  onChange: EventHandler[];
  onSave: EventHandler[];
  error?: string;
}

interface FormLibrariesAnalyzerProps {
  data: FormLibrariesData | null;
  onClose: () => void;
}

const FormLibrariesAnalyzer: React.FC<FormLibrariesAnalyzerProps> = ({ data, onClose }) => {
  const [activeTab, setActiveTab] = useState<'onload' | 'onchange' | 'onsave' | 'libraries'>('onload');
  const [copiedText, setCopiedText] = useState<string>('');

  if (!data) return null;

  // Check if there's an error
  if (data.error) {
    return (
      <div className="d365-libraries-overlay">
        <div className="d365-libraries-modal">
          <div className="d365-libraries-header">
            <h2>JavaScript Libraries & Event Handlers</h2>
            <button className="d365-libraries-close" onClick={onClose} title="Close">✕</button>
          </div>
          <div className="d365-libraries-content">
            <div className="d365-libraries-error">
              <div className="d365-libraries-error-icon">⚠️</div>
              <div className="d365-libraries-error-message">{data.error}</div>
              <div className="d365-libraries-error-hint">
                Make sure you are on a D365 entity form page (e.g., Account, Contact, etc.)
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  const handleCopy = async (text: string, label: string) => {
    try {
      await navigator.clipboard.writeText(text);
      setCopiedText(label);
      setTimeout(() => setCopiedText(''), 2000);
    } catch (error) {
      console.error('Failed to copy:', error);
    }
  };

  const handleCopyAll = async (handlers: EventHandler[], eventType: string) => {
    const text = handlers
      .map(h => `${h.library}.${h.functionName} (${h.target})`)
      .join('\n');
    await handleCopy(text, `${eventType}-all`);
  };

  const renderEventHandlers = (handlers: EventHandler[], emptyMessage: string) => {
    if (handlers.length === 0) {
      return <div className="d365-libraries-empty">{emptyMessage}</div>;
    }

    // Group handlers by library
    const groupedByLibrary = handlers.reduce((acc, handler) => {
      if (!acc[handler.library]) {
        acc[handler.library] = [];
      }
      acc[handler.library].push(handler);
      return acc;
    }, {} as Record<string, EventHandler[]>);

    return (
      <div className="d365-libraries-handlers">
        {Object.entries(groupedByLibrary).map(([library, libHandlers]) => (
          <div key={library} className="d365-libraries-group">
            <div className="d365-libraries-group-header">
              <span className="d365-libraries-library-name">{library}</span>
              <span className="d365-libraries-count">({libHandlers.length})</span>
            </div>
            <div className="d365-libraries-functions">
              {libHandlers.map((handler, idx) => (
                <div key={idx} className="d365-libraries-function-item">
                  <div className="d365-libraries-function-main">
                    <span className="d365-libraries-function-name">{handler.functionName}</span>
                    {handler.type === 'field' && (
                      <span className="d365-libraries-field-target"> → {handler.target}</span>
                    )}
                    {!handler.enabled && (
                      <span className="d365-libraries-disabled"> (disabled)</span>
                    )}
                  </div>
                  <button
                    className="d365-libraries-copy-btn"
                    onClick={() => handleCopy(`${library}.${handler.functionName}`, `${library}.${handler.functionName}`)}
                    title="Copy function reference"
                  >
                    {copiedText === `${library}.${handler.functionName}` ? (
                      <span className="d365-copy-success">✓</span>
                    ) : (
                      <img
                        src={chrome.runtime.getURL('icons/rg_copy.svg')}
                        alt="Copy"
                        className="d365-copy-icon"
                      />
                    )}
                  </button>
                </div>
              ))}
            </div>
          </div>
        ))}
      </div>
    );
  };

  return (
    <div className="d365-libraries-overlay">
      <div className="d365-libraries-modal">
        <div className="d365-libraries-header">
          <h2>JavaScript Libraries & Event Handlers</h2>
          <button className="d365-libraries-close" onClick={onClose} title="Close">✕</button>
        </div>

        <div className="d365-libraries-tabs">
          <button
            className={`d365-libraries-tab ${activeTab === 'onload' ? 'd365-libraries-tab-active' : ''}`}
            onClick={() => setActiveTab('onload')}
          >
            OnLoad ({data.onLoad.length})
          </button>
          <button
            className={`d365-libraries-tab ${activeTab === 'onchange' ? 'd365-libraries-tab-active' : ''}`}
            onClick={() => setActiveTab('onchange')}
          >
            OnChange ({data.onChange.length})
          </button>
          <button
            className={`d365-libraries-tab ${activeTab === 'onsave' ? 'd365-libraries-tab-active' : ''}`}
            onClick={() => setActiveTab('onsave')}
          >
            OnSave ({data.onSave.length})
          </button>
          <button
            className={`d365-libraries-tab ${activeTab === 'libraries' ? 'd365-libraries-tab-active' : ''}`}
            onClick={() => setActiveTab('libraries')}
          >
            Libraries ({data.libraries.length})
          </button>
        </div>

        <div className="d365-libraries-content">
          {activeTab === 'onload' && (
            <div className="d365-libraries-tab-content">
              <div className="d365-libraries-tab-header">
                <h3>Form OnLoad Event Handlers</h3>
                {data.onLoad.length > 0 && (
                  <button
                    className="d365-libraries-copy-all-btn"
                    onClick={() => handleCopyAll(data.onLoad, 'onload')}
                  >
                    {copiedText === 'onload-all' ? '✓ Copied' : 'Copy All'}
                  </button>
                )}
              </div>
              {renderEventHandlers(data.onLoad, 'No OnLoad handlers registered')}
            </div>
          )}

          {activeTab === 'onchange' && (
            <div className="d365-libraries-tab-content">
              <div className="d365-libraries-tab-header">
                <h3>Field OnChange Event Handlers</h3>
                {data.onChange.length > 0 && (
                  <button
                    className="d365-libraries-copy-all-btn"
                    onClick={() => handleCopyAll(data.onChange, 'onchange')}
                  >
                    {copiedText === 'onchange-all' ? '✓ Copied' : 'Copy All'}
                  </button>
                )}
              </div>
              {renderEventHandlers(data.onChange, 'No OnChange handlers registered')}
            </div>
          )}

          {activeTab === 'onsave' && (
            <div className="d365-libraries-tab-content">
              <div className="d365-libraries-tab-header">
                <h3>Form OnSave Event Handlers</h3>
                {data.onSave.length > 0 && (
                  <button
                    className="d365-libraries-copy-all-btn"
                    onClick={() => handleCopyAll(data.onSave, 'onsave')}
                  >
                    {copiedText === 'onsave-all' ? '✓ Copied' : 'Copy All'}
                  </button>
                )}
              </div>
              {renderEventHandlers(data.onSave, 'No OnSave handlers registered')}
            </div>
          )}

          {activeTab === 'libraries' && (
            <div className="d365-libraries-tab-content">
              <div className="d365-libraries-tab-header">
                <h3>Form Libraries</h3>
                {data.libraries.length > 0 && (
                  <button
                    className="d365-libraries-copy-all-btn"
                    onClick={() => handleCopy(data.libraries.map(l => l.name).join('\n'), 'libraries-all')}
                  >
                    {copiedText === 'libraries-all' ? '✓ Copied' : 'Copy All'}
                  </button>
                )}
              </div>
              {data.libraries.length === 0 ? (
                <div className="d365-libraries-empty">No libraries loaded</div>
              ) : (
                <div className="d365-libraries-list">
                  {data.libraries.map((lib, idx) => (
                    <div key={idx} className="d365-libraries-list-item">
                      <span className="d365-libraries-list-name">{lib.name}</span>
                      <button
                        className="d365-libraries-copy-btn"
                        onClick={() => handleCopy(lib.name, lib.name)}
                        title="Copy library name"
                      >
                        {copiedText === lib.name ? (
                          <span className="d365-copy-success">✓</span>
                        ) : (
                          <img
                            src={chrome.runtime.getURL('icons/rg_copy.svg')}
                            alt="Copy"
                            className="d365-copy-icon"
                          />
                        )}
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default FormLibrariesAnalyzer;
