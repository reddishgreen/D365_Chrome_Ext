import React, { useState, useEffect, useMemo, useRef } from 'react';
import './ImpersonationSelector.css';

export interface SystemUser {
  systemuserid: string;
  fullname: string;
  domainname: string;
  internalemailaddress?: string;
}

export interface ImpersonationData {
  users: SystemUser[];
  error?: string;
}

interface ImpersonationSelectorProps {
  data: ImpersonationData | null;
  onClose: () => void;
  onSelect: (user: SystemUser) => void;
  onRefresh: () => void;
  currentImpersonation: SystemUser | null;
}

const ImpersonationSelector: React.FC<ImpersonationSelectorProps> = ({
  data,
  onClose,
  onSelect,
  onRefresh,
  currentImpersonation
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedUser, setSelectedUser] = useState<SystemUser | null>(null);
  const [showResults, setShowResults] = useState(false);
  const [activeIndex, setActiveIndex] = useState(-1);
  const inputRef = useRef<HTMLInputElement>(null);
  const resultsRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    // Focus input on mount
    if (inputRef.current) {
      inputRef.current.focus();
    }
  }, []);

  const filteredUsers = useMemo(() => {
    if (!data?.users) return [];
    
    const term = searchTerm.trim().toLowerCase();
    if (!term) return data.users;

    return data.users.filter(user => {
      return (
        user.fullname?.toLowerCase().includes(term) ||
        user.domainname?.toLowerCase().includes(term) ||
        user.internalemailaddress?.toLowerCase().includes(term)
      );
    });
  }, [data?.users, searchTerm]);

  const handleSelect = (user: SystemUser) => {
    setSelectedUser(user);
    setSearchTerm(user.fullname);
    setShowResults(false);
  };

  const handleSubmit = () => {
    if (selectedUser) {
      onSelect(selectedUser);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (!showResults || filteredUsers.length === 0) return;

    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault();
        setActiveIndex(prev => Math.min(prev + 1, filteredUsers.length - 1));
        break;
      case 'ArrowUp':
        e.preventDefault();
        setActiveIndex(prev => Math.max(prev - 1, 0));
        break;
      case 'Enter':
        e.preventDefault();
        if (activeIndex >= 0 && activeIndex < filteredUsers.length) {
          handleSelect(filteredUsers[activeIndex]);
        }
        break;
      case 'Escape':
        setShowResults(false);
        break;
    }
  };

  const handleBlur = (e: React.FocusEvent) => {
    // Delay hiding results to allow click events to fire
    setTimeout(() => {
      if (!resultsRef.current?.contains(document.activeElement)) {
        setShowResults(false);
      }
    }, 200);
  };

  const renderBody = () => {
    if (!data) {
      return (
        <div className="d365-impersonate-content">
          <div className="d365-impersonate-loading">
            <div className="d365-impersonate-spinner"></div>
            <p>Loading users...</p>
          </div>
        </div>
      );
    }

    if (data.error) {
      return (
        <div className="d365-impersonate-content">
          <div className="d365-dialog-error">
            <div className="d365-dialog-error-icon">⚠</div>
            <div className="d365-dialog-error-message">{data.error}</div>
            <div className="d365-dialog-error-hint">
              Make sure you have permission to view system users.
            </div>
          </div>
        </div>
      );
    }

    return (
      <div className="d365-impersonate-content">
        {currentImpersonation && (
          <div className="d365-impersonate-current">
            <span className="d365-impersonate-current-label">Currently impersonating:</span>
            <span className="d365-impersonate-current-user">{currentImpersonation.fullname}</span>
          </div>
        )}

        <div className="d365-impersonate-form">
          <label className="d365-impersonate-label">Select a user to impersonate:</label>
          
          <div className="d365-impersonate-search">
            <input
              ref={inputRef}
              type="text"
              className="d365-impersonate-input"
              placeholder="Search by name or email..."
              value={searchTerm}
              onChange={(e) => {
                setSearchTerm(e.target.value);
                setSelectedUser(null);
                setShowResults(true);
                setActiveIndex(-1);
              }}
              onFocus={() => setShowResults(true)}
              onBlur={handleBlur}
              onKeyDown={handleKeyDown}
            />
            
            {showResults && filteredUsers.length > 0 && (
              <div className="d365-impersonate-results" ref={resultsRef}>
                {filteredUsers.slice(0, 50).map((user, index) => (
                  <button
                    key={user.systemuserid}
                    type="button"
                    className={`d365-impersonate-result ${index === activeIndex ? 'd365-impersonate-result--active' : ''}`}
                    onClick={() => handleSelect(user)}
                    onMouseEnter={() => setActiveIndex(index)}
                  >
                    <span className="d365-impersonate-result-name">{user.fullname}</span>
                    <span className="d365-impersonate-result-domain">{user.domainname}</span>
                  </button>
                ))}
                {filteredUsers.length > 50 && (
                  <div className="d365-impersonate-more">
                    +{filteredUsers.length - 50} more users. Type to narrow results.
                  </div>
                )}
              </div>
            )}
            
            {showResults && searchTerm && filteredUsers.length === 0 && (
              <div className="d365-impersonate-results">
                <div className="d365-impersonate-no-results">No users found</div>
              </div>
            )}
          </div>

          {selectedUser && (
            <div className="d365-impersonate-selected">
              <div className="d365-impersonate-selected-label">Selected:</div>
              <div className="d365-impersonate-selected-user">
                <strong>{selectedUser.fullname}</strong>
                <span>{selectedUser.domainname}</span>
              </div>
            </div>
          )}

          <div className="d365-impersonate-info">
            <p><strong>Note:</strong> Impersonation adds the MSCRMCallerID header to Web API calls.</p>
            <p>Records you create/modify will show the impersonated user as creator/modifier.</p>
            <p>You need the "Act on Behalf of Another User" privilege for this to work.</p>
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="d365-dialog-overlay">
      <div className="d365-dialog-modal d365-impersonate-modal">
        <div className="d365-dialog-header d365-impersonate-header">
          <h2>Impersonate User</h2>
          <div className="d365-impersonate-header-actions">
            <button
              className="d365-trace-refresh"
              onClick={onRefresh}
              title="Refresh user list"
            >
              ↻
            </button>
            <button className="d365-dialog-close" onClick={onClose} title="Close">
              ×
            </button>
          </div>
        </div>

        {renderBody()}

        <div className="d365-impersonate-footer">
          <button
            className="d365-impersonate-btn d365-impersonate-btn-cancel"
            onClick={onClose}
          >
            Cancel
          </button>
          <button
            className="d365-impersonate-btn d365-impersonate-btn-submit"
            onClick={handleSubmit}
            disabled={!selectedUser}
          >
            Start Impersonating
          </button>
        </div>
      </div>
    </div>
  );
};

export default ImpersonationSelector;

