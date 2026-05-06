import React, { useState } from 'react';
import './PrivilegeDebugger.css';
import CloseIcon from './CloseIcon';

export interface PrivilegeAccessRights {
  ReadAccess: boolean;
  WriteAccess: boolean;
  DeleteAccess: boolean;
  AppendAccess: boolean;
  AppendToAccess: boolean;
  AssignAccess: boolean;
  ShareAccess: boolean;
  CreateAccess?: boolean;
}

export interface PrivilegeRoleSummary {
  roleId: string;
  roleName: string;
  businessUnitName?: string;
  privileges: {
    accessRight: string;
    depth: string;
    depthRank: number;
  }[];
}

export interface PrivilegeRecordInfo {
  entityName: string;
  recordId: string;
  recordName?: string;
  ownerType?: string;
  ownerName?: string;
  ownerBusinessUnit?: string;
  recordBusinessUnit?: string;
}

export interface PrivilegeUserInfo {
  userId: string;
  fullName: string;
  businessUnitId: string;
  businessUnitName?: string;
  isDisabled?: boolean;
  teams: { id: string; name: string }[];
}

export interface PrivilegeDebugData {
  inputEntityName: string;
  inputRecordId: string;
  effectiveAccess?: PrivilegeAccessRights;
  record?: PrivilegeRecordInfo;
  user?: PrivilegeUserInfo;
  roles?: PrivilegeRoleSummary[];
  diagnosis?: string[];
  error?: string;
}

interface PrivilegeDebuggerProps {
  data: PrivilegeDebugData | null;
  defaultEntityName?: string;
  defaultRecordId?: string;
  onClose: () => void;
  onRun: (entityName: string, recordId: string) => Promise<void>;
  isLoading?: boolean;
}

const ACCESS_LABELS: { key: keyof PrivilegeAccessRights; label: string }[] = [
  { key: 'ReadAccess', label: 'Read' },
  { key: 'WriteAccess', label: 'Write' },
  { key: 'DeleteAccess', label: 'Delete' },
  { key: 'AppendAccess', label: 'Append' },
  { key: 'AppendToAccess', label: 'Append To' },
  { key: 'AssignAccess', label: 'Assign' },
  { key: 'ShareAccess', label: 'Share' },
];

const PrivilegeDebugger: React.FC<PrivilegeDebuggerProps> = ({
  data,
  defaultEntityName,
  defaultRecordId,
  onClose,
  onRun,
  isLoading,
}) => {
  const [entityName, setEntityName] = useState(defaultEntityName || '');
  const [recordId, setRecordId] = useState(defaultRecordId || '');

  const handleRun = async () => {
    const trimmedEntity = entityName.trim();
    const trimmedId = recordId.trim().replace(/[{}]/g, '');
    if (!trimmedEntity || !trimmedId) return;
    await onRun(trimmedEntity, trimmedId);
  };

  return (
    <div className="pd-overlay">
      <div className="pd-container">
        <div className="pd-header">
          <div>
            <h2>Privilege Debugger</h2>
            <div className="pd-subtitle">Diagnose record access for the current user</div>
          </div>
          <div className="pd-actions">
            {isLoading && <span className="pd-loading-pill">Checking...</span>}
            <button onClick={onClose} className="pd-icon-btn" title="Close" aria-label="Close">
              <CloseIcon />
            </button>
          </div>
        </div>

        <div className="pd-input-bar">
          <div className="pd-field">
            <label>Entity logical name</label>
            <input
              type="text"
              className="pd-input"
              placeholder="e.g. account"
              value={entityName}
              onChange={(e) => setEntityName(e.target.value)}
            />
          </div>
          <div className="pd-field">
            <label>Record GUID</label>
            <input
              type="text"
              className="pd-input"
              placeholder="00000000-0000-0000-0000-000000000000"
              value={recordId}
              onChange={(e) => setRecordId(e.target.value)}
            />
          </div>
          <button
            className="pd-run-btn"
            onClick={handleRun}
            disabled={isLoading || !entityName.trim() || !recordId.trim()}
          >
            {isLoading ? 'Checking...' : 'Run check'}
          </button>
        </div>

        {data?.error ? (
          <div className="pd-error">{data.error}</div>
        ) : !data ? (
          <div className="pd-empty">Enter an entity and record GUID, then click Run check.</div>
        ) : (
          <div className="pd-content">
            {data.effectiveAccess && (
              <section className="pd-section">
                <h3>Effective access</h3>
                <div className="pd-access-grid">
                  {ACCESS_LABELS.map(({ key, label }) => {
                    const granted = !!data.effectiveAccess?.[key];
                    return (
                      <div
                        key={key}
                        className={`pd-access-cell ${granted ? 'pd-access-on' : 'pd-access-off'}`}
                      >
                        <span className="pd-access-mark">{granted ? '✓' : '✕'}</span>
                        <span className="pd-access-label">{label}</span>
                      </div>
                    );
                  })}
                </div>
              </section>
            )}

            {data.diagnosis && data.diagnosis.length > 0 && (
              <section className="pd-section">
                <h3>Diagnosis</h3>
                <ul className="pd-diagnosis-list">
                  {data.diagnosis.map((line, i) => (
                    <li key={i}>{line}</li>
                  ))}
                </ul>
              </section>
            )}

            <div className="pd-cols">
              {data.record && (
                <section className="pd-section pd-col">
                  <h3>Record</h3>
                  <dl className="pd-dl">
                    <dt>Entity</dt>
                    <dd>{data.record.entityName}</dd>
                    <dt>Record</dt>
                    <dd>
                      {data.record.recordName ? (
                        <>
                          {data.record.recordName}
                          <div className="pd-mono-sub">{data.record.recordId}</div>
                        </>
                      ) : (
                        <code>{data.record.recordId}</code>
                      )}
                    </dd>
                    <dt>Owner</dt>
                    <dd>
                      {data.record.ownerName || '-'}
                      {data.record.ownerType && (
                        <span className="pd-tag">{data.record.ownerType}</span>
                      )}
                    </dd>
                    <dt>Owner BU</dt>
                    <dd>{data.record.ownerBusinessUnit || '-'}</dd>
                    <dt>Record BU</dt>
                    <dd>{data.record.recordBusinessUnit || '-'}</dd>
                  </dl>
                </section>
              )}

              {data.user && (
                <section className="pd-section pd-col">
                  <h3>Current user</h3>
                  <dl className="pd-dl">
                    <dt>Name</dt>
                    <dd>{data.user.fullName}</dd>
                    <dt>User ID</dt>
                    <dd>
                      <code>{data.user.userId}</code>
                    </dd>
                    <dt>BU</dt>
                    <dd>{data.user.businessUnitName || data.user.businessUnitId}</dd>
                    <dt>Disabled</dt>
                    <dd>{data.user.isDisabled ? 'Yes' : 'No'}</dd>
                    <dt>Teams</dt>
                    <dd>
                      {data.user.teams.length === 0 ? (
                        <span className="pd-muted">none</span>
                      ) : (
                        <div className="pd-tag-list">
                          {data.user.teams.map((t) => (
                            <span key={t.id} className="pd-tag">
                              {t.name}
                            </span>
                          ))}
                        </div>
                      )}
                    </dd>
                  </dl>
                </section>
              )}
            </div>

            {data.roles && data.roles.length > 0 && (
              <section className="pd-section">
                <h3>Roles &amp; privileges on {data.inputEntityName}</h3>
                <div className="pd-roles">
                  {data.roles.map((r) => (
                    <div key={r.roleId} className="pd-role">
                      <div className="pd-role-header">
                        <span className="pd-role-name">{r.roleName}</span>
                        {r.businessUnitName && (
                          <span className="pd-role-bu">{r.businessUnitName}</span>
                        )}
                      </div>
                      {r.privileges.length === 0 ? (
                        <div className="pd-muted">No privileges granted on this entity by this role.</div>
                      ) : (
                        <div className="pd-priv-grid">
                          {r.privileges.map((p, i) => (
                            <div key={i} className="pd-priv-cell">
                              <span className="pd-priv-name">{p.accessRight}</span>
                              <span className={`pd-priv-depth pd-priv-depth-${p.depthRank}`}>{p.depth}</span>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              </section>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

export default PrivilegeDebugger;
