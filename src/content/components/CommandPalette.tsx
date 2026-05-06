import React, { useEffect, useMemo, useRef, useState } from 'react';
import './CommandPalette.css';

export interface PaletteCommand {
  id: string;
  label: string;
  description?: string;
  keywords?: string;
  section?: string;
  isCustom?: boolean;
  run: () => void;
}

export interface CustomCommand {
  id: string;
  label: string;
  url: string;
  description?: string;
}

interface CommandPaletteProps {
  open: boolean;
  commands: PaletteCommand[];
  pinnedIds: string[];
  customCommands: CustomCommand[];
  keyBindings: Record<string, string>;
  onClose: () => void;
  onTogglePin: (id: string) => void;
  onSetBinding: (id: string, chord: string | null) => void;
  onAddCustomCommand: (cmd: { label: string; url: string; description?: string }) => void;
  onDeleteCustomCommand: (id: string) => void;
}

const fuzzyScore = (haystack: string, needle: string): number => {
  if (!needle) return 1;
  const h = haystack.toLowerCase();
  const n = needle.toLowerCase();
  if (h.includes(n)) return 100 - (h.indexOf(n) / Math.max(1, h.length)) * 30;
  let score = 0;
  let hi = 0;
  for (const ch of n) {
    const idx = h.indexOf(ch, hi);
    if (idx === -1) return 0;
    score += 1 - (idx - hi) * 0.05;
    hi = idx + 1;
  }
  return Math.max(score, 1);
};

const StarIcon: React.FC<{ filled: boolean }> = ({ filled }) => (
  <svg
    xmlns="http://www.w3.org/2000/svg"
    width="14"
    height="14"
    viewBox="0 0 24 24"
    fill={filled ? 'currentColor' : 'none'}
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden="true"
  >
    <polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2" />
  </svg>
);

const KeyboardIcon: React.FC = () => (
  <svg
    xmlns="http://www.w3.org/2000/svg"
    width="14"
    height="14"
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden="true"
  >
    <rect x="2" y="6" width="20" height="12" rx="2" />
    <path d="M6 10h.01" />
    <path d="M10 10h.01" />
    <path d="M14 10h.01" />
    <path d="M18 10h.01" />
    <path d="M6 14h12" />
  </svg>
);

const TrashIcon: React.FC = () => (
  <svg
    xmlns="http://www.w3.org/2000/svg"
    width="13"
    height="13"
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden="true"
  >
    <polyline points="3 6 5 6 21 6" />
    <path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6" />
    <path d="M10 11v6" />
    <path d="M14 11v6" />
  </svg>
);

const PlusIcon: React.FC = () => (
  <svg
    xmlns="http://www.w3.org/2000/svg"
    width="14"
    height="14"
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden="true"
  >
    <line x1="12" y1="5" x2="12" y2="19" />
    <line x1="5" y1="12" x2="19" y2="12" />
  </svg>
);

// Display-friendly version of a normalised chord (e.g. "ctrl+shift+d" → "Ctrl+Shift+D").
export const formatChord = (chord: string): string => {
  if (!chord) return '';
  return chord
    .split('+')
    .map((part) => {
      switch (part) {
        case 'ctrl':
          return 'Ctrl';
        case 'shift':
          return 'Shift';
        case 'alt':
          return 'Alt';
        case 'meta':
          return navigator.platform.toUpperCase().includes('MAC') ? '⌘' : 'Win';
        default:
          if (part.length === 1) return part.toUpperCase();
          return part.charAt(0).toUpperCase() + part.slice(1);
      }
    })
    .join('+');
};

// Normalise a KeyboardEvent into a stable chord string ("ctrl+shift+d").
// Returns null for pure-modifier or no-modifier-no-fn-key presses.
export const normaliseChord = (e: KeyboardEvent | React.KeyboardEvent): string | null => {
  const key = e.key;
  if (['Control', 'Shift', 'Alt', 'Meta'].includes(key)) return null;
  const parts: string[] = [];
  if (e.ctrlKey) parts.push('ctrl');
  if (e.shiftKey) parts.push('shift');
  if (e.altKey) parts.push('alt');
  if (e.metaKey) parts.push('meta');

  const isFn = /^F\d{1,2}$/.test(key);
  // Require at least one modifier, OR a function key on its own.
  if (parts.length === 0 && !isFn) return null;

  if (key.length === 1) {
    parts.push(key.toLowerCase());
  } else {
    parts.push(key);
  }
  return parts.join('+');
};

const RESERVED_BROWSER_CHORDS = new Set([
  'ctrl+t', 'ctrl+w', 'ctrl+n', 'ctrl+shift+t', 'ctrl+shift+n', 'ctrl+shift+w',
  'ctrl+l', 'ctrl+shift+i', 'ctrl+shift+j', 'ctrl+shift+c',
]);

const CommandPalette: React.FC<CommandPaletteProps> = ({
  open,
  commands,
  pinnedIds,
  customCommands,
  keyBindings,
  onClose,
  onTogglePin,
  onSetBinding,
  onAddCustomCommand,
  onDeleteCustomCommand,
}) => {
  const [query, setQuery] = useState('');
  const [activeIdx, setActiveIdx] = useState(0);
  const [showAddForm, setShowAddForm] = useState(false);
  const [newLabel, setNewLabel] = useState('');
  const [newUrl, setNewUrl] = useState('');
  const [newDesc, setNewDesc] = useState('');
  const [formError, setFormError] = useState('');
  const [bindingFor, setBindingFor] = useState<string | null>(null);
  const [bindingHint, setBindingHint] = useState<string>('');
  const inputRef = useRef<HTMLInputElement | null>(null);
  const listRef = useRef<HTMLDivElement | null>(null);

  const pinnedSet = useMemo(() => new Set(pinnedIds), [pinnedIds]);

  useEffect(() => {
    if (open) {
      setQuery('');
      setActiveIdx(0);
      setShowAddForm(false);
      setNewLabel('');
      setNewUrl('');
      setNewDesc('');
      setFormError('');
      setBindingFor(null);
      setBindingHint('');
      setTimeout(() => inputRef.current?.focus(), 0);
    }
  }, [open]);

  // While capturing a chord, intercept keydown at window level so the user can
  // press literally any chord without it being swallowed by the search input.
  useEffect(() => {
    if (!open || !bindingFor) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        e.stopPropagation();
        setBindingFor(null);
        setBindingHint('');
        return;
      }
      if (e.key === 'Backspace' || e.key === 'Delete') {
        e.preventDefault();
        e.stopPropagation();
        onSetBinding(bindingFor, null);
        setBindingFor(null);
        setBindingHint('');
        return;
      }
      const chord = normaliseChord(e);
      if (!chord) return;
      e.preventDefault();
      e.stopPropagation();
      if (RESERVED_BROWSER_CHORDS.has(chord)) {
        setBindingHint(`${formatChord(chord)} is reserved by Chrome — pick something else.`);
        return;
      }
      // Detect collisions with another command's binding
      const collidingId = Object.entries(keyBindings).find(
        ([id, c]) => c === chord && id !== bindingFor
      )?.[0];
      if (collidingId) {
        const colliding = commands.find((c) => c.id === collidingId);
        if (
          !window.confirm(
            `${formatChord(chord)} is already bound to "${colliding?.label || collidingId}". Re-bind to this command?`
          )
        ) {
          setBindingHint('Pick a different chord.');
          return;
        }
        onSetBinding(collidingId, null);
      }
      onSetBinding(bindingFor, chord);
      setBindingFor(null);
      setBindingHint('');
    };
    window.addEventListener('keydown', onKey, true);
    return () => window.removeEventListener('keydown', onKey, true);
  }, [open, bindingFor, keyBindings, commands, onSetBinding]);

  const filtered = useMemo(() => {
    const list = commands;
    if (!query.trim()) {
      const pinnedOrder = new Map(pinnedIds.map((id, i) => [id, i]));
      return [...list].sort((a, b) => {
        const aPinned = pinnedSet.has(a.id);
        const bPinned = pinnedSet.has(b.id);
        if (aPinned && bPinned) return (pinnedOrder.get(a.id) ?? 0) - (pinnedOrder.get(b.id) ?? 0);
        if (aPinned) return -1;
        if (bPinned) return 1;
        return 0;
      });
    }
    const q = query.trim();
    return list
      .map((c) => {
        const text = `${c.label} ${c.description || ''} ${c.keywords || ''} ${c.section || ''}`;
        let score = fuzzyScore(text, q);
        if (pinnedSet.has(c.id)) score += 10;
        return { cmd: c, score };
      })
      .filter((x) => x.score > 0)
      .sort((a, b) => b.score - a.score)
      .map((x) => x.cmd);
  }, [commands, query, pinnedSet, pinnedIds]);

  useEffect(() => {
    if (activeIdx >= filtered.length) setActiveIdx(0);
  }, [filtered, activeIdx]);

  useEffect(() => {
    if (!listRef.current) return;
    const el = listRef.current.querySelector<HTMLDivElement>(`[data-idx="${activeIdx}"]`);
    if (el) el.scrollIntoView({ block: 'nearest' });
  }, [activeIdx]);

  if (!open) return null;

  const runCommand = (cmd: PaletteCommand) => {
    onClose();
    setTimeout(() => {
      try {
        cmd.run();
      } catch (err) {
        console.error('[D365 Helper] Command failed:', err);
      }
    }, 0);
  };

  const handleKey = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (showAddForm || bindingFor) return;
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setActiveIdx((i) => Math.min(i + 1, Math.max(0, filtered.length - 1)));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setActiveIdx((i) => Math.max(i - 1, 0));
    } else if (e.key === 'Enter') {
      e.preventDefault();
      const cmd = filtered[activeIdx];
      if (cmd) runCommand(cmd);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      onClose();
    }
  };

  const handleSaveCustom = () => {
    const label = newLabel.trim();
    const url = newUrl.trim();
    if (!label) {
      setFormError('Label is required.');
      return;
    }
    let normalisedUrl = url;
    if (!/^https?:\/\//i.test(url)) {
      if (url.startsWith('/')) {
        normalisedUrl = window.location.origin + url;
      } else if (url) {
        normalisedUrl = 'https://' + url;
      } else {
        setFormError('URL is required.');
        return;
      }
    }
    try {
      new URL(normalisedUrl);
    } catch {
      setFormError('That doesn’t look like a valid URL.');
      return;
    }
    onAddCustomCommand({ label, url: normalisedUrl, description: newDesc.trim() || undefined });
    setShowAddForm(false);
    setNewLabel('');
    setNewUrl('');
    setNewDesc('');
    setFormError('');
  };

  return (
    <div className="cp-overlay" onMouseDown={onClose}>
      <div className="cp-panel" onMouseDown={(e) => e.stopPropagation()}>
        <div className="cp-input-row">
          <span className="cp-input-icon">⌘</span>
          <input
            ref={inputRef}
            className="cp-input"
            placeholder={bindingFor ? 'Capturing keyboard shortcut…' : 'Type a command...'}
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={handleKey}
            spellCheck={false}
            disabled={showAddForm || !!bindingFor}
          />
          <span className="cp-hint">Esc</span>
        </div>

        {bindingFor && (
          <div className="cp-binding-banner">
            <strong>Press a key combination</strong> for{' '}
            <em>{commands.find((c) => c.id === bindingFor)?.label || bindingFor}</em>.{' '}
            <span className="cp-binding-help">Backspace clears, Esc cancels.</span>
            {bindingHint && <div className="cp-binding-hint">{bindingHint}</div>}
          </div>
        )}

        {showAddForm ? (
          <div className="cp-add-form">
            <div className="cp-add-title">Add custom command</div>
            <label className="cp-add-field">
              <span>Label</span>
              <input
                type="text"
                value={newLabel}
                onChange={(e) => setNewLabel(e.target.value)}
                placeholder="e.g. Open Power Automate"
                autoFocus
              />
            </label>
            <label className="cp-add-field">
              <span>URL</span>
              <input
                type="text"
                value={newUrl}
                onChange={(e) => setNewUrl(e.target.value)}
                placeholder="https://make.powerautomate.com or /main.aspx?..."
              />
            </label>
            <label className="cp-add-field">
              <span>Description (optional)</span>
              <input
                type="text"
                value={newDesc}
                onChange={(e) => setNewDesc(e.target.value)}
                placeholder="Shown in the palette"
              />
            </label>
            {formError && <div className="cp-add-error">{formError}</div>}
            <div className="cp-add-actions">
              <button className="cp-btn cp-btn-secondary" onClick={() => setShowAddForm(false)}>
                Cancel
              </button>
              <button className="cp-btn cp-btn-primary" onClick={handleSaveCustom}>
                Add
              </button>
            </div>
          </div>
        ) : (
          <>
            <div ref={listRef} className="cp-list">
              {filtered.length === 0 ? (
                <div className="cp-empty">No matches</div>
              ) : (
                filtered.map((cmd, i) => {
                  const isPinned = pinnedSet.has(cmd.id);
                  const chord = keyBindings[cmd.id];
                  const isCapturing = bindingFor === cmd.id;
                  return (
                    <div
                      key={cmd.id}
                      data-idx={i}
                      className={`cp-item ${i === activeIdx ? 'cp-item-active' : ''} ${
                        isPinned ? 'cp-item-pinned' : ''
                      }`}
                      onMouseEnter={() => !bindingFor && setActiveIdx(i)}
                      onClick={() => !bindingFor && runCommand(cmd)}
                    >
                      <button
                        type="button"
                        className={`cp-pin-btn ${isPinned ? 'cp-pinned' : ''}`}
                        onClick={(e) => {
                          e.stopPropagation();
                          if (bindingFor) return;
                          onTogglePin(cmd.id);
                        }}
                        title={isPinned ? 'Unpin' : 'Pin to compact toolbar'}
                        aria-label={isPinned ? 'Unpin' : 'Pin'}
                      >
                        <StarIcon filled={isPinned} />
                      </button>
                      <div className="cp-item-main">
                        <div className="cp-item-label">{cmd.label}</div>
                        {cmd.description && <div className="cp-item-desc">{cmd.description}</div>}
                      </div>
                      {chord && !isCapturing && (
                        <kbd className="cp-chord">{formatChord(chord)}</kbd>
                      )}
                      {isCapturing && <kbd className="cp-chord cp-chord-capturing">…</kbd>}
                      <button
                        type="button"
                        className={`cp-bind-btn ${isCapturing ? 'cp-bind-active' : ''}`}
                        onClick={(e) => {
                          e.stopPropagation();
                          setBindingFor(isCapturing ? null : cmd.id);
                          setBindingHint('');
                        }}
                        title={chord ? 'Change shortcut' : 'Set keyboard shortcut'}
                        aria-label="Set shortcut"
                      >
                        <KeyboardIcon />
                      </button>
                      {cmd.section && <span className="cp-item-section">{cmd.section}</span>}
                      {cmd.isCustom && (
                        <button
                          type="button"
                          className="cp-delete-btn"
                          onClick={(e) => {
                            e.stopPropagation();
                            onDeleteCustomCommand(cmd.id);
                          }}
                          title="Delete custom command"
                          aria-label="Delete"
                        >
                          <TrashIcon />
                        </button>
                      )}
                    </div>
                  );
                })
              )}
            </div>
            <div className="cp-list-footer">
              <button
                type="button"
                className="cp-add-link"
                onClick={() => {
                  setShowAddForm(true);
                  setFormError('');
                }}
                title="Add a custom URL command"
              >
                <PlusIcon /> Add custom command
              </button>
            </div>
          </>
        )}

        <div className="cp-footer">
          <span><kbd>↑</kbd> <kbd>↓</kbd> navigate</span>
          <span><kbd>↵</kbd> run</span>
          <span><kbd>★</kbd> pin</span>
          <span><kbd>⌨</kbd> bind</span>
          <span><kbd>Esc</kbd> close</span>
        </div>
      </div>
    </div>
  );
};

export default CommandPalette;
