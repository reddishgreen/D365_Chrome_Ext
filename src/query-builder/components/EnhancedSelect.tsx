import React, { useState, useRef, useEffect } from 'react';

interface Option {
  value: string;
  label: string;
  description?: string;
  group?: string;
}

interface EnhancedSelectProps {
  options: Option[];
  value: string;
  onChange: (value: string) => void;
  placeholder?: string;
  disabled?: boolean;
  searchable?: boolean;
  className?: string;
  width?: string;
  size?: 'small' | 'medium' | 'large';
}

const EnhancedSelect: React.FC<EnhancedSelectProps> = ({
  options,
  value,
  onChange,
  placeholder = 'Select...',
  disabled = false,
  searchable = true,
  className = '',
  width,
  size = 'medium'
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [activeIndex, setActiveIndex] = useState(-1);
  const containerRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const listRef = useRef<HTMLDivElement>(null);

  const selectedOption = options.find(opt => opt.value === value);

  const filteredOptions = searchable && searchTerm
    ? options.filter(opt =>
        opt.label.toLowerCase().includes(searchTerm.toLowerCase()) ||
        opt.value.toLowerCase().includes(searchTerm.toLowerCase()) ||
        opt.description?.toLowerCase().includes(searchTerm.toLowerCase())
      )
    : options;

  // Group options if they have a group property
  const groupedOptions = filteredOptions.reduce((acc, option) => {
    const group = option.group || 'default';
    if (!acc[group]) acc[group] = [];
    acc[group].push(option);
    return acc;
  }, {} as Record<string, Option[]>);

  const handleSelect = (optionValue: string) => {
    onChange(optionValue);
    setIsOpen(false);
    setSearchTerm('');
    setActiveIndex(-1);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      if (!isOpen) {
        setIsOpen(true);
      } else {
        setActiveIndex(prev => (prev < filteredOptions.length - 1 ? prev + 1 : prev));
      }
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setActiveIndex(prev => (prev > 0 ? prev - 1 : 0));
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (activeIndex >= 0 && activeIndex < filteredOptions.length) {
        handleSelect(filteredOptions[activeIndex].value);
      }
    } else if (e.key === 'Escape') {
      setIsOpen(false);
      setSearchTerm('');
    }
  };

  // Scroll active item into view
  useEffect(() => {
    if (activeIndex >= 0 && listRef.current) {
      const activeElement = listRef.current.children[activeIndex] as HTMLElement;
      if (activeElement) {
        activeElement.scrollIntoView({ block: 'nearest' });
      }
    }
  }, [activeIndex]);

  // Close on click outside
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
        setSearchTerm('');
      }
    };

    if (isOpen) {
      document.addEventListener('mousedown', handleClickOutside);
      return () => document.removeEventListener('mousedown', handleClickOutside);
    }
  }, [isOpen]);

  const sizeClasses = {
    small: 'enhanced-select--small',
    medium: 'enhanced-select--medium',
    large: 'enhanced-select--large'
  };

  return (
    <div
      ref={containerRef}
      className={`enhanced-select ${sizeClasses[size]} ${className} ${disabled ? 'enhanced-select--disabled' : ''} ${isOpen ? 'enhanced-select--open' : ''}`}
      style={{ width }}
    >
      <div
        className="enhanced-select__control"
        onClick={() => !disabled && setIsOpen(!isOpen)}
      >
        {searchable && isOpen ? (
          <input
            ref={inputRef}
            type="text"
            className="enhanced-select__search-input"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            onKeyDown={handleKeyDown}
            placeholder={placeholder}
            autoFocus
          />
        ) : (
          <div className="enhanced-select__value">
            {selectedOption ? selectedOption.label : placeholder}
          </div>
        )}
        <div className="enhanced-select__arrow">
          <svg width="12" height="12" viewBox="0 0 12 12" fill="currentColor">
            <path d="M2 4l4 4 4-4" stroke="currentColor" strokeWidth="1.5" fill="none" strokeLinecap="round" strokeLinejoin="round"/>
          </svg>
        </div>
      </div>

      {isOpen && (
        <div className="enhanced-select__menu">
          <div className="enhanced-select__menu-list" ref={listRef}>
            {filteredOptions.length === 0 ? (
              <div className="enhanced-select__no-options">No options found</div>
            ) : (
              Object.entries(groupedOptions).map(([group, groupOptions]) => (
                <div key={group} className="enhanced-select__group">
                  {group !== 'default' && (
                    <div className="enhanced-select__group-label">{group}</div>
                  )}
                  {groupOptions.map((option, index) => {
                    const globalIndex = filteredOptions.indexOf(option);
                    return (
                      <div
                        key={option.value}
                        className={`enhanced-select__option ${
                          option.value === value ? 'enhanced-select__option--selected' : ''
                        } ${
                          globalIndex === activeIndex ? 'enhanced-select__option--active' : ''
                        }`}
                        onClick={() => handleSelect(option.value)}
                        onMouseEnter={() => setActiveIndex(globalIndex)}
                      >
                        <div className="enhanced-select__option-label">{option.label}</div>
                        {option.description && (
                          <div className="enhanced-select__option-description">{option.description}</div>
                        )}
                        {option.value === value && (
                          <div className="enhanced-select__option-checkmark">✓</div>
                        )}
                      </div>
                    );
                  })}
                </div>
              ))
            )}
          </div>
        </div>
      )}
    </div>
  );
};

export default EnhancedSelect;
