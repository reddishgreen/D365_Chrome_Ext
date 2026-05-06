import React from 'react';

interface CloseIconProps {
  size?: number;
  strokeWidth?: number;
}

/**
 * Standard close (X) icon used by all modal headers.
 * Plain SVG — no transform/rotate hover effects.
 */
const CloseIcon: React.FC<CloseIconProps> = ({ size = 16, strokeWidth = 2 }) => (
  <svg
    xmlns="http://www.w3.org/2000/svg"
    width={size}
    height={size}
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth={strokeWidth}
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden="true"
    focusable="false"
  >
    <line x1="6" y1="6" x2="18" y2="18" />
    <line x1="18" y1="6" x2="6" y2="18" />
  </svg>
);

export default CloseIcon;
