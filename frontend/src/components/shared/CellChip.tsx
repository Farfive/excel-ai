import React from 'react';
import { navigateTo } from '../../services/excel';
import styles from './shared.module.css';

interface CellChipProps {
  address: string;
  highlight?: boolean;
}

export function CellChip({ address, highlight }: CellChipProps): React.ReactElement {
  const handleClick = () => {
    navigateTo(address).catch((e) => console.error('Navigate failed:', e));
  };

  return (
    <button
      className={`${styles.cellChip} ${highlight ? styles['cellChip--highlight'] : ''}`}
      onClick={handleClick}
      role="button"
      aria-label={`Navigate to ${address}`}
      type="button"
    >
      {address}
    </button>
  );
}
