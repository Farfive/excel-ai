import React, { useState } from 'react';
import { Anomaly } from '../../types';
import { CellChip } from '../shared/CellChip';
import styles from './Anomalies.module.css';

interface AnomalyRowProps {
  anomaly: Anomaly;
}

export function AnomalyRow({ anomaly }: AnomalyRowProps): React.ReactElement {
  const [expanded, setExpanded] = useState(false);

  const severityLabel = anomaly.severity.toUpperCase();
  const severityClass = `anomaly__badge--${anomaly.severity}`;

  return (
    <div className={styles.anomalyRow}>
      <div
        className={styles.anomalyRow__header}
        onClick={() => setExpanded((v) => !v)}
        role="button"
        tabIndex={0}
        aria-expanded={expanded}
        aria-label={`${severityLabel} anomaly at ${anomaly.cell}`}
        onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') setExpanded((v) => !v); }}
      >
        <span
          className={`${styles['anomaly__badge']} ${styles[severityClass]}`}
          aria-label={`Severity: ${severityLabel}`}
        >
          {severityLabel}
        </span>
        <CellChip address={anomaly.cell} />
        <span className={styles.anomalyRow__value}>{String(anomaly.value ?? '')}</span>
        <span className={styles.anomalyRow__reason}>
          {anomaly.formula ? `Formula: ${anomaly.formula.slice(0, 30)}` : 'Numeric outlier'}
        </span>
      </div>

      <div className={`${styles.anomalyRow__expand} ${expanded ? styles['anomalyRow__expand--open'] : ''}`}>
        <div className={styles.anomalyRow__detail}>
          <div>cell: <span>{anomaly.cell}</span></div>
          {anomaly.formula && <div>formula: <span>{anomaly.formula}</span></div>}
          {anomaly.sheet_name && <div>sheet: <span>{anomaly.sheet_name}</span></div>}
          {anomaly.named_range && <div>named range: <span>{anomaly.named_range}</span></div>}
          <div>anomaly score: <span>{anomaly.anomaly_score.toFixed(4)}</span></div>
        </div>
      </div>
    </div>
  );
}
