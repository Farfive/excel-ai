import React, { useState } from 'react';
import { Anomaly } from '../../types';
import { getAnomalies } from '../../services/api';
import { AnomalyRow } from './AnomalyRow';
import styles from './Anomalies.module.css';

interface AnomaliesPanelProps {
  workbookUuid: string | null;
}

export function AnomaliesPanel({ workbookUuid }: AnomaliesPanelProps): React.ReactElement {
  const [anomalies, setAnomalies] = useState<Anomaly[]>([]);
  const [scanning, setScanning] = useState(false);
  const [scanned, setScanned] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleScan = async () => {
    if (!workbookUuid) return;
    setScanning(true);
    setError(null);
    try {
      const results = await getAnomalies(workbookUuid);
      const sorted = [...results].sort((a, b) => {
        const order = { high: 0, medium: 1, low: 2 };
        return order[a.severity] - order[b.severity];
      });
      setAnomalies(sorted);
      setScanned(true);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : 'Scan failed');
    } finally {
      setScanning(false);
    }
  };

  return (
    <div className={styles.anomaliesPanel}>
      <div className={styles.anomaliesPanel__header}>
        <button
          className={styles['anomaliesPanel__scan-btn']}
          onClick={handleScan}
          disabled={scanning || !workbookUuid}
          type="button"
          aria-label="Scan for anomalies"
        >
          {scanning ? 'Scanning…' : 'Scan for Anomalies'}
        </button>
        {scanning && (
          <div className={styles.anomaliesPanel__progress}>
            <div className={styles['anomaliesPanel__progress-bar']} />
          </div>
        )}
        {error && (
          <div style={{ color: 'var(--accent)', fontSize: 11, marginTop: 6 }}>{error}</div>
        )}
      </div>

      <div className={styles.anomaliesPanel__list}>
        {scanned && anomalies.length === 0 && (
          <div className={styles.anomaliesPanel__empty}>
            <span className={styles.anomaliesPanel__empty_icon} aria-hidden="true">✓</span>
            No anomalies detected
          </div>
        )}
        {!scanned && !scanning && (
          <div className={styles.anomaliesPanel__empty} style={{ color: 'var(--text-muted)' }}>
            Run scan to detect anomalies
          </div>
        )}
        {anomalies.map((anomaly) => (
          <AnomalyRow key={anomaly.cell} anomaly={anomaly} />
        ))}
      </div>
    </div>
  );
}
