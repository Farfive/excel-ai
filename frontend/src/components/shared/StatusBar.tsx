import React, { useEffect, useState } from 'react';
import { getHealth } from '../../services/api';
import { WorkbookInfo } from '../../types';
import styles from './shared.module.css';

interface StatusBarProps {
  workbookInfo: WorkbookInfo;
}

export function StatusBar({ workbookInfo }: StatusBarProps): React.ReactElement {
  const [ollamaOnline, setOllamaOnline] = useState(false);
  const [model, setModel] = useState('');

  useEffect(() => {
    getHealth()
      .then((h) => {
        setOllamaOnline(h.components.ollama.status === 'ok');
        setModel(h.model);
      })
      .catch(() => setOllamaOnline(false));

    const interval = setInterval(() => {
      getHealth()
        .then((h) => {
          setOllamaOnline(h.components.ollama.status === 'ok');
          setModel(h.model);
        })
        .catch(() => setOllamaOnline(false));
    }, 30000);

    return () => clearInterval(interval);
  }, []);

  return (
    <div className={styles.statusBar} role="status" aria-label="Workbook status">
      <span
        className={`${styles.statusBar__dot} ${ollamaOnline ? styles['statusBar__dot--online'] : styles['statusBar__dot--offline']}`}
        aria-label={ollamaOnline ? 'Ollama online' : 'Ollama offline'}
      />
      <span>
        {workbookInfo.cell_count} cells · {workbookInfo.cluster_count} clusters
        {model ? ` · ${model}` : ''}
      </span>
    </div>
  );
}
