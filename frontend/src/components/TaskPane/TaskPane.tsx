import React, { useState, useEffect } from 'react';
import { useWorkbook } from '../../hooks/useWorkbook';
import { useExcelEvents } from '../../hooks/useExcelEvents';
import { ChatPanel } from '../Chat/ChatPanel';
import { AnomaliesPanel } from '../Anomalies/AnomaliesPanel';
import { AnalysisPanel } from '../Analysis/AnalysisPanel';
import { UploadZone } from '../shared/UploadZone';
import { StatusBar } from '../shared/StatusBar';
import { getHealth } from '../../services/api';
import styles from './TaskPane.module.css';

type Tab = 'CHAT' | 'ANOMALIES' | 'ANALYSIS';

export function TaskPane(): React.ReactElement {
  const { workbookInfo, isUploading, uploadError, upload } = useWorkbook();
  const [activeTab, setActiveTab] = useState<Tab>('CHAT');
  const [ollamaOnline, setOllamaOnline] = useState(false);

  useExcelEvents(workbookInfo?.workbook_uuid ?? null);

  useEffect(() => {
    getHealth()
      .then((h) => setOllamaOnline(h.components.ollama.status === 'ok'))
      .catch(() => setOllamaOnline(false));
  }, []);

  return (
    <div className={styles.taskPane}>
      <div className={styles.taskPane__header}>
        <span className={styles.taskPane__title}>Excel AI</span>
        <div className={styles.taskPane__health} aria-label={ollamaOnline ? 'Ollama online' : 'Ollama offline'}>
          <span
            className={`${styles.taskPane__health_dot} ${ollamaOnline ? styles['taskPane__health-dot--online'] : styles['taskPane__health-dot--offline']}`}
          />
        </div>
      </div>

      <div className={styles.taskPane__upload}>
        <UploadZone
          workbookInfo={workbookInfo}
          isUploading={isUploading}
          uploadError={uploadError}
          onUpload={upload}
        />
      </div>

      {workbookInfo && <StatusBar workbookInfo={workbookInfo} />}

      <div className={styles.taskPane__tabs}>
        <button
          className={`${styles.taskPane__tab} ${activeTab === 'CHAT' ? styles['taskPane__tab--active'] : ''}`}
          onClick={() => setActiveTab('CHAT')}
          type="button"
          aria-selected={activeTab === 'CHAT'}
          role="tab"
        >
          Chat
        </button>
        <button
          className={`${styles.taskPane__tab} ${activeTab === 'ANOMALIES' ? styles['taskPane__tab--active'] : ''}`}
          onClick={() => setActiveTab('ANOMALIES')}
          type="button"
          aria-selected={activeTab === 'ANOMALIES'}
          role="tab"
        >
          Anomalies
        </button>
        <button
          className={`${styles.taskPane__tab} ${activeTab === 'ANALYSIS' ? styles['taskPane__tab--active'] : ''}`}
          onClick={() => setActiveTab('ANALYSIS')}
          type="button"
          aria-selected={activeTab === 'ANALYSIS'}
          role="tab"
        >
          Analysis
        </button>
      </div>

      <div className={styles.taskPane__content}>
        {activeTab === 'CHAT' && (
          <ChatPanel workbookUuid={workbookInfo?.workbook_uuid ?? null} />
        )}
        {activeTab === 'ANOMALIES' && (
          <AnomaliesPanel workbookUuid={workbookInfo?.workbook_uuid ?? null} />
        )}
        {activeTab === 'ANALYSIS' && (
          <AnalysisPanel workbookUuid={workbookInfo?.workbook_uuid ?? null} />
        )}
      </div>
    </div>
  );
}
