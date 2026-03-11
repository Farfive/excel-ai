import React, { useState, useRef, useCallback, useEffect } from 'react';
import { WorkbookInfo } from '../../types';
import { useWorkbook } from '../../hooks/useWorkbook';
import { getHealth } from '../../services/api';
import { ExcelViewer } from '../ExcelViewer/ExcelViewer';
import { ChatSidebar } from '../ChatSidebar/ChatSidebar';
import styles from './App.module.css';

export function App(): React.ReactElement {
  const { workbookInfo, isUploading, uploadError, upload, fileName } = useWorkbook();
  const [leftWidth, setLeftWidth] = useState(62); // percent
  const [isDragging, setIsDragging] = useState(false);
  const [llmOnline, setLlmOnline] = useState(false);
  const [model, setModel] = useState('');
  const containerRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    getHealth()
      .then(h => { setLlmOnline(h.components.ollama.status === 'ok'); setModel(h.model); })
      .catch(() => setLlmOnline(false));
    const interval = setInterval(() => {
      getHealth()
        .then(h => { setLlmOnline(h.components.ollama.status === 'ok'); setModel(h.model); })
        .catch(() => setLlmOnline(false));
    }, 30000);
    return () => clearInterval(interval);
  }, []);

  const handleMouseDown = useCallback((e: React.MouseEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  useEffect(() => {
    if (!isDragging) return;
    const onMove = (e: MouseEvent) => {
      const container = containerRef.current;
      if (!container) return;
      const rect = container.getBoundingClientRect();
      const pct = ((e.clientX - rect.left) / rect.width) * 100;
      setLeftWidth(Math.min(80, Math.max(30, pct)));
    };
    const onUp = () => setIsDragging(false);
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
    return () => { window.removeEventListener('mousemove', onMove); window.removeEventListener('mouseup', onUp); };
  }, [isDragging]);

  const handleUploadClick = () => fileInputRef.current?.click();
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (f) upload(f);
    e.target.value = '';
  };

  return (
    <div className={styles.app}>
      {/* ── Top bar ── */}
      <div className={styles.topbar}>
        <div className={styles.topbar__left}>
          <div className={styles.topbar__logo}>
            <div className={styles['topbar__logo-mark']}>X</div>
            <span className={styles['topbar__logo-text']}>Excel AI</span>
          </div>
          {workbookInfo && (
            <>
              <div className={styles.topbar__divider} />
              <span className={styles.topbar__filename}>{fileName || workbookInfo.workbook_uuid.slice(0, 12) + '…'}</span>
              <span className={styles.topbar__badge}>{workbookInfo.cell_count.toLocaleString()} cells</span>
              <span className={styles.topbar__badge}>{workbookInfo.cluster_count} clusters</span>
              {workbookInfo.anomaly_count > 0 && (
                <span className={`${styles.topbar__badge} ${styles['topbar__badge--warn']}`}>
                  {workbookInfo.anomaly_count} anomalies
                </span>
              )}
            </>
          )}
        </div>
        <div className={styles.topbar__right}>
          {model && <span className={styles.topbar__model}>{model}</span>}
          <div className={styles.topbar__status}>
            <span className={`${styles.topbar__dot} ${llmOnline ? styles['topbar__dot--online'] : styles['topbar__dot--offline']}`} />
            <span>{llmOnline ? 'online' : 'offline'}</span>
          </div>
          <button className={styles['topbar__upload-btn']} onClick={handleUploadClick} type="button">
            <svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor">
              <path d="M8 1.5a.5.5 0 0 1 .5.5v9.793l2.646-2.647a.5.5 0 0 1 .708.708l-3.5 3.5a.5.5 0 0 1-.708 0l-3.5-3.5a.5.5 0 1 1 .708-.708L7.5 11.793V2a.5.5 0 0 1 .5-.5z" transform="rotate(180 8 8)"/>
              <path d="M.5 9.9a.5.5 0 0 1 .5.5v2.1A1.5 1.5 0 0 0 2.5 14h11a1.5 1.5 0 0 0 1.5-1.5v-2.1a.5.5 0 0 1 1 0v2.1a2.5 2.5 0 0 1-2.5 2.5h-11A2.5 2.5 0 0 1 0 12.5v-2.1a.5.5 0 0 1 .5-.5z"/>
            </svg>
            {workbookInfo ? 'Replace' : 'Upload .xlsx'}
          </button>
          <input ref={fileInputRef} type="file" accept=".xlsx" style={{ display: 'none' }} onChange={handleFileChange} />
        </div>
      </div>

      {/* ── Main workspace ── */}
      <div className={styles.workspace} ref={containerRef}>
        {/* Left: Excel Viewer */}
        <div className={styles.leftPanel} style={{ width: `${leftWidth}%` }}>
          <ExcelViewer
            workbookInfo={workbookInfo}
            isUploading={isUploading}
            uploadError={uploadError}
            onUpload={upload}
          />
        </div>

        {/* Resizer */}
        <div
          className={`${styles.resizer} ${isDragging ? styles['resizer--dragging'] : ''}`}
          onMouseDown={handleMouseDown}
        />

        {/* Right: Chat */}
        <div className={styles.rightPanel} style={{ width: `${100 - leftWidth}%` }}>
          <ChatSidebar
            workbookUuid={workbookInfo?.workbook_uuid ?? null}
          />
        </div>
      </div>
    </div>
  );
}
