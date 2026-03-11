import React, { useRef, useState } from 'react';
import { WorkbookInfo } from '../../types';
import styles from './shared.module.css';

interface UploadZoneProps {
  workbookInfo: WorkbookInfo | null;
  isUploading: boolean;
  uploadError: string | null;
  onUpload: (file: File) => void;
}

export function UploadZone({ workbookInfo, isUploading, uploadError, onUpload }: UploadZoneProps): React.ReactElement {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);

  const handleFile = (file: File) => {
    if (file.name.endsWith('.xlsx')) {
      onUpload(file);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) handleFile(file);
  };

  if (workbookInfo) {
    return (
      <div className={styles.uploadZone}>
        <div className={styles.uploadZone__filename}>{workbookInfo.workbook_uuid.slice(0, 8)}…</div>
        <div className={styles.uploadZone__stats}>
          {workbookInfo.cell_count} cells · {workbookInfo.cluster_count} clusters · {workbookInfo.anomaly_count} anomalies
        </div>
        <div className={styles.uploadZone__stats}>
          Processed in {workbookInfo.processing_time_ms}ms
        </div>
      </div>
    );
  }

  return (
    <div
      className={`${styles.uploadZone} ${dragging ? styles['uploadZone--dragging'] : ''}`}
      onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={handleDrop}
      onClick={() => inputRef.current?.click()}
      role="button"
      tabIndex={0}
      aria-label="Upload Excel workbook"
      onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') inputRef.current?.click(); }}
    >
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx"
        style={{ display: 'none' }}
        onChange={handleChange}
        aria-hidden="true"
      />
      {isUploading ? (
        <div className={styles.spinner} aria-label="Uploading..." />
      ) : (
        <>
          <span className={styles.uploadZone__label}>Drop .xlsx file or click to upload</span>
          <span className={styles.uploadZone__hint}>.xlsx only</span>
        </>
      )}
      {uploadError && (
        <span style={{ color: 'var(--accent)', fontSize: 11, marginTop: 4 }}>{uploadError}</span>
      )}
    </div>
  );
}
