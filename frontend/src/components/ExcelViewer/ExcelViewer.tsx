import React, { useState, useRef, useEffect } from 'react';
import { WorkbookInfo } from '../../types';
import { getAnomalies, getWorkbookGrid, revertChanges, downloadWorkbook, getAuditTrail, AuditChange, diffWorkbooks } from '../../services/api';
import { onGridRefresh, onGridDiff, CellChange } from '../../services/gridBus';
import { AnalysisPanel } from '../Analysis/AnalysisPanel';
import styles from './ExcelViewer.module.css';

interface RawCell {
  v: string;
  f: boolean;
  t: string;
  nr?: string;
}

interface SheetData {
  name: string;
  rows: RawCell[][];
  colHeaders: string[];
}

interface ExcelViewerProps {
  workbookInfo: WorkbookInfo | null;
  isUploading: boolean;
  uploadError: string | null;
  onUpload: (f: File) => void;
}

type ViewTab = 'grid' | 'analysis' | 'audit' | 'compare';

interface DiffCellItem { cell: string; sheet: string; type: string; old_value: unknown; new_value: unknown; delta_pct: number | null; impact: number; }
interface DiffSheetItem { sheet: string; added: number; removed: number; modified: number; formula_changes: number; value_changes: number; max_delta_pct: number | null; }
interface DiffResult { summary: string; total_changes: number; sheets_added: string[]; sheets_removed: string[]; sheet_diffs: DiffSheetItem[]; cell_diffs: DiffCellItem[]; high_impact: DiffCellItem[]; }

function isNumeric(v: string): boolean {
  if (!v) return false;
  return !isNaN(Number(v.replace(/[,$%]/g, '')));
}

function formatDiffVal(v: unknown): string {
  if (v === null || v === undefined) return '—';
  if (typeof v === 'number') return Number.isInteger(v) ? String(v) : v.toFixed(4).replace(/\.?0+$/, '');
  return String(v);
}

export function ExcelViewer({ workbookInfo, isUploading, uploadError, onUpload }: ExcelViewerProps): React.ReactElement {
  const [dragging, setDragging] = useState(false);
  const [activeSheet, setActiveSheet] = useState(0);
  const [viewTab, setViewTab] = useState<ViewTab>('grid');
  const [sheets, setSheets] = useState<SheetData[]>([]);
  const [loadingGrid, setLoadingGrid] = useState(false);
  const [anomalyCells, setAnomalyCells] = useState<Set<string>>(new Set());
  const [pendingChanges, setPendingChanges] = useState<CellChange[]>([]);
  const [auditChanges, setAuditChanges] = useState<AuditChange[]>([]);
  const [loadingAudit, setLoadingAudit] = useState(false);
  const [diffResult, setDiffResult] = useState<DiffResult | null>(null);
  const [loadingDiff, setLoadingDiff] = useState(false);
  const [diffError, setDiffError] = useState<string | null>(null);
  const diffInputRef = useRef<HTMLInputElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const prevUuid = useRef<string | null>(null);

  const pendingMap = useRef<Map<string, CellChange>>(new Map());
  useEffect(() => {
    const m = new Map<string, CellChange>();
    pendingChanges.forEach(c => m.set(c.cell, c));
    pendingMap.current = m;
  }, [pendingChanges]);

  const fetchGrid = (uuid: string) => {
    setLoadingGrid(true);
    getWorkbookGrid(uuid)
      .then(data => setSheets(data.sheets || []))
      .catch(() => {})
      .finally(() => setLoadingGrid(false));
  };

  useEffect(() => {
    if (!workbookInfo) { setSheets([]); setAnomalyCells(new Set()); setPendingChanges([]); return; }
    if (workbookInfo.workbook_uuid === prevUuid.current) return;
    prevUuid.current = workbookInfo.workbook_uuid;
    setActiveSheet(0);
    setPendingChanges([]);

    const uuid = workbookInfo.workbook_uuid;

    getAnomalies(uuid)
      .then(anomalies => {
        const cellSet = new Set(
          anomalies.map((a: any) =>
            a.sheet_name ? `${a.sheet_name}!${a.cell}` : a.cell
          )
        );
        setAnomalyCells(cellSet);
      })
      .catch(() => {});

    fetchGrid(uuid);
  }, [workbookInfo]);

  // Subscribe to gridBus: refresh grid when AI finishes writing
  useEffect(() => {
    const unsub = onGridRefresh(() => {
      if (workbookInfo) fetchGrid(workbookInfo.workbook_uuid);
    });
    return unsub;
  }, [workbookInfo]);

  // Subscribe to gridBus: receive diff data for pending changes highlight
  useEffect(() => {
    const unsub = onGridDiff((changes: CellChange[]) => {
      setPendingChanges(prev => [...prev, ...changes]);
    });
    return unsub;
  }, []);

  const handleAccept = () => {
    setPendingChanges([]);
  };

  const handleReject = async () => {
    if (!workbookInfo || pendingChanges.length === 0) return;
    try {
      await revertChanges(workbookInfo.workbook_uuid, pendingChanges);
      fetchGrid(workbookInfo.workbook_uuid);
    } catch (e) {
      console.error('Revert failed:', e);
    }
    setPendingChanges([]);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragging(false);
    const f = e.dataTransfer.files[0];
    if (f?.name.endsWith('.xlsx')) onUpload(f);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const f = e.target.files?.[0];
    if (f) onUpload(f);
    e.target.value = '';
  };

  // ── No workbook: drop zone ──
  if (!workbookInfo && !isUploading) {
    return (
      <div
        className={`${styles.dropzone} ${dragging ? styles['dropzone--dragging'] : ''}`}
        onDragOver={e => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={handleDrop}
        onClick={() => inputRef.current?.click()}
        role="button"
        tabIndex={0}
        onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') inputRef.current?.click(); }}
      >
        <input ref={inputRef} type="file" accept=".xlsx" style={{ display: 'none' }} onChange={handleFileChange} />
        <svg className={styles.dropzone__icon} width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
          <polyline points="14 2 14 8 20 8"/>
          <line x1="12" y1="18" x2="12" y2="12"/>
          <line x1="9" y1="15" x2="15" y2="15"/>
        </svg>
        <div className={styles.dropzone__title}>Drop your Excel file here</div>
        <div className={styles.dropzone__sub}>or click to browse · .xlsx only</div>
        {uploadError && <div className={styles.dropzone__error}>{uploadError}</div>}
      </div>
    );
  }

  if (isUploading) {
    return (
      <div className={styles.dropzone}>
        <div className={styles.dropzone__spinner} />
        <div className={styles.dropzone__title}>Processing workbook…</div>
        <div className={styles.dropzone__sub}>Building dependency graph and embeddings</div>
      </div>
    );
  }

  const currentSheet = sheets[activeSheet];

  return (
    <div className={styles.viewer}>
      {/* Toolbar */}
      <div className={styles.toolbar}>
        <span className={styles.toolbar__label}>Workbook</span>
        <span className={styles.toolbar__badge}>{workbookInfo!.cell_count.toLocaleString()} cells</span>
        <span className={styles.toolbar__badge}>{workbookInfo!.cluster_count} clusters</span>
        {workbookInfo!.anomaly_count > 0 && (
          <span className={`${styles.toolbar__badge} ${styles['toolbar__badge--accent']}`}>
            {workbookInfo!.anomaly_count} anomalies
          </span>
        )}
        <button
          className={styles.toolbar__download}
          onClick={() => downloadWorkbook(workbookInfo!.workbook_uuid)}
          type="button"
          title="Download .xlsx with current changes"
        >
          <svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor">
            <path d="M8 1.5a.5.5 0 0 1 .5.5v9.793l2.646-2.647a.5.5 0 0 1 .708.708l-3.5 3.5a.5.5 0 0 1-.708 0l-3.5-3.5a.5.5 0 1 1 .708-.708L7.5 11.793V2a.5.5 0 0 1 .5-.5z"/>
            <path d="M.5 9.9a.5.5 0 0 1 .5.5v2.1A1.5 1.5 0 0 0 2.5 14h11a1.5 1.5 0 0 0 1.5-1.5v-2.1a.5.5 0 0 1 1 0v2.1a2.5 2.5 0 0 1-2.5 2.5h-11A2.5 2.5 0 0 1 0 12.5v-2.1a.5.5 0 0 1 .5-.5z"/>
          </svg>
          Download .xlsx
        </button>
      </div>

      {/* View tabs */}
      <div className={styles.analysisTabs}>
        <button
          className={`${styles.analysisTab} ${viewTab === 'grid' ? styles['analysisTab--active'] : ''}`}
          onClick={() => setViewTab('grid')}
          type="button"
        >Grid</button>
        <button
          className={`${styles.analysisTab} ${viewTab === 'analysis' ? styles['analysisTab--active'] : ''}`}
          onClick={() => setViewTab('analysis')}
          type="button"
        >Analysis</button>
        <button
          className={`${styles.analysisTab} ${viewTab === 'audit' ? styles['analysisTab--active'] : ''}`}
          onClick={() => {
            setViewTab('audit');
            if (workbookInfo) {
              setLoadingAudit(true);
              getAuditTrail(workbookInfo.workbook_uuid)
                .then(r => setAuditChanges(r.changes))
                .catch(() => {})
                .finally(() => setLoadingAudit(false));
            }
          }}
          type="button"
        >Audit Trail</button>
        <button
          className={`${styles.analysisTab} ${viewTab === 'compare' ? styles['analysisTab--active'] : ''}`}
          onClick={() => setViewTab('compare')}
          type="button"
        >Compare</button>
      </div>

      {viewTab === 'compare' ? (
        <div className={styles.auditWrap} style={{ padding: 16 }}>
          <input
            ref={diffInputRef}
            type="file"
            accept=".xlsx"
            style={{ display: 'none' }}
            onChange={async (e) => {
              const f = e.target.files?.[0];
              if (!f || !workbookInfo) return;
              setLoadingDiff(true);
              setDiffError(null);
              setDiffResult(null);
              try {
                const res = await diffWorkbooks(workbookInfo.workbook_uuid, f);
                setDiffResult(res as DiffResult);
              } catch (err) {
                setDiffError(err instanceof Error ? err.message : 'Diff failed');
              } finally {
                setLoadingDiff(false);
                if (diffInputRef.current) diffInputRef.current.value = '';
              }
            }}
          />
          <div style={{ marginBottom: 12 }}>
            <button
              className={styles.toolbar__download}
              onClick={() => diffInputRef.current?.click()}
              disabled={loadingDiff}
              type="button"
              style={{ fontSize: 12, padding: '6px 14px' }}
            >
              {loadingDiff ? 'Comparing…' : 'Upload second .xlsx to compare'}
            </button>
          </div>
          {diffError && (
            <div style={{ color: 'var(--accent)', fontFamily: 'var(--font-mono)', fontSize: 11, marginBottom: 8 }}>{diffError}</div>
          )}
          {diffResult && (
            <div>
              <div style={{ fontFamily: 'var(--font-mono)', fontSize: 12, color: 'var(--text-primary)', fontWeight: 600, marginBottom: 8 }}>{diffResult.summary}</div>
              {diffResult.sheets_added.length > 0 && (
                <div style={{ fontFamily: 'var(--font-mono)', fontSize: 11, color: 'var(--green)', marginBottom: 4 }}>+ Sheets added: {diffResult.sheets_added.join(', ')}</div>
              )}
              {diffResult.sheets_removed.length > 0 && (
                <div style={{ fontFamily: 'var(--font-mono)', fontSize: 11, color: '#f87171', marginBottom: 4 }}>- Sheets removed: {diffResult.sheets_removed.join(', ')}</div>
              )}
              {diffResult.sheet_diffs.length > 0 && (
                <table className={styles.auditTable} style={{ marginBottom: 12 }}>
                  <thead><tr><th>Sheet</th><th>Added</th><th>Removed</th><th>Modified</th><th>Formula Δ</th><th>Value Δ</th><th>Max Δ%</th></tr></thead>
                  <tbody>
                    {diffResult.sheet_diffs.map(sd => (
                      <tr key={sd.sheet}>
                        <td>{sd.sheet}</td>
                        <td style={{ color: 'var(--green)' }}>{sd.added || '—'}</td>
                        <td style={{ color: '#f87171' }}>{sd.removed || '—'}</td>
                        <td>{sd.modified || '—'}</td>
                        <td>{sd.formula_changes || '—'}</td>
                        <td>{sd.value_changes || '—'}</td>
                        <td>{sd.max_delta_pct != null ? `${sd.max_delta_pct}%` : '—'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
              {diffResult.high_impact.length > 0 && (
                <>
                  <div style={{ fontFamily: 'var(--font-mono)', fontSize: 11, fontWeight: 600, color: 'var(--accent)', marginBottom: 4 }}>High-Impact Changes ({diffResult.high_impact.length})</div>
                  <table className={styles.auditTable} style={{ marginBottom: 12 }}>
                    <thead><tr><th>Cell</th><th>Type</th><th>Old</th><th>New</th><th>Δ%</th><th>Impact</th></tr></thead>
                    <tbody>
                      {diffResult.high_impact.map(cd => (
                        <tr key={cd.cell}>
                          <td>{cd.cell}</td>
                          <td>{cd.type}</td>
                          <td className={styles.auditOld}>{formatDiffVal(cd.old_value)}</td>
                          <td className={styles.auditNew}>{formatDiffVal(cd.new_value)}</td>
                          <td>{cd.delta_pct != null ? `${cd.delta_pct}%` : '—'}</td>
                          <td>{cd.impact.toFixed(1)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </>
              )}
              {diffResult.cell_diffs.length > 0 && (
                <>
                  <div style={{ fontFamily: 'var(--font-mono)', fontSize: 11, fontWeight: 600, color: 'var(--text-secondary)', marginBottom: 4 }}>All Changes ({diffResult.total_changes}, showing top {Math.min(diffResult.cell_diffs.length, 500)})</div>
                  <table className={styles.auditTable}>
                    <thead><tr><th>Cell</th><th>Type</th><th>Old</th><th>New</th><th>Δ%</th></tr></thead>
                    <tbody>
                      {diffResult.cell_diffs.slice(0, 200).map(cd => (
                        <tr key={cd.cell}>
                          <td>{cd.cell}</td>
                          <td>{cd.type}</td>
                          <td className={styles.auditOld}>{formatDiffVal(cd.old_value)}</td>
                          <td className={styles.auditNew}>{formatDiffVal(cd.new_value)}</td>
                          <td>{cd.delta_pct != null ? `${cd.delta_pct}%` : '—'}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </>
              )}
            </div>
          )}
        </div>
      ) : viewTab === 'audit' ? (
        <div className={styles.auditWrap}>
          {loadingAudit ? (
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: 120, color: 'var(--text-secondary)', fontFamily: 'var(--font-mono)', fontSize: 12 }}>
              Loading audit trail…
            </div>
          ) : auditChanges.length === 0 ? (
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: 120, color: 'var(--text-muted)', fontFamily: 'var(--font-mono)', fontSize: 12 }}>
              No changes recorded yet.
            </div>
          ) : (
            <table className={styles.auditTable}>
              <thead>
                <tr>
                  <th>#</th>
                  <th>Time</th>
                  <th>Cell</th>
                  <th>Old</th>
                  <th>New</th>
                  <th>Action</th>
                  <th>By</th>
                </tr>
              </thead>
              <tbody>
                {auditChanges.map((c, i) => (
                  <tr key={c.id}>
                    <td>{i + 1}</td>
                    <td>{c.timestamp.slice(11, 19)}</td>
                    <td>{c.cell}</td>
                    <td className={styles.auditOld}>{formatDiffVal(c.old_value)}</td>
                    <td className={styles.auditNew}>{formatDiffVal(c.new_value)}</td>
                    <td style={{ maxWidth: 200, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{c.reason}</td>
                    <td>{c.approved_by}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      ) : viewTab === 'analysis' ? (
        <div style={{ flex: 1, overflow: 'auto' }}>
          <AnalysisPanel workbookUuid={workbookInfo?.workbook_uuid ?? null} />
        </div>
      ) : (
        <>
          {/* Sheet tabs */}
          {sheets.length > 0 && (
            <div className={styles.sheetTabs}>
              {sheets.map((s, i) => (
                <button
                  key={s.name}
                  className={`${styles.sheetTab} ${i === activeSheet ? styles['sheetTab--active'] : ''}`}
                  onClick={() => setActiveSheet(i)}
                  type="button"
                >{s.name}</button>
              ))}
            </div>
          )}

          {/* Grid */}
          <div className={styles.gridWrap}>
            {loadingGrid ? (
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100%', gap: 10, color: 'var(--text-secondary)', fontFamily: 'var(--font-mono)', fontSize: 12 }}>
                <div className={styles.dropzone__spinner} style={{ width: 16, height: 16 }} />
                Loading grid data…
              </div>
            ) : currentSheet ? (
              <table className={styles.grid}>
                <thead>
                  <tr>
                    <th className={styles.grid__rowNum}>#</th>
                    {currentSheet.colHeaders.map(h => (
                      <th key={h}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {currentSheet.rows.map((row, r) => (
                    <tr key={r}>
                      <td className={styles.grid__rowNum}>{r + 1}</td>
                      {row.map((cell, c) => {
                        const colLtr = currentSheet.colHeaders[c] || '';
                        const addr = `${currentSheet.name}!${colLtr}${r + 1}`;
                        const isAnomaly = anomalyCells.has(addr);
                        const pending = pendingMap.current.get(addr);
                        const isPending = !!pending;
                        const isEmpty = !cell.v && !isPending;
                        const isFormula = cell.f;
                        const isNum = isNumeric(cell.v);
                        const isHeader = r === 0 && !isFormula && !isEmpty;
                        const cls = [
                          styles.grid__cell,
                          isFormula ? styles['grid__cell--formula'] : styles['grid__cell--hardcoded'],
                          isHeader ? styles['grid__cell--header'] : '',
                          isNum ? styles['grid__cell--number'] : '',
                          isEmpty ? styles['grid__cell--empty'] : '',
                          isAnomaly ? styles['grid__cell--anomaly'] : '',
                          isPending ? styles['grid__cell--pending'] : '',
                        ].filter(Boolean).join(' ');
                        return (
                          <td key={c} className={cls} title={isPending ? `${formatDiffVal(pending.before)} → ${formatDiffVal(pending.after)}` : (cell.nr ? `Named: ${cell.nr}` : cell.v)}>
                            {isPending ? (
                              <>
                                <span className={styles['cellOld']}>{formatDiffVal(pending.before)}</span>
                                <span className={styles['cellNew']}>{formatDiffVal(pending.after)}</span>
                              </>
                            ) : cell.v}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : (
              <div style={{ padding: 32, textAlign: 'center', color: 'var(--text-secondary)', fontFamily: 'var(--font-mono)', fontSize: 12 }}>
                No data
              </div>
            )}
          </div>

          {/* Pending changes bar */}
          {pendingChanges.length > 0 && (
            <div className={styles.pendingBar}>
              <span className={styles.pendingBar__label}>
                <span className={styles.pendingBar__count}>{pendingChanges.length}</span> cell{pendingChanges.length !== 1 ? 's' : ''} changed by AI
              </span>
              <button className={styles.pendingBar__accept} onClick={handleAccept} type="button">
                <svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor"><path d="M13.854 3.646a.5.5 0 0 1 0 .708l-7 7a.5.5 0 0 1-.708 0l-3.5-3.5a.5.5 0 1 1 .708-.708L6.5 10.293l6.646-6.647a.5.5 0 0 1 .708 0z"/></svg>
                Accept
              </button>
              <button className={styles.pendingBar__reject} onClick={handleReject} type="button">
                <svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor"><path d="M4.646 4.646a.5.5 0 0 1 .708 0L8 7.293l2.646-2.647a.5.5 0 0 1 .708.708L8.707 8l2.647 2.646a.5.5 0 0 1-.708.708L8 8.707l-2.646 2.647a.5.5 0 0 1-.708-.708L7.293 8 4.646 5.354a.5.5 0 0 1 0-.708z"/></svg>
                Reject
              </button>
            </div>
          )}

          {/* Legend */}
          <div className={styles.legend}>
            <div className={styles.legend__item}>
              <div className={`${styles.legend__dot} ${styles['legend__dot--formula']}`} />
              <span>Formula</span>
            </div>
            <div className={styles.legend__item}>
              <div className={`${styles.legend__dot} ${styles['legend__dot--hardcoded']}`} />
              <span>Value</span>
            </div>
            {anomalyCells.size > 0 && (
              <div className={styles.legend__item}>
                <div className={`${styles.legend__dot} ${styles['legend__dot--anomaly']}`} />
                <span>Anomaly ({anomalyCells.size})</span>
              </div>
            )}
            {pendingChanges.length > 0 && (
              <div className={styles.legend__item}>
                <div className={`${styles.legend__dot} ${styles['legend__dot--changed']}`} />
                <span>Pending ({pendingChanges.length})</span>
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
}
