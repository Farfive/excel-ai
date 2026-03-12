import React, { useState, useRef, useEffect, useCallback } from 'react';
import { WorkbookInfo } from '../../types';
import { getAnomalies, getWorkbookGrid, revertChanges, downloadWorkbook, getAuditTrail, AuditChange, diffWorkbooks, editCell } from '../../services/api';
import { onGridRefresh, onGridDiff, CellChange } from '../../services/gridBus';
import { AnalysisPanel } from '../Analysis/AnalysisPanel';
import { WorkbookSummary } from '../WorkbookSummary/WorkbookSummary';
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

type ViewTab = 'summary' | 'grid' | 'analysis' | 'audit' | 'compare';

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
  const gridWrapRef = useRef<HTMLDivElement>(null);
  const cellEditRef = useRef<HTMLInputElement>(null);
  const formulaInputRef = useRef<HTMLInputElement>(null);

  // ── Cell selection & editing state ──
  const [selectedCell, setSelectedCell] = useState<{row: number; col: number} | null>(null);
  const [editingCell, setEditingCell] = useState<{row: number; col: number} | null>(null);
  const [editValue, setEditValue] = useState('');
  const [saving, setSaving] = useState(false);

  // ── Context menu state ──
  const [ctxMenu, setCtxMenu] = useState<{x: number; y: number; row: number; col: number} | null>(null);

  // ── Clipboard state ──
  const [clipboard, setClipboard] = useState<{value: string; isFormula: boolean; isCut: boolean} | null>(null);

  // ── Undo/Redo state ──
  const undoStack = useRef<Array<{cell: string; sheet: string; oldVal: string; newVal: string}>>([]); 
  const redoStack = useRef<Array<{cell: string; sheet: string; oldVal: string; newVal: string}>>([]); 

  // ── Find & Replace state ──
  const [showFind, setShowFind] = useState(false);
  const [findText, setFindText] = useState('');
  const [replaceText, setReplaceText] = useState('');
  const [findMatches, setFindMatches] = useState<Array<{row: number; col: number}>>([]); 
  const [findIdx, setFindIdx] = useState(-1);
  const findInputRef = useRef<HTMLInputElement>(null);

  // ── Column widths state (for resize) ──
  const [colWidths, setColWidths] = useState<Record<number, number>>({});
  const resizingCol = useRef<{col: number; startX: number; startW: number} | null>(null);

  // ── Sort state ──
  const [sortCol, setSortCol] = useState<number | null>(null);
  const [sortDir, setSortDir] = useState<'asc' | 'desc'>('asc');

  const currentSheet = sheets[activeSheet] || null;

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
    setViewTab('summary');

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

  // ── Cell editing helpers ──
  const getCellAddr = useCallback((row: number, col: number): string => {
    if (!currentSheet) return '';
    return `${currentSheet.colHeaders[col] || ''}${row + 1}`;
  }, [sheets, activeSheet]);

  const getRawCell = useCallback((row: number, col: number): RawCell | null => {
    const s = sheets[activeSheet];
    if (!s || !s.rows[row]) return null;
    return s.rows[row][col] || null;
  }, [sheets, activeSheet]);

  const startEditing = useCallback((row: number, col: number, initialValue?: string) => {
    const cell = getRawCell(row, col);
    const val = initialValue !== undefined ? initialValue : (cell?.f ? `=${cell.v}` : (cell?.v ?? ''));
    setEditingCell({ row, col });
    setEditValue(val);
    setTimeout(() => cellEditRef.current?.focus(), 0);
  }, [getRawCell]);

  const cancelEditing = useCallback(() => {
    setEditingCell(null);
    setEditValue('');
    setTimeout(() => gridWrapRef.current?.focus(), 0);
  }, []);

  const commitEdit = useCallback(async (moveRow = 0, moveCol = 0) => {
    if (!editingCell || !workbookInfo || !currentSheet || saving) return;
    const addr = getCellAddr(editingCell.row, editingCell.col);
    const oldCell = getRawCell(editingCell.row, editingCell.col);
    const oldVal = oldCell?.v ?? '';

    // If value unchanged, just cancel
    const newVal = editValue;
    if (newVal === oldVal || (newVal === `=${oldVal}` && oldCell?.f)) {
      setEditingCell(null);
      setEditValue('');
    } else {
      setSaving(true);
      try {
        pushUndo(addr, currentSheet.name, oldVal, newVal);
        await editCell(workbookInfo.workbook_uuid, currentSheet.name, addr, newVal);
        fetchGrid(workbookInfo.workbook_uuid);
      } catch (e) {
        console.error('Cell edit failed:', e);
      } finally {
        setSaving(false);
      }
      setEditingCell(null);
      setEditValue('');
    }

    // Move selection
    if (moveRow !== 0 || moveCol !== 0) {
      setSelectedCell(prev => {
        if (!prev) return prev;
        const maxRow = currentSheet.rows.length - 1;
        const maxCol = currentSheet.colHeaders.length - 1;
        const nr = Math.max(0, Math.min(maxRow, prev.row + moveRow));
        const nc = Math.max(0, Math.min(maxCol, prev.col + moveCol));
        return { row: nr, col: nc };
      });
    }
    setTimeout(() => gridWrapRef.current?.focus(), 0);
  }, [editingCell, editValue, workbookInfo, sheets, activeSheet, saving, getCellAddr, getRawCell]);

  // ── Undo / Redo helpers ──
  const pushUndo = useCallback((cell: string, sheet: string, oldVal: string, newVal: string) => {
    undoStack.current.push({ cell, sheet, oldVal, newVal });
    redoStack.current = [];
  }, []);

  const handleUndo = useCallback(async () => {
    const entry = undoStack.current.pop();
    if (!entry || !workbookInfo) return;
    redoStack.current.push(entry);
    await editCell(workbookInfo.workbook_uuid, entry.sheet, entry.cell, entry.oldVal);
    fetchGrid(workbookInfo.workbook_uuid);
  }, [workbookInfo]);

  const handleRedo = useCallback(async () => {
    const entry = redoStack.current.pop();
    if (!entry || !workbookInfo) return;
    undoStack.current.push(entry);
    await editCell(workbookInfo.workbook_uuid, entry.sheet, entry.cell, entry.newVal);
    fetchGrid(workbookInfo.workbook_uuid);
  }, [workbookInfo]);

  // ── Clipboard helpers ──
  const handleCopy = useCallback((cut = false) => {
    if (!selectedCell || !currentSheet) return;
    const raw = getRawCell(selectedCell.row, selectedCell.col);
    if (!raw) return;
    const val = raw.f ? `=${raw.v}` : raw.v;
    setClipboard({ value: val, isFormula: raw.f, isCut: cut });
    try { navigator.clipboard.writeText(raw.v); } catch {}
  }, [selectedCell, currentSheet, getRawCell]);

  const handlePaste = useCallback(async () => {
    if (!selectedCell || !workbookInfo || !currentSheet) return;
    let val = clipboard?.value ?? '';
    if (!val) {
      try { val = await navigator.clipboard.readText(); } catch { return; }
    }
    const addr = getCellAddr(selectedCell.row, selectedCell.col);
    const oldCell = getRawCell(selectedCell.row, selectedCell.col);
    pushUndo(addr, currentSheet.name, oldCell?.v ?? '', val);
    await editCell(workbookInfo.workbook_uuid, currentSheet.name, addr, val);
    if (clipboard?.isCut) setClipboard(null);
    fetchGrid(workbookInfo.workbook_uuid);
  }, [selectedCell, clipboard, workbookInfo, currentSheet, getCellAddr, getRawCell, pushUndo]);

  // ── Find & Replace logic ──
  const runFind = useCallback((text: string) => {
    if (!currentSheet || !text) { setFindMatches([]); setFindIdx(-1); return; }
    const matches: Array<{row: number; col: number}> = [];
    const lower = text.toLowerCase();
    currentSheet.rows.forEach((row, r) => {
      row.forEach((cell, c) => {
        if (cell.v && cell.v.toLowerCase().includes(lower)) matches.push({ row: r, col: c });
      });
    });
    setFindMatches(matches);
    setFindIdx(matches.length > 0 ? 0 : -1);
    if (matches.length > 0) setSelectedCell(matches[0]);
  }, [currentSheet]);

  const findNext = useCallback(() => {
    if (findMatches.length === 0) return;
    const next = (findIdx + 1) % findMatches.length;
    setFindIdx(next);
    setSelectedCell(findMatches[next]);
  }, [findMatches, findIdx]);

  const findPrev = useCallback(() => {
    if (findMatches.length === 0) return;
    const prev = (findIdx - 1 + findMatches.length) % findMatches.length;
    setFindIdx(prev);
    setSelectedCell(findMatches[prev]);
  }, [findMatches, findIdx]);

  const handleReplaceOne = useCallback(async () => {
    if (findIdx < 0 || !workbookInfo || !currentSheet) return;
    const match = findMatches[findIdx];
    const addr = getCellAddr(match.row, match.col);
    const raw = getRawCell(match.row, match.col);
    const oldV = raw?.v ?? '';
    const newV = oldV.replace(new RegExp(findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i'), replaceText);
    pushUndo(addr, currentSheet.name, oldV, newV);
    await editCell(workbookInfo.workbook_uuid, currentSheet.name, addr, newV);
    fetchGrid(workbookInfo.workbook_uuid);
    runFind(findText);
  }, [findIdx, findMatches, findText, replaceText, workbookInfo, currentSheet, getCellAddr, getRawCell, pushUndo, runFind]);

  const handleReplaceAll = useCallback(async () => {
    if (findMatches.length === 0 || !workbookInfo || !currentSheet) return;
    for (const match of findMatches) {
      const addr = getCellAddr(match.row, match.col);
      const raw = getRawCell(match.row, match.col);
      const oldV = raw?.v ?? '';
      const newV = oldV.replace(new RegExp(findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'), replaceText);
      await editCell(workbookInfo.workbook_uuid, currentSheet.name, addr, newV);
    }
    fetchGrid(workbookInfo.workbook_uuid);
    setFindMatches([]);
    setFindIdx(-1);
  }, [findMatches, findText, replaceText, workbookInfo, currentSheet, getCellAddr, getRawCell]);

  // ── Column sort ──
  const handleSort = useCallback((colIdx: number, dir: 'asc' | 'desc') => {
    setSortCol(colIdx);
    setSortDir(dir);
    setCtxMenu(null);
  }, []);

  const sortedRows = (() => {
    if (!currentSheet) return [];
    if (sortCol === null) return currentSheet.rows;
    const rows = [...currentSheet.rows];
    const headerRow = rows.shift();
    rows.sort((a, b) => {
      const va = a[sortCol]?.v ?? '';
      const vb = b[sortCol]?.v ?? '';
      const na = parseFloat(va.replace(/[,$%]/g, ''));
      const nb = parseFloat(vb.replace(/[,$%]/g, ''));
      if (!isNaN(na) && !isNaN(nb)) return sortDir === 'asc' ? na - nb : nb - na;
      return sortDir === 'asc' ? va.localeCompare(vb) : vb.localeCompare(va);
    });
    if (headerRow) rows.unshift(headerRow);
    return rows;
  })();

  // ── Context menu handler ──
  const handleContextMenu = useCallback((e: React.MouseEvent, row: number, col: number) => {
    e.preventDefault();
    setSelectedCell({ row, col });
    setCtxMenu({ x: e.clientX, y: e.clientY, row, col });
  }, []);

  // Close context menu on click anywhere
  useEffect(() => {
    if (!ctxMenu) return;
    const close = () => setCtxMenu(null);
    window.addEventListener('click', close);
    return () => window.removeEventListener('click', close);
  }, [ctxMenu]);

  // ── Column resize handlers ──
  const handleResizeStart = useCallback((e: React.MouseEvent, colIdx: number) => {
    e.preventDefault();
    e.stopPropagation();
    const startW = colWidths[colIdx] || 90;
    resizingCol.current = { col: colIdx, startX: e.clientX, startW };

    const onMove = (ev: MouseEvent) => {
      if (!resizingCol.current) return;
      const delta = ev.clientX - resizingCol.current.startX;
      const newW = Math.max(40, resizingCol.current.startW + delta);
      setColWidths(prev => ({ ...prev, [resizingCol.current!.col]: newW }));
    };
    const onUp = () => {
      resizingCol.current = null;
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  }, [colWidths]);

  // ── Keyboard handler for grid ──
  const handleGridKeyDown = useCallback((e: React.KeyboardEvent) => {
    if (!currentSheet) return;
    const maxRow = currentSheet.rows.length - 1;
    const maxCol = currentSheet.colHeaders.length - 1;

    // If editing a cell
    if (editingCell) {
      if (e.key === 'Escape') { e.preventDefault(); cancelEditing(); return; }
      if (e.key === 'Enter') { e.preventDefault(); commitEdit(1, 0); return; }
      if (e.key === 'Tab') { e.preventDefault(); commitEdit(0, e.shiftKey ? -1 : 1); return; }
      return; // let other keys go to input
    }

    // If a cell is selected but not editing
    if (selectedCell) {
      const { row, col } = selectedCell;
      if (e.key === 'ArrowDown')  { e.preventDefault(); setSelectedCell({ row: Math.min(maxRow, row + 1), col }); return; }
      if (e.key === 'ArrowUp')    { e.preventDefault(); setSelectedCell({ row: Math.max(0, row - 1), col }); return; }
      if (e.key === 'ArrowRight') { e.preventDefault(); setSelectedCell({ row, col: Math.min(maxCol, col + 1) }); return; }
      if (e.key === 'ArrowLeft')  { e.preventDefault(); setSelectedCell({ row, col: Math.max(0, col - 1) }); return; }
      if (e.key === 'Tab')        { e.preventDefault(); setSelectedCell({ row, col: Math.min(maxCol, col + (e.shiftKey ? -1 : 1)) }); return; }
      if (e.key === 'Enter')      { e.preventDefault(); startEditing(row, col); return; }
      if (e.key === 'F2')         { e.preventDefault(); startEditing(row, col); return; }
      if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); startEditing(row, col, ''); return; }
      if (e.key === 'Escape')     { e.preventDefault(); setSelectedCell(null); return; }
      // Cmd/Ctrl shortcuts
      const mod = e.metaKey || e.ctrlKey;
      if (mod && e.key === 'c') { e.preventDefault(); handleCopy(false); return; }
      if (mod && e.key === 'x') { e.preventDefault(); handleCopy(true); return; }
      if (mod && e.key === 'v') { e.preventDefault(); handlePaste(); return; }
      if (mod && e.key === 'z' && !e.shiftKey) { e.preventDefault(); handleUndo(); return; }
      if (mod && (e.key === 'y' || (e.key === 'z' && e.shiftKey))) { e.preventDefault(); handleRedo(); return; }
      if (mod && e.key === 'f') { e.preventDefault(); setShowFind(true); setTimeout(() => findInputRef.current?.focus(), 0); return; }

      // Printable character → start editing with that char
      if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
        e.preventDefault();
        startEditing(row, col, e.key);
        return;
      }
    }
  }, [selectedCell, editingCell, editValue, currentSheet, startEditing, cancelEditing, commitEdit, handleCopy, handlePaste, handleUndo, handleRedo]);

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
          className={`${styles.analysisTab} ${viewTab === 'summary' ? styles['analysisTab--active'] : ''}`}
          onClick={() => setViewTab('summary')}
          type="button"
        >Summary</button>
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

      {viewTab === 'summary' ? (
        <div style={{ flex: 1, overflow: 'auto' }}>
          <WorkbookSummary
            workbookUuid={workbookInfo!.workbook_uuid}
            onSheetClick={(name) => {
              const idx = sheets.findIndex(s => s.name === name);
              if (idx >= 0) { setActiveSheet(idx); setViewTab('grid'); }
            }}
          />
        </div>
      ) : viewTab === 'compare' ? (
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
                <div style={{ fontFamily: 'var(--font-mono)', fontSize: 11, color: '#DC2626', marginBottom: 4 }}>- Sheets removed: {diffResult.sheets_removed.join(', ')}</div>
              )}
              {diffResult.sheet_diffs.length > 0 && (
                <table className={styles.auditTable} style={{ marginBottom: 12 }}>
                  <thead><tr><th>Sheet</th><th>Added</th><th>Removed</th><th>Modified</th><th>Formula Δ</th><th>Value Δ</th><th>Max Δ%</th></tr></thead>
                  <tbody>
                    {diffResult.sheet_diffs.map(sd => (
                      <tr key={sd.sheet}>
                        <td>{sd.sheet}</td>
                        <td style={{ color: 'var(--green)' }}>{sd.added || '—'}</td>
                        <td style={{ color: '#DC2626' }}>{sd.removed || '—'}</td>
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
                  onClick={() => { setActiveSheet(i); setSelectedCell(null); setEditingCell(null); }}
                  type="button"
                >{s.name}</button>
              ))}
            </div>
          )}

          {/* Formula bar */}
          {currentSheet && (
            <div className={styles.formulaBar}>
              <span className={styles.formulaBar__addr}>
                {selectedCell ? getCellAddr(selectedCell.row, selectedCell.col) : ''}
              </span>
              <span className={styles.formulaBar__fx}>fx</span>
              <input
                ref={formulaInputRef}
                className={styles.formulaBar__input}
                value={editingCell ? editValue : (
                  selectedCell ? (() => {
                    const rc = getRawCell(selectedCell.row, selectedCell.col);
                    return rc ? (rc.f ? `=${rc.v}` : rc.v) : '';
                  })() : ''
                )}
                onChange={e => {
                  if (editingCell) {
                    setEditValue(e.target.value);
                  } else if (selectedCell) {
                    startEditing(selectedCell.row, selectedCell.col, e.target.value);
                  }
                }}
                onFocus={() => {
                  if (selectedCell && !editingCell) {
                    startEditing(selectedCell.row, selectedCell.col);
                  }
                }}
                onKeyDown={e => {
                  if (e.key === 'Enter') { e.preventDefault(); commitEdit(1, 0); }
                  if (e.key === 'Escape') { e.preventDefault(); cancelEditing(); }
                  if (e.key === 'Tab') { e.preventDefault(); commitEdit(0, e.shiftKey ? -1 : 1); }
                }}
                readOnly={!editingCell && !selectedCell}
              />
            </div>
          )}

          {/* Find & Replace bar */}
          {showFind && (
            <div className={styles.findBar}>
              <input
                ref={findInputRef}
                className={styles.findBar__input}
                placeholder="Find…"
                value={findText}
                onChange={e => { setFindText(e.target.value); runFind(e.target.value); }}
                onKeyDown={e => {
                  if (e.key === 'Enter') { e.preventDefault(); e.shiftKey ? findPrev() : findNext(); }
                  if (e.key === 'Escape') { e.preventDefault(); setShowFind(false); setFindMatches([]); setFindIdx(-1); }
                }}
              />
              <input
                className={styles.findBar__input}
                placeholder="Replace…"
                value={replaceText}
                onChange={e => setReplaceText(e.target.value)}
                onKeyDown={e => {
                  if (e.key === 'Escape') { e.preventDefault(); setShowFind(false); setFindMatches([]); setFindIdx(-1); }
                }}
              />
              <span className={styles.findBar__count}>
                {findMatches.length > 0 ? `${findIdx + 1}/${findMatches.length}` : 'No matches'}
              </span>
              <button className={styles.findBar__btn} onClick={findPrev} type="button" title="Previous">▲</button>
              <button className={styles.findBar__btn} onClick={findNext} type="button" title="Next">▼</button>
              <button className={styles.findBar__btn} onClick={handleReplaceOne} type="button" title="Replace">Replace</button>
              <button className={styles.findBar__btn} onClick={handleReplaceAll} type="button" title="Replace All">All</button>
              <button className={styles.findBar__btn} onClick={() => { setShowFind(false); setFindMatches([]); setFindIdx(-1); }} type="button" title="Close">✕</button>
            </div>
          )}

          {/* Grid */}
          <div
            className={styles.gridWrap}
            ref={gridWrapRef}
            tabIndex={0}
            onKeyDown={handleGridKeyDown}
          >
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
                    {currentSheet.colHeaders.map((h, ci) => (
                      <th
                        key={h}
                        className={`${selectedCell?.col === ci ? styles['grid__colHeader--selected'] : ''} ${sortCol === ci ? styles['grid__colHeader--sorted'] : ''}`}
                        style={colWidths[ci] ? { width: colWidths[ci], minWidth: colWidths[ci], maxWidth: colWidths[ci] } : undefined}
                        onClick={() => handleSort(ci, sortCol === ci && sortDir === 'asc' ? 'desc' : 'asc')}
                      >
                        {h}
                        {sortCol === ci && <span className={styles.grid__sortIcon}>{sortDir === 'asc' ? ' ▲' : ' ▼'}</span>}
                        <span
                          className={styles.grid__resizeHandle}
                          onMouseDown={e => handleResizeStart(e, ci)}
                        />
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {sortedRows.map((row, r) => (
                    <tr key={r}>
                      <td className={`${styles.grid__rowNum} ${selectedCell?.row === r ? styles['grid__rowNum--selected'] : ''}`}>{r + 1}</td>
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
                        const isSel = selectedCell?.row === r && selectedCell?.col === c;
                        const isEditing = editingCell?.row === r && editingCell?.col === c;
                        const isFindMatch = findMatches.some(m => m.row === r && m.col === c);
                        const cls = [
                          styles.grid__cell,
                          isFormula ? styles['grid__cell--formula'] : styles['grid__cell--hardcoded'],
                          isHeader ? styles['grid__cell--header'] : '',
                          isNum ? styles['grid__cell--number'] : '',
                          isEmpty ? styles['grid__cell--empty'] : '',
                          isAnomaly ? styles['grid__cell--anomaly'] : '',
                          isPending ? styles['grid__cell--pending'] : '',
                          isSel && !isEditing ? styles['grid__cell--selected'] : '',
                          isEditing ? styles['grid__cell--editing'] : '',
                          isFindMatch ? styles['grid__cell--findMatch'] : '',
                          selectedCell && !isSel && selectedCell.col === c ? styles.grid__colHighlight : '',
                          selectedCell && !isSel && selectedCell.row === r ? styles.grid__rowHighlight : '',
                        ].filter(Boolean).join(' ');
                        return (
                          <td
                            key={c}
                            className={cls}
                            style={colWidths[c] ? { width: colWidths[c], minWidth: colWidths[c], maxWidth: colWidths[c] } : undefined}
                            title={isPending ? `${formatDiffVal(pending.before)} → ${formatDiffVal(pending.after)}` : (cell.nr ? `Named: ${cell.nr}` : cell.v)}
                            onClick={() => {
                              if (isEditing) return;
                              setSelectedCell({ row: r, col: c });
                              if (editingCell) commitEdit();
                            }}
                            onDoubleClick={() => startEditing(r, c)}
                            onContextMenu={e => handleContextMenu(e, r, c)}
                          >
                            {isEditing ? (
                              <input
                                ref={cellEditRef}
                                value={editValue}
                                onChange={e => setEditValue(e.target.value)}
                                onKeyDown={e => {
                                  if (e.key === 'Enter') { e.preventDefault(); commitEdit(1, 0); }
                                  if (e.key === 'Tab') { e.preventDefault(); commitEdit(0, e.shiftKey ? -1 : 1); }
                                  if (e.key === 'Escape') { e.preventDefault(); cancelEditing(); }
                                }}
                                onBlur={() => {
                                  setTimeout(() => {
                                    if (editingCell?.row === r && editingCell?.col === c) commitEdit();
                                  }, 100);
                                }}
                                autoFocus
                              />
                            ) : isPending ? (
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

          {/* Context menu */}
          {ctxMenu && (
            <div className={styles.ctxMenu} style={{ left: ctxMenu.x, top: ctxMenu.y }}>
              <button className={styles.ctxMenu__item} onClick={() => { handleCopy(false); setCtxMenu(null); }} type="button">
                Copy
                <span className={styles.ctxMenu__shortcut}>⌘C</span>
              </button>
              <button className={styles.ctxMenu__item} onClick={() => { handleCopy(true); setCtxMenu(null); }} type="button">
                Cut
                <span className={styles.ctxMenu__shortcut}>⌘X</span>
              </button>
              <button className={styles.ctxMenu__item} onClick={() => { handlePaste(); setCtxMenu(null); }} type="button">
                Paste
                <span className={styles.ctxMenu__shortcut}>⌘V</span>
              </button>
              <div className={styles.ctxMenu__sep} />
              <button className={styles.ctxMenu__item} onClick={() => { startEditing(ctxMenu.row, ctxMenu.col, ''); setCtxMenu(null); }} type="button">
                Clear cell
              </button>
              <div className={styles.ctxMenu__sep} />
              <button className={styles.ctxMenu__item} onClick={() => handleSort(ctxMenu.col, 'asc')} type="button">
                Sort A → Z
              </button>
              <button className={styles.ctxMenu__item} onClick={() => handleSort(ctxMenu.col, 'desc')} type="button">
                Sort Z → A
              </button>
            </div>
          )}

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
