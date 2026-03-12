import React, { useState, useRef, useEffect, useCallback, useMemo } from 'react';
import { WorkbookInfo } from '../../types';
import {
  getAnomalies, getWorkbookGrid, revertChanges, downloadWorkbook,
  getAuditTrail, AuditChange, diffWorkbooks, editCell,
  formatCells, insertRow, insertCol, deleteRow, deleteCol,
  GridCell, CellStyleData,
  setDataValidation, getDataValidation, DataValidationRule,
  setNamedRange, deleteNamedRange, getNamedRanges,
  goalSeek, textToColumns, removeDuplicates,
} from '../../services/api';
import { onGridRefresh, onGridDiff, CellChange } from '../../services/gridBus';
import { AnalysisPanel } from '../Analysis/AnalysisPanel';
import { WorkbookSummary } from '../WorkbookSummary/WorkbookSummary';
import { FormatToolbar } from '../FormatToolbar/FormatToolbar';
import { ChartPanel } from '../ChartPanel/ChartPanel';
import styles from './ExcelViewer.module.css';

type RawCell = GridCell;

interface SheetData {
  name: string;
  rows: RawCell[][];
  colHeaders: string[];
}

interface CellPos { row: number; col: number; }

interface SelectionRange {
  start: CellPos;
  end: CellPos;
}

function normalizeRange(r: SelectionRange): { r1: number; c1: number; r2: number; c2: number } {
  return {
    r1: Math.min(r.start.row, r.end.row),
    c1: Math.min(r.start.col, r.end.col),
    r2: Math.max(r.start.row, r.end.row),
    c2: Math.max(r.start.col, r.end.col),
  };
}

function isCellInSelection(row: number, col: number, sel: SelectionRange | null): boolean {
  if (!sel) return false;
  const { r1, c1, r2, c2 } = normalizeRange(sel);
  return row >= r1 && row <= r2 && col >= c1 && col <= c2;
}

function cellStyleToCSS(s?: CellStyleData): React.CSSProperties | undefined {
  if (!s) return undefined;
  const css: React.CSSProperties = {};
  if (s.b) css.fontWeight = 'bold';
  if (s.i) css.fontStyle = 'italic';
  if (s.u) css.textDecoration = 'underline';
  if (s.fs) css.fontSize = `${s.fs}px`;
  if (s.fn) css.fontFamily = s.fn;
  if (s.fc) css.color = s.fc;
  if (s.bg) css.backgroundColor = s.bg;
  if (s.ha) css.textAlign = s.ha as any;
  if (s.va) css.verticalAlign = s.va === 'center' ? 'middle' : s.va as any;
  if (s.wt) { css.whiteSpace = 'pre-wrap'; css.wordBreak = 'break-word'; }
  if (s.ind) css.paddingLeft = `${s.ind * 12}px`;
  if (s.bt) css.borderTop = s.bt === 'thick' ? '2px solid #000' : s.bt === 'double' ? '3px double #000' : '1px solid #000';
  if (s.bb) css.borderBottom = s.bb === 'thick' ? '2px solid #000' : s.bb === 'double' ? '3px double #000' : '1px solid #000';
  if (s.bl) css.borderLeft = s.bl === 'thick' ? '2px solid #000' : '1px solid #000';
  if (s.br) css.borderRight = s.br === 'thick' ? '2px solid #000' : '1px solid #000';
  return Object.keys(css).length > 0 ? css : undefined;
}

const EXCEL_FUNCTIONS = [
  'SUM','AVERAGE','COUNT','COUNTA','COUNTBLANK','MAX','MIN','IF','IFERROR','IFNA',
  'VLOOKUP','HLOOKUP','INDEX','MATCH','XLOOKUP','SUMIF','SUMIFS','COUNTIF','COUNTIFS',
  'AVERAGEIF','AVERAGEIFS','LEFT','RIGHT','MID','LEN','TRIM','UPPER','LOWER','PROPER',
  'CONCATENATE','TEXTJOIN','SUBSTITUTE','FIND','SEARCH','VALUE','TEXT','DATE','TODAY',
  'NOW','YEAR','MONTH','DAY','DATEDIF','ROUND','ROUNDUP','ROUNDDOWN','ABS','SQRT',
  'POWER','MOD','INT','RAND','RANDBETWEEN','AND','OR','NOT','TRUE','FALSE',
  'PMT','FV','PV','NPV','IRR','RATE','OFFSET','INDIRECT','ROW','COLUMN',
  'ROWS','COLUMNS','TRANSPOSE','UNIQUE','SORT','FILTER','SEQUENCE','LET','LAMBDA',
];

interface ExcelViewerProps {
  workbookInfo: WorkbookInfo | null;
  isUploading: boolean;
  uploadError: string | null;
  onUpload: (f: File) => void;
}

type ViewTab = 'summary' | 'grid' | 'charts' | 'analysis' | 'audit' | 'compare';

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

  // ── Multi-cell selection range ──
  const [selRange, setSelRange] = useState<SelectionRange | null>(null);
  const isDragging = useRef(false);

  // ── Filtering state ──
  const [filterCol, setFilterCol] = useState<number | null>(null);
  const [filterValues, setFilterValues] = useState<Set<string>>(new Set());
  const [filterOpen, setFilterOpen] = useState<number | null>(null);
  const filterRef = useRef<HTMLDivElement>(null);

  // ── Freeze panes ──
  const [freezeRow, setFreezeRow] = useState(0);
  const [freezeCol, setFreezeCol] = useState(0);

  // ── Formula autocomplete ──
  const [formulaSuggestions, setFormulaSuggestions] = useState<string[]>([]);
  const [formulaSugIdx, setFormulaSugIdx] = useState(-1);

  // ── Row heights ──
  const [rowHeights, setRowHeights] = useState<Record<number, number>>({});

  // ── Conditional formatting rules ──
  interface CondRule { id: string; type: 'gt' | 'lt' | 'eq' | 'between' | 'text' | 'duplicate' | 'colorScale'; col: number; v1: string; v2?: string; color: string; bgColor: string; }
  const [condRules, setCondRules] = useState<CondRule[]>([]);
  const [showCondFmt, setShowCondFmt] = useState(false);

  // ── Cell comments ──
  const [comments, setComments] = useState<Record<string, string>>({});
  const [editingComment, setEditingComment] = useState<string | null>(null);
  const [commentText, setCommentText] = useState('');

  // ── Paste special ──
  const [showPasteSpecial, setShowPasteSpecial] = useState(false);

  // ── Hidden rows/cols ──
  const [hiddenRows, setHiddenRows] = useState<Set<number>>(new Set());
  const [hiddenCols, setHiddenCols] = useState<Set<number>>(new Set());

  // ── Data tools dialogs ──
  const [showDataValidation, setShowDataValidation] = useState(false);
  const [showNamedRanges, setShowNamedRanges] = useState(false);
  const [showGoalSeek, setShowGoalSeek] = useState(false);
  const [showTextToCol, setShowTextToCol] = useState(false);
  const [showRemoveDups, setShowRemoveDups] = useState(false);
  const [namedRangesData, setNamedRangesData] = useState<Record<string, string>>({});
  const [goalSeekResult, setGoalSeekResult] = useState<string | null>(null);
  const [dataToolMsg, setDataToolMsg] = useState<string | null>(null);

  // ── Grouping / outlining ──
  const [collapsedGroups, setCollapsedGroups] = useState<Set<string>>(new Set());

  const currentSheet = sheets[activeSheet] || null;

  // ── Conditional formatting evaluation ──
  const evalCondFmt = useCallback((row: number, col: number, value: string): React.CSSProperties | null => {
    if (condRules.length === 0) return null;
    const num = parseFloat(value.replace(/[,$%]/g, ''));
    for (const rule of condRules) {
      if (rule.col !== -1 && rule.col !== col) continue;
      let match = false;
      if (rule.type === 'gt' && !isNaN(num) && num > parseFloat(rule.v1)) match = true;
      if (rule.type === 'lt' && !isNaN(num) && num < parseFloat(rule.v1)) match = true;
      if (rule.type === 'eq' && value === rule.v1) match = true;
      if (rule.type === 'between' && !isNaN(num) && rule.v2) {
        if (num >= parseFloat(rule.v1) && num <= parseFloat(rule.v2)) match = true;
      }
      if (rule.type === 'text' && value.toLowerCase().includes(rule.v1.toLowerCase())) match = true;
      if (rule.type === 'duplicate') {
        if (!currentSheet) continue;
        let count = 0;
        for (const r of currentSheet.rows) {
          if (r[col]?.v === value) count++;
        }
        if (count > 1) match = true;
      }
      if (match) {
        const css: React.CSSProperties = {};
        if (rule.bgColor) css.backgroundColor = rule.bgColor;
        if (rule.color) css.color = rule.color;
        return css;
      }
    }
    return null;
  }, [condRules, currentSheet]);

  // ── Selection helpers for multi-cell ──
  const getSelectedCells = useCallback((): string[] => {
    if (!currentSheet) return [];
    if (selRange) {
      const { r1, c1, r2, c2 } = normalizeRange(selRange);
      const cells: string[] = [];
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          cells.push(`${currentSheet.name}!${currentSheet.colHeaders[c]}${r + 1}`);
        }
      }
      return cells;
    }
    if (selectedCell) {
      return [`${currentSheet.name}!${currentSheet.colHeaders[selectedCell.col]}${selectedCell.row + 1}`];
    }
    return [];
  }, [currentSheet, selRange, selectedCell]);

  // ── Current style for selection (from first selected cell) ──
  const currentSelStyle = useMemo((): CellStyleData | null => {
    if (!currentSheet || !selectedCell) return null;
    const cell = currentSheet.rows[selectedCell.row]?.[selectedCell.col];
    return cell?.s || null;
  }, [currentSheet, selectedCell]);

  // ── Format handler ──
  const handleFormat = useCallback(async (stylePatch: Partial<CellStyleData>) => {
    if (!workbookInfo) return;
    const cells = getSelectedCells();
    if (cells.length === 0) return;
    try {
      await formatCells(workbookInfo.workbook_uuid, cells, stylePatch);
      fetchGrid(workbookInfo.workbook_uuid);
    } catch (e) {
      console.error('Format failed:', e);
    }
  }, [workbookInfo, getSelectedCells]);

  // ── Status bar calculations ──
  const statusBarInfo = useMemo(() => {
    if (!currentSheet) return null;
    const cells: number[] = [];
    let count = 0;
    if (selRange) {
      const { r1, c1, r2, c2 } = normalizeRange(selRange);
      for (let r = r1; r <= r2; r++) {
        for (let c = c1; c <= c2; c++) {
          const cell = currentSheet.rows[r]?.[c];
          if (cell?.v) {
            count++;
            const n = parseFloat(cell.v.replace(/[,$%]/g, ''));
            if (!isNaN(n)) cells.push(n);
          }
        }
      }
    } else if (selectedCell) {
      const cell = currentSheet.rows[selectedCell.row]?.[selectedCell.col];
      if (cell?.v) {
        count = 1;
        const n = parseFloat(cell.v.replace(/[,$%]/g, ''));
        if (!isNaN(n)) cells.push(n);
      }
    }
    if (cells.length === 0) return { count, sum: 0, avg: 0, min: 0, max: 0, numCount: 0 };
    const sum = cells.reduce((a, b) => a + b, 0);
    return {
      count,
      numCount: cells.length,
      sum,
      avg: sum / cells.length,
      min: Math.min(...cells),
      max: Math.max(...cells),
    };
  }, [currentSheet, selRange, selectedCell]);

  // ── Filtered rows ──
  const filteredRows = useMemo(() => {
    if (!currentSheet) return [];
    if (filterCol === null || filterValues.size === 0) return currentSheet.rows;
    return currentSheet.rows.filter((row, idx) => {
      if (idx === 0) return true;
      const val = row[filterCol]?.v || '';
      return filterValues.has(val);
    });
  }, [currentSheet, filterCol, filterValues]);

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
      return;
    }

    // If a cell is selected but not editing
    if (selectedCell) {
      const { row, col } = selectedCell;
      const mod = e.metaKey || e.ctrlKey;

      // Shift+Arrow → extend selection range
      if (e.shiftKey && ['ArrowDown','ArrowUp','ArrowRight','ArrowLeft'].includes(e.key)) {
        e.preventDefault();
        const anchor = selRange ? selRange.start : { row, col };
        const end = selRange ? { ...selRange.end } : { row, col };
        if (e.key === 'ArrowDown')  end.row = Math.min(maxRow, end.row + 1);
        if (e.key === 'ArrowUp')    end.row = Math.max(0, end.row - 1);
        if (e.key === 'ArrowRight') end.col = Math.min(maxCol, end.col + 1);
        if (e.key === 'ArrowLeft')  end.col = Math.max(0, end.col - 1);
        setSelRange({ start: anchor, end });
        return;
      }

      // Arrow keys → single cell move (clears range)
      if (e.key === 'ArrowDown')  { e.preventDefault(); setSelectedCell({ row: Math.min(maxRow, row + 1), col }); setSelRange(null); return; }
      if (e.key === 'ArrowUp')    { e.preventDefault(); setSelectedCell({ row: Math.max(0, row - 1), col }); setSelRange(null); return; }
      if (e.key === 'ArrowRight') { e.preventDefault(); setSelectedCell({ row, col: Math.min(maxCol, col + 1) }); setSelRange(null); return; }
      if (e.key === 'ArrowLeft')  { e.preventDefault(); setSelectedCell({ row, col: Math.max(0, col - 1) }); setSelRange(null); return; }
      if (e.key === 'Tab')        { e.preventDefault(); setSelectedCell({ row, col: Math.min(maxCol, col + (e.shiftKey ? -1 : 1)) }); setSelRange(null); return; }
      if (e.key === 'Enter')      { e.preventDefault(); startEditing(row, col); return; }
      if (e.key === 'F2')         { e.preventDefault(); startEditing(row, col); return; }
      if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); startEditing(row, col, ''); return; }
      if (e.key === 'Escape')     { e.preventDefault(); setSelectedCell(null); setSelRange(null); return; }

      // Ctrl+A → select all
      if (mod && e.key === 'a') { e.preventDefault(); setSelRange({ start: { row: 0, col: 0 }, end: { row: maxRow, col: maxCol } }); return; }

      // Formatting shortcuts
      if (mod && e.key === 'b') { e.preventDefault(); handleFormat({ b: !(currentSelStyle?.b) }); return; }
      if (mod && e.key === 'i') { e.preventDefault(); handleFormat({ i: !(currentSelStyle?.i) }); return; }
      if (mod && e.key === 'u') { e.preventDefault(); handleFormat({ u: !(currentSelStyle?.u) }); return; }

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
  }, [selectedCell, editingCell, editValue, currentSheet, selRange, currentSelStyle, startEditing, cancelEditing, commitEdit, handleCopy, handlePaste, handleUndo, handleRedo, handleFormat]);

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
        <button className={styles.toolbar__btn} onClick={() => setShowCondFmt(!showCondFmt)} type="button" title="Conditional Formatting" style={{ marginLeft: 'auto' }}>Cond.Fmt</button>
        <button className={styles.toolbar__btn} onClick={() => setShowDataValidation(true)} type="button" title="Data Validation">Validation</button>
        <button className={styles.toolbar__btn} onClick={async () => {
          if (!workbookInfo) return;
          const res = await getNamedRanges(workbookInfo.workbook_uuid);
          setNamedRangesData(res.named_ranges);
          setShowNamedRanges(true);
        }} type="button" title="Named Ranges">Names</button>
        <button className={styles.toolbar__btn} onClick={() => setShowGoalSeek(true)} type="button" title="Goal Seek">Goal Seek</button>
        <button className={styles.toolbar__btn} onClick={() => setShowTextToCol(true)} type="button" title="Text to Columns">Txt→Col</button>
        <button className={styles.toolbar__btn} onClick={() => setShowRemoveDups(true)} type="button" title="Remove Duplicates">Rm Dups</button>
        <button className={styles.toolbar__btn} onClick={() => window.print()} type="button" title="Print / PDF">Print</button>
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
          className={`${styles.analysisTab} ${viewTab === 'charts' ? styles['analysisTab--active'] : ''}`}
          onClick={() => setViewTab('charts')}
          type="button"
        >Charts</button>
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
      ) : viewTab === 'charts' ? (
        <div style={{ flex: 1, overflow: 'auto' }}>
          {currentSheet ? (
            <ChartPanel
              rows={currentSheet.rows}
              colHeaders={currentSheet.colHeaders}
              sheetName={currentSheet.name}
            />
          ) : (
            <div style={{ padding: 32, textAlign: 'center', color: 'var(--text-muted)', fontFamily: 'var(--font-mono)', fontSize: 12 }}>
              No sheet data available
            </div>
          )}
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
                  onClick={() => { setActiveSheet(i); setSelectedCell(null); setEditingCell(null); setSelRange(null); }}
                  type="button"
                >{s.name}</button>
              ))}
            </div>
          )}

          {/* Format Toolbar (Ribbon) */}
          <FormatToolbar
            currentStyle={currentSelStyle}
            hasSelection={!!selectedCell || !!selRange}
            onFormat={handleFormat}
          />

          {/* Formula bar */}
          {currentSheet && (
            <div className={styles.formulaBar}>
              <span className={styles.formulaBar__addr}>
                {selRange
                  ? `${currentSheet.colHeaders[normalizeRange(selRange).c1]}${normalizeRange(selRange).r1 + 1}:${currentSheet.colHeaders[normalizeRange(selRange).c2]}${normalizeRange(selRange).r2 + 1}`
                  : selectedCell ? getCellAddr(selectedCell.row, selectedCell.col) : ''}
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
                  const val = e.target.value;
                  if (editingCell) {
                    setEditValue(val);
                    if (val.startsWith('=')) {
                      const match = val.match(/=([A-Z]+)\(?$/i);
                      if (match) {
                        const prefix = match[1].toUpperCase();
                        setFormulaSuggestions(EXCEL_FUNCTIONS.filter(f => f.startsWith(prefix)).slice(0, 8));
                        setFormulaSugIdx(0);
                      } else {
                        setFormulaSuggestions([]);
                      }
                    } else {
                      setFormulaSuggestions([]);
                    }
                  } else if (selectedCell) {
                    startEditing(selectedCell.row, selectedCell.col, val);
                  }
                }}
                onFocus={() => {
                  if (selectedCell && !editingCell) {
                    startEditing(selectedCell.row, selectedCell.col);
                  }
                }}
                onKeyDown={e => {
                  if (formulaSuggestions.length > 0) {
                    if (e.key === 'ArrowDown') { e.preventDefault(); setFormulaSugIdx(i => Math.min(formulaSuggestions.length - 1, i + 1)); return; }
                    if (e.key === 'ArrowUp') { e.preventDefault(); setFormulaSugIdx(i => Math.max(0, i - 1)); return; }
                    if (e.key === 'Tab' || e.key === 'Enter') {
                      if (formulaSugIdx >= 0 && formulaSugIdx < formulaSuggestions.length) {
                        e.preventDefault();
                        const fn = formulaSuggestions[formulaSugIdx];
                        const eqIdx = editValue.lastIndexOf('=');
                        const newVal = editValue.slice(0, eqIdx + 1) + fn + '(';
                        setEditValue(newVal);
                        setFormulaSuggestions([]);
                        return;
                      }
                    }
                    if (e.key === 'Escape') { setFormulaSuggestions([]); return; }
                  }
                  if (e.key === 'Enter') { e.preventDefault(); commitEdit(1, 0); }
                  if (e.key === 'Escape') { e.preventDefault(); cancelEditing(); }
                  if (e.key === 'Tab') { e.preventDefault(); commitEdit(0, e.shiftKey ? -1 : 1); }
                }}
                readOnly={!editingCell && !selectedCell}
              />
              {formulaSuggestions.length > 0 && (
                <div className={styles.formulaSuggestions}>
                  {formulaSuggestions.map((fn, i) => (
                    <div
                      key={fn}
                      className={`${styles.formulaSuggestion} ${i === formulaSugIdx ? styles['formulaSuggestion--active'] : ''}`}
                      onMouseDown={e => {
                        e.preventDefault();
                        const eqIdx = editValue.lastIndexOf('=');
                        setEditValue(editValue.slice(0, eqIdx + 1) + fn + '(');
                        setFormulaSuggestions([]);
                      }}
                    >{fn}</div>
                  ))}
                </div>
              )}
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
            onMouseUp={() => { isDragging.current = false; }}
            onMouseLeave={() => { isDragging.current = false; }}
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
                    {currentSheet.colHeaders.map((h, ci) => {
                      if (hiddenCols.has(ci)) return null;
                      return (
                      <th
                        key={h}
                        className={`${selectedCell?.col === ci || isCellInSelection(0, ci, selRange) ? styles['grid__colHeader--selected'] : ''} ${sortCol === ci ? styles['grid__colHeader--sorted'] : ''}`}
                        style={colWidths[ci] ? { width: colWidths[ci], minWidth: colWidths[ci], maxWidth: colWidths[ci] } : undefined}
                        onClick={() => handleSort(ci, sortCol === ci && sortDir === 'asc' ? 'desc' : 'asc')}
                      >
                        {h}
                        {sortCol === ci && <span className={styles.grid__sortIcon}>{sortDir === 'asc' ? ' ▲' : ' ▼'}</span>}
                        {filterCol === ci && <span className={styles.grid__filterIcon}>▾</span>}
                        <span
                          className={styles.grid__filterBtn}
                          onClick={e => { e.stopPropagation(); setFilterOpen(filterOpen === ci ? null : ci); }}
                          title="Filter"
                        >▼</span>
                        {filterOpen === ci && (
                          <div className={styles.filterDropdown} ref={filterRef} onClick={e => e.stopPropagation()}>
                            <div className={styles.filterDropdown__header}>
                              <button className={styles.findBar__btn} onClick={() => { setFilterCol(null); setFilterValues(new Set()); setFilterOpen(null); }} type="button">Clear</button>
                              <button className={styles.findBar__btn} onClick={() => setFilterOpen(null)} type="button">✕</button>
                            </div>
                            <div className={styles.filterDropdown__list}>
                              {(() => {
                                const vals = new Set<string>();
                                currentSheet.rows.forEach((row, ri) => { if (ri > 0 && row[ci]?.v) vals.add(row[ci].v); });
                                return [...vals].sort().map(val => (
                                  <label key={val} className={styles.filterDropdown__item}>
                                    <input
                                      type="checkbox"
                                      checked={filterCol === ci ? filterValues.has(val) : true}
                                      onChange={e => {
                                        const newVals = new Set(filterCol === ci ? filterValues : vals);
                                        if (e.target.checked) newVals.add(val); else newVals.delete(val);
                                        setFilterCol(ci);
                                        setFilterValues(newVals);
                                      }}
                                    />
                                    <span>{val || '(blank)'}</span>
                                  </label>
                                ));
                              })()}
                            </div>
                          </div>
                        )}
                        <span
                          className={styles.grid__resizeHandle}
                          onMouseDown={e => handleResizeStart(e, ci)}
                          onDoubleClick={e => {
                            e.stopPropagation();
                            let maxW = 60;
                            currentSheet.rows.forEach(row => {
                              const v = row[ci]?.v || '';
                              maxW = Math.max(maxW, v.length * 8 + 16);
                            });
                            setColWidths(prev => ({ ...prev, [ci]: Math.min(maxW, 400) }));
                          }}
                        />
                      </th>
                      );
                    })}
                  </tr>
                </thead>
                <tbody>
                  {sortedRows.map((row, r) => (
                    <tr key={r} style={rowHeights[r] ? { height: rowHeights[r] } : undefined}>
                      <td
                        className={`${styles.grid__rowNum} ${selectedCell?.row === r || isCellInSelection(r, 0, selRange) ? styles['grid__rowNum--selected'] : ''}`}
                        onMouseDown={() => {
                          const maxCol = currentSheet.colHeaders.length - 1;
                          setSelectedCell({ row: r, col: 0 });
                          setSelRange({ start: { row: r, col: 0 }, end: { row: r, col: maxCol } });
                        }}
                      >{r + 1}</td>
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
                        const isInRange = isCellInSelection(r, c, selRange);
                        const isEditing = editingCell?.row === r && editingCell?.col === c;
                        const isFindMatch = findMatches.some(m => m.row === r && m.col === c);
                        const isMergedSlave = !!cell.mm;
                        if (isMergedSlave) return null;
                        const cls = [
                          styles.grid__cell,
                          isFormula ? styles['grid__cell--formula'] : styles['grid__cell--hardcoded'],
                          isHeader ? styles['grid__cell--header'] : '',
                          isNum ? styles['grid__cell--number'] : '',
                          isEmpty ? styles['grid__cell--empty'] : '',
                          isAnomaly ? styles['grid__cell--anomaly'] : '',
                          isPending ? styles['grid__cell--pending'] : '',
                          isSel && !isEditing ? styles['grid__cell--selected'] : '',
                          isInRange && !isSel ? styles['grid__cell--inRange'] : '',
                          isEditing ? styles['grid__cell--editing'] : '',
                          isFindMatch ? styles['grid__cell--findMatch'] : '',
                          !isSel && !isInRange && selectedCell?.col === c ? styles.grid__colHighlight : '',
                          !isSel && !isInRange && selectedCell?.row === r ? styles.grid__rowHighlight : '',
                        ].filter(Boolean).join(' ');
                        const cellCss = cellStyleToCSS(cell.s);
                        const condCss = evalCondFmt(r, c, cell.v);
                        const widthStyle = colWidths[c] ? { width: colWidths[c], minWidth: colWidths[c], maxWidth: colWidths[c] } : undefined;
                        const mergedStyle = { ...widthStyle, ...cellCss, ...condCss };
                        const hasComment = !!comments[addr];
                        return (
                          <td
                            key={c}
                            className={cls}
                            style={mergedStyle}
                            title={isPending ? `${formatDiffVal(pending.before)} → ${formatDiffVal(pending.after)}` : (cell.nr ? `Named: ${cell.nr}` : cell.v)}
                            onMouseDown={e => {
                              if (isEditing) return;
                              if (e.shiftKey && selectedCell) {
                                setSelRange({ start: { row: selectedCell.row, col: selectedCell.col }, end: { row: r, col: c } });
                              } else {
                                setSelectedCell({ row: r, col: c });
                                setSelRange(null);
                                isDragging.current = true;
                              }
                              if (editingCell) commitEdit();
                            }}
                            onMouseEnter={() => {
                              if (isDragging.current && selectedCell) {
                                setSelRange({ start: { row: selectedCell.row, col: selectedCell.col }, end: { row: r, col: c } });
                              }
                            }}
                            onDoubleClick={() => startEditing(r, c)}
                            onContextMenu={e => handleContextMenu(e, r, c)}
                          >
                            {isEditing ? (
                              <input
                                ref={cellEditRef}
                                value={editValue}
                                onChange={e => {
                                  setEditValue(e.target.value);
                                  if (e.target.value.startsWith('=')) {
                                    const match = e.target.value.match(/=([A-Z]+)\(?$/i);
                                    if (match) {
                                      const prefix = match[1].toUpperCase();
                                      setFormulaSuggestions(EXCEL_FUNCTIONS.filter(f => f.startsWith(prefix)).slice(0, 8));
                                      setFormulaSugIdx(0);
                                    } else {
                                      setFormulaSuggestions([]);
                                    }
                                  }
                                }}
                                onKeyDown={e => {
                                  if (formulaSuggestions.length > 0) {
                                    if (e.key === 'ArrowDown') { e.preventDefault(); setFormulaSugIdx(i => Math.min(formulaSuggestions.length - 1, i + 1)); return; }
                                    if (e.key === 'ArrowUp') { e.preventDefault(); setFormulaSugIdx(i => Math.max(0, i - 1)); return; }
                                    if (e.key === 'Tab') {
                                      if (formulaSugIdx >= 0) {
                                        e.preventDefault();
                                        const fn = formulaSuggestions[formulaSugIdx];
                                        const eqIdx = editValue.lastIndexOf('=');
                                        setEditValue(editValue.slice(0, eqIdx + 1) + fn + '(');
                                        setFormulaSuggestions([]);
                                        return;
                                      }
                                    }
                                  }
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
                            ) : /^https?:\/\//i.test(cell.v) ? (
                              <a
                                href={cell.v}
                                target="_blank"
                                rel="noopener noreferrer"
                                className={styles['grid__cell--link']}
                                onClick={e => e.stopPropagation()}
                              >{cell.v}</a>
                            ) : cell.v}
                            {hasComment && <span className={styles.commentIndicator} title={comments[addr]} />}
                            {formulaSuggestions.length > 0 && isEditing && (
                              <div className={styles.formulaSuggestions} style={{ position: 'absolute', top: '100%', left: 0, zIndex: 100 }}>
                                {formulaSuggestions.map((fn, i) => (
                                  <div
                                    key={fn}
                                    className={`${styles.formulaSuggestion} ${i === formulaSugIdx ? styles['formulaSuggestion--active'] : ''}`}
                                    onMouseDown={e => {
                                      e.preventDefault();
                                      const eqIdx = editValue.lastIndexOf('=');
                                      setEditValue(editValue.slice(0, eqIdx + 1) + fn + '(');
                                      setFormulaSuggestions([]);
                                    }}
                                  >{fn}</div>
                                ))}
                              </div>
                            )}
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
              <button className={styles.ctxMenu__item} onClick={async () => {
                if (!selectedCell || !workbookInfo || !currentSheet) { setCtxMenu(null); return; }
                let val = '';
                try { val = await navigator.clipboard.readText(); } catch { setCtxMenu(null); return; }
                const numVal = parseFloat(val.replace(/[,$%]/g, ''));
                const addr = getCellAddr(selectedCell.row, selectedCell.col);
                const oldCell = getRawCell(selectedCell.row, selectedCell.col);
                pushUndo(addr, currentSheet.name, oldCell?.v ?? '', isNaN(numVal) ? val : String(numVal));
                await editCell(workbookInfo.workbook_uuid, currentSheet.name, addr, isNaN(numVal) ? val : String(numVal));
                fetchGrid(workbookInfo.workbook_uuid);
                setCtxMenu(null);
              }} type="button">Paste values only</button>
              <div className={styles.ctxMenu__sep} />
              <button className={styles.ctxMenu__item} onClick={() => { startEditing(ctxMenu.row, ctxMenu.col, ''); setCtxMenu(null); }} type="button">
                Clear cell
              </button>
              <div className={styles.ctxMenu__sep} />
              <button className={styles.ctxMenu__item} onClick={async () => {
                if (!workbookInfo || !currentSheet) return;
                await insertRow(workbookInfo.workbook_uuid, currentSheet.name, ctxMenu.row + 1);
                fetchGrid(workbookInfo.workbook_uuid);
                setCtxMenu(null);
              }} type="button">Insert row above</button>
              <button className={styles.ctxMenu__item} onClick={async () => {
                if (!workbookInfo || !currentSheet) return;
                await insertRow(workbookInfo.workbook_uuid, currentSheet.name, ctxMenu.row + 2);
                fetchGrid(workbookInfo.workbook_uuid);
                setCtxMenu(null);
              }} type="button">Insert row below</button>
              <button className={styles.ctxMenu__item} onClick={async () => {
                if (!workbookInfo || !currentSheet) return;
                await insertCol(workbookInfo.workbook_uuid, currentSheet.name, ctxMenu.col + 1);
                fetchGrid(workbookInfo.workbook_uuid);
                setCtxMenu(null);
              }} type="button">Insert column left</button>
              <button className={styles.ctxMenu__item} onClick={async () => {
                if (!workbookInfo || !currentSheet) return;
                await insertCol(workbookInfo.workbook_uuid, currentSheet.name, ctxMenu.col + 2);
                fetchGrid(workbookInfo.workbook_uuid);
                setCtxMenu(null);
              }} type="button">Insert column right</button>
              <div className={styles.ctxMenu__sep} />
              <button className={styles.ctxMenu__item} onClick={async () => {
                if (!workbookInfo || !currentSheet) return;
                await deleteRow(workbookInfo.workbook_uuid, currentSheet.name, ctxMenu.row + 1);
                fetchGrid(workbookInfo.workbook_uuid);
                setCtxMenu(null);
              }} type="button" style={{ color: '#DC2626' }}>Delete row</button>
              <button className={styles.ctxMenu__item} onClick={async () => {
                if (!workbookInfo || !currentSheet) return;
                await deleteCol(workbookInfo.workbook_uuid, currentSheet.name, ctxMenu.col + 1);
                fetchGrid(workbookInfo.workbook_uuid);
                setCtxMenu(null);
              }} type="button" style={{ color: '#DC2626' }}>Delete column</button>
              <div className={styles.ctxMenu__sep} />
              <button className={styles.ctxMenu__item} onClick={() => handleSort(ctxMenu.col, 'asc')} type="button">
                Sort A → Z
              </button>
              <button className={styles.ctxMenu__item} onClick={() => handleSort(ctxMenu.col, 'desc')} type="button">
                Sort Z → A
              </button>
              <div className={styles.ctxMenu__sep} />
              <button className={styles.ctxMenu__item} onClick={() => { setFreezeRow(ctxMenu.row + 1); setFreezeCol(ctxMenu.col + 1); setCtxMenu(null); }} type="button">
                Freeze panes here
              </button>
              <button className={styles.ctxMenu__item} onClick={() => { setFreezeRow(0); setFreezeCol(0); setCtxMenu(null); }} type="button">
                Unfreeze panes
              </button>
              <div className={styles.ctxMenu__sep} />
              <button className={styles.ctxMenu__item} onClick={() => {
                const addr = `${currentSheet?.name}!${currentSheet?.colHeaders[ctxMenu.col]}${ctxMenu.row + 1}`;
                setEditingComment(addr);
                setCommentText(comments[addr] || '');
                setCtxMenu(null);
              }} type="button">{comments[`${currentSheet?.name}!${currentSheet?.colHeaders[ctxMenu.col]}${ctxMenu.row + 1}`] ? 'Edit comment' : 'Add comment'}</button>
              {comments[`${currentSheet?.name}!${currentSheet?.colHeaders[ctxMenu.col]}${ctxMenu.row + 1}`] && (
                <button className={styles.ctxMenu__item} onClick={() => {
                  const addr = `${currentSheet?.name}!${currentSheet?.colHeaders[ctxMenu.col]}${ctxMenu.row + 1}`;
                  setComments(prev => { const n = { ...prev }; delete n[addr]; return n; });
                  setCtxMenu(null);
                }} type="button" style={{ color: '#DC2626' }}>Delete comment</button>
              )}
              <div className={styles.ctxMenu__sep} />
              <button className={styles.ctxMenu__item} onClick={() => {
                setHiddenRows(prev => new Set([...prev, ctxMenu.row]));
                setCtxMenu(null);
              }} type="button">Hide row</button>
              <button className={styles.ctxMenu__item} onClick={() => {
                setHiddenCols(prev => new Set([...prev, ctxMenu.col]));
                setCtxMenu(null);
              }} type="button">Hide column</button>
              {(hiddenRows.size > 0 || hiddenCols.size > 0) && (
                <button className={styles.ctxMenu__item} onClick={() => {
                  setHiddenRows(new Set());
                  setHiddenCols(new Set());
                  setCtxMenu(null);
                }} type="button">Unhide all</button>
              )}
            </div>
          )}

          {/* Comment editor */}
          {editingComment && (
            <div className={styles.commentEditor}>
              <div className={styles.commentEditor__header}>
                <span>Comment: {editingComment}</span>
                <button onClick={() => setEditingComment(null)} type="button">✕</button>
              </div>
              <textarea
                className={styles.commentEditor__input}
                value={commentText}
                onChange={e => setCommentText(e.target.value)}
                autoFocus
                rows={3}
              />
              <div className={styles.commentEditor__actions}>
                <button className={styles.findBar__btn} onClick={() => {
                  if (commentText.trim()) {
                    setComments(prev => ({ ...prev, [editingComment]: commentText.trim() }));
                  } else {
                    setComments(prev => { const n = { ...prev }; delete n[editingComment]; return n; });
                  }
                  setEditingComment(null);
                }} type="button">Save</button>
                <button className={styles.findBar__btn} onClick={() => setEditingComment(null)} type="button">Cancel</button>
              </div>
            </div>
          )}

          {/* Conditional Formatting panel */}
          {showCondFmt && (
            <div className={styles.condFmtPanel}>
              <div className={styles.condFmtPanel__header}>
                <span>Conditional Formatting Rules</span>
                <button onClick={() => setShowCondFmt(false)} type="button">✕</button>
              </div>
              {condRules.map(rule => (
                <div key={rule.id} className={styles.condFmtRule}>
                  <span style={{ backgroundColor: rule.bgColor, color: rule.color, padding: '1px 6px', borderRadius: 2, fontSize: 10 }}>
                    {rule.type === 'gt' ? `> ${rule.v1}` : rule.type === 'lt' ? `< ${rule.v1}` : rule.type === 'eq' ? `= ${rule.v1}` :
                     rule.type === 'between' ? `${rule.v1} – ${rule.v2}` : rule.type === 'text' ? `"${rule.v1}"` : rule.type === 'duplicate' ? 'Duplicates' : 'Color Scale'}
                  </span>
                  <span style={{ fontSize: 10, color: 'var(--text-muted)' }}>Col: {rule.col === -1 ? 'All' : currentSheet?.colHeaders[rule.col] || rule.col}</span>
                  <button onClick={() => setCondRules(prev => prev.filter(r => r.id !== rule.id))} type="button" style={{ background: 'none', border: 'none', color: '#DC2626', cursor: 'pointer', fontSize: 11 }}>✕</button>
                </div>
              ))}
              <div className={styles.condFmtAdd}>
                <select id="cfType" className={styles.condFmtSelect}><option value="gt">&gt;</option><option value="lt">&lt;</option><option value="eq">=</option><option value="between">Between</option><option value="text">Contains</option><option value="duplicate">Duplicates</option></select>
                <input id="cfV1" placeholder="Value" className={styles.condFmtInput} />
                <input id="cfV2" placeholder="Value 2" className={styles.condFmtInput} style={{ width: 50 }} />
                <input id="cfBg" type="color" defaultValue="#FECACA" style={{ width: 24, height: 24, border: 'none', cursor: 'pointer' }} />
                <input id="cfFc" type="color" defaultValue="#DC2626" style={{ width: 24, height: 24, border: 'none', cursor: 'pointer' }} />
                <button className={styles.findBar__btn} onClick={() => {
                  const type = (document.getElementById('cfType') as HTMLSelectElement).value as CondRule['type'];
                  const v1 = (document.getElementById('cfV1') as HTMLInputElement).value;
                  const v2 = (document.getElementById('cfV2') as HTMLInputElement).value;
                  const bgColor = (document.getElementById('cfBg') as HTMLInputElement).value;
                  const color = (document.getElementById('cfFc') as HTMLInputElement).value;
                  if (!v1 && type !== 'duplicate') return;
                  setCondRules(prev => [...prev, { id: Date.now().toString(), type, col: selectedCell?.col ?? -1, v1, v2, bgColor, color }]);
                }} type="button">Add</button>
              </div>
            </div>
          )}

          {/* Data Validation dialog */}
          {showDataValidation && (
            <div className={styles.condFmtPanel}>
              <div className={styles.condFmtPanel__header}>
                <span>Data Validation</span>
                <button onClick={() => setShowDataValidation(false)} type="button">✕</button>
              </div>
              <div className={styles.condFmtAdd} style={{ flexDirection: 'column', alignItems: 'flex-start', gap: 6 }}>
                <label style={{ fontSize: 10, color: 'var(--text-muted)' }}>Type</label>
                <select id="dvType" className={styles.condFmtSelect} style={{ width: '100%' }}>
                  <option value="list">Dropdown List</option>
                  <option value="number">Number Range</option>
                  <option value="text_length">Text Length</option>
                </select>
                <label style={{ fontSize: 10, color: 'var(--text-muted)' }}>Values (comma-separated for list)</label>
                <input id="dvValues" className={styles.condFmtInput} style={{ width: '100%' }} placeholder="e.g. Yes,No,Maybe" />
                <label style={{ fontSize: 10, color: 'var(--text-muted)' }}>Min / Max (for number)</label>
                <div style={{ display: 'flex', gap: 4 }}>
                  <input id="dvMin" className={styles.condFmtInput} placeholder="Min" style={{ width: 80 }} />
                  <input id="dvMax" className={styles.condFmtInput} placeholder="Max" style={{ width: 80 }} />
                </div>
                <label style={{ fontSize: 10, color: 'var(--text-muted)' }}>Input message</label>
                <input id="dvMsg" className={styles.condFmtInput} style={{ width: '100%' }} placeholder="Help text..." />
                <button className={styles.findBar__btn} onClick={async () => {
                  if (!workbookInfo) return;
                  const cells = getSelectedCells();
                  if (cells.length === 0) { setDataToolMsg('Select cells first'); return; }
                  const type = (document.getElementById('dvType') as HTMLSelectElement).value as 'list' | 'number' | 'text_length';
                  const valStr = (document.getElementById('dvValues') as HTMLInputElement).value;
                  const min = parseFloat((document.getElementById('dvMin') as HTMLInputElement).value) || undefined;
                  const max = parseFloat((document.getElementById('dvMax') as HTMLInputElement).value) || undefined;
                  const message = (document.getElementById('dvMsg') as HTMLInputElement).value;
                  await setDataValidation(workbookInfo.workbook_uuid, cells, {
                    type, values: valStr ? valStr.split(',').map(s => s.trim()) : [], min, max, message,
                  });
                  setDataToolMsg(`Validation set on ${cells.length} cell(s)`);
                  setShowDataValidation(false);
                }} type="button" style={{ marginTop: 4 }}>Apply Validation</button>
              </div>
            </div>
          )}

          {/* Named Ranges dialog */}
          {showNamedRanges && (
            <div className={styles.condFmtPanel}>
              <div className={styles.condFmtPanel__header}>
                <span>Named Ranges</span>
                <button onClick={() => setShowNamedRanges(false)} type="button">✕</button>
              </div>
              {Object.entries(namedRangesData).map(([name, range]) => (
                <div key={name} className={styles.condFmtRule}>
                  <span style={{ fontWeight: 600, fontSize: 11 }}>{name}</span>
                  <span style={{ fontSize: 10, color: 'var(--text-muted)', flex: 1 }}>{range}</span>
                  <button onClick={async () => {
                    if (!workbookInfo) return;
                    await deleteNamedRange(workbookInfo.workbook_uuid, name);
                    setNamedRangesData(prev => { const n = { ...prev }; delete n[name]; return n; });
                  }} type="button" style={{ background: 'none', border: 'none', color: '#DC2626', cursor: 'pointer', fontSize: 11 }}>✕</button>
                </div>
              ))}
              <div className={styles.condFmtAdd}>
                <input id="nrName" className={styles.condFmtInput} placeholder="Name" style={{ width: 80 }} />
                <input id="nrRange" className={styles.condFmtInput} placeholder="Sheet1!A1:B10" style={{ width: 120 }} />
                <button className={styles.findBar__btn} onClick={async () => {
                  if (!workbookInfo) return;
                  const name = (document.getElementById('nrName') as HTMLInputElement).value.trim();
                  const range = (document.getElementById('nrRange') as HTMLInputElement).value.trim();
                  if (!name || !range) return;
                  await setNamedRange(workbookInfo.workbook_uuid, name, range);
                  setNamedRangesData(prev => ({ ...prev, [name]: range }));
                  (document.getElementById('nrName') as HTMLInputElement).value = '';
                  (document.getElementById('nrRange') as HTMLInputElement).value = '';
                }} type="button">Add</button>
              </div>
            </div>
          )}

          {/* Goal Seek dialog */}
          {showGoalSeek && (
            <div className={styles.condFmtPanel}>
              <div className={styles.condFmtPanel__header}>
                <span>Goal Seek</span>
                <button onClick={() => { setShowGoalSeek(false); setGoalSeekResult(null); }} type="button">✕</button>
              </div>
              <div className={styles.condFmtAdd} style={{ flexDirection: 'column', alignItems: 'flex-start', gap: 6 }}>
                <label style={{ fontSize: 10, color: 'var(--text-muted)' }}>Target cell (e.g. Sheet1!B10)</label>
                <input id="gsTarget" className={styles.condFmtInput} style={{ width: '100%' }} placeholder="Sheet1!B10" />
                <label style={{ fontSize: 10, color: 'var(--text-muted)' }}>Goal value</label>
                <input id="gsGoal" className={styles.condFmtInput} style={{ width: '100%' }} placeholder="1000" />
                <label style={{ fontSize: 10, color: 'var(--text-muted)' }}>Changing cell (e.g. Sheet1!A1)</label>
                <input id="gsChanging" className={styles.condFmtInput} style={{ width: '100%' }} placeholder="Sheet1!A1" />
                <button className={styles.findBar__btn} onClick={async () => {
                  if (!workbookInfo) return;
                  const target = (document.getElementById('gsTarget') as HTMLInputElement).value.trim();
                  const goal = parseFloat((document.getElementById('gsGoal') as HTMLInputElement).value);
                  const changing = (document.getElementById('gsChanging') as HTMLInputElement).value.trim();
                  if (!target || isNaN(goal) || !changing) return;
                  try {
                    const res = await goalSeek(workbookInfo.workbook_uuid, target, goal, changing);
                    setGoalSeekResult(`Result: ${res.result_value.toFixed(4)} (achieved: ${res.achieved_value}, diff: ${res.difference.toFixed(6)})`);
                    fetchGrid(workbookInfo.workbook_uuid);
                  } catch (e) {
                    setGoalSeekResult(`Error: ${e instanceof Error ? e.message : 'Failed'}`);
                  }
                }} type="button">Seek</button>
                {goalSeekResult && <div style={{ fontSize: 11, color: 'var(--text-primary)', marginTop: 4 }}>{goalSeekResult}</div>}
              </div>
            </div>
          )}

          {/* Text to Columns dialog */}
          {showTextToCol && (
            <div className={styles.condFmtPanel}>
              <div className={styles.condFmtPanel__header}>
                <span>Text to Columns</span>
                <button onClick={() => setShowTextToCol(false)} type="button">✕</button>
              </div>
              <div className={styles.condFmtAdd} style={{ flexDirection: 'column', alignItems: 'flex-start', gap: 6 }}>
                <label style={{ fontSize: 10, color: 'var(--text-muted)' }}>Select a column in the grid, then choose delimiter</label>
                <select id="ttcDelim" className={styles.condFmtSelect} style={{ width: '100%' }}>
                  <option value=",">Comma (,)</option>
                  <option value=";">Semicolon (;)</option>
                  <option value=" ">Space</option>
                  <option value="\t">Tab</option>
                  <option value="|">Pipe (|)</option>
                  <option value="-">Dash (-)</option>
                </select>
                <button className={styles.findBar__btn} onClick={async () => {
                  if (!workbookInfo || !currentSheet || selectedCell === null) { setDataToolMsg('Select a cell in the column first'); return; }
                  const delimiter = (document.getElementById('ttcDelim') as HTMLSelectElement).value;
                  try {
                    const res = await textToColumns(workbookInfo.workbook_uuid, currentSheet.name, selectedCell.col + 1, delimiter);
                    setDataToolMsg(`Split into ${res.split_into_columns} columns (${res.cells_written} cells)`);
                    fetchGrid(workbookInfo.workbook_uuid);
                    setShowTextToCol(false);
                  } catch (e) {
                    setDataToolMsg(`Error: ${e instanceof Error ? e.message : 'Failed'}`);
                  }
                }} type="button">Split</button>
              </div>
            </div>
          )}

          {/* Remove Duplicates dialog */}
          {showRemoveDups && (
            <div className={styles.condFmtPanel}>
              <div className={styles.condFmtPanel__header}>
                <span>Remove Duplicates</span>
                <button onClick={() => setShowRemoveDups(false)} type="button">✕</button>
              </div>
              <div className={styles.condFmtAdd} style={{ flexDirection: 'column', alignItems: 'flex-start', gap: 6 }}>
                <label style={{ fontSize: 10, color: 'var(--text-muted)' }}>Removes duplicate rows (keeps first occurrence). Compares all columns.</label>
                <button className={styles.findBar__btn} onClick={async () => {
                  if (!workbookInfo || !currentSheet) return;
                  try {
                    const res = await removeDuplicates(workbookInfo.workbook_uuid, currentSheet.name);
                    setDataToolMsg(`Removed ${res.removed} duplicate row(s). ${res.remaining_rows} rows remaining.`);
                    fetchGrid(workbookInfo.workbook_uuid);
                    setShowRemoveDups(false);
                  } catch (e) {
                    setDataToolMsg(`Error: ${e instanceof Error ? e.message : 'Failed'}`);
                  }
                }} type="button" style={{ color: '#DC2626' }}>Remove Duplicates Now</button>
              </div>
            </div>
          )}

          {/* Data tool message toast */}
          {dataToolMsg && (
            <div className={styles.dataToolToast} onClick={() => setDataToolMsg(null)}>
              {dataToolMsg}
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

          {/* Status bar */}
          <div className={styles.statusBar}>
            <div className={styles.statusBar__left}>
              {selectedCell && (
                <span>
                  {selRange
                    ? `${normalizeRange(selRange).r2 - normalizeRange(selRange).r1 + 1}R × ${normalizeRange(selRange).c2 - normalizeRange(selRange).c1 + 1}C`
                    : getCellAddr(selectedCell.row, selectedCell.col)}
                </span>
              )}
              {freezeRow > 0 && <span className={styles.statusBar__badge}>Frozen {freezeRow}R {freezeCol}C</span>}
            </div>
            <div className={styles.statusBar__right}>
              {statusBarInfo && statusBarInfo.numCount > 0 && (
                <>
                  <span>Sum: <strong>{statusBarInfo.sum.toLocaleString(undefined, { maximumFractionDigits: 2 })}</strong></span>
                  <span>Avg: <strong>{statusBarInfo.avg.toLocaleString(undefined, { maximumFractionDigits: 2 })}</strong></span>
                  <span>Count: <strong>{statusBarInfo.count}</strong></span>
                  <span>Min: <strong>{statusBarInfo.min.toLocaleString(undefined, { maximumFractionDigits: 2 })}</strong></span>
                  <span>Max: <strong>{statusBarInfo.max.toLocaleString(undefined, { maximumFractionDigits: 2 })}</strong></span>
                </>
              )}
              {statusBarInfo && statusBarInfo.numCount === 0 && statusBarInfo.count > 0 && (
                <span>Count: <strong>{statusBarInfo.count}</strong></span>
              )}
            </div>
          </div>
        </>
      )}
    </div>
  );
}
