export interface CellChange {
  cell: string;   // e.g. "Assumptions!B4"
  before: unknown;
  after: unknown;
}

type RefreshListener = () => void;
type DiffListener = (changes: CellChange[]) => void;

const refreshListeners: Set<RefreshListener> = new Set();
const diffListeners: Set<DiffListener> = new Set();

export function onGridRefresh(fn: RefreshListener): () => void {
  refreshListeners.add(fn);
  return () => refreshListeners.delete(fn);
}

export function emitGridRefresh(): void {
  refreshListeners.forEach(fn => fn());
}

export function onGridDiff(fn: DiffListener): () => void {
  diffListeners.add(fn);
  return () => diffListeners.delete(fn);
}

export function emitGridDiff(changes: CellChange[]): void {
  diffListeners.forEach(fn => fn(changes));
}
