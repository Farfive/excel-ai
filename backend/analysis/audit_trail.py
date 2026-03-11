"""
Audit Trail
Pełna historia zmian w workbooku — kto (agent), co, kiedy, dlaczego.
Snapshot before/after, rollback capability, diff generation.
"""
import logging
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional
from copy import deepcopy

logger = logging.getLogger(__name__)


@dataclass
class ChangeRecord:
    change_id: str
    timestamp: str
    cell: str
    sheet: str
    old_value: Any
    new_value: Any
    old_formula: Optional[str]
    new_formula: Optional[str]
    reason: str
    agent_step: Optional[int]
    approved_by: str
    rollback_available: bool = True


@dataclass
class Snapshot:
    snapshot_id: str
    timestamp: str
    label: str
    cell_values: Dict[str, Any]
    cell_formulas: Dict[str, Optional[str]]
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class AuditReport:
    total_changes: int
    changes: List[ChangeRecord]
    snapshots: List[Snapshot]
    cells_modified: int
    risk_changes: List[ChangeRecord]


class AuditTrail:
    def __init__(self) -> None:
        self.changes: List[ChangeRecord] = []
        self.snapshots: List[Snapshot] = []
        self._change_counter = 0

    def take_snapshot(
        self,
        label: str,
        cell_values: Dict[str, Any],
        cell_formulas: Dict[str, Optional[str]],
        metadata: Optional[Dict[str, Any]] = None,
    ) -> Snapshot:
        self._change_counter += 1
        snapshot = Snapshot(
            snapshot_id=f"snap_{self._change_counter}_{datetime.utcnow().strftime('%H%M%S')}",
            timestamp=datetime.utcnow().isoformat(),
            label=label,
            cell_values=deepcopy(cell_values),
            cell_formulas=deepcopy(cell_formulas),
            metadata=metadata or {},
        )
        self.snapshots.append(snapshot)
        logger.info(f"Snapshot taken: {snapshot.snapshot_id} ({label})")
        return snapshot

    def record_change(
        self,
        cell: str,
        sheet: str,
        old_value: Any,
        new_value: Any,
        old_formula: Optional[str],
        new_formula: Optional[str],
        reason: str,
        agent_step: Optional[int] = None,
        approved_by: str = "user",
    ) -> ChangeRecord:
        self._change_counter += 1
        record = ChangeRecord(
            change_id=f"chg_{self._change_counter}_{datetime.utcnow().strftime('%H%M%S')}",
            timestamp=datetime.utcnow().isoformat(),
            cell=cell,
            sheet=sheet,
            old_value=old_value,
            new_value=new_value,
            old_formula=old_formula,
            new_formula=new_formula,
            reason=reason,
            agent_step=agent_step,
            approved_by=approved_by,
        )
        self.changes.append(record)
        logger.info(f"Change recorded: {cell} {old_value} → {new_value}")
        return record

    def record_batch(
        self,
        writes: List[Dict[str, Any]],
        reason: str,
        approved_by: str = "user",
    ) -> List[ChangeRecord]:
        records = []
        for w in writes:
            r = self.record_change(
                cell=w.get("cell", ""),
                sheet=w.get("sheet", ""),
                old_value=w.get("old_value"),
                new_value=w.get("new_value"),
                old_formula=w.get("old_formula"),
                new_formula=w.get("new_formula"),
                reason=reason,
                agent_step=w.get("step"),
                approved_by=approved_by,
            )
            records.append(r)
        return records

    def get_cell_history(self, cell: str) -> List[ChangeRecord]:
        return [c for c in self.changes if c.cell == cell]

    def get_changes_since(self, since_iso: str) -> List[ChangeRecord]:
        return [c for c in self.changes if c.timestamp >= since_iso]

    def rollback_to_snapshot(self, snapshot_id: str) -> Optional[Snapshot]:
        target = None
        for s in self.snapshots:
            if s.snapshot_id == snapshot_id:
                target = s
                break
        if not target:
            logger.warning(f"Snapshot {snapshot_id} not found")
            return None

        self.record_change(
            cell="*",
            sheet="*",
            old_value="current_state",
            new_value=f"rollback_to_{snapshot_id}",
            old_formula=None,
            new_formula=None,
            reason=f"Rollback to snapshot: {target.label}",
            approved_by="system",
        )
        logger.info(f"Rollback to snapshot: {snapshot_id} ({target.label})")
        return target

    def diff_snapshots(self, snap_id_a: str, snap_id_b: str) -> List[Dict[str, Any]]:
        snap_a = None
        snap_b = None
        for s in self.snapshots:
            if s.snapshot_id == snap_id_a:
                snap_a = s
            if s.snapshot_id == snap_id_b:
                snap_b = s

        if not snap_a or not snap_b:
            return []

        diffs: List[Dict[str, Any]] = []
        all_cells = set(snap_a.cell_values.keys()) | set(snap_b.cell_values.keys())

        for cell in sorted(all_cells):
            val_a = snap_a.cell_values.get(cell)
            val_b = snap_b.cell_values.get(cell)
            form_a = snap_a.cell_formulas.get(cell)
            form_b = snap_b.cell_formulas.get(cell)

            if val_a != val_b or form_a != form_b:
                delta = None
                delta_pct = None
                if isinstance(val_a, (int, float)) and isinstance(val_b, (int, float)):
                    delta = val_b - val_a
                    if val_a != 0:
                        delta_pct = round((val_b - val_a) / abs(val_a) * 100, 2)

                diffs.append({
                    "cell": cell,
                    "value_before": val_a,
                    "value_after": val_b,
                    "formula_before": form_a,
                    "formula_after": form_b,
                    "delta": delta,
                    "delta_pct": delta_pct,
                    "change_type": "modified" if cell in snap_a.cell_values and cell in snap_b.cell_values
                                   else ("added" if cell not in snap_a.cell_values else "removed"),
                })

        return diffs

    def generate_report(self) -> AuditReport:
        cells_modified = len(set(c.cell for c in self.changes if c.cell != "*"))

        risk_changes = []
        for c in self.changes:
            if c.cell == "*":
                continue
            if isinstance(c.old_value, (int, float)) and isinstance(c.new_value, (int, float)):
                if c.old_value != 0:
                    delta_pct = abs((c.new_value - c.old_value) / c.old_value * 100)
                    if delta_pct > 20:
                        risk_changes.append(c)

        return AuditReport(
            total_changes=len(self.changes),
            changes=self.changes,
            snapshots=self.snapshots,
            cells_modified=cells_modified,
            risk_changes=risk_changes,
        )

    def to_markdown(self) -> str:
        lines = [
            f"## Audit Trail — {datetime.utcnow().strftime('%Y-%m-%d %H:%M UTC')}",
            "",
            f"Total changes: {len(self.changes)} | Snapshots: {len(self.snapshots)}",
            "",
            "| # | Time | Cell | Old | New | Reason | Approved |",
            "|---|------|------|-----|-----|--------|----------|",
        ]
        for i, c in enumerate(self.changes, 1):
            time_short = c.timestamp[11:19] if len(c.timestamp) > 19 else c.timestamp
            lines.append(
                f"| {i} | {time_short} | {c.cell} | {c.old_value} | {c.new_value} | {c.reason[:40]} | {c.approved_by} |"
            )
        return "\n".join(lines)
