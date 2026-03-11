"""
Version Diff Engine
Porównanie dwóch wersji tego samego workbooka — co się zmieniło, gdzie, o ile.
Przydatne do code review modeli finansowych i audytu zmian między wersjami.
"""
import logging
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

from parser.xlsx_parser import XLSXParser, WorkbookData

logger = logging.getLogger(__name__)


@dataclass
class CellDiff:
    cell: str
    sheet: str
    change_type: str
    old_value: Any
    new_value: Any
    old_formula: Optional[str]
    new_formula: Optional[str]
    delta: Optional[float]
    delta_pct: Optional[float]
    impact_score: float


@dataclass
class SheetDiff:
    sheet_name: str
    cells_added: int
    cells_removed: int
    cells_modified: int
    formula_changes: int
    value_changes: int
    max_delta_pct: Optional[float]


@dataclass
class VersionDiffReport:
    file_a: str
    file_b: str
    total_changes: int
    sheets_added: List[str]
    sheets_removed: List[str]
    sheet_diffs: List[SheetDiff]
    cell_diffs: List[CellDiff]
    high_impact_changes: List[CellDiff]
    summary: str


class VersionDiffer:
    def __init__(self) -> None:
        self.parser = XLSXParser()

    def diff_files(self, file_a: str, file_b: str) -> VersionDiffReport:
        wb_a = self.parser.parse(file_a)
        wb_b = self.parser.parse(file_b)
        return self.diff_workbooks(wb_a, wb_b, file_a, file_b)

    def diff_workbooks(
        self,
        wb_a: WorkbookData,
        wb_b: WorkbookData,
        label_a: str = "version_a",
        label_b: str = "version_b",
    ) -> VersionDiffReport:
        sheets_a = set(wb_a.sheets)
        sheets_b = set(wb_b.sheets)
        sheets_added = sorted(sheets_b - sheets_a)
        sheets_removed = sorted(sheets_a - sheets_b)

        cell_diffs: List[CellDiff] = []
        all_cells = set(wb_a.cells.keys()) | set(wb_b.cells.keys())

        for addr in all_cells:
            cell_a = wb_a.cells.get(addr)
            cell_b = wb_b.cells.get(addr)

            if cell_a is None and cell_b is not None:
                cell_diffs.append(CellDiff(
                    cell=addr,
                    sheet=cell_b.sheet_name,
                    change_type="added",
                    old_value=None,
                    new_value=cell_b.value,
                    old_formula=None,
                    new_formula=cell_b.formula,
                    delta=None,
                    delta_pct=None,
                    impact_score=0.0,
                ))
            elif cell_a is not None and cell_b is None:
                cell_diffs.append(CellDiff(
                    cell=addr,
                    sheet=cell_a.sheet_name,
                    change_type="removed",
                    old_value=cell_a.value,
                    new_value=None,
                    old_formula=cell_a.formula,
                    new_formula=None,
                    delta=None,
                    delta_pct=None,
                    impact_score=0.0,
                ))
            elif cell_a is not None and cell_b is not None:
                val_changed = cell_a.value != cell_b.value
                formula_changed = cell_a.formula != cell_b.formula

                if not val_changed and not formula_changed:
                    continue

                delta = None
                delta_pct = None
                if isinstance(cell_a.value, (int, float)) and isinstance(cell_b.value, (int, float)):
                    delta = cell_b.value - cell_a.value
                    if cell_a.value != 0:
                        delta_pct = round((cell_b.value - cell_a.value) / abs(cell_a.value) * 100, 2)

                change_type = "modified"
                if formula_changed and not val_changed:
                    change_type = "formula_changed"
                elif val_changed and not formula_changed:
                    change_type = "value_changed"

                impact = 0.0
                if delta_pct is not None:
                    impact = abs(delta_pct)
                if formula_changed:
                    impact += 10.0

                cell_diffs.append(CellDiff(
                    cell=addr,
                    sheet=cell_a.sheet_name,
                    change_type=change_type,
                    old_value=cell_a.value,
                    new_value=cell_b.value,
                    old_formula=cell_a.formula,
                    new_formula=cell_b.formula,
                    delta=delta,
                    delta_pct=delta_pct,
                    impact_score=impact,
                ))

        cell_diffs.sort(key=lambda d: d.impact_score, reverse=True)

        sheet_diff_map: Dict[str, SheetDiff] = {}
        for sheet in sorted(sheets_a | sheets_b):
            sheet_diff_map[sheet] = SheetDiff(
                sheet_name=sheet,
                cells_added=0,
                cells_removed=0,
                cells_modified=0,
                formula_changes=0,
                value_changes=0,
                max_delta_pct=None,
            )

        for d in cell_diffs:
            sd = sheet_diff_map.get(d.sheet)
            if not sd:
                continue
            if d.change_type == "added":
                sd.cells_added += 1
            elif d.change_type == "removed":
                sd.cells_removed += 1
            else:
                sd.cells_modified += 1
            if d.old_formula != d.new_formula:
                sd.formula_changes += 1
            if d.old_value != d.new_value:
                sd.value_changes += 1
            if d.delta_pct is not None:
                if sd.max_delta_pct is None or abs(d.delta_pct) > abs(sd.max_delta_pct):
                    sd.max_delta_pct = d.delta_pct

        sheet_diffs = [sd for sd in sheet_diff_map.values()
                       if sd.cells_added + sd.cells_removed + sd.cells_modified > 0]
        sheet_diffs.sort(key=lambda s: s.cells_modified + s.cells_added + s.cells_removed, reverse=True)

        high_impact = [d for d in cell_diffs if d.impact_score > 20]

        parts = [f"{len(cell_diffs)} cell changes"]
        if sheets_added:
            parts.append(f"{len(sheets_added)} sheets added")
        if sheets_removed:
            parts.append(f"{len(sheets_removed)} sheets removed")
        if high_impact:
            parts.append(f"{len(high_impact)} high-impact changes")
        summary = f"Version diff: {', '.join(parts)}"

        logger.info(summary)

        return VersionDiffReport(
            file_a=label_a,
            file_b=label_b,
            total_changes=len(cell_diffs),
            sheets_added=sheets_added,
            sheets_removed=sheets_removed,
            sheet_diffs=sheet_diffs,
            cell_diffs=cell_diffs,
            high_impact_changes=high_impact,
            summary=summary,
        )
