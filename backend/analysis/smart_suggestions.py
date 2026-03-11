"""
Smart Suggestions Engine
Proaktywne wykrywanie typowych błędów modelowania i sugestie usprawnień:
- Brakujące dyskontowanie w DCF
- Niespójne stopy wzrostu
- Hardcoded wartości które powinny być wzorami
- Brak walidacji (np. suma wag ≠ 100%)
- Potencjalne duplikaty formuł do konsolidacji
- Missing links między arkuszami
"""
import logging
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Set, Tuple
from collections import Counter

import networkx as nx

from parser.xlsx_parser import WorkbookData

logger = logging.getLogger(__name__)


@dataclass
class Suggestion:
    priority: str
    category: str
    title: str
    description: str
    affected_cells: List[str]
    suggested_action: str
    estimated_effort: str
    confidence: float


@dataclass
class SuggestionsReport:
    total_suggestions: int
    high_priority: int
    medium_priority: int
    low_priority: int
    suggestions: List[Suggestion]
    model_maturity_score: float


class SmartSuggestionsEngine:
    def __init__(self, graph: nx.DiGraph, workbook_data: WorkbookData) -> None:
        self.graph = graph
        self.workbook_data = workbook_data
        self.suggestions: List[Suggestion] = []

    def run(self) -> SuggestionsReport:
        self.suggestions = []
        self._detect_missing_discount_factors()
        self._detect_inconsistent_growth_rates()
        self._detect_hardcoded_patterns()
        self._detect_duplicate_formulas()
        self._detect_missing_cross_sheet_links()
        self._detect_weight_validation()
        self._detect_unprotected_key_cells()
        self._detect_date_convention_issues()

        high = sum(1 for s in self.suggestions if s.priority == "high")
        medium = sum(1 for s in self.suggestions if s.priority == "medium")
        low = sum(1 for s in self.suggestions if s.priority == "low")

        max_score = 100.0
        penalty = high * 12 + medium * 5 + low * 1
        maturity = max(0.0, min(100.0, max_score - penalty))

        logger.info(f"Smart suggestions: {len(self.suggestions)} suggestions, maturity={maturity:.0f}/100")

        return SuggestionsReport(
            total_suggestions=len(self.suggestions),
            high_priority=high,
            medium_priority=medium,
            low_priority=low,
            suggestions=self.suggestions,
            model_maturity_score=maturity,
        )

    def _detect_missing_discount_factors(self) -> None:
        dcf_keywords = {"dcf", "npv", "discount", "wacc", "irr", "present value", "pv"}
        dcf_sheets: Set[str] = set()
        for sheet in self.workbook_data.sheets:
            if any(kw in sheet.lower() for kw in dcf_keywords):
                dcf_sheets.add(sheet)

        if not dcf_sheets:
            return

        for sheet in dcf_sheets:
            has_discount_ref = False
            sheet_cells = [
                (addr, cell) for addr, cell in self.workbook_data.cells.items()
                if cell.sheet_name == sheet and cell.formula
            ]
            for addr, cell in sheet_cells:
                if not cell.formula:
                    continue
                formula_lower = cell.formula.lower()
                if any(kw in formula_lower for kw in ["discount", "wacc", "1+", "/(1", "npv"]):
                    has_discount_ref = True
                    break
                for pred in self.graph.predecessors(addr):
                    pred_nr = self.graph.nodes.get(pred, {}).get("named_range", "")
                    if pred_nr and any(kw in pred_nr.lower() for kw in dcf_keywords):
                        has_discount_ref = True
                        break

            if not has_discount_ref and len(sheet_cells) > 3:
                self.suggestions.append(Suggestion(
                    priority="high",
                    category="missing_discount",
                    title=f"No discounting detected in '{sheet}'",
                    description=f"Sheet '{sheet}' appears to be a DCF model but no discount factor or NPV formula was found",
                    affected_cells=[addr for addr, _ in sheet_cells[:5]],
                    suggested_action="Add discount factor calculation using WACC or required rate of return",
                    estimated_effort="medium",
                    confidence=0.7,
                ))

    def _detect_inconsistent_growth_rates(self) -> None:
        growth_cells: List[Tuple[str, float]] = []
        for node in self.graph.nodes():
            nd = self.graph.nodes[node]
            nr = nd.get("named_range", "")
            val = nd.get("value")
            if not isinstance(val, (int, float)):
                continue
            if nr and any(kw in nr.lower() for kw in ["growth", "rate", "cagr", "yoy"]):
                growth_cells.append((node, float(val)))
                continue
            if nd.get("is_hardcoded") and -1.0 < val < 1.0 and val != 0:
                successors = list(self.graph.successors(node))
                for s in successors:
                    s_formula = self.graph.nodes.get(s, {}).get("formula", "")
                    if s_formula and "*" in s_formula:
                        growth_cells.append((node, float(val)))
                        break

        if len(growth_cells) < 2:
            return

        values = [v for _, v in growth_cells]
        mean_val = sum(values) / len(values)
        for cell, val in growth_cells:
            if abs(val - mean_val) > abs(mean_val) * 3 and abs(val - mean_val) > 0.05:
                self.suggestions.append(Suggestion(
                    priority="medium",
                    category="inconsistent_growth",
                    title=f"Unusual growth rate in {cell}",
                    description=f"Value {val:.2%} is significantly different from average growth rate {mean_val:.2%}",
                    affected_cells=[cell],
                    suggested_action="Verify this growth rate is intentional and document the reasoning",
                    estimated_effort="low",
                    confidence=0.6,
                ))

    def _detect_hardcoded_patterns(self) -> None:
        sheet_sequences: Dict[str, List[Tuple[str, int, int, float]]] = {}
        for addr, cell in self.workbook_data.cells.items():
            if not cell.is_hardcoded or not isinstance(cell.value, (int, float)):
                continue
            s = cell.sheet_name
            if s not in sheet_sequences:
                sheet_sequences[s] = []
            sheet_sequences[s].append((addr, cell.row, cell.col, float(cell.value)))

        for sheet, cells in sheet_sequences.items():
            by_row: Dict[int, List[Tuple[str, int, float]]] = {}
            for addr, row, col, val in cells:
                by_row.setdefault(row, []).append((addr, col, val))

            for row, row_cells in by_row.items():
                if len(row_cells) < 4:
                    continue
                row_cells.sort(key=lambda x: x[1])
                vals = [v for _, _, v in row_cells]

                if len(set(vals)) == 1 and vals[0] != 0:
                    self.suggestions.append(Suggestion(
                        priority="medium",
                        category="constant_row",
                        title=f"Row {row} in '{sheet}' has identical values",
                        description=f"{len(vals)} cells all contain {vals[0]} — could be replaced with a single reference",
                        affected_cells=[addr for addr, _, _ in row_cells],
                        suggested_action="Use a single assumption cell and reference it across the row",
                        estimated_effort="low",
                        confidence=0.8,
                    ))

                is_linear = True
                if len(vals) >= 3:
                    diffs = [vals[i+1] - vals[i] for i in range(len(vals)-1)]
                    if len(set(round(d, 6) for d in diffs)) == 1 and diffs[0] != 0:
                        self.suggestions.append(Suggestion(
                            priority="low",
                            category="linear_pattern",
                            title=f"Linear pattern in row {row} of '{sheet}'",
                            description=f"Values increase by {diffs[0]} each period — could be a formula",
                            affected_cells=[addr for addr, _, _ in row_cells],
                            suggested_action=f"Replace with formula: =previous_cell+{diffs[0]}",
                            estimated_effort="low",
                            confidence=0.85,
                        ))

    def _detect_duplicate_formulas(self) -> None:
        formula_map: Dict[str, List[str]] = {}
        for addr, cell in self.workbook_data.cells.items():
            if not cell.formula:
                continue
            normalized = re.sub(r'[A-Z]+[0-9]+', 'REF', cell.formula.upper())
            normalized = re.sub(r'\s+', '', normalized)
            formula_map.setdefault(normalized, []).append(addr)

        for pattern, cells in formula_map.items():
            if len(cells) < 5:
                continue
            sheets = set(self.workbook_data.cells[c].sheet_name for c in cells if c in self.workbook_data.cells)
            if len(sheets) > 1:
                self.suggestions.append(Suggestion(
                    priority="low",
                    category="duplicate_formula",
                    title=f"Same formula pattern repeated {len(cells)} times across {len(sheets)} sheets",
                    description=f"Consider creating a helper function or consolidating into a single calculation block",
                    affected_cells=cells[:10],
                    suggested_action="Review if these can be consolidated or if a LAMBDA/helper column would simplify",
                    estimated_effort="medium",
                    confidence=0.5,
                ))

    def _detect_missing_cross_sheet_links(self) -> None:
        sheet_connections: Dict[str, Set[str]] = {}
        for sheet in self.workbook_data.sheets:
            sheet_connections[sheet] = set()

        for addr, cell in self.workbook_data.cells.items():
            if not cell.formula:
                continue
            for pred in self.graph.predecessors(addr):
                pred_sheet = self.graph.nodes.get(pred, {}).get("sheet_name", "")
                if pred_sheet and pred_sheet != cell.sheet_name:
                    sheet_connections[cell.sheet_name].add(pred_sheet)

        isolated_sheets = [
            s for s in self.workbook_data.sheets
            if not sheet_connections.get(s) and
            sum(1 for c in self.workbook_data.cells.values() if c.sheet_name == s) > 5
        ]

        other_sheets_with_content = [
            s for s in self.workbook_data.sheets
            if sum(1 for c in self.workbook_data.cells.values() if c.sheet_name == s) > 5
        ]

        if len(isolated_sheets) > 0 and len(other_sheets_with_content) > 1:
            for sheet in isolated_sheets:
                has_outputs_to = any(sheet in conns for conns in sheet_connections.values())
                if not has_outputs_to:
                    self.suggestions.append(Suggestion(
                        priority="medium",
                        category="isolated_sheet",
                        title=f"Sheet '{sheet}' has no cross-sheet connections",
                        description=f"This sheet doesn't reference or get referenced by other sheets",
                        affected_cells=[],
                        suggested_action="Verify this sheet is needed; if so, connect its outputs to the main model",
                        estimated_effort="medium",
                        confidence=0.6,
                    ))

    def _detect_weight_validation(self) -> None:
        weight_keywords = {"weight", "alloc", "split", "share", "pct", "percent", "%"}

        for sheet in self.workbook_data.sheets:
            col_groups: Dict[int, List[Tuple[str, float]]] = {}
            for addr, cell in self.workbook_data.cells.items():
                if cell.sheet_name != sheet:
                    continue
                if not isinstance(cell.value, (int, float)):
                    continue
                nr = cell.named_range or ""
                if not any(kw in nr.lower() for kw in weight_keywords):
                    if not (0 < abs(cell.value) < 1):
                        continue
                col_groups.setdefault(cell.col, []).append((addr, float(cell.value)))

            for col, cells in col_groups.items():
                if len(cells) < 3:
                    continue
                values = [v for _, v in cells]
                total = sum(values)
                if 0.95 < total < 1.05 and total != 1.0:
                    self.suggestions.append(Suggestion(
                        priority="medium",
                        category="weight_validation",
                        title=f"Weights in column {col} of '{sheet}' sum to {total:.4f} (not exactly 1.0)",
                        description="Weight/allocation percentages should typically sum to exactly 100%",
                        affected_cells=[addr for addr, _ in cells],
                        suggested_action="Add a validation row that checks SUM = 100% and flag if not",
                        estimated_effort="low",
                        confidence=0.75,
                    ))

    def _detect_unprotected_key_cells(self) -> None:
        for node in self.graph.nodes():
            nd = self.graph.nodes[node]
            pr = nd.get("pagerank", 0.0)
            if pr < 0.5:
                continue
            in_degree = self.graph.in_degree(node)
            out_degree = self.graph.out_degree(node)
            if isinstance(out_degree, int) and out_degree > 5 and nd.get("is_hardcoded"):
                self.suggestions.append(Suggestion(
                    priority="high",
                    category="unprotected_key_input",
                    title=f"Critical input {node} drives {out_degree} cells but is unprotected",
                    description=f"Cell {node} (PageRank={pr:.2f}) is a key driver. Accidental changes could cascade through the model.",
                    affected_cells=[node],
                    suggested_action="Add data validation, named range, and/or cell protection to prevent accidental edits",
                    estimated_effort="low",
                    confidence=0.9,
                ))

    def _detect_date_convention_issues(self) -> None:
        date_cells: Dict[str, List[Tuple[str, Any]]] = {}
        for addr, cell in self.workbook_data.cells.items():
            if cell.data_type == "date" or (isinstance(cell.value, (int, float)) and 40000 < cell.value < 50000):
                date_cells.setdefault(cell.sheet_name, []).append((addr, cell.value))

        for sheet, cells in date_cells.items():
            if len(cells) < 3:
                continue
            values = sorted([v for _, v in cells if isinstance(v, (int, float))])
            if len(values) < 3:
                continue
            diffs = [values[i+1] - values[i] for i in range(len(values)-1)]
            unique_diffs = set(int(d) for d in diffs if d > 0)

            has_monthly = any(28 <= d <= 31 for d in unique_diffs)
            has_quarterly = any(89 <= d <= 92 for d in unique_diffs)
            has_annual = any(364 <= d <= 366 for d in unique_diffs)

            mixed_count = sum([has_monthly, has_quarterly, has_annual])
            if mixed_count > 1:
                self.suggestions.append(Suggestion(
                    priority="medium",
                    category="mixed_date_convention",
                    title=f"Mixed date intervals in '{sheet}'",
                    description=f"Found {'monthly' if has_monthly else ''} {'quarterly' if has_quarterly else ''} {'annual' if has_annual else ''} intervals in the same sheet",
                    affected_cells=[addr for addr, _ in cells[:5]],
                    suggested_action="Standardize date convention within each sheet or clearly separate different periodicities",
                    estimated_effort="medium",
                    confidence=0.65,
                ))
