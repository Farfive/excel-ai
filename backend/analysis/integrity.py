"""
Model Integrity Checker
Wykrywa typowe błędy w modelach finansowych:
- Broken references (formuły odwołujące się do pustych komórek)
- Dangling named ranges
- Inconsistent time series (różne długości wierszy w tej samej tabeli)
- Hardcoded values w sheet'ach z formułami (powinny być w Assumptions)
- Unused inputs (hardcoded cells bez downstream dependents)
- Circular references
- Inconsistent sign conventions (+/- w revenue vs costs)
"""
import logging
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Set

import networkx as nx

from parser.xlsx_parser import WorkbookData

logger = logging.getLogger(__name__)


@dataclass
class IntegrityIssue:
    severity: str
    category: str
    cell: str
    sheet: str
    message: str
    suggestion: str
    details: Optional[Dict[str, Any]] = None


@dataclass
class IntegrityReport:
    total_issues: int
    critical: int
    warning: int
    info: int
    issues: List[IntegrityIssue]
    model_health_score: float
    summary: str


class IntegrityChecker:
    def __init__(self, graph: nx.DiGraph, workbook_data: WorkbookData) -> None:
        self.graph = graph
        self.workbook_data = workbook_data
        self.issues: List[IntegrityIssue] = []

    def run(self) -> IntegrityReport:
        self.issues = []
        self._check_broken_references()
        self._check_dangling_named_ranges()
        self._check_hardcoded_in_formula_sheets()
        self._check_unused_inputs()
        self._check_circular_references()
        self._check_inconsistent_time_series()
        self._check_sign_conventions()
        self._check_formula_complexity()

        critical = sum(1 for i in self.issues if i.severity == "critical")
        warning = sum(1 for i in self.issues if i.severity == "warning")
        info = sum(1 for i in self.issues if i.severity == "info")

        max_score = 100.0
        penalty = critical * 15 + warning * 5 + info * 1
        health = max(0.0, min(100.0, max_score - penalty))

        total_cells = len(self.workbook_data.cells)
        summary_parts = []
        if critical > 0:
            summary_parts.append(f"{critical} critical issues found")
        if warning > 0:
            summary_parts.append(f"{warning} warnings")
        if info > 0:
            summary_parts.append(f"{info} informational notes")
        if not summary_parts:
            summary_parts.append("No issues found")
        summary = f"Model health: {health:.0f}/100. {total_cells} cells analyzed. " + ", ".join(summary_parts) + "."

        logger.info(f"Integrity check: {len(self.issues)} issues, health={health:.0f}/100")

        return IntegrityReport(
            total_issues=len(self.issues),
            critical=critical,
            warning=warning,
            info=info,
            issues=self.issues,
            model_health_score=health,
            summary=summary,
        )

    def _check_broken_references(self) -> None:
        for node in self.graph.nodes():
            nd = self.graph.nodes[node]
            formula = nd.get("formula")
            if not formula:
                continue
            for pred in self.graph.predecessors(node):
                pred_nd = self.graph.nodes.get(pred, {})
                if pred_nd.get("data_type") == "empty" and pred_nd.get("value") is None:
                    self.issues.append(IntegrityIssue(
                        severity="critical",
                        category="broken_reference",
                        cell=node,
                        sheet=nd.get("sheet_name", ""),
                        message=f"Formula references empty cell {pred}",
                        suggestion=f"Check if {pred} should contain a value or if the reference in {node} is incorrect",
                        details={"formula": formula, "referenced_cell": pred},
                    ))

    def _check_dangling_named_ranges(self) -> None:
        for name, dest in self.workbook_data.named_ranges.items():
            if "!" in dest:
                sheet, cell = dest.split("!", 1)
                sheet = sheet.strip("'")
                full_addr = f"{sheet}!{cell}"
                if full_addr not in self.workbook_data.cells:
                    self.issues.append(IntegrityIssue(
                        severity="warning",
                        category="dangling_named_range",
                        cell=full_addr,
                        sheet=sheet,
                        message=f"Named range '{name}' points to non-existent cell {full_addr}",
                        suggestion=f"Update or remove named range '{name}'",
                        details={"named_range": name, "destination": dest},
                    ))

    def _check_hardcoded_in_formula_sheets(self) -> None:
        sheet_formula_ratio: Dict[str, Dict[str, int]] = {}
        for addr, cell in self.workbook_data.cells.items():
            s = cell.sheet_name
            if s not in sheet_formula_ratio:
                sheet_formula_ratio[s] = {"formula": 0, "hardcoded": 0, "total": 0}
            sheet_formula_ratio[s]["total"] += 1
            if cell.formula:
                sheet_formula_ratio[s]["formula"] += 1
            elif cell.is_hardcoded:
                sheet_formula_ratio[s]["hardcoded"] += 1

        assumption_sheets = {"assumptions", "inputs", "parameters", "settings", "config"}

        for sheet, counts in sheet_formula_ratio.items():
            if sheet.lower() in assumption_sheets:
                continue
            if counts["total"] < 5:
                continue
            formula_ratio = counts["formula"] / max(counts["total"], 1)
            if formula_ratio > 0.5 and counts["hardcoded"] > 3:
                for addr, cell in self.workbook_data.cells.items():
                    if cell.sheet_name != sheet or not cell.is_hardcoded:
                        continue
                    if not isinstance(cell.value, (int, float)):
                        continue
                    out_degree = self.graph.out_degree(addr) if addr in self.graph else 0
                    if out_degree > 0:
                        self.issues.append(IntegrityIssue(
                            severity="warning",
                            category="hardcoded_in_formula_sheet",
                            cell=addr,
                            sheet=sheet,
                            message=f"Hardcoded value {cell.value} in formula-heavy sheet '{sheet}'",
                            suggestion=f"Move this assumption to an 'Assumptions' sheet and reference it",
                            details={
                                "value": cell.value,
                                "dependents_count": out_degree,
                                "sheet_formula_ratio": round(formula_ratio, 2),
                            },
                        ))

    def _check_unused_inputs(self) -> None:
        for node in self.graph.nodes():
            nd = self.graph.nodes[node]
            if not nd.get("is_hardcoded"):
                continue
            if not isinstance(nd.get("value"), (int, float)):
                continue
            out_degree = self.graph.out_degree(node)
            if out_degree == 0:
                self.issues.append(IntegrityIssue(
                    severity="info",
                    category="unused_input",
                    cell=node,
                    sheet=nd.get("sheet_name", ""),
                    message=f"Hardcoded value {nd.get('value')} has no dependents",
                    suggestion="Remove if unused, or connect to relevant formulas",
                    details={"value": nd.get("value"), "named_range": nd.get("named_range")},
                ))

    def _check_circular_references(self) -> None:
        try:
            cycles = list(nx.simple_cycles(self.graph))
            for cycle in cycles[:10]:
                cycle_str = " → ".join(cycle[:5])
                if len(cycle) > 5:
                    cycle_str += f" → ... ({len(cycle)} cells total)"
                for cell in cycle:
                    nd = self.graph.nodes.get(cell, {})
                    self.issues.append(IntegrityIssue(
                        severity="critical",
                        category="circular_reference",
                        cell=cell,
                        sheet=nd.get("sheet_name", ""),
                        message=f"Part of circular reference: {cycle_str}",
                        suggestion="Break the circular dependency by replacing one reference with a hardcoded value or intermediate calculation",
                        details={"cycle_length": len(cycle), "cycle_cells": cycle[:10]},
                    ))
                    break
        except Exception as e:
            logger.warning(f"Circular reference check failed: {e}")

    def _check_inconsistent_time_series(self) -> None:
        sheet_rows: Dict[str, Set[int]] = {}
        sheet_cols: Dict[str, Set[int]] = {}
        for addr, cell in self.workbook_data.cells.items():
            s = cell.sheet_name
            if s not in sheet_rows:
                sheet_rows[s] = set()
                sheet_cols[s] = set()
            sheet_rows[s].add(cell.row)
            sheet_cols[s].add(cell.col)

        for sheet in self.workbook_data.sheets:
            if sheet not in sheet_cols:
                continue
            cols = sorted(sheet_cols[sheet])
            if len(cols) < 4:
                continue

            row_lengths: Dict[int, int] = {}
            for addr, cell in self.workbook_data.cells.items():
                if cell.sheet_name != sheet:
                    continue
                if cell.formula is not None:
                    row_lengths[cell.row] = row_lengths.get(cell.row, 0) + 1

            if not row_lengths:
                continue
            lengths = list(row_lengths.values())
            most_common = max(set(lengths), key=lengths.count)

            for row, length in row_lengths.items():
                if length < most_common and most_common - length > 2:
                    self.issues.append(IntegrityIssue(
                        severity="warning",
                        category="inconsistent_time_series",
                        cell=f"{sheet}!A{row}",
                        sheet=sheet,
                        message=f"Row {row} has {length} formula cells vs expected {most_common}",
                        suggestion=f"Check if row {row} is missing formulas in some periods",
                        details={"row": row, "actual_length": length, "expected_length": most_common},
                    ))

    def _check_sign_conventions(self) -> None:
        cluster_signs: Dict[int, Dict[str, int]] = {}
        for node in self.graph.nodes():
            nd = self.graph.nodes[node]
            val = nd.get("value")
            if not isinstance(val, (int, float)) or val == 0:
                continue
            cid = nd.get("cluster_id", 0)
            if cid not in cluster_signs:
                cluster_signs[cid] = {"positive": 0, "negative": 0}
            if val > 0:
                cluster_signs[cid]["positive"] += 1
            else:
                cluster_signs[cid]["negative"] += 1

        for cid, signs in cluster_signs.items():
            total = signs["positive"] + signs["negative"]
            if total < 5:
                continue
            minority = min(signs["positive"], signs["negative"])
            if 1 <= minority <= 2 and total > 10:
                minority_sign = "negative" if signs["negative"] <= 2 else "positive"
                for node in self.graph.nodes():
                    nd = self.graph.nodes[node]
                    if nd.get("cluster_id") != cid:
                        continue
                    val = nd.get("value")
                    if not isinstance(val, (int, float)):
                        continue
                    if (minority_sign == "negative" and val < 0) or (minority_sign == "positive" and val > 0):
                        self.issues.append(IntegrityIssue(
                            severity="info",
                            category="sign_convention",
                            cell=node,
                            sheet=nd.get("sheet_name", ""),
                            message=f"Value {val} has unusual sign for this group (cluster {nd.get('cluster_name', cid)})",
                            suggestion="Verify sign convention is intentional",
                            details={"value": val, "cluster": nd.get("cluster_name"), "cluster_id": cid},
                        ))

    def _check_formula_complexity(self) -> None:
        for addr, cell in self.workbook_data.cells.items():
            if not cell.formula:
                continue
            formula = cell.formula

            nesting = formula.count("(")
            func_count = len(re.findall(r'[A-Z]+\(', formula))
            length = len(formula)

            complexity_score = nesting * 2 + func_count * 1.5 + (length / 50)

            if complexity_score > 15:
                severity = "critical" if complexity_score > 30 else "warning"
                self.issues.append(IntegrityIssue(
                    severity=severity,
                    category="formula_complexity",
                    cell=addr,
                    sheet=cell.sheet_name,
                    message=f"Complex formula (score={complexity_score:.0f}): {nesting} nesting levels, {func_count} functions, {length} chars",
                    suggestion="Consider breaking into helper cells for readability and auditability",
                    details={
                        "formula": formula[:200],
                        "complexity_score": round(complexity_score, 1),
                        "nesting_depth": nesting,
                        "function_count": func_count,
                        "formula_length": length,
                    },
                ))
