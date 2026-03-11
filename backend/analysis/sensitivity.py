"""
Sensitivity Analysis Engine
Automatyczna analiza wrażliwości — jak zmiana inputów wpływa na outputy.
Perturbuje kluczowe hardcoded cells (±1%, ±5%, ±10%) i mierzy wpływ na top PageRank cells.
"""
import logging
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import networkx as nx

from parser.xlsx_parser import WorkbookData

logger = logging.getLogger(__name__)


@dataclass
class SensitivityResult:
    input_cell: str
    input_value: float
    input_named_range: Optional[str]
    output_cell: str
    output_value: float
    output_named_range: Optional[str]
    perturbation_pct: float
    estimated_impact: float
    impact_direction: str
    elasticity: float


@dataclass
class SensitivityReport:
    input_cells_tested: int
    output_cells_monitored: int
    results: List[SensitivityResult]
    tornado_chart_data: List[Dict[str, Any]]
    top_drivers: List[Dict[str, Any]]


class SensitivityAnalyzer:
    def __init__(self, graph: nx.DiGraph, workbook_data: WorkbookData) -> None:
        self.graph = graph
        self.workbook_data = workbook_data

    def find_input_cells(self, max_inputs: int = 20) -> List[str]:
        candidates = []
        for node in self.graph.nodes():
            nd = self.graph.nodes[node]
            if not nd.get("is_hardcoded"):
                continue
            val = nd.get("value")
            if not isinstance(val, (int, float)):
                continue
            if val == 0:
                continue
            out_degree = self.graph.out_degree(node)
            pr = nd.get("pagerank", 0.0)
            score = out_degree * 0.6 + pr * 100
            candidates.append((node, score, val))
        candidates.sort(key=lambda x: x[1], reverse=True)
        return [c[0] for c in candidates[:max_inputs]]

    def find_output_cells(self, max_outputs: int = 10) -> List[str]:
        candidates = []
        for node in self.graph.nodes():
            nd = self.graph.nodes[node]
            if nd.get("formula") is None:
                continue
            val = nd.get("value")
            if not isinstance(val, (int, float)):
                continue
            pr = nd.get("pagerank", 0.0)
            in_degree = self.graph.in_degree(node)
            out_degree = self.graph.out_degree(node)
            score = pr * 100 + in_degree * 0.3 - out_degree * 0.1
            candidates.append((node, score, val))
        candidates.sort(key=lambda x: x[1], reverse=True)
        return [c[0] for c in candidates[:max_outputs]]

    def _is_reachable(self, input_cell: str, output_cell: str) -> bool:
        try:
            return nx.has_path(self.graph, input_cell, output_cell)
        except (nx.NetworkXError, nx.NodeNotFound):
            return False

    def _estimate_impact(
        self, input_cell: str, output_cell: str, perturbation_pct: float
    ) -> float:
        try:
            path_length = nx.shortest_path_length(self.graph, input_cell, output_cell)
        except (nx.NetworkXNoPath, nx.NodeNotFound):
            return 0.0

        input_val = self.graph.nodes[input_cell].get("value", 0)
        output_val = self.graph.nodes[output_cell].get("value", 0)

        if not isinstance(input_val, (int, float)) or not isinstance(output_val, (int, float)):
            return 0.0
        if input_val == 0 or output_val == 0:
            return 0.0

        decay = 0.85 ** path_length
        input_pr = self.graph.nodes[input_cell].get("pagerank", 0.0)
        amplification = 1.0 + input_pr * 2.0

        delta_input = abs(input_val * perturbation_pct / 100.0)
        estimated_delta_output = delta_input * decay * amplification

        impact_pct = (estimated_delta_output / abs(output_val)) * 100.0
        return round(impact_pct, 4)

    def run(
        self,
        perturbation_pcts: Optional[List[float]] = None,
        max_inputs: int = 15,
        max_outputs: int = 8,
    ) -> SensitivityReport:
        if perturbation_pcts is None:
            perturbation_pcts = [-10.0, -5.0, -1.0, 1.0, 5.0, 10.0]

        input_cells = self.find_input_cells(max_inputs)
        output_cells = self.find_output_cells(max_outputs)

        results: List[SensitivityResult] = []

        for inp in input_cells:
            inp_nd = self.graph.nodes[inp]
            inp_val = inp_nd.get("value", 0)
            inp_nr = inp_nd.get("named_range")

            for out in output_cells:
                if not self._is_reachable(inp, out):
                    continue

                out_nd = self.graph.nodes[out]
                out_val = out_nd.get("value", 0)
                out_nr = out_nd.get("named_range")

                for pct in perturbation_pcts:
                    impact = self._estimate_impact(inp, out, pct)
                    if impact == 0.0:
                        continue

                    direction = "positive" if pct > 0 else "negative"
                    elasticity = abs(impact / pct) if pct != 0 else 0.0

                    results.append(SensitivityResult(
                        input_cell=inp,
                        input_value=float(inp_val),
                        input_named_range=inp_nr,
                        output_cell=out,
                        output_value=float(out_val),
                        output_named_range=out_nr,
                        perturbation_pct=pct,
                        estimated_impact=impact,
                        impact_direction=direction,
                        elasticity=round(elasticity, 4),
                    ))

        tornado_data = self._build_tornado_data(results, output_cells)
        top_drivers = self._find_top_drivers(results)

        logger.info(
            f"Sensitivity analysis: {len(input_cells)} inputs × {len(output_cells)} outputs = {len(results)} data points"
        )

        return SensitivityReport(
            input_cells_tested=len(input_cells),
            output_cells_monitored=len(output_cells),
            results=results,
            tornado_chart_data=tornado_data,
            top_drivers=top_drivers,
        )

    def _build_tornado_data(
        self, results: List[SensitivityResult], output_cells: List[str]
    ) -> List[Dict[str, Any]]:
        tornado: List[Dict[str, Any]] = []
        for out in output_cells:
            out_results = [r for r in results if r.output_cell == out]
            by_input: Dict[str, Dict[str, float]] = {}
            for r in out_results:
                key = r.input_cell
                if key not in by_input:
                    by_input[key] = {"low": 0.0, "high": 0.0}
                if r.perturbation_pct < 0:
                    by_input[key]["low"] = min(by_input[key]["low"], -r.estimated_impact)
                else:
                    by_input[key]["high"] = max(by_input[key]["high"], r.estimated_impact)

            bars = []
            for inp, vals in by_input.items():
                spread = vals["high"] - vals["low"]
                bars.append({
                    "input_cell": inp,
                    "input_named_range": self.graph.nodes[inp].get("named_range"),
                    "low_impact_pct": vals["low"],
                    "high_impact_pct": vals["high"],
                    "spread": spread,
                })
            bars.sort(key=lambda x: x["spread"], reverse=True)

            tornado.append({
                "output_cell": out,
                "output_named_range": self.graph.nodes[out].get("named_range"),
                "bars": bars[:10],
            })
        return tornado

    def _find_top_drivers(self, results: List[SensitivityResult]) -> List[Dict[str, Any]]:
        driver_scores: Dict[str, float] = {}
        driver_info: Dict[str, Dict] = {}
        for r in results:
            key = r.input_cell
            driver_scores[key] = driver_scores.get(key, 0.0) + abs(r.estimated_impact)
            if key not in driver_info:
                driver_info[key] = {
                    "cell": key,
                    "value": r.input_value,
                    "named_range": r.input_named_range,
                    "affected_outputs": set(),
                    "max_elasticity": 0.0,
                }
            driver_info[key]["affected_outputs"].add(r.output_cell)
            driver_info[key]["max_elasticity"] = max(
                driver_info[key]["max_elasticity"], r.elasticity
            )

        sorted_drivers = sorted(driver_scores.items(), key=lambda x: x[1], reverse=True)

        top = []
        for cell, score in sorted_drivers[:10]:
            info = driver_info[cell]
            top.append({
                "cell": cell,
                "value": info["value"],
                "named_range": info["named_range"],
                "total_impact_score": round(score, 4),
                "affected_outputs_count": len(info["affected_outputs"]),
                "max_elasticity": info["max_elasticity"],
                "risk_level": "high" if score > 50 else ("medium" if score > 10 else "low"),
            })
        return top
