"""
Scenario Manager
Zarządzanie scenariuszami (Base Case, Upside, Downside, Custom).
Snapshot kluczowych inputów → porównanie outputów między scenariuszami.
"""
import logging
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional
from copy import deepcopy

import networkx as nx

from parser.xlsx_parser import WorkbookData

logger = logging.getLogger(__name__)


@dataclass
class Scenario:
    name: str
    description: str
    created_at: str
    input_overrides: Dict[str, float]
    computed_outputs: Dict[str, float] = field(default_factory=dict)
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class ScenarioComparison:
    output_cell: str
    output_named_range: Optional[str]
    base_value: float
    scenarios: Dict[str, float]
    deltas: Dict[str, float]
    delta_pcts: Dict[str, float]


@dataclass
class ScenarioReport:
    scenarios: List[Scenario]
    comparisons: List[ScenarioComparison]
    summary: str


class ScenarioManager:
    def __init__(self, graph: nx.DiGraph, workbook_data: WorkbookData) -> None:
        self.graph = graph
        self.workbook_data = workbook_data
        self.scenarios: Dict[str, Scenario] = {}

    def create_base_case(self) -> Scenario:
        inputs: Dict[str, float] = {}
        for node in self.graph.nodes():
            nd = self.graph.nodes[node]
            if nd.get("is_hardcoded") and isinstance(nd.get("value"), (int, float)):
                out_degree = self.graph.out_degree(node)
                if out_degree > 0:
                    inputs[node] = float(nd["value"])

        outputs = self._get_current_outputs()
        scenario = Scenario(
            name="Base Case",
            description="Current workbook values",
            created_at=datetime.utcnow().isoformat(),
            input_overrides=inputs,
            computed_outputs=outputs,
            metadata={"auto_generated": True, "input_count": len(inputs)},
        )
        self.scenarios["Base Case"] = scenario
        return scenario

    def create_scenario(
        self,
        name: str,
        description: str,
        input_overrides: Dict[str, float],
    ) -> Scenario:
        outputs = self._estimate_outputs(input_overrides)
        scenario = Scenario(
            name=name,
            description=description,
            created_at=datetime.utcnow().isoformat(),
            input_overrides=input_overrides,
            computed_outputs=outputs,
            metadata={"input_count": len(input_overrides)},
        )
        self.scenarios[name] = scenario
        return scenario

    def create_perturbation_scenario(
        self,
        name: str,
        description: str,
        perturbation_pct: float,
    ) -> Scenario:
        base = self.scenarios.get("Base Case")
        if not base:
            base = self.create_base_case()

        overrides: Dict[str, float] = {}
        for cell, val in base.input_overrides.items():
            overrides[cell] = val * (1.0 + perturbation_pct / 100.0)

        return self.create_scenario(name, description, overrides)

    def compare(self, scenario_names: Optional[List[str]] = None) -> ScenarioReport:
        if "Base Case" not in self.scenarios:
            self.create_base_case()

        if scenario_names is None:
            scenario_names = list(self.scenarios.keys())

        base = self.scenarios["Base Case"]
        output_cells = list(base.computed_outputs.keys())

        comparisons: List[ScenarioComparison] = []
        for out_cell in output_cells:
            base_val = base.computed_outputs.get(out_cell, 0.0)
            out_nd = self.graph.nodes.get(out_cell, {})
            out_nr = out_nd.get("named_range")

            scenario_vals: Dict[str, float] = {}
            deltas: Dict[str, float] = {}
            delta_pcts: Dict[str, float] = {}

            for sname in scenario_names:
                if sname == "Base Case":
                    continue
                s = self.scenarios.get(sname)
                if not s:
                    continue
                sval = s.computed_outputs.get(out_cell, base_val)
                scenario_vals[sname] = sval
                deltas[sname] = round(sval - base_val, 4)
                if base_val != 0:
                    delta_pcts[sname] = round((sval - base_val) / abs(base_val) * 100, 2)
                else:
                    delta_pcts[sname] = 0.0

            if scenario_vals:
                comparisons.append(ScenarioComparison(
                    output_cell=out_cell,
                    output_named_range=out_nr,
                    base_value=base_val,
                    scenarios=scenario_vals,
                    deltas=deltas,
                    delta_pcts=delta_pcts,
                ))

        comparisons.sort(key=lambda c: max(abs(v) for v in c.delta_pcts.values()) if c.delta_pcts else 0, reverse=True)

        summary_parts = [f"{len(self.scenarios)} scenarios"]
        if comparisons:
            max_delta = max(max(abs(v) for v in c.delta_pcts.values()) for c in comparisons if c.delta_pcts)
            summary_parts.append(f"max delta: {max_delta:.1f}%")
        summary = "Scenario comparison: " + ", ".join(summary_parts)

        return ScenarioReport(
            scenarios=[self.scenarios[n] for n in scenario_names if n in self.scenarios],
            comparisons=comparisons,
            summary=summary,
        )

    def _get_current_outputs(self) -> Dict[str, float]:
        outputs: Dict[str, float] = {}
        candidates = []
        for node in self.graph.nodes():
            nd = self.graph.nodes[node]
            if nd.get("formula") is None:
                continue
            val = nd.get("value")
            if not isinstance(val, (int, float)):
                continue
            pr = nd.get("pagerank", 0.0)
            candidates.append((node, pr, float(val)))

        candidates.sort(key=lambda x: x[1], reverse=True)
        for cell, _, val in candidates[:20]:
            outputs[cell] = val
        return outputs

    def _estimate_outputs(self, input_overrides: Dict[str, float]) -> Dict[str, float]:
        base = self.scenarios.get("Base Case")
        if not base:
            return self._get_current_outputs()

        estimated: Dict[str, float] = dict(base.computed_outputs)

        for out_cell in estimated:
            total_delta_pct = 0.0
            for inp_cell, new_val in input_overrides.items():
                old_val = base.input_overrides.get(inp_cell, 0.0)
                if old_val == 0:
                    continue
                input_delta_pct = (new_val - old_val) / abs(old_val)

                try:
                    if nx.has_path(self.graph, inp_cell, out_cell):
                        path_len = nx.shortest_path_length(self.graph, inp_cell, out_cell)
                        decay = 0.85 ** path_len
                        inp_pr = self.graph.nodes.get(inp_cell, {}).get("pagerank", 0.0)
                        amplification = 1.0 + inp_pr
                        total_delta_pct += input_delta_pct * decay * amplification
                except (nx.NetworkXError, nx.NodeNotFound):
                    continue

            base_val = estimated[out_cell]
            estimated[out_cell] = round(base_val * (1.0 + total_delta_pct), 4)

        return estimated

    def to_dict(self) -> Dict[str, Any]:
        return {
            name: {
                "name": s.name,
                "description": s.description,
                "created_at": s.created_at,
                "input_overrides": s.input_overrides,
                "computed_outputs": s.computed_outputs,
                "metadata": s.metadata,
            }
            for name, s in self.scenarios.items()
        }
