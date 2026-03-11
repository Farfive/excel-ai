import logging
import re
import uuid
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Any

import networkx as nx

from parser.xlsx_parser import WorkbookData

logger = logging.getLogger(__name__)

FORMULA_PATTERNS = [
    (re.compile(r"=NPV\(([^,]+),([^)]+)\)", re.IGNORECASE), "NPV using {0} as discount rate applied to cash flows {1}"),
    (re.compile(r"=IRR\(([^)]+)\)", re.IGNORECASE), "IRR of cash flows {0}"),
    (re.compile(r"=SUM\(([^)]+)\)\s*\*\s*([^,)\s]+)", re.IGNORECASE), "Sum of {0} multiplied by {1}"),
    (re.compile(r"=SUM\(([^)]+)\)", re.IGNORECASE), "Sum of {0}"),
    (re.compile(r"=SUMIF\(([^,]+),([^,]+),([^)]+)\)", re.IGNORECASE), "Sum of {2} where {0} matches {1}"),
    (re.compile(r"=IF\(([^,]+),([^,]+),([^)]+)\)", re.IGNORECASE), "If {0} then {1} else {2}"),
    (re.compile(r"=VLOOKUP\(([^,]+),([^,]+),([^,]+)", re.IGNORECASE), "Lookup {0} in {1} returning column {2}"),
    (re.compile(r"=AVERAGE\(([^)]+)\)", re.IGNORECASE), "Average of {0}"),
    (re.compile(r"=COUNT\(([^)]+)\)", re.IGNORECASE), "Count of {0}"),
    (re.compile(r"=MIN\(([^)]+)\)", re.IGNORECASE), "Minimum of {0}"),
    (re.compile(r"=MAX\(([^)]+)\)", re.IGNORECASE), "Maximum of {0}"),
    (re.compile(r"=PRODUCT\(([^)]+)\)", re.IGNORECASE), "Product of {0}"),
]


def _explain_formula(formula: str) -> str:
    if not formula:
        return ""
    for pattern, template in FORMULA_PATTERNS:
        m = pattern.match(formula.strip())
        if m:
            try:
                return template.format(*m.groups())
            except Exception:
                pass
    return formula


@dataclass
class Chunk:
    chunk_id: str
    cluster_id: int
    cluster_name: str
    text: str
    metadata: Dict[str, Any] = field(default_factory=dict)
    embedding: Optional[List[float]] = None


class ChunkBuilder:
    def build_chunks(
        self,
        workbook_data: WorkbookData,
        graph: nx.DiGraph,
        workbook_uuid: str,
    ) -> List[Chunk]:
        workbook_context = self._build_workbook_context(workbook_data, graph)[:200]

        clusters: Dict[int, List[str]] = {}
        for node in graph.nodes():
            cid = graph.nodes[node].get("cluster_id", 0)
            clusters.setdefault(cid, []).append(node)

        chunks: List[Chunk] = []

        for cluster_id, nodes in clusters.items():
            if not nodes:
                continue

            cluster_name = graph.nodes[nodes[0]].get("cluster_name", f"cluster_{cluster_id}")

            sheets_in_cluster = list({graph.nodes[n].get("sheet_name", "") for n in nodes if graph.nodes[n].get("sheet_name")})
            sheet_label = sheets_in_cluster[0] if sheets_in_cluster else "Unknown"

            addrs = [n for n in nodes if "!" in n]
            if addrs:
                col_nums = []
                row_nums = []
                for a in addrs:
                    try:
                        _, cell_part = a.split("!", 1)
                        col_str = re.sub(r"[^A-Za-z]", "", cell_part)
                        row_str = re.sub(r"[^0-9]", "", cell_part)
                        if col_str and row_str:
                            col_num = 0
                            for ch in col_str.upper():
                                col_num = col_num * 26 + (ord(ch) - ord("A") + 1)
                            col_nums.append(col_num)
                            row_nums.append(int(row_str))
                    except Exception:
                        pass
                if col_nums and row_nums:
                    def _num_to_col(n: int) -> str:
                        result = ""
                        while n > 0:
                            n, rem = divmod(n - 1, 26)
                            result = chr(ord("A") + rem) + result
                        return result
                    min_addr = f"{_num_to_col(min(col_nums))}{min(row_nums)}"
                    max_addr = f"{_num_to_col(max(col_nums))}{max(row_nums)}"
                    range_label = f"{min_addr}:{max_addr}"
                else:
                    range_label = "various"
            else:
                range_label = "various"

            top_nodes = sorted(
                nodes,
                key=lambda n: graph.nodes[n].get("pagerank", 0),
                reverse=True,
            )[:5]

            key_cells_lines = []
            cell_addresses = []
            for node in top_nodes:
                nd = graph.nodes[node]
                addr_short = node.split("!", 1)[1] if "!" in node else node
                nr = nd.get("named_range") or ""
                nr_part = f" ({nr})" if nr else ""
                val = nd.get("value")
                formula = nd.get("formula") or ""
                formula_exp = _explain_formula(formula) if formula else ""
                formula_part = f" [{formula_exp}]" if formula_exp else (f" [{formula}]" if formula else "")
                key_cells_lines.append(f"- {addr_short}{nr_part}: {val}{formula_part}")
                cell_addresses.append(node)

            upstream_clusters = set()
            downstream_clusters = set()
            for node in nodes:
                for pred in graph.predecessors(node):
                    pred_cluster = graph.nodes[pred].get("cluster_name", "")
                    if pred_cluster and pred_cluster != cluster_name:
                        upstream_clusters.add(pred_cluster)
                for succ in graph.successors(node):
                    succ_cluster = graph.nodes[succ].get("cluster_name", "")
                    if succ_cluster and succ_cluster != cluster_name:
                        downstream_clusters.add(succ_cluster)

            anomaly_lines = []
            for node in nodes:
                nd = graph.nodes[node]
                if nd.get("is_anomaly"):
                    addr_short = node.split("!", 1)[1] if "!" in node else node
                    anomaly_lines.append(
                        f"! {addr_short}: {nd.get('value')} is outlier (score={nd.get('anomaly_score', 0):.3f})"
                    )

            parts = [
                f"CONTEXT: {workbook_context}",
                f"CLUSTER: {cluster_name} | Sheet: {sheet_label} | Range: {range_label}",
                "KEY CELLS:",
                *key_cells_lines,
                "DEPENDENCIES:",
                f"Receives from: {', '.join(upstream_clusters) or 'none'}",
                f"Feeds into: {', '.join(downstream_clusters) or 'none'}",
            ]
            if anomaly_lines:
                parts.append("ANOMALIES:")
                parts.extend(anomaly_lines)

            text = "\n".join(parts)

            avg_pagerank = (
                sum(graph.nodes[n].get("pagerank", 0) for n in nodes) / len(nodes)
                if nodes else 0.0
            )

            chunk = Chunk(
                chunk_id=str(uuid.uuid4()),
                cluster_id=cluster_id,
                cluster_name=cluster_name,
                text=text,
                metadata={
                    "workbook_uuid": workbook_uuid,
                    "cluster_id": cluster_id,
                    "cluster_name": cluster_name,
                    "sheet": sheet_label,
                    "cell_range": range_label,
                    "cell_addresses": cell_addresses,
                    "avg_pagerank": avg_pagerank,
                },
            )
            chunks.append(chunk)

        chunks.sort(key=lambda c: c.metadata.get("avg_pagerank", 0), reverse=True)
        logger.info(f"Built {len(chunks)} chunks for workbook {workbook_uuid}")
        return chunks

    def _build_workbook_context(self, workbook_data: WorkbookData, graph: nx.DiGraph) -> str:
        sheet_names = workbook_data.sheets
        cell_count = workbook_data.metadata.get("cell_count", 0)
        named_ranges = list(workbook_data.named_ranges.keys())[:5]
        nr_str = ", ".join(named_ranges) if named_ranges else "none"
        return (
            f"Workbook with {cell_count} cells across sheets: {', '.join(sheet_names)}. "
            f"Key named ranges: {nr_str}. "
            f"Graph: {graph.number_of_nodes()} nodes, {graph.number_of_edges()} edges."
        )
