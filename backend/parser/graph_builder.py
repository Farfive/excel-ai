import json
import logging
import os
import re
from typing import Dict, List, Optional, Tuple

import networkx as nx
import numpy as np
from sklearn.ensemble import IsolationForest

from parser.xlsx_parser import WorkbookData, CellData

logger = logging.getLogger(__name__)

CELL_REF_PATTERN = re.compile(
    r"(?:'([^']+)'|([A-Za-z][A-Za-z0-9 ]*))"  # sheet name optionally quoted
    r"!"
    r"(\$?[A-Z]+\$?[0-9]+(?::\$?[A-Z]+\$?[0-9]+)?)"  # cell or range
    r"|"
    r"(\$?[A-Z]+\$?[0-9]+(?::\$?[A-Z]+\$?[0-9]+)?)"  # bare cell or range
    , re.IGNORECASE
)

SINGLE_CELL_PATTERN = re.compile(r"^\$?([A-Z]+)\$?([0-9]+)$", re.IGNORECASE)
RANGE_PATTERN = re.compile(r"^\$?([A-Z]+)\$?([0-9]+):\$?([A-Z]+)\$?([0-9]+)$", re.IGNORECASE)


def _expand_range(sheet: str, range_str: str) -> List[str]:
    m = RANGE_PATTERN.match(range_str.strip())
    if not m:
        cell = range_str.replace("$", "").upper()
        return [f"{sheet}!{cell}"]
    col_start = _col_to_num(m.group(1))
    row_start = int(m.group(2))
    col_end = _col_to_num(m.group(3))
    row_end = int(m.group(4))
    result = []
    for r in range(row_start, row_end + 1):
        for c in range(col_start, col_end + 1):
            result.append(f"{sheet}!{_num_to_col(c)}{r}")
    return result


def _col_to_num(col: str) -> int:
    col = col.upper().replace("$", "")
    result = 0
    for ch in col:
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result


def _num_to_col(n: int) -> str:
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(ord("A") + rem) + result
    return result


def _parse_formula_refs(formula: str, current_sheet: str, named_ranges: Dict[str, str]) -> List[str]:
    refs = []
    if not formula:
        return refs

    for m in CELL_REF_PATTERN.finditer(formula):
        quoted_sheet = m.group(1)
        unquoted_sheet = m.group(2)
        cross_ref = m.group(3)
        bare_ref = m.group(4)

        if cross_ref:
            sheet = (quoted_sheet or unquoted_sheet or "").strip()
            refs.extend(_expand_range(sheet, cross_ref))
        elif bare_ref:
            refs.extend(_expand_range(current_sheet, bare_ref))

    for nr_name, nr_dest in named_ranges.items():
        if re.search(r'\b' + re.escape(nr_name) + r'\b', formula, re.IGNORECASE):
            if "!" in nr_dest:
                sheet_part, cell_part = nr_dest.split("!", 1)
                sheet_part = sheet_part.strip("'")
                refs.extend(_expand_range(sheet_part, cell_part))

    return list(set(refs))


class DependencyGraphBuilder:
    def build(self, workbook_data: WorkbookData) -> nx.DiGraph:
        G = nx.DiGraph()

        for addr, cell in workbook_data.cells.items():
            G.add_node(addr, **{
                "value": cell.value,
                "formula": cell.formula,
                "data_type": cell.data_type,
                "named_range": cell.named_range,
                "is_hardcoded": cell.is_hardcoded,
                "sheet_name": cell.sheet_name,
                "row": cell.row,
                "col": cell.col,
                "is_merged": cell.is_merged,
            })

        for addr, cell in workbook_data.cells.items():
            if not cell.formula:
                continue
            deps = _parse_formula_refs(cell.formula, cell.sheet_name, workbook_data.named_ranges)
            for dep in deps:
                if dep != addr:
                    if dep not in G:
                        G.add_node(dep, value=None, formula=None, data_type="empty",
                                   named_range=None, is_hardcoded=False,
                                   sheet_name=dep.split("!")[0] if "!" in dep else "",
                                   row=0, col=0, is_merged=False)
                    G.add_edge(dep, addr)

        logger.info(f"Built dependency graph: {G.number_of_nodes()} nodes, {G.number_of_edges()} edges")
        return G

    def run_algorithms(self, G: nx.DiGraph, workbook_data: WorkbookData) -> nx.DiGraph:
        try:
            topo_order = list(nx.topological_sort(G))
            G.graph["topological_order"] = topo_order
        except nx.NetworkXUnfeasible:
            logger.warning("Circular references detected — topological sort skipped")
            G.graph["topological_order"] = list(G.nodes())
            workbook_data.metadata["has_circular_refs"] = True

        try:
            pr = nx.pagerank(G, alpha=0.85)
            max_pr = max(pr.values()) if pr else 1.0
            min_pr = min(pr.values()) if pr else 0.0
            rng = max_pr - min_pr if max_pr != min_pr else 1.0
            for node, score in pr.items():
                G.nodes[node]["pagerank"] = (score - min_pr) / rng
        except Exception as e:
            logger.warning(f"PageRank failed: {e}")
            for node in G.nodes():
                G.nodes[node]["pagerank"] = 0.0

        try:
            import community as community_louvain
            undirected = G.to_undirected()
            partition = community_louvain.best_partition(undirected)
            for node, cluster_id in partition.items():
                G.nodes[node]["cluster_id"] = cluster_id
            self._name_clusters(G, workbook_data)
        except Exception as e:
            logger.warning(f"Louvain clustering failed: {e}")
            for node in G.nodes():
                G.nodes[node]["cluster_id"] = 0
                G.nodes[node]["cluster_name"] = "default"

        try:
            numeric_nodes = [
                n for n in G.nodes()
                if isinstance(G.nodes[n].get("value"), (int, float))
            ]
            if len(numeric_nodes) >= 10:
                values = np.array([[G.nodes[n]["value"]] for n in numeric_nodes], dtype=float)
                iso = IsolationForest(contamination=0.05, random_state=42)
                preds = iso.fit_predict(values)
                scores = iso.decision_function(values)
                for i, node in enumerate(numeric_nodes):
                    G.nodes[node]["is_anomaly"] = bool(preds[i] == -1)
                    G.nodes[node]["anomaly_score"] = float(scores[i])
                for n in G.nodes():
                    if "is_anomaly" not in G.nodes[n]:
                        G.nodes[n]["is_anomaly"] = False
                        G.nodes[n]["anomaly_score"] = 0.0
            else:
                for n in G.nodes():
                    G.nodes[n]["is_anomaly"] = False
                    G.nodes[n]["anomaly_score"] = 0.0
        except Exception as e:
            logger.warning(f"Isolation Forest failed: {e}")
            for n in G.nodes():
                G.nodes[n]["is_anomaly"] = False
                G.nodes[n]["anomaly_score"] = 0.0

        return G

    def _name_clusters(self, G: nx.DiGraph, workbook_data: WorkbookData) -> None:
        cluster_names: Dict[int, str] = {}
        clusters: Dict[int, List[str]] = {}
        for node in G.nodes():
            cid = G.nodes[node].get("cluster_id", 0)
            clusters.setdefault(cid, []).append(node)

        for cid, nodes in clusters.items():
            name = None
            for node in nodes:
                nr = G.nodes[node].get("named_range")
                if nr:
                    name = nr
                    break
            if not name:
                sheet_counts: Dict[str, int] = {}
                for node in nodes:
                    sheet = G.nodes[node].get("sheet_name", "")
                    sheet_counts[sheet] = sheet_counts.get(sheet, 0) + 1
                if sheet_counts:
                    name = max(sheet_counts, key=lambda k: sheet_counts[k])
                else:
                    name = f"cluster_{cid}"
            cluster_names[cid] = name

        for node in G.nodes():
            cid = G.nodes[node].get("cluster_id", 0)
            G.nodes[node]["cluster_name"] = cluster_names.get(cid, f"cluster_{cid}")

    def export(self, G: nx.DiGraph, output_dir: str) -> None:
        os.makedirs(output_dir, exist_ok=True)

        graph_data = nx.node_link_data(G)
        with open(os.path.join(output_dir, "graph.json"), "w") as f:
            json.dump(graph_data, f, default=str)

        clusters: Dict[str, List[str]] = {}
        for node in G.nodes():
            cname = G.nodes[node].get("cluster_name", "default")
            clusters.setdefault(cname, []).append(node)
        with open(os.path.join(output_dir, "clusters.json"), "w") as f:
            json.dump(clusters, f)

        anomalies = [
            {"cell": n, "value": G.nodes[n].get("value"), "anomaly_score": G.nodes[n].get("anomaly_score", 0.0)}
            for n in G.nodes()
            if G.nodes[n].get("is_anomaly")
        ]
        with open(os.path.join(output_dir, "anomalies.json"), "w") as f:
            json.dump(anomalies, f, default=str)

        logger.info(f"Graph exported to {output_dir}")
