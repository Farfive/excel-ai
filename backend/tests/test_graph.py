import pytest
import networkx as nx
from parser.xlsx_parser import XLSXParser
from parser.graph_builder import DependencyGraphBuilder


def test_graph_edges_correct(sample_xlsx_path):
    parser = XLSXParser()
    workbook_data = parser.parse(sample_xlsx_path)
    builder = DependencyGraphBuilder()
    G = builder.build(workbook_data)
    assert G.number_of_nodes() > 0


def test_topological_order_valid(sample_xlsx_path):
    parser = XLSXParser()
    workbook_data = parser.parse(sample_xlsx_path)
    builder = DependencyGraphBuilder()
    G = builder.build(workbook_data)
    G = builder.run_algorithms(G, workbook_data)
    topo = G.graph.get("topological_order", [])
    assert isinstance(topo, list)
    assert len(topo) > 0


def test_pagerank_values_normalized(sample_xlsx_path):
    parser = XLSXParser()
    workbook_data = parser.parse(sample_xlsx_path)
    builder = DependencyGraphBuilder()
    G = builder.build(workbook_data)
    G = builder.run_algorithms(G, workbook_data)
    for node in G.nodes():
        pr = G.nodes[node].get("pagerank", 0.0)
        assert 0.0 <= pr <= 1.0, f"PageRank out of range for {node}: {pr}"


def test_louvain_assigns_cluster_id(sample_xlsx_path):
    parser = XLSXParser()
    workbook_data = parser.parse(sample_xlsx_path)
    builder = DependencyGraphBuilder()
    G = builder.build(workbook_data)
    G = builder.run_algorithms(G, workbook_data)
    for node in G.nodes():
        assert "cluster_id" in G.nodes[node], f"cluster_id missing for {node}"


def test_isolation_forest_runs(sample_xlsx_path):
    parser = XLSXParser()
    workbook_data = parser.parse(sample_xlsx_path)
    builder = DependencyGraphBuilder()
    G = builder.build(workbook_data)
    G = builder.run_algorithms(G, workbook_data)
    for node in G.nodes():
        assert "is_anomaly" in G.nodes[node]
        assert isinstance(G.nodes[node]["is_anomaly"], bool)


def test_graph_export_valid(sample_xlsx_path, tmp_path):
    parser = XLSXParser()
    workbook_data = parser.parse(sample_xlsx_path)
    builder = DependencyGraphBuilder()
    G = builder.build(workbook_data)
    G = builder.run_algorithms(G, workbook_data)
    output_dir = str(tmp_path / "graph_output")
    builder.export(G, output_dir)
    import os, json
    assert os.path.exists(os.path.join(output_dir, "graph.json"))
    assert os.path.exists(os.path.join(output_dir, "clusters.json"))
    assert os.path.exists(os.path.join(output_dir, "anomalies.json"))
    with open(os.path.join(output_dir, "graph.json")) as f:
        data = json.load(f)
    assert "nodes" in data
