import pytest
from rag.chunker import ChunkBuilder, Chunk
from parser.xlsx_parser import XLSXParser
from parser.graph_builder import DependencyGraphBuilder


def test_chunks_built(sample_xlsx_path):
    parser = XLSXParser()
    workbook_data = parser.parse(sample_xlsx_path)
    builder = DependencyGraphBuilder()
    G = builder.build(workbook_data)
    G = builder.run_algorithms(G, workbook_data)
    chunk_builder = ChunkBuilder()
    chunks = chunk_builder.build_chunks(workbook_data, G, "test-uuid-123")
    assert len(chunks) > 0
    for chunk in chunks:
        assert isinstance(chunk, Chunk)
        assert chunk.text
        assert chunk.chunk_id


def test_chunk_text_contains_context(sample_xlsx_path):
    parser = XLSXParser()
    workbook_data = parser.parse(sample_xlsx_path)
    builder = DependencyGraphBuilder()
    G = builder.build(workbook_data)
    G = builder.run_algorithms(G, workbook_data)
    chunk_builder = ChunkBuilder()
    chunks = chunk_builder.build_chunks(workbook_data, G, "test-uuid-123")
    for chunk in chunks:
        assert "CONTEXT:" in chunk.text
        assert "CLUSTER:" in chunk.text


def test_chunks_sorted_by_pagerank(sample_xlsx_path):
    parser = XLSXParser()
    workbook_data = parser.parse(sample_xlsx_path)
    builder = DependencyGraphBuilder()
    G = builder.build(workbook_data)
    G = builder.run_algorithms(G, workbook_data)
    chunk_builder = ChunkBuilder()
    chunks = chunk_builder.build_chunks(workbook_data, G, "test-uuid-123")
    if len(chunks) > 1:
        pageranks = [c.metadata.get("avg_pagerank", 0) for c in chunks]
        assert pageranks == sorted(pageranks, reverse=True)
