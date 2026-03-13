"""
Formula Chain Unrolling — traces formula dependencies up to N levels deep
and produces a human-readable chain that gets injected into LLM context.

Example output:
  DCF!B20 [Enterprise Value] = SUM(B8:B19) + B21
    ├─ DCF!B8 [PV FCF 2021] = B5/(1+Assumptions!B13)^1
    │   └─ Assumptions!B13 [WACC] = 0.10
    ├─ DCF!B21 [PV Terminal Value] = B22/(1+Assumptions!B13)^13
    │   └─ DCF!B22 [Terminal Value] = B19*(1+Assumptions!B14)/(Assumptions!B13-Assumptions!B14)
"""

import logging
from collections import deque
from typing import Any, Dict, List, Optional, Set, Tuple

import networkx as nx

logger = logging.getLogger(__name__)

MAX_DEPTH = 3
MAX_NODES = 50
MAX_CHAIN_CHARS = 3000


def unroll_formula_chain(
    seed_addresses: List[str],
    graph: nx.DiGraph,
    workbook_data: Any,
    max_depth: int = MAX_DEPTH,
    direction: str = "both",
) -> str:
    """Trace formula dependencies from seed cells and produce readable chain text.

    Args:
        seed_addresses: Starting cell addresses (e.g. from top retrieved chunks)
        graph: Dependency DiGraph (edge = dep → formula cell)
        workbook_data: WorkbookData with .cells
        max_depth: How many levels deep to trace
        direction: 'up' (predecessors/inputs), 'down' (successors/outputs), 'both'

    Returns:
        Human-readable formula chain string for LLM context injection.
    """
    if not seed_addresses or not graph:
        return ""

    visited: Set[str] = set()
    chain_lines: List[str] = []

    # Collect all chain nodes via BFS
    queue: deque = deque()
    for addr in seed_addresses:
        if addr in graph:
            queue.append((addr, 0, None))

    while queue and len(visited) < MAX_NODES:
        addr, depth, parent = queue.popleft()
        if addr in visited:
            continue
        visited.add(addr)

        cell = workbook_data.cells.get(addr)
        line = _format_cell(addr, cell, depth)
        chain_lines.append(line)

        if depth >= max_depth:
            continue

        # Trace upstream (inputs to this formula)
        if direction in ("up", "both"):
            for pred in graph.predecessors(addr):
                if pred not in visited:
                    queue.append((pred, depth + 1, addr))

        # Trace downstream (cells that use this cell)
        if direction in ("down", "both"):
            for succ in graph.successors(addr):
                if succ not in visited:
                    queue.append((succ, depth + 1, addr))

    if not chain_lines:
        return ""

    # Build the output
    header = f"FORMULA CHAIN ({len(chain_lines)} cells, depth={max_depth}):"
    body = "\n".join(chain_lines)

    result = f"{header}\n{body}"
    if len(result) > MAX_CHAIN_CHARS:
        result = result[:MAX_CHAIN_CHARS] + f"\n[...truncated, {len(chain_lines)} total cells]"

    logger.info(f"Formula chain: {len(chain_lines)} cells from {len(seed_addresses)} seeds")
    return result


def _format_cell(addr: str, cell, depth: int) -> str:
    """Format a single cell for the chain output."""
    indent = "  " * depth
    prefix = "├─ " if depth > 0 else ""

    if cell is None:
        return f"{indent}{prefix}{addr} [missing]"

    parts = [f"{indent}{prefix}{addr}"]

    # Add row label if available
    label = _get_row_label(addr, cell)
    if label:
        parts[0] += f" [{label}]"

    if cell.formula:
        parts[0] += f" = {cell.formula}"
    elif cell.value is not None:
        val = cell.value
        if isinstance(val, float):
            if abs(val) < 1:
                parts[0] += f" = {val:.2%}"
            elif abs(val) > 1e6:
                parts[0] += f" = ${val/1e6:.1f}M"
            else:
                parts[0] += f" = {val:,.2f}"
        else:
            parts[0] += f" = {str(val)[:80]}"

    if cell.named_range:
        parts[0] += f" (named: {cell.named_range})"

    return parts[0]


def _get_row_label(addr: str, cell) -> Optional[str]:
    """Try to infer a label for the cell from its row context."""
    if isinstance(cell.value, str) and not cell.formula:
        return cell.value[:40]
    return None


def extract_seed_addresses(chunks: List[Dict], max_seeds: int = 10) -> List[str]:
    """Extract the most important cell addresses from top retrieved chunks.

    Prioritizes cells that appear in formulas and cross-sheet references.
    """
    seen: Set[str] = set()
    seeds: List[str] = []

    for chunk in chunks[:5]:
        addrs = chunk.get("metadata", {}).get("cell_addresses", [])
        if isinstance(addrs, str):
            addrs = addrs.split(",") if addrs else []
        for addr in addrs:
            addr = addr.strip()
            if addr and addr not in seen:
                seen.add(addr)
                seeds.append(addr)
                if len(seeds) >= max_seeds:
                    return seeds

    return seeds
