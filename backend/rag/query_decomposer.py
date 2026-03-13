"""
Query Decomposition — breaks complex queries into focused sub-queries.

For simple queries (direct lookups, single-concept), passes through unchanged.
For complex queries (multi-hop, what-if, comparison), generates 2-4 sub-queries
that are each retrieved independently and results merged via RRF.
"""

import json
import logging
import re
from typing import List, Optional

logger = logging.getLogger(__name__)

# Patterns that indicate a simple, single-concept query
_SIMPLE_PATTERNS = [
    re.compile(r"^(what is|co to jest|jaki jest|jaka jest|ile wynosi|podaj|read)\s+\S+", re.I),
    re.compile(r"^(show|pokaż|wyświetl)\s+(the\s+)?\S+\s*(sheet|arkusz)?$", re.I),
]

# Patterns that indicate complex multi-hop queries
_COMPLEX_INDICATORS = [
    "how does", "what if", "co jeśli", "jak wpływa",
    "affect", "impact", "wpływa", "zależy",
    "compare", "porównaj", "difference", "różnica",
    "explain how", "wyjaśnij jak", "trace", "śledź",
    "why", "dlaczego", "relationship", "relacja",
    "sensitivity", "scenario", "what happens",
]


def is_complex_query(query: str) -> bool:
    """Heuristic check if query needs decomposition."""
    q_lower = query.lower()
    for pattern in _SIMPLE_PATTERNS:
        if pattern.match(query):
            return False
    for indicator in _COMPLEX_INDICATORS:
        if indicator in q_lower:
            return True
    # Multi-clause queries (contains "and", "then", commas with verbs)
    if " and " in q_lower and len(query.split()) > 8:
        return True
    return False


async def decompose_query(query: str, ollama_client, max_sub: int = 4) -> List[str]:
    """Decompose a complex query into focused sub-queries using LLM.

    Returns a list of sub-queries. For simple queries, returns [query].
    """
    if not is_complex_query(query):
        return [query]

    system = (
        "You are a query decomposition expert for financial Excel workbooks. "
        "Break the user's complex question into 2-4 focused sub-queries that together "
        "cover all aspects of the original question. Each sub-query should target "
        "a specific sheet, metric, or relationship in the workbook. "
        "Return ONLY a JSON array of strings, nothing else. "
        "Example: [\"What is WACC?\", \"What cells depend on WACC?\", \"How is Enterprise Value calculated?\"]"
    )

    try:
        response = await ollama_client.chat(
            messages=[{"role": "user", "content": f"Decompose this question: {query}"}],
            system=system,
            temperature=0.2,
        )
        # Parse JSON array from response
        cleaned = response.strip()
        # Handle markdown code blocks
        cleaned = re.sub(r"```(?:json)?", "", cleaned).strip()
        sub_queries = json.loads(cleaned)

        if isinstance(sub_queries, list) and 1 < len(sub_queries) <= max_sub:
            sub_queries = [str(sq).strip() for sq in sub_queries if str(sq).strip()]
            if sub_queries:
                logger.info(
                    f"Query decomposed into {len(sub_queries)} sub-queries: "
                    f"{[sq[:50] for sq in sub_queries]}"
                )
                return sub_queries
    except Exception as e:
        logger.warning(f"Query decomposition failed: {e}")

    return [query]
