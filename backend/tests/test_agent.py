import pytest
import pytest_asyncio
from unittest.mock import AsyncMock, MagicMock
from agent.excel_agent import ExcelAgent
from agent.tools import ExcelTools, ToolResult


class MockOllama:
    async def chat(self, messages, system=None, temperature=0.1):
        return '[{"step": 1, "tool": "read_range", "args": {"sheet": "Sheet1", "range": "A1"}, "reason": "Read value"}]'

    async def stream_chat(self, messages, system=None):
        yield "The value in A1 is 0.085"


class MockOllamaBadJson:
    async def chat(self, messages, system=None, temperature=0.1):
        return "This is not valid JSON at all!!!"

    async def stream_chat(self, messages, system=None):
        yield "fallback answer"


class MockOllamaReflectBlock:
    call_count = 0

    async def chat(self, messages, system=None, temperature=0.1):
        self.call_count += 1
        if self.call_count == 1:
            return '[{"step": 1, "tool": "write_range", "args": {"sheet": "Sheet1", "range": "A1", "values": {"Sheet1!A1": 999}}, "reason": "Write value"}]'
        return '{"ok": false, "concern": "IRR jumped 50pp after this change"}'

    async def stream_chat(self, messages, system=None):
        yield "blocked"


class MockOllamaReflectOk:
    call_count = 0

    async def chat(self, messages, system=None, temperature=0.1):
        self.call_count += 1
        if self.call_count == 1:
            return '[{"step": 1, "tool": "write_range", "args": {"sheet": "Sheet1", "range": "A1", "values": {"Sheet1!A1": 0.09}}, "reason": "Adjust discount rate"}]'
        return '{"ok": true, "concern": ""}'

    async def stream_chat(self, messages, system=None):
        yield "Change applied"


class MockRetriever:
    async def retrieve(self, query, workbook_uuid, k=5):
        return []

    def build_context(self, chunks, query, max_chars=3000):
        return "mock context"


def make_tools():
    import networkx as nx
    from parser.xlsx_parser import WorkbookData
    G = nx.DiGraph()
    G.add_node("Sheet1!A1", value=0.085, formula=None, data_type="number",
               named_range="discount_rate", is_hardcoded=True, sheet_name="Sheet1",
               row=1, col=1, is_merged=False, pagerank=0.5, cluster_id=0,
               cluster_name="Assumptions", is_anomaly=False, anomaly_score=0.0)
    G.graph["topological_order"] = ["Sheet1!A1"]
    wb = WorkbookData()
    return ExcelTools(workbook_state={}, graph=G, workbook_data=wb)


@pytest.mark.asyncio
async def test_plan_returns_valid_json():
    ollama = MockOllama()
    retriever = MockRetriever()
    tools = make_tools()
    agent = ExcelAgent(ollama=ollama, retriever=retriever, tools=tools)
    plan = await agent.plan("What is A1?", "context")
    assert isinstance(plan, list)
    assert len(plan) > 0
    assert "tool" in plan[0]


@pytest.mark.asyncio
async def test_plan_fallback_on_bad_json():
    ollama = MockOllamaBadJson()
    retriever = MockRetriever()
    tools = make_tools()
    agent = ExcelAgent(ollama=ollama, retriever=retriever, tools=tools)
    plan = await agent.plan("What is A1?", "context")
    assert isinstance(plan, list)
    assert len(plan) > 0


@pytest.mark.asyncio
async def test_reflect_blocks_on_ok_false():
    ollama = MockOllamaReflectBlock()
    retriever = MockRetriever()
    tools = make_tools()
    agent = ExcelAgent(ollama=ollama, retriever=retriever, tools=tools)

    step = {"step": 1, "tool": "write_range", "args": {"sheet": "Sheet1", "range": "A1", "values": {"Sheet1!A1": 999}}, "reason": "Write"}
    result = ToolResult(tool_name="write_range", success=True, data={"ordered_writes": []})
    ok, concern = await agent.reflect(step, result)
    assert ok is False
    assert concern


@pytest.mark.asyncio
async def test_reflect_continues_on_ok_true():
    ollama = MockOllamaReflectOk()
    retriever = MockRetriever()
    tools = make_tools()
    agent = ExcelAgent(ollama=ollama, retriever=retriever, tools=tools)

    step = {"step": 1, "tool": "write_range", "args": {"sheet": "Sheet1", "range": "A1", "values": {"Sheet1!A1": 0.09}}, "reason": "Adjust rate"}
    result = ToolResult(tool_name="write_range", success=True, data={"ordered_writes": []})
    ok, concern = await agent.reflect(step, result)
    assert ok is True
