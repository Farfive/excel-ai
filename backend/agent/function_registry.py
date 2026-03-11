TOOLS_SCHEMA = [
    {
        "name": "read_range",
        "description": "Read cell values from the workbook for a given sheet and range.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet": {"type": "string", "description": "Sheet name"},
                "range": {"type": "string", "description": "Cell range e.g. A1:B10 or A1"},
            },
            "required": ["sheet", "range"],
        },
    },
    {
        "name": "write_range",
        "description": "Write values to cells in the workbook. Returns ordered write instructions respecting formula dependencies.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet": {"type": "string", "description": "Sheet name"},
                "range": {"type": "string", "description": "Cell range e.g. A1:B10 or A1"},
                "values": {
                    "type": "object",
                    "description": "Dict mapping cell address (SheetName!A1) to value",
                    "additionalProperties": True,
                },
            },
            "required": ["sheet", "range", "values"],
        },
    },
    {
        "name": "get_dependencies",
        "description": "Get upstream and downstream dependencies for a cell up to 2 hops.",
        "parameters": {
            "type": "object",
            "properties": {
                "cell": {"type": "string", "description": "Full cell address e.g. Sheet1!A1"},
            },
            "required": ["cell"],
        },
    },
    {
        "name": "find_anomalies",
        "description": "Find anomalous cells in the workbook using Isolation Forest scores.",
        "parameters": {
            "type": "object",
            "properties": {
                "sheet": {"type": "string", "description": "Optional sheet name to filter by"},
            },
            "required": [],
        },
    },
    {
        "name": "explain_formula",
        "description": "Explain a cell formula in plain English including its dependencies.",
        "parameters": {
            "type": "object",
            "properties": {
                "cell": {"type": "string", "description": "Full cell address e.g. Sheet1!A1"},
            },
            "required": ["cell"],
        },
    },
    {
        "name": "generate_change_log",
        "description": "Generate a markdown change log table from a list of changes made.",
        "parameters": {
            "type": "object",
            "properties": {
                "changes": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "cell": {"type": "string"},
                            "old_value": {},
                            "new_value": {},
                            "reason": {"type": "string"},
                        },
                    },
                    "description": "List of changes with cell, old_value, new_value, reason",
                },
            },
            "required": ["changes"],
        },
    },
    {
        "name": "run_sensitivity",
        "description": "Run sensitivity analysis — how changes in key inputs affect outputs. Returns tornado chart data and top drivers.",
        "parameters": {
            "type": "object",
            "properties": {
                "max_inputs": {"type": "integer", "description": "Max input cells to test (default 15)"},
                "max_outputs": {"type": "integer", "description": "Max output cells to monitor (default 8)"},
            },
            "required": [],
        },
    },
    {
        "name": "run_integrity_check",
        "description": "Check model integrity — broken refs, dangling named ranges, formula complexity, circular refs, unused inputs, sign conventions.",
        "parameters": {
            "type": "object",
            "properties": {},
            "required": [],
        },
    },
    {
        "name": "run_smart_suggestions",
        "description": "Get smart suggestions for model improvement — missing discounting, inconsistent growth rates, hardcoded patterns, duplicate formulas.",
        "parameters": {
            "type": "object",
            "properties": {},
            "required": [],
        },
    },
    {
        "name": "create_scenario",
        "description": "Create a scenario with custom input overrides and see estimated impact on outputs.",
        "parameters": {
            "type": "object",
            "properties": {
                "name": {"type": "string", "description": "Scenario name e.g. Upside, Downside"},
                "description": {"type": "string", "description": "Scenario description"},
                "perturbation_pct": {"type": "number", "description": "Perturbation % to apply to all inputs (e.g. 10 for +10%, -5 for -5%)"},
            },
            "required": ["name"],
        },
    },
    {
        "name": "compare_scenarios",
        "description": "Compare all created scenarios against Base Case — shows deltas and delta percentages for each output.",
        "parameters": {
            "type": "object",
            "properties": {},
            "required": [],
        },
    },
]
