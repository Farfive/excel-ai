from typing import Any, Dict, List, Optional
from pydantic import BaseModel


class UploadResponse(BaseModel):
    workbook_uuid: str
    cell_count: int
    cluster_count: int
    anomaly_count: int
    processing_time_ms: int


class AskRequest(BaseModel):
    question: str
    approved_plan: Optional[List[Dict[str, Any]]] = None
    mode: str = "agent"


class DeltaRequest(BaseModel):
    changed_cells: List[str]


class StateRequest(BaseModel):
    cells: Dict[str, Any]


class StateResponse(BaseModel):
    updated: int


class AnomalyItem(BaseModel):
    cell: str
    value: Any
    anomaly_score: float
    formula: Optional[str] = None
    sheet_name: Optional[str] = None
    named_range: Optional[str] = None
    severity: str = "low"


class AnomaliesResponse(BaseModel):
    anomalies: List[AnomalyItem]


class ExplainResponse(BaseModel):
    cell: str
    value: Any
    formula: Optional[str]
    named_range: Optional[str]
    cluster: Optional[str]
    pagerank: float
    depends_on: List[Dict[str, Any]]


class DeltaResponse(BaseModel):
    chunks_updated: int
    time_ms: int


class HealthResponse(BaseModel):
    status: str
    components: Dict[str, Any]
    model: str


# --- Sensitivity Analysis ---

class SensitivityRequest(BaseModel):
    perturbation_pcts: Optional[List[float]] = None
    max_inputs: int = 15
    max_outputs: int = 8


class SensitivityResultItem(BaseModel):
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


class SensitivityResponse(BaseModel):
    input_cells_tested: int
    output_cells_monitored: int
    results: List[SensitivityResultItem]
    tornado_chart_data: List[Dict[str, Any]]
    top_drivers: List[Dict[str, Any]]


# --- Integrity Check ---

class IntegrityIssueItem(BaseModel):
    severity: str
    category: str
    cell: str
    sheet: str
    message: str
    suggestion: str
    details: Optional[Dict[str, Any]] = None


class IntegrityResponse(BaseModel):
    total_issues: int
    critical: int
    warning: int
    info: int
    issues: List[IntegrityIssueItem]
    model_health_score: float
    summary: str


# --- Scenarios ---

class ScenarioCreateRequest(BaseModel):
    name: str
    description: str = ""
    input_overrides: Dict[str, float]


class ScenarioPerturbRequest(BaseModel):
    name: str
    description: str = ""
    perturbation_pct: float


class ScenarioItem(BaseModel):
    name: str
    description: str
    created_at: str
    input_overrides: Dict[str, float]
    computed_outputs: Dict[str, float]
    metadata: Dict[str, Any] = {}


class ScenarioComparisonItem(BaseModel):
    output_cell: str
    output_named_range: Optional[str]
    base_value: float
    scenarios: Dict[str, float]
    deltas: Dict[str, float]
    delta_pcts: Dict[str, float]


class ScenarioCompareResponse(BaseModel):
    scenarios: List[ScenarioItem]
    comparisons: List[ScenarioComparisonItem]
    summary: str


# --- Smart Suggestions ---

class SuggestionItem(BaseModel):
    priority: str
    category: str
    title: str
    description: str
    affected_cells: List[str]
    suggested_action: str
    estimated_effort: str
    confidence: float


class SuggestionsResponse(BaseModel):
    total_suggestions: int
    high_priority: int
    medium_priority: int
    low_priority: int
    suggestions: List[SuggestionItem]
    model_maturity_score: float


# --- Version Diff ---

class CellDiffItem(BaseModel):
    cell: str
    sheet: str
    change_type: str
    old_value: Any
    new_value: Any
    old_formula: Optional[str]
    new_formula: Optional[str]
    delta: Optional[float]
    delta_pct: Optional[float]
    impact_score: float


class SheetDiffItem(BaseModel):
    sheet_name: str
    cells_added: int
    cells_removed: int
    cells_modified: int
    formula_changes: int
    value_changes: int
    max_delta_pct: Optional[float]


class VersionDiffResponse(BaseModel):
    file_a: str
    file_b: str
    total_changes: int
    sheets_added: List[str]
    sheets_removed: List[str]
    sheet_diffs: List[SheetDiffItem]
    cell_diffs: List[CellDiffItem]
    high_impact_changes: List[CellDiffItem]
    summary: str


# --- Audit Trail ---

class AuditChangeItem(BaseModel):
    change_id: str
    timestamp: str
    cell: str
    sheet: str
    old_value: Any
    new_value: Any
    reason: str
    approved_by: str


class AuditSnapshotItem(BaseModel):
    snapshot_id: str
    timestamp: str
    label: str


class AuditResponse(BaseModel):
    total_changes: int
    changes: List[AuditChangeItem]
    snapshots: List[AuditSnapshotItem]
    cells_modified: int
    risk_changes: List[AuditChangeItem]
