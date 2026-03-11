import logging
import tempfile
import os
from typing import Any, Dict

from fastapi import APIRouter, Depends, HTTPException, UploadFile, File

from api.dependencies import (
    get_workbook_graphs, get_workbook_data_cache,
    get_audit_trails, get_scenario_managers,
)
from api.models import (
    SensitivityRequest, SensitivityResponse, SensitivityResultItem,
    IntegrityResponse, IntegrityIssueItem,
    ScenarioCreateRequest, ScenarioPerturbRequest,
    ScenarioCompareResponse, ScenarioItem, ScenarioComparisonItem,
    SuggestionsResponse, SuggestionItem,
    VersionDiffResponse, CellDiffItem, SheetDiffItem,
    AuditResponse, AuditChangeItem, AuditSnapshotItem,
)
from analysis.sensitivity import SensitivityAnalyzer
from analysis.integrity import IntegrityChecker
from analysis.scenarios import ScenarioManager
from analysis.smart_suggestions import SmartSuggestionsEngine
from analysis.version_diff import VersionDiffer
from analysis.audit_trail import AuditTrail

router = APIRouter(prefix="/workbook", tags=["analysis"])
logger = logging.getLogger(__name__)


def _get_graph_and_data(uuid: str, graphs, data_cache):
    if uuid not in graphs:
        raise HTTPException(status_code=404, detail=f"Workbook {uuid} not found. Upload first.")
    return graphs[uuid], data_cache[uuid]


@router.post("/{uuid}/sensitivity", response_model=SensitivityResponse)
async def run_sensitivity(
    uuid: str,
    request: SensitivityRequest,
    graphs=Depends(get_workbook_graphs),
    data_cache=Depends(get_workbook_data_cache),
):
    graph, workbook_data = _get_graph_and_data(uuid, graphs, data_cache)
    analyzer = SensitivityAnalyzer(graph, workbook_data)
    report = analyzer.run(
        perturbation_pcts=request.perturbation_pcts,
        max_inputs=request.max_inputs,
        max_outputs=request.max_outputs,
    )
    return SensitivityResponse(
        input_cells_tested=report.input_cells_tested,
        output_cells_monitored=report.output_cells_monitored,
        results=[SensitivityResultItem(
            input_cell=r.input_cell,
            input_value=r.input_value,
            input_named_range=r.input_named_range,
            output_cell=r.output_cell,
            output_value=r.output_value,
            output_named_range=r.output_named_range,
            perturbation_pct=r.perturbation_pct,
            estimated_impact=r.estimated_impact,
            impact_direction=r.impact_direction,
            elasticity=r.elasticity,
        ) for r in report.results],
        tornado_chart_data=report.tornado_chart_data,
        top_drivers=report.top_drivers,
    )


@router.get("/{uuid}/integrity", response_model=IntegrityResponse)
async def run_integrity_check(
    uuid: str,
    graphs=Depends(get_workbook_graphs),
    data_cache=Depends(get_workbook_data_cache),
):
    graph, workbook_data = _get_graph_and_data(uuid, graphs, data_cache)
    checker = IntegrityChecker(graph, workbook_data)
    report = checker.run()
    return IntegrityResponse(
        total_issues=report.total_issues,
        critical=report.critical,
        warning=report.warning,
        info=report.info,
        issues=[IntegrityIssueItem(
            severity=i.severity,
            category=i.category,
            cell=i.cell,
            sheet=i.sheet,
            message=i.message,
            suggestion=i.suggestion,
            details=i.details,
        ) for i in report.issues],
        model_health_score=report.model_health_score,
        summary=report.summary,
    )


@router.get("/{uuid}/suggestions", response_model=SuggestionsResponse)
async def run_smart_suggestions(
    uuid: str,
    graphs=Depends(get_workbook_graphs),
    data_cache=Depends(get_workbook_data_cache),
):
    graph, workbook_data = _get_graph_and_data(uuid, graphs, data_cache)
    engine = SmartSuggestionsEngine(graph, workbook_data)
    report = engine.run()
    return SuggestionsResponse(
        total_suggestions=report.total_suggestions,
        high_priority=report.high_priority,
        medium_priority=report.medium_priority,
        low_priority=report.low_priority,
        suggestions=[SuggestionItem(
            priority=s.priority,
            category=s.category,
            title=s.title,
            description=s.description,
            affected_cells=s.affected_cells,
            suggested_action=s.suggested_action,
            estimated_effort=s.estimated_effort,
            confidence=s.confidence,
        ) for s in report.suggestions],
        model_maturity_score=report.model_maturity_score,
    )


@router.post("/{uuid}/scenarios", response_model=ScenarioItem)
async def create_scenario(
    uuid: str,
    request: ScenarioCreateRequest,
    graphs=Depends(get_workbook_graphs),
    data_cache=Depends(get_workbook_data_cache),
    scenario_mgrs=Depends(get_scenario_managers),
):
    graph, workbook_data = _get_graph_and_data(uuid, graphs, data_cache)
    if uuid not in scenario_mgrs:
        scenario_mgrs[uuid] = ScenarioManager(graph, workbook_data)
    mgr = scenario_mgrs[uuid]
    if "Base Case" not in mgr.scenarios:
        mgr.create_base_case()
    s = mgr.create_scenario(request.name, request.description, request.input_overrides)
    return ScenarioItem(
        name=s.name, description=s.description, created_at=s.created_at,
        input_overrides=s.input_overrides, computed_outputs=s.computed_outputs,
        metadata=s.metadata,
    )


@router.post("/{uuid}/scenarios/perturbation", response_model=ScenarioItem)
async def create_perturbation_scenario(
    uuid: str,
    request: ScenarioPerturbRequest,
    graphs=Depends(get_workbook_graphs),
    data_cache=Depends(get_workbook_data_cache),
    scenario_mgrs=Depends(get_scenario_managers),
):
    graph, workbook_data = _get_graph_and_data(uuid, graphs, data_cache)
    if uuid not in scenario_mgrs:
        scenario_mgrs[uuid] = ScenarioManager(graph, workbook_data)
    mgr = scenario_mgrs[uuid]
    s = mgr.create_perturbation_scenario(request.name, request.description, request.perturbation_pct)
    return ScenarioItem(
        name=s.name, description=s.description, created_at=s.created_at,
        input_overrides=s.input_overrides, computed_outputs=s.computed_outputs,
        metadata=s.metadata,
    )


@router.get("/{uuid}/scenarios/compare", response_model=ScenarioCompareResponse)
async def compare_scenarios(
    uuid: str,
    graphs=Depends(get_workbook_graphs),
    data_cache=Depends(get_workbook_data_cache),
    scenario_mgrs=Depends(get_scenario_managers),
):
    graph, workbook_data = _get_graph_and_data(uuid, graphs, data_cache)
    if uuid not in scenario_mgrs:
        scenario_mgrs[uuid] = ScenarioManager(graph, workbook_data)
    mgr = scenario_mgrs[uuid]
    if "Base Case" not in mgr.scenarios:
        mgr.create_base_case()
    report = mgr.compare()
    return ScenarioCompareResponse(
        scenarios=[ScenarioItem(
            name=s.name, description=s.description, created_at=s.created_at,
            input_overrides=s.input_overrides, computed_outputs=s.computed_outputs,
            metadata=s.metadata,
        ) for s in report.scenarios],
        comparisons=[ScenarioComparisonItem(
            output_cell=c.output_cell, output_named_range=c.output_named_range,
            base_value=c.base_value, scenarios=c.scenarios,
            deltas=c.deltas, delta_pcts=c.delta_pcts,
        ) for c in report.comparisons],
        summary=report.summary,
    )


@router.post("/{uuid}/diff", response_model=VersionDiffResponse)
async def diff_workbooks(
    uuid: str,
    file_b: UploadFile = File(...),
    graphs=Depends(get_workbook_graphs),
    data_cache=Depends(get_workbook_data_cache),
):
    if uuid not in graphs:
        raise HTTPException(status_code=404, detail=f"Workbook {uuid} not found.")

    if not file_b.filename or not file_b.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx files supported")

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        content = await file_b.read()
        tmp.write(content)
        tmp_path = tmp.name

    try:
        wb_a = data_cache[uuid]
        differ = VersionDiffer()
        wb_b_data = differ.parser.parse(tmp_path)
        report = differ.diff_workbooks(wb_a, wb_b_data, label_a="current", label_b=file_b.filename)
        return VersionDiffResponse(
            file_a=report.file_a, file_b=report.file_b,
            total_changes=report.total_changes,
            sheets_added=report.sheets_added, sheets_removed=report.sheets_removed,
            sheet_diffs=[SheetDiffItem(
                sheet_name=s.sheet_name, cells_added=s.cells_added,
                cells_removed=s.cells_removed, cells_modified=s.cells_modified,
                formula_changes=s.formula_changes, value_changes=s.value_changes,
                max_delta_pct=s.max_delta_pct,
            ) for s in report.sheet_diffs],
            cell_diffs=[CellDiffItem(
                cell=d.cell, sheet=d.sheet, change_type=d.change_type,
                old_value=d.old_value, new_value=d.new_value,
                old_formula=d.old_formula, new_formula=d.new_formula,
                delta=d.delta, delta_pct=d.delta_pct, impact_score=d.impact_score,
            ) for d in report.cell_diffs[:200]],
            high_impact_changes=[CellDiffItem(
                cell=d.cell, sheet=d.sheet, change_type=d.change_type,
                old_value=d.old_value, new_value=d.new_value,
                old_formula=d.old_formula, new_formula=d.new_formula,
                delta=d.delta, delta_pct=d.delta_pct, impact_score=d.impact_score,
            ) for d in report.high_impact_changes],
            summary=report.summary,
        )
    finally:
        os.unlink(tmp_path)


@router.get("/{uuid}/audit", response_model=AuditResponse)
async def get_audit_trail(
    uuid: str,
    audit_trails_cache=Depends(get_audit_trails),
):
    trail = audit_trails_cache.get(uuid)
    if not trail:
        return AuditResponse(
            total_changes=0, changes=[], snapshots=[],
            cells_modified=0, risk_changes=[],
        )
    report = trail.generate_report()
    return AuditResponse(
        total_changes=report.total_changes,
        changes=[AuditChangeItem(
            change_id=c.change_id, timestamp=c.timestamp,
            cell=c.cell, sheet=c.sheet,
            old_value=c.old_value, new_value=c.new_value,
            reason=c.reason, approved_by=c.approved_by,
        ) for c in report.changes],
        snapshots=[AuditSnapshotItem(
            snapshot_id=s.snapshot_id, timestamp=s.timestamp, label=s.label,
        ) for s in report.snapshots],
        cells_modified=report.cells_modified,
        risk_changes=[AuditChangeItem(
            change_id=c.change_id, timestamp=c.timestamp,
            cell=c.cell, sheet=c.sheet,
            old_value=c.old_value, new_value=c.new_value,
            reason=c.reason, approved_by=c.approved_by,
        ) for c in report.risk_changes],
    )
