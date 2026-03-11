import { WorkbookInfo, Anomaly, SSEEvent, PlanStep, HealthStatus } from '../types';

const BASE_URL = 'http://localhost:8000';

export async function uploadWorkbook(file: File): Promise<WorkbookInfo> {
  const formData = new FormData();
  formData.append('file', file);

  const res = await fetch(`${BASE_URL}/workbook/upload`, {
    method: 'POST',
    body: formData,
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: 'Upload failed' }));
    throw new Error(err.detail || `HTTP ${res.status}`);
  }

  return res.json();
}

export async function* askQuestion(
  uuid: string,
  question: string,
  approvedPlan?: PlanStep[],
  mode: 'agent' | 'plan' = 'agent',
): AsyncGenerator<SSEEvent> {
  const res = await fetch(`${BASE_URL}/workbook/${uuid}/ask`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ question, approved_plan: approvedPlan ?? null, mode }),
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: 'Request failed' }));
    throw new Error(err.detail || `HTTP ${res.status}`);
  }

  const reader = res.body?.getReader();
  if (!reader) throw new Error('No response body');

  const decoder = new TextDecoder();
  let buffer = '';

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;

    buffer += decoder.decode(value, { stream: true });
    const lines = buffer.split('\n');
    buffer = lines.pop() ?? '';

    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed.startsWith('data:')) continue;
      const jsonStr = trimmed.slice(5).trim();
      if (!jsonStr) continue;
      try {
        const event: SSEEvent = JSON.parse(jsonStr);
        yield event;
      } catch {
        // malformed line — skip
      }
    }
  }
}

export async function getAnomalies(uuid: string): Promise<Anomaly[]> {
  const res = await fetch(`${BASE_URL}/workbook/${uuid}/anomalies`);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const data = await res.json();
  return (data.anomalies ?? []).map((a: Anomaly) => ({
    ...a,
    severity: anomalyScoreToSeverity(a.anomaly_score),
  }));
}

function anomalyScoreToSeverity(score: number): 'high' | 'medium' | 'low' {
  if (score < -0.1) return 'high';
  if (score < 0.0) return 'medium';
  return 'low';
}

export async function downloadWorkbook(uuid: string): Promise<void> {
  const res = await fetch(`${BASE_URL}/workbook/${uuid}/download`);
  if (!res.ok) throw new Error(`Download failed: HTTP ${res.status}`);
  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `workbook_${uuid.slice(0, 8)}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export async function confirmDelete(uuid: string, sheet: string, range: string): Promise<{ success: boolean; data: unknown; error: string | null }> {
  const res = await fetch(`${BASE_URL}/workbook/${uuid}/confirm-delete`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ sheet, range }),
  });
  if (!res.ok) throw new Error(`Confirm delete failed: HTTP ${res.status}`);
  return res.json();
}

export async function revertChanges(uuid: string, changes: { cell: string; before: unknown }[]): Promise<{ reverted: number }> {
  const res = await fetch(`${BASE_URL}/workbook/${uuid}/revert`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(changes),
  });
  if (!res.ok) throw new Error(`Revert failed: HTTP ${res.status}`);
  return res.json();
}

export async function sendDeltaUpdate(uuid: string, changedCells: string[]): Promise<void> {
  const res = await fetch(`${BASE_URL}/workbook/${uuid}/delta`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ changed_cells: changedCells }),
  });
  if (!res.ok) throw new Error(`Delta update failed: HTTP ${res.status}`);
}

export async function updateWorkbookState(uuid: string, cells: Record<string, unknown>): Promise<void> {
  const res = await fetch(`${BASE_URL}/workbook/${uuid}/state`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ cells }),
  });
  if (!res.ok) throw new Error(`State update failed: HTTP ${res.status}`);
}

export interface AuditChange {
  id: string;
  timestamp: string;
  cell: string;
  sheet: string;
  old_value: unknown;
  new_value: unknown;
  reason: string;
  approved_by: string;
  step: number | null;
}

export async function getAuditTrail(uuid: string): Promise<{ changes: AuditChange[]; total: number }> {
  const res = await fetch(`${BASE_URL}/workbook/${uuid}/audit-trail`);
  if (!res.ok) throw new Error(`Audit trail failed: HTTP ${res.status}`);
  return res.json();
}

export async function getHealth(): Promise<HealthStatus> {
  const res = await fetch(`${BASE_URL}/health`);
  if (!res.ok) throw new Error(`Health check failed: HTTP ${res.status}`);
  return res.json();
}

export async function runSensitivity(
  uuid: string,
  maxInputs: number = 15,
  maxOutputs: number = 8,
): Promise<any> {
  const res = await fetch(`${BASE_URL}/workbook/${uuid}/sensitivity`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ max_inputs: maxInputs, max_outputs: maxOutputs }),
  });
  if (!res.ok) throw new Error(`Sensitivity failed: HTTP ${res.status}`);
  return res.json();
}

export async function validateFormulas(uuid: string): Promise<any> {
  const res = await fetch(`${BASE_URL}/workbook/${uuid}/validate-formulas`);
  if (!res.ok) throw new Error(`Formula validation failed: HTTP ${res.status}`);
  return res.json();
}

export async function runIntegrityCheck(uuid: string): Promise<any> {
  const res = await fetch(`${BASE_URL}/workbook/${uuid}/integrity`);
  if (!res.ok) throw new Error(`Integrity check failed: HTTP ${res.status}`);
  return res.json();
}

export async function runSmartSuggestions(uuid: string): Promise<any> {
  const res = await fetch(`${BASE_URL}/workbook/${uuid}/suggestions`);
  if (!res.ok) throw new Error(`Suggestions failed: HTTP ${res.status}`);
  return res.json();
}

export async function createScenario(
  uuid: string,
  name: string,
  description: string,
  inputOverrides: Record<string, number>,
): Promise<any> {
  const res = await fetch(`${BASE_URL}/workbook/${uuid}/scenarios`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, description, input_overrides: inputOverrides }),
  });
  if (!res.ok) throw new Error(`Create scenario failed: HTTP ${res.status}`);
  return res.json();
}

export async function createPerturbationScenario(
  uuid: string,
  name: string,
  description: string,
  perturbationPct: number,
): Promise<any> {
  const res = await fetch(`${BASE_URL}/workbook/${uuid}/scenarios/perturbation`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, description, perturbation_pct: perturbationPct }),
  });
  if (!res.ok) throw new Error(`Perturbation scenario failed: HTTP ${res.status}`);
  return res.json();
}

export async function compareScenarios(uuid: string): Promise<any> {
  const res = await fetch(`${BASE_URL}/workbook/${uuid}/scenarios/compare`);
  if (!res.ok) throw new Error(`Compare scenarios failed: HTTP ${res.status}`);
  return res.json();
}

export async function diffWorkbooks(uuid: string, fileB: File): Promise<any> {
  const formData = new FormData();
  formData.append('file_b', fileB);
  const res = await fetch(`${BASE_URL}/workbook/${uuid}/diff`, {
    method: 'POST',
    body: formData,
  });
  if (!res.ok) throw new Error(`Diff failed: HTTP ${res.status}`);
  return res.json();
}

export async function getWorkbookGrid(uuid: string): Promise<{
  sheets: Array<{
    name: string;
    colHeaders: string[];
    rows: Array<Array<{ v: string; f: boolean; t: string; nr?: string }>>;
  }>;
}> {
  const res = await fetch(`${BASE_URL}/workbook/${uuid}/grid`);
  if (!res.ok) throw new Error(`Grid fetch failed: HTTP ${res.status}`);
  return res.json();
}
