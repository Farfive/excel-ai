export interface WorkbookInfo {
  workbook_uuid: string;
  cell_count: number;
  cluster_count: number;
  anomaly_count: number;
  processing_time_ms: number;
}

export interface CellDiff {
  cell: string;
  before: unknown;
  after: unknown;
}

export interface ToolStep {
  step: number;
  tool: string;
  args: Record<string, unknown>;
  reason: string;
  status: 'pending' | 'running' | 'done' | 'error' | 'blocked' | 'warning';
  data?: unknown;
  diff?: CellDiff[];
  error?: string;
  concern?: string;
}

export interface PlanStep {
  step: number;
  tool: string;
  args: Record<string, unknown>;
  reason: string;
}

export interface ChatMessage {
  id: string;
  role: 'user' | 'assistant';
  content: string;
  timestamp: Date;
  toolSteps?: ToolStep[];
  isStreaming?: boolean;
  changeLog?: string;
}

export interface Anomaly {
  cell: string;
  value: unknown;
  anomaly_score: number;
  formula?: string;
  sheet_name?: string;
  named_range?: string;
  severity: 'high' | 'medium' | 'low';
}

export type SSEEventType =
  | 'retrieving'
  | 'planning'
  | 'plan_ready'
  | 'executing_plan'
  | 'executing'
  | 'tool_result'
  | 'reflection_warning'
  | 'reflection_blocked'
  | 'answer_start'
  | 'answer'
  | 'done'
  | 'error';

export interface SSEEvent {
  event: SSEEventType;
  data: {
    message?: string;
    plan?: PlanStep[];
    requiresApproval?: boolean;
    step?: number;
    tool?: string;
    args?: Record<string, unknown>;
    reason?: string;
    status?: string;
    success?: boolean;
    data?: unknown;
    error?: string;
    concern?: string;
    step_count?: number;
    chunk?: string;
    full_answer?: string;
  };
}

export interface HealthStatus {
  status: string;
  components: {
    ollama: { status: string };
    chroma: { status: string; detail?: string };
    embedder: { status: string };
    db: { status: string };
  };
  model: string;
}
