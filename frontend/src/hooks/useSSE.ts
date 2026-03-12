import { useState, useCallback, useRef } from 'react';
import { ChatMessage, ToolStep, PlanStep, SSEEvent, ImpactPreview } from '../types';
import { askQuestion } from '../services/api';
import { emitGridRefresh, emitGridDiff, CellChange } from '../services/gridBus';

let msgIdCounter = 0;
function nextId(): string {
  return `msg_${++msgIdCounter}_${Date.now()}`;
}

export type AgentMode = 'agent' | 'plan' | 'ask';

interface UseSSEReturn {
  messages: ChatMessage[];
  isStreaming: boolean;
  pendingPlan: PlanStep[] | null;
  impactPreview: ImpactPreview | null;
  mode: AgentMode;
  setMode: (m: AgentMode) => void;
  sendMessage: (question: string, approvedPlan?: PlanStep[]) => Promise<void>;
  approvePlan: () => Promise<void>;
  rejectPlan: () => void;
}

export function useSSE(workbookUuid: string | null): UseSSEReturn {
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [isStreaming, setIsStreaming] = useState(false);
  const [pendingPlan, setPendingPlan] = useState<PlanStep[] | null>(null);
  const [impactPreview, setImpactPreview] = useState<ImpactPreview | null>(null);
  const [mode, setMode] = useState<AgentMode>('agent');
  const pendingQuestionRef = useRef<string>('');
  const modeRef = useRef<AgentMode>(mode);
  modeRef.current = mode;

  const sendMessage = useCallback(async (question: string, approvedPlan?: PlanStep[]) => {
    if (!workbookUuid) return;
    const currentMode = modeRef.current;

    if (!approvedPlan) {
      pendingQuestionRef.current = question;
      const userMsg: ChatMessage = {
        id: nextId(),
        role: 'user',
        content: question,
        timestamp: new Date(),
      };
      setMessages(prev => [...prev, userMsg]);
    }

    const assistantMsgId = nextId();
    const assistantMsg: ChatMessage = {
      id: assistantMsgId,
      role: 'assistant',
      content: '',
      timestamp: new Date(),
      toolSteps: [],
      isStreaming: true,
    };
    setMessages(prev => [...prev, assistantMsg]);

    setIsStreaming(true);

    const collectedDiffs: CellChange[] = [];

    try {
      for await (const event of askQuestion(workbookUuid, question, approvedPlan, currentMode)) {
        handleEvent(event, assistantMsgId);
        const DIFF_TOOLS = ['write_range', 'write_formula', 'add_column', 'add_row', 'delete_range', 'create_sheet', 'pivot_table'];
        if (event.event === 'tool_result' && event.data.success &&
            DIFF_TOOLS.includes(event.data.tool ?? '')) {
          const resultData = event.data.data as Record<string, unknown> | undefined;
          const diff = resultData?.diff as CellChange[] | undefined;
          if (diff && diff.length > 0) {
            collectedDiffs.push(...diff);
          }
        }
      }
    } catch (e: unknown) {
      const errMsg = e instanceof Error ? e.message : 'Unknown error';
      setMessages(prev =>
        prev.map(m =>
          m.id === assistantMsgId
            ? { ...m, content: `Error: ${errMsg}`, isStreaming: false }
            : m
        )
      );
    } finally {
      setIsStreaming(false);
      setMessages(prev =>
        prev.map(m => m.id === assistantMsgId ? { ...m, isStreaming: false } : m)
      );
      if (collectedDiffs.length > 0) {
        emitGridRefresh();
        setTimeout(() => emitGridDiff(collectedDiffs), 300);
      }
    }
  }, [workbookUuid]);

  function handleEvent(event: SSEEvent, assistantMsgId: string) {
    const { event: type, data } = event;

    if (type === 'plan_ready' && data.requiresApproval && data.plan) {
      setPendingPlan(data.plan);
      setImpactPreview(data.impact_preview ?? null);
      setIsStreaming(false);
      setMessages(prev =>
        prev.map(m =>
          m.id === assistantMsgId
            ? { ...m, content: 'Plan ready — please review and approve below.', isStreaming: false }
            : m
        )
      );
      return;
    }

    if (type === 'executing_plan') {
      setMessages(prev =>
        prev.map(m =>
          m.id === assistantMsgId
            ? { ...m, content: data.message || `Executing ${data.step_count || ''} steps...` }
            : m
        )
      );
      return;
    }

    if (type === 'executing') {
      const step: ToolStep = {
        step: data.step ?? 0,
        tool: data.tool ?? '',
        args: (data.args as Record<string, unknown>) ?? {},
        reason: data.reason ?? '',
        status: 'running',
      };
      setMessages(prev =>
        prev.map(m => {
          if (m.id !== assistantMsgId) return m;
          const existing = m.toolSteps ?? [];
          return { ...m, toolSteps: [...existing, step] };
        })
      );
    }

    if (type === 'tool_result') {
      const resultData = data.data as Record<string, unknown> | undefined;
      const diff = (resultData?.diff as import('../types').CellDiff[] | undefined) ?? undefined;
      setMessages(prev =>
        prev.map(m => {
          if (m.id !== assistantMsgId) return m;
          const steps = (m.toolSteps ?? []).map(s =>
            s.step === data.step
              ? { ...s, status: data.success ? 'done' as const : 'error' as const, data: data.data, diff, error: data.error }
              : s
          );
          return { ...m, toolSteps: steps };
        })
      );
    }

    if (type === 'reflection_warning') {
      setMessages(prev =>
        prev.map(m => {
          if (m.id !== assistantMsgId) return m;
          const steps = (m.toolSteps ?? []).map(s =>
            s.step === data.step ? { ...s, status: 'warning' as const, concern: data.concern } : s
          );
          return { ...m, toolSteps: steps };
        })
      );
    }

    if (type === 'reflection_blocked') {
      setMessages(prev =>
        prev.map(m => {
          if (m.id !== assistantMsgId) return m;
          const steps = (m.toolSteps ?? []).map(s =>
            s.step === data.step ? { ...s, status: 'blocked' as const, concern: data.concern } : s
          );
          return { ...m, toolSteps: steps, isStreaming: false };
        })
      );
      setIsStreaming(false);
    }

    if (type === 'answer') {
      setMessages(prev =>
        prev.map(m =>
          m.id === assistantMsgId ? { ...m, content: m.content + (data.chunk ?? '') } : m
        )
      );
    }

    if (type === 'done') {
      setMessages(prev =>
        prev.map(m =>
          m.id === assistantMsgId
            ? { ...m, isStreaming: false, changeLog: typeof data.full_answer === 'string' ? undefined : undefined }
            : m
        )
      );
    }
  }

  const approvePlan = useCallback(async () => {
    if (!pendingPlan) return;
    const plan = pendingPlan;
    const question = pendingQuestionRef.current;
    setPendingPlan(null);
    setImpactPreview(null);
    await sendMessage(question, plan);
  }, [pendingPlan, sendMessage]);

  const rejectPlan = useCallback(() => {
    setPendingPlan(null);
    setImpactPreview(null);
    setMessages(prev => {
      const last = prev[prev.length - 1];
      if (last && last.role === 'assistant' && last.isStreaming) {
        return prev.map(m =>
          m.id === last.id ? { ...m, content: 'Plan cancelled.', isStreaming: false } : m
        );
      }
      return prev;
    });
  }, []);

  return { messages, isStreaming, pendingPlan, impactPreview, mode, setMode, sendMessage, approvePlan, rejectPlan };
}
