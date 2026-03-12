import React, { useRef, useEffect, useState } from 'react';
import { useSSE, AgentMode } from '../../hooks/useSSE';
import { ChatMessage, ToolStep, PlanStep, CellDiff, ImpactPreview, ImpactKPI } from '../../types';
import { confirmDelete } from '../../services/api';
import { emitGridRefresh, emitGridDiff } from '../../services/gridBus';
import styles from './ChatSidebar.module.css';

interface Props {
  workbookUuid: string | null;
}

const MODE_CONFIG = {
  agent: {
    icon: (
      <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
        <polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/>
      </svg>
    ),
    title: 'Agent Mode',
    sub: <>AI automatically plans and executes actions.<br />Just describe what you want done.</>,
    hints: [
      'Change WACC to 9% and update all downstream cells',
      'Add a new "Scenarios" sheet with Base / Bull / Bear cases',
      'Fix all hardcoded values in the Assumptions sheet',
      'Recalculate Revenue Growth Rate for 2026 to 15%',
    ],
  },
  plan: {
    icon: (
      <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
        <path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>
      </svg>
    ),
    title: 'Plan Mode',
    sub: <>AI shows execution plan + impact analysis<br />before making any changes. You approve first.</>,
    hints: [
      'Update Revenue Growth Rate to 20% — show impact',
      'Increase EBITDA margin assumptions by 2pp',
      'Change Discount Rate and show effect on Enterprise Value',
      'Adjust CapEx % of Revenue across all years',
    ],
  },
  ask: {
    icon: (
      <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
        <circle cx="12" cy="12" r="10"/><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/><line x1="12" y1="17" x2="12.01" y2="17"/>
      </svg>
    ),
    title: 'Ask Mode',
    sub: <>Read-only Q&amp;A — AI answers questions<br />and never modifies your data.</>,
    hints: [
      'What is the DCF implied share price?',
      'Explain the revenue assumptions on the Assumptions sheet',
      'What are the top 3 anomalies in this model?',
      'What drives the biggest change in Enterprise Value?',
    ],
  },
};

function EmptyState({ mode, workbookUuid, onHintClick }: { mode: AgentMode; workbookUuid: string | null; onHintClick: (h: string) => void }) {
  if (!workbookUuid) {
    return (
      <div className={styles.empty}>
        <svg className={styles.empty__icon} width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5">
          <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
        </svg>
        <div className={styles.empty__title}>Upload a workbook to start</div>
        <div className={styles.empty__sub}>Upload an .xlsx file to analyse,<br />modify or ask questions about it.</div>
      </div>
    );
  }
  const cfg = MODE_CONFIG[mode];
  return (
    <div className={styles.empty}>
      <div className={styles.empty__icon}>{cfg.icon}</div>
      <div className={styles.empty__title}>{cfg.title}</div>
      <div className={styles.empty__sub}>{cfg.sub}</div>
      <div className={styles.empty__hints}>
        {cfg.hints.map(h => (
          <button key={h} className={styles.empty__hint} onClick={() => onHintClick(h)} type="button">
            {h}
          </button>
        ))}
      </div>
    </div>
  );
}

function formatTime(d: Date): string {
  return d.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
}

function formatVal(v: unknown): string {
  if (v === null || v === undefined) return '—';
  if (typeof v === 'number') return Number.isInteger(v) ? String(v) : v.toFixed(4).replace(/\.?0+$/, '');
  return String(v);
}

function DiffTable({ diff }: { diff: CellDiff[] }) {
  return (
    <div className={styles.diff}>
      {diff.map(d => (
        <div key={d.cell} className={styles.diff__row}>
          <span className={styles.diff__cell}>{d.cell}</span>
          <span className={`${styles.diff__val} ${styles['diff__val--before']}`}>- {formatVal(d.before)}</span>
          <span className={`${styles.diff__val} ${styles['diff__val--after']}`}>+ {formatVal(d.after)}</span>
        </div>
      ))}
    </div>
  );
}

const WRITE_TOOLS = ['write_range', 'write_formula', 'add_column', 'add_row', 'delete_range'];

interface DeletePreviewItem { cell: string; value: unknown; formula?: string; data_type?: string; }

function DeletePreview({ data, onConfirm, onCancel }: { data: { preview: DeletePreviewItem[]; message: string }; onConfirm: () => void; onCancel: () => void }) {
  return (
    <div className={styles.deletePreview}>
      <div className={styles.deletePreview__header}>{data.message}</div>
      <div className={styles.deletePreview__table}>
        {data.preview.map((item: DeletePreviewItem) => (
          <div key={item.cell} className={styles.deletePreview__row}>
            <span className={styles.deletePreview__cell}>{item.cell}</span>
            <span className={styles.deletePreview__val}>{formatVal(item.value)}</span>
          </div>
        ))}
      </div>
      <div className={styles.deletePreview__actions}>
        <button className={styles.deletePreview__confirm} onClick={onConfirm} type="button">
          Confirm Delete
        </button>
        <button className={styles.deletePreview__cancel} onClick={onCancel} type="button">
          Cancel
        </button>
      </div>
    </div>
  );
}

function ToolStepRow({ step, workbookUuid, onDeleteConfirmed }: { step: ToolStep; workbookUuid: string | null; onDeleteConfirmed?: () => void }) {
  const hasDiff = WRITE_TOOLS.includes(step.tool) && step.diff && step.diff.length > 0;
  const resultData = step.data as Record<string, unknown> | undefined;
  const isDeletePreview = step.tool === 'delete_range' && resultData?.requires_confirmation === true;
  const [deleteHandled, setDeleteHandled] = useState(false);

  const handleConfirmDelete = async () => {
    if (!workbookUuid || !step.args) return;
    try {
      const res = await confirmDelete(workbookUuid, step.args.sheet as string, step.args.range as string);
      if (res.success) {
        setDeleteHandled(true);
        emitGridRefresh();
        const diff = (res.data as Record<string, unknown>)?.diff as import('../../services/gridBus').CellChange[] | undefined;
        if (diff && diff.length > 0) {
          setTimeout(() => emitGridDiff(diff), 300);
        }
        onDeleteConfirmed?.();
      }
    } catch (e) {
      console.error('Confirm delete failed:', e);
    }
  };

  return (
    <div className={styles.toolStep}>
      <span className={`${styles.toolStep__dot} ${styles[`toolStep__dot--${step.status}`]}`} />
      <div style={{ flex: 1 }}>
        <span className={styles.toolStep__name}>{step.tool}</span>
        <span style={{ color: 'var(--text-muted)', marginLeft: 6 }}>{step.reason}</span>
        {step.concern && (
          <div className={`${styles.toolStep__concern} ${step.status === 'blocked' ? styles['toolStep__concern--blocked'] : styles['toolStep__concern--warning']}`}>
            {step.concern}
          </div>
        )}
        {hasDiff && <DiffTable diff={step.diff!} />}
        {isDeletePreview && !deleteHandled && (
          <DeletePreview
            data={resultData as { preview: DeletePreviewItem[]; message: string }}
            onConfirm={handleConfirmDelete}
            onCancel={() => setDeleteHandled(true)}
          />
        )}
        {isDeletePreview && deleteHandled && (
          <div style={{ fontSize: 10, color: 'var(--text-muted)', fontFamily: 'var(--font-mono)', marginTop: 4 }}>Deletion handled.</div>
        )}
      </div>
    </div>
  );
}

function PlanView({ plan, impact, onApprove, onReject }: { plan: PlanStep[]; impact: ImpactPreview | null; onApprove: () => void; onReject: () => void }) {
  const [expandedStep, setExpandedStep] = useState<number | null>(null);
  const writeSteps = plan.filter(s => ['write_range', 'write_formula', 'add_column', 'add_row', 'delete_range'].includes(s.tool));
  const isReadOnly = writeSteps.length === 0;
  return (
    <div className={styles.plan}>
      <div className={styles.plan__header}>
        <span className={styles.plan__title}>Execution Plan</span>
        <span className={styles.plan__subtitle}>{plan.length} step{plan.length !== 1 ? 's' : ''} · requires approval</span>
      </div>

      {/* Impact Preview */}
      {impact && !isReadOnly && (
        <div className={styles.impact}>
          <div className={styles.impact__header}>
            <svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor"><path d="M8 1a.5.5 0 0 1 .5.5v6h6a.5.5 0 0 1 0 1h-6v6a.5.5 0 0 1-1 0v-6h-6a.5.5 0 0 1 0-1h6v-6A.5.5 0 0 1 8 1z"/></svg>
            Impact Analysis
            <span className={styles.impact__badge}>{impact.cells_affected} cells affected</span>
          </div>
          {impact.written.length > 0 && (
            <div className={styles.impact__writes}>
              {impact.written.map(w => (
                <div key={w.cell} className={styles.impact__write}>
                  <span className={styles.impact__writeCell}>{w.cell.split('!')[1] ?? w.cell}</span>
                  <span className={styles.impact__writeArrow}>→</span>
                  <span className={styles.impact__writeVal}>{w.new_value}</span>
                </div>
              ))}
            </div>
          )}
          {impact.kpis.length > 0 && (
            <>
              <div className={styles.impact__kpiTitle}>Affected KPIs (current values)</div>
              <div className={styles.impact__kpis}>
                {impact.kpis.map((kpi: ImpactKPI) => (
                  <div key={kpi.cell} className={styles.impact__kpi}>
                    <div className={styles.impact__kpiLabel}>{kpi.label}</div>
                    <div className={styles.impact__kpiValue}>{kpi.current_value}</div>
                    <div className={styles.impact__kpiSheet}>{kpi.sheet}</div>
                  </div>
                ))}
              </div>
            </>
          )}
          {impact.kpis.length === 0 && impact.cells_affected > 0 && (
            <div className={styles.impact__noKpi}>
              {impact.cells_affected} formula cells will recalculate
            </div>
          )}
        </div>
      )}

      <div className={styles.plan__steps}>
        {plan.map(step => {
          const hasArgs = step.args && Object.keys(step.args).length > 0;
          const isExpanded = expandedStep === step.step;
          return (
            <div
              key={step.step}
              className={styles.plan__step}
              onClick={() => hasArgs && setExpandedStep(isExpanded ? null : step.step)}
              style={{ cursor: hasArgs ? 'pointer' : 'default' }}
            >
              <div className={styles['plan__step-num']}>{step.step}</div>
              <div className={styles['plan__step-body']}>
                <div className={styles['plan__step-tool']}>
                  {step.tool}
                  {hasArgs && (
                    <span style={{ marginLeft: 6, fontSize: 9, color: 'var(--text-muted)' }}>
                      {isExpanded ? '▲' : '▼'}
                    </span>
                  )}
                </div>
                <div className={styles['plan__step-reason']}>{step.reason}</div>
                {hasArgs && isExpanded && (
                  <div className={styles['plan__step-args']}>
                    {JSON.stringify(step.args, null, 2)}
                  </div>
                )}
              </div>
            </div>
          );
        })}
      </div>
      <div className={styles.plan__actions}>
        <button className={styles.plan__approve} onClick={onApprove} type="button">
          <svg width="13" height="13" viewBox="0 0 16 16" fill="currentColor">
            <path d="M13.854 3.646a.5.5 0 0 1 0 .708l-7 7a.5.5 0 0 1-.708 0l-3.5-3.5a.5.5 0 1 1 .708-.708L6.5 10.293l6.646-6.647a.5.5 0 0 1 .708 0z"/>
          </svg>
          Approve & Run
        </button>
        <button className={styles.plan__reject} onClick={onReject} type="button">
          Cancel
        </button>
      </div>
    </div>
  );
}

function MessageRow({ message, workbookUuid }: { message: ChatMessage; workbookUuid: string | null }) {
  const isUser = message.role === 'user';
  return (
    <div className={`${styles.msgRow} ${isUser ? styles['msgRow--user'] : styles['msgRow--assistant']}`}>
      <div className={styles.msgRow__label}>{isUser ? 'You' : 'AI'}</div>
      {message.toolSteps && message.toolSteps.length > 0 && (
        <div className={styles.toolSteps}>
          {message.toolSteps.map(s => <ToolStepRow key={s.step} step={s} workbookUuid={workbookUuid} />)}
        </div>
      )}
      {message.content && (
        <div className={`${styles.bubble} ${isUser ? styles['bubble--user'] : styles['bubble--assistant']}`}>
          {message.content}
          {message.isStreaming && <span className={styles.bubble__cursor} />}
        </div>
      )}
      <div className={styles.msgRow__time}>{formatTime(message.timestamp)}</div>
    </div>
  );
}

export function ChatSidebar({ workbookUuid }: Props): React.ReactElement {
  const { messages, isStreaming, pendingPlan, impactPreview, mode, setMode, sendMessage, approvePlan, rejectPlan } = useSSE(workbookUuid);
  const [input, setInput] = useState('');
  const bottomRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, isStreaming, pendingPlan]);

  const handleSend = () => {
    const q = input.trim();
    if (!q || isStreaming || !workbookUuid) return;
    setInput('');
    if (textareaRef.current) textareaRef.current.style.height = 'auto';
    sendMessage(q);
  };

  const handleKey = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSend(); }
  };

  const handleChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setInput(e.target.value);
    const ta = textareaRef.current;
    if (ta) { ta.style.height = 'auto'; ta.style.height = Math.min(ta.scrollHeight, 120) + 'px'; }
  };

  const handleClear = () => {
    window.location.reload();
  };

  const showDots = isStreaming && messages.length > 0 && messages[messages.length - 1]?.content === '';

  return (
    <div className={styles.sidebar}>
      {/* Header */}
      <div className={styles.header}>
        <span className={styles.header__title}>AI Chat</span>
        <div className={styles.modeToggle}>
          <button
            className={`${styles.modeToggle__btn} ${mode === 'agent' ? styles['modeToggle__btn--active'] : ''}`}
            onClick={() => setMode('agent')}
            type="button"
            title="Agent mode: AI automatically executes actions"
          >
            ⚡ Agent
          </button>
          <button
            className={`${styles.modeToggle__btn} ${mode === 'plan' ? styles['modeToggle__btn--active'] : ''}`}
            onClick={() => setMode('plan')}
            type="button"
            title="Plan mode: AI shows plan for approval before executing"
          >
            📋 Plan
          </button>
          <button
            className={`${styles.modeToggle__btn} ${mode === 'ask' ? styles['modeToggle__btn--active'] : ''}`}
            onClick={() => setMode('ask')}
            type="button"
            title="Ask mode: read-only Q&A, AI never modifies data"
          >
            💬 Ask
          </button>
        </div>
        {messages.length > 0 && (
          <button className={styles.header__clear} onClick={handleClear} type="button">Clear</button>
        )}
      </div>

      {/* Messages */}
      <div className={styles.messages}>
        {messages.length === 0 ? (
          <EmptyState
            key={mode}
            mode={mode}
            workbookUuid={workbookUuid}
            onHintClick={(h) => { setInput(h); textareaRef.current?.focus(); }}
          />
        ) : (
          messages.map(m => <MessageRow key={m.id} message={m} workbookUuid={workbookUuid} />)
        )}

        {showDots && (
          <div className={styles.dots}>
            <span className={styles.dot} />
            <span className={styles.dot} />
            <span className={styles.dot} />
          </div>
        )}

        <div ref={bottomRef} />
      </div>

      {/* Plan approval */}
      {pendingPlan && (
        <PlanView plan={pendingPlan} impact={impactPreview} onApprove={approvePlan} onReject={rejectPlan} />
      )}

      {/* Input */}
      <div className={styles.inputArea}>
        <div className={styles.inputWrap}>
          <textarea
            ref={textareaRef}
            className={styles.textarea}
            value={input}
            onChange={handleChange}
            onKeyDown={handleKey}
            placeholder={!workbookUuid ? 'Upload a workbook first' : mode === 'ask' ? 'Ask a question (read-only, no changes)…' : 'Ask anything about the workbook…'}
            disabled={!workbookUuid || isStreaming}
            rows={1}
          />
          <div className={styles.inputFooter}>
            {isStreaming ? (
              <div className={styles.streamBadge}>
                <span className={styles.streamBadge__dot} />
                Generating…
              </div>
            ) : (
              <span className={styles.inputHint}>Enter to send · Shift+Enter for newline</span>
            )}
            <button
              className={styles.sendBtn}
              onClick={handleSend}
              disabled={!input.trim() || isStreaming || !workbookUuid}
              type="button"
            >
              <svg width="11" height="11" viewBox="0 0 16 16" fill="currentColor">
                <path d="M15.854.146a.5.5 0 0 1 .11.54l-5.819 14.547a.75.75 0 0 1-1.329.124l-3.178-4.995L.643 7.184a.75.75 0 0 1 .124-1.33L15.314.037a.5.5 0 0 1 .54.11z"/>
              </svg>
              Send
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
