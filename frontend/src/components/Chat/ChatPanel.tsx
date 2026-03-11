import React, { useEffect, useRef, useState } from 'react';
import { useSSE } from '../../hooks/useSSE';
import { ChatMessage } from './ChatMessage';
import { PlanApproval } from './PlanApproval';
import styles from './Chat.module.css';

interface ChatPanelProps {
  workbookUuid: string | null;
}

export function ChatPanel({ workbookUuid }: ChatPanelProps): React.ReactElement {
  const { messages, isStreaming, pendingPlan, sendMessage, approvePlan, rejectPlan } = useSSE(workbookUuid);
  const [input, setInput] = useState('');
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, isStreaming]);

  const handleSend = () => {
    const q = input.trim();
    if (!q || isStreaming || !workbookUuid) return;
    setInput('');
    sendMessage(q);
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const handleTextareaChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setInput(e.target.value);
    const ta = textareaRef.current;
    if (ta) {
      ta.style.height = 'auto';
      ta.style.height = Math.min(ta.scrollHeight, 80) + 'px';
    }
  };

  const showDots = isStreaming && messages.length > 0 && messages[messages.length - 1]?.content === '';

  return (
    <div className={styles.chat}>
      <div className={styles.chat__messages}>
        {messages.length === 0 && !workbookUuid && (
          <div className={styles.chat__empty}>Upload a workbook to start</div>
        )}
        {messages.length === 0 && workbookUuid && (
          <div className={styles.chat__empty}>Ask a question about the workbook</div>
        )}
        {messages.map((msg) => (
          <ChatMessage key={msg.id} message={msg} />
        ))}
        {showDots && (
          <div className={styles.chat__dots}>
            <span className={styles.chat__dot} />
            <span className={styles.chat__dot} />
            <span className={styles.chat__dot} />
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      {pendingPlan && (
        <PlanApproval plan={pendingPlan} onApprove={approvePlan} onReject={rejectPlan} />
      )}

      <div className={styles['chat__input-bar']}>
        <textarea
          ref={textareaRef}
          className={styles.chat__textarea}
          value={input}
          onChange={handleTextareaChange}
          onKeyDown={handleKeyDown}
          placeholder={workbookUuid ? 'Ask about the workbook…' : 'Upload a workbook first'}
          disabled={!workbookUuid || isStreaming}
          rows={1}
          aria-label="Chat input"
        />
        <button
          className={styles['chat__send-btn']}
          onClick={handleSend}
          disabled={!input.trim() || isStreaming || !workbookUuid}
          type="button"
          aria-label="Send message"
        >
          Send
        </button>
      </div>
    </div>
  );
}
