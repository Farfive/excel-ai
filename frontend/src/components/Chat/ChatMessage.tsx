import React from 'react';
import { ChatMessage as ChatMessageType } from '../../types';
import { CellChip } from '../shared/CellChip';
import { ToolStep } from './ToolStep';
import styles from './Chat.module.css';

interface ChatMessageProps {
  message: ChatMessageType;
}

const CELL_REF_REGEX = /\b([A-Z]{1,2}[0-9]+(?::[A-Z]{1,2}[0-9]+)?)\b/g;

function renderContentWithCellChips(content: string): React.ReactNode[] {
  const parts: React.ReactNode[] = [];
  let lastIndex = 0;
  let match: RegExpExecArray | null;
  CELL_REF_REGEX.lastIndex = 0;

  while ((match = CELL_REF_REGEX.exec(content)) !== null) {
    if (match.index > lastIndex) {
      parts.push(content.slice(lastIndex, match.index));
    }
    parts.push(<CellChip key={`${match[1]}-${match.index}`} address={match[1]} />);
    lastIndex = match.index + match[0].length;
  }

  if (lastIndex < content.length) {
    parts.push(content.slice(lastIndex));
  }

  return parts;
}

export function ChatMessage({ message }: ChatMessageProps): React.ReactElement {
  const isUser = message.role === 'user';
  const timeStr = message.timestamp.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });

  return (
    <div className={`${styles.message} ${isUser ? styles['message--user'] : styles['message--assistant']}`}>
      <div className={styles.message__bubble}>
        {renderContentWithCellChips(message.content)}
        {message.isStreaming && <span className={styles.message__cursor} aria-hidden="true" />}
      </div>

      {message.toolSteps && message.toolSteps.length > 0 && (
        <div className={styles.message__tools}>
          {message.toolSteps.map((step) => (
            <ToolStep key={step.step} step={step} />
          ))}
        </div>
      )}

      <span className={styles.message__time}>{timeStr}</span>
    </div>
  );
}
