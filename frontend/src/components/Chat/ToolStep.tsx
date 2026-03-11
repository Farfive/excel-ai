import React from 'react';
import { ToolStep as ToolStepType } from '../../types';
import styles from './Chat.module.css';

interface ToolStepProps {
  step: ToolStepType;
}

export function ToolStep({ step }: ToolStepProps): React.ReactElement {
  const argsPreview = Object.entries(step.args)
    .slice(0, 2)
    .map(([k, v]) => `${k}=${JSON.stringify(v)}`)
    .join(', ');

  return (
    <div className={styles.toolStep}>
      <span className={`${styles.toolStep__indicator} ${styles[`toolStep__indicator--${step.status}`]}`} aria-hidden="true" />
      <div className={styles.toolStep__label}>
        <span>→ {step.tool}({argsPreview})</span>
        {step.status === 'warning' && step.concern && (
          <div className={styles['toolStep__concern--warning']}>{step.concern}</div>
        )}
        {step.status === 'blocked' && step.concern && (
          <div className={styles['toolStep__concern--blocked']}>{step.concern}</div>
        )}
        {step.error && (
          <div className={styles['toolStep__concern--blocked']}>{step.error}</div>
        )}
      </div>
    </div>
  );
}
