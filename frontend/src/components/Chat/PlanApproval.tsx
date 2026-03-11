import React from 'react';
import { PlanStep } from '../../types';
import styles from './Chat.module.css';

interface PlanApprovalProps {
  plan: PlanStep[];
  onApprove: () => void;
  onReject: () => void;
}

export function PlanApproval({ plan, onApprove, onReject }: PlanApprovalProps): React.ReactElement {
  return (
    <div className={styles.planApproval}>
      <div className={styles.planApproval__title}>Execution Plan — Review &amp; Approve</div>
      <div className={styles.planApproval__steps}>
        {plan.map((step) => (
          <div key={step.step} className={styles.planApproval__step}>
            <span className={styles['planApproval__step-num']}>{step.step}</span>
            <div>
              <div className={styles['planApproval__step-tool']}>{step.tool}</div>
              <div className={styles['planApproval__step-reason']}>{step.reason}</div>
            </div>
          </div>
        ))}
      </div>
      <div className={styles.planApproval__actions}>
        <button className={styles.planApproval__approve} onClick={onApprove} type="button">
          Approve &amp; Execute
        </button>
        <button className={styles.planApproval__cancel} onClick={onReject} type="button">
          Cancel
        </button>
      </div>
    </div>
  );
}
