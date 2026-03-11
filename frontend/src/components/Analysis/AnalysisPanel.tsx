import React, { useState } from 'react';
import {
  runSensitivity, runIntegrityCheck, runSmartSuggestions,
  createPerturbationScenario, compareScenarios,
  validateFormulas,
} from '../../services/api';
import { CellChip } from '../shared/CellChip';
import styles from './Analysis.module.css';

interface AnalysisPanelProps {
  workbookUuid: string | null;
}

type AnalysisView = 'none' | 'integrity' | 'sensitivity' | 'suggestions' | 'scenarios' | 'formulas';

export function AnalysisPanel({ workbookUuid }: AnalysisPanelProps): React.ReactElement {
  const [activeView, setActiveView] = useState<AnalysisView>('none');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [data, setData] = useState<any>(null);

  const runAnalysis = async (view: AnalysisView) => {
    if (!workbookUuid) return;
    setActiveView(view);
    setLoading(true);
    setError(null);
    setData(null);
    try {
      let result;
      switch (view) {
        case 'integrity':
          result = await runIntegrityCheck(workbookUuid);
          break;
        case 'sensitivity':
          result = await runSensitivity(workbookUuid);
          break;
        case 'suggestions':
          result = await runSmartSuggestions(workbookUuid);
          break;
        case 'scenarios':
          await createPerturbationScenario(workbookUuid, 'Upside +10%', 'All inputs +10%', 10);
          await createPerturbationScenario(workbookUuid, 'Downside -10%', 'All inputs -10%', -10);
          await createPerturbationScenario(workbookUuid, 'Stress -25%', 'All inputs -25%', -25);
          result = await compareScenarios(workbookUuid);
          break;
        case 'formulas':
          result = await validateFormulas(workbookUuid);
          break;
      }
      setData(result);
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : 'Analysis failed');
    } finally {
      setLoading(false);
    }
  };

  const scoreColor = (score: number) => {
    if (score >= 80) return styles['scoreCard__value--good'];
    if (score >= 50) return styles['scoreCard__value--warn'];
    return styles['scoreCard__value--bad'];
  };

  return (
    <div className={styles.analysisPanel}>
      <div className={styles.analysisPanel__section}>
        <div className={styles['analysisPanel__section-title']}>Model Analysis</div>
        <div className={styles['analysisPanel__btn-row']}>
          <button
            className={`${styles.analysisPanel__btn} ${activeView === 'integrity' ? styles['analysisPanel__btn--active'] : ''}`}
            onClick={() => runAnalysis('integrity')}
            disabled={loading || !workbookUuid}
            type="button"
          >
            Integrity
          </button>
          <button
            className={`${styles.analysisPanel__btn} ${activeView === 'sensitivity' ? styles['analysisPanel__btn--active'] : ''}`}
            onClick={() => runAnalysis('sensitivity')}
            disabled={loading || !workbookUuid}
            type="button"
          >
            Sensitivity
          </button>
          <button
            className={`${styles.analysisPanel__btn} ${activeView === 'suggestions' ? styles['analysisPanel__btn--active'] : ''}`}
            onClick={() => runAnalysis('suggestions')}
            disabled={loading || !workbookUuid}
            type="button"
          >
            Suggestions
          </button>
          <button
            className={`${styles.analysisPanel__btn} ${activeView === 'scenarios' ? styles['analysisPanel__btn--active'] : ''}`}
            onClick={() => runAnalysis('scenarios')}
            disabled={loading || !workbookUuid}
            type="button"
          >
            Scenarios
          </button>
          <button
            className={`${styles.analysisPanel__btn} ${activeView === 'formulas' ? styles['analysisPanel__btn--active'] : ''}`}
            onClick={() => runAnalysis('formulas')}
            disabled={loading || !workbookUuid}
            type="button"
          >
            Formulas
          </button>
        </div>
      </div>

      {error && <div className={styles.analysisPanel__error}>{error}</div>}

      {loading && (
        <div className={styles.analysisPanel__loading}>
          Analyzing...
        </div>
      )}

      <div className={styles.analysisPanel__results}>
        {!loading && !data && activeView === 'none' && (
          <div className={styles.analysisPanel__empty}>
            Select an analysis to run
          </div>
        )}

        {!loading && data && activeView === 'integrity' && (
          <>
            <div className={styles.scoreCard}>
              <div className={`${styles.scoreCard__value} ${scoreColor(data.model_health_score)}`}>
                {data.model_health_score?.toFixed(0)}
              </div>
              <div className={styles.scoreCard__label}>
                Model Health Score<br />
                {data.critical} critical, {data.warning} warnings, {data.info} info
              </div>
            </div>
            {data.issues?.map((issue: any, i: number) => (
              <div key={i} className={styles.issueRow}>
                <span className={`${styles.issueRow__badge} ${styles[`issueRow__badge--${issue.severity}`]}`}>
                  {issue.severity}
                </span>
                <div className={styles.issueRow__content}>
                  <div className={styles.issueRow__message}>{issue.message}</div>
                  <div className={styles.issueRow__cell}>{issue.cell}</div>
                  <div className={styles.issueRow__suggestion}>{issue.suggestion}</div>
                </div>
              </div>
            ))}
          </>
        )}

        {!loading && data && activeView === 'sensitivity' && (
          <>
            <div className={styles.scoreCard}>
              <div className={styles.scoreCard__value} style={{ color: 'var(--text-primary)' }}>
                {data.input_cells_tested}
              </div>
              <div className={styles.scoreCard__label}>
                Inputs tested against {data.output_cells_monitored} outputs<br />
                {data.results?.length ?? 0} data points generated
              </div>
            </div>
            <div className={styles['analysisPanel__section-title']} style={{ marginTop: 8 }}>Top Drivers</div>
            {data.top_drivers?.map((driver: any, i: number) => {
              const maxScore = data.top_drivers?.[0]?.total_impact_score || 1;
              const pct = Math.min(100, (driver.total_impact_score / maxScore) * 100);
              return (
                <div key={i} className={styles.driverRow}>
                  <div className={styles.driverRow__cell}>
                    <CellChip address={driver.cell} />
                  </div>
                  <div className={styles.driverRow__bar}>
                    <div className={styles.driverRow__barFill} style={{ width: `${pct}%` }} />
                  </div>
                  <div className={styles.driverRow__score}>
                    {driver.total_impact_score?.toFixed(1)}
                  </div>
                </div>
              );
            })}
          </>
        )}

        {!loading && data && activeView === 'suggestions' && (
          <>
            <div className={styles.scoreCard}>
              <div className={`${styles.scoreCard__value} ${scoreColor(data.model_maturity_score)}`}>
                {data.model_maturity_score?.toFixed(0)}
              </div>
              <div className={styles.scoreCard__label}>
                Model Maturity Score<br />
                {data.high_priority} high, {data.medium_priority} medium, {data.low_priority} low
              </div>
            </div>
            {data.suggestions?.map((s: any, i: number) => (
              <div key={i} className={styles.issueRow}>
                <span className={`${styles.issueRow__badge} ${styles[`issueRow__badge--${s.priority}`]}`}>
                  {s.priority}
                </span>
                <div className={styles.issueRow__content}>
                  <div className={styles.issueRow__message}>{s.title}</div>
                  <div className={styles.issueRow__suggestion}>{s.suggested_action}</div>
                </div>
              </div>
            ))}
          </>
        )}

        {!loading && data && activeView === 'formulas' && (
          <>
            <div className={styles.scoreCard}>
              <div className={`${styles.scoreCard__value} ${scoreColor(data.consistency_score)}`}>
                {data.consistency_score?.toFixed(0)}
              </div>
              <div className={styles.scoreCard__label}>
                Formula Consistency Score<br />
                {data.formula_count} formulas checked · {data.errors} errors, {data.warnings} warnings, {data.info} info
              </div>
            </div>
            {data.issues?.map((issue: any, i: number) => (
              <div key={i} className={styles.issueRow}>
                <span className={`${styles.issueRow__badge} ${styles[`issueRow__badge--${issue.severity === 'error' ? 'critical' : issue.severity}`]}`}>
                  {issue.severity}
                </span>
                <div className={styles.issueRow__content}>
                  <div className={styles.issueRow__message}>{issue.message}</div>
                  <div className={styles.issueRow__cell}>{issue.cell}</div>
                  {issue.expected && (
                    <div className={styles.issueRow__suggestion}>Expected: {issue.expected}</div>
                  )}
                </div>
              </div>
            ))}
          </>
        )}

        {!loading && data && activeView === 'scenarios' && (
          <>
            <div className={styles.scoreCard}>
              <div className={styles.scoreCard__value} style={{ color: 'var(--text-primary)' }}>
                {data.scenarios?.length ?? 0}
              </div>
              <div className={styles.scoreCard__label}>
                Scenarios compared<br />
                {data.summary}
              </div>
            </div>
            {data.comparisons?.slice(0, 15).map((comp: any, i: number) => {
              const entries = Object.entries(comp.delta_pcts || {});
              return (
                <div key={i} style={{ marginBottom: 6 }}>
                  <div style={{ fontFamily: 'var(--font-mono)', fontSize: 10, color: 'var(--text-secondary)', marginBottom: 2 }}>
                    <CellChip address={comp.output_cell} /> base: {comp.base_value?.toFixed(2)}
                  </div>
                  {entries.map(([name, pct]: [string, any]) => (
                    <div key={name} className={styles.scenarioRow}>
                      <span className={styles.scenarioRow__name}>{name}</span>
                      <span className={`${styles.scenarioRow__delta} ${pct >= 0 ? styles['scenarioRow__delta--positive'] : styles['scenarioRow__delta--negative']}`}>
                        {pct >= 0 ? '+' : ''}{pct?.toFixed(2)}%
                      </span>
                    </div>
                  ))}
                </div>
              );
            })}
          </>
        )}
      </div>
    </div>
  );
}
