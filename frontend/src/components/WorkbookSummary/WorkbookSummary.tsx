import React, { useEffect, useState } from 'react';
import { getSummary, WorkbookSummaryData, SheetSummary } from '../../services/api';
import styles from './WorkbookSummary.module.css';

interface Props {
  workbookUuid: string;
  onSheetClick?: (sheetName: string) => void;
}

function formatNumber(v: number | string): string {
  if (typeof v === 'string') return v;
  if (Math.abs(v) >= 1_000_000) return `${(v / 1_000_000).toFixed(1)}M`;
  if (Math.abs(v) >= 1_000) return `${(v / 1_000).toFixed(1)}K`;
  if (Number.isInteger(v)) return v.toLocaleString();
  return v.toFixed(2);
}

function qualityColor(score: number): string {
  if (score >= 80) return '#16A34A';
  if (score >= 50) return '#D97706';
  return '#DC2626';
}

function qualityLabel(score: number): { text: string; desc: string } {
  if (score >= 90) return { text: 'Excellent', desc: 'Well-structured model with strong formula coverage and no significant issues.' };
  if (score >= 80) return { text: 'Good', desc: 'Solid model structure. Minor improvements possible.' };
  if (score >= 65) return { text: 'Fair', desc: 'Some structural gaps — check empty cells and formula coverage.' };
  if (score >= 50) return { text: 'Needs Review', desc: 'Significant issues detected — high empty cell ratio or anomalies.' };
  return { text: 'Poor', desc: 'Critical quality problems — model needs restructuring.' };
}

function groupDeps(deps: { from: string; to: string }[]): Map<string, string[]> {
  const grouped = new Map<string, string[]>();
  for (const dep of deps) {
    const list = grouped.get(dep.from) || [];
    list.push(dep.to);
    grouped.set(dep.from, list);
  }
  return grouped;
}

function QualityRing({ score }: { score: number }) {
  const r = 26;
  const circ = 2 * Math.PI * r;
  const offset = circ - (score / 100) * circ;
  const color = qualityColor(score);

  return (
    <div className={styles.qualityRing}>
      <svg width="64" height="64" viewBox="0 0 64 64">
        <circle className={styles.qualityRing__bg} cx="32" cy="32" r={r} />
        <circle
          className={styles.qualityRing__fill}
          cx="32" cy="32" r={r}
          stroke={color}
          strokeDasharray={circ}
          strokeDashoffset={offset}
        />
      </svg>
      <span className={styles.qualityRing__text}>{score}</span>
    </div>
  );
}

function SheetCard({ sheet, onClick }: { sheet: SheetSummary; onClick: () => void }) {
  return (
    <div className={styles.sheetCard} onClick={onClick} role="button" tabIndex={0}>
      <div className={styles.sheetCard__header}>
        <span className={styles.sheetCard__name}>{sheet.name}</span>
        <span className={styles.sheetCard__size}>{sheet.rows}r × {sheet.cols}c</span>
      </div>

      <div className={styles.sheetCard__stats}>
        <span className={styles.sheetCard__stat}>{sheet.cell_count} cells</span>
        <span className={`${styles.sheetCard__stat} ${styles['sheetCard__stat--blue']}`}>
          {sheet.formula_count} formulas
        </span>
        {sheet.anomaly_count > 0 && (
          <span className={`${styles.sheetCard__stat} ${styles['sheetCard__stat--warn']}`}>
            {sheet.anomaly_count} anomalies
          </span>
        )}
      </div>

      {sheet.key_metrics.length > 0 && (
        <div className={styles.metrics}>
          {sheet.key_metrics.slice(0, 5).map((m, i) => (
            <div className={styles.metricRow} key={i}>
              <span className={styles.metricRow__label}>{m.label}</span>
              <span className={`${styles.metricRow__value} ${m.has_formula ? styles.metricRow__formula : ''}`}>
                {formatNumber(m.value)}
              </span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export function WorkbookSummary({ workbookUuid, onSheetClick }: Props): React.ReactElement {
  const [data, setData] = useState<WorkbookSummaryData | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    setLoading(true);
    setError(null);

    getSummary(workbookUuid)
      .then(d => { if (!cancelled) setData(d); })
      .catch(e => { if (!cancelled) setError(e.message); })
      .finally(() => { if (!cancelled) setLoading(false); });

    return () => { cancelled = true; };
  }, [workbookUuid]);

  if (loading) {
    return (
      <div className={styles.loading}>
        <div className={styles.spinner} />
        Loading summary...
      </div>
    );
  }

  if (error) {
    return <div className={styles.error}>{error}</div>;
  }

  if (!data) return <div className={styles.error}>No data</div>;

  return (
    <div className={styles.dashboard}>
      {/* ── Top stats ── */}
      <div className={styles.statsRow}>
        <div className={styles.statCard}>
          <span className={styles.statCard__value}>{data.sheet_count}</span>
          <span className={styles.statCard__label}>Sheets</span>
        </div>
        <div className={styles.statCard}>
          <span className={styles.statCard__value}>{data.total_cells.toLocaleString()}</span>
          <span className={styles.statCard__label}>Total Cells</span>
        </div>
        <div className={styles.statCard}>
          <span className={styles.statCard__value}>{data.total_formulas.toLocaleString()}</span>
          <span className={styles.statCard__label}>Formulas</span>
        </div>
        <div className={styles.statCard}>
          <span className={styles.statCard__value}>{data.formula_pct}%</span>
          <span className={styles.statCard__label}>Formula Coverage</span>
        </div>
        <div className={styles.statCard}>
          <span className={styles.statCard__value}>{data.total_anomalies}</span>
          <span className={styles.statCard__label}>Anomalies</span>
        </div>
      </div>

      {/* ── Quality score ── */}
      <div className={styles.qualityCard}>
        <QualityRing score={data.quality_score} />
        <div className={styles.qualityInfo}>
          <span className={styles.qualityInfo__title}>
            Data Quality Score — <span style={{ color: qualityColor(data.quality_score) }}>{qualityLabel(data.quality_score).text}</span>
          </span>
          <span className={styles.qualityInfo__desc}>{qualityLabel(data.quality_score).desc}</span>
          <div className={styles.qualityInfo__detail}>
            <span className={styles.qualityInfo__item}>
              <span className={styles.qualityInfo__dot} style={{ background: '#2563EB' }} />
              {data.formula_pct}% formulas
            </span>
            <span className={styles.qualityInfo__item}>
              <span className={styles.qualityInfo__dot} style={{ background: '#ADB5BD' }} />
              {data.empty_pct}% empty
            </span>
            <span className={styles.qualityInfo__item}>
              <span className={styles.qualityInfo__dot} style={{ background: '#D97706' }} />
              {data.total_anomalies} anomalies
            </span>
            <span className={styles.qualityInfo__item}>
              <span className={styles.qualityInfo__dot} style={{ background: '#7C3AED' }} />
              {data.named_ranges} named ranges
            </span>
          </div>
        </div>
      </div>

      {/* ── Cross-sheet dependencies (grouped) ── */}
      {data.cross_sheet_deps.length > 0 && (
        <>
          <div className={styles.sectionHeader}>Cross-Sheet Dependencies ({data.cross_sheet_deps.length})</div>
          <div className={styles.depsGroupWrap}>
            {Array.from(groupDeps(data.cross_sheet_deps).entries()).map(([source, targets]) => (
              <div className={styles.depGroup} key={source}>
                <span className={styles.depGroup__source}>{source}</span>
                <span className={styles.depGroup__arrow}>→</span>
                <div className={styles.depGroup__targets}>
                  {targets.map(t => (
                    <span className={styles.depGroup__target} key={t}>{t}</span>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </>
      )}

      {/* ── Sheet cards ── */}
      <div className={styles.sectionHeader}>Sheets Overview ({data.sheet_count})</div>
      <div className={styles.sheetsGrid}>
        {data.sheets.map(sheet => (
          <SheetCard
            key={sheet.name}
            sheet={sheet}
            onClick={() => onSheetClick?.(sheet.name)}
          />
        ))}
      </div>
    </div>
  );
}
