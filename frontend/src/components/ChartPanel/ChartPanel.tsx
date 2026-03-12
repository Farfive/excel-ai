import React, { useState, useMemo } from 'react';
import {
  Chart as ChartJS,
  CategoryScale, LinearScale, BarElement, PointElement, LineElement,
  ArcElement, Filler, Tooltip, Legend, Title,
} from 'chart.js';
import { Bar, Line, Pie, Scatter, Doughnut } from 'react-chartjs-2';
import { GridCell } from '../../services/api';
import styles from './ChartPanel.module.css';

ChartJS.register(
  CategoryScale, LinearScale, BarElement, PointElement, LineElement,
  ArcElement, Filler, Tooltip, Legend, Title,
);

type ChartType = 'bar' | 'line' | 'pie' | 'scatter' | 'doughnut' | 'area';

interface Props {
  rows: GridCell[][];
  colHeaders: string[];
  sheetName: string;
}

const COLORS = [
  '#2563EB', '#F97316', '#16A34A', '#DC2626', '#8B5CF6',
  '#EC4899', '#14B8A6', '#EAB308', '#6366F1', '#F43F5E',
];

function parseNum(v: string): number | null {
  if (!v) return null;
  const n = parseFloat(v.replace(/[,$%]/g, ''));
  return isNaN(n) ? null : n;
}

export function ChartPanel({ rows, colHeaders, sheetName }: Props): React.ReactElement {
  const [chartType, setChartType] = useState<ChartType>('bar');
  const [labelCol, setLabelCol] = useState(0);
  const [dataCols, setDataCols] = useState<number[]>([1]);

  const numericCols = useMemo(() => {
    const result: number[] = [];
    colHeaders.forEach((_, ci) => {
      let numCount = 0;
      for (let r = 1; r < Math.min(rows.length, 20); r++) {
        if (parseNum(rows[r]?.[ci]?.v || '') !== null) numCount++;
      }
      if (numCount > 0) result.push(ci);
    });
    return result;
  }, [rows, colHeaders]);

  const labels = useMemo(() => {
    return rows.slice(1).map(row => row[labelCol]?.v || '').filter(Boolean);
  }, [rows, labelCol]);

  const datasets = useMemo(() => {
    return dataCols.map((ci, idx) => {
      const header = rows[0]?.[ci]?.v || colHeaders[ci] || `Col ${ci}`;
      const data = rows.slice(1).map(row => parseNum(row[ci]?.v || '') ?? 0);
      const color = COLORS[idx % COLORS.length];
      return {
        label: header,
        data,
        backgroundColor: chartType === 'pie' || chartType === 'doughnut'
          ? data.map((_, i) => COLORS[i % COLORS.length])
          : color + '99',
        borderColor: color,
        borderWidth: chartType === 'line' || chartType === 'area' ? 2 : 1,
        fill: chartType === 'area',
        tension: 0.3,
        pointRadius: chartType === 'scatter' ? 4 : 2,
      };
    });
  }, [rows, dataCols, colHeaders, chartType]);

  const chartData = { labels, datasets };

  const options = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { position: 'top' as const, labels: { font: { size: 11 } } },
      title: { display: true, text: `${sheetName} Chart`, font: { size: 13 } },
    },
    scales: chartType !== 'pie' && chartType !== 'doughnut' ? {
      x: { grid: { color: 'rgba(0,0,0,0.06)' } },
      y: { grid: { color: 'rgba(0,0,0,0.06)' }, beginAtZero: true },
    } : undefined,
  };

  if (rows.length < 2) {
    return <div className={styles.empty}>No data to chart. Upload a workbook with at least 2 rows.</div>;
  }

  const toggleDataCol = (ci: number) => {
    setDataCols(prev =>
      prev.includes(ci) ? prev.filter(c => c !== ci) : [...prev, ci]
    );
  };

  return (
    <div className={styles.wrap}>
      <div className={styles.controls}>
        <span className={styles.label}>Type</span>
        <div className={styles.typeRow}>
          {(['bar', 'line', 'area', 'pie', 'doughnut', 'scatter'] as ChartType[]).map(t => (
            <button
              key={t}
              className={`${styles.btn} ${chartType === t ? styles['btn--active'] : ''}`}
              onClick={() => setChartType(t)}
              type="button"
            >{t.charAt(0).toUpperCase() + t.slice(1)}</button>
          ))}
        </div>
      </div>

      <div className={styles.controls}>
        <span className={styles.label}>Labels</span>
        <select
          className={styles.select}
          value={labelCol}
          onChange={e => setLabelCol(Number(e.target.value))}
        >
          {colHeaders.map((h, ci) => (
            <option key={ci} value={ci}>{h} — {rows[0]?.[ci]?.v || h}</option>
          ))}
        </select>
      </div>

      <div className={styles.controls}>
        <span className={styles.label}>Data columns</span>
        {numericCols.map(ci => (
          <button
            key={ci}
            className={`${styles.btn} ${dataCols.includes(ci) ? styles['btn--active'] : ''}`}
            onClick={() => toggleDataCol(ci)}
            type="button"
          >{colHeaders[ci]} — {rows[0]?.[ci]?.v || colHeaders[ci]}</button>
        ))}
      </div>

      <div className={styles.chartContainer}>
        {chartType === 'bar' && <Bar data={chartData} options={options as any} />}
        {(chartType === 'line' || chartType === 'area') && <Line data={chartData} options={options as any} />}
        {chartType === 'pie' && <Pie data={chartData} options={options as any} />}
        {chartType === 'doughnut' && <Doughnut data={chartData} options={options as any} />}
        {chartType === 'scatter' && <Scatter data={{
          datasets: dataCols.map((ci, idx) => ({
            label: rows[0]?.[ci]?.v || colHeaders[ci],
            data: rows.slice(1).map((row, ri) => ({
              x: parseNum(row[labelCol]?.v || '') ?? ri,
              y: parseNum(row[ci]?.v || '') ?? 0,
            })),
            backgroundColor: COLORS[idx % COLORS.length] + '99',
            borderColor: COLORS[idx % COLORS.length],
            pointRadius: 4,
          })),
        }} options={options as any} />}
      </div>
    </div>
  );
}
