import React, { useState, useRef, useEffect } from 'react';
import { CellStyleData } from '../../services/api';
import styles from './FormatToolbar.module.css';

interface Props {
  currentStyle: CellStyleData | null;
  hasSelection: boolean;
  onFormat: (style: Partial<CellStyleData>) => void;
}

const FONTS = ['Calibri', 'Arial', 'Times New Roman', 'Courier New', 'Verdana', 'Georgia', 'Tahoma', 'Trebuchet MS'];
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, 72];
const NUM_FORMATS: Array<{ label: string; value: string }> = [
  { label: 'General', value: 'General' },
  { label: '0', value: '0' },
  { label: '0.00', value: '0.00' },
  { label: '#,##0', value: '#,##0' },
  { label: '#,##0.00', value: '#,##0.00' },
  { label: '$#,##0', value: '$#,##0.00' },
  { label: '0%', value: '0%' },
  { label: '0.00%', value: '0.00%' },
  { label: 'mm/dd/yyyy', value: 'mm/dd/yyyy' },
];

function toHex(color: string | undefined): string {
  if (!color) return '';
  let c = color.replace('#', '').trim();
  if (c.length === 8) c = c.slice(2);
  if (c.length === 6) return `#${c}`;
  if (c.length === 3) return `#${c[0]}${c[0]}${c[1]}${c[1]}${c[2]}${c[2]}`;
  return '';
}

export function FormatToolbar({ currentStyle, hasSelection, onFormat }: Props): React.ReactElement {
  const [showBorders, setShowBorders] = useState(false);
  const borderRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!showBorders) return;
    const close = (e: MouseEvent) => {
      if (borderRef.current && !borderRef.current.contains(e.target as Node)) setShowBorders(false);
    };
    window.addEventListener('mousedown', close);
    return () => window.removeEventListener('mousedown', close);
  }, [showBorders]);

  const s = currentStyle || {};
  const disabled = !hasSelection;

  return (
    <div className={styles.ribbon}>
      {/* Font family */}
      <select
        className={`${styles.select} ${styles.fontSelect}`}
        value={s.fn || 'Calibri'}
        onChange={e => onFormat({ fn: e.target.value })}
        disabled={disabled}
        title="Font"
      >
        {FONTS.map(f => <option key={f} value={f}>{f}</option>)}
      </select>

      {/* Font size */}
      <select
        className={`${styles.select} ${styles.sizeSelect}`}
        value={s.fs || 11}
        onChange={e => onFormat({ fs: Number(e.target.value) })}
        disabled={disabled}
        title="Font Size"
      >
        {SIZES.map(sz => <option key={sz} value={sz}>{sz}</option>)}
      </select>

      <span className={styles.sep} />

      {/* Bold */}
      <button
        className={`${styles.btn} ${s.b ? styles['btn--active'] : ''}`}
        onClick={() => onFormat({ b: !s.b })}
        disabled={disabled}
        title="Bold (Ctrl+B)"
        type="button"
      ><strong>B</strong></button>

      {/* Italic */}
      <button
        className={`${styles.btn} ${s.i ? styles['btn--active'] : ''}`}
        onClick={() => onFormat({ i: !s.i })}
        disabled={disabled}
        title="Italic (Ctrl+I)"
        type="button"
      ><em>I</em></button>

      {/* Underline */}
      <button
        className={`${styles.btn} ${s.u ? styles['btn--active'] : ''}`}
        onClick={() => onFormat({ u: !s.u })}
        disabled={disabled}
        title="Underline (Ctrl+U)"
        type="button"
      ><span style={{ textDecoration: 'underline' }}>U</span></button>

      <span className={styles.sep} />

      {/* Font color */}
      <div className={styles.colorWrap} title="Font Color">
        <span className={styles.colorLabel}>A</span>
        <div
          className={styles.colorSwatch}
          style={{ background: toHex(s.fc) || '#000000' }}
        />
        <input
          type="color"
          className={styles.colorInput}
          value={toHex(s.fc) || '#000000'}
          onChange={e => onFormat({ fc: e.target.value })}
          disabled={disabled}
        />
      </div>

      {/* Background color */}
      <div className={styles.colorWrap} title="Fill Color">
        <svg width="12" height="12" viewBox="0 0 16 16" fill={toHex(s.bg) || '#FFFF00'} stroke="var(--text-muted)" strokeWidth="1">
          <rect x="1" y="1" width="14" height="14" rx="2" />
        </svg>
        <div
          className={styles.colorSwatch}
          style={{ background: toHex(s.bg) || '#FFFF00' }}
        />
        <input
          type="color"
          className={styles.colorInput}
          value={toHex(s.bg) || '#FFFFFF'}
          onChange={e => onFormat({ bg: e.target.value })}
          disabled={disabled}
        />
      </div>

      <span className={styles.sep} />

      {/* Horizontal alignment */}
      <button
        className={`${styles.btn} ${s.ha === 'left' ? styles['btn--active'] : ''}`}
        onClick={() => onFormat({ ha: 'left' })}
        disabled={disabled}
        title="Align Left"
        type="button"
      >
        <svg viewBox="0 0 16 16" fill="currentColor"><path d="M2 3h12v1H2zm0 3h8v1H2zm0 3h12v1H2zm0 3h8v1H2z"/></svg>
      </button>
      <button
        className={`${styles.btn} ${s.ha === 'center' ? styles['btn--active'] : ''}`}
        onClick={() => onFormat({ ha: 'center' })}
        disabled={disabled}
        title="Align Center"
        type="button"
      >
        <svg viewBox="0 0 16 16" fill="currentColor"><path d="M2 3h12v1H2zm2 3h8v1H4zm-2 3h12v1H2zm2 3h8v1H4z"/></svg>
      </button>
      <button
        className={`${styles.btn} ${s.ha === 'right' ? styles['btn--active'] : ''}`}
        onClick={() => onFormat({ ha: 'right' })}
        disabled={disabled}
        title="Align Right"
        type="button"
      >
        <svg viewBox="0 0 16 16" fill="currentColor"><path d="M2 3h12v1H2zm4 3h8v1H6zm-4 3h12v1H2zm4 3h8v1H6z"/></svg>
      </button>

      <span className={styles.sep} />

      {/* Wrap text */}
      <button
        className={`${styles.btn} ${s.wt ? styles['btn--active'] : ''}`}
        onClick={() => onFormat({ wt: !s.wt })}
        disabled={disabled}
        title="Wrap Text"
        type="button"
        style={{ fontSize: 9, width: 'auto', padding: '0 4px' }}
      >↵</button>

      {/* Borders */}
      <div style={{ position: 'relative' }} ref={borderRef}>
        <button
          className={styles.btn}
          onClick={() => setShowBorders(!showBorders)}
          disabled={disabled}
          title="Borders"
          type="button"
        >
          <svg viewBox="0 0 16 16" fill="currentColor"><path d="M1 1h14v14H1zm1 1v12h12V2zm5 0v12M2 8h12" stroke="currentColor" strokeWidth="0.5"/></svg>
        </button>
        {showBorders && (
          <div className={styles.borderMenu}>
            <button className={styles.btn} onClick={() => { onFormat({ bb: 'thin' }); setShowBorders(false); }} type="button" title="Bottom">▁</button>
            <button className={styles.btn} onClick={() => { onFormat({ bt: 'thin' }); setShowBorders(false); }} type="button" title="Top">▔</button>
            <button className={styles.btn} onClick={() => { onFormat({ bl: 'thin' }); setShowBorders(false); }} type="button" title="Left">▏</button>
            <button className={styles.btn} onClick={() => { onFormat({ br: 'thin' }); setShowBorders(false); }} type="button" title="Right">▕</button>
            <button className={styles.btn} onClick={() => { onFormat({ bt: 'thin', bb: 'thin', bl: 'thin', br: 'thin' }); setShowBorders(false); }} type="button" title="All">▣</button>
            <button className={styles.btn} onClick={() => { onFormat({ bt: '', bb: '', bl: '', br: '' }); setShowBorders(false); }} type="button" title="None">⊘</button>
            <button className={styles.btn} onClick={() => { onFormat({ bb: 'thick' }); setShowBorders(false); }} type="button" title="Thick Bottom" style={{ fontWeight: 800 }}>▁</button>
            <button className={styles.btn} onClick={() => { onFormat({ bt: 'thick', bb: 'thick', bl: 'thick', br: 'thick' }); setShowBorders(false); }} type="button" title="Thick All" style={{ fontWeight: 800 }}>▣</button>
            <button className={styles.btn} onClick={() => { onFormat({ bb: 'double' }); setShowBorders(false); }} type="button" title="Double Bottom">=</button>
          </div>
        )}
      </div>

      <span className={styles.sep} />

      {/* Number format */}
      <select
        className={`${styles.select} ${styles.numFmtSelect}`}
        value={s.nf || 'General'}
        onChange={e => onFormat({ nf: e.target.value })}
        disabled={disabled}
        title="Number Format"
      >
        {NUM_FORMATS.map(nf => <option key={nf.value} value={nf.value}>{nf.label}</option>)}
      </select>

      {/* Indent */}
      <button
        className={styles.btn}
        onClick={() => onFormat({ ind: Math.max(0, (s.ind || 0) - 1) })}
        disabled={disabled}
        title="Decrease Indent"
        type="button"
        style={{ fontSize: 10 }}
      >⇤</button>
      <button
        className={styles.btn}
        onClick={() => onFormat({ ind: (s.ind || 0) + 1 })}
        disabled={disabled}
        title="Increase Indent"
        type="button"
        style={{ fontSize: 10 }}
      >⇥</button>
    </div>
  );
}
