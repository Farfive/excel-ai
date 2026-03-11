import React from 'react';
import ReactDOM from 'react-dom/client';
import { App } from './components/App/App';
import './styles/theme.css';
import './styles/global.css';

let initialized = false;

function renderApp() {
  if (initialized) return;
  initialized = true;
  const container = document.getElementById('app');
  if (!container) return;
  const root = ReactDOM.createRoot(container);
  root.render(<App />);
}

function init() {
  try {
    const g = window as any;
    if (g.Office && typeof g.Office.onReady === 'function') {
      g.Office.onReady(() => renderApp());
      setTimeout(() => renderApp(), 2500);
    } else {
      renderApp();
    }
  } catch {
    renderApp();
  }
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
