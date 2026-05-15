import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';

// One-time cleanup: unregister any old service workers and clear their caches.
// This defends against stale builds being served to users whose browsers cached an older version.
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.getRegistrations().then(regs => regs.forEach(r => r.unregister())).catch(() => {});
}
if ('caches' in window) {
  caches.keys().then(keys => keys.forEach(k => caches.delete(k))).catch(() => {});
}

ReactDOM.createRoot(document.getElementById('root')).render(<App />);
