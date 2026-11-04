import React, { StrictMode, Component, ErrorInfo, ReactNode } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App.tsx';
import './index.css';

interface ErrorBoundaryProps {
  children: ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
  error: Error | null;
}

class ErrorBoundary extends React.Component<ErrorBoundaryProps, ErrorBoundaryState> {
  props: ErrorBoundaryProps;
  state: ErrorBoundaryState = { hasError: false, error: null };

  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.props = props;
  }

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    console.error('Unhandled runtime error caught by ErrorBoundary:', error, errorInfo);
  }

  handleReload = () => {
    window.location.reload();
  };

  handleClearCacheAndReload = () => {
    try {
      sessionStorage.clear();
      // Keep main records safe in IndexedDB / localStorage
    } catch {}
    window.location.reload();
  };

  render() {
    if (this.state.hasError) {
      return (
        <div style={{
          minHeight: '100vh',
          backgroundColor: '#0f172a',
          color: '#ffffff',
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          justifyContent: 'center',
          padding: '24px',
          textAlign: 'center',
          fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif'
        }}>
          <div style={{
            maxWidth: '500px',
            backgroundColor: '#1e293b',
            border: '1px solid #334155',
            borderRadius: '24px',
            padding: '32px',
            boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.5)'
          }}>
            <div style={{
              width: '56px',
              height: '56px',
              backgroundColor: 'rgba(239, 68, 68, 0.2)',
              border: '2px solid rgba(239, 68, 68, 0.4)',
              borderRadius: '18px',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              margin: '0 auto 16px auto',
              color: '#f87171',
              fontSize: '28px'
            }}>
              ⚠️
            </div>
            <h2 style={{ fontSize: '18px', fontWeight: 900, textTransform: 'uppercase', margin: '0 0 8px 0', color: '#f8fafc' }}>
              စနစ်လုပ်ဆောင်ချက် ခေတ္တချို့ယွင်းချက်ဖြစ်ပေါ်သည်
            </h2>
            <p style={{ fontSize: '12px', color: '#94a3b8', margin: '0 0 20px 0', lineHeight: 1.6 }}>
              အချက်အလက်များသည် Local Database တွင် လုံခြုံစွာ ရှိနေပါသည်။ ကျေးဇူးပြု၍ စနစ်အား ပြန်လည် Reload ပြုလုပ်ပါ။
            </p>
            {this.state.error?.message && (
              <div style={{
                backgroundColor: '#0f172a',
                border: '1px solid #334155',
                borderRadius: '12px',
                padding: '10px 14px',
                fontSize: '11px',
                fontFamily: 'monospace',
                color: '#fca5a5',
                marginBottom: '20px',
                textAlign: 'left',
                overflowX: 'auto',
                maxHeight: '100px'
              }}>
                {this.state.error.message}
              </div>
            )}
            <div style={{ display: 'flex', gap: '10px', justifyContent: 'center' }}>
              <button
                onClick={this.handleReload}
                style={{
                  backgroundColor: '#3b82f6',
                  color: '#ffffff',
                  border: 'none',
                  borderRadius: '12px',
                  padding: '10px 20px',
                  fontSize: '12px',
                  fontWeight: 'bold',
                  cursor: 'pointer'
                }}
              >
                🔄 Reload App (ပြန်လည်ဖွင့်မည်)
              </button>
            </div>
          </div>
        </div>
      );
    }

    return this.props.children;
  }
}

// Force Update Version Check for Next Launch across all devices
const BUILD_VERSION = 'v11.0_20261104_force_update_v2';

async function checkAndApplyForceUpdate() {
  try {
    const storedVersion = localStorage.getItem('app_build_version');
    if (storedVersion !== BUILD_VERSION) {
      console.log(`[Force Update] Upgrading app build from ${storedVersion || 'legacy'} to ${BUILD_VERSION}...`);
      if ('caches' in window) {
        const keys = await caches.keys();
        await Promise.all(keys.map(k => caches.delete(k)));
      }
      if ('serviceWorker' in navigator) {
        const registrations = await navigator.serviceWorker.getRegistrations();
        for (const reg of registrations) {
          await reg.unregister();
        }
      }
      localStorage.setItem('app_build_version', BUILD_VERSION);
      // Automatically refresh to load fresh bundle
      window.location.reload();
      return;
    }
  } catch (e) {
    console.error('Force update check failed:', e);
    localStorage.setItem('app_build_version', BUILD_VERSION);
  }
}

// Execute update check immediately on next launch
checkAndApplyForceUpdate();

// Register PWA Service Worker & Monitor for Updates
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('/sw.js').then(
      (reg) => {
        console.log('PWA ServiceWorker registered with scope: ', reg.scope);
        reg.update().catch(() => {});
        reg.onupdatefound = () => {
          const installingWorker = reg.installing;
          if (installingWorker) {
            installingWorker.onstatechange = () => {
              if (installingWorker.state === 'installed' && navigator.serviceWorker.controller) {
                console.log('New update installed! Reloading app...');
                window.location.reload();
              }
            };
          }
        };
      },
      (err) => console.log('PWA ServiceWorker registration failed: ', err)
    );
  });
}

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </StrictMode>,
);
