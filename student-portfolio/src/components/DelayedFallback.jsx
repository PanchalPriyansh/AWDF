import { useState, useEffect } from 'react';

/**
 * DelayedFallback Component (Practical 8 - Supplementary Problem 2)
 * Prevents jarring UI flicker on fast connections by delaying the loading indicator by 300ms.
 * If the lazy-loaded chunk arrives within 300ms, no loading indicator is rendered.
 * If loading exceeds 300ms (e.g. on 3G), a clean spinner is displayed.
 */
function DelayedFallback({ delay = 300, label = 'Loading page resources...' }) {
  const [show, setShow] = useState(false);

  useEffect(() => {
    const timer = setTimeout(() => {
      setShow(true);
    }, delay);

    return () => clearTimeout(timer);
  }, [delay]);

  if (!show) {
    // Return null during the initial delay window to eliminate flickering
    return null;
  }

  return (
    <div
      className="delayed-fallback-wrapper"
      role="status"
      aria-live="polite"
      style={{
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        minHeight: '40vh',
        padding: '2rem',
        gap: '1rem',
      }}
    >
      <div
        className="lazy-spinner-ring"
        style={{
          width: '48px',
          height: '48px',
          borderRadius: '50%',
          border: '4px solid #e2e8f0',
          borderTopColor: '#1d4ed8',
          animation: 'spin 0.8s linear infinite',
        }}
      />
      <p style={{ color: 'var(--text-muted, #64748b)', fontSize: '0.95rem', fontWeight: 500 }}>
        {label}
      </p>
    </div>
  );
}

export default DelayedFallback;
