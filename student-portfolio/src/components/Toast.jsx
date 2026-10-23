import { useEffect } from 'react';

function Toast({ message, type = 'success', onClose, duration = 4000 }) {
  useEffect(() => {
    if (!message) return;
    const timer = setTimeout(() => {
      if (onClose) onClose();
    }, duration);

    return () => clearTimeout(timer);
  }, [message, duration, onClose]);

  if (!message) return null;

  const isSuccess = type === 'success';

  return (
    <div
      className={`toast-notification ${isSuccess ? 'toast-success' : 'toast-error'}`}
      role="status"
      aria-live="polite"
    >
      <span className="toast-icon">{isSuccess ? '✅' : '⚠️'}</span>
      <p className="toast-message">{message}</p>
      <button
        type="button"
        className="toast-close-btn"
        onClick={onClose}
        aria-label="Dismiss notification"
      >
        ✕
      </button>
    </div>
  );
}

export default Toast;
