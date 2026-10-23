function ErrorMessage({ message, onRetry }) {
  return (
    <div className="error-message-card" role="alert" aria-live="assertive">
      <div className="error-icon">⚠️</div>
      <h3 className="error-title">Failed to Load Repositories</h3>
      <p className="error-detail">{message || 'An unexpected error occurred while fetching data.'}</p>
      {/* Supplementary Problem 1: Retry Button */}
      {onRetry && (
        <button
          type="button"
          className="btn-retry"
          onClick={onRetry}
          aria-label="Retry fetching repositories"
        >
          🔄 Try Again
        </button>
      )}
    </div>
  );
}

export default ErrorMessage;
