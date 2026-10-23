function Spinner({ message = 'Loading repositories from GitHub...' }) {
  return (
    <div className="spinner-container" role="status" aria-live="polite">
      {/* Semantic progress element adhering to modern web standards */}
      <progress aria-label={message} className="loading-spinner"></progress>
      <p className="spinner-text">{message}</p>
    </div>
  );
}

export default Spinner;
