function ConfirmModal({ isOpen, title, message, onConfirm, onCancel, isProcessing }) {
  if (!isOpen) return null;

  return (
    <div className="modal-backdrop" role="dialog" aria-modal="true" aria-labelledby="modal-title">
      <div className="modal-card">
        <div className="modal-header">
          <span className="modal-icon">🗑️</span>
          <h3 id="modal-title" className="modal-title">
            {title || 'Confirm Deletion'}
          </h3>
        </div>
        <p className="modal-message">{message || 'Are you sure you want to proceed with this deletion?'}</p>
        <div className="modal-actions">
          <button
            type="button"
            className="btn-secondary"
            onClick={onCancel}
            disabled={isProcessing}
          >
            Cancel
          </button>
          <button
            type="button"
            className="btn-danger-confirm"
            onClick={onConfirm}
            disabled={isProcessing}
          >
            {isProcessing ? 'Deleting...' : 'Delete Task'}
          </button>
        </div>
      </div>
    </div>
  );
}

export default ConfirmModal;
