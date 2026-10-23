import { useState } from 'react';

function Contact({ studentEmail, studentGithub }) {
  // Controlled input states (Problem Definition & Step 5)
  const [name, setName] = useState('');
  const [email, setEmail] = useState('');
  const [message, setMessage] = useState('');

  // UI Visibility Toggle state (Problem Definition & Step 6)
  const [showTips, setShowTips] = useState(false);

  // Form submission feedback state
  const [submittedData, setSubmittedData] = useState(null);

  const handleSubmit = (e) => {
    e.preventDefault();
    if (message.trim()) {
      setSubmittedData({ name, email, message });
    }
  };

  const handleReset = () => {
    setName('');
    setEmail('');
    setMessage('');
    setSubmittedData(null);
  };

  const MAX_CHARS = 500;

  return (
    <div className="portfolio-content page-contact">
      <div className="section-title-wrap">
        <h1 className="page-heading">Contact Me</h1>
        <p className="page-subheading">
          Have a question or want to collaborate on a software project? Send a message below!
        </p>
        <div className="section-underline"></div>
      </div>

      <div className="contact-layout">
        {/* Contact Form Section */}
        <section className="contact-card form-section" aria-labelledby="form-heading">
          <div className="form-header-bar">
            <h2 id="form-heading" className="contact-card-title">Send a Message</h2>
            {/* Toggle UI Visibility Button (Step 6 & Problem Definition) */}
            <button
              type="button"
              className="btn-toggle-tips"
              onClick={() => setShowTips((prev) => !prev)}
              aria-expanded={showTips}
            >
              {showTips ? '🙈 Hide Tips' : '💡 Help & Guidelines'}
            </button>
          </div>

          {/* Collapsible Element Toggled via useState (Step 6) */}
          {showTips && (
            <aside className="help-tooltip-box" role="region" aria-label="Contact Guidelines">
              <strong>Quick Tips for Outreach:</strong>
              <ul>
                <li>Mention your subject, project idea, or collaboration scope clearly.</li>
                <li>You can also reach out directly via email: <code>{studentEmail}</code>.</li>
                <li>Form input is reflected live on the preview card on the right in real time.</li>
              </ul>
            </aside>
          )}

          <form onSubmit={handleSubmit} className="contact-form">
            <div className="form-group">
              <label htmlFor="contact-name">Your Name</label>
              <input
                id="contact-name"
                type="text"
                className="form-input"
                placeholder="e.g. John Doe"
                value={name}
                onChange={(e) => setName(e.target.value)}
                required
              />
            </div>

            <div className="form-group">
              <label htmlFor="contact-email">Your Email Address</label>
              <input
                id="contact-email"
                type="email"
                className="form-input"
                placeholder="e.g. john@example.com"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                required
              />
            </div>

            <div className="form-group">
              <label htmlFor="contact-message">Message</label>
              {/* Controlled input tied to useState (Step 5) */}
              <textarea
                id="contact-message"
                className="form-input form-textarea"
                rows="5"
                placeholder="Type your message here..."
                maxLength={MAX_CHARS}
                value={message}
                onChange={(e) => setMessage(e.target.value)}
                required
              />
              {/* Live Character Count (Supplementary Problem 3) */}
              <div className="char-counter" aria-live="polite">
                <span>Live Character Count:</span>
                <strong className={message.length >= MAX_CHARS ? 'char-limit' : ''}>
                  {message.length} / {MAX_CHARS}
                </strong>
              </div>
            </div>

            <div className="form-actions">
              <button type="submit" className="btn-primary" disabled={!message.trim()}>
                Send Message 🚀
              </button>
              <button type="button" className="btn-secondary" onClick={handleReset}>
                Reset
              </button>
            </div>
          </form>

          {submittedData && (
            <div className="submission-success-banner" role="status">
              ✅ Thank you, <strong>{submittedData.name}</strong>! Your message has been captured.
            </div>
          )}
        </section>

        {/* Live Real-time Preview Section (Problem Definition) */}
        <aside className="contact-card preview-section" aria-labelledby="preview-heading">
          <h2 id="preview-heading" className="contact-card-title">Live Input Preview</h2>
          <p className="preview-subtext">
            Captures and displays your input reactively in real time:
          </p>

          <div className="live-preview-box">
            <div className="preview-field">
              <span className="preview-label">Name:</span>
              <span className="preview-value">{name || <em className="placeholder-text">Waiting for input...</em>}</span>
            </div>
            <div className="preview-field">
              <span className="preview-label">Email:</span>
              <span className="preview-value">{email || <em className="placeholder-text">Waiting for input...</em>}</span>
            </div>
            <div className="preview-field preview-field-message">
              <span className="preview-label">Message:</span>
              <div className="preview-message-content">
                {message ? (
                  <p>{message}</p>
                ) : (
                  <em className="placeholder-text">Your typed message will appear here live...</em>
                )}
              </div>
            </div>
          </div>

          <div className="direct-connect-box">
            <h3>Direct Contact Details</h3>
            <p><strong>Email:</strong> <a href={`mailto:${studentEmail}`}>{studentEmail}</a></p>
            <p><strong>GitHub:</strong> <a href={`https://github.com/${studentGithub}`} target="_blank" rel="noreferrer">@{studentGithub}</a></p>
          </div>
        </aside>
      </div>
    </div>
  );
}

export default Contact;
