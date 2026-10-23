import React from 'react';

/**
 * Footer Component (Memoized for React DevTools Profiler Optimization - Practical 8 Supplementary 3)
 */
function Footer({ name, year = 2026, email, github }) {
  return (
    <footer id="contact" className="portfolio-footer">
      <div className="footer-container">
        <h3 className="footer-title">Get In Touch</h3>
        <p className="footer-subtitle">
          Interested in working together or have questions? Feel free to reach out.
        </p>
        <div className="footer-links">
          {email && (
            <a href={`mailto:${email}`} className="footer-chip">
              ✉️ {email}
            </a>
          )}
          {github && (
            <a
              href={`https://github.com/${github}`}
              target="_blank"
              rel="noreferrer"
              className="footer-chip"
            >
              💻 github.com/{github}
            </a>
          )}
        </div>
        <div className="footer-bottom">
          <p>© {year} {name}. All rights reserved. Built with React & Vite.</p>
        </div>
      </div>
    </footer>
  );
}

// React.memo prevents unnecessary re-renders when parent states (theme/routes) change
export default React.memo(Footer);
