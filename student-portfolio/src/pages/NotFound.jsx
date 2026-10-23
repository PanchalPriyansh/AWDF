import { Link } from 'react-router-dom';

function NotFound() {
  return (
    <div className="portfolio-content page-not-found">
      <div className="not-found-card">
        <span className="error-code">404</span>
        <h1 className="error-title">Page Not Found</h1>
        <p className="error-description">
          Oops! The page or route you are looking for does not exist in this student portfolio.
        </p>
        <div className="not-found-actions">
          <Link to="/" className="btn-primary">
            ← Return to Home
          </Link>
          <Link to="/projects" className="btn-secondary">
            Browse Projects
          </Link>
        </div>
      </div>
    </div>
  );
}

export default NotFound;
