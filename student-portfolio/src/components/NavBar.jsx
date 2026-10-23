import React from 'react';
import { NavLink, Link } from 'react-router-dom';

/**
 * NavBar Component (Memoized for React DevTools Profiler Optimization - Practical 8 Supplementary 3)
 */
function NavBar({ brand = 'Priyansh Panchal', darkMode = false, onToggleTheme }) {
  const navItems = [
    { path: '/', label: 'Home' },
    { path: '/projects', label: 'Tasks' },
    { path: '/contact', label: 'Contact' },
  ];

  return (
    <nav className="portfolio-nav" aria-label="Main Application Navigation">
      <div className="nav-container">
        {/* Brand link navigating Home using Link to avoid full page reload */}
        <Link to="/" className="nav-brand">
          {brand}
        </Link>

        <div className="nav-controls">
          <ul className="nav-list">
            {navItems.map((item) => (
              <li key={item.path} className="nav-item">
                {/* React Router NavLink for seamless client-side routing */}
                <NavLink
                  to={item.path}
                  end={item.path === '/'}
                  className={({ isActive }) =>
                    `nav-link ${isActive ? 'active' : ''}`
                  }
                >
                  {item.label}
                </NavLink>
              </li>
            ))}
          </ul>

          {/* Dark / Light Mode Toggle Button */}
          {onToggleTheme && (
            <button
              type="button"
              className="theme-toggle-btn"
              onClick={onToggleTheme}
              aria-label={`Switch to ${darkMode ? 'Light' : 'Dark'} mode`}
              title={`Switch to ${darkMode ? 'Light' : 'Dark'} mode`}
            >
              {darkMode ? '☀️ Light' : '🌙 Dark'}
            </button>
          )}
        </div>
      </div>
    </nav>
  );
}

// React.memo prevents re-rendering when props have not changed
export default React.memo(NavBar);
