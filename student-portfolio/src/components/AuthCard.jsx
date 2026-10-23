import { useState } from 'react';
import { loginUser, registerUser } from '../services/api';

/**
 * Authentication Card Component (Practical 7 - Steps 3 & 4)
 * Handles client-side user registration and login with JWT token acquisition.
 */
function AuthCard({ onAuthSuccess, showToast }) {
  const [isLoginView, setIsLoginView] = useState(true);
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');

  const handleSubmit = async (e) => {
    e.preventDefault();
    setErrorMsg('');

    if (!email.trim() || !password.trim()) {
      setErrorMsg('Please enter both email and password.');
      return;
    }

    if (password.length < 6) {
      setErrorMsg('Password must be at least 6 characters long.');
      return;
    }

    setLoading(true);

    try {
      if (isLoginView) {
        // Step 4: Login and acquire JWT
        const res = await loginUser(email.trim(), password);
        showToast('Login successful! JWT token saved.', 'success');
        if (onAuthSuccess) {
          onAuthSuccess(res.user, res.token);
        }
      } else {
        // Step 3: Register user with hashed password
        const res = await registerUser(email.trim(), password);
        showToast(res.message || 'Registration successful! Logging you in...', 'success');
        // Automatically log in after registration
        const loginRes = await loginUser(email.trim(), password);
        if (onAuthSuccess) {
          onAuthSuccess(loginRes.user, loginRes.token);
        }
      }
    } catch (err) {
      setErrorMsg(err.message || 'Authentication failed. Please verify credentials.');
      showToast(err.message || 'Authentication error', 'error');
    } finally {
      setLoading(false);
    }
  };

  const handleDemoFill = () => {
    setEmail('panchalpriyansh3257@gmail.com');
    setPassword('Priyansh@123');
  };

  return (
    <div className="auth-card" style={{
      maxWidth: '460px',
      margin: '2rem auto',
      padding: '2rem',
      backgroundColor: 'var(--card-bg, #ffffff)',
      borderRadius: '12px',
      boxShadow: '0 4px 20px rgba(0,0,0,0.08)',
      border: '1px solid var(--border-color, #e2e8f0)',
      textAlign: 'left'
    }}>
      <div style={{ textAlign: 'center', marginBottom: '1.5rem' }}>
        <h2 style={{ margin: '0 0 0.5rem', fontSize: '1.5rem' }}>
          {isLoginView ? '🔐 Secure User Login' : '📝 Create New Account'}
        </h2>
        <p style={{ margin: 0, color: 'var(--text-muted, #64748b)', fontSize: '0.9rem' }}>
          JWT-protected Task Management System (Practical 7)
        </p>
      </div>

      {/* Tabs */}
      <div style={{ display: 'flex', gap: '8px', marginBottom: '1.5rem', borderBottom: '1px solid #e2e8f0', paddingBottom: '8px' }}>
        <button
          type="button"
          onClick={() => { setIsLoginView(true); setErrorMsg(''); }}
          style={{
            flex: 1,
            padding: '8px 12px',
            border: 'none',
            borderRadius: '6px',
            cursor: 'pointer',
            fontWeight: 600,
            backgroundColor: isLoginView ? '#1d4ed8' : 'transparent',
            color: isLoginView ? '#ffffff' : '#64748b',
            transition: 'all 0.2s ease',
          }}
        >
          Login
        </button>
        <button
          type="button"
          onClick={() => { setIsLoginView(false); setErrorMsg(''); }}
          style={{
            flex: 1,
            padding: '8px 12px',
            border: 'none',
            borderRadius: '6px',
            cursor: 'pointer',
            fontWeight: 600,
            backgroundColor: !isLoginView ? '#1d4ed8' : 'transparent',
            color: !isLoginView ? '#ffffff' : '#64748b',
            transition: 'all 0.2s ease',
          }}
        >
          Register
        </button>
      </div>

      {errorMsg && (
        <div style={{
          backgroundColor: '#fef2f2',
          color: '#991b1b',
          border: '1px solid #fecaca',
          borderRadius: '6px',
          padding: '10px 14px',
          marginBottom: '1rem',
          fontSize: '0.875rem'
        }}>
          ⚠️ {errorMsg}
        </div>
      )}

      <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: '1rem' }}>
        <div>
          <label htmlFor="auth-email" style={{ display: 'block', fontWeight: 500, fontSize: '0.875rem', marginBottom: '4px' }}>
            Email Address
          </label>
          <input
            id="auth-email"
            type="email"
            required
            placeholder="e.g. panchalpriyansh3257@gmail.com"
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            style={{
              width: '100%',
              padding: '10px 12px',
              border: '1px solid #cbd5e1',
              borderRadius: '6px',
              fontSize: '0.95rem',
              boxSizing: 'border-box'
            }}
          />
        </div>

        <div>
          <label htmlFor="auth-password" style={{ display: 'block', fontWeight: 500, fontSize: '0.875rem', marginBottom: '4px' }}>
            Password (min 6 characters)
          </label>
          <input
            id="auth-password"
            type="password"
            required
            placeholder="Enter secure password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            style={{
              width: '100%',
              padding: '10px 12px',
              border: '1px solid #cbd5e1',
              borderRadius: '6px',
              fontSize: '0.95rem',
              boxSizing: 'border-box'
            }}
          />
        </div>

        <button
          type="submit"
          disabled={loading}
          style={{
            padding: '12px',
            backgroundColor: '#1d4ed8',
            color: '#ffffff',
            border: 'none',
            borderRadius: '6px',
            fontWeight: 600,
            fontSize: '1rem',
            cursor: loading ? 'not-allowed' : 'pointer',
            opacity: loading ? 0.7 : 1,
            transition: 'background 0.2s ease',
          }}
        >
          {loading ? 'Authenticating...' : (isLoginView ? 'Sign In (Generate JWT)' : 'Create Account (Hash with bcrypt)')}
        </button>

        <button
          type="button"
          onClick={handleDemoFill}
          style={{
            padding: '8px',
            backgroundColor: 'transparent',
            color: '#64748b',
            border: '1px dashed #cbd5e1',
            borderRadius: '6px',
            fontSize: '0.8rem',
            cursor: 'pointer',
          }}
        >
          Fill Demo Credentials (Priyansh Panchal)
        </button>
      </form>
    </div>
  );
}

export default AuthCard;
