import React, { useState } from 'react';
import { useAuth } from '../contexts/AuthContext';
import './Login.css';

export default function Login() {
  const { signInWithEmployeeCode } = useAuth();
  const [employeeCode, setEmployeeCode] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [submitting, setSubmitting] = useState(false);

  async function handleSubmit(e) {
    e.preventDefault();
    setError('');
    setSubmitting(true);
    const { error } = await signInWithEmployeeCode(employeeCode.trim(), password);
    setSubmitting(false);
    if (error) {
      setError('Employee code or password is incorrect.');
    }
  }

  return (
    <div className="login-screen">
      <div className="login-card">
        <div className="login-mark">EXCELLENCE</div>
        <h1>Sign in</h1>
        <p className="login-sub">Use your employee code and password.</p>

        <form onSubmit={handleSubmit}>
          <label htmlFor="employeeCode">Employee code</label>
          <input
            id="employeeCode"
            type="text"
            inputMode="numeric"
            autoComplete="username"
            value={employeeCode}
            onChange={(e) => setEmployeeCode(e.target.value)}
            required
            autoFocus
          />

          <label htmlFor="password">Password</label>
          <input
            id="password"
            type="password"
            autoComplete="current-password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            required
          />

          {error && <div className="login-error" role="alert">{error}</div>}

          <button type="submit" disabled={submitting}>
            {submitting ? 'Signing in…' : 'Sign in'}
          </button>
        </form>

        <p className="login-hint">
          First time signing in? Your password is your employee code.
          You'll be asked to set a new one.
        </p>
      </div>
    </div>
  );
}
