import React from 'react';
import { Navigate } from 'react-router-dom';
import { useAuth } from '../contexts/AuthContext';
import ChangePassword from './ChangePassword';

export default function ProtectedRoute({ children }) {
  const { session, profile, loading } = useAuth();

  if (loading) {
    return <div className="centered-message">Loading…</div>;
  }

  if (!session) {
    return <Navigate to="/login" replace />;
  }

  if (!profile) {
    // Signed in to Supabase Auth, but no matching app_users row — this
    // shouldn't happen for anyone created via the provisioning script, but
    // fail with a clear, actionable message rather than a blank screen.
    return (
      <div className="centered-message">
        Your account isn't set up yet. Please contact your administrator.
      </div>
    );
  }

  if (profile.is_default_password) {
    return <ChangePassword />;
  }

  return children;
}
